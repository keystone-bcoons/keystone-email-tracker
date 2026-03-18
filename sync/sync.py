"""
Keystone Email Response Tracker — Sync Script
Pulls email data from Microsoft 365 Graph API → Supabase

Required environment variables:
  AZURE_TENANT_ID       - from Azure App Registration
  AZURE_CLIENT_ID       - from Azure App Registration
  AZURE_CLIENT_SECRET   - from Azure App Registration
  SUPABASE_URL          - from Supabase project settings
  SUPABASE_SERVICE_KEY  - service_role key (not anon key)
  TARGET_HOURS          - response time target (default: 48)
  LOOKBACK_DAYS         - how far back on first run (default: 90)
"""

import os
import logging
import re
import time
from datetime import datetime, timezone, timedelta
from typing import Optional
import requests
from supabase import create_client, Client

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)

# ─────────────────────────────────────────
# Config
# ─────────────────────────────────────────
TENANT_ID      = os.environ["AZURE_TENANT_ID"]
CLIENT_ID      = os.environ["AZURE_CLIENT_ID"]
CLIENT_SECRET  = os.environ["AZURE_CLIENT_SECRET"]
SUPABASE_URL   = os.environ["SUPABASE_URL"]
SUPABASE_KEY   = os.environ["SUPABASE_SERVICE_KEY"]
TARGET_HOURS   = int(os.environ.get("TARGET_HOURS", 48))
LOOKBACK_DAYS  = int(os.environ.get("LOOKBACK_DAYS", 90))

GRAPH_BASE     = "https://graph.microsoft.com/v1.0"
GRAPH_SCOPES   = "https://graph.microsoft.com/.default"
PAGE_SIZE      = 999

# ─────────────────────────────────────────
# Automated sender filtering
# ─────────────────────────────────────────
# Local-part prefixes that indicate automated/system senders
AUTOMATED_PREFIXES = {
    "noreply",
    "no-reply",
    "no_reply",
    "donotreply",
    "do-not-reply",
    "do_not_reply",
    "notifications",
    "notification",
    "alerts",
    "alert",
    "automated",
    "automailer",
    "auto",
    "mailer",
    "mailer-daemon",
    "mailerdaemon",
    "bounce",
    "bounces",
    "info",
    "support",
    "noreply+",
}

# Domains that are known automated/system senders
AUTOMATED_DOMAINS = {
    "breezeway.io",
    "turno.com",
    "ownerrez.com",
    "hostaway.com",
    "guesty.com",
    "lodgify.com",
    "vacasa.com",
    "vrbo.com",
    "airbnb.com",
    "booking.com",
    "tripadvisor.com",
    "homeaway.com",
    # Generic automation platforms
    "sendgrid.net",
    "mailchimp.com",
    "mandrillapp.com",
    "amazonses.com",
    "bounce.stripe.com",
    "em.stripe.com",
}

# Regex for catching common automated patterns not covered by exact prefix matching
_AUTOMATED_RE = re.compile(
    r"^(noreply|no[_\-]?reply|donotreply|do[_\-]?not[_\-]?reply|"
    r"notification[s]?|alert[s]?|automated?|auto|mailer|bounce[s]?|"
    r"postmaster|daemon|system|robot|bot)[+\-_.]?",
    re.IGNORECASE,
)


def is_automated_sender(email: str) -> bool:
    """Return True if the email address looks like an automated/system sender."""
    if not email:
        return False
    email = email.strip().lower()
    if "@" not in email:
        return False
    local, domain = email.split("@", 1)

    # Domain-based block
    if domain in AUTOMATED_DOMAINS:
        return True

    # Exact prefix match
    if local in AUTOMATED_PREFIXES:
        return True

    # Regex match on local part
    if _AUTOMATED_RE.match(local):
        return True

    return False


# ─────────────────────────────────────────
# Auth
# ─────────────────────────────────────────
def get_access_token() -> str:
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    resp = requests.post(url, data={
        "grant_type":    "client_credentials",
        "client_id":     CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope":         GRAPH_SCOPES,
    })
    resp.raise_for_status()
    return resp.json()["access_token"]


def graph_get(token: str, url: str, params: dict = None, retries: int = 3) -> dict:
    headers = {"Authorization": f"Bearer {token}"}
    for attempt in range(retries):
        try:
            resp = requests.get(url, headers=headers, params=params, timeout=60)
            if resp.status_code == 504:
                wait = 30 * (attempt + 1)
                log.warning(f"504 timeout, retrying in {wait}s (attempt {attempt+1}/{retries})")
                time.sleep(wait)
                continue
            resp.raise_for_status()
            return resp.json()
        except requests.HTTPError as e:
            if attempt == retries - 1:
                raise
            time.sleep(30)
    raise Exception(f"Failed after {retries} retries: {url}")


# ─────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────
def business_days_between(start: datetime, end: datetime) -> float:
    if end <= start:
        return 0.0
    total_seconds = 0.0
    current = start
    while current < end:
        if current.weekday() < 5:
            next_tick = min(current + timedelta(hours=1), end)
            total_seconds += (next_tick - current).total_seconds()
        current += timedelta(hours=1)
    return total_seconds / 86400


def parse_dt(s: str) -> Optional[datetime]:
    if not s:
        return None
    try:
        return datetime.fromisoformat(s.replace("Z", "+00:00"))
    except Exception:
        return None


def extract_email(address_obj: dict) -> tuple[str, str]:
    if not address_obj:
        return "", ""
    ea = address_obj.get("emailAddress", address_obj)
    return ea.get("address", "").lower(), ea.get("name", "")


# ─────────────────────────────────────────
# Graph API
# ─────────────────────────────────────────
def get_all_users(token: str) -> list[dict]:
    users = []
    url = f"{GRAPH_BASE}/users"
    params = {
        "$select": "id,displayName,mail,userPrincipalName",
        "$filter": "assignedLicenses/$count ne 0 and accountEnabled eq true",
        "$count": "true",
        "$top": 999,
    }
    while url:
        resp = requests.get(url, headers={
            "Authorization": f"Bearer {token}",
            "ConsistencyLevel": "eventual",
        }, params=params)
        resp.raise_for_status()
        data = resp.json()
        for u in data.get("value", []):
            email = (u.get("mail") or u.get("userPrincipalName", "")).lower()
            if email:
                users.append({"id": u["id"], "email": email, "name": u.get("displayName", "")})
        url = data.get("@odata.nextLink")
        params = None
    log.info(f"Found {len(users)} users")
    return users


def fetch_inbox_messages(token: str, user_id: str, since: datetime) -> list[dict]:
    """
    Fetch received messages across ALL folders (not just Inbox).
    Karbon moves cleared emails to Archive, so Inbox-only would miss the vast
    majority of messages.  /users/{id}/messages covers every folder.
    """
    messages = []
    since_str = since.strftime("%Y-%m-%dT%H:%M:%SZ")
    url = f"{GRAPH_BASE}/users/{user_id}/messages"
    params = {
        "$select": "id,subject,from,toRecipients,receivedDateTime,sentDateTime,conversationId,internetMessageId",
        "$filter": f"receivedDateTime ge {since_str} and isDraft eq false",
        "$top": PAGE_SIZE,
        "$orderby": "receivedDateTime asc",
    }
    while url:
        try:
            data = graph_get(token, url, params)
        except Exception as e:
            log.warning(f"Skipping remaining messages for {user_id}: {e}")
            break
        messages.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
        params = None
    return messages


def fetch_sent_messages(token: str, user_id: str, since: datetime) -> list[dict]:
    """Fetch messages from the Sent Items folder, filtered by sentDateTime."""
    messages = []
    since_str = since.strftime("%Y-%m-%dT%H:%M:%SZ")
    url = f"{GRAPH_BASE}/users/{user_id}/mailFolders/SentItems/messages"
    params = {
        "$select": "id,subject,from,toRecipients,receivedDateTime,sentDateTime,conversationId,internetMessageId",
        "$filter": f"sentDateTime ge {since_str}",
        "$top": PAGE_SIZE,
        "$orderby": "sentDateTime asc",
    }
    while url:
        try:
            data = graph_get(token, url, params)
        except Exception as e:
            log.warning(f"Skipping remaining sent messages for {user_id}: {e}")
            break
        messages.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
        params = None
    return messages


def merge_and_deduplicate(inbox: list[dict], sent: list[dict]) -> list[dict]:
    """
    Merge inbox and sent messages, deduplicating by Graph message ID.
    Graph message IDs are unique within a mailbox, so no cross-folder
    duplicates are expected — but this guards against any overlap.
    """
    seen: set[str] = set()
    merged: list[dict] = []
    for m in inbox + sent:
        mid = m.get("id", "")
        if mid and mid not in seen:
            seen.add(mid)
            merged.append(m)
    return merged


# ─────────────────────────────────────────
# Response time calculation
# ─────────────────────────────────────────
def calculate_responses_from_db(sb: Client, team_email: str, team_name: str, since: datetime) -> list[dict]:
    """
    Match inbound → outbound pairs by querying the DB for ALL messages stored
    for this team member since `since`.  This is more correct than matching
    only within the current fetch batch because:
      - A client email and the reply may arrive in different sync runs.
      - After a full re-fetch we want to recalculate all pairs cleanly.
    """
    all_msgs = []
    offset = 0
    batch_size = 1000
    while True:
        result = (
            sb.table("email_messages")
            .select("id,direction,conversation_id,client_email,client_name,subject,received_at")
            .eq("team_member_email", team_email)
            .gte("received_at", since.isoformat())
            .order("received_at", desc=False)
            .range(offset, offset + batch_size - 1)
            .execute()
        )
        batch = result.data or []
        all_msgs.extend(batch)
        if len(batch) < batch_size:
            break
        offset += batch_size

    by_conv: dict[str, list[dict]] = {}
    for m in all_msgs:
        cid = m.get("conversation_id", "")
        if not cid:
            continue
        by_conv.setdefault(cid, []).append(m)

    responses = []
    for conv_id, thread in by_conv.items():
        thread.sort(key=lambda x: x["received_at"])
        for i, msg in enumerate(thread):
            if msg["direction"] != "inbound":
                continue
            reply = next(
                (m for m in thread[i + 1:] if m["direction"] == "outbound"),
                None,
            )
            if not reply:
                continue
            inbound_at = parse_dt(msg["received_at"])
            replied_at = parse_dt(reply["received_at"])
            if not inbound_at or not replied_at:
                continue
            resp_days = business_days_between(inbound_at, replied_at)
            within_target = resp_days <= (TARGET_HOURS / 24)
            responses.append({
                "inbound_message_id": msg["id"],
                "reply_message_id":   reply["id"],
                "team_member_email":  team_email,
                "team_member_name":   team_name,
                "client_email":       msg.get("client_email", ""),
                "client_name":        msg.get("client_name", ""),
                "subject":            (msg.get("subject") or ""),
                "inbound_at":         msg["received_at"],
                "replied_at":         reply["received_at"],
                "response_days":      round(resp_days, 4),
                "within_target":      within_target,
                "target_hours":       TARGET_HOURS,
            })
    return responses


# ─────────────────────────────────────────
# Supabase writes
# ─────────────────────────────────────────
def upsert_messages(sb: Client, records: list[dict]):
    if not records:
        return
    for i in range(0, len(records), 500):
        sb.table("email_messages").upsert(records[i:i+500], on_conflict="id").execute()


def upsert_responses(sb: Client, records: list[dict]):
    if not records:
        return
    for i in range(0, len(records), 500):
        sb.table("email_responses").upsert(
            records[i:i+500], on_conflict="inbound_message_id,reply_message_id"
        ).execute()


def get_last_synced(sb: Client, email: str) -> datetime:
    result = sb.table("sync_state").select("last_synced_at").eq("team_member_email", email).execute()
    if result.data:
        return parse_dt(result.data[0]["last_synced_at"])
    return datetime.now(timezone.utc) - timedelta(days=LOOKBACK_DAYS)


def set_last_synced(sb: Client, email: str):
    sb.table("sync_state").upsert({
        "team_member_email": email,
        "last_synced_at":    datetime.now(timezone.utc).isoformat(),
        "updated_at":        datetime.now(timezone.utc).isoformat(),
    }, on_conflict="team_member_email").execute()


# ─────────────────────────────────────────
# Main
# ─────────────────────────────────────────
def main():
    log.info("Starting sync")
    token = get_access_token()
    sb = create_client(SUPABASE_URL, SUPABASE_KEY)

    users = get_all_users(token)

    for user in users:
        email = user["email"]
        name  = user["name"]
        log.info(f"Syncing {email}")

        since = get_last_synced(sb, email)
        log.info(f"  Fetching messages since {since.date()}")

        # Fetch inbox (receivedDateTime) and sent items (sentDateTime) separately,
        # then merge and deduplicate by Graph message ID.
        inbox_messages = fetch_inbox_messages(token, user["id"], since)
        sent_messages  = fetch_sent_messages(token, user["id"], since)
        log.info(f"  Fetched {len(inbox_messages)} inbox + {len(sent_messages)} sent messages")

        raw_messages = merge_and_deduplicate(inbox_messages, sent_messages)
        log.info(f"  {len(raw_messages)} messages after deduplication")

        message_records = []
        skipped_automated = 0

        for m in raw_messages:
            from_email, from_name = extract_email(m.get("from", {}))
            is_outbound = from_email == email
            direction   = "outbound" if is_outbound else "inbound"

            # For inbound messages, skip automated/system senders entirely.
            # Outbound messages from the team member are always kept.
            if not is_outbound and is_automated_sender(from_email):
                skipped_automated += 1
                continue

            client_email, client_name = "", ""
            if is_outbound:
                recipients = m.get("toRecipients", [])
                if recipients:
                    client_email, client_name = extract_email(recipients[0])
                # Also skip outbound emails sent TO automated/system addresses
                # (e.g. replies to notification threads that go back to a no-reply)
                if is_automated_sender(client_email):
                    skipped_automated += 1
                    continue
            else:
                client_email = from_email
                client_name  = from_name

            # Use sentDateTime for outbound, receivedDateTime for inbound.
            # Fall back to whichever field is available.
            if is_outbound:
                received_at = parse_dt(m.get("sentDateTime") or m.get("receivedDateTime"))
            else:
                received_at = parse_dt(m.get("receivedDateTime") or m.get("sentDateTime"))

            if not received_at:
                continue

            rec = {
                "id":                  m["id"],
                "team_member_email":   email,
                "team_member_name":    name,
                "client_email":        client_email,
                "client_name":         client_name,
                "subject":             (m.get("subject") or "")[:500],
                "received_at":         received_at.isoformat(),
                "direction":           direction,
                "conversation_id":     m.get("conversationId", ""),
                "internet_message_id": m.get("internetMessageId", ""),
            }
            message_records.append(rec)

        log.info(f"  Skipped {skipped_automated} automated-sender messages")
        upsert_messages(sb, message_records)
        log.info(f"  Upserted {len(message_records)} messages")

        # Match responses from the full DB record for this user (not just this batch).
        # This handles the case where the inbound email and the reply were fetched
        # in separate sync runs (e.g. daily incremental syncs).
        responses = calculate_responses_from_db(sb, email, name, since)
        upsert_responses(sb, responses)
        log.info(f"  Upserted {len(responses)} response records")

        set_last_synced(sb, email)

    log.info("Sync complete")


if __name__ == "__main__":
    main()
