import { useState, useEffect, useMemo } from "react";
import { createClient } from "@supabase/supabase-js";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Cell,
} from "recharts";

const supabase = createClient(
  import.meta.env.VITE_SUPABASE_URL,
  import.meta.env.VITE_SUPABASE_ANON_KEY
);

const EXCLUDED_DOMAINS = [
  "keystonebookkeepers.com",
  "topkey.io",
];

const PAGE_SIZE = 50;

const AUTO_REPLY_PATTERNS = [
  "automatic reply", "out of office", "undeliverable", "delivery failed",
  "mail delivery", "noreply", "no-reply", "do not reply", "donotreply",
  "mailer-daemon", "postmaster", "notification", "[govos]", "auto-reply",
  "auto reply", "away from", "vacation reply",
];

function isAutoReply(subject = "", clientEmail = "") {
  const s = subject.toLowerCase();
  const e = clientEmail.toLowerCase();
  return (
    AUTO_REPLY_PATTERNS.some(p => s.includes(p)) ||
    e.includes("noreply") || e.includes("no-reply") ||
    e.includes("donotreply") || e.includes("notifications@") ||
    e.includes("mailer-daemon")
  );
}

function isInternal(clientEmail = "") {
  const e = clientEmail.toLowerCase();
  return EXCLUDED_DOMAINS.some(domain => e.includes(domain));
}

function fmtDays(d) {
  if (d === null || d === undefined) return "—";
  const totalMins = Math.round(d * 24 * 60);
  if (totalMins < 60) return `${totalMins}m`;
  const hours = Math.floor(totalMins / 60);
  const mins  = totalMins % 60;
  if (hours < 24) return mins > 0 ? `${hours}h ${mins}m` : `${hours}h`;
  const days = Math.floor(hours / 24);
  const remH = hours % 24;
  return remH > 0 ? `${days}d ${remH}h` : `${days}d`;
}

function bucketLabel(days) {
  if (days < 1)  return "< 1 Day";
  if (days < 2)  return "1–2 Days";
  if (days < 5)  return "2–5 Days";
  if (days < 10) return "5–10 Days";
  return "> 10 Days";
}

const BUCKET_ORDER = ["< 1 Day", "1–2 Days", "2–5 Days", "5–10 Days", "> 10 Days"];

function KpiCard({ label, value, sub }) {
  return (
    <div className="kpi-card">
      <div className="kpi-value">{value}</div>
      <div className="kpi-label">{label}</div>
      {sub && <div className="kpi-sub">{sub}</div>}
    </div>
  );
}

function CustomTooltip({ active, payload, label }) {
  if (!active || !payload?.length) return null;
  return (
    <div className="chart-tooltip">
      <div className="tooltip-label">{label}</div>
      <div className="tooltip-value">{payload[0].value} responses</div>
    </div>
  );
}

function exportCSV(rows) {
  const headers = ["Date", "Team Member", "Client Email", "Client Name", "Subject", "Response Time", "Within Target"];
  const lines = [
    headers.join(","),
    ...rows.map(r => [
      new Date(r.inbound_at).toLocaleDateString(),
      r.team_member_name || r.team_member_email,
      r.client_email || "",
      r.client_name || "",
      `"${(r.subject || "").replace(/"/g, '""')}"`,
      fmtDays(r.response_days),
      r.within_target ? "Yes" : "No",
    ].join(",")),
  ];
  const blob = new Blob([lines.join("\n")], { type: "text/csv" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `email-response-time-${new Date().toISOString().split("T")[0]}.csv`;
  a.click();
  URL.revokeObjectURL(url);
}

function SortIcon({ active, dir }) {
  if (!active) return <span className="sort-icon sort-inactive">↕</span>;
  return <span className="sort-icon sort-active">{dir === "asc" ? "↑" : "↓"}</span>;
}

export default function App() {
  const [responses, setResponses]   = useState([]);
  const [loading, setLoading]       = useState(true);
  const [loadingMsg, setLoadingMsg] = useState("Loading…");
  const [error, setError]           = useState(null);
  const [lastSynced, setLastSynced] = useState(null);
  const [page, setPage]             = useState(1);
  const [inboundCount, setInboundCount] = useState(null);

  // Top filters
  const [dateFrom, setDateFrom] = useState(() => {
    const d = new Date(); d.setMonth(d.getMonth() - 3);
    return d.toISOString().split("T")[0];
  });
  const [dateTo, setDateTo]           = useState(() => new Date().toISOString().split("T")[0]);
  const [teamMember, setTeamMember]   = useState("All");
  const [targetHours, setTargetHours] = useState(48);
  const [emailSource, setEmailSource] = useState("clients");
  const [excludeAutoReply, setExcludeAutoReply] = useState(true);

  // Column filters
  const [colClient,  setColClient]  = useState("");
  const [colSubject, setColSubject] = useState("");
  const [colTarget,  setColTarget]  = useState("all");

  // Sorting
  const [sortCol, setSortCol] = useState("inbound_at");
  const [sortDir, setSortDir] = useState("desc");

  function handleSort(col) {
    if (sortCol === col) setSortDir(d => d === "asc" ? "desc" : "asc");
    else { setSortCol(col); setSortDir("desc"); }
    setPage(1);
  }

  useEffect(() => {
    async function load() {
      setLoading(true);
      setLoadingMsg("Loading…");
      setError(null);
      setPage(1);
      try {
        let allData = [];
        let from = 0;
        const batchSize = 1000;
        while (true) {
          setLoadingMsg(`Loading… (${allData.length} records)`);
          const { data, error: err } = await supabase
            .from("email_responses")
            .select("*")
            .gte("inbound_at", dateFrom)
            .lte("inbound_at", dateTo + "T23:59:59Z")
            .order("inbound_at", { ascending: false })
            .range(from, from + batchSize - 1);
          if (err) throw err;
          allData = [...allData, ...(data || [])];
          if (!data || data.length < batchSize) break;
          from += batchSize;
        }
        setResponses(allData);

        const { data: syncData } = await supabase
          .from("sync_state")
          .select("last_synced_at")
          .order("last_synced_at", { ascending: false })
          .limit(1);
        if (syncData?.[0]) setLastSynced(new Date(syncData[0].last_synced_at));
      } catch (e) {
        setError(e.message);
      } finally {
        setLoading(false);
      }
    }
    load();
  }, [dateFrom, dateTo]);

  // Separately fetch true inbound email count from email_messages table.
  // This is the real "Total Emails" — not just the subset that have a matched response.
  useEffect(() => {
    async function loadInboundCount() {
      setInboundCount(null);
      let q = supabase
        .from("email_messages")
        .select("*", { count: "exact", head: true })
        .eq("direction", "inbound")
        .gte("received_at", dateFrom)
        .lte("received_at", dateTo + "T23:59:59Z");
      if (teamMember !== "All") {
        q = q.eq("team_member_name", teamMember);
      }
      const { count } = await q;
      setInboundCount(count);
    }
    loadInboundCount();
  }, [dateFrom, dateTo, teamMember]);

  const teamMembers = useMemo(() => {
    const names = [...new Set(responses.map(r => r.team_member_name || r.team_member_email))];
    return ["All", ...names.sort()];
  }, [responses]);

  const filtered = useMemo(() => {
    return responses
      .filter(r => teamMember === "All" || r.team_member_name === teamMember || r.team_member_email === teamMember)
      .filter(r => emailSource === "all" || !isInternal(r.client_email || ""))
      .filter(r => !excludeAutoReply || !isAutoReply(r.subject, r.client_email))
      .map(r => ({ ...r, within_target: r.response_days <= targetHours / 24 }));
  }, [responses, teamMember, targetHours, emailSource, excludeAutoReply]);

  const tableRows = useMemo(() => {
    let rows = filtered
      .filter(r => !colClient  || (r.client_name || r.client_email || "").toLowerCase().includes(colClient.toLowerCase()))
      .filter(r => !colSubject || (r.subject || "").toLowerCase().includes(colSubject.toLowerCase()))
      .filter(r => colTarget === "all" || (colTarget === "yes" ? r.within_target : !r.within_target));

    rows = [...rows].sort((a, b) => {
      let av, bv;
      if      (sortCol === "inbound_at") {
