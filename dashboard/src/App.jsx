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
      const { count, error: countErr } = await q;
      if (countErr) {
        console.error("inboundCount query error:", countErr);
      } else {
        console.log("inboundCount result:", count, "| member:", teamMember, "| from:", dateFrom, "| to:", dateTo);
      }
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
      if      (sortCol === "inbound_at") { av = a.inbound_at; bv = b.inbound_at; }
      else if (sortCol === "client")     { av = (a.client_name || a.client_email || "").toLowerCase(); bv = (b.client_name || b.client_email || "").toLowerCase(); }
      else if (sortCol === "member")     { av = (a.team_member_name || "").toLowerCase(); bv = (b.team_member_name || "").toLowerCase(); }
      else if (sortCol === "subject")    { av = (a.subject || "").toLowerCase(); bv = (b.subject || "").toLowerCase(); }
      else if (sortCol === "response")   { av = a.response_days; bv = b.response_days; }
      else if (sortCol === "target")     { av = a.within_target ? 1 : 0; bv = b.within_target ? 1 : 0; }
      if (av < bv) return sortDir === "asc" ? -1 : 1;
      if (av > bv) return sortDir === "asc" ? 1 : -1;
      return 0;
    });
    return rows;
  }, [filtered, colClient, colSubject, colTarget, sortCol, sortDir]);

  const totalResponses = filtered.length;
  const withinTarget   = filtered.filter(r => r.within_target).length;
  const pctWithin      = totalResponses ? (withinTarget / totalResponses * 100).toFixed(1) : "0.0";
  const avgDays        = totalResponses ? filtered.reduce((s, r) => s + r.response_days, 0) / totalResponses : 0;
  // inboundCount comes from a direct email_messages query — this is the true total.
  const totalEmails = inboundCount ?? new Set(filtered.map(r => r.inbound_message_id)).size;

  const chartData = useMemo(() => {
    const counts = Object.fromEntries(BUCKET_ORDER.map(b => [b, 0]));
    filtered.forEach(r => { counts[bucketLabel(r.response_days)]++; });
    return BUCKET_ORDER.map(b => ({ label: b, count: counts[b] }));
  }, [filtered]);

  const barColor = label =>
    label === "< 1 Day" || label === "1–2 Days" ? "#2563eb" : "#93c5fd";

  const totalPages = Math.max(1, Math.ceil(tableRows.length / PAGE_SIZE));
  const paginated  = tableRows.slice((page - 1) * PAGE_SIZE, page * PAGE_SIZE);
  const hasColFilters = colClient || colSubject || colTarget !== "all";

  return (
    <div className="app">
      <header className="header">
        <div className="header-left">
          <div className="logo">KS</div>
          <div>
            <h1 className="header-title">Email Response Time</h1>
            <p className="header-sub">Keystone Bookkeepers · Internal Dashboard</p>
          </div>
        </div>
        <div className="header-right">
          {lastSynced && (
            <span className="sync-badge">
              Last synced {lastSynced.toLocaleDateString("en-US", { month: "short", day: "numeric", hour: "2-digit", minute: "2-digit" })}
            </span>
          )}
        </div>
      </header>

      <div className="top-filters">
        <div className="top-filter-group">
          <label className="filter-label">Date Range</label>
          <div className="date-range">
            <input type="date" className="filter-input" value={dateFrom} onChange={e => { setDateFrom(e.target.value); setPage(1); }} />
            <span className="date-sep">→</span>
            <input type="date" className="filter-input" value={dateTo} onChange={e => { setDateTo(e.target.value); setPage(1); }} />
          </div>
        </div>
        <div className="top-filter-group">
          <label className="filter-label">Team Member</label>
          <select className="filter-input" value={teamMember} onChange={e => { setTeamMember(e.target.value); setPage(1); }}>
            {teamMembers.map(m => <option key={m}>{m}</option>)}
          </select>
        </div>
        <div className="top-filter-group">
          <label className="filter-label">Target Hours</label>
          <input type="number" className="filter-input filter-input-sm" value={targetHours} min={1} max={240}
            onChange={e => { setTargetHours(Number(e.target.value)); setPage(1); }} />
        </div>
        <div className="top-filter-group">
          <label className="filter-label">Include Emails From</label>
          <select className="filter-input" value={emailSource} onChange={e => { setEmailSource(e.target.value); setPage(1); }}>
            <option value="clients">Clients Only</option>
            <option value="all">All</option>
          </select>
        </div>
        <div className="top-filter-group">
          <label className="filter-label">Auto-Replies</label>
          <select className="filter-input" value={excludeAutoReply ? "exclude" : "include"}
            onChange={e => { setExcludeAutoReply(e.target.value === "exclude"); setPage(1); }}>
            <option value="exclude">Exclude</option>
            <option value="include">Include</option>
          </select>
        </div>
        <div className="top-filter-group top-filter-note">
          <span className="filter-note">Excludes weekend days</span>
        </div>
      </div>

      <div className="main">
        {error && <div className="error-banner">Failed to load data: {error}</div>}

        <div className="kpi-row">
          <KpiCard label="Total Emails"      value={(loading && inboundCount === null) ? "—" : totalEmails.toLocaleString()} />
          <KpiCard label="Total Responses"   value={loading ? "—" : totalResponses.toLocaleString()} />
          <KpiCard label="Within Target"     value={loading ? "—" : withinTarget.toLocaleString()} />
          <KpiCard label="% Within Target"   value={loading ? "—" : `${pctWithin}%`} />
          <KpiCard label="Avg Response Time" value={loading ? loadingMsg : fmtDays(avgDays)} sub={loading ? "" : "business hours"} />
        </div>

        <div className="chart-card">
          <h2 className="chart-title">Response Time Distribution</h2>
          <p className="chart-sub">Count of email responses by business day bucket</p>
          {loading ? (
            <div className="chart-loading">{loadingMsg}</div>
          ) : totalResponses === 0 ? (
            <div className="chart-loading">No responses found for selected filters.</div>
          ) : (
            <ResponsiveContainer width="100%" height={300}>
              <BarChart data={chartData} margin={{ top: 10, right: 20, left: 0, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#e5e7eb" vertical={false} />
                <XAxis dataKey="label" tick={{ fill: "#6b7280", fontSize: 13, fontFamily: "inherit" }} axisLine={false} tickLine={false} />
                <YAxis tick={{ fill: "#6b7280", fontSize: 12, fontFamily: "inherit" }} axisLine={false} tickLine={false} width={40} />
                <Tooltip content={<CustomTooltip />} cursor={{ fill: "#f3f4f6" }} />
                <Bar dataKey="count" radius={[4, 4, 0, 0]} maxBarSize={100}>
                  {chartData.map(entry => (
                    <Cell key={entry.label} fill={barColor(entry.label)} />
                  ))}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          )}
        </div>

        {!loading && filtered.length > 0 && (
          <div className="table-card">
            <div className="table-header">
              <div className="table-header-left">
                <h2 className="chart-title">Recent Responses</h2>
                {hasColFilters && (
                  <button className="clear-filters-btn" onClick={() => { setColClient(""); setColSubject(""); setColTarget("all"); }}>
                    ✕ Clear column filters
                  </button>
                )}
              </div>
              <button className="export-btn" onClick={() => exportCSV(tableRows)}>
                ↓ Export CSV
              </button>
            </div>
            <div className="table-wrap">
              <table className="resp-table">
                <thead>
                  <tr className="th-labels">
                    <th onClick={() => handleSort("inbound_at")} className="th-sortable">
                      Date <SortIcon active={sortCol === "inbound_at"} dir={sortDir} />
                    </th>
                    <th onClick={() => handleSort("member")} className="th-sortable">
                      Team Member <SortIcon active={sortCol === "member"} dir={sortDir} />
                    </th>
                    <th onClick={() => handleSort("client")} className="th-sortable">
                      Client <SortIcon active={sortCol === "client"} dir={sortDir} />
                    </th>
                    <th onClick={() => handleSort("subject")} className="th-sortable">
                      Subject <SortIcon active={sortCol === "subject"} dir={sortDir} />
                    </th>
                    <th onClick={() => handleSort("response")} className="th-sortable">
                      Response Time <SortIcon active={sortCol === "response"} dir={sortDir} />
                    </th>
                    <th onClick={() => handleSort("target")} className="th-sortable">
                      Within Target <SortIcon active={sortCol === "target"} dir={sortDir} />
                    </th>
                  </tr>
                  <tr className="th-filters">
                    <td></td>
                    <td></td>
                    <td>
                      <input className="col-filter-input" placeholder="Filter client…"
                        value={colClient} onChange={e => { setColClient(e.target.value); setPage(1); }} />
                    </td>
                    <td>
                      <input className="col-filter-input" placeholder="Filter subject…"
                        value={colSubject} onChange={e => { setColSubject(e.target.value); setPage(1); }} />
                    </td>
                    <td></td>
                    <td>
                      <select className="col-filter-input" value={colTarget}
                        onChange={e => { setColTarget(e.target.value); setPage(1); }}>
                        <option value="all">All</option>
                        <option value="yes">✓ Within</option>
                        <option value="no">✗ Over</option>
                      </select>
                    </td>
                  </tr>
                </thead>
                <tbody>
                  {paginated.map(r => (
                    <tr key={r.id}>
                      <td>{new Date(r.inbound_at).toLocaleDateString()}</td>
                      <td>{r.team_member_name || r.team_member_email}</td>
                      <td className="td-muted">{r.client_name || r.client_email || "—"}</td>
                      <td className="td-subject">{r.subject || "—"}</td>
                      <td className="td-mono">{fmtDays(r.response_days)}</td>
                      <td>
                        <span className={r.within_target ? "badge-green" : "badge-red"}>
                          {r.within_target ? "✓" : "✗"}
                        </span>
                      </td>
                    </tr>
                  ))}
                  {paginated.length === 0 && (
                    <tr><td colSpan={6} className="td-empty">No results match your filters.</td></tr>
                  )}
                </tbody>
              </table>
            </div>

            <div className="pagination">
              <span className="pagination-info">
                Showing {tableRows.length === 0 ? 0 : ((page - 1) * PAGE_SIZE) + 1}–{Math.min(page * PAGE_SIZE, tableRows.length)} of {tableRows.length.toLocaleString()} responses
              </span>
              <div className="pagination-controls">
                <button className="page-btn" onClick={() => setPage(1)} disabled={page === 1}>«</button>
                <button className="page-btn" onClick={() => setPage(p => p - 1)} disabled={page === 1}>‹ Prev</button>
                <span className="page-current">Page {page} of {totalPages}</span>
                <button className="page-btn" onClick={() => setPage(p => p + 1)} disabled={page === totalPages}>Next ›</button>
                <button className="page-btn" onClick={() => setPage(totalPages)} disabled={page === totalPages}>»</button>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
