import { useState, useEffect, useMemo } from "react";
import { createClient } from "@supabase/supabase-js";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Cell,
} from "recharts";

const supabase = createClient(
  import.meta.env.VITE_SUPABASE_URL,
  import.meta.env.VITE_SUPABASE_ANON_KEY
);

const TARGET_HOURS = 48;

function fmtDays(d) {
  if (d < 1) return `${Math.round(d * 24)}h`;
  return `${d.toFixed(1)}d`;
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

export default function App() {
  const [responses, setResponses]     = useState([]);
  const [loading, setLoading]         = useState(true);
  const [error, setError]             = useState(null);
  const [lastSynced, setLastSynced]   = useState(null);

  const [dateFrom, setDateFrom] = useState(() => {
    const d = new Date(); d.setMonth(d.getMonth() - 3);
    return d.toISOString().split("T")[0];
  });
  const [dateTo, setDateTo]           = useState(() => new Date().toISOString().split("T")[0]);
  const [teamMember, setTeamMember]   = useState("All");
  const [targetHours, setTargetHours] = useState(TARGET_HOURS);

  useEffect(() => {
    async function load() {
      setLoading(true);
      setError(null);
      try {
        const { data, error: err } = await supabase
          .from("email_responses")
          .select("*")
          .gte("inbound_at", dateFrom)
          .lte("inbound_at", dateTo + "T23:59:59Z")
          .order("inbound_at", { ascending: false });

        if (err) throw err;
        setResponses(data || []);

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

  const teamMembers = useMemo(() => {
    const names = [...new Set(responses.map(r => r.team_member_name || r.team_member_email))];
    return ["All", ...names.sort()];
  }, [responses]);

  const filtered = useMemo(() => {
    return responses
      .filter(r => teamMember === "All" || r.team_member_name === teamMember || r.team_member_email === teamMember)
      .map(r => ({ ...r, within_target: r.response_days <= targetHours / 24 }));
  }, [responses, teamMember, targetHours]);

  const totalResponses = filtered.length;
  const withinTarget   = filtered.filter(r => r.within_target).length;
  const pctWithin      = totalResponses ? (withinTarget / totalResponses * 100).toFixed(1) : "0.0";
  const avgDays        = totalResponses
    ? filtered.reduce((s, r) => s + r.response_days, 0) / totalResponses : 0;

  const totalEmails = useMemo(() =>
    new Set(filtered.map(r => r.inbound_message_id)).size, [filtered]);

  const chartData = useMemo(() => {
    const counts = Object.fromEntries(BUCKET_ORDER.map(b => [b, 0]));
    filtered.forEach(r => { counts[bucketLabel(r.response_days)]++; });
    return BUCKET_ORDER.map(b => ({ label: b, count: counts[b] }));
  }, [filtered]);

  const barColor = (label) =>
    label === "< 1 Day" || label === "1–2 Days" ? "#2563eb" : "#93c5fd";

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

      <div className="layout">
        <aside className="sidebar">
          <div className="filter-section">
            <label className="filter-label">Date Range</label>
            <input type="date" className="filter-input" value={dateFrom} onChange={e => setDateFrom(e.target.value)} />
            <input type="date" className="filter-input" value={dateTo} onChange={e => setDateTo(e.target.value)} />
            <p className="filter-note">Excludes weekend days</p>
          </div>
          <div className="filter-section">
            <label className="filter-label">Team Member</label>
            <select className="filter-input" value={teamMember} onChange={e => setTeamMember(e.target.value)}>
              {teamMembers.map(m => <option key={m}>{m}</option>)}
            </select>
          </div>
          <div className="filter-section">
            <label className="filter-label">Target Hours</label>
            <input
              type="number"
              className="filter-input"
              value={targetHours}
              min={1} max={240}
              onChange={e => setTargetHours(Number(e.target.value))}
            />
          </div>
        </aside>

        <main className="main">
          {error && <div className="error-banner">Failed to load data: {error}</div>}

          <div className="kpi-row">
            <KpiCard label="Total Emails"              value={loading ? "—" : totalEmails.toLocaleString()} />
            <KpiCard label="Total Responses"           value={loading ? "—" : totalResponses.toLocaleString()} />
            <KpiCard label="Responses Within Target"   value={loading ? "—" : withinTarget.toLocaleString()} />
            <KpiCard label="% Within Target"           value={loading ? "—" : `${pctWithin}%`} />
            <KpiCard label="Avg Response Time"         value={loading ? "—" : fmtDays(avgDays)} sub="business days" />
          </div>

          <div className="chart-card">
            <h2 className="chart-title">Response Time Distribution</h2>
            <p className="chart-sub">Count of email responses by business day bucket</p>
            {loading ? (
              <div className="chart-loading">Loading…</div>
            ) : totalResponses === 0 ? (
              <div className="chart-loading">No responses found for selected filters.</div>
            ) : (
              <ResponsiveContainer width="100%" height={320}>
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
              <h2 className="chart-title">Recent Responses</h2>
              <div className="table-wrap">
                <table className="resp-table">
                  <thead>
                    <tr>
                      <th>Date</th>
                      <th>Team Member</th>
                      <th>Client</th>
                      <th>Subject</th>
                      <th>Response Time</th>
                      <th>Within Target</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filtered.slice(0, 100).map(r => (
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
                  </tbody>
                </table>
                {filtered.length > 100 && (
                  <p className="table-note">Showing 100 of {filtered.length.toLocaleString()} responses</p>
                )}
              </div>
            </div>
          )}
        </main>
      </div>
    </div>
  );
}
