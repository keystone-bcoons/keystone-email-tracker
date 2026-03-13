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
          .order("last_synced_at", { ascending: fals
