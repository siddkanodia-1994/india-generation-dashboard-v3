import React, { useEffect, useMemo, useRef, useState } from "react";
import {
  Bar,
  BarChart,
  CartesianGrid,
  Legend,
  Line,
  LineChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from "recharts";

/**
 * India Electricity Generation Tracker
 *
 * IMPORTANT BEHAVIOR (as requested):
 * - ALWAYS loads default data from GitHub-served CSV: /data/generation.csv
 * - DOES NOT use localStorage
 * - If you manually add/import during a session, it updates charts immediately
 * - On refresh, it reloads from /data/generation.csv again
 *
 * CSV format (header optional):
 *   date,generation_gwh
 *   Date format: DD-MM-YYYY
 */

function clamp(n: number, min: number, max: number) {
  return Math.min(max, Math.max(min, n));
}

// ----------------------
// Date helpers
// ----------------------

function parseISOKey(s: string) {
  const ok = /^\d{4}-\d{2}-\d{2}$/.test(s);
  if (!ok) return null;
  const d = new Date(s + "T00:00:00Z");
  return Number.isNaN(d.getTime()) ? null : s;
}

// Parse DD-MM-YYYY -> ISO YYYY-MM-DD (also accept ISO for backward compatibility)
function parseInputDate(s: unknown) {
  if (typeof s !== "string") return null;
  const t = s.trim();

  if (/^\d{2}-\d{2}-\d{4}$/.test(t)) {
    const [dd, mm, yyyy] = t.split("-").map(Number);
    const d = new Date(Date.UTC(yyyy, mm - 1, dd));
    if (Number.isNaN(d.getTime())) return null;
    if (d.getUTCFullYear() !== yyyy || d.getUTCMonth() !== mm - 1 || d.getUTCDate() !== dd) return null;
    return `${yyyy}-${String(mm).padStart(2, "0")}-${String(dd).padStart(2, "0")}`;
  }

  if (/^\d{4}-\d{2}-\d{2}$/.test(t)) return parseISOKey(t);

  return null;
}

function formatDDMMYYYY(iso: string) {
  if (!iso || typeof iso !== "string" || !/^\d{4}-\d{2}-\d{2}$/.test(iso)) return "—";
  const [y, m, d] = iso.split("-");
  return `${d}-${m}-${y}`;
}

function isoMinusDays(iso: string, days: number) {
  const d = new Date(iso + "T00:00:00Z");
  d.setUTCDate(d.getUTCDate() - days);
  return d.toISOString().slice(0, 10);
}

function isoPlusDays(iso: string, days: number) {
  const d = new Date(iso + "T00:00:00Z");
  d.setUTCDate(d.getUTCDate() + days);
  return d.toISOString().slice(0, 10);
}

// Week starts on Monday
function startOfWeekISO(iso: string) {
  const d = new Date(iso + "T00:00:00Z");
  const dow = d.getUTCDay(); // 0=Sun,1=Mon,...
  const diffToMon = dow === 0 ? -6 : 1 - dow;
  d.setUTCDate(d.getUTCDate() + diffToMon);
  return d.toISOString().slice(0, 10);
}

// ----------------------
// Number formatting
// ----------------------

function fmtNum(x: number | null | undefined, digits = 2) {
  if (x == null || Number.isNaN(x)) return "—";
  return new Intl.NumberFormat("en-IN", {
    minimumFractionDigits: digits,
    maximumFractionDigits: digits,
  }).format(x);
}

function fmtPct(x: number | null | undefined, digits = 2) {
  if (x == null || Number.isNaN(x)) return "—";
  const sign = x > 0 ? "+" : "";
  return `${sign}${fmtNum(x, digits)}%`;
}

function asFiniteNumber(v: unknown): number | null {
  if (typeof v === "number" && Number.isFinite(v)) return v;
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

// ----------------------
// Aggregation + growth
// ----------------------

function monthKey(isoDate: string) {
  return isoDate.slice(0, 7); // YYYY-MM
}

function addMonths(ym: string, delta: number) {
  const [y, m] = ym.split("-").map(Number);
  const d = new Date(Date.UTC(y, m - 1 + delta, 1));
  return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, "0")}`;
}

function getYear(ym: string) {
  return Number(ym.slice(0, 4));
}

function getMonth(ym: string) {
  return Number(ym.slice(5, 7));
}

function safeDiv(n: number, d: number | null | undefined) {
  if (d == null || d === 0) return null;
  return n / d;
}

function growthPct(curr: number, prev: number) {
  const r = safeDiv(curr - prev, prev);
  return r == null ? null : r * 100;
}

function sortISO(a: string, b: string) {
  return a < b ? -1 : a > b ? 1 : 0;
}

function mergeRecords(
  existingMap: Map<string, number>,
  incoming: Array<{ date: string; generation_gwh: number }>
) {
  const next = new Map(existingMap);
  for (const r of incoming) next.set(r.date, r.generation_gwh);
  return next;
}

function downloadCSV(filename: string, text: string) {
  const blob = new Blob([text], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

// ----------------------
// Data types
// ----------------------

type DailyPoint = { date: string; generation_gwh: number };

type DailyChartPoint = {
  label: string;
  units: number;
  prev_year_units: number | null;
  yoy_pct: number | null;
  mom_pct: number | null;
};

// ----------------------
// KPI + monthly computation
// ----------------------

function computeKPIs(sortedDaily: DailyPoint[]) {
  if (sortedDaily.length === 0) {
    return {
      latest: null as DailyPoint | null,
      latestYoY: null as number | null,
      avg7: null as number | null,
      avg7YoY: null as number | null,
      avg30: null as number | null,
      avg30YoY: null as number | null,
      ytdTotal: null as number | null,
      ytdYoY: null as number | null,
      mtdAvg: null as number | null,
      mtdYoY: null as number | null,
    };
  }

  const dailyLookup = new Map(sortedDaily.map((d) => [d.date, d.generation_gwh] as const));
  const latest = sortedDaily[sortedDaily.length - 1];

  const isoAddYears = (iso: string, deltaYears: number) => {
    const y = Number(iso.slice(0, 4));
    const m = Number(iso.slice(5, 7));
    const d = Number(iso.slice(8, 10));

    const tryDt = new Date(Date.UTC(y + deltaYears, m - 1, d));
    if (tryDt.getUTCFullYear() === y + deltaYears && tryDt.getUTCMonth() === m - 1 && tryDt.getUTCDate() === d) {
      return tryDt.toISOString().slice(0, 10);
    }
    const lastDay = new Date(Date.UTC(y + deltaYears, m, 0));
    return lastDay.toISOString().slice(0, 10);
  };

  const sumAndCountInclusive = (startIso: string, endIso: string) => {
    if (startIso > endIso) return { sum: null as number | null, count: 0 };

    let sum = 0;
    let count = 0;
    let cur = startIso;

    while (cur <= endIso) {
      const v = dailyLookup.get(cur);
      if (v != null) {
        sum += v;
        count += 1;
      }
      cur = isoPlusDays(cur, 1);
    }

    return { sum: count ? sum : null, count };
  };

  const avgForLastNDaysEnding = (endIso: string, nDays: number) => {
    const startIso = isoMinusDays(endIso, nDays - 1);
    const { sum, count } = sumAndCountInclusive(startIso, endIso);
    return { startIso, endIso, avg: sum != null && count ? sum / count : null };
  };

  // Latest day YoY
  const prevYearDate = isoAddYears(latest.date, -1);
  const prevYearVal = dailyLookup.get(prevYearDate) ?? null;
  const latestYoY = prevYearVal != null ? growthPct(latest.generation_gwh, prevYearVal) : null;

  // 7D avg + YoY
  const last7 = avgForLastNDaysEnding(latest.date, 7);
  const py7 = sumAndCountInclusive(isoAddYears(last7.startIso, -1), isoAddYears(last7.endIso, -1));
  const avg7 = last7.avg;
  const avg7PY = py7.sum != null && py7.count ? py7.sum / py7.count : null;
  const avg7YoY = avg7 != null && avg7PY != null ? growthPct(avg7, avg7PY) : null;

  // 30D avg + YoY
  const last30 = avgForLastNDaysEnding(latest.date, 30);
  const py30 = sumAndCountInclusive(isoAddYears(last30.startIso, -1), isoAddYears(last30.endIso, -1));
  const avg30 = last30.avg;
  const avg30PY = py30.sum != null && py30.count ? py30.sum / py30.count : null;
  const avg30YoY = avg30 != null && avg30PY != null ? growthPct(avg30, avg30PY) : null;

  // FY-YTD (Apr 1 to latest)
  const latestY = Number(latest.date.slice(0, 4));
  const latestM = Number(latest.date.slice(5, 7));
  const fyStartYear = latestM >= 4 ? latestY : latestY - 1;
  const ytdStart = `${fyStartYear}-04-01`;

  const ytd = sumAndCountInclusive(ytdStart, latest.date);
  const ytdTotal = ytd.sum;

  const ytdPYStart = `${fyStartYear - 1}-04-01`;
  const ytdPYEnd = isoAddYears(latest.date, -1);
  const ytdPY = sumAndCountInclusive(ytdPYStart, ytdPYEnd);
  const ytdTotalPY = ytdPY.sum;
  const ytdYoY = ytdTotal != null && ytdTotalPY != null ? growthPct(ytdTotal, ytdTotalPY) : null;

  // MTD avg + YoY
  const thisMonthStart = `${latest.date.slice(0, 7)}-01`;
  const mtd = sumAndCountInclusive(thisMonthStart, latest.date);
  const mtdAvg = mtd.sum != null && mtd.count ? mtd.sum / mtd.count : null;

  const mtdPY = sumAndCountInclusive(isoAddYears(thisMonthStart, -1), isoAddYears(latest.date, -1));
  const mtdAvgPY = mtdPY.sum != null && mtdPY.count ? mtdPY.sum / mtdPY.count : null;
  const mtdYoY = mtdAvg != null && mtdAvgPY != null ? growthPct(mtdAvg, mtdAvgPY) : null;

  return { latest, latestYoY, avg7, avg7YoY, avg30, avg30YoY, ytdTotal, ytdYoY, mtdAvg, mtdYoY };
}

function buildMonthDayMap(sortedDaily: DailyPoint[]) {
  const map = new Map<string, { total: number; maxDay: number; byDay: Map<number, number> }>();
  for (const d of sortedDaily) {
    const m = monthKey(d.date);
    const day = Number(d.date.slice(8, 10));
    if (!map.has(m)) map.set(m, { total: 0, maxDay: 0, byDay: new Map() });
    const rec = map.get(m)!;
    rec.total += d.generation_gwh;
    rec.maxDay = Math.max(rec.maxDay, day);
    rec.byDay.set(day, (rec.byDay.get(day) || 0) + d.generation_gwh);
  }
  return map;
}

function sumMonthUpToDay(monthRec: { byDay: Map<number, number> } | undefined, dayLimit: number) {
  if (!monthRec) return null;
  let s = 0;
  let hasAny = false;
  for (let day = 1; day <= dayLimit; day++) {
    const v = monthRec.byDay.get(day);
    if (v != null) {
      s += v;
      hasAny = true;
    }
  }
  return hasAny ? s : null;
}

function toMonthly(sortedDaily: DailyPoint[]) {
  const monthMap = buildMonthDayMap(sortedDaily);
  const months = Array.from(monthMap.keys()).sort(sortISO);

  const out = months.map((m) => ({
    month: m,
    total_gwh: monthMap.get(m)!.total,
    max_day: monthMap.get(m)!.maxDay,
    yoy_pct: null as number | null,
    mom_pct: null as number | null,
  }));

  for (const r of out) {
    const prevMonth = addMonths(r.month, -1);
    const prevMonthRec = monthMap.get(prevMonth);
    const prevComparableMoM = sumMonthUpToDay(prevMonthRec, r.max_day);
    r.mom_pct = prevComparableMoM != null ? growthPct(r.total_gwh, prevComparableMoM) : null;

    const prevYearMonth = `${getYear(r.month) - 1}-${String(getMonth(r.month)).padStart(2, "0")}`;
    const prevYearRec = monthMap.get(prevYearMonth);
    const prevComparableYoY = sumMonthUpToDay(prevYearRec, r.max_day);
    r.yoy_pct = prevComparableYoY != null ? growthPct(r.total_gwh, prevComparableYoY) : null;
  }

  return out;
}

// ----------------------
// CSV helpers
// ----------------------

function csvParse(text: string) {
  const lines = text
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean);

  const rows: string[][] = [];
  for (const line of lines) {
    const cols = line.split(",").map((c) => c.trim());
    if (cols.length >= 2) rows.push(cols);
  }

  // Drop header if present
  if (rows.length) {
    const h0 = rows[0][0].toLowerCase();
    const h1 = rows[0][1].toLowerCase();
    if (h0.includes("date") && (h1.includes("gen") || h1.includes("gwh") || h1.includes("unit"))) rows.shift();
  }

  const parsed: Array<{ date: string; generation_gwh: number }> = [];
  const errors: string[] = [];

  for (let i = 0; i < rows.length; i++) {
    const [dRaw, gRaw] = rows[i];
    const date = parseInputDate(dRaw);
    const g = Number(String(gRaw).replace(/,/g, ""));

    if (!date) {
      errors.push(`Row ${i + 1}: invalid date '${dRaw}' (expected DD-MM-YYYY)`);
      continue;
    }
    if (!Number.isFinite(g) || g < 0) {
      errors.push(`Row ${i + 1}: invalid generation '${gRaw}' (expected non-negative number)`);
      continue;
    }
    parsed.push({ date, generation_gwh: g });
  }

  return { parsed, errors };
}

function sampleCSV() {
  return [
    "date,generation_gwh",
    "18-12-2025,4140",
    "19-12-2025,4215",
    "20-12-2025,4198",
  ].join("\n");
}

// ----------------------
// UI components
// ----------------------

function Card({
  title,
  children,
  right,
}: {
  title: string;
  children: React.ReactNode;
  right?: React.ReactNode;
}) {
  return (
    <div className="rounded-2xl bg-white shadow-sm ring-1 ring-slate-200">
      <div className="flex items-start justify-between gap-3 border-b border-slate-100 p-4">
        <div className="text-sm font-semibold text-slate-800">{title}</div>
        {right ? <div className="text-sm text-slate-600">{right}</div> : null}
      </div>
      <div className="p-4">{children}</div>
    </div>
  );
}

function Stat({
  label,
  value,
  sub,
  accent,
}: {
  label: string;
  value: string;
  sub?: string | null;
  accent?: "ytd" | null;
}) {
  const valueClass =
    accent === "ytd"
      ? "mt-1 text-2xl font-semibold text-rose-700"
      : "mt-1 text-2xl font-semibold text-slate-900";

  return (
    <div className="rounded-2xl bg-slate-50 p-4 ring-1 ring-slate-200">
      <div className="text-xs font-medium text-slate-500">{label}</div>
      <div className={valueClass}>{value}</div>
      {sub ? <div className="mt-1 text-xs text-slate-500">{sub}</div> : null}
    </div>
  );
}

function EmptyState({ onLoadSample }: { onLoadSample: () => void }) {
  return (
    <div className="rounded-2xl border border-dashed border-slate-300 bg-white p-8 text-center">
      <div className="mx-auto max-w-xl">
        <div className="text-lg font-semibold text-slate-900">No data yet</div>
        <div className="mt-2 text-sm text-slate-600">
          Upload the default CSV to <span className="font-mono">public/data/generation.csv</span> or import a CSV.
        </div>
        <div className="mt-5 flex flex-wrap justify-center gap-3">
          <button
            onClick={onLoadSample}
            className="rounded-xl bg-slate-900 px-4 py-2 text-sm font-semibold text-white hover:bg-slate-800"
          >
            Load sample data
          </button>
        </div>
      </div>
    </div>
  );
}

// ----------------------
// App
// ----------------------

export default function App() {
  // DATA STATE (in-memory; always reset by GitHub CSV on refresh)
  const [dataMap, setDataMap] = useState<Map<string, number>>(new Map());

  const [date, setDate] = useState(() => {
    const t = new Date();
    const dd = String(t.getDate()).padStart(2, "0");
    const mm = String(t.getMonth() + 1).padStart(2, "0");
    const yyyy = t.getFullYear();
    return `${dd}-${mm}-${yyyy}`;
  });

  const [gwh, setGwh] = useState("");
  const [msg, setMsg] = useState<string | null>(null);
  const [errors, setErrors] = useState<string[]>([]);
  const [rangeDays, setRangeDays] = useState(120);

  // Status for GitHub CSV load
  const [defaultLoadStatus, setDefaultLoadStatus] = useState<string>("Loading /data/generation.csv ...");

  // Slicer controls (ISO for date input)
  const [fromIso, setFromIso] = useState("");
  const [toIso, setToIso] = useState("");
  const [aggFreq, setAggFreq] = useState<"daily" | "weekly" | "monthly" | "rolling30">("daily");

  // Series toggles
  const [showUnitsSeries, setShowUnitsSeries] = useState(true);
  const [showPrevYearSeries, setShowPrevYearSeries] = useState(true);
  const [showYoYSeries, setShowYoYSeries] = useState(true);
  const [showMoMSeries, setShowMoMSeries] = useState(true);

  const [showControlLines, setShowControlLines] = useState(false);

  const fileRef = useRef<HTMLInputElement | null>(null);

  // ✅ ALWAYS LOAD FROM GITHUB CSV ON PAGE LOAD
  useEffect(() => {
    let cancelled = false;

    async function loadDefaultCSV() {
      try {
        setDefaultLoadStatus("Loading /data/generation.csv ...");
        const res = await fetch("/data/generation.csv", { cache: "no-store" });
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        const text = await res.text();

        const { parsed, errors: errs } = csvParse(text);
        if (cancelled) return;

        if (!parsed.length) {
          setDataMap(new Map());
          setErrors(errs.length ? errs.slice(0, 12) : ["No valid rows found in /data/generation.csv"]);
          setDefaultLoadStatus("Loaded, but no valid rows were found in /data/generation.csv");
          return;
        }

        const m = new Map<string, number>();
        for (const r of parsed) m.set(r.date, r.generation_gwh);

        setDataMap(m);
        setErrors(errs.slice(0, 12));
        setDefaultLoadStatus(`Loaded ${parsed.length} rows from /data/generation.csv`);
      } catch (e) {
        if (cancelled) return;
        setDataMap(new Map());
        setErrors(["Could not load /data/generation.csv. Ensure it exists at public/data/generation.csv in GitHub."]);
        setDefaultLoadStatus("Failed to load /data/generation.csv");
      }
    }

    loadDefaultCSV();
    return () => {
      cancelled = true;
    };
  }, []);

  const sortedDaily = useMemo<DailyPoint[]>(() => {
    return Array.from(dataMap.entries())
      .map(([d, g]) => ({ date: d, generation_gwh: g }))
      .sort((a, b) => sortISO(a.date, b.date));
  }, [dataMap]);

  // Initialize slicer range once data is present
  useEffect(() => {
    if (!sortedDaily.length) return;
    const lastIso = sortedDaily[sortedDaily.length - 1].date;
    if (!toIso) setToIso(lastIso);
    if (!fromIso) setFromIso(isoMinusDays(lastIso, clamp(rangeDays, 7, 3650)));
  }, [sortedDaily, toIso, fromIso, rangeDays]);

  const anyTotalsShown = showUnitsSeries || showPrevYearSeries;
  const anyPctShown = showYoYSeries || showMoMSeries;

  const dailyForChart = useMemo<DailyChartPoint[]>(() => {
    if (!sortedDaily.length) return [];

    const lastIso = sortedDaily[sortedDaily.length - 1].date;
    const effectiveTo = toIso || lastIso;
    const effectiveFrom = fromIso || isoMinusDays(lastIso, clamp(rangeDays, 7, 3650));

    const f = effectiveFrom <= effectiveTo ? effectiveFrom : effectiveTo;
    const t = effectiveFrom <= effectiveTo ? effectiveTo : effectiveFrom;

    const filtered = sortedDaily.filter((d) => d.date >= f && d.date <= t);
    const dailyLookup = new Map(sortedDaily.map((d) => [d.date, d.generation_gwh]));

    const sumRangeInclusive = (startIso: string, endIso: string) => {
      if (startIso > endIso) return null;
      let s = 0;
      let hasAny = false;
      let cur = startIso;
      while (cur <= endIso) {
        const v = dailyLookup.get(cur);
        if (v != null) {
          s += v;
          hasAny = true;
        }
        cur = isoPlusDays(cur, 1);
      }
      return hasAny ? s : null;
    };

    if (aggFreq === "daily") {
      const sameDayPrevYear = (iso: string) => `${Number(iso.slice(0, 4)) - 1}${iso.slice(4)}`;

      const sameDayPrevMonth = (iso: string) => {
        const y = Number(iso.slice(0, 4));
        const m = Number(iso.slice(5, 7));
        const d = Number(iso.slice(8, 10));
        const dt = new Date(Date.UTC(y, m - 2, d));
        const iso2 = dt.toISOString().slice(0, 10);
        return Number(iso2.slice(8, 10)) === d ? iso2 : null;
      };

      return filtered.map((d) => {
        const pyDate = sameDayPrevYear(d.date);
        const pmDate = sameDayPrevMonth(d.date);
        const py = dailyLookup.get(pyDate) ?? null;
        const pm = pmDate ? dailyLookup.get(pmDate) ?? null : null;

        return {
          label: formatDDMMYYYY(d.date),
          units: Number(d.generation_gwh.toFixed(2)),
          prev_year_units: py != null ? Number(py.toFixed(2)) : null,
          yoy_pct: py != null ? Number(growthPct(d.generation_gwh, py)!.toFixed(2)) : null,
          mom_pct: pm != null ? Number(growthPct(d.generation_gwh, pm)!.toFixed(2)) : null,
        };
      });
    }

    if (aggFreq === "rolling30") {
      const points: DailyChartPoint[] = [];
      let cur = f;
      while (cur <= t) {
        const start = isoMinusDays(cur, 29);
        const currSum = sumRangeInclusive(start, cur);

        // Compare to the equivalent 30D window ending one year earlier (approx by 365 days).
        const curPrevYear = isoMinusDays(cur, 365);
        const startPrevYear = isoMinusDays(curPrevYear, 29);
        const prevSum = sumRangeInclusive(startPrevYear, curPrevYear);

        points.push({
          label: formatDDMMYYYY(cur),
          units: currSum != null ? Number(currSum.toFixed(2)) : 0,
          prev_year_units: prevSum != null ? Number(prevSum.toFixed(2)) : null,
          yoy_pct: currSum != null && prevSum != null ? Number(growthPct(currSum, prevSum)!.toFixed(2)) : null,
          mom_pct: null,
        });

        cur = isoPlusDays(cur, 1);
      }
      return points;
    }

    if (aggFreq === "weekly") {
      const weekMap = new Map<string, number>();
      const weekOffsets = new Map<string, Set<number>>();

      for (const d of filtered) {
        const wk = startOfWeekISO(d.date);
        weekMap.set(wk, (weekMap.get(wk) || 0) + d.generation_gwh);

        const off = Math.floor(
          (new Date(d.date + "T00:00:00Z").getTime() - new Date(wk + "T00:00:00Z").getTime()) / 86400000
        );
        if (!weekOffsets.has(wk)) weekOffsets.set(wk, new Set());
        weekOffsets.get(wk)!.add(off);
      }

      const weeks = Array.from(weekMap.keys()).sort(sortISO);

      const sumWeekByOffsets = (weekStartIso: string, offsetsSet: Set<number>) => {
        let s = 0;
        let hasAny = false;
        for (const off of offsetsSet) {
          const key = isoPlusDays(weekStartIso, off);
          const v = dailyLookup.get(key);
          if (v != null) {
            s += v;
            hasAny = true;
          }
        }
        return hasAny ? s : null;
      };

      return weeks.map((wk) => {
        const curr = weekMap.get(wk)!;
        const offs = weekOffsets.get(wk) || new Set<number>();

        const prevWkYoY = isoMinusDays(wk, 364);
        const prevYoY = sumWeekByOffsets(prevWkYoY, offs);

        const prevWkWoW = isoMinusDays(wk, 7);
        const prevWoW = sumWeekByOffsets(prevWkWoW, offs);

        return {
          label: `Wk of ${formatDDMMYYYY(wk)}`,
          units: Number(curr.toFixed(2)),
          prev_year_units: prevYoY != null ? Number(prevYoY.toFixed(2)) : null,
          yoy_pct: prevYoY != null ? Number(growthPct(curr, prevYoY)!.toFixed(2)) : null,
          mom_pct: prevWoW != null ? Number(growthPct(curr, prevWoW)!.toFixed(2)) : null,
        };
      });
    }

    // monthly
    const mMap = new Map<string, number>();
    const monthDays = new Map<string, Set<number>>();

    for (const d of filtered) {
      const mk = monthKey(d.date);
      mMap.set(mk, (mMap.get(mk) || 0) + d.generation_gwh);

      const day = Number(d.date.slice(8, 10));
      if (!monthDays.has(mk)) monthDays.set(mk, new Set());
      monthDays.get(mk)!.add(day);
    }

    const months = Array.from(mMap.keys()).sort(sortISO);

    const sumMonthByDaySet = (ym: string, daySet: Set<number>) => {
      const y = ym.slice(0, 4);
      const m = ym.slice(5, 7);
      let s = 0;
      let hasAny = false;
      for (const day of daySet) {
        const key = `${y}-${m}-${String(day).padStart(2, "0")}`;
        const v = dailyLookup.get(key);
        if (v != null) {
          s += v;
          hasAny = true;
        }
      }
      return hasAny ? s : null;
    };

    return months.map((m) => {
      const curr = mMap.get(m)!;
      const days = monthDays.get(m) || new Set<number>();

      const prevYearMonth = `${getYear(m) - 1}-${String(getMonth(m)).padStart(2, "0")}`;
      const prevYoY = sumMonthByDaySet(prevYearMonth, days);

      const prevMonth = addMonths(m, -1);
      const prevMoM = sumMonthByDaySet(prevMonth, days);

      return {
        label: m,
        units: Number(curr.toFixed(2)),
        prev_year_units: prevYoY != null ? Number(prevYoY.toFixed(2)) : null,
        yoy_pct: prevYoY != null ? Number(growthPct(curr, prevYoY)!.toFixed(2)) : null,
        mom_pct: prevMoM != null ? Number(growthPct(curr, prevMoM)!.toFixed(2)) : null,
      };
    });
  }, [sortedDaily, rangeDays, fromIso, toIso, aggFreq]);

  const monthly = useMemo(() => toMonthly(sortedDaily), [sortedDaily]);

  const monthlyForChart = useMemo(() => {
    if (!monthly.length) return [];
    return monthly.slice(Math.max(0, monthly.length - 24)).map((m) => ({
      month: m.month,
      total_units: Number(m.total_gwh.toFixed(2)),
      yoy_pct: m.yoy_pct != null ? Number(m.yoy_pct.toFixed(2)) : null,
      mom_pct: m.mom_pct != null ? Number(m.mom_pct.toFixed(2)) : null,
    }));
  }, [monthly]);

  const kpis = useMemo(() => computeKPIs(sortedDaily), [sortedDaily]);

  function upsertOne() {
    setMsg(null);
    setErrors([]);

    const iso = parseInputDate(date);
    if (!iso) {
      setErrors(["Please enter a valid date (DD-MM-YYYY)."]);
      return;
    }

    const v = Number(String(gwh).replace(/,/g, ""));
    if (!Number.isFinite(v) || v < 0) {
      setErrors(["Please enter a valid non-negative number for generation."]);
      return;
    }

    setDataMap((prev) => {
      const next = new Map(prev);
      next.set(iso, v);
      return next;
    });

    setMsg(`Saved ${formatDDMMYYYY(iso)}: ${fmtNum(v)} units (session only; refresh reloads GitHub CSV)`);
    setGwh("");
  }

  function removeDate(isoDate: string) {
    setDataMap((prev) => {
      const next = new Map(prev);
      next.delete(isoDate);
      return next;
    });
  }

  async function importCSV(file?: File) {
    setMsg(null);
    setErrors([]);
    if (!file) return;

    try {
      const text = await file.text();
      const { parsed, errors: errs } = csvParse(text);

      if (errs.length) setErrors(errs.slice(0, 12));
      if (!parsed.length) {
        setErrors((e) => (e.length ? e : ["No valid rows found in CSV."]));
        return;
      }

      setDataMap((prev) => mergeRecords(prev, parsed));
      setMsg(`Imported ${parsed.length} rows (session only; refresh reloads GitHub CSV).`);
    } catch {
      setErrors(["Could not read CSV."]);
    } finally {
      if (fileRef.current) fileRef.current.value = "";
    }
  }

  function exportCSV() {
    const header = "date,generation_gwh";
    const lines = sortedDaily.map((d) => `${formatDDMMYYYY(d.date)},${d.generation_gwh}`);
    downloadCSV(
      `india_generation_${new Date().toISOString().slice(0, 10)}.csv`,
      [header, ...lines].join("\n")
    );
  }

  function loadSample() {
    const { parsed } = csvParse(sampleCSV());
    setDataMap((prev) => mergeRecords(prev, parsed));
    setMsg("Loaded sample data (session only; refresh reloads GitHub CSV).");
  }

  const hasData = sortedDaily.length > 0;

  return (
    <div className="min-h-screen bg-slate-50">
      <div className="mx-auto max-w-7xl px-4 py-8">
        <div className="flex flex-col gap-2 sm:flex-row sm:items-end sm:justify-between">
          <div>
            <div className="text-2xl font-semibold text-slate-900">India Electricity Generation Dashboard</div>
            <div className="mt-1 text-sm text-slate-600">
              Default source: <span className="font-mono">/data/generation.csv</span> (GitHub)
            </div>
            <div className="mt-1 text-xs text-slate-500">{defaultLoadStatus}</div>
          </div>

          <div className="flex flex-wrap gap-2">
            <button
              onClick={() => downloadCSV("sample_generation.csv", sampleCSV())}
              className="rounded-xl bg-white px-3 py-2 text-sm font-semibold text-slate-700 ring-1 ring-slate-200 hover:bg-slate-50"
            >
              Download sample CSV
            </button>

            <button
              onClick={exportCSV}
              disabled={!hasData}
              className="rounded-xl bg-white px-3 py-2 text-sm font-semibold text-slate-700 ring-1 ring-slate-200 hover:bg-slate-50 disabled:opacity-50"
            >
              Export CSV
            </button>
          </div>
        </div>

        <div className="mt-6 grid grid-cols-1 gap-4 lg:grid-cols-3">
          <Card title="Add / Update a day">
            <div className="grid grid-cols-1 gap-3">
              <label className="text-xs font-medium text-slate-600">Date (DD-MM-YYYY)</label>
              <input
                type="text"
                placeholder="DD-MM-YYYY"
                value={date}
                onChange={(e) => setDate(e.target.value)}
                className="w-full rounded-xl border border-slate-200 bg-white px-3 py-2 text-sm text-slate-900 outline-none focus:ring-2 focus:ring-slate-300"
              />

              <label className="mt-1 text-xs font-medium text-slate-600">Generation (units / MU)</label>
              <input
                inputMode="decimal"
                placeholder="e.g., 4200"
                value={gwh}
                onChange={(e) => setGwh(e.target.value)}
                className="w-full rounded-xl border border-slate-200 bg-white px-3 py-2 text-sm text-slate-900 outline-none focus:ring-2 focus:ring-slate-300"
              />

              <button
                onClick={upsertOne}
                className="mt-1 rounded-xl bg-slate-900 px-4 py-2 text-sm font-semibold text-white hover:bg-slate-800"
              >
                Save day (session only)
              </button>

              <div className="mt-2">
                <div className="text-xs font-medium text-slate-600">Import CSV (session only)</div>
                <div className="mt-2 flex items-center gap-2">
                  <input
                    ref={fileRef}
                    type="file"
                    accept=".csv,text/csv"
                    onChange={(e) => importCSV(e.target.files?.[0])}
                    className="block w-full text-sm text-slate-700 file:mr-3 file:rounded-xl file:border-0 file:bg-slate-900 file:px-3 file:py-2 file:text-sm file:font-semibold file:text-white hover:file:bg-slate-800"
                  />
                </div>
                <div className="mt-2 text-xs text-slate-500">
                  Supported: <span className="font-mono">date,generation_gwh</span> (DD-MM-YYYY, number)
                </div>
              </div>

              {msg ? (
                <div className="mt-2 rounded-xl bg-emerald-50 p-3 text-sm text-emerald-800 ring-1 ring-emerald-200">
                  {msg}
                </div>
              ) : null}

              {errors.length ? (
                <div className="mt-2 rounded-xl bg-rose-50 p-3 text-sm text-rose-800 ring-1 ring-rose-200">
                  <div className="font-semibold">Import / input issues</div>
                  <ul className="mt-1 list-disc pl-5">
                    {errors.map((e, i) => (
                      <li key={i}>{e}</li>
                    ))}
                  </ul>
                </div>
              ) : null}

              <div className="mt-2 text-xs text-slate-500">
                To make changes permanent: edit <span className="font-mono">public/data/generation.csv</span> in GitHub and
                commit. Vercel will redeploy automatically.
              </div>
            </div>
          </Card>

          <Card title="Quick stats" right={hasData ? `Records: ${sortedDaily.length}` : null}>
            {!hasData ? (
              <EmptyState onLoadSample={loadSample} />
            ) : (
              <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
                <Stat
                  label="Latest day"
                  value={kpis.latest ? formatDDMMYYYY(kpis.latest.date) : "—"}
                  sub={kpis.latest ? `${fmtNum(kpis.latest.generation_gwh)} units` : null}
                />
                <Stat label="Latest YoY (same day)" value={fmtPct(kpis.latestYoY, 2)} sub="vs same date last year (if available)" />
                <Stat
                  label="Current 7-Day Average Units"
                  value={kpis.avg7 != null ? `${fmtNum(kpis.avg7)} units` : "—"}
                  sub={kpis.avg7YoY != null ? `${fmtPct(kpis.avg7YoY, 2)} YoY` : "YoY: —"}
                />
                <Stat
                  label="Current 30-Day Average Units"
                  value={kpis.avg30 != null ? `${fmtNum(kpis.avg30)} units` : "—"}
                  sub={kpis.avg30YoY != null ? `${fmtPct(kpis.avg30YoY, 2)} YoY` : "YoY: —"}
                />
                <Stat
                  label="YTD Total Units (from 1 Apr)"
                  value={kpis.ytdTotal != null ? `${fmtNum(kpis.ytdTotal)} units` : "—"}
                  sub={kpis.ytdYoY != null ? `${fmtPct(kpis.ytdYoY, 2)} YoY` : "YoY: —"}
                  accent="ytd"
                />
                <Stat
                  label="MTD Average Units"
                  value={kpis.mtdAvg != null ? `${fmtNum(kpis.mtdAvg)} units` : "—"}
                  sub={kpis.mtdYoY != null ? `${fmtPct(kpis.mtdYoY, 2)} YoY` : "YoY: —"}
                />
              </div>
            )}
          </Card>

          <Card title="Recent entries">
            {!hasData ? (
              <div className="text-sm text-slate-600">Once you add data, the most recent entries will appear here.</div>
            ) : (
              <div className="max-h-[420px] overflow-auto rounded-xl ring-1 ring-slate-200">
                <table className="w-full border-collapse bg-white text-left text-sm">
                  <thead className="sticky top-0 bg-slate-50">
                    <tr>
                      <th className="px-3 py-2 text-xs font-semibold text-slate-600">Date</th>
                      <th className="px-3 py-2 text-xs font-semibold text-slate-600">Generation (units)</th>
                      <th className="px-3 py-2 text-xs font-semibold text-slate-600"></th>
                    </tr>
                  </thead>
                  <tbody>
                    {sortedDaily
                      .slice(-25)
                      .reverse()
                      .map((r) => (
                        <tr key={r.date} className="border-t border-slate-100">
                          <td className="px-3 py-2 font-medium text-slate-900">{formatDDMMYYYY(r.date)}</td>
                          <td className="px-3 py-2 text-slate-700">{fmtNum(r.generation_gwh)}</td>
                          <td className="px-3 py-2 text-right">
                            <button
                              onClick={() => removeDate(r.date)}
                              className="rounded-lg px-2 py-1 text-xs font-semibold text-rose-700 hover:bg-rose-50"
                            >
                              Remove
                            </button>
                          </td>
                        </tr>
                      ))}
                  </tbody>
                </table>
              </div>
            )}
          </Card>
        </div>

        <div className="mt-6 grid grid-cols-1 gap-4 lg:grid-cols-2">
          <Card
            title="Daily generation"
            right={
              hasData ? (
                <div className="flex items-center gap-2">
                  <span className="text-xs text-slate-500">Range</span>
                  <select
                    value={rangeDays}
                    onChange={(e) => {
                      const v = Number(e.target.value);
                      setRangeDays(v);
                      if (sortedDaily.length) {
                        const lastIso = sortedDaily[sortedDaily.length - 1].date;
                        setToIso(lastIso);
                        setFromIso(isoMinusDays(lastIso, clamp(v, 7, 3650)));
                      }
                    }}
                    className="rounded-xl border border-slate-200 bg-white px-2 py-1 text-sm text-slate-700"
                  >
                    <option value={60}>Last 60 days</option>
                    <option value={120}>Last 120 days</option>
                    <option value={365}>Last 12 months</option>
                    <option value={730}>Last 24 months</option>
                    <option value={1825}>Last 5 years</option>
                    <option value={3650}>Last 10 years</option>
                  </select>
                </div>
              ) : null
            }
          >
            {!hasData ? (
              <div className="text-sm text-slate-600">Add data to see the chart.</div>
            ) : (
              <>
                <div className="mb-3 rounded-2xl bg-slate-50 p-3 ring-1 ring-slate-200">
                  <div className="grid grid-cols-1 gap-3 sm:grid-cols-3 sm:items-end">
                    <div>
                      <div className="text-xs font-medium text-slate-600">From</div>
                      <input
                        type="date"
                        value={fromIso}
                        onChange={(e) => setFromIso(e.target.value)}
                        className="mt-1 w-full rounded-xl border border-slate-200 bg-white px-3 py-2 text-sm text-slate-900 outline-none focus:ring-2 focus:ring-slate-300"
                      />
                      <div className="mt-1 text-[11px] text-slate-500">{fromIso ? formatDDMMYYYY(fromIso) : ""}</div>
                    </div>

                    <div>
                      <div className="text-xs font-medium text-slate-600">To</div>
                      <input
                        type="date"
                        value={toIso}
                        onChange={(e) => setToIso(e.target.value)}
                        className="mt-1 w-full rounded-xl border border-slate-200 bg-white px-3 py-2 text-sm text-slate-900 outline-none focus:ring-2 focus:ring-slate-300"
                      />
                      <div className="mt-1 text-[11px] text-slate-500">{toIso ? formatDDMMYYYY(toIso) : ""}</div>
                    </div>

                    <div>
                      <div className="text-xs font-medium text-slate-600">View as</div>
                      <select
                        value={aggFreq}
                        onChange={(e) => setAggFreq(e.target.value as any)}
                        className="mt-1 w-full rounded-xl border border-slate-200 bg-white px-3 py-2 text-sm text-slate-700"
                      >
                        <option value="daily">Daily</option>
                        <option value="weekly">Weekly (sum)</option>
                        <option value="monthly">Monthly (sum)</option>
                        <option value="rolling30">Last 30 Days Rolling Sum (YoY Demand Growth)</option>
                      </select>

                      <div className="mt-2 rounded-xl bg-white p-3 ring-1 ring-slate-200">
                        <div className="flex items-center justify-between gap-2">
                          <div className="text-xs font-semibold text-slate-700">Series toggles</div>
                          <label className="flex items-center gap-2 text-[12px] text-slate-700">
                            <input
                              type="checkbox"
                              checked={showControlLines}
                              onChange={(e) => setShowControlLines(e.target.checked)}
                              className="h-4 w-4 rounded border-slate-300"
                            />
                            <span className="font-medium">Show Control Lines</span>
                          </label>
                        </div>

                        <div className="mt-2 grid grid-cols-2 gap-2 text-[12px] text-slate-700">
                          <label className="flex items-center gap-2">
                            <input
                              type="checkbox"
                              checked={showUnitsSeries}
                              onChange={(e) => setShowUnitsSeries(e.target.checked)}
                              className="h-4 w-4 rounded border-slate-300"
                            />
                            <span className="font-medium">Total Current</span>
                          </label>

                          <label className="flex items-center gap-2">
                            <input
                              type="checkbox"
                              checked={showPrevYearSeries}
                              onChange={(e) => setShowPrevYearSeries(e.target.checked)}
                              className="h-4 w-4 rounded border-slate-300"
                            />
                            <span className="font-medium">Total (previous year)</span>
                          </label>

                          <label className="flex items-center gap-2">
                            <input
                              type="checkbox"
                              checked={showYoYSeries}
                              onChange={(e) => setShowYoYSeries(e.target.checked)}
                              className="h-4 w-4 rounded border-slate-300"
                            />
                            <span className="font-medium">YoY %</span>
                          </label>

                          <label className="flex items-center gap-2">
                            <input
                              type="checkbox"
                              checked={showMoMSeries}
                              onChange={(e) => setShowMoMSeries(e.target.checked)}
                              className="h-4 w-4 rounded border-slate-300"
                            />
                            <span className="font-medium">MoM %</span>
                          </label>
                        </div>

                        <div className="mt-2 flex flex-wrap gap-2">
                          <button
                            type="button"
                            onClick={() => {
                              setShowUnitsSeries(false);
                              setShowPrevYearSeries(false);
                              setShowMoMSeries(false);
                              setShowYoYSeries(true);
                            }}
                            className="rounded-lg bg-slate-900 px-2 py-1 text-[12px] font-semibold text-white hover:bg-slate-800"
                          >
                            YoY% only
                          </button>
                          <button
                            type="button"
                            onClick={() => {
                              setShowUnitsSeries(true);
                              setShowPrevYearSeries(true);
                              setShowMoMSeries(false);
                              setShowYoYSeries(false);
                            }}
                            className="rounded-lg bg-white px-2 py-1 text-[12px] font-semibold text-slate-700 ring-1 ring-slate-200 hover:bg-slate-50"
                          >
                            Totals only
                          </button>
                        </div>
                      </div>

                      <div className="mt-1 text-[11px] text-slate-500">
                        Weekly/Monthly uses same day-window for YoY and prior-period % (prevents partial-period distortion).
                      </div>
                    </div>
                  </div>
                </div>

                <div className="h-[340px]">
                  <ResponsiveContainer width="100%" height="100%">
                    <LineChart data={dailyForChart} margin={{ top: 10, right: 18, bottom: 10, left: 8 }}>
                      <CartesianGrid strokeDasharray="3 3" />
                      <XAxis dataKey="label" tick={{ fontSize: 12 }} minTickGap={24} />

                      {anyTotalsShown ? (
                        <YAxis yAxisId="left" tick={{ fontSize: 12 }} tickFormatter={(v) => fmtNum(asFiniteNumber(v) ?? null, 2)} />
                      ) : null}
                      {anyPctShown ? (
                        <YAxis
                          yAxisId="right"
                          orientation="right"
                          tick={{ fontSize: 12 }}
                          tickFormatter={(v) => fmtPct(asFiniteNumber(v) ?? null, 2)}
                        />
                      ) : null}

                      <Tooltip
                        formatter={(v: any, name: any, item: any) => {
                          const key = (item && (item.dataKey as string)) || (name as string);
                          const num = asFiniteNumber(v);

                          if (key === "units") return [`${fmtNum(num ?? null, 2)} units`, "Total Current"];
                          if (key === "prev_year_units") return [`${fmtNum(num ?? null, 2)} units`, "Total (previous year)"];
                          if (key === "yoy_pct") return [fmtPct(num ?? null, 2), "YoY %"];
                          if (key === "mom_pct") return [fmtPct(num ?? null, 2), aggFreq === "weekly" ? "WoW %" : "MoM %"];
                          if (num != null) return [fmtNum(num, 2), String(name)];
                          return [v, String(name)];
                        }}
                      />
                      <Legend />

                      {showUnitsSeries ? (
                        <Line
                          yAxisId="left"
                          type="monotone"
                          dataKey="units"
                          name="Total Current"
                          dot={false}
                          strokeWidth={2}
                          stroke="#dc2626"
                        />
                      ) : null}

                      {showPrevYearSeries ? (
                        <Line
                          yAxisId="left"
                          type="monotone"
                          dataKey="prev_year_units"
                          name="Total (previous year)"
                          dot={false}
                          strokeWidth={2}
                          stroke="#64748b"
                        />
                      ) : null}

                      {showYoYSeries ? (
                        <Line yAxisId="right" type="monotone" dataKey="yoy_pct" name="YoY %" dot={false} strokeWidth={2} stroke="#16a34a" />
                      ) : null}

                      {showMoMSeries ? (
                        <Line yAxisId="right" type="monotone" dataKey="mom_pct" name={aggFreq === "weekly" ? "WoW %" : "MoM %"} dot={false} strokeWidth={2} stroke="#dc2626" />
                      ) : null}
                    </LineChart>
                  </ResponsiveContainer>
                </div>
              </>
            )}
          </Card>

          <Card title="Monthly totals + growth">
            {!hasData ? (
              <div className="text-sm text-slate-600">Add data to see monthly totals and growth.</div>
            ) : (
              <div className="space-y-4">
                <div className="h-[260px]">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart data={monthlyForChart} margin={{ top: 10, right: 18, bottom: 10, left: 8 }}>
                      <CartesianGrid strokeDasharray="3 3" />
                      <XAxis dataKey="month" tick={{ fontSize: 12 }} minTickGap={18} />
                      <YAxis tick={{ fontSize: 12 }} tickFormatter={(v) => fmtNum(asFiniteNumber(v) ?? null, 2)} />
                      <Tooltip
                        formatter={(v: any, n: any) => {
                          const num = asFiniteNumber(v);
                          if (n === "total_units") return [`${fmtNum(num ?? null, 2)} units`, "Monthly total"];
                          if (n === "yoy_pct") return [fmtPct(num ?? null, 2), "YoY"];
                          if (n === "mom_pct") return [fmtPct(num ?? null, 2), "MoM"];
                          return [v, n];
                        }}
                      />
                      <Legend />
                      <Bar dataKey="total_units" name="Monthly total (units)" fill="#dc2626" />
                    </BarChart>
                  </ResponsiveContainer>
                </div>

                <div className="h-[240px]">
                  <ResponsiveContainer width="100%" height="100%">
                    <LineChart data={monthlyForChart} margin={{ top: 10, right: 18, bottom: 10, left: 8 }}>
                      <CartesianGrid strokeDasharray="3 3" />
                      <XAxis dataKey="month" tick={{ fontSize: 12 }} minTickGap={18} />
                      <YAxis tick={{ fontSize: 12 }} tickFormatter={(v) => fmtPct(asFiniteNumber(v) ?? null, 2)} />
                      <Tooltip
                        formatter={(v: any, n: any) => {
                          const num = asFiniteNumber(v);
                          if (n === "yoy_pct") return [fmtPct(num ?? null, 2), "YoY"];
                          if (n === "mom_pct") return [fmtPct(num ?? null, 2), "MoM"];
                          return [v, n];
                        }}
                      />
                      <Legend />
                      <Line type="monotone" dataKey="yoy_pct" name="YoY %" dot={false} strokeWidth={2} stroke="#16a34a" />
                      <Line type="monotone" dataKey="mom_pct" name="MoM %" dot={false} strokeWidth={2} stroke="#dc2626" />
                    </LineChart>
                  </ResponsiveContainer>
                </div>
              </div>
            )}
          </Card>
        </div>

        <div className="mt-6">
          <Card title="Monthly table (last 24 months)">
            {!hasData ? (
              <div className="text-sm text-slate-600">Add data to see the monthly table.</div>
            ) : (
              <div className="overflow-auto rounded-xl ring-1 ring-slate-200">
                <table className="w-full border-collapse bg-white text-left text-sm">
                  <thead className="sticky top-0 bg-slate-50">
                    <tr>
                      <th className="px-3 py-2 text-xs font-semibold text-slate-600">Month</th>
                      <th className="px-3 py-2 text-xs font-semibold text-slate-600">Total (units)</th>
                      <th className="px-3 py-2 text-xs font-semibold text-slate-600">MoM%</th>
                      <th className="px-3 py-2 text-xs font-semibold text-slate-600">YoY%</th>
                    </tr>
                  </thead>
                  <tbody>
                    {monthlyForChart
                      .slice()
                      .reverse()
                      .map((m) => (
                        <tr key={m.month} className="border-t border-slate-100">
                          <td className="px-3 py-2 font-medium text-slate-900">{m.month}</td>
                          <td className="px-3 py-2 text-slate-700">{fmtNum(m.total_units, 2)}</td>
                          <td className="px-3 py-2 text-slate-700">{fmtPct(m.mom_pct, 2)}</td>
                          <td className="px-3 py-2 text-slate-700">{fmtPct(m.yoy_pct, 2)}</td>
                        </tr>
                      ))}
                  </tbody>
                </table>
              </div>
            )}
          </Card>
        </div>
      </div>
    </div>
  );
}
