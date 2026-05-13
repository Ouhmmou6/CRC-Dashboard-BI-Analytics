"use client";

/* eslint-disable @typescript-eslint/no-explicit-any */
import { useTheme } from "next-themes";
import Link from "next/link";
import { type ReactNode, useEffect, useMemo, useRef, useState } from "react";
import {
  Area,
  AreaChart,
  Bar,
  BarChart,
  CartesianGrid,
  Cell,
  Legend,
  Pie,
  PieChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from "recharts";

import {
  applyFilters,
  chartSemanticPalette,
  classifyStatus,
  dailySeries,
  defaultDashboardFilters,
  globalKpis,
  isEmptyField,
  monthlySeries,
  operatorRanking,
  pivotMétierParRégion,
  pivotNatureParRégion,
  pivotRésultatParRégion,
  type DashboardFilters,
} from "@/lib/crc-analytics";
import {
  AGGREGATE_VOLUME_BAR,
  compareResultBuckets,
  getResultColor,
} from "@/lib/constants/chart-colors";
import {
  REGION_COLORS,
  REGION_ORDER,
  REGION_SHORT,
} from "@/lib/crc-constants";
import { displayLabel, parseImportedFile } from "@/lib/crc-parser";
import {
  type CrcChartKey,
  type CrcKpiKey,
  type CrcReportConfig,
  type CrcTableKey,
  type ExcelSheetKey,
  type PdfBundleKey,
  type PptxSlideKey,
  type RawColumnKey,
  CRC_CHART_KEYS,
  CRC_KPI_KEYS,
  CRC_TABLE_KEYS,
  RAW_COLUMN_KEYS,
  defaultCrcReportConfig,
  loadCrcReportConfig,
  resetCrcReportConfig,
  saveCrcReportConfig,
} from "@/lib/crc-report-config";
import { CrcRegionPivotWidget } from "@/components/crc/crc-region-pivot-widget";
import { CrcRegionResultCardWidget } from "@/components/crc/crc-region-result-card-widget";
import { CrcRawPreviewWidget } from "@/components/crc/crc-raw-preview-widget";
import { CrcTeleopStatsWidget } from "@/components/crc/crc-teleop-stats-widget";
import { captureRefToDataUrl } from "@/lib/crc-export-engine";
import {
  exportInvestigationExcel,
  exportInvestigationPdf,
  exportInvestigationPptx,
  type InvestigationColumnKey,
} from "@/lib/crc-export-engine";
import {
  defaultCrcColumnVisibility,
  type CrcColumnVisibilityState,
  type RegionPivotId,
  type TeleOpMetricKey,
} from "@/lib/crc-export-helpers";
import { exportCrcExcel } from "@/lib/export-excel";
import { exportCrcPdf } from "@/lib/export-pdf";
import { exportCrcPowerPoint } from "@/lib/export-pptx";

import type { CrcRow, ParseDebug } from "@/lib/crc-types";
import { ThemeToggle } from "@/components/theme-toggle";

function GlassCard({
  title,
  subtitle,
  accent,
  children,
  action,
}: {
  title?: string;
  subtitle?: string;
  accent?: string;
  children: React.ReactNode;
  action?: React.ReactNode;
}) {
  return (
    <section className="glass-panel p-5 sm:p-6 hover:-translate-y-0.5 hover:shadow-2xl hover:shadow-slate-900/10 dark:hover:shadow-slate-950/40 motion-safe:animate-fade-in">
      <div className="flex flex-col gap-1 sm:flex-row sm:items-start sm:justify-between mb-4">
        <div>
          {title && (
            <h3 className="text-lg font-semibold tracking-tight text-slate-900 dark:text-slate-50">
              {title}
            </h3>
          )}
          {subtitle && (
            <p className="text-sm text-slate-600 dark:text-slate-400 mt-0.5">{subtitle}</p>
          )}
          {accent && (
            <p className="text-xs font-medium text-sky-600 dark:text-sky-300 mt-1">{accent}</p>
          )}
        </div>
        {action}
      </div>
      {children}
    </section>
  );
}

const INVESTIGATION_COLS: InvestigationColumnKey[] = [
  "datetime",
  "teleop",
  "resultat",
  "metier",
  "region",
  "phone",
  "nature",
];

function formatDateTime(d: Date | null) {
  if (!d) return "";
  return `${d.toLocaleDateString("fr-FR")} à ${d.toLocaleTimeString("fr-FR", { hour: "2-digit", minute: "2-digit" })}`;
}

function parseMulti(select: HTMLSelectElement) {
  return [...select.selectedOptions].map((o) => o.value);
}

const CHART_LABEL_FR: Record<CrcChartKey, string> = {
  geoBars: "Histogramme régions canon",
  geoDonut: "Donut géographique",
  statusPie: "Camembert résultats (libellés standardisés)",
  dailyArea: "Courbes cumulées par jour",
  monthlyBars: "Barres empilées par mois",
  trendLine: "Tendance totale jour",
  teleopBars: "Classement téléopérateurs (diagramme)",
  regionCards: "Cartes région Drâa / Laâyoune / Souss / Faux appels",
};

const TABLE_LABEL_FR: Record<CrcTableKey, string> = {
  pivotResult: "Pivot Résultat × régions",
  pivotMetier: "Pivot Métier × régions",
  pivotNature: "Pivot Nature × régions",
  teleOpStats: "Table Statistiques téléopérateurs",
  rawPreview: "Grille brute (extrait)",
};

const PDF_BLOCK_LABELS: { id: PdfBundleKey; label: string }[] = [
  { id: "kpisCards", label: "PDF — Bandeau KPI (cartes synthèse)" },
  { id: "summaryCharts", label: "PDF — Graphes recap globaux (régions / résultat)" },
  { id: "pivotResultTable", label: "PDF — Table Pivot Résultat × régions" },
  { id: "trendsCharts", label: "PDF — Courbes journalière & mensuelle" },
  { id: "metierNatureTables", label: "PDF — Tables Métier & Nature" },
  { id: "teleopPage", label: "PDF — Page téléopérateurs" },
];

const EXCEL_SHEET_LABELS: { id: ExcelSheetKey; label: string }[] = [
  { id: "readme", label: "Feuille README" },
  { id: "kpis", label: "Synthèse KPI" },
  { id: "flat", label: "Données brutes" },
  { id: "pivotResult", label: "Pivot Résultat × régions" },
  { id: "pivotMetier", label: "Pivot Métier × régions" },
  { id: "pivotNature", label: "Pivot Nature × régions" },
  { id: "operators", label: "Classement téléopérateurs" },
  { id: "daily", label: "Évolution journalière" },
  { id: "monthly", label: "Évolution mensuelle" },
];

const PPTX_BLOCK_LABELS: { id: PptxSlideKey; label: string }[] = [
  { id: "cover", label: "Diapo couverture" },
  { id: "kpi", label: "Diapo KPI régions" },
  { id: "results", label: "Diapo résultats + tendance" },
  { id: "metier", label: "Diapo métiers" },
  { id: "operators", label: "Diapo téléopérateurs" },
  { id: "definitions", label: "Diapo définitions + mensuel" },
];

const KPI_LABEL_FR: Record<CrcKpiKey, string> = {
  totalVolume: "Total interactions",
  abandons: "Appels abandonnés",
  decrochesInterrompus: "Appels décrochés interrompus",
  informes: "Clients informés",
  tickets: "Tickets transmis",
  teleopsDistinct: "Téléopérateurs actifs",
  pctInformes: "Part clients informés",
  pctTickets: "% tickets transmis",
  coverage: "Couverture filtre",
};

const RAW_PREVIEW_COLUMNS: {
  key: RawColumnKey;
  label: string;
  tdClass?: string;
  cell: (r: CrcRow) => ReactNode;
}[] = [
  {
    key: "date",
    label: "Date",
    cell: (r) => (r.date ? r.date.toLocaleDateString("fr-FR") : ""),
  },
  {
    key: "resultatRaw",
    label: "Résultat (Excel)",
    cell: (r) => (
      <span className="font-semibold" style={{ color: getResultColor(r.résultat) }}>
        {r.résultatRaw || "—"}
      </span>
    ),
  },
  {
    key: "teleop",
    label: "Téléopérateur",
    cell: (r) => r.téléopérateur,
  },
  { key: "metier", label: "Métier", cell: (r) => r.metier },
  {
    key: "nature",
    label: "Nature",
    cell: (r) => r.natureRéclamation,
  },
  {
    key: "regionCanon",
    label: "Région canon",
    cell: (r) => (
      <span className="font-semibold" style={{ color: REGION_COLORS[r.régionCanon] }}>
        {REGION_SHORT[r.régionCanon]}
      </span>
    ),
  },
  {
    key: "regionsource",
    label: "Région brute",
    cell: (r) => r.regions,
  },
  {
    key: "valid",
    label: "Valide ?",
    cell: (r) => (
      <span className={r.valid ? "" : "text-rose-600 dark:text-rose-400"}>
        {r.valid ? "oui" : "non"}
      </span>
    ),
  },
  {
    key: "phone",
    label: "Téléphone",
    tdClass: "font-mono text-[10px]",
    cell: (r) => r.téléphone,
  },
];

export default function CrcDashboard() {
  const { resolvedTheme } = useTheme();
  const [mounted, setMounted] = useState(false);
  useEffect(() => setMounted(true), []);

  const [reportConfig, setReportConfig] = useState<CrcReportConfig>(() =>
    typeof window !== "undefined" ? loadCrcReportConfig() : defaultCrcReportConfig(),
  );

  useEffect(() => {
    setReportConfig(loadCrcReportConfig());
  }, []);

  const persistReportConfig = (next: CrcReportConfig) => {
    setReportConfig(next);
    saveCrcReportConfig(next);
  };

  const isDark = mounted && resolvedTheme === "dark";
  const palette = useMemo(() => chartSemanticPalette(isDark), [isDark]);

  const [rows, setRows] = useState<CrcRow[]>([]);
  const [debug, setDebug] = useState<ParseDebug | null>(null);
  const [sourceLabel, setSourceLabel] = useState("");
  const [busy, setBusy] = useState(false);
  const [filters, setFilters] = useState<DashboardFilters>(defaultDashboardFilters());
  const [debugOpen, setDebugOpen] = useState(false);
  const [exportUsesFilters, setExportUsesFilters] = useState(true);

  const onFile = async (files: FileList | null) => {
    if (!files?.length) return;
    setBusy(true);
    try {
      const parsed = await parseImportedFile(files[0]);
      setRows(parsed.rows);
      setDebug(parsed.debug);
      setSourceLabel(files[0].name);
      const ops = [...new Set(parsed.rows.map((r) => r.téléopérateur))].sort((a, b) =>
        a.localeCompare(b, "fr"),
      );
      const res = [...new Set(parsed.rows.map((r) => r.résultat))].sort((a, b) =>
        a.localeCompare(b, "fr"),
      );
      /** reset filtres fichier */
      setFilters({
        ...defaultDashboardFilters(),
        téléopérateurs: ops,
        résultats: res,
        régions: [...REGION_ORDER],
      });
      setColumnVisibility(defaultCrcColumnVisibility());
    } finally {
      setBusy(false);
    }
  };

  /** empty télé/rés filt set = désactivé ⇒ pas de sous-ensemble forcé après init */
  const effectiveFilters = useMemo(() => {
    const téléAll = [...new Set(rows.map((r) => r.téléopérateur))];
    const résAll = [...new Set(rows.map((r) => r.résultat))];
    return {
      ...filters,
      téléopérateurs:
        filters.téléopérateurs.length > 0 &&
        filters.téléopérateurs.length < téléAll.length &&
        téléAll.length > 0
          ? filters.téléopérateurs
          : [],
      résultats:
        filters.résultats.length > 0 &&
        filters.résultats.length < résAll.length &&
        résAll.length > 0
          ? filters.résultats
          : [],
    };
  }, [filters, rows]);

  const filteredRows = useMemo(() => applyFilters(rows, effectiveFilters), [rows, effectiveFilters]);
  const [metierResultatFilter, setMetierResultatFilter] = useState("all");

  const kpis = useMemo(() => globalKpis(filteredRows), [filteredRows]);
  const pivotResult = useMemo(() => pivotRésultatParRégion(filteredRows), [filteredRows]);
  const pivotMet = useMemo(() => {
  const scoped =
    metierResultatFilter === "all"
      ? filteredRows
      : filteredRows.filter(
          (r) => r.résultat === metierResultatFilter,
        );

  return pivotMétierParRégion(scoped);
}, [filteredRows, metierResultatFilter]);
  const pivotNat = useMemo(() => pivotNatureParRégion(filteredRows), [filteredRows]);
  const sérieJour = useMemo(() => dailySeries(filteredRows), [filteredRows]);
  const sérieMois = useMemo(() => monthlySeries(filteredRows), [filteredRows]);
  const téléopRanking = useMemo(() => operatorRanking(filteredRows), [filteredRows]);

  const téléOptions = [...new Set(rows.map((r) => r.téléopérateur))].sort((a, b) =>
    a.localeCompare(b, "fr"),
  );
  const résOptions = [...new Set(rows.map((r) => r.résultat))].sort((a, b) =>
    a.localeCompare(b, "fr"),
  );
  const metierResultatOptions = useMemo(
  () =>
    Array.from(
      new Set(
        filteredRows
          .map((r) => r.résultat)
          .filter(Boolean),
      ),
    ).sort((a, b) => a.localeCompare(b, "fr")),
  [filteredRows],
);

  const exportDataset = exportUsesFilters ? filteredRows : rows;

  const [columnVisibility, setColumnVisibility] = useState<CrcColumnVisibilityState>(defaultCrcColumnVisibility);

  const pivotResultTableRef = useRef<HTMLDivElement>(null);
  const pivotResultChartRef = useRef<HTMLDivElement>(null);
  const pivotMetierTableRef = useRef<HTMLDivElement>(null);
  const pivotMetierChartRef = useRef<HTMLDivElement>(null);
  const pivotNatureTableRef = useRef<HTMLDivElement>(null);
  const pivotNatureChartRef = useRef<HTMLDivElement>(null);
  const teleTableRef = useRef<HTMLDivElement>(null);
  const teleChartRef = useRef<HTMLDivElement>(null);

  const exportColumnVisibility = useMemo(
    () => ({
      pivotResultRegions: columnVisibility.pivotRegions.pivotResult,
      pivotMetierRegions: columnVisibility.pivotRegions.pivotMetier,
      pivotNatureRegions: columnVisibility.pivotRegions.pivotNature,
      teleOpMetrics: columnVisibility.teleOpMetrics,
    }),
    [columnVisibility],
  );

  const togglePivotRegion = (id: RegionPivotId, short: string) => {
    setColumnVisibility((v) => {
      const prevMap = { ...v.pivotRegions[id] };
      const visible = prevMap[short] !== false;
      prevMap[short] = !visible;
      return {
        ...v,
        pivotRegions: { ...v.pivotRegions, [id]: prevMap },
      };
    });
  };

  const toggleTeleOpMetric = (key: TeleOpMetricKey) => {
    setColumnVisibility((v) => {
      const m = { ...v.teleOpMetrics };
      m[key] = !(m[key] !== false);
      return { ...v, teleOpMetrics: m };
    });
  };

  const toggleRawColumn = (key: RawColumnKey) => {
    persistReportConfig({
      ...reportConfig,
      rawColumns: { ...reportConfig.rawColumns, [key]: !reportConfig.rawColumns[key] },
    });
  };

  const collectWidgetChartImages = async () => {
    const snap = async (chart: HTMLDivElement | null, table: HTMLDivElement | null) =>
      captureRefToDataUrl(chart ?? table);
    const [pivotResult, pivotMetier, pivotNature, teleOp] = await Promise.all([
      snap(pivotResultChartRef.current, pivotResultTableRef.current),
      snap(pivotMetierChartRef.current, pivotMetierTableRef.current),
      snap(pivotNatureChartRef.current, pivotNatureTableRef.current),
      snap(teleChartRef.current, teleTableRef.current),
    ]);
    const out: Partial<Record<"pivotResult" | "pivotMetier" | "pivotNature" | "teleOp", string>> = {};
    if (pivotResult) out.pivotResult = pivotResult;
    if (pivotMetier) out.pivotMetier = pivotMetier;
    if (pivotNature) out.pivotNature = pivotNature;
    if (teleOp) out.teleOp = teleOp;
    return out;
  };

  const exportFileBase = useMemo(
    () => `CRC_${(sourceLabel || "export").replace(/\W+/g, "_")}`,
    [sourceLabel],
  );

  const [detailKpi, setDetailKpi] = useState<null | "abandons" | "decroches" | "informes" | "tickets">(null);
  const [detailRegions, setDetailRegions] = useState<Record<string, boolean>>(
    Object.fromEntries(REGION_ORDER.map((r) => [r, true])),
  );
  const [fauxPage, setFauxPage] = useState(1);
  const [fauxResultatFilter, setFauxResultatFilter] = useState("all");

  const fauxResultatOptions = useMemo(
  () =>
    Array.from(
      new Set(
        filteredRows
          .map((r) => r.résultat)
          .filter(Boolean),
      ),
    ).sort((a, b) => a.localeCompare(b, "fr")),
  [filteredRows],
);

const fauxTraitementsRows = useMemo(() => {
  let rowsFiltered = filteredRows.filter(
    (r) => isEmptyField(r.regions) || isEmptyField(r.metier),
  );

  if (fauxResultatFilter !== "all") {
    rowsFiltered = rowsFiltered.filter(
      (r) => r.résultat === fauxResultatFilter,
    );
  }

  return rowsFiltered;
}, [filteredRows, fauxResultatFilter]);

  const fauxPerPage = 25;
  const fauxPaged = useMemo(
    () => fauxTraitementsRows.slice((fauxPage - 1) * fauxPerPage, fauxPage * fauxPerPage),
    [fauxTraitementsRows, fauxPage],
  );

  const detailRows = useMemo(() => {
    if (!detailKpi) return [] as CrcRow[];
    const byKpi = filteredRows.filter((r) => {
      const st = classifyStatus(r.résultat);
      if (detailKpi === "abandons") return st === "abandon";
      if (detailKpi === "decroches") return st === "appel_abandonne";
      if (detailKpi === "informes") return st === "client_informe";
      return st === "ticket_transmis";
    });
    return byKpi.filter((r) => detailRegions[r.régionCanon] !== false);
  }, [detailKpi, filteredRows, detailRegions]);

const chartTooltip = (
  <Tooltip
    contentStyle={{
      backgroundColor: palette.tooltipBg,
      borderRadius: 12,
      border: `1px solid ${palette.grid}`,
    }}
    labelStyle={{ color: palette.fg }}
    formatter={(value: number, name: string) => {
      const labels: Record<string, string> = {
        volume: "Volume",
        informés: "Clients informés",
        tickets: "Tickets transmis",
        abandons: "Appels abandonnés",
        appelsDécrochésInterrompus:
          "Appels décrochés interrompus",
      };

      return [
        `${Number(value)?.toLocaleString("fr-FR") ?? 0} lignes`,
        labels[name] || name,
      ];
    }}
  />
);

  const donutRég = REGION_ORDER.map((rg) => ({
    name: REGION_SHORT[rg],
    value: kpis.appelsParRégion.get(rg) ?? 0,
    fill: REGION_COLORS[rg],
  }));

  const résultatPieData = useMemo(() => {
    const buckets = filteredRows.map((r) => r.résultat);
    const uniq = [...new Set(buckets)].sort(compareResultBuckets);
    return uniq
      .map((name) => ({
        name,
        value: buckets.filter((x) => x === name).length,
        fill: getResultColor(name),
      }))
      .filter((d) => d.value > 0);
  }, [filteredRows]);

const téléBar = téléopRanking.slice(0, 12).map((o) => ({
  nom: o.name.length > 18 ? `${o.name.slice(0, 17)}…` : o.name,

  volume: o.volume,

  informés: o.informés,

  tickets: o.tickets,

  abandons: o.abandons,

  appelsDécrochésInterrompus:
    o.appelsDécrochésInterrompus ?? 0,
}));

  const exportTitlePdf =
    reportConfig.reportTitle.trim() ||
    "Département Clientèle et Suivi des Performances CRC";

  const kpiTiles = useMemo(() => {
    type Item = { key: CrcKpiKey; title: string; subtitle: string; body: ReactNode };
    const tiles: Item[] = [
      {
        key: "totalVolume",
        title: KPI_LABEL_FR.totalVolume,
        subtitle: "Lignes présentes après filtres cockpit.",
        body: kpis.totalAppels,
      },
      {
        key: "abandons",
        title: KPI_LABEL_FR.abandons,
        subtitle: "Lignes dont le résultat normalisé est « Appels abandonnés ».",
        body: kpis.appelsAbandonnés,
      },
      {
        key: "decrochesInterrompus",
        title: KPI_LABEL_FR.decrochesInterrompus,
        subtitle: "Lignes dont le résultat normalisé est « Appels décrochés interrompus ».",
        body: kpis.appelsDécrochésInterrompus,
      },
      {
        key: "informes",
        title: KPI_LABEL_FR.informes,
        subtitle: "Lignes dont le résultat normalisé est « Clients informés ».",
        body: kpis.clientsInformés,
      },
      {
        key: "tickets",
        title: KPI_LABEL_FR.tickets,
        subtitle: "Lignes dont le résultat normalisé est « Tickets transmis ».",
        body: kpis.ticketsTransmis,
      },
      {
        key: "teleopsDistinct",
        title: KPI_LABEL_FR.teleopsDistinct,
        subtitle: "Identifiants distincts post filtres cockpit.",
        body: new Set(filteredRows.map((r) => r.téléopérateur)).size,
      },
      {
        key: "pctInformes",
        title: KPI_LABEL_FR.pctInformes,
        subtitle: "clients informés / volume filtré",
        body: filteredRows.length
          ? `${((kpis.clientsInformés / filteredRows.length) * 100).toFixed(1)} %`
          : "—",
      },
      {
        key: "pctTickets",
        title: KPI_LABEL_FR.pctTickets,
        subtitle: "tickets transmis / volume filtré",
        body: filteredRows.length
          ? `${((kpis.ticketsTransmis / filteredRows.length) * 100).toFixed(1)} %`
          : "—",
      },
      {
        key: "coverage",
        title: KPI_LABEL_FR.coverage,
        subtitle: "préservation intégrale base source",
        body: `${filteredRows.length} / ${rows.length}`,
      },
    ];
    return tiles.filter((t) => reportConfig.kpis[t.key]);
  }, [filteredRows, kpis, reportConfig.kpis, rows.length]);

  return (
    <div className="max-w-[1640px] mx-auto px-4 pb-20 pt-10 space-y-8">
      <header className="flex flex-wrap items-center gap-4 justify-between">
        <div>
          <p className="text-[11px] uppercase tracking-[0.22em] text-sky-600 dark:text-sky-300 font-bold">
            Axilus CRC Operational
          </p>
          <h1 className="text-3xl sm:text-4xl font-bold text-slate-900 dark:text-white tracking-tight">
            Tableau de bord CRC & relations clients
          </h1>
          <p className="text-sm text-slate-600 dark:text-slate-400 max-w-2xl mt-2 leading-relaxed">
            Données issues des exports bureau Axilus : traitements, contacts répétés et réclamations
            multiplex conservés intégralement (pas de suppression de lignes ou d&apos;
            identifiants fictifs obsolètes).
          </p>
        </div>
        <div className="flex flex-wrap items-center gap-2">
          <Link
            href="/business-intelligence"
            className="rounded-full px-3 py-2 text-xs font-semibold border border-slate-300 dark:border-slate-600"
          >
            BI
          </Link>
          <ThemeToggle />
          <button
            type="button"
            onClick={() => setDebugOpen((x) => !x)}
            className="rounded-full px-4 py-2 text-xs font-semibold bg-slate-900 text-white dark:bg-white dark:text-slate-950 shadow hover:opacity-90 transition"
          >
            {debugOpen ? "Fermer debug" : "Mode développeur"}
          </button>
        </div>
      </header>

      <GlassCard title="Imports & diffusion" subtitle="CSV / Excel Axilus (.csv, .xls, .xlsx)">
        <div
          className="flex flex-col gap-3 lg:flex-row lg:items-center flex-wrap"
          onDragOver={(e) => {
            e.preventDefault();
          }}
          onDrop={(e) => {
            e.preventDefault();
            void onFile(e.dataTransfer.files);
          }}
        >
          <label className="inline-flex px-6 py-3 rounded-3xl cursor-pointer bg-gradient-to-r from-sky-500 via-indigo-500 to-violet-500 text-white text-sm font-semibold shadow-lg shadow-indigo-500/30 hover:brightness-105 active:translate-y-[1px] transition">
            <input
              type="file"
              accept=".csv,.xls,.xlsx"
              className="hidden"
              onChange={(e) => void onFile(e.target.files)}
            />
            Importer l&apos;export Axilus
          </label>
          {busy ? <span className="text-sm text-slate-500 animate-pulse">Analyse workbook…</span> : null}
          {sourceLabel ? (
            <span className="text-xs px-4 py-1.5 rounded-full border border-emerald-200 dark:border-emerald-800 bg-emerald-50/70 dark:bg-emerald-950/40 font-mono text-emerald-800 dark:text-emerald-300">
              {sourceLabel}
            </span>
          ) : null}
          <label className="ml-auto lg:ml-0 inline-flex gap-2 text-xs items-center text-slate-600 dark:text-slate-400">
            <input
              type="checkbox"
              checked={exportUsesFilters}
              onChange={(e) => setExportUsesFilters(e.target.checked)}
            />
            Exports reflètent les filtres actifs
          </label>
          <div className="flex gap-2 flex-wrap">
            <button
              type="button"
              disabled={!rows.length}
              onClick={() =>
                exportCrcExcel(
                  exportDataset,
                  sourceLabel || "crc_export",
                  reportConfig.exportExcelSheets,
                  exportColumnVisibility,
                )
              }
              className="rounded-2xl px-4 py-2 text-xs font-bold bg-emerald-600 hover:bg-emerald-500 disabled:opacity-35 text-white"
            >
              Excel
            </button>
            <button
              type="button"
              disabled={!rows.length}
              onClick={() => {
                void (async () => {
                  const widgetChartImages = await collectWidgetChartImages();
                  await exportCrcPdf(exportDataset, exportTitlePdf, exportFileBase, {
                    subtitle: reportConfig.reportSubtitle.trim() || undefined,
                    logoOverride: reportConfig.logoDataUrl,
                    pdf: reportConfig.exportPdf,
                    columnVisibility: exportColumnVisibility,
                    widgetChartImages,
                  });
                })();
              }}
              className="rounded-2xl px-4 py-2 text-xs font-bold bg-sky-600 hover:bg-sky-500 disabled:opacity-35 text-white"
            >
              PDF
            </button>
            <button
              type="button"
              disabled={!rows.length}
              onClick={() => {
                void exportCrcPowerPoint(
                  exportDataset,
                  reportConfig.reportTitle.trim() || `CRC — Pilotage (${sourceLabel || "Axilus"})`,
                  `CRC_ppt_${(sourceLabel || "Axilus").replace(/\W+/g, "_")}`,
                  {
                    subtitle: reportConfig.reportSubtitle.trim() || undefined,
                    logoOverride: reportConfig.logoDataUrl,
                    columnVisibility: exportColumnVisibility,
                    dashboard: {
                      charts: reportConfig.charts,
                      tables: reportConfig.tables,
                      kpis: reportConfig.kpis,
                      rawColumns: reportConfig.rawColumns,
                      exportPptx: reportConfig.exportPptx,
                    },
                  },
                );
              }}
              className="rounded-2xl px-4 py-2 text-xs font-bold bg-purple-600 hover:bg-purple-500 disabled:opacity-35 text-white"
            >
              PowerPoint
            </button>
          </div>
        </div>
      </GlassCard>

      <GlassCard
        title="Configuration du rapport"
        subtitle="Personnaliser titres/logo, colonnes grille, briques cockpit et sections exportées (PDF/PPTX/XLSX) — aucune ligne de code requise."
      >
        <div className="grid gap-6 xl:grid-cols-12">
          <div className="space-y-4 xl:col-span-5">
            <div>
              <p className="text-xs uppercase font-semibold text-slate-600 dark:text-slate-400 mb-1">
                Titre (exports structurés)
              </p>
              <input
                className="w-full rounded-2xl border border-slate-200 dark:border-slate-700 bg-white/80 dark:bg-slate-950/60 px-4 py-2 text-sm text-slate-900 dark:text-slate-50"
                value={reportConfig.reportTitle}
                placeholder="Laissé vide ⇒ libellés CRC automatiques."
                onChange={(e) => persistReportConfig({ ...reportConfig, reportTitle: e.target.value })}
              />
            </div>
            <div>
              <p className="text-xs uppercase font-semibold text-slate-600 dark:text-slate-400 mb-1">
                Sous-titre / mention légale
              </p>
              <input
                className="w-full rounded-2xl border border-slate-200 dark:border-slate-700 bg-white/80 dark:bg-slate-950/60 px-4 py-2 text-sm text-slate-900 dark:text-slate-50"
                value={reportConfig.reportSubtitle}
                placeholder='Exemple : "Agence XYZ — novembre 2025"'
                onChange={(e) => persistReportConfig({ ...reportConfig, reportSubtitle: e.target.value })}
              />
            </div>
            <div>
              <p className="text-xs uppercase font-semibold text-slate-600 dark:text-slate-400 mb-2">
                Logo export (prioritaire sur `/srm-logo.png`)
              </p>
              <div className="flex flex-wrap items-center gap-2">
                <label className="rounded-2xl border border-dashed border-slate-300 dark:border-slate-600 px-4 py-2 text-xs font-semibold cursor-pointer hover:bg-slate-50 dark:hover:bg-slate-800/70">
                  Importer un logo PNG/JPEG
                  <input
                    type="file"
                    accept="image/png,image/jpeg,image/webp"
                    className="hidden"
                    onChange={(ev) => {
                      const file = ev.target.files?.[0];
                      if (!file) return;
                      const rd = new FileReader();
                      rd.onload = () =>
                        persistReportConfig({
                          ...reportConfig,
                          logoDataUrl: String(rd.result || ""),
                        });
                      rd.readAsDataURL(file);
                    }}
                  />
                </label>
                <button
                  type="button"
                  className="rounded-2xl border border-slate-300 px-4 py-2 text-xs dark:border-slate-600 disabled:opacity-40"
                  disabled={!reportConfig.logoDataUrl}
                  onClick={() => persistReportConfig({ ...reportConfig, logoDataUrl: null })}
                >
                  Réinitialiser logo upload
                </button>
              </div>
            </div>
          </div>
          <div className="space-y-5 xl:col-span-7">
            <div>
              <p className="text-xs uppercase font-semibold text-slate-600 dark:text-slate-400 mb-2">
                Afficher sur le tableau de bord
              </p>
              <div className="grid sm:grid-cols-2 gap-2 max-h-[200px] overflow-y-auto pr-1">
                {CRC_CHART_KEYS.map((k) => (
                  <label
                    key={k}
                    className="flex items-start gap-2 text-xs rounded-xl border border-slate-100 dark:border-slate-700/70 px-2 py-2"
                  >
                    <input
                      type="checkbox"
                      checked={reportConfig.charts[k]}
                      onChange={(e) =>
                        persistReportConfig({
                          ...reportConfig,
                          charts: { ...reportConfig.charts, [k]: e.target.checked },
                        })
                      }
                    />
                    <span>{CHART_LABEL_FR[k]}</span>
                  </label>
                ))}
                <div className="sm:col-span-2 mt-2 text-[11px] font-semibold text-slate-500 uppercase">
                  Tableaux cockpit
                </div>
                {CRC_TABLE_KEYS.map((k) => (
                  <label
                    key={k}
                    className="flex items-start gap-2 text-xs rounded-xl border border-slate-100 dark:border-slate-700/70 px-2 py-2"
                  >
                    <input
                      type="checkbox"
                      checked={reportConfig.tables[k]}
                      onChange={(e) =>
                        persistReportConfig({
                          ...reportConfig,
                          tables: { ...reportConfig.tables, [k]: e.target.checked },
                        })
                      }
                    />
                    <span>{TABLE_LABEL_FR[k]}</span>
                  </label>
                ))}
                <div className="sm:col-span-2 mt-2 text-[11px] font-semibold text-slate-500 uppercase">
                  Cartes KPI
                </div>
                {CRC_KPI_KEYS.map((k) => (
                  <label
                    key={k}
                    className="flex items-start gap-2 text-xs rounded-xl border border-slate-100 dark:border-slate-700/70 px-2 py-2"
                  >
                    <input
                      type="checkbox"
                      checked={reportConfig.kpis[k]}
                      onChange={(e) =>
                        persistReportConfig({
                          ...reportConfig,
                          kpis: { ...reportConfig.kpis, [k]: e.target.checked },
                        })
                      }
                    />
                    <span>{KPI_LABEL_FR[k]}</span>
                  </label>
                ))}
              </div>
            </div>
            <div>
              <p className="text-xs uppercase font-semibold text-slate-600 dark:text-slate-400 mb-2">
                Colonnes visibles grille brute
              </p>
              <div className="flex flex-wrap gap-2">
                {RAW_COLUMN_KEYS.map((col) => (
                  <label
                    key={col}
                    className="flex items-center gap-1 px-3 py-1 rounded-full border border-slate-200 dark:border-slate-700 text-[11px]"
                  >
                    <input
                      type="checkbox"
                      checked={reportConfig.rawColumns[col]}
                      onChange={(e) =>
                        persistReportConfig({
                          ...reportConfig,
                          rawColumns: { ...reportConfig.rawColumns, [col]: e.target.checked },
                        })
                      }
                    />
                    {RAW_PREVIEW_COLUMNS.find((c) => c.key === col)?.label ?? col}
                  </label>
                ))}
              </div>
            </div>
          </div>
          <div className="xl:col-span-12 grid lg:grid-cols-3 gap-4 pt-4 border-t border-slate-200/70 dark:border-slate-700/70">
            <div>
              <p className="text-xs uppercase font-semibold mb-2 text-orange-700 dark:text-orange-200">
                Sections PDF export
              </p>
              <div className="space-y-2 max-h-48 overflow-y-auto pr-1">
                {PDF_BLOCK_LABELS.map(({ id, label }) => (
                  <label key={id} className="flex gap-2 text-xs items-start leading-snug">
                    <input
                      type="checkbox"
                      checked={reportConfig.exportPdf[id]}
                      onChange={(e) =>
                        persistReportConfig({
                          ...reportConfig,
                          exportPdf: { ...reportConfig.exportPdf, [id]: e.target.checked },
                        })
                      }
                    />
                    <span>{label}</span>
                  </label>
                ))}
              </div>
            </div>
            <div>
              <p className="text-xs uppercase font-semibold mb-2 text-emerald-700 dark:text-emerald-200">
                Feuilles Excel export
              </p>
              <div className="space-y-2 max-h-48 overflow-y-auto pr-1">
                {EXCEL_SHEET_LABELS.map(({ id, label }) => (
                  <label key={id} className="flex gap-2 text-xs items-start leading-snug">
                    <input
                      type="checkbox"
                      checked={reportConfig.exportExcelSheets[id]}
                      onChange={(e) =>
                        persistReportConfig({
                          ...reportConfig,
                          exportExcelSheets: { ...reportConfig.exportExcelSheets, [id]: e.target.checked },
                        })
                      }
                    />
                    <span>{label}</span>
                  </label>
                ))}
              </div>
            </div>
            <div className="space-y-2">
              <p className="text-xs uppercase font-semibold mb-2 text-purple-700 dark:text-purple-200">
                Diapositives PowerPoint
              </p>
              <div className="space-y-2 max-h-48 overflow-y-auto pr-1">
                {PPTX_BLOCK_LABELS.map(({ id, label }) => (
                  <label key={id} className="flex gap-2 text-xs items-start leading-snug">
                    <input
                      type="checkbox"
                      checked={reportConfig.exportPptx[id]}
                      onChange={(e) =>
                        persistReportConfig({
                          ...reportConfig,
                          exportPptx: { ...reportConfig.exportPptx, [id]: e.target.checked },
                        })
                      }
                    />
                    <span>{label}</span>
                  </label>
                ))}
              </div>
              <button
                type="button"
                onClick={() => setReportConfig(resetCrcReportConfig())}
                className="mt-4 w-full rounded-2xl border border-slate-300 dark:border-slate-600 px-4 py-2 text-xs font-semibold"
              >
                Réinitialiser la configuration cockpit
              </button>
              <p className="text-[11px] text-slate-500 mt-2">
                La configuration est stockée sur ce navigateur (`localStorage`). Les exports reflètent
                immédiatement ces cases cochées.
              </p>
            </div>
          </div>
        </div>
      </GlassCard>

      <section className="glass-panel px-6 py-5 rounded-3xl border border-amber-400/35 bg-gradient-to-r from-orange-500/15 via-transparent to-transparent dark:from-orange-950/55">
        <h2 className="text-base font-semibold text-amber-950 dark:text-orange-50 mb-2">
          Interprétation opérationnelle des rubriques Résultat
        </h2>
        <ul className="grid sm:grid-cols-2 gap-3 text-sm text-slate-900 dark:text-slate-50">
          <li>
            <span className="font-semibold text-orange-700 dark:text-orange-200">Appels abandonnés</span>
            {' — '}appel entré mais aucune prise téléconseiller.
          </li>
          <li>
            <span className="font-semibold text-orange-700 dark:text-orange-200">Appels décrochés interrompus</span>
            {' — '}liaison coupée alors que la file était active.
          </li>
          <li>
            <span className="font-semibold text-emerald-700 dark:text-emerald-200">Clients informés</span>
            {' — '}résolution informationnelle constatée.
          </li>
          <li>
            <span className="font-semibold text-purple-700 dark:text-purple-200">Tickets transmis</span>
            {' — '}transfert officiel dans le workflow back-office ou partenaires.
          </li>
        </ul>
      </section>

      {rows.length === 0 && (
        <GlassCard title="Aucun fichier sélectionné" subtitle="Chargez vos colonnes Axilus pour activer tous les graphes enterprise.">
          <p className="text-sm text-slate-600 dark:text-slate-400 leading-relaxed">
            Colonnes attendues parmi autres :{' '}
            <code className="text-[11px] bg-slate-100 dark:bg-slate-800 px-2 py-0.5 rounded">
              Campagne/Nom • Date • Page3/Nature • Téléopérateur • Résultat • Page3/Regions •
              Page3/Metier
            </code>
            . Labels normalisés (sans préfixes Header/ ou Page3/) pour l&apos;ensemble du cockpit.
          </p>
        </GlassCard>
      )}

      {rows.length > 0 && (
        <>
          <GlassCard title="Filtres cockpit" subtitle="Segmentation analytique uniquement ; la grille source conserve toutes ses lignes.">
            <div className="grid gap-4 lg:grid-cols-[1.05fr_minmax(0,1fr)]">
              <div className="space-y-4">
                <label className="flex gap-2 text-sm leading-snug">
                  <input
                    type="checkbox"
                    checked={filters.onlyValid}
                    onChange={(e) =>
                      setFilters((f) => ({
                        ...f,
                        onlyValid: e.target.checked,
                      }))
                    }
                  />
                  <span>Analyser seulement les lignes&nbsp;<strong>validées</strong> (date lisible et au moins une dimension opérationnelle renseignée).</span>
                </label>

                <div>
                  <p className="text-xs font-bold uppercase text-slate-500 mb-2">Bassins RCC</p>
                  <div className="flex flex-wrap gap-x-5 gap-y-2 text-xs font-medium">
                    {REGION_ORDER.map((rg) => (
                      <label key={rg} className="inline-flex items-center gap-2">
                        <input
                          type="checkbox"
                          checked={filters.régions.includes(rg)}
                          onChange={(e) => {
                            setFilters((prev) => {
                              if (e.target.checked)
                                return { ...prev, régions: [...new Set([...prev.régions, rg])] };
                              return {
                                ...prev,
                                régions: prev.régions.filter((x) => x !== rg),
                              };
                            });
                          }}
                        />
                        <span style={{ color: REGION_COLORS[rg] }}>{REGION_SHORT[rg]}</span>
                      </label>
                    ))}
                  </div>
                </div>

                <div>
                  <p className="text-xs font-bold uppercase text-slate-500 mb-1">Bornes dates</p>
                  <div className="flex flex-wrap gap-2">
                    <input
                      type="date"
                      className="rounded-xl border px-3 py-1.5 text-sm bg-white/80 dark:bg-slate-900/70 dark:border-slate-600"
                      value={filters.dateFrom ?? ""}
                      onChange={(e) =>
                        setFilters((f) => ({
                          ...f,
                          dateFrom: e.target.value || null,
                        }))
                      }
                    />
                    <input
                      type="date"
                      className="rounded-xl border px-3 py-1.5 text-sm bg-white/80 dark:bg-slate-900/70 dark:border-slate-600"
                      value={filters.dateTo ?? ""}
                      onChange={(e) =>
                        setFilters((f) => ({
                          ...f,
                          dateTo: e.target.value || null,
                        }))
                      }
                    />
                  </div>
                </div>

                <p className="text-[11px] text-slate-500">
                  Multi-sélection : utilisez CTRL / Maj sur Windows ; CMD sur macOS pour affiner téléopérator / résultat.
                </p>
              </div>

              <div className="grid sm:grid-cols-2 gap-3">
                <div className="flex flex-col gap-1">
                  <span className="text-xs uppercase font-semibold text-slate-600 dark:text-slate-400">
                    Téléopérateurs
                  </span>
                  <select
                    multiple
                    className="h-52 rounded-2xl border border-slate-200 dark:border-slate-700 px-3 py-2 text-xs bg-white/80 dark:bg-slate-950/65"
                    value={filters.téléopérateurs}
                    onChange={(e) =>
                      setFilters((f) => ({
                        ...f,
                        téléopérateurs: parseMulti(e.target),
                      }))
                    }
                  >
                    {téléOptions.map((t) => (
                      <option key={t} value={t}>
                        {t}
                      </option>
                    ))}
                  </select>
                </div>
                <div className="flex flex-col gap-1">
                  <span className="text-xs uppercase font-semibold text-slate-600 dark:text-slate-400">
                    Résultat (normalisé)
                  </span>
                  <select
                    multiple
                    className="h-52 rounded-2xl border border-slate-200 dark:border-slate-700 px-3 py-2 text-xs bg-white/80 dark:bg-slate-950/65"
                    value={filters.résultats}
                    onChange={(e) =>
                      setFilters((f) => ({
                        ...f,
                        résultats: parseMulti(e.target),
                      }))
                    }
                  >
                    {résOptions.map((r) => (
                      <option key={r} value={r}>
                        {r}
                      </option>
                    ))}
                  </select>
                </div>
              </div>
            </div>
          </GlassCard>

          {kpiTiles.length ? (
            <div className="grid gap-3 sm:grid-cols-2 xl:grid-cols-4">
              {kpiTiles.map((tile) => {
                const clickable = ["abandons", "decrochesInterrompus", "informes", "tickets"].includes(tile.key);
                return (
                  <GlassCard
                    key={tile.key}
                    title={tile.title}
                    subtitle={tile.subtitle}
                    action={
                      clickable ? (
                        <button
                          type="button"
                          onClick={() =>
                            setDetailKpi(
                              tile.key === "abandons"
                                ? "abandons"
                                : tile.key === "decrochesInterrompus"
                                  ? "decroches"
                                  : tile.key === "informes"
                                    ? "informes"
                                    : "tickets",
                            )
                          }
                          className="rounded-full px-3 py-1 text-[11px] font-semibold border border-slate-300 dark:border-slate-600"
                        >
                          Détails
                        </button>
                      ) : null
                    }
                  >
                    <div className="text-4xl font-bold tabular-nums text-indigo-600 dark:text-indigo-300 mt-2">
                      {tile.body}
                    </div>
                  </GlassCard>
                );
              })}
            </div>
          ) : null}

          <div className="grid xl:grid-cols-3 gap-4 items-stretch">
            {reportConfig.charts.geoBars ? (
              <GlassCard title="Interactions par région canon" subtitle="Colonnes alignées Reporting Power BI">
                <div className="h-72">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart data={REGION_ORDER.map((rg) => ({ lib: REGION_SHORT[rg], v: kpis.appelsParRégion.get(rg) ?? 0 }))}>
                      <CartesianGrid strokeDasharray="3 3" stroke={palette.grid} vertical={false} />
                      <XAxis tick={{ fill: palette.muted, fontSize: 11 }} dataKey="lib" />
                      <YAxis tick={{ fill: palette.muted, fontSize: 11 }} />
                      {chartTooltip}
                      <Bar dataKey="v" radius={[10, 10, 4, 4]}>
                        {REGION_ORDER.map((rg) => (
                          <Cell key={rg} fill={REGION_COLORS[rg]} opacity={isDark ? 0.94 : 0.92} />
                        ))}
                      </Bar>
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </GlassCard>
            ) : null}
            {reportConfig.charts.geoDonut ? (
              <GlassCard title="Donut géographique" subtitle={REGION_ORDER.map((r) => REGION_SHORT[r]).join(" • ")}>
                <div className="h-72">
                  <ResponsiveContainer width="100%" height="100%">
                    <PieChart>
                      <Pie data={donutRég} dataKey="value" innerRadius={48} outerRadius={98} stroke="none">
                        {donutRég.map((d, idx) => (
                          <Cell key={idx} fill={d.fill} />
                        ))}
                      </Pie>
                      <Legend formatter={(value) => <span style={{ color: palette.fg }}>{value}</span>} />
                      {chartTooltip}
                    </PieChart>
                  </ResponsiveContainer>
                </div>
              </GlassCard>
            ) : null}
            {reportConfig.charts.statusPie ? (
              <GlassCard title="Répartition KPI statuts critiques" subtitle="Couleurs Résultat figées sur tout le cockpit">
                <div className="h-72">
                  <ResponsiveContainer width="100%" height="100%">
                    <PieChart>
                      <Pie data={résultatPieData} dataKey="value" cx="48%" outerRadius={90} stroke="none">
                        {résultatPieData.map((d) => (
                          <Cell key={d.name} fill={d.fill} />
                        ))}
                      </Pie>
                      <Legend formatter={(v) => <span style={{ color: palette.fg }}>{v}</span>} />
                      {chartTooltip}
                    </PieChart>
                  </ResponsiveContainer>
                </div>
              </GlassCard>
            ) : null}
          </div>

          {(reportConfig.charts.dailyArea || reportConfig.charts.monthlyBars) && (
            <div className="grid xl:grid-cols-2 gap-4">
              {reportConfig.charts.dailyArea ? (
                <GlassCard title="Courbe cumul jour / régions" subtitle="Stacks translucides harmonisées">
                  <div className="h-[340px]">
                    <ResponsiveContainer width="100%" height="100%">
                      <AreaChart data={sérieJour}>
                        <CartesianGrid stroke={palette.grid} strokeDasharray="4 8" vertical={false} />
                        <XAxis hide={sérieJour.length > 30} tick={{ fill: palette.muted, fontSize: 10 }} dataKey="jour" />
                        <YAxis tick={{ fill: palette.muted }} />
                        <Legend />
                        {chartTooltip}
                        {REGION_ORDER.map((rg) => (
                          <Area
                            key={rg}
                            type="monotone"
                            stackId="r"
                            dataKey={REGION_SHORT[rg]}
                            stroke={REGION_COLORS[rg]}
                            fill={`${REGION_COLORS[rg]}73`}
                          />
                        ))}
                      </AreaChart>
                    </ResponsiveContainer>
                  </div>
                </GlassCard>
              ) : null}

              {reportConfig.charts.monthlyBars ? (
                <GlassCard title="Barres empilées par mois" subtitle="Colonnes région canon + Mois métier métier hors colonne Dates">
                  <div className="h-[340px]">
                    <ResponsiveContainer width="100%" height="100%">
                      <BarChart data={sérieMois}>
                        <CartesianGrid stroke={palette.grid} strokeDasharray="4 8" vertical={false} />
                        <XAxis
                          dataKey="mois"
                          interval={0}
                          angle={-15}
                          textAnchor="end"
                          height={70}
                          tick={{ fontSize: 12 }}
                        />
                        {/* <XAxis hide={sérieMois.length > 22} tick={{ fill: palette.muted, fontSize: 10 }} dataKey="mois" /> */}
                        <YAxis tick={{ fill: palette.muted }} />
                        <Legend />
                        {chartTooltip}
                        {REGION_ORDER.map((rg) => (
                          <Bar key={rg} stackId="m" dataKey={REGION_SHORT[rg]} fill={REGION_COLORS[rg]} />
                        ))}
                      </BarChart>
                    </ResponsiveContainer>
                  </div>
                </GlassCard>
              ) : null}
            </div>
          )}

          {reportConfig.charts.trendLine ? (
            <GlassCard title="Tendance brute totale journée" subtitle="Ligne générale hors stack">
              <div className="h-64">
                <ResponsiveContainer width="100%" height="100%">
                  <AreaChart data={sérieJour}>
                    <CartesianGrid stroke={palette.grid} strokeDasharray="3 8" vertical={false} />
                    <XAxis dataKey="jour" tick={{ fill: palette.muted, fontSize: 10 }} hide={sérieJour.length > 28} />
                    <YAxis tick={{ fill: palette.muted }} />
                    {chartTooltip}
                    <Area
                      type="stepAfter"
                      dataKey="total"
                      stroke={AGGREGATE_VOLUME_BAR}
                      strokeWidth={2}
                      fillOpacity={0.22}
                      fill={AGGREGATE_VOLUME_BAR}
                    />
                  </AreaChart>
                </ResponsiveContainer>
              </div>
            </GlassCard>
          ) : null}

          {reportConfig.tables.pivotResult ? (
            <CrcRegionPivotWidget
              widgetId="pivotResult"
              title="1. Pivot Résultat × Régions"
              subtitle="Colonnes géographiques Drâa • Laâyoune • Souss • Faux appels · familles depuis le résultat normalisé."
              labelHeader="Résultat"
              rowLabelKey="name"
              rows={pivotResult}
              regionVisibility={columnVisibility.pivotRegions.pivotResult}
              onToggleRegion={(s) => togglePivotRegion("pivotResult", s)}
              isDark={isDark}
              palette={palette}
              exportBasename={exportFileBase}
              exportTableRef={pivotResultTableRef}
              exportChartRef={pivotResultChartRef}
            />
          ) : null}
            {reportConfig.tables.pivotMetier ? (
              <GlassCard
                title="2. Métier × Régions"
                subtitle="Métiers issus Page3 après normalisation d'étiquettes"
              >
                <div className="mb-4 flex items-center gap-3 flex-wrap">
                  <label className="text-xs font-semibold uppercase text-slate-500">
                    Résultat
                  </label>

                  <select
                    value={metierResultatFilter}
                    onChange={(e) =>
                      setMetierResultatFilter(e.target.value)
                    }
                    className="rounded-xl border border-slate-300 dark:border-slate-600 px-3 py-2 text-sm bg-white dark:bg-slate-900"
                  >
                    <option value="all">Tous</option>

                    {metierResultatOptions.map((r) => (
                      <option key={r} value={r}>
                        {r}
                      </option>
                    ))}
                  </select>

                  <span className="text-xs text-slate-500">
                    {pivotMet.length} métiers
                  </span>
                </div>

                <CrcRegionPivotWidget
                  widgetId="pivotMetier"
                  title=""
                  subtitle=""
                  labelHeader="Métier"
                  rowLabelKey="métier"
                  rows={pivotMet as any[]}
                  regionVisibility={columnVisibility.pivotRegions.pivotMetier}
                  onToggleRegion={(s) =>
                    togglePivotRegion("pivotMetier", s)
                  }
                  isDark={isDark}
                  palette={palette}
                  exportBasename={exportFileBase}
                  exportTableRef={pivotMetierTableRef}
                  exportChartRef={pivotMetierChartRef}
                />
              </GlassCard>
            ) : null}
          {reportConfig.tables.pivotNature ? (
            <CrcRegionPivotWidget
              widgetId="pivotNature"
              title="3. Nature de réclamation × Régions"
              subtitle={`${pivotNat.length} natures différentes · focus qualité dossiers`}
              labelHeader="Nature de Réclamation"
              rowLabelKey="nature"
              rows={pivotNat as any[]}
              regionVisibility={columnVisibility.pivotRegions.pivotNature}
              onToggleRegion={(s) => togglePivotRegion("pivotNature", s)}
              isDark={isDark}
              palette={palette}
              exportBasename={exportFileBase}
              exportTableRef={pivotNatureTableRef}
              exportChartRef={pivotNatureChartRef}
            />
          ) : null}

          {reportConfig.charts.teleopBars ? (
            <GlassCard title="Classement téléopérateurs" subtitle="Groupes lisibles volume / KPI qualité sans double comptabilisation graphique">
            <div className="h-[440px]">
              <ResponsiveContainer width="100%" height="100%">
                <BarChart layout="vertical" data={téléBar} margin={{ left: 28, right: 12 }}>
                  <CartesianGrid stroke={palette.grid} horizontal strokeDasharray="3 8" vertical={false} />
                  <XAxis tick={{ fill: palette.muted }} type="number" />
                  <YAxis type="category" dataKey="nom" tick={{ fill: palette.fg, fontSize: 11 }} width={128} />
                  {chartTooltip}
                  <Legend />
                  <Bar
                    dataKey="volume"
                    fill={AGGREGATE_VOLUME_BAR}
                    radius={[0, 4, 4, 0]}
                    maxBarSize={18}
                  />
<Bar
  dataKey="informés"
  fill={getResultColor("Clients informés")}
  name="Clients informés"
  maxBarSize={14}
/>

<Bar
  dataKey="tickets"
  fill={getResultColor("Tickets transmis")}
  name="Tickets transmis"
  maxBarSize={14}
/>

<Bar
  dataKey="abandons"
  fill={getResultColor("Appels abandonnés")}
  name="Appels abandonnés"
  maxBarSize={14}
/>

<Bar
  dataKey="appelsDécrochésInterrompus"
  fill={getResultColor("Appels décrochés interrompus")}
  name="Appels décrochés interrompus"
  maxBarSize={14}
/>
                </BarChart>
              </ResponsiveContainer>
            </div>
          </GlassCard>
          ) : null}

          {reportConfig.tables.teleOpStats ? (
            <CrcTeleopStatsWidget
              rows={téléopRanking}
              metricVisibility={columnVisibility.teleOpMetrics}
              onToggleMetric={toggleTeleOpMetric}
              isDark={isDark}
              palette={{ muted: palette.muted, grid: palette.grid, tooltipBg: palette.tooltipBg }}
              exportBasename={exportFileBase}
              exportTableRef={teleTableRef}
              exportChartRef={teleChartRef}
            />
          ) : null}

          {reportConfig.charts.regionCards ? (
            <div className="grid xl:grid-cols-3 gap-4">
              {REGION_ORDER.map((rg) => (
                <CrcRegionResultCardWidget
                  key={rg}
                  region={rg}
                  rows={filteredRows}
                  palette={palette}
                  exportBasename={exportFileBase}
                />
              ))}
            </div>
          ) : null}

          {reportConfig.tables.rawPreview ? (
            <CrcRawPreviewWidget
              rows={filteredRows}
              rawColumns={reportConfig.rawColumns}
              onToggleColumn={toggleRawColumn}
              columnCells={RAW_PREVIEW_COLUMNS}
              isDark={isDark}
              palette={palette}
              exportBasename={exportFileBase}
            />
          ) : null}
        </>
      )}

      {rows.length > 0 ? (
        <GlassCard
          title="Faux traitements"
          subtitle="Lignes avec Région vide ou Métier vide (investigation opérationnelle)."
          action={
            <div className="flex gap-2">
              <button
                type="button"
                className="rounded-full px-3 py-1 text-[11px] font-semibold border border-slate-300 dark:border-slate-600"
                onClick={() =>
                  exportInvestigationExcel(fauxTraitementsRows, INVESTIGATION_COLS, `${exportFileBase}_faux_traitements`)
                }
              >
                Excel
              </button>
              <button
                type="button"
                className="rounded-full px-3 py-1 text-[11px] font-semibold border border-slate-300 dark:border-slate-600"
                onClick={() =>
                  void exportInvestigationPdf(
                    "Faux traitements",
                    fauxTraitementsRows,
                    INVESTIGATION_COLS,
                    `${exportFileBase}_faux_traitements`,
                  )
                }
              >
                PDF
              </button>
              <button
                type="button"
                className="rounded-full px-3 py-1 text-[11px] font-semibold border border-slate-300 dark:border-slate-600"
                onClick={() =>
                  void exportInvestigationPptx(
                    "Faux traitements",
                    fauxTraitementsRows,
                    INVESTIGATION_COLS,
                    `${exportFileBase}_faux_traitements`,
                  )
                }
              >
                PPTX
              </button>
            </div>
          }
        >
          <div className="mb-4 flex items-center gap-3 flex-wrap">
            <label className="text-xs font-semibold uppercase text-slate-500">
              Résultat
            </label>

            <select
              value={fauxResultatFilter}
              onChange={(e) => {
                setFauxResultatFilter(e.target.value);
                setFauxPage(1);
              }}
              className="rounded-xl border border-slate-300 dark:border-slate-600 px-3 py-2 text-sm bg-white dark:bg-slate-900"
            >
              <option value="all">Tous</option>

              {fauxResultatOptions.map((r) => (
                <option key={r} value={r}>
                  {r}
                </option>
              ))}
            </select>

            <span className="text-xs text-slate-500">
              {fauxTraitementsRows.length} lignes
            </span>
          </div>
          <div className="overflow-auto rounded-2xl border border-slate-200 dark:border-slate-700">
            <table className="min-w-[980px] w-full text-xs">
              <thead className="bg-slate-900 text-white">
                <tr>
                  {["Date + Heure", "Téléopérateur", "Résultat", "Région", "Métier", "Nature", "Téléphone"].map((h) => (
                    <th key={h} className="px-2 py-2 text-left whitespace-nowrap">
                      {h}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {fauxPaged.map((r, i) => (
                  <tr key={`${r.rawIndex}-${i}`} className={i % 2 ? "bg-slate-50 dark:bg-slate-900/50" : ""}>
                    <td className="px-2 py-1">{formatDateTime(r.date)}</td>
                    <td className="px-2 py-1">{r.téléopérateur}</td>
                    <td className="px-2 py-1">{r.résultat}</td>
                    <td className="px-2 py-1">{REGION_SHORT[r.régionCanon]}</td>
                    <td className="px-2 py-1">{r.metier}</td>
                    <td className="px-2 py-1">{r.natureRéclamation}</td>
                    <td className="px-2 py-1 font-mono">{r.téléphone}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div className="mt-3 flex items-center gap-2">
            <button
              type="button"
              disabled={fauxPage <= 1}
              onClick={() => setFauxPage((p) => Math.max(1, p - 1))}
              className="rounded-full px-3 py-1 text-[11px] border border-slate-300 disabled:opacity-40"
            >
              Précédent
            </button>
            <span className="text-xs text-slate-500">
              Page {fauxPage} / {Math.max(1, Math.ceil(fauxTraitementsRows.length / fauxPerPage))}
            </span>
            <button
              type="button"
              disabled={fauxPage >= Math.ceil(Math.max(1, fauxTraitementsRows.length) / fauxPerPage)}
              onClick={() =>
                setFauxPage((p) => Math.min(Math.max(1, Math.ceil(fauxTraitementsRows.length / fauxPerPage)), p + 1))
              }
              className="rounded-full px-3 py-1 text-[11px] border border-slate-300 disabled:opacity-40"
            >
              Suivant
            </button>
          </div>
        </GlassCard>
      ) : null}

      {detailKpi ? (
        <div className="fixed inset-0 z-50 bg-slate-950/50 backdrop-blur-sm p-4 overflow-auto">
          <div className="max-w-[1400px] mx-auto mt-8 glass-panel p-5">
            <div className="flex items-center justify-between mb-4">
              <h3 className="text-lg font-semibold text-slate-900 dark:text-slate-50">
                Détail KPI — {detailKpi === "abandons" ? "Appels abandonnés" : detailKpi === "decroches" ? "Appels décrochés interrompus" : detailKpi === "informes" ? "Clients informés" : "Tickets transmis"}
              </h3>
              <button type="button" onClick={() => setDetailKpi(null)} className="rounded-full px-3 py-1 border border-slate-300">
                Fermer
              </button>
            </div>
            <div className="flex flex-wrap gap-3 mb-3 items-center">
              {REGION_ORDER.map((rg) => (
                <label key={rg} className="text-xs inline-flex items-center gap-1">
                  <input
                    type="checkbox"
                    checked={detailRegions[rg] !== false}
                    onChange={() => setDetailRegions((s) => ({ ...s, [rg]: !(s[rg] !== false) }))}
                  />
                  <span style={{ color: REGION_COLORS[rg] }}>{REGION_SHORT[rg]}</span>
                </label>
              ))}
              <button
                type="button"
                className="rounded-full px-3 py-1 text-[11px] border border-slate-300"
                onClick={() => exportInvestigationExcel(detailRows, INVESTIGATION_COLS, `${exportFileBase}_detail_${detailKpi}`)}
              >
                Export Excel
              </button>
              <button
                type="button"
                className="rounded-full px-3 py-1 text-[11px] border border-slate-300"
                onClick={() =>
                  void exportInvestigationPdf(
                    `Détail KPI ${detailKpi}`,
                    detailRows,
                    INVESTIGATION_COLS,
                    `${exportFileBase}_detail_${detailKpi}`,
                  )
                }
              >
                Export PDF
              </button>
              <button
                type="button"
                className="rounded-full px-3 py-1 text-[11px] border border-slate-300"
                onClick={() =>
                  void exportInvestigationPptx(
                    `Détail KPI ${detailKpi}`,
                    detailRows,
                    INVESTIGATION_COLS,
                    `${exportFileBase}_detail_${detailKpi}`,
                  )
                }
              >
                Export PPTX
              </button>
            </div>
            <div className="max-h-[65vh] overflow-auto rounded-2xl border border-slate-200 dark:border-slate-700">
              <table className="min-w-[1100px] w-full text-xs">
                <thead className="bg-slate-900 text-white sticky top-0">
                  <tr>
                    {["Date + Heure", "Téléopérateur", "Résultat brut Excel", "Métier", "Région", "Téléphone", "Nature de réclamation"].map((h) => (
                      <th key={h} className="px-2 py-2 text-left whitespace-nowrap">{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {detailRows.map((r, i) => (
                    <tr key={`${r.rawIndex}-${i}`} className={i % 2 ? "bg-slate-50 dark:bg-slate-900/50" : ""}>
                      <td className="px-2 py-1">{formatDateTime(r.date)}</td>
                      <td className="px-2 py-1">{r.téléopérateur}</td>
                      <td className="px-2 py-1">{r.résultat}</td>
                      <td className="px-2 py-1">{r.metier}</td>
                      <td className="px-2 py-1">{REGION_SHORT[r.régionCanon]}</td>
                      <td className="px-2 py-1 font-mono">{r.téléphone}</td>
                      <td className="px-2 py-1">{r.natureRéclamation}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      ) : null}

      {debugOpen && (
        <GlassCard title="Mode développeur — Parsing Axilus">
          {!debug ? (
            <p className="text-sm text-slate-500">Aucun jeu chargé.</p>
          ) : (
            <div className="space-y-4 text-xs font-mono">
              <div>
                <p className="font-bold uppercase text-slate-500 mb-2">Feuilles</p>
                <p>{debug.sheets.join(" | ")}</p>
              </div>
              <div>
                <p className="font-bold uppercase text-slate-500 mb-2">En-têtes détectés</p>
                <div className="flex flex-wrap gap-1">{debug.detectedHeaders.map((h) => (
                  <span key={h} className="px-2 py-0.5 rounded bg-sky-500/15 text-[10px] border border-sky-600/40">
                    {h}
                  </span>
                ))}</div>
              </div>
              <div>
                <p className="font-bold uppercase text-slate-500 mb-2">Mappings normalisés</p>
                <pre className="max-h-40 overflow-auto p-4 rounded-2xl bg-slate-950 text-green-400 text-[11px]">
                  {debug.normalizedHeaders
                    .map((m) => `"${m.original}" → [${String(m.normalizedKey)}] label UI: "${displayLabel(m.original)}"`)
                    .join("\n")}
                </pre>
              </div>
              <div>
                <p className="font-bold uppercase text-slate-500 mb-2">Échantillon JSON (lignes 1-5)</p>
                <pre className="max-h-64 overflow-auto p-4 rounded-2xl bg-slate-900 text-yellow-400 text-[10px]">
                  {JSON.stringify(
                    rows.slice(0, 5).map((r) => ({
                      rawIndex: r.rawIndex,
                      valid: r.valid,
                      motif: r.validationReason ?? null,
                      date: r.date ? r.date.toISOString() : null,
                      résultat: r.résultat,
                      résultatRaw: r.résultatRaw,
                      régionCanon: r.régionCanon,
                      téléopérateur: r.téléopérateur,
                    })),
                    null,
                    2,
                  )}
                </pre>
              </div>

              <div className="grid sm:grid-cols-4 gap-2">
                {[
                  ["Brutes parsées", debug.parsedRows],
                  ["Valides métier/date", debug.validRows],
                  ["À contrôler", debug.invalidRows],
                  ["Somme contrôle", `${debug.validRows + debug.invalidRows}`],
                ].map(([label, count]) => (
                  <div key={String(label)} className="p-4 rounded-xl border border-white/60 dark:border-white/10 bg-white/40 dark:bg-slate-900/55">
                    <p className="text-[11px] text-slate-500 uppercase">{label}</p>
                    <p className="text-xl font-semibold">{String(count)}</p>
                  </div>
                ))}
              </div>
              <div>
                <p className="font-bold uppercase text-slate-500 mb-2">Parsing logs</p>
                <pre className="max-h-60 overflow-auto p-4 rounded-2xl bg-slate-50 dark:bg-black/55 text-emerald-800 dark:text-emerald-400">
                  {debug.logs.join("\n")}
                </pre>
              </div>
            </div>
          )}
        </GlassCard>
      )}
    </div>
  );
}
