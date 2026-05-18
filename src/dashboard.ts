// =============================================================================
// dashboard.ts — Premium Bento Dashboard: ECharts + Live Async Data Binding
// Tree-shakeable imports for Vite / GitHub Pages bundle.
// =============================================================================

import * as echarts from 'echarts/core';
import { BarChart, PieChart, RadarChart, TreemapChart, EffectScatterChart, FunnelChart, SunburstChart, SankeyChart, LineChart } from 'echarts/charts';
import {
  GridComponent,
  TooltipComponent,
  LegendComponent,
  TitleComponent,
  VisualMapComponent,
  GraphicComponent,
  DataZoomComponent,
  RadarComponent,
} from 'echarts/components';
import { CanvasRenderer } from 'echarts/renderers';
import type { ComposeOption } from 'echarts/core';
import type {
  BarSeriesOption,
  PieSeriesOption,
  RadarSeriesOption,
  TreemapSeriesOption,
  EffectScatterSeriesOption,
  FunnelSeriesOption,
  SunburstSeriesOption,
  SankeySeriesOption,
  LineSeriesOption,
} from 'echarts/charts';
import type {
  GridComponentOption,
  TooltipComponentOption,
  LegendComponentOption,
  TitleComponentOption,
  VisualMapComponentOption,
  GraphicComponentOption,
  DataZoomComponentOption,
  RadarComponentOption,
} from 'echarts/components';

echarts.use([
  BarChart,
  PieChart,
  RadarChart,
  TreemapChart,
  EffectScatterChart,
  FunnelChart,
  SunburstChart,
  SankeyChart,
  LineChart,
  GridComponent,
  TooltipComponent,
  LegendComponent,
  TitleComponent,
  VisualMapComponent,
  GraphicComponent,
  DataZoomComponent,
  RadarComponent,
  CanvasRenderer,
]);

type ECOption = ComposeOption<
  | BarSeriesOption
  | PieSeriesOption
  | RadarSeriesOption
  | SunburstSeriesOption
  | TreemapSeriesOption
  | EffectScatterSeriesOption
  | FunnelSeriesOption
  | SankeySeriesOption
  | LineSeriesOption
  | GridComponentOption
  | TooltipComponentOption
  | LegendComponentOption
  | TitleComponentOption
  | VisualMapComponentOption
  | GraphicComponentOption
  | DataZoomComponentOption
  | RadarComponentOption
>;

// =============================================================================
// Design tokens — resolved at runtime so charts honour CSS variable overrides.
// =============================================================================
function cssVar(name: string): string {
  return getComputedStyle(document.documentElement).getPropertyValue(name).trim();
}

function palette(): string[] {
  return [
    cssVar('--info')    || '#3b82f6',
    cssVar('--accent')  || '#6366f1',
    cssVar('--success') || '#22c55e',
    cssVar('--danger')  || '#ef4444',
    cssVar('--purple')  || '#a855f7',
    cssVar('--warning') || '#f59e0b',
  ];
}

// =============================================================================
// Shared tooltip extraCssText (glassmorphism)
// =============================================================================
const GLASS_TOOLTIP =
  'backdrop-filter: blur(12px); ' +
  'background-color: rgba(255, 255, 255, 0.65); ' +
  'border-radius: 12px; ' +
  'border: 1px solid rgba(255, 255, 255, 0.4); ' +
  'box-shadow: 0 8px 16px rgba(0,0,0,0.08);';

// =============================================================================
// Mount chart instances — deferred until DOM is ready so containers have
// their CSS-driven dimensions before ECharts measures them.
// =============================================================================
const CHART_BY_DOM_ID: Record<string, echarts.ECharts> = {};
let dashboardResizeObserver: ResizeObserver | null = null;

function ensureDashboardResizeObserver(): void {
  if (dashboardResizeObserver) return;
  if (typeof ResizeObserver === 'undefined') {
    console.warn('[dashboard] ResizeObserver unavailable; using window resize fallback only.');
    return;
  }
  dashboardResizeObserver = new ResizeObserver((entries) => {
    for (const entry of entries) {
      const id = entry.target && (entry.target as HTMLElement).id;
      if (!id) continue;
      const chart = CHART_BY_DOM_ID[id];
      if (!chart) continue;
      try { chart.resize(); } catch { /* noop */ }
    }
  });
}

function mountChart(id: string): echarts.ECharts | null {
  const el = document.getElementById(id);
  if (!el) {
    console.warn(`[dashboard] mount point #${id} not found`);
    return null;
  }
  
  // If the element is hidden (e.g. inside a collapsed section), clientWidth/Height will be 0.
  // We can temporarily ensure it has dimensions, or just rely on CSS and resize() later.
  // ECharts warns if it initializes on a 0x0 element. 
  // Let's provide an explicit width/height to init if clientWidth/Height is 0.
  const width = el.clientWidth || 400;
  const height = el.clientHeight || parseInt(getComputedStyle(el).minHeight, 10) || 300;
  
  // Force the element to have a real height before init so ECharts doesn't
  // measure a 0-px container and produce a squished/invisible chart.
  if (!el.offsetHeight) {
    el.style.height = `${height}px`;
  }

  const existing = CHART_BY_DOM_ID[id];
  if (existing) {
    try { existing.dispose(); } catch { /* noop */ }
  }
  const chart = echarts.init(el, undefined, { renderer: 'canvas', width, height });
  CHART_BY_DOM_ID[id] = chart;
  ensureDashboardResizeObserver();
  if (dashboardResizeObserver) {
    try { dashboardResizeObserver.observe(el); } catch { /* noop */ }
  }
  return chart;
}

// Chart instances — populated inside initDashboardCharts() after DOM ready.
let chartOccupancy:   echarts.ECharts | null = null;
let chartMoveOuts:    echarts.ECharts | null = null;
let chartPortfolio:   echarts.ECharts | null = null;
let chartVelocity:    echarts.ECharts | null = null;
let chartPmBar:       echarts.ECharts | null = null;
let chartStatusDonut: echarts.ECharts | null = null;
let chartPmLoad:      echarts.ECharts | null = null;
let chartWoType:      echarts.ECharts | null = null;
let chartUrgency:     echarts.ECharts | null = null;
let ALL_CHARTS:       echarts.ECharts[]      = [];

// =============================================================================
// Shared axis / grid style helpers
// =============================================================================

/** Subtle dashed split lines — keeps grid clean without visual noise. */
function yAxisSplitLine() {
  return {
    splitLine: {
      lineStyle: { type: 'dashed' as const, color: 'rgba(0,0,0,0.06)' },
    },
  };
}

function xAxisNoSplitLine() {
  return { splitLine: { show: false } };
}

/** Shared axis label style used across all bar charts. */
function axisLabelStyle() {
  return {
    color: 'rgba(0,0,0,0.45)',
    fontSize: 11,
    fontFamily: 'var(--font-mono, "Inconsolata", monospace)',
  };
}

// =============================================================================
// Gradient helpers
// =============================================================================

/** Top-to-bottom LinearGradient: full color at top fading to ~18% opacity. */
function barGradient(hex: string): echarts.graphic.LinearGradient {
  return new echarts.graphic.LinearGradient(0, 0, 0, 1, [
    { offset: 0,   color: hex },
    { offset: 0.6, color: hex + '99' }, // mid-fade ~60% opacity
    { offset: 1,   color: hex + '18' }, // near-transparent bottom
  ]);
}

/** Radial glow gradient — used as bar background/shadow fill on hover. */
function barRadialGlow(hex: string): echarts.graphic.RadialGradient {
  return new echarts.graphic.RadialGradient(0.5, 0.5, 1, [
    { offset: 0,   color: hex + 'cc' },
    { offset: 1,   color: hex + '00' },
  ]);
}

/** Shared itemStyle for every bar series — gradient fill + colored shadow. */
function barItemStyle(color: string): object {
  return {
    color: barGradient(color),
    borderRadius: [6, 6, 0, 0] as [number, number, number, number],
    shadowBlur: 12,
    shadowColor: color + '66',   // ~40% opacity glow
    shadowOffsetY: 4,
  };
}

/** Emphasis itemStyle: glow intensifies on hover. */
function barEmphasisStyle(color: string): object {
  return {
    color: barGradient(color),
    shadowBlur: 24,
    shadowColor: color + 'aa',   // ~67% opacity deep glow
    shadowOffsetY: 6,
  };
}

/** Glowing line style for radar/line series. */
function glowLineStyle(color: string, width = 2): object {
  return {
    color,
    width,
    shadowBlur: 10,
    shadowColor: color + 'cc',
  };
}

// =============================================================================
// Tooltip formatter helpers
// =============================================================================

/** Bar tooltip: colored dot + label + count with "WOs" suffix. */
function barTooltipFormatter(dotColor: string) {
  return (params: { name: string; value: number | string; seriesName?: string }) => {
    const val = Number(params.value ?? 0);
    return (
      `<div style="display:flex;align-items:center;gap:7px;padding:2px 0">` +
        `<span style="display:inline-block;width:8px;height:8px;border-radius:50%;` +
          `background:${dotColor};box-shadow:0 0 0 2px ${dotColor}33;flex-shrink:0"></span>` +
        `<span style="font-weight:600;color:#19202f">${params.name}</span>` +
      `</div>` +
      `<div style="margin-top:4px;font-size:13px;font-weight:700;` +
        `letter-spacing:-0.3px;color:#19202f">` +
        `${val.toLocaleString()} <span style="font-size:10px;font-weight:500;` +
          `color:rgba(0,0,0,0.4);letter-spacing:0">WOs</span>` +
      `</div>`
    );
  };
}

/** Donut tooltip: colored dot + segment name + count + percentage. */
function donutTooltipFormatter() {
  return (params: { name: string; value: number; percent: number; color: string }) => {
    return (
      `<div style="display:flex;align-items:center;gap:7px;padding:2px 0">` +
        `<span style="display:inline-block;width:8px;height:8px;border-radius:50%;` +
          `background:${params.color};box-shadow:0 0 0 2px ${params.color}33;flex-shrink:0"></span>` +
        `<span style="font-weight:600;color:#19202f">${params.name}</span>` +
      `</div>` +
      `<div style="margin-top:4px;font-size:13px;font-weight:700;letter-spacing:-0.3px;color:#19202f">` +
        `${params.value.toLocaleString()} ` +
        `<span style="font-size:10px;font-weight:500;color:rgba(0,0,0,0.4)">WOs</span>` +
        `<span style="margin-left:8px;font-size:11px;font-weight:600;` +
          `color:${params.color}">${params.percent.toFixed(1)}%</span>` +
      `</div>`
    );
  };
}

// =============================================================================
// Central donut graphic overlay (total count in the hole)
// =============================================================================
function donutCenterGraphic(total: number, centerX: string, label: string): object[] {
  return [
    {
      type: 'text',
      left: centerX,
      top: 'center',
      style: {
        text: total.toLocaleString(),
        textAlign: 'center',
        fill: '#19202f',
        fontSize: 22,
        fontWeight: 700,
        fontFamily: 'var(--font-display, "Bricolage Grotesque", sans-serif)',
        lineHeight: 26,
      },
    },
    {
      type: 'text',
      left: centerX,
      top: '58%',
      style: {
        text: label,
        textAlign: 'center',
        fill: 'rgba(0,0,0,0.4)',
        fontSize: 10,
        fontWeight: 500,
        fontFamily: 'var(--font-mono, "Inconsolata", monospace)',
        textTransform: 'uppercase',
        letterSpacing: 0.8,
      },
    },
  ];
}

// =============================================================================
// Skeleton / placeholder options (shown while loading)
// =============================================================================
function makePlaceholderBar(): ECOption {
  return {
    backgroundColor: 'transparent',
    tooltip: { show: false },
    grid: { left: '3%', right: '4%', bottom: '3%', containLabel: true },
    xAxis: {
      type: 'category',
      data: [],
      axisTick: { show: false },
      axisLine: { show: false },
      ...xAxisNoSplitLine(),
    },
    yAxis: { type: 'value', ...yAxisSplitLine() },
    series: [{
      type: 'bar',
      data: [],
      itemStyle: { color: 'rgba(0,0,0,0.06)', borderRadius: [6, 6, 0, 0] },
    }],
  };
}

function makePlaceholderDonut(): ECOption {
  return {
    backgroundColor: 'transparent',
    tooltip: { show: false },
    legend: { show: false },
    series: [{
      type: 'pie',
      radius: ['48%', '72%'],
      data: [],
      itemStyle: { color: 'rgba(0,0,0,0.06)', borderRadius: 8 },
      label: { show: false },
      labelLine: { show: false },
    }],
  };
}

// =============================================================================
// Resize handler (registered once; ALL_CHARTS populated by initDashboardCharts)
// =============================================================================
window.addEventListener('resize', () => {
  ALL_CHARTS.forEach(c => c.resize());
});

// Resolved at runtime from localStorage (same key app.js uses).
function resolveProxyUrl(): string {
  const defaultProxy = 'https://afproxy.val.run';
  try {
    return (localStorage.getItem('hm_proxy_url') || '').trim() || defaultProxy;
  } catch {
    return defaultProxy;
  }
}

function buildWorkOrdersUrl(): string {
  const proxy = resolveProxyUrl();
  if (proxy) {
    const sep = proxy.includes('?') ? '&' : '?';
    return `${proxy}${sep}action=work_orders&days=180`;
  }
  return 'https://afproxy.val.run/?action=work_orders&days=180';
}

// =============================================================================
// Auth header helper (mirrors app.js getProxyAccessToken logic)
// =============================================================================
function proxyAuthHeaders(): Record<string, string> {
  try {
    const token =
      localStorage.getItem('hm_access_token') ||
      sessionStorage.getItem('hm_access_token') ||
      localStorage.getItem('hm_auth_token') ||
      localStorage.getItem('hm_device_token') ||
      localStorage.getItem('hm_proxy_token') ||
      '';
    if (token) return { Authorization: `Bearer ${token}` };
  } catch { /* non-fatal */ }
  return {};
}

// =============================================================================
// Data aggregation helpers
// =============================================================================
interface WoRecord {
  // Raw proxy field names (direct fetch)
  pm_name?: string;
  property_manager?: string;
  PropertyManager?: string;
  status?: string;
  Status?: string;
  category?: string;
  Category?: string;
  type?: string;
  Type?: string;
  priority?: string;
  Priority?: string;
  vendor_name?: string;
  vendor?: string;
  Vendor?: string;
  created_at?: string;
  CreatedAt?: string;
  created?: string;
  date_created?: string;
  // Normalized shape (from app.js WORK_ORDERS array)
  propertyManager?: string;    // normalized pm
  assignedUser?: string;       // normalized assigned user (maps to raw assigned_user)
  vendorName?: string;         // normalized vendor name
  vendorId?: string;
  description?: string;
  propertyName?: string;
  unitId?: string;
  // Shared
  [key: string]: unknown;
}

function pmField(r: WoRecord): string {
  return String(
    r.pm_name || r.property_manager || r.propertyManager || r.PropertyManager || 'Unassigned'
  ).trim() || 'Unassigned';
}

function statusField(r: WoRecord): string {
  return String(r.status || r.Status || 'Unknown').trim() || 'Unknown';
}

function categoryField(r: WoRecord): string {
  // 'type' in normalized WOs is work_order_type
  return String(r.category || r.Category || r.type || r.Type || 'Other').trim() || 'Other';
}

function priorityField(r: WoRecord): string {
  return String(r.priority || r.Priority || 'Normal').trim() || 'Normal';
}

function vendorField(r: WoRecord): string {
  return String(r.vendor_name || r.vendorName || r.vendor || r.Vendor || '').trim();
}

/** Generic group-by counter returning { label → count } sorted descending by count. */
function groupCount(
  rows: WoRecord[],
  keyFn: (r: WoRecord) => string,
  topN = 20,
): { names: string[]; counts: number[] } {
  const map: Record<string, number> = rows.reduce((acc, r) => {
    const k = keyFn(r);
    acc[k] = (acc[k] || 0) + 1;
    return acc;
  }, {} as Record<string, number>);

  const sorted = Object.entries(map)
    .sort((a, b) => b[1] - a[1])
    .slice(0, topN);

  return {
    names:  sorted.map(([k]) => k),
    counts: sorted.map(([, v]) => v),
  };
}

/** Count completed/closed WOs per PM — case-insensitive partial match so
 *  AppFolio status strings like "Completed", "completed", "Closed", etc. all hit. */
function isClosedStatus(s: string): boolean {
  const lc = s.toLowerCase();
  return lc.includes('complet') || lc.includes('closed') || lc.includes('done') || lc.includes('resolved');
}

function pmClosedCounts(rows: WoRecord[]): { names: string[]; counts: number[] } {
  const closed = rows.filter(r => isClosedStatus(statusField(r)));
  // Fall back to ALL rows if zero closed found — avoids an empty bar chart.
  return groupCount(closed.length ? closed : rows, pmField, 15);
}

/** Map grouped counts to ECharts pie data array. */
function toPieData(
  rows: WoRecord[],
  keyFn: (r: WoRecord) => string,
): Array<{ name: string; value: number }> {
  const map: Record<string, number> = rows.reduce((acc, r) => {
    const k = keyFn(r);
    acc[k] = (acc[k] || 0) + 1;
    return acc;
  }, {} as Record<string, number>);

  return Object.entries(map)
    .sort((a, b) => b[1] - a[1])
    .map(([name, value]) => ({ name, value }));
}

// =============================================================================
// Chart option builders — Advanced Analytics edition
// =============================================================================

// ── Shared axis config used by bar builders ───────────────────────────────────
function premiumXAxis(names: string[]): object {
  return {
    type: 'category',
    data: names,
    axisTick: { show: false },
    axisLine: { show: false },
    axisLabel: {
      ...axisLabelStyle(),
      rotate: names.length > 5 ? 30 : 0,
      overflow: 'truncate',
      width: 80,
    },
    ...xAxisNoSplitLine(),
  };
}

function premiumYAxis(name: string): object {
  return {
    type: 'value',
    name,
    nameTextStyle: { fontSize: 10, color: 'rgba(0,0,0,0.4)', fontFamily: axisLabelStyle().fontFamily },
    axisLabel: { ...axisLabelStyle() },
    ...yAxisSplitLine(),
  };
}

// ── PM Radar — multi-axis PM comparison from real WO data ────────────────────
//
// Four axes, all derived from the work_orders endpoint:
//   Open WOs     — volume of unresolved work (lower = better throughput)
//   Closed WOs   — resolved work in the window (higher = more productive)
//   Urgent WOs   — pressure / fire-fighting load (lower = healthier portfolio)
//   Unassigned   — WOs not yet routed to a vendor (lower = faster dispatch)
//
// `max` values are computed from the actual data so the shape is always
// proportional — no axis ever shows 0/N with a hardcoded maximum.

interface PmRadarRow {
  name: string;
  open: number;
  closed: number;
  urgent: number;
  unassigned: number;
}

/** Build per-PM metric rows from the raw WO results array. */
function buildPmRadarRows(results: WoRecord[]): PmRadarRow[] {
  const map: Record<string, PmRadarRow> = {};
  for (const r of results) {
    const pm = pmField(r);
    if (!map[pm]) map[pm] = { name: pm, open: 0, closed: 0, urgent: 0, unassigned: 0 };
    const row = map[pm];
    const closed = isClosedStatus(statusField(r));
    if (closed) {
      row.closed++;
    } else {
      row.open++;
      const pri = priorityField(r).toLowerCase();
      if (pri === 'urgent' || pri === 'emergency') row.urgent++;
      if (!vendorField(r)) row.unassigned++;
    }
  }
  // Sort by total WO volume desc, cap at top 8 PMs so the radar stays readable
  return Object.values(map)
    .sort((a, b) => (b.open + b.closed) - (a.open + a.closed))
    .slice(0, 8);
}

function buildPmRadarOption(results: WoRecord[]): ECOption {
  const colors = palette();
  const rows = buildPmRadarRows(results);
  if (!rows.length) return makePlaceholderBar();

  // Dynamic maximums — use the highest value across all PMs, floored at 1.
  const maxOpen       = Math.max(1, ...rows.map(r => r.open));
  const maxClosed     = Math.max(1, ...rows.map(r => r.closed));
  const maxUrgent     = Math.max(1, ...rows.map(r => r.urgent));
  const maxUnassigned = Math.max(1, ...rows.map(r => r.unassigned));

  // Assign each PM a color from the palette, cycling if > 6 PMs
  const seriesColors = rows.map((_, i) => colors[i % colors.length]);

  const radarAxisLabel = {
    color: 'rgba(0,0,0,0.45)',
    fontSize: 10,
    fontFamily: axisLabelStyle().fontFamily,
  };

  return {
    backgroundColor: 'transparent',
    color: seriesColors,
    tooltip: {
      trigger: 'item',
      extraCssText: GLASS_TOOLTIP,
      formatter: (p: unknown) => {
        const params = p as { name: string; value: number[]; color: string };
        const [open, closed, urgent, unassigned] = params.value;
        const dotColor = params.color;
        return (
          `<div style="display:flex;align-items:center;gap:7px;padding:2px 0 4px">` +
            `<span style="display:inline-block;width:8px;height:8px;border-radius:50%;` +
              `background:${dotColor};box-shadow:0 0 0 2px ${dotColor}33;flex-shrink:0"></span>` +
            `<span style="font-weight:700;color:#19202f;font-size:12px">${params.name}</span>` +
          `</div>` +
          `<table style="font-size:11px;border-collapse:collapse;width:100%">` +
            `<tr><td style="color:rgba(0,0,0,0.45);padding:1px 8px 1px 0">Open WOs</td><td style="font-weight:600;color:#19202f;text-align:right">${open}</td></tr>` +
            `<tr><td style="color:rgba(0,0,0,0.45);padding:1px 8px 1px 0">Closed WOs</td><td style="font-weight:600;color:#22c55e;text-align:right">${closed}</td></tr>` +
            `<tr><td style="color:rgba(0,0,0,0.45);padding:1px 8px 1px 0">Urgent</td><td style="font-weight:600;color:#ef4444;text-align:right">${urgent}</td></tr>` +
            `<tr><td style="color:rgba(0,0,0,0.45);padding:1px 8px 1px 0">Unassigned</td><td style="font-weight:600;color:#f59e0b;text-align:right">${unassigned}</td></tr>` +
          `</table>`
        );
      },
    },
    legend: {
      bottom: 0,
      textStyle: { fontSize: 10, color: 'rgba(0,0,0,0.5)', fontFamily: axisLabelStyle().fontFamily },
      itemWidth: 8,
      itemHeight: 8,
      icon: 'circle',
      // Only show legend if ≤ 5 PMs — beyond that it becomes a scrolling wall of text
      show: rows.length <= 5,
    },
    radar: {
      shape: 'circle',
      center: ['50%', '48%'],
      radius: rows.length <= 4 ? '62%' : '55%',
      splitNumber: 4,
      indicator: [
        { name: 'Open WOs',   max: maxOpen },
        { name: 'Closed',     max: maxClosed },
        { name: 'Urgent',     max: maxUrgent },
        { name: 'Unassigned', max: maxUnassigned },
      ],
      axisName: { ...radarAxisLabel, padding: [0, 4] },
      splitLine: { lineStyle: { color: 'rgba(0,0,0,0.06)', type: 'dashed' } },
      splitArea: { show: true, areaStyle: {
        color: ['rgba(0,0,0,0.015)', 'rgba(0,0,0,0.03)'],
      }},
      axisLine: { lineStyle: { color: 'rgba(0,0,0,0.1)' } },
    } as RadarComponentOption,
    series: [
      {
        type: 'radar',
        data: rows.map((row, i) => ({
          name: row.name,
          value: [row.open, row.closed, row.urgent, row.unassigned],
          symbol: 'circle',
          symbolSize: 6,
          lineStyle: glowLineStyle(seriesColors[i % seriesColors.length], 2),
          areaStyle: {
            color: new echarts.graphic.RadialGradient(0.5, 0.5, 1, [
              { offset: 0,   color: seriesColors[i % seriesColors.length] + '55' },
              { offset: 1,   color: seriesColors[i % seriesColors.length] + '11' },
            ]),
          },
          itemStyle: {
            color: seriesColors[i % seriesColors.length],
            shadowBlur: 6,
            shadowColor: seriesColors[i % seriesColors.length] + 'cc',
          },
        })),
        emphasis: {
          lineStyle: { width: 3.5, shadowBlur: 16 },
          areaStyle: { opacity: 0.45 },
        },
      },
    ] as RadarSeriesOption[],
  };
}

// ── WO Type Treemap — category × urgency hierarchy ────────────────────────────
//
// Structure: each leaf = one category.  Color is driven by the proportion of
// urgent WOs inside that category (high urgency → warm red, low → cool indigo),
// giving the viewer an immediate "where is the fire?" read.

interface TreemapLeaf { name: string; value: number; urgentCount: number }

function buildWoTypeTreemapOption(results: WoRecord[]): ECOption {
  const colors = palette();
  // Build per-category totals and urgent sub-counts
  const catMap: Record<string, TreemapLeaf> = {};
  for (const r of results) {
    if (isClosedStatus(statusField(r))) continue;   // open WOs only
    const cat = categoryField(r);
    if (!catMap[cat]) catMap[cat] = { name: cat, value: 0, urgentCount: 0 };
    catMap[cat].value++;
    const pri = String(r.priority || r.Priority || '').toLowerCase();
    if (pri === 'urgent' || pri === 'emergency') catMap[cat].urgentCount++;
  }

  const leaves = Object.values(catMap).sort((a, b) => b.value - a.value);
  if (!leaves.length) return makePlaceholderDonut();

  // Map urgency ratio to a color: 0% urgent → indigo, 100% urgent → red
  function urgencyColor(leaf: TreemapLeaf): string {
    const ratio = leaf.value > 0 ? leaf.urgentCount / leaf.value : 0;
    // Lerp between info-blue and danger-red in hex space
    if (ratio >= 0.5)  return colors[3]; // --danger
    if (ratio >= 0.25) return colors[5]; // --warning
    if (ratio > 0)     return colors[1]; // --accent
    return colors[0];                    // --info (no urgent)
  }

  const treemapData = leaves.map(leaf => {
    const base = urgencyColor(leaf);
    return {
      name: leaf.name,
      value: leaf.value,
      itemStyle: {
        color: base + 'dd',           // ~87% opacity fill
        borderColor: base + '44',     // subtle same-hue border
        borderWidth: 1,
        shadowBlur: 8,
        shadowColor: base + '44',
        shadowOffsetY: 2,
      },
      label: { show: leaf.value >= 2 },
    };
  });

  return {
    backgroundColor: 'transparent',
    tooltip: {
      trigger: 'item',
      extraCssText: GLASS_TOOLTIP,
      formatter: (p: unknown) => {
        const params = p as { name: string; value: number; color: string };
        const leaf = catMap[params.name];
        const urgentPct = leaf && leaf.value > 0
          ? ((leaf.urgentCount / leaf.value) * 100).toFixed(0)
          : '0';
        return (
          `<div style="display:flex;align-items:center;gap:7px;padding:2px 0">` +
            `<span style="display:inline-block;width:8px;height:8px;border-radius:2px;` +
              `background:${params.color};flex-shrink:0"></span>` +
            `<span style="font-weight:700;color:#19202f">${params.name}</span>` +
          `</div>` +
          `<div style="margin-top:4px;font-size:13px;font-weight:700;color:#19202f">` +
            `${params.value.toLocaleString()} ` +
            `<span style="font-size:10px;font-weight:500;color:rgba(0,0,0,0.4)">open WOs</span>` +
          `</div>` +
          (leaf?.urgentCount
            ? `<div style="margin-top:2px;font-size:11px;color:#ef4444;font-weight:600">${urgentPct}% urgent</div>`
            : '')
        );
      },
    },
    series: [
      {
        type: 'treemap',
        data: treemapData,
        roam: false,
        leafDepth: 1,
        width: '100%',
        height: '100%',
        breadcrumb: { show: false },
        label: {
          show: true,
          position: 'insideTopLeft' as const,
          formatter: (p: { name: string; value: number }) =>
            `{name|${p.name}}\n{val|${p.value} WOs}`,
          rich: {
            name: {
              fontSize: 11,
              fontWeight: 600,
              color: 'rgba(255,255,255,0.92)',
              fontFamily: axisLabelStyle().fontFamily,
              lineHeight: 16,
            },
            val: {
              fontSize: 10,
              color: 'rgba(255,255,255,0.65)',
              fontFamily: axisLabelStyle().fontFamily,
            },
          },
        },
        upperLabel: { show: false },
        levels: [
          {
            itemStyle: {
              borderColor: 'rgba(255,255,255,0.2)',
              borderWidth: 2,
              gapWidth: 3,
            },
          },
        ],
        emphasis: {
          itemStyle: {
            shadowBlur: 24,
            shadowColor: 'rgba(0,0,0,0.35)',
            borderColor: 'rgba(255,255,255,0.8)',
            borderWidth: 2,
          },
        },
      },
    ] as TreemapSeriesOption[],
  };
}

// ── PM Load bar (open WOs per PM) + visual map + DataZoom ────────────────────
function buildPmLoadBarOption(names: string[], counts: number[]): ECOption {
  const colors = palette();
  const baseColor = colors[0];
  // Show 8 PMs at a time; if there are more the slider reveals the rest
  const windowPct = names.length > 0 ? Math.min(100, Math.round((8 / names.length) * 100)) : 100;

  return {
    backgroundColor: 'transparent',
    tooltip: {
      trigger: 'axis',
      axisPointer: { type: 'none' },
      extraCssText: GLASS_TOOLTIP,
      formatter: (p: unknown) => {
        const params = (p as { name: string; value: number }[])[0];
        return barTooltipFormatter(baseColor)({ name: params.name, value: params.value });
      },
    },
    // DataZoom: slider at bottom + mouse-wheel scroll inside the chart
    dataZoom: names.length > 8 ? [
      {
        type: 'slider',
        show: true,
        xAxisIndex: 0,
        start: 0,
        end: windowPct,
        height: 18,
        bottom: 0,
        borderColor: 'transparent',
        backgroundColor: 'rgba(0,0,0,0.04)',
        fillerColor: baseColor + '22',
        handleStyle: { color: baseColor, borderColor: baseColor },
        textStyle: { color: 'rgba(0,0,0,0.35)', fontSize: 10 },
        brushSelect: false,
      },
      { type: 'inside', xAxisIndex: 0, start: 0, end: windowPct },
    ] as DataZoomComponentOption[] : undefined,
    grid: {
      left: '3%',
      right: '4%',
      bottom: names.length > 8 ? '14%' : '10%',
      top: '8%',
      containLabel: true,
    },
    xAxis: premiumXAxis(names),
    yAxis: premiumYAxis('Open WOs'),
    series: [
      {
        type: 'bar',
        data: counts.map((v, i) => {
          const c = v >= 16 ? colors[3] : v <= 4 ? colors[2] : baseColor;
          return {
            value: v,
            itemStyle: barItemStyle(c) as BarSeriesOption['itemStyle'],
          };
        }),
        barMaxWidth: 48,
        emphasis: {
          focus: 'self',
          itemStyle: { shadowBlur: 24, shadowOffsetY: 6 } as BarSeriesOption['itemStyle'],
        },
      },
    ],
  };
}

// ── Open WOs by Age (horizontal bar, oldest → most urgent) ───────────────────
//
// Buckets: 0-7d (green/fresh), 8-14d, 15-30d, 31-60d, 60+d (red/critical)
// Bars sorted oldest-first so the most attention-needing items sit at the top.
//
interface AgeBucket { label: string; min: number; max: number; color: string; }

const WO_AGE_BUCKETS: AgeBucket[] = [
  { label: '60+ days',  min: 60,  max: Infinity, color: '#ef4444' },  // critical
  { label: '31–60 days', min: 31, max: 59,        color: '#f97316' },  // warning-high
  { label: '15–30 days', min: 15, max: 30,        color: '#f59e0b' },  // warning
  { label: '8–14 days',  min: 8,  max: 14,        color: '#6366f1' },  // info
  { label: '0–7 days',   min: 0,  max: 7,         color: '#22c55e' },  // fresh
];

function buildOpenWoByAgeOption(
  bucketCounts: number[],  // parallel to WO_AGE_BUCKETS
): ECOption {
  const labels = WO_AGE_BUCKETS.map(b => b.label);
  const total = bucketCounts.reduce((s, v) => s + v, 0);

  return {
    backgroundColor: 'transparent',
    tooltip: {
      trigger: 'axis',
      axisPointer: { type: 'none' },
      extraCssText: GLASS_TOOLTIP,
      formatter: (p: unknown) => {
        const params = (p as { name: string; value: number; color: string }[])[0];
        const pct = total > 0 ? ((params.value / total) * 100).toFixed(0) : '0';
        return (
          `<div style="display:flex;align-items:center;gap:7px;padding:2px 0">` +
            `<span style="display:inline-block;width:8px;height:8px;border-radius:50%;` +
              `background:${params.color};flex-shrink:0"></span>` +
            `<span style="font-weight:600;color:#19202f">${params.name}</span>` +
          `</div>` +
          `<div style="margin-top:4px;font-size:13px;font-weight:700;color:#19202f">` +
            `${params.value.toLocaleString()} ` +
            `<span style="font-size:10px;font-weight:500;color:rgba(0,0,0,0.4)">WOs</span>` +
            `<span style="margin-left:8px;font-size:11px;font-weight:600;color:${params.color}">${pct}%</span>` +
          `</div>`
        );
      },
    },
    grid: { left: '3%', right: '10%', top: '4%', bottom: '4%', containLabel: true },
    xAxis: {
      type: 'value',
      axisLabel: { ...axisLabelStyle(), fontSize: 10 },
      ...yAxisSplitLine(),  // split lines on value axis
      axisLine: { show: false },
      axisTick: { show: false },
    },
    yAxis: {
      type: 'category',
      data: labels,
      inverse: false,
      axisLabel: { ...axisLabelStyle(), fontSize: 11 },
      axisTick: { show: false },
      axisLine: { show: false },
      splitLine: { show: false },
    },
    series: [
      {
        type: 'bar',
        data: bucketCounts.map((v, i) => ({
          value: v,
          itemStyle: barItemStyle(WO_AGE_BUCKETS[i].color) as BarSeriesOption['itemStyle'],
        })),
        barMaxWidth: 32,
        label: {
          show: true,
          position: 'right' as const,
          formatter: (p: { value: unknown }) => (typeof p.value === 'number' && p.value > 0 ? String(p.value) : ''),
          color: 'rgba(0,0,0,0.45)',
          fontSize: 10,
          fontFamily: axisLabelStyle().fontFamily,
        },
        emphasis: {
          focus: 'self',
          itemStyle: { shadowBlur: 14, shadowOffsetY: 4 } as BarSeriesOption['itemStyle'],
        },
      },
    ],
  };
}
// ── WO Status donut ───────────────────────────────────────────────────────────
function buildStatusDonutOption(data: Array<{ name: string; value: number }>): ECOption {
  const colors = palette();
  const total = data.reduce((s, d) => s + d.value, 0);
  return {
    backgroundColor: 'transparent',
    color: colors,
    tooltip: {
      trigger: 'item',
      extraCssText: GLASS_TOOLTIP,
      formatter: donutTooltipFormatter() as never,
    },
    legend: {
      orient: 'vertical',
      right: '2%',
      top: 'center',
      textStyle: { fontSize: 11, color: 'rgba(0,0,0,0.55)', fontFamily: axisLabelStyle().fontFamily },
      itemWidth: 8,
      itemHeight: 8,
      icon: 'circle',
    },
    graphic: donutCenterGraphic(total, '50%', 'Total WOs'),
    series: [
      {
        type: 'pie',
        radius: ['52%', '74%'],
        center: ['50%', '48%'],
        padAngle: 3,
        data: data.map((d, i) => ({
          ...d,
          itemStyle: {
            color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
              { offset: 0,   color: colors[i % colors.length] },
              { offset: 1,   color: colors[i % colors.length] + 'bb' },
            ]),
            borderRadius: 8,
            borderWidth: 2,
            borderColor: 'rgba(255,255,255,0.15)',
            shadowBlur: 14,
            shadowColor: colors[i % colors.length] + '55',
            shadowOffsetY: 4,
          },
        })),
        label: { show: false },
        labelLine: { show: false },
        emphasis: {
          scale: true,
          scaleSize: 7,
          itemStyle: {
            shadowBlur: 28,
            shadowOffsetY: 0,
            shadowColor: 'rgba(0,0,0,0.3)',
          },
        },
      },
    ],
  };
}

// ── Urgency / Priority bar ────────────────────────────────────────────────────
function buildUrgencyBarOption(names: string[], counts: number[]): ECOption {
  const colors = palette();
  // Per-priority semantic colors: Urgent→danger, High→warning, Normal→info, Low→success
  const priorityColorMap: Record<string, string> = {
    urgent: colors[3],  // --danger
    high:   colors[5],  // --warning
    normal: colors[0],  // --info
    low:    colors[2],  // --success
  };

  return {
    backgroundColor: 'transparent',
    tooltip: {
      trigger: 'axis',
      axisPointer: { type: 'none' },
      extraCssText: GLASS_TOOLTIP,
      formatter: (p: unknown) => {
        const params = (p as { name: string; value: number }[])[0];
        const dotColor = priorityColorMap[params.name.toLowerCase()] || colors[0];
        return barTooltipFormatter(dotColor)({ name: params.name, value: params.value });
      },
    },
    grid: { left: '3%', right: '4%', bottom: '3%', top: '8%', containLabel: true },
    xAxis: premiumXAxis(names),
    yAxis: premiumYAxis('Count'),
    series: [
      {
        type: 'bar',
        barMaxWidth: 56,
        data: counts.map((v, i) => {
          const c = priorityColorMap[names[i]?.toLowerCase()] || colors[0];
          return {
            value: v,
            itemStyle: barItemStyle(c) as BarSeriesOption['itemStyle'],
          };
        }),
        emphasis: {
          focus: 'self',
          itemStyle: { shadowBlur: 18, shadowOffsetY: 6 } as BarSeriesOption['itemStyle'],
        },
      },
    ],
  };
}

// =============================================================================
// Vacancy KPI Scatter — exposed on window for app.js to call
// =============================================================================
//
// Each bubble = one vacant unit.
// X axis    = days vacant (0 → max)
// Y axis    = property name (categorical, one lane per property)
// Size      = sqft when available, uniform 12px otherwise
// Color     = severity: <30d green, 30–59d amber, 60+d red
//
// Data shape fed from app.js (matches what loadPropertyVacancies already has):
//   { property: string, unit: string, days: number, sqft: number|null, rent: number|null }[]

interface VacancyDot {
  property: string;
  unit: string;
  days: number;
  sqft: number | null;
  rent: number | null;
}


function buildVacancyBarOption(dots: VacancyDot[]): ECOption {
  if (!dots.length) {
    return {
      backgroundColor: 'transparent',
      graphic: [{
        type: 'text',
        left: 'center', top: 'middle',
        style: { text: 'No vacancy data', fill: 'rgba(0,0,0,0.3)', fontSize: 13 },
      }],
    };
  }

  const font = 'var(--font-mono, "Inconsolata", monospace)';

  // ── Aggregate per property ────────────────────────────────────────────────
  const propMap: Record<string, { fresh: number; elevated: number; critical: number; units: string[] }> = {};
  for (const d of dots) {
    if (!propMap[d.property]) propMap[d.property] = { fresh: 0, elevated: 0, critical: 0, units: [] };
    const row = propMap[d.property];
    if (d.days >= 60)      { row.critical++;  }
    else if (d.days >= 30) { row.elevated++;  }
    else                   { row.fresh++;     }
    row.units.push(`${d.unit} (${d.days}d)`);
  }

  // Sort properties: most critical units first, then total vacant desc
  const properties = Object.keys(propMap).sort((a, b) => {
    const da = propMap[a], db = propMap[b];
    if (db.critical !== da.critical) return db.critical - da.critical;
    if (db.elevated !== da.elevated) return db.elevated - da.elevated;
    return (db.fresh + db.elevated + db.critical) - (da.fresh + da.elevated + da.critical);
  });

  const freshData    = properties.map(p => propMap[p].fresh);
  const elevatedData = properties.map(p => propMap[p].elevated);
  const criticalData = properties.map(p => propMap[p].critical);

  const BAR_RADIUS = [0, 3, 3, 0] as [number, number, number, number];

  return {
    backgroundColor: 'transparent',
    animation: true,
    animationDuration: 500,
    tooltip: {
      trigger: 'axis',
      axisPointer: { type: 'shadow' },
      extraCssText: GLASS_TOOLTIP,
      formatter: (params: unknown) => {
        const list = params as Array<{ seriesName: string; value: number; color: string }>;
        const prop = properties[list[0] ? (list[0] as unknown as { dataIndex: number }).dataIndex : 0] || '';
        const total = (propMap[prop]?.fresh ?? 0) + (propMap[prop]?.elevated ?? 0) + (propMap[prop]?.critical ?? 0);
        let html = `<div style="font-weight:700;font-size:12px;color:#19202f;margin-bottom:6px">${prop}</div>`;
        for (const p of list) {
          if (!p.value) continue;
          html += `<div style="display:flex;align-items:center;gap:6px;margin:2px 0">` +
            `<span style="display:inline-block;width:8px;height:8px;border-radius:2px;background:${p.color};flex-shrink:0"></span>` +
            `<span style="font-size:11px;color:rgba(0,0,0,0.6)">${p.seriesName}:</span>` +
            `<span style="font-weight:700;font-size:11px;color:#19202f">${p.value}</span>` +
          `</div>`;
        }
        html += `<div style="font-size:11px;color:rgba(0,0,0,0.45);margin-top:4px;border-top:1px solid rgba(0,0,0,0.07);padding-top:4px">${total} unit${total !== 1 ? 's' : ''} vacant</div>`;
        return html;
      },
    },
    legend: {
      data: ['< 30 days', '30–59 days', '60+ days'],
      bottom: 0,
      itemWidth: 12,
      itemHeight: 8,
      textStyle: { fontSize: 10, color: 'rgba(0,0,0,0.45)', fontFamily: font },
    },
    grid: { left: 8, right: 16, top: 8, bottom: 36, containLabel: true },
    xAxis: {
      type: 'value',
      minInterval: 1,
      axisLabel: { fontSize: 10, color: 'rgba(0,0,0,0.35)', fontFamily: font },
      splitLine: { lineStyle: { type: 'dashed' as const, color: 'rgba(0,0,0,0.06)' } },
      axisLine: { show: false },
      axisTick: { show: false },
    },
    yAxis: {
      type: 'category',
      data: properties,
      inverse: true,
      axisLabel: {
        fontSize: 11,
        color: 'rgba(0,0,0,0.55)',
        fontFamily: font,
        overflow: 'truncate',
        width: 120,
      },
      axisTick: { show: false },
      axisLine: { show: false },
    },
    series: [
      {
        name: '< 30 days',
        type: 'bar',
        stack: 'vacancies',
        barMaxWidth: 20,
        itemStyle: {
          color: new echarts.graphic.LinearGradient(0, 0, 1, 0, [
            { offset: 0, color: '#22c55e' },
            { offset: 1, color: '#4ade80' },
          ]),
          borderRadius: BAR_RADIUS,
        },
        emphasis: { itemStyle: { shadowBlur: 8, shadowColor: '#22c55e66' } },
        data: freshData,
      },
      {
        name: '30–59 days',
        type: 'bar',
        stack: 'vacancies',
        barMaxWidth: 20,
        itemStyle: {
          color: new echarts.graphic.LinearGradient(0, 0, 1, 0, [
            { offset: 0, color: '#f59e0b' },
            { offset: 1, color: '#fbbf24' },
          ]),
          borderRadius: BAR_RADIUS,
        },
        emphasis: { itemStyle: { shadowBlur: 8, shadowColor: '#f59e0b66' } },
        data: elevatedData,
      },
      {
        name: '60+ days',
        type: 'bar',
        stack: 'vacancies',
        barMaxWidth: 20,
        itemStyle: {
          color: new echarts.graphic.LinearGradient(0, 0, 1, 0, [
            { offset: 0, color: '#ef4444' },
            { offset: 1, color: '#f87171' },
          ]),
          borderRadius: BAR_RADIUS,
        },
        emphasis: { itemStyle: { shadowBlur: 8, shadowColor: '#ef444466' } },
        data: criticalData,
      },
    ],
  };
}

// Expose on window so app.js (non-module) can call it
// Keep old name as alias so any cached references still work
(window as unknown as Record<string, unknown>).buildVacancyBarOption = buildVacancyBarOption;
(window as unknown as Record<string, unknown>).buildVacancyScatterOption = buildVacancyBarOption;
(window as unknown as Record<string, unknown>).echartsCore = echarts;

// =============================================================================
// Inspection Map — EffectScatter on abstract 0-100 XY grid (Phoenix GPS bounds)
// =============================================================================

/** Maps real GPS coords to a 0–100 abstract grid (Phoenix metro bounds). */
export function normalizeGpsToGrid(lat: number, lon: number): [number, number] {
  const LON_MIN = -112.35, LON_MAX = -111.75;
  const LAT_MIN =  33.25,  LAT_MAX =  33.75;
  const x = Math.max(0, Math.min(100, ((lon - LON_MIN) / (LON_MAX - LON_MIN)) * 100));
  const y = Math.max(0, Math.min(100, ((lat - LAT_MIN) / (LAT_MAX - LAT_MIN)) * 100));
  return [Math.round(x * 10) / 10, Math.round(y * 10) / 10];
}

interface InspectionRecord {
  property_name?: string;
  propertyName?: string;
  property_id?: string | number;
  propertyId?: string | number;
  last_inspection_date?: string;
  lastInspectionDate?: string;
  _lat?: number;
  _lon?: number;
  _x?: number;
  _y?: number;
}

/**
 * Builds an EffectScatter chart option for the inspection property map.
 * Each point is a property; size = days since last inspection; color = urgency.
 * Pass geocoded data: each record should have _x/_y (0-100 grid) or _lat/_lon.
 */
export function buildInspectionMapOption(inspections: InspectionRecord[]): ECOption {
  const today = Date.now();
  const MS_DAY = 86400000;

  const data = inspections
    .filter(r => r._x != null && r._y != null)
    .map(r => {
      const name = r.property_name || r.propertyName || 'Unknown';
      const lastDate = r.last_inspection_date || r.lastInspectionDate || '';
      const daysAgo = lastDate
        ? Math.round((today - new Date(lastDate).getTime()) / MS_DAY)
        : 999;
      const id = r.property_id || r.propertyId || '';
      // symbolSize: 12 base + scale with days (max 40)
      const size = Math.min(40, 12 + Math.round(daysAgo / 10));
      // color: green <30d, amber 30-89d, red 90+d
      const color = daysAgo < 30 ? '#22c55e' : daysAgo < 90 ? '#f59e0b' : '#ef4444';
      return {
        name,
        value: [r._x, r._y, daysAgo, id],
        symbolSize: size,
        itemStyle: { color },
        label: { show: false },
      };
    });

  return {
    backgroundColor: 'transparent',
    tooltip: {
      trigger: 'item',
      formatter: (p: any) => {
        const [, , days, id] = p.value as [number, number, number, string | number];
        return `<b>${p.name}</b><br/>Last insp: ${days === 999 ? 'Never' : days + ' days ago'}${id ? `<br/>ID: ${id}` : ''}`;
      },
    },
    grid: { top: 10, bottom: 10, left: 10, right: 10, containLabel: false },
    xAxis: {
      type: 'value', min: 0, max: 100,
      splitLine: { show: false },
      axisLine: { show: false },
      axisTick: { show: false },
      axisLabel: { show: false },
    },
    yAxis: {
      type: 'value', min: 0, max: 100,
      splitLine: { show: false },
      axisLine: { show: false },
      axisTick: { show: false },
      axisLabel: { show: false },
    },
    series: [{
      type: 'effectScatter',
      rippleEffect: { scale: 2.5, brushType: 'stroke' },
      data,
      zlevel: 1,
    }],
  };
}

(window as unknown as Record<string, unknown>).buildInspectionMapOption = buildInspectionMapOption;
(window as unknown as Record<string, unknown>).normalizeGpsToGrid = normalizeGpsToGrid;

// =============================================================================
// Turnover Pipeline Funnel
// =============================================================================

interface TurnPipeEntry {
  stages?: {
    upcoming?:     { done: boolean };
    moveout?:      { done: boolean };
    inspection?:   { done: boolean };
    wo_created?:   { done: boolean };
    est_received?: { done: boolean };
    assigned?:     { done: boolean };
    work_done?:    { done: boolean };
  };
  isCompleted?: boolean;
  isClosed?: boolean;
}

/**
 * Builds a Funnel chart option from TURN_PIPE_DATA showing how many turns
 * are at each pipeline stage (cumulative — each stage includes all later stages).
 */
export function buildTurnoverPipelineOption(turns: TurnPipeEntry[]): ECOption {
  const active = turns.filter(t => !t.isClosed);
  const total  = active.length || 1; // avoid div-zero

  const stageDefs: Array<{ key: keyof NonNullable<TurnPipeEntry['stages']>; label: string; color: string }> = [
    { key: 'upcoming',     label: 'Upcoming',    color: '#818cf8' },
    { key: 'moveout',      label: 'Move-Out',    color: '#60a5fa' },
    { key: 'inspection',   label: 'Inspection',  color: '#34d399' },
    { key: 'wo_created',   label: 'WOs Created', color: '#fbbf24' },
    { key: 'est_received', label: 'Est Received',color: '#fb923c' },
    { key: 'assigned',     label: 'Assigned',    color: '#f87171' },
    { key: 'work_done',    label: 'Work Done',   color: '#a78bfa' },
  ];

  const funnelData = stageDefs.map(def => {
    const count = active.filter(t => t.stages?.[def.key]?.done).length;
    return {
      name: def.label,
      value: Math.round((count / total) * 100),
      _count: count,
      itemStyle: { color: def.color },
    };
  });

  return {
    backgroundColor: 'transparent',
    tooltip: {
      trigger: 'item',
      formatter: (p: any) => `<b>${p.name}</b><br/>${p.data._count} turns (${p.value}%)`,
    },
    series: [{
      type: 'funnel',
      left: '10%',
      width: '80%',
      top: 20,
      bottom: 20,
      sort: 'none',
      gap: 3,
      label: {
        show: true,
        position: 'inside',
        formatter: (p: any) => `${p.name}: ${p.data._count}`,
        color: '#fff',
        fontSize: 11,
        fontWeight: 600,
      },
      data: funnelData,
    }],
  };
}

(window as unknown as Record<string, unknown>).buildTurnoverPipelineOption = buildTurnoverPipelineOption;

// =============================================================================
// Portfolio Health Sunburst
// Levels: Property (inner) → Tenant Status (middle) → Unit (outer)
// Value: Market Rent (slice size); $0/vacant units use fallback 100 so they
// remain visible — matching the server-side shapePortfolioPulse behaviour.
// =============================================================================

interface SunburstNode {
  name: string;
  value?: number;
  children?: SunburstNode[];
}

// Status → color mapping (matches common AppFolio tenant_status values)
const STATUS_COLORS: Record<string, string> = {
  'Current':          '#22c55e',
  'Occupied':         '#22c55e',
  'Vacant':           '#ef4444',
  'Notice':           '#f59e0b',
  'Notice Unrented':  '#f59e0b',
  'Notice Rented':    '#a3e635',
  'Eviction':         '#dc2626',
  'Past Due':         '#fb923c',
  'Month-to-Month':   '#60a5fa',
};

function colorizeStatus(nodes: SunburstNode[]): SunburstNode[] {
  return nodes.map(propNode => ({
    ...propNode,
    itemStyle: { color: '#6366f1', borderWidth: 2, borderColor: '#1e1b4b' },
    children: (propNode.children || []).map(statusNode => ({
      ...statusNode,
      itemStyle: {
        color: STATUS_COLORS[statusNode.name] ?? '#94a3b8',
        borderWidth: 1,
        borderColor: '#0f172a',
      },
      children: (statusNode.children || []).map(unitNode => ({
        ...unitNode,
        itemStyle: { opacity: 0.85 },
      })),
    })),
  }));
}

export function buildPortfolioSunburstOption(data: SunburstNode[]): ECOption {
  const colored = colorizeStatus(data);

  return {
    backgroundColor: 'transparent',
    tooltip: {
      trigger: 'item',
      formatter: (p: any) => {
        const val = p.value != null ? `$${Number(p.value).toLocaleString()}` : '';
        const path = (p.treePathInfo as Array<{ name: string }>)
          ?.slice(1).map(n => n.name).join(' › ') ?? p.name;
        return `<b>${path}</b>${val ? `<br/>Rent: ${val}` : ''}`;
      },
    },
    series: [{
      type: 'sunburst',
      data: colored,
      radius: ['8%', '95%'],
      sort: 'desc',
      emphasis: {
        focus: 'ancestor',
        itemStyle: { shadowBlur: 10, shadowColor: 'rgba(99,102,241,0.6)' },
      },
      levels: [
        {},
        // Level 1 — Property
        {
          r0: '8%', r: '35%',
          label: { rotate: 'tangential', fontSize: 10, fontWeight: 700, color: '#e0e7ff' },
          itemStyle: { borderWidth: 2 },
        },
        // Level 2 — Tenant Status
        {
          r0: '35%', r: '68%',
          label: { fontSize: 9, color: '#fff', fontWeight: 600 },
          itemStyle: { borderWidth: 1 },
        },
        // Level 3 — Unit (outer ring, labels outside)
        {
          r0: '68%', r: '72%',
          label: { position: 'outside', fontSize: 8, color: '#94a3b8', silent: true },
          itemStyle: { borderWidth: 0.5 },
        },
      ],
    }],
  };
}

(window as unknown as Record<string, unknown>).buildPortfolioSunburstOption = buildPortfolioSunburstOption;

// =============================================================================
// Main async data fetch + render
// =============================================================================
async function fetchAndRenderDashboardData(): Promise<void> {
  // Show loading spinners on all charts.
  ALL_CHARTS.forEach(c =>
    c.showLoading('default', {
      text: 'Loading…',
      color: cssVar('--accent') || '#6366f1',
      maskColor: 'rgba(255,255,255,0.6)',
    })
  );

  const baseUrl = resolveProxyUrl();
  const headers = { Accept: 'application/json', ...proxyAuthHeaders() };

  // ── 1. Occupancy Doughnut ─────────────────────────────────────────
  try {
    const resp = await fetch(`${baseUrl}?action=chart_occupancy`, { headers });
    if (resp.ok) {
      const data = await resp.json();
      if (chartOccupancy && Array.isArray(data) && data.length > 0) {
        chartOccupancy.setOption((window as any).buildOccupancyDonutOption(data));
        chartOccupancy.hideLoading();
        const meta = document.getElementById('dashOccupancyMeta');
        if (meta) meta.textContent = `Total: ${data.reduce((s: number, d: any) => s + (d.value || 0), 0)} units`;
      } else if (chartOccupancy) {
        chartOccupancy.hideLoading();
        chartOccupancy.setOption((window as any).buildOccupancyDonutOption([]));
        const meta = document.getElementById('dashOccupancyMeta');
        if (meta) meta.textContent = 'No data';
      }
    } else if (chartOccupancy) {
      chartOccupancy.hideLoading();
    }
  } catch (e) {
    console.error('[dashboard] Occupancy fetch failed:', e);
    if (chartOccupancy) chartOccupancy.hideLoading();
  }

  // ── 2. Move-Outs Bar ──────────────────────────────────────────────
  try {
    const resp = await fetch(`${baseUrl}?action=chart_moveouts`, { headers });
    if (resp.ok) {
      const data = await resp.json();
      if (chartMoveOuts && data && Array.isArray(data.labels)) {
        chartMoveOuts.setOption((window as any).buildMoveOutsBarOption(data.labels, data.values || []));
        chartMoveOuts.hideLoading();
        const meta = document.getElementById('dashMoveOutsMeta');
        const vals = data.values || [];
        if (meta) meta.textContent = `Next 90 days: ${vals.reduce((a: number, b: number) => a + b, 0)} move-outs`;
      } else if (chartMoveOuts) {
        chartMoveOuts.hideLoading();
        const meta = document.getElementById('dashMoveOutsMeta');
        if (meta) meta.textContent = 'No data';
      }
    } else if (chartMoveOuts) {
      chartMoveOuts.hideLoading();
    }
  } catch (e) {
    console.error('[dashboard] Move-Outs fetch failed:', e);
    if (chartMoveOuts) chartMoveOuts.hideLoading();
  }

  // ── 3. Portfolio Treemap ─────────────────────────────────────────
  try {
    const resp = await fetch(`${baseUrl}?action=chart_portfolio_owner`, { headers });
    if (resp.ok) {
      const data = await resp.json();
      if (chartPortfolio && Array.isArray(data) && data.length > 0) {
        chartPortfolio.setOption((window as any).buildPortfolioTreemapOption(data));
        chartPortfolio.hideLoading();
        const meta = document.getElementById('dashPortfolioTreeMeta');
        if (meta) meta.textContent = `${data.length} owners/groups`;
      } else if (chartPortfolio) {
        chartPortfolio.hideLoading();
        const meta = document.getElementById('dashPortfolioTreeMeta');
        if (meta) meta.textContent = 'No data';
      }
    } else if (chartPortfolio) {
      chartPortfolio.hideLoading();
    }
  } catch (e) {
    console.error('[dashboard] Portfolio fetch failed:', e);
    if (chartPortfolio) chartPortfolio.hideLoading();
  }

  // ── 4. Leasing Velocity Area Chart ───────────────────────────────
  try {
    const resp = await fetch(`${baseUrl}?action=chart_leasing_velocity`, { headers });
    if (resp.ok) {
      const data = await resp.json();
      if (chartVelocity && data && Array.isArray(data.dates)) {
        chartVelocity.setOption((window as any).buildLeasingVelocityOption(data.dates, data.moveIns || [], data.moveOuts || []));
        chartVelocity.hideLoading();
        const meta = document.getElementById('dashVelocityMeta');
        if (meta) {
          const totalIn = (data.moveIns || []).reduce((a: number, b: number) => a + b, 0);
          const totalOut = (data.moveOuts || []).reduce((a: number, b: number) => a + b, 0);
          meta.textContent = `In: ${totalIn} · Out: ${totalOut}`;
        }
      } else if (chartVelocity) {
        chartVelocity.hideLoading();
        const meta = document.getElementById('dashVelocityMeta');
        if (meta) meta.textContent = 'No data';
      }
    } else if (chartVelocity) {
      chartVelocity.hideLoading();
    }
  } catch (e) {
    console.error('[dashboard] Velocity fetch failed:', e);
    if (chartVelocity) chartVelocity.hideLoading();
  }

  // ===========================================================================
  // Original WO Processing (re-integrated)
  // ===========================================================================
  let results: WoRecord[] = [];
  const appWOs = (window as unknown as Record<string, unknown>).WORK_ORDERS;
  if (Array.isArray(appWOs) && (appWOs as WoRecord[]).length > 0) {
    results = appWOs as WoRecord[];
    console.debug('[dashboard] using WORK_ORDERS from app.js (' + results.length + ' rows)');
  } else {
    // Fallback: fetch raw data directly from the proxy.
    try {
      // @ts-ignore - buildWorkOrdersUrl is defined earlier in the file
      const url = buildWorkOrdersUrl();
      const resp = await fetch(url, { headers });
      if (resp.ok) {
        const body = await resp.json() as { ok: boolean; results: WoRecord[]; count: number };
        if (body.ok) results = Array.isArray(body.results) ? body.results : [];
      }
    } catch (err) {
      console.error('[dashboard] fetch WOs failed:', err);
    }
  }

  // If we have WO results, process and render them
  if (results.length > 0) {
    // ── PM Performance Radar (multi-axis per PM from real WO data) ───────────────
    if (chartPmBar) {
      // @ts-ignore - buildPmRadarOption is defined elsewhere
      chartPmBar.setOption(buildPmRadarOption(results));
      chartPmBar.hideLoading();
    }

    // ── Open WOs by Age (replaces WO Status donut) ───────────────────────────
    if (chartStatusDonut) {
      const today = new Date();
      // @ts-ignore
      const openRows = results.filter(r => !isClosedStatus(statusField(r)));
      // @ts-ignore
      const bucketCounts = WO_AGE_BUCKETS.map(bucket => {
        return openRows.filter(r => {
          const created = r.created_at || r.CreatedAt || r.created || r.date_created || '';
          if (!created) return false;
          const age = Math.floor((today.getTime() - new Date(String(created)).getTime()) / 86400000);
          return age >= bucket.min && age <= bucket.max;
        }).length;
      });
      const metaEl = document.getElementById('kpiVacancyGroupsSub');
      if (metaEl) {
        const total = bucketCounts.reduce((s: number, v: number) => s + v, 0);
        const critical = bucketCounts[0]; // 60+ days bucket
        metaEl.textContent = total > 0
          ? `${total} open WO${total !== 1 ? 's' : ''}${critical > 0 ? ` · ${critical} critical (60+ d)` : ''}`
          : 'No open work orders';
      }
      // @ts-ignore
      chartStatusDonut.setOption(buildOpenWoByAgeOption(bucketCounts));
      chartStatusDonut.hideLoading();
    }

    // ── PM Load Bar (ALL open WOs per PM, not just closed) ────────────────────
    if (chartPmLoad) {
      // @ts-ignore
      const openRows = results.filter(r => !isClosedStatus(statusField(r)));
      // @ts-ignore
      const { names, counts } = groupCount(openRows, pmField, 15);
      // @ts-ignore
      chartPmLoad.setOption(buildPmLoadBarOption(names, counts));
      chartPmLoad.hideLoading();
    }

    // ── WO Type Treemap (category × urgency) ──────────────────────────────────
    if (chartWoType) {
      // @ts-ignore
      chartWoType.setOption(buildWoTypeTreemapOption(results));
      chartWoType.hideLoading();
    }

    // ── Urgency Bar (WOs by priority) ─────────────────────────────────────────
    if (chartUrgency) {
      // @ts-ignore
      const { names, counts } = groupCount(results, priorityField, 8);
      // @ts-ignore
      chartUrgency.setOption(buildUrgencyBarOption(names, counts));
      chartUrgency.hideLoading();
    }
  }

  // Force a resize pass after all options are set and rendering is complete
  requestAnimationFrame(() => {
    requestAnimationFrame(() => {
      ALL_CHARTS.forEach(c => {
        try { c.resize(); } catch (_) { /* ignore resize during render */ }
      });
    });
  });
}

// =============================================================================
// NEW PM CHART BUILDERS — Property Manager Dashboard
// =============================================================================

// ── 1. Real-Time Occupancy Doughnut ────────────────────────────────────────
export function buildOccupancyDonutOption(
  data: Array<{ name: string; value: number }>,
): ECOption {
  const colorMap: Record<string, string> = {
    Occupied: "#00E676",
    Notice: "#FFEA00",
    "Vacant Unrented": "#FF1744",
    "Vacant Rented": "#2979FF",
  };
  const safeData = Array.isArray(data) ? data : [];
  const colors = safeData.map(d => colorMap[d.name] || "#888");

  return {
    backgroundColor: "transparent",
    tooltip: { trigger: "item" as const },
    color: colors,
    series: [{
      name: "Portfolio Occupancy",
      type: "pie" as const,
      radius: safeData.length > 0 ? ["40%", "70%"] : ["0%", "0%"],
      avoidLabelOverlap: false,
      itemStyle: { borderRadius: 10, borderColor: "#121212", borderWidth: 2 },
      label: { show: safeData.length > 0, color: "#fff", formatter: "{b}\n{c}" },
      data: safeData,
    }],
  };
}

// ── 2. Upcoming Move-Outs Forecast (Bar) ──────────────────────────────────
export function buildMoveOutsBarOption(
  labels: string[],
  values: number[],
): ECOption {
  return {
    backgroundColor: "transparent",
    tooltip: { trigger: "axis" as const, backgroundColor: "rgba(20,20,20,0.9)" },
    grid: { left: "5%", right: "5%", bottom: "10%", top: "15%", containLabel: true },
    xAxis: {
      type: "category" as const,
      data: labels,
      axisLabel: { color: "#888" },
    },
    yAxis: {
      type: "value" as const,
      splitLine: { lineStyle: { color: "#333" } },
      axisLabel: { color: "#888" },
    },
    series: [{
      data: values,
      type: "bar" as const,
      barWidth: "40%",
      itemStyle: {
        borderRadius: [5, 5, 0, 0],
        color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
          { offset: 0, color: "#FF9100" },
          { offset: 1, color: "#FF3D00" },
        ]),
      },
    }],
  };
}

// ── 3. Portfolio Distribution Treemap ───────────────────────────────────────
export function buildPortfolioTreemapOption(
  data: Array<{ name: string; value: number }>,
): ECOption {
  const safeData = Array.isArray(data) ? data.filter(d => d && typeof d === 'object' && 'name' in d) : [];

  return {
    backgroundColor: "transparent",
    tooltip: { formatter: "{b}: {c} Units" },
    series: [{
      type: "treemap" as const,
      data: safeData,
      roam: false,
      nodeClick: false as any,
      breadcrumb: { show: false },
      itemStyle: { borderColor: "#121212", borderWidth: 2, gapWidth: 1 },
      colorMappingBy: "value" as const,
      color: ["#3F51B5", "#673AB7", "#9C27B0", "#E91E63"],
    }],
  };
}

// ── 4. Leasing Velocity (Area Chart) ───────────────────────────────────────
export function buildLeasingVelocityOption(
  dates: string[],
  moveIns: number[],
  moveOuts: number[],
): ECOption {
  return {
    backgroundColor: "transparent",
    tooltip: { trigger: "axis" as const, backgroundColor: "rgba(20,20,20,0.9)" },
    legend: {
      data: ["Move-Ins", "Move-Outs"],
      textStyle: { color: "#e0e0e0" },
      top: "0%",
    },
    grid: { left: "5%", right: "5%", bottom: "10%", top: "15%", containLabel: true },
    xAxis: {
      type: "category" as const,
      boundaryGap: false,
      data: dates,
      axisLabel: { color: "#888" },
    },
    yAxis: {
      type: "value" as const,
      splitLine: { lineStyle: { color: "#333" } },
      axisLabel: { color: "#888" },
    },
    series: [
      {
        name: "Move-Ins",
        type: "line" as const,
        smooth: true,
        lineStyle: { width: 3, color: "#00E676" },
        areaStyle: {
          color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
            { offset: 0, color: "rgba(0,230,118,0.5)" },
            { offset: 1, color: "rgba(0,230,118,0)" },
          ]),
        },
        data: moveIns,
      },
      {
        name: "Move-Outs",
        type: "line" as const,
        smooth: true,
        lineStyle: { width: 3, color: "#FF1744" },
        areaStyle: {
          color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
            { offset: 0, color: "rgba(255,23,68,0.5)" },
            { offset: 1, color: "rgba(255,23,68,0)" },
          ]),
        },
        data: moveOuts,
      },
    ],
  };
}

// =============================================================================
// WO Sankey Flow — Property Group → Type → Status → Assigned
// Shows how work orders flow through the operational pipeline.
// =============================================================================

interface WoRecord {
  property_group_id?: string;
  property_group?: string;
  propertyGroupId?: string;
  wo_type?: string;
  work_order_type?: string;
  Type?: string;
  type?: string;
  status?: string;
  Status?: string;
  vendor_id?: string;
  vendorId?: string;
  VendorId?: string;
  assigned_to?: string;
  assignedTo?: string;
  assigned_user?: string;
  assigned_users?: any;
  AssignedUsers?: any;
  vendor?: string;
  vendor_name?: string;
  vendorName?: string;
  priority?: string;
  Priority?: string;
}

const SANKEY_COLORS: Record<string, string> = {
  // Property groups
  Phoenix: '#3b82f6',
  Tucson: '#8b5cf6',
  // Types
  Internal: '#6366f1',
  Resident: '#f59e0b',
  // Statuses
  New: '#94a3b8',
  Assigned: '#3b82f6',
  Scheduled: '#6366f1',
  Estimated: '#8b5cf6',
  'In Progress': '#a855f7',
  'Work Completed': '#22c55e',
  Completed: '#22c55e',
  Canceled: '#ef4444',
  // Assignment
  Unassigned: '#ef4444',
  Urgent: '#dc2626',
};

function sankeyColor(key: string): string {
  const k = String(key || '').trim();
  return SANKEY_COLORS[k] || cssVar('--accent') || '#6366f1';
}

/**
 * Builds a Sankey diagram showing WO flow:
 *   Property Group → Type (Internal/Resident) → Status → Assignee
 *
 * Data comes from the v0-synced WORK_ORDERS global.
 * Each node is a category; link thickness = count of WOs in that path.
 */
export function buildWoSankeyOption(wos: WoRecord[]): ECOption {
  const nodes: { name: string; itemStyle: { color: string } }[] = [];
  const links: { source: string; target: string; value: number }[] = [];
  const nodeSet = new Set<string>();

  function addNode(name: string, color: string): void {
    if (nodeSet.has(name)) return;
    nodeSet.add(name);
    nodes.push({ name, itemStyle: { color } });
  }

  function addLink(source: string, target: string): void {
    const key = `${source}→${target}`;
    const existing = links.find(l => l.source === source && l.target === target);
    if (existing) {
      existing.value++;
    } else {
      links.push({ source, target, value: 1 });
    }
  }

  const groupNames: Record<string, string> = {};

  for (const wo of wos) {
    // Layer 1: Property Group
    const rawGroup = String(
      wo.property_group || wo.property_group_id || wo.propertyGroupId || 'Unassigned Group',
    ).trim();
    const group = groupNames[rawGroup] || rawGroup.split('-')[0] || rawGroup;
    if (!groupNames[rawGroup]) groupNames[rawGroup] = group;

    // Layer 2: WO Type
    const woType = String(wo.wo_type || wo.Type || wo.type || wo.work_order_type || 'Other').trim();

    // Layer 3: Status
    const status = String(wo.status || wo.Status || 'Unknown').trim();

    // Layer 4: Assignee (vendor name or assigned user or Unassigned)
    const assignedUsers = Array.isArray(wo.assigned_users || wo.AssignedUsers) ? (wo.assigned_users || wo.AssignedUsers) : [];
    const vendorName = String(wo.vendor || wo.vendor_name || wo.vendorName || '').trim();
    const primaryAssignee = assignedUsers.length > 0
      ? String(assignedUsers[0]?.Name || [assignedUsers[0]?.FirstName, assignedUsers[0]?.LastName].filter(Boolean).join(' ')).trim()
      : '';
    const assignee = vendorName || primaryAssignee || wo.assigned_to || wo.assigned_user || 'Unassigned';

    // Prefix layers to avoid name collisions across layers
    const groupNode = `📁 ${group}`;
    const typeNode = `🏷 ${woType}`;
    const statusNode = `⚡ ${status}`;
    const assigneeNode = `👤 ${assignee.length > 20 ? assignee.slice(0, 18) + '…' : assignee}`;

    addNode(groupNode, sankeyColor(group));
    addNode(typeNode, sankeyColor(woType));
    addNode(statusNode, sankeyColor(status));
    addNode(assigneeNode, assignee === 'Unassigned' ? '#ef4444' : '#6366f1');

    addLink(groupNode, typeNode);
    addLink(typeNode, statusNode);
    addLink(statusNode, assigneeNode);
  }

  return {
    tooltip: {
      trigger: 'item' as const,
      triggerOn: 'mousemove' as const,
      formatter: (params: any) => {
        if (params.dataType === 'edge') {
          return `<div style="font-weight:600;color:#19202f">${params.data.source.split(' ').slice(1).join(' ')}</div>` +
            `<div style="font-size:12px;color:rgba(0,0,0,0.5)">→ ${params.data.target.split(' ').slice(1).join(' ')}</div>` +
            `<div style="margin-top:4px;font-size:14px;font-weight:700;color:#19202f">${params.data.value} WOs</div>`;
        }
        const name = params.name.split(' ').slice(1).join(' ');
        return `<div style="font-weight:600;color:#19202f">${name}</div>`;
      },
      extraCssText: GLASS_TOOLTIP,
    },
    series: [{
      type: 'sankey' as const,
      data: nodes,
      links,
      emphasis: {
        focus: 'adjacency' as const,
        lineStyle: { opacity: 0.7 },
      },
      lineStyle: {
        color: 'gradient' as const,
        curveness: 0.5,
        opacity: 0.35,
      },
      itemStyle: {
        borderWidth: 0,
        borderRadius: 4,
      },
      label: {
        color: cssVar('--text') || '#e2e8f0',
        fontSize: 11,
        fontFamily: 'var(--font-mono, "Inconsolata", monospace)',
      },
      layoutIterations: 64,
      nodeWidth: 20,
      nodeGap: 12,
    }],
    animationDuration: 1200,
    animationEasing: 'cubicInOut' as const,
  };
}

// Expose all chart builders on window for use from app.js
(window as any).buildOccupancyDonutOption = buildOccupancyDonutOption;
(window as any).buildMoveOutsBarOption = buildMoveOutsBarOption;
(window as any).buildPortfolioTreemapOption = buildPortfolioTreemapOption;
(window as any).buildLeasingVelocityOption = buildLeasingVelocityOption;
(window as any).buildWoSankeyOption = buildWoSankeyOption;

// =============================================================================
// Init — deferred to DOMContentLoaded so every chart container has its
// CSS-driven dimensions before ECharts measures them.
// =============================================================================
function initDashboardCharts(): void {
  // Mount chart instances now that the DOM is fully painted.
  // Note: These IDs may not exist on all routes/views - mountChart returns null if not found.
  chartOccupancy   = mountChart('dashOccupancyDonut');
  chartMoveOuts    = mountChart('dashMoveOutsBar');
  chartPortfolio   = mountChart('dashPortfolioTreemap');
  chartVelocity    = mountChart('dashLeasingVelocity');
  chartPmBar       = mountChart('dashMainChart');
  chartStatusDonut = mountChart('dashAgingChart');
  chartPmLoad      = mountChart('dashPmLoadChart');
  chartWoType      = mountChart('dashWoTypeChart');
  chartUrgency     = mountChart('dashUrgencyChart');

  ALL_CHARTS = [
    chartOccupancy,
    chartMoveOuts,
    chartPortfolio,
    chartVelocity,
    chartPmBar,
    chartStatusDonut,
    chartPmLoad,
    chartWoType,
    chartUrgency,
  ].filter(Boolean) as echarts.ECharts[];

  // Show lightweight placeholders while the API call is in-flight.
  chartOccupancy?.setOption(makePlaceholderDonut());
  chartMoveOuts?.setOption(makePlaceholderBar());
  chartPortfolio?.setOption(makePlaceholderBar());
  chartVelocity?.setOption(makePlaceholderBar());
  chartPmBar?.setOption(makePlaceholderBar());
  chartStatusDonut?.setOption(makePlaceholderBar());
  chartPmLoad?.setOption(makePlaceholderBar());
  chartWoType?.setOption(makePlaceholderDonut());
  chartUrgency?.setOption(makePlaceholderBar());

  fetchAndRenderDashboardData();
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', initDashboardCharts);
} else {
  // DOM already ready (script loaded with defer/async or after DOMContentLoaded).
  initDashboardCharts();
}
