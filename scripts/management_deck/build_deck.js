// Management performance deck — Mitsui Thailand hotel portfolio, H1 2026
// All charts are native PowerPoint charts (fully editable).
const pptxgen = require("pptxgenjs");
const data = require("./deck_data.json");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5

// ---------- palette ----------
const NAVY = "1E2761";      // 2026 / primary
const SKY = "7FA4DC";       // 2025
const ICE = "C9D6EE";       // 2024
const GREY = "AEB6C2";      // 2019 reference
const GOLD = "D9A441";      // budget
const INK = "26292E";       // body text
const MUT = "6B7280";       // muted text
const GOOD = "2E7D32";
const BAD = "C0392B";
const PANEL = "F4F6FB";     // light panel fill
const SEG_COLORS = ["1E2761", "4A6FC4", "7FA4DC", "AFC6EC", "D9A441",
                    "8A5A83", "2F8F83", "C97B4A", "8FA35C", "AEB6C2"];
const F = "Calibri";
const W = 13.33, H = 7.5, M = 0.6;

const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const DAYS = [31, 28, 31, 30, 31, 30];

// ---------- helpers ----------
const fmtM = (v) => (v >= 1e6 ? (v / 1e6).toFixed(2) + "M" : Math.round(v / 1000) + "k");
const fmtPct = (v, dp = 1) => (v * 100).toFixed(dp) + "%";
const fmtDelta = (v, dp = 1) => (v >= 0 ? "+" : "") + (v * 100).toFixed(dp) + "%";
const fmtTHB = (v) => "฿" + Math.round(v).toLocaleString("en-US");
const grow = (now, prev) => now / prev - 1;

function header(slide, section, title, subtitle) {
  slide.background = { color: "FFFFFF" };
  slide.addText(section.toUpperCase(), { x: M, y: 0.32, w: 9, h: 0.3, margin: 0,
    fontFace: F, fontSize: 11, color: GOLD, bold: true, charSpacing: 2 });
  slide.addText(title, { x: M, y: 0.55, w: W - 2 * M, h: 0.62, margin: 0,
    fontFace: F, fontSize: 28, color: NAVY, bold: true });
  if (subtitle) slide.addText(subtitle, { x: M, y: 1.16, w: W - 2 * M, h: 0.32,
    margin: 0, fontFace: F, fontSize: 12.5, color: MUT });
  slide.addText("Mitsui Thailand Portfolio · H1 2026", { x: W - 4.4, y: 0.34, w: 3.8, h: 0.3,
    margin: 0, align: "right", fontFace: F, fontSize: 9, color: MUT });
}

function tile(slide, x, y, w, h, label, big, sub, subColor) {
  slide.addShape("roundRect", { x, y, w, h, rectRadius: 0.06, fill: { color: PANEL },
    line: { color: "E2E7F2", width: 0.75 } });
  slide.addText(label, { x: x + 0.16, y: y + 0.1, w: w - 0.32, h: 0.3, margin: 0,
    fontFace: F, fontSize: 10.5, color: MUT, bold: true });
  slide.addText(big, { x: x + 0.16, y: y + 0.34, w: w - 0.32, h: 0.52, margin: 0,
    fontFace: F, fontSize: 24, color: NAVY, bold: true });
  if (sub) slide.addText(sub, { x: x + 0.16, y: y + 0.84, w: w - 0.32, h: h - 0.9,
    margin: 0, fontFace: F, fontSize: 10.5, color: subColor || MUT });
}

function note(slide, x, y, w, text) {
  slide.addText(text, { x, y, w, h: 0.48, margin: 0, fontFace: F, fontSize: 10.5,
    color: MUT, italic: true });
}

const axStyle = {
  catAxisLabelColor: MUT, catAxisLabelFontSize: 10, catAxisLabelFontFace: F,
  valAxisLabelColor: MUT, valAxisLabelFontSize: 9.5, valAxisLabelFontFace: F,
  valGridLine: { color: "E5E8EF", size: 0.5 }, catGridLine: { style: "none" },
  legendPos: "b", legendFontSize: 10, legendFontFace: F, legendColor: INK,
  chartColors: [NAVY],
};

function divider(num, title, items) {
  const s = pres.addSlide();
  s.background = { color: NAVY };
  s.addText(num, { x: M, y: 2.15, w: 2.2, h: 1.7, margin: 0, fontFace: F,
    fontSize: 96, color: GOLD, bold: true });
  s.addText(title, { x: 2.75, y: 2.5, w: 9.6, h: 1.0, margin: 0, fontFace: F,
    fontSize: 40, color: "FFFFFF", bold: true });
  if (items) s.addText(items, { x: 2.78, y: 3.55, w: 9.2, h: 0.9, margin: 0,
    fontFace: F, fontSize: 14, color: "AFC0E8" });
  s.addText("Mitsui Thailand Portfolio · H1 2026", { x: M, y: H - 0.75, w: 6, h: 0.3,
    margin: 0, fontFace: F, fontSize: 10, color: "7C8BC0" });
}

// ============================================================
// 1. TITLE
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: NAVY };
  s.addText("MITSUI THAILAND HOTEL PORTFOLIO", { x: M, y: 2.15, w: 12, h: 0.4,
    margin: 0, fontFace: F, fontSize: 15, color: GOLD, bold: true, charSpacing: 3 });
  s.addText("Performance Report — H1 2026", { x: M, y: 2.55, w: 12.2, h: 1.05,
    margin: 0, fontFace: F, fontSize: 48, color: "FFFFFF", bold: true });
  s.addText("Tourist arrivals · Market (STR) · Portfolio performance vs budget · Segmentation & nationality",
    { x: M, y: 3.7, w: 11.5, h: 0.4, margin: 0, fontFace: F, fontSize: 15, color: "AFC0E8" });
  s.addText("Somerset Rama 9 (SR9)  ·  Ascott Embassy Sathorn (AES)  ·  lyf Sukhumvit 8 (LYF)  ·  Somerset Pattaya (SP)",
    { x: M, y: 6.35, w: 11.5, h: 0.35, margin: 0, fontFace: F, fontSize: 12, color: "7C8BC0" });
  s.addText("Data through June 2026 (arrivals & STR through May) · Prepared July 2026",
    { x: M, y: 6.72, w: 11.5, h: 0.35, margin: 0, fontFace: F, fontSize: 11, color: "7C8BC0" });
}

// ============================================================
// 2. EXECUTIVE SUMMARY
// ============================================================
const A = data.arrivals;
const ytdOf = (arr) => { // YTD aggregates from portfolio month dicts
  const sold = arr.reduce((t, m) => t + m.sold, 0);
  const ra = arr.reduce((t, m) => t + m.ra, 0);
  const rev = arr.reduce((t, m) => t + m.revenue, 0);
  return { occ: sold / ra, adr: rev / sold, revpar: rev / ra, revenue: rev };
};
const PF = { act: ytdOf(data.portfolio.act), bg: ytdOf(data.portfolio.bg), ly: ytdOf(data.portfolio.ly) };
{
  const s = pres.addSlide();
  header(s, "Executive summary", "H1 2026 at a glance",
    "Arrivals YTD = Jan–May vs same period 2025 · Portfolio YTD = Jan–Jun vs budget");
  const tw = (W - 2 * M - 0.6) / 3, th = 1.42;
  const g = (o) => (o >= 0 ? GOOD : BAD);
  const r1 = [
    ["THAILAND ARRIVALS (MOTS)", fmtM(A.mots.ytd["2026"]), fmtDelta(grow(A.mots.ytd["2026"], A.mots.ytd["2025"])) + " vs 2025 YTD"],
    ["CHINESE ARRIVALS", fmtM(A.chinese.ytd["2026"]), fmtDelta(grow(A.chinese.ytd["2026"], A.chinese.ytd["2025"])) + " vs 2025 · still " + fmtPct(A.chinese.ytd["2026"] / A.chinese.ytd["2019"], 0) + " of 2019"],
    ["INDIA ARRIVALS", fmtM(A.india.ytd["2026"]), fmtDelta(grow(A.india.ytd["2026"], A.india.ytd["2025"])) + " vs 2025 — record high"],
  ];
  const r2 = [
    ["PORTFOLIO OCCUPANCY YTD", fmtPct(PF.act.occ), fmtDelta(PF.act.occ - PF.bg.occ, 1).replace("%", " pts") + " vs budget · " + fmtDelta(PF.act.occ - PF.ly.occ, 1).replace("%", " pts") + " vs LY"],
    ["PORTFOLIO ADR YTD", fmtTHB(PF.act.adr), fmtDelta(grow(PF.act.adr, PF.bg.adr)) + " vs budget · " + fmtDelta(grow(PF.act.adr, PF.ly.adr)) + " vs LY"],
    ["PORTFOLIO REVPAR YTD", fmtTHB(PF.act.revpar), fmtDelta(grow(PF.act.revpar, PF.bg.revpar)) + " vs budget · " + fmtDelta(grow(PF.act.revpar, PF.ly.revpar)) + " vs LY"],
  ];
  r1.forEach((t, i) => tile(s, M + i * (tw + 0.3), 1.72, tw, th, t[0], t[1], t[2],
    i === 0 ? BAD : GOOD));
  r2.forEach((t, i) => tile(s, M + i * (tw + 0.3), 3.3, tw, th, t[0], t[1], t[2],
    [PF.act.occ - PF.bg.occ, grow(PF.act.adr, PF.bg.adr), grow(PF.act.revpar, PF.bg.revpar)][i] >= 0 ? GOOD : BAD));
  s.addText([
    { text: "Key messages", options: { bold: true, color: NAVY, fontSize: 14, breakLine: true, paraSpaceAfter: 6 } },
    { text: "Market demand is mixed: total arrivals are slightly below last year, but China (+18%), India (+8%) and long-haul are holding the recovery together — Middle East is down on the war effect.", options: { bullet: true, breakLine: true, paraSpaceAfter: 4 } },
    { text: "Portfolio RevPAR is " + fmtDelta(grow(PF.act.revpar, PF.bg.revpar)) + " vs budget on " + fmtPct(PF.act.occ) + " occupancy — rate (ADR, −3.4%) is the main gap; occupancy is close to plan.", options: { bullet: true, breakLine: true, paraSpaceAfter: 4 } },
    { text: "The market is softening — Bangkok and Pattaya compset RevPAR are both below last year — and the portfolio is holding up slightly better than the market (−0.6% vs LY).", options: { bullet: true } },
  ], { x: M, y: 4.95, w: W - 2 * M, h: 1.9, margin: 0, fontFace: F, fontSize: 12.5, color: INK });
}

// ============================================================
// SECTION 1 — ARRIVALS
// ============================================================
divider("1", "Thailand Tourist Arrivals",
  "Total (MOTS) · Chinese · India · Middle East · Long-haul (EU + USA)");

const fmtN = (v) => (v == null ? "\u2014" : Math.round(v).toLocaleString("en-US"));

// Monthly numbers table (years as rows), like the Summary sheet
function monthTable(s, x, y, w, series, nMonths, opts) {
  opts = opts || {};
  const fs = opts.fontSize || 9.5, rh = opts.rowH || 0.26;
  const zebra = (i) => ({ color: i % 2 ? PANEL : "FFFFFF" });
  const hdrCell = (t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: NAVY }, align: "center" } });
  const rows = [[hdrCell(opts.corner || "")].concat(
    MONTHS.slice(0, nMonths).map(hdrCell), [hdrCell(opts.totalLabel || "YTD")])];
  series.forEach((sr, i) => {
    const vals = sr.values.slice(0, nMonths);
    const ytd = sr.total != null ? sr.total : vals.reduce((t, v) => t + (v || 0), 0);
    rows.push([{ text: sr.name, options: { bold: true, color: INK, fill: zebra(i) } }].concat(
      vals.map((v) => ({ text: fmtN(v), options: { color: INK, align: "right", fill: zebra(i) } })),
      [{ text: fmtN(ytd), options: { bold: true, color: NAVY, align: "right", fill: zebra(i) } }]));
  });
  const y26 = series.find((sr) => sr.name === "2026"), y25 = series.find((sr) => sr.name === "2025");
  if (y26 && y25) {
    const n26 = y26.values.filter((v) => v != null).length;
    const d = MONTHS.slice(0, nMonths).map((_, i) =>
      (y26.values[i] != null && y25.values[i] ? y26.values[i] / y25.values[i] - 1 : null));
    const s26 = y26.values.slice(0, n26).reduce((t, v) => t + (v || 0), 0);
    const s25 = y25.values.slice(0, n26).reduce((t, v) => t + (v || 0), 0);
    const dy = s25 ? s26 / s25 - 1 : null;
    const dCell = (v, bold) => ({ text: v == null ? "\u2014" : fmtDelta(v, bold ? 1 : 0),
      options: { bold: !!bold, align: "right", color: v == null ? MUT : (v >= 0 ? GOOD : BAD), fill: { color: "FFFFFF" } } });
    rows.push([{ text: "26 vs 25", options: { bold: true, color: MUT, fill: { color: "FFFFFF" } } }]
      .concat(d.map((v) => dCell(v, false)), [dCell(dy, true)]));
  }
  const lw = opts.labelW || 1.0, tw = opts.totalW || 1.1;
  const mw = (w - lw - tw) / nMonths;
  s.addTable(rows, { x, y, w, colW: [lw].concat(Array(nMonths).fill(mw), [tw]),
    fontFace: F, fontSize: fs, border: { type: "solid", color: "E2E7F2", pt: 0.5 },
    rowH: rh, valign: "middle" });
}

function arrivalSlide(title, subtitle, series, ytdTiles, noteText, valFmt) {
  const s = pres.addSlide();
  header(s, "Section 1 · Tourist arrivals", title, subtitle);
  const nCats = Math.max(...series.map((x) => x.values.length));
  const withTable = nCats <= 6;
  s.addChart(pres.ChartType.bar, series.map((x) => ({
    name: x.name, labels: MONTHS.slice(0, nCats), values: x.values,
  })), Object.assign({}, axStyle, {
    x: M, y: 1.75, w: 8.2, h: withTable ? 3.35 : 4.9, barGapWidthPct: 60,
    chartColors: series.map((x) => x.color),
    valAxisNumFmt: valFmt || "#,##0,K", showLegend: true, legendPos: "b",
  }));
  if (withTable) monthTable(s, M, 5.22, 8.2, series, nCats);
  let ty = 1.85;
  ytdTiles.forEach((t) => { tile(s, 9.15, ty, 3.55, 1.32, t[0], t[1], t[2], t[3]); ty += 1.48; });
  if (noteText) note(s, M, withTable ? 6.98 : 6.75, 11.9, noteText);
  return s;
}

// MOTS
arrivalSlide("Total international arrivals (MOTS)", "Monthly arrivals, all nationalities — 2026 data through May",
  [{ name: "2024", values: A.mots["2024"], color: ICE },
   { name: "2025", values: A.mots["2025"], color: SKY },
   { name: "2026", values: A.mots["2026"], color: NAVY }],
  [["YTD JAN–MAY 2026", fmtM(A.mots.ytd["2026"]), fmtDelta(grow(A.mots.ytd["2026"], A.mots.ytd["2025"])) + " vs 2025 (" + fmtM(A.mots.ytd["2025"]) + ")", BAD],
   ["YTD JAN–MAY 2025", fmtM(A.mots.ytd["2025"]), fmtDelta(grow(A.mots.ytd["2025"], A.mots.ytd["2024"])) + " vs 2024", MUT],
   ["VS PRE-COVID (2019)", fmtPct(A.mots.ytd["2026"] / 16746308, 0), "of Jan–May 2019 level", MUT]],
  "Source: Ministry of Tourism & Sports (MOTS). 2026 slightly below 2025 — January was weak (−12% vs Jan 2025), February–May broadly flat.");

// MOTS monthly detail table (like the Summary sheet)
{
  const s = pres.addSlide();
  header(s, "Section 1 · Tourist arrivals", "Total arrivals (MOTS) — monthly numbers by year",
    "International arrivals by month, 2018\u20132026 \u00b7 unit: persons \u00b7 same layout as the Summary sheet");
  const series = Object.keys(A.mots_full).map((y) => ({
    name: y, values: A.mots_full[y].m, total: A.mots_full[y].total }));
  monthTable(s, M, 1.8, W - 2 * M, series, 12,
    { fontSize: 8, rowH: 0.31, labelW: 0.8, totalW: 1.05, corner: "Year", totalLabel: "Total" });
  note(s, M, 5.65, 11.9,
    "2026 total = Jan\u2013May. \u201826 vs 25\u2019 row compares each month; the Total cell compares Jan\u2013May of both years (\u22122.3%). COVID years 2020\u20132022 shown for context.");
}

// Chinese
arrivalSlide("Chinese arrivals", "Jan–May by year — 2019 shown as pre-COVID reference",
  [{ name: "2019", values: A.chinese["2019"], color: GREY },
   { name: "2024", values: A.chinese["2024"], color: ICE },
   { name: "2025", values: A.chinese["2025"], color: SKY },
   { name: "2026", values: A.chinese["2026"], color: NAVY }],
  [["YTD JAN–MAY 2026", fmtM(A.chinese.ytd["2026"]), fmtDelta(grow(A.chinese.ytd["2026"], A.chinese.ytd["2025"])) + " vs 2025", GOOD],
   ["VS 2019 LEVEL", fmtPct(A.chinese.ytd["2026"] / A.chinese.ytd["2019"], 0), "of pre-COVID Jan–May", MUT],
   ["VS 2024 LEVEL", fmtPct(A.chinese.ytd["2026"] / A.chinese.ytd["2024"], 0), "of Jan–May 2024", MUT]],
  "Chinese demand is recovering vs a weak 2025 (safety concerns faded) but is still only ~48% of pre-COVID volume.");

// India
arrivalSlide("India arrivals", "Jan–May by year — India keeps setting record highs",
  [{ name: "2024", values: A.india["2024"], color: ICE },
   { name: "2025", values: A.india["2025"], color: SKY },
   { name: "2026", values: A.india["2026"], color: NAVY }],
  [["YTD JAN–MAY 2026", fmtM(A.india.ytd["2026"]), fmtDelta(grow(A.india.ytd["2026"], A.india.ytd["2025"])) + " vs 2025", GOOD],
   ["YTD JAN–MAY 2025", fmtM(A.india.ytd["2025"]), fmtDelta(grow(A.india.ytd["2025"], A.india.ytd["2024"])) + " vs 2024", MUT],
   ["TREND", "3rd year", "of consecutive YTD growth", GOOD]],
  "India is the strongest growth market: +24% in January alone; Apr–May flat against a high 2025 base.");

// Middle East
arrivalSlide("Middle East arrivals", "Jan–May by year — 2019 shown as pre-COVID reference",
  [{ name: "2019", values: A.mideast["2019"], color: GREY },
   { name: "2024", values: A.mideast["2024"], color: ICE },
   { name: "2025", values: A.mideast["2025"], color: SKY },
   { name: "2026", values: A.mideast["2026"], color: NAVY }],
  [["YTD JAN–MAY 2026", fmtM(A.mideast.ytd["2026"]), fmtDelta(grow(A.mideast.ytd["2026"], A.mideast.ytd["2025"])) + " vs 2025", BAD],
   ["VS 2019 LEVEL", fmtPct(A.mideast.ytd["2026"] / A.mideast.ytd["2019"], 0), "of pre-COVID Jan–May", MUT],
   ["WORST MONTH", "Apr −57%", "vs April 2025", BAD]],
  "Middle East volumes dropped sharply from February — regional war effect. Watch for recovery in H2.", "#,##0");

// Longhaul
{
  const s = pres.addSlide();
  header(s, "Section 1 · Tourist arrivals", "Long-haul arrivals (EU + USA)",
    "Monthly arrivals Jan–May — Europe (left) and USA (right)");
  const mk = (key, x, ttl) => s.addChart(pres.ChartType.bar, [
    { name: "2024", labels: MONTHS.slice(0, 5), values: A[key]["2024"] },
    { name: "2025", labels: MONTHS.slice(0, 5), values: A[key]["2025"] },
    { name: "2026", labels: MONTHS.slice(0, 5), values: A[key]["2026"] },
  ], Object.assign({}, axStyle, {
    x, y: 2.0, w: 4.1, h: 3.5, chartColors: [ICE, SKY, NAVY],
    showLegend: true, legendPos: "b", valAxisNumFmt: "#,##0,K",
    showTitle: true, title: ttl, titleFontSize: 13, titleColor: NAVY, titleFontFace: F,
  }));
  mk("lh_eu", M, "Europe");
  mk("lh_usa", 4.95, "USA");
  const c = A.lh_all;
  tile(s, 9.35, 2.2, 3.35, 1.32, "LONG-HAUL YTD 2026", fmtM(c.ytd["2026"]),
    fmtDelta(grow(c.ytd["2026"], c.ytd["2025"])) + " vs 2025", GOOD);
  tile(s, 9.35, 3.68, 3.35, 1.32, "EUROPE YTD", fmtM(A.lh_eu["2026"].reduce((a, b) => a + b, 0)),
    fmtDelta(grow(A.lh_eu["2026"].reduce((a, b) => a + b, 0), A.lh_eu["2025"].reduce((a, b) => a + b, 0))) + " vs 2025", GOOD);
  tile(s, 9.35, 5.16, 3.35, 1.32, "USA YTD", fmtM(A.lh_usa["2026"].reduce((a, b) => a + b, 0)),
    fmtDelta(grow(A.lh_usa["2026"].reduce((a, b) => a + b, 0), A.lh_usa["2025"].reduce((a, b) => a + b, 0))) + " vs 2025", GOOD);
  monthTable(s, M, 5.68, 8.45, [
    { name: "2024", values: A.lh_all["2024"] },
    { name: "2025", values: A.lh_all["2025"] },
    { name: "2026", values: A.lh_all["2026"] },
  ], 5, { corner: "EU+USA", rowH: 0.25, fontSize: 9 });
  note(s, 9.35, 5.75, 3.4, "Long-haul stable at record levels — strong Jan–Feb offset a softer April.");
}

// ============================================================
// SECTION 2 — STR
// ============================================================
divider("2", "Competitive Market (STR)", "Bangkok compset · Pattaya compset — Occ, ADR, RevPAR");

function strSlide(cityKey, cityName, subtitle, noteText) {
  const s = pres.addSlide();
  header(s, "Section 2 · Market (STR)", cityName + " compset — Occ / ADR / RevPAR", subtitle);
  const d = data.str[cityKey];
  const defs = [
    ["Occupancy (%)", "occ", "#,##0\"%\"", (v) => v],
    ["ADR (THB)", "adr", "#,##0", (v) => v],
    ["RevPAR (THB)", "revpar", "#,##0", (v) => v],
  ];
  defs.forEach((def, i) => {
    const [ttl, key, fmt] = def;
    s.addChart(pres.ChartType.bar, [
      { name: "2025", labels: MONTHS.slice(0, 5), values: d[key]["2025"].slice(0, 5) },
      { name: "2026", labels: MONTHS.slice(0, 5), values: d[key]["2026"] },
    ], Object.assign({}, axStyle, {
      x: M + i * 4.15, y: 2.0, w: 3.95, h: 4.35, chartColors: [SKY, NAVY],
      showLegend: true, legendPos: "b", valAxisNumFmt: fmt,
      showTitle: true, title: ttl, titleFontSize: 13, titleColor: NAVY, titleFontFace: F,
    }));
  });
  // YTD deltas Jan-May
  const avg = (a) => a.slice(0, 5).reduce((x, y) => x + y, 0) / 5;
  const dOcc = avg(d.occ["2026"]) - avg(d.occ["2025"]);
  const dAdr = grow(avg(d.adr["2026"]), avg(d.adr["2025"]));
  const dRvp = grow(avg(d.revpar["2026"]), avg(d.revpar["2025"]));
  s.addText("Jan–May avg vs 2025:   Occ " + (dOcc >= 0 ? "+" : "") + dOcc.toFixed(1) + " pts   ·   ADR " +
    fmtDelta(dAdr) + "   ·   RevPAR " + fmtDelta(dRvp),
    { x: M, y: 6.45, w: 11.9, h: 0.3, margin: 0, fontFace: F, fontSize: 12.5, bold: true,
      color: dRvp >= 0 ? GOOD : BAD });
  note(s, M, 6.78, 11.9, noteText);
  return s;
}
strSlide("bangkok", "Bangkok", "Monthly compset performance, 2025 vs 2026 (Jan–May) — ADR excludes breakfast",
  "Bangkok market: occupancy holding above last year, but rate is down — the market is buying occupancy with price.");
strSlide("pattaya", "Pattaya", "Monthly compset performance, 2025 vs 2026 (Jan–May)",
  "Pattaya market: strong high season (Jan–Feb) then softening from March; rates broadly below 2025.");

// ============================================================
// SECTION 3 — PERFORMANCE VS BUDGET
// ============================================================
divider("3", "Performance vs Budget", "Portfolio and each property — Occ, ADR, RevPAR · monthly & YTD");

function perfSlide(sectionTitle, subtitle, actArr, bgArr, lyArr, ytd, noteText) {
  const s = pres.addSlide();
  header(s, "Section 3 · Performance vs budget", sectionTitle, subtitle);
  const n = actArr.length;
  const lbl = MONTHS.slice(0, n);
  s.addChart([
    { type: pres.ChartType.bar,
      data: [{ name: "RevPAR 2026", labels: lbl, values: actArr.map((m) => m.revpar) }],
      options: { chartColors: [NAVY], barGapWidthPct: 80 } },
    { type: pres.ChartType.line,
      data: [{ name: "Budget", labels: lbl, values: bgArr.map((m) => m.revpar) },
             { name: "2025", labels: lbl, values: lyArr.map((m) => m.revpar) }],
      options: { chartColors: [GOLD, GREY], lineSize: 2.5, lineSmooth: false,
                 lineDataSymbol: "circle", lineDataSymbolSize: 5 } },
  ], Object.assign({}, axStyle, {
    x: M, y: 1.85, w: 7.9, h: 3.0, showLegend: true, legendPos: "b",
    valAxisNumFmt: "#,##0", showTitle: true, title: "Monthly RevPAR (THB) — actual vs budget vs 2025",
    titleFontSize: 12, titleColor: NAVY, titleFontFace: F,
  }));
  // 2026 monthly numbers table (like the Summary sheet)
  {
    const zebra = (i) => ({ color: i % 2 ? PANEL : "FFFFFF" });
    const hc = (t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: NAVY }, align: "center" } });
    const vRow = (label, vals, ytdText, f, i) => [
      { text: label, options: { bold: true, color: INK, fill: zebra(i) } }].concat(
      vals.map((v) => ({ text: f(v), options: { color: INK, align: "right", fill: zebra(i) } })),
      [{ text: ytdText, options: { bold: true, color: NAVY, align: "right", fill: zebra(i) } }]);
    const dRow = (label, vals, ytdV) => [
      { text: label, options: { bold: true, color: MUT, fill: { color: "FFFFFF" } } }].concat(
      vals.map((v) => ({ text: fmtDelta(v, 0), options: { align: "right", color: v >= 0 ? GOOD : BAD, fill: { color: "FFFFFF" } } })),
      [{ text: fmtDelta(ytdV, 1), options: { bold: true, align: "right", color: ytdV >= 0 ? GOOD : BAD, fill: { color: "FFFFFF" } } }]);
    const nf = (v) => Math.round(v).toLocaleString("en-US");
    const t = [[hc("2026")].concat(lbl.map(hc), [hc("YTD")])];
    t.push(vRow("Occupancy", actArr.map((m) => m.occ), fmtPct(ytd.act.occ), (v) => (v * 100).toFixed(1) + "%", 0));
    t.push(vRow("ADR (\u0e3f)", actArr.map((m) => m.adr), nf(ytd.act.adr), nf, 1));
    t.push(vRow("RevPAR (\u0e3f)", actArr.map((m) => m.revpar), nf(ytd.act.revpar), nf, 0));
    t.push(dRow("RevPAR vs BG", actArr.map((m, i) => m.revpar / bgArr[i].revpar - 1), grow(ytd.act.revpar, ytd.bg.revpar)));
    t.push(dRow("RevPAR vs 25", actArr.map((m, i) => m.revpar / lyArr[i].revpar - 1), grow(ytd.act.revpar, ytd.ly.revpar)));
    s.addTable(t, { x: M, y: 5.05, w: 7.9,
      colW: [1.45].concat(Array(n).fill(5.22 / n), [1.23]),
      fontFace: F, fontSize: 9, border: { type: "solid", color: "E2E7F2", pt: 0.5 },
      rowH: 0.25, valign: "middle" });
  }
  const rows = [
    ["OCCUPANCY · YTD", fmtPct(ytd.act.occ),
      fmtDelta(ytd.act.occ - ytd.bg.occ).replace("%", " pts") + " vs budget · " + fmtDelta(ytd.act.occ - ytd.ly.occ).replace("%", " pts") + " vs LY",
      ytd.act.occ >= ytd.bg.occ ? GOOD : BAD],
    ["ADR · YTD", fmtTHB(ytd.act.adr),
      fmtDelta(grow(ytd.act.adr, ytd.bg.adr)) + " vs budget · " + fmtDelta(grow(ytd.act.adr, ytd.ly.adr)) + " vs LY",
      ytd.act.adr >= ytd.bg.adr ? GOOD : BAD],
    ["REVPAR · YTD", fmtTHB(ytd.act.revpar),
      fmtDelta(grow(ytd.act.revpar, ytd.bg.revpar)) + " vs budget · " + fmtDelta(grow(ytd.act.revpar, ytd.ly.revpar)) + " vs LY",
      ytd.act.revpar >= ytd.bg.revpar ? GOOD : BAD],
  ];
  let ty = 2.0;
  rows.forEach((t) => { tile(s, 8.85, ty, 3.85, 1.38, t[0], t[1], t[2], t[3]); ty += 1.54; });
  if (noteText) note(s, M, 6.75, 11.9, noteText);
}

// portfolio
perfSlide("Portfolio — all four properties", "Weighted by room inventory (1,406 keys) · YTD = Jan–Jun 2026",
  data.portfolio.act, data.portfolio.bg, data.portfolio.ly, PF,
  "Portfolio RevPAR is ~4% behind budget, driven mainly by rate (ADR −3.4%); occupancy is near plan (−0.7 pts). Vs last year RevPAR is broadly flat. April was the softest month (Songkran).");

// per property
const propYtd = (p, pre) => {
  let sold = 0, ra = 0, rev = 0;
  for (let i = 0; i < 6; i++) {
    const occ = p[pre + "occ"][i], adr = p[pre + "adr"][i];
    const ra_i = p.rooms * DAYS[i];
    ra += ra_i; sold += occ * ra_i; rev += occ * ra_i * adr;
  }
  return { occ: sold / ra, adr: rev / sold, revpar: rev / ra };
};
const PROP_NOTES = {
  SR9: "SR9 runs close to budget all year; Feb outperformed strongly, May–Jun slightly behind on rate.",
  AES: "AES: February was exceptional (92% occupancy, well above budget); April–May fell behind budget in a softer Bangkok market.",
  LYF: "LYF: strong Q1 above budget, then April dropped sharply (occupancy −32 pts vs budget) — recovering since May.",
  SP: "SP: ahead of budget in Jan–Feb high season; Mar–Jun broadly on plan in a softer Pattaya market.",
};
for (const key of ["SR9", "AES", "LYF", "SP"]) {
  const p = data.perf[key];
  const act = [], bg = [], ly = [];
  for (let i = 0; i < 6; i++) {
    act.push({ occ: p.occ[i], adr: p.adr[i], revpar: p.revpar[i] });
    bg.push({ occ: p.bg_occ[i], adr: p.bg_adr[i], revpar: p.bg_revpar[i] });
    ly.push({ occ: p.ly_occ[i], adr: p.ly_adr[i], revpar: p.ly_revpar[i] });
  }
  perfSlide(p.name + " (" + key + ")", p.rooms + " keys · monthly + YTD Jan–Jun vs budget and last year",
    act, bg, ly, { act: propYtd(p, ""), bg: propYtd(p, "bg_"), ly: propYtd(p, "ly_") },
    PROP_NOTES[key]);
}

// ============================================================
// SECTION 4 — SEGMENTATION & NATIONALITY
// ============================================================
divider("4", "Segmentation & Nationality", "Room nights by segment and guest nationality — each property, monthly & YTD");

const SEG_SHORT = {
  "Corporate SS": "Corporate SS", "Corporate LS": "Corporate LS",
  "Corporate Short Stay": "Corporate SS", "Corporate Long Stay": "Corporate LS",
  "Online Business": "Online", "Online Business (Dynamic Rate)": "Online",
  "ASR": "ASR", "Wholesale": "Wholesale", "Wholesale (Static Rate)": "Wholesale",
  "Group Corporate": "Group Corp", "Group Leisure": "Group Leisure",
  "Group Series": "Group Series", "Corporate Group": "Corporate Group",
  "Wholesale Group": "Wholesale Grp", "Medical": "Medical",
  "Employee Travel": "Employee", "Corporate Group with Banque": "Corp Grp Banque",
};

for (const key of ["SR9", "AES", "LYF", "SP"]) {
  const p = data.perf[key];
  const segs = data.seg[key].filter((r) => r.ytd > 0);
  const s = pres.addSlide();
  header(s, "Section 4 · Segmentation", p.name + " (" + key + ") — room nights by segment",
    "Monthly mix Jan–Jun 2026 (left) and YTD share (right)");
  s.addChart(pres.ChartType.bar,
    segs.map((r, i) => ({ name: SEG_SHORT[r.label] || r.label, labels: MONTHS.slice(0, 6), values: r.m })),
    Object.assign({}, axStyle, {
      x: M, y: 1.9, w: 7.3, h: 4.7, barGrouping: "stacked", barGapWidthPct: 45,
      chartColors: SEG_COLORS.slice(0, segs.length),
      showLegend: true, legendPos: "b", valAxisNumFmt: "#,##0",
      showTitle: true, title: "Room nights per month (stacked)", titleFontSize: 13,
      titleColor: NAVY, titleFontFace: F,
    }));
  s.addChart(pres.ChartType.doughnut,
    [{ name: "YTD", labels: segs.map((r) => SEG_SHORT[r.label] || r.label),
       values: segs.map((r) => r.ytd) }],
    { x: 8.35, y: 1.9, w: 4.4, h: 4.7, chartColors: SEG_COLORS.slice(0, segs.length),
      showLegend: true, legendPos: "r", legendFontSize: 9.5, legendFontFace: F, legendColor: INK,
      dataLabelColor: "FFFFFF", dataLabelFontSize: 9, dataLabelFontFace: F,
      showPercent: true, holeSize: 55,
      showTitle: true, title: "YTD mix (" + segs.reduce((t, r) => t + r.ytd, 0).toLocaleString("en-US") + " RN)",
      titleFontSize: 13, titleColor: NAVY, titleFontFace: F });
  const top = segs.slice().sort((a, b) => b.ytd - a.ytd)[0];
  note(s, M, 6.72, 11.9, "Largest segment YTD: " + (SEG_SHORT[top.label] || top.label) + " (" +
    fmtPct(top.ytd / segs.reduce((t, r) => t + r.ytd, 0), 0) + " of room nights).");
}

// nationality slides
for (const key of ["SR9", "AES", "LYF", "SP"]) {
  const p = data.perf[key];
  const counts = data.nat_counts[key];
  const shares = data.nat[key];
  const s = pres.addSlide();
  header(s, "Section 4 · Nationality", p.name + " (" + key + ") — guest nationality",
    "Top 10 by room nights, YTD Jan–Apr 2026 vs 2025 (left) · monthly share of room nights (right)");
  s.addChart(pres.ChartType.bar, [
    { name: "2025", labels: counts.map((c) => c.label), values: counts.map((c) => c.y2025) },
    { name: "2026", labels: counts.map((c) => c.label), values: counts.map((c) => c.y2026) },
  ], Object.assign({}, axStyle, {
    x: M, y: 1.9, w: 6.1, h: 4.85, barDir: "bar", chartColors: [SKY, NAVY],
    showLegend: true, legendPos: "b", valAxisNumFmt: "#,##0",
    catAxisLabelFontSize: 9.5, barGapWidthPct: 40,
    showTitle: true, title: "Room nights YTD (Jan–Apr)", titleFontSize: 13,
    titleColor: NAVY, titleFontFace: F,
  }));
  // monthly share table (top 6 by avg share)
  const top6 = shares.filter((r) => r.avg != null).sort((a, b) => b.avg - a.avg).slice(0, 6);
  const trows = [[
    { text: "Share of RN", options: { bold: true, color: "FFFFFF", fill: { color: NAVY } } },
    ...MONTHS.slice(0, 6).map((m) => ({ text: m, options: { bold: true, color: "FFFFFF", fill: { color: NAVY }, align: "center" } })),
  ]];
  top6.forEach((r, ri) => {
    trows.push([
      { text: r.label, options: { bold: true, color: INK, fill: { color: ri % 2 ? PANEL : "FFFFFF" } } },
      ...r.m.map((v) => ({ text: v == null ? "—" : (v * 100).toFixed(0) + "%",
        options: { color: INK, align: "center", fill: { color: ri % 2 ? PANEL : "FFFFFF" } } })),
    ]);
  });
  s.addTable(trows, { x: 7.0, y: 2.15, w: 5.75, colW: [1.85, 0.65, 0.65, 0.65, 0.65, 0.65, 0.65],
    fontFace: F, fontSize: 10, border: { type: "solid", color: "E2E7F2", pt: 0.5 },
    rowH: 0.34, valign: "middle" });
  const winner = counts.slice().sort((a, b) => (b.y2026 - b.y2025) - (a.y2026 - a.y2025))[0];
  const loser = counts.slice().sort((a, b) => (a.y2026 - a.y2025) - (b.y2026 - b.y2025))[0];
  note(s, M, 6.85, 11.9, "Biggest YTD gain: " + winner.label + " (+" + (winner.y2026 - winner.y2025).toLocaleString("en-US") +
    " RN) · biggest decline: " + loser.label + " (" + (loser.y2026 - loser.y2025).toLocaleString("en-US") + " RN).");
}

// ============================================================
// SOURCES / NOTES
// ============================================================
{
  const s = pres.addSlide();
  header(s, "Appendix", "Sources & definitions");
  s.addText([
    { text: "Data sources", options: { bold: true, color: NAVY, fontSize: 14, breakLine: true, paraSpaceAfter: 6 } },
    { text: "Arrivals: Ministry of Tourism & Sports (MOTS) and AOT statistics, as compiled in the Segment Half-year workbook (tabs Arrival / Summary-1). 2026 data through May.", options: { bullet: true, breakLine: true, paraSpaceAfter: 4 } },
    { text: "STR / compset: Bangkok and Pattaya competitive-set Occ, ADR, RevPAR from the Compset tab (ADR excludes breakfast).", options: { bullet: true, breakLine: true, paraSpaceAfter: 4 } },
    { text: "Property performance: Summary tab of the reconciled workbook (Segment_Half_year_version_1_ALL-reconciled.xlsx) — actuals, budget (BG rows) and 2025. Segment data reconciled to each property's source system, July 2026.", options: { bullet: true, breakLine: true, paraSpaceAfter: 4 } },
    { text: "Segmentation & nationality: property tabs (2026 Jan–Jun) and Summary-1 top-10 nationality tables (Jan–Apr).", options: { bullet: true, breakLine: true, paraSpaceAfter: 10 } },
    { text: "Definitions & notes", options: { bold: true, color: NAVY, fontSize: 14, breakLine: true, paraSpaceAfter: 6 } },
    { text: "ADR = room revenue ÷ rooms sold; RevPAR = room revenue ÷ rooms available; Occ = rooms sold ÷ rooms available. YTD = Jan–Jun 2026 unless stated.", options: { bullet: true, breakLine: true, paraSpaceAfter: 4 } },
    { text: "Portfolio figures weight the four properties by room inventory (SR9 464 · AES 403 · SP 333 · LYF 206 = 1,406 keys).", options: { bullet: true, breakLine: true, paraSpaceAfter: 4 } },
    { text: "AES & SP June 2026 revenue estimated from Occ × ADR (source revenue available through May).", options: { bullet: true, breakLine: true, paraSpaceAfter: 4 } },
    { text: "All charts are native PowerPoint charts — right-click → Edit Data to change any number.", options: { bullet: true } },
  ], { x: M, y: 1.8, w: W - 2 * M, h: 5.4, margin: 0, fontFace: F, fontSize: 12, color: INK });
}

pres.writeFile({ fileName: "/home/user/mae-workspace/output/Portfolio_Performance_Report_H1-2026.pptx" })
  .then(() => console.log("deck written"));
