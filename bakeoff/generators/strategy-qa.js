const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Q3 2026 Strategic Review — NovaCrest";

// ─── COLOR PALETTE ───────────────────────────────────────────────
const C = {
  navy:    "0F2044",   // primary dark bg
  blue:    "1A3A6B",   // secondary bg
  accent:  "2E7FD9",   // accent blue
  teal:    "0D9488",   // positive/green signal
  amber:   "D97706",   // warning
  red:     "DC2626",   // negative
  white:   "FFFFFF",
  offwhite:"F4F6FA",
  ltgray:  "E2E8F0",
  midgray: "94A3B8",
  darkgray:"334155",
  text:    "1E293B",
};

const W = 10, H = 5.625;

// ─── HELPERS ─────────────────────────────────────────────────────
function hdr(slide, title, sub) {
  // Top bar
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: 0.52, fill: { color: C.navy }, line: { color: C.navy } });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0.52, w: W, h: 0.03, fill: { color: C.accent }, line: { color: C.accent } });
  slide.addText(title, { x: 0.35, y: 0, w: 7.5, h: 0.52, fontSize: 17, bold: true, color: C.white, valign: "middle", margin: 0 });
  if (sub) slide.addText(sub, { x: 0, y: 0, w: W - 0.3, h: 0.52, fontSize: 10, color: "8BA7C7", valign: "middle", align: "right", margin: 0 });
  // Footer
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: H - 0.28, w: W, h: 0.28, fill: { color: C.navy }, line: { color: C.navy } });
  slide.addText("NovaCrest  |  Q3 2026 Board Review  |  CONFIDENTIAL", { x: 0.3, y: H - 0.28, w: 6, h: 0.28, fontSize: 7.5, color: "6B82A0", valign: "middle", margin: 0 });
  slide.addText("March 31, 2026", { x: 0, y: H - 0.28, w: W - 0.3, h: 0.28, fontSize: 7.5, color: "6B82A0", align: "right", valign: "middle", margin: 0 });
}

function kpiCard(slide, x, y, w, h, label, value, sub, valueColor, bgColor) {
  bgColor = bgColor || C.offwhite;
  slide.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: bgColor }, line: { color: C.ltgray, width: 0.8 } });
  slide.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.07, h, fill: { color: valueColor || C.accent }, line: { color: valueColor || C.accent } });
  slide.addText(value, { x: x + 0.14, y: y + 0.08, w: w - 0.18, h: h * 0.48, fontSize: 21, bold: true, color: valueColor || C.accent, valign: "middle", margin: 0 });
  slide.addText(label, { x: x + 0.14, y: y + h * 0.52, w: w - 0.18, h: h * 0.26, fontSize: 8.5, bold: true, color: C.darkgray, margin: 0 });
  if (sub) slide.addText(sub, { x: x + 0.14, y: y + h * 0.74, w: w - 0.18, h: h * 0.22, fontSize: 9, color: C.midgray, margin: 0 });
}

function sectionLabel(slide, x, y, w, text, color) {
  color = color || C.accent;
  slide.addShape(pres.shapes.RECTANGLE, { x, y, w, h: 0.22, fill: { color }, line: { color } });
  slide.addText(text, { x: x + 0.08, y, w: w - 0.1, h: 0.22, fontSize: 8, bold: true, color: C.white, valign: "middle", margin: 0 });
}

function tbl(slide, rows, opts) {
  // rows[0] = header
  const colW = opts.colW;
  const x = opts.x, y = opts.y, h = opts.h || 0.28;
  const hdrFill = opts.hdrFill || C.navy;
  const altFill = opts.altFill || C.offwhite;

  rows.forEach((row, ri) => {
    let cx = x;
    row.forEach((cell, ci) => {
      const cw = colW[ci];
      const isHdr = ri === 0;
      const bg = isHdr ? hdrFill : (ri % 2 === 0 ? C.white : altFill);
      slide.addShape(pres.shapes.RECTANGLE, { x: cx, y: y + ri * h, w: cw, h, fill: { color: bg }, line: { color: C.ltgray, width: 0.5 } });
      const cellText = String(cell.text !== undefined ? cell.text : cell);
      const align = cell.align || (ci === 0 ? "left" : "center");
      const fc = isHdr ? C.white : (cell.color || C.text);
      const fs = opts.fontSize || 9;
      slide.addText(cellText, { x: cx + 0.06, y: y + ri * h, w: cw - 0.08, h, fontSize: fs, color: fc, bold: isHdr || cell.bold, align, valign: "middle", margin: 0 });
      cx += cw;
    });
  });
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 1 — TITLE + EXECUTIVE SUMMARY
// ═══════════════════════════════════════════════════════════════
{
  const slide = pres.addSlide();
  // Full dark bg
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: H, fill: { color: C.navy }, line: { color: C.navy } });
  // Accent stripe left
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.12, h: H, fill: { color: C.accent }, line: { color: C.accent } });
  // Right panel
  slide.addShape(pres.shapes.RECTANGLE, { x: 5.8, y: 0, w: 4.2, h: H, fill: { color: C.blue }, line: { color: C.blue } });

  // Title block
  slide.addText("Q3 2026 STRATEGIC REVIEW", { x: 0.35, y: 0.5, w: 5.2, h: 0.48, fontSize: 22, bold: true, color: C.white, charSpacing: 2, margin: 0 });
  slide.addText("NovaCrest  ·  Board of Directors", { x: 0.35, y: 1.02, w: 5.2, h: 0.3, fontSize: 13, color: "8BA7C7", margin: 0 });
  slide.addText("March 31, 2026  ·  Confidential", { x: 0.35, y: 1.32, w: 5.2, h: 0.25, fontSize: 10, color: "5C7A9B", margin: 0 });

  // Company snapshot strip
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.35, y: 1.72, w: 5.2, h: 0.03, fill: { color: C.accent }, line: { color: C.accent } });

  const snapItems = [
    ["ARR", "$15.6M"], ["Stage", "Series B"], ["Headcount", "112"], ["Founded", "2021, SF"],
  ];
  snapItems.forEach((item, i) => {
    const sx = 0.35 + i * 1.3;
    slide.addShape(pres.shapes.RECTANGLE, { x: sx, y: 1.78, w: 1.25, h: 0.68, fill: { color: "0D1E3A" }, line: { color: C.accent, width: 0.5 } });
    slide.addText(item[1], { x: sx + 0.07, y: 1.82, w: 1.12, h: 0.3, fontSize: 13, bold: true, color: C.accent, margin: 0 });
    slide.addText(item[0], { x: sx + 0.07, y: 2.1, w: 1.12, h: 0.3, fontSize: 8, color: C.midgray, margin: 0 });
  });

  // CEO / Board line
  slide.addText("CEO: Sarah Chen  ·  CFO: David Park  ·  CTO: Marcus Williams  ·  VP Sales: Rachel Torres", {
    x: 0.35, y: 2.58, w: 5.2, h: 0.26, fontSize: 9, color: "5C7A9B", margin: 0
  });
  slide.addText("Board: Jeff Blackwell (Insight), Priya Sharma (Accel), Tom Nguyen (Independent)", {
    x: 0.35, y: 2.84, w: 5.2, h: 0.26, fontSize: 9, color: "5C7A9B", margin: 0
  });

  // ── RIGHT PANEL ──
  // Wins
  slide.addShape(pres.shapes.RECTANGLE, { x: 5.95, y: 0.28, w: 3.8, h: 0.22, fill: { color: C.teal }, line: { color: C.teal } });
  slide.addText("✓  Q3 KEY WINS", { x: 6.05, y: 0.28, w: 3.6, h: 0.22, fontSize: 8.5, bold: true, color: C.white, valign: "middle", margin: 0 });
  const wins = [
    "ARR +10.6% QoQ → $15.6M  (+4.2% vs plan)",
    "Pipeline $8.6M — best quarter ever; win rate 28.4%",
    "VP CS + Head of Partnerships + 2 ML PhDs hired",
  ];
  wins.forEach((w, i) => {
    slide.addText("→  " + w, { x: 6.05, y: 0.54 + i * 0.32, w: 3.7, h: 0.29, fontSize: 8.5, color: C.white, margin: 0 });
  });

  // Concerns
  slide.addShape(pres.shapes.RECTANGLE, { x: 5.95, y: 1.56, w: 3.8, h: 0.22, fill: { color: C.amber }, line: { color: C.amber } });
  slide.addText("⚠  CONCERNS", { x: 6.05, y: 1.56, w: 3.6, h: 0.22, fontSize: 8.5, bold: true, color: C.white, valign: "middle", margin: 0 });
  const concerns = [
    "Churn spike: $500K in Q3 (+66% QoQ); 58% preventable",
    "Competitive: DataForge SMB module; Zenith fundraise + hiring",
  ];
  concerns.forEach((c, i) => {
    slide.addText("→  " + c, { x: 6.05, y: 1.82 + i * 0.32, w: 3.7, h: 0.29, fontSize: 8.5, color: "FFD580", margin: 0 });
  });

  // Board Decision
  slide.addShape(pres.shapes.RECTANGLE, { x: 5.95, y: 2.56, w: 3.8, h: 0.22, fill: { color: C.red }, line: { color: C.red } });
  slide.addText("●  BOARD DECISION REQUIRED", { x: 6.05, y: 2.56, w: 3.6, h: 0.22, fontSize: 8.5, bold: true, color: C.white, valign: "middle", margin: 0 });
  slide.addText("→  Approve $2.5M CS investment to address preventable churn\n→  Approve usage-based SMB pricing tier ($499/mo)\n→  Discuss Series C timing (Q2 2027 vs $30M ARR milestone)", {
    x: 6.05, y: 2.82, w: 3.7, h: 0.72, fontSize: 8.5, color: "FCA5A5", margin: 0
  });

  // Agenda
  slide.addShape(pres.shapes.RECTANGLE, { x: 5.95, y: 3.7, w: 3.8, h: 0.22, fill: { color: "162A4A" }, line: { color: C.accent, width: 0.5 } });
  slide.addText("AGENDA", { x: 6.05, y: 3.7, w: 3.6, h: 0.22, fontSize: 8, bold: true, color: C.accent, valign: "middle", margin: 0 });
  const agenda = ["1  Revenue Dashboard", "2  GTM + Customer Health", "3  Product + Team", "4  Financial Outlook + Competitive", "5  Board Asks"];
  agenda.forEach((a, i) => {
    slide.addText(a, { x: 6.05, y: 3.96 + i * 0.26, w: 3.7, h: 0.24, fontSize: 8, color: C.midgray, margin: 0 });
  });

  // Footer
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: H - 0.28, w: W, h: 0.28, fill: { color: "080F1C" }, line: { color: "080F1C" } });
  slide.addText("NovaCrest  |  Q3 2026 Board Review  |  CONFIDENTIAL", { x: 0.3, y: H - 0.28, w: 9.4, h: 0.28, fontSize: 7.5, color: "3D5470", valign: "middle", margin: 0 });
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 2 — REVENUE DASHBOARD
// ═══════════════════════════════════════════════════════════════
{
  const slide = pres.addSlide();
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: H, fill: { color: C.offwhite }, line: { color: C.offwhite } });
  hdr(slide, "SLIDE 2  ·  Revenue Dashboard", "Q3 2026 Actuals vs Plan");

  const TOP = 0.65;

  // ── KPI Cards Row ──
  const cards = [
    { label: "ARR", value: "$15.6M", sub: "+10.6% QoQ  ·  +4.2% vs plan", vc: C.accent },
    { label: "Net Revenue Retention", value: "118%", sub: "Target 120%+  ·  ↓ from 121% (churn)", vc: C.amber },
    { label: "Gross Margin", value: "78.2%", sub: "Target 80%+  ·  Improving ↑", vc: C.teal },
    { label: "Magic Number", value: "1.12x", sub: "Q2: 0.98  ·  Target >0.75  ·  Strong", vc: C.teal },
    { label: "CAC Payback", value: "13.2 mo", sub: "Target <16 mo  ·  Improving ↑", vc: C.teal },
  ];
  const cw = (W - 0.3) / cards.length;
  cards.forEach((c, i) => {
    kpiCard(slide, 0.15 + i * cw, TOP, cw - 0.08, 0.82, c.label, c.value, c.sub, c.vc);
  });

  // ── ARR Trend Chart ──
  sectionLabel(slide, 0.15, TOP + 0.92, 3.5, "ARR TREND  ($M)  ·  Q1 2025 – Q3 2026");

  const arrData = [{ name: "ARR ($M)", labels: ["Q1'25","Q2'25","Q3'25","Q4'25","Q1'26","Q2'26","Q3'26"], values: [7.2, 8.5, 9.8, 11.4, 12.8, 14.1, 15.6] }];
  slide.addChart(pres.charts.LINE, arrData, {
    x: 0.15, y: TOP + 1.18, w: 3.55, h: 2.32,
    chartColors: [C.accent],
    lineSize: 2.5,
    lineSmooth: false,
    showValue: true,
    dataLabelFontSize: 8,
    dataLabelColor: C.darkgray,
    chartArea: { fill: { color: C.white }, border: { color: C.ltgray, width: 0.5 } },
    catAxisLabelColor: C.midgray,
    valAxisLabelColor: C.midgray,
    catAxisLabelFontSize: 8,
    valAxisLabelFontSize: 8,
    valGridLine: { color: C.ltgray, size: 0.5 },
    catGridLine: { style: "none" },
    showLegend: false,
  });

  // ── Quarterly Revenue Breakdown Chart ──
  sectionLabel(slide, 3.85, TOP + 0.92, 2.9, "QUARTERLY REVENUE  ($M)");
  const revData = [
    { name: "Subscription", labels: ["Q1'26","Q2'26","Q3'26"], values: [3.1, 3.4, 3.8] },
    { name: "Services",     labels: ["Q1'26","Q2'26","Q3'26"], values: [0.3, 0.3, 0.3] },
  ];
  slide.addChart(pres.charts.BAR, revData, {
    x: 3.85, y: TOP + 1.18, w: 2.9, h: 2.32,
    barDir: "col",
    barGrouping: "stacked",
    chartColors: [C.accent, "93C5FD"],
    showValue: true,
    dataLabelFontSize: 7.5,
    dataLabelColor: C.white,
    chartArea: { fill: { color: C.white }, border: { color: C.ltgray, width: 0.5 } },
    catAxisLabelColor: C.midgray,
    valAxisLabelColor: C.midgray,
    catAxisLabelFontSize: 8,
    valAxisLabelFontSize: 8,
    valGridLine: { color: C.ltgray, size: 0.5 },
    catGridLine: { style: "none" },
    showLegend: true,
    legendPos: "b",
    legendFontSize: 7.5,
  });

  // ── Unit Economics Table ──
  sectionLabel(slide, 6.72, TOP + 0.92, 3.13, "UNIT ECONOMICS");
  const ueRows = [
    ["Metric", "Q3 2026", "Q2 2026", "Target", "Trend"],
    ["Gross Margin", "78.2%", "77.5%", "80%+", { text: "↑ Improving", color: C.teal }],
    ["NRR", "118%", "121%", "120%+", { text: "↓ Churn", color: C.amber }],
    ["Logo Retention", "92.4%", "94.1%", "95%+", { text: "↓ Risk", color: C.red }],
    ["LTV:CAC", "4.8x", "4.6x", "4.0x+", { text: "↑ Healthy", color: C.teal }],
    ["CAC Payback", "13.2 mo", "14.1 mo", "<16 mo", { text: "↑ Improving", color: C.teal }],
    ["Magic Number", "1.12", "0.98", ">0.75", { text: "↑ Strong", color: C.teal }],
  ];
  tbl(slide, ueRows, {
    x: 6.72, y: TOP + 1.18, h: 0.3,
    colW: [0.85, 0.57, 0.57, 0.57, 0.57],
    fontSize: 9,
    hdrFill: C.navy,
  });

  // ARR Bridge table below chart
  sectionLabel(slide, 0.15, TOP + 3.5, 6.6, "ARR BRIDGE  —  Q3 2026");
  const bridgeRows = [
    ["Component", "Q1 2026", "Q2 2026", "Q3 2026", "QoQ Δ", "vs Plan"],
    ["Beginning ARR", "$11.4M", "$12.8M", "$14.1M", "—", "—"],
    [{ text: "+ New ARR", color: C.teal }, "$1.4M", "$1.6M", { text: "$1.9M", color: C.teal }, "+18.8%", { text: "+12% ↑", color: C.teal }],
    [{ text: "+ Expansion ARR", color: C.teal }, "$0.6M", "$0.7M", { text: "$0.8M", color: C.teal }, "+14.3%", "On plan"],
    [{ text: "− Churn ARR", color: C.red }, "($0.3M)", "($0.3M)", { text: "($0.5M)", color: C.red }, { text: "+66.7% ⚠", color: C.red }, { text: "Behind ↓", color: C.red }],
    [{ text: "= Ending ARR", bold: true }, "$12.8M", "$14.1M", { text: "$15.6M", bold: true, color: C.accent }, "+10.6%", { text: "+4.2% ↑", color: C.teal }],
  ];
  tbl(slide, bridgeRows, {
    x: 0.15, y: TOP + 3.72, h: 0.155,
    colW: [1.6, 0.88, 0.88, 0.88, 0.88, 0.88],
    fontSize: 9,
    hdrFill: C.navy,
  });
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 3 — GTM + CUSTOMER HEALTH
// ═══════════════════════════════════════════════════════════════
{
  const slide = pres.addSlide();
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: H, fill: { color: C.offwhite }, line: { color: C.offwhite } });
  hdr(slide, "SLIDE 3  ·  GTM Performance & Customer Health", "Q3 2026");

  const TOP = 0.65;

  // ── Sales KPIs row ──
  const gtmCards = [
    { label: "Pipeline Generated", value: "$8.6M", sub: "+21.1% QoQ  ·  Best ever", vc: C.accent },
    { label: "Win Rate", value: "28.4%", sub: "+2.3pp QoQ  ·  Enterprise improving", vc: C.teal },
    { label: "Avg Deal Size", value: "$42K ACV", sub: "+10.5%  ·  Moving upmarket ✓", vc: C.teal },
    { label: "Quota Attainment", value: "112%", sub: "+18pp vs Q2  ·  Q2 was 94%", vc: C.teal },
    { label: "Sales Cycle", value: "68 days", sub: "-5.6% QoQ  ·  New demo flow", vc: C.teal },
  ];
  const cw = (W - 0.3) / gtmCards.length;
  gtmCards.forEach((c, i) => {
    kpiCard(slide, 0.15 + i * cw, TOP, cw - 0.08, 0.78, c.label, c.value, c.sub, c.vc);
  });

  // ── Channel Performance ──
  sectionLabel(slide, 0.15, TOP + 0.88, 4.55, "CHANNEL PERFORMANCE  —  Q3 2026");
  const chanRows = [
    ["Channel", "Pipeline", "Deals Won", "Avg ACV", "CAC", "Notes"],
    ["Outbound SDR", "$3.8M", "18", "$52K", "$24.1K", "Scaled 4→6 SDRs in Q2"],
    ["Inbound Marketing", "$2.4M", "14", "$34K", "$12.8K", "Content + paid search"],
    [{ text: "Partner / Referral", color: C.teal }, "$1.5M", "5", { text: "$68K", color: C.teal }, "$8.4K", { text: "SAP + Siemens ramping ↑", color: C.teal }],
    [{ text: "TOTAL", bold: true }, { text: "$8.6M", bold: true }, { text: "45", bold: true }, { text: "$42K", bold: true }, "$15.8K", ""],
  ];
  tbl(slide, chanRows, {
    x: 0.15, y: TOP + 1.12, h: 0.28,
    colW: [1.2, 0.72, 0.72, 0.65, 0.65, 1.5],
    fontSize: 9,
    hdrFill: C.navy,
  });

  // ── CAC Trend chart ──
  sectionLabel(slide, 4.82, TOP + 0.88, 5.03, "BLENDED CAC TREND  ($K)");
  const cacData = [
    { name: "Blended",     labels: ["Q1'25","Q2'25","Q3'25","Q4'25","Q1'26","Q2'26","Q3'26"], values: [22.4,20.1,19.8,18.2,17.6,16.4,15.8] },
    { name: "Enterprise",  labels: ["Q1'25","Q2'25","Q3'25","Q4'25","Q1'26","Q2'26","Q3'26"], values: [38.2,34.6,32.1,28.4,26.8,24.2,22.6] },
    { name: "Mid-Market",  labels: ["Q1'25","Q2'25","Q3'25","Q4'25","Q1'26","Q2'26","Q3'26"], values: [14.8,13.2,13.6,12.4,12.1,11.8,11.2] },
  ];
  slide.addChart(pres.charts.LINE, cacData, {
    x: 4.82, y: TOP + 1.12, w: 5.03, h: 1.82,
    chartColors: [C.accent, C.teal, C.amber],
    lineSize: 2,
    showValue: false,
    chartArea: { fill: { color: C.white }, border: { color: C.ltgray, width: 0.5 } },
    catAxisLabelColor: C.midgray, valAxisLabelColor: C.midgray,
    catAxisLabelFontSize: 7.5, valAxisLabelFontSize: 7.5,
    valGridLine: { color: C.ltgray, size: 0.5 }, catGridLine: { style: "none" },
    showLegend: true, legendPos: "r", legendFontSize: 7.5,
    valAxisTitle: true, valAxisTitleText: 'Churn ($K)', valAxisTitleFontSize: 8,
  });

  // ── Churn Deep Dive ──
  sectionLabel(slide, 0.15, TOP + 2.57, 5.8, "CHURN DEEP DIVE  —  $500K IN Q3  (+66% QoQ)  ·  58% PREVENTABLE", C.red);
  const churnRows = [
    ["Account", "Segment", "ARR Lost", "Reason", { text: "Preventable?", align: "center" }],
    ["Apex Manufacturing", "Enterprise", "$145K", "Acquired — contract not renewed post-M&A", { text: "No", color: C.midgray }],
    ["4 SMB accounts", "SMB", "$118K", "Price sensitivity; 3→DataForge, 1 went OOB", { text: "Yes — 2-3 saveable", color: C.red }],
    ["TechFab Solutions", "Mid-Market", "$48K", "Poor implementation; never achieved time-to-value", { text: "Yes — CS failure", color: C.red }],
    ["3 SMB accounts", "SMB", "$82K", "Low usage (<10% adoption); never onboarded", { text: "Yes — onboarding gap", color: C.red }],
    [{ text: "TOTAL / PREVENTABLE", bold: true }, "", { text: "$500K", bold: true, color: C.red }, { text: "$290K of $500K (58%) preventable", bold: true }, { text: "$290K recoverable", bold: true, color: C.red }],
  ];
  tbl(slide, churnRows, {
    x: 0.15, y: TOP + 2.81, h: 0.31,
    colW: [1.38, 0.82, 0.65, 2.42, 1.2],
    fontSize: 9,
    hdrFill: "7F1D1D",
  });

  // ── At-Risk Accounts ──
  sectionLabel(slide, 6.1, TOP + 2.57, 3.75, "AT-RISK ACCOUNTS  —  Q4 WATCH LIST", C.amber);
  const riskRows = [
    ["Account", "ARR", "Risk Signal", "Owner"],
    [{ text: "Sterling Industries", color: C.red }, { text: "$210K", color: C.red }, "Champion left; new VP evaluating alternatives", "R. Torres"],
    ["ClearPath Systems", "$72K", "Renewal in 60 days; competitor POC underway", "D. Park"],
    ["Midwest Components", "$58K", "Usage -40% Aug; tickets 3x — CS engaged", "J. Lin"],
  ];
  tbl(slide, riskRows, {
    x: 6.1, y: TOP + 2.81, h: 0.40,
    colW: [0.95, 0.55, 1.70, 0.55],
    fontSize: 9,
    hdrFill: C.amber,
  });
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 4 — PRODUCT + TEAM
// ═══════════════════════════════════════════════════════════════
{
  const slide = pres.addSlide();
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: H, fill: { color: C.offwhite }, line: { color: C.offwhite } });
  hdr(slide, "SLIDE 4  ·  Product & Engineering / Team", "Q3 2026 Shipped · Q4 Roadmap · Headcount");

  const TOP = 0.65;

  // ── Q3 Shipped ──
  sectionLabel(slide, 0.15, TOP, 5.5, "Q3 SHIPPED FEATURES");
  const shipRows = [
    ["Feature", "30-Day Adoption", "Revenue Signal"],
    [{ text: "Predictive Maintenance v2", bold: true }, "67% of enterprise accounts enabled", { text: "$1.2M pipeline; 3 deals cite directly", color: C.teal }],
    ["Self-Serve Dashboard Builder", "340 dashboards by 89 accounts", "CS ticket volume -18%"],
    [{ text: "SAP Integration (native)", color: C.accent }, "12 connected; 8 in progress", { text: "Partner channel accelerant", color: C.accent }],
    [{ text: "SOC 2 Type II ✓  (cert 8/15)", color: C.teal }, "N/A — compliance milestone", { text: "Unblocked 4 deals ($380K pipeline)", color: C.teal }],
  ];
  tbl(slide, shipRows, {
    x: 0.15, y: TOP + 0.24, h: 0.32,
    colW: [1.7, 2.05, 1.8],
    fontSize: 9,
    hdrFill: C.navy,
  });

  // ── Q4 Roadmap ──
  sectionLabel(slide, 0.15, TOP + 1.9, 5.5, "Q4 ROADMAP PRIORITIES");
  const roadRows = [
    ["Pri", "Feature", "Status", "Expected Impact"],
    [{ text: "P0", bold: true, color: C.red }, "Multi-tenant analytics (enterprise isolation)", "In dev — 60%", { text: "2 deals $200K+ blocked on this", color: C.red }],
    [{ text: "P0", bold: true, color: C.red }, "Siemens MindSphere integration", "Design done; dev starting", { text: "Opens $4M TAM in industrial IoT", color: C.teal }],
    [{ text: "P1", bold: true, color: C.amber }, "Usage-based pricing tier (SMB)", "Spec complete", { text: "Addresses 58% of preventable churn", color: C.amber }],
    [{ text: "P1", bold: true, color: C.amber }, "Customer health score dashboard", "Prototype built", "Early warning for at-risk accounts"],
    ["P2", "Mobile app (read-only dashboards)", "Scoping", "Frequently requested; low revenue impact"],
  ];
  tbl(slide, roadRows, {
    x: 0.15, y: TOP + 2.14, h: 0.295,
    colW: [0.35, 2.0, 1.35, 1.85],
    fontSize: 9,
    hdrFill: C.navy,
  });

  // ── Headcount by Function ──
  sectionLabel(slide, 5.8, TOP, 4.05, "HEADCOUNT BY FUNCTION");
  const hcRows = [
    ["Function", "Q2", "Q3", "Open", "Q4 Target"],
    ["Engineering", "42", "46", "4", "50"],
    ["Product & Design", "8", "9", "1", "10"],
    ["Sales", "18", "22", "3", "25"],
    ["Customer Success", "12", "14", "2", "16"],
    ["Marketing", "8", "9", "1", "10"],
    ["G&A", "10", "12", "1", "13"],
    [{ text: "TOTAL", bold: true }, { text: "98", bold: true }, { text: "112", bold: true }, { text: "12", bold: true, color: C.amber }, { text: "124", bold: true, color: C.teal }],
  ];
  tbl(slide, hcRows, {
    x: 5.8, y: TOP + 0.24, h: 0.29,
    colW: [1.42, 0.55, 0.55, 0.55, 0.88],
    fontSize: 9,
    hdrFill: C.navy,
  });

  // Headcount bar chart
  slide.addChart(pres.charts.BAR, [
    { name: "Q3 2026", labels: ["Eng","Prod","Sales","CS","Mktg","G&A"], values: [46,9,22,14,9,12] },
    { name: "Q4 Target", labels: ["Eng","Prod","Sales","CS","Mktg","G&A"], values: [50,10,25,16,10,13] },
  ], {
    x: 5.8, y: TOP + 2.61, w: 4.05, h: 1.70,
    barDir: "col",
    barGrouping: "clustered",
    chartColors: [C.accent, "93C5FD"],
    chartArea: { fill: { color: C.white }, border: { color: C.ltgray, width: 0.5 } },
    catAxisLabelColor: C.midgray, valAxisLabelColor: C.midgray,
    catAxisLabelFontSize: 9, valAxisLabelFontSize: 9,
    valGridLine: { color: C.ltgray, size: 0.5 }, catGridLine: { style: "none" },
    showLegend: true, legendPos: "b", legendFontSize: 9,
    showValue: true, dataLabelFontSize: 9,
  });

  // Key hires
  sectionLabel(slide, 0.15, TOP + 3.42, 5.5, "KEY HIRES — Q3 2026");
  const hireRows = [
    ["Role", "Background", "Focus"],
    [{ text: "VP Customer Success (9/1)", bold: true }, "Previously at Datadog — built CS org from 20→80", { text: "Own churn reduction; build SMB onboarding motion", color: C.teal }],
    [{ text: "Head of Partnerships (8/15)", bold: true }, "Previously at Siemens — owns SAP + Siemens channel", { text: "Accelerate partner pipeline ($1.5M → $5M target)", color: C.accent }],
    ["2 Senior ML Engineers", "PhD hires from Georgia Tech", "Predictive Maintenance v3; model accuracy roadmap"],
  ];
  tbl(slide, hireRows, {
    x: 0.15, y: TOP + 3.44, h: 0.32,
    colW: [1.47, 2.30, 1.73],
    fontSize: 9,
    hdrFill: C.navy,
  });

  // Attrition note
  slide.addText("⚠  Q3 Attrition: 6 departures (5.4% quarterly) — 3 engineering to FAANG. Comp bands adjusted +12% for senior engineers.", {
    x: 0.15, y: H - 0.28, w: 5.5, h: 0.28, fontSize: 9, color: C.amber, margin: 0
  });
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 5 — FINANCIAL OUTLOOK + COMPETITIVE
// ═══════════════════════════════════════════════════════════════
{
  const slide = pres.addSlide();
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: H, fill: { color: C.offwhite }, line: { color: C.offwhite } });
  hdr(slide, "SLIDE 5  ·  Financial Outlook & Competitive Landscape", "FY2026E / Cash Position / Market Dynamics");

  const TOP = 0.65;

  // ── P&L Table ──
  sectionLabel(slide, 0.15, TOP, 5.8, "P&L SUMMARY  ($K)");
  const plRows = [
    ["Line Item", "Q1 2026", "Q2 2026", "Q3 2026", "Q3 YoY", "FY2026E"],
    ["Revenue", "$3,400", "$3,700", "$4,100", "+64%", "$15,400"],
    ["COGS", "($740)", "($830)", "($890)", "+52%", "($3,340)"],
    [{ text: "Gross Profit", bold: true }, { text: "$2,660", bold: true }, { text: "$2,870", bold: true }, { text: "$3,210", bold: true, color: C.teal }, { text: "+68%", color: C.teal }, { text: "$12,060", bold: true }],
    ["Gross Margin", "78.2%", "77.6%", "78.3%", "+2.1pp", "78.3%"],
    ["S&M", "($1,420)", "($1,580)", "($1,720)", "+48%", "($6,480)"],
    ["R&D", "($1,340)", "($1,480)", "($1,620)", "+55%", "($6,040)"],
    ["G&A", "($480)", "($510)", "($540)", "+38%", "($2,060)"],
    [{ text: "Total OpEx", bold: true }, "($3,240)", "($3,570)", "($3,880)", "+50%", "($14,580)"],
    [{ text: "Net Income", bold: true }, { text: "($580)", color: C.red }, { text: "($700)", color: C.red }, { text: "($670)", color: C.amber }, { text: "Improving", color: C.teal }, { text: "($2,520)", color: C.amber }],
    ["Monthly Burn", "$193K", "$233K", "$223K", "—", "—"],
  ];
  tbl(slide, plRows, {
    x: 0.15, y: TOP + 0.24, h: 0.278,
    colW: [1.38, 0.82, 0.82, 0.82, 0.72, 0.82],
    fontSize: 9,
    hdrFill: C.navy,
  });

  // ── Cash Position KPIs ──
  sectionLabel(slide, 0.15, TOP + 3.35, 5.8, "CASH POSITION & EFFICIENCY");
  const cashCards = [
    { label: "Cash on Hand", value: "$18.2M", sub: "Series B $32M (Mar 2025)", vc: C.teal },
    { label: "Monthly Burn", value: "$223K", sub: "Q3 avg — declining QoQ", vc: C.accent },
    { label: "Runway", value: "81 months", sub: "Comfortable — not a near-term factor", vc: C.teal },
    { label: "Burn Multiple", value: "0.36x", sub: "Excellent (<1x = efficient)", vc: C.teal },
  ];
  const ccw = 5.8 / cashCards.length;
  cashCards.forEach((c, i) => {
    kpiCard(slide, 0.15 + i * ccw, TOP + 3.59, ccw - 0.08, 0.78, c.label, c.value, c.sub, c.vc);
  });

  // ── Burn Chart ──
  sectionLabel(slide, 6.1, TOP, 3.75, "MONTHLY BURN & REVENUE TREND  ($K)");
  slide.addChart(pres.charts.BAR, [
    { name: "Monthly Burn ($K)", labels: ["Q1'26","Q2'26","Q3'26"], values: [193, 233, 223] },
  ], {
    x: 6.1, y: TOP + 0.24, w: 3.75, h: 1.58,
    barDir: "col",
    chartColors: [C.red],
    showValue: true,
    dataLabelFontSize: 8,
    chartArea: { fill: { color: C.white }, border: { color: C.ltgray, width: 0.5 } },
    catAxisLabelColor: C.midgray, valAxisLabelColor: C.midgray,
    catAxisLabelFontSize: 8, valAxisLabelFontSize: 8,
    valGridLine: { color: C.ltgray, size: 0.5 }, catGridLine: { style: "none" },
    showLegend: false,
    valAxisTitle: true, valAxisTitleText: '$K', valAxisTitleFontSize: 8,
  });

  // ── Competitive Table ──
  sectionLabel(slide, 6.1, TOP + 1.96, 3.75, "COMPETITIVE LANDSCAPE");
  const compRows = [
    ["Dimension", "NovaCrest", "DataForge", "Acme", "Zenith AI"],
    ["Stage", "Series B $32M", "Series C $85M", "Public $2.1B", "Series A $18M"],
    ["ARR (est)", "$15.6M", "~$45M", "~$180M", "~$4M"],
    ["Target Mkt", "Mid-mkt mfg", "SMB-Mid horiz.", "Enterprise all", "Mid-mkt mfg"],
    ["Avg ACV", "$30-250K", "$2.4-48K", "$100K-500K+", "$20-80K"],
    ["Win Rate", "—", { text: "62% ✓", color: C.teal }, { text: "34% ✗", color: C.red }, { text: "55% ✓", color: C.teal }],
    ["Key Threat", "—", { text: "$199/mo SMB module", color: C.amber }, { text: "Acme Lite Q1'27", color: C.amber }, { text: "VC hiring blitz", color: C.red }],
  ];
  tbl(slide, compRows, {
    x: 6.1, y: TOP + 2.2, h: 0.27,
    colW: [0.76, 0.82, 0.78, 0.70, 0.69],
    fontSize: 9,
    hdrFill: C.navy,
  });
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 6 — BOARD ASKS
// ═══════════════════════════════════════════════════════════════
{
  const slide = pres.addSlide();
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: H, fill: { color: C.navy }, line: { color: C.navy } });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0.55, w: W, h: 0.03, fill: { color: C.accent }, line: { color: C.accent } });

  // Header
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: 0.55, fill: { color: "080F1C" }, line: { color: "080F1C" } });
  slide.addText("BOARD ASKS  ·  Q3 2026 REVIEW  ·  THREE ITEMS REQUIRING ACTION", {
    x: 0.3, y: 0, w: 9.4, h: 0.55, fontSize: 14, bold: true, color: C.white, valign: "middle", charSpacing: 1.5, margin: 0
  });

  const askData = [
    {
      num: "01",
      label: "APPROVE",
      color: C.red,
      title: "$2.5M Incremental Customer Success Investment",
      x: 0.15, y: 0.72, w: 3.1,
      bullets: [
        "Hire 8 FTEs: 4 onboarding specialists, 2 renewal managers, 1 SMB CS lead, 1 CS ops",
        "$290K of Q3 churn ($500K total) was preventable — weak onboarding, no down-tier",
        "Sterling ($210K), ClearPath ($72K), Midwest ($58K) at risk in Q4 alone",
      ],
      metrics: [
        ["Expected ARR Saved", "~$700K/yr"],
        ["Payback Period", "14 months"],
        ["ROI (3-yr)", "~280%"],
        ["If Not Approved", "Q4 churn ≥ $600K"],
      ],
      urgency: "URGENT — Q4 at-risk pipeline is $340K; delay = churn compounds",
    },
    {
      num: "02",
      label: "APPROVE",
      color: C.amber,
      title: "Usage-Based SMB Pricing Tier  ($499/mo entry)",
      x: 3.45, y: 0.72, w: 3.1,
      bullets: [
        "Current minimum $2.5K/mo losing 60+ SMB prospects/quarter on price",
        "Addresses SMB onboarding gap — usage-based aligns cost to adoption",
        "DataForge $199/mo module is pulling SMB deals away now",
      ],
      metrics: [
        ["New ARR (12 mo)", "+$1.2M"],
        ["Down-tier risk", "~($200K)"],
        ["Net ARR impact", "+$1.0M"],
        ["Timeline", "Q1 2027 launch"],
      ],
      urgency: "COMPETITIVE — DataForge manufacturing module launched in Q3",
    },
    {
      num: "03",
      label: "DISCUSS",
      color: C.accent,
      title: "Series C Timing — Q2 2027 vs $30M ARR Milestone",
      x: 6.75, y: 0.72, w: 3.1,
      bullets: [
        "Current runway: 81 months — not forced; can raise from position of strength",
        "Option A: Q2 2027 at $22-25M ARR — momentum story, growth rate compelling",
        "Option B: Wait for $30M ARR — stronger metrics but 2028 timeline, market risk",
      ],
      metrics: [
        ["Current ARR", "$15.6M"],
        ["Cash on Hand", "$18.2M"],
        ["Burn Multiple", "0.36x"],
        ["Target Raise", "TBD — board input"],
      ],
      urgency: "STRATEGIC — Board input requested on timing and investor targets",
    },
  ];

  askData.forEach(ask => {
    // Card bg
    slide.addShape(pres.shapes.RECTANGLE, { x: ask.x, y: ask.y, w: ask.w, h: 4.38, fill: { color: "0A1628" }, line: { color: ask.color, width: 1.2 } });
    // Top accent
    slide.addShape(pres.shapes.RECTANGLE, { x: ask.x, y: ask.y, w: ask.w, h: 0.06, fill: { color: ask.color }, line: { color: ask.color } });

    // Number + label
    slide.addText(ask.num, { x: ask.x + 0.12, y: ask.y + 0.1, w: 0.55, h: 0.42, fontSize: 26, bold: true, color: ask.color, margin: 0 });
    slide.addShape(pres.shapes.RECTANGLE, { x: ask.x + 0.65, y: ask.y + 0.18, w: 0.72, h: 0.22, fill: { color: ask.color }, line: { color: ask.color } });
    slide.addText(ask.label, { x: ask.x + 0.66, y: ask.y + 0.18, w: 0.7, h: 0.22, fontSize: 7.5, bold: true, color: C.white, valign: "middle", align: "center", margin: 0 });

    // Title
    slide.addText(ask.title, { x: ask.x + 0.12, y: ask.y + 0.52, w: ask.w - 0.22, h: 0.5, fontSize: 10, bold: true, color: C.white, margin: 0 });

    // Rationale
    slide.addShape(pres.shapes.RECTANGLE, { x: ask.x + 0.12, y: ask.y + 1.06, w: ask.w - 0.22, h: 0.18, fill: { color: "141F35" }, line: { color: "141F35" } });
    slide.addText("RATIONALE", { x: ask.x + 0.14, y: ask.y + 1.06, w: ask.w - 0.24, h: 0.18, fontSize: 7, bold: true, color: ask.color, valign: "middle", margin: 0 });
    ask.bullets.forEach((b, bi) => {
      slide.addText("·  " + b, { x: ask.x + 0.14, y: ask.y + 1.27 + bi * 0.35, w: ask.w - 0.24, h: 0.32, fontSize: 7.8, color: "B0C4DE", margin: 0 });
    });

    // Metrics
    slide.addShape(pres.shapes.RECTANGLE, { x: ask.x + 0.12, y: ask.y + 2.37, w: ask.w - 0.22, h: 0.18, fill: { color: "141F35" }, line: { color: "141F35" } });
    slide.addText("KEY METRICS", { x: ask.x + 0.14, y: ask.y + 2.37, w: ask.w - 0.24, h: 0.18, fontSize: 7, bold: true, color: ask.color, valign: "middle", margin: 0 });
    ask.metrics.forEach((m, mi) => {
      const mx = ask.x + 0.12 + (mi % 2) * ((ask.w - 0.22) / 2);
      const my = ask.y + 2.58 + Math.floor(mi / 2) * 0.52;
      const mw = (ask.w - 0.26) / 2;
      slide.addShape(pres.shapes.RECTANGLE, { x: mx, y: my, w: mw - 0.04, h: 0.48, fill: { color: "0F1E36" }, line: { color: "1E3A5F", width: 0.5 } });
      slide.addText(m[1], { x: mx + 0.06, y: my + 0.04, w: mw - 0.14, h: 0.24, fontSize: 11, bold: true, color: ask.color, margin: 0 });
      slide.addText(m[0], { x: mx + 0.06, y: my + 0.27, w: mw - 0.14, h: 0.18, fontSize: 7, color: C.midgray, margin: 0 });
    });

    // Urgency
    slide.addShape(pres.shapes.RECTANGLE, { x: ask.x + 0.12, y: ask.y + 3.66, w: ask.w - 0.22, h: 0.52, fill: { color: "0F1E36" }, line: { color: ask.color, width: 0.5 } });
    slide.addText(ask.urgency, { x: ask.x + 0.17, y: ask.y + 3.7, w: ask.w - 0.3, h: 0.44, fontSize: 7.8, color: ask.color, margin: 0 });
  });

  // Footer
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: H - 0.28, w: W, h: 0.28, fill: { color: "080F1C" }, line: { color: "080F1C" } });
  slide.addText("NovaCrest  |  Q3 2026 Board Review  |  CONFIDENTIAL  |  For Board Use Only", {
    x: 0.3, y: H - 0.28, w: 9.4, h: 0.28, fontSize: 7.5, color: "3D5470", valign: "middle", margin: 0
  });
}

// ─── WRITE FILE ───────────────────────────────────────────────
pres.writeFile({ fileName: "outputs/strategy-qa.pptx" })
  .then(() => console.log("✓ outputs/strategy-qa.pptx written"))
  .catch(e => { console.error(e); process.exit(1); });
