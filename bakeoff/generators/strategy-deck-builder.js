"use strict";
const pptxgen = require("pptxgenjs");

// ============================================================
// NovaCrest Q3 2026 Strategic Review — Board Deck
// ============================================================

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.author = "NovaCrest";
pres.title = "NovaCrest Q3 2026 Strategic Review";

// ============================================================
// LAYER 1: CONSTANTS
// ============================================================
const W = 10;
const H = 5.625;
const PAD = 0.5;
const TITLE_H = 0.5;
const BODY_TOP = 0.62;
const FOOTER_Y = 5.35;
const CONTENT_BOTTOM = 5.23;
const SECTION_GAP = 0.12;
const MIN_FONT = 9;

const C = {
  darkNavy:   "1A2744",
  navy:       "162340",
  teal:       "0EA5E9",
  tealDark:   "0284C7",
  amber:      "F59E0B",
  amberLight: "FEF3C7",
  green:      "10B981",
  greenLight: "D1FAE5",
  red:        "EF4444",
  white:      "FFFFFF",
  lightBg:    "F8FAFC",
  cardBg:     "FFFFFF",
  border:     "E2E8F0",
  slate:      "64748B",
  darkSlate:  "334155",
  midGray:    "94A3B8",
  textDark:   "1E293B",
  rowAlt:     "F1F5F9",
  blue:       "3B82F6",
  blueDark:   "1D4ED8",
};

const FONT = {
  head: "Calibri",
  sub:  "Calibri",
  body: "Calibri",
};

// ============================================================
// LAYER 2: HELPERS
// ============================================================

function addDarkSlide() {
  const slide = pres.addSlide();
  slide.background = { color: C.darkNavy };
  return slide;
}

function addLightSlide() {
  const slide = pres.addSlide();
  slide.background = { color: C.lightBg };
  return slide;
}

// Returns y after header block
function addHeader(slide, title, dark = false) {
  const color = dark ? C.white : C.textDark;
  const accentColor = dark ? C.teal : C.teal;
  // Accent bar
  slide.addShape("rect", {
    x: 0, y: 0, w: W, h: 0.06,
    fill: { color: accentColor },
  });
  slide.addText(title, {
    x: PAD, y: 0.1, w: W - PAD * 2, h: TITLE_H,
    fontSize: 15, fontFace: FONT.head,
    color: color, bold: true,
    valign: "middle", margin: 0,
  });
  return BODY_TOP;
}

// Footer
function addFooter(slide, num, total, dark = false) {
  const color = dark ? C.midGray : C.slate;
  slide.addText(`NovaCrest  |  Q3 2026 Strategic Review  |  CONFIDENTIAL`, {
    x: PAD, y: FOOTER_Y, w: 6.5, h: 0.22,
    fontSize: 7, fontFace: FONT.body,
    color: color, align: "left", margin: 0,
  });
  slide.addText(`${num} / ${total}`, {
    x: W - 1.2, y: FOOTER_Y, w: 0.9, h: 0.22,
    fontSize: 7, fontFace: FONT.body,
    color: color, align: "right", margin: 0,
  });
}

// Section label — returns y + 0.24
function addSectionLabel(slide, label, x, y, w, dark = false) {
  const color = dark ? C.teal : C.tealDark;
  slide.addText(label.toUpperCase(), {
    x, y, w, h: 0.2,
    fontSize: 8, fontFace: FONT.body,
    color: color, bold: true, margin: 0,
  });
  slide.addShape("rect", {
    x, y: y + 0.2, w, h: 0.015,
    fill: { color: color },
  });
  return y + 0.24;
}

// KPI stat card
function addKpiCard(slide, x, y, w, h, value, label, sub, valueColor, dark = false) {
  const bg = dark ? C.navy : C.cardBg;
  const labelColor = dark ? C.midGray : C.slate;
  const subColor = dark ? C.teal : C.tealDark;
  slide.addShape("rect", {
    x, y, w, h,
    fill: { color: bg },
    line: { color: dark ? "2A3A5E" : C.border, width: 0.5 },
  });
  slide.addText(value, {
    x: x + 0.08, y: y + 0.06, w: w - 0.16, h: h * 0.46,
    fontSize: 18, fontFace: FONT.head,
    color: valueColor || (dark ? C.white : C.textDark),
    bold: true, align: "center", valign: "middle", margin: 0,
  });
  slide.addText(label, {
    x: x + 0.06, y: y + h * 0.5, w: w - 0.12, h: 0.22,
    fontSize: 8.5, fontFace: FONT.body,
    color: labelColor, bold: true, align: "center", margin: 0,
  });
  if (sub) {
    slide.addText(sub, {
      x: x + 0.06, y: y + h * 0.5 + 0.22, w: w - 0.12, h: 0.2,
      fontSize: 8, fontFace: FONT.body,
      color: subColor, align: "center", margin: 0,
    });
  }
}

// Bullet list — returns y after last bullet
function addBullets(slide, items, x, y, w, dark = false) {
  const color = dark ? "D0DCF0" : C.darkSlate;
  items.forEach((item) => {
    const bullet = item.bullet || "•";
    const bColor = item.bulletColor || (dark ? C.teal : C.tealDark);
    slide.addText(`${bullet}  ${item.text}`, {
      x, y, w, h: item.h || 0.26,
      fontSize: item.fontSize || 9.5, fontFace: FONT.body,
      color: item.color || color,
      bold: item.bold || false,
      margin: 0, valign: "middle",
    });
    y += item.h || 0.26;
  });
  return y;
}

// Table helper
function makeRow(cells, isHeader, altRow, colWidths) {
  return cells.map((c, i) => {
    const isObj = typeof c === "object" && c !== null && c.text !== undefined;
    const text = isObj ? c.text : String(c);
    const overrides = isObj ? (c.options || {}) : {};
    const baseFill = isHeader ? C.tealDark : (altRow ? C.rowAlt : C.cardBg);
    const baseColor = isHeader ? C.white : C.darkSlate;
    return {
      text,
      options: {
        fill: { color: overrides.fill?.color || baseFill },
        color: overrides.color || baseColor,
        fontSize: isHeader ? 8 : MIN_FONT,
        fontFace: FONT.body,
        bold: isHeader || overrides.bold || false,
        align: (i === 0 ? "left" : "center"),
        valign: "middle",
        margin: [0, 4, 0, 4],
        ...overrides,
        fill: { color: overrides.fill?.color || baseFill },
      },
    };
  });
}

function addTable(slide, headers, rows, x, y, w, colW, rowH, opts = {}) {
  const tableRows = [
    makeRow(headers, true, false, colW),
    ...rows.map((r, i) => makeRow(r, false, i % 2 === 1, colW)),
  ];
  slide.addTable(tableRows, {
    x, y, w, colW,
    rowH: rowH || 0.26,
    border: { type: "solid", pt: 0.3, color: C.border },
    autoPage: false,
    ...opts,
  });
}

// ============================================================
// LAYER 3: SLIDES
// ============================================================

// ------------------------------------------------------------------
// SLIDE 1: Title + Executive Summary (dark)
// ------------------------------------------------------------------
function slide01() {
  const slide = addDarkSlide();
  addFooter(slide, 1, 6, true);

  // Left column: title block
  slide.addText("CONFIDENTIAL", {
    x: PAD, y: 0.08, w: 3, h: 0.22,
    fontSize: 7.5, fontFace: FONT.body,
    color: C.amber, bold: true, margin: 0,
  });

  slide.addText("NovaCrest", {
    x: PAD, y: 0.28, w: 5.5, h: 0.38,
    fontSize: 11, fontFace: FONT.body,
    color: C.teal, bold: true, margin: 0,
  });

  slide.addText("Q3 2026 Strategic Review", {
    x: PAD, y: 0.62, w: 5.5, h: 0.72,
    fontSize: 26, fontFace: FONT.head,
    color: C.white, bold: true, margin: 0,
    lineSpacingMultiple: 0.9,
  });

  // Company context
  slide.addText("Series B SaaS  |  $32M (March 2025)  |  B2B Predictive Analytics for Manufacturing  |  118 Employees", {
    x: PAD, y: 1.38, w: 5.5, h: 0.26,
    fontSize: 8.5, fontFace: FONT.body,
    color: C.midGray, margin: 0,
  });

  // Board members
  slide.addText("Board: Sarah Chen (CEO)  ·  Jeff Blackwell (Insight)  ·  Priya Sharma (Accel)  ·  Tom Nguyen (Independent)", {
    x: PAD, y: 1.68, w: 5.5, h: 0.22,
    fontSize: 7.5, fontFace: FONT.body,
    color: C.midGray, margin: 0,
  });

  // Separator line
  slide.addShape("rect", {
    x: PAD, y: 1.96, w: 4.2, h: 0.018,
    fill: { color: C.teal },
  });

  // WINS section
  let y = 2.08;
  y = addSectionLabel(slide, "3 Wins", PAD, y, 4.5, true);
  y += 0.04;
  const wins = [
    { text: "ARR +10.6% QoQ to $15.6M — ahead of plan", bulletColor: C.green, color: "A7F3D0" },
    { text: "Pipeline record: $8.6M generated (best-ever quarter)", bulletColor: C.green, color: "A7F3D0" },
    { text: "Q3 hiring complete: VP CS + Head Partnerships", bulletColor: C.green, color: "A7F3D0" },
  ];
  wins.forEach((w2) => {
    slide.addText(`✓  ${w2.text}`, {
      x: PAD, y, w: 4.5, h: 0.26,
      fontSize: 9.5, fontFace: FONT.body,
      color: w2.color, margin: 0, valign: "middle",
    });
    y += 0.27;
  });

  // CONCERNS
  y += 0.06;
  y = addSectionLabel(slide, "2 Concerns", PAD, y, 4.5, true);
  y += 0.04;
  const concerns = [
    { text: "Churn spike: $500K (up from $300K) — $290K preventable", color: C.amber },
    { text: "Competitive pressure: DataForge mfg module + Zenith Series A", color: C.amber },
  ];
  concerns.forEach((c2) => {
    slide.addText(`⚠  ${c2.text}`, {
      x: PAD, y, w: 4.5, h: 0.26,
      fontSize: 9.5, fontFace: FONT.body,
      color: c2.color, margin: 0, valign: "middle",
    });
    y += 0.27;
  });

  // DECISION
  y += 0.06;
  y = addSectionLabel(slide, "Decision Required", PAD, y, 4.5, true);
  y += 0.04;
  slide.addText("▶  CS Investment: Approve $2.5M to address $290K preventable churn", {
    x: PAD, y, w: 4.5, h: 0.26,
    fontSize: 9.5, fontFace: FONT.body,
    color: "93C5FD", margin: 0, valign: "middle",
  });

  // Right side: KPI summary cards
  const kpis = [
    { value: "$15.6M", label: "ARR", sub: "+10.6% QoQ", vc: C.green },
    { value: "118%", label: "NRR", sub: "Q2: 121%", vc: C.teal },
    { value: "$8.6M", label: "Pipeline", sub: "Record quarter", vc: C.green },
    { value: "$18.2M", label: "Cash", sub: "81mo runway", vc: C.white },
  ];

  const cardW = 2.0;
  const cardH = 1.08;
  const cardGap = 0.12;
  const cardStartX = 5.8;
  const cardStartY = 0.52;

  kpis.forEach((k, i) => {
    const cx = cardStartX + (i % 2) * (cardW + cardGap);
    const cy = cardStartY + Math.floor(i / 2) * (cardH + cardGap);
    addKpiCard(slide, cx, cy, cardW, cardH, k.value, k.label, k.sub, k.vc, true);
  });

  // Separator
  slide.addShape("rect", {
    x: 5.7, y: 2.85, w: 4.15, h: 0.015,
    fill: { color: "2A3A5E" },
  });

  // Meeting date label
  slide.addText("Board Meeting  ·  Q3 2026  ·  March 30, 2026", {
    x: 5.8, y: 2.9, w: 4.0, h: 0.22,
    fontSize: 8, fontFace: FONT.body,
    color: C.midGray, margin: 0,
  });

  // Quick metrics 2x2 table
  const metricsY = 3.14;
  const tbl = [
    ["Gross Margin", "78.2%", "Burn Rate", "$223K/mo"],
    ["Magic Number", "1.12", "Runway", "81 months"],
    ["Logo Retention", "92.4%", "Employees", "118 (+20 Q3)"],
  ];
  tbl.forEach((row, ri) => {
    const rowY = metricsY + ri * 0.3;
    for (let ci = 0; ci < 4; ci += 2) {
      slide.addText(row[ci], {
        x: 5.8 + (ci / 2) * 2.08, y: rowY, w: 1.0, h: 0.26,
        fontSize: 8, fontFace: FONT.body,
        color: C.midGray, margin: 0, valign: "middle",
      });
      slide.addText(row[ci + 1], {
        x: 5.8 + (ci / 2) * 2.08 + 1.02, y: rowY, w: 1.0, h: 0.26,
        fontSize: 9, fontFace: FONT.body,
        color: C.white, bold: true, margin: 0, valign: "middle",
      });
    }
  });
}

// ------------------------------------------------------------------
// SLIDE 2: Revenue Dashboard (light)
// ------------------------------------------------------------------
function slide02() {
  const slide = addLightSlide();
  let y = addHeader(slide, "Revenue Dashboard — Q3 2026", false);
  addFooter(slide, 2, 6, false);

  // Top KPI cards
  const kpis = [
    { value: "$15.6M", label: "ARR", sub: "+10.6% QoQ", vc: C.green },
    { value: "118%", label: "NRR", sub: "Q2: 121% ↓", vc: C.amber },
    { value: "78.2%", label: "Gross Margin", sub: "Target: 80%+", vc: C.textDark },
    { value: "1.12", label: "Magic Number", sub: "Q2: 0.98 ↑", vc: C.green },
  ];

  const cardW = 2.1;
  const cardH = 0.9;
  const gapX = 0.13;
  const kpiStartX = PAD;
  kpis.forEach((k, i) => {
    addKpiCard(slide, kpiStartX + i * (cardW + gapX), y, cardW, cardH, k.value, k.label, k.sub, k.vc, false);
  });
  y += cardH + SECTION_GAP + 0.08;

  // ARR trend chart (bar)
  const chartY = y;
  const chartH = 1.82;
  y = addSectionLabel(slide, "ARR Trend ($M)", PAD, chartY, 4.6, false);

  slide.addChart(pres.charts.BAR, [
    {
      name: "ARR ($M)",
      labels: ["Q1'25", "Q2'25", "Q3'25", "Q4'25", "Q1'26", "Q2'26", "Q3'26"],
      values: [7.2, 8.5, 9.8, 11.4, 12.8, 14.1, 15.6],
    },
  ], {
    x: PAD, y: chartY + 0.26, w: 4.6, h: chartH - 0.26,
    chartColors: [C.teal],
    showLegend: true, legendPos: "b", legendFontSize: 7,
    showValue: true, dataLabelFontSize: 7, dataLabelColor: C.white,
    catAxisLabelFontSize: 7.5,
    valAxisLabelFontSize: 7.5,
    valAxisMinVal: 0,
    valAxisMaxVal: 18,
    showTitle: false,
    border: { pt: 0, color: C.lightBg },
  });

  y = chartY + chartH + SECTION_GAP + 0.06;

  // Unit economics table (right column starts at chart x)
  const tableX = 5.4;
  const tableW = W - tableX - PAD;

  const ueY = addSectionLabel(slide, "Unit Economics", tableX, BODY_TOP + cardH + SECTION_GAP + 0.08, tableW, false);

  const ueHeaders = ["Metric", "Q3 2026", "Q2 2026", "Target", "Trend"];
  const ueColW = [1.35, 0.72, 0.72, 0.72, 0.69];
  const mkTrend = (text, up) => ({
    text,
    options: { color: up ? C.green : C.red, bold: true },
  });
  const ueRows = [
    ["Gross Margin", "78.2%", "77.5%", "80%+", mkTrend("↑ Improv", true)],
    ["NRR", "118%", "121%", "120%+", mkTrend("↓ Churn", false)],
    ["Logo Retention", "92.4%", "94.1%", "95%+", mkTrend("↓ Declin", false)],
    ["LTV:CAC", "4.8x", "4.6x", "4.0x+", { text: "Healthy", options: { color: C.green } }],
    ["CAC Payback", "13.2 mo", "14.1 mo", "<16 mo", mkTrend("↑ Improv", true)],
    ["Magic Number", "1.12", "0.98", ">0.75", { text: "Strong", options: { color: C.green } }],
  ];

  addTable(slide, ueHeaders, ueRows, tableX, ueY, tableW, ueColW, 0.27);

  // Cash / burn snapshot below chart
  const snapY = Math.max(y, ueY + 0.27 * 7 + SECTION_GAP);
  const snapData = [
    { label: "Cash", value: "$18.2M" },
    { label: "Burn/mo", value: "$223K" },
    { label: "Runway", value: "81 mo" },
    { label: "Burn Multiple", value: "0.36x" },
  ];
  const snapW = (W - PAD * 2) / 4;
  snapData.forEach((s, i) => {
    const sx = PAD + i * snapW;
    slide.addShape("rect", {
      x: sx + 0.04, y: snapY, w: snapW - 0.08, h: 0.5,
      fill: { color: i === 2 ? "D1FAE5" : C.cardBg },
      line: { color: C.border, width: 0.4 },
    });
    slide.addText(s.label, {
      x: sx + 0.08, y: snapY + 0.02, w: snapW * 0.45, h: 0.2,
      fontSize: 7.5, fontFace: FONT.body,
      color: C.slate, margin: 0,
    });
    slide.addText(s.value, {
      x: sx + 0.08, y: snapY + 0.22, w: snapW - 0.16, h: 0.22,
      fontSize: 11, fontFace: FONT.head,
      color: i === 2 ? C.green : C.textDark, bold: true, margin: 0,
    });
  });
}

// ------------------------------------------------------------------
// SLIDE 3: GTM + Customer Health (light)
// ------------------------------------------------------------------
function slide03() {
  const slide = addLightSlide();
  let y = addHeader(slide, "GTM Performance + Customer Health — Q3 2026", false);
  addFooter(slide, 3, 6, false);

  const colLX = PAD;
  const colLW = 4.5;
  const colRX = 5.3;
  const colRW = W - colRX - PAD;

  // LEFT COLUMN — Sales Metrics
  let ly = addSectionLabel(slide, "Sales Performance", colLX, y, colLW, false);

  const salesHeaders = ["Metric", "Q3", "Q2", "QoQ"];
  const salesColW = [1.6, 0.9, 0.9, 0.9];
  const salesRows = [
    ["Pipeline Generated", "$8.6M", "$7.1M", { text: "+21.1%", options: { color: C.green, bold: true } }],
    ["Win Rate", "28.4%", "26.1%", { text: "+2.3pp", options: { color: C.green, bold: true } }],
    ["Avg Deal Size", "$42K ACV", "$38K ACV", { text: "+10.5%", options: { color: C.green, bold: true } }],
    ["Sales Cycle", "68 days", "72 days", { text: "-5.6%", options: { color: C.green, bold: true } }],
    ["Quota Attainment", "112%", "94%", { text: "+18pp", options: { color: C.green, bold: true } }],
  ];
  addTable(slide, salesHeaders, salesRows, colLX, ly, colLW, salesColW, 0.26);
  ly += 0.26 * 6 + SECTION_GAP + 0.1;

  // Channel breakdown
  ly = addSectionLabel(slide, "Pipeline by Channel", colLX, ly, colLW, false);
  const chanHeaders = ["Channel", "Pipeline", "Deals", "Avg ACV", "CAC"];
  const chanColW = [1.35, 0.78, 0.5, 0.6, 0.65];
  const chanRows = [
    ["Outbound SDR", "$3.8M", "18", "$52K", "$24.1K"],
    ["Inbound Mktg", "$2.4M", "14", "$34K", "$12.8K"],
    ["PLG/Self-Serve", "$0.9M", "8", "$18K", "$6.2K"],
    ["Partner/Referral", "$1.5M", "5", "$68K", { text: "$8.4K", options: { color: C.green } }],
  ];
  addTable(slide, chanHeaders, chanRows, colLX, ly, colLW, chanColW, 0.26);

  // RIGHT COLUMN — Churn deep dive
  let ry = addSectionLabel(slide, "Churn Deep Dive — Q3 ($500K Total)", colRX, y, colRW, false);

  const churnHeaders = ["Account", "ARR Lost", "Reason", "Prev?"];
  const churnColW = [1.2, 0.72, 1.6, 0.55];
  const mkPrev = (text, color) => ({ text, options: { color, bold: true } });
  const churnRows = [
    ["Apex Manufacturing", "$145K", "M&A", mkPrev("No", C.slate)],
    ["Precision Dynamics", { text: "$62K", options: { color: C.amber } }, "In-house analytics", mkPrev("Partial", C.amber)],
    ["4 SMB accounts", { text: "$118K", options: { color: C.red } }, "Price / DataForge", mkPrev("YES", C.red)],
    ["TechFab Solutions", { text: "$48K", options: { color: C.red } }, "Poor implementation", mkPrev("YES", C.red)],
    ["3 SMB accounts", { text: "$82K", options: { color: C.red } }, "Low usage/onboarding", mkPrev("YES", C.red)],
    ["Consolidated Parts", "$45K", "Budget cuts", mkPrev("No", C.slate)],
  ];
  addTable(slide, churnHeaders, churnRows, colRX, ry, colRW, churnColW, 0.265);
  ry += 0.265 * 7 + SECTION_GAP + 0.08;

  // Callout box
  slide.addShape("rect", {
    x: colRX, y: ry, w: colRW, h: 0.72,
    fill: { color: "FFF7ED" },
    line: { color: C.amber, width: 1 },
  });
  slide.addText("$290K of $500K (58%) preventable", {
    x: colRX + 0.1, y: ry + 0.06, w: colRW - 0.2, h: 0.22,
    fontSize: 10, fontFace: FONT.head,
    color: C.amber, bold: true, margin: 0,
  });
  slide.addText("Root causes: weak SMB onboarding  ·  no down-tier option  ·  1 CS implementation failure", {
    x: colRX + 0.1, y: ry + 0.29, w: colRW - 0.2, h: 0.36,
    fontSize: 8.5, fontFace: FONT.body,
    color: C.darkSlate, margin: 0,
  });
}

// ------------------------------------------------------------------
// SLIDE 4: Product + Team (light)
// ------------------------------------------------------------------
function slide04() {
  const slide = addLightSlide();
  let y = addHeader(slide, "Product Delivery + Team — Q3 2026", false);
  addFooter(slide, 4, 6, false);

  const colLX = PAD;
  const colLW = 4.5;
  const colRX = 5.3;
  const colRW = W - colRX - PAD;

  // LEFT: Shipped features
  let ly = addSectionLabel(slide, "Q3 Shipped", colLX, y, colLW, false);
  const shipHeaders = ["Feature", "Adoption", "Impact"];
  const shipColW = [1.55, 1.4, 1.35];
  const shipRows = [
    ["Predictive Maint v2", "67% enterprise", "$1.2M pipeline"],
    ["Self-Serve Dashboard", "340 dashboards, 89 accts", "CS tickets -18%"],
    ["SAP Integration", "12 connected", "Partner accelerant"],
    ["SOC 2 Type II", "Completed 8/15", { text: "4 deals unlocked ($380K)", options: { color: C.green } }],
  ];
  addTable(slide, shipHeaders, shipRows, colLX, ly, colLW, shipColW, 0.28);
  ly += 0.28 * 5 + SECTION_GAP + 0.1;

  // LEFT: Headcount
  ly = addSectionLabel(slide, "Headcount by Function", colLX, ly, colLW, false);
  const hcHeaders = ["Function", "Q2", "Q3", "Open", "Q4 Target"];
  const hcColW = [1.2, 0.7, 0.7, 0.65, 0.85];
  const hcRows = [
    ["Engineering", "42", "46", "4", "50"],
    ["Sales", "18", "22", "3", "25"],
    ["CS", "12", "14", "2", "16"],
    ["Other", "26", "30", "3", "33"],
    [
      { text: "Total", options: { bold: true } },
      { text: "98", options: { bold: true } },
      { text: "112", options: { bold: true } },
      { text: "12", options: { bold: true } },
      { text: "124", options: { bold: true, color: C.tealDark } },
    ],
  ];
  addTable(slide, hcHeaders, hcRows, colLX, ly, colLW, hcColW, 0.265);

  // RIGHT: Q4 Roadmap
  let ry = addSectionLabel(slide, "Q4 Roadmap", colRX, y, colRW, false);
  const rdHeaders = ["Pri", "Feature", "Status", "Impact"];
  const rdColW = [0.38, 1.42, 0.95, 1.3];
  const mkP = (text, color) => ({ text, options: { color, bold: true, fill: { color: "F8FAFC" } } });
  const rdRows = [
    [mkP("P0", C.red), "Multi-tenant analytics", "Dev 60%", "2 deals $200K+"],
    [mkP("P0", C.red), "Siemens MindSphere", "Design → Dev", "$4M TAM"],
    [mkP("P1", C.amber), "Usage-based pricing (SMB)", "Spec complete", { text: "Fixes 58% prev churn", options: { color: C.green } }],
    [mkP("P1", C.amber), "Customer health score", "Prototype", "Early warning system"],
    [mkP("P2", C.slate), "Mobile app", "Scoping", "Freq. requested"],
  ];
  addTable(slide, rdHeaders, rdRows, colRX, ry, colRW, rdColW, 0.28);
  ry += 0.28 * 6 + SECTION_GAP + 0.1;

  // Headcount mini-chart (bar) on right
  ry = addSectionLabel(slide, "Quarterly Headcount Growth", colRX, ry, colRW, false);
  slide.addChart(pres.charts.BAR, [
    {
      name: "Headcount",
      labels: ["Q1'25", "Q2'25", "Q3'25", "Q4'25", "Q1'26", "Q2'26", "Q3'26"],
      values: [68, 74, 80, 86, 92, 98, 112],
    },
  ], {
    x: colRX, y: ry, w: colRW, h: 1.52,
    chartColors: [C.teal],
    showLegend: true, legendPos: "b", legendFontSize: 7,
    showValue: true, dataLabelFontSize: 7, dataLabelColor: C.white,
    catAxisLabelFontSize: 7,
    valAxisLabelFontSize: 7,
    valAxisMinVal: 0,
    showTitle: false,
  });
}

// ------------------------------------------------------------------
// SLIDE 5: Financial Outlook + Competitive (light)
// ------------------------------------------------------------------
function slide05() {
  const slide = addLightSlide();
  let y = addHeader(slide, "Financial Outlook + Competitive Landscape — Q3 2026", false);
  addFooter(slide, 5, 6, false);

  const colLX = PAD;
  const colLW = 5.3;
  const colRX = 6.1;
  const colRW = W - colRX - PAD;

  // LEFT: P&L
  let ly = addSectionLabel(slide, "P&L Summary ($M)", colLX, y, colLW, false);
  const plHeaders = ["Line Item", "Q1", "Q2", "Q3", "Q3 YoY", "FY2026E"];
  const plColW = [1.3, 0.62, 0.62, 0.62, 0.72, 0.72];
  const mkYoY = (text, positive) => ({ text, options: { color: positive ? C.green : C.red, bold: true } });
  const plRows = [
    ["Revenue", "$3.4M", "$3.7M", "$4.1M", mkYoY("+64%", true), "$15.4M"],
    ["Gross Profit", "$2.66M", "$2.87M", "$3.21M", mkYoY("+68%", true), "$12.06M"],
    ["Gross Margin", "78.2%", "77.6%", "78.3%", mkYoY("+2.1pp", true), "78.3%"],
    [
      "Net Income",
      { text: "($580K)", options: { color: C.red } },
      { text: "($700K)", options: { color: C.red } },
      { text: "($670K)", options: { color: C.amber } },
      { text: "Improved", options: { color: C.green } },
      { text: "($2.52M)", options: { color: C.amber } },
    ],
  ];
  addTable(slide, plHeaders, plRows, colLX, ly, colLW, plColW, 0.28);
  ly += 0.28 * 5 + SECTION_GAP + 0.08;

  // Cash position bar
  ly = addSectionLabel(slide, "Cash Position", colLX, ly, colLW, false);
  const cashItems = [
    { label: "Cash on Hand", value: "$18.2M", color: C.green },
    { label: "Monthly Burn", value: "$223K", color: C.amber },
    { label: "Runway", value: "81 months", color: C.green },
    { label: "Burn Multiple", value: "0.36x  (Excellent)", color: C.green },
  ];
  cashItems.forEach((ci) => {
    slide.addText(`${ci.label}:`, {
      x: colLX, y: ly, w: 1.4, h: 0.26,
      fontSize: 9, fontFace: FONT.body,
      color: C.slate, margin: 0, valign: "middle",
    });
    slide.addText(ci.value, {
      x: colLX + 1.42, y: ly, w: 2.0, h: 0.26,
      fontSize: 9.5, fontFace: FONT.body,
      color: ci.color, bold: true, margin: 0, valign: "middle",
    });
    ly += 0.27;
  });

  // Revenue chart
  ly += SECTION_GAP;
  ly = addSectionLabel(slide, "Quarterly Revenue Trend", colLX, ly, colLW, false);
  const chartBottom = CONTENT_BOTTOM;
  const chartH = chartBottom - ly;
  if (chartH > 0.6) {
    slide.addChart(pres.charts.LINE, [
      {
        name: "Revenue ($M)",
        labels: ["Q1'25", "Q2'25", "Q3'25", "Q4'25", "Q1'26", "Q2'26", "Q3'26"],
        values: [2.2, 2.55, 2.65, 2.8, 3.4, 3.7, 4.1],
      },
    ], {
      x: colLX, y: ly, w: colLW, h: Math.min(chartH, 1.05),
      chartColors: [C.teal],
      showLegend: true, legendPos: "b", legendFontSize: 7,
      showValue: false,
      catAxisLabelFontSize: 7,
      valAxisLabelFontSize: 7,
      showTitle: false,
      lineDataSymbol: "circle",
      lineDataSymbolSize: 5,
    });
  }

  // RIGHT: Competitive
  let ry = addSectionLabel(slide, "Competitive Landscape", colRX, y, colRW, false);
  const compHeaders = ["", "NovaCrest", "DataForge", "Acme Analytics", "Zenith AI"];
  const compColW = [0.72, 0.7, 0.7, 0.7, 0.7];
  const mkC = (text, highlight) => ({
    text,
    options: highlight ? { color: C.teal, bold: true } : {},
  });
  const compRows = [
    ["Stage", mkC("Series B", true), "Series C", "Public", "Series A"],
    ["ARR", mkC("$15.6M", true), "$28M", "$120M", "$4M"],
    ["Market", mkC("Mfg (pure)", true), "Multi-vert", "Enterprise", "Mfg AI"],
    ["Differentiator", mkC("Domain depth", true), "Breadth", "Scale", "ML speed"],
    ["Win vs", mkC("Zenith: 72%", true), mkC("vs NC: 31%", false), "No overlap", "Early stage"],
    ["Threat Level", mkC("Low", true), { text: "HIGH", options: { color: C.red, bold: true } }, "Low", { text: "WATCH", options: { color: C.amber, bold: true } }],
  ];
  addTable(slide, compHeaders, compRows, colRX, ry, colRW, compColW, 0.275);
  ry += 0.275 * 7 + SECTION_GAP + 0.1;

  // Key risks
  ry = addSectionLabel(slide, "Key Risks", colRX, ry, colRW, false);
  const risks = [
    { text: "DataForge mfg module — monitor win/loss tightly", color: C.red },
    { text: "SMB churn continues without pricing fix (P1)", color: C.amber },
    { text: "CS capacity: 2 FTE open roles slow to close", color: C.amber },
    { text: "Series C timing: raise momentum vs optionality", color: C.slate },
  ];
  risks.forEach((r) => {
    slide.addText(`▸  ${r.text}`, {
      x: colRX, y: ry, w: colRW, h: 0.25,
      fontSize: 8.5, fontFace: FONT.body,
      color: r.color, margin: 0, valign: "middle",
    });
    ry += 0.26;
  });
}

// ------------------------------------------------------------------
// SLIDE 6: Board Asks (dark)
// ------------------------------------------------------------------
function slide06() {
  const slide = addDarkSlide();
  let y = addHeader(slide, "Board Asks — Q3 2026", true);
  addFooter(slide, 6, 6, true);

  const cardH = 1.52;
  const cardW = W - PAD * 2;
  const cardGap = 0.1;

  const asks = [
    {
      tag: "APPROVE",
      num: "#1",
      tagColor: C.green,
      title: "Approve $2.5M CS Investment",
      details: [
        "8 FTEs: 4 onboarding specialists · 2 renewal managers · 1 SMB CS lead · 1 CS ops",
        "Addresses $290K/yr preventable churn — saves ~$700K ARR/yr at full effectiveness",
        "Payback: 14 months  ·  Risk without action: Q4 churn could reach $600K+",
      ],
      roi: "ROI: $700K/yr saved  |  14mo payback  |  NRR recovery from 118% → 122%+",
    },
    {
      tag: "APPROVE",
      num: "#2",
      tagColor: C.green,
      title: "Approve Usage-Based SMB Pricing Tier ($499/mo entry point)",
      details: [
        "Current floor $2.5K/mo loses 60+ qualified prospects per quarter on price alone",
        "Down-tier risk on existing base: ~$200K ARR · Net impact positive within 2 quarters",
        "+$1.2M new ARR in 12 months at current pipeline conversion rates",
      ],
      roi: "ROI: +$1.2M net new ARR in 12 mo  |  Opens 300+ prospect ICP segment",
    },
    {
      tag: "DISCUSS",
      num: "#3",
      tagColor: C.teal,
      title: "Series C Timing: Raise Q2 2027 ($22–25M ARR) or Wait for $30M?",
      details: [
        "Current: 81mo runway — no urgency, but growth is re-accelerating post-Q3",
        "Raise Q2 2027: capitalize on momentum, ARR at $22–25M, competitive window vs DataForge",
        "Wait for $30M: stronger terms, higher valuation, but risk of market timing / DataForge gap",
      ],
      roi: "Board input needed: timing · size · lead investor strategy · existing LP follow-on",
    },
  ];

  asks.forEach((ask, i) => {
    const cardY = y + i * (cardH + cardGap);
    if (cardY + cardH > CONTENT_BOTTOM + 0.05) return;

    // Card background
    slide.addShape("rect", {
      x: PAD, y: cardY, w: cardW, h: cardH,
      fill: { color: C.navy },
      line: { color: "2A3A5E", width: 0.6 },
    });

    // Left accent bar
    slide.addShape("rect", {
      x: PAD, y: cardY, w: 0.06, h: cardH,
      fill: { color: ask.tagColor },
    });

    // Tag badge
    slide.addShape("rect", {
      x: PAD + 0.14, y: cardY + 0.1, w: 0.72, h: 0.22,
      fill: { color: ask.tagColor },
    });
    slide.addText(ask.tag, {
      x: PAD + 0.14, y: cardY + 0.1, w: 0.72, h: 0.22,
      fontSize: 7, fontFace: FONT.body,
      color: C.darkNavy, bold: true, align: "center", valign: "middle", margin: 0,
    });

    // Ask number
    slide.addText(ask.num, {
      x: PAD + 0.9, y: cardY + 0.1, w: 0.35, h: 0.22,
      fontSize: 9, fontFace: FONT.body,
      color: C.midGray, margin: 0,
    });

    // Title
    slide.addText(ask.title, {
      x: PAD + 0.14, y: cardY + 0.34, w: cardW - 0.28, h: 0.28,
      fontSize: 11.5, fontFace: FONT.head,
      color: C.white, bold: true, margin: 0,
    });

    // Details
    ask.details.forEach((d, di) => {
      slide.addText(`▸  ${d}`, {
        x: PAD + 0.18, y: cardY + 0.64 + di * 0.22, w: cardW * 0.75 - 0.1, h: 0.2,
        fontSize: 8.5, fontFace: FONT.body,
        color: "B0C4DE", margin: 0, valign: "middle",
      });
    });

    // ROI box (right side)
    slide.addShape("rect", {
      x: PAD + cardW * 0.76, y: cardY + 0.34, w: cardW * 0.22, h: cardH - 0.44,
      fill: { color: C.darkNavy },
      line: { color: ask.tagColor, width: 0.5 },
    });
    slide.addText(ask.roi, {
      x: PAD + cardW * 0.76 + 0.08, y: cardY + 0.38, w: cardW * 0.22 - 0.12, h: cardH - 0.52,
      fontSize: 8, fontFace: FONT.body,
      color: ask.tagColor, margin: 0, valign: "middle",
    });
  });
}

// ============================================================
// MAIN — BUILD ALL SLIDES
// ============================================================
slide01();
slide02();
slide03();
slide04();
slide05();
slide06();

pres.writeFile({ fileName: "outputs/strategy-deck-builder.pptx" })
  .then(() => console.log("Done: outputs/strategy-deck-builder.pptx"))
  .catch((err) => { console.error(err); process.exit(1); });
