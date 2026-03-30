/**
 * Q3 2026 Strategic Review — Board Deck Generator
 * Series B SaaS company ($15M ARR, 120 employees)
 * 12 slides, investor-ready aesthetic
 */

const fs = require("fs");
const path = require("path");
const pptxgen = require("../node_modules/pptxgenjs");

// ────────────────────────────────────────────
// LAYER 1: Constants & Design Tokens
// ────────────────────────────────────────────
const W = 10, H = 5.625;
const PAD = 0.5;
const TITLE_H = 0.5;
const BODY_TOP = TITLE_H + 0.25;
const BODY_W = W - PAD * 2;
const FOOTER_Y = 5.35;
const CONTENT_BOTTOM = FOOTER_Y - 0.12;
const SECTION_GAP = 0.12;
const MIN_FONT = 9;

// Colors — conservative, investor-ready palette
const C = {
  navy:       "1B2A4A",
  charcoal:   "2D3748",
  darkGray:   "4A5568",
  medGray:    "718096",
  lightGray:  "E2E8F0",
  veryLight:  "F7FAFC",
  white:      "FFFFFF",
  accent:     "2B6CB0",   // muted blue accent
  accentLight:"BEE3F8",
  green:      "276749",
  greenLight: "C6F6D5",
  red:        "9B2C2C",
  redLight:   "FED7D7",
  amber:      "975A16",
  amberLight: "FEFCBF",
};

// Fonts
const FONT = "Calibri";
const TITLE_FS = 20;
const SUBTITLE_FS = 12;
const BODY_FS = 10;
const SMALL_FS = 9;
const TABLE_HEADER_FS = 9;
const TABLE_BODY_FS = 9;

// ────────────────────────────────────────────
// LAYER 1b: Utilities
// ────────────────────────────────────────────
function trimText(text, maxChars) {
  if (text.length <= maxChars) return text;
  const trimmed = text.substring(0, maxChars - 1).replace(/[\s,.;:]+$/, "");
  const lastSpace = trimmed.lastIndexOf(" ");
  return lastSpace > maxChars * 0.5 ? trimmed.substring(0, lastSpace) : trimmed;
}

function fitBullets(items, max, chars) {
  const result = items.slice(0, max).map(b => trimText(b, chars));
  if (items.length > max) {
    console.log(`  fitBullets: dropped ${items.length - max} items`);
  }
  return result;
}

function estimateLines(text, fontSize, boxW) {
  const charsPerLine = Math.floor(boxW * 72 / (fontSize * 0.6));
  return Math.ceil(text.length / charsPerLine);
}

// ────────────────────────────────────────────
// LAYER 2: Helpers
// ────────────────────────────────────────────
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Q3 2026 Strategic Review";
pres.author = "Acme SaaS Inc.";

pres.defineSlideMaster({
  title: "MASTER",
  background: { color: C.white },
  objects: [],
});

pres.defineSlideMaster({
  title: "DARK",
  background: { color: C.navy },
  objects: [],
});

function addHeader(slide, title, opts = {}) {
  const color = opts.dark ? C.white : C.navy;
  slide.addText(title, {
    x: PAD, y: 0.18, w: BODY_W, h: TITLE_H,
    fontSize: TITLE_FS, fontFace: FONT, color, bold: true,
    align: "left", valign: "middle", margin: 0, fit: "shrink",
  });
  // Accent underline
  slide.addShape(pres.shapes.LINE, {
    x: PAD, y: 0.7, w: 1.2, h: 0,
    line: { color: opts.dark ? C.accentLight : C.accent, width: 2.5 },
  });
}

function addFooter(slide, pageNum, opts = {}) {
  const color = opts.dark ? C.medGray : C.medGray;
  // Thin line
  slide.addShape(pres.shapes.LINE, {
    x: PAD, y: FOOTER_Y, w: BODY_W, h: 0,
    line: { color: C.lightGray, width: 0.5 },
  });
  slide.addText(`Acme SaaS Inc.  |  Confidential  |  ${pageNum}`, {
    x: PAD, y: FOOTER_Y + 0.02, w: BODY_W, h: 0.22,
    fontSize: 7, fontFace: FONT, color, align: "right", margin: 0,
  });
}

function addSectionLabel(slide, text, y, opts = {}) {
  const color = opts.color || C.charcoal;
  slide.addText(text.toUpperCase(), {
    x: PAD, y, w: BODY_W, h: 0.22,
    fontSize: 9, fontFace: FONT, color, bold: true,
    charSpacing: 2, margin: 0,
  });
  slide.addShape(pres.shapes.LINE, {
    x: PAD, y: y + 0.22, w: BODY_W, h: 0,
    line: { color: C.lightGray, width: 0.5 },
  });
  return y + 0.3;
}

function addBullets(slide, items, x, y, w, h, fs, opts = {}) {
  const fontSize = Math.max(fs || BODY_FS, MIN_FONT);
  const color = opts.color || C.charcoal;
  const textArr = items.map((item, i) => ({
    text: item,
    options: {
      bullet: true,
      breakLine: i < items.length - 1,
      fontSize,
      color,
    },
  }));
  slide.addText(textArr, {
    x, y, w, h,
    fontFace: FONT, valign: "top", margin: [0, 0, 0, 4],
    lineSpacingMultiple: 1.15,
  });
}

function addSubHeader(slide, text, x, y, w) {
  slide.addText(text, {
    x, y, w, h: 0.25,
    fontSize: 11, fontFace: FONT, color: C.navy, bold: true,
    margin: 0,
  });
  return y + 0.28;
}

function twoColumnLayout(gap) {
  gap = gap || 0.3;
  const colW = (BODY_W - gap) / 2;
  return { leftX: PAD, rightX: PAD + colW + gap, colW };
}

function cardShadow() {
  return { type: "outer", blur: 3, offset: 1, color: "000000", opacity: 0.08, angle: 135 };
}

function addCard(slide, title, body, x, y, w, h, opts = {}) {
  const accentColor = opts.accent || C.accent;
  // Card background
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h,
    fill: { color: C.white },
    line: { color: C.lightGray, width: 0.5 },
    rectRadius: 0.06,
    shadow: cardShadow(),
  });
  // Top accent bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x: x + 0.01, y, w: w - 0.02, h: 0.04,
    fill: { color: accentColor },
  });
  // Title
  slide.addText(title, {
    x: x + 0.12, y: y + 0.08, w: w - 0.24, h: 0.22,
    fontSize: 10, fontFace: FONT, color: C.navy, bold: true, margin: 0,
  });
  // Body
  if (typeof body === "string") {
    slide.addText(body, {
      x: x + 0.12, y: y + 0.32, w: w - 0.24, h: h - 0.44,
      fontSize: MIN_FONT, fontFace: FONT, color: C.darkGray, margin: 0,
      valign: "top", lineSpacingMultiple: 1.15,
    });
  } else if (Array.isArray(body)) {
    addBullets(slide, body, x + 0.12, y + 0.32, w - 0.24, h - 0.44, MIN_FONT);
  }
}

function addKPIBox(slide, label, value, x, y, w, h, opts = {}) {
  const bgColor = opts.bg || C.veryLight;
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h,
    fill: { color: bgColor },
    line: { color: C.lightGray, width: 0.5 },
    rectRadius: 0.06,
  });
  slide.addText(value, {
    x, y: y + 0.06, w, h: h * 0.5,
    fontSize: 18, fontFace: FONT, color: opts.valueColor || C.navy,
    bold: true, align: "center", margin: 0,
  });
  slide.addText(label, {
    x, y: y + h * 0.5, w, h: h * 0.4,
    fontSize: 8, fontFace: FONT, color: C.medGray,
    align: "center", margin: 0,
  });
}

function addChip(slide, text, x, y, opts = {}) {
  const bg = opts.bg || C.accentLight;
  const color = opts.color || C.accent;
  const chipW = Math.max(text.length * 0.065 + 0.2, 0.6);
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w: chipW, h: 0.22,
    fill: { color: bg }, rectRadius: 0.11,
  });
  slide.addText(text, {
    x, y, w: chipW, h: 0.22,
    fontSize: 7, fontFace: FONT, color, bold: true,
    align: "center", valign: "middle", margin: 0,
  });
  return x + chipW + 0.08;
}

// Table helper with consistent styling
function styledTable(slide, headers, rows, x, y, w, opts = {}) {
  const colW = opts.colW || headers.map(() => w / headers.length);
  const headerRow = headers.map(h => ({
    text: h,
    options: {
      fill: { color: C.navy },
      color: C.white,
      bold: true,
      fontSize: TABLE_HEADER_FS,
      fontFace: FONT,
      align: "left",
      valign: "middle",
      margin: [2, 4, 2, 4],
    },
  }));
  const dataRows = rows.map((row, ri) =>
    row.map((cell, ci) => {
      const isObj = typeof cell === "object" && cell !== null && cell.text !== undefined;
      return {
        text: isObj ? cell.text : String(cell),
        options: {
          fill: { color: ri % 2 === 0 ? C.veryLight : C.white },
          color: (isObj && cell.color) || C.charcoal,
          bold: (isObj && cell.bold) || false,
          fontSize: TABLE_BODY_FS,
          fontFace: FONT,
          align: (isObj && cell.align) || "left",
          valign: "middle",
          margin: [2, 4, 2, 4],
        },
      };
    })
  );
  slide.addTable([headerRow, ...dataRows], {
    x, y, w,
    colW,
    border: { pt: 0.5, color: C.lightGray },
  });
}

// ────────────────────────────────────────────
// LAYER 3: Slide Definitions (12 Slides)
// ────────────────────────────────────────────

let pageNum = 0;

// =============================================
// SLIDE 1: Cover
// =============================================
{
  pageNum++;
  const s = pres.addSlide({ masterName: "DARK" });

  // Subtle geometric accent — large circle
  s.addShape(pres.shapes.OVAL, {
    x: 6.5, y: -1.5, w: 5.5, h: 5.5,
    fill: { color: C.accent, transparency: 88 },
  });
  s.addShape(pres.shapes.OVAL, {
    x: 7.5, y: 2.0, w: 3.5, h: 3.5,
    fill: { color: C.accent, transparency: 92 },
  });

  // Logo placeholder
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: PAD, y: 0.6, w: 1.8, h: 0.55,
    fill: { color: C.accent },
    rectRadius: 0.06,
  });
  s.addText("ACME", {
    x: PAD, y: 0.6, w: 1.8, h: 0.55,
    fontSize: 16, fontFace: FONT, color: C.white, bold: true,
    align: "center", valign: "middle", margin: 0,
    charSpacing: 4,
  });

  let y = 1.65;
  s.addText("Q3 2026\nStrategic Review", {
    x: PAD, y, w: 7, h: 1.2,
    fontSize: 36, fontFace: FONT, color: C.white, bold: true,
    margin: 0, lineSpacingMultiple: 1.05,
  });
  y += 1.35;

  s.addText("Board of Directors Meeting", {
    x: PAD, y, w: 6, h: 0.35,
    fontSize: 16, fontFace: FONT, color: C.accentLight, margin: 0,
  });
  y += 0.45;

  s.addText("September 15, 2026  |  Series B  |  $15M ARR  |  120 Employees", {
    x: PAD, y, w: 7, h: 0.25,
    fontSize: 11, fontFace: FONT, color: C.medGray, margin: 0,
  });
  y += 0.5;

  s.addText("CONFIDENTIAL", {
    x: PAD, y, w: 2, h: 0.22,
    fontSize: 8, fontFace: FONT, color: C.medGray, margin: 0,
    charSpacing: 3,
  });
}

// =============================================
// SLIDE 2: Executive Summary
// =============================================
{
  pageNum++;
  const s = pres.addSlide({ masterName: "MASTER" });
  addHeader(s, "Executive Summary");
  addFooter(s, pageNum);

  let y = BODY_TOP + 0.12;

  // Key Wins
  y = addSubHeader(s, "Key Wins", PAD, y, BODY_W);
  const wins = [
    "Crossed $15M ARR milestone, up 18% QoQ driven by enterprise segment expansion",
    "Launched AI-powered analytics module; 42% of existing customers adopted within 8 weeks",
    "Closed 3 new enterprise logos (>$200K ACV each), including Fortune 500 financial services firm",
  ];
  addBullets(s, wins, PAD + 0.15, y, BODY_W - 0.15, 0.72, BODY_FS, { color: C.green });
  y += 0.78;

  // Concerns
  y = addSubHeader(s, "Concerns", PAD, y, BODY_W);
  const concerns = [
    "Mid-market churn increased to 3.2% monthly (up from 2.1%) — onboarding quality gap identified",
    "Engineering hiring 40% behind plan; 6 senior roles open >90 days, impacting roadmap velocity",
  ];
  addBullets(s, concerns, PAD + 0.15, y, BODY_W - 0.15, 0.52, BODY_FS, { color: C.red });
  y += 0.58;

  // Decision Needed
  y = addSubHeader(s, "Decision Needed", PAD, y, BODY_W);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: PAD + 0.15, y, w: BODY_W - 0.3, h: 0.55,
    fill: { color: "FFF5F5" },
    line: { color: C.red, width: 1 },
    rectRadius: 0.06,
  });
  s.addText("Approve $2.5M additional investment in customer success team (8 FTEs) to address mid-market churn before it impacts net retention below 110%. Board vote required.", {
    x: PAD + 0.3, y: y + 0.04, w: BODY_W - 0.6, h: 0.47,
    fontSize: BODY_FS, fontFace: FONT, color: C.red, margin: 0,
    valign: "middle", lineSpacingMultiple: 1.15,
  });
}

// =============================================
// SLIDE 3: Revenue Dashboard
// =============================================
{
  pageNum++;
  const s = pres.addSlide({ masterName: "MASTER" });
  addHeader(s, "Revenue Dashboard");
  addFooter(s, pageNum);

  let y = BODY_TOP + 0.05;

  // KPI row
  const kpiW = (BODY_W - 0.36) / 4;
  const kpis = [
    { label: "ARR", value: "$15.0M", bg: C.veryLight, vc: C.navy },
    { label: "QoQ Growth", value: "+18%", bg: C.greenLight, vc: C.green },
    { label: "Net Retention", value: "118%", bg: C.veryLight, vc: C.navy },
    { label: "Pipeline", value: "$8.2M", bg: C.accentLight, vc: C.accent },
  ];
  kpis.forEach((k, i) => {
    addKPIBox(s, k.label, k.value, PAD + i * (kpiW + 0.12), y, kpiW, 0.65, { bg: k.bg, valueColor: k.vc });
  });
  y += 0.8;

  // ARR growth chart (left)
  const { leftX, rightX, colW } = twoColumnLayout(0.4);
  y = addSectionLabel(s, "ARR Trend ($M)", y);

  s.addChart(pres.charts.BAR, [{
    name: "ARR",
    labels: ["Q1 '25", "Q2 '25", "Q3 '25", "Q4 '25", "Q1 '26", "Q2 '26", "Q3 '26"],
    values: [8.5, 9.2, 10.1, 11.0, 12.2, 12.7, 15.0],
  }], {
    x: leftX, y, w: colW, h: 2.8,
    barDir: "col",
    chartColors: [C.accent],
    valGridLine: { color: C.lightGray, size: 0.5 },
    catGridLine: { style: "none" },
    showValue: true,
    dataLabelPosition: "outEnd",
    dataLabelFontSize: 7,
    dataLabelColor: C.charcoal,
    catAxisLabelFontSize: 8,
    valAxisLabelFontSize: 7,
    showLegend: false,
    valAxisHidden: true,
  });

  // Pipeline by segment table (right)
  styledTable(s,
    ["Segment", "Pipeline", "Win Rate", "Avg Deal"],
    [
      ["Enterprise", "$4.1M", "32%", "$185K"],
      ["Mid-Market", "$2.8M", "28%", "$45K"],
      ["SMB", "$1.3M", "41%", "$12K"],
    ],
    rightX, y + 0.1, colW,
    { colW: [colW * 0.28, colW * 0.24, colW * 0.24, colW * 0.24] }
  );

  // Net retention note
  s.addText("Net retention driven by 24% expansion revenue in enterprise segment; SMB net retention flat at 101%.", {
    x: rightX, y: y + 1.3, w: colW, h: 0.45,
    fontSize: 8, fontFace: FONT, color: C.medGray, margin: 0,
    italic: true, valign: "top",
  });
}

// =============================================
// SLIDE 4: Product Update
// =============================================
{
  pageNum++;
  const s = pres.addSlide({ masterName: "MASTER" });
  addHeader(s, "Product Update");
  addFooter(s, pageNum);

  let y = BODY_TOP + 0.05;
  const { leftX, rightX, colW } = twoColumnLayout(0.3);

  // Left: Shipped Features
  y = addSubHeader(s, "Shipped This Quarter", leftX, y, colW);
  const shipped = [
    "AI Analytics Module — predictive churn scoring, usage anomaly detection",
    "SSO/SCIM provisioning for enterprise (SOC2 prerequisite)",
    "API v3 with webhook support and rate limiting overhaul",
    "Dashboard redesign — 35% reduction in time-to-insight (user studies)",
  ];
  addBullets(s, shipped, leftX + 0.1, y, colW - 0.1, 1.2, MIN_FONT);
  const rightY0 = BODY_TOP + 0.05;

  // Right: Adoption Metrics
  let ry = addSubHeader(s, "Adoption Metrics", rightX, rightY0, colW);
  styledTable(s,
    ["Feature", "Adoption", "Trend"],
    [
      ["AI Analytics", "42%", { text: "+12% MoM", color: C.green }],
      ["SSO/SCIM", "68%", { text: "Enterprise", color: C.accent }],
      ["API v3", "31%", { text: "+8% MoM", color: C.green }],
      ["New Dashboard", "89%", { text: "Stable", color: C.medGray }],
    ],
    rightX, ry, colW,
    { colW: [colW * 0.38, colW * 0.24, colW * 0.38] }
  );

  // Roadmap Priorities (full width, bottom half)
  y = BODY_TOP + 0.05 + 0.28 + 1.2 + SECTION_GAP;
  y = addSectionLabel(s, "Q4 Roadmap Priorities", y);
  const cards = [
    { title: "Multi-tenant Analytics", body: "Platform analytics for reseller partners; $1.2M pipeline attached", accent: C.accent },
    { title: "HIPAA Compliance", body: "Healthcare vertical unlock; 4 enterprise prospects waiting on compliance", accent: C.green },
    { title: "Workflow Automation", body: "No-code automation builder; top-requested feature by mid-market segment", accent: C.amber },
  ];
  const cardW = (BODY_W - 0.24) / 3;
  cards.forEach((c, i) => {
    addCard(s, c.title, c.body, PAD + i * (cardW + 0.12), y, cardW, 0.85, { accent: c.accent });
  });
}

// =============================================
// SLIDE 5: Go-to-Market
// =============================================
{
  pageNum++;
  const s = pres.addSlide({ masterName: "MASTER" });
  addHeader(s, "Go-to-Market Performance");
  addFooter(s, pageNum);

  let y = BODY_TOP + 0.05;

  // KPI row
  const kpiW = (BODY_W - 0.36) / 4;
  const kpis = [
    { label: "Blended CAC", value: "$18.2K", bg: C.veryLight, vc: C.navy },
    { label: "CAC Payback", value: "14 mo", bg: C.amberLight, vc: C.amber },
    { label: "LTV:CAC", value: "4.2x", bg: C.greenLight, vc: C.green },
    { label: "Expansion Rev", value: "$1.8M", bg: C.accentLight, vc: C.accent },
  ];
  kpis.forEach((k, i) => {
    addKPIBox(s, k.label, k.value, PAD + i * (kpiW + 0.12), y, kpiW, 0.65, { bg: k.bg, valueColor: k.vc });
  });
  y += 0.82;

  const { leftX, rightX, colW } = twoColumnLayout(0.4);

  // CAC Trend chart (left)
  s.addChart(pres.charts.LINE, [
    { name: "Enterprise CAC", labels: ["Q1", "Q2", "Q3", "Q4", "Q1", "Q2", "Q3"], values: [32, 30, 28, 27, 25, 24, 22] },
    { name: "Mid-Market CAC", labels: ["Q1", "Q2", "Q3", "Q4", "Q1", "Q2", "Q3"], values: [15, 14, 14, 15, 16, 17, 18] },
  ], {
    x: leftX, y, w: colW, h: 2.5,
    chartColors: [C.accent, C.amber],
    valGridLine: { color: C.lightGray, size: 0.5 },
    catGridLine: { style: "none" },
    showLegend: true,
    legendPos: "b",
    legendFontSize: 7,
    catAxisLabelFontSize: 8,
    valAxisLabelFontSize: 7,
    lineDataSymbol: "circle",
    lineDataSymbolSize: 5,
  });

  // Channel performance table (right)
  s.addText("Channel Performance", {
    x: rightX, y: y - 0.02, w: colW, h: 0.22,
    fontSize: 10, fontFace: FONT, color: C.navy, bold: true, margin: 0,
  });

  styledTable(s,
    ["Channel", "Revenue", "% of New", "CAC"],
    [
      ["Outbound Sales", "$3.2M", "38%", "$24K"],
      ["Inbound/Content", "$2.1M", "25%", "$11K"],
      ["Partnerships", "$1.6M", "19%", "$15K"],
      ["PLG/Self-Serve", "$1.5M", "18%", "$8K"],
    ],
    rightX, y + 0.25, colW,
    { colW: [colW * 0.32, colW * 0.22, colW * 0.22, colW * 0.24] }
  );

  // Expansion note
  s.addText("Expansion revenue grew 34% QoQ, now representing 22% of total new ARR. Cross-sell into AI Analytics module driving bulk of expansion.", {
    x: rightX, y: y + 1.7, w: colW, h: 0.4,
    fontSize: 8, fontFace: FONT, color: C.medGray, margin: 0, italic: true,
  });
}

// =============================================
// SLIDE 6: Customer Health
// =============================================
{
  pageNum++;
  const s = pres.addSlide({ masterName: "MASTER" });
  addHeader(s, "Customer Health");
  addFooter(s, pageNum);

  let y = BODY_TOP + 0.05;
  const { leftX, rightX, colW } = twoColumnLayout(0.4);

  // NPS chart (left)
  s.addText("NPS by Segment", {
    x: leftX, y, w: colW, h: 0.22,
    fontSize: 10, fontFace: FONT, color: C.navy, bold: true, margin: 0,
  });
  s.addChart(pres.charts.BAR, [{
    name: "NPS",
    labels: ["Enterprise", "Mid-Market", "SMB", "Overall"],
    values: [62, 38, 45, 48],
  }], {
    x: leftX, y: y + 0.25, w: colW, h: 1.8,
    barDir: "bar",
    chartColors: [C.accent],
    valGridLine: { color: C.lightGray, size: 0.5 },
    catGridLine: { style: "none" },
    showValue: true,
    dataLabelPosition: "outEnd",
    dataLabelFontSize: 8,
    dataLabelColor: C.charcoal,
    catAxisLabelFontSize: 9,
    valAxisHidden: true,
    showLegend: false,
  });

  // Churn analysis (right)
  s.addText("Monthly Churn Rate", {
    x: rightX, y, w: colW, h: 0.22,
    fontSize: 10, fontFace: FONT, color: C.navy, bold: true, margin: 0,
  });
  s.addChart(pres.charts.LINE, [{
    name: "Churn %",
    labels: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep"],
    values: [1.8, 1.9, 2.0, 2.1, 2.3, 2.5, 2.8, 3.0, 3.2],
  }], {
    x: rightX, y: y + 0.25, w: colW, h: 1.8,
    chartColors: [C.red],
    valGridLine: { color: C.lightGray, size: 0.5 },
    catGridLine: { style: "none" },
    showLegend: false,
    catAxisLabelFontSize: 7,
    valAxisLabelFontSize: 7,
    lineDataSymbol: "circle",
    lineDataSymbolSize: 4,
  });

  // Accounts at Risk table (full width)
  y += 2.25;
  y = addSectionLabel(s, "Top Accounts at Risk", y);
  styledTable(s,
    ["Account", "ARR", "Risk Signal", "Action Plan", "Owner"],
    [
      ["GlobalBank Corp", "$420K", { text: "Champion departed", color: C.red }, "Exec sponsor meeting scheduled 9/20", "VP Sales"],
      ["MedTech Solutions", "$280K", { text: "Usage down 40%", color: C.red }, "CSM-led reactivation campaign launched", "CS Lead"],
      ["RetailPro Inc", "$195K", { text: "Contract up 10/31", color: C.amber }, "Renewal proposal with AI module incentive", "AE"],
    ],
    PAD, y, BODY_W,
    { colW: [BODY_W * 0.18, BODY_W * 0.1, BODY_W * 0.2, BODY_W * 0.38, BODY_W * 0.14] }
  );
}

// =============================================
// SLIDE 7: Team & Hiring
// =============================================
{
  pageNum++;
  const s = pres.addSlide({ masterName: "MASTER" });
  addHeader(s, "Team & Hiring");
  addFooter(s, pageNum);

  let y = BODY_TOP + 0.05;
  const { leftX, rightX, colW } = twoColumnLayout(0.4);

  // Headcount by function chart (left)
  s.addText("Headcount by Function (120 total)", {
    x: leftX, y, w: colW, h: 0.22,
    fontSize: 10, fontFace: FONT, color: C.navy, bold: true, margin: 0,
  });
  s.addChart(pres.charts.BAR, [{
    name: "Headcount",
    labels: ["Engineering", "Sales", "CS/Support", "Marketing", "G&A", "Product"],
    values: [45, 28, 18, 14, 9, 6],
  }], {
    x: leftX, y: y + 0.25, w: colW, h: 2.2,
    barDir: "bar",
    chartColors: [C.navy],
    valGridLine: { color: C.lightGray, size: 0.5 },
    catGridLine: { style: "none" },
    showValue: true,
    dataLabelPosition: "outEnd",
    dataLabelFontSize: 8,
    dataLabelColor: C.charcoal,
    catAxisLabelFontSize: 8,
    valAxisHidden: true,
    showLegend: false,
  });

  // Open roles & attrition (right)
  s.addText("Open Roles (14 total)", {
    x: rightX, y, w: colW, h: 0.22,
    fontSize: 10, fontFace: FONT, color: C.navy, bold: true, margin: 0,
  });

  styledTable(s,
    ["Role", "Count", "Days Open", "Status"],
    [
      ["Sr. Engineers", "6", ">90", { text: "Critical", color: C.red, bold: true }],
      ["AEs", "3", "45", { text: "Pipeline", color: C.amber }],
      ["CS Managers", "2", "30", { text: "Screening", color: C.accent }],
      ["Product Mgrs", "2", "60", { text: "Final round", color: C.green }],
      ["Marketing", "1", "15", { text: "New", color: C.medGray }],
    ],
    rightX, y + 0.25, colW,
    { colW: [colW * 0.32, colW * 0.15, colW * 0.22, colW * 0.31] }
  );

  // Attrition section
  y += 2.65;
  y = addSectionLabel(s, "Attrition & Retention", y);
  const { leftX: lx2, rightX: rx2, colW: cw2 } = twoColumnLayout(0.4);
  const attrCards = [
    { title: "Q3 Attrition", body: "4.2% quarterly (annualized 16.8%)\n3 engineering departures to FAANG\n1 sales departure (performance)", accent: C.red },
    { title: "Retention Actions", body: "Equity refresh for top 20% performers\nEngineering career ladder launched\nRemote-first policy expanded globally", accent: C.green },
  ];
  addCard(s, attrCards[0].title, attrCards[0].body, lx2, y, cw2, 0.85, { accent: attrCards[0].accent });
  addCard(s, attrCards[1].title, attrCards[1].body, rx2, y, cw2, 0.85, { accent: attrCards[1].accent });
}

// =============================================
// SLIDE 8: Competitive Landscape
// =============================================
{
  pageNum++;
  const s = pres.addSlide({ masterName: "MASTER" });
  addHeader(s, "Competitive Landscape");
  addFooter(s, pageNum);

  let y = BODY_TOP + 0.1;

  // Competitive positioning table
  styledTable(s,
    ["", "Acme (Us)", "CompetitorA", "CompetitorB", "CompetitorC"],
    [
      [{ text: "Target Segment", bold: true }, "Mid-Market + Enterprise", "Enterprise only", "SMB focus", "Horizontal"],
      [{ text: "AI/ML Capabilities", bold: true }, { text: "Native (GA)", color: C.green, bold: true }, "Partnership", "None", "Beta"],
      [{ text: "Avg ACV", bold: true }, "$52K", "$180K", "$8K", "$24K"],
      [{ text: "Net Retention", bold: true }, { text: "118%", color: C.green, bold: true }, "125%", "95%", "108%"],
      [{ text: "Funding", bold: true }, "Series B ($32M)", "Series D ($120M)", "Series A ($15M)", "Series C ($65M)"],
      [{ text: "Est. ARR", bold: true }, "$15M", "$85M", "$6M", "$38M"],
      [{ text: "Key Weakness", bold: true }, "Scale/brand", "Slow to ship", "Churn / upmarket", "No vertical depth"],
    ],
    PAD, y, BODY_W,
    { colW: [BODY_W * 0.18, BODY_W * 0.22, BODY_W * 0.20, BODY_W * 0.20, BODY_W * 0.20] }
  );

  // Competitive insights at bottom
  y += 2.95;
  y = addSectionLabel(s, "Competitive Takeaways", y);
  const insights = [
    "Win rate vs CompetitorA improved to 38% (from 29%) — AI module is key differentiator in head-to-head",
    "CompetitorC raising Series D; expect increased enterprise competition in Q4",
    "Our moat: vertical-specific workflows + native AI analytics; competitors require third-party integrations",
  ];
  addBullets(s, insights, PAD + 0.1, y, BODY_W - 0.1, 0.8, MIN_FONT);
}

// =============================================
// SLIDE 9: Financial Outlook
// =============================================
{
  pageNum++;
  const s = pres.addSlide({ masterName: "MASTER" });
  addHeader(s, "Financial Outlook");
  addFooter(s, pageNum);

  let y = BODY_TOP + 0.05;

  // KPI row
  const kpiW = (BODY_W - 0.48) / 5;
  const kpis = [
    { label: "Monthly Burn", value: "$1.1M", bg: C.veryLight, vc: C.navy },
    { label: "Cash Balance", value: "$18.4M", bg: C.veryLight, vc: C.navy },
    { label: "Runway", value: "16.7 mo", bg: C.greenLight, vc: C.green },
    { label: "Gross Margin", value: "78%", bg: C.veryLight, vc: C.navy },
    { label: "Breakeven", value: "Q2 2028", bg: C.amberLight, vc: C.amber },
  ];
  kpis.forEach((k, i) => {
    addKPIBox(s, k.label, k.value, PAD + i * (kpiW + 0.12), y, kpiW, 0.65, { bg: k.bg, valueColor: k.vc });
  });
  y += 0.82;

  const { leftX, rightX, colW } = twoColumnLayout(0.4);

  // Burn & revenue trend (left)
  s.addChart(pres.charts.BAR, [
    { name: "Revenue", labels: ["Q1 '26", "Q2 '26", "Q3 '26", "Q4 '26E", "Q1 '27E"], values: [3.1, 3.2, 3.75, 4.2, 4.8] },
    { name: "OpEx", labels: ["Q1 '26", "Q2 '26", "Q3 '26", "Q4 '26E", "Q1 '27E"], values: [3.8, 3.9, 4.1, 4.3, 4.4] },
  ], {
    x: leftX, y, w: colW, h: 2.5,
    barDir: "col",
    barGrouping: "clustered",
    chartColors: [C.accent, C.medGray],
    valGridLine: { color: C.lightGray, size: 0.5 },
    catGridLine: { style: "none" },
    showLegend: true,
    legendPos: "b",
    legendFontSize: 7,
    catAxisLabelFontSize: 8,
    valAxisLabelFontSize: 7,
    showValue: true,
    dataLabelFontSize: 7,
    dataLabelPosition: "outEnd",
    dataLabelColor: C.charcoal,
  });

  // Path to profitability (right)
  s.addText("Path to Profitability", {
    x: rightX, y: y - 0.02, w: colW, h: 0.22,
    fontSize: 10, fontFace: FONT, color: C.navy, bold: true, margin: 0,
  });

  styledTable(s,
    ["Metric", "Current", "Q4 '26E", "Q2 '27E"],
    [
      ["Revenue/mo", "$1.25M", "$1.40M", "$1.60M"],
      ["OpEx/mo", "$1.37M", "$1.43M", "$1.47M"],
      ["Net Burn/mo", { text: "-$120K", color: C.red }, { text: "-$30K", color: C.amber }, { text: "+$130K", color: C.green }],
      ["Headcount", "120", "132", "140"],
      ["Rev/Employee", "$125K", "$127K", "$137K"],
    ],
    rightX, y + 0.25, colW,
    { colW: [colW * 0.32, colW * 0.22, colW * 0.23, colW * 0.23] }
  );
}

// =============================================
// SLIDE 10: Key Risks & Mitigations
// =============================================
{
  pageNum++;
  const s = pres.addSlide({ masterName: "MASTER" });
  addHeader(s, "Key Risks & Mitigations");
  addFooter(s, pageNum);

  let y = BODY_TOP + 0.1;

  const risks = [
    {
      risk: "Mid-market churn acceleration",
      severity: "HIGH",
      sevColor: C.red,
      sevBg: C.redLight,
      impact: "Could reduce NRR below 110%, impacting growth narrative for Series C",
      mitigation: "CS team expansion (board ask); onboarding overhaul underway; automated health scoring shipping Q4",
    },
    {
      risk: "Engineering hiring shortfall",
      severity: "HIGH",
      sevColor: C.red,
      sevBg: C.redLight,
      impact: "Roadmap velocity at 60% of plan; HIPAA compliance may slip to Q1 2027",
      mitigation: "Engaged 2 specialist recruiters; adjusted comp bands +15%; exploring acqui-hire targets",
    },
    {
      risk: "CompetitorC Series D funding",
      severity: "MED",
      sevColor: C.amber,
      sevBg: C.amberLight,
      impact: "Increased enterprise competition expected Q4; potential pricing pressure",
      mitigation: "Accelerating AI moat; deepening vertical integrations; locking in renewals ahead of competitive push",
    },
    {
      risk: "Macro environment / budget freezes",
      severity: "MED",
      sevColor: C.amber,
      sevBg: C.amberLight,
      impact: "Enterprise deal cycles extending 15-20%; 2 deals slipped from Q3 to Q4",
      mitigation: "Flexible pricing/payment terms; ROI calculator for procurement; multi-year discount incentives",
    },
  ];

  // Build as table
  const headerRow = [
    { text: "Risk", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: TABLE_HEADER_FS, fontFace: FONT, align: "left", valign: "middle", margin: [2, 4, 2, 4] } },
    { text: "Severity", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: TABLE_HEADER_FS, fontFace: FONT, align: "center", valign: "middle", margin: [2, 4, 2, 4] } },
    { text: "Impact", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: TABLE_HEADER_FS, fontFace: FONT, align: "left", valign: "middle", margin: [2, 4, 2, 4] } },
    { text: "Mitigation", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: TABLE_HEADER_FS, fontFace: FONT, align: "left", valign: "middle", margin: [2, 4, 2, 4] } },
  ];
  const dataRows = risks.map((r, ri) => [
    { text: r.risk, options: { fill: { color: ri % 2 === 0 ? C.veryLight : C.white }, color: C.charcoal, bold: true, fontSize: TABLE_BODY_FS, fontFace: FONT, margin: [2, 4, 2, 4] } },
    { text: r.severity, options: { fill: { color: r.sevBg }, color: r.sevColor, bold: true, fontSize: TABLE_BODY_FS, fontFace: FONT, align: "center", margin: [2, 4, 2, 4] } },
    { text: r.impact, options: { fill: { color: ri % 2 === 0 ? C.veryLight : C.white }, color: C.darkGray, fontSize: TABLE_BODY_FS, fontFace: FONT, margin: [2, 4, 2, 4] } },
    { text: r.mitigation, options: { fill: { color: ri % 2 === 0 ? C.veryLight : C.white }, color: C.darkGray, fontSize: TABLE_BODY_FS, fontFace: FONT, margin: [2, 4, 2, 4] } },
  ]);

  s.addTable([headerRow, ...dataRows], {
    x: PAD, y, w: BODY_W,
    colW: [BODY_W * 0.2, BODY_W * 0.1, BODY_W * 0.35, BODY_W * 0.35],
    border: { pt: 0.5, color: C.lightGray },
  });
}

// =============================================
// SLIDE 11: Board Asks
// =============================================
{
  pageNum++;
  const s = pres.addSlide({ masterName: "MASTER" });
  addHeader(s, "Board Asks");
  addFooter(s, pageNum);

  let y = BODY_TOP + 0.15;

  s.addText("What we need from the board this quarter", {
    x: PAD, y, w: BODY_W, h: 0.28,
    fontSize: 12, fontFace: FONT, color: C.medGray, italic: true, margin: 0,
  });
  y += 0.4;

  const asks = [
    {
      num: "1",
      title: "Approve CS Team Investment ($2.5M)",
      desc: "8 additional FTEs across Customer Success, onboarding, and support. Required to address mid-market churn before it impacts net retention below 110%. Payback expected within 9 months via retained revenue.",
      urgency: "Vote Required",
      urgencyColor: C.red,
      urgencyBg: C.redLight,
    },
    {
      num: "2",
      title: "Introductions to Enterprise Prospects",
      desc: "Board network activation for 3 target accounts: MegaRetail, HealthFirst Systems, and Pacific Financial. Combined TAM of $850K ACV. Warm introductions to CTO/CIO level contacts.",
      urgency: "This Quarter",
      urgencyColor: C.amber,
      urgencyBg: C.amberLight,
    },
    {
      num: "3",
      title: "Series C Timing Discussion",
      desc: "With 16.7 months of runway and clear path to breakeven, discuss optimal Series C timing. Management recommends Q2 2027 raise at $20M+ ARR for strongest valuation leverage.",
      urgency: "Discussion",
      urgencyColor: C.accent,
      urgencyBg: C.accentLight,
    },
  ];

  asks.forEach((ask) => {
    // Card
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: PAD, y, w: BODY_W, h: 1.05,
      fill: { color: C.white },
      line: { color: C.lightGray, width: 0.75 },
      rectRadius: 0.06,
      shadow: cardShadow(),
    });

    // Number circle
    s.addShape(pres.shapes.OVAL, {
      x: PAD + 0.15, y: y + 0.15, w: 0.4, h: 0.4,
      fill: { color: C.navy },
    });
    s.addText(ask.num, {
      x: PAD + 0.15, y: y + 0.15, w: 0.4, h: 0.4,
      fontSize: 16, fontFace: FONT, color: C.white, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    // Title + urgency chip
    s.addText(ask.title, {
      x: PAD + 0.7, y: y + 0.1, w: BODY_W - 2.2, h: 0.3,
      fontSize: 13, fontFace: FONT, color: C.navy, bold: true, margin: 0,
    });

    // Urgency chip
    const chipW = ask.urgency.length * 0.07 + 0.3;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: PAD + BODY_W - chipW - 0.2, y: y + 0.14, w: chipW, h: 0.24,
      fill: { color: ask.urgencyBg },
      rectRadius: 0.12,
    });
    s.addText(ask.urgency, {
      x: PAD + BODY_W - chipW - 0.2, y: y + 0.14, w: chipW, h: 0.24,
      fontSize: 8, fontFace: FONT, color: ask.urgencyColor, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    // Description
    s.addText(ask.desc, {
      x: PAD + 0.7, y: y + 0.42, w: BODY_W - 1.0, h: 0.55,
      fontSize: MIN_FONT, fontFace: FONT, color: C.darkGray, margin: 0,
      valign: "top", lineSpacingMultiple: 1.2,
    });

    y += 1.15;
  });
}

// =============================================
// SLIDE 12: Appendix — Detailed Financials
// =============================================
{
  pageNum++;
  const s = pres.addSlide({ masterName: "MASTER" });
  addHeader(s, "Appendix: Detailed Financials");
  addFooter(s, pageNum);

  let y = BODY_TOP + 0.05;

  // P&L summary table (dense is OK for appendix)
  const plHeaders = ["($K)", "Q1 '26", "Q2 '26", "Q3 '26", "Q3 YoY", "FY '26E"];
  const plRows = [
    [{ text: "Revenue", bold: true }, "3,050", "3,175", "3,750", { text: "+48%", color: C.green }, "13,650"],
    ["  Subscriptions", "2,850", "2,975", "3,500", "+51%", "12,800"],
    ["  Services", "200", "200", "250", "+25%", "850"],
    [{ text: "COGS", bold: true }, "(670)", "(698)", "(825)", "+41%", "(3,003)"],
    [{ text: "Gross Profit", bold: true }, { text: "2,380", bold: true }, { text: "2,477", bold: true }, { text: "2,925", bold: true }, { text: "+51%", color: C.green, bold: true }, { text: "10,647", bold: true }],
    [{ text: "Gross Margin", bold: true }, "78.0%", "78.0%", "78.0%", "--", "78.0%"],
    ["", "", "", "", "", ""],
    [{ text: "OpEx", bold: true }, "", "", "", "", ""],
    ["  R&D", "(1,520)", "(1,560)", "(1,640)", "+32%", "(6,400)"],
    ["  Sales & Marketing", "(1,140)", "(1,170)", "(1,230)", "+28%", "(4,720)"],
    ["  G&A", "(380)", "(390)", "(410)", "+18%", "(1,580)"],
    [{ text: "Total OpEx", bold: true }, { text: "(3,040)", bold: true }, { text: "(3,120)", bold: true }, { text: "(3,280)", bold: true }, { text: "+29%", bold: true }, { text: "(12,700)", bold: true }],
    ["", "", "", "", "", ""],
    [{ text: "Net Income", bold: true }, { text: "(660)", color: C.red, bold: true }, { text: "(643)", color: C.red, bold: true }, { text: "(355)", color: C.amber, bold: true }, { text: "+64% impr", color: C.green, bold: true }, { text: "(2,053)", bold: true }],
  ];

  // Build fully styled rows
  const tHeaderRow = plHeaders.map(h => ({
    text: h,
    options: {
      fill: { color: C.navy }, color: C.white, bold: true,
      fontSize: 8, fontFace: FONT, align: h === "($K)" ? "left" : "right",
      valign: "middle", margin: [1, 3, 1, 3],
    },
  }));

  const tDataRows = plRows.map((row, ri) =>
    row.map((cell, ci) => {
      const isObj = typeof cell === "object" && cell !== null && cell.text !== undefined;
      const text = isObj ? cell.text : String(cell);
      const isBold = isObj ? cell.bold : false;
      const color = isObj && cell.color ? cell.color : C.charcoal;
      // Highlight gross profit and net income rows
      const isHighlight = isBold && (text.includes("2,925") || text.includes("(355)") || text.includes("10,647") || text.includes("(2,053)"));
      return {
        text,
        options: {
          fill: { color: isHighlight ? C.veryLight : (ri % 2 === 0 ? C.white : C.white) },
          color,
          bold: isBold || false,
          fontSize: 8,
          fontFace: FONT,
          align: ci === 0 ? "left" : "right",
          valign: "middle",
          margin: [1, 3, 1, 3],
        },
      };
    })
  );

  s.addTable([tHeaderRow, ...tDataRows], {
    x: PAD, y, w: BODY_W,
    colW: [BODY_W * 0.28, BODY_W * 0.14, BODY_W * 0.14, BODY_W * 0.14, BODY_W * 0.14, BODY_W * 0.16],
    border: { pt: 0.25, color: C.lightGray },
  });

  // Cash position note
  s.addText("Cash: $18.4M  |  Monthly Burn: $118K  |  Runway: 16.7 months  |  Last funding: Series B ($22M, Mar 2025)", {
    x: PAD, y: CONTENT_BOTTOM - 0.25, w: BODY_W, h: 0.2,
    fontSize: 8, fontFace: FONT, color: C.medGray, margin: 0, italic: true,
  });
}

// ────────────────────────────────────────────
// OUTPUT
// ────────────────────────────────────────────
const outDir = path.join(__dirname, "output");
if (!fs.existsSync(outDir)) {
  fs.mkdirSync(outDir, { recursive: true });
}

pres.writeFile({ fileName: path.join(outDir, "q3-board-review.pptx") })
  .then(() => {
    console.log("Generated: output/q3-board-review.pptx");
    console.log(`Slides: ${pageNum}`);
  })
  .catch((err) => {
    console.error("Error generating deck:", err);
    process.exit(1);
  });
