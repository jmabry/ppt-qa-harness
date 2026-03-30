const pptxgen = require("pptxgenjs");

// ============================================================
// Project Chimera — Monolith to Microservices
// VP Engineering Review
// ============================================================

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.author = "Engineering Architecture";
pres.title = "Project Chimera — Monolith Decomposition Plan";

// ============================================================
// COLOR PALETTE — Deep Ocean Technical
// ============================================================
const C = {
  bgDark:      "0D1B2A",   // near-black navy
  bgLight:     "F0F7FF",   // light blue-white
  primary:     "0E4D92",   // deep blue
  secondary:   "1C7293",   // teal blue
  accent:      "00B4D8",   // bright teal accent
  accentGold:  "F59E0B",   // gold for callouts
  textDark:    "E8F4FD",   // text on dark
  textLight:   "1A2942",   // text on light
  textMid:     "3D5A80",   // mid-tone text
  textMuted:   "64748B",   // muted text
  white:       "FFFFFF",
  offWhite:    "F8FBFF",
  cardBg:      "FFFFFF",
  border:      "CBD5E1",
  // Status colors
  critical:    "DC2626",   // red
  high:        "EA580C",   // orange
  medium:      "D97706",   // amber
  low:         "16A34A",   // green
  p0:          "DC2626",
  p1:          "D97706",
  p2:          "16A34A",
  // Rejection colors
  rejected:    "DC2626",
  insufficient:"D97706",
  selected:    "16A34A",
  green:       "16A34A",
  amber:       "D97706",
  red:         "DC2626",
};

const FONT = {
  head: "Arial Black",
  sub:  "Arial",
  body: "Calibri",
};

const SLIDE_W = 10;
const SLIDE_H = 5.625;
const MARGIN = 0.45;
const CONTENT_W = SLIDE_W - 2 * MARGIN;
const TOTAL_SLIDES = 6;

// ============================================================
// HELPERS
// ============================================================

function addDarkSlide() {
  const slide = pres.addSlide();
  slide.background = { color: C.bgDark };
  return slide;
}

function addLightSlide() {
  const slide = pres.addSlide();
  slide.background = { color: C.bgLight };
  return slide;
}

function addFooter(slide, num, dark = false) {
  const color = dark ? "4A7BA7" : C.textMuted;
  slide.addText(`${num} / ${TOTAL_SLIDES}`, {
    x: SLIDE_W - 1.2, y: SLIDE_H - 0.32,
    w: 0.9, h: 0.22,
    fontSize: 8, fontFace: FONT.body,
    color, align: "right", margin: 0,
  });
  slide.addText("PROJECT CHIMERA  |  VP ENGINEERING REVIEW  |  CONFIDENTIAL", {
    x: MARGIN, y: SLIDE_H - 0.32,
    w: 5, h: 0.22,
    fontSize: 7, fontFace: FONT.body,
    color, align: "left", margin: 0,
    charSpacing: 0.5,
  });
}

// Left-border accent bar motif on light slides
function addAccentBar(slide, x, y, h, opts = {}) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w: 0.07, h,
    fill: { color: opts.color || C.accent },
    line: { type: "none" },
  });
}

// Light slide header with left-border motif
function addLightHeader(slide, title, subtitle) {
  // Header background band
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: SLIDE_W, h: 0.78,
    fill: { color: C.primary },
    line: { type: "none" },
  });
  // Accent stripe
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.12, h: 0.78,
    fill: { color: C.accent },
    line: { type: "none" },
  });
  slide.addText(title, {
    x: 0.28, y: 0.09,
    w: CONTENT_W, h: 0.38,
    fontSize: 17, fontFace: FONT.head,
    color: C.white, bold: true,
    align: "left", margin: 0,
  });
  if (subtitle) {
    slide.addText(subtitle, {
      x: 0.28, y: 0.46,
      w: CONTENT_W, h: 0.26,
      fontSize: 9.5, fontFace: FONT.body,
      color: C.textDark, align: "left", margin: 0,
    });
  }
}

// Dark slide header
function addDarkHeader(slide, label, title) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.12, h: SLIDE_H,
    fill: { color: C.accent },
    line: { type: "none" },
  });
  if (label) {
    slide.addText(label.toUpperCase(), {
      x: 0.3, y: 0.38,
      w: CONTENT_W, h: 0.28,
      fontSize: 9, fontFace: FONT.sub,
      color: C.accent, bold: true,
      align: "left", margin: 0,
      charSpacing: 2,
    });
  }
  slide.addText(title, {
    x: 0.3, y: 0.62,
    w: CONTENT_W * 0.85, h: 0.7,
    fontSize: 26, fontFace: FONT.head,
    color: C.white, bold: true,
    align: "left", margin: 0,
  });
}

// ============================================================
// SLIDE 1 — TITLE (dark)
// ============================================================
{
  const slide = addDarkSlide();

  // Left accent bar (full height)
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.18, h: SLIDE_H,
    fill: { color: C.accent },
    line: { type: "none" },
  });

  // Background decorative — subtle grid lines
  for (let i = 1; i <= 4; i++) {
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.18, y: i * (SLIDE_H / 5),
      w: SLIDE_W, h: 0.008,
      fill: { color: "1A3550" },
      line: { type: "none" },
    });
  }

  // Main title
  slide.addText("Project Chimera", {
    x: 0.45, y: 0.62,
    w: 7.5, h: 1.1,
    fontSize: 54, fontFace: FONT.head,
    color: C.white, bold: true,
    align: "left", margin: 0,
  });

  // Accent line under title
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.45, y: 1.76,
    w: 3.8, h: 0.05,
    fill: { color: C.accent },
    line: { type: "none" },
  });

  // Subtitle
  slide.addText("Monolith Decomposition Plan — VP Engineering Review", {
    x: 0.45, y: 1.9,
    w: 7.5, h: 0.36,
    fontSize: 13.5, fontFace: FONT.sub,
    color: C.textDark, align: "left", margin: 0,
  });

  // 3 stat callouts
  const stats = [
    { val: "340K", label: "Lines of Code\n(monolith surface area)" },
    { val: "22",   label: "Engineers\nDelivering on single codebase" },
    { val: "$133K", label: "Per Quarter\nEstimated CI/deploy friction cost" },
  ];
  const cardW = 2.55;
  const cardH = 1.18;
  const cardY = 3.05;
  const cardGap = 0.2;
  const startX = 0.45;

  stats.forEach((s, i) => {
    const cx = startX + i * (cardW + cardGap);
    // Card bg
    slide.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: cardY, w: cardW, h: cardH,
      fill: { color: "112233" },
      line: { color: "1C4A6E", pt: 1 },
    });
    // Top accent stripe
    slide.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: cardY, w: cardW, h: 0.06,
      fill: { color: C.accent },
      line: { type: "none" },
    });
    // Value
    slide.addText(s.val, {
      x: cx, y: cardY + 0.08,
      w: cardW, h: 0.65,
      fontSize: 44, fontFace: FONT.head,
      color: C.accent, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
    // Label
    slide.addText(s.label, {
      x: cx + 0.1, y: cardY + 0.72,
      w: cardW - 0.2, h: 0.42,
      fontSize: 9, fontFace: FONT.body,
      color: C.textDark, align: "center",
      valign: "top", margin: 0,
      lineSpacingMultiple: 1.2,
    });
  });

  // Bottom decorative strip
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: SLIDE_H - 0.28, w: SLIDE_W, h: 0.28,
    fill: { color: "091524" },
    line: { type: "none" },
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: SLIDE_H - 0.28, w: 2.5, h: 0.28,
    fill: { color: C.secondary },
    line: { type: "none" },
  });
  slide.addText("Strangler Fig Pattern  ·  Phase 1–3  ·  M1–M6 2026", {
    x: 0.28, y: SLIDE_H - 0.28,
    w: 6, h: 0.28,
    fontSize: 8, fontFace: FONT.body,
    color: C.white, align: "left",
    valign: "middle", margin: 0,
  });

  addFooter(slide, 1, true);
}

// ============================================================
// SLIDE 2 — CURRENT STATE & PAIN POINTS (light)
// ============================================================
{
  const slide = addLightSlide();
  addLightHeader(slide, "Current State & Pain Points", "DORA metrics reveal systemic delivery constraints — Q1 2026 benchmark");
  addFooter(slide, 2, false);

  // Left panel: DORA metrics table
  const leftX = MARGIN;
  const tableY = 0.95;
  const tableW = 5.1;

  // Section label
  slide.addText("DORA METRICS — Q1 2026 BENCHMARK", {
    x: leftX + 0.1, y: tableY,
    w: tableW - 0.1, h: 0.22,
    fontSize: 8.5, fontFace: FONT.sub,
    color: C.primary, bold: true,
    align: "left", margin: 0, charSpacing: 0.8,
  });
  addAccentBar(slide, leftX, tableY, 0.22);

  // Table header
  const colHeaders = ["METRIC", "CURRENT", "ELITE P50", "STATUS"];
  const colWidths = [1.7, 1.05, 1.05, 0.88];
  const colXs = [];
  let cx = leftX + 0.1;
  colWidths.forEach(w => { colXs.push(cx); cx += w; });

  const headerRowY = tableY + 0.26;
  slide.addShape(pres.shapes.RECTANGLE, {
    x: leftX + 0.1, y: headerRowY, w: tableW - 0.1, h: 0.26,
    fill: { color: C.primary },
    line: { type: "none" },
  });
  colHeaders.forEach((h, i) => {
    slide.addText(h, {
      x: colXs[i], y: headerRowY,
      w: colWidths[i], h: 0.26,
      fontSize: 8, fontFace: FONT.sub,
      color: C.white, bold: true,
      align: i === 0 ? "left" : "center",
      valign: "middle", margin: [0, 3, 0, 3],
    });
  });

  // DORA rows
  const doraRows = [
    { metric: "Deploy Frequency",    current: "1.8/wk",   elite: "1/day",     status: "LAGGING",  sColor: C.high },
    { metric: "Lead Time for Change",current: "12.3 days", elite: "1–7 days",  status: "LAGGING",  sColor: C.high },
    { metric: "Change Failure Rate", current: "18.2%",    elite: "<15%",       status: "FAILING",  sColor: C.critical },
    { metric: "Mean Time to Recover",current: "4.2 hrs",  elite: "<1 hr",      status: "LAGGING",  sColor: C.high },
    { metric: "Build Time (CI)",     current: "47 min",   elite: "<15 min",    status: "FAILING",  sColor: C.critical },
  ];

  doraRows.forEach((row, ri) => {
    const rowY = headerRowY + 0.26 + ri * 0.32;
    const rowBg = ri % 2 === 0 ? C.white : "EDF4FF";
    slide.addShape(pres.shapes.RECTANGLE, {
      x: leftX + 0.1, y: rowY, w: tableW - 0.1, h: 0.32,
      fill: { color: rowBg },
      line: { color: C.border, pt: 0.5 },
    });
    // Metric name
    slide.addText(row.metric, {
      x: colXs[0], y: rowY,
      w: colWidths[0], h: 0.32,
      fontSize: 9, fontFace: FONT.body,
      color: C.textLight, align: "left",
      valign: "middle", margin: [0, 4, 0, 4],
    });
    // Current
    slide.addText(row.current, {
      x: colXs[1], y: rowY,
      w: colWidths[1], h: 0.32,
      fontSize: 9.5, fontFace: FONT.sub,
      color: C.textLight, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
    // Elite
    slide.addText(row.elite, {
      x: colXs[2], y: rowY,
      w: colWidths[2], h: 0.32,
      fontSize: 9, fontFace: FONT.body,
      color: C.textMuted, align: "center",
      valign: "middle", margin: 0,
    });
    // Status badge
    slide.addShape(pres.shapes.RECTANGLE, {
      x: colXs[3] + 0.05, y: rowY + 0.07,
      w: colWidths[3] - 0.1, h: 0.18,
      fill: { color: row.sColor },
      line: { type: "none" },
    });
    slide.addText(row.status, {
      x: colXs[3] + 0.05, y: rowY + 0.07,
      w: colWidths[3] - 0.1, h: 0.18,
      fontSize: 7.5, fontFace: FONT.sub,
      color: C.white, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
  });

  // Insight callout below table
  const calloutY = headerRowY + 0.26 + doraRows.length * 0.32 + 0.12;
  addAccentBar(slide, leftX, calloutY, 0.52, { color: C.accentGold });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: leftX + 0.1, y: calloutY, w: tableW - 0.1, h: 0.52,
    fill: { color: "FFFBEB" },
    line: { color: C.accentGold, pt: 0.5 },
  });
  slide.addText("6.8% CI flakiness = 'retry and pray' culture\nTeams spend 9 min avg per failed pipeline waiting for reruns", {
    x: leftX + 0.22, y: calloutY + 0.04,
    w: tableW - 0.32, h: 0.44,
    fontSize: 9, fontFace: FONT.body,
    color: "92400E", align: "left",
    valign: "middle", margin: 0,
    lineSpacingMultiple: 1.3,
    italic: false,
  });

  // Right panel: Cost breakdown
  const rightX = 5.7;
  const rightW = SLIDE_W - rightX - MARGIN;

  slide.addText("COST IMPACT ANALYSIS", {
    x: rightX + 0.1, y: tableY,
    w: rightW, h: 0.22,
    fontSize: 8.5, fontFace: FONT.sub,
    color: C.primary, bold: true,
    align: "left", margin: 0, charSpacing: 0.8,
  });
  addAccentBar(slide, rightX, tableY, 0.22);

  // Big cost number
  slide.addShape(pres.shapes.RECTANGLE, {
    x: rightX + 0.1, y: tableY + 0.26,
    w: rightW, h: 1.0,
    fill: { color: C.primary },
    line: { type: "none" },
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: rightX + 0.1, y: tableY + 0.26,
    w: 0.07, h: 1.0,
    fill: { color: C.accent },
    line: { type: "none" },
  });
  slide.addText("$133K", {
    x: rightX + 0.2, y: tableY + 0.3,
    w: rightW - 0.12, h: 0.56,
    fontSize: 46, fontFace: FONT.head,
    color: C.accent, bold: true,
    align: "center", valign: "middle", margin: 0,
  });
  slide.addText("Quarterly Engineering Waste", {
    x: rightX + 0.2, y: tableY + 0.84,
    w: rightW - 0.12, h: 0.28,
    fontSize: 9.5, fontFace: FONT.sub,
    color: C.textDark, align: "center",
    valign: "middle", margin: 0,
  });

  // Cost breakdown lines
  const breakdownItems = [
    ["1,400 eng-hrs / qtr", ""],
    ["× $95 / hr (blended rate)", ""],
    ["= $133,000 quarterly", "friction cost"],
    ["= $532,000 annualized", "opportunity cost"],
  ];
  breakdownItems.forEach((item, ii) => {
    const itemY = tableY + 1.36 + ii * 0.3;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: rightX + 0.1, y: itemY,
      w: rightW, h: 0.29,
      fill: { color: ii % 2 === 0 ? C.white : "EDF4FF" },
      line: { color: C.border, pt: 0.5 },
    });
    slide.addText(item[0], {
      x: rightX + 0.2, y: itemY,
      w: rightW - 0.3, h: 0.29,
      fontSize: ii >= 2 ? 9.5 : 9,
      fontFace: ii >= 2 ? FONT.sub : FONT.body,
      color: ii >= 2 ? C.primary : C.textLight,
      bold: ii >= 2,
      align: "left", valign: "middle", margin: 0,
    });
    if (item[1]) {
      slide.addText(item[1], {
        x: rightX + 0.1, y: itemY,
        w: rightW - 0.1, h: 0.29,
        fontSize: 8, fontFace: FONT.body,
        color: C.textMuted,
        align: "right", valign: "middle", margin: [0, 5, 0, 0],
      });
    }
  });

  // Sources note
  const srcY = tableY + 1.36 + breakdownItems.length * 0.3 + 0.08;
  slide.addText("Source: Engineering time-tracking Q4 2025, 22-engineer base", {
    x: rightX + 0.1, y: srcY,
    w: rightW, h: 0.2,
    fontSize: 7.5, fontFace: FONT.body,
    color: C.textMuted, align: "left", margin: 0,
    italic: true,
  });
}

// ============================================================
// SLIDE 3 — TARGET ARCHITECTURE & SERVICE BOUNDARIES (light)
// ============================================================
{
  const slide = addLightSlide();
  addLightHeader(slide, "Target Architecture & Service Boundaries", "Six bounded-context services — Strangler Fig extraction sequence");
  addFooter(slide, 3, false);

  const tableY = 0.92;
  const tableX = MARGIN;
  const tableW = CONTENT_W;

  // Table section label
  slide.addText("SERVICE BOUNDARY DEFINITIONS", {
    x: tableX + 0.1, y: tableY,
    w: tableW - 0.1, h: 0.22,
    fontSize: 8.5, fontFace: FONT.sub,
    color: C.primary, bold: true, align: "left", margin: 0, charSpacing: 0.8,
  });
  addAccentBar(slide, tableX, tableY, 0.22);

  // Header row
  const svcCols = ["SERVICE", "SQUAD", "DATABASE", "API STYLE", "PRIORITY"];
  const svcWidths = [1.65, 1.5, 1.6, 1.45, 0.9];
  const svcXs = [];
  let sxBase = tableX + 0.1;
  svcWidths.forEach(w => { svcXs.push(sxBase); sxBase += w; });

  const hdrY = tableY + 0.26;
  slide.addShape(pres.shapes.RECTANGLE, {
    x: tableX + 0.1, y: hdrY, w: tableW - 0.1, h: 0.26,
    fill: { color: C.primary }, line: { type: "none" },
  });
  svcCols.forEach((h, i) => {
    slide.addText(h, {
      x: svcXs[i], y: hdrY,
      w: svcWidths[i], h: 0.26,
      fontSize: 8, fontFace: FONT.sub,
      color: C.white, bold: true,
      align: i === 0 ? "left" : "center",
      valign: "middle", margin: [0, 3, 0, 3],
    });
  });

  // Service rows
  const services = [
    { name: "User Service",      squad: "Identity",      db: "PostgreSQL",   api: "REST + gRPC",   priority: "P0", pColor: C.p0 },
    { name: "Order Service",     squad: "Commerce",      db: "PostgreSQL",   api: "REST",          priority: "P1", pColor: C.p1 },
    { name: "Payment Service",   squad: "Commerce",      db: "PostgreSQL",   api: "REST + Events", priority: "P0", pColor: C.p0 },
    { name: "Search Service",    squad: "Discovery",     db: "Elasticsearch",api: "REST",          priority: "P1", pColor: C.p1 },
    { name: "Notification Svc",  squad: "Platform",      db: "Redis + PG",   api: "Events",        priority: "P2", pColor: C.p2 },
    { name: "Analytics Service", squad: "Data",          db: "ClickHouse",   api: "gRPC",          priority: "P2", pColor: C.p2 },
  ];

  services.forEach((svc, ri) => {
    const rowY = hdrY + 0.26 + ri * 0.3;
    const rowBg = ri % 2 === 0 ? C.white : "EDF4FF";
    slide.addShape(pres.shapes.RECTANGLE, {
      x: tableX + 0.1, y: rowY, w: tableW - 0.1, h: 0.3,
      fill: { color: rowBg }, line: { color: C.border, pt: 0.5 },
    });
    // Service name
    slide.addText(svc.name, {
      x: svcXs[0], y: rowY, w: svcWidths[0], h: 0.3,
      fontSize: 9.5, fontFace: FONT.sub, color: C.primary,
      bold: true, align: "left", valign: "middle", margin: [0, 4, 0, 4],
    });
    [svc.squad, svc.db, svc.api].forEach((val, ci) => {
      slide.addText(val, {
        x: svcXs[ci + 1], y: rowY, w: svcWidths[ci + 1], h: 0.3,
        fontSize: 9, fontFace: FONT.body, color: C.textLight,
        align: "center", valign: "middle", margin: 0,
      });
    });
    // Priority badge
    slide.addShape(pres.shapes.RECTANGLE, {
      x: svcXs[4] + 0.12, y: rowY + 0.07,
      w: svcWidths[4] - 0.24, h: 0.18,
      fill: { color: svc.pColor }, line: { type: "none" },
    });
    slide.addText(svc.priority, {
      x: svcXs[4] + 0.12, y: rowY + 0.07,
      w: svcWidths[4] - 0.24, h: 0.18,
      fontSize: 8.5, fontFace: FONT.sub,
      color: C.white, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
  });

  // Architecture decisions callout — right of table or below
  const adY = hdrY + 0.26 + services.length * 0.3 + 0.14;
  addAccentBar(slide, tableX, adY, 0.72);
  slide.addShape(pres.shapes.RECTANGLE, {
    x: tableX + 0.1, y: adY, w: tableW - 0.1, h: 0.72,
    fill: { color: "EFF6FF" }, line: { color: C.accent, pt: 0.75 },
  });
  slide.addText("KEY ARCHITECTURAL DECISIONS", {
    x: tableX + 0.22, y: adY + 0.06,
    w: tableW - 0.32, h: 0.2,
    fontSize: 8.5, fontFace: FONT.sub,
    color: C.primary, bold: true, align: "left", margin: 0, charSpacing: 0.5,
  });
  const decisions = [
    "API Gateway: Kong (rate limiting, auth offload, canary routing)",
    "Service Mesh: Istio (mTLS, circuit breakers, traffic shaping, observability)",
    "Event Bus: Kafka 3-broker cluster (ordered delivery, at-least-once, compacted topics per service)",
    "Data Isolation: Each service owns its schema — no cross-service DB joins; CDC via Debezium for sync",
  ];
  decisions.forEach((d, di) => {
    slide.addText(`·  ${d}`, {
      x: tableX + 0.22, y: adY + 0.26 + di * 0.115,
      w: tableW - 0.32, h: 0.13,
      fontSize: 9, fontFace: FONT.body,
      color: C.textLight, align: "left", margin: 0,
      lineSpacingMultiple: 1.1,
    });
  });
}

// ============================================================
// SLIDE 4 — WHY STRANGLER FIG — ADR (dark)
// ============================================================
{
  const slide = addDarkSlide();

  // Left accent bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.12, h: SLIDE_H,
    fill: { color: C.accent }, line: { type: "none" },
  });

  // Header
  slide.addText("ARCHITECTURE DECISION RECORD", {
    x: 0.3, y: 0.28,
    w: CONTENT_W, h: 0.24,
    fontSize: 9, fontFace: FONT.sub,
    color: C.accent, bold: true,
    align: "left", margin: 0, charSpacing: 2,
  });
  slide.addText("Why Strangler Fig?", {
    x: 0.3, y: 0.5,
    w: CONTENT_W, h: 0.56,
    fontSize: 24, fontFace: FONT.head,
    color: C.white, bold: true,
    align: "left", margin: 0,
  });

  // Horizontal divider
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.3, y: 1.1, w: CONTENT_W, h: 0.03,
    fill: { color: "1C3A52" }, line: { type: "none" },
  });

  // 3 evaluation columns
  const cols = [
    {
      badge: "REJECTED",
      badgeColor: C.rejected,
      icon: "✗",
      title: "Full Rewrite",
      details: [
        "14–18 months estimated",
        "Dedicated team of 8",
        "Feature freeze required",
        "Zero revenue delivery during migration",
      ],
      evidence: '"Netscape, Basecamp: 2–3× longer than estimated. Risk of losing product-market fit mid-rewrite."',
      evidenceColor: "FCA5A5",
    },
    {
      badge: "INSUFFICIENT",
      badgeColor: C.insufficient,
      icon: "−",
      title: "Modularize Monolith",
      details: [
        "Enforced module boundaries only",
        "Atlas Modular Q4 2025: -30% conflicts",
        "Deploy frequency: unchanged",
        "DB coupling problem persists",
      ],
      evidence: '"Boundary violations drop but deployment bottleneck remains. One bad migration still blocks all teams."',
      evidenceColor: "FDE68A",
    },
    {
      badge: "SELECTED",
      badgeColor: C.selected,
      icon: "✓",
      title: "Strangler Fig",
      details: [
        "Incremental, bounded-risk extraction",
        "Revenue delivery continues throughout",
        "Proven: User Service Q1 2026",
        "3 weeks, zero incidents",
      ],
      evidence: '"Validated internal proof-of-concept. Monolith stays runnable as safety net throughout migration."',
      evidenceColor: "6EE7B7",
    },
  ];

  const colW = 2.95;
  const colGap = 0.12;
  const colStartX = 0.3;
  const colBodyY = 1.2;

  cols.forEach((col, ci) => {
    const colX = colStartX + ci * (colW + colGap);

    // Column background
    slide.addShape(pres.shapes.RECTANGLE, {
      x: colX, y: colBodyY, w: colW, h: 3.55,
      fill: { color: "0A1624" },
      line: { color: col.badgeColor, pt: 1 },
    });

    // Top badge
    slide.addShape(pres.shapes.RECTANGLE, {
      x: colX, y: colBodyY, w: colW, h: 0.32,
      fill: { color: col.badgeColor }, line: { type: "none" },
    });
    slide.addText(`${col.icon}  ${col.badge}`, {
      x: colX, y: colBodyY, w: colW, h: 0.32,
      fontSize: 10.5, fontFace: FONT.sub,
      color: C.white, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    // Option title
    slide.addText(col.title, {
      x: colX + 0.12, y: colBodyY + 0.38,
      w: colW - 0.24, h: 0.34,
      fontSize: 13.5, fontFace: FONT.head,
      color: C.white, bold: true,
      align: "left", valign: "middle", margin: 0,
    });

    // Bullet details
    col.details.forEach((d, di) => {
      slide.addText(`·  ${d}`, {
        x: colX + 0.12, y: colBodyY + 0.78 + di * 0.26,
        w: colW - 0.24, h: 0.26,
        fontSize: 9.5, fontFace: FONT.body,
        color: C.textDark, align: "left",
        valign: "middle", margin: 0,
        lineSpacingMultiple: 1.1,
      });
    });

    // Evidence quote
    const quoteY = colBodyY + 0.78 + col.details.length * 0.26 + 0.1;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: colX + 0.1, y: quoteY,
      w: colW - 0.2, h: 0.75,
      fill: { color: "0F2030" },
      line: { color: "1E4060", pt: 0.5 },
    });
    slide.addText(col.evidence, {
      x: colX + 0.18, y: quoteY + 0.06,
      w: colW - 0.36, h: 0.63,
      fontSize: 8.5, fontFace: FONT.body,
      color: col.evidenceColor, align: "left",
      valign: "middle", margin: 0,
      italic: true, lineSpacingMultiple: 1.2,
    });
  });

  // Bottom notes
  const noteY = SLIDE_H - 0.72;
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.3, y: noteY, w: CONTENT_W, h: 0.04,
    fill: { color: "1C3A52" }, line: { type: "none" },
  });
  slide.addText("Scope: Reporting + Admin modules remain in monolith. Goal is deployment velocity — not zero-monolith.", {
    x: 0.3, y: noteY + 0.08,
    w: CONTENT_W, h: 0.2,
    fontSize: 9, fontFace: FONT.body,
    color: C.textDark, align: "left", margin: 0,
  });
  slide.addText("Timeline confidence: 65% Phase 1–2  ·  40% Phase 3 (ClickHouse learning curve factored in)", {
    x: 0.3, y: noteY + 0.28,
    w: CONTENT_W, h: 0.2,
    fontSize: 8.5, fontFace: FONT.body,
    color: "7AA5C2", align: "left", margin: 0, italic: true,
  });

  addFooter(slide, 4, true);
}

// ============================================================
// SLIDE 5 — MIGRATION ROADMAP (light)
// ============================================================
{
  const slide = addLightSlide();
  addLightHeader(slide, "Migration Roadmap", "Strangler Fig extraction — M1–M6 2026 · Three-phase delivery");
  addFooter(slide, 5, false);

  // Phase bar
  const phases = [
    { label: "Phase 1: Foundation", sub: "Months 1–2", color: C.primary },
    { label: "Phase 2: Commerce Core", sub: "Months 3–4", color: C.secondary },
    { label: "Phase 3: Platform", sub: "Months 5–6", color: "21295C" },
  ];
  const phaseBarY = 0.88;
  const phaseW = (CONTENT_W - 0.1) / 3;
  phases.forEach((ph, pi) => {
    const phX = MARGIN + pi * (phaseW + 0.04);
    slide.addShape(pres.shapes.RECTANGLE, {
      x: phX, y: phaseBarY, w: phaseW, h: 0.44,
      fill: { color: ph.color }, line: { type: "none" },
    });
    if (pi === 0) {
      slide.addShape(pres.shapes.RECTANGLE, {
        x: phX, y: phaseBarY, w: 0.07, h: 0.44,
        fill: { color: C.accent }, line: { type: "none" },
      });
    }
    slide.addText(ph.label, {
      x: phX + (pi === 0 ? 0.1 : 0.05), y: phaseBarY + 0.04,
      w: phaseW - 0.1, h: 0.22,
      fontSize: 10, fontFace: FONT.sub,
      color: C.white, bold: true, align: "center", margin: 0,
    });
    slide.addText(ph.sub, {
      x: phX + 0.05, y: phaseBarY + 0.24,
      w: phaseW - 0.1, h: 0.16,
      fontSize: 8.5, fontFace: FONT.body,
      color: "B8D4F0", align: "center", margin: 0,
    });
  });

  // Phase task tables
  const phaseData = [
    {
      tasks: [
        ["API Gateway (Kong)", "Setup", "LOW"],
        ["OTel Instrumentation", "Observability", "LOW"],
        ["User Service Extract", "Core P0", "MEDIUM"],
        ["Kafka Cluster Provision", "Infra", "MEDIUM"],
        ["DB Migration Tooling", "Flyway+CDC", "LOW"],
      ],
    },
    {
      tasks: [
        ["Order Service Extract", "Commerce P1", "HIGH"],
        ["Payment Service Extract", "Commerce P0", "CRITICAL"],
        ["Checkout Saga Pattern", "Resilience", "HIGH"],
        ["Search Extraction", "Discovery P1", "MEDIUM"],
      ],
    },
    {
      tasks: [
        ["Notification Service", "Platform P2", "MEDIUM"],
        ["Analytics (ClickHouse)", "Data P2", "HIGH"],
        ["Monolith Decommission", "Cleanup", "MEDIUM"],
        ["Chaos Engineering", "Resilience", "LOW"],
      ],
    },
  ];

  const riskColorMap = {
    CRITICAL: C.critical,
    HIGH:     C.high,
    MEDIUM:   C.medium,
    LOW:      C.low,
  };

  const taskTableY = phaseBarY + 0.52;
  phaseData.forEach((phase, pi) => {
    const tX = MARGIN + pi * (phaseW + 0.04);
    const tW = phaseW;

    // Sub-header
    const subHdrCols = ["TASK", "TYPE", "RISK"];
    const subColW = [1.55, 0.72, 0.58];
    const subColXBase = tX;
    const subColXs = [];
    let xc = subColXBase;
    subColW.forEach(w => { subColXs.push(xc); xc += w; });

    slide.addShape(pres.shapes.RECTANGLE, {
      x: tX, y: taskTableY, w: tW, h: 0.24,
      fill: { color: phases[pi].color }, line: { type: "none" },
    });
    subHdrCols.forEach((h, hi) => {
      slide.addText(h, {
        x: subColXs[hi], y: taskTableY, w: subColW[hi], h: 0.24,
        fontSize: 7.5, fontFace: FONT.sub,
        color: C.white, bold: true,
        align: hi === 0 ? "left" : "center",
        valign: "middle", margin: [0, 3, 0, 3],
      });
    });

    phase.tasks.forEach((task, ti) => {
      const rowY = taskTableY + 0.24 + ti * 0.3;
      const rowBg = ti % 2 === 0 ? C.white : "EDF4FF";
      slide.addShape(pres.shapes.RECTANGLE, {
        x: tX, y: rowY, w: tW, h: 0.3,
        fill: { color: rowBg }, line: { color: C.border, pt: 0.5 },
      });
      // Task name
      slide.addText(task[0], {
        x: subColXs[0], y: rowY, w: subColW[0], h: 0.3,
        fontSize: 8.5, fontFace: FONT.body,
        color: C.textLight, align: "left",
        valign: "middle", margin: [0, 3, 0, 3],
      });
      // Type
      slide.addText(task[1], {
        x: subColXs[1], y: rowY, w: subColW[1], h: 0.3,
        fontSize: 8, fontFace: FONT.body,
        color: C.textMuted, align: "center",
        valign: "middle", margin: 0,
      });
      // Risk badge
      const rColor = riskColorMap[task[2]] || C.low;
      slide.addShape(pres.shapes.RECTANGLE, {
        x: subColXs[2] + 0.03, y: rowY + 0.07,
        w: subColW[2] - 0.06, h: 0.16,
        fill: { color: rColor }, line: { type: "none" },
      });
      slide.addText(task[2], {
        x: subColXs[2] + 0.03, y: rowY + 0.07,
        w: subColW[2] - 0.06, h: 0.16,
        fontSize: 6.5, fontFace: FONT.sub,
        color: C.white, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
    });
  });

  // Progress milestone notes
  const milesY = taskTableY + 0.24 + 5 * 0.3 + 0.1;
  addAccentBar(slide, MARGIN, milesY, 0.32);
  slide.addShape(pres.shapes.RECTANGLE, {
    x: MARGIN + 0.1, y: milesY, w: CONTENT_W - 0.1, h: 0.32,
    fill: { color: "EFF6FF" }, line: { color: C.accent, pt: 0.5 },
  });
  slide.addText("Gate criteria: Each phase requires >95% test coverage on extracted service, P95 latency <150ms, and zero SEV-1 incidents for 5 business days before proceeding.", {
    x: MARGIN + 0.22, y: milesY + 0.04,
    w: CONTENT_W - 0.32, h: 0.24,
    fontSize: 8.5, fontFace: FONT.body,
    color: C.primary, align: "left",
    valign: "middle", margin: 0, lineSpacingMultiple: 1.2,
  });
}

// ============================================================
// SLIDE 6 — RISK MATRIX + SUCCESS METRICS (light)
// ============================================================
{
  const slide = addLightSlide();
  addLightHeader(slide, "Risk Matrix & Success Metrics", "Six identified risks · Eight before/after delivery targets at 12 months");
  addFooter(slide, 6, false);

  const contentY = 0.92;
  const leftW = 4.6;
  const rightW = CONTENT_W - leftW - 0.15;
  const rightX = MARGIN + leftW + 0.15;

  // ---- LEFT: Risk table ----
  slide.addText("RISK REGISTER", {
    x: MARGIN + 0.1, y: contentY,
    w: leftW, h: 0.22,
    fontSize: 8.5, fontFace: FONT.sub,
    color: C.primary, bold: true, align: "left", margin: 0, charSpacing: 0.8,
  });
  addAccentBar(slide, MARGIN, contentY, 0.22);

  const riskCols = ["RISK", "SEV", "MITIGATION"];
  const riskColWidths = [1.52, 0.64, 2.28];
  const riskColXs = [];
  let rxc = MARGIN + 0.1;
  riskColWidths.forEach(w => { riskColXs.push(rxc); rxc += w; });

  const riskHdrY = contentY + 0.26;
  slide.addShape(pres.shapes.RECTANGLE, {
    x: MARGIN + 0.1, y: riskHdrY, w: leftW - 0.1, h: 0.26,
    fill: { color: C.primary }, line: { type: "none" },
  });
  riskCols.forEach((h, i) => {
    slide.addText(h, {
      x: riskColXs[i], y: riskHdrY, w: riskColWidths[i], h: 0.26,
      fontSize: 8, fontFace: FONT.sub,
      color: C.white, bold: true,
      align: i === 0 ? "left" : "center",
      valign: "middle", margin: [0, 3, 0, 3],
    });
  });

  const risks = [
    { risk: "Saga transaction failures",  sev: "CRITICAL", sevColor: C.critical, mit: "Compensation logic + monolith fallback flag" },
    { risk: "Kafka instability",          sev: "HIGH",     sevColor: C.high,     mit: "Confluent Cloud managed + specialist hire Q1" },
    { risk: "Data inconsistency",         sev: "HIGH",     sevColor: C.high,     mit: "CDC Debezium + nightly reconciliation jobs" },
    { risk: "Service latency increase",   sev: "MEDIUM",   sevColor: C.medium,   mit: "200ms budget enforcement, Istio circuit breakers" },
    { risk: "Team cognitive load",        sev: "MEDIUM",   sevColor: C.medium,   mit: "Phased tooling adoption, inner-source docs" },
    { risk: "PCI compliance scope creep", sev: "MEDIUM",   sevColor: C.medium,   mit: "Security review M2; Stripe hosted checkout" },
  ];

  risks.forEach((r, ri) => {
    const rowY = riskHdrY + 0.26 + ri * 0.34;
    const rowBg = ri % 2 === 0 ? C.white : "EDF4FF";
    slide.addShape(pres.shapes.RECTANGLE, {
      x: MARGIN + 0.1, y: rowY, w: leftW - 0.1, h: 0.34,
      fill: { color: rowBg }, line: { color: C.border, pt: 0.5 },
    });
    slide.addText(r.risk, {
      x: riskColXs[0], y: rowY, w: riskColWidths[0], h: 0.34,
      fontSize: 9, fontFace: FONT.body,
      color: C.textLight, align: "left",
      valign: "middle", margin: [0, 4, 0, 4],
    });
    // Severity badge
    slide.addShape(pres.shapes.RECTANGLE, {
      x: riskColXs[1] + 0.03, y: rowY + 0.08,
      w: riskColWidths[1] - 0.06, h: 0.18,
      fill: { color: r.sevColor }, line: { type: "none" },
    });
    slide.addText(r.sev, {
      x: riskColXs[1] + 0.03, y: rowY + 0.08,
      w: riskColWidths[1] - 0.06, h: 0.18,
      fontSize: 7, fontFace: FONT.sub,
      color: C.white, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
    slide.addText(r.mit, {
      x: riskColXs[2], y: rowY, w: riskColWidths[2], h: 0.34,
      fontSize: 8.5, fontFace: FONT.body,
      color: C.textMuted, align: "left",
      valign: "middle", margin: [0, 3, 0, 3],
      lineSpacingMultiple: 1.1,
    });
  });

  // ---- RIGHT: Before/After table ----
  slide.addText("12-MONTH TARGETS", {
    x: rightX + 0.1, y: contentY,
    w: rightW, h: 0.22,
    fontSize: 8.5, fontFace: FONT.sub,
    color: C.primary, bold: true, align: "left", margin: 0, charSpacing: 0.8,
  });
  addAccentBar(slide, rightX, contentY, 0.22);

  const metCols = ["METRIC", "NOW", "TARGET"];
  const metColWidths = [1.52, 0.7, 0.94];
  const metColXs = [];
  let mxc = rightX + 0.1;
  metColWidths.forEach(w => { metColXs.push(mxc); mxc += w; });

  const metHdrY = contentY + 0.26;
  slide.addShape(pres.shapes.RECTANGLE, {
    x: rightX + 0.1, y: metHdrY, w: rightW - 0.1, h: 0.26,
    fill: { color: C.primary }, line: { type: "none" },
  });
  metCols.forEach((h, i) => {
    slide.addText(h, {
      x: metColXs[i], y: metHdrY, w: metColWidths[i], h: 0.26,
      fontSize: 8, fontFace: FONT.sub,
      color: C.white, bold: true,
      align: i === 0 ? "left" : "center",
      valign: "middle", margin: [0, 3, 0, 3],
    });
  });

  const metrics = [
    { metric: "Deploy Frequency",    now: "1.8/wk",  target: "1+/day" },
    { metric: "Lead Time",           now: "12.3d",   target: "<2d" },
    { metric: "Change Failure Rate", now: "18.2%",   target: "<8%" },
    { metric: "MTTR",                now: "4.2 hrs", target: "<30 min" },
    { metric: "P99 API Latency",     now: "820ms",   target: "<200ms" },
    { metric: "Build Time (CI)",     now: "47 min",  target: "<8 min" },
    { metric: "CI Flakiness",        now: "6.8%",    target: "<1%" },
    { metric: "Dev Satisfaction",    now: "5.8/10",  target: "8.0/10" },
  ];

  metrics.forEach((m, mi) => {
    const rowY = metHdrY + 0.26 + mi * 0.29;
    const rowBg = mi % 2 === 0 ? C.white : "EDF4FF";
    slide.addShape(pres.shapes.RECTANGLE, {
      x: rightX + 0.1, y: rowY, w: rightW - 0.1, h: 0.29,
      fill: { color: rowBg }, line: { color: C.border, pt: 0.5 },
    });
    slide.addText(m.metric, {
      x: metColXs[0], y: rowY, w: metColWidths[0], h: 0.29,
      fontSize: 9, fontFace: FONT.body,
      color: C.textLight, align: "left",
      valign: "middle", margin: [0, 4, 0, 4],
    });
    slide.addText(m.now, {
      x: metColXs[1], y: rowY, w: metColWidths[1], h: 0.29,
      fontSize: 9, fontFace: FONT.sub,
      color: C.high, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
    slide.addText(m.target, {
      x: metColXs[2], y: rowY, w: metColWidths[2], h: 0.29,
      fontSize: 9, fontFace: FONT.sub,
      color: C.green, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
  });
}

// ============================================================
// SAVE
// ============================================================
pres.writeFile({ fileName: "outputs/software-anthropic.pptx" })
  .then(() => console.log("✓ Saved: outputs/software-anthropic.pptx"))
  .catch(err => { console.error("Error:", err); process.exit(1); });
