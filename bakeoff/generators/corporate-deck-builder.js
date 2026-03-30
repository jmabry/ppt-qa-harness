// ============================================================
// United Airlines Holdings — FY2025 Results & 2026 Outlook
// deck-builder convention: constants → helpers → slides
// ============================================================
"use strict";

const pptxgen = require("pptxgenjs");
const path = require("path");

// ============================================================
// LAYER 1: CONSTANTS
// ============================================================
const W = 10;
const H = 5.625;
const PAD = 0.5;
const TITLE_H = 0.5;
const FOOTER_Y = 5.35;
const CONTENT_BOTTOM = FOOTER_Y - 0.12;   // 5.23
const BODY_TOP = PAD + TITLE_H + 0.1;      // 1.1 after header on light slides
const MIN_FONT = 9;

const COL = {
  navy:      "1A2744",
  gold:      "D4A843",
  white:     "FFFFFF",
  bgLight:   "F5F6FA",
  bgCard:    "FFFFFF",
  navyMid:   "253563",
  navyLight: "3A4F7A",
  gray:      "8A96AE",
  grayDark:  "4A5568",
  grayLight: "CBD5E0",
  ice:       "D6E4F7",
  green:     "2D7D46",
  amber:     "B7791F",
  red:       "C0392B",
  stripe:    "EBF0FA",
};

const FONT_HEAD = "Calibri";
const FONT_BODY = "Calibri";

// ============================================================
// LAYER 2: HELPERS  (Y-position chaining: helpers return next Y)
// ============================================================
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.author = "United Airlines Holdings, Inc.";
pres.title  = "UAL FY2025 Results & 2026 Outlook";

// --- addHeader: returns next Y (BODY_TOP equiv after header bar)
function addHeader(slide, title, dark = false) {
  if (dark) {
    // Dark slides: just a bottom rule line + title text, no filled bar
    slide.addShape(pres.ShapeType.rect, {
      x: 0, y: 0, w: W, h: 0.06,
      fill: { color: COL.gold },
    });
    slide.addText(title, {
      x: PAD, y: 0.12, w: W - PAD * 2, h: TITLE_H,
      fontSize: 18, fontFace: FONT_HEAD,
      color: COL.white, bold: true, align: "left",
    });
    return 0.12 + TITLE_H + 0.12;  // ~0.74
  } else {
    // Light slides: navy bar
    slide.addShape(pres.ShapeType.rect, {
      x: 0, y: 0, w: W, h: TITLE_H + 0.2,
      fill: { color: COL.navy },
    });
    slide.addText(title, {
      x: PAD, y: 0.1, w: W - PAD * 2, h: TITLE_H,
      fontSize: 15, fontFace: FONT_HEAD,
      color: COL.white, bold: true, align: "left",
    });
    return TITLE_H + 0.2 + 0.12;  // ~0.82
  }
}

// --- addFooter: adds footer bar, no return value needed
function addFooter(slide, dark = false) {
  const fg = dark ? COL.ice : COL.grayDark;
  slide.addShape(pres.ShapeType.rect, {
    x: 0, y: FOOTER_Y, w: W, h: H - FOOTER_Y,
    fill: { color: dark ? COL.navyMid : COL.grayLight, transparency: dark ? 60 : 40 },
  });
  slide.addText("United Airlines Holdings  |  NASDAQ: UAL  |  FY2025 Investor Presentation", {
    x: PAD, y: FOOTER_Y + 0.02, w: W - PAD * 2, h: 0.2,
    fontSize: 7, fontFace: FONT_BODY, color: fg, align: "left",
  });
}

// --- addSectionLabel: label above a section, returns y + 0.24
function addSectionLabel(slide, text, y, dark = false) {
  slide.addText(text.toUpperCase(), {
    x: PAD, y, w: W - PAD * 2, h: 0.2,
    fontSize: MIN_FONT, fontFace: FONT_HEAD, bold: true,
    color: dark ? COL.gold : COL.navy,
    align: "left",
  });
  return y + 0.24;
}

// --- addBullets: renders bullet text, returns y + consumed height
function addBullets(slide, items, y, opts = {}) {
  const h = opts.h || items.length * 0.28;
  const bullets = items.map((txt) => ({ text: txt, options: { bullet: { code: "2022" } } }));
  slide.addText(bullets, {
    x: opts.x || PAD, y, w: opts.w || W - PAD * 2, h,
    fontSize: opts.fontSize || MIN_FONT, fontFace: FONT_BODY,
    color: opts.color || COL.grayDark,
    lineSpacingMultiple: 1.25,
    paraSpaceBefore: 1,
  });
  return y + h + 0.06;
}

// --- addKpiCard: single KPI box (dark variant)
function addKpiCard(slide, x, y, w, h, value, label, opts = {}) {
  const bg = opts.bg || COL.navyMid;
  slide.addShape(pres.ShapeType.rect, {
    x, y, w, h, fill: { color: bg },
    line: { color: COL.gold, pt: 1.5 },
  });
  slide.addText(value, {
    x, y: y + h * 0.1, w, h: h * 0.5,
    fontSize: opts.valueSz || 22, fontFace: FONT_HEAD,
    color: opts.valueColor || COL.gold,
    bold: true, align: "center", valign: "middle",
  });
  slide.addText(label, {
    x, y: y + h * 0.6, w, h: h * 0.35,
    fontSize: opts.labelSz || 11, fontFace: FONT_BODY,
    color: COL.ice, align: "center", valign: "top",
    lineSpacingMultiple: 1.1,
  });
}

// --- addLightKpiCard: single KPI box (light variant)
function addLightKpiCard(slide, x, y, w, h, value, label, opts = {}) {
  slide.addShape(pres.ShapeType.rect, {
    x, y, w, h, fill: { color: COL.bgCard },
    line: { color: COL.grayLight, pt: 0.75 },
  });
  slide.addShape(pres.ShapeType.rect, {
    x, y, w, h: 0.04, fill: { color: opts.accent || COL.navy },
  });
  slide.addText(value, {
    x, y: y + 0.08, w, h: h * 0.52,
    fontSize: opts.valueSz || 20, fontFace: FONT_HEAD,
    color: opts.valueColor || COL.navy,
    bold: true, align: "center", valign: "middle",
  });
  slide.addText(label, {
    x, y: y + h * 0.6, w, h: h * 0.38,
    fontSize: MIN_FONT, fontFace: FONT_BODY,
    color: COL.grayDark, align: "center", valign: "top",
    lineSpacingMultiple: 1.1,
  });
}

// --- addTable: styled table helper
function addTable(slide, rows, x, y, w, opts = {}) {
  const colW = opts.colW;
  const rowH = opts.rowH || 0.28;
  const mapped = rows.map((row, ri) =>
    row.map((cell, ci) => {
      const isHeader = ri === 0;
      const isFirstCol = ci === 0;
      return {
        text: String(cell),
        options: {
          fill: { color: isHeader ? COL.navy : (ri % 2 === 0 ? COL.bgCard : COL.stripe) },
          color: isHeader ? COL.white : (isFirstCol ? COL.navy : COL.grayDark),
          fontSize: opts.fontSize || MIN_FONT,
          fontFace: FONT_BODY,
          bold: isHeader || (isFirstCol && !isHeader),
          align: (isFirstCol && !isHeader) ? "left" : "center",
          valign: "middle",
          margin: [2, 4, 2, 4],
        },
      };
    })
  );
  slide.addTable(mapped, {
    x, y, w,
    colW,
    rowH,
    border: { pt: 0.5, color: COL.grayLight },
  });
  return y + rows.length * rowH + 0.05;
}

// --- callout box (insight/note)
function addCallout(slide, text, x, y, w, h, opts = {}) {
  const bg = opts.bg || COL.navy;
  const fg = opts.fg || COL.white;
  slide.addShape(pres.ShapeType.rect, {
    x, y, w, h, fill: { color: bg },
    line: { color: opts.border || COL.gold, pt: 1 },
  });
  slide.addText(text, {
    x, y, w, h,
    fontSize: opts.fontSize || MIN_FONT, fontFace: FONT_BODY,
    color: fg, align: "center", valign: "middle",
    bold: opts.bold || false,
    lineSpacingMultiple: 1.2,
    margin: [4, 6, 4, 6],
  });
}

// ============================================================
// LAYER 3: SLIDES
// ============================================================

// ------------------------------------------------------------
// SLIDE 1 — Title (dark)
// ------------------------------------------------------------
{
  const slide = pres.addSlide();
  slide.background = { color: COL.navy };

  // Decorative gold bar top
  slide.addShape(pres.ShapeType.rect, {
    x: 0, y: 0, w: W, h: 0.07, fill: { color: COL.gold },
  });
  // Side accent
  slide.addShape(pres.ShapeType.rect, {
    x: 6.8, y: 0, w: 3.2, h: H, fill: { color: COL.navyMid },
  });
  // Vertical gold rule
  slide.addShape(pres.ShapeType.rect, {
    x: 6.8, y: 0, w: 0.04, h: H, fill: { color: COL.gold },
  });

  // UAL logo text placeholder
  slide.addText("UAL", {
    x: 7.1, y: 0.3, w: 2.6, h: 0.7,
    fontSize: 48, fontFace: FONT_HEAD, color: COL.gold,
    bold: true, align: "center",
  });
  slide.addText("NASDAQ: UAL", {
    x: 7.1, y: 0.9, w: 2.6, h: 0.28,
    fontSize: 11, fontFace: FONT_BODY, color: COL.ice,
    align: "center",
  });

  // Main title
  slide.addText("United Airlines Holdings", {
    x: PAD, y: 1.1, w: 6.1, h: 0.7,
    fontSize: 32, fontFace: FONT_HEAD,
    color: COL.white, bold: true, align: "left",
  });
  slide.addText("FY2025 Results & 2026 Outlook", {
    x: PAD, y: 1.78, w: 6.1, h: 0.45,
    fontSize: 20, fontFace: FONT_HEAD,
    color: COL.gold, bold: false, align: "left",
  });
  slide.addText("Institutional Investor Presentation", {
    x: PAD, y: 2.32, w: 6.1, h: 0.3,
    fontSize: 13, fontFace: FONT_BODY,
    color: COL.ice, align: "left",
  });
  slide.addText("January 2026", {
    x: PAD, y: 2.72, w: 6.1, h: 0.28,
    fontSize: 12, fontFace: FONT_BODY,
    color: COL.gray, align: "left",
  });

  // Bottom bar
  slide.addShape(pres.ShapeType.rect, {
    x: 0, y: FOOTER_Y, w: W, h: H - FOOTER_Y,
    fill: { color: "000000", transparency: 40 },
  });
  slide.addText("Star Alliance Member  |  World's Largest Airline by ASMs  |  Chicago, IL", {
    x: PAD, y: FOOTER_Y + 0.03, w: W - PAD * 2, h: 0.2,
    fontSize: 8, fontFace: FONT_BODY, color: COL.ice, align: "left",
  });
}

// ------------------------------------------------------------
// SLIDE 2 — Investment Thesis (dark)
// ------------------------------------------------------------
{
  const slide = pres.addSlide();
  slide.background = { color: COL.navy };

  let y = addHeader(slide, "Investment Thesis", true);
  addFooter(slide, true);

  // 3 KPI cards
  const cards = [
    { v: "$59.1B", l: "FY2025 Revenue" },
    { v: "$3.4B",  l: "FY2025 Net Income" },
    { v: "$10.62", l: "FY2025 Adj. EPS" },
  ];
  const cw = 2.6, ch = 1.1, gap = 0.25;
  const startX = PAD;
  cards.forEach((c, i) => {
    addKpiCard(slide, startX + i * (cw + gap), y, cw, ch, c.v, c.l, { valueSz: 26 });
  });

  y += ch + 0.2;

  y = addSectionLabel(slide, "Strategic Investment Case", y, true);

  slide.addText(
    "Labor headwinds now largely behind. Fleet efficiency is compounding as Boeing MAX and Airbus A321neo deliveries continue. 2026 marks the first year United is expected to materially exceed the 2019 peak EPS of $12.05 — driven by structural margin expansion, premium revenue growth, and disciplined capacity deployment.",
    {
      x: PAD, y, w: W - PAD * 2, h: 1.2,
      fontSize: 12, fontFace: FONT_BODY,
      color: COL.ice, align: "left",
      lineSpacingMultiple: 1.4,
    }
  );

  // Accent bar left
  slide.addShape(pres.ShapeType.rect, {
    x: PAD - 0.12, y, w: 0.05, h: 1.2,
    fill: { color: COL.gold },
  });
}

// ------------------------------------------------------------
// SLIDE 3 — Multi-Year Financial Summary (light)
// ------------------------------------------------------------
{
  const slide = pres.addSlide();
  slide.background = { color: COL.bgLight };

  let y = addHeader(slide, "Multi-Year Financial Summary  |  FY2021–FY2026E", false);
  addFooter(slide, false);

  const rows = [
    ["Year", "Revenue ($B)", "Op Inc ($B)", "Net Inc ($B)", "Adj EPS", "Adj EBITDA ($B)", "ASMs (B)", "Fleet"],
    ["FY2021",  "24.6", "(1.0)", "(1.9)", "neg",    "—",   "178.7", "~1,250"],
    ["FY2022",  "45.0", "2.3",  "0.7",   "$10.61", "—",   "247.9", "~1,300"],
    ["FY2023",  "53.7", "4.2",  "2.6",   "$10.05", "7.9", "291.3", "1,358"],
    ["FY2024",  "57.1", "5.1",  "3.1",   "$10.61", "8.2", "311.2", "1,406"],
    ["FY2025",  "59.1", "4.7",  "3.4",   "$10.62", "8.1", "330.3", "1,490"],
    ["FY2026E", "~62",  "—",    "—",     "$12–14", "—",   "~349",  "~1,610E"],
  ];

  const colW = [0.85, 1.15, 1.0, 1.1, 0.9, 1.2, 1.0, 0.8];
  addTable(slide, rows, PAD, y, W - PAD * 2, { colW, rowH: 0.31, fontSize: 9 });
}

// ------------------------------------------------------------
// SLIDE 4 — FY2025 Quarterly P&L (light)
// ------------------------------------------------------------
{
  const slide = pres.addSlide();
  slide.background = { color: COL.bgLight };

  let y = addHeader(slide, "FY2025 Quarterly Profit & Loss", false);
  addFooter(slide, false);

  const qRows = [
    ["Metric",          "Q1 2025", "Q2 2025", "Q3 2025", "Q4 2025", "FY2025"],
    ["Revenue ($B)",    "12.5",    "14.2",    "15.0",    "17.4",    "59.1"],
    ["Op Income ($B)",  "0.3",     "1.5",     "1.7",     "1.2",     "4.7"],
    ["Net Income ($B)", "0.0",     "1.0",     "1.2",     "1.2",     "3.4"],
    ["Adj EPS",         "$0.91",   "$3.51",   "$3.98",   "$2.22",   "$10.62"],
  ];

  const colW = [1.7, 1.3, 1.3, 1.3, 1.3, 1.1];
  y = addTable(slide, qRows, PAD, y, W - PAD * 2, { colW, rowH: 0.32, fontSize: 9 });

  // Insight callouts
  const callouts = [
    { t: "Q4 Net Income $1.2B\n+49% YoY", c: COL.green },
    { t: "FY2025 Adj EPS $10.62\nvs $10.61 in FY2024", c: COL.navy },
    { t: "Q3 strongest quarter\nEPS $3.98 record", c: COL.navyMid },
  ];
  const cw = (W - PAD * 2 - 0.2) / 3;
  callouts.forEach((c, i) => {
    addCallout(slide, c.t, PAD + i * (cw + 0.1), y + 0.05, cw, 0.7,
      { bg: c.c, fg: COL.white, fontSize: 10, bold: true });
  });
}

// ------------------------------------------------------------
// SLIDE 5 — Unit Economics (light)
// ------------------------------------------------------------
{
  const slide = pres.addSlide();
  slide.background = { color: COL.bgLight };

  let y = addHeader(slide, "Unit Economics  |  FY2025", false);
  addFooter(slide, false);

  y = addSectionLabel(slide, "Cost & Revenue per ASM", y, false);

  const rows = [
    ["Metric",                    "FY2025 Value", "vs FY2024"],
    ["CASM (¢)",                  "15.96",        "+1.2%"],
    ["CASM ex-Fuel (¢)",          "11.20",        "+2.1%"],
    ["PRASM (¢)",                 "16.18",        "+1.8%"],
    ["TRASM (¢)",                 "17.89",        "+1.5%"],
    ["Fuel Cost / Gallon ($)",    "$2.62",        "-4.0%"],
    ["Fuel Burn (B gallons)",     "4.0",          "+3.8%"],
  ];

  const colW = [3.2, 2.0, 1.8];
  y = addTable(slide, rows, PAD, y, 7.0, { colW, rowH: 0.30, fontSize: 9 });

  // Fuel sensitivity callout
  addCallout(slide,
    "Fuel Sensitivity: $0.10/gal change ≈ $40M annual fuel cost impact",
    PAD, y + 0.1, W - PAD * 2, 0.45,
    { bg: "FEF3C7", fg: COL.grayDark, border: COL.amber, fontSize: 10, bold: true }
  );
}

// ------------------------------------------------------------
// SLIDE 6 — Revenue Composition (light)
// ------------------------------------------------------------
{
  const slide = pres.addSlide();
  slide.background = { color: COL.bgLight };

  let y = addHeader(slide, "Revenue Composition  |  FY2025", false);
  addFooter(slide, false);

  y = addSectionLabel(slide, "Revenue Breakdown by Segment", y, false);

  const rows = [
    ["Segment",          "Revenue", "% of Total", "YoY Change"],
    ["Passenger",        "$53.4B",  "90.4%",      "+3.5%"],
    ["Cargo",            "$1.8B",   "3.0%",        "+6.2%"],
    ["Other / Ancillary","$3.9B",   "6.6%",       "+5.4%"],
    ["Total",            "$59.1B",  "100%",       "+3.5%"],
  ];
  const colW = [2.8, 1.6, 1.4, 1.4];
  y = addTable(slide, rows, PAD, y, 7.2, { colW, rowH: 0.30, fontSize: 9 });

  y += 0.1;
  addCallout(slide,
    "Premium Revenue +8% YoY — Polaris & Premium Plus outperforming. MileagePlus loyalty program contributes significant ancillary revenue and provides $20B+ in collateral value.",
    PAD, y, W - PAD * 2, 0.65,
    { bg: COL.navy, fg: COL.white, border: COL.gold, fontSize: 10 }
  );
}

// ------------------------------------------------------------
// SLIDE 7 — Fleet Modernization (light)
// ------------------------------------------------------------
{
  const slide = pres.addSlide();
  slide.background = { color: COL.bgLight };

  let y = addHeader(slide, "Fleet Modernization", false);
  addFooter(slide, false);

  // KPI mini-cards
  const cards = [
    { v: "1,490", l: "Total Fleet\n(+84 YoY)" },
    { v: "200+",  l: "737 MAX\nDeliveries Pending" },
    { v: "~20%",  l: "A321neo CASM\nImprovement vs Legacy" },
  ];
  const cw = 2.6, ch = 0.9, gap = 0.25;
  cards.forEach((c, i) => {
    addLightKpiCard(slide, PAD + i * (cw + gap), y, cw, ch, c.v, c.l, { valueSz: 22 });
  });
  y += ch + 0.18;

  const rows = [
    ["Aircraft Type",    "Count", "Avg Age (yr)", "Fuel Efficiency vs Replaced"],
    ["Boeing 737 MAX 8", "~280",  "2.1",          "+14% vs 737-900ER"],
    ["Boeing 737 MAX 9", "~190",  "1.8",          "+14% vs 737-900ER"],
    ["Airbus A321neo",   "~60",   "1.5",          "+20% vs A319/A320"],
    ["Boeing 787-9/10",  "~120",  "4.2",          "+25% vs 747/767"],
    ["Other types",      "~840",  "11.3",         "Various"],
  ];
  const colW = [2.2, 0.8, 1.3, 2.7];
  y = addTable(slide, rows, PAD, y, W - PAD * 2, { colW, rowH: 0.28, fontSize: 9 });

  addCallout(slide,
    "Risk: Boeing production delays continue to impact MAX deliveries. UAL working with Boeing on revised schedules; mitigation includes leasing and retaining older 737s.",
    PAD, y + 0.08, W - PAD * 2, 0.5,
    { bg: "FDECEA", fg: COL.grayDark, border: COL.red, fontSize: MIN_FONT }
  );
}

// ------------------------------------------------------------
// SLIDE 8 — New Products & Strategy (light)
// ------------------------------------------------------------
{
  const slide = pres.addSlide();
  slide.background = { color: COL.bgLight };

  let y = addHeader(slide, "New Products & Strategy", false);
  addFooter(slide, false);

  // 2x2 card grid
  const prodCards = [
    {
      title: "Polaris Business Class",
      body:  "787 premium long-haul cabins — lie-flat suites, direct aisle access, elevated dining. Highest customer satisfaction scores in United history. Capacity expanding on key transcon and transpacific routes.",
    },
    {
      title: "United Clubs Expansion",
      body:  "Multi-year program to add and renovate lounge space at key hubs. New flagship clubs at ORD, EWR, IAH. Membership revenue growing double-digits; exclusivity maintained via capacity management.",
    },
    {
      title: "Basic Economy Optimization",
      body:  "Refined yield management between Basic and regular economy. Driving higher PRASM while retaining price-sensitive travelers. Ancillary attach rates increasing with improved digital merchandising.",
    },
    {
      title: "Transcon 'Elevated' Product",
      body:  "Premium transcontinental service ORD–LAX/SFO and EWR–LAX/SFO. Targeted at corporate and leisure premium travelers. Competing directly with AA Flagship and DL premium offerings.",
    },
  ];

  const cardW = (W - PAD * 2 - 0.15) / 2;
  const cardH = (CONTENT_BOTTOM - y - 0.05) / 2 - 0.1;

  prodCards.forEach((pc, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const cx = PAD + col * (cardW + 0.15);
    const cy = y + row * (cardH + 0.1);

    slide.addShape(pres.ShapeType.rect, {
      x: cx, y: cy, w: cardW, h: cardH,
      fill: { color: COL.bgCard },
      line: { color: COL.grayLight, pt: 0.75 },
    });
    slide.addShape(pres.ShapeType.rect, {
      x: cx, y: cy, w: cardW, h: 0.04,
      fill: { color: COL.gold },
    });
    slide.addText(pc.title, {
      x: cx + 0.1, y: cy + 0.08, w: cardW - 0.2, h: 0.28,
      fontSize: 11, fontFace: FONT_HEAD, bold: true,
      color: COL.navy, align: "left",
    });
    slide.addText(pc.body, {
      x: cx + 0.1, y: cy + 0.34, w: cardW - 0.2, h: cardH - 0.4,
      fontSize: MIN_FONT, fontFace: FONT_BODY,
      color: COL.grayDark, align: "left",
      lineSpacingMultiple: 1.25,
    });
  });
}

// ------------------------------------------------------------
// SLIDE 9 — Starlink Connectivity (dark)
// ------------------------------------------------------------
{
  const slide = pres.addSlide();
  slide.background = { color: COL.navy };

  let y = addHeader(slide, "Starlink In-Flight Connectivity", true);
  addFooter(slide, true);

  // Timeline
  y = addSectionLabel(slide, "Rollout Timeline", y, true);

  const milestones = [
    { label: "2025 H1", desc: "Rollout commenced across narrowbody fleet" },
    { label: "End-2025", desc: "200+ aircraft equipped; customer availability live" },
    { label: "2026 Full Fleet", desc: "Target: entire mainline fleet connected" },
  ];

  const mw = (W - PAD * 2 - 0.3) / 3;
  milestones.forEach((m, i) => {
    const mx = PAD + i * (mw + 0.15);
    slide.addShape(pres.ShapeType.rect, {
      x: mx, y, w: mw, h: 0.8,
      fill: { color: COL.navyMid },
      line: { color: COL.gold, pt: 1 },
    });
    slide.addText(m.label, {
      x: mx + 0.06, y: y + 0.04, w: mw - 0.12, h: 0.26,
      fontSize: 11, fontFace: FONT_HEAD, bold: true,
      color: COL.gold, align: "center",
    });
    slide.addText(m.desc, {
      x: mx + 0.06, y: y + 0.3, w: mw - 0.12, h: 0.44,
      fontSize: MIN_FONT, fontFace: FONT_BODY,
      color: COL.ice, align: "center", lineSpacingMultiple: 1.2,
    });
  });
  y += 0.8 + 0.22;

  // Key metrics
  y = addSectionLabel(slide, "Key Metrics & Revenue Opportunity", y, true);

  const kpis = [
    { v: "100 Mbps", l: "Per-aircraft bandwidth\n(SpaceX Starlink)" },
    { v: "$500M+",   l: "Ancillary revenue\nopportunity (est.)" },
    { v: "~2,000+",  l: "Planned aircraft\nby end-2026" },
    { v: "Free Wi-Fi",l: "For MileagePlus\nElite members" },
  ];
  const kw = (W - PAD * 2 - 0.45) / 4;
  kpis.forEach((k, i) => {
    addKpiCard(slide, PAD + i * (kw + 0.15), y, kw, 1.05, k.v, k.l,
      { valueSz: 15, labelSz: 11 });
  });
}

// ------------------------------------------------------------
// SLIDE 10 — Operational Performance (light)
// ------------------------------------------------------------
{
  const slide = pres.addSlide();
  slide.background = { color: COL.bgLight };

  let y = addHeader(slide, "Operational Performance  |  FY2025", false);
  addFooter(slide, false);

  // KPI cards row
  const opCards = [
    { v: "83%",   l: "On-Time Arrival\n(DOT A14)" },
    { v: "99.3%", l: "Completion\nFactor" },
    { v: "#1",    l: "DOT Customer\nComplaints (major carriers)" },
    { v: "+6%",   l: "Corporate Travel\nRecovery YoY" },
  ];
  const cw = (W - PAD * 2 - 0.45) / 4;
  const ch = 1.0;
  opCards.forEach((c, i) => {
    addLightKpiCard(slide, PAD + i * (cw + 0.15), y, cw, ch, c.v, c.l, { valueSz: 20 });
  });
  y += ch + 0.2;

  y = addSectionLabel(slide, "Demand Signals & Forward Indicators", y, false);

  const rows = [
    ["Indicator",                    "Status",       "Commentary"],
    ["Forward Bookings",             "+5–7% YoY",    "Domestic and international leisure strong"],
    ["Corporate Travel Recovery",    "~95% of 2019", "Tech and financial services leading"],
    ["Premium Cabin Load Factor",    "81%",          "Polaris and P+ running near full"],
    ["International Long-Haul LF",  "86%",          "Trans-Pacific and Trans-Atlantic robust"],
    ["Loyalty Enrollments",          "+12% YoY",     "MileagePlus 110M+ members"],
  ];
  const colW = [2.6, 1.6, 4.8];
  addTable(slide, rows, PAD, y, W - PAD * 2, { colW, rowH: 0.28, fontSize: 9 });
}

// ------------------------------------------------------------
// SLIDE 11 — Balance Sheet & Deleveraging (light)
// ------------------------------------------------------------
{
  const slide = pres.addSlide();
  slide.background = { color: COL.bgLight };

  let y = addHeader(slide, "Balance Sheet & Deleveraging Progress", false);
  addFooter(slide, false);

  const rows = [
    ["Metric",                  "FY2023",  "FY2024",  "FY2025"],
    ["Total Debt ($B)",         "$33.8",   "$31.5",   "$29.2"],
    ["Net Debt ($B)",           "~$30.0",  "~$27.8",  "~$26.0"],
    ["Adj Net Leverage (x)",    "3.5x",    "2.9x",    "2.6x"],
    ["Interest Coverage (x)",   "3.1x",    "3.8x",    "4.5x"],
    ["Cash & Equivalents ($B)", "$6.0",    "$6.3",    "$6.8"],
  ];
  const colW = [2.8, 1.6, 1.6, 1.6];
  y = addTable(slide, rows, PAD, y, 7.6, { colW, rowH: 0.30, fontSize: 9 });

  y += 0.1;
  addCallout(slide,
    "MileagePlus program: $20B+ estimated value — provides significant off-balance-sheet asset quality. Secured notes backed by MileagePlus IP and brand agreements.",
    PAD, y, W - PAD * 2, 0.55,
    { bg: COL.navy, fg: COL.white, border: COL.gold, fontSize: 10 }
  );
}

// ------------------------------------------------------------
// SLIDE 12 — Cash Flow & Capital Allocation (light)
// ------------------------------------------------------------
{
  const slide = pres.addSlide();
  slide.background = { color: COL.bgLight };

  let y = addHeader(slide, "Cash Flow & Capital Allocation  |  FY2025", false);
  addFooter(slide, false);

  // Cash flow KPIs
  const cfCards = [
    { v: "$5.8B",  l: "Operating\nCash Flow" },
    { v: "($3.7B)", l: "Capital\nExpenditures" },
    { v: "$2.1B",  l: "Free\nCash Flow" },
  ];
  const cw = 2.5, ch = 1.0, gap = 0.3;
  cfCards.forEach((c, i) => {
    addLightKpiCard(slide, PAD + i * (cw + gap), y, cw, ch, c.v, c.l,
      { valueSz: 22, accent: i === 2 ? COL.green : COL.navy });
  });
  y += ch + 0.2;

  y = addSectionLabel(slide, "Capital Allocation Priority Stack", y, false);

  const rows = [
    ["Priority",  "Initiative",                            "FY2025 Deployed", "Rationale"],
    ["1st",       "Debt Reduction",                        "$2.3B",           "De-lever to <2.0x target by 2027"],
    ["2nd",       "Fleet Investment (Capex)",              "$3.7B",           "Efficiency & growth capex"],
    ["3rd",       "Liquidity Buffer",                      "$0.5B",           "Maintain $6B+ cash target"],
    ["4th",       "Share Buybacks",                        "$0.5B",           "Authorized $1B program"],
  ];
  const colW = [0.65, 2.8, 1.5, 3.05];
  addTable(slide, rows, PAD, y, W - PAD * 2, { colW, rowH: 0.28, fontSize: 9 });
}

// ------------------------------------------------------------
// SLIDE 13 — 2026 Guidance (dark)
// ------------------------------------------------------------
{
  const slide = pres.addSlide();
  slide.background = { color: COL.navy };

  let y = addHeader(slide, "FY2026 Guidance", true);
  addFooter(slide, true);

  const kpis = [
    { v: "$12–14",  l: "Adjusted EPS\nGuidance Range" },
    { v: "~5.5%",   l: "Capacity Growth\n(ASM YoY)" },
    { v: "+LSD%",   l: "CASM ex-Fuel\nGrowth (low single digits)" },
    { v: "~15%+",   l: "Adj EBITDA\nMargin Target" },
    { v: "~5–6%",   l: "Revenue Growth\nYoY" },
    { v: "~$4.5B",  l: "Capital\nExpenditures" },
  ];

  const kw = (W - PAD * 2 - 0.5) / 3;
  const kh = (CONTENT_BOTTOM - y - 0.05) / 2 - 0.1;

  kpis.forEach((k, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    addKpiCard(
      slide,
      PAD + col * (kw + 0.25),
      y + row * (kh + 0.1),
      kw, kh,
      k.v, k.l,
      { valueSz: 22, labelSz: 11 }
    );
  });
}

// ------------------------------------------------------------
// SLIDE 14 — Risk Matrix (light)
// ------------------------------------------------------------
{
  const slide = pres.addSlide();
  slide.background = { color: COL.bgLight };

  let y = addHeader(slide, "Risk Matrix  |  FY2026 Key Risks", false);
  addFooter(slide, false);

  const rows = [
    ["Risk",                        "Likelihood", "Impact", "Mitigation"],
    ["Fuel Price Spike",            "Medium",     "High",   "Fuel hedging program; fuel efficient fleet investments"],
    ["Boeing Delivery Delays",      "High",       "Medium", "Leasing optionality; retaining older MAX capacity"],
    ["Macro Recession / Demand",    "Low-Medium", "High",   "Premium/loyalty revenue diversification; flex capacity"],
    ["Labor Disruption",            "Low",        "High",   "New contracts settled 2024–25; stable through 2027+"],
    ["Premium Competition",         "Medium",     "Medium", "Product differentiation; MileagePlus loyalty moat"],
    ["Regulatory / Policy Changes", "Low",        "Low",    "Proactive government engagement; legal reserves maintained"],
  ];
  const colW = [2.2, 1.2, 0.9, 4.7];
  addTable(slide, rows, PAD, y, W - PAD * 2, { colW, rowH: 0.3, fontSize: 9 });
}

// ------------------------------------------------------------
// SLIDE 15 — ESG & Sustainability (light)
// ------------------------------------------------------------
{
  const slide = pres.addSlide();
  slide.background = { color: COL.bgLight };

  let y = addHeader(slide, "ESG & Sustainability  |  Progress to 2030 Targets", false);
  addFooter(slide, false);

  const rows = [
    ["Metric",                     "FY2023",     "FY2025",     "2030 Target"],
    ["CO2 Emissions (MT, millions)","99.4",       "101.2",      "Target: -50% vs 2019"],
    ["SAF Usage (% of fuel)",       "0.1%",       "0.5%",       "10% by 2030"],
    ["Fleet Fuel Efficiency Idx",   "Base 100",   "94.3",       "85 (15% improvement)"],
    ["Scope 1 CO2 per ASM (g)",     "71.4",       "68.9",       "~55 by 2030"],
    ["Carbon Offsets Purchased",    "~2.0M ton",  "~3.5M ton",  "Growing program"],
  ];
  const colW = [2.8, 1.5, 1.5, 3.2];
  y = addTable(slide, rows, PAD, y, W - PAD * 2, { colW, rowH: 0.28, fontSize: 9 });

  y += 0.1;

  // ESG program cards
  const esgCards = [
    { t: "SAF Partnerships", b: "Long-term SAF offtake agreements with Neste, World Energy, and others. Path to 10% SAF by 2030." },
    { t: "Carbon Offsets",   b: "Nature-based and technology-based carbon removal portfolio. Verified Verra/Gold Standard credits." },
    { t: "Fleet Efficiency", b: "Boeing MAX and A321neo average 15–20% better fuel burn than replaced aircraft types." },
  ];
  const ew = (W - PAD * 2 - 0.2) / 3;
  const eh = Math.min(CONTENT_BOTTOM - y - 0.05, 1.0);
  esgCards.forEach((c, i) => {
    const ex = PAD + i * (ew + 0.1);
    slide.addShape(pres.ShapeType.rect, {
      x: ex, y: y + 0.05, w: ew, h: eh,
      fill: { color: COL.bgCard },
      line: { color: COL.grayLight, pt: 0.75 },
    });
    slide.addShape(pres.ShapeType.rect, {
      x: ex, y: y + 0.05, w: ew, h: 0.04,
      fill: { color: COL.green },
    });
    slide.addText(c.t, {
      x: ex + 0.08, y: y + 0.1, w: ew - 0.16, h: 0.22,
      fontSize: 10, fontFace: FONT_HEAD, bold: true,
      color: COL.navy, align: "left",
    });
    slide.addText(c.b, {
      x: ex + 0.08, y: y + 0.34, w: ew - 0.16, h: eh - 0.36,
      fontSize: MIN_FONT, fontFace: FONT_BODY,
      color: COL.grayDark, align: "left", lineSpacingMultiple: 1.2,
    });
  });
}

// ------------------------------------------------------------
// SLIDE 16 — Appendix (light)
// ------------------------------------------------------------
{
  const slide = pres.addSlide();
  slide.background = { color: COL.bgLight };

  let y = addHeader(slide, "Appendix  |  Reference Data", false);
  addFooter(slide, false);

  // Hub network table
  y = addSectionLabel(slide, "Hub Network", y, false);
  const hubRows = [
    ["Hub", "IATA", "Role",          "Daily Departures (est.)"],
    ["Chicago O'Hare",    "ORD", "Primary domestic/intl", "~500"],
    ["Newark/New York",   "EWR", "Northeast gateway",     "~400"],
    ["Houston Intercontinental","IAH","Latin America/Gulf","~350"],
    ["Denver",            "DEN", "Mountain West hub",     "~280"],
    ["San Francisco",     "SFO", "Trans-Pacific gateway", "~260"],
    ["Los Angeles",       "LAX", "West Coast hub",        "~200"],
    ["Washington Dulles", "IAD", "Government/premium",    "~180"],
  ];
  const colW = [2.4, 0.7, 2.4, 2.0];
  y = addTable(slide, hubRows, PAD, y, 7.5, { colW, rowH: 0.27, fontSize: 8.5 });

  // Workforce & long-term targets side by side
  const remaining = CONTENT_BOTTOM - y - 0.1;
  if (remaining > 0.5) {
    y += 0.08;
    y = addSectionLabel(slide, "Workforce & Long-Term Targets", y, false);
    slide.addText(
      "Employees: ~100,000 (FY2025)\nUnionized: ~84%\nPilot headcount: ~16,000+\nNew contracts settled: ALPA, IAM, AFA (2024-25)\n\nLong-term targets:\n• Adj EPS $15+ by 2028\n• Net leverage <2.0x by 2027\n• CASM ex-fuel flat to -1% by 2027\n• ASM growth 4–6% annually",
      {
        x: PAD, y, w: W - PAD * 2, h: Math.min(remaining - 0.05, 1.05),
        fontSize: MIN_FONT, fontFace: FONT_BODY,
        color: COL.grayDark, align: "left",
        lineSpacingMultiple: 1.3,
      }
    );
  }
}

// ============================================================
// WRITE OUTPUT
// ============================================================
const outputPath = path.resolve(__dirname, "../outputs/corporate-deck-builder.pptx");
pres.writeFile({ fileName: outputPath })
  .then(() => {
    console.log("SUCCESS: " + outputPath);
  })
  .catch((err) => {
    console.error("ERROR:", err);
    process.exit(1);
  });
