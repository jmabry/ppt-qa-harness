const pptxgen = require("pptxgenjs");

// ============================================================
// United Airlines Holdings — FY2025 Results & 2026 Outlook
// Institutional Investor Presentation
// ============================================================

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.author = "United Airlines Holdings, Inc.";
pres.title = "UAL FY2025 Results & 2026 Outlook";

// ============================================================
// COLOR PALETTE — Midnight Executive
// ============================================================
const C = {
  navy:       "1E2761",
  deepNavy:   "0F1535",
  midNavy:    "162050",
  ice:        "CADCFC",
  white:      "FFFFFF",
  offWhite:   "F0F3FA",
  lightGray:  "8B99BD",
  medGray:    "5A6A94",
  accent:     "3B82F6",  // bright blue accent
  accentDark: "2563EB",
  green:      "10B981",
  red:        "EF4444",
  amber:      "F59E0B",
  gold:       "FBBF24",
  charcoal:   "1E293B",
  slate:      "64748B",
  cardBg:     "F8FAFC",
  cardBorder: "E2E8F0",
};

// ============================================================
// TYPOGRAPHY
// ============================================================
const FONT = {
  head: "Arial Black",
  sub:  "Arial",
  body: "Calibri",
};

// ============================================================
// HELPERS
// ============================================================
const SLIDE_W = 10;
const SLIDE_H = 5.625;
const MARGIN = 0.55;
const CONTENT_W = SLIDE_W - 2 * MARGIN;

function makeShadow(opts = {}) {
  return {
    type: "outer",
    color: "000000",
    blur: opts.blur || 4,
    offset: opts.offset || 1,
    angle: opts.angle || 135,
    opacity: opts.opacity || 0.12,
  };
}

function addDarkSlide() {
  const slide = pres.addSlide();
  slide.background = { color: C.deepNavy };
  return slide;
}

function addLightSlide() {
  const slide = pres.addSlide();
  slide.background = { color: C.offWhite };
  return slide;
}

// Slide number footer
function addFooter(slide, num, total, dark = false) {
  const color = dark ? C.lightGray : C.slate;
  slide.addText(`${num} / ${total}`, {
    x: SLIDE_W - 1.5,
    y: SLIDE_H - 0.35,
    w: 1.2,
    h: 0.25,
    fontSize: 8,
    fontFace: FONT.body,
    color: color,
    align: "right",
    margin: 0,
  });
  slide.addText("UNITED AIRLINES HOLDINGS  |  NASDAQ: UAL", {
    x: MARGIN,
    y: SLIDE_H - 0.35,
    w: 4,
    h: 0.25,
    fontSize: 7,
    fontFace: FONT.body,
    color: color,
    align: "left",
    margin: 0,
    charSpacing: 1,
  });
}

// Section title on a dark slide
function addSectionTitle(slide, title, subtitle) {
  slide.addText(title.toUpperCase(), {
    x: MARGIN,
    y: 1.6,
    w: CONTENT_W,
    h: 0.7,
    fontSize: 28,
    fontFace: FONT.head,
    color: C.white,
    bold: true,
    align: "left",
    margin: 0,
    charSpacing: 3,
  });
  if (subtitle) {
    slide.addText(subtitle, {
      x: MARGIN,
      y: 2.35,
      w: CONTENT_W * 0.75,
      h: 0.5,
      fontSize: 13,
      fontFace: FONT.body,
      color: C.ice,
      align: "left",
      margin: 0,
      lineSpacingMultiple: 1.3,
    });
  }
}

// Light slide header
function addLightHeader(slide, title, subtitle) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: SLIDE_W, h: 0.85,
    fill: { color: C.navy },
  });
  slide.addText(title.toUpperCase(), {
    x: MARGIN,
    y: 0.12,
    w: CONTENT_W,
    h: 0.42,
    fontSize: 16,
    fontFace: FONT.head,
    color: C.white,
    bold: true,
    align: "left",
    margin: 0,
    charSpacing: 2,
  });
  if (subtitle) {
    slide.addText(subtitle, {
      x: MARGIN,
      y: 0.5,
      w: CONTENT_W,
      h: 0.3,
      fontSize: 10,
      fontFace: FONT.body,
      color: C.ice,
      align: "left",
      margin: 0,
    });
  }
}

// Stat card on dark background
function addStatCard(slide, x, y, w, h, value, label, opts = {}) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w, h,
    fill: { color: opts.bg || C.midNavy },
    shadow: makeShadow(),
  });
  slide.addText(value, {
    x, y: y + h * 0.12, w, h: h * 0.52,
    fontSize: opts.valueFontSize || 30,
    fontFace: FONT.head,
    color: opts.valueColor || C.accent,
    bold: true,
    align: "center",
    valign: "middle",
    margin: 0,
  });
  slide.addText(label, {
    x: x + 0.1, y: y + h * 0.6, w: w - 0.2, h: h * 0.35,
    fontSize: opts.labelFontSize || 9,
    fontFace: FONT.body,
    color: opts.labelColor || C.ice,
    align: "center",
    valign: "top",
    margin: 0,
    lineSpacingMultiple: 1.2,
  });
}

// Stat card on light background
function addLightStatCard(slide, x, y, w, h, value, label, opts = {}) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w, h,
    fill: { color: C.white },
    shadow: makeShadow(),
  });
  // Left accent bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w: 0.06, h,
    fill: { color: opts.accentColor || C.accent },
  });
  slide.addText(value, {
    x, y: y + h * 0.08, w, h: h * 0.52,
    fontSize: opts.valueFontSize || 28,
    fontFace: FONT.head,
    color: opts.valueColor || C.navy,
    bold: true,
    align: "center",
    valign: "middle",
    margin: 0,
  });
  slide.addText(label, {
    x: x + 0.1, y: y + h * 0.58, w: w - 0.2, h: h * 0.38,
    fontSize: 9,
    fontFace: FONT.body,
    color: C.slate,
    align: "center",
    valign: "top",
    margin: 0,
    lineSpacingMultiple: 1.15,
  });
}

// Table helper for light slides
function addStyledTable(slide, rows, x, y, w, opts = {}) {
  const headerOpts = {
    fill: { color: C.navy },
    color: C.white,
    bold: true,
    fontSize: 8.5,
    fontFace: FONT.body,
    align: "center",
    valign: "middle",
  };
  const cellOpts = (rowIdx) => ({
    fill: { color: rowIdx % 2 === 0 ? C.white : C.offWhite },
    color: C.charcoal,
    fontSize: 8.5,
    fontFace: FONT.body,
    align: "center",
    valign: "middle",
  });

  const tableRows = rows.map((row, rIdx) =>
    row.map((cell) => ({
      text: String(cell),
      options: rIdx === 0 ? headerOpts : cellOpts(rIdx),
    }))
  );

  slide.addTable(tableRows, {
    x, y, w,
    border: { pt: 0.5, color: C.cardBorder },
    colW: opts.colW,
    rowH: opts.rowH || 0.3,
    margin: [2, 4, 2, 4],
  });
}

const TOTAL_SLIDES = 18;

// ============================================================
// SLIDE 1 — TITLE
// ============================================================
{
  const slide = addDarkSlide();

  // Large background shape for visual interest
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: SLIDE_W, h: SLIDE_H,
    fill: { color: C.deepNavy },
  });
  // Top accent bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: SLIDE_W, h: 0.06,
    fill: { color: C.accent },
  });
  // Decorative side element
  slide.addShape(pres.shapes.RECTANGLE, {
    x: SLIDE_W - 2.5, y: 0, w: 2.5, h: SLIDE_H,
    fill: { color: C.navy, transparency: 60 },
  });

  slide.addText("UNITED AIRLINES", {
    x: MARGIN,
    y: 1.0,
    w: 7,
    h: 0.7,
    fontSize: 38,
    fontFace: FONT.head,
    color: C.white,
    bold: true,
    align: "left",
    margin: 0,
    charSpacing: 4,
  });
  slide.addText("HOLDINGS, INC.", {
    x: MARGIN,
    y: 1.6,
    w: 7,
    h: 0.5,
    fontSize: 22,
    fontFace: FONT.sub,
    color: C.ice,
    align: "left",
    margin: 0,
    charSpacing: 6,
  });
  slide.addText("FY2025 Full-Year Results & 2026 Outlook", {
    x: MARGIN,
    y: 2.5,
    w: 7,
    h: 0.4,
    fontSize: 16,
    fontFace: FONT.body,
    color: C.accent,
    align: "left",
    margin: 0,
  });
  slide.addText("Investor Presentation  |  January 20, 2026", {
    x: MARGIN,
    y: 3.0,
    w: 7,
    h: 0.35,
    fontSize: 12,
    fontFace: FONT.body,
    color: C.lightGray,
    align: "left",
    margin: 0,
  });

  // Bottom info bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: SLIDE_H - 0.55, w: SLIDE_W, h: 0.55,
    fill: { color: C.navy },
  });
  slide.addText("NASDAQ: UAL    |    Star Alliance Member    |    World's Largest Airline by ASMs", {
    x: MARGIN,
    y: SLIDE_H - 0.5,
    w: CONTENT_W,
    h: 0.4,
    fontSize: 9,
    fontFace: FONT.body,
    color: C.lightGray,
    align: "center",
    margin: 0,
    charSpacing: 1,
  });
}

// ============================================================
// SLIDE 2 — INVESTMENT THESIS
// ============================================================
{
  const slide = addDarkSlide();
  addSectionTitle(slide, "Investment Thesis", null);
  addFooter(slide, 2, TOTAL_SLIDES, true);

  // Accent bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x: MARGIN, y: 2.3, w: 0.8, h: 0.04,
    fill: { color: C.accent },
  });

  const thesisText = [
    { text: "United Airlines is the world's largest airline by ASMs, operating from seven fortress hubs with #1 market share in six of seven. ", options: { fontSize: 12, fontFace: FONT.body, color: C.ice, breakLine: true, lineSpacingMultiple: 1.5 } },
    { text: "\n", options: { fontSize: 6, breakLine: true } },
    { text: "The 2022\u20132025 EPS plateau ($10.05\u2013$10.62) masks a deliberate transformation: revenue grew 31% while the company absorbed $10B+ in pilot contract costs and $2B+ in flight attendant increases \u2014 the largest labor cost step-up in airline history. That cost absorption cycle is now largely complete.", options: { fontSize: 12, fontFace: FONT.body, color: C.ice, breakLine: true, lineSpacingMultiple: 1.5 } },
    { text: "\n", options: { fontSize: 6, breakLine: true } },
    { text: "2026 guidance of $12\u201314 EPS (+22% at midpoint) marks the first time UAL would materially exceed its 2019 peak of $12.05, signaling the multi-year fleet investment cycle is producing returns rather than absorbing costs. Premium revenue (+11% YoY), a loyalty program potentially worth $22B+ standalone, 630+ aircraft on order, and a clear path to investment-grade credit create a compounding value creation story that the market has not fully re-rated.", options: { fontSize: 12, fontFace: FONT.body, color: C.ice, lineSpacingMultiple: 1.5 } },
  ];

  slide.addText(thesisText, {
    x: MARGIN,
    y: 2.55,
    w: CONTENT_W * 0.85,
    h: 2.7,
    margin: 0,
    valign: "top",
  });

  // Right side stat highlights
  const statX = 8.0;
  addStatCard(slide, statX, 1.0, 1.55, 0.95, "$59.1B", "FY2025 Revenue");
  addStatCard(slide, statX, 2.1, 1.55, 0.95, "$13.00", "FY2026E EPS\nMidpoint", { valueColor: C.green });
  addStatCard(slide, statX, 3.2, 1.55, 0.95, "630+", "Aircraft on Order");
  addStatCard(slide, statX, 4.3, 1.55, 0.95, "2.2\u00D7", "Net Leverage\n\u2192 <2.0\u00D7 Target");
}

// ============================================================
// SLIDE 3 — MULTI-YEAR FINANCIAL PERFORMANCE
// ============================================================
{
  const slide = addLightSlide();
  addLightHeader(slide, "Multi-Year Financial Recovery Arc", "From COVID trough through 2025 actual and 2026 guidance \u2014 the EPS plateau explained");
  addFooter(slide, 3, TOTAL_SLIDES);

  const rows = [
    ["Metric", "FY2021", "FY2022", "FY2023", "FY2024", "FY2025", "FY2026E"],
    ["Revenue ($B)", "$24.6", "$45.0", "$53.7", "$57.1", "$59.1", "\u2014"],
    ["Adj. EPS", "neg.", "$10.61", "$10.05", "$10.61", "$10.62", "$12\u201314"],
    ["Net Income ($B)", "($2.0)", "$0.7", "$2.6", "$3.1", "$3.4", "\u2014"],
    ["Adj. Pre-Tax Margin", "neg.", "~6%", "8.0%", "8.1%", "7.8%", "~10%+"],
    ["ASMs (B)", "178.7", "247.9", "291.3", "311.2", "330.3", "~349"],
    ["Fleet (year-end)", "\u2014", "~1,300", "1,358", "1,406", "1,490", "~1,610"],
    ["Employees", "~75K", "~92K", "103K", "107K", "113K", "\u2014"],
  ];
  addStyledTable(slide, rows, MARGIN, 1.05, CONTENT_W, { rowH: 0.28 });

  // Key insight callout
  slide.addShape(pres.shapes.RECTANGLE, {
    x: MARGIN, y: 3.5, w: CONTENT_W, h: 1.7,
    fill: { color: C.white },
    shadow: makeShadow(),
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: MARGIN, y: 3.5, w: 0.06, h: 1.7,
    fill: { color: C.amber },
  });
  slide.addText("KEY CONTEXT", {
    x: MARGIN + 0.2, y: 3.55, w: 2, h: 0.28,
    fontSize: 9, fontFace: FONT.sub, color: C.amber, bold: true, margin: 0, charSpacing: 2,
  });
  slide.addText(
    "EPS flat at ~$10.60 for four years is not a value trap \u2014 it\u2019s a deliberate cost absorption cycle. Revenue grew 31% since 2022 on 33% ASM growth. The pilot contract ($10B cumulative, 34\u201340% raises) and flight attendant deal (20\u201328% raises) created a multi-year earnings plateau now largely behind the company. The 2026 $12\u201314 guidance is the first time UAL expects to materially exceed the 2019 peak of $12.05 EPS, signaling the investment cycle is finally producing returns.",
    {
      x: MARGIN + 0.2, y: 3.82, w: CONTENT_W - 0.4, h: 1.3,
      fontSize: 9.5, fontFace: FONT.body, color: C.charcoal, margin: 0,
      lineSpacingMultiple: 1.35,
    }
  );
}

// ============================================================
// SLIDE 4 — FY2025 QUARTERLY P&L BREAKDOWN
// ============================================================
{
  const slide = addLightSlide();
  addLightHeader(slide, "FY2025 Quarterly P&L", "Seasonal cadence, Q4 comp challenge, and the guidance revision/outperformance narrative");
  addFooter(slide, 4, TOTAL_SLIDES);

  const rows = [
    ["Quarter", "Total Rev", "YoY", "Pax Rev", "Cargo", "Other", "Net Income", "Adj. EPS", "Pre-Tax Mgn", "ASMs (B)"],
    ["Q1 2025", "$13.2B", "+5.4%", "$11.9B", "$0.43B", "$0.92B", "$0.39B", "$0.91", "3.0%", "75.2"],
    ["Q2 2025", "$15.2B", "+1.7%", "$13.8B", "$0.43B", "$0.97B", "$0.97B", "$3.87", "11.0%", "84.3"],
    ["Q3 2025", "$15.2B", "+2.6%", "$13.8B", "$0.43B", "$0.98B", "$0.95B", "$2.78", "8.0%", "87.4"],
    ["Q4 2025", "$15.4B", "+4.8%", "$13.9B", "$0.49B", "$0.98B", "$1.02B", "$3.10", "8.5%", "83.4"],
    ["FY 2025", "$59.1B", "+3.5%", "$53.4B", "$1.78B", "$3.85B", "$3.40B", "$10.62", "7.8%", "330.3"],
  ];
  addStyledTable(slide, rows, MARGIN, 1.0, CONTENT_W, { rowH: 0.28 });

  // Two insight cards below
  const cardY = 2.85;
  const cardW = (CONTENT_W - 0.2) / 2;

  // Card 1: Q4 comp
  slide.addShape(pres.shapes.RECTANGLE, {
    x: MARGIN, y: cardY, w: cardW, h: 1.55,
    fill: { color: C.white }, shadow: makeShadow(),
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: MARGIN, y: cardY, w: 0.06, h: 1.55,
    fill: { color: C.red },
  });
  slide.addText("Q4 COMP CHALLENGE", {
    x: MARGIN + 0.15, y: cardY + 0.06, w: cardW - 0.3, h: 0.22,
    fontSize: 8.5, fontFace: FONT.sub, color: C.red, bold: true, margin: 0, charSpacing: 1.5,
  });
  slide.addText(
    "Q4 2025 vs Q4 2024: revenue +4.8% but EPS \u20134.9% ($3.10 vs $3.26) and margin \u2013120 bps (8.5% vs 9.7%). Capacity grew faster than unit revenue recovered; holiday demand that supercharged Q4 2024 had partially normalized. This is a comps problem, not structural deterioration.",
    {
      x: MARGIN + 0.15, y: cardY + 0.32, w: cardW - 0.3, h: 1.15,
      fontSize: 9, fontFace: FONT.body, color: C.charcoal, margin: 0, lineSpacingMultiple: 1.3,
    }
  );

  // Card 2: Guidance track record
  slide.addShape(pres.shapes.RECTANGLE, {
    x: MARGIN + cardW + 0.2, y: cardY, w: cardW, h: 1.55,
    fill: { color: C.white }, shadow: makeShadow(),
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: MARGIN + cardW + 0.2, y: cardY, w: 0.06, h: 1.55,
    fill: { color: C.green },
  });
  slide.addText("GUIDANCE TRACK RECORD", {
    x: MARGIN + cardW + 0.35, y: cardY + 0.06, w: cardW - 0.3, h: 0.22,
    fontSize: 8.5, fontFace: FONT.sub, color: C.green, bold: true, margin: 0, charSpacing: 1.5,
  });
  slide.addText(
    "Initial 2025 guide: $11.50\u2013$13.50. Mid-year cut (domestic PRASM softness + macro): $9.00\u2013$11.00. Actual: $10.62 \u2014 at/near top of revised range, above consensus. This conservative re-guidance + outperformance pattern has been consistent since 2022. The $12\u201314 range for 2026 was issued in a similar cautious window.",
    {
      x: MARGIN + cardW + 0.35, y: cardY + 0.32, w: cardW - 0.3, h: 1.15,
      fontSize: 9, fontFace: FONT.body, color: C.charcoal, margin: 0, lineSpacingMultiple: 1.3,
    }
  );

  // GAAP vs Adjusted note
  slide.addText(
    "Note: Q2 GAAP vs. adjusted gap was the largest ($447M in special charges, primarily fleet write-offs for early retirements). Full-year GAAP EPS of $10.20 vs. $10.62 adjusted.",
    {
      x: MARGIN, y: 4.55, w: CONTENT_W, h: 0.35,
      fontSize: 8, fontFace: FONT.body, color: C.slate, italic: true, margin: 0,
    }
  );
}

// ============================================================
// SLIDE 5 — UNIT ECONOMICS DEEP DIVE
// ============================================================
{
  const slide = addLightSlide();
  addLightHeader(slide, "Unit Economics", "TRASM decline was a 2025 industry phenomenon, not structural demand deterioration");
  addFooter(slide, 5, TOTAL_SLIDES);

  // Annual table
  slide.addText("ANNUAL TRENDS", {
    x: MARGIN, y: 0.95, w: 3, h: 0.22,
    fontSize: 8.5, fontFace: FONT.sub, color: C.navy, bold: true, margin: 0, charSpacing: 2,
  });
  const annualRows = [
    ["Year", "TRASM", "PRASM", "Yield", "CASM", "CASM-ex", "Fuel/Gal", "Load Factor"],
    ["FY2023", "18.44\u00A2", "\u2014", "20.07\u00A2", "16.99\u00A2", "12.03\u00A2", "$3.01", "83.9%"],
    ["FY2024", "18.34\u00A2", "16.66\u00A2", "20.05\u00A2", "16.70\u00A2", "12.58\u00A2", "$2.65", "83.1%"],
    ["FY2025", "17.88\u00A2", "16.18\u00A2", "19.67\u00A2", "16.46\u00A2", "12.64\u00A2", "$2.44", "82.2%"],
  ];
  addStyledTable(slide, annualRows, MARGIN, 1.18, CONTENT_W * 0.55, { rowH: 0.26 });

  // Quarterly table
  slide.addText("QUARTERLY FY2025", {
    x: MARGIN + CONTENT_W * 0.58, y: 0.95, w: 3, h: 0.22,
    fontSize: 8.5, fontFace: FONT.sub, color: C.navy, bold: true, margin: 0, charSpacing: 2,
  });
  const qtrRows = [
    ["Qtr", "TRASM", "PRASM", "CASM-ex", "Fuel/Gal", "LF"],
    ["Q1", "17.58\u00A2", "15.78\u00A2", "13.17\u00A2", "$2.53", "79.2%"],
    ["Q2", "18.06\u00A2", "16.40\u00A2", "12.36\u00A2", "$2.34", "83.1%"],
    ["Q3", "17.42\u00A2", "15.80\u00A2", "12.15\u00A2", "$2.43", "84.4%"],
    ["Q4", "18.47\u00A2", "16.71\u00A2", "12.94\u00A2", "$2.49", "81.9%"],
  ];
  addStyledTable(slide, qtrRows, MARGIN + CONTENT_W * 0.58, 1.18, CONTENT_W * 0.42, { rowH: 0.26 });

  // Insight card
  const insY = 2.65;
  slide.addShape(pres.shapes.RECTANGLE, {
    x: MARGIN, y: insY, w: CONTENT_W, h: 2.2,
    fill: { color: C.white }, shadow: makeShadow(),
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: MARGIN, y: insY, w: 0.06, h: 2.2,
    fill: { color: C.accent },
  });

  // 3 key points in columns
  const colW3 = (CONTENT_W - 0.4) / 3;
  const pts = [
    { title: "TRASM \u20132.5% YoY", color: C.red, body: "Industry overcapacity, not demand weakness. Passengers hit record 181M. UAL responded by cutting 4 pts of domestic capacity mid-year and retiring 21 aircraft ahead of plan." },
    { title: "CASM-ex +0.5%", color: C.amber, body: "Fuel saved ~$360M (price/gal: $2.65\u2192$2.44). Labor consumed +$969M. Net: costs grew slightly faster than revenue. Labor headwinds are known and peaking; fuel efficiency compounds with new aircraft." },
    { title: "2026 THESIS", color: C.green, body: "TRASM inflects up as industry capacity normalizes. CASM-ex stabilizes as deliveries bring efficiency. The gap between revenue and cost lines widens into margin expansion. That is the entire 2026 thesis." },
  ];
  pts.forEach((p, i) => {
    const px = MARGIN + 0.2 + i * (colW3 + 0.05);
    slide.addText(p.title, {
      x: px, y: insY + 0.12, w: colW3 - 0.1, h: 0.25,
      fontSize: 9, fontFace: FONT.sub, color: p.color, bold: true, margin: 0, charSpacing: 1,
    });
    slide.addText(p.body, {
      x: px, y: insY + 0.42, w: colW3 - 0.1, h: 1.7,
      fontSize: 8.5, fontFace: FONT.body, color: C.charcoal, margin: 0, lineSpacingMultiple: 1.3,
    });
  });

  // Fuel sensitivity callout
  slide.addText(
    "FUEL SENSITIVITY: UAL carries no financial fuel hedges. Every $0.10/gal move = ~$466M annual P&L impact (4.663B gal consumed). A return to 2023\u2019s $3.01/gal = ~$2.7B headwind vs. ~$4.6B adj. pre-tax income.",
    {
      x: MARGIN, y: SLIDE_H - 0.5, w: CONTENT_W, h: 0.35,
      fontSize: 8, fontFace: FONT.body, color: C.red, italic: true, margin: 0, bold: true,
    }
  );
}

// ============================================================
// SLIDE 6 — OPERATING COST BRIDGE
// ============================================================
{
  const slide = addLightSlide();
  addLightHeader(slide, "Operating Expense Composition", "OpEx +4.6% vs. revenue +3.5% \u2014 the composition tells different stories about durability");
  addFooter(slide, 6, TOTAL_SLIDES);

  const rows = [
    ["Line Item", "FY2025", "FY2024", "YoY \u0394", "YoY %", "% of Total"],
    ["Salaries & related", "$17.65B", "$16.68B", "+$969M", "+5.8%", "32.5%"],
    ["Aircraft fuel", "$11.40B", "$11.76B", "\u2013$360M", "\u20133.1%", "21.0%"],
    ["Landing fees & rent", "$3.85B", "$3.44B", "+$412M", "+12.0%", "7.1%"],
    ["Maintenance", "$3.29B", "$3.06B", "+$231M", "+7.5%", "6.1%"],
    ["Depreciation", "$2.94B", "$2.93B", "+$11M", "+0.4%", "5.4%"],
    ["Regional capacity", "$2.69B", "$2.52B", "+$177M", "+7.0%", "5.0%"],
    ["Distribution", "$2.11B", "$2.23B", "\u2013$122M", "\u20135.5%", "3.9%"],
    ["Other", "$9.92B", "$9.05B", "+$866M", "+9.6%", "18.2%"],
    ["Total OpEx", "$54.36B", "$51.97B", "+$2.39B", "+4.6%", "100%"],
  ];
  addStyledTable(slide, rows, MARGIN, 1.0, CONTENT_W, { rowH: 0.24 });

  // Three callout cards below
  const cy = 3.55;
  const cw = (CONTENT_W - 0.3) / 3;
  const cards = [
    {
      title: "LABOR: KNOWN & BOUNDED",
      color: C.navy,
      text: "+$2.9B in two years. Pilot contract ($10B value, 34\u201340% raises) front-loaded \u2014 steps down by 2026\u201327. FA deal removes multi-year overhang. Gauge-up: labor cost per seat served rising less than cost per employee.",
    },
    {
      title: "LANDING FEES: GOOD COST",
      color: C.accent,
      text: "+12% is almost entirely self-inflicted: $2.7B EWR Terminal A, ORD T2 renovation, DEN gate expansion. Correlation between this line and NPS improvement is not coincidental. Deferred premium spending in the OpEx line.",
    },
    {
      title: "DISTRIBUTION: HIDDEN TAILWIND",
      color: C.green,
      text: "\u20135.5% as booking mix shifts to direct channels. App handles 85% of check-ins. Direct bookings capture more ancillary revenue + better loyalty data. Engagement metrics accelerating, not plateauing.",
    },
  ];
  cards.forEach((c, i) => {
    const cx = MARGIN + i * (cw + 0.15);
    slide.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: cy, w: cw, h: 1.7,
      fill: { color: C.white }, shadow: makeShadow(),
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: cy, w: cw, h: 0.04,
      fill: { color: c.color },
    });
    slide.addText(c.title, {
      x: cx + 0.1, y: cy + 0.12, w: cw - 0.2, h: 0.2,
      fontSize: 8, fontFace: FONT.sub, color: c.color, bold: true, margin: 0, charSpacing: 1,
    });
    slide.addText(c.text, {
      x: cx + 0.1, y: cy + 0.38, w: cw - 0.2, h: 1.25,
      fontSize: 8.5, fontFace: FONT.body, color: C.charcoal, margin: 0, lineSpacingMultiple: 1.25,
    });
  });
}

// ============================================================
// SLIDE 7 — REGIONAL REVENUE MIX
// ============================================================
{
  const slide = addLightSlide();
  addLightHeader(slide, "Revenue by Region", "A two-speed airline: domestic fare softness vs. international structural strength");
  addFooter(slide, 7, TOTAL_SLIDES);

  // Annual table
  slide.addText("ANNUAL PASSENGER REVENUE", {
    x: MARGIN, y: 0.95, w: 4, h: 0.2,
    fontSize: 8.5, fontFace: FONT.sub, color: C.navy, bold: true, margin: 0, charSpacing: 1.5,
  });
  const annRows = [
    ["Region", "FY2023", "FY2024", "FY2025", "3-Yr CAGR"],
    ["Domestic", "$31.9B", "$28.5B", "~$31.6B", "\u20130.4%"],
    ["Atlantic \u2014 Europe", "$8.1B", "$8.1B", "~$10.1B", "+11.9%"],
    ["Atlantic \u2014 MEIA", "$1.0B", "$1.1B", "~$1.3B", "+14.0%"],
    ["Pacific", "$4.4B", "$5.2B", "~$6.9B", "+24.5%"],
    ["Latin America", "$3.7B", "$4.8B", "~$5.5B", "+22.8%"],
  ];
  addStyledTable(slide, annRows, MARGIN, 1.17, CONTENT_W * 0.48, { rowH: 0.24 });

  // Q4 detail
  slide.addText("Q4 2025 DETAIL", {
    x: MARGIN + CONTENT_W * 0.52, y: 0.95, w: 4, h: 0.2,
    fontSize: 8.5, fontFace: FONT.sub, color: C.navy, bold: true, margin: 0, charSpacing: 1.5,
  });
  const q4Rows = [
    ["Region", "Q4 Rev", "Q4 YoY", "Q4 PRASM \u0394", "Q4 ASM \u0394"],
    ["Domestic", "$8,301M", "+2.0%", "\u20131.9%", "+4.0%"],
    ["Europe", "$2,274M", "+8.7%", "+1.3%", "+7.3%"],
    ["MEIA", "$410M", "+58.7%", "+1.7%", "+56.0%"],
    ["Pacific", "$1,626M", "+10.1%", "+4.2%", "+5.7%"],
    ["Latin Amer.", "$1,317M", "+0.5%", "\u20137.6%", "+8.7%"],
  ];
  addStyledTable(slide, q4Rows, MARGIN + CONTENT_W * 0.52, 1.17, CONTENT_W * 0.48, { rowH: 0.24 });

  // Insight cards
  const iy = 2.85;
  const iw = (CONTENT_W - 0.15) / 2;

  // Card 1
  slide.addShape(pres.shapes.RECTANGLE, {
    x: MARGIN, y: iy, w: iw, h: 2.25,
    fill: { color: C.white }, shadow: makeShadow(),
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: MARGIN, y: iy, w: 0.06, h: 2.25,
    fill: { color: C.accent },
  });
  slide.addText("MEIA: EMERGING GROWTH VECTOR", {
    x: MARGIN + 0.15, y: iy + 0.08, w: iw - 0.3, h: 0.2,
    fontSize: 8.5, fontFace: FONT.sub, color: C.accent, bold: true, margin: 0, charSpacing: 1,
  });
  slide.addText(
    "+58.7% Q4 revenue with +56% capacity deployment \u2014 all new route ramps (EWR\u2013Dubai, EWR\u2013Delhi, SFO\u2013Bangalore). India is a structural opportunity: world\u2019s most populous country with rapidly growing outbound travel. UAL is the most aggressive U.S. airline in India/Gulf. A321XLR in 2026 gives more tools to expand beyond gateway hubs.\n\nPACIFIC: +24.5% 3-Yr CAGR. \u201COnly U.S. carrier\u201D monopoly routes to Bangkok, Adelaide, Ho Chi Minh City carry premium business passengers with no alternative.",
    {
      x: MARGIN + 0.15, y: iy + 0.33, w: iw - 0.3, h: 1.85,
      fontSize: 8.5, fontFace: FONT.body, color: C.charcoal, margin: 0, lineSpacingMultiple: 1.25,
    }
  );

  // Card 2
  slide.addShape(pres.shapes.RECTANGLE, {
    x: MARGIN + iw + 0.15, y: iy, w: iw, h: 2.25,
    fill: { color: C.white }, shadow: makeShadow(),
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: MARGIN + iw + 0.15, y: iy, w: 0.06, h: 2.25,
    fill: { color: C.amber },
  });
  slide.addText("DOMESTIC: THE VARIABLE THAT MATTERS", {
    x: MARGIN + iw + 0.3, y: iy + 0.08, w: iw - 0.3, h: 0.2,
    fontSize: 8.5, fontFace: FONT.sub, color: C.amber, bold: true, margin: 0, charSpacing: 1,
  });
  slide.addText(
    "PRASM negative in every quarter of 2025 (worst: \u20137.0% in Q2). Industry added seats faster than domestic fares could absorb. But domestic is 54%+ of pax revenue \u2014 even modest PRASM improvement on that base compounds quickly into earnings leverage.\n\nLATIN AMERICA: Genuine softness (PRASM \u201313.5% Q3, \u20137.6% Q4). Brazilian carrier competition + LatAm macro headwinds. Only ~10% of pax revenue \u2014 manageable drag, not a thesis-breaker.",
    {
      x: MARGIN + iw + 0.3, y: iy + 0.33, w: iw - 0.3, h: 1.85,
      fontSize: 8.5, fontFace: FONT.body, color: C.charcoal, margin: 0, lineSpacingMultiple: 1.25,
    }
  );
}

// ============================================================
// SLIDE 8 — REVENUE QUALITY: PREMIUM & LOYALTY
// ============================================================
{
  const slide = addLightSlide();
  addLightHeader(slide, "Revenue Quality \u2014 Premium & Loyalty", "The quality-of-earnings story that gets the least credit in bear-case analyses");
  addFooter(slide, 8, TOTAL_SLIDES);

  // Top stat cards row
  const statY = 1.05;
  const sw = (CONTENT_W - 0.6) / 5;
  const stats = [
    { val: "+11%", lbl: "Premium Cabin\nFY2025 YoY" },
    { val: "+9%", lbl: "MileagePlus\nFY2025 YoY" },
    { val: "+12%", lbl: "Chase Card Spend\nFY2025 YoY" },
    { val: "27.4M", lbl: "Premium Seats\nFlown FY2025" },
    { val: "130M+", lbl: "MileagePlus\nMembers" },
  ];
  stats.forEach((s, i) => {
    addLightStatCard(slide, MARGIN + i * (sw + 0.15), statY, sw, 0.85, s.val, s.lbl, {
      accentColor: i < 3 ? C.accent : C.green,
      valueFontSize: 22,
    });
  });

  // Two-column detail
  const detY = 2.15;
  const colW = (CONTENT_W - 0.2) / 2;

  // Left: Premium
  slide.addShape(pres.shapes.RECTANGLE, {
    x: MARGIN, y: detY, w: colW, h: 2.95,
    fill: { color: C.white }, shadow: makeShadow(),
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: MARGIN, y: detY, w: colW, h: 0.04,
    fill: { color: C.accent },
  });
  slide.addText("PREMIUM MIX SHIFT", {
    x: MARGIN + 0.12, y: detY + 0.12, w: colW - 0.24, h: 0.2,
    fontSize: 9, fontFace: FONT.sub, color: C.navy, bold: true, margin: 0, charSpacing: 1,
  });
  slide.addText([
    { text: "Premium +11% vs. Basic Economy +5%", options: { bold: true, breakLine: true, fontSize: 9, fontFace: FONT.body, color: C.charcoal } },
    { text: " \u2014 deliberate mix shift toward higher-yield passengers less likely to trade down in a downturn.\n\n", options: { fontSize: 9, fontFace: FONT.body, color: C.charcoal, breakLine: true } },
    { text: "Signature Interior rollout: ", options: { bold: true, fontSize: 9, fontFace: FONT.body, color: C.charcoal } },
    { text: "119 aircraft (68% of NB fleet) with new seatback screens + larger bins. NPS +10 pts on equipped aircraft.\n\n", options: { fontSize: 9, fontFace: FONT.body, color: C.charcoal, breakLine: true } },
    { text: "Premium seats per N. American departure +40% since 2021", options: { bold: true, fontSize: 9, fontFace: FONT.body, color: C.charcoal } },
    { text: " via gauge-up: retiring 50-seat RJs, replacing with mainline aircraft carrying more premium rows.", options: { fontSize: 9, fontFace: FONT.body, color: C.charcoal } },
  ], {
    x: MARGIN + 0.12, y: detY + 0.38, w: colW - 0.24, h: 2.45,
    margin: 0, lineSpacingMultiple: 1.25,
  });

  // Right: Loyalty
  slide.addShape(pres.shapes.RECTANGLE, {
    x: MARGIN + colW + 0.2, y: detY, w: colW, h: 2.95,
    fill: { color: C.white }, shadow: makeShadow(),
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: MARGIN + colW + 0.2, y: detY, w: colW, h: 0.04,
    fill: { color: C.green },
  });
  slide.addText("LOYALTY & MILEAGEPLUS", {
    x: MARGIN + colW + 0.32, y: detY + 0.12, w: colW - 0.24, h: 0.2,
    fontSize: 9, fontFace: FONT.sub, color: C.navy, bold: true, margin: 0, charSpacing: 1,
  });
  slide.addText([
    { text: "1M+ new Chase cards/yr ", options: { bold: true, fontSize: 9, fontFace: FONT.body, color: C.charcoal } },
    { text: "for 3rd consecutive year. Card spend +12% YoY on large base \u2014 a compounding flywheel.\n\n", options: { fontSize: 9, fontFace: FONT.body, color: C.charcoal, breakLine: true } },
    { text: "2020 MileagePlus valuation: ~$22B ", options: { bold: true, fontSize: 9, fontFace: FONT.body, color: C.charcoal } },
    { text: "(used as collateral for $6.8B COVID financing, now fully repaid). Program has grown since. UAL market cap: ~$28B. Market ascribes minimal standalone value above embedded.\n\n", options: { fontSize: 9, fontFace: FONT.body, color: C.charcoal, breakLine: true } },
    { text: "April 2026: ", options: { bold: true, fontSize: 9, fontFace: FONT.body, color: C.charcoal } },
    { text: "New Chase card benefits: 2\u00D7 miles for cardholders, 10\u201315% discount on award redemptions.\n\n", options: { fontSize: 9, fontFace: FONT.body, color: C.charcoal, breakLine: true } },
    { text: "CEO target: 2\u00D7 MileagePlus profits by 2030.", options: { bold: true, fontSize: 9, fontFace: FONT.body, color: C.accent } },
  ], {
    x: MARGIN + colW + 0.32, y: detY + 0.38, w: colW - 0.24, h: 2.45,
    margin: 0, lineSpacingMultiple: 1.25,
  });
}

// ============================================================
// SLIDE 9 — FLEET MODERNIZATION
// ============================================================
{
  const slide = addLightSlide();
  addLightHeader(slide, "Fleet Modernization & United Next", "630+ new aircraft by 2034 \u2014 the largest fleet transformation in UAL history");
  addFooter(slide, 9, TOTAL_SLIDES);

  // Fleet growth stats
  const gY = 1.05;
  const gw = (CONTENT_W - 0.45) / 4;
  const gStats = [
    { val: "1,490", lbl: "Fleet Dec 2025" },
    { val: "~1,610", lbl: "Fleet Dec 2026E" },
    { val: "~120", lbl: "2026 Deliveries\n(Largest since 1988)" },
    { val: "174", lbl: "Avg Seats/Departure\n(vs. 151 in 2019)" },
  ];
  gStats.forEach((s, i) => {
    addLightStatCard(slide, MARGIN + i * (gw + 0.15), gY, gw, 0.85, s.val, s.lbl, {
      accentColor: C.accent,
      valueFontSize: 22,
    });
  });

  // Order book table
  slide.addText("ORDER BOOK THROUGH 2034", {
    x: MARGIN, y: 2.1, w: 4, h: 0.2,
    fontSize: 8.5, fontFace: FONT.sub, color: C.navy, bold: true, margin: 0, charSpacing: 1.5,
  });
  const fleetRows = [
    ["Aircraft", "Role", "Remaining / Target"],
    ["Boeing 787-9 Elevated", "Long-haul premium intl", "47 next wave; 30+ by end 2027"],
    ["Boeing 737 MAX (8/9/10)", "Domestic / short-haul NB", "200+ remaining"],
    ["Airbus A321neo (standard)", "General domestic", "18+ remaining"],
    ["Airbus A321neo Coastliner", "Domestic transcon lie-flat", "50 ordered; 40 by Apr 2028"],
    ["Airbus A321XLR", "New transatlantic + S. Am.", "50 ordered; 25+ by 2028"],
    ["CRJ450 (SkyWest)", "Premium regional upgrade", "70 total; 50 by 2028"],
  ];
  addStyledTable(slide, fleetRows, MARGIN, 2.32, CONTENT_W * 0.6, { rowH: 0.22 });

  // Boeing risk + gauge-up callout
  const rX = MARGIN + CONTENT_W * 0.63;
  const rW = CONTENT_W * 0.37;
  slide.addShape(pres.shapes.RECTANGLE, {
    x: rX, y: 2.32, w: rW, h: 1.15,
    fill: { color: C.white }, shadow: makeShadow(),
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: rX, y: 2.32, w: 0.06, h: 1.15,
    fill: { color: C.red },
  });
  slide.addText("BOEING DELIVERY RISK", {
    x: rX + 0.12, y: 2.37, w: rW - 0.24, h: 0.18,
    fontSize: 8, fontFace: FONT.sub, color: C.red, bold: true, margin: 0, charSpacing: 1,
  });
  slide.addText(
    "FAA cap of 38 MAX/month post-door plug incident. MAX 10 certification pending (200 aircraft). Each month of delay = ~$30\u201340M deferred efficiency benefits. A321neo/XLR orders explicitly hedge Boeing timing uncertainty.",
    {
      x: rX + 0.12, y: 2.58, w: rW - 0.24, h: 0.85,
      fontSize: 8.5, fontFace: FONT.body, color: C.charcoal, margin: 0, lineSpacingMultiple: 1.25,
    }
  );

  slide.addShape(pres.shapes.RECTANGLE, {
    x: rX, y: 3.55, w: rW, h: 0.95,
    fill: { color: C.white }, shadow: makeShadow(),
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: rX, y: 3.55, w: 0.06, h: 0.95,
    fill: { color: C.green },
  });
  slide.addText("GAUGE-UP ECONOMICS", {
    x: rX + 0.12, y: 3.6, w: rW - 0.24, h: 0.18,
    fontSize: 8, fontFace: FONT.sub, color: C.green, bold: true, margin: 0, charSpacing: 1,
  });
  slide.addText(
    "Each new 737 MAX / A321neo burns ~20% less fuel per seat and requires fewer crew per seat served. 50-seat RJs replaced by 150+ seat mainline \u2014 a labor productivity strategy disguised as fleet strategy.",
    {
      x: rX + 0.12, y: 3.81, w: rW - 0.24, h: 0.65,
      fontSize: 8.5, fontFace: FONT.body, color: C.charcoal, margin: 0, lineSpacingMultiple: 1.25,
    }
  );

  // Capex context
  slide.addText(
    "CAPEX: $5.9B in 2025 | <$8.0B guided 2026 | $7\u20139B/yr multi-year guidance as delivery ramp continues through 2027",
    {
      x: MARGIN, y: 4.7, w: CONTENT_W, h: 0.3,
      fontSize: 8.5, fontFace: FONT.body, color: C.slate, italic: true, bold: true, margin: 0,
    }
  );
}

// ============================================================
// SLIDE 10 — NEW PRODUCT ARCHITECTURE
// ============================================================
{
  const slide = addLightSlide();
  addLightHeader(slide, "New Product Architecture: 2026\u20132028", "Four new aircraft products that redefine UAL\u2019s competitive position");
  addFooter(slide, 10, TOTAL_SLIDES);

  const cardW = (CONTENT_W - 0.3) / 2;
  const cardH = 2.0;
  const prodCards = [
    {
      x: MARGIN, y: 1.0,
      color: C.navy,
      title: "787-9 ELEVATED",
      sub: "Inaugural: SFO\u2013Singapore, Apr 22, 2026",
      body: "222 seats / 44.6% premium \u2014 highest among U.S. widebodies\n\nPolaris Studio: 8 suites, 25% larger than standard Polaris, 27\u201D 4K OLED (largest on any U.S. carrier), privacy door, double-bed center config\n\nPolaris: 56 lie-flat | PPL: 35 | E+: 39 | Econ: 84\nAll cabins: 4K OLED, Bluetooth, wireless charging\n\nTarget: 30+ aircraft by end 2027",
    },
    {
      x: MARGIN + cardW + 0.15, y: 1.0,
      color: C.accent,
      title: "A321neo COASTLINER",
      sub: "Launch summer 2026: SFO/LAX \u2194 EWR/JFK",
      body: "161 seats / First lie-flat on any U.S. domestic narrowbody\n\nPolaris: 20 (1-1 lie-flat) \u2014 directly targets Delta One and JetBlue Mint on transcon routes where UAL currently has no competitive answer\n\nPremium Plus: 12 (2-2) \u2014 first PPL on domestic NB\nEconomy: 129 with rear snack bar\n\n50 ordered; 40 by April 2028",
    },
    {
      x: MARGIN, y: 1.0 + cardH + 0.15,
      color: C.green,
      title: "A321XLR",
      sub: "Launch summer 2026: new European & S. American routes",
      body: "~150 seats / 4,700 nm range unlocks entirely new markets\n\nPolaris: 20 all-aisle-access lie-flat with privacy door\nPPL: 12 | Economy: ~118\n\nRange enables direct EWR to Bari, Split, Glasgow, Santiago de Compostela, and deeper S. American cities \u2014 routes neither 757-200 nor standard A321 can serve. Not incremental \u2014 new addressable market.\n\n50 ordered; 25+ by 2028",
    },
    {
      x: MARGIN + cardW + 0.15, y: 1.0 + cardH + 0.15,
      color: C.amber,
      title: "CRJ450 REGIONAL PREMIUM",
      sub: "Launch fall 2026: DEN + ORD hubs (SkyWest)",
      body: "41 seats / Credible premium product on regional routes\n\nUnited First: 7 seats, 37\u201D pitch, overhead luggage closet (industry first on regional). Economy Plus: 16. Economy: 18.\n\nFree Starlink Wi-Fi for MileagePlus members.\n\nConverts regional from \u201Cmiddle-seat-blocked + snack box\u201D to real premium. Critical for business travelers connecting through DEN/ORD.\n\n70 fleet; 50 by 2028",
    },
  ];

  prodCards.forEach((c) => {
    slide.addShape(pres.shapes.RECTANGLE, {
      x: c.x, y: c.y, w: cardW, h: cardH,
      fill: { color: C.white }, shadow: makeShadow(),
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x: c.x, y: c.y, w: cardW, h: 0.04,
      fill: { color: c.color },
    });
    slide.addText(c.title, {
      x: c.x + 0.1, y: c.y + 0.1, w: cardW * 0.4, h: 0.2,
      fontSize: 9, fontFace: FONT.head, color: c.color, bold: true, margin: 0,
    });
    slide.addText(c.sub, {
      x: c.x + cardW * 0.4 + 0.1, y: c.y + 0.1, w: cardW * 0.55, h: 0.2,
      fontSize: 7.5, fontFace: FONT.body, color: C.slate, italic: true, align: "right", margin: 0,
    });
    slide.addText(c.body, {
      x: c.x + 0.1, y: c.y + 0.36, w: cardW - 0.2, h: cardH - 0.46,
      fontSize: 8, fontFace: FONT.body, color: C.charcoal, margin: 0, lineSpacingMultiple: 1.2,
    });
  });
}

// ============================================================
// SLIDE 11 — STARLINK + KINECTIVE MEDIA
// ============================================================
{
  const slide = addDarkSlide();
  addFooter(slide, 11, TOTAL_SLIDES, true);

  slide.addText("STARLINK + KINECTIVE MEDIA", {
    x: MARGIN, y: 0.3, w: CONTENT_W, h: 0.45,
    fontSize: 22, fontFace: FONT.head, color: C.white, bold: true, margin: 0, charSpacing: 3,
  });
  slide.addText("Connectivity as infrastructure for advertising and loyalty data \u2014 a nascent high-margin revenue stream", {
    x: MARGIN, y: 0.75, w: CONTENT_W * 0.8, h: 0.3,
    fontSize: 10, fontFace: FONT.body, color: C.ice, margin: 0,
  });

  // Left: Starlink timeline
  const leftW = CONTENT_W * 0.42;
  slide.addShape(pres.shapes.RECTANGLE, {
    x: MARGIN, y: 1.2, w: leftW, h: 3.6,
    fill: { color: C.midNavy },
    shadow: makeShadow(),
  });
  slide.addText("STARLINK ROLLOUT", {
    x: MARGIN + 0.15, y: 1.3, w: leftW - 0.3, h: 0.22,
    fontSize: 9, fontFace: FONT.sub, color: C.accent, bold: true, margin: 0, charSpacing: 2,
  });
  const timeline = [
    "May 2025: First E175 regional with Starlink",
    "Oct 15, 2025: First mainline commercial flight (EWR)",
    "Feb 2026: 300+ mainline aircraft equipped",
    "End 2026: 800+ total aircraft (~25%+ of departures)",
    "End 2027: Full fleet completion",
    "Install pace: 50+ aircraft per month",
    "Free for all MileagePlus members",
  ];
  const tlItems = timeline.map((t, i) => ({
    text: t,
    options: {
      bullet: true,
      breakLine: i < timeline.length - 1,
      fontSize: 9,
      fontFace: FONT.body,
      color: C.ice,
      paraSpaceAfter: 6,
    },
  }));
  slide.addText(tlItems, {
    x: MARGIN + 0.15, y: 1.6, w: leftW - 0.3, h: 3.0,
    margin: 0,
  });

  // Right: Kinective Media
  const rightX = MARGIN + leftW + 0.15;
  const rightW = CONTENT_W - leftW - 0.15;
  slide.addShape(pres.shapes.RECTANGLE, {
    x: rightX, y: 1.2, w: rightW, h: 3.6,
    fill: { color: C.midNavy },
    shadow: makeShadow(),
  });
  slide.addText("KINECTIVE MEDIA NETWORK", {
    x: rightX + 0.15, y: 1.3, w: rightW - 0.3, h: 0.22,
    fontSize: 9, fontFace: FONT.sub, color: C.gold, bold: true, margin: 0, charSpacing: 2,
  });
  slide.addText([
    { text: "Three assets no other airline possesses simultaneously:\n\n", options: { fontSize: 9.5, fontFace: FONT.body, color: C.ice, bold: true, breakLine: true } },
    { text: "227,000+ seatback screens", options: { fontSize: 9, fontFace: FONT.body, color: C.gold, bold: true } },
    { text: " across a captive audience\n\n", options: { fontSize: 9, fontFace: FONT.body, color: C.ice, breakLine: true } },
    { text: "108M unique annual flyers", options: { fontSize: 9, fontFace: FONT.body, color: C.gold, bold: true } },
    { text: " with known demographics\n\n", options: { fontSize: 9, fontFace: FONT.body, color: C.ice, breakLine: true } },
    { text: "130M MileagePlus members", options: { fontSize: 9, fontFace: FONT.body, color: C.gold, bold: true } },
    { text: " with deep first-party transaction data \u2014 miles, redemptions, card spend, hotel/car bookings\n\n", options: { fontSize: 9, fontFace: FONT.body, color: C.ice, breakLine: true } },
    { text: "Better audience profile than most digital ad platforms", options: { fontSize: 9, fontFace: FONT.body, color: C.ice, bold: true } },
    { text: " \u2014 behavioral + purchase-intent data, not just demographics.\n\n", options: { fontSize: 9, fontFace: FONT.body, color: C.ice, breakLine: true } },
    { text: "CFO: Kinective \"will really start to accelerate in \u201926 and beyond.\" ", options: { fontSize: 9, fontFace: FONT.body, color: C.accent, italic: true } },
    { text: "Media margins dramatically exceed seat revenue margins. The market currently values this business at approximately zero.", options: { fontSize: 9, fontFace: FONT.body, color: C.ice } },
  ], {
    x: rightX + 0.15, y: 1.6, w: rightW - 0.3, h: 3.0,
    margin: 0, lineSpacingMultiple: 1.15,
  });

  // NPS callout at bottom
  slide.addText(
    "NPS IMPACT: Signature Interior aircraft score +10 NPS pts vs. planes they replace. For a 40\u00D7/yr business traveler, that translates to retention and card spend.",
    {
      x: MARGIN, y: SLIDE_H - 0.55, w: CONTENT_W, h: 0.3,
      fontSize: 8, fontFace: FONT.body, color: C.lightGray, italic: true, margin: 0,
    }
  );
}

// ============================================================
// SLIDE 12 — OPERATIONAL PERFORMANCE
// ============================================================
{
  const slide = addLightSlide();
  addLightHeader(slide, "Operational Performance", "Record passengers, record on-time, record NPS \u2014 and the demand signals that matter most for 2026");
  addFooter(slide, 12, TOTAL_SLIDES);

  // Stat cards row
  const opsY = 1.05;
  const opsW = (CONTENT_W - 0.6) / 5;
  const opsStats = [
    { val: "181.1M", lbl: "Passengers\nCompany Record" },
    { val: "#2", lbl: "D:14 On-Time\nIndustry Rank" },
    { val: "303", lbl: "Daily Widebody\nDepartures (Record)" },
    { val: "1M+", lbl: "ConnectionSaver\nRescues (+42% YoY)" },
    { val: "85%", lbl: "Digital Check-In\nRate (Q1 Record)" },
  ];
  opsStats.forEach((s, i) => {
    addLightStatCard(slide, MARGIN + i * (opsW + 0.15), opsY, opsW, 0.85, s.val, s.lbl, {
      accentColor: i === 0 ? C.green : C.accent,
      valueFontSize: 20,
    });
  });

  // Key milestones card
  const mY = 2.15;
  slide.addShape(pres.shapes.RECTANGLE, {
    x: MARGIN, y: mY, w: CONTENT_W, h: 1.35,
    fill: { color: C.white }, shadow: makeShadow(),
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: MARGIN, y: mY, w: CONTENT_W, h: 0.04,
    fill: { color: C.green },
  });
  slide.addText("FORWARD-LOOKING DEMAND SIGNALS", {
    x: MARGIN + 0.15, y: mY + 0.1, w: 4, h: 0.2,
    fontSize: 9, fontFace: FONT.sub, color: C.green, bold: true, margin: 0, charSpacing: 1.5,
  });
  slide.addText([
    { text: "Week ending Jan 4, 2026: ", options: { bold: true, fontSize: 9.5, fontFace: FONT.body, color: C.charcoal } },
    { text: "Highest flown revenue week in UAL history.\n", options: { fontSize: 9.5, fontFace: FONT.body, color: C.charcoal, breakLine: true } },
    { text: "Week ending Jan 11, 2026: ", options: { bold: true, fontSize: 9.5, fontFace: FONT.body, color: C.charcoal } },
    { text: "Simultaneously the highest ticketing week AND the highest business sales week ever recorded.\n\n", options: { fontSize: 9.5, fontFace: FONT.body, color: C.charcoal, breakLine: true } },
    { text: "Ticketing data is forward-looking \u2014 it reflects purchases for future travel. Record business bookings entering 2026 is the single strongest real-time indicator that 2025\u2019s domestic PRASM softness was temporary, not structural. Business travel leads leisure in recovery cycles.", options: { fontSize: 9.5, fontFace: FONT.body, color: C.slate } },
  ], {
    x: MARGIN + 0.15, y: mY + 0.35, w: CONTENT_W - 0.3, h: 0.9,
    margin: 0, lineSpacingMultiple: 1.3,
  });

  // Additional operational detail
  const adY = 3.7;
  const adW = (CONTENT_W - 0.3) / 3;
  const opCards = [
    { title: "COMPLETION", text: "Best system completion factor in company history. United Express achieved 134 days with zero cancellations.", color: C.accent },
    { title: "NPS", text: "Company-record NPS in Q4 2025. November 2025 = highest-ever monthly NPS. Signature Interior: +10 pts vs. old interiors.", color: C.navy },
    { title: "NETWORK", text: "13 new international destinations in 2025 (Nuuk, Ulaanbaatar, Faro, Palermo, Bangkok, Adelaide...) + 29 domestic/Canadian routes.", color: C.green },
  ];
  opCards.forEach((c, i) => {
    const cx = MARGIN + i * (adW + 0.15);
    slide.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: adY, w: adW, h: 1.45,
      fill: { color: C.white }, shadow: makeShadow(),
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: adY, w: 0.06, h: 1.45,
      fill: { color: c.color },
    });
    slide.addText(c.title, {
      x: cx + 0.15, y: adY + 0.08, w: adW - 0.3, h: 0.2,
      fontSize: 8, fontFace: FONT.sub, color: c.color, bold: true, margin: 0, charSpacing: 2,
    });
    slide.addText(c.text, {
      x: cx + 0.15, y: adY + 0.32, w: adW - 0.3, h: 1.05,
      fontSize: 8.5, fontFace: FONT.body, color: C.charcoal, margin: 0, lineSpacingMultiple: 1.25,
    });
  });
}

// ============================================================
// SLIDE 13 — BALANCE SHEET & DELEVERAGING
// ============================================================
{
  const slide = addLightSlide();
  addLightHeader(slide, "Balance Sheet & Deleveraging", "MileagePlus bond payoff completed \u2014 investment-grade path increasingly visible");
  addFooter(slide, 13, TOTAL_SLIDES);

  const rows = [
    ["Metric", "Dec 2023", "Dec 2024", "Mar 2025", "Jun 2025", "Sep 2025", "Dec 2025", "2026 Target"],
    ["Total Liquidity ($B)", "$16.1", "$17.4", "$18.3", "$18.6", "$16.3", "$15.2", "\u2014"],
    ["Total Debt + Leases ($B)", "$29.3", "$28.7", "$27.7", "$27.1", "$25.4", "$25.0", "\u2014"],
    ["Net Leverage", "2.9\u00D7", "2.4\u00D7", "2.0\u00D7", "2.0\u00D7", "2.1\u00D7", "2.2\u00D7", "<2.0\u00D7"],
    ["S&P Rating", "BB", "BB+", "BB+", "BB+", "BB+", "BB+", "IG"],
    ["Interest Expense ($B)", "\u2014", "$1.629", "\u2014", "\u2014", "\u2014", "$1.373", "<$1.2E"],
  ];
  addStyledTable(slide, rows, MARGIN, 1.0, CONTENT_W, { rowH: 0.26 });

  // MileagePlus payoff callout
  const mpY = 2.65;
  slide.addShape(pres.shapes.RECTANGLE, {
    x: MARGIN, y: mpY, w: CONTENT_W * 0.5, h: 1.5,
    fill: { color: C.white }, shadow: makeShadow(),
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: MARGIN, y: mpY, w: 0.06, h: 1.5,
    fill: { color: C.green },
  });
  slide.addText("MILEAGEPLUS BOND PAYOFF", {
    x: MARGIN + 0.15, y: mpY + 0.06, w: CONTENT_W * 0.5 - 0.3, h: 0.2,
    fontSize: 9, fontFace: FONT.sub, color: C.green, bold: true, margin: 0, charSpacing: 1.5,
  });
  slide.addText(
    "Q3 2025: Prepaid remaining $1.5B of MileagePlus secured bonds, fully retiring $6.8B in MileagePlus-backed COVID financing by July 2025. Frees the MileagePlus asset, eliminates collateral restrictions, signals management confidence. Interest expense \u201315.7% YoY ($1.629B \u2192 $1.373B) \u2014 flows directly to EPS.",
    {
      x: MARGIN + 0.15, y: mpY + 0.3, w: CONTENT_W * 0.5 - 0.3, h: 1.1,
      fontSize: 9, fontFace: FONT.body, color: C.charcoal, margin: 0, lineSpacingMultiple: 1.3,
    }
  );

  // IG path
  const igX = MARGIN + CONTENT_W * 0.53;
  const igW = CONTENT_W * 0.47;
  slide.addShape(pres.shapes.RECTANGLE, {
    x: igX, y: mpY, w: igW, h: 1.5,
    fill: { color: C.white }, shadow: makeShadow(),
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: igX, y: mpY, w: 0.06, h: 1.5,
    fill: { color: C.accent },
  });
  slide.addText("INVESTMENT-GRADE PATH", {
    x: igX + 0.15, y: mpY + 0.06, w: igW - 0.3, h: 0.2,
    fontSize: 9, fontFace: FONT.sub, color: C.accent, bold: true, margin: 0, charSpacing: 1.5,
  });
  slide.addText(
    "At ~$3\u20134B/yr deleveraging, UAL reaches <2.0\u00D7 by year-end 2026. S&P IG criteria for airlines: 2.0\u20132.5\u00D7 with earnings stability. Three consequences: (1) Lower aircraft financing cost (\u201350\u201375 bps on $25B = $125\u2013190M/yr); (2) IG mandates open larger investor universe; (3) Multiple re-rating from \u201CCOVID bankruptcy risk\u201D to \u201Cinvestment-quality enterprise.\u201D Delta already carries BBB\u2013.",
    {
      x: igX + 0.15, y: mpY + 0.3, w: igW - 0.3, h: 1.1,
      fontSize: 9, fontFace: FONT.body, color: C.charcoal, margin: 0, lineSpacingMultiple: 1.3,
    }
  );

  // Deleveraging pace
  slide.addText(
    "Adjusted net debt: $18.8B (Q1 2025) vs. $22.5B (Q1 2024). Deleveraging pace: ~$3\u20134B per year. The question is not whether IG happens \u2014 but whether to own the stock before or after the rating action.",
    {
      x: MARGIN, y: 4.35, w: CONTENT_W, h: 0.45,
      fontSize: 9, fontFace: FONT.body, color: C.navy, bold: true, italic: true, margin: 0, lineSpacingMultiple: 1.3,
    }
  );
}

// ============================================================
// SLIDE 14 — CAPITAL ALLOCATION
// ============================================================
{
  const slide = addLightSlide();
  addLightHeader(slide, "Cash Flow & Capital Allocation", "CapEx bulge through 2027 delivery ramp, then normalization to $3\u20134B sustained FCF");
  addFooter(slide, 14, TOTAL_SLIDES);

  // Cash flow table
  const cfRows = [
    ["Metric", "FY2023", "FY2024", "FY2025", "FY2026E"],
    ["Operating Cash Flow ($B)", "$6.9", "$9.4", "$8.4", "\u2014"],
    ["Net CapEx ($B)", "$7.9", "$5.6", "$5.9", "<$8.0"],
    ["Free Cash Flow ($B)", "($1.0)", "$3.4", "$2.7", "~$2.7"],
    ["Multi-Yr CapEx Guidance", "\u2014", "\u2014", "\u2014", "$7\u20139B/yr"],
  ];
  addStyledTable(slide, cfRows, MARGIN, 1.0, CONTENT_W * 0.55, { rowH: 0.26 });

  // Buyback table
  slide.addText("SHARE REPURCHASES", {
    x: MARGIN + CONTENT_W * 0.58, y: 0.95, w: 3, h: 0.2,
    fontSize: 8.5, fontFace: FONT.sub, color: C.navy, bold: true, margin: 0, charSpacing: 1.5,
  });
  const buyRows = [
    ["Year / Period", "Amount", "Note"],
    ["FY2024", "$81M", "Capital preservation"],
    ["FY2025", "$640M", "Breakout year"],
    ["Q1 2025", "$451M", "Concentrated"],
    ["Q3\u2013Q4 2025", "$48M", "Slowed; CapEx accel."],
    ["Authorization", "$1.5B", "Approved Oct 2024"],
    ["Diluted Shares", "328.5M", "\u2192 ~325M FY2026E"],
  ];
  addStyledTable(slide, buyRows, MARGIN + CONTENT_W * 0.58, 1.17, CONTENT_W * 0.42, { rowH: 0.22 });

  // Priority stack
  const psY = 2.6;
  slide.addShape(pres.shapes.RECTANGLE, {
    x: MARGIN, y: psY, w: CONTENT_W * 0.45, h: 1.6,
    fill: { color: C.white }, shadow: makeShadow(),
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: MARGIN, y: psY, w: CONTENT_W * 0.45, h: 0.04,
    fill: { color: C.navy },
  });
  slide.addText("CAPITAL ALLOCATION PRIORITY", {
    x: MARGIN + 0.12, y: psY + 0.1, w: CONTENT_W * 0.45 - 0.24, h: 0.2,
    fontSize: 9, fontFace: FONT.sub, color: C.navy, bold: true, margin: 0, charSpacing: 1,
  });
  const priorities = [
    { text: "1. Fleet CapEx", options: { bold: true, breakLine: true, fontSize: 9, fontFace: FONT.body, color: C.charcoal, paraSpaceAfter: 4 } },
    { text: "   Non-negotiable given firm orders", options: { breakLine: true, fontSize: 8.5, fontFace: FONT.body, color: C.slate, paraSpaceAfter: 8 } },
    { text: "2. Debt Reduction", options: { bold: true, breakLine: true, fontSize: 9, fontFace: FONT.body, color: C.charcoal, paraSpaceAfter: 4 } },
    { text: "   $3.7B total debt reduction in FY2025", options: { breakLine: true, fontSize: 8.5, fontFace: FONT.body, color: C.slate, paraSpaceAfter: 8 } },
    { text: "3. Share Buybacks", options: { bold: true, breakLine: true, fontSize: 9, fontFace: FONT.body, color: C.charcoal, paraSpaceAfter: 4 } },
    { text: "   $640M in 2025; likely higher in 2026", options: { breakLine: true, fontSize: 8.5, fontFace: FONT.body, color: C.slate, paraSpaceAfter: 8 } },
    { text: "4. No Dividend", options: { bold: true, breakLine: true, fontSize: 9, fontFace: FONT.body, color: C.charcoal, paraSpaceAfter: 4 } },
    { text: "   None planned near-term", options: { fontSize: 8.5, fontFace: FONT.body, color: C.slate } },
  ];
  slide.addText(priorities, {
    x: MARGIN + 0.12, y: psY + 0.35, w: CONTENT_W * 0.45 - 0.24, h: 1.2,
    margin: 0,
  });

  // FCF normalization thesis
  const fcfX = MARGIN + CONTENT_W * 0.48;
  const fcfW = CONTENT_W * 0.52;
  slide.addShape(pres.shapes.RECTANGLE, {
    x: fcfX, y: psY, w: fcfW, h: 1.6,
    fill: { color: C.white }, shadow: makeShadow(),
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: fcfX, y: psY, w: 0.06, h: 1.6,
    fill: { color: C.green },
  });
  slide.addText("POST-2027 FCF NORMALIZATION", {
    x: fcfX + 0.15, y: psY + 0.1, w: fcfW - 0.3, h: 0.2,
    fontSize: 9, fontFace: FONT.sub, color: C.green, bold: true, margin: 0, charSpacing: 1,
  });
  slide.addText(
    "FCF step-down ($3.4B \u2192 $2.7B) despite higher net income is a CapEx story. ~120 aircraft in 2026 at $50\u201370M each creates a bulge persisting through 2027.\n\nOnce backlog thins (~2028), CapEx normalizes to $5\u20136B maintenance-plus-growth vs. $8\u20139B operating CF = sustained $3\u20134B FCF.\n\nFCF/share: $3.5B on ~320M shares (2027\u201328E) \u2248 ~$11/share. At 5\u20138\u00D7 airline sector FCF multiples, that implies significant equity value vs. current price.",
    {
      x: fcfX + 0.15, y: psY + 0.35, w: fcfW - 0.3, h: 1.2,
      fontSize: 8.5, fontFace: FONT.body, color: C.charcoal, margin: 0, lineSpacingMultiple: 1.25,
    }
  );

  // TTM highlight
  slide.addText(
    "Q1 2025 TTM peak: Operating CF $10.3B, Free CF $5.0B \u2014 temporarily boosted by advance ticket sales and delivery timing, but illustrates underlying cash generation capacity.",
    {
      x: MARGIN, y: 4.4, w: CONTENT_W, h: 0.4,
      fontSize: 8.5, fontFace: FONT.body, color: C.slate, italic: true, margin: 0,
    }
  );
}

// ============================================================
// SLIDE 15 — SUSTAINABILITY
// ============================================================
{
  const slide = addLightSlide();
  addLightHeader(slide, "Sustainability & Climate", "Honest framing: absolute emissions rise with growth; intensity falls with fleet efficiency");
  addFooter(slide, 15, TOTAL_SLIDES);

  // Emissions table
  slide.addText("GHG EMISSIONS", {
    x: MARGIN, y: 0.95, w: 3, h: 0.2,
    fontSize: 8.5, fontFace: FONT.sub, color: C.navy, bold: true, margin: 0, charSpacing: 1.5,
  });
  const emRows = [
    ["Category", "FY2021", "FY2022", "FY2023", "FY2024", "\u03942023\u201324"],
    ["Scope 1 (M MT CO2e)", "21.4", "30.4", "36.6", "38.5", "+5.3%"],
    ["Scope 2 Market (K MT)", "161", "149", "144", "134", "\u20136.6%"],
    ["Scope 3 (M MT)", "12.2", "13.3", "12.7", "13.6", "+7.2%"],
    ["CO2e / M ASMs", "187.5", "176.2", "169.0", "167.3", "\u20131.0%"],
  ];
  addStyledTable(slide, emRows, MARGIN, 1.17, CONTENT_W * 0.55, { rowH: 0.24 });

  // SAF table
  slide.addText("SAF PROGRESS", {
    x: MARGIN + CONTENT_W * 0.58, y: 0.95, w: 3, h: 0.2,
    fontSize: 8.5, fontFace: FONT.sub, color: C.navy, bold: true, margin: 0, charSpacing: 1.5,
  });
  const safRows = [
    ["Year", "SAF Gal (M)", "GHG Reduction", "CO2e Avoided (MT)"],
    ["2021", "0.6", "83.4%", "5,953"],
    ["2022", "2.9", "88.2%", "29,362"],
    ["2023", "7.3", "82.4%", "68,370"],
    ["2024", "13.6", "83.5%", "126,174"],
  ];
  addStyledTable(slide, safRows, MARGIN + CONTENT_W * 0.58, 1.17, CONTENT_W * 0.42, { rowH: 0.24 });

  // Commitments and honest framing
  const cY = 2.6;
  const cW2 = (CONTENT_W - 0.15) / 2;

  slide.addShape(pres.shapes.RECTANGLE, {
    x: MARGIN, y: cY, w: cW2, h: 2.1,
    fill: { color: C.white }, shadow: makeShadow(),
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: MARGIN, y: cY, w: cW2, h: 0.04,
    fill: { color: C.green },
  });
  slide.addText("COMMITMENTS", {
    x: MARGIN + 0.12, y: cY + 0.1, w: cW2 - 0.24, h: 0.18,
    fontSize: 8.5, fontFace: FONT.sub, color: C.green, bold: true, margin: 0, charSpacing: 1.5,
  });
  slide.addText([
    { text: "Net zero GHG by 2050", options: { bold: true, breakLine: true, fontSize: 9, fontFace: FONT.body, color: C.charcoal, paraSpaceAfter: 4 } },
    { text: "First global airline to commit without traditional offsets\n", options: { fontSize: 8.5, fontFace: FONT.body, color: C.slate, breakLine: true, paraSpaceAfter: 6 } },
    { text: "50% emissions intensity reduction by 2035", options: { bold: true, breakLine: true, fontSize: 9, fontFace: FONT.body, color: C.charcoal, paraSpaceAfter: 4 } },
    { text: "vs. 2019 baseline, SBTi-validated\n", options: { fontSize: 8.5, fontFace: FONT.body, color: C.slate, breakLine: true, paraSpaceAfter: 6 } },
    { text: "CDP Score: A\u2013  |  $200M+ UAL SAF Fund\n", options: { bold: true, breakLine: true, fontSize: 9, fontFace: FONT.body, color: C.charcoal, paraSpaceAfter: 6 } },
    { text: "Eco-Skies Alliance: 50+ corporate partners co-investing in SAF \u2014 socializes supply development cost across corporate customers", options: { fontSize: 8.5, fontFace: FONT.body, color: C.slate } },
  ], {
    x: MARGIN + 0.12, y: cY + 0.32, w: cW2 - 0.24, h: 1.7,
    margin: 0,
  });

  slide.addShape(pres.shapes.RECTANGLE, {
    x: MARGIN + cW2 + 0.15, y: cY, w: cW2, h: 2.1,
    fill: { color: C.white }, shadow: makeShadow(),
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: MARGIN + cW2 + 0.15, y: cY, w: cW2, h: 0.04,
    fill: { color: C.amber },
  });
  slide.addText("HONEST INVESTOR FRAMING", {
    x: MARGIN + cW2 + 0.27, y: cY + 0.1, w: cW2 - 0.24, h: 0.18,
    fontSize: 8.5, fontFace: FONT.sub, color: C.amber, bold: true, margin: 0, charSpacing: 1.5,
  });
  slide.addText(
    "SAF at 13.6M gal vs. 4.2B+ total fuel = <0.35%. Doubling annually since 2021, but the gap between current penetration and 2050 target is enormous.\n\nAbsolute emissions will rise with capacity growth for the foreseeable future. 2035 target is intensity-based (per RTK/ASM) \u2014 achievable with newer aircraft. 2050 net zero depends on technologies that don\u2019t exist at commercial scale today (green hydrogen, next-gen SAF, carbon removal).\n\nUAL Ventures ($200M+): Twelve, ZeroAvia, Dimensional Energy, Cemvita, Svante \u2014 shaping the supply curve, not just procurement. Genuine strategic investment, not greenwashing.",
    {
      x: MARGIN + cW2 + 0.27, y: cY + 0.32, w: cW2 - 0.24, h: 1.7,
      fontSize: 8.5, fontFace: FONT.body, color: C.charcoal, margin: 0, lineSpacingMultiple: 1.2,
    }
  );
}

// ============================================================
// SLIDE 16 — 2026 GUIDANCE & RISK/REWARD
// ============================================================
{
  const slide = addDarkSlide();
  addFooter(slide, 16, TOTAL_SLIDES, true);

  slide.addText("2026 GUIDANCE & RISK MATRIX", {
    x: MARGIN, y: 0.3, w: CONTENT_W, h: 0.45,
    fontSize: 22, fontFace: FONT.head, color: C.white, bold: true, margin: 0, charSpacing: 3,
  });

  // Guidance stats
  const gStatY = 0.9;
  const gsw = (CONTENT_W - 0.75) / 6;
  const gItems = [
    { val: "$12\u201314", lbl: "Adj. EPS\nGuidance" },
    { val: "$13.00", lbl: "EPS\nMidpoint" },
    { val: "~10%+", lbl: "Pre-Tax\nMargin" },
    { val: "~$2.7B", lbl: "Free Cash\nFlow" },
    { val: "<$8.0B", lbl: "Net\nCapEx" },
    { val: "<2.0\u00D7", lbl: "Year-End\nLeverage" },
  ];
  gItems.forEach((g, i) => {
    addStatCard(slide, MARGIN + i * (gsw + 0.15), gStatY, gsw, 0.8, g.val, g.lbl, {
      valueFontSize: 18,
      labelFontSize: 8,
    });
  });

  // Risk matrix
  slide.addText("KEY RISKS TO THE 2026 THESIS", {
    x: MARGIN, y: 1.9, w: 4, h: 0.22,
    fontSize: 9, fontFace: FONT.sub, color: C.ice, bold: true, margin: 0, charSpacing: 2,
  });

  const riskRows = [
    ["Risk", "Magnitude", "Mitigant"],
    ["GDP recession", "\u20131.5\u20132% rev per 1% GDP", "Premium/intl mix; loyalty is contract-based"],
    ["Fuel spike", "$0.10/gal = ~$466M; unhedged", "New aircraft burn ~20% less/seat"],
    ["Boeing MAX delays", "~$30\u201340M/mo deferred efficiency", "A321neo/XLR hedge; existing fleet performing"],
    ["Domestic PRASM stays negative", "Margin expansion stalls", "Mid-year capacity cuts showed discipline"],
    ["MAX 10 certification slip", "200 aircraft in tranche waiting", "Can defer CapEx; no hard date committed"],
    ["MileagePlus re-contract (2029)", "Chase deal terms uncertain", "Delta/AmEx parity gives pricing leverage"],
  ];

  const rHeaderOpts = {
    fill: { color: C.midNavy },
    color: C.accent,
    bold: true,
    fontSize: 8,
    fontFace: FONT.body,
    align: "center",
    valign: "middle",
  };
  const rCellOpts = (rIdx) => ({
    fill: { color: rIdx % 2 === 0 ? C.deepNavy : C.midNavy },
    color: C.ice,
    fontSize: 8,
    fontFace: FONT.body,
    align: "left",
    valign: "middle",
  });
  const rTable = riskRows.map((row, rIdx) =>
    row.map((cell) => ({
      text: String(cell),
      options: rIdx === 0 ? { ...rHeaderOpts, align: "center" } : rCellOpts(rIdx),
    }))
  );
  slide.addTable(rTable, {
    x: MARGIN, y: 2.15, w: CONTENT_W,
    border: { pt: 0.5, color: C.midNavy },
    colW: [2.2, 2.5, 4.2],
    rowH: 0.27,
    margin: [2, 5, 2, 5],
  });

  // Bottom thesis
  slide.addShape(pres.shapes.RECTANGLE, {
    x: MARGIN, y: 4.15, w: CONTENT_W, h: 0.85,
    fill: { color: C.midNavy },
    shadow: makeShadow(),
  });
  slide.addText([
    { text: "THE BULL CASE: ", options: { bold: true, fontSize: 10, fontFace: FONT.sub, color: C.accent } },
    { text: "If TRASM inflects positive in Q1\u2013Q2 2026 as industry capacity tightens, UAL tracks toward $14 EPS or higher. At the current multiple, that represents significant equity upside. ", options: { fontSize: 10, fontFace: FONT.body, color: C.ice } },
    { text: "The base case of $13 midpoint is already a ~22% EPS growth year \u2014 which in a normal market commands a premium multiple for an airline finally breaking through into double-digit pre-tax margins.", options: { fontSize: 10, fontFace: FONT.body, color: C.ice } },
  ], {
    x: MARGIN + 0.15, y: 4.2, w: CONTENT_W - 0.3, h: 0.75,
    margin: 0, lineSpacingMultiple: 1.3,
  });
}

// ============================================================
// SLIDE 17 — APPENDIX: FINANCIAL STATEMENTS
// ============================================================
{
  const slide = addLightSlide();
  addLightHeader(slide, "Appendix A \u2014 Financial Detail", "Three-year income statement, hub network, workforce, and fleet history");
  addFooter(slide, 17, TOTAL_SLIDES);

  // Hub system table
  slide.addText("HUB NETWORK (2026 SCHEDULE)", {
    x: MARGIN, y: 0.95, w: 4, h: 0.2,
    fontSize: 8, fontFace: FONT.sub, color: C.navy, bold: true, margin: 0, charSpacing: 1.5,
  });
  const hubRows = [
    ["Hub", "Code", "2026 Dep.", "2026 Seats", "Position"],
    ["Chicago O'Hare", "ORD", "250,306", "28.6M", "#1 (31 consec. qtrs)"],
    ["Denver", "DEN", "194,200", "\u2014", "#1 (11 consec. qtrs)"],
    ["Houston Bush", "IAH", "176,734", "22.0M", "#1"],
    ["Newark", "EWR", "141,087", "20.8M", "#1 NYC intl gateway"],
    ["San Francisco", "SFO", "104,442", "\u2014", "#1"],
    ["Washington Dulles", "IAD", "95,037", "11.8M", "#1"],
    ["Los Angeles", "LAX", "49,717", "8.1M", "#2"],
    ["System Total", "\u2014", "1,896,990", "242.6M", "Largest U.S. by ASMs"],
  ];
  addStyledTable(slide, hubRows, MARGIN, 1.17, CONTENT_W * 0.55, { rowH: 0.2 });

  // Workforce & labor
  slide.addText("WORKFORCE & LABOR CONTRACTS", {
    x: MARGIN + CONTENT_W * 0.58, y: 0.95, w: 4, h: 0.2,
    fontSize: 8, fontFace: FONT.sub, color: C.navy, bold: true, margin: 0, charSpacing: 1.5,
  });
  const laborRows = [
    ["Date", "Employees", "Salaries ($B)"],
    ["Dec 2023", "103,300", "$14.8"],
    ["Dec 2024", "107,300", "$16.7"],
    ["Dec 2025", "113,200", "$17.6"],
  ];
  addStyledTable(slide, laborRows, MARGIN + CONTENT_W * 0.58, 1.17, CONTENT_W * 0.42, { rowH: 0.22 });

  const lcRows = [
    ["Union", "Group", "Status"],
    ["ALPA", "Pilots", "Ratified mid-2023; ~$10B; through ~2027"],
    ["AFA", "Flight Attendants", "TA early 2025; 20\u201328% raise"],
    ["IBT", "Mechanics", "Ratified 2024; substantial"],
    ["\u2014", "Ground Workers", "Updated 2024\u201325; inflation-adj."],
  ];
  addStyledTable(slide, lcRows, MARGIN + CONTENT_W * 0.58, 2.15, CONTENT_W * 0.42, { rowH: 0.22 });

  // Long-term targets
  slide.addText("LONG-TERM UNITED NEXT TARGETS", {
    x: MARGIN + CONTENT_W * 0.58, y: 3.35, w: 4, h: 0.2,
    fontSize: 8, fontFace: FONT.sub, color: C.navy, bold: true, margin: 0, charSpacing: 1.5,
  });
  const ltItems = [
    { text: "Investment-grade credit rating", options: { bullet: true, breakLine: true, fontSize: 8.5, fontFace: FONT.body, color: C.charcoal, paraSpaceAfter: 3 } },
    { text: "Double-digit pre-tax margin (10%+) sustained", options: { bullet: true, breakLine: true, fontSize: 8.5, fontFace: FONT.body, color: C.charcoal, paraSpaceAfter: 3 } },
    { text: "Net leverage <2.0\u00D7 by year-end 2026", options: { bullet: true, breakLine: true, fontSize: 8.5, fontFace: FONT.body, color: C.charcoal, paraSpaceAfter: 3 } },
    { text: "Double MileagePlus profits by 2030", options: { bullet: true, breakLine: true, fontSize: 8.5, fontFace: FONT.body, color: C.charcoal, paraSpaceAfter: 3 } },
    { text: "630+ new aircraft by 2034", options: { bullet: true, breakLine: true, fontSize: 8.5, fontFace: FONT.body, color: C.charcoal, paraSpaceAfter: 3 } },
    { text: "Full-fleet Starlink by end 2027", options: { bullet: true, breakLine: true, fontSize: 8.5, fontFace: FONT.body, color: C.charcoal, paraSpaceAfter: 3 } },
    { text: "Net zero emissions by 2050", options: { bullet: true, fontSize: 8.5, fontFace: FONT.body, color: C.charcoal } },
  ];
  slide.addText(ltItems, {
    x: MARGIN + CONTENT_W * 0.58, y: 3.55, w: CONTENT_W * 0.42, h: 1.6,
    margin: 0,
  });

  // Fleet count history
  slide.addText("FLEET TRAJECTORY", {
    x: MARGIN, y: 3.35, w: 3, h: 0.2,
    fontSize: 8, fontFace: FONT.sub, color: C.navy, bold: true, margin: 0, charSpacing: 1.5,
  });
  const fleetHist = [
    ["Period", "Fleet Count"],
    ["Dec 2023", "1,358"],
    ["Dec 2024", "1,406"],
    ["Q1 2025", "1,442"],
    ["Q2 2025", "1,473"],
    ["Q3 2025", "1,486"],
    ["Dec 2025", "1,490"],
    ["Dec 2026E", "~1,610"],
  ];
  addStyledTable(slide, fleetHist, MARGIN, 3.55, CONTENT_W * 0.25, { rowH: 0.2 });

  // Macro sensitivity
  slide.addText("MACRO: Every \u20131% GDP \u2248 \u20131.5\u20132% revenue. Intl revenue (>35% pax) provides structural diversification. Premium/loyalty (~60%+ of rev) provides pricing buffer.", {
    x: MARGIN + CONTENT_W * 0.27, y: 3.55, w: CONTENT_W * 0.28, h: 1.5,
    fontSize: 8, fontFace: FONT.body, color: C.slate, margin: 0, lineSpacingMultiple: 1.3,
  });
}

// ============================================================
// SLIDE 18 — APPENDIX: DETAILED UNIT ECONOMICS & FUEL
// ============================================================
{
  const slide = addLightSlide();
  addLightHeader(slide, "Appendix B \u2014 Unit Economics & Operating Detail", "Quarterly unit economics, fuel statistics, revenue quality detail, and new route additions");
  addFooter(slide, 18, TOTAL_SLIDES);

  // Full quarterly unit economics
  slide.addText("QUARTERLY UNIT ECONOMICS \u2014 FY2025", {
    x: MARGIN, y: 0.95, w: 5, h: 0.2,
    fontSize: 8, fontFace: FONT.sub, color: C.navy, bold: true, margin: 0, charSpacing: 1.5,
  });
  const ueRows = [
    ["Quarter", "TRASM", "PRASM", "Yield", "CASM", "CASM-ex", "Fuel/Gal", "Load Factor"],
    ["Q1 2025", "17.58\u00A2", "15.78\u00A2", "19.93\u00A2", "16.77\u00A2", "13.17\u00A2", "$2.53", "79.2%"],
    ["Q2 2025", "18.06\u00A2", "16.40\u00A2", "19.74\u00A2", "16.49\u00A2", "12.36\u00A2", "$2.34", "83.1%"],
    ["Q3 2025", "17.42\u00A2", "15.80\u00A2", "18.73\u00A2", "15.82\u00A2", "12.15\u00A2", "$2.43", "84.4%"],
    ["Q4 2025", "18.47\u00A2", "16.71\u00A2", "20.41\u00A2", "16.81\u00A2", "12.94\u00A2", "$2.49", "81.9%"],
    ["FY2025", "17.88\u00A2", "16.18\u00A2", "19.67\u00A2", "16.46\u00A2", "12.64\u00A2", "$2.44", "82.2%"],
    ["FY2024", "18.34\u00A2", "16.66\u00A2", "20.05\u00A2", "16.70\u00A2", "12.58\u00A2", "$2.65", "83.1%"],
    ["FY2023", "18.44\u00A2", "\u2014", "20.07\u00A2", "16.99\u00A2", "12.03\u00A2", "$3.01", "83.9%"],
  ];
  addStyledTable(slide, ueRows, MARGIN, 1.17, CONTENT_W, { rowH: 0.22 });

  // Revenue quality
  slide.addText("REVENUE QUALITY", {
    x: MARGIN, y: 3.0, w: 3, h: 0.2,
    fontSize: 8, fontFace: FONT.sub, color: C.navy, bold: true, margin: 0, charSpacing: 1.5,
  });
  const rqRows = [
    ["Segment", "Q4 2025 YoY", "FY2025 YoY"],
    ["Premium cabin", "+9%", "+11%"],
    ["MileagePlus / Loyalty", "+10%", "+9%"],
    ["Basic Economy", "+7%", "+5%"],
    ["Chase co-brand spend", "+14%", "+12%"],
  ];
  addStyledTable(slide, rqRows, MARGIN, 3.2, CONTENT_W * 0.33, { rowH: 0.22 });

  // Key MileagePlus stats
  slide.addText("MILEAGEPLUS DETAIL", {
    x: MARGIN + CONTENT_W * 0.36, y: 3.0, w: 3, h: 0.2,
    fontSize: 8, fontFace: FONT.sub, color: C.navy, bold: true, margin: 0, charSpacing: 1.5,
  });
  const mpRows = [
    ["Metric", "Value"],
    ["Total members", "130M+"],
    ["Flight awards redeemed", "10.9M (10.3% RPMs)"],
    ["Non-flight awards", "4.3M"],
    ["New Chase cards/yr", "1M+ (3rd consec. yr)"],
    ["Chase agreement", "Through 2029"],
    ["2020 collateral valuation", "~$22B"],
  ];
  addStyledTable(slide, mpRows, MARGIN + CONTENT_W * 0.36, 3.2, CONTENT_W * 0.31, { rowH: 0.22 });

  // 2025 new routes
  slide.addText("2025\u20132026 ROUTE ADDITIONS", {
    x: MARGIN + CONTENT_W * 0.7, y: 3.0, w: 3, h: 0.2,
    fontSize: 8, fontFace: FONT.sub, color: C.navy, bold: true, margin: 0, charSpacing: 1,
  });
  slide.addText([
    { text: "2025: ", options: { bold: true, fontSize: 8, fontFace: FONT.body, color: C.charcoal, breakLine: false } },
    { text: "Nuuk, Ulaanbaatar, Faro, Palermo, Bilbao, Madeira, Bangkok, Adelaide, Ho Chi Minh City + 29 domestic/Canadian\n\n", options: { fontSize: 8, fontFace: FONT.body, color: C.slate, breakLine: true } },
    { text: "2026: ", options: { bold: true, fontSize: 8, fontFace: FONT.body, color: C.charcoal, breakLine: false } },
    { text: "Bari, Split, Glasgow, Santiago de Compostela, Seoul expansion, Reykjavik expansion\n\n", options: { fontSize: 8, fontFace: FONT.body, color: C.slate, breakLine: true } },
    { text: "\"Only U.S. carrier\" routes: Bangkok, Ho Chi Minh City, Adelaide, Tepic", options: { italic: true, fontSize: 8, fontFace: FONT.body, color: C.accent } },
  ], {
    x: MARGIN + CONTENT_W * 0.7, y: 3.2, w: CONTENT_W * 0.3, h: 1.9,
    margin: 0, lineSpacingMultiple: 1.2,
  });

  // CASM-ex bridge
  slide.addText(
    "CASM-ex Bridge (Q1 2025): CASM 16.77\u00A2 \u2192 less fuel (3.59)\u00A2 \u2192 less profit sharing (0.06)\u00A2 \u2192 less third-party (0.09)\u00A2 \u2192 add back special charges 0.14\u00A2 \u2192 CASM-ex 13.17\u00A2",
    {
      x: MARGIN, y: SLIDE_H - 0.4, w: CONTENT_W, h: 0.25,
      fontSize: 7.5, fontFace: FONT.body, color: C.slate, italic: true, margin: 0,
    }
  );
}

// ============================================================
// WRITE FILE
// ============================================================
pres.writeFile({ fileName: "outputs/corporate-anthropic.pptx" })
  .then(() => console.log("Created: outputs/corporate-anthropic.pptx"))
  .catch((err) => console.error("Error:", err));
