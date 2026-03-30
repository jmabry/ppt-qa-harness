const pptxgen = require("pptxgenjs");

// ============================================================
// NovaCrest Q3 2026 Strategic Review — CONFIDENTIAL
// Board Presentation
// ============================================================

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.author = "NovaCrest";
pres.title = "NovaCrest Q3 2026 Strategic Review";

// ============================================================
// COLOR PALETTE — Charcoal Minimal
// ============================================================
const C = {
  bgDark:    "1A1A2E",
  bgLight:   "F8FAFC",
  primary:   "16213E",
  accent:    "0F3460",
  highlight: "533483",
  green:     "10B981",
  amber:     "F59E0B",
  red:       "EF4444",
  white:     "FFFFFF",
  offWhite:  "F8FAFC",
  slate:     "64748B",
  lightGray: "94A3B8",
  cardBg:    "FFFFFF",
  cardBorder:"E2E8F0",
  charcoal:  "1E293B",
  rowAlt:    "EFF6FF",
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
// LAYOUT CONSTANTS
// ============================================================
const SLIDE_W = 10;
const SLIDE_H = 5.625;
const MARGIN = 0.45;
const CONTENT_W = SLIDE_W - 2 * MARGIN;

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

// Corner accent blocks for dark slides
function addCornerAccents(slide) {
  // Top-left corner
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.18, h: 0.18,
    fill: { color: C.highlight },
    line: { color: C.highlight, width: 0 },
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.06, h: 1.2,
    fill: { color: C.accent },
    line: { color: C.accent, width: 0 },
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 1.2, h: 0.06,
    fill: { color: C.accent },
    line: { color: C.accent, width: 0 },
  });
  // Bottom-right corner
  slide.addShape(pres.shapes.RECTANGLE, {
    x: SLIDE_W - 0.18, y: SLIDE_H - 0.18, w: 0.18, h: 0.18,
    fill: { color: C.highlight },
    line: { color: C.highlight, width: 0 },
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: SLIDE_W - 0.06, y: SLIDE_H - 1.2, w: 0.06, h: 1.2,
    fill: { color: C.accent },
    line: { color: C.accent, width: 0 },
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: SLIDE_W - 1.2, y: SLIDE_H - 0.06, w: 1.2, h: 0.06,
    fill: { color: C.accent },
    line: { color: C.accent, width: 0 },
  });
}

// Left-border accent bar for light slide content sections
function addAccentBar(slide, x, y, h, color) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x: x,
    y: y,
    w: 0.055,
    h: h,
    fill: { color: color || C.accent },
    line: { color: color || C.accent, width: 0 },
  });
}

// Slide header bar for light slides
function addLightHeader(slide, title, subtitle) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: SLIDE_W, h: 0.72,
    fill: { color: C.primary },
    line: { color: C.primary, width: 0 },
  });
  // Accent highlight line under header
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0.72, w: SLIDE_W, h: 0.04,
    fill: { color: C.highlight },
    line: { color: C.highlight, width: 0 },
  });
  slide.addText(title.toUpperCase(), {
    x: MARGIN,
    y: 0.1,
    w: subtitle ? CONTENT_W * 0.65 : CONTENT_W,
    h: 0.38,
    fontSize: 15,
    fontFace: FONT.head,
    color: C.white,
    bold: true,
    align: "left",
    margin: 0,
    charSpacing: 1.5,
  });
  if (subtitle) {
    slide.addText(subtitle, {
      x: MARGIN,
      y: 0.48,
      w: CONTENT_W,
      h: 0.22,
      fontSize: 9,
      fontFace: FONT.body,
      color: C.lightGray,
      align: "left",
      margin: 0,
    });
  }
}

// Footer for all slides
function addFooter(slide, num, total, dark = false) {
  const color = dark ? C.lightGray : C.slate;
  slide.addText(`NOVACREST  |  Q3 2026 STRATEGIC REVIEW  |  CONFIDENTIAL`, {
    x: MARGIN,
    y: SLIDE_H - 0.28,
    w: CONTENT_W - 1,
    h: 0.22,
    fontSize: 7,
    fontFace: FONT.body,
    color: color,
    align: "left",
    margin: 0,
    charSpacing: 0.5,
  });
  slide.addText(`${num} / ${total}`, {
    x: SLIDE_W - 1.1,
    y: SLIDE_H - 0.28,
    w: 0.8,
    h: 0.22,
    fontSize: 7.5,
    fontFace: FONT.body,
    color: color,
    align: "right",
    margin: 0,
  });
}

// Section label (accent-colored uppercase label)
function addSectionLabel(slide, x, y, w, label, color) {
  slide.addText(label.toUpperCase(), {
    x: x + 0.1,
    y: y,
    w: w,
    h: 0.22,
    fontSize: 8,
    fontFace: FONT.head,
    color: color || C.accent,
    bold: true,
    align: "left",
    margin: 0,
    charSpacing: 1.5,
  });
}

// KPI card with left-border accent bar (light slides)
function addKpiCard(slide, x, y, w, h, value, label, sub, opts = {}) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w, h,
    fill: { color: C.white },
    line: { color: C.cardBorder, width: 0.5 },
    shadow: { type: "outer", color: "000000", blur: 3, offset: 1, angle: 135, opacity: 0.08 },
  });
  // Left accent bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w: 0.055, h,
    fill: { color: opts.accentColor || C.accent },
    line: { color: opts.accentColor || C.accent, width: 0 },
  });
  slide.addText(value, {
    x: x + 0.12, y: y + 0.08, w: w - 0.16, h: h * 0.48,
    fontSize: opts.valueFontSize || 22,
    fontFace: FONT.head,
    color: opts.valueColor || C.primary,
    bold: true,
    align: "left",
    valign: "middle",
    margin: 0,
  });
  slide.addText(label, {
    x: x + 0.12, y: y + h * 0.54, w: w - 0.16, h: h * 0.24,
    fontSize: 8,
    fontFace: FONT.body,
    color: C.charcoal,
    bold: true,
    align: "left",
    margin: 0,
  });
  if (sub) {
    slide.addText(sub, {
      x: x + 0.12, y: y + h * 0.76, w: w - 0.16, h: h * 0.22,
      fontSize: 8,
      fontFace: FONT.body,
      color: opts.subColor || C.slate,
      align: "left",
      margin: 0,
    });
  }
}

// Dark stat callout for title slide
function addDarkStatCallout(slide, x, y, w, h, value, label, opts = {}) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w, h,
    fill: { color: opts.bg || C.primary },
    line: { color: C.accent, width: 1 },
  });
  // Left accent bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w: 0.06, h,
    fill: { color: opts.accentColor || C.highlight },
    line: { color: opts.accentColor || C.highlight, width: 0 },
  });
  slide.addText(value, {
    x: x + 0.1, y: y + 0.05, w: w - 0.14, h: h * 0.55,
    fontSize: opts.valueFontSize || 26,
    fontFace: FONT.head,
    color: opts.valueColor || C.white,
    bold: true,
    align: "left",
    valign: "middle",
    margin: 0,
  });
  slide.addText(label, {
    x: x + 0.1, y: y + h * 0.6, w: w - 0.14, h: h * 0.36,
    fontSize: 8.5,
    fontFace: FONT.body,
    color: C.lightGray,
    align: "left",
    valign: "top",
    margin: 0,
    lineSpacingMultiple: 1.2,
  });
}

// Styled table: first row = header
function addStyledTable(slide, rows, x, y, w, opts = {}) {
  const colCount = rows[0].length;
  const colW = opts.colW || Array(colCount).fill(w / colCount);

  const defaultHeaderFill = opts.headerFill || C.primary;
  const defaultHeaderColor = opts.headerColor || C.white;

  const tableRows = rows.map((row, rIdx) =>
    row.map((cell, cIdx) => {
      let cellText = cell;
      let cellOpts = {};

      if (typeof cell === "object" && cell !== null) {
        cellText = cell.text !== undefined ? cell.text : "";
        cellOpts = cell.opts || {};
      }

      if (rIdx === 0) {
        return {
          text: String(cellText),
          options: {
            fill: { color: cellOpts.fill || defaultHeaderFill },
            color: cellOpts.color || defaultHeaderColor,
            bold: true,
            fontSize: opts.headerFontSize || 8.5,
            fontFace: FONT.body,
            align: cellOpts.align || (cIdx === 0 ? "left" : "center"),
            valign: "middle",
          },
        };
      }

      const altRow = rIdx % 2 === 0 ? C.offWhite : C.white;
      const baseFill = cellOpts.fill || altRow;
      return {
        text: String(cellText),
        options: {
          fill: { color: baseFill },
          color: cellOpts.color || C.charcoal,
          bold: cellOpts.bold || false,
          fontSize: opts.bodyFontSize || 8.5,
          fontFace: FONT.body,
          align: cellOpts.align || (cIdx === 0 ? "left" : "center"),
          valign: "middle",
          italic: cellOpts.italic || false,
        },
      };
    })
  );

  slide.addTable(tableRows, {
    x, y, w,
    border: { pt: 0.5, color: C.cardBorder },
    colW,
    rowH: opts.rowH || 0.285,
    margin: [2, 5, 2, 5],
  });
}

const TOTAL_SLIDES = 6;

// ============================================================
// SLIDE 1 — TITLE (dark)
// ============================================================
{
  const slide = addDarkSlide();
  addCornerAccents(slide);

  // Subtle background gradient block on right
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5.5, y: 0, w: 4.5, h: SLIDE_H,
    fill: { color: C.primary },
    line: { color: C.primary, width: 0 },
  });
  // Vertical divider accent
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5.5, y: 0, w: 0.05, h: SLIDE_H,
    fill: { color: C.highlight },
    line: { color: C.highlight, width: 0 },
  });

  // Company name — dominant
  slide.addText("NovaCrest", {
    x: MARGIN,
    y: 0.55,
    w: 5.0,
    h: 0.85,
    fontSize: 52,
    fontFace: FONT.head,
    color: C.white,
    bold: true,
    align: "left",
    margin: 0,
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: MARGIN, y: 1.42, w: 2.6, h: 0.045,
    fill: { color: C.highlight },
    line: { color: C.highlight, width: 0 },
  });
  slide.addText("Q3 2026 Strategic Review — CONFIDENTIAL", {
    x: MARGIN,
    y: 1.52,
    w: 5.0,
    h: 0.32,
    fontSize: 12,
    fontFace: FONT.sub,
    color: C.lightGray,
    align: "left",
    margin: 0,
    charSpacing: 0.5,
  });
  slide.addText("October 2026", {
    x: MARGIN,
    y: 1.9,
    w: 3,
    h: 0.28,
    fontSize: 10,
    fontFace: FONT.body,
    color: C.slate,
    align: "left",
    margin: 0,
  });

  // Left column stats
  const lcx = 0.38;
  const rcx = 5.65;
  const cw = 1.95;
  const sh = 0.72;
  const sy = 2.42;
  const gap = 0.16;

  addDarkStatCallout(slide, lcx, sy,           cw, sh, "$15.6M", "ARR",          { accentColor: C.green,     valueFontSize: 28 });
  addDarkStatCallout(slide, lcx, sy + sh + gap, cw, sh, "118%",   "Net Revenue Retention", { accentColor: C.amber, valueFontSize: 28 });
  addDarkStatCallout(slide, lcx, sy + 2*(sh+gap), cw, sh, "78.2%", "Gross Margin",         { accentColor: C.green,     valueFontSize: 28 });

  // Right column stats
  addDarkStatCallout(slide, rcx, sy,           cw, sh, "$18.2M", "Cash on Hand",   { accentColor: C.green,   valueFontSize: 28 });
  addDarkStatCallout(slide, rcx, sy + sh + gap, cw, sh, "81 mo",  "Runway",         { accentColor: C.green,   valueFontSize: 28 });
  addDarkStatCallout(slide, rcx, sy + 2*(sh+gap), cw, sh, "0.36x", "Burn Multiple (Excellent <1x)", { accentColor: C.highlight, valueFontSize: 28 });

  // Board footer
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: SLIDE_H - 0.38, w: SLIDE_W, h: 0.38,
    fill: { color: C.primary },
    line: { color: C.primary, width: 0 },
  });
  slide.addText("Board of Directors — October 2026  |  Prepared by: Finance & Strategy  |  Do not distribute", {
    x: MARGIN,
    y: SLIDE_H - 0.33,
    w: CONTENT_W,
    h: 0.27,
    fontSize: 7.5,
    fontFace: FONT.body,
    color: C.lightGray,
    align: "center",
    margin: 0,
  });
}

// ============================================================
// SLIDE 2 — Revenue Dashboard (light)
// ============================================================
{
  const slide = addLightSlide();
  addLightHeader(slide, "Revenue Dashboard", "NovaCrest Q3 2026 — ARR, Unit Economics & Retention");
  addFooter(slide, 2, TOTAL_SLIDES, false);

  const cardY = 0.88;
  const cardH = 0.82;
  const cardW = 2.2;
  const gap = 0.1;
  const startX = MARGIN;

  // 4 KPI cards
  addKpiCard(slide, startX,                    cardY, cardW, cardH, "$15.6M", "ARR", "+10.6% QoQ",         { accentColor: C.green, subColor: C.green });
  addKpiCard(slide, startX + (cardW+gap),      cardY, cardW, cardH, "118%",   "Net Revenue Retention", "↓ from 121% — monitor", { accentColor: C.amber, subColor: C.amber });
  addKpiCard(slide, startX + 2*(cardW+gap),    cardY, cardW, cardH, "78.2%",  "Gross Margin", "Target: 80%+", { accentColor: C.accent, subColor: C.slate });
  addKpiCard(slide, startX + 3*(cardW+gap),    cardY, cardW, cardH, "1.12",   "Magic Number", "Strong (>0.75)", { accentColor: C.green, subColor: C.green });

  // ARR Trend Chart (manual bar chart)
  const chartY = 1.84;
  const chartH = 2.05;
  const chartX = MARGIN;
  const chartW = 4.35;

  // Chart background + left accent
  addAccentBar(slide, chartX, chartY, chartH, C.accent);
  slide.addShape(pres.shapes.RECTANGLE, {
    x: chartX + 0.055, y: chartY, w: chartW - 0.055, h: chartH,
    fill: { color: C.white },
    line: { color: C.cardBorder, width: 0.5 },
  });
  slide.addText("ARR GROWTH TREND ($M)", {
    x: chartX + 0.18, y: chartY + 0.07, w: chartW - 0.25, h: 0.22,
    fontSize: 8,
    fontFace: FONT.head,
    color: C.primary,
    bold: true,
    align: "left",
    margin: 0,
    charSpacing: 1,
  });

  // Bar chart data
  const bars = [
    { label: "Q1'25", val: 7.2 },
    { label: "Q2'25", val: 8.5 },
    { label: "Q3'25", val: 9.8 },
    { label: "Q4'25", val: 11.4 },
    { label: "Q1'26", val: 12.8 },
    { label: "Q2'26", val: 14.1 },
    { label: "Q3'26", val: 15.6 },
  ];
  const maxVal = 16.0;
  const barAreaX = chartX + 0.25;
  const barAreaY = chartY + 0.38;
  const barAreaH = chartH - 0.72;
  const barAreaW = chartW - 0.38;
  const barW = (barAreaW / bars.length) * 0.58;
  const barSpacing = barAreaW / bars.length;

  bars.forEach((b, i) => {
    const bh = (b.val / maxVal) * barAreaH;
    const bx = barAreaX + i * barSpacing + (barSpacing - barW) / 2;
    const by = barAreaY + barAreaH - bh;

    // Gradient effect: last bar is brighter
    const fillColor = i === bars.length - 1 ? C.highlight : C.accent;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: bx, y: by, w: barW, h: bh,
      fill: { color: fillColor },
      line: { color: fillColor, width: 0 },
    });
    // Value label
    slide.addText(`$${b.val}M`, {
      x: bx - 0.04, y: by - 0.2, w: barW + 0.08, h: 0.18,
      fontSize: 7.5,
      fontFace: FONT.body,
      color: i === bars.length - 1 ? C.highlight : C.primary,
      bold: i === bars.length - 1,
      align: "center",
      margin: 0,
    });
    // Quarter label
    slide.addText(b.label, {
      x: bx - 0.04, y: barAreaY + barAreaH + 0.02, w: barW + 0.08, h: 0.18,
      fontSize: 7.5,
      fontFace: FONT.body,
      color: C.slate,
      align: "center",
      margin: 0,
    });
  });

  // Unit Economics Table
  const tblX = chartX + chartW + 0.2;
  const tblY = chartY;
  const tblW = SLIDE_W - tblX - MARGIN;

  addAccentBar(slide, tblX, tblY, chartH, C.highlight);
  slide.addText("UNIT ECONOMICS", {
    x: tblX + 0.12, y: tblY + 0.07, w: tblW - 0.15, h: 0.22,
    fontSize: 8,
    fontFace: FONT.head,
    color: C.primary,
    bold: true,
    align: "left",
    margin: 0,
    charSpacing: 1,
  });

  const ueCols = [1.4, 0.7, 0.72, 0.72, 0.84];
  const ueRows = [
    ["Metric", "Q3", "Q2", "Target", "Status"],
    [
      "Gross Margin",
      "78.2%", "77.5%", "80%+",
      { text: "On track ↑", opts: { color: C.green, bold: true } }
    ],
    [
      "NRR",
      "118%", "121%", "120%+",
      { text: "Concern ↓", opts: { color: C.amber, bold: true } }
    ],
    [
      "Logo Retention",
      "92.4%", "94.1%", "95%+",
      { text: "Concern ↓", opts: { color: C.amber, bold: true } }
    ],
    [
      "LTV:CAC",
      "4.8x", "4.6x", "4.0x+",
      { text: "Healthy ✓", opts: { color: C.green, bold: true } }
    ],
    [
      "CAC Payback",
      "13.2 mo", "14.1 mo", "<16 mo",
      { text: "On track ↑", opts: { color: C.green, bold: true } }
    ],
    [
      "Magic Number",
      "1.12", "0.98", ">0.75",
      { text: "Strong ✓", opts: { color: C.green, bold: true } }
    ],
  ];

  addStyledTable(slide, ueRows, tblX + 0.08, tblY + 0.33, tblW - 0.08, {
    colW: ueCols,
    rowH: 0.255,
    bodyFontSize: 8.5,
  });

  // QoQ growth callout
  slide.addText("QoQ ARR growth: +10.6%  |  YoY growth: +59.2%  |  Logo retention decline requires immediate CS investment", {
    x: MARGIN,
    y: SLIDE_H - 0.45,
    w: CONTENT_W,
    h: 0.2,
    fontSize: 8,
    fontFace: FONT.body,
    color: C.slate,
    align: "left",
    margin: 0,
    italic: true,
  });
}

// ============================================================
// SLIDE 3 — GTM + Customer Health (light)
// ============================================================
{
  const slide = addLightSlide();
  addLightHeader(slide, "GTM + Customer Health", "Pipeline, Channel Performance & Q3 Churn Analysis");
  addFooter(slide, 3, TOTAL_SLIDES, false);

  const topY = 0.88;
  const colL = MARGIN;
  const colR = SLIDE_W / 2 + 0.05;
  const colW = SLIDE_W / 2 - MARGIN - 0.15;

  // ---- LEFT: GO-TO-MARKET ----
  addAccentBar(slide, colL, topY, 3.94, C.accent);
  addSectionLabel(slide, colL, topY + 0.04, colW, "Go-To-Market", C.accent);

  // Pipeline metrics table
  const pipeCols = [1.45, 0.7, 0.7, 0.9];
  const pipeRows = [
    ["Metric", "Q3", "Q2", "Change"],
    ["Pipeline", "$8.6M", "$7.1M", { text: "+21.1% ↑", opts: { color: C.green, bold: true } }],
    ["Win Rate", "28.4%", "26.1%", { text: "+2.3pp ↑", opts: { color: C.green, bold: true } }],
    ["Avg Deal", "$42K", "$38K", { text: "+10.5% ↑", opts: { color: C.green, bold: true } }],
    ["Sales Cycle", "68d", "72d", { text: "−5.6% ↑", opts: { color: C.green, bold: true } }],
    ["Quota Attainment", "112%", "94%", { text: "+18pp ↑", opts: { color: C.green, bold: true } }],
  ];
  addStyledTable(slide, pipeRows, colL + 0.08, topY + 0.27, colW - 0.08, {
    colW: pipeCols,
    rowH: 0.275,
    bodyFontSize: 8.5,
  });

  // Channel table
  const chanY = topY + 0.27 + 6 * 0.275 + 0.14;
  addAccentBar(slide, colL, chanY, 0.22 + 5 * 0.265, C.highlight);
  addSectionLabel(slide, colL, chanY + 0.03, colW, "Channel Breakdown", C.highlight);

  const chanCols = [1.1, 0.65, 0.55, 0.55, 0.8];
  const chanRows = [
    ["Channel", "Pipeline", "Deals", "ACV", "CAC"],
    ["SDR",     "$3.8M", "18", "$52K", "$24.1K"],
    ["Inbound", "$2.4M", "14", "$34K", "$12.8K"],
    ["PLG",     "$0.9M", "8",  "$18K", "$6.2K"],
    ["Partner", "$1.5M", "5",  "$68K", "$8.4K"],
  ];
  addStyledTable(slide, chanRows, colL + 0.08, chanY + 0.25, colW - 0.08, {
    colW: chanCols,
    rowH: 0.262,
    bodyFontSize: 8.5,
    headerFill: C.accent,
  });

  // ---- RIGHT: CHURN ANALYSIS ----
  const rightH = 3.94;
  addAccentBar(slide, colR, topY, rightH, C.red);
  addSectionLabel(slide, colR, topY + 0.04, colW, "Churn Analysis — Q3", C.red);

  // Churn callout
  slide.addShape(pres.shapes.RECTANGLE, {
    x: colR + 0.07, y: topY + 0.27, w: colW - 0.08, h: 0.35,
    fill: { color: "FEF2F2" },
    line: { color: C.red, width: 0.75 },
  });
  slide.addText("Q3 Churn: $500K gross  (+67% vs $300K in Q2)", {
    x: colR + 0.13, y: topY + 0.30, w: colW - 0.18, h: 0.28,
    fontSize: 9.5,
    fontFace: FONT.sub,
    color: C.red,
    bold: true,
    align: "left",
    margin: 0,
  });

  // Churn breakdown table
  const churnCols = [1.2, 0.52, 0.88, 0.88];
  const churnRows = [
    ["Account", "ARR", "Category", "Preventable?"],
    ["Apex Mfg", "$145K", "M&A", { text: "No", opts: { color: C.green } }],
    ["Precision Dynamics", "$62K", "In-house build", { text: "Partially", opts: { color: C.amber } }],
    ["4 SMB accounts", "$118K", "Price (DataForge)", { text: "Yes", opts: { color: C.red, bold: true } }],
    ["TechFab", "$48K", "Bad impl — CS fail", { text: "Yes — CS", opts: { color: C.red, bold: true } }],
    ["3 SMB accounts", "$82K", "Never onboarded", { text: "Yes", opts: { color: C.red, bold: true } }],
    ["Consolidated Parts", "$45K", "Budget cut", { text: "No", opts: { color: C.green } }],
  ];
  addStyledTable(slide, churnRows, colR + 0.08, topY + 0.66, colW - 0.08, {
    colW: churnCols,
    rowH: 0.262,
    bodyFontSize: 8.5,
    headerFill: C.primary,
  });

  // 58% preventable callout
  const calloutY = topY + 0.66 + 8 * 0.262 + 0.06;
  slide.addShape(pres.shapes.RECTANGLE, {
    x: colR + 0.07, y: calloutY, w: colW - 0.08, h: 0.38,
    fill: { color: C.primary },
    line: { color: C.primary, width: 0 },
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: colR + 0.07, y: calloutY, w: 0.045, h: 0.38,
    fill: { color: C.red },
    line: { color: C.red, width: 0 },
  });
  slide.addText("58% PREVENTABLE — $290K of $500K", {
    x: colR + 0.16, y: calloutY + 0.06, w: colW - 0.2, h: 0.26,
    fontSize: 11,
    fontFace: FONT.head,
    color: C.white,
    bold: true,
    align: "left",
    margin: 0,
  });

  // At-risk Q4 table
  const arY = calloutY + 0.44;
  addAccentBar(slide, colR, arY, 0.22 + 4 * 0.265, C.amber);
  addSectionLabel(slide, colR, arY + 0.03, colW, "At-Risk Q4 Accounts", C.amber);
  const arCols = [1.1, 0.52, 1.86];
  const arRows = [
    ["Account", "ARR", "Risk Signal"],
    ["Sterling Mfg", "$210K", "Champion departed — outreach stalled"],
    ["Midwest Parts", "$58K", "Usage down -40% last 60 days"],
    ["ClearPath", "$72K", "Renewal + competitive POC active"],
  ];
  addStyledTable(slide, arRows, colR + 0.08, arY + 0.25, colW - 0.08, {
    colW: arCols,
    rowH: 0.262,
    bodyFontSize: 8.5,
    headerFill: C.amber,
    headerColor: C.white,
  });
}

// ============================================================
// SLIDE 4 — Product + Team (light)
// ============================================================
{
  const slide = addLightSlide();
  addLightHeader(slide, "Product + Team", "Q3 Shipped Features, Q4 Roadmap & Headcount");
  addFooter(slide, 4, TOTAL_SLIDES, false);

  const topY = 0.88;

  // SHIPPED FEATURES
  addAccentBar(slide, MARGIN, topY, 0.22 + 5 * 0.27, C.green);
  addSectionLabel(slide, MARGIN, topY + 0.04, CONTENT_W / 2, "Q3 Shipped Features", C.green);

  const featureCols = [2.05, 1.8, 2.0];
  const featureRows = [
    ["Feature", "Adoption", "Revenue Impact"],
    ["Predictive Maintenance v2", "67% enterprise accounts", "$1.2M pipeline (3 deals)"],
    ["Self-Serve Dashboards", "340 boards, 89 accounts", "CS tickets −18%"],
    ["SAP Integration (native)", "12 accounts connected", "Partner channel accelerant"],
    ["SOC 2 Type II", "Cert issued 8/15/26", "$380K pipeline unblocked"],
  ];
  addStyledTable(slide, featureRows, MARGIN + 0.08, topY + 0.27, 5.85, {
    colW: featureCols,
    rowH: 0.27,
    bodyFontSize: 8.5,
    headerFill: C.primary,
  });

  // Q4 ROADMAP
  const rdY = topY;
  const rdX = MARGIN + 5.85 + 0.2;
  const rdW = SLIDE_W - rdX - MARGIN;
  addAccentBar(slide, rdX, rdY, 0.22 + 5 * 0.27, C.highlight);
  addSectionLabel(slide, rdX, rdY + 0.04, rdW, "Q4 Roadmap", C.highlight);

  const rdCols = [0.42, 1.55, 0.9, 1.1];
  const rdRows = [
    ["Pri", "Feature", "Status", "Impact"],
    [
      { text: "P0", opts: { color: C.red, bold: true } },
      "Multi-tenant analytics",
      "60% dev",
      "2 deals $200K+ ea"
    ],
    [
      { text: "P0", opts: { color: C.red, bold: true } },
      "Siemens MindSphere",
      "Dev starting",
      "$4M TAM"
    ],
    [
      { text: "P1", opts: { color: C.amber, bold: true } },
      "Usage-based pricing",
      "Spec complete",
      "Fix 58% prev churn"
    ],
    [
      { text: "P1", opts: { color: C.amber, bold: true } },
      "Customer health score",
      "Prototype",
      "Early warning"
    ],
  ];
  addStyledTable(slide, rdRows, rdX + 0.08, rdY + 0.27, rdW - 0.08, {
    colW: rdCols,
    rowH: 0.27,
    bodyFontSize: 8.5,
    headerFill: C.accent,
  });

  // HEADCOUNT
  const hcY = topY + 0.22 + 5 * 0.27 + 0.18;
  addAccentBar(slide, MARGIN, hcY, 0.22 + 8 * 0.268, C.accent);
  addSectionLabel(slide, MARGIN, hcY + 0.04, CONTENT_W * 0.65, "Headcount", C.accent);

  const hcCols = [1.75, 0.6, 0.6, 0.6, 0.75];
  const hcRows = [
    ["Function", "Q2", "Q3", "Open", "Q4 Target"],
    ["Engineering",        "42", "46", "4", "50"],
    ["Product & Design",   "8",  "9",  "1", "10"],
    ["Sales",              "18", "22", "3", "25"],
    ["Customer Success",   "12", "14", "2", "16"],
    ["Marketing",          "8",  "9",  "1", "10"],
    ["G&A",                "10", "12", "1", "13"],
    [
      { text: "Total", opts: { bold: true } },
      { text: "98", opts: { bold: true } },
      { text: "112", opts: { bold: true } },
      { text: "12", opts: { bold: true } },
      { text: "124", opts: { bold: true } }
    ],
  ];
  addStyledTable(slide, hcRows, MARGIN + 0.08, hcY + 0.27, 4.3, {
    colW: hcCols,
    rowH: 0.268,
    bodyFontSize: 8.5,
    headerFill: C.primary,
  });

  // Key hires callout
  const khX = MARGIN + 4.3 + 0.2;
  const khY = hcY;
  const khW = SLIDE_W - khX - MARGIN;
  addAccentBar(slide, khX, khY, 0.22 + 8 * 0.268, C.green);
  addSectionLabel(slide, khX, khY + 0.04, khW, "Key Q3 Hires", C.green);

  const keyHires = [
    { role: "VP Customer Success", detail: "ex-Datadog — heads $700K churn recovery plan" },
    { role: "Head of Partnerships", detail: "ex-Siemens — leads MindSphere integration" },
    { role: "2× Senior ML Engineers", detail: "PhD, Georgia Tech — Predictive Maintenance v3" },
  ];

  keyHires.forEach((kh, i) => {
    const ky = khY + 0.3 + i * 0.72;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: khX + 0.07, y: ky, w: khW - 0.08, h: 0.64,
      fill: { color: C.white },
      line: { color: C.cardBorder, width: 0.5 },
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x: khX + 0.07, y: ky, w: 0.045, h: 0.64,
      fill: { color: C.green },
      line: { color: C.green, width: 0 },
    });
    slide.addText(kh.role, {
      x: khX + 0.17, y: ky + 0.08, w: khW - 0.25, h: 0.22,
      fontSize: 9.5,
      fontFace: FONT.sub,
      color: C.primary,
      bold: true,
      align: "left",
      margin: 0,
    });
    slide.addText(kh.detail, {
      x: khX + 0.17, y: ky + 0.32, w: khW - 0.25, h: 0.26,
      fontSize: 8.5,
      fontFace: FONT.body,
      color: C.slate,
      align: "left",
      margin: 0,
    });
  });
}

// ============================================================
// SLIDE 5 — Financial Outlook + Competitive (light)
// ============================================================
{
  const slide = addLightSlide();
  addLightHeader(slide, "Financial Outlook + Competitive Position", "P&L Summary, Cash Position & Market Landscape");
  addFooter(slide, 5, TOTAL_SLIDES, false);

  const topY = 0.88;

  // P&L TABLE
  addAccentBar(slide, MARGIN, topY, 0.22 + 8 * 0.28, C.accent);
  addSectionLabel(slide, MARGIN, topY + 0.04, 5.6, "P&L Summary ($K)", C.accent);

  const plCols = [1.65, 0.7, 0.7, 0.7, 0.72, 0.93];
  const plRows = [
    ["Line Item", "Q1", "Q2", "Q3", "Q3 YoY", "FY2026E"],
    ["Revenue",      "$3,400", "$3,700", "$4,100", { text: "+64%", opts: { color: C.green, bold: true } }, "$15,400"],
    ["COGS",         "($740)", "($830)", "($890)", { text: "+52%", opts: { color: C.slate } }, "($3,340)"],
    ["Gross Profit", "$2,660", "$2,870", "$3,210", { text: "+68%", opts: { color: C.green, bold: true } }, "$12,060"],
    ["S&M",          "($1,420)", "($1,580)", "($1,720)", { text: "+48%", opts: { color: C.slate } }, "($6,480)"],
    ["R&D",          "($1,340)", "($1,480)", "($1,620)", { text: "+55%", opts: { color: C.slate } }, "($6,040)"],
    [
      { text: "Net Income", opts: { bold: true } },
      { text: "($580)", opts: { color: C.red } },
      { text: "($700)", opts: { color: C.red } },
      { text: "($670)", opts: { color: C.amber } },
      { text: "improved", opts: { color: C.green, bold: true } },
      { text: "($2,520)", opts: { color: C.red } }
    ],
  ];
  addStyledTable(slide, plRows, MARGIN + 0.08, topY + 0.27, 5.4, {
    colW: plCols,
    rowH: 0.28,
    bodyFontSize: 8.5,
    headerFill: C.primary,
  });

  // Cash position callouts
  const cashY = topY + 0.27 + 8 * 0.28 + 0.1;
  const cashStats = [
    { val: "$18.2M", lbl: "Cash on Hand" },
    { val: "$223K", lbl: "Monthly Burn" },
    { val: "81 mo", lbl: "Runway" },
    { val: "0.36x", lbl: "Burn Multiple" },
  ];
  const cashW = 1.26;
  const cashH = 0.5;
  cashStats.forEach((s, i) => {
    addKpiCard(slide, MARGIN + i * (cashW + 0.06), cashY, cashW, cashH, s.val, s.lbl, null, {
      accentColor: C.green,
      valueFontSize: 14,
    });
  });

  // COMPETITIVE TABLE
  const compX = MARGIN + 5.4 + 0.22;
  const compW = SLIDE_W - compX - MARGIN;
  addAccentBar(slide, compX, topY, 0.22 + 6 * 0.32, C.highlight);
  addSectionLabel(slide, compX, topY + 0.04, compW, "Competitive Landscape", C.highlight);

  const compCols = [1.38, 0.75, 0.75, 0.75, 0.75];
  const compRows = [
    ["", "NovaCrest", "DataForge", "Acme Analytics", "Zenith AI"],
    ["ARR",       "$15.6M", "~$45M",  "~$180M",  "~$4M"],
    ["Stage",     "Series B", "Series C", "Public", "Series A"],
    ["Target",    "Mid-mkt mfg", "SMB-Mid", "Enterprise", "Mid-mkt mfg"],
    [
      "Win Rate vs",
      "—",
      { text: "62% (depth)", opts: { color: C.amber } },
      { text: "34% (brand)", opts: { color: C.green } },
      { text: "55% (threat)", opts: { color: C.red } }
    ],
    [
      "Key Threat",
      "—",
      "Mfg module $199/mo",
      "Acme Lite Q1'27",
      "VC hiring spree"
    ],
  ];
  addStyledTable(slide, compRows, compX + 0.08, topY + 0.27, compW - 0.08, {
    colW: compCols,
    rowH: 0.32,
    bodyFontSize: 8,
    headerFill: C.accent,
  });

  // Win rate legend note
  slide.addText("Win rate = our win rate when competing head-to-head. Zenith AI is a growing threat in same ICP.", {
    x: compX + 0.08,
    y: topY + 0.27 + 7 * 0.32 + 0.05,
    w: compW - 0.12,
    h: 0.22,
    fontSize: 7.5,
    fontFace: FONT.body,
    color: C.slate,
    align: "left",
    margin: 0,
    italic: true,
  });
}

// ============================================================
// SLIDE 6 — Board Asks (dark)
// ============================================================
{
  const slide = addDarkSlide();
  addCornerAccents(slide);

  // Header
  slide.addText("Board Asks", {
    x: MARGIN,
    y: 0.3,
    w: 6,
    h: 0.56,
    fontSize: 32,
    fontFace: FONT.head,
    color: C.white,
    bold: true,
    align: "left",
    margin: 0,
    charSpacing: 1,
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: MARGIN, y: 0.88, w: 2.4, h: 0.045,
    fill: { color: C.highlight },
    line: { color: C.highlight, width: 0 },
  });
  slide.addText("October 2026  |  NovaCrest Board of Directors", {
    x: MARGIN,
    y: 0.94,
    w: 6,
    h: 0.24,
    fontSize: 9,
    fontFace: FONT.body,
    color: C.lightGray,
    align: "left",
    margin: 0,
  });

  // Ask blocks
  const asks = [
    {
      action: "APPROVE",
      actionColor: C.green,
      title: "$2.5M CS Investment — 8 FTEs",
      detail: "4 onboarding specialists  |  2 renewal managers  |  1 SMB CS lead  |  1 CS ops",
      roi: "Reduce preventable churn 60% → save ~$700K ARR/yr  |  Payback: 14 months",
      risk: "Risk of inaction: Q4 churn forecast $600K+ based on current at-risk pipeline",
    },
    {
      action: "APPROVE",
      actionColor: C.green,
      title: "SMB Usage-Based Pricing Tier — $499/mo entry",
      detail: "vs current $2.5K/mo minimum — 60+ prospects/qtr currently lost on price",
      roi: "+$1.2M new ARR in 12 months  |  Down-tier risk ~$200K (net positive)",
      risk: "Solves 58% of preventable churn via down-tier option; no implementation risk",
    },
    {
      action: "DISCUSS",
      actionColor: C.amber,
      title: "Series C Timing — 81-month runway, grow from strength",
      detail: "Option A: Q2 2027 at $22–25M ARR  |  Option B: Wait for $30M ARR milestone",
      roi: "Raise size: $40–60M  |  Board input needed: timing, size, investor targets",
      risk: "Accelerating growth may compress optionality — raise from strength preferred",
    },
  ];

  const askStartY = 1.28;
  const askH = 1.28;
  const askGap = 0.12;

  asks.forEach((ask, i) => {
    const ay = askStartY + i * (askH + askGap);

    // Card background
    slide.addShape(pres.shapes.RECTANGLE, {
      x: MARGIN, y: ay, w: CONTENT_W, h: askH,
      fill: { color: C.primary },
      line: { color: C.accent, width: 0.5 },
    });

    // Left accent bar
    slide.addShape(pres.shapes.RECTANGLE, {
      x: MARGIN, y: ay, w: 0.06, h: askH,
      fill: { color: ask.actionColor },
      line: { color: ask.actionColor, width: 0 },
    });

    // Action badge
    slide.addShape(pres.shapes.RECTANGLE, {
      x: MARGIN + 0.14, y: ay + 0.12, w: 0.88, h: 0.3,
      fill: { color: ask.actionColor },
      line: { color: ask.actionColor, width: 0 },
    });
    slide.addText(ask.action, {
      x: MARGIN + 0.14, y: ay + 0.12, w: 0.88, h: 0.3,
      fontSize: 9.5,
      fontFace: FONT.head,
      color: C.white,
      bold: true,
      align: "center",
      valign: "middle",
      margin: 0,
    });

    // Title
    slide.addText(ask.title, {
      x: MARGIN + 1.1, y: ay + 0.12, w: CONTENT_W - 1.18, h: 0.3,
      fontSize: 13,
      fontFace: FONT.sub,
      color: C.white,
      bold: true,
      align: "left",
      valign: "middle",
      margin: 0,
    });

    // Detail
    slide.addText(ask.detail, {
      x: MARGIN + 0.14, y: ay + 0.48, w: CONTENT_W - 0.2, h: 0.22,
      fontSize: 9,
      fontFace: FONT.body,
      color: C.lightGray,
      align: "left",
      margin: 0,
    });

    // ROI line
    slide.addShape(pres.shapes.RECTANGLE, {
      x: MARGIN + 0.14, y: ay + 0.73, w: 0.06, h: 0.2,
      fill: { color: C.green },
      line: { color: C.green, width: 0 },
    });
    slide.addText(`ROI: ${ask.roi}`, {
      x: MARGIN + 0.26, y: ay + 0.73, w: CONTENT_W - 0.32, h: 0.22,
      fontSize: 8.5,
      fontFace: FONT.body,
      color: C.lightGray,
      align: "left",
      margin: 0,
    });

    // Risk/note line
    slide.addShape(pres.shapes.RECTANGLE, {
      x: MARGIN + 0.14, y: ay + 0.97, w: 0.06, h: 0.2,
      fill: { color: C.amber },
      line: { color: C.amber, width: 0 },
    });
    slide.addText(ask.risk, {
      x: MARGIN + 0.26, y: ay + 0.97, w: CONTENT_W - 0.32, h: 0.22,
      fontSize: 8.5,
      fontFace: FONT.body,
      color: C.lightGray,
      align: "left",
      margin: 0,
      italic: true,
    });
  });

  addFooter(slide, 6, TOTAL_SLIDES, true);
}

// ============================================================
// SAVE
// ============================================================
pres.writeFile({ fileName: "outputs/strategy-anthropic.pptx" })
  .then(() => console.log("Saved: outputs/strategy-anthropic.pptx"))
  .catch((err) => { console.error("Error:", err); process.exit(1); });
