const pptxgen = require("pptxgenjs");

// ── Theme: Midnight Navy (board/strategy) ──
const theme = {
  primary: "1A2744",
  secondary: "162050",
  accent: "0EA5E9",
  light: "E0F2FE",
  bg: "0A0F1E",
};

const TITLE_FONT = "Trebuchet MS";
const BODY_FONT = "Calibri";

// ── Helpers ──

function pageBadge(slide, pres, num) {
  slide.addShape(pres.shapes.OVAL, {
    x: 9.3, y: 5.1, w: 0.4, h: 0.4,
    fill: { color: theme.secondary },
  });
  slide.addText(String(num), {
    x: 9.3, y: 5.1, w: 0.4, h: 0.4,
    fontSize: 12, fontFace: BODY_FONT,
    color: theme.accent, bold: true,
    align: "center", valign: "middle",
  });
}

function sectionLabel(slide, text, x, y, w, darkBg) {
  slide.addShape("rect", {
    x, y, w, h: 0.03,
    fill: { color: theme.accent },
  });
  slide.addText(text.toUpperCase(), {
    x, y: y + 0.06, w, h: 0.28,
    fontSize: 9, fontFace: BODY_FONT,
    color: darkBg ? theme.accent : theme.primary,
    bold: true,
    charSpacing: 3,
  });
}

function makeTableOpts(x, y, w, colW, opts = {}) {
  return {
    x, y, w, colW,
    fontSize: 9, fontFace: BODY_FONT,
    color: opts.lightBg ? "1A2744" : "D8E8F4",
    border: { type: "solid", pt: 0.5, color: opts.lightBg ? "AACCDD" : "1a3a5c" },
    rowH: opts.rowH || 0.27,
    autoPage: false,
    ...opts,
  };
}

function headerRow(cells, lightBg) {
  return cells.map((c) => ({
    text: c,
    options: {
      bold: true,
      color: lightBg ? "FFFFFF" : "0A0F1E",
      fill: { color: theme.accent },
      fontSize: 9,
      fontFace: BODY_FONT,
      align: "center",
      valign: "middle",
    },
  }));
}

function dataRow(cells, opts = {}) {
  const lightBg = opts.lightBg;
  const altDark  = opts.alt ? "0D1A30" : "0A1425";
  const altLight = opts.alt ? "EBF5FA" : "F5FBFE";
  return cells.map((c, i) => {
    const isObj = typeof c === "object" && c !== null && c.text !== undefined;
    const text = isObj ? c.text : String(c);
    const cellOpts = isObj ? (c.options || {}) : {};
    return {
      text,
      options: {
        fill: { color: lightBg ? altLight : altDark },
        fontSize: 9,
        fontFace: BODY_FONT,
        color: lightBg ? "1A2744" : "C8DCF0",
        valign: "middle",
        align: i === 0 ? "left" : "center",
        ...cellOpts,
      },
    };
  });
}

function calloutBox(slide, x, y, w, h, title, body, darkBg) {
  slide.addShape("rect", {
    x, y, w, h,
    fill: { color: darkBg ? "0D1E3A" : "EBF5FA" },
    line: { color: theme.accent, pt: 1.5 },
  });
  if (title) {
    slide.addText(title, {
      x: x + 0.12, y: y + 0.07, w: w - 0.24, h: 0.22,
      fontSize: 9, fontFace: BODY_FONT,
      color: theme.accent, bold: true,
    });
  }
  slide.addText(body, {
    x: x + 0.12, y: y + (title ? 0.26 : 0.08), w: w - 0.24, h: h - (title ? 0.32 : 0.16),
    fontSize: 8.5, fontFace: BODY_FONT,
    color: darkBg ? "C8DCF0" : "1A2744",
    valign: "top",
    wrap: true,
  });
}

// ════════════════════════════════════════════════════════════════
// SLIDE 1: Title — dark bg, no badge
// ════════════════════════════════════════════════════════════════
function slide01(pres) {
  const slide = pres.addSlide();
  slide.background = { color: theme.bg };

  // Top accent bar
  slide.addShape("rect", {
    x: 0, y: 0, w: 10, h: 0.07,
    fill: { color: theme.accent },
  });

  // Left decorative stripe
  slide.addShape("rect", {
    x: 0, y: 0.07, w: 0.06, h: 5.43,
    fill: { color: theme.secondary },
  });

  // Company label
  slide.addText("NOVACREST", {
    x: 0.35, y: 0.25, w: 6, h: 0.35,
    fontSize: 11, fontFace: BODY_FONT,
    color: theme.accent, bold: true,
    charSpacing: 5,
  });

  // Main title
  slide.addText("Q3 2026\nStrategic Review", {
    x: 0.35, y: 0.55, w: 6.2, h: 1.6,
    fontSize: 40, fontFace: TITLE_FONT,
    color: "FFFFFF", bold: true,
    lineSpacingMultiple: 0.95,
  });

  // Subtitle
  slide.addText("Series B  ·  $15.6M ARR  ·  118 Employees  |  CONFIDENTIAL", {
    x: 0.35, y: 2.2, w: 7, h: 0.32,
    fontSize: 11, fontFace: BODY_FONT,
    color: "7FA8C4",
  });

  // Divider line
  slide.addShape("rect", {
    x: 0.35, y: 2.58, w: 9.3, h: 0.02,
    fill: { color: theme.secondary },
  });

  // 3 KPI mini cards
  const kpis = [
    { label: "ARR", value: "$15.6M", sub: "Q3 2026" },
    { label: "NRR", value: "118%", sub: "Net Revenue Retention" },
    { label: "Runway", value: "81 mo", sub: "$18.2M cash" },
  ];

  kpis.forEach((k, i) => {
    const cx = 0.35 + i * 3.15;
    slide.addShape("rect", {
      x: cx, y: 2.72, w: 3.0, h: 1.1,
      fill: { color: theme.primary },
      line: { color: theme.accent, pt: 0.8 },
    });
    slide.addText(k.value, {
      x: cx + 0.1, y: 2.78, w: 2.8, h: 0.6,
      fontSize: 30, fontFace: TITLE_FONT,
      color: theme.accent, bold: true,
      align: "center", valign: "middle",
    });
    slide.addText(k.label, {
      x: cx + 0.1, y: 3.33, w: 2.8, h: 0.22,
      fontSize: 9, fontFace: BODY_FONT,
      color: "FFFFFF", bold: true,
      align: "center",
    });
    slide.addText(k.sub, {
      x: cx + 0.1, y: 3.54, w: 2.8, h: 0.2,
      fontSize: 7.5, fontFace: BODY_FONT,
      color: "6A8FAB", align: "center",
    });
  });

  // Board members footer
  slide.addShape("rect", {
    x: 0, y: 5.0, w: 10, h: 0.5,
    fill: { color: theme.secondary },
  });
  slide.addText("Board:  Sarah Chen (Sequoia, Lead)  ·  David Park (Andreessen Horowitz)  ·  Marcus Webb (Independent)  ·  Priya Nair (Founder/CEO)  ·  James Liu (Founder/CTO)", {
    x: 0.3, y: 5.05, w: 9.4, h: 0.35,
    fontSize: 8, fontFace: BODY_FONT,
    color: "8AAEC8", align: "center", valign: "middle",
  });
}

// ════════════════════════════════════════════════════════════════
// SLIDE 2: Revenue Dashboard — dark bg
// ════════════════════════════════════════════════════════════════
function slide02(pres) {
  const slide = pres.addSlide();
  slide.background = { color: theme.bg };

  // Top bar
  slide.addShape("rect", { x: 0, y: 0, w: 10, h: 0.07, fill: { color: theme.accent } });

  // Slide title
  slide.addText("Revenue Dashboard", {
    x: 0.35, y: 0.12, w: 7, h: 0.42,
    fontSize: 24, fontFace: TITLE_FONT,
    color: "FFFFFF", bold: true,
  });

  sectionLabel(slide, "FINANCIAL PERFORMANCE", 0.35, 0.58, 9.3, true);

  // 4 KPI stat cards
  const kpis = [
    { label: "ARR", value: "$15.6M", sub: "+10.6% QoQ" },
    { label: "NRR", value: "118%", sub: "Target: 120%+" },
    { label: "Gross Margin", value: "78.2%", sub: "Target: 80%+" },
    { label: "Magic Number", value: "1.12", sub: "Target: >0.75  ✓" },
  ];

  kpis.forEach((k, i) => {
    const cx = 0.2 + i * 2.44;
    slide.addShape("rect", {
      x: cx, y: 0.95, w: 2.28, h: 0.95,
      fill: { color: theme.primary },
      line: { color: theme.accent, pt: 0.8 },
    });
    slide.addText(k.value, {
      x: cx + 0.08, y: 0.98, w: 2.12, h: 0.52,
      fontSize: 22, fontFace: TITLE_FONT,
      color: theme.accent, bold: true,
      align: "center", valign: "middle",
    });
    slide.addText(k.label, {
      x: cx + 0.08, y: 1.48, w: 2.12, h: 0.2,
      fontSize: 8.5, fontFace: BODY_FONT,
      color: "FFFFFF", bold: true, align: "center",
    });
    slide.addText(k.sub, {
      x: cx + 0.08, y: 1.67, w: 2.12, h: 0.18,
      fontSize: 7.5, fontFace: BODY_FONT,
      color: "6A8FAB", align: "center",
    });
  });

  // ARR Trend section
  sectionLabel(slide, "ARR GROWTH TREND", 0.2, 2.0, 4.5, true);

  const arrRows = [
    headerRow(["Quarter", "ARR"], false),
    dataRow(["Q1 '25", "$7.2M"]),
    dataRow(["Q2 '25", "$8.5M"], { alt: true }),
    dataRow(["Q3 '25", "$9.8M"]),
    dataRow(["Q4 '25", "$11.4M"], { alt: true }),
    dataRow(["Q1 '26", "$12.8M"]),
    dataRow(["Q2 '26", "$14.1M"], { alt: true }),
    dataRow(["Q3 '26", { text: "$15.6M", options: { color: theme.accent, bold: true } }]),
  ];

  slide.addTable(arrRows, makeTableOpts(0.2, 2.36, 3.0, [1.5, 1.5], { rowH: 0.27 }));

  // ARR bar chart visual
  const barData = [7.2, 8.5, 9.8, 11.4, 12.8, 14.1, 15.6];
  const quarters = ["Q1'25","Q2'25","Q3'25","Q4'25","Q1'26","Q2'26","Q3'26"];
  const maxVal = 15.6;
  const barX = 3.35, barY = 2.36, barW = 1.28, maxH = 2.0;

  barData.forEach((v, i) => {
    const bh = (v / maxVal) * maxH;
    const bx = barX + i * (barW / barData.length + 0.02);
    const bw = barW / barData.length - 0.02;
    const by = barY + maxH - bh;
    const isLast = i === barData.length - 1;
    slide.addShape("rect", {
      x: bx, y: by, w: bw, h: bh,
      fill: { color: isLast ? theme.accent : "1E4A78" },
    });
    slide.addText(quarters[i], {
      x: bx - 0.02, y: barY + maxH + 0.04, w: bw + 0.04, h: 0.2,
      fontSize: 6.5, fontFace: BODY_FONT,
      color: "7FA8C4", align: "center",
    });
  });

  // Unit economics table — right side
  sectionLabel(slide, "UNIT ECONOMICS", 4.8, 2.0, 4.8, true);

  const ueRows = [
    headerRow(["Metric", "Q3 2026", "Q2 2026", "Target", "Trend"], false),
    dataRow(["Gross Margin", "78.2%", "77.5%", "80%+", { text: "↑", options: { color: "#22C55E", bold: true } }]),
    dataRow(["NRR", "118%", "121%", "120%+", { text: "↓ churn", options: { color: "#F59E0B", bold: true } }], { alt: true }),
    dataRow(["LTV:CAC", "4.8x", "4.6x", "4.0x+", { text: "Healthy", options: { color: "#22C55E" } }]),
    dataRow(["CAC Payback", "13.2 mo", "14.1 mo", "<16 mo", { text: "↑", options: { color: "#22C55E", bold: true } }], { alt: true }),
    dataRow(["Magic Number", "1.12", "0.98", ">0.75", { text: "Strong", options: { color: "#22C55E" } }]),
  ];

  slide.addTable(ueRows, makeTableOpts(4.8, 2.36, 4.85, [1.6, 0.9, 0.9, 0.8, 0.65], { rowH: 0.31 }));

  // Cash callout
  calloutBox(slide, 0.2, 4.42, 3.1, 0.9,
    "CASH & BURN",
    "$18.2M cash  ·  $223K/mo burn  ·  81 mo runway  ·  0.36x burn multiple",
    true);

  // Insight callout
  calloutBox(slide, 3.45, 4.42, 6.2, 0.9,
    "KEY INSIGHT",
    "Magic Number jumped from 0.98 → 1.12 QoQ — sales efficiency is accelerating. NRR dipped 3pp due to Q3 churn cluster ($500K); see slide 3 for churn analysis and remediation plan.",
    true);

  pageBadge(slide, pres, 2);
}

// ════════════════════════════════════════════════════════════════
// SLIDE 3: GTM + Customer Health — light bg, two columns
// ════════════════════════════════════════════════════════════════
function slide03(pres) {
  const slide = pres.addSlide();
  slide.background = { color: "F0F7FB" };

  // Top bar
  slide.addShape("rect", { x: 0, y: 0, w: 10, h: 0.07, fill: { color: theme.accent } });

  slide.addText("GTM & Customer Health", {
    x: 0.35, y: 0.12, w: 7, h: 0.40,
    fontSize: 24, fontFace: TITLE_FONT,
    color: theme.primary, bold: true,
  });

  // Column divider
  slide.addShape("rect", {
    x: 5.0, y: 0.55, w: 0.02, h: 4.95,
    fill: { color: "C0D8E8" },
  });

  // ─ LEFT COLUMN: Sales metrics ─
  sectionLabel(slide, "GO-TO-MARKET", 0.25, 0.58, 4.6, false);

  const salesRows = [
    headerRow(["Metric", "Q3", "Q2", "QoQ"], true),
    dataRow(["Pipeline", "$8.6M", "$7.1M", { text: "+21%", options: { color: "#16A34A", bold: true } }], { lightBg: true }),
    dataRow(["Win Rate", "28.4%", "26.1%", { text: "+2.3pp", options: { color: "#16A34A" } }], { alt: true, lightBg: true }),
    dataRow(["Avg ACV", "$42K", "$38K", { text: "+10.5%", options: { color: "#16A34A" } }], { lightBg: true }),
    dataRow(["Sales Cycle", "68d", "72d", { text: "-5.6%", options: { color: "#16A34A" } }], { alt: true, lightBg: true }),
    dataRow(["Quota Attain.", "112%", "94%", { text: "+18pp", options: { color: "#16A34A", bold: true } }], { lightBg: true }),
  ];

  slide.addTable(salesRows, makeTableOpts(0.25, 0.96, 4.6, [1.45, 1.05, 1.05, 1.05], { rowH: 0.28, lightBg: true }));

  sectionLabel(slide, "CHANNEL MIX", 0.25, 2.68, 4.6, false);

  const chanRows = [
    headerRow(["Channel", "Pipeline", "Deals", "ACV", "CAC"], true),
    dataRow(["SDR", "$3.8M", "18", "$52K", "$24.1K"], { lightBg: true }),
    dataRow(["Inbound", "$2.4M", "14", "$34K", "$12.8K"], { alt: true, lightBg: true }),
    dataRow(["PLG", "$0.9M", "8", "$18K", "$6.2K"], { lightBg: true }),
    dataRow(["Partner", "$1.5M", "5", "$68K", "$8.4K"], { alt: true, lightBg: true }),
  ];

  slide.addTable(chanRows, makeTableOpts(0.25, 3.04, 4.6, [0.9, 1.0, 0.7, 0.9, 1.1], { rowH: 0.28, lightBg: true }));

  // ─ RIGHT COLUMN: Churn analysis ─
  sectionLabel(slide, "CHURN DEEP DIVE — Q3", 5.15, 0.58, 4.6, false);

  // Churn summary line
  slide.addText("$500K total churn (up from $300K Q2) — requires immediate response", {
    x: 5.15, y: 0.96, w: 4.6, h: 0.26,
    fontSize: 8.5, fontFace: BODY_FONT,
    color: "#B45309", bold: true,
  });

  const churnRows = [
    headerRow(["Account", "ARR", "Preventable?"], true),
    dataRow(["Apex Mfg", "$145K", "No (M&A)"], { lightBg: true }),
    dataRow(["Precision Dynamics", "$62K", "Partially"], { alt: true, lightBg: true }),
    dataRow(["4 SMB accts", "$118K", { text: "Yes — price", options: { color: "#DC2626" } }], { lightBg: true }),
    dataRow(["TechFab Solutions", "$48K", { text: "Yes — CS failure", options: { color: "#DC2626" } }], { alt: true, lightBg: true }),
    dataRow(["3 SMB accts", "$82K", { text: "Yes — onboarding", options: { color: "#DC2626" } }], { lightBg: true }),
    dataRow(["Consolidated Parts", "$45K", "No — budget"], { alt: true, lightBg: true }),
  ];

  slide.addTable(churnRows, makeTableOpts(5.15, 1.26, 4.6, [1.85, 0.95, 1.8], { rowH: 0.27, lightBg: true }));

  // Callout: preventable
  calloutBox(slide, 5.15, 3.2, 4.6, 0.45,
    null,
    "$290K / 58% preventable — CS capacity + onboarding gap are primary drivers",
    false);

  // At-risk Q4 section
  sectionLabel(slide, "AT-RISK Q4 ACCOUNTS", 5.15, 3.74, 4.6, false);

  const atRiskRows = [
    headerRow(["Account", "ARR", "Risk Factor"], true),
    dataRow(["Sterling Industries", "$210K", "Champion left"], { lightBg: true }),
    dataRow(["Midwest Components", "$58K", "Usage drop -40%"], { alt: true, lightBg: true }),
    dataRow(["ClearPath Logistics", "$72K", "Renewal + POC needed"], { lightBg: true }),
  ];

  slide.addTable(atRiskRows, makeTableOpts(5.15, 4.08, 4.6, [1.85, 0.85, 1.9], { rowH: 0.27, lightBg: true }));

  pageBadge(slide, pres, 3);
}

// ════════════════════════════════════════════════════════════════
// SLIDE 4: Product + Team — light bg
// ════════════════════════════════════════════════════════════════
function slide04(pres) {
  const slide = pres.addSlide();
  slide.background = { color: "F0F7FB" };

  slide.addShape("rect", { x: 0, y: 0, w: 10, h: 0.07, fill: { color: theme.accent } });

  slide.addText("Product & Team", {
    x: 0.35, y: 0.12, w: 7, h: 0.40,
    fontSize: 24, fontFace: TITLE_FONT,
    color: theme.primary, bold: true,
  });

  sectionLabel(slide, "PRODUCT & ENGINEERING", 0.25, 0.58, 9.5, false);

  // Q3 shipped features — left
  sectionLabel(slide, "SHIPPED Q3", 0.25, 0.97, 4.55, false);

  const shippedRows = [
    headerRow(["Feature", "Adoption", "Revenue Impact"], true),
    dataRow(["Predictive Maintenance v2", "67% enterprise", "$1.2M pipeline"], { lightBg: true }),
    dataRow(["Self-Serve Dashboards", "340 dashboards", "CS tickets -18%"], { alt: true, lightBg: true }),
    dataRow(["SAP Integration", "12 accounts", "Partner accelerant"], { lightBg: true }),
    dataRow(["SOC 2 Type II", "Cert 8/15", { text: "Unblocked $380K", options: { color: "#16A34A", bold: true } }], { alt: true, lightBg: true }),
  ];

  slide.addTable(shippedRows, makeTableOpts(0.25, 1.3, 4.55, [1.9, 1.3, 1.35], { rowH: 0.29, lightBg: true }));

  // Q4 roadmap — right
  sectionLabel(slide, "Q4 ROADMAP", 5.05, 0.97, 4.55, false);

  const roadmapRows = [
    headerRow(["P", "Feature", "Status", "Impact"], true),
    dataRow([{ text: "P0", options: { color: "#DC2626", bold: true } }, "Multi-tenant analytics", "60% dev", "2 deals $200K+"], { lightBg: true }),
    dataRow([{ text: "P0", options: { color: "#DC2626", bold: true } }, "Siemens MindSphere", "Dev starting", "$4M TAM"], { alt: true, lightBg: true }),
    dataRow([{ text: "P1", options: { color: "#D97706", bold: true } }, "Usage-based pricing", "Spec done", "SMB churn fix"], { lightBg: true }),
    dataRow([{ text: "P1", options: { color: "#D97706", bold: true } }, "Customer health score", "Prototype", "Early warning"], { alt: true, lightBg: true }),
  ];

  slide.addTable(roadmapRows, makeTableOpts(5.05, 1.3, 4.55, [0.35, 1.85, 1.1, 1.25], { rowH: 0.29, lightBg: true }));

  // Headcount table
  sectionLabel(slide, "HEADCOUNT", 0.25, 2.82, 9.5, false);

  const hcRows = [
    headerRow(["Function", "Q2", "Q3", "Open", "Q4 Target"], true),
    dataRow(["Engineering", "42", "46", "4", "50"], { lightBg: true }),
    dataRow(["Sales", "18", "22", "3", "25"], { alt: true, lightBg: true }),
    dataRow(["Customer Success", "12", "14", "2", "16"], { lightBg: true }),
    dataRow(["All Others", "26", "30", "3", "33"], { alt: true, lightBg: true }),
    dataRow([{ text: "Total", options: { bold: true } }, { text: "98", options: { bold: true } }, { text: "112", options: { bold: true } }, { text: "12", options: { bold: true } }, { text: "124", options: { bold: true, color: theme.accent } }], { lightBg: true }),
  ];

  slide.addTable(hcRows, makeTableOpts(0.25, 3.15, 5.0, [1.85, 0.8, 0.8, 0.75, 0.8], { rowH: 0.27, lightBg: true }));

  // Key Q3 hires callout — right of headcount
  slide.addShape("rect", {
    x: 5.45, y: 3.15, w: 4.2, h: 1.7,
    fill: { color: "EBF5FA" },
    line: { color: theme.accent, pt: 1.5 },
  });
  slide.addText("KEY Q3 HIRES", {
    x: 5.58, y: 3.22, w: 4.0, h: 0.22,
    fontSize: 8.5, fontFace: BODY_FONT,
    color: theme.accent, bold: true, charSpacing: 2,
  });

  const hires = [
    "VP Customer Success — ex-Datadog (scaled CS org 20→80), starts Q4",
    "Head of Partnerships — ex-Siemens (SAP + Siemens channel)",
    "2× Senior ML Engineers — PhD, Georgia Tech (Predictive Maintenance roadmap)",
  ];

  hires.forEach((h, i) => {
    slide.addText("▸  " + h, {
      x: 5.58, y: 3.48 + i * 0.38, w: 4.0, h: 0.35,
      fontSize: 8.5, fontFace: BODY_FONT,
      color: "1A2744",
      wrap: true,
    });
  });

  // Net headcount growth callout
  calloutBox(slide, 0.25, 4.9, 9.4, 0.42,
    null,
    "Net headcount: 98 → 112 (+14 in Q3). 12 open roles funded. Largest growth: Engineering (+4) and Sales (+4). CS investment is the board ask on slide 6.",
    false);

  pageBadge(slide, pres, 4);
}

// ════════════════════════════════════════════════════════════════
// SLIDE 5: Financial Outlook + Competitive — light bg
// ════════════════════════════════════════════════════════════════
function slide05(pres) {
  const slide = pres.addSlide();
  slide.background = { color: "F0F7FB" };

  slide.addShape("rect", { x: 0, y: 0, w: 10, h: 0.07, fill: { color: theme.accent } });

  slide.addText("Financials & Competitive", {
    x: 0.35, y: 0.12, w: 7, h: 0.40,
    fontSize: 24, fontFace: TITLE_FONT,
    color: theme.primary, bold: true,
  });

  sectionLabel(slide, "FINANCIALS & COMPETITIVE", 0.25, 0.58, 9.5, false);

  // P&L table
  sectionLabel(slide, "P&L SUMMARY", 0.25, 0.97, 9.5, false);

  const plRows = [
    headerRow(["Line Item", "Q1", "Q2", "Q3", "Q3 YoY", "FY2026E"], true),
    dataRow(["Revenue", "$3.4M", "$3.7M", { text: "$4.1M", options: { bold: true } }, { text: "+64%", options: { color: "#16A34A", bold: true } }, "$15.4M"], { lightBg: true }),
    dataRow(["Gross Profit", "$2.66M", "$2.87M", "$3.21M", { text: "+68%", options: { color: "#16A34A" } }, "$12.06M"], { alt: true, lightBg: true }),
    dataRow(["Gross Margin %", "78.2%", "77.6%", "78.3%", "+2.1pp", "78.3%"], { lightBg: true }),
    dataRow(["S&M", "($1.42M)", "($1.58M)", "($1.72M)", "+48%", "($6.48M)"], { alt: true, lightBg: true }),
    dataRow(["R&D", "($1.34M)", "($1.48M)", "($1.62M)", "+55%", "($6.04M)"], { lightBg: true }),
    dataRow([{ text: "Net Income", options: { bold: true } }, "($580K)", "($700K)", { text: "($670K)", options: { bold: true } }, { text: "improved", options: { color: "#16A34A" } }, "($2.52M)"], { alt: true, lightBg: true }),
  ];

  slide.addTable(plRows, makeTableOpts(0.25, 1.3, 9.5, [1.85, 1.1, 1.1, 1.1, 1.1, 1.25], { rowH: 0.3, lightBg: true }));

  // Cash metrics callout row
  const cashMetrics = [
    { label: "Cash", value: "$18.2M" },
    { label: "Monthly Burn", value: "$223K" },
    { label: "Runway", value: "81 months" },
    { label: "Burn Multiple", value: "0.36x" },
  ];

  cashMetrics.forEach((m, i) => {
    const cx = 0.25 + i * 2.38;
    slide.addShape("rect", {
      x: cx, y: 3.44, w: 2.2, h: 0.7,
      fill: { color: theme.primary },
      line: { color: theme.accent, pt: 0.8 },
    });
    slide.addText(m.value, {
      x: cx + 0.08, y: 3.48, w: 2.04, h: 0.38,
      fontSize: 18, fontFace: TITLE_FONT,
      color: theme.accent, bold: true,
      align: "center", valign: "middle",
    });
    slide.addText(m.label, {
      x: cx + 0.08, y: 3.84, w: 2.04, h: 0.22,
      fontSize: 8.5, fontFace: BODY_FONT,
      color: "FFFFFF", align: "center", bold: true,
    });
  });

  // Competitive table
  sectionLabel(slide, "COMPETITIVE WIN RATES", 0.25, 4.23, 9.5, false);

  const compRows = [
    headerRow(["Competitor", "Win Rate", "Key Threat", "Our Counter"], true),
    dataRow(["DataForge", { text: "62%", options: { color: "#16A34A", bold: true } }, "SMB price pressure", "Superior integration depth"], { lightBg: true }),
    dataRow(["Acme Analytics", { text: "34%", options: { color: "#DC2626", bold: true } }, "Enterprise Lite tier", "Manufacturing domain expertise"], { alt: true, lightBg: true }),
    dataRow(["Zenith AI", { text: "55%", options: { color: "#D97706", bold: true } }, "VC-funded hiring blitz", "SOC 2 certified + customer ROI"], { lightBg: true }),
  ];

  slide.addTable(compRows, makeTableOpts(0.25, 4.56, 9.5, [1.6, 1.0, 2.5, 2.8], { rowH: 0.28, lightBg: true }));

  pageBadge(slide, pres, 5);
}

// ════════════════════════════════════════════════════════════════
// SLIDE 6: Board Asks — dark bg
// ════════════════════════════════════════════════════════════════
function slide06(pres) {
  const slide = pres.addSlide();
  slide.background = { color: theme.bg };

  slide.addShape("rect", { x: 0, y: 0, w: 10, h: 0.07, fill: { color: theme.accent } });

  slide.addText("Board Asks — Q3 2026", {
    x: 0.35, y: 0.12, w: 7, h: 0.40,
    fontSize: 24, fontFace: TITLE_FONT,
    color: "FFFFFF", bold: true,
  });

  sectionLabel(slide, "BOARD ASKS — Q3 2026", 0.35, 0.58, 9.3, true);

  // 3 ask cards
  const asks = [
    {
      tag: "APPROVE",
      tagColor: "22C55E",
      title: "$2.5M CS Investment",
      subtitle: "8 FTEs over 12 months",
      bullets: [
        "VP CS (ex-Datadog) onboarding now — team build Q4-Q1",
        "Projected churn savings: $700K/yr based on Q3 analysis",
        "Payback period: 14 months. IRR > 80% at current NRR trajectory",
        "Immediate actions: Onboarding revamp + health score tooling",
        "Risk of inaction: $340K+ at-risk ARR in Q4 alone",
      ],
    },
    {
      tag: "APPROVE",
      tagColor: "22C55E",
      title: "SMB Usage-Based Pricing",
      subtitle: "$499/mo self-serve tier",
      bullets: [
        "PLG motion: 8 deals, $18K ACV, $6.2K CAC already proven",
        "Projected new ARR: +$1.2M in first 12 months post-launch",
        "Addresses 58% preventable churn from SMB price sensitivity",
        "Product spec complete — engineering 6 weeks to launch",
        "Net revenue positive: expansion > contraction modeled at 2:1",
      ],
    },
    {
      tag: "DISCUSS",
      tagColor: "F59E0B",
      title: "Series C Timing",
      subtitle: "Raise now vs. wait for $30M ARR",
      bullets: [
        "Current: 81mo runway, $15.6M ARR, accelerating growth (+10.6% QoQ)",
        "Option A: Raise Q2 2027 at $22-25M ARR — stronger SaaS comps",
        "Option B: Wait for $30M ARR — higher valuation, more dilution risk",
        "Key inputs: Siemens integration TAM, board investor relationships",
        "Board input requested: timing preference, target size, lead investor",
      ],
    },
  ];

  asks.forEach((ask, i) => {
    const cx = 0.22 + i * 3.25;
    const cardW = 3.1;

    // Card background
    slide.addShape("rect", {
      x: cx, y: 0.97, w: cardW, h: 4.28,
      fill: { color: theme.primary },
      line: { color: ask.tagColor, pt: 1.5 },
    });

    // Tag badge
    slide.addShape("rect", {
      x: cx + 0.15, y: 1.05, w: 0.85, h: 0.28,
      fill: { color: ask.tagColor },
    });
    slide.addText(ask.tag, {
      x: cx + 0.15, y: 1.05, w: 0.85, h: 0.28,
      fontSize: 8.5, fontFace: BODY_FONT,
      color: "000000", bold: true,
      align: "center", valign: "middle",
    });

    // Title
    slide.addText(ask.title, {
      x: cx + 0.12, y: 1.38, w: cardW - 0.24, h: 0.42,
      fontSize: 14, fontFace: TITLE_FONT,
      color: "FFFFFF", bold: true,
      lineSpacingMultiple: 0.95,
    });

    // Subtitle
    slide.addText(ask.subtitle, {
      x: cx + 0.12, y: 1.78, w: cardW - 0.24, h: 0.24,
      fontSize: 9, fontFace: BODY_FONT,
      color: theme.accent,
    });

    // Divider
    slide.addShape("rect", {
      x: cx + 0.12, y: 2.06, w: cardW - 0.24, h: 0.02,
      fill: { color: "1E3F6A" },
    });

    // Bullets
    ask.bullets.forEach((b, j) => {
      slide.addText("▸  " + b, {
        x: cx + 0.12, y: 2.14 + j * 0.42, w: cardW - 0.24, h: 0.4,
        fontSize: 8.5, fontFace: BODY_FONT,
        color: "C0D8F0",
        valign: "top",
        wrap: true,
      });
    });
  });

  // Bottom footer
  slide.addShape("rect", {
    x: 0, y: 5.27, w: 10, h: 0.23,
    fill: { color: theme.secondary },
  });
  slide.addText("NovaCrest Confidential  ·  Q3 2026 Board Meeting  ·  Prepared by Priya Nair, CEO", {
    x: 0.3, y: 5.27, w: 9.4, h: 0.23,
    fontSize: 7.5, fontFace: BODY_FONT,
    color: "6A8FAB", align: "center", valign: "middle",
  });

  pageBadge(slide, pres, 6);
}

// ════════════════════════════════════════════════════════════════
// MAIN
// ════════════════════════════════════════════════════════════════
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 10 x 5.63

slide01(pres);
slide02(pres);
slide03(pres);
slide04(pres);
slide05(pres);
slide06(pres);

pres.writeFile({ fileName: "outputs/strategy-minimax.pptx" })
  .then(() => console.log("✓ outputs/strategy-minimax.pptx written"))
  .catch((err) => { console.error(err); process.exit(1); });
