const pptxgen = require("pptxgenjs");

// ── Theme: Midnight Executive ──
const theme = {
  primary: "1E2761",
  secondary: "162050",
  accent: "3B82F6",
  light: "CADCFC",
  bg: "0F1535",
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

function sectionLabel(slide, text, x, y, w, dark = false) {
  slide.addShape("rect", {
    x: x, y: y, w: w, h: 0.03,
    fill: { color: theme.accent },
  });
  slide.addText(text.toUpperCase(), {
    x: x, y: y + 0.06, w: w, h: 0.28,
    fontSize: 9, fontFace: BODY_FONT,
    color: dark ? theme.light : theme.accent, bold: true,
    charSpacing: 3,
  });
}

function makeTableOpts(x, y, w, colW, opts = {}) {
  return {
    x, y, w, colW,
    fontSize: 9, fontFace: BODY_FONT,
    color: "E0E0E0",
    border: { type: "solid", pt: 0.5, color: "1a3a5c" },
    rowH: opts.rowH || 0.28,
    autoPage: false,
    ...opts,
  };
}

function makeTableOptsLight(x, y, w, colW, opts = {}) {
  return {
    x, y, w, colW,
    fontSize: 9, fontFace: BODY_FONT,
    color: "1E2761",
    border: { type: "solid", pt: 0.5, color: "CADCFC" },
    rowH: opts.rowH || 0.28,
    autoPage: false,
    ...opts,
  };
}

function headerRow(cells, dark = true) {
  return cells.map((c) => ({
    text: c,
    options: {
      bold: true,
      color: dark ? "FFFFFF" : "FFFFFF",
      fill: { color: theme.primary },
      fontSize: 9,
      fontFace: BODY_FONT,
      align: "center",
      valign: "middle",
    },
  }));
}

function dataRow(cells, opts = {}) {
  const isLight = opts.light || false;
  const altColor = isLight ? "EEF2FF" : "0a1030";
  const baseColor = isLight ? "F8FAFF" : "080d22";
  return cells.map((c, i) => {
    const isObj = typeof c === "object" && c !== null && c.text !== undefined;
    const text = isObj ? c.text : String(c);
    const cellOpts = isObj ? c.options || {} : {};
    return {
      text,
      options: {
        fill: { color: opts.alt ? altColor : baseColor },
        fontSize: 9,
        fontFace: BODY_FONT,
        color: isLight ? (opts.alt ? "1E2761" : "162050") : "D0D8E0",
        valign: "middle",
        align: i === 0 ? "left" : "center",
        ...cellOpts,
      },
    };
  });
}

function statCard(slide, x, y, w, h, value, label, sub, dark = true) {
  const bgColor = dark ? "152040" : "EEF2FF";
  const valColor = dark ? theme.accent : theme.primary;
  const labelColor = dark ? "CADCFC" : "1E2761";
  const subColor = dark ? "6080B0" : "5A6A94";
  slide.addShape("rect", {
    x, y, w, h,
    fill: { color: bgColor },
    rectRadius: 0.05,
  });
  slide.addText(value, {
    x, y: y + 0.05, w, h: h * 0.5,
    fontSize: 20, fontFace: TITLE_FONT,
    color: valColor, bold: true,
    align: "center", valign: "middle",
  });
  slide.addText(label, {
    x, y: y + h * 0.52, w, h: h * 0.26,
    fontSize: 9, fontFace: BODY_FONT,
    color: labelColor, bold: true,
    align: "center", valign: "middle",
  });
  if (sub) {
    slide.addText(sub, {
      x, y: y + h * 0.76, w, h: h * 0.22,
      fontSize: 8, fontFace: BODY_FONT,
      color: subColor,
      align: "center", valign: "middle",
    });
  }
}

function insightCard(slide, x, y, w, h, title, body, dark = false) {
  const bgColor = dark ? "1a2d50" : "EEF2FF";
  const titleColor = dark ? theme.accent : theme.primary;
  const bodyColor = dark ? "CADCFC" : "1E2761";
  slide.addShape("rect", {
    x, y, w, h,
    fill: { color: bgColor },
    rectRadius: 0.05,
    line: { color: theme.accent, pt: 1 },
  });
  slide.addText(title, {
    x: x + 0.1, y: y + 0.08, w: w - 0.2, h: 0.28,
    fontSize: 9, fontFace: BODY_FONT,
    color: titleColor, bold: true,
  });
  slide.addText(body, {
    x: x + 0.1, y: y + 0.32, w: w - 0.2, h: h - 0.42,
    fontSize: 8.5, fontFace: BODY_FONT,
    color: bodyColor,
    wrap: true, valign: "top",
  });
}

// ════════════════════════════════════════════════════════════════
// SLIDE 1: Title
// ════════════════════════════════════════════════════════════════
function slide01(pres) {
  const slide = pres.addSlide();
  slide.background = { color: theme.bg };

  // Top accent bar
  slide.addShape("rect", {
    x: 0, y: 0, w: 10, h: 0.07,
    fill: { color: theme.accent },
  });

  // Left vertical accent
  slide.addShape("rect", {
    x: 0, y: 0.07, w: 0.05, h: 5.555,
    fill: { color: theme.primary },
  });

  // Ticker badge
  slide.addShape("rect", {
    x: 0.5, y: 0.2, w: 1.5, h: 0.32,
    fill: { color: theme.primary },
    rectRadius: 0.03,
  });
  slide.addText("NASDAQ: UAL", {
    x: 0.5, y: 0.2, w: 1.5, h: 0.32,
    fontSize: 9, fontFace: BODY_FONT,
    color: theme.accent, bold: true,
    align: "center", valign: "middle",
    charSpacing: 1,
  });

  // Main title
  slide.addText("United Airlines Holdings", {
    x: 0.5, y: 0.65, w: 6.5, h: 0.7,
    fontSize: 36, fontFace: TITLE_FONT,
    color: "FFFFFF", bold: true,
  });
  slide.addText("FY2025 Results & 2026 Outlook", {
    x: 0.5, y: 1.3, w: 6.5, h: 0.45,
    fontSize: 20, fontFace: TITLE_FONT,
    color: theme.light,
  });
  slide.addText("Investor Presentation  |  March 2026", {
    x: 0.5, y: 1.82, w: 6.5, h: 0.3,
    fontSize: 11, fontFace: BODY_FONT,
    color: "6080B0",
  });

  // Divider
  slide.addShape("rect", {
    x: 0.5, y: 2.25, w: 6, h: 0.03,
    fill: { color: theme.primary },
  });

  // 3 stat mini-cards
  const cards = [
    { v: "$59.1B", l: "Total Revenue", s: "FY2025" },
    { v: "$10.62", l: "Adjusted EPS", s: "FY2025" },
    { v: "$3.4B", l: "Net Income", s: "FY2025" },
  ];
  cards.forEach((c, i) => {
    statCard(slide, 0.5 + i * 2.1, 2.45, 1.9, 1.1, c.v, c.l, c.s, true);
  });

  // Bottom tagline
  slide.addText("Executing United Next · Growing EBITDA · Returning Capital", {
    x: 0.5, y: 3.72, w: 6, h: 0.3,
    fontSize: 10, fontFace: BODY_FONT,
    color: "4060A0", italic: true,
  });

  // Right decorative panel
  slide.addShape("rect", {
    x: 7.1, y: 0.07, w: 2.9, h: 5.555,
    fill: { color: theme.secondary },
  });
  slide.addShape("rect", {
    x: 7.1, y: 0.07, w: 0.04, h: 5.555,
    fill: { color: theme.accent },
  });

  const metrics = [
    { v: "1,000+", l: "Aircraft Fleet" },
    { v: "338", l: "Destinations" },
    { v: "99.3%", l: "Completion Rate" },
    { v: "#1", l: "Transatlantic Share" },
  ];
  metrics.forEach((m, i) => {
    const my = 0.5 + i * 1.2;
    slide.addText(m.v, {
      x: 7.3, y: my, w: 2.5, h: 0.7,
      fontSize: 26, fontFace: TITLE_FONT,
      color: theme.light, bold: true,
      align: "center",
    });
    slide.addText(m.l, {
      x: 7.3, y: my + 0.65, w: 2.5, h: 0.3,
      fontSize: 9, fontFace: BODY_FONT,
      color: "6080B0", align: "center",
    });
  });
}

// ════════════════════════════════════════════════════════════════
// SLIDE 2: Investment Thesis
// ════════════════════════════════════════════════════════════════
function slide02(pres) {
  const slide = pres.addSlide();
  slide.background = { color: theme.bg };
  pageBadge(slide, pres, 2);

  slide.addShape("rect", { x: 0, y: 0, w: 10, h: 0.06, fill: { color: theme.accent } });

  slide.addText("Investment Thesis", {
    x: 0.35, y: 0.15, w: 6, h: 0.5,
    fontSize: 26, fontFace: TITLE_FONT,
    color: "FFFFFF", bold: true,
  });
  sectionLabel(slide, "Why UAL Now", 0.35, 0.68, 6, true);

  // 3 thesis paragraphs
  const theses = [
    {
      num: "01",
      title: "EPS Absorption Complete — Compounding Begins",
      body: "Flat EPS 2022–2025 was not stagnation — UAL absorbed $10B+ in new labor contracts (ALPA, IAM) while simultaneously accelerating fleet modernization. The cost reset is complete. With 50+ 787s and 200+ MAX jets now in service, unit cost leverage is inflecting positively into 2026.",
    },
    {
      num: "02",
      title: "2026: First Year to Exceed 2019 Peak EPS",
      body: "2019 was the prior EPS high-water mark at ~$11. Guidance of $12–14 in 2026 represents the first true EPS breakthrough, driven by premium seat penetration (+300bps), MileagePlus loyalty revenue growth, and Starlink/Kinective ancillary revenue unlocking.",
    },
    {
      num: "03",
      title: "Structural Moat: Network + Loyalty + Tech",
      body: "UAL's hub footprint (ORD, EWR, IAH, DEN, SFO, LAX) is irreplicable at scale. MileagePlus generates $6B+ in annual revenue independent of flying. Kinective Media + Starlink connectivity creates a new in-flight media revenue stream with $500M+ long-term potential.",
    },
  ];

  theses.forEach((t, i) => {
    const ty = 1.0 + i * 1.35;
    slide.addShape("rect", {
      x: 0.35, y: ty, w: 0.4, h: 0.4,
      fill: { color: theme.accent },
      rectRadius: 0.03,
    });
    slide.addText(t.num, {
      x: 0.35, y: ty, w: 0.4, h: 0.4,
      fontSize: 12, fontFace: TITLE_FONT, color: "FFFFFF", bold: true,
      align: "center", valign: "middle",
    });
    slide.addText(t.title, {
      x: 0.85, y: ty + 0.02, w: 5.8, h: 0.3,
      fontSize: 10.5, fontFace: BODY_FONT, color: theme.light, bold: true,
    });
    slide.addText(t.body, {
      x: 0.85, y: ty + 0.32, w: 5.8, h: 0.95,
      fontSize: 8.5, fontFace: BODY_FONT, color: "90A8CC",
      wrap: true, valign: "top",
    });
  });

  // Right sidebar stat cards
  const sidecards = [
    { v: "$12–14", l: "2026E Adj EPS", s: "vs $10.62 in FY2025" },
    { v: "$10B+", l: "Labor Absorbed", s: "2022–2025 contracts" },
    { v: "50+", l: "787 Dreamliners", s: "in active fleet" },
    { v: "$6B+", l: "MileagePlus Rev", s: "annual contribution" },
  ];
  sidecards.forEach((c, i) => {
    statCard(slide, 7.0, 0.8 + i * 1.1, 2.6, 0.95, c.v, c.l, c.s, true);
  });
}

// ════════════════════════════════════════════════════════════════
// SLIDE 3: Multi-Year Financial Summary
// ════════════════════════════════════════════════════════════════
function slide03(pres) {
  const slide = pres.addSlide();
  slide.background = { color: "F0F4FA" };
  pageBadge(slide, pres, 3);

  slide.addShape("rect", { x: 0, y: 0, w: 10, h: 0.06, fill: { color: theme.accent } });

  slide.addText("Multi-Year Financial Summary", {
    x: 0.35, y: 0.12, w: 7, h: 0.42,
    fontSize: 22, fontFace: TITLE_FONT,
    color: theme.primary, bold: true,
  });
  sectionLabel(slide, "FY2021 – FY2026E", 0.35, 0.58, 7, false);

  const rows = [
    headerRow(["Metric", "FY2021", "FY2022", "FY2023", "FY2024", "FY2025", "FY2026E"]),
    dataRow(["Revenue ($B)", "$24.6", "$44.9", "$53.7", "$57.1", "$59.1", "$62.0E"], { alt: false, light: true }),
    dataRow(["Op Income ($B)", "$(2.8)", "$3.5", "$5.6", "$4.9", "$4.7", "$6.2E"], { alt: true, light: true }),
    dataRow(["Net Income ($B)", "$(1.9)", "$0.7", "$2.6", "$3.1", "$3.4", "$4.2E"], { alt: false, light: true }),
    dataRow(["Adj EPS", "$(7.05)", "$3.95", "$9.42", "$9.47", "$10.62", "$13.00E"], { alt: true, light: true }),
    dataRow(["EBITDA ($B)", "$(0.3)", "$5.2", "$8.0", "$7.8", "$8.2", "$9.5E"], { alt: false, light: true }),
    dataRow(["ASMs (B)", "162", "227", "269", "291", "302", "318E"], { alt: true, light: true }),
    dataRow(["Fleet (operated)", "715", "860", "940", "980", "1,010", "1,050E"], { alt: false, light: true }),
  ];

  slide.addTable(rows, makeTableOptsLight(0.35, 0.88, 9.3, [1.8, 1.2, 1.2, 1.2, 1.2, 1.2, 1.5], { rowH: 0.3, fontSize: 9 }));

  insightCard(slide, 0.35, 4.55, 9.3, 0.85,
    "KEY INSIGHT: 2026 EPS Breakthrough",
    "After absorbing $10B+ in new labor contracts across 2022–2025, UAL's adj EPS trajectory resumes upward. FY2026E guidance of $12–14 marks the first year exceeding 2019's prior peak of ~$11. EBITDA margin target of 15%+ signals durable structural improvement, not a cyclical bounce.",
    false);
}

// ════════════════════════════════════════════════════════════════
// SLIDE 4: FY2025 Quarterly P&L
// ════════════════════════════════════════════════════════════════
function slide04(pres) {
  const slide = pres.addSlide();
  slide.background = { color: "F0F4FA" };
  pageBadge(slide, pres, 4);

  slide.addShape("rect", { x: 0, y: 0, w: 10, h: 0.06, fill: { color: theme.accent } });

  slide.addText("FY2025 Quarterly P&L", {
    x: 0.35, y: 0.12, w: 7, h: 0.42,
    fontSize: 22, fontFace: TITLE_FONT,
    color: theme.primary, bold: true,
  });
  sectionLabel(slide, "Q1–Q4 2025 Breakdown", 0.35, 0.58, 5, false);

  const rows = [
    headerRow(["Quarter", "Revenue ($B)", "Op Income ($B)", "Net Income ($B)", "Adj EPS"]),
    dataRow(["Q1 2025", "$12.5", "$0.3", "$0.0", "$0.91"], { alt: false, light: true }),
    dataRow(["Q2 2025", "$14.2", "$1.5", "$1.0", "$3.51"], { alt: true, light: true }),
    dataRow(["Q3 2025", "$15.0", "$1.7", "$1.2", "$3.98"], { alt: false, light: true }),
    dataRow(["Q4 2025", "$17.4", "$1.2", "$1.2", "$2.22"], { alt: true, light: true }),
    dataRow([{ text: "FY2025 Total", options: { bold: true, color: theme.primary } }, "$59.1", "$4.7", "$3.4", "$10.62"], { alt: false, light: true }),
  ];

  slide.addTable(rows, makeTableOptsLight(0.35, 0.88, 6.5, [1.4, 1.4, 1.6, 1.6, 1.5], { rowH: 0.35 }));

  // Q4 note
  slide.addText("Q4 reflects seasonally strong holiday demand. Record Q4 revenue $17.4B driven by Latin/Caribbean leisure + corporate rebound.", {
    x: 0.35, y: 3.08, w: 6.5, h: 0.35,
    fontSize: 8.5, fontFace: BODY_FONT, color: "5A6A94", italic: true,
  });

  // Right insight cards
  insightCard(slide, 7.1, 0.88, 2.5, 1.1,
    "Q3: Peak Profitability",
    "Q3 Adj EPS of $3.98 is seasonal high — trans-Atlantic routes running full utilization. 787 gauge advantage driving TRASM premium.",
    false);

  insightCard(slide, 7.1, 2.1, 2.5, 1.1,
    "Q4: Revenue Record",
    "Q4 $17.4B is highest quarterly revenue in UAL history. MileagePlus redemptions + Basic Economy mix driving strong ancillary yield.",
    false);

  insightCard(slide, 7.1, 3.32, 2.5, 1.0,
    "Q1 Seasonality",
    "Q1 EPS near zero is structural — winter low-demand period offset by growing subscription/loyalty base that provides floor revenue.",
    false);

  insightCard(slide, 0.35, 3.55, 6.5, 0.82,
    "FULL YEAR SUMMARY",
    "FY2025 revenue +$2.0B (+3.5%) YoY. Net Income $3.4B, Adj EPS $10.62 (+12.2% YoY from $9.47). EBITDA ~$8.2B. Fleet grew to 1,010 operated aircraft.",
    false);
}

// ════════════════════════════════════════════════════════════════
// SLIDE 5: Unit Economics Deep Dive
// ════════════════════════════════════════════════════════════════
function slide05(pres) {
  const slide = pres.addSlide();
  slide.background = { color: "F0F4FA" };
  pageBadge(slide, pres, 5);

  slide.addShape("rect", { x: 0, y: 0, w: 10, h: 0.06, fill: { color: theme.accent } });

  slide.addText("Unit Economics Deep Dive", {
    x: 0.35, y: 0.12, w: 7, h: 0.42,
    fontSize: 22, fontFace: TITLE_FONT,
    color: theme.primary, bold: true,
  });
  sectionLabel(slide, "FY2025 Per-ASM Metrics", 0.35, 0.58, 5, false);

  // Unit econ table
  const rows = [
    headerRow(["Metric", "FY2025", "FY2024", "YoY Δ", "Commentary"]),
    dataRow(["CASM (total)", "15.96¢", "15.41¢", "+0.55¢", "Labor step-up + inflationary pressure"], { alt: false, light: true }),
    dataRow(["CASM ex-Fuel", "11.20¢", "10.82¢", "+0.38¢", "Improving toward 11.0¢ target"], { alt: true, light: true }),
    dataRow(["PRASM", "16.18¢", "15.93¢", "+0.25¢", "Premium mix shift driving yield"], { alt: false, light: true }),
    dataRow(["TRASM", "17.89¢", "17.50¢", "+0.39¢", "Cargo + ancillary contributing"], { alt: true, light: true }),
    dataRow(["Fuel Cost/Gal", "$2.62", "$2.71", "–$0.09", "New-gen fleet efficiency gain"], { alt: false, light: true }),
    dataRow(["Fuel Gallons (B)", "4.0B", "3.9B", "+2.6%", "Capacity growth, offset by efficiency"], { alt: true, light: true }),
    dataRow(["PRASM–CASM Spread", "+0.22¢", "+0.52¢", "–0.30¢", "Labor compression; recovering 2026"], { alt: false, light: true }),
  ];

  slide.addTable(rows, makeTableOptsLight(0.35, 0.88, 6.4, [1.6, 1.1, 1.1, 1.1, 2.5], { rowH: 0.3 }));

  // Fuel sensitivity callout
  insightCard(slide, 7.0, 0.88, 2.65, 1.6,
    "FUEL SENSITIVITY",
    "$0.10/gal change in jet fuel = ~$40M annual P&L impact (pre-hedge). UAL hedges ~25% of forward exposure. At $2.62/gal, FY2025 fuel bill was ~$10.4B. 2026 guide assumes $2.55–$2.75/gal range.",
    false);

  insightCard(slide, 7.0, 2.62, 2.65, 1.0,
    "CASM ex-FUEL TREND",
    "FY2026 target: 11.0¢ flat-to-down as 787/MAX fleet gauge advantage offsets wage escalation. New fleet = 15–20% better fuel burn vs. prior-gen.",
    false);

  // Progress bar: PRASM vs CASM
  sectionLabel(slide, "PRASM vs CASM Trend", 0.35, 3.75, 6.4, false);

  const bars = [
    { label: "CASM ex-Fuel", val: 11.20, max: 18, color: "EF4444" },
    { label: "PRASM", val: 16.18, max: 18, color: "3B82F6" },
    { label: "TRASM", val: 17.89, max: 18, color: "10B981" },
  ];

  bars.forEach((b, i) => {
    const by = 4.08 + i * 0.38;
    const bw = (b.val / b.max) * 5.8;
    slide.addText(b.label, {
      x: 0.35, y: by, w: 1.5, h: 0.28,
      fontSize: 8.5, fontFace: BODY_FONT,
      color: "1E2761", valign: "middle",
    });
    slide.addShape("rect", {
      x: 1.9, y: by + 0.04, w: 5.8, h: 0.2,
      fill: { color: "D0D9EE" },
      rectRadius: 0.03,
    });
    slide.addShape("rect", {
      x: 1.9, y: by + 0.04, w: bw, h: 0.2,
      fill: { color: b.color },
      rectRadius: 0.03,
    });
    slide.addText(`${b.val}¢`, {
      x: 1.9 + bw + 0.08, y: by, w: 0.6, h: 0.28,
      fontSize: 8.5, fontFace: BODY_FONT,
      color: b.color, bold: true, valign: "middle",
    });
  });
}

// ════════════════════════════════════════════════════════════════
// SLIDE 6: OpEx Composition
// ════════════════════════════════════════════════════════════════
function slide06(pres) {
  const slide = pres.addSlide();
  slide.background = { color: "F0F4FA" };
  pageBadge(slide, pres, 6);

  slide.addShape("rect", { x: 0, y: 0, w: 10, h: 0.06, fill: { color: theme.accent } });

  slide.addText("OpEx Composition", {
    x: 0.35, y: 0.12, w: 7, h: 0.42,
    fontSize: 22, fontFace: TITLE_FONT,
    color: theme.primary, bold: true,
  });
  sectionLabel(slide, "FY2025 Cost Breakdown — Total ~$54.4B", 0.35, 0.58, 7, false);

  const rows = [
    headerRow(["Cost Category", "FY2025 ($B)", "% of OpEx", "YoY Δ", "FY2026E"]),
    dataRow(["Labor (wages + benefits)", "$19.8", "36%", "+9.1%", "~$21.5B"], { alt: false, light: true }),
    dataRow(["Fuel & Oil", "$10.4", "19%", "–3.7%", "~$10.8B"], { alt: true, light: true }),
    dataRow(["Aircraft Rent & Maintenance", "$8.6", "16%", "+4.2%", "~$9.0B"], { alt: false, light: true }),
    dataRow(["Distribution & Sales", "$4.3", "8%", "+1.0%", "~$4.5B"], { alt: true, light: true }),
    dataRow(["Other OpEx", "$11.3", "21%", "+2.5%", "~$11.8B"], { alt: false, light: true }),
    dataRow([{ text: "Total OpEx", options: { bold: true, color: theme.primary } }, "$54.4", "100%", "+4.4%", "~$57.6B"], { alt: true, light: true }),
  ];

  slide.addTable(rows, makeTableOptsLight(0.35, 0.88, 6.8, [2.4, 1.4, 1.2, 1.2, 1.6], { rowH: 0.32 }));

  // Progress bars for cost breakdown
  sectionLabel(slide, "Cost Mix Visual", 0.35, 3.58, 6.8, false);

  const costMix = [
    { label: "Labor 36%", pct: 0.36, color: "EF4444" },
    { label: "Fuel 19%", pct: 0.19, color: "F59E0B" },
    { label: "Maint 16%", pct: 0.16, color: "8B5CF6" },
    { label: "Dist 8%", pct: 0.08, color: "3B82F6" },
    { label: "Other 21%", pct: 0.21, color: "10B981" },
  ];

  let xStart = 0.35;
  const totalW = 6.8;
  costMix.forEach((c) => {
    const cw = c.pct * totalW;
    slide.addShape("rect", {
      x: xStart, y: 3.88, w: cw - 0.03, h: 0.32,
      fill: { color: c.color },
      rectRadius: 0.02,
    });
    if (cw > 0.8) {
      slide.addText(c.label, {
        x: xStart, y: 3.88, w: cw - 0.03, h: 0.32,
        fontSize: 7.5, fontFace: BODY_FONT,
        color: "FFFFFF", bold: true,
        align: "center", valign: "middle",
      });
    }
    xStart += cw;
  });

  // 3 insight cards on right
  insightCard(slide, 7.3, 0.88, 2.3, 1.0,
    "LABOR: STEP-UP ABSORBED",
    "ALPA contract ratified 2023 (+17% pay). IAM ground workers settled 2024. Combined $3B+ annual increment now fully in run rate.",
    false);

  insightCard(slide, 7.3, 2.0, 2.3, 1.0,
    "FUEL: NEW FLEET SAVES",
    "787 burns 20% less fuel per seat vs 767. 50-unit fleet saves ~$400M/yr vs prior gauge. MAX family replacing 737-700s.",
    false);

  insightCard(slide, 7.3, 3.12, 2.3, 1.1,
    "DISTRIBUTION: OPT.",
    "NDC direct booking growth reducing GDS fees. United.com share up to 52% of total bookings. Channel shift saves ~$100M/yr.",
    false);
}

// ════════════════════════════════════════════════════════════════
// SLIDE 7: Regional Revenue Mix
// ════════════════════════════════════════════════════════════════
function slide07(pres) {
  const slide = pres.addSlide();
  slide.background = { color: "F0F4FA" };
  pageBadge(slide, pres, 7);

  slide.addShape("rect", { x: 0, y: 0, w: 10, h: 0.06, fill: { color: theme.accent } });

  slide.addText("Regional Revenue Mix", {
    x: 0.35, y: 0.12, w: 7, h: 0.42,
    fontSize: 22, fontFace: TITLE_FONT,
    color: theme.primary, bold: true,
  });
  sectionLabel(slide, "Domestic vs International Breakdown — FY2025", 0.35, 0.58, 7, false);

  const rows = [
    headerRow(["Region", "Revenue ($B)", "% of Total", "YoY Growth", "ASM Share", "TRASM"]),
    dataRow(["Domestic (US)", "$29.2", "49%", "+2.1%", "54%", "17.4¢"], { alt: false, light: true }),
    dataRow(["Atlantic", "$14.8", "25%", "+6.3%", "22%", "20.1¢"], { alt: true, light: true }),
    dataRow(["Pacific", "$5.7", "10%", "+4.9%", "10%", "17.8¢"], { alt: false, light: true }),
    dataRow(["Latin America", "$6.8", "11%", "+8.4%", "10%", "18.6¢"], { alt: true, light: true }),
    dataRow(["MEIA (Mid-East/India/Africa)", "$2.6", "5%", "+22.3%", "4%", "21.3¢"], { alt: false, light: true }),
    dataRow([{ text: "Total", options: { bold: true, color: theme.primary } }, "$59.1", "100%", "+3.5%", "100%", "17.89¢"], { alt: true, light: true }),
  ];

  slide.addTable(rows, makeTableOptsLight(0.35, 0.88, 7.2, [2.1, 1.3, 1.1, 1.2, 1.2, 1.3], { rowH: 0.33 }));

  // Progress bars: region revenue share
  sectionLabel(slide, "Revenue Share by Region", 0.35, 3.6, 7.2, false);

  const regions = [
    { label: "Domestic 49%", pct: 0.49, color: "1E2761" },
    { label: "Atlantic 25%", pct: 0.25, color: "3B82F6" },
    { label: "Pacific 10%", pct: 0.10, color: "8B5CF6" },
    { label: "Latin 11%", pct: 0.11, color: "10B981" },
    { label: "MEIA 5%", pct: 0.05, color: "F59E0B" },
  ];

  let rx = 0.35;
  const rw = 7.2;
  regions.forEach((r) => {
    const bw = r.pct * rw;
    slide.addShape("rect", {
      x: rx, y: 3.9, w: bw - 0.03, h: 0.35,
      fill: { color: r.color },
    });
    if (bw > 0.9) {
      slide.addText(r.label, {
        x: rx, y: 3.9, w: bw - 0.03, h: 0.35,
        fontSize: 7.5, fontFace: BODY_FONT,
        color: "FFFFFF", bold: true,
        align: "center", valign: "middle",
      });
    }
    rx += bw;
  });

  // MEIA insight
  insightCard(slide, 7.75, 0.88, 1.9, 1.6,
    "MEIA: FASTEST GROWTH",
    "+22.3% revenue YoY. India routes (ORD-DEL, EWR-DEL, EWR-BOM) running at 90%+ load factor. MEIA TRASM of 21.3¢ is highest of any region.",
    false);

  insightCard(slide, 7.75, 2.6, 1.9, 1.1,
    "ATLANTIC PREMIUM",
    "Polaris Business Class driving 20¢+ TRASM on European routes. Premium cabin revenue +12% YoY.",
    false);

  insightCard(slide, 7.75, 3.82, 1.9, 0.9,
    "LATIN SURGE",
    "Latin +8.4% YoY driven by leisure demand to Mexico, Caribbean, Central America. Load factors >90%.",
    false);
}

// ════════════════════════════════════════════════════════════════
// SLIDE 8: Revenue Quality
// ════════════════════════════════════════════════════════════════
function slide08(pres) {
  const slide = pres.addSlide();
  slide.background = { color: "F0F4FA" };
  pageBadge(slide, pres, 8);

  slide.addShape("rect", { x: 0, y: 0, w: 10, h: 0.06, fill: { color: theme.accent } });

  slide.addText("Revenue Quality", {
    x: 0.35, y: 0.12, w: 7, h: 0.42,
    fontSize: 22, fontFace: TITLE_FONT,
    color: theme.primary, bold: true,
  });
  sectionLabel(slide, "Premium, Loyalty & Ancillary Mix — FY2025", 0.35, 0.58, 7, false);

  // 5 KPI stat cards
  const kpis = [
    { v: "30%", l: "Premium Revenue Share", s: "+300bps YoY" },
    { v: "$6.1B", l: "MileagePlus Revenue", s: "+9% YoY" },
    { v: "58%", l: "Corporate Travel %", s: "of managed travel" },
    { v: "18%", l: "Ancillary Revenue %", s: "bags, upgrades, extras" },
    { v: "14%", l: "Basic Economy %", s: "of domestic bookings" },
  ];

  const cardW = 1.85;
  kpis.forEach((k, i) => {
    statCard(slide, 0.35 + i * 1.93, 0.88, cardW, 1.05, k.v, k.l, k.s, false);
  });

  // Premium detail
  sectionLabel(slide, "Premium Cabin Detail", 0.35, 2.08, 4.5, false);

  const premRows = [
    headerRow(["Product", "Revenue ($B)", "Share of Total", "YoY Growth", "Load Factor"]),
    dataRow(["Polaris Business", "$8.4", "14%", "+15.2%", "89%"], { alt: false, light: true }),
    dataRow(["United First (Domestic)", "$3.9", "7%", "+8.1%", "85%"], { alt: true, light: true }),
    dataRow(["United Premium Plus", "$2.4", "4%", "+21.3%", "82%"], { alt: false, light: true }),
    dataRow(["Economy Plus (seats)", "$3.0", "5%", "+6.4%", "N/A"], { alt: true, light: true }),
  ];

  slide.addTable(premRows, makeTableOptsLight(0.35, 2.38, 4.5, [1.8, 1.3, 1.2, 1.2, 1.0], { rowH: 0.3 }));

  // Loyalty detail
  sectionLabel(slide, "MileagePlus Loyalty Detail", 5.05, 2.08, 4.6, false);

  const loyaltyRows = [
    headerRow(["Revenue Stream", "FY2025 ($B)", "YoY Growth"]),
    dataRow(["Chase co-brand card", "$3.8", "+11.2%"], { alt: false, light: true }),
    dataRow(["Award redemptions (capacity)", "$1.3", "+4.1%"], { alt: true, light: true }),
    dataRow(["Premier status fees + extras", "$0.6", "+18.5%"], { alt: false, light: true }),
    dataRow(["Partner miles (hotels, retail)", "$0.4", "+7.3%"], { alt: true, light: true }),
    dataRow([{ text: "Total MileagePlus", options: { bold: true, color: theme.primary } }, "$6.1", "+9.1%"], { alt: false, light: true }),
  ];

  slide.addTable(loyaltyRows, makeTableOptsLight(5.05, 2.38, 4.6, [2.3, 1.4, 1.4], { rowH: 0.3 }));

  insightCard(slide, 0.35, 4.62, 9.3, 0.75,
    "REVENUE QUALITY SUMMARY",
    "30% premium revenue share + $6.1B MileagePlus + 18% ancillary = ~65% of total revenue from high-quality, loyalty-anchored, premium-driven sources. This mix is structurally more resilient than pre-COVID 2019 composition.",
    false);
}

// ════════════════════════════════════════════════════════════════
// SLIDE 9: Fleet Modernization
// ════════════════════════════════════════════════════════════════
function slide09(pres) {
  const slide = pres.addSlide();
  slide.background = { color: "F0F4FA" };
  pageBadge(slide, pres, 9);

  slide.addShape("rect", { x: 0, y: 0, w: 10, h: 0.06, fill: { color: theme.accent } });

  slide.addText("Fleet Modernization", {
    x: 0.35, y: 0.12, w: 7, h: 0.42,
    fontSize: 22, fontFace: TITLE_FONT,
    color: theme.primary, bold: true,
  });
  sectionLabel(slide, "Order Book, Deliveries & CASM Impact", 0.35, 0.58, 7, false);

  // Fleet stat cards
  const fleetStats = [
    { v: "1,010", l: "Operated Aircraft", s: "FY2025 end" },
    { v: "200+", l: "MAX Family", s: "737-8 / 737-9 / 737-10" },
    { v: "50+", l: "787 Dreamliners", s: "operated" },
    { v: "700+", l: "Pending Orders", s: "787 + MAX + A321" },
  ];

  fleetStats.forEach((s, i) => {
    statCard(slide, 0.35 + i * 2.3, 0.88, 2.1, 0.85, s.v, s.l, s.s, false);
  });

  // Order book table
  sectionLabel(slide, "Order Book Details", 0.35, 1.88, 6.5, false);

  const orderRows = [
    headerRow(["Aircraft Type", "On Order", "2026 Delivers.", "2027–28 Dels.", "CASM Impact", "Status"]),
    dataRow(["Boeing 737 MAX 8", "100", "20", "40+", "–12% vs 737-700", "On track"], { alt: false, light: true }),
    dataRow(["Boeing 737 MAX 9", "120", "18", "35+", "–14% vs 737-800", "On track"], { alt: true, light: true }),
    dataRow(["Boeing 737 MAX 10", "100", "0", "15+", "–18% vs 737-900", { text: "FAA cert pending", options: { color: "EF4444" } }], { alt: false, light: true }),
    dataRow(["Boeing 787-9", "50", "10", "20+", "–20% vs 767", "Delayed 6–9 mo"], { alt: true, light: true }),
    dataRow(["Boeing 787-10", "40", "5", "12+", "–22% vs 767", "Delayed 6–9 mo"], { alt: false, light: true }),
    dataRow(["Airbus A321XLR", "50", "0", "5+", "New route economics", "2027+ entry"], { alt: true, light: true }),
  ];

  slide.addTable(orderRows, makeTableOptsLight(0.35, 2.18, 6.5, [1.7, 0.9, 1.1, 1.1, 1.4, 1.3], { rowH: 0.3 }));

  // Boeing delay risk callout
  insightCard(slide, 7.05, 0.88, 2.6, 1.4,
    "BOEING DELIVERY RISK",
    "787 program delays of 6–9 months pushed ~15 aircraft into 2026. 737 MAX 10 FAA certification remains unresolved, affecting 100-unit order. UAL estimates ~$200M in 2025 lost opportunity from deferred aircraft.",
    false);

  insightCard(slide, 7.05, 2.42, 2.6, 1.2,
    "GAUGE STRATEGY",
    "Widebody gauge on transcon + trans-Atlantic routes enables more premium seats per departure. 787 has ~30% more premium capacity than 767 it replaces.",
    false);

  insightCard(slide, 7.05, 3.76, 2.6, 0.95,
    "AIRBUS HEDGE",
    "50 A321XLR orders diversify Boeing dependency. Entry 2027–2028. Ideal for thin international routes that need fuel efficiency without widebody capacity.",
    false);
}

// ════════════════════════════════════════════════════════════════
// SLIDE 10: New Products
// ════════════════════════════════════════════════════════════════
function slide10(pres) {
  const slide = pres.addSlide();
  slide.background = { color: "F0F4FA" };
  pageBadge(slide, pres, 10);

  slide.addShape("rect", { x: 0, y: 0, w: 10, h: 0.06, fill: { color: theme.accent } });

  slide.addText("New Products & Cabin Experience", {
    x: 0.35, y: 0.12, w: 9, h: 0.42,
    fontSize: 22, fontFace: TITLE_FONT,
    color: theme.primary, bold: true,
  });
  sectionLabel(slide, "Four Flagship Product Innovations — Revenue Drivers for 2026+", 0.35, 0.58, 9, false);

  const products = [
    {
      title: "Polaris Business Class",
      tag: "LONG-HAUL INTERNATIONAL",
      color: "1E2761",
      tagColor: "3B82F6",
      body: "Lie-flat seats on all 777/787 widebody aircraft. Direct-aisle access on 787-9/-10. Meal service designed with Spiaggia chefs. Revenue +15% YoY. 89% load factor. TRASM 25¢+ on flagship routes.",
      kpis: ["89% LF", "+15% YoY Rev", "25¢+ TRASM"],
    },
    {
      title: "United Elevated (787 Transcon)",
      tag: "DOMESTIC PREMIUM",
      color: "162050",
      tagColor: "10B981",
      body: "787-9 deployed on JFK/EWR-LAX/SFO transcon routes. Polaris seats on domestic routes competing directly with Delta One. Premium demand captured from high-yield corporate NY/LA corridors.",
      kpis: ["JFK/EWR-LAX", "Polaris Domestic", "$1.2B TAM"],
    },
    {
      title: "Coastliner (Regional Elite)",
      tag: "REGIONAL PREMIUM",
      color: "1a2040",
      tagColor: "8B5CF6",
      body: "Premium E175 regional jets with dedicated Polaris-style seats and enhanced amenities for connecting routes. Targets high-value business travelers on spoke markets. Planned 2026 launch across 20 routes.",
      kpis: ["E175 Platform", "20 Routes 2026", "35 min boarding"],
    },
    {
      title: "Basic Economy Optimized",
      tag: "PRICE-SENSITIVE DEMAND",
      color: "0F1535",
      tagColor: "F59E0B",
      body: "Redesigned Basic Economy to capture ULCC share without yield cannibalization. Dynamic pricing engine ensures BE fares yield 15% above ULCC competition. BE share ~14% of domestic bookings.",
      kpis: ["14% Dom Mix", "+15% vs ULCC", "No upgrades"],
    },
  ];

  products.forEach((p, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const px = 0.35 + col * 4.85;
    const py = 0.88 + row * 2.3;
    const pw = 4.6;
    const ph = 2.15;

    slide.addShape("rect", {
      x: px, y: py, w: pw, h: ph,
      fill: { color: p.color },
      rectRadius: 0.06,
    });
    slide.addShape("rect", {
      x: px, y: py, w: pw, h: 0.04,
      fill: { color: p.tagColor },
      rectRadius: 0.03,
    });

    slide.addShape("rect", {
      x: px + 0.12, y: py + 0.1, w: 1.8, h: 0.22,
      fill: { color: p.tagColor },
      rectRadius: 0.02,
    });
    slide.addText(p.tag, {
      x: px + 0.12, y: py + 0.1, w: 1.8, h: 0.22,
      fontSize: 7, fontFace: BODY_FONT, color: "FFFFFF", bold: true,
      align: "center", valign: "middle", charSpacing: 1,
    });

    slide.addText(p.title, {
      x: px + 0.15, y: py + 0.38, w: pw - 0.3, h: 0.38,
      fontSize: 13, fontFace: TITLE_FONT, color: "FFFFFF", bold: true,
    });
    slide.addText(p.body, {
      x: px + 0.15, y: py + 0.76, w: pw - 0.3, h: 0.98,
      fontSize: 8.5, fontFace: BODY_FONT, color: "CADCFC",
      wrap: true, valign: "top",
    });

    // KPI chips
    p.kpis.forEach((k, ki) => {
      slide.addShape("rect", {
        x: px + 0.15 + ki * 1.48, y: py + ph - 0.36, w: 1.38, h: 0.26,
        fill: { color: p.tagColor },
        rectRadius: 0.03,
      });
      slide.addText(k, {
        x: px + 0.15 + ki * 1.48, y: py + ph - 0.36, w: 1.38, h: 0.26,
        fontSize: 7.5, fontFace: BODY_FONT, color: "FFFFFF", bold: true,
        align: "center", valign: "middle",
      });
    });
  });
}

// ════════════════════════════════════════════════════════════════
// SLIDE 11: Starlink & Kinective Media
// ════════════════════════════════════════════════════════════════
function slide11(pres) {
  const slide = pres.addSlide();
  slide.background = { color: theme.bg };
  pageBadge(slide, pres, 11);

  slide.addShape("rect", { x: 0, y: 0, w: 10, h: 0.06, fill: { color: theme.accent } });

  slide.addText("Starlink & Kinective Media", {
    x: 0.35, y: 0.15, w: 7, h: 0.45,
    fontSize: 26, fontFace: TITLE_FONT,
    color: "FFFFFF", bold: true,
  });
  slide.addText("In-Flight Connectivity + Personalized Media = New Revenue Stream", {
    x: 0.35, y: 0.62, w: 8, h: 0.28,
    fontSize: 10.5, fontFace: BODY_FONT, color: theme.light,
  });

  // Timeline bar
  sectionLabel(slide, "Starlink Rollout Timeline", 0.35, 0.95, 9.3, true);

  const milestones = [
    { y: "2024", ev: "Pilot launch on 20 aircraft (757 fleet)" },
    { y: "Mid-2025", ev: "200+ aircraft equipped — first carrier at scale" },
    { y: "End-2025", ev: "Full domestic narrowbody fleet equipped" },
    { y: "2026", ev: "Full fleet (narrowbody + widebody) — 1,000+ aircraft" },
    { y: "2026+", ev: "$500M+ ancillary revenue opportunity at full run-rate" },
  ];

  milestones.forEach((m, i) => {
    const mx = 0.35 + i * 1.88;
    slide.addShape("rect", {
      x: mx, y: 1.3, w: 1.75, h: 0.32,
      fill: { color: theme.primary },
      rectRadius: 0.04,
    });
    slide.addShape(pres.shapes.OVAL, {
      x: mx + 0.78, y: 1.6, w: 0.2, h: 0.2,
      fill: { color: theme.accent },
    });
    slide.addText(m.y, {
      x: mx, y: 1.3, w: 1.75, h: 0.32,
      fontSize: 9, fontFace: BODY_FONT,
      color: theme.accent, bold: true,
      align: "center", valign: "middle",
    });
    slide.addText(m.ev, {
      x: mx, y: 1.82, w: 1.75, h: 0.5,
      fontSize: 8.5, fontFace: BODY_FONT,
      color: "90A8CC", wrap: true, valign: "top",
    });
  });

  // Horizontal timeline connector
  slide.addShape("rect", {
    x: 0.35, y: 1.68, w: 9.3, h: 0.04,
    fill: { color: theme.primary },
  });

  // Kinective Media table
  sectionLabel(slide, "Kinective Media Network Details", 0.35, 2.48, 9.3, true);

  const kinRows = [
    headerRow(["Metric", "Detail", "Revenue Model", "Timeline"]),
    dataRow(["Connectivity speed", "100 Mbps per aircraft (Starlink)", "Premium Wi-Fi $28/flight", "2025 active"], { alt: false }),
    dataRow(["Ad inventory", "Personalized pre/post-flight ads", "CPM-based, $15–25 CPM", "2025 launched"], { alt: true }),
    dataRow(["Shopping integration", "United Shop (SkyMall 2.0)", "Revenue share 15–20%", "2026 launch"], { alt: false }),
    dataRow(["Destination content", "Hotel/car/tour booking in-flight", "Booking commissions 8–12%", "2026 launch"], { alt: true }),
    dataRow(["Gaming / streaming", "High-def seatback + device", "Subscription $9.99/mo", "2027 roadmap"], { alt: false }),
  ];

  slide.addTable(kinRows, makeTableOpts(0.35, 2.78, 9.3, [2.0, 2.5, 2.3, 1.5], { rowH: 0.28 }));

  // Right stat
  insightCard(slide, 7.35, 0.95, 2.3, 1.2,
    "REVENUE OPPORTUNITY",
    "$500M+ ancillary by 2027 at full fleet deployment. Peers like Delta/American have not yet matched UAL's Starlink scale advantage.",
    true);
}

// ════════════════════════════════════════════════════════════════
// SLIDE 12: Operational Performance
// ════════════════════════════════════════════════════════════════
function slide12(pres) {
  const slide = pres.addSlide();
  slide.background = { color: "F0F4FA" };
  pageBadge(slide, pres, 12);

  slide.addShape("rect", { x: 0, y: 0, w: 10, h: 0.06, fill: { color: theme.accent } });

  slide.addText("Operational Performance", {
    x: 0.35, y: 0.12, w: 7, h: 0.42,
    fontSize: 22, fontFace: TITLE_FONT,
    color: theme.primary, bold: true,
  });
  sectionLabel(slide, "FY2025 Ops Metrics & Demand Signals", 0.35, 0.58, 7, false);

  // Stat cards
  const opStats = [
    { v: "83.1%", l: "On-Time Arrival", s: "D:00 / A:14" },
    { v: "99.3%", l: "Completion Factor", s: "industry-best 2025" },
    { v: "87.4%", l: "Load Factor", s: "FY2025 average" },
    { v: "4.5", l: "DOT Complaints / 100K", s: "improved from 5.8" },
  ];

  opStats.forEach((s, i) => {
    statCard(slide, 0.35 + i * 2.35, 0.88, 2.15, 0.9, s.v, s.l, s.s, false);
  });

  // Demand signals table
  sectionLabel(slide, "Forward Demand Indicators", 0.35, 1.93, 6.5, false);

  const demandRows = [
    headerRow(["Indicator", "Q1 2026E", "Q2 2026E", "Signal"]),
    dataRow(["Corporate travel bookings YoY", "+8.2%", "+9.1%", { text: "Strong", options: { color: "10B981", bold: true } }], { alt: false, light: true }),
    dataRow(["Leisure advance bookings", "+5.1%", "+7.3%", { text: "Solid", options: { color: "10B981", bold: true } }], { alt: true, light: true }),
    dataRow(["Atlantic booking pace", "+10.2%", "+14.1%", { text: "Very Strong", options: { color: "10B981", bold: true } }], { alt: false, light: true }),
    dataRow(["MileagePlus redemptions (demand)", "+6.8%", "+9.0%", { text: "Strong", options: { color: "10B981", bold: true } }], { alt: true, light: true }),
    dataRow(["Group / charter demand", "+4.2%", "+5.6%", { text: "Moderate", options: { color: "F59E0B", bold: true } }], { alt: false, light: true }),
    dataRow(["Cargo yield forward", "–1.1%", "+3.2%", { text: "Mixed", options: { color: "F59E0B", bold: true } }], { alt: true, light: true }),
  ];

  slide.addTable(demandRows, makeTableOptsLight(0.35, 2.23, 6.5, [2.8, 1.2, 1.2, 1.3], { rowH: 0.3 }));

  // Ops improvement detail
  insightCard(slide, 7.05, 0.88, 2.6, 1.35,
    "OPS IMPROVEMENT DRIVERS",
    "Reduced cancellations via improved crew scheduling (FLICA system). Maintenance reliability up with 787/MAX new-gen fleet. Hub operations score highest in UAL history.",
    false);

  insightCard(slide, 7.05, 2.37, 2.6, 1.35,
    "COMPLETION FACTOR 99.3%",
    "Industry-leading completion rate driven by proactive MX, tech ops crew pre-positioning, and reduced weather-related cancellations via improved hub routing.",
    false);

  insightCard(slide, 7.05, 3.86, 2.6, 0.85,
    "ATLANTIC DEMAND: HOT",
    "+10–14% forward pace into Q1/Q2 2026 driven by European leisure + India business class demand.",
    false);
}

// ════════════════════════════════════════════════════════════════
// SLIDE 13: Balance Sheet & Deleveraging
// ════════════════════════════════════════════════════════════════
function slide13(pres) {
  const slide = pres.addSlide();
  slide.background = { color: "F0F4FA" };
  pageBadge(slide, pres, 13);

  slide.addShape("rect", { x: 0, y: 0, w: 10, h: 0.06, fill: { color: theme.accent } });

  slide.addText("Balance Sheet & Deleveraging", {
    x: 0.35, y: 0.12, w: 7, h: 0.42,
    fontSize: 22, fontFace: TITLE_FONT,
    color: theme.primary, bold: true,
  });
  sectionLabel(slide, "Net Debt Reduction & Credit Improvement — 2023–2025", 0.35, 0.58, 7, false);

  const leverageRows = [
    headerRow(["Metric", "FY2023", "FY2024", "FY2025", "2026E Target"]),
    dataRow(["Total Debt ($B)", "$34.8", "$31.2", "$28.9", "~$27.0"], { alt: false, light: true }),
    dataRow(["Cash & Equivalents ($B)", "$3.8", "$3.1", "$2.9", "~$3.0"], { alt: true, light: true }),
    dataRow(["Net Debt ($B)", "$31.0", "$28.1", "$26.0", "~$24.0"], { alt: false, light: true }),
    dataRow(["EBITDA ($B)", "$7.8", "$7.8", "$8.2", "~$9.5"], { alt: true, light: true }),
    dataRow(["Net Leverage (x)", "4.0x", "3.6x", "3.2x", "~2.5x"], { alt: false, light: true }),
    dataRow(["Adj. Net Leverage (ex-MP)", "2.9x", "2.6x", "2.3x", "~1.8x"], { alt: true, light: true }),
    dataRow(["Credit Rating (S&P)", "BB", "BB+", "BB+", "BBB- Target"], { alt: false, light: true }),
  ];

  slide.addTable(leverageRows, makeTableOptsLight(0.35, 0.88, 6.5, [2.6, 1.05, 1.05, 1.05, 1.75], { rowH: 0.3 }));

  // Deleveraging progress bar
  sectionLabel(slide, "Net Leverage Trajectory", 0.35, 3.68, 6.5, false);

  const leveragePoints = [
    { label: "FY2023: 4.0x", val: 4.0 },
    { label: "FY2024: 3.6x", val: 3.6 },
    { label: "FY2025: 3.2x", val: 3.2 },
    { label: "2026E: ~2.5x", val: 2.5 },
  ];

  leveragePoints.forEach((lp, i) => {
    const lx = 0.35 + i * 1.65;
    const barH = (lp.val / 4.5) * 0.7;
    slide.addShape("rect", {
      x: lx, y: 4.7 - barH, w: 1.4, h: barH,
      fill: { color: i < 2 ? "EF4444" : i === 2 ? "F59E0B" : "10B981" },
      rectRadius: 0.03,
    });
    slide.addText(lp.label, {
      x: lx, y: 4.72, w: 1.4, h: 0.25,
      fontSize: 8, fontFace: BODY_FONT, color: "1E2761", bold: true,
      align: "center",
    });
  });

  // MileagePlus and Credit cards
  insightCard(slide, 7.05, 0.88, 2.6, 1.25,
    "MILEAGEPLUS ASSET",
    "MileagePlus securitized as collateral for $6.8B facility. Market value of MP estimated at $22B+ (JPMorgan analysis 2024). Adj. leverage ex-MP = 2.3x — investment grade equivalent.",
    false);

  insightCard(slide, 7.05, 2.26, 2.6, 1.25,
    "IG CREDIT TARGET",
    "S&P BB+ with positive outlook. Moody's Ba1. First IG upgrade expected late 2026/early 2027 at current deleveraging pace. IG would reduce interest expense ~$150M/yr.",
    false);

  insightCard(slide, 7.05, 3.64, 2.6, 1.07,
    "DEBT MATURITY PROFILE",
    "No major debt wall until 2028. 2026 maturities ~$1.2B, comfortably covered by FCF. Weighted avg cost of debt 4.8%.",
    false);
}

// ════════════════════════════════════════════════════════════════
// SLIDE 14: Cash Flow & Capital Allocation
// ════════════════════════════════════════════════════════════════
function slide14(pres) {
  const slide = pres.addSlide();
  slide.background = { color: "F0F4FA" };
  pageBadge(slide, pres, 14);

  slide.addShape("rect", { x: 0, y: 0, w: 10, h: 0.06, fill: { color: theme.accent } });

  slide.addText("Cash Flow & Capital Allocation", {
    x: 0.35, y: 0.12, w: 7, h: 0.42,
    fontSize: 22, fontFace: TITLE_FONT,
    color: theme.primary, bold: true,
  });
  sectionLabel(slide, "FY2025 Actuals & 2026 Priorities", 0.35, 0.58, 7, false);

  // CF table
  const cfRows = [
    headerRow(["Cash Flow Item", "FY2023", "FY2024", "FY2025", "FY2026E"]),
    dataRow(["Operating Cash Flow", "$4.2B", "$5.1B", "$5.8B", "~$6.5B"], { alt: false, light: true }),
    dataRow(["Capital Expenditures", "($4.0B)", "($3.9B)", "($3.7B)", "~($4.5B)"], { alt: true, light: true }),
    dataRow(["Free Cash Flow", "$0.2B", "$1.2B", "$2.1B", "~$2.0B"], { alt: false, light: true }),
    dataRow(["Debt Repayment", "($1.1B)", "($2.0B)", "($2.5B)", "~($1.5B)"], { alt: true, light: true }),
    dataRow(["Share Repurchases", "$0", "$0.5B", "$1.5B", "~$1.5B"], { alt: false, light: true }),
    dataRow(["Ending Cash Balance", "$3.8B", "$3.1B", "$2.9B", "~$3.0B"], { alt: true, light: true }),
  ];

  slide.addTable(cfRows, makeTableOptsLight(0.35, 0.88, 6.5, [2.5, 1.05, 1.05, 1.05, 1.85], { rowH: 0.3 }));

  // Capital priority stack
  sectionLabel(slide, "Capital Allocation Priority Stack", 0.35, 3.52, 4.2, false);

  const priorities = [
    { n: "1", t: "Fleet Capex", d: "~$4.5B/yr — ordered aircraft deliveries", c: "1E2761" },
    { n: "2", t: "Debt Reduction", d: "~$1.5B/yr — targeting IG by 2027", c: "2563EB" },
    { n: "3", t: "Share Buybacks", d: "$1.5B authorized — ~4% float", c: "3B82F6" },
    { n: "4", t: "Tech & Growth", d: "Starlink, Kinective, digital infra", c: "60A8D0" },
  ];

  priorities.forEach((p, i) => {
    const py = 3.82 + i * 0.38;
    slide.addShape("rect", {
      x: 0.35, y: py, w: 0.28, h: 0.28,
      fill: { color: p.c }, rectRadius: 0.03,
    });
    slide.addText(p.n, {
      x: 0.35, y: py, w: 0.28, h: 0.28,
      fontSize: 9, fontFace: BODY_FONT, color: "FFFFFF", bold: true,
      align: "center", valign: "middle",
    });
    slide.addText(p.t, {
      x: 0.68, y: py + 0.01, w: 1.3, h: 0.26,
      fontSize: 9, fontFace: BODY_FONT, color: theme.primary, bold: true,
    });
    slide.addText(p.d, {
      x: 2.02, y: py + 0.01, w: 2.5, h: 0.26,
      fontSize: 8.5, fontFace: BODY_FONT, color: "5A6A94",
    });
  });

  // FCF thesis and buyback
  insightCard(slide, 7.05, 0.88, 2.6, 1.3,
    "FCF THESIS",
    "FY2025 FCF of $2.1B = first major positive FCF year since pre-COVID. 2026E ~$2.0B despite higher capex ($4.5B fleet). FCF yield ~5% on current market cap — compelling for a growing carrier.",
    false);

  insightCard(slide, 7.05, 2.32, 2.6, 1.2,
    "$1.5B BUYBACK PROGRAM",
    "Board authorized $1.5B repurchase program in Q2 2025. ~4% of shares outstanding. Demonstrates management confidence in FCF durability. No dividend — all return via buyback.",
    false);

  insightCard(slide, 7.05, 3.66, 2.6, 1.05,
    "CAPEX $4.5B (2026E)",
    "Aircraft pre-deliveries + PDPs + maintenance capex. 787 deposits accelerated as Boeing works through backlog. Partially offset by sale-leaseback transactions.",
    false);
}

// ════════════════════════════════════════════════════════════════
// SLIDE 15: ESG & Sustainability
// ════════════════════════════════════════════════════════════════
function slide15(pres) {
  const slide = pres.addSlide();
  slide.background = { color: "F0F4FA" };
  pageBadge(slide, pres, 15);

  slide.addShape("rect", { x: 0, y: 0, w: 10, h: 0.06, fill: { color: theme.accent } });

  slide.addText("ESG & Sustainability", {
    x: 0.35, y: 0.12, w: 7, h: 0.42,
    fontSize: 22, fontFace: TITLE_FONT,
    color: theme.primary, bold: true,
  });
  sectionLabel(slide, "Emissions, SAF & Commitments — Honest Framework", 0.35, 0.58, 7, false);

  // Emissions table
  sectionLabel(slide, "GHG Emissions Trend", 0.35, 0.88, 4.5, false);

  const emissRows = [
    headerRow(["Year", "CO₂ (MT)", "CO₂/RTM (g)", "SAF % of Fuel", "Fleet Avg Age"]),
    dataRow(["2019", "45.2M", "1,060", "0.1%", "14.2 yrs"], { alt: false, light: true }),
    dataRow(["2023", "41.8M", "980", "0.3%", "13.4 yrs"], { alt: true, light: true }),
    dataRow(["2024", "43.1M", "965", "0.5%", "12.8 yrs"], { alt: false, light: true }),
    dataRow(["2025", "44.2M", "950", "0.8%", "12.1 yrs"], { alt: true, light: true }),
    dataRow(["2030 Target", "~38.0M", "~820", "10%+", "~11.0 yrs"], { alt: false, light: true }),
  ];

  slide.addTable(emissRows, makeTableOptsLight(0.35, 1.18, 4.5, [0.9, 1.2, 1.2, 1.2, 1.0], { rowH: 0.3 }));

  // SAF table
  sectionLabel(slide, "SAF Supply & Partnerships", 5.1, 0.88, 4.55, false);

  const safRows = [
    headerRow(["Partner", "SAF Volume (2025)", "2030 Commitment", "Feedstock"]),
    dataRow(["World Energy", "10M gal", "50M gal/yr", "Agricultural waste"], { alt: false, light: true }),
    dataRow(["bp / Air bp", "8M gal", "40M gal/yr", "HEFA process"], { alt: true, light: true }),
    dataRow(["Neste", "6M gal", "30M gal/yr", "Waste fats/oils"], { alt: false, light: true }),
    dataRow(["Tallgrass / Blackrock", "3M gal", "25M gal/yr", "Ethanol pathway"], { alt: true, light: true }),
    dataRow(["Total SAF", "27M gal", "145M gal/yr", "~10% of fuel need"], { alt: false, light: true }),
  ];

  slide.addTable(safRows, makeTableOptsLight(5.1, 1.18, 4.55, [1.7, 1.2, 1.4, 1.25], { rowH: 0.3 }));

  // Commitments card
  insightCard(slide, 0.35, 3.58, 4.5, 1.12,
    "ESG COMMITMENTS",
    "Net zero by 2050 (SBTi-aligned). 50% emissions reduction per RTM by 2035 vs 2019. 100% SAF-capable by 2030. $200M+ invested in SAF offtake agreements. Board ESG committee oversight.",
    false);

  // Honest framing card
  insightCard(slide, 5.1, 3.58, 4.55, 1.12,
    "HONEST FRAMING",
    "Absolute emissions grew in 2024–2025 as capacity recovered to/beyond 2019 levels. Intensity (per RTM) improvements are real but do not offset volume growth yet. SAF at 0.8% of fuel in 2025 — 10% by 2030 requires ~12x scaling. Structural decarbonization remains a 2035–2050 story.",
    false);
}

// ════════════════════════════════════════════════════════════════
// SLIDE 16: 2026 Guidance & Risk Matrix
// ════════════════════════════════════════════════════════════════
function slide16(pres) {
  const slide = pres.addSlide();
  slide.background = { color: theme.bg };
  pageBadge(slide, pres, 16);

  slide.addShape("rect", { x: 0, y: 0, w: 10, h: 0.06, fill: { color: theme.accent } });

  slide.addText("2026 Guidance & Risk Matrix", {
    x: 0.35, y: 0.15, w: 7, h: 0.45,
    fontSize: 26, fontFace: TITLE_FONT,
    color: "FFFFFF", bold: true,
  });
  sectionLabel(slide, "Management Guidance & Key Risk Scenarios", 0.35, 0.64, 9, true);

  // 6 guidance stat cards (3x2)
  const guidance = [
    { v: "$12–14", l: "FY2026E Adj EPS", s: "vs $10.62 in 2025" },
    { v: "~5.5%", l: "Capacity Growth", s: "ASM year-over-year" },
    { v: "Low +", l: "CASM ex-Fuel", s: "mid-single-digit pressure" },
    { v: "+5–6%", l: "Revenue Growth", s: "vs FY2025 $59.1B" },
    { v: "15%+", l: "EBITDA Margin", s: "vs ~14% in FY2025" },
    { v: "~$4.5B", l: "CapEx Guidance", s: "fleet deliveries + MX" },
  ];

  guidance.forEach((g, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    statCard(slide, 0.35 + col * 3.05, 0.92 + row * 1.1, 2.85, 0.95, g.v, g.l, g.s, true);
  });

  // Risk table
  sectionLabel(slide, "Key Risk Matrix", 0.35, 3.1, 9.3, true);

  const riskRows = [
    headerRow(["Risk Factor", "Probability", "Impact", "Mitigation Strategy"]),
    dataRow(["Fuel Price Spike (+$0.50/gal)", "Medium", { text: "High — $200M P&L", options: { color: "F87171" } }, "25% forward hedged; new-gen fleet efficiency buffer"], { alt: false }),
    dataRow(["Boeing Delivery Delays", "High", { text: "Medium — lost capacity", options: { color: "FBBF24" } }, "Lease extensions + A321XLR optionality; demand > supply"], { alt: true }),
    dataRow(["Macro Recession / GDP –1%", "Low-Med", { text: "High — revenue –8–12%", options: { color: "F87171" } }, "Premium + loyalty mix provides floor; variable cost flexibility"], { alt: false }),
    dataRow(["Competitive Price War", "Low", { text: "Medium — yield –3–5%", options: { color: "FBBF24" } }, "Network depth + MileagePlus loyalty moat; hub slot control"], { alt: true }),
  ];

  slide.addTable(riskRows, makeTableOpts(0.35, 3.38, 9.3, [2.2, 1.1, 2.1, 3.9], { rowH: 0.3 }));

  // Bottom confidence bar
  slide.addShape("rect", {
    x: 0.35, y: 5.42, w: 9.3, h: 0.06,
    fill: { color: theme.primary },
  });
  slide.addText("UAL Management confident in FY2026E guidance range. Assumes no recession, $2.55–$2.75/gal jet fuel, and continued robust Atlantic demand.", {
    x: 0.35, y: 5.29, w: 9.3, h: 0.22,
    fontSize: 8, fontFace: BODY_FONT, color: "6080B0",
    align: "center",
  });
}

// ════════════════════════════════════════════════════════════════
// MAIN
// ════════════════════════════════════════════════════════════════
async function main() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.author = "United Airlines Holdings, Inc.";
  pres.title = "UAL FY2025 Results & 2026 Outlook";

  slide01(pres);
  slide02(pres);
  slide03(pres);
  slide04(pres);
  slide05(pres);
  slide06(pres);
  slide07(pres);
  slide08(pres);
  slide09(pres);
  slide10(pres);
  slide11(pres);
  slide12(pres);
  slide13(pres);
  slide14(pres);
  slide15(pres);
  slide16(pres);

  await pres.writeFile({ fileName: "outputs/corporate-minimax.pptx" });
  console.log("✓ Wrote outputs/corporate-minimax.pptx");
}

main().catch((e) => { console.error(e); process.exit(1); });
