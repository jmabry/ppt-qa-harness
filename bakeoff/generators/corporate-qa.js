'use strict';
const PptxGenJS = require('pptxgenjs');

// ─── Color System ───────────────────────────────────────────────────────────
const C = {
  navy:      '1E3A5F',
  gold:      'C9A84C',
  lightBlue: 'D6E4F0',
  darkSlate: '1A2332',
  white:     'FFFFFF',
  textDark:  '1A2332',
  textMed:   '374151',
  textLight: '6B7280',
  navyLight: '2A4F7F',
  goldLight: 'E8C872',
  redAlert:  'C0392B',
};

// ─── Helpers ────────────────────────────────────────────────────────────────
function makeShadow() {
  return { type: 'outer', color: '000000', blur: 4, offset: 2, angle: 45, opacity: 0.25 };
}

function titleSlide(prs, bg) {
  const sld = prs.addSlide();
  sld.background = { color: bg || C.navy };
  return sld;
}

function contentSlide(prs) {
  const sld = prs.addSlide();
  sld.background = { color: C.white };
  return sld;
}

function addNavyHeader(sld, title, subtitle) {
  // Full-width navy header band
  sld.addShape(prs.ShapeType.rect, {
    x: 0, y: 0, w: 10, h: 1.05,
    fill: { color: C.navy },
  });
  sld.addText(title, {
    x: 0.3, y: 0.08, w: 9.4, h: 0.65,
    fontSize: 26, bold: true, color: C.white, fontFace: 'Calibri',
    valign: 'middle',
  });
  if (subtitle) {
    sld.addText(subtitle, {
      x: 0.3, y: 0.68, w: 9.4, h: 0.32,
      fontSize: 11, color: C.goldLight, fontFace: 'Calibri Light', valign: 'top',
    });
  }
  // Gold strip under header
  sld.addShape(prs.ShapeType.rect, {
    x: 0, y: 1.05, w: 10, h: 0.055,
    fill: { color: C.gold },
  });
}

// Helper to add a callout box
function addCallout(sld, x, y, w, h, label, bodyLines, bgColor, textColor) {
  bgColor  = bgColor  || C.lightBlue;
  textColor = textColor || C.textDark;
  sld.addShape(prs.ShapeType.rect, {
    x, y, w, h,
    fill: { color: bgColor },
    line: { color: bgColor },
    shadow: makeShadow(),
  });
  const arr = [];
  if (label) {
    arr.push({ text: label + '\n', options: { bold: true, fontSize: 10, color: textColor, breakLine: false } });
  }
  bodyLines.forEach((ln, i) => {
    arr.push({ text: ln, options: { fontSize: 9, color: textColor, breakLine: i < bodyLines.length - 1 } });
  });
  sld.addText(arr, { x: x + 0.1, y: y + 0.08, w: w - 0.2, h: h - 0.16, valign: 'top', wrap: true });
}

// Helper: simple styled table
function styledTable(sld, rows, x, y, w, colW, rowH, headColor, headTextColor, bodyBg, bodyText, fontSize) {
  fontSize = fontSize || 8.5;
  headColor     = headColor     || C.navy;
  headTextColor = headTextColor || C.white;
  bodyBg        = bodyBg        || C.white;
  bodyText      = bodyText      || C.textDark;

  // Build pptxgenjs table rows
  const tableRows = rows.map((row, rIdx) => {
    return row.map((cell, cIdx) => {
      const isHead = rIdx === 0;
      const isAlt  = !isHead && rIdx % 2 === 0;
      return {
        text: String(cell),
        options: {
          bold: isHead,
          fontSize,
          color: isHead ? headTextColor : bodyText,
          fill: isHead ? headColor : (isAlt ? 'EDF4FB' : bodyBg),
          align: cIdx === 0 ? 'left' : 'center',
          valign: 'middle',
          margin: [3, 4, 3, 4],
        },
      };
    });
  });

  sld.addTable(tableRows, {
    x, y, w,
    colW,
    rowH,
    border: { type: 'solid', color: 'D0DCE8', pt: 0.5 },
  });
}

// ─── Main ────────────────────────────────────────────────────────────────────
const prs = new PptxGenJS();
prs.layout = 'LAYOUT_WIDE'; // 10 × 5.625

// ════════════════════════════════════════════════════════════
// SLIDE 1 — Title
// ════════════════════════════════════════════════════════════
{
  const sld = prs.addSlide();
  sld.background = { color: C.navy };

  // Dark-slate top strip
  sld.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: 10, h: 0.35, fill: { color: C.darkSlate } });

  // Top label
  sld.addText('INSTITUTIONAL INVESTOR PRESENTATION', {
    x: 0.4, y: 0.05, w: 9.2, h: 0.27,
    fontSize: 9, color: C.goldLight, fontFace: 'Calibri', bold: true, charSpacing: 2,
  });

  // UAL logo circle
  sld.addShape(prs.ShapeType.ellipse, {
    x: 0.45, y: 0.8, w: 1.1, h: 1.1,
    fill: { color: C.gold }, shadow: makeShadow(),
  });
  sld.addText('UAL', {
    x: 0.45, y: 0.95, w: 1.1, h: 0.7,
    fontSize: 22, bold: true, color: C.navy, fontFace: 'Calibri', align: 'center',
  });

  // Main title
  sld.addText('United Airlines Holdings', {
    x: 0.4, y: 1.75, w: 9.2, h: 1.1,
    fontSize: 48, bold: true, color: C.white, fontFace: 'Calibri',
  });

  // Subtitle
  sld.addText('FY2025 Results & 2026 Outlook', {
    x: 0.4, y: 2.8, w: 9.2, h: 0.6,
    fontSize: 26, color: C.goldLight, fontFace: 'Calibri Light',
  });

  // Gold accent bar
  sld.addShape(prs.ShapeType.rect, { x: 0, y: 3.6, w: 10, h: 0.08, fill: { color: C.gold } });

  // Details row
  sld.addShape(prs.ShapeType.rect, { x: 0, y: 3.68, w: 10, h: 0.9, fill: { color: C.darkSlate } });

  sld.addText([
    { text: 'UAL  |  NASDAQ  |  January 20, 2026', options: { fontSize: 13, color: C.white, bold: false } },
  ], { x: 0.4, y: 3.78, w: 9.2, h: 0.55, valign: 'middle' });

  // Key stats strip — fits within 5.625" slide height
  const stats = [
    ['$59.1B', 'FY2025 Revenue'],
    ['$10.62', 'Adj. EPS'],
    ['181.1M', 'Passengers'],
    ['1,490', 'Fleet Size'],
    ['$12–$14', '2026E EPS'],
  ];
  const sw = 10 / stats.length;
  stats.forEach(([val, lbl], i) => {
    sld.addShape(prs.ShapeType.rect, {
      x: i * sw, y: 4.5, w: sw, h: 1.12,
      fill: { color: i % 2 === 0 ? C.navyLight : C.navy },
    });
    sld.addText(val, {
      x: i * sw + 0.05, y: 4.54, w: sw - 0.1, h: 0.48,
      fontSize: 18, bold: true, color: C.gold, fontFace: 'Calibri', align: 'center',
    });
    sld.addText(lbl, {
      x: i * sw + 0.05, y: 5.02, w: sw - 0.1, h: 0.26,
      fontSize: 9, color: C.lightBlue, fontFace: 'Calibri Light', align: 'center',
    });
  });
}

// ════════════════════════════════════════════════════════════
// SLIDE 2 — Investment Thesis
// ════════════════════════════════════════════════════════════
{
  const sld = contentSlide(prs);
  addNavyHeader(sld, 'Investment Thesis', 'Four pillars supporting UAL re-rating');

  const pillars = [
    {
      title: '1. Scale & Hub Dominance',
      lines: [
        'Largest U.S. airline by ASMs (330.3B)',
        '#1 at ORD, DEN, IAH, EWR, SFO, IAD',
        'Only U.S. carrier to Bangkok, Ho Chi Minh City, Adelaide',
        'Network moat compounds with every new route',
      ],
    },
    {
      title: '2. Premium / Loyalty Revenue Quality',
      lines: [
        'Premium cabin +11% in 2025 — growing 2× basic economy',
        '130M MileagePlus members',
        'Chase co-brand adding 1M+ cards/yr',
        'Signature Interior rolling out across fleet',
      ],
    },
    {
      title: '3. Balance Sheet Normalization',
      lines: [
        'Net leverage 2.2× (Dec\'25) → <2.0× (Dec\'26)',
        'MileagePlus bonds fully retired',
        'Investment-grade trajectory',
        'Interest expense –16% YoY',
      ],
    },
    {
      title: '4. 2026 EPS Inflection',
      lines: [
        '$12–$14 EPS guidance (vs. $10.62 in 2025): +13%–+32%',
        'First time materially exceeding 2019 peak of $12.05',
        'Labor headwinds largely absorbed',
        'TRASM inflection + aircraft efficiency = margin expansion',
      ],
    },
  ];

  const boxW = 4.75;
  const boxH = 1.95;
  const positions = [
    { x: 0.25, y: 1.2 },
    { x: 5.05, y: 1.2 },
    { x: 0.25, y: 3.22 },
    { x: 5.05, y: 3.22 },
  ];
  const colors = [C.navy, C.navyLight, C.navyLight, C.navy];

  pillars.forEach((p, i) => {
    const { x, y } = positions[i];
    sld.addShape(prs.ShapeType.rect, {
      x, y, w: boxW, h: boxH,
      fill: { color: colors[i] },
      shadow: makeShadow(),
    });
    // Gold title bar
    sld.addShape(prs.ShapeType.rect, {
      x, y, w: boxW, h: 0.35,
      fill: { color: C.gold },
    });
    sld.addText(p.title, {
      x: x + 0.12, y: y + 0.04, w: boxW - 0.24, h: 0.3,
      fontSize: 10, bold: true, color: C.darkSlate, fontFace: 'Calibri',
    });
    p.lines.forEach((ln, j) => {
      sld.addText('\u2022  ' + ln, {
        x: x + 0.15, y: y + 0.42 + j * 0.31, w: boxW - 0.3, h: 0.29,
        fontSize: 9, color: C.white, fontFace: 'Calibri Light',
      });
    });
  });

  // Bottom tagline
  sld.addShape(prs.ShapeType.rect, { x: 0, y: 5.17, w: 10, h: 0.455, fill: { color: C.lightBlue } });
  sld.addText(
    'UAL enters 2026 with labor contracts settled, the largest delivery pipeline in its history, record booking trends, and a balance sheet on track for investment-grade — the convergence of tailwinds that have been building since 2022.',
    { x: 0.3, y: 5.19, w: 9.4, h: 0.42, fontSize: 9, color: C.textDark, italic: true, valign: 'middle', wrap: true }
  );
}

// ════════════════════════════════════════════════════════════
// SLIDE 3 — Multi-Year Financial Performance
// ════════════════════════════════════════════════════════════
{
  const sld = contentSlide(prs);
  addNavyHeader(sld, 'Recovery Arc: 2021–2026E', 'Six-year revenue and earnings trajectory');

  const rows = [
    ['Metric', 'FY2021', 'FY2022', 'FY2023', 'FY2024', 'FY2025', 'FY2026E'],
    ['Total Revenue ($B)', '$24.6', '$45.0', '$53.7', '$57.1', '$59.1', '—'],
    ['Adj. EPS', 'Neg.', '$10.61', '$10.05', '$10.61', '$10.62', '$12–$14'],
    ['Adj. Pre-Tax Margin', 'Neg.', '~6%', '8.0%', '8.1%', '7.8%', '~10%+'],
    ['ASMs (B)', '178.7', '247.9', '291.3', '311.2', '330.3', '~349'],
    ['Fleet', '—', '~1,300', '1,358', '1,406', '1,490', '~1,610E'],
  ];

  styledTable(sld, rows, 0.25, 1.18, 9.5,
    [2.2, 1.22, 1.22, 1.22, 1.22, 1.22, 1.2],
    0.41, C.navy, C.white, C.white, C.textDark, 9
  );

  // Revenue growth visual bar strip
  const yrs = ['2021', '2022', '2023', '2024', '2025', '2026E'];
  const revs = [24.6, 45.0, 53.7, 57.1, 59.1, 63.0];
  const maxRev = 65;
  const barX0 = 0.55, barY0 = 3.6, barW = 9.2, barH = 0.72;
  const bw = barW / yrs.length - 0.09;
  sld.addText('Revenue ($B) — scaled to $65B', {
    x: 0.55, y: 3.47, w: 6, h: 0.15, fontSize: 9, color: C.textLight, italic: true,
  });
  // Y-axis unit label
  sld.addText('Revenue ($B)', {
    x: 0.0, y: 3.55, w: 0.55, h: 0.72,
    fontSize: 9, color: C.textLight, italic: true, align: 'center', valign: 'middle',
    rotate: 270,
  });
  yrs.forEach((yr, i) => {
    const frac = revs[i] / maxRev;
    const bh = barH * frac;
    const bx = barX0 + i * (bw + 0.09);
    const by = barY0 + barH - bh;
    sld.addShape(prs.ShapeType.rect, {
      x: bx, y: by, w: bw, h: bh,
      fill: { color: i === 5 ? C.gold : (i === 4 ? C.navyLight : C.navyLight) },
    });
    // Always place label inside bar (all bars tall enough at this scale)
    sld.addText('$' + revs[i] + 'B', {
      x: bx, y: by + 0.04, w: bw, h: 0.18,
      fontSize: 9, color: i === 5 ? C.darkSlate : C.white, bold: true, align: 'center',
    });
    sld.addText(yr, {
      x: bx, y: barY0 + barH + 0.03, w: bw, h: 0.18, fontSize: 9, color: C.textMed, align: 'center',
    });
  });

  // Key callout — extended to fill slide bottom
  sld.addShape(prs.ShapeType.rect, {
    x: 0.25, y: 4.52, w: 9.5, h: 1.1,
    fill: { color: C.gold }, shadow: makeShadow(),
  });
  sld.addText(
    'EPS Plateau 2022–2025: $10.61 → $10.05 → $10.61 → $10.62 reflects absorption of $10B+ in cumulative pilot/FA contract obligations — NOT operational underperformance. Revenue +31% since 2022. The plateau ends in 2026.',
    { x: 0.4, y: 4.55, w: 9.2, h: 1.04, fontSize: 9, color: C.darkSlate, bold: false, italic: true, valign: 'middle', wrap: true }
  );
}

// ════════════════════════════════════════════════════════════
// SLIDE 4 — FY2025 Quarterly P&L
// ════════════════════════════════════════════════════════════
{
  const sld = contentSlide(prs);
  addNavyHeader(sld, 'FY2025 Quarterly Performance', 'Revenue, earnings, and margin by quarter');

  const rows = [
    ['Quarter', 'Total Revenue', 'YoY', 'Net Income', 'Adj. EPS', 'Pre-Tax Margin', 'ASMs (B)'],
    ['Q1 2025', '$13.213B', '+5.4%', '$0.387B', '$0.91', '3.0%', '75.155'],
    ['Q2 2025', '$15.236B', '+1.7%', '$0.973B', '$3.87', '11.0%', '84.347'],
    ['Q3 2025', '$15.225B', '+2.6%', '$0.949B', '$2.78', '8.0%', '87.417'],
    ['Q4 2025', '$15.394B', '+4.8%', '$1.024B', '$3.10', '8.5%', '83.365'],
    ['FY 2025', '$59.068B', '+3.5%', '$3.400B', '$10.62', '7.8%', '330.284'],
  ];

  styledTable(sld, rows, 0.25, 1.18, 9.5,
    [1.5, 1.55, 0.9, 1.3, 1.2, 1.6, 1.45],
    0.4, C.navy, C.white, C.white, C.textDark, 9
  );

  // Three callout boxes
  const callouts = [
    {
      label: 'Q4 Context',
      body: 'Revenue +4.8% but EPS –4.9% vs. Q4\'24 ($3.10 vs. $3.26). Comps problem + temporary domestic fare softness, not structural.',
      bg: C.lightBlue,
    },
    {
      label: 'Guidance Story',
      body: 'Initial $11.50–$13.50 → revised to $9.00–$11.00 mid-year → delivered $10.62 at/near top of revised range.',
      bg: C.navy,
    },
    {
      label: 'Q2 GAAP vs. Adj.',
      body: '$447M fleet write-offs pulled GAAP EPS below adjusted; retirements economically rational.',
      bg: C.gold,
    },
  ];

  callouts.forEach((c, i) => {
    const x = 0.25 + i * 3.22;
    const boxH = 5.47 - 3.9; // fill to near slide bottom
    sld.addShape(prs.ShapeType.rect, {
      x, y: 3.9, w: 3.12, h: boxH,
      fill: { color: c.bg }, shadow: makeShadow(),
    });
    // Contrasting title strip
    sld.addShape(prs.ShapeType.rect, {
      x, y: 3.9, w: 3.12, h: 0.32,
      fill: { color: c.bg === C.gold ? C.darkSlate : C.gold },
    });
    sld.addText(c.label, {
      x: x + 0.1, y: 3.9, w: 3.0, h: 0.32,
      fontSize: 9.5, bold: true, color: c.bg === C.gold ? C.goldLight : C.darkSlate,
      fontFace: 'Calibri', valign: 'middle',
    });
    sld.addText(c.body, {
      x: x + 0.12, y: 4.26, w: 2.9, h: boxH - 0.4,
      fontSize: 9, color: c.bg === C.navy ? C.white : C.darkSlate, wrap: true, valign: 'top',
    });
  });
}

// ════════════════════════════════════════════════════════════
// SLIDE 5 — Unit Economics Deep Dive
// ════════════════════════════════════════════════════════════
{
  const sld = contentSlide(prs);
  addNavyHeader(sld, 'Unit Economics: The TRASM/CASM Story', 'Annual trends and 2025 quarterly breakdown');

  // Annual table
  sld.addShape(prs.ShapeType.rect, { x: 0.25, y: 1.18, w: 2.5, h: 0.22, fill: { color: C.navy } });
  sld.addText('Annual Trends', { x: 0.25, y: 1.18, w: 2.5, h: 0.22, fontSize: 9, bold: true, color: C.white, align: 'center', valign: 'middle' });

  const annualRows = [
    ['Year', 'TRASM', 'PRASM', 'CASM-ex', 'Fuel/Gal', 'Load Factor'],
    ['FY2023', '18.44¢', '—', '12.03¢', '$3.01', '83.9%'],
    ['FY2024', '18.34¢', '16.66¢', '12.58¢', '$2.65', '83.1%'],
    ['FY2025', '17.88¢', '16.18¢', '12.64¢', '$2.44', '82.2%'],
  ];
  styledTable(sld, annualRows, 0.25, 1.4, 5.8,
    [1.0, 0.92, 0.92, 0.96, 0.96, 1.04], 0.37,
    C.navyLight, C.white, C.white, C.textDark, 8.5
  );

  // Quarterly table
  sld.addShape(prs.ShapeType.rect, { x: 0.25, y: 2.98, w: 2.8, h: 0.22, fill: { color: C.navy } });
  sld.addText('2025 Quarterly', { x: 0.25, y: 2.98, w: 2.8, h: 0.22, fontSize: 9, bold: true, color: C.white, align: 'center', valign: 'middle' });

  const qRows = [
    ['Quarter', 'TRASM', 'PRASM', 'CASM-ex', 'Fuel/Gal', 'Load Factor'],
    ['Q1', '17.58¢', '15.78¢', '13.17¢', '$2.53', '79.2%'],
    ['Q2', '18.06¢', '16.40¢', '12.36¢', '$2.34', '83.1%'],
    ['Q3', '17.42¢', '15.80¢', '12.15¢', '$2.43', '84.4%'],
    ['Q4', '18.47¢', '16.71¢', '12.94¢', '$2.49', '81.9%'],
  ];
  styledTable(sld, qRows, 0.25, 3.2, 5.8,
    [1.0, 0.92, 0.92, 0.96, 0.96, 1.04], 0.36,
    C.navyLight, C.white, C.white, C.textDark, 8.5
  );

  // Right-side callouts
  // TRASM callout
  sld.addShape(prs.ShapeType.rect, { x: 6.3, y: 1.18, w: 3.45, h: 1.32, fill: { color: C.lightBlue }, shadow: makeShadow() });
  sld.addText('TRASM –2.5% in 2025: NOT a demand problem', { x: 6.42, y: 1.23, w: 3.22, h: 0.3, fontSize: 9, bold: true, color: C.navy });
  sld.addText('Record 181M passengers. Industry capacity grew faster than unit revenue. Response: UAL pulled 4 ppts domestic capacity mid-year, retired 21 aircraft early.', {
    x: 6.42, y: 1.54, w: 3.22, h: 0.9, fontSize: 9, color: C.textDark, wrap: true, valign: 'top',
  });

  // Fuel sensitivity
  sld.addShape(prs.ShapeType.rect, { x: 6.3, y: 2.62, w: 3.45, h: 0.82, fill: { color: C.gold }, shadow: makeShadow() });
  sld.addText('Unhedged by Policy', { x: 6.42, y: 2.67, w: 3.22, h: 0.25, fontSize: 9, bold: true, color: C.darkSlate });
  sld.addText('$0.10/gal = ~$466M annual P&L impact at 4.663B gallons consumed', {
    x: 6.42, y: 2.93, w: 3.22, h: 0.46, fontSize: 9, color: C.darkSlate, wrap: true,
  });

  // 2026 thesis
  sld.addShape(prs.ShapeType.rect, { x: 6.3, y: 3.55, w: 3.45, h: 1.95, fill: { color: C.navy }, shadow: makeShadow() });
  sld.addText('2026 Thesis', { x: 6.42, y: 3.6, w: 3.22, h: 0.28, fontSize: 9, bold: true, color: C.gold });
  sld.addText('TRASM inflects positive as industry capacity normalizes → CASM-ex stabilizes with aircraft deliveries → gap widens = margin expansion toward 10%+ pre-tax margin target', {
    x: 6.42, y: 3.9, w: 3.22, h: 1.52, fontSize: 9, color: C.white, wrap: true, valign: 'top',
  });
}

// ════════════════════════════════════════════════════════════
// SLIDE 6 — Operating Cost Bridge
// ════════════════════════════════════════════════════════════
{
  const sld = contentSlide(prs);
  addNavyHeader(sld, 'Cost Structure: Composition Matters', 'FY2025 vs. FY2024 operating expense detail');

  const rows = [
    ['Line Item', 'FY2025', 'FY2024', 'YoY $', 'YoY %', 'Interpretation'],
    ['Salaries & related', '$17.647B', '$16.678B', '+$969M', '+5.8%', 'Bounded — pilot contract peaks in early years'],
    ['Aircraft fuel', '$11.396B', '$11.756B', '–$360M', '–3.1%', 'Genuine tailwind; new aircraft insulates further'],
    ['Landing fees & rent', '$3.849B', '$3.437B', '+$412M', '+12.0%', 'Good cost: EWR T-A, ORD, DEN upgrades'],
    ['Maintenance', '$3.294B', '$3.063B', '+$231M', '+7.5%', 'Fleet growth + aging aircraft in transition'],
    ['Depreciation', '$2.939B', '$2.928B', '+$11M', '+0.4%', 'Stable'],
    ['Regional capacity', '$2.693B', '$2.516B', '+$177M', '+7.0%', 'Volume growth'],
    ['Distribution', '$2.109B', '$2.231B', '–$122M', '–5.5%', 'Tailwind: direct channel shift'],
    ['Other', '$9.919B', '$9.053B', '+$866M', '+9.6%', '—'],
    ['Total OpEx', '$54.356B', '$51.967B', '+$2.389B', '+4.6%', 'vs. Revenue +3.5%'],
  ];

  styledTable(sld, rows, 0.25, 1.18, 9.5,
    [1.9, 1.15, 1.15, 1.0, 0.8, 3.5],
    0.36, C.navy, C.white, C.white, C.textDark, 8
  );

  // Three callout boxes at bottom
  const cbs = [
    { lbl: 'Labor', txt: 'Up $2.9B in 2 yrs. 32.5% of OpEx. Gauge-up means labor cost per SEAT rises less than cost per employee — 150 pax vs. 50 on replaced planes.', bg: C.lightBlue },
    { lbl: 'Landing Fees', txt: '+12% = UAL\'s own $2.7B EWR Terminal A + ORD T2 satellite + DEN gates. Premium product in OpEx, not just CapEx.', bg: C.navyLight },
    { lbl: 'Distribution –5.5%', txt: 'Direct channels handle 85%+ of check-ins. GDS fee avoidance has legs as app engagement accelerates.', bg: C.gold },
  ];
  cbs.forEach((c, i) => {
    const x = 0.25 + i * 3.22;
    sld.addShape(prs.ShapeType.rect, { x, y: 4.95, w: 3.12, h: 0.55, fill: { color: c.bg }, shadow: makeShadow() });
    sld.addText(c.lbl + ': ' + c.txt, {
      x: x + 0.1, y: 4.97, w: 2.95, h: 0.51,
      fontSize: 9, color: c.bg === C.navyLight ? C.white : C.darkSlate, wrap: true, valign: 'middle',
    });
  });
}

// ════════════════════════════════════════════════════════════
// SLIDE 7 — Regional Revenue Mix
// ════════════════════════════════════════════════════════════
{
  const sld = contentSlide(prs);
  addNavyHeader(sld, 'Two-Speed Airline: Regional Revenue Trends', 'International outperformance driving revenue quality improvement');

  const rows = [
    ['Region', 'FY2023', 'FY2024', 'FY2025E', '3-Yr CAGR'],
    ['Domestic', '$31.874B', '$28.458B', '~$31.6B', '–0.4%'],
    ['Atlantic — Europe', '$8.068B', '$8.099B', '~$10.1B', '+11.9%'],
    ['Atlantic — MEIA', '$1.000B', '$1.141B', '~$1.3B', '+14.0%'],
    ['Pacific', '$4.437B', '$5.226B', '~$6.88B', '+24.5%'],
    ['Latin America', '$3.667B', '$4.827B', '~$5.53B', '+22.8%'],
  ];

  styledTable(sld, rows, 0.25, 1.18, 5.5,
    [1.9, 0.9, 0.9, 0.9, 0.9], 0.41,
    C.navy, C.white, C.white, C.textDark, 9
  );

  // Right: Q4 spotlight callouts
  const spots = [
    { region: 'MEIA', color: C.navy, txt: '+58.7% Q4 revenue | +56% capacity | EWR-Dubai, EWR-Delhi, SFO-Bangalore ramping | India = structural opportunity' },
    { region: 'Pacific', color: C.navyLight, txt: '+10.1% Q4 | Only U.S. carrier to Bangkok, Adelaide, Ho Chi Minh City | Monopoly routes support yield' },
    { region: 'Domestic', color: C.lightBlue, txt: '+2.0% Q4 revenue but PRASM –1.9% | Soft all year | Mid-year capacity discipline in progress' },
    { region: 'Latin', color: C.textMed, txt: '+0.5% Q4 but PRASM –7.6% | Brazilian competition + macro | ~10% of pax revenue' },
  ];

  spots.forEach((s, i) => {
    const x = 5.95;
    const y = 1.18 + i * 1.02;
    const boxH = i === 3 ? 1.25 : 0.95;
    sld.addShape(prs.ShapeType.rect, { x, y, w: 3.8, h: boxH, fill: { color: s.color }, shadow: makeShadow() });
    sld.addText(s.region, { x: x + 0.12, y: y + 0.07, w: 3.6, h: 0.26, fontSize: 10, bold: true, color: s.color === C.lightBlue ? C.navy : C.gold });
    sld.addText(s.txt, { x: x + 0.12, y: y + 0.35, w: 3.6, h: boxH - 0.39, fontSize: 9, color: s.color === C.lightBlue ? C.textDark : C.white, wrap: true, valign: 'top' });
  });

  // Growth drivers callout — fills the left-side empty region below the table
  sld.addShape(prs.ShapeType.rect, { x: 0.25, y: 3.72, w: 5.5, h: 1.05, fill: { color: C.navyLight }, shadow: makeShadow() });
  sld.addText('2026 Route Expansion', { x: 0.37, y: 3.77, w: 5.28, h: 0.26, fontSize: 9.5, bold: true, color: C.gold });
  sld.addText('New routes in 2025–2026: EWR-Dubai (daily), EWR-Delhi (daily), SFO-Bangalore, SFO-Singapore (787-9 Elevated April 2026), ORD-Melbourne. Pacific + MEIA corridors represent the highest PRASM expansion opportunity in the portfolio.', {
    x: 0.37, y: 4.05, w: 5.28, h: 0.68, fontSize: 11, color: C.white, wrap: true, valign: 'top',
  });

  // A321XLR callout — full-width, extended to fill slide bottom
  sld.addShape(prs.ShapeType.rect, { x: 0.25, y: 4.85, w: 9.5, h: 0.62, fill: { color: C.gold } });
  sld.addText('A321XLR (4,700 nm range) unlocks routes impossible for 757 or A321neo: EWR to Bari, Split, Glasgow, Santiago de Compostela — point-to-point international with narrowbody economics.', {
    x: 0.4, y: 4.87, w: 9.2, h: 0.58, fontSize: 9, color: C.darkSlate, italic: true, valign: 'middle', wrap: true,
  });
}

// ════════════════════════════════════════════════════════════
// SLIDE 8 — Revenue Quality: Premium & Loyalty
// ════════════════════════════════════════════════════════════
{
  const sld = contentSlide(prs);
  addNavyHeader(sld, 'Revenue Quality: Premium & Loyalty Flywheel', 'Structural shift to higher-margin revenue streams');

  // Left panel — Premium
  sld.addShape(prs.ShapeType.rect, { x: 0.25, y: 1.18, w: 4.5, h: 0.3, fill: { color: C.navy } });
  sld.addText('Premium Cabin Performance', { x: 0.25, y: 1.18, w: 4.5, h: 0.3, fontSize: 10, bold: true, color: C.white, align: 'center', valign: 'middle' });

  const premStats = [
    ['Premium cabin FY2025', '+11% YoY'],
    ['Basic Economy FY2025', '+5% YoY'],
    ['Premium seats flown', '27.4M (12% of seats)'],
    ['Premium seats/N.Am. dep. since 2021', '+40%'],
    ['Narrowbodies w/ Signature Interior', '119 (68% of fleet)'],
    ['Signature Interior NPS premium', '+10 pts'],
  ];
  premStats.forEach(([lbl, val], i) => {
    const bg = i % 2 === 0 ? C.lightBlue : 'EDF4FB';
    sld.addShape(prs.ShapeType.rect, { x: 0.25, y: 1.48 + i * 0.42, w: 4.5, h: 0.4, fill: { color: bg } });
    sld.addText(lbl, { x: 0.35, y: 1.52 + i * 0.42, w: 2.8, h: 0.32, fontSize: 9, color: C.textDark, valign: 'middle' });
    sld.addText(val, { x: 3.2, y: 1.52 + i * 0.42, w: 1.5, h: 0.32, fontSize: 9, bold: true, color: C.navy, align: 'right', valign: 'middle' });
  });

  // Right panel — MileagePlus
  sld.addShape(prs.ShapeType.rect, { x: 5.0, y: 1.18, w: 4.75, h: 0.3, fill: { color: C.navyLight } });
  sld.addText('MileagePlus / Loyalty', { x: 5.0, y: 1.18, w: 4.75, h: 0.3, fontSize: 10, bold: true, color: C.white, align: 'center', valign: 'middle' });

  const loyaltyStats = [
    ['Total members', '130M+'],
    ['Chase co-brand spend growth Q4', '+14%'],
    ['Chase co-brand spend growth FY2025', '+12%'],
    ['New cards/year (3rd consecutive yr)', '1M+'],
    ['Chase deal expiry', '2029'],
    ['CEO target: MileagePlus profit growth', '2× by 2030'],
  ];
  loyaltyStats.forEach(([lbl, val], i) => {
    const bg = i % 2 === 0 ? C.lightBlue : 'EDF4FB';
    sld.addShape(prs.ShapeType.rect, { x: 5.0, y: 1.48 + i * 0.42, w: 4.75, h: 0.4, fill: { color: bg } });
    sld.addText(lbl, { x: 5.1, y: 1.52 + i * 0.42, w: 3.2, h: 0.32, fontSize: 9, color: C.textDark, valign: 'middle' });
    sld.addText(val, { x: 8.25, y: 1.52 + i * 0.42, w: 1.4, h: 0.32, fontSize: 9, bold: true, color: C.navy, align: 'right', valign: 'middle' });
  });

  // Gold valuation callout
  sld.addShape(prs.ShapeType.rect, { x: 0.25, y: 4.25, w: 9.5, h: 1.2, fill: { color: C.gold }, shadow: makeShadow() });
  sld.addText('MileagePlus Standalone Valuation Opportunity', { x: 0.4, y: 4.3, w: 9.2, h: 0.3, fontSize: 11, bold: true, color: C.darkSlate });
  sld.addText(
    'In 2020, bankers valued MileagePlus at ~$22B as loan collateral — and the program has grown since. UAL total market cap today: ~$28B. This implies minimal standalone loyalty value is priced in — a significant re-rating opportunity vs. the Delta/AmEx model. New April 2026: primary cardholders earn 2\u00d7 more miles; 10\u201315% discount on award redemptions.',
    { x: 0.4, y: 4.62, w: 9.2, h: 0.75, fontSize: 9, color: C.darkSlate, wrap: true, valign: 'top' }
  );
}

// ════════════════════════════════════════════════════════════
// SLIDE 9 — Fleet Modernization
// ════════════════════════════════════════════════════════════
{
  const sld = contentSlide(prs);
  addNavyHeader(sld, 'United Next Fleet Transformation', 'Largest delivery pipeline in UAL history');

  // Fleet timeline
  const ftData = [
    { yr: 'Dec\'23', n: 1358 },
    { yr: 'Dec\'24', n: 1406 },
    { yr: 'Dec\'25', n: 1490 },
    { yr: 'Dec\'26E', n: 1610 },
  ];
  ftData.forEach((f, i) => {
    const x = 0.35 + i * 2.38;
    const isLast = i === ftData.length - 1;
    sld.addShape(prs.ShapeType.rect, { x, y: 1.18, w: 2.15, h: 0.8, fill: { color: isLast ? C.gold : C.navy }, shadow: makeShadow() });
    sld.addText(String(f.n), { x, y: 1.22, w: 2.15, h: 0.45, fontSize: 22, bold: true, color: isLast ? C.darkSlate : C.white, align: 'center' });
    sld.addText(f.yr, { x, y: 1.65, w: 2.15, h: 0.25, fontSize: 9, color: isLast ? C.darkSlate : C.lightBlue, align: 'center' });
  });

  // Delivery highlights
  sld.addShape(prs.ShapeType.rect, { x: 0.25, y: 2.1, w: 4.5, h: 0.5, fill: { color: C.lightBlue } });
  sld.addText('2025 Deliveries: 71 narrowbody + 11 widebody = 82 aircraft', { x: 0.35, y: 2.13, w: 4.3, h: 0.44, fontSize: 9, color: C.navy, bold: true, valign: 'middle' });
  sld.addShape(prs.ShapeType.rect, { x: 5.0, y: 2.1, w: 4.75, h: 0.5, fill: { color: C.navyLight } });
  sld.addText('2026 Plan: ~100 NB + ~20 WB = ~120 aircraft (largest U.S. widebody intake since 1988)', { x: 5.1, y: 2.13, w: 4.55, h: 0.44, fontSize: 9, color: C.white, bold: true, valign: 'middle', wrap: true });

  // Order book table
  const rows = [
    ['Aircraft', 'Purpose', 'Orders Remaining'],
    ['Boeing 787-9 Elevated', 'Long-haul premium', '47+ next wave'],
    ['Boeing 737 MAX 8/9/10', 'Domestic narrowbody', '200+'],
    ['Airbus A321neo', 'General domestic', '18+'],
    ['A321neo Coastliner', 'Domestic transcon lie-flat', '50 total, 40 by Apr\'28'],
    ['Airbus A321XLR', 'New intl routes', '50 total, 25+ by 2028'],
    ['CRJ450 (SkyWest)', 'Premium regional', '70 total, 50 by 2028'],
    ['Total United Next', '', '630+ by 2034'],
  ];
  styledTable(sld, rows, 0.25, 2.72, 6.5,
    [2.2, 2.3, 2.0], 0.33,
    C.navy, C.white, C.white, C.textDark, 8.5
  );

  // Right side callouts
  sld.addShape(prs.ShapeType.rect, { x: 6.95, y: 2.72, w: 2.8, h: 1.2, fill: { color: C.lightBlue }, shadow: makeShadow() });
  sld.addText('Gauge-Up Impact', { x: 7.05, y: 2.77, w: 2.6, h: 0.28, fontSize: 9, bold: true, color: C.navy });
  sld.addText('Avg mainline seats/dep: 151 (2019) → 174 (2025). 50-seat RJs → 150+ seat mainline = 20% better fuel/seat, lower crew cost per passenger.', {
    x: 7.05, y: 3.07, w: 2.6, h: 0.8, fontSize: 9, color: C.textDark, wrap: true, valign: 'top',
  });

  sld.addShape(prs.ShapeType.rect, { x: 6.95, y: 4.03, w: 2.8, h: 1.4, fill: { color: C.gold }, shadow: makeShadow() });
  sld.addText('Boeing Risk', { x: 7.05, y: 4.08, w: 2.6, h: 0.28, fontSize: 9, bold: true, color: C.darkSlate });
  sld.addText('FAA 38 MAX/month production cap. MAX 10 certification pending — 200 aircraft waiting. Each month delay = ~$30–40M deferred efficiency. Mitigation: A321neo/XLR Airbus orders provide operational hedge.', {
    x: 7.05, y: 4.38, w: 2.6, h: 0.98, fontSize: 9, color: C.darkSlate, wrap: true, valign: 'top',
  });
}

// ════════════════════════════════════════════════════════════
// SLIDE 10 — New Product Architecture
// ════════════════════════════════════════════════════════════
{
  const sld = contentSlide(prs);
  addNavyHeader(sld, 'New Product Architecture: 2026–2028 Launches', 'Four differentiated products entering service');

  const products = [
    {
      name: '787-9 Elevated',
      when: 'April 2026 | SFO-Singapore, SFO-LHR',
      bullets: [
        '8 Polaris Studio suites (27" 4K OLED)',
        '56 United Polaris + 35 Premium Plus',
        '222 total | 44.6% premium — highest among U.S. widebodies',
        'Target: 30+ aircraft by end 2027',
      ],
      color: C.navy,
    },
    {
      name: 'A321neo Coastliner',
      when: 'Summer 2026 | SFO/LAX \u2194 EWR/JFK',
      bullets: [
        '20 Polaris — FIRST lie-flat on any U.S. domestic narrowbody',
        '12 Premium Plus — FIRST on domestic narrowbody',
        '161 total | 50 ordered, 40 by Apr\'28',
        'Target: 10,000+ pax/day on these corridors',
      ],
      color: C.navyLight,
    },
    {
      name: 'A321XLR',
      when: 'Summer 2026 | New European + S. American routes',
      bullets: [
        '20 Polaris (all-aisle access + privacy door)',
        '12 Premium Plus | ~118 Economy',
        '~150 total | All screens 4K OLED',
        'Range: 4,700 nm — Unlocks Bari, Split, Glasgow, Santiago',
      ],
      color: C.navy,
    },
    {
      name: 'CRJ450 Premium Regional',
      when: 'Fall 2026 | DEN + ORD hubs',
      bullets: [
        '7 United First (37" pitch) — overhead luggage closets (industry first)',
        '16 Economy Plus (33-34") | 18 Economy (31")',
        'Free Starlink for MileagePlus members',
        '70 eventual fleet | 50 by 2028',
      ],
      color: C.navyLight,
    },
  ];

  const bw = 4.75, bh = 2.2;
  products.forEach((p, i) => {
    const x = i % 2 === 0 ? 0.25 : 5.05;
    const y = i < 2 ? 1.18 : 3.2;
    sld.addShape(prs.ShapeType.rect, { x, y, w: bw, h: bh, fill: { color: p.color }, shadow: makeShadow() });
    sld.addShape(prs.ShapeType.rect, { x, y, w: bw, h: 0.32, fill: { color: C.gold } });
    sld.addText(p.name, { x: x + 0.12, y: y + 0.03, w: bw - 0.24, h: 0.28, fontSize: 10, bold: true, color: C.darkSlate });
    sld.addText(p.when, { x: x + 0.12, y: y + 0.36, w: bw - 0.24, h: 0.22, fontSize: 9, color: C.goldLight, italic: true });
    p.bullets.forEach((b, j) => {
      sld.addText('\u2022  ' + b, { x: x + 0.15, y: y + 0.61 + j * 0.32, w: bw - 0.28, h: 0.3, fontSize: 9, color: C.white, wrap: true });
    });
  });
}

// ════════════════════════════════════════════════════════════
// SLIDE 11 — Starlink + Kinective Media
// ════════════════════════════════════════════════════════════
{
  const sld = contentSlide(prs);
  addNavyHeader(sld, 'Connectivity as a Business: Starlink + Kinective Media', 'Monetizing the world\'s largest captive audience');

  // Left: Starlink rollout
  sld.addShape(prs.ShapeType.rect, { x: 0.25, y: 1.18, w: 4.5, h: 0.3, fill: { color: C.navy } });
  sld.addText('Starlink Rollout Timeline', { x: 0.25, y: 1.18, w: 4.5, h: 0.3, fontSize: 10, bold: true, color: C.white, align: 'center', valign: 'middle' });

  const events = [
    ['May 2025', 'First E175 regional enters service'],
    ['Oct 15, 2025', 'First mainline commercial flight (EWR)'],
    ['Feb 2026', '300+ mainline aircraft equipped'],
    ['End 2026', '800+ total aircraft'],
    ['End 2027', 'Full fleet (~50+ installs/month)'],
    ['All time', 'Free for all MileagePlus members'],
  ];
  events.forEach(([dt, txt], i) => {
    const y = 1.52 + i * 0.48;
    const bg = i % 2 === 0 ? C.lightBlue : 'EDF4FB';
    sld.addShape(prs.ShapeType.rect, { x: 0.25, y, w: 4.5, h: 0.44, fill: { color: bg } });
    sld.addText(dt, { x: 0.35, y: y + 0.04, w: 1.2, h: 0.36, fontSize: 9, bold: true, color: C.navy, valign: 'middle' });
    sld.addText(txt, { x: 1.6, y: y + 0.04, w: 3.05, h: 0.36, fontSize: 9, color: C.textDark, valign: 'middle' });
  });

  // Right: Kinective Media
  sld.addShape(prs.ShapeType.rect, { x: 5.0, y: 1.18, w: 4.75, h: 0.3, fill: { color: C.navyLight } });
  sld.addText('Kinective Media — The Real Story', { x: 5.0, y: 1.18, w: 4.75, h: 0.3, fontSize: 10, bold: true, color: C.white, align: 'center', valign: 'middle' });

  sld.addText('Three assets no competitor possesses simultaneously:', { x: 5.1, y: 1.53, w: 4.55, h: 0.25, fontSize: 9, bold: true, color: C.navy });

  const kinAssets = [
    ['227,000+', 'seatback screens — captive audience'],
    ['108M', 'unique annual flyers with known demographics'],
    ['130M', 'MileagePlus members with first-party transaction data'],
  ];
  kinAssets.forEach(([num, lbl], i) => {
    sld.addShape(prs.ShapeType.rect, { x: 5.0, y: 1.82 + i * 0.52, w: 4.75, h: 0.48, fill: { color: i % 2 === 0 ? C.lightBlue : 'EDF4FB' } });
    sld.addText(num, { x: 5.1, y: 1.87 + i * 0.52, w: 1.2, h: 0.36, fontSize: 14, bold: true, color: C.navy, align: 'center' });
    sld.addText(lbl, { x: 6.35, y: 1.87 + i * 0.52, w: 3.3, h: 0.36, fontSize: 9, color: C.textDark, valign: 'middle' });
  });

  // CFO quote
  sld.addShape(prs.ShapeType.rect, { x: 5.0, y: 3.4, w: 4.75, h: 1.1, fill: { color: C.navy }, shadow: makeShadow() });
  sld.addText('CFO on Q4 earnings call:', { x: 5.12, y: 3.44, w: 4.52, h: 0.22, fontSize: 11, bold: true, color: C.gold });
  sld.addText('"Kinective will really start to accelerate in \'26 and beyond." Media margins dramatically higher than seat revenue. High-margin ad revenue changes the EBITDA profile in a way seat revenue cannot.', {
    x: 5.12, y: 3.68, w: 4.52, h: 0.78, fontSize: 11, color: C.white, wrap: true, valign: 'top', italic: true,
  });

  // Bottom callout
  sld.addShape(prs.ShapeType.rect, { x: 0.25, y: 4.58, w: 9.5, h: 0.87, fill: { color: C.gold } });
  sld.addText('Competitor Comparison: Delta has SkyMiles advertising. No competitor has UAL\'s level of integration between connectivity + behavioral data + programmatic capability. Market is currently valuing Kinective at approximately zero — a hidden asset not reflected in current multiples.', {
    x: 0.4, y: 4.62, w: 9.2, h: 0.8, fontSize: 9, color: C.darkSlate, wrap: true, valign: 'middle',
  });
}

// ════════════════════════════════════════════════════════════
// SLIDE 12 — Operational Performance
// ════════════════════════════════════════════════════════════
{
  const sld = contentSlide(prs);
  addNavyHeader(sld, 'Operational Excellence: Record Performance in 2025', 'Reliability and customer metrics at all-time highs');

  // Big stat boxes
  const stats = [
    { val: '181.1M', lbl: 'Passengers\n(Company Record)', sub: '+4.3% YoY' },
    { val: '#2', lbl: 'D:14 On-Time\nIndustry Rank', sub: 'Best-ever system completion' },
    { val: '303', lbl: 'Avg Daily Widebody\nDepartures', sub: 'Company record' },
    { val: '1M+', lbl: 'ConnectionSaver\nRescues', sub: '+42% YoY' },
    { val: '85%', lbl: 'Digital Check-In\nRate', sub: 'Q1 record' },
    { val: 'Record', lbl: 'NPS Score\nQ4 2025', sub: 'Nov\'25 highest ever' },
  ];

  // 3×2 grid with improved spacing to fill slide
  const sw = 3.1, sh = 1.38;
  const rowGap = 0.12;
  const gridTop = 1.2;
  stats.forEach((s, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const x = 0.25 + col * (sw + 0.1);
    const y = gridTop + row * (sh + rowGap);
    sld.addShape(prs.ShapeType.rect, { x, y, w: sw, h: sh, fill: { color: i % 2 === 0 ? C.navy : C.navyLight }, shadow: makeShadow() });
    sld.addText(s.val, { x: x + 0.1, y: y + 0.1, w: sw - 0.2, h: 0.5, fontSize: 26, bold: true, color: C.gold, align: 'center' });
    sld.addText(s.lbl, { x: x + 0.1, y: y + 0.62, w: sw - 0.2, h: 0.46, fontSize: 9, color: C.white, align: 'center', wrap: true });
    sld.addText(s.sub, { x: x + 0.1, y: y + 1.12, w: sw - 0.2, h: 0.22, fontSize: 11, color: C.goldLight, align: 'center' });
  });

  // Early 2026 demand signals — positioned after grid
  const demandY = gridTop + 2 * (sh + rowGap) + 0.1;
  sld.addShape(prs.ShapeType.rect, { x: 0.25, y: demandY, w: 9.5, h: 5.47 - demandY, fill: { color: C.gold }, shadow: makeShadow() });
  sld.addText('Early 2026 Demand Signals — Real-Time Leading Indicators', { x: 0.4, y: demandY + 0.06, w: 9.2, h: 0.28, fontSize: 11, bold: true, color: C.darkSlate });
  sld.addText(
    'Week of Jan 4, 2026: Highest flown revenue week in UAL history.  |  Week of Jan 11, 2026: Highest ticketing week AND highest business sales week ever recorded simultaneously.\n\nBusiness travel leads leisure in recovery cycles — entering 2026 with record business bookings is the strongest real-time indicator that $12–$14 EPS guidance is achievable.',
    { x: 0.4, y: demandY + 0.37, w: 9.2, h: 5.47 - demandY - 0.42, fontSize: 9, color: C.darkSlate, wrap: true, valign: 'top' }
  );
}

// ════════════════════════════════════════════════════════════
// SLIDE 13 — Balance Sheet & Deleveraging
// ════════════════════════════════════════════════════════════
{
  const sld = contentSlide(prs);
  addNavyHeader(sld, 'Balance Sheet Normalization: Path to Investment Grade', 'Leverage declining; MileagePlus fully unencumbered');

  const rows = [
    ['Period', 'Total Liquidity', 'Total Debt + Leases', 'Net Leverage', 'S&P Rating'],
    ['Dec 2023', '$16.1B', '$29.3B', '2.9\u00d7', 'BB'],
    ['Dec 2024', '$17.4B', '$28.7B', '2.4\u00d7', 'BB+'],
    ['Mar 2025', '$18.3B', '$27.7B', '2.0\u00d7', 'BB+'],
    ['Jun 2025', '$18.6B', '$27.1B', '2.0\u00d7', 'BB+'],
    ['Sep 2025', '$16.3B', '$25.4B', '2.1\u00d7', 'BB+'],
    ['Dec 2025', '$15.2B', '$25.0B', '2.2\u00d7', 'BB+'],
    ['Dec 2026E', '—', '—', '<2.0\u00d7 target', 'IG target'],
  ];
  styledTable(sld, rows, 0.25, 1.18, 6.2,
    [1.2, 1.5, 1.5, 1.0, 1.0], 0.38,
    C.navy, C.white, C.white, C.textDark, 9
  );

  // Key events
  sld.addShape(prs.ShapeType.rect, { x: 6.65, y: 1.18, w: 3.1, h: 1.82, fill: { color: C.lightBlue }, shadow: makeShadow() });
  sld.addText('Key Events', { x: 6.77, y: 1.24, w: 2.88, h: 0.28, fontSize: 9.5, bold: true, color: C.navy });
  const evts = [
    'Q3 2025: Prepaid final $1.5B of MileagePlus bonds — Full $6.8B COVID-era collateralized debt retired',
    'MileagePlus asset now fully unencumbered',
    'Interest expense: $1.629B (2024) → $1.373B (2025) = –15.7% — flows straight to EPS',
  ];
  evts.forEach((e, i) => {
    sld.addText('\u2022  ' + e, { x: 6.77, y: 1.55 + i * 0.43, w: 2.88, h: 0.4, fontSize: 9, color: C.textDark, wrap: true, valign: 'top' });
  });

  // IG impact callout
  sld.addShape(prs.ShapeType.rect, { x: 0.25, y: 4.15, w: 9.5, h: 1.3, fill: { color: C.gold }, shadow: makeShadow() });
  sld.addText('Three Consequences of Investment-Grade Upgrade', { x: 0.4, y: 4.19, w: 9.2, h: 0.3, fontSize: 11, bold: true, color: C.darkSlate });
  const igItems = [
    '1. Lower financing cost on $25B fleet debt → $125–190M annual interest savings',
    '2. IG mandate funds open to UAL equity — materially larger institutional buyer universe',
    '3. Multiple re-rating: COVID-risk narrative to investment-quality enterprise complete. Delta carries BBB-. UAL at 2.2\u00d7 leverage and falling — probability of upgrade is high. Question: own before or after the rating action.',
  ];
  igItems.forEach((it, i) => {
    sld.addText(it, { x: 0.4, y: 4.52 + i * 0.28, w: 9.2, h: 0.26, fontSize: 9, color: C.darkSlate, wrap: true });
  });
}

// ════════════════════════════════════════════════════════════
// SLIDE 14 — Capital Allocation
// ════════════════════════════════════════════════════════════
{
  const sld = contentSlide(prs);
  addNavyHeader(sld, 'Capital Allocation: CapEx Bulge → FCF Normalization', 'Fleet investment peaking; FCF trajectory improving post-2027');

  // Cash flow table
  const rows = [
    ['Metric', 'FY2023', 'FY2024', 'FY2025', 'FY2026E'],
    ['Operating Cash Flow ($B)', '$6.911', '$9.4', '$8.4', '—'],
    ['Net CapEx ($B)', '$7.948', '$5.6', '$5.874', '<$8.0'],
    ['Free Cash Flow ($B)', '($1.037)', '$3.4', '$2.7', '~$2.7'],
  ];
  styledTable(sld, rows, 0.25, 1.18, 5.5,
    [2.2, 0.82, 0.82, 0.82, 0.84], 0.42,
    C.navy, C.white, C.white, C.textDark, 9.5
  );

  // Priority stack
  sld.addShape(prs.ShapeType.rect, { x: 5.95, y: 1.18, w: 3.8, h: 0.3, fill: { color: C.navy } });
  sld.addText('Capital Priority Stack', { x: 5.95, y: 1.18, w: 3.8, h: 0.3, fontSize: 10, bold: true, color: C.white, align: 'center', valign: 'middle' });

  const pItems = [
    { n: '1', lbl: 'Fleet CapEx (Non-negotiable)', txt: '$5.9B in 2025, <$8B guided 2026; 120 aircraft deliveries', bg: C.navy },
    { n: '2', lbl: 'Debt Reduction', txt: '$3.7B total debt reduction FY2025; $3–4B/yr pace', bg: C.navyLight },
    { n: '3', lbl: 'Share Buybacks', txt: 'FY2024: $81M → FY2025: $640M → expanding. $1.5B authorization outstanding. ~325M shares 2026E.', bg: C.lightBlue },
    { n: '4', lbl: 'No Dividend (near-term)', txt: 'Free cash flow directed to fleet + delevering', bg: 'EDF4FB' },
  ];
  pItems.forEach((p, i) => {
    const y = 1.52 + i * 0.72;
    sld.addShape(prs.ShapeType.rect, { x: 5.95, y, w: 3.8, h: 0.68, fill: { color: p.bg } });
    sld.addShape(prs.ShapeType.rect, { x: 5.95, y, w: 0.35, h: 0.68, fill: { color: C.gold } });
    sld.addText(p.n, { x: 5.95, y, w: 0.35, h: 0.68, fontSize: 12, bold: true, color: C.darkSlate, align: 'center', valign: 'middle' });
    sld.addText(p.lbl, { x: 6.35, y: y + 0.05, w: 3.35, h: 0.25, fontSize: 9, bold: true, color: ['EDF4FB', C.lightBlue].includes(p.bg) ? C.navy : C.white });
    sld.addText(p.txt, { x: 6.35, y: y + 0.31, w: 3.35, h: 0.33, fontSize: 9, color: ['EDF4FB', C.lightBlue].includes(p.bg) ? C.textDark : C.white, wrap: true });
  });

  // Post-2027 callout
  sld.addShape(prs.ShapeType.rect, { x: 0.25, y: 4.22, w: 5.5, h: 1.2, fill: { color: C.gold }, shadow: makeShadow() });
  sld.addText('Post-2027 FCF Normalization Thesis', { x: 0.4, y: 4.27, w: 5.2, h: 0.3, fontSize: 10, bold: true, color: C.darkSlate });
  sld.addText('CapEx bulge persists through 2027. Post-2028 steady-state: $5–6B CapEx on $8–9B operating CF = $3–4B sustained FCF. At 320M shares = ~$10–12/share FCF. Airlines historically trade at 5–8\u00d7 FCF. At normalized FCF, current share price implies significant upside for investors who can look through the delivery cycle.', {
    x: 0.4, y: 4.59, w: 5.2, h: 0.8, fontSize: 9, color: C.darkSlate, wrap: true, valign: 'top',
  });
}

// ════════════════════════════════════════════════════════════
// SLIDE 15 — Sustainability
// ════════════════════════════════════════════════════════════
{
  const sld = contentSlide(prs);
  addNavyHeader(sld, 'Sustainability: Honest Assessment & Commitments', 'Science-based targets; no greenwashing');

  // Left: Commitments
  sld.addShape(prs.ShapeType.rect, { x: 0.25, y: 1.18, w: 4.5, h: 0.3, fill: { color: C.navy } });
  sld.addText('Commitments', { x: 0.25, y: 1.18, w: 4.5, h: 0.3, fontSize: 10, bold: true, color: C.white, align: 'center', valign: 'middle' });

  const cmts = [
    'Net zero GHG by 2050 — first global airline without traditional offsets',
    '50% emissions intensity reduction by 2035 vs. 2019 (SBTi-validated May 2023)',
    'CDP Climate Disclosure: A-',
  ];
  cmts.forEach((c, i) => {
    sld.addShape(prs.ShapeType.rect, { x: 0.25, y: 1.52 + i * 0.42, w: 4.5, h: 0.38, fill: { color: i % 2 === 0 ? C.lightBlue : 'EDF4FB' } });
    sld.addText('\u2022  ' + c, { x: 0.35, y: 1.55 + i * 0.42, w: 4.3, h: 0.32, fontSize: 9, color: C.textDark, wrap: true });
  });

  // Intensity trend
  sld.addShape(prs.ShapeType.rect, { x: 0.25, y: 2.82, w: 4.5, h: 0.28, fill: { color: C.navyLight } });
  sld.addText('Emissions Intensity (per M ASMs)', { x: 0.25, y: 2.82, w: 4.5, h: 0.28, fontSize: 9, bold: true, color: C.white, align: 'center', valign: 'middle' });

  // Y-axis unit label (rotated 90°)
  sld.addText('CO2e / M ASM', {
    x: 0.25, y: 3.0, w: 0.95, h: 0.22,
    fontSize: 9, color: C.textMed, bold: true, align: 'center',
    rotate: 270,
  });

  const intData = [['2021','187.5'], ['2022','176.2'], ['2023','169.0'], ['2024','167.3']];
  const barW = 4.5 / intData.length - 0.12;
  intData.forEach(([yr, val], i) => {
    const norm = parseFloat(val) / 190;
    const bh = 0.85 * norm;
    const bx = 0.25 + i * (barW + 0.12);
    const by = 3.1 + 0.85 * (1 - norm);
    sld.addShape(prs.ShapeType.rect, { x: bx, y: by, w: barW, h: bh, fill: { color: i === 3 ? C.navy : C.lightBlue } });
    // Data label above bar
    sld.addText(val, { x: bx, y: by - 0.22, w: barW, h: 0.2, fontSize: 9, bold: true, color: i === 3 ? C.navy : C.navyLight, align: 'center' });
    // Data label inside bar (near top)
    if (bh > 0.22) {
      sld.addText(val, { x: bx, y: by + 0.04, w: barW, h: 0.2, fontSize: 9, bold: true, color: i === 3 ? C.white : C.navy, align: 'center' });
    }
    sld.addText(yr, { x: bx, y: 3.97, w: barW, h: 0.18, fontSize: 9, color: C.textMed, align: 'center' });
  });
  // Legend for the bar chart colors
  sld.addShape(prs.ShapeType.rect, { x: 0.25, y: 4.15, w: 0.18, h: 0.12, fill: { color: C.lightBlue } });
  sld.addText('2021–2023', { x: 0.46, y: 4.14, w: 1.1, h: 0.14, fontSize: 9, color: C.textLight });
  sld.addShape(prs.ShapeType.rect, { x: 1.65, y: 4.15, w: 0.18, h: 0.12, fill: { color: C.navy } });
  sld.addText('2024 (latest)', { x: 1.86, y: 4.14, w: 1.1, h: 0.14, fontSize: 9, color: C.textLight });

  // SAF table
  sld.addShape(prs.ShapeType.rect, { x: 0.25, y: 4.32, w: 4.5, h: 0.25, fill: { color: C.navy } });
  sld.addText('SAF Progress', { x: 0.25, y: 4.32, w: 4.5, h: 0.25, fontSize: 9, bold: true, color: C.white, align: 'center', valign: 'middle' });
  const safRows = [
    ['Year', 'SAF Gallons', 'CO2e Avoided'],
    ['2021', '0.6M', '5,953 MT'],
    ['2022', '2.9M', '29,362 MT'],
    ['2023', '7.3M', '68,370 MT'],
    ['2024', '13.6M', '126,174 MT'],
  ];
  styledTable(sld, safRows, 0.25, 4.57, 4.5, [1.0, 1.75, 1.75], 0.18, C.navyLight, C.white, C.white, C.textDark, 8);

  // Right: Honest Assessment
  sld.addShape(prs.ShapeType.rect, { x: 5.0, y: 1.18, w: 4.75, h: 2.1, fill: { color: C.gold }, shadow: makeShadow() });
  sld.addText('Honest Assessment for Institutional Investors', { x: 5.12, y: 1.23, w: 4.52, h: 0.3, fontSize: 10, bold: true, color: C.darkSlate });
  sld.addText(
    'Absolute emissions rising with capacity growth (Scope 1: +5.3% in 2024). 13.6M SAF gallons vs. 4.2B+ total consumed = <0.35%. The 2050 net zero target depends on technologies not yet at commercial scale (green hydrogen, next-gen SAF, carbon removal). The 2035 intensity target is achievable with newer aircraft. Do not model material near-term emissions reductions into the financial case.',
    { x: 5.12, y: 1.57, w: 4.52, h: 1.65, fontSize: 9, color: C.darkSlate, wrap: true, valign: 'top' }
  );

  // Ventures
  sld.addShape(prs.ShapeType.rect, { x: 5.0, y: 3.42, w: 4.75, h: 2.0, fill: { color: C.lightBlue }, shadow: makeShadow() });
  sld.addText('UAL SAF Ventures Portfolio ($200M+)', { x: 5.12, y: 3.47, w: 4.52, h: 0.28, fontSize: 9.5, bold: true, color: C.navy });
  const ventures = [
    'Twelve (CO2-to-SAF) — direct air capture pathway',
    'ZeroAvia (hydrogen-electric) — regional aviation',
    'Dimensional Energy, Cemvita, Svante (carbon capture)',
    'Eco-Skies Alliance: 50+ corporate partners co-investing in SAF supply',
  ];
  ventures.forEach((v, i) => {
    sld.addText('\u2022  ' + v, { x: 5.12, y: 3.78 + i * 0.4, w: 4.52, h: 0.37, fontSize: 9, color: C.textDark, wrap: true });
  });
}

// ════════════════════════════════════════════════════════════
// SLIDE 16 — 2026 Guidance & Risk/Reward
// ════════════════════════════════════════════════════════════
{
  const sld = contentSlide(prs);
  addNavyHeader(sld, '2026 Guidance & Risk/Reward Framework', 'EPS inflection supported by multiple convergent tailwinds');

  // Guidance table
  const gRows = [
    ['Metric', 'Q1 2026E', 'FY2026E', 'vs. FY2025'],
    ['Adj. EPS', '$1.00–$1.50', '$12.00–$14.00', '+13% to +32%'],
    ['EPS midpoint', '$1.25', '$13.00', '+22%'],
    ['Adj. Pre-Tax Margin', '—', '~10%+', '+220+ bps'],
    ['Free Cash Flow', '—', '~$2.7B', 'flat'],
    ['Net CapEx', '—', '<$8.0B', '—'],
    ['Net Leverage (target)', '—', '<2.0\u00d7', '—'],
  ];
  styledTable(sld, gRows, 0.25, 1.18, 5.5,
    [2.2, 1.1, 1.1, 1.1], 0.38,
    C.navy, C.white, C.white, C.textDark, 9
  );

  // Bull case
  sld.addShape(prs.ShapeType.rect, { x: 0.25, y: 4.05, w: 5.5, h: 1.35, fill: { color: C.gold }, shadow: makeShadow() });
  sld.addText('Bull Case', { x: 0.4, y: 4.1, w: 5.2, h: 0.28, fontSize: 10.5, bold: true, color: C.darkSlate });
  sld.addText('If TRASM inflects positive in Q1–Q2 2026 as industry capacity tightens → tracking toward $14+ EPS → at current multiple = significant equity upside. 2019 EPS peak: $12.05. $14 would be first material all-time high.', {
    x: 0.4, y: 4.4, w: 5.2, h: 0.95, fontSize: 9, color: C.darkSlate, wrap: true, valign: 'top',
  });

  // Risk matrix
  sld.addShape(prs.ShapeType.rect, { x: 6.0, y: 1.18, w: 3.75, h: 0.3, fill: { color: C.navy } });
  sld.addText('Risk Matrix', { x: 6.0, y: 1.18, w: 3.75, h: 0.3, fontSize: 10, bold: true, color: C.white, align: 'center', valign: 'middle' });

  const risks = [
    { risk: 'GDP Recession', mag: '–1.5–2.0% rev/–1% GDP', mit: 'Premium/intl mix; loyalty contract-based' },
    { risk: 'Fuel Spike', mag: '$0.10/gal = ~$466M; unhedged', mit: 'New aircraft –20% fuel/seat' },
    { risk: 'Boeing MAX Delays', mag: '~$30–40M/month deferred', mit: 'A321neo/XLR Airbus hedge' },
    { risk: 'Domestic PRASM', mag: 'Margin expansion stalls', mit: 'Mid-year capacity cuts demonstrated' },
    { risk: 'MAX 10 Cert Slip', mag: '200 aircraft waiting', mit: 'Can defer CapEx if needed' },
    { risk: 'Chase Re-contract', mag: '2029 — terms uncertain', mit: 'Delta/AmEx parity gives leverage' },
  ];
  risks.forEach((r, i) => {
    const y = 1.52 + i * 0.58;
    const bg = i % 2 === 0 ? 'FFF8EC' : C.lightBlue;
    sld.addShape(prs.ShapeType.rect, { x: 6.0, y, w: 3.75, h: 0.54, fill: { color: bg } });
    sld.addText(r.risk, { x: 6.1, y: y + 0.03, w: 1.4, h: 0.22, fontSize: 9, bold: true, color: C.navy });
    sld.addText(r.mag, { x: 6.1, y: y + 0.26, w: 1.4, h: 0.24, fontSize: 9, color: C.textMed, wrap: true });
    sld.addText(r.mit, { x: 7.55, y: y + 0.06, w: 2.12, h: 0.42, fontSize: 9, color: C.textDark, wrap: true, valign: 'middle' });
  });
}

// ════════════════════════════════════════════════════════════
// SLIDE 17 — Appendix: Financial Statements
// ════════════════════════════════════════════════════════════
{
  const sld = prs.addSlide();
  sld.background = { color: 'F4F7FA' };

  sld.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: 10, h: 0.85, fill: { color: C.navy } });
  sld.addText('Appendix: Multi-Year Income Statement', { x: 0.3, y: 0.08, w: 9.4, h: 0.55, fontSize: 20, bold: true, color: C.white, fontFace: 'Calibri', valign: 'middle' });
  sld.addText('FY2023–FY2025 | Figures in billions unless noted', { x: 0.3, y: 0.6, w: 9.4, h: 0.22, fontSize: 9, color: C.goldLight, valign: 'top' });
  sld.addShape(prs.ShapeType.rect, { x: 0, y: 0.85, w: 10, h: 0.04, fill: { color: C.gold } });

  const rows = [
    ['Line Item', 'FY2023', 'FY2024', 'FY2025'],
    ['Passenger Revenue', '—', '$51.829B', '$53.436B'],
    ['Cargo Revenue', '—', '$1.743B', '$1.779B'],
    ['Other Revenue', '—', '$3.491B', '$3.853B'],
    ['Total Revenue', '$53.717B', '$57.063B', '$59.068B'],
    ['Salaries', '$14.787B', '$16.678B', '$17.647B'],
    ['Fuel', '$12.651B', '$11.756B', '$11.396B'],
    ['Landing Fees', '—', '$3.437B', '$3.849B'],
    ['Maintenance', '$2.736B', '$3.063B', '$3.294B'],
    ['D&A', '$2.671B', '$2.928B', '$2.939B'],
    ['Regional Capacity', '—', '$2.516B', '$2.693B'],
    ['Distribution', '—', '$2.231B', '$2.109B'],
    ['Other OpEx', '—', '$9.053B', '$9.919B'],
    ['Total OpEx', '$49.506B', '$51.967B', '$54.356B'],
    ['Operating Income', '$4.211B', '$5.096B', '$4.713B'],
    ['Net Income', '$2.618B', '$3.149B', '$3.400B'],
    ['Adj. Diluted EPS', '$10.05', '$10.61', '$10.62'],
    ['Adj. Pre-Tax Margin', '8.0%', '8.1%', '7.8%'],
    ['Adj. EBITDA', '$7.938B', '$8.211B', '$8.076B'],
    ['Adj. EBITDA Margin', '14.8%', '14.4%', '13.7%'],
    ['Operating CF', '$6.911B', '$9.4B', '$8.4B'],
    ['Net CapEx', '$7.948B', '$5.6B', '$5.874B'],
    ['Free Cash Flow', '($1.037B)', '$3.4B', '$2.7B'],
  ];

  styledTable(sld, rows, 0.25, 0.97, 9.5,
    [3.5, 2.0, 2.0, 2.0], 0.22,
    C.navy, C.white, C.white, C.textDark, 8
  );
}

// ════════════════════════════════════════════════════════════
// SLIDE 18 — Appendix: Unit Economics & Fleet
// ════════════════════════════════════════════════════════════
{
  const sld = prs.addSlide();
  sld.background = { color: 'F4F7FA' };

  sld.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: 10, h: 0.85, fill: { color: C.navy } });
  sld.addText('Appendix: Unit Economics & Fleet Reference', { x: 0.3, y: 0.08, w: 9.4, h: 0.55, fontSize: 20, bold: true, color: C.white, fontFace: 'Calibri', valign: 'middle' });
  sld.addText('Annual + Quarterly Unit Economics | Fleet & Workforce | CASM-ex Reconciliation', { x: 0.3, y: 0.6, w: 9.4, h: 0.22, fontSize: 9, color: C.goldLight });
  sld.addShape(prs.ShapeType.rect, { x: 0, y: 0.85, w: 10, h: 0.04, fill: { color: C.gold } });

  // Annual unit economics
  sld.addShape(prs.ShapeType.rect, { x: 0.25, y: 0.95, w: 3.0, h: 0.22, fill: { color: C.navyLight } });
  sld.addText('Annual', { x: 0.25, y: 0.95, w: 3.0, h: 0.22, fontSize: 9, bold: true, color: C.white, align: 'center', valign: 'middle' });

  const annRows = [
    ['Year', 'TRASM', 'PRASM', 'Yield', 'CASM', 'CASM-ex', 'Fuel/Gal', 'Load Factor'],
    ['FY2023', '18.44\u00a2', '—', '20.07\u00a2', '16.99\u00a2', '12.03\u00a2', '$3.01', '83.9%'],
    ['FY2024', '18.34\u00a2', '16.66\u00a2', '20.05\u00a2', '16.70\u00a2', '12.58\u00a2', '$2.65', '83.1%'],
    ['FY2025', '17.88\u00a2', '16.18\u00a2', '19.67\u00a2', '16.46\u00a2', '12.64\u00a2', '$2.44', '82.2%'],
  ];
  styledTable(sld, annRows, 0.25, 1.17, 9.5,
    [1.0, 1.0, 1.0, 1.0, 1.0, 1.1, 1.1, 2.3], 0.28,
    C.navy, C.white, C.white, C.textDark, 8
  );

  // Quarterly unit economics
  sld.addShape(prs.ShapeType.rect, { x: 0.25, y: 2.4, w: 3.0, h: 0.22, fill: { color: C.navyLight } });
  sld.addText('2025 Quarterly', { x: 0.25, y: 2.4, w: 3.0, h: 0.22, fontSize: 9, bold: true, color: C.white, align: 'center', valign: 'middle' });

  const qRows = [
    ['Quarter', 'TRASM', 'PRASM', 'Yield', 'CASM', 'CASM-ex', 'Fuel/Gal', 'Load Factor'],
    ['Q1', '17.58\u00a2', '15.78\u00a2', '19.93\u00a2', '16.77\u00a2', '13.17\u00a2', '$2.53', '79.2%'],
    ['Q2', '18.06\u00a2', '16.40\u00a2', '19.74\u00a2', '16.49\u00a2', '12.36\u00a2', '$2.34', '83.1%'],
    ['Q3', '17.42\u00a2', '15.80\u00a2', '18.73\u00a2', '15.82\u00a2', '12.15\u00a2', '$2.43', '84.4%'],
    ['Q4', '18.47\u00a2', '16.71\u00a2', '20.41\u00a2', '16.81\u00a2', '12.94\u00a2', '$2.49', '81.9%'],
  ];
  styledTable(sld, qRows, 0.25, 2.62, 9.5,
    [1.0, 1.0, 1.0, 1.0, 1.0, 1.1, 1.1, 2.3], 0.24,
    C.navyLight, C.white, C.white, C.textDark, 8
  );

  // Fleet & workforce table — label spans full table width (4.5")
  // qRows: 5 rows × 0.24 = 1.20; table ends at 2.62 + 1.20 = 3.82; add 0.12" gap → fleet header at 3.94
  sld.addShape(prs.ShapeType.rect, { x: 0.25, y: 3.94, w: 4.5, h: 0.22, fill: { color: C.navy } });
  sld.addText('Fleet & Workforce', { x: 0.25, y: 3.94, w: 4.5, h: 0.22, fontSize: 9, bold: true, color: C.white, align: 'center', valign: 'middle' });

  const fleetRows = [
    ['Year', 'Fleet', 'Employees', 'Salaries ($B)'],
    ['2021', '—', '~75,000', '—'],
    ['2022', '~1,300', '~92,000', '—'],
    ['2023', '1,358', '103,300', '$14.787B'],
    ['2024', '1,406', '107,300', '$16.678B'],
    ['2025', '1,490', '113,200', '$17.647B'],
  ];
  styledTable(sld, fleetRows, 0.25, 4.16, 4.5,
    [0.8, 0.8, 1.3, 1.6], 0.20,
    C.navy, C.white, C.white, C.textDark, 8
  );

  // CASM-ex reconciliation
  sld.addShape(prs.ShapeType.rect, { x: 5.0, y: 3.94, w: 4.75, h: 1.58, fill: { color: C.lightBlue } });
  sld.addText('CASM-ex Reconciliation (Q1 2025 example)', { x: 5.12, y: 3.99, w: 4.52, h: 0.25, fontSize: 9, bold: true, color: C.navy });
  const recon = [
    'CASM: 16.77\u00a2',
    'Less fuel: (3.59)\u00a2',
    'Less profit sharing: (0.06)\u00a2',
    'Less third-party: (0.09)\u00a2',
    'Add back special charges: +0.14\u00a2',
    'CASM-ex: 13.17\u00a2',
  ];
  recon.forEach((r, i) => {
    const bold = i === 0 || i === recon.length - 1;
    sld.addText(r, { x: 5.12, y: 4.26 + i * 0.20, w: 4.52, h: 0.20, fontSize: 9, bold, color: bold ? C.navy : C.textDark });
  });
  sld.addText('Fuel consumption: 4.663B gallons in 2025', { x: 5.12, y: 5.42, w: 4.52, h: 0.18, fontSize: 9, color: C.textLight, italic: true });
}

// ─── Save ────────────────────────────────────────────────────────────────────
prs.writeFile({ fileName: 'outputs/corporate-qa.pptx' })
  .then(() => console.log('Saved: outputs/corporate-qa.pptx'))
  .catch(err => { console.error(err); process.exit(1); });
