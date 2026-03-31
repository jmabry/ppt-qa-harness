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
  altRow:    'EDF4FB',
  pageBg:    'F4F7FA',
};

// ─── Helpers ────────────────────────────────────────────────────────────────
// ALWAYS use factory – never reuse shadow objects
function makeShadow() {
  return { type: 'outer', color: '000000', blur: 4, offset: 2, angle: 45, opacity: 0.25 };
}

function contentSlide(prs) {
  const sld = prs.addSlide();
  sld.background = { color: C.white };
  return sld;
}

// Canvas: LAYOUT_16x9 = 10" wide × 5.625" tall
function addNavyHeader(sld, title, subtitle) {
  const headerH = subtitle ? 1.05 : 0.82;
  sld.addShape(prs.ShapeType.rect, {
    x: 0, y: 0, w: 10, h: headerH,
    fill: { color: C.navy },
  });
  sld.addText(title, {
    x: 0.3, y: 0.08, w: 9.4, h: 0.6,
    fontSize: 24, bold: true, color: C.white, fontFace: 'Calibri',
    valign: 'middle', margin: 0,
  });
  if (subtitle) {
    sld.addText(subtitle, {
      x: 0.3, y: 0.68, w: 9.4, h: 0.3,
      fontSize: 10, color: C.goldLight, fontFace: 'Calibri Light', valign: 'top', margin: 0,
    });
  }
  // Gold divider strip
  sld.addShape(prs.ShapeType.rect, {
    x: 0, y: headerH, w: 10, h: 0.045,
    fill: { color: C.gold },
  });
  return headerH + 0.045;
}

// Helper to add a callout box with label+body
function addCallout(sld, x, y, w, h, label, bodyLines, bgColor, textColor) {
  bgColor   = bgColor   || C.lightBlue;
  textColor = textColor || C.textDark;
  sld.addShape(prs.ShapeType.rect, {
    x, y, w, h,
    fill: { color: bgColor },
    line: { color: bgColor },
    shadow: makeShadow(),
  });
  const arr = [];
  if (label) {
    arr.push({ text: label, options: { bold: true, fontSize: 10, color: textColor, breakLine: true } });
  }
  bodyLines.forEach((ln, i) => {
    arr.push({ text: ln, options: { fontSize: 9, color: textColor, breakLine: i < bodyLines.length - 1 } });
  });
  sld.addText(arr, { x: x + 0.1, y: y + 0.08, w: w - 0.2, h: h - 0.16, valign: 'top', wrap: true });
}

// Simple styled table
function styledTable(sld, rows, x, y, w, colW, rowH, headColor, headTextColor, bodyBg, bodyText, fontSize) {
  fontSize      = fontSize      || 8.5;
  headColor     = headColor     || C.navy;
  headTextColor = headTextColor || C.white;
  bodyBg        = bodyBg        || C.white;
  bodyText      = bodyText      || C.textDark;

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
          fill: isHead ? headColor : (isAlt ? C.altRow : bodyBg),
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
prs.layout = 'LAYOUT_16x9'; // 10" × 5.625"

// ════════════════════════════════════════════════════════════
// SLIDE 1 — Title
// ════════════════════════════════════════════════════════════
{
  const sld = prs.addSlide();
  sld.background = { color: C.navy };

  // Top strip
  sld.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: 10, h: 0.35, fill: { color: C.darkSlate } });
  sld.addText('INSTITUTIONAL INVESTOR PRESENTATION', {
    x: 0.4, y: 0.05, w: 9.2, h: 0.27,
    fontSize: 8.5, color: C.goldLight, fontFace: 'Calibri', bold: true, charSpacing: 2, margin: 0,
  });

  // UAL logo circle
  sld.addShape(prs.ShapeType.ellipse, {
    x: 0.45, y: 0.75, w: 1.05, h: 1.05,
    fill: { color: C.gold }, shadow: makeShadow(),
  });
  sld.addText('UAL', {
    x: 0.45, y: 0.88, w: 1.05, h: 0.65,
    fontSize: 21, bold: true, color: C.navy, fontFace: 'Calibri', align: 'center', margin: 0,
  });

  // Main title
  sld.addText('United Airlines Holdings', {
    x: 0.4, y: 1.65, w: 9.2, h: 0.95,
    fontSize: 44, bold: true, color: C.white, fontFace: 'Calibri', margin: 0,
  });

  // Subtitle
  sld.addText('FY2025 Results & 2026 Outlook', {
    x: 0.4, y: 2.62, w: 9.2, h: 0.55,
    fontSize: 24, color: C.goldLight, fontFace: 'Calibri Light', margin: 0,
  });

  // Gold accent bar
  sld.addShape(prs.ShapeType.rect, { x: 0, y: 3.28, w: 10, h: 0.07, fill: { color: C.gold } });

  // Details row
  sld.addShape(prs.ShapeType.rect, { x: 0, y: 3.35, w: 10, h: 0.72, fill: { color: C.darkSlate } });
  sld.addText('UAL  |  NASDAQ  |  January 20, 2026', {
    x: 0.4, y: 3.42, w: 9.2, h: 0.48,
    fontSize: 12, color: C.white, valign: 'middle', margin: 0,
  });

  // Key stats strip — 5 columns, bottom band within 5.625" height
  const stats = [
    ['$59.1B', 'FY2025 Revenue'],
    ['$10.62', 'Adj. EPS'],
    ['181.1M', 'Passengers'],
    ['1,490', 'Fleet Size'],
    ['$12–$14', '2026E EPS'],
  ];
  const sw = 10 / stats.length; // 2.0" each
  stats.forEach(([val, lbl], i) => {
    sld.addShape(prs.ShapeType.rect, {
      x: i * sw, y: 4.18, w: sw, h: 1.445,
      fill: { color: i % 2 === 0 ? C.navyLight : C.navy },
    });
    sld.addText(val, {
      x: i * sw + 0.05, y: 4.24, w: sw - 0.1, h: 0.54,
      fontSize: 18, bold: true, color: C.gold, fontFace: 'Calibri', align: 'center', margin: 0,
    });
    sld.addText(lbl, {
      x: i * sw + 0.05, y: 4.8, w: sw - 0.1, h: 0.3,
      fontSize: 8, color: C.lightBlue, fontFace: 'Calibri Light', align: 'center', margin: 0,
    });
  });
}

// ════════════════════════════════════════════════════════════
// SLIDE 2 — Investment Thesis
// ════════════════════════════════════════════════════════════
{
  const sld = contentSlide(prs);
  const contentY = addNavyHeader(sld, 'Investment Thesis', 'Four pillars supporting UAL re-rating');

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
        "Net leverage 2.2× (Dec'25) → <2.0× (Dec'26)",
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

  const boxW = 4.72;
  const boxH = 1.88;
  const gap = 0.28;
  const positions = [
    { x: 0.14, y: contentY + 0.12 },
    { x: 0.14 + boxW + gap, y: contentY + 0.12 },
    { x: 0.14, y: contentY + 0.12 + boxH + 0.14 },
    { x: 0.14 + boxW + gap, y: contentY + 0.12 + boxH + 0.14 },
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
      x, y, w: boxW, h: 0.33,
      fill: { color: C.gold },
    });
    sld.addText(p.title, {
      x: x + 0.12, y: y + 0.03, w: boxW - 0.24, h: 0.28,
      fontSize: 9.5, bold: true, color: C.darkSlate, fontFace: 'Calibri', margin: 0,
    });
    p.lines.forEach((ln, j) => {
      sld.addText([{ text: ln, options: { bullet: true } }], {
        x: x + 0.12, y: y + 0.38 + j * 0.36, w: boxW - 0.24, h: 0.34,
        fontSize: 8.5, color: C.white, fontFace: 'Calibri Light', margin: 0,
      });
    });
  });

  // Bottom tagline
  const tagY = positions[2].y + boxH + 0.1;
  sld.addShape(prs.ShapeType.rect, { x: 0, y: tagY, w: 10, h: 5.625 - tagY, fill: { color: C.lightBlue } });
  sld.addText(
    'UAL enters 2026 with labor contracts settled, the largest delivery pipeline in its history, record booking trends, and a balance sheet on track for investment-grade — the convergence of tailwinds that have been building since 2022.',
    { x: 0.3, y: tagY + 0.07, w: 9.4, h: 5.625 - tagY - 0.14, fontSize: 8.5, color: C.textDark, italic: true, valign: 'middle', wrap: true, margin: 0 }
  );
}

// ════════════════════════════════════════════════════════════
// SLIDE 3 — Multi-Year Financial Performance
// ════════════════════════════════════════════════════════════
{
  const sld = contentSlide(prs);
  const contentY = addNavyHeader(sld, 'Recovery Arc: 2021–2026E', 'Six-year revenue and earnings trajectory');

  const rows = [
    ['Metric', 'FY2021', 'FY2022', 'FY2023', 'FY2024', 'FY2025', 'FY2026E'],
    ['Total Revenue ($B)', '$24.6', '$45.0', '$53.7', '$57.1', '$59.1', '—'],
    ['Adj. EPS', 'Neg.', '$10.61', '$10.05', '$10.61', '$10.62', '$12–$14'],
    ['Adj. Pre-Tax Margin', 'Neg.', '~6%', '8.0%', '8.1%', '7.8%', '~10%+'],
    ['ASMs (B)', '178.7', '247.9', '291.3', '311.2', '330.3', '~349'],
    ['Fleet', '—', '~1,300', '1,358', '1,406', '1,490', '~1,610E'],
  ];

  styledTable(sld, rows, 0.25, contentY + 0.1, 9.5,
    [2.1, 1.28, 1.28, 1.28, 1.28, 1.28, 1.0],
    0.38, C.navy, C.white, C.white, C.textDark, 9
  );

  // Revenue bar chart
  const yrs  = ['2021', '2022', '2023', '2024', '2025'];
  const revs = [24.6, 45.0, 53.7, 57.1, 59.1];
  const maxRev = 62;
  const barAreaX = 0.25, barAreaY = 3.42, barAreaW = 9.5, barAreaH = 0.8;
  const bw = barAreaW / yrs.length - 0.12;

  sld.addText('Revenue ($B) — scaled to $62B', {
    x: 0.25, y: 3.28, w: 6, h: 0.16, fontSize: 7.5, color: C.textLight, italic: true, margin: 0,
  });

  yrs.forEach((yr, i) => {
    const frac = revs[i] / maxRev;
    const bh = barAreaH * frac;
    const bx = barAreaX + i * (bw + 0.12);
    const by = barAreaY + barAreaH - bh;
    sld.addShape(prs.ShapeType.rect, {
      x: bx, y: by, w: bw, h: bh,
      fill: { color: i === 4 ? C.gold : C.navyLight },
    });
    sld.addText('$' + revs[i] + 'B', {
      x: bx, y: by + 0.04, w: bw, h: 0.2,
      fontSize: 7.5, color: C.white, bold: true, align: 'center', margin: 0,
    });
    sld.addText(yr, {
      x: bx, y: barAreaY + barAreaH + 0.03, w: bw, h: 0.18,
      fontSize: 7.5, color: C.textMed, align: 'center', margin: 0,
    });
  });

  // Callout
  const callY = barAreaY + barAreaH + 0.25;
  sld.addShape(prs.ShapeType.rect, {
    x: 0.25, y: callY, w: 9.5, h: 5.625 - callY - 0.05,
    fill: { color: C.gold }, shadow: makeShadow(),
  });
  sld.addText(
    'EPS Plateau 2022–2025: $10.61 → $10.05 → $10.61 → $10.62 reflects absorption of $10B+ in cumulative pilot/FA contract obligations — NOT operational underperformance. Revenue +31% since 2022. The plateau ends in 2026.',
    { x: 0.4, y: callY + 0.08, w: 9.1, h: 5.625 - callY - 0.18, fontSize: 9, color: C.darkSlate, bold: false, italic: true, valign: 'middle', wrap: true, margin: 0 }
  );
}

// ════════════════════════════════════════════════════════════
// SLIDE 4 — FY2025 Quarterly P&L
// ════════════════════════════════════════════════════════════
{
  const sld = contentSlide(prs);
  const contentY = addNavyHeader(sld, 'FY2025 Quarterly Performance', 'Revenue, earnings, and margin by quarter');

  const rows = [
    ['Quarter', 'Total Revenue', 'YoY', 'Net Income', 'Adj. EPS', 'Pre-Tax Margin', 'ASMs (B)'],
    ['Q1 2025', '$13.213B', '+5.4%', '$0.387B', '$0.91', '3.0%', '75.155'],
    ['Q2 2025', '$15.236B', '+1.7%', '$0.973B', '$3.87', '11.0%', '84.347'],
    ['Q3 2025', '$15.225B', '+2.6%', '$0.949B', '$2.78', '8.0%', '87.417'],
    ['Q4 2025', '$15.394B', '+4.8%', '$1.024B', '$3.10', '8.5%', '83.365'],
    ['FY 2025', '$59.068B', '+3.5%', '$3.400B', '$10.62', '7.8%', '330.284'],
  ];

  styledTable(sld, rows, 0.25, contentY + 0.1, 9.5,
    [1.3, 1.4, 0.78, 1.2, 1.06, 1.46, 2.3],
    0.4, C.navy, C.white, C.white, C.textDark, 9
  );

  // Three callout boxes
  const callouts = [
    {
      label: 'Q4 Context',
      body: "Revenue +4.8% but EPS –4.9% vs. Q4'24 ($3.10 vs. $3.26). Comps problem + temporary domestic fare softness, not structural.",
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

  const callBoxW = 3.12;
  const callBoxH = 5.545 - (contentY + 0.1 + 0.4 * 6 + 0.18);
  const callBoxY = contentY + 0.1 + 0.4 * 6 + 0.18;

  callouts.forEach((c, i) => {
    const x = 0.25 + i * (callBoxW + 0.06);
    sld.addShape(prs.ShapeType.rect, {
      x, y: callBoxY, w: callBoxW, h: callBoxH,
      fill: { color: c.bg }, shadow: makeShadow(),
    });
    // Contrasting title strip
    const stripColor = c.bg === C.gold ? C.darkSlate : C.gold;
    sld.addShape(prs.ShapeType.rect, {
      x, y: callBoxY, w: callBoxW, h: 0.3,
      fill: { color: stripColor },
    });
    sld.addText(c.label, {
      x: x + 0.1, y: callBoxY, w: callBoxW - 0.2, h: 0.3,
      fontSize: 9, bold: true,
      color: c.bg === C.gold ? C.goldLight : C.darkSlate,
      fontFace: 'Calibri', valign: 'middle', margin: 0,
    });
    sld.addText(c.body, {
      x: x + 0.1, y: callBoxY + 0.35, w: callBoxW - 0.2, h: callBoxH - 0.4,
      fontSize: 8.5, color: c.bg === C.navy ? C.white : C.darkSlate, wrap: true, valign: 'top', margin: 0,
    });
  });
}

// ════════════════════════════════════════════════════════════
// SLIDE 5 — Unit Economics Deep Dive
// ════════════════════════════════════════════════════════════
{
  const sld = contentSlide(prs);
  const contentY = addNavyHeader(sld, 'Unit Economics: The TRASM/CASM Story', 'Annual trends and 2025 quarterly breakdown');

  // Annual table header label
  sld.addShape(prs.ShapeType.rect, { x: 0.25, y: contentY + 0.1, w: 2.6, h: 0.22, fill: { color: C.navy } });
  sld.addText('Annual Trends', { x: 0.25, y: contentY + 0.1, w: 2.6, h: 0.22, fontSize: 8.5, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0 });

  const annualRows = [
    ['Year', 'TRASM', 'PRASM', 'CASM-ex', 'Fuel/Gal', 'Load Factor'],
    ['FY2023', '18.44¢', '—', '12.03¢', '$3.01', '83.9%'],
    ['FY2024', '18.34¢', '16.66¢', '12.58¢', '$2.65', '83.1%'],
    ['FY2025', '17.88¢', '16.18¢', '12.64¢', '$2.44', '82.2%'],
  ];
  styledTable(sld, annualRows, 0.25, contentY + 0.32, 5.8,
    [1.0, 0.94, 0.94, 0.98, 0.98, 0.96], 0.36,
    C.navyLight, C.white, C.white, C.textDark, 8.5
  );

  // Quarterly table header label
  sld.addShape(prs.ShapeType.rect, { x: 0.25, y: contentY + 1.85, w: 2.8, h: 0.22, fill: { color: C.navy } });
  sld.addText('2025 Quarterly', { x: 0.25, y: contentY + 1.85, w: 2.8, h: 0.22, fontSize: 8.5, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0 });

  const qRows = [
    ['Quarter', 'TRASM', 'PRASM', 'CASM-ex', 'Fuel/Gal', 'Load Factor'],
    ['Q1', '17.58¢', '15.78¢', '13.17¢', '$2.53', '79.2%'],
    ['Q2', '18.06¢', '16.40¢', '12.36¢', '$2.34', '83.1%'],
    ['Q3', '17.42¢', '15.80¢', '12.15¢', '$2.43', '84.4%'],
    ['Q4', '18.47¢', '16.71¢', '12.94¢', '$2.49', '81.9%'],
  ];
  styledTable(sld, qRows, 0.25, contentY + 2.07, 5.8,
    [1.0, 0.94, 0.94, 0.98, 0.98, 0.96], 0.35,
    C.navyLight, C.white, C.white, C.textDark, 8.5
  );

  // Right-side callouts
  const rx = 6.3, rw = 3.45;

  sld.addShape(prs.ShapeType.rect, { x: rx, y: contentY + 0.1, w: rw, h: 1.32, fill: { color: C.lightBlue }, shadow: makeShadow() });
  sld.addText('TRASM –2.5% in 2025: NOT a demand problem', { x: rx + 0.12, y: contentY + 0.15, w: rw - 0.24, h: 0.3, fontSize: 9, bold: true, color: C.navy, margin: 0 });
  sld.addText('Record 181M passengers. Industry capacity grew faster than unit revenue. Response: UAL pulled 4 ppts domestic capacity mid-year, retired 21 aircraft early.', {
    x: rx + 0.12, y: contentY + 0.47, w: rw - 0.24, h: 0.88, fontSize: 8.5, color: C.textDark, wrap: true, valign: 'top', margin: 0,
  });

  sld.addShape(prs.ShapeType.rect, { x: rx, y: contentY + 1.52, w: rw, h: 0.76, fill: { color: C.gold }, shadow: makeShadow() });
  sld.addText('Unhedged by Policy', { x: rx + 0.12, y: contentY + 1.57, w: rw - 0.24, h: 0.25, fontSize: 9, bold: true, color: C.darkSlate, margin: 0 });
  sld.addText('$0.10/gal = ~$466M annual P&L impact at 4.663B gallons consumed', {
    x: rx + 0.12, y: contentY + 1.84, w: rw - 0.24, h: 0.4, fontSize: 8.5, color: C.darkSlate, wrap: true, margin: 0,
  });

  const thesisY = contentY + 2.38;
  const thesisH = 5.545 - thesisY;
  sld.addShape(prs.ShapeType.rect, { x: rx, y: thesisY, w: rw, h: thesisH, fill: { color: C.navy }, shadow: makeShadow() });
  sld.addText('2026 Thesis', { x: rx + 0.12, y: thesisY + 0.08, w: rw - 0.24, h: 0.28, fontSize: 9, bold: true, color: C.gold, margin: 0 });
  sld.addText('TRASM inflects positive as industry capacity normalizes → CASM-ex stabilizes with aircraft deliveries → gap widens = margin expansion toward 10%+ pre-tax margin target', {
    x: rx + 0.12, y: thesisY + 0.38, w: rw - 0.24, h: thesisH - 0.48, fontSize: 8.5, color: C.white, wrap: true, valign: 'top', margin: 0,
  });
}

// ════════════════════════════════════════════════════════════
// SLIDE 6 — Operating Cost Bridge
// ════════════════════════════════════════════════════════════
{
  const sld = contentSlide(prs);
  const contentY = addNavyHeader(sld, 'Cost Structure: Composition Matters', 'FY2025 vs. FY2024 operating expense detail');

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

  styledTable(sld, rows, 0.25, contentY + 0.1, 9.5,
    [1.85, 1.14, 1.14, 0.98, 0.78, 3.61],
    0.35, C.navy, C.white, C.white, C.textDark, 8
  );

  // Three callout boxes at bottom
  const cbs = [
    { lbl: 'Labor', txt: "Up $2.9B in 2 yrs. 32.5% of OpEx. Gauge-up means labor cost per SEAT rises less than cost per employee — 150 pax vs. 50 on replaced planes.", bg: C.lightBlue },
    { lbl: 'Landing Fees', txt: "+12% = UAL's own $2.7B EWR Terminal A + ORD T2 satellite + DEN gates. Premium product in OpEx, not just CapEx.", bg: C.navyLight },
    { lbl: 'Distribution –5.5%', txt: 'Direct channels handle 85%+ of check-ins. GDS fee avoidance has legs as app engagement accelerates.', bg: C.gold },
  ];

  const tableBottom = contentY + 0.1 + 0.35 * 10 + 0.06;
  const cbH = 5.545 - tableBottom - 0.08;
  const cbW = (9.5 - 0.1) / 3;

  cbs.forEach((c, i) => {
    const x = 0.25 + i * (cbW + 0.05);
    sld.addShape(prs.ShapeType.rect, { x, y: tableBottom, w: cbW, h: cbH, fill: { color: c.bg }, shadow: makeShadow() });
    sld.addText(c.lbl + ': ' + c.txt, {
      x: x + 0.1, y: tableBottom + 0.06, w: cbW - 0.2, h: cbH - 0.12,
      fontSize: 7.8, color: c.bg === C.navyLight ? C.white : C.darkSlate, wrap: true, valign: 'middle', margin: 0,
    });
  });
}

// ════════════════════════════════════════════════════════════
// SLIDE 7 — Regional Revenue Mix
// ════════════════════════════════════════════════════════════
{
  const sld = contentSlide(prs);
  const contentY = addNavyHeader(sld, 'Two-Speed Airline: Regional Revenue Trends', 'International outperformance driving revenue quality improvement');

  const rows = [
    ['Region', 'FY2023', 'FY2024', 'FY2025E', '3-Yr CAGR'],
    ['Domestic', '$31.874B', '$28.458B', '~$31.6B', '–0.4%'],
    ['Atlantic — Europe', '$8.068B', '$8.099B', '~$10.1B', '+11.9%'],
    ['Atlantic — MEIA', '$1.000B', '$1.141B', '~$1.3B', '+14.0%'],
    ['Pacific', '$4.437B', '$5.226B', '~$6.88B', '+24.5%'],
    ['Latin America', '$3.667B', '$4.827B', '~$5.53B', '+22.8%'],
  ];

  styledTable(sld, rows, 0.25, contentY + 0.1, 5.55,
    [1.9, 0.92, 0.92, 0.92, 0.89], 0.4,
    C.navy, C.white, C.white, C.textDark, 9
  );

  // Right: Q4 spotlight callouts
  const spots = [
    { region: 'MEIA', color: C.navy, txt: '+58.7% Q4 revenue | +56% capacity | EWR-Dubai, EWR-Delhi, SFO-Bangalore ramping | India = structural opportunity' },
    { region: 'Pacific', color: C.navyLight, txt: '+10.1% Q4 | Only U.S. carrier to Bangkok, Adelaide, Ho Chi Minh City | Monopoly routes support yield' },
    { region: 'Domestic', color: C.lightBlue, txt: '+2.0% Q4 revenue but PRASM –1.9% | Soft all year | Mid-year capacity discipline in progress' },
    { region: 'Latin', color: C.textMed, txt: '+0.5% Q4 but PRASM –7.6% | Brazilian competition + macro | ~10% of pax revenue' },
  ];

  const spotH = 0.9;
  const spotGap = 0.06;
  spots.forEach((s, i) => {
    const x = 6.0;
    const y = contentY + 0.1 + i * (spotH + spotGap);
    sld.addShape(prs.ShapeType.rect, { x, y, w: 3.75, h: spotH, fill: { color: s.color }, shadow: makeShadow() });
    sld.addText(s.region, { x: x + 0.12, y: y + 0.06, w: 3.52, h: 0.26, fontSize: 10, bold: true, color: s.color === C.lightBlue ? C.navy : C.gold, margin: 0 });
    sld.addText(s.txt, { x: x + 0.12, y: y + 0.35, w: 3.52, h: 0.5, fontSize: 8, color: s.color === C.lightBlue ? C.textDark : C.white, wrap: true, valign: 'top', margin: 0 });
  });

  // A321XLR callout at bottom
  const xlrY = contentY + 0.1 + 4 * (spotH + spotGap) + 0.1;
  sld.addShape(prs.ShapeType.rect, { x: 0.25, y: xlrY, w: 9.5, h: 5.545 - xlrY, fill: { color: C.gold } });
  sld.addText('A321XLR (4,700 nm range) unlocks routes impossible for 757 or A321neo: EWR to Bari, Split, Glasgow, Santiago de Compostela — point-to-point international with narrowbody economics.', {
    x: 0.4, y: xlrY + 0.06, w: 9.2, h: 5.545 - xlrY - 0.12, fontSize: 8.5, color: C.darkSlate, italic: true, valign: 'middle', wrap: true, margin: 0,
  });
}

// ════════════════════════════════════════════════════════════
// SLIDE 8 — Revenue Quality: Premium & Loyalty
// ════════════════════════════════════════════════════════════
{
  const sld = contentSlide(prs);
  const contentY = addNavyHeader(sld, 'Revenue Quality: Premium & Loyalty Flywheel', 'Structural shift to higher-margin revenue streams');

  // Left panel — Premium
  sld.addShape(prs.ShapeType.rect, { x: 0.25, y: contentY + 0.1, w: 4.55, h: 0.3, fill: { color: C.navy } });
  sld.addText('Premium Cabin Performance', { x: 0.25, y: contentY + 0.1, w: 4.55, h: 0.3, fontSize: 10, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0 });

  const premStats = [
    ['Premium cabin FY2025', '+11% YoY'],
    ['Basic Economy FY2025', '+5% YoY'],
    ['Premium seats flown', '27.4M (12% of seats)'],
    ['Premium seats/N.Am. dep. since 2021', '+40%'],
    ['Narrowbodies w/ Signature Interior', '119 (68% of fleet)'],
    ['Signature Interior NPS premium', '+10 pts'],
  ];
  premStats.forEach(([lbl, val], i) => {
    const bg = i % 2 === 0 ? C.lightBlue : C.altRow;
    sld.addShape(prs.ShapeType.rect, { x: 0.25, y: contentY + 0.44 + i * 0.4, w: 4.55, h: 0.38, fill: { color: bg } });
    sld.addText(lbl, { x: 0.35, y: contentY + 0.47 + i * 0.4, w: 2.85, h: 0.3, fontSize: 8.5, color: C.textDark, valign: 'middle', margin: 0 });
    sld.addText(val, { x: 3.25, y: contentY + 0.47 + i * 0.4, w: 1.48, h: 0.3, fontSize: 9, bold: true, color: C.navy, align: 'right', valign: 'middle', margin: 0 });
  });

  // Right panel — MileagePlus
  sld.addShape(prs.ShapeType.rect, { x: 5.05, y: contentY + 0.1, w: 4.7, h: 0.3, fill: { color: C.navyLight } });
  sld.addText('MileagePlus / Loyalty', { x: 5.05, y: contentY + 0.1, w: 4.7, h: 0.3, fontSize: 10, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0 });

  const loyaltyStats = [
    ['Total members', '130M+'],
    ['Chase co-brand spend growth Q4', '+14%'],
    ['Chase co-brand spend growth FY2025', '+12%'],
    ['New cards/year (3rd consecutive yr)', '1M+'],
    ['Chase deal expiry', '2029'],
    ['CEO target: MileagePlus profit growth', '2× by 2030'],
  ];
  loyaltyStats.forEach(([lbl, val], i) => {
    const bg = i % 2 === 0 ? C.lightBlue : C.altRow;
    sld.addShape(prs.ShapeType.rect, { x: 5.05, y: contentY + 0.44 + i * 0.4, w: 4.7, h: 0.38, fill: { color: bg } });
    sld.addText(lbl, { x: 5.15, y: contentY + 0.47 + i * 0.4, w: 3.25, h: 0.3, fontSize: 8.5, color: C.textDark, valign: 'middle', margin: 0 });
    sld.addText(val, { x: 8.3, y: contentY + 0.47 + i * 0.4, w: 1.4, h: 0.3, fontSize: 9, bold: true, color: C.navy, align: 'right', valign: 'middle', margin: 0 });
  });

  // Gold valuation callout
  const callY = contentY + 0.44 + 6 * 0.4 + 0.1;
  sld.addShape(prs.ShapeType.rect, { x: 0.25, y: callY, w: 9.5, h: 5.545 - callY, fill: { color: C.gold }, shadow: makeShadow() });
  sld.addText('MileagePlus Standalone Valuation Opportunity', { x: 0.4, y: callY + 0.08, w: 9.2, h: 0.28, fontSize: 11, bold: true, color: C.darkSlate, margin: 0 });
  sld.addText(
    'In 2020, bankers valued MileagePlus at ~$22B as loan collateral — and the program has grown since. UAL total market cap today: ~$28B. This implies minimal standalone loyalty value is priced in — a significant re-rating opportunity vs. the Delta/AmEx model. New April 2026: primary cardholders earn 2× more miles; 10–15% discount on award redemptions.',
    { x: 0.4, y: callY + 0.4, w: 9.2, h: 5.545 - callY - 0.48, fontSize: 8.5, color: C.darkSlate, wrap: true, valign: 'top', margin: 0 }
  );
}

// ════════════════════════════════════════════════════════════
// SLIDE 9 — Fleet Modernization
// ════════════════════════════════════════════════════════════
{
  const sld = contentSlide(prs);
  const contentY = addNavyHeader(sld, 'United Next Fleet Transformation', 'Largest delivery pipeline in UAL history');

  // Fleet timeline — 4 equal boxes
  const ftData = [
    { yr: "Dec'23", n: 1358 },
    { yr: "Dec'24", n: 1406 },
    { yr: "Dec'25", n: 1490 },
    { yr: "Dec'26E", n: 1610 },
  ];
  const ftW = 2.3, ftH = 0.82, ftGap = 0.06;
  const ftTotalW = ftData.length * ftW + (ftData.length - 1) * ftGap;
  const ftX0 = (10 - ftTotalW) / 2;

  ftData.forEach((f, i) => {
    const x = ftX0 + i * (ftW + ftGap);
    const isLast = i === ftData.length - 1;
    sld.addShape(prs.ShapeType.rect, { x, y: contentY + 0.12, w: ftW, h: ftH, fill: { color: isLast ? C.gold : C.navy }, shadow: makeShadow() });
    sld.addText(String(f.n), { x, y: contentY + 0.18, w: ftW, h: 0.46, fontSize: 22, bold: true, color: isLast ? C.darkSlate : C.white, align: 'center', margin: 0 });
    sld.addText(f.yr, { x, y: contentY + 0.65, w: ftW, h: 0.24, fontSize: 9, color: isLast ? C.darkSlate : C.lightBlue, align: 'center', margin: 0 });
  });

  // Delivery highlights
  const dlY = contentY + 0.12 + ftH + 0.1;
  sld.addShape(prs.ShapeType.rect, { x: 0.25, y: dlY, w: 4.6, h: 0.44, fill: { color: C.lightBlue } });
  sld.addText('2025 Deliveries: 71 narrowbody + 11 widebody = 82 aircraft', { x: 0.35, y: dlY + 0.04, w: 4.4, h: 0.36, fontSize: 9, color: C.navy, bold: true, valign: 'middle', margin: 0 });
  sld.addShape(prs.ShapeType.rect, { x: 5.1, y: dlY, w: 4.65, h: 0.44, fill: { color: C.navyLight } });
  sld.addText('2026 Plan: ~100 NB + ~20 WB = ~120 aircraft (largest U.S. widebody intake since 1988)', { x: 5.2, y: dlY + 0.04, w: 4.45, h: 0.36, fontSize: 9, color: C.white, bold: true, valign: 'middle', wrap: true, margin: 0 });

  // Order book table
  const rows = [
    ['Aircraft', 'Purpose', 'Orders Remaining'],
    ['Boeing 787-9 Elevated', 'Long-haul premium', '47+ next wave'],
    ['Boeing 737 MAX 8/9/10', 'Domestic narrowbody', '200+'],
    ['Airbus A321neo', 'General domestic', '18+'],
    ['A321neo Coastliner', 'Domestic transcon lie-flat', "50 total, 40 by Apr'28"],
    ['Airbus A321XLR', 'New intl routes', '50 total, 25+ by 2028'],
    ['CRJ450 (SkyWest)', 'Premium regional', '70 total, 50 by 2028'],
    ['Total United Next', '', '630+ by 2034'],
  ];
  const tableY = dlY + 0.44 + 0.1;
  styledTable(sld, rows, 0.25, tableY, 6.5,
    [2.2, 2.3, 2.0], 0.31,
    C.navy, C.white, C.white, C.textDark, 8.5
  );

  // Right side callouts
  const rcx = 7.0, rcw = 2.75;
  sld.addShape(prs.ShapeType.rect, { x: rcx, y: tableY, w: rcw, h: 1.24, fill: { color: C.lightBlue }, shadow: makeShadow() });
  sld.addText('Gauge-Up Impact', { x: rcx + 0.1, y: tableY + 0.06, w: rcw - 0.2, h: 0.26, fontSize: 9, bold: true, color: C.navy, margin: 0 });
  sld.addText('Avg mainline seats/dep: 151 (2019) → 174 (2025). 50-seat RJs → 150+ seat mainline = 20% better fuel/seat, lower crew cost per passenger.', {
    x: rcx + 0.1, y: tableY + 0.34, w: rcw - 0.2, h: 0.86, fontSize: 8, color: C.textDark, wrap: true, valign: 'top', margin: 0,
  });

  const rc2y = tableY + 1.24 + 0.1;
  sld.addShape(prs.ShapeType.rect, { x: rcx, y: rc2y, w: rcw, h: 5.545 - rc2y, fill: { color: C.gold }, shadow: makeShadow() });
  sld.addText('Boeing Risk', { x: rcx + 0.1, y: rc2y + 0.06, w: rcw - 0.2, h: 0.26, fontSize: 9, bold: true, color: C.darkSlate, margin: 0 });
  sld.addText('FAA 38 MAX/month production cap. MAX 10 certification pending — 200 aircraft waiting. Each month delay = ~$30–40M deferred efficiency. Mitigation: A321neo/XLR Airbus orders provide operational hedge.', {
    x: rcx + 0.1, y: rc2y + 0.34, w: rcw - 0.2, h: 5.545 - rc2y - 0.4, fontSize: 7.8, color: C.darkSlate, wrap: true, valign: 'top', margin: 0,
  });
}

// ════════════════════════════════════════════════════════════
// SLIDE 10 — New Product Architecture
// ════════════════════════════════════════════════════════════
{
  const sld = contentSlide(prs);
  const contentY = addNavyHeader(sld, 'New Product Architecture: 2026–2028 Launches', 'Four differentiated products entering service');

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
      when: 'Summer 2026 | SFO/LAX ↔ EWR/JFK',
      bullets: [
        '20 Polaris — FIRST lie-flat on any U.S. domestic narrowbody',
        '12 Premium Plus — FIRST on domestic narrowbody',
        "161 total | 50 ordered, 40 by Apr'28",
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

  const bw = 4.73, bh = 2.08;
  const rowGap = 0.12, colGap = 0.29;
  products.forEach((p, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.14 + col * (bw + colGap);
    const y = contentY + 0.1 + row * (bh + rowGap);
    sld.addShape(prs.ShapeType.rect, { x, y, w: bw, h: bh, fill: { color: p.color }, shadow: makeShadow() });
    sld.addShape(prs.ShapeType.rect, { x, y, w: bw, h: 0.32, fill: { color: C.gold } });
    sld.addText(p.name, { x: x + 0.12, y: y + 0.03, w: bw - 0.24, h: 0.27, fontSize: 10, bold: true, color: C.darkSlate, margin: 0 });
    sld.addText(p.when, { x: x + 0.12, y: y + 0.36, w: bw - 0.24, h: 0.22, fontSize: 8, color: C.goldLight, italic: true, margin: 0 });
    p.bullets.forEach((b, j) => {
      sld.addText([{ text: b, options: { bullet: true } }], {
        x: x + 0.12, y: y + 0.61 + j * 0.35, w: bw - 0.24, h: 0.33,
        fontSize: 8.5, color: C.white, wrap: true, margin: 0,
      });
    });
  });
}

// ════════════════════════════════════════════════════════════
// SLIDE 11 — Starlink + Kinective Media
// ════════════════════════════════════════════════════════════
{
  const sld = contentSlide(prs);
  const contentY = addNavyHeader(sld, 'Starlink + Kinective Media', 'Connectivity as infrastructure — monetizing the world\'s largest captive audience');

  // Left: Starlink rollout
  // contentY ≈ 1.095; left column rows: 6 × 0.40 = 2.4" → bottom at 1.095+0.1+0.28+0.04+2.4 = 3.915
  sld.addShape(prs.ShapeType.rect, { x: 0.25, y: contentY + 0.1, w: 4.55, h: 0.28, fill: { color: C.navy } });
  sld.addText('Starlink Rollout Timeline', { x: 0.25, y: contentY + 0.1, w: 4.55, h: 0.28, fontSize: 10, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0 });

  const events = [
    ['May 2025', 'First E175 regional enters service'],
    ['Oct 15, 2025', 'First mainline commercial flight (EWR)'],
    ['Feb 2026', '300+ mainline aircraft equipped'],
    ['End 2026', '800+ aircraft (50+ installs/month)'],
    ['End 2027', 'Full fleet completion'],
    ['All time', 'FREE for all MileagePlus members'],
  ];
  events.forEach(([dt, txt], i) => {
    const y = contentY + 0.42 + i * 0.40;
    const bg = i % 2 === 0 ? C.lightBlue : C.altRow;
    sld.addShape(prs.ShapeType.rect, { x: 0.25, y, w: 4.55, h: 0.36, fill: { color: bg } });
    sld.addText(dt, { x: 0.35, y: y + 0.03, w: 1.3, h: 0.3, fontSize: 7.5, bold: true, color: C.navy, valign: 'middle', margin: 0 });
    sld.addText(txt, { x: 1.7, y: y + 0.03, w: 3.0, h: 0.3, fontSize: 8, color: C.textDark, valign: 'middle', margin: 0 });
  });

  // Right: Kinective Media
  // right column: header at contentY+0.1, assets 3×0.44=1.32, quote 0.9, total ~2.56"
  sld.addShape(prs.ShapeType.rect, { x: 5.05, y: contentY + 0.1, w: 4.7, h: 0.28, fill: { color: C.navyLight } });
  sld.addText('Kinective Media — The Real Story', { x: 5.05, y: contentY + 0.1, w: 4.7, h: 0.28, fontSize: 10, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0 });

  sld.addText('Three assets no competitor possesses simultaneously:', { x: 5.15, y: contentY + 0.44, w: 4.5, h: 0.22, fontSize: 9, bold: true, color: C.navy, margin: 0 });

  const kinAssets = [
    ['227,000+', 'seatback screens — captive audience'],
    ['108M', 'unique annual flyers with known demographics'],
    ['130M', 'MileagePlus members with first-party purchase data'],
  ];
  kinAssets.forEach(([num, lbl], i) => {
    const ky = contentY + 0.70 + i * 0.44;
    sld.addShape(prs.ShapeType.rect, { x: 5.05, y: ky, w: 4.7, h: 0.40, fill: { color: i % 2 === 0 ? C.lightBlue : C.altRow } });
    sld.addText(num, { x: 5.15, y: ky + 0.04, w: 1.25, h: 0.32, fontSize: 13, bold: true, color: C.navy, align: 'center', margin: 0 });
    sld.addText(lbl, { x: 6.45, y: ky + 0.04, w: 3.2, h: 0.32, fontSize: 8, color: C.textDark, valign: 'middle', margin: 0 });
  });

  // CFO quote — starts after kinAssets: contentY + 0.70 + 3*0.44 + 0.08 = contentY + 2.10
  const quoteY = contentY + 0.70 + 3 * 0.44 + 0.08;
  sld.addShape(prs.ShapeType.rect, { x: 5.05, y: quoteY, w: 4.7, h: 0.88, fill: { color: C.navy }, shadow: makeShadow() });
  sld.addText('CFO on Q4 earnings call:', { x: 5.15, y: quoteY + 0.06, w: 4.5, h: 0.2, fontSize: 8, bold: true, color: C.gold, margin: 0 });
  sld.addText('"Kinective will really start to accelerate in \'26 and beyond." Media margins are dramatically higher than seat revenue margins. Even $200–400M of high-margin ad revenue materially changes the EBITDA profile.', {
    x: 5.15, y: quoteY + 0.28, w: 4.5, h: 0.56, fontSize: 8, color: C.white, wrap: true, valign: 'top', italic: true, margin: 0,
  });

  // Bottom callout — starts after both columns, max of their bottoms
  // Left column bottom: contentY + 0.42 + 6*0.40 = contentY + 2.82
  // Right column bottom: quoteY + 0.88 = contentY + 2.10 + 0.88 = contentY + 2.98
  const btmY = contentY + 2.98 + 0.1;   // ≈ 1.095 + 2.98 + 0.1 = 4.175
  sld.addShape(prs.ShapeType.rect, { x: 0.25, y: btmY, w: 9.5, h: 5.545 - btmY, fill: { color: C.gold } });
  sld.addText("Market currently values Kinective at approximately ZERO. Better audience profile than most digital ad platforms — behavioral and purchase-intent data, not just demographic. Even $200–400M of high-margin ad revenue materially changes EBITDA. Competitor Delta has SkyMiles advertising — no competitor has UAL's level of connectivity + behavioral data integration.", {
    x: 0.4, y: btmY + 0.08, w: 9.2, h: 5.545 - btmY - 0.16, fontSize: 8.5, color: C.darkSlate, wrap: true, valign: 'middle', margin: 0,
  });
}

// ════════════════════════════════════════════════════════════
// SLIDE 12 — Operational Performance
// ════════════════════════════════════════════════════════════
{
  const sld = contentSlide(prs);
  const contentY = addNavyHeader(sld, 'Operational Excellence: Record Performance in 2025', 'Reliability and customer metrics at all-time highs');

  const stats = [
    { val: '181.1M', lbl: 'Passengers\n(Company Record)', sub: '+4.3% YoY' },
    { val: '#2', lbl: 'D:14 On-Time\nIndustry Rank', sub: 'Best-ever system completion' },
    { val: '303', lbl: 'Avg Daily Widebody\nDepartures', sub: 'Company record' },
    { val: '1M+', lbl: 'ConnectionSaver\nRescues', sub: '+42% YoY' },
    { val: '85%', lbl: 'Digital Check-In\nRate', sub: 'Q1 record' },
    { val: 'Record', lbl: 'NPS Score\nQ4 2025', sub: "Nov'25 highest ever" },
  ];

  // 3×2 grid
  const sw = 3.1, sh = 1.25;
  const rowGap = 0.08;
  const gridTop = contentY + 0.12;
  stats.forEach((s, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const x = 0.25 + col * (sw + 0.09);
    const y = gridTop + row * (sh + rowGap);
    sld.addShape(prs.ShapeType.rect, { x, y, w: sw, h: sh, fill: { color: i % 2 === 0 ? C.navy : C.navyLight }, shadow: makeShadow() });
    sld.addText(s.val, { x: x + 0.1, y: y + 0.1, w: sw - 0.2, h: 0.48, fontSize: 26, bold: true, color: C.gold, align: 'center', margin: 0 });
    sld.addText(s.lbl, { x: x + 0.1, y: y + 0.6, w: sw - 0.2, h: 0.4, fontSize: 9, color: C.white, align: 'center', wrap: true, margin: 0 });
    sld.addText(s.sub, { x: x + 0.1, y: y + 1.03, w: sw - 0.2, h: 0.18, fontSize: 7.5, color: C.goldLight, align: 'center', margin: 0 });
  });

  // Demand signals
  const demandY = gridTop + 2 * (sh + rowGap) + 0.1;
  sld.addShape(prs.ShapeType.rect, { x: 0.25, y: demandY, w: 9.5, h: 5.545 - demandY, fill: { color: C.gold }, shadow: makeShadow() });
  sld.addText('Early 2026 Demand Signals — Real-Time Leading Indicators', { x: 0.4, y: demandY + 0.06, w: 9.2, h: 0.28, fontSize: 11, bold: true, color: C.darkSlate, margin: 0 });
  sld.addText(
    'Week of Jan 4, 2026: Highest flown revenue week in UAL history.  |  Week of Jan 11, 2026: Highest ticketing week AND highest business sales week ever recorded simultaneously.\n\nBusiness travel leads leisure in recovery cycles — entering 2026 with record business bookings is the strongest real-time indicator that $12–$14 EPS guidance is achievable.',
    { x: 0.4, y: demandY + 0.38, w: 9.2, h: 5.545 - demandY - 0.46, fontSize: 8.5, color: C.darkSlate, wrap: true, valign: 'top', margin: 0 }
  );
}

// ════════════════════════════════════════════════════════════
// SLIDE 13 — Balance Sheet & Deleveraging
// ════════════════════════════════════════════════════════════
{
  const sld = contentSlide(prs);
  const contentY = addNavyHeader(sld, 'Balance Sheet Normalization: Path to Investment Grade', 'Leverage declining; MileagePlus fully unencumbered');

  const rows = [
    ['Period', 'Total Liquidity', 'Total Debt + Leases', 'Net Leverage', 'S&P Rating'],
    ['Dec 2023', '$16.1B', '$29.3B', '2.9×', 'BB'],
    ['Dec 2024', '$17.4B', '$28.7B', '2.4×', 'BB+'],
    ['Mar 2025', '$18.3B', '$27.7B', '2.0×', 'BB+'],
    ['Jun 2025', '$18.6B', '$27.1B', '2.0×', 'BB+'],
    ['Sep 2025', '$16.3B', '$25.4B', '2.1×', 'BB+'],
    ['Dec 2025', '$15.2B', '$25.0B', '2.2×', 'BB+'],
    ['Dec 2026E', '—', '—', '<2.0× target', 'IG target'],
  ];
  styledTable(sld, rows, 0.25, contentY + 0.1, 6.25,
    [1.22, 1.52, 1.52, 1.0, 0.99], 0.37,
    C.navy, C.white, C.white, C.textDark, 9
  );

  // Key events panel
  sld.addShape(prs.ShapeType.rect, { x: 6.7, y: contentY + 0.1, w: 3.05, h: 1.92, fill: { color: C.lightBlue }, shadow: makeShadow() });
  sld.addText('Key Events', { x: 6.82, y: contentY + 0.16, w: 2.82, h: 0.28, fontSize: 9.5, bold: true, color: C.navy, margin: 0 });
  const evts = [
    'Q3 2025: Prepaid final $1.5B of MileagePlus bonds — Full $6.8B COVID-era collateralized debt retired',
    'MileagePlus asset now fully unencumbered',
    'Interest expense: $1.629B (2024) → $1.373B (2025) = –15.7% — flows straight to EPS',
  ];
  evts.forEach((e, i) => {
    sld.addText([{ text: e, options: { bullet: true } }], {
      x: 6.82, y: contentY + 0.5 + i * 0.44, w: 2.82, h: 0.42,
      fontSize: 8, color: C.textDark, wrap: true, valign: 'top', margin: 0,
    });
  });

  // IG impact callout
  const igY = contentY + 0.1 + 0.37 * 8 + 0.12;
  sld.addShape(prs.ShapeType.rect, { x: 0.25, y: igY, w: 9.5, h: 5.545 - igY, fill: { color: C.gold }, shadow: makeShadow() });
  sld.addText('Three Consequences of Investment-Grade Upgrade', { x: 0.4, y: igY + 0.07, w: 9.2, h: 0.28, fontSize: 11, bold: true, color: C.darkSlate, margin: 0 });
  const igItems = [
    '1. Lower financing cost on $25B fleet debt → $125–190M annual interest savings',
    '2. IG mandate funds open to UAL equity — materially larger institutional buyer universe',
    '3. Multiple re-rating: COVID-risk narrative to investment-quality enterprise complete. Delta carries BBB-. UAL at 2.2× leverage and falling — probability of upgrade is high. Question: own before or after the rating action.',
  ];
  igItems.forEach((it, i) => {
    sld.addText(it, { x: 0.4, y: igY + 0.4 + i * 0.3, w: 9.2, h: 0.28, fontSize: 8.5, color: C.darkSlate, wrap: true, margin: 0 });
  });
}

// ════════════════════════════════════════════════════════════
// SLIDE 14 — Capital Allocation
// ════════════════════════════════════════════════════════════
{
  const sld = contentSlide(prs);
  const contentY = addNavyHeader(sld, 'Capital Allocation: CapEx Bulge → FCF Normalization', 'Fleet investment peaking; FCF trajectory improving post-2027');

  // Cash flow table
  const rows = [
    ['Metric', 'FY2023', 'FY2024', 'FY2025', 'FY2026E'],
    ['Operating Cash Flow ($B)', '$6.911', '$9.4', '$8.4', '—'],
    ['Net CapEx ($B)', '$7.948', '$5.6', '$5.874', '<$8.0'],
    ['Free Cash Flow ($B)', '($1.037)', '$3.4', '$2.7', '~$2.7'],
  ];
  styledTable(sld, rows, 0.25, contentY + 0.1, 5.55,
    [2.2, 0.84, 0.84, 0.84, 0.83], 0.42,
    C.navy, C.white, C.white, C.textDark, 9.5
  );

  // Priority stack header
  sld.addShape(prs.ShapeType.rect, { x: 6.0, y: contentY + 0.1, w: 3.75, h: 0.3, fill: { color: C.navy } });
  sld.addText('Capital Priority Stack', { x: 6.0, y: contentY + 0.1, w: 3.75, h: 0.3, fontSize: 10, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0 });

  const pItems = [
    { n: '1', lbl: 'Fleet CapEx (Non-negotiable)', txt: '$5.9B in 2025, <$8B guided 2026; 120 aircraft deliveries', bg: C.navy },
    { n: '2', lbl: 'Debt Reduction', txt: '$3.7B total debt reduction FY2025; $3–4B/yr pace', bg: C.navyLight },
    { n: '3', lbl: 'Share Buybacks', txt: 'FY2024: $81M → FY2025: $640M → expanding. $1.5B authorization outstanding. ~325M shares 2026E.', bg: C.lightBlue },
    { n: '4', lbl: 'No Dividend (near-term)', txt: 'Free cash flow directed to fleet + delevering', bg: C.altRow },
  ];
  pItems.forEach((p, i) => {
    const y = contentY + 0.44 + i * 0.7;
    sld.addShape(prs.ShapeType.rect, { x: 6.0, y, w: 3.75, h: 0.66, fill: { color: p.bg } });
    sld.addShape(prs.ShapeType.rect, { x: 6.0, y, w: 0.34, h: 0.66, fill: { color: C.gold } });
    sld.addText(p.n, { x: 6.0, y, w: 0.34, h: 0.66, fontSize: 12, bold: true, color: C.darkSlate, align: 'center', valign: 'middle', margin: 0 });
    sld.addText(p.lbl, { x: 6.38, y: y + 0.05, w: 3.3, h: 0.25, fontSize: 9, bold: true, color: [C.altRow, C.lightBlue].includes(p.bg) ? C.navy : C.white, margin: 0 });
    sld.addText(p.txt, { x: 6.38, y: y + 0.32, w: 3.3, h: 0.3, fontSize: 7.8, color: [C.altRow, C.lightBlue].includes(p.bg) ? C.textDark : C.white, wrap: true, margin: 0 });
  });

  // Post-2027 callout
  const fcfY = contentY + 0.1 + 0.42 * 4 + 0.16;
  sld.addShape(prs.ShapeType.rect, { x: 0.25, y: fcfY, w: 5.55, h: 5.545 - fcfY, fill: { color: C.gold }, shadow: makeShadow() });
  sld.addText('Post-2027 FCF Normalization Thesis', { x: 0.4, y: fcfY + 0.08, w: 5.25, h: 0.28, fontSize: 10, bold: true, color: C.darkSlate, margin: 0 });
  sld.addText('CapEx bulge persists through 2027. Post-2028 steady-state: $5–6B CapEx on $8–9B operating CF = $3–4B sustained FCF. At 320M shares = ~$10–12/share FCF. Airlines historically trade at 5–8× FCF. At normalized FCF, current share price implies significant upside for investors who can look through the delivery cycle.', {
    x: 0.4, y: fcfY + 0.4, w: 5.25, h: 5.545 - fcfY - 0.48, fontSize: 8, color: C.darkSlate, wrap: true, valign: 'top', margin: 0,
  });
}

// ════════════════════════════════════════════════════════════
// SLIDE 15 — Sustainability
// ════════════════════════════════════════════════════════════
{
  const sld = contentSlide(prs);
  const contentY = addNavyHeader(sld, 'Sustainability: Progress, Ambition & Honest Framing', 'Science-based targets; no greenwashing');

  // ── LEFT COLUMN ──────────────────────────────────────────
  // contentY ≈ 1.095. Available: 5.545 - 1.095 = 4.45"
  // Commitments header
  sld.addShape(prs.ShapeType.rect, { x: 0.25, y: contentY + 0.12, w: 4.55, h: 0.3, fill: { color: C.navy } });
  sld.addText('Key Commitments', { x: 0.25, y: contentY + 0.12, w: 4.55, h: 0.3, fontSize: 10, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0 });

  const cmts = [
    'Net zero GHG by 2050 — first global airline without traditional offsets',
    '50% emissions intensity reduction by 2035 vs. 2019 (SBTi May 2023)',
    'CDP Climate Score: A-',
  ];
  cmts.forEach((c, i) => {
    sld.addShape(prs.ShapeType.rect, { x: 0.25, y: contentY + 0.46 + i * 0.42, w: 4.55, h: 0.38, fill: { color: i % 2 === 0 ? C.lightBlue : C.altRow } });
    sld.addText([{ text: c, options: { bullet: true } }], {
      x: 0.35, y: contentY + 0.48 + i * 0.42, w: 4.35, h: 0.34, fontSize: 8.5, color: C.textDark, wrap: true, margin: 0,
    });
  });

  // Intensity bar chart — starts at contentY + 0.46 + 3*0.42 = contentY + 1.72
  const intHdrY = contentY + 1.72;
  sld.addShape(prs.ShapeType.rect, { x: 0.25, y: intHdrY, w: 4.55, h: 0.26, fill: { color: C.navyLight } });
  sld.addText('CO2 Intensity (MT per M ASMs) — Declining', {
    x: 0.25, y: intHdrY, w: 4.55, h: 0.26, fontSize: 9, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0,
  });

  const intData = [['2021','187.5'], ['2022','176.2'], ['2023','169.0'], ['2024','167.3']];
  // bar area height 0.72"; value label placed INSIDE bar near top
  const intBarAreaH = 0.72;
  const intBarBaseY = intHdrY + 0.26 + intBarAreaH;
  const intBarW = 4.55 / intData.length - 0.14;
  intData.forEach(([yr, val], i) => {
    const norm = parseFloat(val) / 192;
    const bh = intBarAreaH * norm;
    const bx = 0.25 + i * (intBarW + 0.14);
    const by = intBarBaseY - bh;
    sld.addShape(prs.ShapeType.rect, { x: bx, y: by, w: intBarW, h: bh, fill: { color: i === 3 ? C.navy : C.lightBlue } });
    // Label inside bar near top
    sld.addText(val, { x: bx, y: by + 0.04, w: intBarW, h: 0.2, fontSize: 7.5, bold: true, color: i === 3 ? C.white : C.navy, align: 'center', margin: 0 });
    sld.addText(yr, { x: bx, y: intBarBaseY + 0.03, w: intBarW, h: 0.18, fontSize: 8, color: C.textMed, align: 'center', margin: 0 });
  });

  // SAF progress table — starts at intBarBaseY + 0.22 = intHdrY + 0.26 + 0.85 + 0.22
  const safHdrY = intBarBaseY + 0.22;
  sld.addShape(prs.ShapeType.rect, { x: 0.25, y: safHdrY, w: 4.55, h: 0.26, fill: { color: C.navy } });
  sld.addText('SAF Progress', { x: 0.25, y: safHdrY, w: 4.55, h: 0.26, fontSize: 9, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0 });
  const safRows = [
    ['Year', 'SAF Gallons', 'CO2e Avoided'],
    ['2021', '0.6M', '5,953 MT'],
    ['2022', '2.9M', '29,362 MT'],
    ['2023', '7.3M', '68,370 MT'],
    ['2024', '13.6M', '126,174 MT'],
  ];
  // 5 rows × 0.22 = 1.10; safHdrY+0.26+1.10 must be ≤ 5.545
  styledTable(sld, safRows, 0.25, safHdrY + 0.26, 4.55, [1.0, 1.78, 1.77], 0.22, C.navyLight, C.white, C.white, C.textDark, 8);

  // ── RIGHT COLUMN ──────────────────────────────────────────
  // Honest Assessment — top 2.1"
  const rightH = 2.1;
  sld.addShape(prs.ShapeType.rect, { x: 5.05, y: contentY + 0.12, w: 4.7, h: rightH, fill: { color: C.gold }, shadow: makeShadow() });
  sld.addText('Honest ESG Framing for Institutional Investors', {
    x: 5.18, y: contentY + 0.17, w: 4.44, h: 0.3, fontSize: 10, bold: true, color: C.darkSlate, margin: 0,
  });
  sld.addText(
    'Absolute emissions RISING with capacity growth — will continue as UAL adds aircraft. 2035 target is intensity-based (achievable with new aircraft). 2050 net zero requires green hydrogen / next-gen SAF not yet at commercial scale. SAF at 13.6M gallons vs. 4.2B+ total = <0.35% penetration. UAV Ventures ($200M+ SAF Fund) is strategic supply-curve investment. Eco-Skies Alliance: 50+ corporate partners co-funding SAF development.',
    { x: 5.18, y: contentY + 0.52, w: 4.44, h: rightH - 0.6, fontSize: 8.5, color: C.darkSlate, wrap: true, valign: 'top', margin: 0 },
  );

  // Ventures panel — fills rest of right column
  const ventY = contentY + 0.12 + rightH + 0.1;
  sld.addShape(prs.ShapeType.rect, { x: 5.05, y: ventY, w: 4.7, h: 5.545 - ventY, fill: { color: C.lightBlue }, shadow: makeShadow() });
  sld.addText('UAL Ventures Portfolio ($200M+)', { x: 5.18, y: ventY + 0.1, w: 4.44, h: 0.28, fontSize: 9.5, bold: true, color: C.navy, margin: 0 });
  const ventures = [
    'Twelve (CO2-to-SAF) — direct air capture pathway',
    'ZeroAvia (hydrogen-electric) — regional aviation',
    'Dimensional Energy, Cemvita, Svante (carbon capture)',
    'Eco-Skies Alliance: 50+ corporate partners co-investing in SAF supply',
  ];
  ventures.forEach((v, i) => {
    sld.addText([{ text: v, options: { bullet: true } }], {
      x: 5.18, y: ventY + 0.42 + i * 0.38, w: 4.44, h: 0.35, fontSize: 8.5, color: C.textDark, wrap: true, margin: 0,
    });
  });
}

// ════════════════════════════════════════════════════════════
// SLIDE 16 — 2026 Guidance & Risk/Reward
// ════════════════════════════════════════════════════════════
{
  const sld = contentSlide(prs);
  const contentY = addNavyHeader(sld, '2026 Guidance & Risk/Reward Framework', 'EPS inflection supported by multiple convergent tailwinds');

  // Guidance table
  const gRows = [
    ['Metric', 'Q1 2026E', 'FY2026E', 'vs. FY2025'],
    ['Adj. EPS', '$1.00–$1.50', '$12.00–$14.00', '+13% to +32%'],
    ['EPS midpoint', '$1.25', '$13.00', '+22%'],
    ['Adj. Pre-Tax Margin', '—', '~10%+', '+220+ bps'],
    ['Free Cash Flow', '—', '~$2.7B', 'flat'],
    ['Net CapEx', '—', '<$8.0B', '—'],
    ['Net Leverage (target)', '—', '<2.0×', '—'],
  ];
  styledTable(sld, gRows, 0.25, contentY + 0.1, 5.55,
    [2.2, 1.1, 1.15, 1.1], 0.37,
    C.navy, C.white, C.white, C.textDark, 9
  );

  // Bull case
  const bullY = contentY + 0.1 + 0.37 * 7 + 0.18;
  sld.addShape(prs.ShapeType.rect, { x: 0.25, y: bullY, w: 5.55, h: 5.545 - bullY, fill: { color: C.gold }, shadow: makeShadow() });
  sld.addText('Bull Case', { x: 0.4, y: bullY + 0.08, w: 5.25, h: 0.28, fontSize: 10.5, bold: true, color: C.darkSlate, margin: 0 });
  sld.addText('If TRASM inflects positive in Q1–Q2 2026 as industry capacity tightens → tracking toward $14+ EPS → at current multiple = significant equity upside. 2019 EPS peak: $12.05. $14 would be first material all-time high.', {
    x: 0.4, y: bullY + 0.4, w: 5.25, h: 5.545 - bullY - 0.48, fontSize: 8.5, color: C.darkSlate, wrap: true, valign: 'top', margin: 0,
  });

  // Risk matrix
  sld.addShape(prs.ShapeType.rect, { x: 6.0, y: contentY + 0.1, w: 3.75, h: 0.3, fill: { color: C.navy } });
  sld.addText('Risk Matrix', { x: 6.0, y: contentY + 0.1, w: 3.75, h: 0.3, fontSize: 10, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0 });

  const risks = [
    { risk: 'GDP Recession', mag: '–1.5–2.0% rev/–1% GDP', mit: 'Premium/intl mix; loyalty contract-based' },
    { risk: 'Fuel Spike', mag: '$0.10/gal = ~$466M; unhedged', mit: 'New aircraft –20% fuel/seat' },
    { risk: 'Boeing MAX Delays', mag: '~$30–40M/month deferred', mit: 'A321neo/XLR Airbus hedge' },
    { risk: 'Domestic PRASM', mag: 'Margin expansion stalls', mit: 'Mid-year capacity cuts demonstrated' },
    { risk: 'MAX 10 Cert Slip', mag: '200 aircraft waiting', mit: 'Can defer CapEx if needed' },
    { risk: 'Chase Re-contract', mag: '2029 — terms uncertain', mit: 'Delta/AmEx parity gives leverage' },
  ];
  risks.forEach((r, i) => {
    const y = contentY + 0.44 + i * 0.52;
    const bg = i % 2 === 0 ? 'FFF8EC' : C.lightBlue;
    sld.addShape(prs.ShapeType.rect, { x: 6.0, y, w: 3.75, h: 0.48, fill: { color: bg } });
    sld.addText(r.risk, { x: 6.1, y: y + 0.03, w: 1.1, h: 0.22, fontSize: 7.5, bold: true, color: C.navy, margin: 0 });
    sld.addText(r.mag, { x: 6.1, y: y + 0.25, w: 1.1, h: 0.2, fontSize: 7, color: C.textMed, wrap: true, margin: 0 });
    sld.addText(r.mit, { x: 7.24, y: y + 0.06, w: 2.44, h: 0.36, fontSize: 7.5, color: C.textDark, wrap: true, valign: 'middle', margin: 0 });
  });
}

// ════════════════════════════════════════════════════════════
// SLIDE 17 — Appendix: Financial Statements
// ════════════════════════════════════════════════════════════
{
  const sld = prs.addSlide();
  sld.background = { color: C.pageBg };

  sld.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: 10, h: 0.85, fill: { color: C.navy } });
  sld.addText('Appendix: Multi-Year Income Statement', { x: 0.3, y: 0.08, w: 9.4, h: 0.5, fontSize: 20, bold: true, color: C.white, fontFace: 'Calibri', valign: 'middle', margin: 0 });
  sld.addText('FY2023–FY2025 | Figures in billions unless noted', { x: 0.3, y: 0.6, w: 9.4, h: 0.22, fontSize: 9, color: C.goldLight, valign: 'top', margin: 0 });
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
    [3.5, 2.0, 2.0, 2.0], 0.175,
    C.navy, C.white, C.white, C.textDark, 7.2
  );
}

// ════════════════════════════════════════════════════════════
// SLIDE 18 — Appendix: Unit Economics & Fleet
// ════════════════════════════════════════════════════════════
{
  const sld = prs.addSlide();
  sld.background = { color: C.pageBg };

  sld.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: 10, h: 0.85, fill: { color: C.navy } });
  sld.addText('Appendix: Unit Economics & Fleet Reference', { x: 0.3, y: 0.08, w: 9.4, h: 0.5, fontSize: 20, bold: true, color: C.white, fontFace: 'Calibri', valign: 'middle', margin: 0 });
  sld.addText('Annual + Quarterly Unit Economics | Fleet & Workforce | CASM-ex Reconciliation', { x: 0.3, y: 0.6, w: 9.4, h: 0.22, fontSize: 9, color: C.goldLight, valign: 'top', margin: 0 });
  sld.addShape(prs.ShapeType.rect, { x: 0, y: 0.85, w: 10, h: 0.04, fill: { color: C.gold } });

  // Annual unit economics
  sld.addShape(prs.ShapeType.rect, { x: 0.25, y: 0.95, w: 3.0, h: 0.22, fill: { color: C.navyLight } });
  sld.addText('Annual', { x: 0.25, y: 0.95, w: 3.0, h: 0.22, fontSize: 8.5, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0 });

  const annRows = [
    ['Year', 'TRASM', 'PRASM', 'Yield', 'CASM', 'CASM-ex', 'Fuel/Gal', 'Load Factor'],
    ['FY2023', '18.44¢', '—', '20.07¢', '16.99¢', '12.03¢', '$3.01', '83.9%'],
    ['FY2024', '18.34¢', '16.66¢', '20.05¢', '16.70¢', '12.58¢', '$2.65', '83.1%'],
    ['FY2025', '17.88¢', '16.18¢', '19.67¢', '16.46¢', '12.64¢', '$2.44', '82.2%'],
  ];
  // Annual: 4 rows × 0.255 = 1.02", starts at 1.17, ends at 2.19
  styledTable(sld, annRows, 0.25, 1.17, 9.5,
    [0.85, 0.85, 0.85, 0.85, 0.85, 0.9, 0.9, 1.45], 0.255,
    C.navy, C.white, C.white, C.textDark, 7.8
  );

  // Quarterly unit economics — label at 2.24, table at 2.46
  // 5 rows × 0.24 = 1.2, ends at 3.66
  sld.addShape(prs.ShapeType.rect, { x: 0.25, y: 2.24, w: 3.0, h: 0.22, fill: { color: C.navyLight } });
  sld.addText('2025 Quarterly', { x: 0.25, y: 2.24, w: 3.0, h: 0.22, fontSize: 8.5, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0 });

  const qRows = [
    ['Quarter', 'TRASM', 'PRASM', 'Yield', 'CASM', 'CASM-ex', 'Fuel/Gal', 'Load Factor'],
    ['Q1', '17.58¢', '15.78¢', '19.93¢', '16.77¢', '13.17¢', '$2.53', '79.2%'],
    ['Q2', '18.06¢', '16.40¢', '19.74¢', '16.49¢', '12.36¢', '$2.34', '83.1%'],
    ['Q3', '17.42¢', '15.80¢', '18.73¢', '15.82¢', '12.15¢', '$2.43', '84.4%'],
    ['Q4', '18.47¢', '16.71¢', '20.41¢', '16.81¢', '12.94¢', '$2.49', '81.9%'],
  ];
  styledTable(sld, qRows, 0.25, 2.46, 9.5,
    [0.85, 0.85, 0.85, 0.85, 0.85, 0.9, 0.9, 1.45], 0.24,
    C.navyLight, C.white, C.white, C.textDark, 7.8
  );

  // Fleet & workforce — label at 3.72, table at 3.94
  // 6 rows × 0.22 = 1.32, ends at 5.26
  sld.addShape(prs.ShapeType.rect, { x: 0.25, y: 3.72, w: 4.55, h: 0.22, fill: { color: C.navy } });
  sld.addText('Fleet & Workforce', { x: 0.25, y: 3.72, w: 4.55, h: 0.22, fontSize: 8.5, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0 });

  const fleetRows = [
    ['Year', 'Fleet', 'Employees', 'Salaries ($B)'],
    ['2021', '—', '~75,000', '—'],
    ['2022', '~1,300', '~92,000', '—'],
    ['2023', '1,358', '103,300', '$14.787B'],
    ['2024', '1,406', '107,300', '$16.678B'],
    ['2025', '1,490', '113,200', '$17.647B'],
  ];
  styledTable(sld, fleetRows, 0.25, 3.94, 4.55,
    [0.82, 0.82, 1.32, 1.59], 0.22,
    C.navy, C.white, C.white, C.textDark, 8
  );

  // CASM-ex reconciliation — same top as fleet section
  // height = 5.625 - 3.72 - 0.08 = 1.825" (but keep safe at 1.8)
  sld.addShape(prs.ShapeType.rect, { x: 5.05, y: 3.72, w: 4.7, h: 1.82, fill: { color: C.lightBlue } });
  sld.addText('CASM-ex Reconciliation (Q1 2025 example)', { x: 5.17, y: 3.77, w: 4.46, h: 0.24, fontSize: 9, bold: true, color: C.navy, margin: 0 });
  const recon = [
    'CASM: 16.77¢',
    'Less fuel: (3.59)¢',
    'Less profit sharing: (0.06)¢',
    'Less third-party: (0.09)¢',
    'Add back special charges: +0.14¢',
    'CASM-ex: 13.17¢',
  ];
  recon.forEach((r, i) => {
    const bold = i === 0 || i === recon.length - 1;
    sld.addText(r, { x: 5.17, y: 4.04 + i * 0.22, w: 4.46, h: 0.21, fontSize: 8.5, bold, color: bold ? C.navy : C.textDark, margin: 0 });
  });
  sld.addText('Fuel consumption: 4.663B gallons in 2025', { x: 5.17, y: 5.38, w: 4.46, h: 0.16, fontSize: 7.5, color: C.textLight, italic: true, margin: 0 });
}

// ─── Save ────────────────────────────────────────────────────────────────────
prs.writeFile({ fileName: 'outputs/corporate-qa.pptx' })
  .then(() => console.log('Saved: outputs/corporate-qa.pptx'))
  .catch(err => { console.error(err); process.exit(1); });
