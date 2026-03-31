const pptxgen = require("pptxgenjs");
const path = require("path");

// ============================================================
// NovaCrest Q3 2026 Strategic Review — Board Deck
// Color Palette: Midnight Executive (navy + ice blue + white)
// ============================================================

const COLORS = {
  navy: "1E2761",
  navyDark: "151C47",
  ice: "CADCFC",
  iceMuted: "E8EFFF",
  white: "FFFFFF",
  offWhite: "F6F8FC",
  charcoal: "2D3142",
  gray: "6B7280",
  grayLight: "9CA3AF",
  grayBorder: "D1D5DB",
  green: "059669",
  greenBg: "ECFDF5",
  red: "DC2626",
  redBg: "FEF2F2",
  amber: "D97706",
  amberBg: "FFFBEB",
  accent: "3B82F6",
  accentDark: "1D4ED8",
};

const FONTS = {
  title: "Georgia",
  body: "Calibri",
};

// Slide dimensions: 10" x 5.625" (16:9)
const SW = 10;
const SH = 5.625;

// Factory functions for reusable shadow/style objects (avoid mutation issue)
const makeCardShadow = () => ({
  type: "outer",
  color: "000000",
  blur: 4,
  offset: 1,
  angle: 135,
  opacity: 0.08,
});

const makeSubtleShadow = () => ({
  type: "outer",
  color: "000000",
  blur: 2,
  offset: 1,
  angle: 135,
  opacity: 0.05,
});

// ============================================================
// HELPER: KPI Card
// ============================================================
function addKPICard(slide, x, y, w, h, label, value, subtext, subtextColor) {
  // Card background
  slide.addShape("rect", {
    x,
    y,
    w,
    h,
    fill: { color: COLORS.white },
    shadow: makeCardShadow(),
  });
  // Left accent bar
  slide.addShape("rect", {
    x,
    y,
    w: 0.06,
    h,
    fill: { color: COLORS.navy },
  });
  // Label
  slide.addText(label.toUpperCase(), {
    x: x + 0.18,
    y: y + 0.08,
    w: w - 0.3,
    h: 0.28,
    fontSize: 8,
    fontFace: FONTS.body,
    color: COLORS.gray,
    bold: true,
    charSpacing: 1.5,
    margin: 0,
  });
  // Value
  slide.addText(value, {
    x: x + 0.18,
    y: y + 0.32,
    w: w - 0.3,
    h: 0.38,
    fontSize: 22,
    fontFace: FONTS.title,
    color: COLORS.charcoal,
    bold: true,
    margin: 0,
  });
  // Subtext
  if (subtext) {
    slide.addText(subtext, {
      x: x + 0.18,
      y: y + 0.68,
      w: w - 0.3,
      h: 0.22,
      fontSize: 9,
      fontFace: FONTS.body,
      color: subtextColor || COLORS.green,
      margin: 0,
    });
  }
}

// ============================================================
// HELPER: Section header bar
// ============================================================
function addSectionHeader(slide, x, y, w, text) {
  slide.addShape("rect", {
    x,
    y,
    w,
    h: 0.32,
    fill: { color: COLORS.navy },
  });
  slide.addText(text.toUpperCase(), {
    x: x + 0.12,
    y,
    w: w - 0.24,
    h: 0.32,
    fontSize: 9,
    fontFace: FONTS.body,
    color: COLORS.white,
    bold: true,
    charSpacing: 2,
    valign: "middle",
    margin: 0,
  });
}

// ============================================================
// HELPER: add a data table
// ============================================================
function addDataTable(slide, x, y, w, headers, rows, opts = {}) {
  const colW = opts.colW || headers.map(() => w / headers.length);
  const headerRow = headers.map((h) => ({
    text: h,
    options: {
      bold: true,
      fontSize: 8,
      fontFace: FONTS.body,
      color: COLORS.white,
      fill: { color: COLORS.navy },
      align: "left",
      valign: "middle",
    },
  }));
  const dataRows = rows.map((row, ri) =>
    row.map((cell, ci) => {
      const isString = typeof cell === "string";
      const cellText = isString ? cell : cell.text;
      const cellOpts = isString ? {} : cell.options || {};
      return {
        text: cellText,
        options: {
          fontSize: 8,
          fontFace: FONTS.body,
          color: cellOpts.color || COLORS.charcoal,
          fill: {
            color:
              cellOpts.fill || (ri % 2 === 0 ? COLORS.offWhite : COLORS.white),
          },
          align: cellOpts.align || (ci === 0 ? "left" : "center"),
          valign: "middle",
          bold: cellOpts.bold || false,
          ...(cellOpts.extra || {}),
        },
      };
    })
  );

  slide.addTable([headerRow, ...dataRows], {
    x,
    y,
    w,
    colW,
    rowH: opts.rowH || 0.26,
    border: { pt: 0.5, color: COLORS.grayBorder },
    autoPage: false,
  });
}

// ============================================================
// BUILD PRESENTATION
// ============================================================
async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "NovaCrest";
  pres.title = "Q3 2026 Strategic Review";

  // ===========================================================
  // SLIDE 1: Title + Executive Summary
  // ===========================================================
  {
    const slide = pres.addSlide();
    slide.background = { color: COLORS.navyDark };

    // Top decorative bar
    slide.addShape("rect", {
      x: 0,
      y: 0,
      w: SW,
      h: 0.06,
      fill: { color: COLORS.accent },
    });

    // Company name
    slide.addText("NOVACREST", {
      x: 0.6,
      y: 0.35,
      w: 4,
      h: 0.35,
      fontSize: 12,
      fontFace: FONTS.body,
      color: COLORS.ice,
      charSpacing: 6,
      bold: true,
      margin: 0,
    });

    // Title
    slide.addText("Q3 2026 Strategic Review", {
      x: 0.6,
      y: 0.72,
      w: 6,
      h: 0.65,
      fontSize: 32,
      fontFace: FONTS.title,
      color: COLORS.white,
      bold: true,
      margin: 0,
    });

    // Subtitle
    slide.addText("Board of Directors  |  September 2026  |  Confidential", {
      x: 0.6,
      y: 1.35,
      w: 6,
      h: 0.3,
      fontSize: 11,
      fontFace: FONTS.body,
      color: COLORS.grayLight,
      margin: 0,
    });

    // Divider line
    slide.addShape("rect", {
      x: 0.6,
      y: 1.82,
      w: 8.8,
      h: 0.015,
      fill: { color: COLORS.accent },
    });

    // === Three-column executive summary ===
    const colX1 = 0.6;
    const colX2 = 3.65;
    const colX3 = 6.7;
    const colW = 2.75;
    const topY = 2.05;

    // Key Wins
    slide.addShape("rect", {
      x: colX1,
      y: topY,
      w: colW,
      h: 2.6,
      fill: { color: COLORS.navy },
    });
    slide.addText("KEY WINS", {
      x: colX1 + 0.15,
      y: topY + 0.08,
      w: colW - 0.3,
      h: 0.25,
      fontSize: 9,
      fontFace: FONTS.body,
      color: COLORS.green,
      bold: true,
      charSpacing: 2,
      margin: 0,
    });
    slide.addText(
      [
        {
          text: "ARR +10.6% QoQ to $15.6M",
          options: {
            bullet: true,
            breakLine: true,
            fontSize: 10,
            color: COLORS.ice,
          },
        },
        {
          text: "+4.2% ahead of plan",
          options: {
            bullet: true,
            indentLevel: 1,
            breakLine: true,
            fontSize: 9,
            color: COLORS.grayLight,
          },
        },
        {
          text: "Pipeline best quarter ever: $8.6M",
          options: {
            bullet: true,
            breakLine: true,
            fontSize: 10,
            color: COLORS.ice,
          },
        },
        {
          text: "+21% QoQ, outbound engine scaling",
          options: {
            bullet: true,
            indentLevel: 1,
            breakLine: true,
            fontSize: 9,
            color: COLORS.grayLight,
          },
        },
        {
          text: "Q3 hiring strong: 118 headcount",
          options: {
            bullet: true,
            breakLine: true,
            fontSize: 10,
            color: COLORS.ice,
          },
        },
        {
          text: "VP CS + Head of Partnerships hired",
          options: {
            bullet: true,
            indentLevel: 1,
            breakLine: true,
            fontSize: 9,
            color: COLORS.grayLight,
          },
        },
        {
          text: "Magic Number 1.12x — sales efficiency strong",
          options: { bullet: true, fontSize: 10, color: COLORS.ice },
        },
      ],
      {
        x: colX1 + 0.15,
        y: topY + 0.38,
        w: colW - 0.3,
        h: 2.15,
        fontFace: FONTS.body,
        paraSpaceAfter: 4,
        valign: "top",
        margin: 0,
      }
    );

    // Concerns
    slide.addShape("rect", {
      x: colX2,
      y: topY,
      w: colW,
      h: 2.6,
      fill: { color: COLORS.navy },
    });
    slide.addText("KEY CONCERNS", {
      x: colX2 + 0.15,
      y: topY + 0.08,
      w: colW - 0.3,
      h: 0.25,
      fontSize: 9,
      fontFace: FONTS.body,
      color: COLORS.red,
      bold: true,
      charSpacing: 2,
      margin: 0,
    });
    slide.addText(
      [
        {
          text: "Churn spike: $500K (+67% QoQ)",
          options: {
            bullet: true,
            breakLine: true,
            fontSize: 10,
            color: COLORS.ice,
          },
        },
        {
          text: "58% was preventable — onboarding + retention gaps",
          options: {
            bullet: true,
            indentLevel: 1,
            breakLine: true,
            fontSize: 9,
            color: COLORS.grayLight,
          },
        },
        {
          text: "NRR declining: 118% (from 121%)",
          options: {
            bullet: true,
            breakLine: true,
            fontSize: 10,
            color: COLORS.ice,
          },
        },
        {
          text: "Logo retention 92.4% vs 95% target",
          options: {
            bullet: true,
            indentLevel: 1,
            breakLine: true,
            fontSize: 9,
            color: COLORS.grayLight,
          },
        },
        {
          text: "Competitive pressure increasing",
          options: {
            bullet: true,
            breakLine: true,
            fontSize: 10,
            color: COLORS.ice,
          },
        },
        {
          text: "DataForge mfg module, Acme Lite, Zenith poaching",
          options: {
            bullet: true,
            indentLevel: 1,
            fontSize: 9,
            color: COLORS.grayLight,
          },
        },
      ],
      {
        x: colX2 + 0.15,
        y: topY + 0.38,
        w: colW - 0.3,
        h: 2.15,
        fontFace: FONTS.body,
        paraSpaceAfter: 4,
        valign: "top",
        margin: 0,
      }
    );

    // Decision Required
    slide.addShape("rect", {
      x: colX3,
      y: topY,
      w: colW,
      h: 2.6,
      fill: { color: COLORS.navy },
    });
    slide.addText("DECISION REQUIRED", {
      x: colX3 + 0.15,
      y: topY + 0.08,
      w: colW - 0.3,
      h: 0.25,
      fontSize: 9,
      fontFace: FONTS.body,
      color: COLORS.amber,
      bold: true,
      charSpacing: 2,
      margin: 0,
    });
    slide.addText(
      [
        {
          text: "$2.5M CS Investment",
          options: {
            bullet: true,
            breakLine: true,
            fontSize: 10,
            color: COLORS.ice,
            bold: true,
          },
        },
        {
          text: "8 FTEs to reduce preventable churn 60%",
          options: {
            bullet: true,
            indentLevel: 1,
            breakLine: true,
            fontSize: 9,
            color: COLORS.grayLight,
          },
        },
        {
          text: "Save ~$700K ARR/yr; 14-mo payback",
          options: {
            bullet: true,
            indentLevel: 1,
            breakLine: true,
            fontSize: 9,
            color: COLORS.grayLight,
          },
        },
        {
          text: "SMB Pricing Tier ($499/mo)",
          options: {
            bullet: true,
            breakLine: true,
            fontSize: 10,
            color: COLORS.ice,
            bold: true,
          },
        },
        {
          text: "+$1.2M new ARR in 12 months",
          options: {
            bullet: true,
            indentLevel: 1,
            breakLine: true,
            fontSize: 9,
            color: COLORS.grayLight,
          },
        },
        {
          text: "Series C Timing Discussion",
          options: {
            bullet: true,
            breakLine: true,
            fontSize: 10,
            color: COLORS.ice,
            bold: true,
          },
        },
        {
          text: "Q2 2027 at $22-25M ARR vs wait for $30M?",
          options: {
            bullet: true,
            indentLevel: 1,
            fontSize: 9,
            color: COLORS.grayLight,
          },
        },
      ],
      {
        x: colX3 + 0.15,
        y: topY + 0.38,
        w: colW - 0.3,
        h: 2.15,
        fontFace: FONTS.body,
        paraSpaceAfter: 4,
        valign: "top",
        margin: 0,
      }
    );

    // Bottom info bar
    slide.addShape("rect", {
      x: 0,
      y: SH - 0.55,
      w: SW,
      h: 0.55,
      fill: { color: COLORS.navy },
    });
    slide.addText(
      "B2B SaaS  |  Predictive Analytics for Manufacturing  |  Series B ($32M)  |  118 Employees  |  $15.6M ARR",
      {
        x: 0.6,
        y: SH - 0.55,
        w: 8.8,
        h: 0.55,
        fontSize: 9,
        fontFace: FONTS.body,
        color: COLORS.grayLight,
        valign: "middle",
        margin: 0,
      }
    );
  }

  // ===========================================================
  // SLIDE 2: Revenue Dashboard
  // ===========================================================
  {
    const slide = pres.addSlide();
    slide.background = { color: COLORS.offWhite };

    // Header bar
    slide.addShape("rect", {
      x: 0,
      y: 0,
      w: SW,
      h: 0.6,
      fill: { color: COLORS.navy },
    });
    slide.addText("Revenue Dashboard", {
      x: 0.5,
      y: 0,
      w: 5,
      h: 0.6,
      fontSize: 20,
      fontFace: FONTS.title,
      color: COLORS.white,
      bold: true,
      valign: "middle",
      margin: 0,
    });
    slide.addText("Q3 2026", {
      x: 7.5,
      y: 0,
      w: 2,
      h: 0.6,
      fontSize: 12,
      fontFace: FONTS.body,
      color: COLORS.ice,
      align: "right",
      valign: "middle",
      margin: 0,
    });

    // KPI Cards Row
    const kpiY = 0.78;
    const kpiH = 0.95;
    const kpiW = 2.1;
    const kpiGap = 0.2;
    const kpiX0 = 0.5;

    addKPICard(
      slide,
      kpiX0,
      kpiY,
      kpiW,
      kpiH,
      "ARR",
      "$15.6M",
      "+10.6% QoQ  |  +4.2% vs plan",
      COLORS.green
    );
    addKPICard(
      slide,
      kpiX0 + kpiW + kpiGap,
      kpiY,
      kpiW,
      kpiH,
      "Net Revenue Retention",
      "118%",
      "Target 120%+  |  Declining",
      COLORS.amber
    );
    addKPICard(
      slide,
      kpiX0 + 2 * (kpiW + kpiGap),
      kpiY,
      kpiW,
      kpiH,
      "Gross Margin",
      "78.2%",
      "Target 80%+  |  Improving",
      COLORS.green
    );
    addKPICard(
      slide,
      kpiX0 + 3 * (kpiW + kpiGap),
      kpiY,
      kpiW,
      kpiH,
      "Magic Number",
      "1.12x",
      "Target >0.75x  |  Strong",
      COLORS.green
    );

    // ARR Growth Chart (left)
    const chartY = 1.92;
    addSectionHeader(slide, 0.5, chartY, 4.2, "ARR Growth Trend ($M)");

    slide.addChart(
      pres.charts.BAR,
      [
        {
          name: "ARR",
          labels: [
            "Q1'25",
            "Q2'25",
            "Q3'25",
            "Q4'25",
            "Q1'26",
            "Q2'26",
            "Q3'26",
          ],
          values: [7.2, 8.5, 9.8, 11.4, 12.8, 14.1, 15.6],
        },
      ],
      {
        x: 0.5,
        y: chartY + 0.32,
        w: 4.2,
        h: 2.55,
        barDir: "col",
        chartColors: [COLORS.navy],
        chartArea: { fill: { color: COLORS.white }, roundedCorners: false },
        catAxisLabelColor: COLORS.gray,
        catAxisLabelFontSize: 7,
        valAxisLabelColor: COLORS.gray,
        valAxisLabelFontSize: 7,
        valGridLine: { color: "E2E8F0", size: 0.5 },
        catGridLine: { style: "none" },
        showValue: true,
        dataLabelPosition: "outEnd",
        dataLabelColor: COLORS.charcoal,
        dataLabelFontSize: 7,
        showLegend: false,
      }
    );

    // Revenue & Key Metrics Table (right)
    addSectionHeader(slide, 5.1, chartY, 4.4, "Revenue & Key Metrics");

    addDataTable(
      slide,
      5.1,
      chartY + 0.32,
      4.4,
      ["Metric", "Q1", "Q2", "Q3", "QoQ", "vs Plan"],
      [
        ["ARR", "$12.8M", "$14.1M", "$15.6M", "+10.6%", "+4.2%"],
        ["New ARR", "$1.4M", "$1.6M", "$1.9M", "+18.8%", "+12%"],
        ["Expansion ARR", "$0.6M", "$0.7M", "$0.8M", "+14.3%", "On plan"],
        [
          {
            text: "Churned ARR",
            options: { color: COLORS.red },
          },
          "($0.3M)",
          "($0.3M)",
          {
            text: "($0.5M)",
            options: { color: COLORS.red, bold: true },
          },
          {
            text: "+66.7%",
            options: { color: COLORS.red },
          },
          {
            text: "Behind",
            options: { color: COLORS.red },
          },
        ],
        ["Net New ARR", "$1.7M", "$2.0M", "$2.2M", "+10.0%", "+6%"],
        ["Gross Revenue", "$3.4M", "$3.7M", "$4.1M", "+10.8%", "+3.8%"],
        ["Subscription", "$3.1M", "$3.4M", "$3.8M", "+11.8%", "+4.1%"],
        ["Services", "$0.3M", "$0.3M", "$0.3M", "Flat", "On plan"],
      ],
      { colW: [0.95, 0.62, 0.62, 0.62, 0.7, 0.89], rowH: 0.24 }
    );

    // Unit Economics row
    addSectionHeader(slide, 0.5, 4.87, 9.0, "Unit Economics");

    // Inline unit economics as a compact row of mini metrics
    const ueY = 5.19;
    const ueData = [
      {
        label: "LTV:CAC",
        val: "4.8x",
        note: "Target >4.0x",
        color: COLORS.green,
      },
      {
        label: "CAC Payback",
        val: "13.2mo",
        note: "Target <16mo",
        color: COLORS.green,
      },
      {
        label: "Logo Retention",
        val: "92.4%",
        note: "Target 95%+",
        color: COLORS.red,
      },
      {
        label: "Gross Margin",
        val: "78.2%",
        note: "Target 80%+",
        color: COLORS.amber,
      },
    ];
    const ueW = 9.0 / ueData.length;
    ueData.forEach((d, i) => {
      const ux = 0.5 + i * ueW;
      slide.addText(
        [
          {
            text: d.label + ": ",
            options: { fontSize: 9, color: COLORS.gray, bold: true },
          },
          {
            text: d.val + "  ",
            options: { fontSize: 11, color: COLORS.charcoal, bold: true },
          },
          {
            text: d.note,
            options: { fontSize: 8, color: d.color },
          },
        ],
        {
          x: ux,
          y: ueY,
          w: ueW,
          h: 0.32,
          fontFace: FONTS.body,
          valign: "middle",
          margin: 0,
        }
      );
    });
  }

  // ===========================================================
  // SLIDE 3: GTM + Customer Health
  // ===========================================================
  {
    const slide = pres.addSlide();
    slide.background = { color: COLORS.offWhite };

    // Header
    slide.addShape("rect", {
      x: 0,
      y: 0,
      w: SW,
      h: 0.6,
      fill: { color: COLORS.navy },
    });
    slide.addText("Go-to-Market & Customer Health", {
      x: 0.5,
      y: 0,
      w: 6,
      h: 0.6,
      fontSize: 20,
      fontFace: FONTS.title,
      color: COLORS.white,
      bold: true,
      valign: "middle",
      margin: 0,
    });

    // GTM KPI cards
    const gkY = 0.78;
    const gkW = 1.72;
    const gkH = 0.78;
    const gkGap = 0.15;
    const gkX0 = 0.5;
    const gkData = [
      {
        label: "Pipeline",
        val: "$8.6M",
        sub: "+21% QoQ",
        color: COLORS.green,
      },
      {
        label: "Win Rate",
        val: "28.4%",
        sub: "+2.3pp QoQ",
        color: COLORS.green,
      },
      {
        label: "Avg ACV",
        val: "$42K",
        sub: "+10.5% QoQ",
        color: COLORS.green,
      },
      {
        label: "Sales Cycle",
        val: "68 days",
        sub: "-5.6% QoQ",
        color: COLORS.green,
      },
      {
        label: "Quota Attain.",
        val: "112%",
        sub: "+18pp QoQ",
        color: COLORS.green,
      },
    ];
    gkData.forEach((d, i) => {
      const gx = gkX0 + i * (gkW + gkGap);
      addKPICard(slide, gx, gkY, gkW, gkH, d.label, d.val, d.sub, d.color);
    });

    // Channel Performance Table (left)
    const secY = 1.72;
    addSectionHeader(slide, 0.5, secY, 4.8, "Channel Performance");
    addDataTable(
      slide,
      0.5,
      secY + 0.32,
      4.8,
      ["Channel", "Pipeline", "Deals", "Avg ACV", "CAC"],
      [
        ["Outbound SDR", "$3.8M", "18", "$52K", "$24.1K"],
        ["Inbound Marketing", "$2.4M", "14", "$34K", "$12.8K"],
        ["PLG / Self-Serve", "$0.9M", "8", "$18K", "$6.2K"],
        ["Partner / Referral", "$1.5M", "5", "$68K", "$8.4K"],
      ],
      { colW: [1.2, 0.9, 0.7, 0.9, 1.1], rowH: 0.24 }
    );

    // CAC Trend chart (right)
    addSectionHeader(slide, 5.6, secY, 3.9, "CAC Trend ($K)");
    slide.addChart(
      pres.charts.LINE,
      [
        {
          name: "Blended",
          labels: [
            "Q1'25",
            "Q2'25",
            "Q3'25",
            "Q4'25",
            "Q1'26",
            "Q2'26",
            "Q3'26",
          ],
          values: [22.4, 20.1, 19.8, 18.2, 17.6, 16.4, 15.8],
        },
        {
          name: "Enterprise",
          labels: [
            "Q1'25",
            "Q2'25",
            "Q3'25",
            "Q4'25",
            "Q1'26",
            "Q2'26",
            "Q3'26",
          ],
          values: [38.2, 34.6, 32.1, 28.4, 26.8, 24.2, 22.6],
        },
      ],
      {
        x: 5.6,
        y: secY + 0.32,
        w: 3.9,
        h: 1.58,
        lineSmooth: true,
        lineSize: 2,
        chartColors: [COLORS.navy, COLORS.accent],
        chartArea: { fill: { color: COLORS.white } },
        catAxisLabelColor: COLORS.gray,
        catAxisLabelFontSize: 6,
        valAxisLabelColor: COLORS.gray,
        valAxisLabelFontSize: 6,
        valGridLine: { color: "E2E8F0", size: 0.5 },
        catGridLine: { style: "none" },
        showLegend: true,
        legendPos: "b",
        legendFontSize: 7,
      }
    );

    // Churn Deep Dive (bottom left)
    const churnY = 3.28;
    addSectionHeader(
      slide,
      0.5,
      churnY,
      4.8,
      "Churn Deep Dive — Q3 ($500K Total)"
    );
    addDataTable(
      slide,
      0.5,
      churnY + 0.32,
      4.8,
      ["Account", "ARR Lost", "Reason", "Preventable?"],
      [
        ["Apex Mfg", "$145K", "Post-M&A non-renewal", "No"],
        ["Precision Dynamics", "$62K", "Switched to in-house", "Partial"],
        [
          "4 SMB accounts",
          "$118K",
          "Price; competitor",
          { text: "Yes", options: { color: COLORS.red, bold: true } },
        ],
        [
          "TechFab Solutions",
          "$48K",
          "Poor implementation",
          { text: "Yes", options: { color: COLORS.red, bold: true } },
        ],
        [
          "3 SMB accounts",
          "$82K",
          "Low usage; not onboarded",
          { text: "Yes", options: { color: COLORS.red, bold: true } },
        ],
        ["Consolidated Parts", "$45K", "Budget cuts", "No"],
      ],
      { colW: [1.1, 0.7, 1.6, 1.4], rowH: 0.22 }
    );

    // At-Risk Accounts (bottom right)
    addSectionHeader(slide, 5.6, churnY, 3.9, "At-Risk Accounts — Q4");
    addDataTable(
      slide,
      5.6,
      churnY + 0.32,
      3.9,
      ["Account", "ARR", "Risk Signal"],
      [
        [
          { text: "Sterling Industries", options: { bold: true } },
          "$210K",
          "Champion left",
        ],
        [
          { text: "Midwest Components", options: { bold: true } },
          "$58K",
          "Usage -40%",
        ],
        [
          { text: "ClearPath Systems", options: { bold: true } },
          "$72K",
          "Competitor POC",
        ],
      ],
      { colW: [1.4, 0.7, 1.8], rowH: 0.24 }
    );

    // Churn summary callout
    const calloutY = churnY + 0.32 + 0.24 * 4 + 0.08;
    slide.addShape("rect", {
      x: 5.6,
      y: calloutY,
      w: 3.9,
      h: 0.42,
      fill: { color: COLORS.redBg },
    });
    slide.addText(
      "$340K ARR at risk in Q4. $290K of Q3 churn (58%) was preventable.",
      {
        x: 5.72,
        y: calloutY,
        w: 3.66,
        h: 0.42,
        fontSize: 8,
        fontFace: FONTS.body,
        color: COLORS.red,
        bold: true,
        valign: "middle",
        margin: 0,
      }
    );
  }

  // ===========================================================
  // SLIDE 4: Product + Team
  // ===========================================================
  {
    const slide = pres.addSlide();
    slide.background = { color: COLORS.offWhite };

    // Header
    slide.addShape("rect", {
      x: 0,
      y: 0,
      w: SW,
      h: 0.6,
      fill: { color: COLORS.navy },
    });
    slide.addText("Product & Engineering  |  Team & Organization", {
      x: 0.5,
      y: 0,
      w: 7,
      h: 0.6,
      fontSize: 20,
      fontFace: FONTS.title,
      color: COLORS.white,
      bold: true,
      valign: "middle",
      margin: 0,
    });

    // Shipped in Q3 (top left)
    const topY = 0.78;
    addSectionHeader(slide, 0.5, topY, 5.6, "Shipped in Q3");
    addDataTable(
      slide,
      0.5,
      topY + 0.32,
      5.6,
      ["Feature", "Impact", "Adoption", "Revenue"],
      [
        [
          "Predictive Maint. v2",
          "+23% accuracy",
          "67% enterprise",
          "$1.2M pipeline",
        ],
        [
          "Dashboard Builder",
          "Self-serve dashboards",
          "340 dashboards/89 accts",
          "-18% CS tickets",
        ],
        [
          "SAP Integration",
          "Native connector",
          "12 connected",
          "Partner accelerant",
        ],
        ["SOC 2 Type II", "Audit complete 8/15", "N/A", "Unblocked $380K"],
      ],
      { colW: [1.35, 1.35, 1.55, 1.35], rowH: 0.24 }
    );

    // Q4 Roadmap (top right)
    addSectionHeader(slide, 6.35, topY, 3.15, "Q4 Roadmap");
    addDataTable(
      slide,
      6.35,
      topY + 0.32,
      3.15,
      ["P", "Feature", "Status"],
      [
        [
          { text: "P0", options: { color: COLORS.red, bold: true } },
          "Multi-tenant analytics",
          "60% done",
        ],
        [
          { text: "P0", options: { color: COLORS.red, bold: true } },
          "Siemens integration",
          "Design done",
        ],
        [
          { text: "P1", options: { color: COLORS.amber, bold: true } },
          "Usage-based pricing",
          "Spec done",
        ],
        [
          { text: "P1", options: { color: COLORS.amber, bold: true } },
          "Health score tool",
          "Prototype",
        ],
        [
          { text: "P2", options: { color: COLORS.gray } },
          "Mobile app",
          "Scoping",
        ],
      ],
      { colW: [0.35, 1.6, 1.2], rowH: 0.22 }
    );

    // Headcount by Function (bottom left)
    const botY = 2.2;
    addSectionHeader(slide, 0.5, botY, 4.6, "Headcount by Function");

    // Stacked bar chart for headcount
    slide.addChart(
      pres.charts.BAR,
      [
        {
          name: "Engineering",
          labels: ["Q2 2026", "Q3 2026", "Q4 Target"],
          values: [42, 46, 50],
        },
        {
          name: "Sales",
          labels: ["Q2 2026", "Q3 2026", "Q4 Target"],
          values: [18, 22, 25],
        },
        {
          name: "CS",
          labels: ["Q2 2026", "Q3 2026", "Q4 Target"],
          values: [12, 14, 16],
        },
        {
          name: "Product+Design",
          labels: ["Q2 2026", "Q3 2026", "Q4 Target"],
          values: [8, 9, 10],
        },
        {
          name: "Marketing",
          labels: ["Q2 2026", "Q3 2026", "Q4 Target"],
          values: [8, 9, 10],
        },
        {
          name: "G&A",
          labels: ["Q2 2026", "Q3 2026", "Q4 Target"],
          values: [10, 12, 13],
        },
      ],
      {
        x: 0.5,
        y: botY + 0.32,
        w: 4.6,
        h: 2.5,
        barDir: "col",
        barGrouping: "stacked",
        chartColors: [
          COLORS.navy,
          COLORS.accent,
          "10B981",
          "8B5CF6",
          COLORS.amber,
          COLORS.grayLight,
        ],
        chartArea: { fill: { color: COLORS.white } },
        catAxisLabelColor: COLORS.gray,
        catAxisLabelFontSize: 8,
        valAxisLabelColor: COLORS.gray,
        valAxisLabelFontSize: 7,
        valGridLine: { color: "E2E8F0", size: 0.5 },
        catGridLine: { style: "none" },
        showLegend: true,
        legendPos: "b",
        legendFontSize: 6,
        showValue: false,
      }
    );

    // Key Hires + Org Notes (bottom right)
    addSectionHeader(slide, 5.35, botY, 4.15, "Key Hires & Org Notes");

    slide.addShape("rect", {
      x: 5.35,
      y: botY + 0.32,
      w: 4.15,
      h: 2.5,
      fill: { color: COLORS.white },
      shadow: makeSubtleShadow(),
    });

    slide.addText(
      [
        {
          text: "Key Q3 Hires",
          options: {
            bold: true,
            fontSize: 10,
            color: COLORS.navy,
            breakLine: true,
          },
        },
        {
          text: "VP Customer Success — ex-Datadog, built CS 20\u219280",
          options: { bullet: true, breakLine: true, fontSize: 9 },
        },
        {
          text: "Head of Partnerships — ex-Siemens, SAP/Siemens channel",
          options: { bullet: true, breakLine: true, fontSize: 9 },
        },
        {
          text: "2 Senior ML Engineers — PhD hires, Georgia Tech",
          options: { bullet: true, breakLine: true, fontSize: 9 },
        },
        {
          text: "",
          options: { breakLine: true, fontSize: 6 },
        },
        {
          text: "Org Health",
          options: {
            bold: true,
            fontSize: 10,
            color: COLORS.navy,
            breakLine: true,
          },
        },
        {
          text: "112 headcount (+14 in Q3); 12 open roles",
          options: { bullet: true, breakLine: true, fontSize: 9 },
        },
        {
          text: "5.4% quarterly attrition (6 departed)",
          options: { bullet: true, breakLine: true, fontSize: 9 },
        },
        {
          text: "3 eng departures to FAANG — adjusted comp +12%",
          options: {
            bullet: true,
            breakLine: true,
            fontSize: 9,
            color: COLORS.amber,
          },
        },
        {
          text: "Engineering: 46 targeting 50 by Q4",
          options: { bullet: true, fontSize: 9 },
        },
      ],
      {
        x: 5.5,
        y: botY + 0.4,
        w: 3.85,
        h: 2.3,
        fontFace: FONTS.body,
        color: COLORS.charcoal,
        paraSpaceAfter: 3,
        valign: "top",
        margin: 0,
      }
    );
  }

  // ===========================================================
  // SLIDE 5: Financial Outlook + Competitive
  // ===========================================================
  {
    const slide = pres.addSlide();
    slide.background = { color: COLORS.offWhite };

    // Header
    slide.addShape("rect", {
      x: 0,
      y: 0,
      w: SW,
      h: 0.6,
      fill: { color: COLORS.navy },
    });
    slide.addText("Financial Outlook & Competitive Landscape", {
      x: 0.5,
      y: 0,
      w: 7,
      h: 0.6,
      fontSize: 20,
      fontFace: FONTS.title,
      color: COLORS.white,
      bold: true,
      valign: "middle",
      margin: 0,
    });

    // P&L Table (left)
    const topY = 0.78;
    addSectionHeader(slide, 0.5, topY, 5.5, "P&L Summary ($K)");
    addDataTable(
      slide,
      0.5,
      topY + 0.32,
      5.5,
      ["Line Item", "Q1'26", "Q2'26", "Q3'26", "Q3 YoY", "FY2026E"],
      [
        ["Revenue", "$3,400", "$3,700", "$4,100", "+64%", "$15,400"],
        ["COGS", "($740)", "($830)", "($890)", "+52%", "($3,340)"],
        [
          { text: "Gross Profit", options: { bold: true } },
          { text: "$2,660", options: { bold: true } },
          { text: "$2,870", options: { bold: true } },
          { text: "$3,210", options: { bold: true } },
          { text: "+68%", options: { bold: true } },
          { text: "$12,060", options: { bold: true } },
        ],
        ["Gross Margin", "78.2%", "77.6%", "78.3%", "+2.1pp", "78.3%"],
        ["S&M", "($1,420)", "($1,580)", "($1,720)", "+48%", "($6,480)"],
        ["R&D", "($1,340)", "($1,480)", "($1,620)", "+55%", "($6,040)"],
        ["G&A", "($480)", "($510)", "($540)", "+38%", "($2,060)"],
        [
          { text: "Net Income", options: { bold: true } },
          { text: "($580)", options: { bold: true, color: COLORS.red } },
          { text: "($700)", options: { bold: true, color: COLORS.red } },
          { text: "($670)", options: { bold: true, color: COLORS.red } },
          { text: "Improved", options: { bold: true, color: COLORS.green } },
          { text: "($2,520)", options: { bold: true, color: COLORS.red } },
        ],
      ],
      { colW: [1.05, 0.82, 0.82, 0.82, 0.82, 1.17], rowH: 0.22 }
    );

    // Cash Position (right top)
    addSectionHeader(slide, 6.25, topY, 3.25, "Cash Position");
    slide.addShape("rect", {
      x: 6.25,
      y: topY + 0.32,
      w: 3.25,
      h: 1.76,
      fill: { color: COLORS.white },
      shadow: makeSubtleShadow(),
    });

    const cashData = [
      { label: "Cash on Hand", val: "$18.2M" },
      { label: "Monthly Burn", val: "$223K" },
      { label: "Runway", val: "81 months" },
      { label: "Burn Multiple", val: "0.36x" },
      { label: "Last Funding", val: "Series B, $32M" },
    ];
    cashData.forEach((d, i) => {
      const cy = topY + 0.38 + i * 0.32;
      slide.addText(d.label, {
        x: 6.38,
        y: cy,
        w: 1.5,
        h: 0.28,
        fontSize: 9,
        fontFace: FONTS.body,
        color: COLORS.gray,
        valign: "middle",
        margin: 0,
      });
      slide.addText(d.val, {
        x: 7.88,
        y: cy,
        w: 1.5,
        h: 0.28,
        fontSize: 10,
        fontFace: FONTS.body,
        color: COLORS.charcoal,
        bold: true,
        align: "right",
        valign: "middle",
        margin: 0,
      });
    });

    // Cash status indicator
    const cashBottomY = topY + 0.32 + 1.76 + 0.06;
    slide.addShape("rect", {
      x: 6.25,
      y: cashBottomY,
      w: 3.25,
      h: 0.28,
      fill: { color: COLORS.greenBg },
    });
    slide.addText("Runway comfortable — no immediate fundraise needed", {
      x: 6.38,
      y: cashBottomY,
      w: 3.0,
      h: 0.28,
      fontSize: 8,
      fontFace: FONTS.body,
      color: COLORS.green,
      bold: true,
      valign: "middle",
      margin: 0,
    });

    // Competitive Landscape (full width bottom)
    const compY = 3.05;
    addSectionHeader(slide, 0.5, compY, 9.0, "Competitive Landscape");
    addDataTable(
      slide,
      0.5,
      compY + 0.32,
      9.0,
      [
        "Dimension",
        "NovaCrest (Us)",
        "DataForge",
        "Acme Analytics",
        "Zenith AI",
      ],
      [
        [
          "Stage",
          { text: "Series B ($32M)", options: { bold: true } },
          "Series C ($85M)",
          "Public ($2.1B)",
          "Series A ($18M)",
        ],
        [
          "ARR",
          { text: "$15.6M", options: { bold: true } },
          "~$45M",
          "~$180M",
          "~$4M",
        ],
        [
          "Target",
          "Mid-mkt mfg",
          "SMB-Mid (horiz.)",
          "Enterprise (all)",
          "Mid-mkt mfg",
        ],
        [
          "Pricing",
          "$30-250K ACV",
          "$2.4-48K ACV",
          "$100-500K+ ACV",
          "$20-80K ACV",
        ],
        [
          "Win Rate vs.",
          "\u2014",
          { text: "62% (win)", options: { color: COLORS.green, bold: true } },
          { text: "34% (lose)", options: { color: COLORS.red, bold: true } },
          {
            text: "55% (watch)",
            options: { color: COLORS.amber, bold: true },
          },
        ],
        [
          "Q3 Move",
          "\u2014",
          "Mfg module ($199/mo)",
          "Acme Lite at $8K",
          "Series A; poaching",
        ],
      ],
      { colW: [1.0, 2.0, 2.0, 2.0, 2.0], rowH: 0.22 }
    );
  }

  // ===========================================================
  // SLIDE 6: Board Asks
  // ===========================================================
  {
    const slide = pres.addSlide();
    slide.background = { color: COLORS.navyDark };

    // Top accent bar
    slide.addShape("rect", {
      x: 0,
      y: 0,
      w: SW,
      h: 0.06,
      fill: { color: COLORS.accent },
    });

    // Title
    slide.addText("Board Asks", {
      x: 0.6,
      y: 0.25,
      w: 5,
      h: 0.5,
      fontSize: 28,
      fontFace: FONTS.title,
      color: COLORS.white,
      bold: true,
      margin: 0,
    });
    slide.addText(
      "Three items for board review \u2014 two for approval, one for discussion",
      {
        x: 0.6,
        y: 0.72,
        w: 8,
        h: 0.3,
        fontSize: 11,
        fontFace: FONTS.body,
        color: COLORS.grayLight,
        margin: 0,
      }
    );

    // === Three Ask Cards ===
    const askY1 = 1.2;
    const askW = 2.85;
    const askH = 3.5;
    const askGap = 0.225;
    const askX1 = 0.6;
    const askX2 = askX1 + askW + askGap;
    const askX3 = askX2 + askW + askGap;

    // Card 1: CS Investment
    slide.addShape("rect", {
      x: askX1,
      y: askY1,
      w: askW,
      h: askH,
      fill: { color: COLORS.navy },
    });
    slide.addShape("rect", {
      x: askX1 + 0.15,
      y: askY1 + 0.15,
      w: 1.0,
      h: 0.26,
      fill: { color: COLORS.green },
    });
    slide.addText("APPROVE", {
      x: askX1 + 0.15,
      y: askY1 + 0.15,
      w: 1.0,
      h: 0.26,
      fontSize: 9,
      fontFace: FONTS.body,
      color: COLORS.white,
      bold: true,
      align: "center",
      valign: "middle",
      margin: 0,
    });
    slide.addText("$2.5M Customer Success Investment", {
      x: askX1 + 0.15,
      y: askY1 + 0.52,
      w: askW - 0.3,
      h: 0.4,
      fontSize: 13,
      fontFace: FONTS.title,
      color: COLORS.white,
      bold: true,
      margin: 0,
    });
    slide.addText(
      [
        {
          text: "8 FTEs: 4 onboarding, 2 renewal mgrs, 1 SMB CS lead, 1 CS ops",
          options: {
            bullet: true,
            breakLine: true,
            fontSize: 9,
            color: COLORS.ice,
          },
        },
        {
          text: "ROI: Reduce preventable churn 60%",
          options: {
            bullet: true,
            breakLine: true,
            fontSize: 9,
            color: COLORS.ice,
          },
        },
        {
          text: "Save ~$700K ARR/year",
          options: {
            bullet: true,
            breakLine: true,
            fontSize: 9,
            color: COLORS.ice,
          },
        },
        {
          text: "Payback: 14 months",
          options: {
            bullet: true,
            breakLine: true,
            fontSize: 9,
            color: COLORS.ice,
          },
        },
        { text: "", options: { breakLine: true, fontSize: 5 } },
        {
          text: "URGENCY",
          options: {
            bold: true,
            fontSize: 8,
            color: COLORS.red,
            breakLine: true,
            charSpacing: 2,
          },
        },
        {
          text: "Q4 churn could reach $600K+ without action. 58% of Q3 churn was preventable.",
          options: { fontSize: 9, color: COLORS.grayLight },
        },
      ],
      {
        x: askX1 + 0.15,
        y: askY1 + 0.95,
        w: askW - 0.3,
        h: 2.4,
        fontFace: FONTS.body,
        paraSpaceAfter: 3,
        valign: "top",
        margin: 0,
      }
    );

    // Card 2: SMB Pricing
    slide.addShape("rect", {
      x: askX2,
      y: askY1,
      w: askW,
      h: askH,
      fill: { color: COLORS.navy },
    });
    slide.addShape("rect", {
      x: askX2 + 0.15,
      y: askY1 + 0.15,
      w: 1.0,
      h: 0.26,
      fill: { color: COLORS.green },
    });
    slide.addText("APPROVE", {
      x: askX2 + 0.15,
      y: askY1 + 0.15,
      w: 1.0,
      h: 0.26,
      fontSize: 9,
      fontFace: FONTS.body,
      color: COLORS.white,
      bold: true,
      align: "center",
      valign: "middle",
      margin: 0,
    });
    slide.addText("Usage-Based SMB Pricing Tier", {
      x: askX2 + 0.15,
      y: askY1 + 0.52,
      w: askW - 0.3,
      h: 0.4,
      fontSize: 13,
      fontFace: FONTS.title,
      color: COLORS.white,
      bold: true,
      margin: 0,
    });
    slide.addText(
      [
        {
          text: "$499/mo entry (vs current $2.5K/mo min)",
          options: {
            bullet: true,
            breakLine: true,
            fontSize: 9,
            color: COLORS.ice,
          },
        },
        {
          text: "Captures 60+ SMB prospects/quarter lost on price",
          options: {
            bullet: true,
            breakLine: true,
            fontSize: 9,
            color: COLORS.ice,
          },
        },
        {
          text: "+$1.2M new ARR in first 12 months",
          options: {
            bullet: true,
            breakLine: true,
            fontSize: 9,
            color: COLORS.ice,
          },
        },
        {
          text: "Provides down-tier option (reduces churn)",
          options: {
            bullet: true,
            breakLine: true,
            fontSize: 9,
            color: COLORS.ice,
          },
        },
        { text: "", options: { breakLine: true, fontSize: 5 } },
        {
          text: "RISK",
          options: {
            bold: true,
            fontSize: 8,
            color: COLORS.amber,
            breakLine: true,
            charSpacing: 2,
          },
        },
        {
          text: "Some existing SMBs may down-tier (~$200K ARR impact). Net positive: +$1.0M ARR in year 1.",
          options: { fontSize: 9, color: COLORS.grayLight },
        },
      ],
      {
        x: askX2 + 0.15,
        y: askY1 + 0.95,
        w: askW - 0.3,
        h: 2.4,
        fontFace: FONTS.body,
        paraSpaceAfter: 3,
        valign: "top",
        margin: 0,
      }
    );

    // Card 3: Series C
    slide.addShape("rect", {
      x: askX3,
      y: askY1,
      w: askW,
      h: askH,
      fill: { color: COLORS.navy },
    });
    slide.addShape("rect", {
      x: askX3 + 0.15,
      y: askY1 + 0.15,
      w: 1.0,
      h: 0.26,
      fill: { color: COLORS.amber },
    });
    slide.addText("DISCUSS", {
      x: askX3 + 0.15,
      y: askY1 + 0.15,
      w: 1.0,
      h: 0.26,
      fontSize: 9,
      fontFace: FONTS.body,
      color: COLORS.white,
      bold: true,
      align: "center",
      valign: "middle",
      margin: 0,
    });
    slide.addText("Series C Timing", {
      x: askX3 + 0.15,
      y: askY1 + 0.52,
      w: askW - 0.3,
      h: 0.4,
      fontSize: 13,
      fontFace: FONTS.title,
      color: COLORS.white,
      bold: true,
      margin: 0,
    });
    slide.addText(
      [
        {
          text: "Current runway: 81 months (comfortable)",
          options: {
            bullet: true,
            breakLine: true,
            fontSize: 9,
            color: COLORS.ice,
          },
        },
        {
          text: "Growth accelerating \u2014 Q3 best quarter",
          options: {
            bullet: true,
            breakLine: true,
            fontSize: 9,
            color: COLORS.ice,
          },
        },
        { text: "", options: { breakLine: true, fontSize: 5 } },
        {
          text: "OPTION A: Q2 2027",
          options: {
            bold: true,
            fontSize: 9,
            color: COLORS.accent,
            breakLine: true,
          },
        },
        {
          text: "Raise at $22-25M ARR. Position of strength. More competitive pressure to address.",
          options: {
            bullet: true,
            breakLine: true,
            fontSize: 9,
            color: COLORS.grayLight,
          },
        },
        { text: "", options: { breakLine: true, fontSize: 5 } },
        {
          text: "OPTION B: H2 2027",
          options: {
            bold: true,
            fontSize: 9,
            color: COLORS.accent,
            breakLine: true,
          },
        },
        {
          text: "Wait for $30M ARR milestone. Higher valuation but more execution risk.",
          options: {
            bullet: true,
            breakLine: true,
            fontSize: 9,
            color: COLORS.grayLight,
          },
        },
        { text: "", options: { breakLine: true, fontSize: 5 } },
        {
          text: "Input needed on timing, target raise size, and investor targets.",
          options: { fontSize: 9, color: COLORS.ice, italic: true },
        },
      ],
      {
        x: askX3 + 0.15,
        y: askY1 + 0.95,
        w: askW - 0.3,
        h: 2.4,
        fontFace: FONTS.body,
        paraSpaceAfter: 3,
        valign: "top",
        margin: 0,
      }
    );

    // Bottom bar
    slide.addShape("rect", {
      x: 0,
      y: SH - 0.45,
      w: SW,
      h: 0.45,
      fill: { color: COLORS.navy },
    });
    slide.addText("NovaCrest  |  Q3 2026 Board Review  |  Confidential", {
      x: 0.6,
      y: SH - 0.45,
      w: 8.8,
      h: 0.45,
      fontSize: 9,
      fontFace: FONTS.body,
      color: COLORS.grayLight,
      valign: "middle",
      margin: 0,
    });
  }

  // Save
  const outPath = path.resolve(__dirname, "../outputs/strategy-anthropic.pptx");
  await pres.writeFile({ fileName: outPath });
  console.log("Saved to", outPath);
}

build().catch((err) => {
  console.error(err);
  process.exit(1);
});
