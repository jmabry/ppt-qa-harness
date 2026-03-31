const pptxgen = require("pptxgenjs");

// ─── Color Palette: Midnight Executive ───
const C = {
  navy:      "1E2761",
  darkNavy:  "141B42",
  ice:       "CADCFC",
  white:     "FFFFFF",
  offWhite:  "F0F2FA",
  accent:    "3B82F6",   // bright blue accent
  accentDim: "2563EB",
  red:       "EF4444",
  amber:     "F59E0B",
  green:     "22C55E",
  gray:      "94A3B8",
  darkGray:  "475569",
  lightGray: "E2E8F0",
  text:      "1E293B",
  textMuted: "64748B",
  critical:  "DC2626",
  high:      "EA580C",
  medium:    "D97706",
  cardBg:    "F8FAFC",
};

const FONT = { head: "Georgia", body: "Calibri" };

// ─── Helpers ───
const makeShadow = () => ({
  type: "outer", color: "000000", blur: 4, offset: 2, angle: 135, opacity: 0.10,
});

function addSlideNumber(slide, num) {
  slide.addText(`${num} / 6`, {
    x: 8.8, y: 5.25, w: 1, h: 0.3,
    fontSize: 9, fontFace: FONT.body, color: C.gray, align: "right",
  });
}

// ─── Table helper ───
function makeTable(slide, headers, rows, opts) {
  const {
    x = 0.5, y = 1.5, w = 9, colW,
    headerFill = C.navy, headerColor = C.white,
    rowFill1 = C.white, rowFill2 = C.offWhite,
    fontSize = 9, headerFontSize = 9,
  } = opts || {};

  const tableRows = [];

  // Header row
  tableRows.push(
    headers.map((h) => ({
      text: h,
      options: {
        bold: true,
        fontSize: headerFontSize,
        fontFace: FONT.body,
        color: headerColor,
        fill: { color: headerFill },
        valign: "middle",
        align: "left",
        margin: [3, 4, 3, 4],
      },
    }))
  );

  // Data rows
  rows.forEach((row, ri) => {
    tableRows.push(
      row.map((cell) => {
        const isObj = typeof cell === "object" && cell !== null && cell.text !== undefined;
        return {
          text: isObj ? cell.text : String(cell),
          options: {
            fontSize,
            fontFace: FONT.body,
            color: isObj && cell.color ? cell.color : C.text,
            bold: isObj && cell.bold ? true : false,
            fill: { color: ri % 2 === 0 ? rowFill1 : rowFill2 },
            valign: "middle",
            align: "left",
            margin: [2, 4, 2, 4],
          },
        };
      })
    );
  });

  const rowH = Array(tableRows.length).fill(0.28);
  rowH[0] = 0.32;

  slide.addTable(tableRows, {
    x, y, w,
    colW,
    rowH,
    border: { pt: 0.5, color: C.lightGray },
  });
}


// ═══════════════════════════════════════════
//  BUILD PRESENTATION
// ═══════════════════════════════════════════
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.author = "Project Chimera";
pres.title = "Project Chimera: Monolith Decomposition Plan";


// ═══════════════════════════════════════════
//  SLIDE 1 — Title + Current State
// ═══════════════════════════════════════════
{
  const slide = pres.addSlide();
  slide.background = { color: C.darkNavy };

  // Title block
  slide.addText("PROJECT CHIMERA", {
    x: 0.6, y: 0.3, w: 9, h: 0.6,
    fontSize: 32, fontFace: FONT.head, color: C.white, bold: true,
    charSpacing: 4,
  });
  slide.addText("Monolith Decomposition Plan", {
    x: 0.6, y: 0.85, w: 6, h: 0.4,
    fontSize: 16, fontFace: FONT.body, color: C.ice, italic: true,
  });

  // System stats row — 4 cards
  const stats = [
    { val: "340K", label: "Lines of Python" },
    { val: "2,800", label: "Database Tables" },
    { val: "14", label: "Django Apps" },
    { val: "22", label: "Backend Engineers" },
  ];
  const cardW = 2.05;
  const cardGap = 0.17;
  const startX = 0.6;
  stats.forEach((s, i) => {
    const cx = startX + i * (cardW + cardGap);
    slide.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: 1.45, w: cardW, h: 0.85,
      fill: { color: C.navy },
      shadow: makeShadow(),
    });
    slide.addText(s.val, {
      x: cx, y: 1.45, w: cardW, h: 0.5,
      fontSize: 22, fontFace: FONT.head, color: C.accent, bold: true,
      align: "center", valign: "bottom", margin: 0,
    });
    slide.addText(s.label, {
      x: cx, y: 1.95, w: cardW, h: 0.35,
      fontSize: 9, fontFace: FONT.body, color: C.gray,
      align: "center", valign: "top", margin: 0,
    });
  });

  // Pain metrics table — dark theme
  const headers = ["Metric", "Current", "DORA P50", "Gap"];
  const rows = [
    ["Deploy frequency", "1.8/week", "1/day – 1/week", { text: "Bottom of medium", color: C.amber }],
    ["Lead time for changes", "12.3 days", "1 day – 1 week", { text: "2x over ceiling", color: C.red }],
    ["Change failure rate", "18.2%", "0–15%", { text: "Above threshold", color: C.red }],
    ["MTTR", "4.2 hours", "< 1 hr – < 1 day", { text: "Functional but slow", color: C.amber }],
    ["CI build time", "47 min", "—", { text: "Context-switch tax", color: C.amber }],
    ["Merge conflicts/week", "3.2 avg", "—", { text: "Coordination tax", color: C.amber }],
    ["Rollback rate", "1 in 5.5", "—", { text: "Low confidence", color: C.red }],
    ["Test flakiness", "6.8%", "—", { text: "Retry & pray", color: C.red }],
  ];

  const tRows = [];
  tRows.push(
    headers.map((h) => ({
      text: h,
      options: {
        bold: true, fontSize: 8.5, fontFace: FONT.body,
        color: C.ice, fill: { color: C.accent },
        valign: "middle", align: "left", margin: [2, 4, 2, 4],
      },
    }))
  );
  rows.forEach((row, ri) => {
    tRows.push(
      row.map((cell) => {
        const isObj = typeof cell === "object" && cell.text !== undefined;
        return {
          text: isObj ? cell.text : String(cell),
          options: {
            fontSize: 8, fontFace: FONT.body,
            color: isObj ? cell.color : C.ice,
            bold: isObj,
            fill: { color: ri % 2 === 0 ? C.darkNavy : C.navy },
            valign: "middle", align: "left", margin: [2, 4, 2, 4],
          },
        };
      })
    );
  });

  slide.addTable(tRows, {
    x: 0.6, y: 2.5, w: 8.8,
    colW: [2.2, 1.5, 2.0, 3.1],
    rowH: Array(tRows.length).fill(0.27),
    border: { pt: 0.5, color: C.navy },
  });

  // Cost callout
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 4.85, w: 8.8, h: 0.5,
    fill: { color: C.navy },
  });
  slide.addText([
    { text: "Cost of monolith friction: ", options: { fontSize: 11, color: C.ice } },
    { text: "~1,400 eng-hours/quarter ", options: { fontSize: 11, color: C.accent, bold: true } },
    { text: "= ", options: { fontSize: 11, color: C.ice } },
    { text: "$133K/quarter wasted", options: { fontSize: 13, color: C.red, bold: true } },
  ], {
    x: 0.6, y: 4.85, w: 8.8, h: 0.5,
    align: "center", valign: "middle", margin: 0,
    fontFace: FONT.body,
  });

  addSlideNumber(slide, 1);
}


// ═══════════════════════════════════════════
//  SLIDE 2 — Target Architecture
// ═══════════════════════════════════════════
{
  const slide = pres.addSlide();
  slide.background = { color: C.offWhite };

  slide.addText("Target Architecture", {
    x: 0.5, y: 0.25, w: 9, h: 0.5,
    fontSize: 28, fontFace: FONT.head, color: C.navy, bold: true,
  });
  slide.addText("Strangler Fig — incremental extraction behind API Gateway", {
    x: 0.5, y: 0.7, w: 9, h: 0.3,
    fontSize: 12, fontFace: FONT.body, color: C.textMuted, italic: true,
  });

  // Service boundary table
  makeTable(
    slide,
    ["Service", "Owner", "DB", "API Style", "Priority"],
    [
      [{ text: "User Service", bold: true }, "Identity", "PostgreSQL (isolated)", "REST + gRPC", { text: "P0", color: C.critical, bold: true }],
      [{ text: "Order Service", bold: true }, "Commerce", "PostgreSQL (isolated)", "REST + events", { text: "P0", color: C.critical, bold: true }],
      [{ text: "Payment Service", bold: true }, "Commerce", "PostgreSQL + Stripe", "REST (sync)", { text: "P1", color: C.high, bold: true }],
      [{ text: "Search Service", bold: true }, "Discovery", "Elasticsearch 8", "REST", { text: "P1", color: C.high, bold: true }],
      [{ text: "Notification Svc", bold: true }, "Platform", "PostgreSQL + SQS", "Async (events)", { text: "P2", color: C.medium, bold: true }],
      [{ text: "Analytics Service", bold: true }, "Data", "ClickHouse", "gRPC (internal)", { text: "P2", color: C.medium, bold: true }],
    ],
    { x: 0.5, y: 1.15, colW: [1.7, 1.2, 2.0, 1.8, 0.8], fontSize: 8.5, headerFontSize: 8.5 }
  );

  // Architecture diagram — simplified visual
  const diagramY = 3.45;
  const diagramH = 2.0;

  // API Gateway bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: diagramY, w: 9, h: 0.35,
    fill: { color: C.accent },
  });
  slide.addText("API Gateway (Kong)  →  Service Mesh (Istio)  →  Rate Limiting  →  Circuit Breakers", {
    x: 0.5, y: diagramY, w: 9, h: 0.35,
    fontSize: 9, fontFace: FONT.body, color: C.white, bold: true,
    align: "center", valign: "middle", margin: 0,
  });

  // Service boxes row 1
  const svcRow1 = [
    { name: "User\nService", color: C.navy },
    { name: "Order\nService", color: C.navy },
    { name: "Payment\nService", color: C.navy },
  ];
  const svcW = 1.7;
  const svcGap = 0.25;
  const row1X = 0.5;
  svcRow1.forEach((svc, i) => {
    const sx = row1X + i * (svcW + svcGap);
    slide.addShape(pres.shapes.RECTANGLE, {
      x: sx, y: diagramY + 0.5, w: svcW, h: 0.55,
      fill: { color: svc.color }, shadow: makeShadow(),
    });
    slide.addText(svc.name, {
      x: sx, y: diagramY + 0.5, w: svcW, h: 0.55,
      fontSize: 9, fontFace: FONT.body, color: C.white, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
  });

  // Kafka event bus bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: diagramY + 1.2, w: 5.6, h: 0.3,
    fill: { color: C.accentDim },
  });
  slide.addText("Event Bus (Kafka) — user.*, order.*, payment.*", {
    x: 0.5, y: diagramY + 1.2, w: 5.6, h: 0.3,
    fontSize: 8, fontFace: FONT.body, color: C.white,
    align: "center", valign: "middle", margin: 0,
  });

  // Service boxes row 2
  const svcRow2 = [
    { name: "Search\nService", color: C.darkGray },
    { name: "Notification\nService", color: C.darkGray },
    { name: "Analytics\nService", color: C.darkGray },
  ];
  svcRow2.forEach((svc, i) => {
    const sx = row1X + i * (svcW + svcGap);
    slide.addShape(pres.shapes.RECTANGLE, {
      x: sx, y: diagramY + 1.65, w: svcW, h: 0.55,
      fill: { color: svc.color }, shadow: makeShadow(),
    });
    slide.addText(svc.name, {
      x: sx, y: diagramY + 1.65, w: svcW, h: 0.55,
      fontSize: 9, fontFace: FONT.body, color: C.white, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
  });

  // Data stores column on right
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 6.3, y: diagramY + 0.5, w: 3.2, h: 1.7,
    fill: { color: C.cardBg },
    line: { color: C.lightGray, width: 1 },
  });
  slide.addText("Isolated Data Stores", {
    x: 6.3, y: diagramY + 0.5, w: 3.2, h: 0.3,
    fontSize: 9, fontFace: FONT.body, color: C.navy, bold: true,
    align: "center", valign: "middle", margin: 0,
  });
  slide.addText([
    { text: "PostgreSQL × 4", options: { breakLine: true, fontSize: 8.5 } },
    { text: "Elasticsearch 8", options: { breakLine: true, fontSize: 8.5 } },
    { text: "ClickHouse (analytics)", options: { breakLine: true, fontSize: 8.5 } },
    { text: "Redis (per-service)", options: { fontSize: 8.5 } },
  ], {
    x: 6.5, y: diagramY + 0.8, w: 2.8, h: 1.2,
    fontFace: FONT.body, color: C.text,
    bullet: true, valign: "top", margin: 0,
  });

  addSlideNumber(slide, 2);
}


// ═══════════════════════════════════════════
//  SLIDE 3 — Why Strangler Fig (Dense Text)
// ═══════════════════════════════════════════
{
  const slide = pres.addSlide();
  slide.background = { color: C.darkNavy };

  slide.addText("Architecture Decision: Why Strangler Fig", {
    x: 0.5, y: 0.25, w: 9, h: 0.5,
    fontSize: 26, fontFace: FONT.head, color: C.white, bold: true,
  });
  slide.addText("Evaluated 3 approaches in 2-week spike (April 14–25, 2026)", {
    x: 0.5, y: 0.7, w: 9, h: 0.3,
    fontSize: 11, fontFace: FONT.body, color: C.gray, italic: true,
  });

  // Three option cards
  const options = [
    {
      title: "Full Rewrite",
      verdict: "REJECTED",
      verdictColor: C.red,
      body: "Greenfield in Go + gRPC. Estimated 14–18 months with 8 engineers. Requires maintaining two systems in parallel. Case studies (Netscape, Basecamp) show rewrites take 2–3x longer than estimated. Our 340K lines contain tested business logic — rewriting means re-discovering every edge case.",
    },
    {
      title: "Modularize Monolith",
      verdict: "INSUFFICIENT",
      verdictColor: C.amber,
      body: "Tried Q4 2025 — \"Atlas Modular\" spent 6 weeks extracting Django apps into packages. Result: merge conflicts dropped 30% but deploy frequency unchanged. Still one artifact, 47-min CI, same blast radius. Addresses code organization, not deployment coupling — our actual bottleneck.",
    },
    {
      title: "Strangler Fig",
      verdict: "SELECTED",
      verdictColor: C.green,
      body: "Incremental extraction behind API Gateway. Each service runs alongside monolith — validate before cutover. Traffic routes back on failure. Risk bounded to one service at a time. Already proven: User Service prototype (Q1 2026) took 3 weeks, ran in parallel 2 weeks, cut over with zero incidents.",
    },
  ];

  const cardW = 2.85;
  const cardGap = 0.15;
  options.forEach((opt, i) => {
    const cx = 0.5 + i * (cardW + cardGap);
    // Card background
    slide.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: 1.15, w: cardW, h: 2.95,
      fill: { color: C.navy },
      shadow: makeShadow(),
    });
    // Title
    slide.addText(opt.title, {
      x: cx + 0.15, y: 1.25, w: cardW - 0.3, h: 0.35,
      fontSize: 13, fontFace: FONT.head, color: C.white, bold: true,
      margin: 0,
    });
    // Verdict badge
    slide.addShape(pres.shapes.RECTANGLE, {
      x: cx + 0.15, y: 1.65, w: 1.2, h: 0.25,
      fill: { color: opt.verdictColor },
    });
    slide.addText(opt.verdict, {
      x: cx + 0.15, y: 1.65, w: 1.2, h: 0.25,
      fontSize: 8, fontFace: FONT.body, color: C.white, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
    // Body text
    slide.addText(opt.body, {
      x: cx + 0.15, y: 2.0, w: cardW - 0.3, h: 2.0,
      fontSize: 9, fontFace: FONT.body, color: C.ice,
      valign: "top", margin: 0, lineSpacingMultiple: 1.15,
    });
  });

  // Bottom section: What we're NOT doing + Timeline confidence
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 4.25, w: 4.3, h: 1.1,
    fill: { color: C.navy },
  });
  slide.addText("What We're NOT Doing", {
    x: 0.65, y: 4.3, w: 4.0, h: 0.25,
    fontSize: 11, fontFace: FONT.body, color: C.accent, bold: true, margin: 0,
  });
  slide.addText("Reporting and Admin modules stay in the monolith indefinitely — low-traffic, low-change. Cost of extraction exceeds benefit. Goal is not \"zero monolith\" — it's removing modules that cause deployment friction for teams that ship the most.", {
    x: 0.65, y: 4.55, w: 4.0, h: 0.75,
    fontSize: 8.5, fontFace: FONT.body, color: C.ice, valign: "top", margin: 0,
    lineSpacingMultiple: 1.15,
  });

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 4.95, y: 4.25, w: 4.55, h: 1.1,
    fill: { color: C.navy },
  });
  slide.addText("Timeline Confidence", {
    x: 5.1, y: 4.3, w: 4.2, h: 0.25,
    fontSize: 11, fontFace: FONT.body, color: C.accent, bold: true, margin: 0,
  });
  slide.addText([
    { text: "Phase 1–2 (core extractions): ", options: { bold: true, color: C.green } },
    { text: "65% confident in 6-month timeline", options: { color: C.ice, breakLine: true } },
    { text: "Phase 3 (platform): ", options: { bold: true, color: C.amber } },
    { text: "40% confident — ClickHouse adoption has unknown learning curve. Padded 2 weeks but may extend.", options: { color: C.ice } },
  ], {
    x: 5.1, y: 4.55, w: 4.2, h: 0.75,
    fontSize: 8.5, fontFace: FONT.body, valign: "top", margin: 0,
    lineSpacingMultiple: 1.15,
  });

  addSlideNumber(slide, 3);
}


// ═══════════════════════════════════════════
//  SLIDE 4 — Migration Plan (6-Month Timeline)
// ═══════════════════════════════════════════
{
  const slide = pres.addSlide();
  slide.background = { color: C.offWhite };

  slide.addText("Migration Plan — 6-Month Phased Rollout", {
    x: 0.5, y: 0.2, w: 9, h: 0.45,
    fontSize: 26, fontFace: FONT.head, color: C.navy, bold: true,
  });

  // Phase timeline bars
  const phases = [
    {
      label: "PHASE 1: FOUNDATION",
      months: "Month 1–2",
      color: C.accent,
      y: 0.75,
      tasks: [
        ["Deploy API Gateway (Kong)", "SRE", "Gateway becomes SPOF — need HA"],
        ["Instrument w/ OpenTelemetry", "SRE + all squads", "~2–3% perf overhead"],
        ["Extract User Service", "Identity", "Session migration (cookie → JWT)"],
        ["Set up Kafka cluster (3 brokers)", "SRE", "No Kafka experience on team"],
        ["Database migration tooling", "DBA", "Data consistency during migration"],
      ],
    },
    {
      label: "PHASE 2: COMMERCE CORE",
      months: "Month 3–4",
      color: C.accentDim,
      y: 2.35,
      tasks: [
        ["Extract Order Service", "Commerce", "47 cross-module imports to untangle"],
        ["Extract Payment Service", "Commerce", "PCI compliance scope change"],
        ["Implement saga pattern (checkout)", "Commerce + Identity", "Complex failure/compensation logic"],
        ["Search Service extraction", "Discovery", "4-hr index rebuild, need zero-downtime"],
      ],
    },
    {
      label: "PHASE 3: PLATFORM & OPTIMIZATION",
      months: "Month 5–6",
      color: C.darkGray,
      y: 3.75,
      tasks: [
        ["Notification Service extraction", "Platform", "Low risk — already loosely coupled"],
        ["Analytics Service extraction", "Data", "18 months historical data migration"],
        ["Decommission monolith modules", "All squads", "Residual coupling in shared utils"],
        ["Perf tuning + chaos engineering", "SRE", "Unknown unknowns in distributed systems"],
      ],
    },
  ];

  phases.forEach((phase) => {
    // Phase header bar
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: phase.y, w: 9, h: 0.3,
      fill: { color: phase.color },
    });
    slide.addText(`${phase.label}  (${phase.months})`, {
      x: 0.6, y: phase.y, w: 8.8, h: 0.3,
      fontSize: 10, fontFace: FONT.body, color: C.white, bold: true,
      valign: "middle", margin: 0,
    });

    // Task rows
    phase.tasks.forEach((task, ti) => {
      const ty = phase.y + 0.33 + ti * 0.26;
      const bgColor = ti % 2 === 0 ? C.white : C.offWhite;
      slide.addShape(pres.shapes.RECTANGLE, {
        x: 0.5, y: ty, w: 9, h: 0.26,
        fill: { color: bgColor },
      });
      // Task name
      slide.addText(task[0], {
        x: 0.6, y: ty, w: 4.0, h: 0.26,
        fontSize: 8.5, fontFace: FONT.body, color: C.text,
        valign: "middle", margin: 0,
      });
      // Owner
      slide.addText(task[1], {
        x: 4.7, y: ty, w: 1.8, h: 0.26,
        fontSize: 8.5, fontFace: FONT.body, color: C.accent, bold: true,
        valign: "middle", margin: 0,
      });
      // Risk
      slide.addText(task[2], {
        x: 6.6, y: ty, w: 2.8, h: 0.26,
        fontSize: 8, fontFace: FONT.body, color: C.textMuted, italic: true,
        valign: "middle", margin: 0,
      });
    });
  });

  // Column headers hint
  slide.addText("TASK", {
    x: 0.6, y: 0.55, w: 4.0, h: 0.2,
    fontSize: 8, fontFace: FONT.body, color: C.textMuted, bold: true, margin: 0,
  });
  slide.addText("OWNER", {
    x: 4.7, y: 0.55, w: 1.8, h: 0.2,
    fontSize: 8, fontFace: FONT.body, color: C.textMuted, bold: true, margin: 0,
  });
  slide.addText("KEY RISK", {
    x: 6.6, y: 0.55, w: 2.8, h: 0.2,
    fontSize: 8, fontFace: FONT.body, color: C.textMuted, bold: true, margin: 0,
  });

  addSlideNumber(slide, 4);
}


// ═══════════════════════════════════════════
//  SLIDE 5 — Risk Matrix + Monitoring
// ═══════════════════════════════════════════
{
  const slide = pres.addSlide();
  slide.background = { color: C.offWhite };

  slide.addText("Risk Matrix & Observability Stack", {
    x: 0.5, y: 0.2, w: 9, h: 0.45,
    fontSize: 26, fontFace: FONT.head, color: C.navy, bold: true,
  });

  // Risk table
  slide.addText("Risk Matrix", {
    x: 0.5, y: 0.7, w: 4, h: 0.3,
    fontSize: 13, fontFace: FONT.body, color: C.navy, bold: true, margin: 0,
  });

  const riskHeaders = ["Risk", "Severity", "Mitigation"];
  const riskRows = [
    [
      "Distributed txn failures (checkout saga)",
      { text: "CRITICAL", color: C.critical, bold: true },
      "Compensation/rollback per step; feature flag fallback to monolith",
    ],
    [
      "Kafka cluster instability",
      { text: "HIGH", color: C.high, bold: true },
      "Hire Kafka specialist; use managed Kafka (Confluent Cloud)",
    ],
    [
      "Data inconsistency (dual-write)",
      { text: "HIGH", color: C.high, bold: true },
      "CDC with Debezium; hourly reconciliation jobs",
    ],
    [
      "Service-to-service latency",
      { text: "MEDIUM", color: C.medium, bold: true },
      "P99 budget 200ms; circuit breakers (500ms timeout); prefer async",
    ],
    [
      "Team cognitive overload",
      { text: "MEDIUM", color: C.medium, bold: true },
      "Phase tool adoption; pair with SRE; weekly architecture office hours",
    ],
    [
      "PCI scope expansion",
      { text: "MEDIUM", color: C.medium, bold: true },
      "Security review in M2; Stripe hosted checkout; document boundaries",
    ],
  ];

  makeTable(slide, riskHeaders, riskRows, {
    x: 0.5, y: 1.0, colW: [2.8, 1.0, 5.2],
    fontSize: 8, headerFontSize: 8.5,
  });

  // Observability stack
  slide.addText("Monitoring & Observability", {
    x: 0.5, y: 3.2, w: 5, h: 0.3,
    fontSize: 13, fontFace: FONT.body, color: C.navy, bold: true, margin: 0,
  });

  const obsHeaders = ["Layer", "Tool", "Alert Threshold"];
  const obsRows = [
    [{ text: "Metrics", bold: true }, "Prometheus + Grafana", "Error rate > 1% (5 min); P99 > 500ms (10 min)"],
    [{ text: "Tracing", bold: true }, "Jaeger + OpenTelemetry", "Trace duration > 2s; span error rate > 5%"],
    [{ text: "Logging", bold: true }, "ELK Stack", "Error log rate > 50/min per service"],
    [{ text: "Health", bold: true }, "K8s probes", "Probe failure → restart; 3 consecutive → page SRE"],
    [{ text: "Synthetic", bold: true }, "Datadog Synthetics", "Critical flow failure → immediate page"],
    [{ text: "Alerting", bold: true }, "PagerDuty", "P1 → page, P2 → Slack, P3 → next business day"],
  ];

  makeTable(slide, obsHeaders, obsRows, {
    x: 0.5, y: 3.5, colW: [1.2, 2.3, 5.5],
    fontSize: 8, headerFontSize: 8.5,
  });

  // Key dashboards callout
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 5.25, w: 9, h: 0.01,
    fill: { color: C.accent },
  });

  addSlideNumber(slide, 5);
}


// ═══════════════════════════════════════════
//  SLIDE 6 — Success Metrics + Team Structure
// ═══════════════════════════════════════════
{
  const slide = pres.addSlide();
  slide.background = { color: C.darkNavy };

  slide.addText("Success Metrics & Team Structure", {
    x: 0.5, y: 0.2, w: 9, h: 0.45,
    fontSize: 26, fontFace: FONT.head, color: C.white, bold: true,
  });

  // Before/After metrics table
  slide.addText("Before / After Targets", {
    x: 0.5, y: 0.7, w: 5, h: 0.25,
    fontSize: 12, fontFace: FONT.body, color: C.accent, bold: true, margin: 0,
  });

  const metricHeaders = ["Metric", "Current", "6-Month", "12-Month"];
  const metricRows = [
    ["Deploy frequency", "1.8/week", { text: "3/week/svc", color: C.amber }, { text: "1+/day/svc", color: C.green }],
    ["Lead time", "12.3 days", { text: "5 days", color: C.amber }, { text: "< 2 days", color: C.green }],
    ["Change failure rate", "18.2%", { text: "12%", color: C.amber }, { text: "< 8%", color: C.green }],
    ["MTTR", "4.2 hours", { text: "1.5 hours", color: C.amber }, { text: "< 30 min", color: C.green }],
    ["P99 API latency", "820ms", { text: "400ms", color: C.amber }, { text: "< 200ms", color: C.green }],
    ["CI build time", "47 min", { text: "15 min/svc", color: C.amber }, { text: "< 8 min", color: C.green }],
    ["Test flakiness", "6.8%", { text: "3%", color: C.amber }, { text: "< 1%", color: C.green }],
    ["Dev satisfaction", "5.8/10", { text: "7.0/10", color: C.amber }, { text: "8.0/10", color: C.green }],
  ];

  // Dark table
  const mRows = [];
  mRows.push(
    metricHeaders.map((h) => ({
      text: h,
      options: {
        bold: true, fontSize: 8.5, fontFace: FONT.body,
        color: C.white, fill: { color: C.accent },
        valign: "middle", align: "left", margin: [2, 4, 2, 4],
      },
    }))
  );
  metricRows.forEach((row, ri) => {
    mRows.push(
      row.map((cell) => {
        const isObj = typeof cell === "object" && cell.text !== undefined;
        return {
          text: isObj ? cell.text : String(cell),
          options: {
            fontSize: 8, fontFace: FONT.body,
            color: isObj ? cell.color : C.ice,
            bold: isObj,
            fill: { color: ri % 2 === 0 ? C.darkNavy : C.navy },
            valign: "middle", align: "left", margin: [2, 4, 2, 4],
          },
        };
      })
    );
  });

  slide.addTable(mRows, {
    x: 0.5, y: 0.95, w: 5.3,
    colW: [1.6, 1.1, 1.3, 1.3],
    rowH: Array(mRows.length).fill(0.25),
    border: { pt: 0.5, color: C.navy },
  });

  // Team structure table
  slide.addText("Target Team Ownership", {
    x: 6.0, y: 0.7, w: 3.8, h: 0.25,
    fontSize: 12, fontFace: FONT.body, color: C.accent, bold: true, margin: 0,
  });

  const teamHeaders = ["Squad", "Services", "Size"];
  const teamRows = [
    [{ text: "Identity", bold: true }, "User, Auth", "4 eng"],
    [{ text: "Commerce", bold: true }, "Order, Payment", "6 eng"],
    [{ text: "Discovery", bold: true }, "Search", "3 eng"],
    [{ text: "Platform", bold: true }, "Notification, Gateway", "4 eng"],
    [{ text: "Data", bold: true }, "Analytics", "3 eng"],
    [{ text: "SRE", bold: true }, "Kafka, K8s, Monitoring", "3 eng"],
  ];

  const tRows = [];
  tRows.push(
    teamHeaders.map((h) => ({
      text: h,
      options: {
        bold: true, fontSize: 8.5, fontFace: FONT.body,
        color: C.white, fill: { color: C.accent },
        valign: "middle", align: "left", margin: [2, 4, 2, 4],
      },
    }))
  );
  teamRows.forEach((row, ri) => {
    tRows.push(
      row.map((cell) => {
        const isObj = typeof cell === "object" && cell.text !== undefined;
        return {
          text: isObj ? cell.text : String(cell),
          options: {
            fontSize: 8, fontFace: FONT.body,
            color: isObj ? C.ice : C.ice,
            bold: isObj ? cell.bold : false,
            fill: { color: ri % 2 === 0 ? C.darkNavy : C.navy },
            valign: "middle", align: "left", margin: [2, 4, 2, 4],
          },
        };
      })
    );
  });

  slide.addTable(tRows, {
    x: 6.0, y: 0.95, w: 3.5,
    colW: [1.1, 1.5, 0.9],
    rowH: Array(tRows.length).fill(0.25),
    border: { pt: 0.5, color: C.navy },
  });

  // Principle callout at bottom
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 6.0, y: 2.85, w: 3.5, h: 0.7,
    fill: { color: C.navy },
  });
  slide.addText([
    { text: "You build it, you run it.", options: { bold: true, color: C.accent, fontSize: 11, breakLine: true } },
    { text: "Each squad owns their SLOs, deploys independently, and carries their own pager.", options: { color: C.ice, fontSize: 8.5 } },
  ], {
    x: 6.15, y: 2.9, w: 3.2, h: 0.6,
    fontFace: FONT.body, valign: "top", margin: 0,
    lineSpacingMultiple: 1.15,
  });

  // Key dashboards at bottom
  slide.addText("Key Grafana Dashboards", {
    x: 0.5, y: 3.65, w: 5, h: 0.25,
    fontSize: 12, fontFace: FONT.body, color: C.accent, bold: true, margin: 0,
  });

  const dashboards = [
    { name: "Service Overview", desc: "Request rate, error %, P50/P95/P99 latency per service" },
    { name: "Kafka Health", desc: "Consumer lag, broker disk, replication status" },
    { name: "Checkout Saga", desc: "Success rate, avg duration, failure breakdown by step" },
    { name: "Deployment Tracker", desc: "Last deploy time, canary status, rollback count" },
  ];

  dashboards.forEach((d, i) => {
    const dy = 3.95 + i * 0.38;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: dy, w: 9, h: 0.35,
      fill: { color: i % 2 === 0 ? C.navy : C.darkNavy },
    });
    slide.addText(d.name, {
      x: 0.65, y: dy, w: 2.5, h: 0.35,
      fontSize: 9, fontFace: FONT.body, color: C.white, bold: true,
      valign: "middle", margin: 0,
    });
    slide.addText(d.desc, {
      x: 3.2, y: dy, w: 6.1, h: 0.35,
      fontSize: 8.5, fontFace: FONT.body, color: C.gray,
      valign: "middle", margin: 0,
    });
  });

  addSlideNumber(slide, 6);
}


// ─── Write file ───
pres.writeFile({ fileName: "outputs/software-anthropic.pptx" })
  .then(() => console.log("Created: outputs/software-anthropic.pptx"))
  .catch((err) => { console.error(err); process.exit(1); });
