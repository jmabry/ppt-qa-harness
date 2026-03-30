const pptxgen = require("pptxgenjs");

// ── Theme: Tech & Night (#9) — deep, luminous, data-driven ──
const theme = {
  primary: "001d3d",
  secondary: "003566",
  accent: "ffc300",
  light: "ffd60a",
  bg: "000814",
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

function sectionLabel(slide, text, x, y, w) {
  slide.addShape("rect", {
    x: x, y: y, w: w, h: 0.03,
    fill: { color: theme.accent },
  });
  slide.addText(text.toUpperCase(), {
    x: x, y: y + 0.08, w: w, h: 0.3,
    fontSize: 10, fontFace: BODY_FONT,
    color: theme.accent, bold: true,
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

function headerRow(cells) {
  return cells.map((c) => ({
    text: c,
    options: {
      bold: true,
      color: theme.bg,
      fill: { color: theme.accent },
      fontSize: 9,
      fontFace: BODY_FONT,
      align: "center",
      valign: "middle",
    },
  }));
}

function dataRow(cells, opts = {}) {
  return cells.map((c, i) => {
    const isObj = typeof c === "object" && c !== null && c.text !== undefined;
    const text = isObj ? c.text : String(c);
    const cellOpts = isObj ? c.options || {} : {};
    return {
      text,
      options: {
        fill: { color: opts.alt ? "001225" : "000e1f" },
        fontSize: 9,
        fontFace: BODY_FONT,
        color: "D0D8E0",
        valign: "middle",
        align: i === 0 ? "left" : "center",
        ...cellOpts,
      },
    };
  });
}

// ════════════════════════════════════════════════════════════════
// SLIDE 1: Title + Current State
// ════════════════════════════════════════════════════════════════
function slide01(pres) {
  const slide = pres.addSlide();
  slide.background = { color: theme.bg };

  // Top accent bar
  slide.addShape("rect", {
    x: 0, y: 0, w: 10, h: 0.06,
    fill: { color: theme.accent },
  });

  // Title block
  slide.addText("PROJECT CHIMERA", {
    x: 0.5, y: 0.25, w: 5.5, h: 0.5,
    fontSize: 14, fontFace: BODY_FONT,
    color: theme.accent, bold: true,
    charSpacing: 5,
  });
  slide.addText("Monolith\nDecomposition\nPlan", {
    x: 0.5, y: 0.7, w: 5.5, h: 1.6,
    fontSize: 38, fontFace: TITLE_FONT,
    color: "FFFFFF", bold: true,
    lineSpacingMultiple: 0.95,
  });
  slide.addText("Atlas: 340K LOC  |  14 Django Apps  |  2,800 Tables  |  22 Engineers", {
    x: 0.5, y: 2.3, w: 5.5, h: 0.3,
    fontSize: 10, fontFace: BODY_FONT,
    color: "8899AA",
  });

  // System stats cards — right side
  const stats = [
    { label: "LOC", value: "340K", sub: "Python 3.11" },
    { label: "DB Tables", value: "2,800", sub: "PostgreSQL 15" },
    { label: "Engineers", value: "22", sub: "4 squads" },
    { label: "Quarterly Waste", value: "$133K", sub: "~1,400 eng-hrs" },
  ];

  stats.forEach((s, i) => {
    const cx = 6.3 + (i % 2) * 1.8;
    const cy = 0.3 + Math.floor(i / 2) * 1.15;
    slide.addShape("rect", {
      x: cx, y: cy, w: 1.6, h: 1.0,
      fill: { color: "001225" },
      rectRadius: 0.05,
    });
    slide.addText(s.value, {
      x: cx, y: cy + 0.08, w: 1.6, h: 0.45,
      fontSize: 22, fontFace: TITLE_FONT,
      color: theme.accent, bold: true,
      align: "center", valign: "middle",
    });
    slide.addText(s.label, {
      x: cx, y: cy + 0.5, w: 1.6, h: 0.22,
      fontSize: 9, fontFace: BODY_FONT,
      color: "FFFFFF", align: "center", valign: "middle",
    });
    slide.addText(s.sub, {
      x: cx, y: cy + 0.72, w: 1.6, h: 0.2,
      fontSize: 8, fontFace: BODY_FONT,
      color: "667788", align: "center", valign: "middle",
    });
  });

  // Pain Metrics table
  sectionLabel(slide, "DORA Metrics & Pain Points", 0.5, 2.7, 9.0);

  const rows = [
    headerRow(["Metric", "Current", "Industry P50", "Gap / Impact"]),
    dataRow(["Deploy frequency", "1.8/week", "1/day - 1/week", "Bottom of medium"]),
    dataRow(["Lead time (commit to prod)", "12.3 days", "1 day - 1 week", { text: "2x over ceiling", options: { color: "FF6B6B" } }], { alt: true }),
    dataRow(["Change failure rate", "18.2%", "0 - 15%", { text: "Above threshold", options: { color: "FF6B6B" } }]),
    dataRow(["MTTR", "4.2 hours", "< 1 hour - 1 day", "Functional but slow"], { alt: true }),
    dataRow(["CI build time", "47 min", "--", "Context-switch tax"]),
    dataRow(["Merge conflicts/week", "3.2 avg (peak 8)", "--", "Cross-team coordination"], { alt: true }),
    dataRow(["Rollback rate", "1 in 5.5 deploys", "--", "Low release confidence"]),
    dataRow(["Test flakiness", "6.8% CI runs", "--", { text: "Retry-and-pray culture", options: { color: "FF6B6B" } }], { alt: true }),
  ];

  slide.addTable(rows, makeTableOpts(0.5, 3.1, 9.0, [2.2, 1.8, 1.8, 3.2], { rowH: 0.26 }));
}

// ════════════════════════════════════════════════════════════════
// SLIDE 2: Target Architecture
// ════════════════════════════════════════════════════════════════
function slide02(pres) {
  const slide = pres.addSlide();
  slide.background = { color: theme.bg };

  slide.addText("Target Architecture", {
    x: 0.5, y: 0.2, w: 6, h: 0.5,
    fontSize: 28, fontFace: TITLE_FONT,
    color: "FFFFFF", bold: true,
  });
  slide.addText("Strangler Fig  |  API Gateway  |  Event-Driven  |  Service per Domain", {
    x: 0.5, y: 0.65, w: 8, h: 0.25,
    fontSize: 10, fontFace: BODY_FONT, color: "667788",
  });

  // Service boundary table — left
  sectionLabel(slide, "Service Boundaries", 0.3, 0.95, 5.9);

  const svcRows = [
    headerRow(["Service", "Squad", "DB", "API", "Priority"]),
    dataRow(["User Service", "Identity", "PostgreSQL", "REST + gRPC", { text: "P0", options: { color: theme.accent, bold: true } }]),
    dataRow(["Order Service", "Commerce", "PostgreSQL", "REST + events", { text: "P0", options: { color: theme.accent, bold: true } }], { alt: true }),
    dataRow(["Payment Service", "Commerce", "PG + Stripe", "REST (sync)", "P1"]),
    dataRow(["Search Service", "Discovery", "Elasticsearch", "REST", "P1"], { alt: true }),
    dataRow(["Notification Svc", "Platform", "PG + SQS", "Async events", "P2"]),
    dataRow(["Analytics Svc", "Data", "ClickHouse", "gRPC internal", "P2"], { alt: true }),
  ];

  slide.addTable(svcRows, makeTableOpts(0.3, 1.3, 5.9, [1.4, 1.0, 1.1, 1.2, 1.2], { rowH: 0.3 }));

  // Architecture diagram — right side
  const dx = 6.5, dw = 3.2;
  sectionLabel(slide, "Architecture", dx, 0.95, dw);

  // API Gateway
  slide.addShape("rect", {
    x: dx + 0.3, y: 1.4, w: 2.6, h: 0.4,
    fill: { color: theme.secondary },
    rectRadius: 0.05,
  });
  slide.addText("API Gateway (Kong)", {
    x: dx + 0.3, y: 1.4, w: 2.6, h: 0.4,
    fontSize: 9, fontFace: BODY_FONT,
    color: "FFFFFF", bold: true,
    align: "center", valign: "middle",
  });

  // Arrow
  slide.addShape("rect", {
    x: dx + 1.5, y: 1.8, w: 0.2, h: 0.2,
    fill: { color: theme.accent },
  });

  // Service boxes
  const svcs = [
    { name: "User", color: "0a3d6b" },
    { name: "Order", color: "0a3d6b" },
    { name: "Payment", color: "0a3d6b" },
    { name: "Search", color: "0d2240" },
    { name: "Notif.", color: "0d2240" },
    { name: "Analytics", color: "0d2240" },
  ];

  svcs.forEach((s, i) => {
    const sx = dx + 0.1 + (i % 3) * 1.05;
    const sy = 2.1 + Math.floor(i / 3) * 0.55;
    slide.addShape("rect", {
      x: sx, y: sy, w: 0.95, h: 0.45,
      fill: { color: s.color },
      rectRadius: 0.04,
    });
    slide.addText(s.name, {
      x: sx, y: sy, w: 0.95, h: 0.45,
      fontSize: 8, fontFace: BODY_FONT,
      color: "FFFFFF", align: "center", valign: "middle",
    });
  });

  // Kafka bar
  slide.addShape("rect", {
    x: dx + 0.1, y: 3.25, w: 3.0, h: 0.35,
    fill: { color: "1a1500" },
    rectRadius: 0.04,
  });
  slide.addText("Event Bus (Kafka)", {
    x: dx + 0.1, y: 3.25, w: 3.0, h: 0.35,
    fontSize: 9, fontFace: BODY_FONT,
    color: theme.accent, bold: true,
    align: "center", valign: "middle",
  });

  // Data stores
  slide.addShape("rect", {
    x: dx + 0.1, y: 3.7, w: 3.0, h: 0.35,
    fill: { color: "0d1520" },
    rectRadius: 0.04,
  });
  slide.addText("PostgreSQL x4  |  ES  |  ClickHouse  |  Redis", {
    x: dx + 0.1, y: 3.7, w: 3.0, h: 0.35,
    fontSize: 8, fontFace: BODY_FONT,
    color: "8899AA",
    align: "center", valign: "middle",
  });

  // Key decisions — bottom left
  sectionLabel(slide, "Key Decisions", 0.3, 3.65, 5.9);

  const principles = [
    { text: "Database-per-service", options: { bold: true, color: theme.accent, fontSize: 9, fontFace: BODY_FONT, breakLine: true } },
    { text: "  No shared tables. Each service owns its data. Cross-service queries via events or API calls.", options: { fontSize: 9, fontFace: BODY_FONT, color: "8899AA", breakLine: true } },
    { text: "Event-driven async", options: { bold: true, color: theme.accent, fontSize: 9, fontFace: BODY_FONT, breakLine: true } },
    { text: "  Kafka for inter-service communication. Sync REST only for user-facing checkout flow.", options: { fontSize: 9, fontFace: BODY_FONT, color: "8899AA", breakLine: true } },
    { text: "Strangler Fig pattern", options: { bold: true, color: theme.accent, fontSize: 9, fontFace: BODY_FONT, breakLine: true } },
    { text: "  Incremental extraction behind API gateway. Monolith stays live. Rollback = route back to monolith.", options: { fontSize: 9, fontFace: BODY_FONT, color: "8899AA" } },
  ];

  slide.addText(principles, {
    x: 0.3, y: 4.0, w: 5.9, h: 1.35,
    valign: "top",
    lineSpacingMultiple: 1.05,
  });

  pageBadge(slide, pres, 2);
}

// ════════════════════════════════════════════════════════════════
// SLIDE 3: Why Strangler Fig (ADR)
// ════════════════════════════════════════════════════════════════
function slide03(pres) {
  const slide = pres.addSlide();
  slide.background = { color: theme.bg };

  slide.addText("Architecture Decision Record", {
    x: 0.5, y: 0.2, w: 7, h: 0.5,
    fontSize: 28, fontFace: TITLE_FONT,
    color: "FFFFFF", bold: true,
  });
  slide.addText("Why Strangler Fig, Not Rewrite, Not Modularize", {
    x: 0.5, y: 0.65, w: 8, h: 0.25,
    fontSize: 11, fontFace: BODY_FONT, color: theme.accent,
  });

  // Three evaluation cards
  const evals = [
    {
      title: "FULL REWRITE",
      verdict: "REJECTED",
      verdictColor: "FF6B6B",
      points: [
        "Estimated 14-18 months with 8 engineers (Go + gRPC greenfield)",
        "Dual-system maintenance: monolith + new system in parallel",
        "340K lines of tested business logic must be re-discovered",
        "Industry data: rewrites take 2-3x longer than estimated",
        "Netscape, Basecamp case studies show high failure rate",
      ],
    },
    {
      title: "MODULARIZE MONOLITH",
      verdict: "INSUFFICIENT",
      verdictColor: "FFB347",
      points: [
        "Tried Q4 2025: 6 weeks extracting Django apps into packages",
        "Merge conflicts dropped 30% but deploy frequency unchanged",
        "47-min CI build didn't improve; blast radius didn't shrink",
        "Addresses code organization, not deployment coupling",
        "Deployment coupling is the actual bottleneck",
      ],
    },
    {
      title: "STRANGLER FIG",
      verdict: "SELECTED",
      verdictColor: "4ADE80",
      points: [
        "Incremental extraction: ship value while de-risking",
        "Each service validated in prod alongside monolith before cutover",
        "If extraction fails, route traffic back to monolith instantly",
        "Q1 2026 prototype: User Service extracted in 3 weeks, zero incidents",
        "Risk bounded to one service at a time, not entire system",
      ],
    },
  ];

  evals.forEach((e, i) => {
    const cx = 0.4 + i * 3.1;
    const cw = 2.9;

    slide.addShape("rect", {
      x: cx, y: 1.05, w: cw, h: 2.7,
      fill: { color: "001225" },
      rectRadius: 0.06,
    });

    slide.addText(e.title, {
      x: cx + 0.15, y: 1.12, w: cw - 0.3, h: 0.28,
      fontSize: 10, fontFace: BODY_FONT,
      color: "FFFFFF", bold: true,
      charSpacing: 2,
    });
    slide.addText(e.verdict, {
      x: cx + 0.15, y: 1.38, w: cw - 0.3, h: 0.22,
      fontSize: 9, fontFace: BODY_FONT,
      color: e.verdictColor, bold: true,
    });

    slide.addShape("rect", {
      x: cx + 0.15, y: 1.62, w: cw - 0.3, h: 0.015,
      fill: { color: "1a3a5c" },
    });

    const textParts = e.points.map((p, pi) => ({
      text: p,
      options: {
        bullet: true,
        breakLine: pi < e.points.length - 1,
        fontSize: 8.5,
        fontFace: BODY_FONT,
        color: "B0BCC8",
        paraSpaceAfter: 4,
      },
    }));

    slide.addText(textParts, {
      x: cx + 0.15, y: 1.7, w: cw - 0.3, h: 1.95,
      valign: "top",
    });
  });

  // Bottom: Scope & Confidence table
  sectionLabel(slide, "Scope & Timeline Confidence", 0.4, 3.9, 9.2);

  const scopeRows = [
    headerRow(["Decision", "Detail", "Confidence"]),
    dataRow([
      "Extracting (6 services)",
      "User, Order, Payment, Search, Notification, Analytics",
      { text: "Phase 1-2: 65%  |  Phase 3: 40%", options: { color: theme.accent } },
    ]),
    dataRow([
      { text: "NOT extracting", options: { bold: true } },
      "Reporting, Admin, Billing (low-traffic, low-change; cost > benefit)",
      "Stay in monolith indefinitely",
    ], { alt: true }),
    dataRow([
      "Timeline risk",
      "Phase 3 depends on ClickHouse adoption (unknown learning curve). Padded 2 weeks.",
      { text: "40% on-time for Phase 3", options: { color: "FFB347" } },
    ]),
  ];

  slide.addTable(scopeRows, makeTableOpts(0.4, 4.25, 9.2, [1.8, 4.6, 2.8], { rowH: 0.32 }));

  pageBadge(slide, pres, 3);
}

// ════════════════════════════════════════════════════════════════
// SLIDE 4: Migration Plan
// ════════════════════════════════════════════════════════════════
function slide04(pres) {
  const slide = pres.addSlide();
  slide.background = { color: theme.bg };

  slide.addText("6-Month Migration Plan", {
    x: 0.5, y: 0.2, w: 7, h: 0.5,
    fontSize: 28, fontFace: TITLE_FONT,
    color: "FFFFFF", bold: true,
  });

  // Phase timeline bar
  const phases = [
    { label: "Phase 1: Foundation", months: "M1-M2", w: 3.0, color: theme.secondary },
    { label: "Phase 2: Commerce Core", months: "M3-M4", w: 3.0, color: "0a3d6b" },
    { label: "Phase 3: Platform & Optimize", months: "M5-M6", w: 3.0, color: "0d2240" },
  ];

  let barX = 0.5;
  phases.forEach((p) => {
    slide.addShape("rect", {
      x: barX, y: 0.75, w: p.w, h: 0.35,
      fill: { color: p.color },
      rectRadius: 0.04,
    });
    slide.addText(p.label + "  (" + p.months + ")", {
      x: barX, y: 0.75, w: p.w, h: 0.35,
      fontSize: 9, fontFace: BODY_FONT,
      color: "FFFFFF", bold: true,
      align: "center", valign: "middle",
    });
    barX += p.w + 0.05;
  });

  // Phase 1 table
  sectionLabel(slide, "Phase 1: Foundation (Month 1-2)", 0.3, 1.2, 4.4);

  const p1Rows = [
    headerRow(["Task", "Owner", "Risk"]),
    dataRow(["Deploy API Gateway (Kong)", "SRE", "SPOF - need HA"]),
    dataRow(["Instrument w/ OpenTelemetry", "SRE + all", "~2-3% overhead"], { alt: true }),
    dataRow(["Extract User Service", "Identity", "Session migration"]),
    dataRow(["Set up Kafka (3 brokers)", "SRE", { text: "No Kafka exp.", options: { color: "FF6B6B" } }], { alt: true }),
    dataRow(["DB migration tooling", "DBA", "Data consistency"]),
  ];

  slide.addTable(p1Rows, makeTableOpts(0.3, 1.55, 4.4, [2.0, 1.0, 1.4], { rowH: 0.25 }));

  // Phase 2 table
  sectionLabel(slide, "Phase 2: Commerce Core (Month 3-4)", 5.0, 1.2, 4.7);

  const p2Rows = [
    headerRow(["Task", "Owner", "Risk"]),
    dataRow(["Extract Order Service", "Commerce", { text: "47 cross-imports", options: { color: "FF6B6B" } }]),
    dataRow(["Extract Payment Service", "Commerce", "PCI scope change"], { alt: true }),
    dataRow(["Implement checkout saga", "Comm+Ident", { text: "Complex failures", options: { color: "FF6B6B" } }]),
    dataRow(["Search Service extraction", "Discovery", "4hr reindex"], { alt: true }),
  ];

  slide.addTable(p2Rows, makeTableOpts(5.0, 1.55, 4.7, [2.2, 1.0, 1.5], { rowH: 0.25 }));

  // Phase 3 table
  sectionLabel(slide, "Phase 3: Platform & Optimization (Month 5-6)", 0.3, 3.1, 4.4);

  const p3Rows = [
    headerRow(["Task", "Owner", "Risk"]),
    dataRow(["Notification Service", "Platform", "Low risk"]),
    dataRow(["Analytics Service", "Data", "18mo data migration"], { alt: true }),
    dataRow(["Decommission monolith modules", "All squads", "Residual coupling"]),
    dataRow(["Perf tuning + chaos eng.", "SRE", "Unknown unknowns"], { alt: true }),
  ];

  slide.addTable(p3Rows, makeTableOpts(0.3, 3.45, 4.4, [2.2, 1.0, 1.2], { rowH: 0.25 }));

  // Critical path — right
  sectionLabel(slide, "Critical Path & Dependencies", 5.0, 3.1, 4.7);

  const depItems = [
    { text: "1. API Gateway must be live before any extraction", options: { bold: true, color: theme.accent, fontSize: 9, fontFace: BODY_FONT, breakLine: true } },
    { text: "   Gateway routes traffic; services are useless without it.", options: { fontSize: 8.5, fontFace: BODY_FONT, color: "8899AA", breakLine: true, paraSpaceAfter: 4 } },
    { text: "2. User Service extracted before Order Service", options: { bold: true, color: theme.accent, fontSize: 9, fontFace: BODY_FONT, breakLine: true } },
    { text: "   Orders depend on user identity. Extract dependency first.", options: { fontSize: 8.5, fontFace: BODY_FONT, color: "8899AA", breakLine: true, paraSpaceAfter: 4 } },
    { text: "3. Kafka operational before Commerce extractions", options: { bold: true, color: theme.accent, fontSize: 9, fontFace: BODY_FONT, breakLine: true } },
    { text: "   Order + Payment services communicate via events.", options: { fontSize: 8.5, fontFace: BODY_FONT, color: "8899AA", breakLine: true, paraSpaceAfter: 4 } },
    { text: "4. Checkout saga requires Order + Payment + User", options: { bold: true, color: theme.accent, fontSize: 9, fontFace: BODY_FONT, breakLine: true } },
    { text: "   All three services must be stable before saga goes live.", options: { fontSize: 8.5, fontFace: BODY_FONT, color: "8899AA", breakLine: true, paraSpaceAfter: 4 } },
    { text: "5. Monolith fallback active until Phase 3 complete", options: { bold: true, color: theme.accent, fontSize: 9, fontFace: BODY_FONT, breakLine: true } },
    { text: "   Feature flag per service; route back on failure.", options: { fontSize: 8.5, fontFace: BODY_FONT, color: "8899AA" } },
  ];

  slide.addText(depItems, {
    x: 5.0, y: 3.45, w: 4.7, h: 1.9,
    valign: "top",
  });

  pageBadge(slide, pres, 4);
}

// ════════════════════════════════════════════════════════════════
// SLIDE 5: Risk Matrix + Observability
// ════════════════════════════════════════════════════════════════
function slide05(pres) {
  const slide = pres.addSlide();
  slide.background = { color: theme.bg };

  slide.addText("Risk Matrix & Observability Stack", {
    x: 0.5, y: 0.2, w: 8, h: 0.5,
    fontSize: 28, fontFace: TITLE_FONT,
    color: "FFFFFF", bold: true,
  });

  // Risk matrix
  sectionLabel(slide, "Risk Matrix", 0.3, 0.72, 9.4);

  function severityCell(text) {
    const colors = { CRITICAL: "FF4444", HIGH: "FF8C42", MEDIUM: "FFB347" };
    return { text, options: { color: colors[text] || "FFFFFF", bold: true } };
  }

  const riskRows = [
    headerRow(["Risk", "L", "I", "Sev.", "Mitigation"]),
    dataRow([
      "Distributed txn failures (checkout saga)",
      "H", "H",
      severityCell("CRITICAL"),
      "Compensation/rollback per saga step; feature flag fallback to monolith",
    ]),
    dataRow([
      "Kafka instability (team inexperience)",
      "M", "H",
      severityCell("HIGH"),
      "Hire Kafka specialist 3mo; use Confluent Cloud managed; start with 3 topics",
    ], { alt: true }),
    dataRow([
      "Data inconsistency (dual-write migration)",
      "H", "M",
      severityCell("HIGH"),
      "CDC with Debezium (not app-level dual-write); hourly reconciliation jobs",
    ]),
    dataRow([
      "Service latency exceeds budget",
      "M", "M",
      severityCell("MEDIUM"),
      "P99 budget 200ms/call; Istio circuit breakers 500ms timeout; async-first",
    ], { alt: true }),
    dataRow([
      "Team cognitive overload (new tools)",
      "M", "M",
      severityCell("MEDIUM"),
      "Phase tool adoption: Kafka M1, Istio M3, ClickHouse M5; weekly arch hours",
    ]),
    dataRow([
      "PCI scope expansion (Payment Svc)",
      "L", "H",
      severityCell("MEDIUM"),
      "Security review M2; use Stripe hosted checkout; document compliance bounds",
    ], { alt: true }),
  ];

  slide.addTable(riskRows, makeTableOpts(0.3, 1.05, 9.4, [2.6, 0.35, 0.35, 0.75, 5.35], { rowH: 0.3 }));

  // Observability stack
  sectionLabel(slide, "Monitoring & Observability Stack", 0.3, 3.25, 9.4);

  const obsRows = [
    headerRow(["Layer", "Tool", "Coverage", "Alert Threshold"]),
    dataRow(["Metrics", "Prometheus + Grafana", "RED: rate, error, duration; saturation", "Err > 1% (5m) | P99 > 500ms (10m)"]),
    dataRow(["Tracing", "Jaeger + OpenTelemetry", "End-to-end request flow across services", "Trace > 2s | Span errors > 5%"], { alt: true }),
    dataRow(["Logging", "ELK Stack", "Structured JSON; correlation IDs", "Error rate > 50/min per service"]),
    dataRow(["Health", "K8s probes", "Liveness + readiness; dependency health", "Probe fail = restart; 3x = page SRE"], { alt: true }),
    dataRow(["Synthetic", "Datadog Synthetics", "Login, search, checkout, payment flows", "Any failure = immediate page"]),
    dataRow(["Alerting", "PagerDuty", "P1 page | P2 Slack | P3 next biz day", "Tiered escalation"], { alt: true }),
  ];

  slide.addTable(obsRows, makeTableOpts(0.3, 3.6, 9.4, [1.0, 1.8, 3.2, 3.4], { rowH: 0.28 }));

  // Dashboards footer
  slide.addShape("rect", {
    x: 0.3, y: 5.0, w: 9.4, h: 0.02,
    fill: { color: "1a3a5c" },
  });
  slide.addText([
    { text: "Key Grafana Dashboards: ", options: { bold: true, color: theme.accent, fontSize: 8.5, fontFace: BODY_FONT } },
    { text: "Service Overview (RED per svc)  |  Kafka Health (consumer lag, replication)  |  Checkout Saga (success rate, step failures)  |  Deploy Tracker (canary, rollbacks)", options: { fontSize: 8.5, fontFace: BODY_FONT, color: "8899AA" } },
  ], {
    x: 0.3, y: 5.05, w: 9.4, h: 0.35,
    valign: "top",
  });

  pageBadge(slide, pres, 5);
}

// ════════════════════════════════════════════════════════════════
// SLIDE 6: Success Metrics + Team Structure
// ════════════════════════════════════════════════════════════════
function slide06(pres) {
  const slide = pres.addSlide();
  slide.background = { color: theme.bg };

  slide.addText("Success Metrics & Team Ownership", {
    x: 0.5, y: 0.2, w: 8, h: 0.5,
    fontSize: 28, fontFace: TITLE_FONT,
    color: "FFFFFF", bold: true,
  });

  // Before/After targets
  sectionLabel(slide, "Before / After Targets", 0.3, 0.72, 9.4);

  function targetCell(text, good) {
    return { text, options: { color: good ? "4ADE80" : "FFFFFF" } };
  }

  const metricRows = [
    headerRow(["Metric", "Current", "6-Month Target", "12-Month Target", "Measurement"]),
    dataRow(["Deploy frequency", "1.8/week", targetCell("3/week per svc", false), targetCell("1+/day per svc", true), "CI/CD pipeline"]),
    dataRow(["Lead time", "12.3 days", targetCell("5 days", false), targetCell("< 2 days", true), "PR merge to deploy"], { alt: true }),
    dataRow(["Change failure rate", "18.2%", targetCell("12%", false), targetCell("< 8%", true), "Hotfix / total deploys"]),
    dataRow(["MTTR", "4.2 hours", targetCell("1.5 hours", false), targetCell("< 30 min", true), "PagerDuty incidents"], { alt: true }),
    dataRow(["P99 API latency", "820ms", targetCell("400ms", false), targetCell("< 200ms", true), "Prometheus histograms"]),
    dataRow(["CI build time", "47 min", targetCell("15 min/svc", false), targetCell("< 8 min", true), "GitHub Actions"], { alt: true }),
    dataRow(["Test flakiness", "6.8%", targetCell("3%", false), targetCell("< 1%", true), "CI failure rate"]),
    dataRow(["Dev satisfaction", "5.8/10", targetCell("7.0/10", false), targetCell("8.0/10", true), "Quarterly survey"], { alt: true }),
  ];

  slide.addTable(metricRows, makeTableOpts(0.3, 1.05, 9.4, [1.5, 1.2, 1.5, 1.5, 1.8], { rowH: 0.26 }));

  // Team ownership
  sectionLabel(slide, "Target Team Structure: You Build It, You Run It", 0.3, 3.5, 9.4);

  const teamRows = [
    headerRow(["Squad", "Services Owned", "Size", "On-Call"]),
    dataRow(["Identity", "User Service, Auth", "4 eng", "Own pager"]),
    dataRow(["Commerce", "Order Service, Payment Service", "6 eng", "Own pager"], { alt: true }),
    dataRow(["Discovery", "Search Service", "3 eng", "Shared w/ Platform"]),
    dataRow(["Platform", "Notification Svc, API Gateway, shared libs", "4 eng", "Own pager"], { alt: true }),
    dataRow(["Data", "Analytics Service", "3 eng", "Shared w/ Platform"]),
    dataRow(["SRE", "Kafka, Kubernetes, monitoring, incident response", "3 eng", "Primary escalation"], { alt: true }),
  ];

  slide.addTable(teamRows, makeTableOpts(0.3, 3.85, 9.4, [1.2, 4.0, 1.0, 3.2], { rowH: 0.26 }));

  // Bottom principle
  slide.addShape("rect", {
    x: 0.3, y: 5.1, w: 9.4, h: 0.02,
    fill: { color: "1a3a5c" },
  });
  slide.addText([
    { text: "Principle: ", options: { bold: true, color: theme.accent, fontSize: 9, fontFace: BODY_FONT } },
    { text: "Each squad owns their service SLOs, deploys independently, and carries their own pager. SRE provides platform, not babysitting. Coordination overhead drops from O(n\u00B2) to O(n).", options: { fontSize: 9, fontFace: BODY_FONT, color: "8899AA" } },
  ], {
    x: 0.3, y: 5.15, w: 9.4, h: 0.35,
    valign: "top",
  });

  pageBadge(slide, pres, 6);
}

// ════════════════════════════════════════════════════════════════
// Compile
// ════════════════════════════════════════════════════════════════
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.author = "Project Chimera";
pres.title = "Project Chimera: Monolith Decomposition Plan";

slide01(pres);
slide02(pres);
slide03(pres);
slide04(pres);
slide05(pres);
slide06(pres);

pres
  .writeFile({ fileName: "./outputs/software-minimax.pptx" })
  .then(() => console.log("Created: outputs/software-minimax.pptx"))
  .catch((err) => {
    console.error("Error:", err);
    process.exit(1);
  });
