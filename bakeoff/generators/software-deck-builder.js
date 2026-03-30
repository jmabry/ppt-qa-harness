"use strict";
const pptxgen = require("pptxgenjs");

// ══════════════════════════════════════════════════════════════
// LAYER 1 — CONSTANTS
// ══════════════════════════════════════════════════════════════
const W            = 10;
const H            = 5.625;
const PAD          = 0.5;
const TITLE_H      = 0.5;
const BODY_TOP     = 0.62;   // TITLE_H + 0.12
const FOOTER_Y     = 5.35;
const CONTENT_BOTTOM = 5.23;
const SECTION_GAP  = 0.12;
const MIN_FONT     = 9;

const C = {
  navy:     "0F2744",   // primary dark navy
  teal:     "0EA5E9",   // accent teal/sky
  white:    "FFFFFF",
  offWhite: "F0F4F8",   // light bg
  lightGray:"CBD5E1",
  slate:    "64748B",
  charcoal: "1E293B",
  midNavy:  "1E3A5F",
  darkCard: "162F4A",
  red:      "DC2626",
  orange:   "EA580C",
  amber:    "D97706",
  green:    "059669",
  rowAlt:   "E8EFF6",
  rowBase:  "FFFFFF",
};

const FONT = {
  head: "Arial",
  body: "Calibri",
};

const TOTAL_SLIDES = 6;

// ══════════════════════════════════════════════════════════════
// LAYER 2 — HELPERS
// ══════════════════════════════════════════════════════════════

function addHeader(slide, title, dark = false) {
  // Top bar
  slide.addShape("rect", {
    x: 0, y: 0, w: W, h: 0.52,
    fill: { color: dark ? C.midNavy : C.navy },
  });
  // Teal accent left stripe
  slide.addShape("rect", {
    x: 0, y: 0, w: 0.08, h: 0.52,
    fill: { color: C.teal },
  });
  slide.addText(title, {
    x: PAD, y: 0.04, w: W - PAD - 0.3, h: 0.44,
    fontSize: 15,
    fontFace: FONT.head,
    color: C.white,
    bold: true,
    align: "left",
    valign: "middle",
    margin: 0,
  });
}

function addFooter(slide, num, total) {
  slide.addText(`Project Chimera — Monolith Decomposition Plan`, {
    x: PAD, y: FOOTER_Y, w: 7, h: 0.22,
    fontSize: 7.5,
    fontFace: FONT.body,
    color: C.slate,
    align: "left",
    margin: 0,
  });
  slide.addText(`${num} / ${total}`, {
    x: W - 1.2, y: FOOTER_Y, w: 1.0, h: 0.22,
    fontSize: 7.5,
    fontFace: FONT.body,
    color: C.slate,
    align: "right",
    margin: 0,
  });
  // Footer rule
  slide.addShape("rect", {
    x: 0, y: FOOTER_Y - 0.04, w: W, h: 0.02,
    fill: { color: C.lightGray },
  });
}

// Returns y + 0.24
function addSectionLabel(slide, text, y, dark = false) {
  slide.addText(text.toUpperCase(), {
    x: PAD, y, w: W - PAD * 2, h: 0.22,
    fontSize: 8,
    fontFace: FONT.body,
    color: dark ? C.teal : C.teal,
    bold: true,
    align: "left",
    margin: 0,
  });
  slide.addShape("rect", {
    x: PAD, y: y + 0.2, w: W - PAD * 2, h: 0.02,
    fill: { color: dark ? C.midNavy : C.lightGray },
  });
  return y + 0.24;
}

function addKpiCard(slide, x, y, w, h, value, label, dark = false) {
  slide.addShape("rect", {
    x, y, w, h,
    fill: { color: dark ? C.darkCard : C.white },
    line: { color: dark ? C.midNavy : C.lightGray, pt: 0.75 },
  });
  // Top teal rule
  slide.addShape("rect", {
    x, y, w, h: 0.04,
    fill: { color: C.teal },
  });
  slide.addText(value, {
    x, y: y + 0.06, w, h: h * 0.52,
    fontSize: 22,
    fontFace: FONT.head,
    color: dark ? C.teal : C.navy,
    bold: true,
    align: "center",
    valign: "middle",
    margin: 0,
  });
  slide.addText(label, {
    x: x + 0.06, y: y + h * 0.58, w: w - 0.12, h: h * 0.36,
    fontSize: MIN_FONT,
    fontFace: FONT.body,
    color: dark ? C.lightGray : C.slate,
    align: "center",
    valign: "top",
    margin: 0,
  });
}

// Generic light-bg table builder
// rows: array of arrays; row[0] = header
// colW: array of column widths
function addTable(slide, rows, x, y, w, colW, rowH = 0.28) {
  const tableRows = rows.map((row, rIdx) =>
    row.map((cell, cIdx) => {
      const isObj = typeof cell === "object" && cell !== null;
      const text  = isObj ? cell.text : String(cell);
      const extra = isObj ? (cell.options || {}) : {};
      if (rIdx === 0) {
        return {
          text,
          options: {
            fill: { color: C.navy },
            color: C.white,
            bold: true,
            fontSize: 8.5,
            fontFace: FONT.body,
            align: cIdx === 0 ? "left" : "center",
            valign: "middle",
            margin: [2, 4, 2, 4],
            ...extra,
          },
        };
      }
      return {
        text,
        options: {
          fill: { color: rIdx % 2 === 1 ? C.rowBase : C.rowAlt },
          color: C.charcoal,
          fontSize: 8.5,
          fontFace: FONT.body,
          align: cIdx === 0 ? "left" : "center",
          valign: "middle",
          margin: [2, 4, 2, 4],
          ...extra,
        },
      };
    })
  );

  slide.addTable(tableRows, {
    x, y, w, colW,
    rowH,
    border: { pt: 0.5, color: C.lightGray },
    autoPage: false,
  });
}

// Dark-bg table
function addDarkTable(slide, rows, x, y, w, colW, rowH = 0.3) {
  const tableRows = rows.map((row, rIdx) =>
    row.map((cell, cIdx) => {
      const isObj = typeof cell === "object" && cell !== null;
      const text  = isObj ? cell.text : String(cell);
      const extra = isObj ? (cell.options || {}) : {};
      if (rIdx === 0) {
        return {
          text,
          options: {
            fill: { color: C.teal },
            color: C.navy,
            bold: true,
            fontSize: 8.5,
            fontFace: FONT.body,
            align: cIdx === 0 ? "left" : "center",
            valign: "middle",
            margin: [2, 4, 2, 4],
            ...extra,
          },
        };
      }
      return {
        text,
        options: {
          fill: { color: rIdx % 2 === 1 ? C.darkCard : C.midNavy },
          color: "D0E4F5",
          fontSize: 8.5,
          fontFace: FONT.body,
          align: cIdx === 0 ? "left" : "center",
          valign: "middle",
          margin: [2, 4, 2, 4],
          ...extra,
        },
      };
    })
  );

  slide.addTable(tableRows, {
    x, y, w, colW,
    rowH,
    border: { pt: 0.5, color: C.midNavy },
    autoPage: false,
  });
}

// Callout / info box
function addCallout(slide, text, x, y, w, h, dark = false) {
  slide.addShape("rect", {
    x, y, w, h,
    fill: { color: dark ? "0A1E34" : "EBF5FB" },
    line: { color: C.teal, pt: 1 },
  });
  slide.addShape("rect", {
    x, y, w: 0.06, h,
    fill: { color: C.teal },
  });
  slide.addText(text, {
    x: x + 0.12, y, w: w - 0.18, h,
    fontSize: 8.5,
    fontFace: FONT.body,
    color: dark ? "B0D4F0" : C.charcoal,
    align: "left",
    valign: "middle",
    margin: [4, 4, 4, 4],
  });
}

// ══════════════════════════════════════════════════════════════
// LAYER 3 — SLIDES
// ══════════════════════════════════════════════════════════════

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.author  = "Engineering";
pres.title   = "Project Chimera: Monolith Decomposition Plan";

// ─────────────────────────────────────────────────────────────
// SLIDE 1 — Title + Current State (dark bg)
// ─────────────────────────────────────────────────────────────
{
  const slide = pres.addSlide();
  slide.background = { color: C.navy };

  // Top teal accent band
  slide.addShape("rect", {
    x: 0, y: 0, w: W, h: 0.07,
    fill: { color: C.teal },
  });

  // Title block
  slide.addText("Project Chimera: Monolith Decomposition Plan", {
    x: PAD, y: 0.12, w: W - PAD - 0.3, h: 0.55,
    fontSize: 24,
    fontFace: FONT.head,
    color: C.white,
    bold: true,
    align: "left",
    valign: "middle",
    margin: 0,
  });
  slide.addText("Presented to VP Engineering  |  April 2026", {
    x: PAD, y: 0.66, w: W - PAD * 2, h: 0.28,
    fontSize: 11,
    fontFace: FONT.body,
    color: C.lightGray,
    align: "left",
    margin: 0,
  });

  // KPI cards row
  const kpiY = 1.0;
  const kpiH = 0.74;
  const kpiW = 2.05;
  const kpiGap = 0.1;
  const kpis = [
    ["340K", "Lines of Code"],
    ["2,800", "DB Tables"],
    ["22", "Backend Engineers"],
    ["$133K/qtr", "Wasted Engineering Cost"],
  ];
  kpis.forEach((kpi, i) => {
    addKpiCard(slide, PAD + i * (kpiW + kpiGap), kpiY, kpiW, kpiH, kpi[0], kpi[1], true);
  });

  // DORA table
  let y = kpiY + kpiH + 0.18;
  y = addSectionLabel(slide, "DORA Metrics — Current State", y, true);

  const doraRows = [
    ["Metric", "Current", "Industry P50", "Gap"],
    ["Deploy frequency",   "1.8/wk",   "1/day–1/wk",      "Bottom of \"medium\""],
    ["Lead time",          "12.3 days", "1–7 days",         "2× over ceiling"],
    ["Change failure rate","18.2%",    "0–15%",            "Above threshold"],
    ["MTTR",               "4.2 hours", "<1 hr – <1 day",  "Functional but slow"],
    ["Build time",         "47 min",   "—",                "Devs context-switch"],
  ];
  const doraColW = [2.4, 1.2, 1.6, 2.4];
  const doraW = doraColW.reduce((a, b) => a + b, 0);
  addDarkTable(slide, doraRows, PAD, y, doraW, doraColW, 0.3);

  addFooter(slide, 1, TOTAL_SLIDES);
}

// ─────────────────────────────────────────────────────────────
// SLIDE 2 — Target Architecture (light bg)
// ─────────────────────────────────────────────────────────────
{
  const slide = pres.addSlide();
  slide.background = { color: C.offWhite };
  addHeader(slide, "Target Architecture: Strangler Fig Pattern");

  let y = BODY_TOP;

  // Service boundaries table
  y = addSectionLabel(slide, "Service Boundaries", y);
  const svcRows = [
    ["Service", "Owner Squad", "Database", "API Style", "Priority"],
    ["User Service",         "Identity",  "PostgreSQL",             "REST + gRPC",    "P0"],
    ["Order Service",        "Commerce",  "PostgreSQL",             "REST + Events",  "P0"],
    ["Payment Service",      "Commerce",  "PostgreSQL + Stripe",    "REST",           "P1"],
    ["Search Service",       "Discovery", "Elasticsearch 8",        "REST",           "P1"],
    ["Notification Service", "Platform",  "PostgreSQL + SQS",       "Async",          "P2"],
    ["Analytics Service",    "Data",      "ClickHouse",             "gRPC",           "P2"],
  ];
  const svcColW = [1.7, 1.2, 1.8, 1.4, 0.8];
  const svcW = svcColW.reduce((a, b) => a + b, 0);
  addTable(slide, svcRows, PAD, y, svcW, svcColW, 0.28);
  y += 0.28 * 7 + 0.18;

  // Architecture diagram (text shapes)
  y = addSectionLabel(slide, "Architecture Flow", y);
  const nodes = [
    { label: "API Gateway", x: PAD },
    { label: "Service Mesh (Istio)", x: 2.0 },
    { label: "6 Microservices", x: 4.1 },
    { label: "Kafka", x: 6.2 },
    { label: "Isolated Data Stores", x: 7.6 },
  ];
  const nodeH = 0.36;
  const nodeW = 1.4;
  nodes.forEach((n, i) => {
    slide.addShape("rect", {
      x: n.x, y, w: nodeW, h: nodeH,
      fill: { color: i === 0 ? C.teal : i === 4 ? C.navy : C.midNavy },
      line: { color: C.teal, pt: 0.5 },
    });
    slide.addText(n.label, {
      x: n.x, y, w: nodeW, h: nodeH,
      fontSize: 8.5,
      fontFace: FONT.body,
      color: C.white,
      bold: true,
      align: "center",
      valign: "middle",
    });
    // Arrow
    if (i < nodes.length - 1) {
      slide.addText("→", {
        x: n.x + nodeW + 0.02, y, w: 0.26, h: nodeH,
        fontSize: 10,
        fontFace: FONT.body,
        color: C.teal,
        align: "center",
        valign: "middle",
      });
    }
  });
  y += nodeH + 0.14;

  // Key decisions callout
  addCallout(
    slide,
    "Key Decisions: Kong/AWS ALB for gateway  |  Istio service mesh  |  Kafka 3-broker cluster",
    PAD, y, W - PAD * 2, 0.36
  );

  addFooter(slide, 2, TOTAL_SLIDES);
}

// ─────────────────────────────────────────────────────────────
// SLIDE 3 — ADR: Why Strangler Fig (light bg)
// ─────────────────────────────────────────────────────────────
{
  const slide = pres.addSlide();
  slide.background = { color: C.offWhite };
  addHeader(slide, "ADR: Why Strangler Fig, Not Rewrite");

  let y = BODY_TOP;

  // 3 evaluation cards side-by-side
  const cardW = (W - PAD * 2 - 0.2) / 3;
  const cardH = 1.54;
  const cardGap = 0.1;
  const cards = [
    {
      status: "REJECTED",
      statusColor: C.red,
      title: "Full Rewrite",
      body: "14–18 months, 8 engineers, parallel systems risk.\n\n\"Basecamp, Netscape — rewrites 2–3× slower than estimated.\"",
    },
    {
      status: "INSUFFICIENT",
      statusColor: C.amber,
      title: "Modularize Only",
      body: "Tried Q4 2025. Merge conflicts −30% but deploy frequency unchanged. Still one artifact.",
    },
    {
      status: "SELECTED",
      statusColor: C.green,
      title: "Strangler Fig",
      body: "Incremental, bounded risk. Already proven: User Service extraction Q1 2026 in 3 weeks, zero customer incidents.",
    },
  ];

  cards.forEach((card, i) => {
    const cx = PAD + i * (cardW + cardGap);
    slide.addShape("rect", {
      x: cx, y, w: cardW, h: cardH,
      fill: { color: C.white },
      line: { color: C.lightGray, pt: 0.75 },
    });
    // Status badge
    slide.addShape("rect", {
      x: cx, y, w: cardW, h: 0.28,
      fill: { color: card.statusColor },
    });
    slide.addText(card.status, {
      x: cx, y, w: cardW, h: 0.28,
      fontSize: 9,
      fontFace: FONT.body,
      color: C.white,
      bold: true,
      align: "center",
      valign: "middle",
    });
    slide.addText(card.title, {
      x: cx + 0.1, y: y + 0.3, w: cardW - 0.2, h: 0.3,
      fontSize: 11,
      fontFace: FONT.head,
      color: C.navy,
      bold: true,
      align: "left",
      valign: "middle",
    });
    slide.addText(card.body, {
      x: cx + 0.1, y: y + 0.62, w: cardW - 0.2, h: cardH - 0.72,
      fontSize: MIN_FONT,
      fontFace: FONT.body,
      color: C.charcoal,
      align: "left",
      valign: "top",
      lineSpacingMultiple: 1.2,
    });
  });

  y += cardH + 0.18;

  // Scope callout
  y = addSectionLabel(slide, "Scope Decisions", y);
  addCallout(
    slide,
    "Reporting + Admin stay in monolith — low-traffic, low-change, extraction cost > benefit",
    PAD, y, W - PAD * 2, 0.36
  );
  y += 0.36 + 0.14;

  // Confidence callout
  addCallout(
    slide,
    "Timeline confidence: 65% confident Phase 1–2 on time  |  40% confident Phase 3 — ClickHouse learning curve",
    PAD, y, W - PAD * 2, 0.36
  );

  addFooter(slide, 3, TOTAL_SLIDES);
}

// ─────────────────────────────────────────────────────────────
// SLIDE 4 — Migration Plan (light bg)
// ─────────────────────────────────────────────────────────────
{
  const slide = pres.addSlide();
  slide.background = { color: C.offWhite };
  addHeader(slide, "6-Month Phased Migration");

  let y = BODY_TOP;

  // Phase timeline bar
  const barY = y;
  const barH = 0.38;
  const phases = [
    { label: "Phase 1 — M1-2", color: C.navy, w: 3.0 },
    { label: "Phase 2 — M3-4", color: C.midNavy, w: 3.0 },
    { label: "Phase 3 — M5-6", color: C.teal, w: 3.0 },
  ];
  let bx = PAD;
  phases.forEach((p, i) => {
    slide.addShape("rect", {
      x: bx, y: barY, w: p.w, h: barH,
      fill: { color: p.color },
    });
    slide.addText(p.label, {
      x: bx, y: barY, w: p.w, h: barH,
      fontSize: 10,
      fontFace: FONT.body,
      color: C.white,
      bold: true,
      align: "center",
      valign: "middle",
    });
    bx += p.w + 0.04;
  });
  y += barH + 0.16;

  // Three phase tables side-by-side
  const tblW = (W - PAD * 2 - 0.2) / 3;
  const tblGap = 0.1;
  const taskColW = [1.6, 0.82, 0.6];

  const phaseData = [
    {
      title: "Phase 1: Foundation",
      rows: [
        ["Task", "Owner", "Risk"],
        ["API Gateway setup",      "Platform", "Med"],
        ["OpenTelemetry rollout",  "SRE",      "Low"],
        ["User Svc extraction",    "Identity", "Med"],
        ["Kafka cluster",          "SRE",      "High"],
        ["DB migration tooling",   "Data",     "Med"],
      ],
    },
    {
      title: "Phase 2: Core Services",
      rows: [
        ["Task", "Owner", "Risk"],
        ["Order Service",    "Commerce",  "High"],
        ["Payment Service",  "Commerce",  "High"],
        ["Saga pattern",     "Commerce",  "High"],
        ["Search Service",   "Discovery", "Med"],
        ["—", "—", "—"],
      ],
    },
    {
      title: "Phase 3: Completion",
      rows: [
        ["Task", "Owner", "Risk"],
        ["Notification Svc",       "Platform", "Med"],
        ["Analytics Svc",          "Data",     "High"],
        ["Decommission modules",   "SRE",      "Med"],
        ["Chaos engineering",      "SRE",      "Med"],
        ["—", "—", "—"],
      ],
    },
  ];

  phaseData.forEach((pd, i) => {
    const tx = PAD + i * (tblW + tblGap);
    slide.addText(pd.title, {
      x: tx, y, w: tblW, h: 0.24,
      fontSize: 9,
      fontFace: FONT.body,
      color: C.navy,
      bold: true,
      align: "left",
      margin: 0,
    });
    addTable(slide, pd.rows, tx, y + 0.26, tblW, taskColW, 0.28);
  });

  addFooter(slide, 4, TOTAL_SLIDES);
}

// ─────────────────────────────────────────────────────────────
// SLIDE 5 — Risk Matrix + Observability (light bg)
// ─────────────────────────────────────────────────────────────
{
  const slide = pres.addSlide();
  slide.background = { color: C.offWhite };
  addHeader(slide, "Risk Matrix & Monitoring Stack");

  const leftW  = 5.3;
  const rightW = W - PAD * 2 - leftW - 0.2;
  const leftX  = PAD;
  const rightX = PAD + leftW + 0.2;

  let ly = BODY_TOP;
  let ry = BODY_TOP;

  // LEFT: Risk table
  ly = addSectionLabel(slide, "Risk Matrix", ly);
  const sevColor = (sev) => {
    if (sev === "CRITICAL") return C.red;
    if (sev === "HIGH")     return C.orange;
    if (sev === "MEDIUM")   return C.amber;
    return C.slate;
  };
  const riskRows = [
    ["Risk", "Like.", "Impact", "Sev.", "Mitigation"],
    ["Saga failures",    "HIGH",   "HIGH",   { text: "CRITICAL", options: { color: C.red,    bold: true, align: "center", valign: "middle", fill: { color: rIdx => rIdx % 2 === 1 ? C.rowBase : C.rowAlt } } }, "Compensation logic + feature flag"],
    ["Kafka instability","MEDIUM", "HIGH",   { text: "HIGH",     options: { color: C.orange,  bold: true, align: "center", valign: "middle" } }, "Confluent Cloud + specialist"],
    ["Data inconsist.",  "HIGH",   "MEDIUM", { text: "HIGH",     options: { color: C.orange,  bold: true, align: "center", valign: "middle" } }, "CDC Debezium + hourly reconcile"],
    ["Service latency",  "MEDIUM", "MEDIUM", { text: "MEDIUM",   options: { color: C.amber,   bold: true, align: "center", valign: "middle" } }, "200ms budget + circuit breakers"],
    ["Team overload",    "MEDIUM", "MEDIUM", { text: "MEDIUM",   options: { color: C.amber,   bold: true, align: "center", valign: "middle" } }, "Phased tools + office hours"],
    ["PCI scope",        "LOW",    "HIGH",   { text: "MEDIUM",   options: { color: C.amber,   bold: true, align: "center", valign: "middle" } }, "Sec review M2 + Stripe checkout"],
  ];
  const riskColW = [1.3, 0.5, 0.55, 0.55, 2.2];

  // Rebuild rows without the lambda (fix: severity color directly)
  const riskRowsClean = [
    ["Risk", "Like.", "Impact", "Sev.", "Mitigation"],
  ];
  const riskData = [
    ["Saga failures",    "HIGH",   "HIGH",   "CRITICAL", "Compensation + feature flag"],
    ["Kafka instability","MEDIUM", "HIGH",   "HIGH",     "Confluent Cloud + specialist"],
    ["Data inconsist.",  "HIGH",   "MEDIUM", "HIGH",     "Debezium CDC + hourly recon"],
    ["Service latency",  "MEDIUM", "MEDIUM", "MEDIUM",   "200ms budget + circuit break"],
    ["Team overload",    "MEDIUM", "MEDIUM", "MEDIUM",   "Phased tools + office hours"],
    ["PCI scope",        "LOW",    "HIGH",   "MEDIUM",   "Sec review M2 + Stripe"],
  ];
  riskData.forEach((r, ri) => {
    riskRowsClean.push([
      r[0],
      r[1],
      r[2],
      { text: r[3], options: { color: sevColor(r[3]), bold: true, align: "center", valign: "middle", fill: { color: ri % 2 === 0 ? C.rowBase : C.rowAlt } } },
      r[4],
    ]);
  });

  addTable(slide, riskRowsClean, leftX, ly, leftW, riskColW, 0.295);
  ly += 0.295 * 7 + 0.14;

  // RIGHT: Observability table
  ry = addSectionLabel(slide, "Observability Stack", ry);
  const obsRows = [
    ["Layer", "Tool", "Coverage", "Alert Threshold"],
    ["Metrics",   "Prometheus + Grafana", "RED metrics",       "Error >1% / 5min"],
    ["Tracing",   "Jaeger + OTEL",        "End-to-end flow",   "Trace >2s"],
    ["Logging",   "ELK Stack",            "Structured JSON",   "Error >50/min"],
    ["Health",    "K8s probes",           "Svc availability",  "3 fail → page"],
    ["Synthetic", "Datadog",              "Critical flows",    "Any fail → page"],
  ];
  const obsColW = [0.75, 1.5, 1.3, 1.45];
  addTable(slide, obsRows, rightX, ry, rightW, obsColW, 0.33);

  addFooter(slide, 5, TOTAL_SLIDES);
}

// ─────────────────────────────────────────────────────────────
// SLIDE 6 — Success Metrics + Team (light bg)
// ─────────────────────────────────────────────────────────────
{
  const slide = pres.addSlide();
  slide.background = { color: C.offWhite };
  addHeader(slide, "Success Metrics & Squad Structure");

  const leftW  = 5.4;
  const rightW = W - PAD * 2 - leftW - 0.2;
  const leftX  = PAD;
  const rightX = PAD + leftW + 0.2;

  let ly = BODY_TOP;
  let ry = BODY_TOP;

  // LEFT: Before/after metrics
  ly = addSectionLabel(slide, "Before / After Targets", ly);
  const metricsRows = [
    ["Metric", "Current", "6-Month", "12-Month"],
    ["Deploy frequency",    "1.8/wk",   "3/wk/svc",  "1+/day/svc"],
    ["Lead time",           "12.3 days","5 days",    "<2 days"],
    ["Change failure rate", "18.2%",    "12%",       "<8%"],
    ["MTTR",                "4.2 hrs",  "1.5 hrs",   "<30 min"],
    ["P99 API latency",     "820ms",    "400ms",     "<200ms"],
    ["Build time",          "47 min",   "15 min",    "<8 min"],
    ["Test flakiness",      "6.8%",     "3%",        "<1%"],
    ["Dev satisfaction",    "5.8/10",   "7.0/10",    "8.0/10"],
  ];
  const metColW = [1.7, 0.88, 0.88, 0.95];
  addTable(slide, metricsRows, leftX, ly, leftW, metColW, 0.27);
  ly += 0.27 * 9 + 0.14;

  // RIGHT: Squad ownership
  ry = addSectionLabel(slide, "Squad Ownership", ry);
  const squadRows = [
    ["Squad", "Services", "Size", "On-Call"],
    ["Identity",  "User, Auth",          "4", "Yes"],
    ["Commerce",  "Order, Payment",       "6", "Yes"],
    ["Discovery", "Search",               "3", "Shared"],
    ["Platform",  "Notifications, GW",    "4", "Yes"],
    ["Data",      "Analytics",            "3", "Shared"],
    ["SRE",       "Kafka, K8s, Monitor",  "3", "Yes (primary)"],
  ];
  const squadColW = [0.9, 1.5, 0.45, 1.05];
  addTable(slide, squadRows, rightX, ry, rightW, squadColW, 0.3);
  ry += 0.3 * 7 + 0.14;

  // Closing callout right-side
  addCallout(
    slide,
    "SRE squad is primary on-call for all infrastructure. Service squads escalate Kafka/K8s issues to SRE first.",
    rightX, ry, rightW, 0.5
  );

  addFooter(slide, 6, TOTAL_SLIDES);
}

// ══════════════════════════════════════════════════════════════
// WRITE FILE
// ══════════════════════════════════════════════════════════════
pres.writeFile({ fileName: "outputs/software-deck-builder.pptx" })
  .then(() => console.log("✓ Wrote outputs/software-deck-builder.pptx"))
  .catch((err) => { console.error(err); process.exit(1); });
