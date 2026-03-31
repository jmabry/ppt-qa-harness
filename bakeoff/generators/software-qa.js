"use strict";
const PptxGenJS = require("pptxgenjs");

// ─── PALETTE ─────────────────────────────────────────────────────────────────
const C = {
  navy:     "0D1B2A",
  navyMid:  "12263A",
  teal:     "0A7E8C",
  tealLt:   "14B8C8",
  ice:      "D6EEF2",
  slate:    "37474F",
  slateD:   "263238",   // darker slate for subtitles on light bg
  muted:    "607D8B",
  offwhite: "F4F8FA",
  white:    "FFFFFF",
  cream:    "EEF4F6",
  red:      "C62828",
  orange:   "E65100",
  amber:    "F57C00",
  green:    "2E7D32",
  border:   "B0BEC5",
  darkCard: "0F2638",
  dark2:    "162B3A",
};

const SW  = 10.0;
const SH  = 5.625;
const LM  = 0.28;
const BM  = 0.30;
const TW  = 9.20;   // table / content width — right edge = 9.48, leaving 0.52" margin

// ─── HELPERS ─────────────────────────────────────────────────────────────────
function makeShadow() {
  return { type: "outer", color: "000000", opacity: 0.10, blur: 5, offset: 2, angle: 135 };
}

function hdr(s, title, sub, dark) {
  s.background = { color: dark ? C.navy : C.offwhite };
  s.addShape("rect", { x: 0, y: 0, w: 0.18, h: SH, fill: { color: C.teal }, line: { color: C.teal, width: 0 } });
  s.addText(title, {
    x: LM, y: 0.17, w: TW, h: 0.52,
    fontSize: 19, bold: true, color: dark ? C.white : C.navy,
    fontFace: "Calibri", margin: 0, valign: "middle"
  });
  if (sub) {
    s.addText(sub, {
      x: LM, y: 0.68, w: TW, h: 0.22,
      fontSize: 9, color: dark ? C.tealLt : C.slateD,  // darker on light bg
      fontFace: "Calibri", margin: 0, italic: true
    });
  }
  s.addShape("rect", {
    x: LM, y: 0.93, w: TW, h: 0.025,
    fill: { color: dark ? C.teal : C.border },
    line: { color: dark ? C.teal : C.border, width: 0 }
  });
  return 0.97;
}

// colW must sum to TW (9.20)
function tbl(s, rows, y, colW, hBg, rH, dark, fontSize) {
  const fs = fontSize || 9;
  const data = rows.map((row, ri) => row.map((cell) => {
    const isH  = ri === 0;
    const even = ri % 2 === 0;
    let bg   = isH ? hBg : (dark ? (even ? C.darkCard : C.dark2) : (even ? C.white : C.cream));
    let fg   = isH ? C.white : (dark ? C.white : C.slate);
    let bold = isH;
    const cs = String(cell);
    if (!isH) {
      if      (cs === "CRITICAL") { bg = C.red;    fg = C.white; bold = true; }
      else if (cs === "HIGH")     { bg = "BF360C"; fg = C.white; bold = true; }
      else if (cs === "MEDIUM")   { bg = C.amber;  fg = C.white; bold = true; }
      else if (cs === "LOW")      { bg = C.green;  fg = C.white; bold = true; }
      else if (cs === "P0")       { bg = C.red;    fg = C.white; bold = true; }
      else if (cs === "P1")       { bg = C.orange; fg = C.white; bold = true; }
      else if (cs === "P2")       { bg = C.green;  fg = C.white; bold = true; }
    }
    const bColor = dark ? "263238" : C.border;
    return {
      text: cs,
      options: {
        fill: { color: bg }, color: fg, bold,
        fontSize: fs, fontFace: "Calibri",
        align: "left", valign: "middle",
        margin: [2, 5, 2, 5],
        border: [
          { type: "solid", pt: 0.5, color: bColor },
          { type: "solid", pt: 0.5, color: bColor },
          { type: "solid", pt: 0.5, color: bColor },
          { type: "solid", pt: 0.5, color: bColor },
        ]
      }
    };
  }));
  s.addTable(data, { x: LM, y, w: colW.reduce((a, b) => a + b, 0), colW, rowH: rH });
}

function sec(s, text, y, color) {
  s.addText(text, {
    x: LM, y, w: TW, h: 0.20,
    fontSize: 8.5, bold: true, color: color || C.teal,
    fontFace: "Calibri", margin: 0
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
//  SLIDE 1 — Title + Current State
// ═══════════════════════════════════════════════════════════════════════════════
function slide1(pres) {
  const s = pres.addSlide();
  s.background = { color: C.navy };
  s.addShape("rect", { x: 0, y: 0, w: 0.18, h: SH, fill: { color: C.teal }, line: { color: C.teal, width: 0 } });

  s.addText("PROJECT CHIMERA", {
    x: LM, y: 0.14, w: TW, h: 0.42,
    fontSize: 28, bold: true, color: C.white, fontFace: "Calibri",
    charSpacing: 4, margin: 0
  });
  s.addText("Monolith Decomposition Plan  ·  Atlas → Microservices  ·  VP Engineering Review", {
    x: LM, y: 0.55, w: TW, h: 0.24,
    fontSize: 9.5, color: C.tealLt, fontFace: "Calibri", margin: 0, italic: true
  });
  s.addShape("rect", {
    x: LM, y: 0.83, w: TW, h: 0.025,
    fill: { color: C.teal }, line: { color: C.teal, width: 0 }
  });

  // Stat blocks
  const stats = [
    ["340K",   "Lines of Python"],
    ["2,800",  "DB tables (shared)"],
    ["14",     "Django apps"],
    ["22",     "Backend engineers"],
    ["47 min", "CI build time"],
    ["$133K",  "Waste / quarter"],
  ];
  const bw  = 1.47;
  const bh  = 0.70;
  const by  = 0.88;
  const gap = (TW - 6 * bw) / 5;
  stats.forEach((st, i) => {
    const x = LM + i * (bw + gap);
    s.addShape("rect", {
      x, y: by, w: bw, h: bh,
      fill: { color: C.darkCard }, line: { color: C.teal, width: 0.75 },
      shadow: makeShadow()
    });
    s.addText(st[0], {
      x: x + 0.05, y: by + 0.06, w: bw - 0.10, h: 0.34,
      fontSize: 20, bold: true, color: C.tealLt, fontFace: "Calibri",
      align: "center", margin: 0
    });
    s.addText(st[1], {
      x: x + 0.05, y: by + 0.41, w: bw - 0.10, h: 0.22,
      fontSize: 9.5, color: C.white, fontFace: "Calibri",
      align: "center", margin: 0
    });
  });

  // DORA table
  const tableY = 1.68;
  sec(s, "DORA METRICS vs INDUSTRY BENCHMARKS (MEASURED)", tableY - 0.18, C.tealLt);
  const doraRows = [
    ["Metric", "Current", "DORA P50", "Gap", "Status"],
    ["Deploy frequency",        "1.8 / week",        "1/day – 1/week",   "Bottom of 'Medium'",       "MEDIUM"],
    ["Lead time (commit→prod)", "12.3 days",          "1 day – 1 week",   "2× over 'Medium' ceiling", "HIGH"],
    ["Change failure rate",     "18.2%",              "0 – 15%",          "Above 'Medium' threshold", "HIGH"],
    ["MTTR",                    "4.2 hours",          "< 1 hr – < 1 day", "Functional but slow",      "MEDIUM"],
    ["CI build time",           "47 min",             "—",                "Devs context-switch",      "HIGH"],
    ["Merge conflicts / week",  "3.2 avg (→8 peak)",  "—",                "Cross-team coord. tax",    "MEDIUM"],
    ["Rollback rate",           "1 in 5.5 deploys",   "—",                "Low release confidence",   "HIGH"],
    ["Test flakiness",          "6.8% CI runs",       "—",                "'Retry and pray' culture", "MEDIUM"],
  ];
  // Metric:2.00, Current:1.30, DORA P50:1.25, Gap:2.75, Status:1.90 = 9.20
  tbl(s, doraRows, tableY, [2.00, 1.30, 1.25, 2.75, 1.90], C.teal, 0.255, true, 9);

  // Cost callout — 0.30" from bottom
  const calloutY = SH - BM - 0.30;
  s.addShape("rect", {
    x: LM, y: calloutY, w: TW, h: 0.30,
    fill: { color: C.darkCard }, line: { color: C.orange, width: 1.5 }
  });
  s.addText(
    "Cost estimate:  ~1,400 engineer-hours/quarter lost to monolith friction  ·  $95/hr loaded = $133K/quarter wasted",
    {
      x: LM + 0.16, y: calloutY, w: TW - 0.32, h: 0.30,
      fontSize: 10, bold: true, color: C.white, fontFace: "Calibri",
      margin: 0, valign: "middle"
    }
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
//  SLIDE 2 — Target Architecture
// ═══════════════════════════════════════════════════════════════════════════════
function slide2(pres) {
  const s = pres.addSlide();
  const y0 = hdr(s, "Target Architecture — Strangler Fig Service Boundaries",
    "Incremental extraction behind API Gateway · No big-bang rewrite · Domain-driven ownership", false);

  sec(s, "SERVICE OWNERSHIP MAP", y0 + 0.02);
  const svcRows = [
    ["Service", "Owner Squad", "Data Store", "API Style", "Priority"],
    ["User Service",         "Identity",  "PostgreSQL (isolated)",  "REST + gRPC",           "P0"],
    ["Order Service",        "Commerce",  "PostgreSQL (isolated)",  "REST + events",         "P0"],
    ["Payment Service",      "Commerce",  "PostgreSQL + Stripe",    "REST (sync checkout)",  "P1"],
    ["Search Service",       "Discovery", "Elasticsearch 8",        "REST",                  "P1"],
    ["Notification Service", "Platform",  "PostgreSQL + SQS",       "Async (event-driven)",  "P2"],
    ["Analytics Service",    "Data",      "ClickHouse",             "gRPC (internal only)",  "P2"],
  ];
  // Service:1.90, Owner:1.25, Store:2.20, API:2.35, Priority:0.50 = 8.20... need to sum to 9.20
  // Service:2.00, Owner:1.30, Store:2.25, API:2.25, Priority:1.40 = 9.20
  tbl(s, svcRows, y0 + 0.22, [2.00, 1.30, 2.25, 2.25, 1.40], C.navyMid, 0.27, false, 9);

  // Architecture diagram
  const diagY = y0 + 0.22 + 0.27 * 7 + 0.12;

  s.addShape("rect", {
    x: LM, y: diagY, w: TW, h: 0.24,
    fill: { color: C.ice }, line: { color: C.border, width: 0.5 }
  });
  s.addText("Clients — Web · Mobile · API", {
    x: LM, y: diagY, w: TW, h: 0.24,
    fontSize: 9, bold: true, color: C.slate, fontFace: "Calibri",
    align: "center", valign: "middle", margin: 0
  });

  const arr1Y = diagY + 0.24;
  s.addShape("line", { x: SW / 2, y: arr1Y, w: 0, h: 0.10, line: { color: C.teal, width: 1.5 } });

  const gwY = arr1Y + 0.10;
  s.addShape("rect", {
    x: 3.20, y: gwY, w: 3.60, h: 0.28,
    fill: { color: C.teal }, line: { color: C.teal, width: 0 }, shadow: makeShadow()
  });
  s.addText("API Gateway (Kong / AWS ALB)  ·  Rate limiting · Auth · Circuit breaker", {
    x: 3.20, y: gwY, w: 3.60, h: 0.28,
    fontSize: 9, bold: true, color: C.white, fontFace: "Calibri",
    align: "center", valign: "middle", margin: 0
  });

  const arr2Y = gwY + 0.28;
  s.addShape("line", { x: SW / 2, y: arr2Y, w: 0, h: 0.10, line: { color: C.teal, width: 1.5 } });

  // Service mesh
  const meshY = arr2Y + 0.10;
  const meshH = 0.65;
  s.addShape("rect", {
    x: LM, y: meshY, w: TW, h: meshH,
    fill: { color: C.cream }, line: { color: C.border, width: 0.75 }
  });
  // Label top-right, well inside boundary
  s.addText("Service Mesh — Istio", {
    x: LM + TW - 1.70, y: meshY + 0.04, w: 1.60, h: 0.16,
    fontSize: 8, bold: true, color: C.muted, fontFace: "Calibri",
    align: "right", margin: 0
  });

  const svcW  = 1.90;
  const svcH  = 0.35;
  const svcGap = (TW - 3 * svcW) / 4;
  const topSvcs = [
    { label: "User Service",    col: C.navyMid },
    { label: "Order Service",   col: C.navyMid },
    { label: "Payment Service", col: C.navyMid },
  ];
  topSvcs.forEach((sv, i) => {
    const sx = LM + svcGap + i * (svcW + svcGap);
    s.addShape("rect", { x: sx, y: meshY + 0.18, w: svcW, h: svcH,
      fill: { color: sv.col }, line: { color: C.teal, width: 0.75 }, shadow: makeShadow() });
    s.addText(sv.label, { x: sx, y: meshY + 0.18, w: svcW, h: svcH,
      fontSize: 9, bold: true, color: C.white, fontFace: "Calibri",
      align: "center", valign: "middle", margin: 0 });
  });

  const kafkaY = meshY + meshH + 0.06;
  s.addShape("rect", {
    x: LM, y: kafkaY, w: TW, h: 0.27,
    fill: { color: "1A237E" }, line: { color: "3949AB", width: 0.75 }
  });
  s.addText("Event Bus — Kafka   ·   Topics:  user.*     order.*     payment.*", {
    x: LM, y: kafkaY, w: TW, h: 0.27,
    fontSize: 9.5, bold: true, color: C.white, fontFace: "Calibri",
    align: "center", valign: "middle", margin: 0
  });

  const botSvcs = [
    { label: "Search Service",       col: "1B5E20" },
    { label: "Notification Service", col: "4A148C" },
    { label: "Analytics Service",    col: "BF360C" },
  ];
  const botY = kafkaY + 0.27 + 0.06;
  botSvcs.forEach((sv, i) => {
    const sx = LM + svcGap + i * (svcW + svcGap);
    s.addShape("rect", { x: sx, y: botY, w: svcW, h: 0.32,
      fill: { color: sv.col }, line: { color: C.border, width: 0.5 }, shadow: makeShadow() });
    s.addText(sv.label, { x: sx, y: botY, w: svcW, h: 0.32,
      fontSize: 9, bold: true, color: C.white, fontFace: "Calibri",
      align: "center", valign: "middle", margin: 0 });
  });

  const dsY = botY + 0.32 + 0.06;
  if (dsY + 0.24 < SH - 0.12) {
    s.addShape("rect", { x: LM, y: dsY, w: TW, h: 0.24,
      fill: { color: C.darkCard }, line: { color: C.border, width: 0.5 } });
    s.addText("Isolated Data Stores:  PostgreSQL ×4   ·   Elasticsearch 8   ·   ClickHouse   ·   Redis (per-svc)", {
      x: LM, y: dsY, w: TW, h: 0.24,
      fontSize: 10, color: C.tealLt, fontFace: "Calibri",
      align: "center", valign: "middle", margin: 0
    });
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
//  SLIDE 3 — Why Strangler Fig (ADR)
// ═══════════════════════════════════════════════════════════════════════════════
function slide3(pres) {
  const s = pres.addSlide();
  const y0 = hdr(s,
    "Architecture Decision Record — Why Strangler Fig, Not Rewrite",
    "2-week evaluation spike Apr 14–25, 2026  ·  Three approaches evaluated  ·  Decision: incremental extraction",
    true
  );

  const cards = [
    {
      title: "Option A — Full Rewrite",
      verdict: "REJECTED",
      vcolor: C.red,
      body: [
        "14–18 months, 8-person team in greenfield Go + gRPC",
        "Two systems run in parallel — monolith still maintained",
        "340K LOC = 340K edge cases to rediscover",
        "Basecamp / Netscape: rewrites take 2–3× longer; usually fail",
      ]
    },
    {
      title: "Option B — Modularise Monolith",
      verdict: "TRIED — FAILED",
      vcolor: C.orange,
      body: [
        "Atlas Modular initiative, Q4 2025 — 6 weeks invested",
        "Merge conflicts ↓30%; deploy frequency: no change",
        "Still one artifact: 47 min CI, same blast radius",
        "Fixes code org, not deployment coupling — wrong target",
      ]
    },
    {
      title: "Option C — Strangler Fig",
      verdict: "CHOSEN",
      vcolor: C.green,
      body: [
        "Each service runs in prod alongside monolith before cut-over",
        "Traffic instantly re-routable to monolith if extraction fails",
        "Proof: User Service prototype Q1 2026 — 3 wks, zero incidents",
        "Incremental value shipped throughout; no feature freeze",
      ]
    },
  ];

  const cardW = (TW - 0.20) / 3;
  const cardH = 2.44;
  const cardY = y0 + 0.06;

  cards.forEach((c, i) => {
    const cx = LM + i * (cardW + 0.10);
    s.addShape("rect", { x: cx, y: cardY, w: cardW, h: cardH,
      fill: { color: C.darkCard }, line: { color: C.teal, width: 0.5 }, shadow: makeShadow() });
    // Header band
    s.addShape("rect", { x: cx, y: cardY, w: cardW, h: 0.30,
      fill: { color: C.navyMid }, line: { color: C.teal, width: 0 } });
    s.addText(c.title, { x: cx + 0.08, y: cardY + 0.04, w: cardW - 0.16, h: 0.24,
      fontSize: 9, bold: true, color: C.white, fontFace: "Calibri", margin: 0, valign: "middle" });
    // Verdict badge
    s.addShape("rect", { x: cx, y: cardY + 0.30, w: cardW, h: 0.24,
      fill: { color: c.vcolor }, line: { color: c.vcolor, width: 0 } });
    s.addText(c.verdict, { x: cx, y: cardY + 0.30, w: cardW, h: 0.24,
      fontSize: 9.5, bold: true, color: C.white, fontFace: "Calibri",
      align: "center", valign: "middle", margin: 0, charSpacing: 1 });
    // 4 bullets at 0.43 spacing
    c.body.forEach((line, li) => {
      s.addText("• " + line, {
        x: cx + 0.10, y: cardY + 0.58 + li * 0.43,
        w: cardW - 0.18, h: 0.41,
        fontSize: 9, color: C.ice, fontFace: "Calibri", margin: 0, valign: "top"
      });
    });
  });

  // Scope & confidence
  const scopeY = cardY + cardH + 0.10;
  s.addShape("rect", { x: LM, y: scopeY, w: TW, h: 0.025,
    fill: { color: C.teal }, line: { color: C.teal, width: 0 } });
  sec(s, "SCOPE BOUNDARIES & TIMELINE CONFIDENCE", scopeY + 0.05, C.tealLt);

  const notRows = [
    ["Decision", "Rationale"],
    ["Reporting & Admin modules stay in monolith",  "Low-traffic, low-change. Extraction cost exceeds benefit."],
    ["Goal is NOT 'zero monolith'",                 "Remove modules causing deployment friction for high-frequency squads."],
    ["Phase 1–2 confidence: 65%",                   "User, Order, Payment, Search extractions well-scoped from Q1 2026 prototype."],
    ["Phase 3 confidence: 40%",                     "Analytics/ClickHouse learning curve unknown. 2-week pad built in; may extend."],
  ];
  tbl(s, notRows, scopeY + 0.25, [3.20, 6.00], C.navyMid, 0.255, true, 9);
}

// ═══════════════════════════════════════════════════════════════════════════════
//  SLIDE 4 — Migration Plan (6-Month Phased)
// ═══════════════════════════════════════════════════════════════════════════════
function slide4(pres) {
  const s = pres.addSlide();
  const y0 = hdr(s, "Migration Plan — 6-Month Phased Rollout",
    "Strangler Fig pattern · Each extraction validated in production before cut-over · Rollback at gateway level", false);

  function phaseHdr(label, months, color, y) {
    s.addShape("rect", { x: LM, y, w: TW, h: 0.24, fill: { color }, line: { color, width: 0 } });
    s.addText(label + "   " + months, {
      x: LM + 0.10, y, w: TW - 0.20, h: 0.24,
      fontSize: 12, bold: true, color: C.white, fontFace: "Calibri", margin: 0, valign: "middle"
    });
  }

  const gap  = 0.06;
  const rH   = 0.245;
  const hdrH = 0.24;
  // Task:2.25, Owner:1.25, Deliverable:2.85, Risk:2.85 = 9.20
  const colW = [2.25, 1.25, 2.85, 2.85];

  // Phase 1
  const p1y = y0 + 0.04;
  phaseHdr("PHASE 1 — Foundation", "Month 1–2", C.teal, p1y);
  const p1rows = [
    ["Task", "Owner", "Deliverable", "Risk"],
    ["Deploy API Gateway (Kong); monolith as upstream",   "SRE",          "All traffic through gateway; HA config day-1",  "Gateway SPOF — need HA"],
    ["Instrument monolith: OpenTelemetry on HTTP+Celery", "SRE + squads", "Distributed tracing live; ~2–3% perf overhead", "Prod perf overhead"],
    ["Extract User Service (auth, permissions, CRUD)",    "Identity",     "User svc behind gateway; Cookie↔JWT dual-mode", "Session migration"],
    ["Deploy Kafka cluster (3 brokers); user.* events",   "SRE",          "Event bus live; Kafka onboarding week 1",        "No Kafka experience"],
    ["DB split tooling: Debezium CDC + reconciliation",   "DBA",          "Per-service schemas; hourly recon jobs",         "Data consistency"],
  ];
  tbl(s, p1rows, p1y + hdrH, colW, C.navyMid, rH, false, 8.5);

  // Phase 2
  const p2y = p1y + hdrH + rH * 6 + gap;
  phaseHdr("PHASE 2 — Commerce Core", "Month 3–4", C.orange, p2y);
  const p2rows = [
    ["Task", "Owner", "Deliverable", "Risk"],
    ["Extract Order Service (orders, carts, pricing)",   "Commerce",         "47 cross-module imports untangled; event-driven", "Highest-coupling svc"],
    ["Extract Payment Service (Stripe, invoices)",       "Commerce",         "PCI scope reviewed; Stripe hosted checkout",      "PCI scope changes"],
    ["Checkout saga: create→reserve→charge",             "Commerce + Ident.","Compensation per step; monolith fallback flag",   "Distributed txn risk"],
    ["Extract Search Service (Elasticsearch)",           "Discovery",        "Zero-downtime reindex; 4-hr rebuild planned",     "Index rebuild time"],
  ];
  tbl(s, p2rows, p2y + hdrH, colW, C.navyMid, rH, false, 8.5);

  // Phase 3
  const p3y = p2y + hdrH + rH * 5 + gap;
  phaseHdr("PHASE 3 — Platform & Optimization", "Month 5–6", "2E7D32", p3y);
  const p3rows = [
    ["Task", "Owner", "Deliverable", "Risk"],
    ["Extract Notification Service (email/SMS/push)",    "Platform",   "SQS-backed templates; already loosely coupled", "Low — minimal coupling"],
    ["Extract Analytics Service (ClickHouse + Kafka)",   "Data",       "18 months historical events migrated",           "ClickHouse ramp-up time"],
    ["Decommission monolith modules; thin orchestration","All squads", "Extracted code removed; shared utils audited",   "Residual coupling"],
    ["Chaos engineering + perf tuning; SRE runbooks",    "SRE",        "Circuit breakers tested; latency targets met",   "Unknown failure modes"],
  ];
  tbl(s, p3rows, p3y + hdrH, colW, C.navyMid, rH, false, 8.5);
}

// ═══════════════════════════════════════════════════════════════════════════════
//  SLIDE 5 — Risk Matrix + Monitoring
// ═══════════════════════════════════════════════════════════════════════════════
function slide5(pres) {
  const s = pres.addSlide();
  const y0 = hdr(s, "Risk Matrix & Observability Stack",
    "Mitigations defined before extraction begins · RED + USE metrics per service · PagerDuty tiered escalation", false);

  sec(s, "RISK REGISTER", y0 + 0.02);
  const riskRows = [
    ["Risk", "Likelihood", "Impact", "Severity", "Mitigation"],
    ["Distributed txn failures (checkout saga)",     "HIGH",   "HIGH",   "CRITICAL", "Compensation per step; monolith fallback flag; integration test suite"],
    ["Kafka cluster instability (no prior exp.)",    "MEDIUM", "HIGH",   "HIGH",     "Managed Kafka (Confluent Cloud); specialist M1–M3; start with 3 topics"],
    ["Data inconsistency during dual-write window",  "HIGH",   "MEDIUM", "HIGH",     "Debezium CDC; hourly reconciliation jobs; rollback procedure documented"],
    ["Service latency exceeds P99 budget (200ms)",   "MEDIUM", "MEDIUM", "MEDIUM",   "Istio circuit breakers (500ms); prefer async events over sync calls"],
    ["Team overload from too many new tools",        "MEDIUM", "MEDIUM", "MEDIUM",   "Phased: Kafka M1 · Istio M3 · ClickHouse M5; weekly arch office hours"],
    ["PCI scope expansion (Payment extraction)",     "LOW",    "HIGH",   "MEDIUM",   "Security review M2 before extraction; Stripe hosted checkout limits PCI surface"],
  ];
  // Risk:2.40, Likelihood:0.88, Impact:0.82, Severity:0.90, Mitigation:4.20 = 9.20
  tbl(s, riskRows, y0 + 0.22, [2.40, 0.88, 0.82, 0.90, 4.20], C.navyMid, 0.27, false, 9);

  const monY = y0 + 0.22 + 0.27 * 7 + 0.10;
  sec(s, "OBSERVABILITY STACK", monY);
  const monRows = [
    ["Layer", "Tool", "Coverage", "Alert Threshold"],
    ["Metrics",   "Prometheus + Grafana",    "RED (rate/error/duration) + USE (CPU, memory, connections)",  "Error > 1% / 5 min · P99 > 500ms / 10 min"],
    ["Tracing",   "Jaeger + OpenTelemetry",  "End-to-end request flow; latency breakdown per hop",          "Trace > 2s · Span error rate > 5%"],
    ["Logging",   "ELK Stack",               "Structured JSON logs; correlation IDs across services",        "Error rate > 50 / min per service"],
    ["Health",    "K8s liveness + readiness","Service, DB, Redis, Kafka dependency availability",            "3 consecutive failures → page SRE"],
    ["Synthetic", "Datadog Synthetics",      "Login · Search · Checkout · Payment critical flows",           "Any failure → immediate page"],
    ["Alerting",  "PagerDuty",               "P1 page · P2 Slack · P3 next business day",                   "Thresholds per row above"],
  ];
  // Layer:1.10, Tool:1.70, Coverage:3.75, Threshold:2.65 = 9.20
  tbl(s, monRows, monY + 0.18, [1.10, 1.70, 3.75, 2.65], C.teal, 0.252, false, 9);
}

// ═══════════════════════════════════════════════════════════════════════════════
//  SLIDE 6 — Success Metrics + Team Structure
// ═══════════════════════════════════════════════════════════════════════════════
function slide6(pres) {
  const s = pres.addSlide();
  const y0 = hdr(s, "Success Metrics & Team Structure — Before / After",
    "DORA targets · Service-level ownership · 'You build it, you run it' operating model", false);

  sec(s, "SUCCESS METRICS — TARGETS", y0 + 0.02);
  const metRows = [
    ["Metric", "Current", "6-Month Target", "12-Month Target", "Measurement"],
    ["Deploy frequency",        "1.8 / week",  "3/week per service",   "1+/day per service", "CI/CD pipeline metrics"],
    ["Lead time (commit→prod)", "12.3 days",   "5 days",               "< 2 days",           "PR merge → deploy timestamp"],
    ["Change failure rate",     "18.2%",       "12%",                  "< 8%",               "Hotfixes / total deploys"],
    ["MTTR",                    "4.2 hours",   "1.5 hours",            "< 30 min",           "PagerDuty incident duration"],
    ["P99 API latency",         "820ms",       "400ms",                "< 200ms",            "Prometheus histograms"],
    ["CI build time",           "47 min",      "15 min (per service)", "< 8 min",            "GitHub Actions duration"],
    ["Test flakiness",          "6.8%",        "3%",                   "< 1%",               "CI fail rate ex. code bugs"],
    ["Developer satisfaction",  "5.8 / 10",    "7.0 / 10",             "8.0 / 10",           "Quarterly eng. survey"],
  ];
  // Metric:2.05, Current:1.15, 6mo:1.75, 12mo:1.60, Measure:2.65 = 9.20
  tbl(s, metRows, y0 + 0.22, [2.05, 1.15, 1.75, 1.60, 2.65], C.navyMid, 0.248, false, 9);

  const teamY = y0 + 0.22 + 0.248 * 9 + 0.09;
  sec(s, "TARGET SQUAD STRUCTURE — SERVICE OWNERSHIP MODEL", teamY);
  const teamRows = [
    ["Squad", "Services Owned", "Size", "On-Call"],
    ["Identity",  "User Service, Auth",                               "4 eng", "Yes — own pager"],
    ["Commerce",  "Order Service, Payment Service",                   "6 eng", "Yes — own pager"],
    ["Discovery", "Search Service",                                   "3 eng", "Shared with Platform"],
    ["Platform",  "Notification Service, API Gateway, shared libs",   "4 eng", "Yes — own pager"],
    ["Data",      "Analytics Service",                                "3 eng", "Shared with Platform"],
    ["SRE",       "Kafka, Kubernetes, monitoring, incident response", "3 eng", "Yes — primary escalation"],
  ];
  // Squad:1.30, Services:4.10, Size:0.95, OnCall:2.85 = 9.20
  tbl(s, teamRows, teamY + 0.18, [1.30, 4.10, 0.95, 2.85], C.teal, 0.250, false, 9);

  // Principle callout
  const omY = teamY + 0.18 + 0.250 * 7 + 0.06;
  if (omY + 0.26 < SH - 0.10) {
    s.addShape("rect", { x: LM, y: omY, w: TW, h: 0.26,
      fill: { color: C.darkCard }, line: { color: C.teal, width: 1 } });
    s.addText(
      "Principle: You build it, you run it.  Each squad owns SLOs, deploys independently, and carries their own pager.  SRE provides platform — not babysitting.",
      { x: LM + 0.14, y: omY, w: TW - 0.28, h: 0.26,
        fontSize: 9, bold: true, color: C.tealLt, fontFace: "Calibri",
        margin: 0, valign: "middle" }
    );
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
//  BUILD
// ═══════════════════════════════════════════════════════════════════════════════
async function buildDeck() {
  const pres = new PptxGenJS();
  pres.layout  = "LAYOUT_16x9";
  pres.author  = "Engineering Architecture Team";
  pres.title   = "Project Chimera: Monolith Decomposition Plan";
  pres.subject = "Atlas → Microservices Migration";

  slide1(pres);
  slide2(pres);
  slide3(pres);
  slide4(pres);
  slide5(pres);
  slide6(pres);

  await pres.writeFile({ fileName: "outputs/software-qa.pptx" });
  console.log("✓ outputs/software-qa.pptx written");
}

buildDeck().catch(err => { console.error(err); process.exit(1); });
