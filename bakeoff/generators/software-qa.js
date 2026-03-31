"use strict";
const PptxGenJS = require("pptxgenjs");

// ─── PALETTE ─────────────────────────────────────────────────────────────────
const C = {
  navy:      "0D1B2A",
  teal:      "0A7E8C",
  tealLt:    "14B8C8",
  ice:       "D6EEF2",
  slate:     "37474F",
  muted:     "607D8B",
  offwhite:  "F4F8FA",
  white:     "FFFFFF",
  cream:     "EEF4F6",
  red:       "C62828",
  orange:    "E65100",
  amber:     "F57C00",
  green:     "2E7D32",
  border:    "B0BEC5",
  darkCard:  "0F2638",
  darkCard2: "162B3A",
};

// Slide dimensions and safe margins
const SW  = 10.0;   // width
const SH  = 5.625;  // height
const LM  = 0.28;   // left margin (after accent bar)
const RM  = 0.18;   // right margin
const BM  = 0.20;   // bottom safe margin
const CW  = SW - LM - RM;  // = 9.54  usable content width

// ─── HELPERS ─────────────────────────────────────────────────────────────────
function shd() { return { type: "outer", color: "000000", opacity: 0.10, blur: 5, offset: 2, angle: 135 }; }

// Standard slide header — returns Y where content starts
function hdr(s, title, sub, dark) {
  s.background = { color: dark ? C.navy : C.offwhite };
  s.addShape("rect", { x: 0, y: 0, w: 0.18, h: SH, fill: { color: C.teal }, line: { color: C.teal, width: 0 } });
  s.addText(title, { x: LM, y: 0.17, w: CW, h: 0.52,
    fontSize: 19, bold: true, color: dark ? C.white : C.navy, fontFace: "Calibri", margin: 0, valign: "middle" });
  if (sub) {
    s.addText(sub, { x: LM, y: 0.67, w: CW, h: 0.22,
      fontSize: 8.5, color: dark ? C.tealLt : C.muted, fontFace: "Calibri", margin: 0, italic: true });
  }
  s.addShape("rect", { x: LM, y: 0.89, w: CW, h: 0.025,
    fill: { color: dark ? C.teal : C.border }, line: { color: dark ? C.teal : C.border, width: 0 } });
  return 0.94; // content-start Y
}

// Section label
function sec(s, text, y, color) {
  s.addText(text, { x: LM, y, w: CW, h: 0.20, fontSize: 9, bold: true, color: color || C.navy, fontFace: "Calibri", margin: 0 });
}

// Generic table renderer
// colW = array of column widths (must sum to CW or desired width)
// dark  = use dark-row colours (for slides with navy bg)
// hBg   = header background colour
// rH    = row height in inches
function tbl(s, rows, y, colW, hBg, rH, dark, fontSize) {
  const fs = fontSize || 8.5;
  const x  = LM;
  const data = rows.map((row, ri) => row.map((cell, _ci) => {
    const isH = ri === 0;
    const even = ri % 2 === 0;
    let bg   = isH ? hBg : (dark ? (even ? C.darkCard : C.darkCard2) : (even ? C.white : C.cream));
    let fg   = isH ? C.white : (dark ? C.white : C.slate);
    let bold = isH;
    const cs = String(cell);
    if (!isH) {
      if      (cs === "CRITICAL") { bg = C.red;    fg = C.white; bold = true; }
      else if (cs === "HIGH")     { bg = "BF360C"; fg = C.white; bold = true; }
      else if (cs === "MEDIUM")   { bg = C.amber;  fg = C.white; bold = true; }
      else if (cs === "LOW")      { bg = C.green;  fg = C.white; bold = true; }
    }
    const bColor = dark ? "263238" : C.border;
    return {
      text: cs,
      options: {
        fill: { color: bg }, color: fg, bold,
        fontSize: fs, fontFace: "Calibri",
        align: "left", valign: "middle",
        margin: [2, 4, 2, 4],
        border: [
          { type: "solid", pt: 0.5, color: bColor },
          { type: "solid", pt: 0.5, color: bColor },
          { type: "solid", pt: 0.5, color: bColor },
          { type: "solid", pt: 0.5, color: bColor },
        ]
      }
    };
  }));
  s.addTable(data, { x, y, w: colW.reduce((a, b) => a + b, 0), colW, rowH: rH });
}

// ═════════════════════════════════════════════════════════════════════════════
//  BUILD
// ═════════════════════════════════════════════════════════════════════════════
async function buildDeck() {
  const pres = new PptxGenJS();
  pres.layout = "LAYOUT_16x9";
  pres.title  = "Project Chimera: Monolith Decomposition Plan";
  pres.author = "Engineering Leadership";

  // ══════════════════════════════════════════════════════════════════════════
  //  SLIDE 1 — Title + Current State  (dark)
  // ══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.navy };
    s.addShape("rect", { x: 0, y: 0, w: 0.22, h: SH, fill: { color: C.teal }, line: { color: C.teal, width: 0 } });

    s.addText("Project Chimera", {
      x: LM, y: 0.14, w: CW, h: 0.60,
      fontSize: 34, bold: true, color: C.white, fontFace: "Calibri", margin: 0
    });
    s.addText("Monolith Decomposition Plan  ·  Atlas → Microservices  ·  VP Engineering + 20 Backend Engineers  ·  March 2026", {
      x: LM, y: 0.73, w: CW, h: 0.20,
      fontSize: 9.5, color: C.tealLt, fontFace: "Calibri", margin: 0, italic: true
    });
    s.addShape("rect", { x: LM, y: 0.93, w: CW, h: 0.025,
      fill: { color: C.teal }, line: { color: C.teal, width: 0 } });

    // 4 stat boxes — y=1.02, h=0.68
    const statW  = 2.29;
    const statGp = 0.08;
    [
      { val: "340K",  lbl: "Lines of Python" },
      { val: "2,800", lbl: "DB Tables (single PG)" },
      { val: "14",    lbl: "Django Apps" },
      { val: "22",    lbl: "Backend Engineers" },
    ].forEach((st, i) => {
      const bx = LM + i * (statW + statGp);
      s.addShape("rect", { x: bx, y: 1.02, w: statW, h: 0.68,
        fill: { color: C.darkCard }, line: { color: C.teal, width: 1 } });
      s.addText(st.val, { x: bx, y: 1.03, w: statW, h: 0.38,
        fontSize: 24, bold: true, color: C.tealLt, fontFace: "Calibri", align: "center", margin: 0 });
      s.addText(st.lbl, { x: bx, y: 1.40, w: statW, h: 0.26,
        fontSize: 9, color: "A0C4CC", fontFace: "Calibri", align: "center", margin: 0 });
    });

    // DORA table  (9 rows × 0.258 = 2.322, starts 2.00, ends 4.322)
    s.addText("DORA Metrics vs. Industry P50 — Measured, Not Anecdotal", {
      x: LM, y: 1.80, w: CW, h: 0.20,
      fontSize: 9.5, bold: true, color: C.tealLt, fontFace: "Calibri", margin: 0
    });
    // 6 rows (header + 5 data) × 0.295 = 1.770, starts 2.00, ends 3.770
    tbl(s, [
      ["Metric",                    "Current (Atlas)",               "DORA Industry P50",  "Status"],
      ["Deploy Frequency",          "1.8 / week  (every 3.9 days)", "1/day – 1/week",     "Bottom of 'Medium'"],
      ["Lead Time (commit → prod)", "12.3 days",                    "1 day – 1 week",     "2× over Medium ceiling"],
      ["Change Failure Rate",       "18.2%",                        "0 – 15%",            "Above Medium threshold"],
      ["Mean Time to Recover",      "4.2 hours",                    "< 1 hr – < 1 day",   "Functional but slow"],
      ["CI Build Time",             "47 min (full suite)",          "—",                  "Forced context-switching"],
    ], 2.00, [2.50, 2.18, 1.68, 3.18], C.teal, 0.295, true, 9);

    // Cost callout — immediately below table (3.770 + 0.12 gap)
    s.addShape("rect", { x: LM, y: 3.90, w: CW, h: 0.60,
      fill: { color: "200B00" }, line: { color: C.orange, width: 1 } });
    s.addText(
      "Cost of Friction:  ~1,400 engineer-hours / quarter  →  ~$133K / quarter wasted\n" +
      "Breakdown: deployment delays + conflict resolution + rollback recovery + flaky-test investigation  @  $95 / hr loaded cost",
      { x: LM + 0.12, y: 3.91, w: CW - 0.24, h: 0.58,
        fontSize: 9, color: C.amber, fontFace: "Calibri", margin: 0, bold: true, valign: "middle" }
    );
  }

  // ══════════════════════════════════════════════════════════════════════════
  //  SLIDE 2 — Target Architecture  (light)
  //  Table width: LM=0.28, colW sums to 5.52 → right edge at 5.80
  //  Diagram: x=5.98, width=3.74 → right edge at 9.72 (safe, RM=0.28 gap)
  // ══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    hdr(s, "Target Architecture — Strangler Fig Decomposition",
      "Domain-driven service boundaries · Kong API Gateway · Istio Service Mesh · Kafka · Isolated data store per service", false);

    // Service boundary table — colW sums to 5.50, right edge = 0.28+5.50 = 5.78
    // Diagram starts at x=5.98, giving 0.20" clearance — avoids Priority column overlap
    tbl(s, [
      ["Service",              "Owner",     "Data Store",              "API",           "Priority"],
      ["User Service",         "Identity",  "PostgreSQL (isolated)",   "REST + gRPC",   "P0 — 1st Extract"],
      ["Order Service",        "Commerce",  "PostgreSQL (isolated)",   "REST + Events", "P0 — High Coupling"],
      ["Payment Service",      "Commerce",  "PostgreSQL + Stripe",     "REST (sync)",   "P1 — After Orders"],
      ["Search Service",       "Discovery", "Elasticsearch 8",         "REST",          "P1 — Low Coupling"],
      ["Notification Svc",     "Platform",  "PostgreSQL + SQS",        "Async/Events",  "P2 — Low Risk"],
      ["Analytics Service",    "Data",      "ClickHouse",              "gRPC (int.)",   "P2 — Read-Only"],
    ], 1.06, [1.36, 0.88, 1.50, 1.02, 0.74], C.teal, 0.268, false);

    // Key principles below table
    [
      ["Strangler Fig:", "Incremental extraction behind gateway — not a big-bang rewrite."],
      ["DB Isolation:",  "Each service owns its schema — no shared tables, no cross-service joins."],
      ["Async-first:",   "Kafka events for consistency; sync only on checkout-critical paths."],
      ["You build it:",  "Each squad owns SLOs, deploys independently, carries their own pager."],
      ["Bounded scope:", "Reporting & Admin stay in monolith — goal is friction removal, not purity."],
    ].forEach((p, i) => {
      const yy = 3.26 + i * 0.26;
      s.addText(p[0], { x: LM, y: yy, w: 1.48, h: 0.25, fontSize: 8.5, bold: true, color: C.teal, fontFace: "Calibri", margin: 0 });
      s.addText(p[1], { x: 1.78, y: yy, w: 4.00, h: 0.25, fontSize: 8.5, color: C.slate, fontFace: "Calibri", margin: 0 });
    });

    // Architecture diagram — x=5.98, fits within right edge (9.72)
    const dx = 5.98, dy = 1.06, dw = 3.74, dh = 4.38;
    s.addShape("rect", { x: dx, y: dy, w: dw, h: dh,
      fill: { color: C.navy }, line: { color: C.teal, width: 1 }, shadow: shd() });
    s.addText("TARGET ARCHITECTURE", {
      x: dx, y: dy+0.06, w: dw, h: 0.20,
      fontSize: 7.5, bold: true, color: C.tealLt, fontFace: "Calibri", align: "center", margin: 0
    });

    const box = (bx, by, bw, bh, txt, fill, border, fs) => {
      s.addShape("rect", { x: bx, y: by, w: bw, h: bh, fill: { color: fill }, line: { color: border || fill, width: 0.75 } });
      s.addText(txt, { x: bx, y: by, w: bw, h: bh, fontSize: fs || 7, color: C.white,
        fontFace: "Calibri", align: "center", valign: "middle", margin: 0 });
    };
    const arr = (ax, ay) => s.addShape("rect", {
      x: ax, y: ay, w: 0.13, h: 0.17, fill: { color: C.muted }, line: { color: C.muted, width: 0 } });

    box(dx+0.74, dy+0.30, 2.26, 0.27, "Clients  (Web · Mobile · API)", C.teal, C.teal, 7.5);
    arr(dx+1.80, dy+0.57);
    box(dx+0.44, dy+0.75, 2.86, 0.42, "API Gateway\nKong · Rate Limit · Auth · Circuit Breaker", C.darkCard, C.tealLt, 7);
    arr(dx+1.80, dy+1.17);
    s.addText("SERVICE MESH  (Istio)", {
      x: dx+0.22, y: dy+1.33, w: 3.30, h: 0.15,
      fontSize: 6, color: C.muted, fontFace: "Calibri", align: "center", margin: 0, italic: true
    });
    ["User\nService","Order\nService","Payment\nService"].forEach((nm, i) => {
      box(dx+0.22+i*1.20, dy+1.48, 1.10, 0.46, nm, "193A4F", C.teal, 7);
    });
    box(dx+0.22, dy+2.04, 3.30, 0.36, "Event Bus  (Kafka)\nTopics: user.* · order.* · payment.*", "1A2F1A", C.green, 7);
    ["Search\nService","Notif.\nService","Analytics\nService"].forEach((nm, i) => {
      box(dx+0.22+i*1.20, dy+2.50, 1.10, 0.46, nm, "193A4F", C.teal, 7);
    });
    box(dx+0.22, dy+3.06, 3.30, 0.42, "Isolated Data Stores\nPostgreSQL×4 · Elasticsearch · ClickHouse · Redis/svc",
      "1C1030", "9C27B0", 6.5);
  }

  // ══════════════════════════════════════════════════════════════════════════
  //  SLIDE 3 — ADR: Why Strangler Fig  (dark)
  // ══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    hdr(s, "Architecture Decision Record — Why Strangler Fig, Not Rewrite",
      "2-week evaluation spike  ·  April 14–25, 2026  ·  Three options scored against risk, timeline, and reversibility", true);

    const OPTS = [
      {
        verdict: "REJECTED", vColor: C.red,
        title: "Option 1: Full Rewrite",
        lines: [
          "Greenfield rebuild in Go + gRPC",
          "Est. 14–18 months, 8 dedicated engineers",
          "Monolith still needs maintenance concurrently",
          "Two systems running in parallel for 12+ months",
          "Industry data: rewrites take 2–3× est. timeline",
          "Netscape, Basecamp: frequently fail to ship",
          "340K lines of tested logic — all edge cases must",
          "  be rediscovered the hard way during rewrite",
        ]
      },
      {
        verdict: "REJECTED", vColor: C.amber,
        title: "Option 2: Modularize Monolith",
        lines: [
          "Tried Q4 2025 as 'Atlas Modular' initiative",
          "6 weeks: Django apps → discrete packages",
          "Merge conflicts ↓30% — deploy freq. unchanged",
          "Still one artifact → 47-min CI build unchanged",
          "Blast radius unchanged — one deploy breaks all",
          "Addresses code coupling, not deploy coupling",
          "Root cause is artifact cohesion, not structure",
          "No path to independent deploys without extraction",
        ]
      },
      {
        verdict: "CHOSEN", vColor: C.green,
        title: "Option 3: Strangler Fig",
        lines: [
          "Incremental extraction behind API Gateway",
          "Each service validated in prod before cutover",
          "Rollback = reroute traffic back to monolith",
          "Risk bounded to one service at a time",
          "Proven in Q1 2026 prototype:",
          "  User Svc: 3 wks extract, 2 wks parallel run",
          "  Cutover with zero customer-facing incidents",
          "Scales to harder extractions (Orders, Payments)",
        ]
      },
    ];

    // Cards: each 3.06 wide, gap 0.09  →  3×3.06 + 2×0.09 = 9.36  +  LM 0.28 = 9.64 ✓
    const cW = 3.06, cGap = 0.09;
    OPTS.forEach((opt, i) => {
      const cx = LM + i * (cW + cGap);
      // Header bar
      s.addShape("rect", { x: cx, y: 1.06, w: cW, h: 0.28, fill: { color: opt.vColor }, line: { color: opt.vColor, width: 0 } });
      s.addText(opt.title, { x: cx+0.08, y: 1.06, w: cW-0.16, h: 0.28,
        fontSize: 8.5, bold: true, color: C.white, fontFace: "Calibri", align: "left", valign: "middle", margin: 0 });
      // Card body
      s.addShape("rect", { x: cx, y: 1.34, w: cW, h: 2.02,
        fill: { color: C.darkCard }, line: { color: opt.vColor, width: 0.75 } });
      opt.lines.forEach((line, li) => {
        s.addText(line, { x: cx+0.10, y: 1.40+li*0.244, w: cW-0.20, h: 0.24,
          fontSize: 8, color: C.white, fontFace: "Calibri", bold: li === 0, margin: 0 });
      });
      // Verdict badge
      s.addShape("rect", { x: cx+0.10, y: 3.38, w: 0.90, h: 0.22,
        fill: { color: opt.vColor }, line: { color: opt.vColor, width: 0 } });
      s.addText(opt.verdict, { x: cx+0.10, y: 3.38, w: 0.90, h: 0.22,
        fontSize: 8, bold: true, color: C.white, fontFace: "Calibri", align: "center", valign: "middle", margin: 0 });
    });

    // ─ Bottom section ─
    // Banner: y=3.62, h=0.24 → 3.86
    // Titles: y=3.90, h=0.20 → 4.10
    // 3 items × 0.34 = 1.02, starting 4.10 → ends 5.12  BM = 0.505" ✓
    s.addShape("rect", { x: LM, y: 3.62, w: CW, h: 0.24,
      fill: { color: C.teal }, line: { color: C.teal, width: 0 } });
    s.addText("SCOPE BOUNDARIES & TIMELINE CONFIDENCE", {
      x: LM+0.10, y: 3.62, w: CW-0.20, h: 0.24,
      fontSize: 8.5, bold: true, color: C.white, fontFace: "Calibri", align: "left", valign: "middle", margin: 0
    });

    // Two columns: colW = (CW-0.12)/2 = 4.71
    const colW3 = (CW - 0.12) / 2;
    [
      {
        title: "What We Are NOT Doing",
        color: C.amber,
        items: [
          "Reporting & Admin stay in monolith; low-traffic, low-change — extraction cost > benefit",
          "Not self-hosting Kafka — Confluent Cloud managed; team has zero Kafka experience",
          "Istio deferred to Month 3; ClickHouse to Month 5 — stagger tools to avoid overload",
        ]
      },
      {
        title: "Timeline Confidence",
        color: C.tealLt,
        items: [
          "65% confident: Phase 1–2 on schedule — validated by Q1 2026 User Service prototype",
          "40% confident: Phase 3 on time — ClickHouse adoption has unknown learning curve",
          "Checkout Saga is highest-risk item; may slip Phase 2 by 1 week — flagged proactively",
        ]
      },
    ].forEach((col, i) => {
      const cx = LM + i * (colW3 + 0.12);
      s.addText(col.title, { x: cx, y: 3.90, w: colW3, h: 0.20,
        fontSize: 9, bold: true, color: col.color, fontFace: "Calibri", margin: 0 });
      col.items.forEach((item, ii) => {
        // 3 items × 0.34 = 1.02, starting 4.10 → ends 5.12  BM = 0.505" ✓
        s.addText("• " + item, { x: cx, y: 4.10+ii*0.34, w: colW3, h: 0.34,
          fontSize: 8.5, color: C.white, fontFace: "Calibri", margin: 0 });
      });
    });
  }

  // ══════════════════════════════════════════════════════════════════════════
  //  SLIDE 4 — Migration Plan  (light)
  //
  //  Layout math (rowH=0.205):
  //  Pills   y=1.06 h=0.22 → 1.28
  //  P1 sec  y=1.32 h=0.18 → 1.50
  //  P1 tbl  y=1.50  6r×0.205=1.230 → 2.730
  //  P2 sec  y=2.77 h=0.18 → 2.95
  //  P2 tbl  y=2.95  5r×0.205=1.025 → 3.975
  //  P3 sec  y=4.02 h=0.18 → 4.20
  //  P3 tbl  y=4.20  5r×0.205=1.025 → 5.225
  //  BM      5.625-5.225 = 0.400" ✓
  // ══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    hdr(s, "Migration Plan — 6-Month Phased Rollout",
      "Phase 1: Foundation (M1–2) · Phase 2: Commerce Core (M3–4) · Phase 3: Platform & Optimization (M5–6)", false);

    const phaseW  = (CW - 0.18) / 3;
    const phaseColors = [C.teal, C.navy, "546E7A"];
    ["Phase 1: Foundation  ·  Months 1–2",
     "Phase 2: Commerce Core  ·  Months 3–4",
     "Phase 3: Platform & Opt.  ·  Months 5–6"].forEach((lbl, i) => {
      const px = LM + i * (phaseW + 0.09);
      s.addShape("rect", { x: px, y: 1.06, w: phaseW, h: 0.22,
        fill: { color: phaseColors[i] }, line: { color: phaseColors[i], width: 0 } });
      s.addText(lbl, { x: px, y: 1.06, w: phaseW, h: 0.22,
        fontSize: 8, bold: true, color: C.white, fontFace: "Calibri", align: "center", valign: "middle", margin: 0 });
    });

    // 3-column tables (Key Risk removed) — sum=9.54=CW
    // Task=3.80, Owner=1.74, Deliverable=4.00
    const tColW = [3.80, 1.74, 4.00];
    const rH = 0.228;

    // Phase 1  (6 rows incl. header) — y=1.50, 6r×0.228=1.368 → 2.868
    sec(s, "Phase 1: Foundation", 1.32, C.teal);
    tbl(s, [
      ["Task",                            "Owner",             "Deliverable"],
      ["Deploy API Gateway (Kong)",       "SRE",               "All traffic via gateway; monolith = upstream"],
      ["Instrument with OpenTelemetry",   "SRE + all squads",  "Distributed tracing: HTTP + Celery tasks"],
      ["Extract User Service",            "Identity squad",    "User CRUD, auth, permissions standalone"],
      ["Stand up Kafka (3 brokers)",      "SRE",               "Event bus; user.created / user.updated topics"],
      ["DB migration tooling (CDC)",      "DBA",               "Debezium CDC; dual-write; reconciliation jobs"],
    ], 1.50, tColW, C.teal, rH, false, 9);

    // Phase 2  (5 rows incl. header) — y=2.95, 5r×0.228=1.140 → 4.090
    sec(s, "Phase 2: Commerce Core", 2.81, C.navy);
    tbl(s, [
      ["Task",                            "Owner",               "Deliverable"],
      ["Extract Order Service",           "Commerce squad",      "Orders, carts, pricing standalone; Kafka consumer"],
      ["Extract Payment Service",         "Commerce squad",      "Stripe, invoicing, refunds standalone"],
      ["Implement Checkout Saga",         "Commerce + Identity", "Distributed txn: order → inventory → charge"],
      ["Extract Search Service",          "Discovery squad",     "Elasticsearch standalone; monolith → Kafka producer"],
    ], 2.96, tColW, C.navy, rH, false, 9);

    // Phase 3  (5 rows incl. header) — y=4.15, 5r×0.228=1.140 → 5.290, BM=0.335"
    sec(s, "Phase 3: Platform & Optimization", 4.02, "546E7A");
    tbl(s, [
      ["Task",                            "Owner",          "Deliverable"],
      ["Extract Notification Service",    "Platform squad", "Email/SMS/push via SQS; template management"],
      ["Extract Analytics Service",       "Data squad",     "ClickHouse ingestion; dashboards migrate to new APIs"],
      ["Decommission monolith modules",   "All squads",     "Remove extracted code; monolith = thin orchestrator"],
      ["Chaos engineering + perf tuning", "SRE",            "Latency targets met; circuit breakers tested; runbooks"],
    ], 4.17, tColW, "546E7A", rH, false, 9);
  }

  // ══════════════════════════════════════════════════════════════════════════
  //  SLIDE 5 — Risk Matrix + Observability  (light)
  //
  //  Layout math (rH=0.248):
  //  Risk sec  y=1.02 h=0.20 → 1.22
  //  Risk tbl  y=1.22  7r×0.248=1.736 → 2.956
  //  Obs sec   y=3.02 h=0.20 → 3.22
  //  Obs tbl   y=3.22  7r×0.248=1.736 → 4.956
  //  Footer    y=5.02 h=0.22 → 5.24
  //  BM        5.625-5.24 = 0.385" ✓
  // ══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    hdr(s, "Risk Matrix & Observability Stack",
      "Severity-coded risks with bounded mitigations · Full-stack monitoring instrumented from Phase 1", false);

    const rH5 = 0.248;

    // Risk: 6 rows (header + 5) × 0.268 = 1.608, starts 1.22, ends 2.828
    sec(s, "Risk Matrix", 1.02, C.navy);
    tbl(s, [
      ["Risk",                                           "Likelihood","Impact","Severity","Mitigation"],
      ["Distributed transaction failures (checkout saga)","HIGH",     "HIGH", "CRITICAL",
        "Compensation/rollback per saga step; feature flag → monolith checkout fallback; integration tests"],
      ["Kafka instability (team inexperience)",           "MEDIUM",   "HIGH", "HIGH",
        "Hire Kafka specialist 3 months; Confluent Cloud managed; start with 3 low-throughput topics"],
      ["Data inconsistency during dual-write",           "HIGH",     "MEDIUM","HIGH",
        "Debezium CDC instead of app dual-write; hourly reconciliation jobs; rollback documented"],
      ["Service-to-service latency overrun",             "MEDIUM",   "MEDIUM","MEDIUM",
        "P99 200ms budget per hop; Istio circuit breakers 500ms; async events over sync calls"],
      ["PCI scope expansion (Payment Service)",          "LOW",      "HIGH", "MEDIUM",
        "Security review M2 before extraction; Stripe hosted checkout minimises PCI surface"],
    ], 1.22, [2.20, 0.98, 0.84, 0.98, 4.54], C.navy, 0.268, false, 9);

    // Obs: 6 rows (header + 5) × 0.268 = 1.608, starts 2.94, ends 4.548
    sec(s, "Observability Stack", 2.90, C.navy);
    tbl(s, [
      ["Layer",         "Tool",                      "Coverage",                                                   "Alert Threshold"],
      ["Metrics",       "Prometheus + Grafana",      "RED (rate, errors, duration); saturation (CPU, mem, conns)", "Error rate >1% / 5 min  ·  P99 >500ms / 10 min"],
      ["Tracing",       "Jaeger + OpenTelemetry",    "End-to-end request flow; latency per service hop",           "Trace >2 s  ·  Span error rate >5%"],
      ["Logging",       "ELK Stack",                 "Structured JSON logs; correlation IDs across services",      "Error log rate >50 / min per service"],
      ["Health Checks", "K8s liveness + readiness",  "Service availability; DB / Redis / Kafka health",            "Probe failure → restart; 3× → page SRE"],
      ["Alerting",      "PagerDuty",                 "Tiered: P1 page → P2 Slack → P3 next business day",          "Per thresholds above"],
    ], 3.06, [1.18, 1.86, 3.54, 2.96], C.teal, 0.268, false, 9);

    // Footer bar — y=4.62, ends 5.06, BM=0.565"
    s.addShape("rect", { x: LM, y: 4.62, w: CW, h: 0.44,
      fill: { color: C.ice }, line: { color: C.teal, width: 0.5 } });
    s.addText(
      "Key Grafana Dashboards:  Service Overview (RED per service)  ·  Kafka Health (consumer lag, broker disk, replication)  ·  Checkout Saga (success %, avg duration, failure by step)  ·  Deploy Tracker (last deploy, canary %, rollback count)",
      { x: LM+0.10, y: 4.62, w: CW-0.20, h: 0.44,
        fontSize: 11, color: C.navy, fontFace: "Calibri", margin: 0, valign: "middle" }
    );
  }

  // ══════════════════════════════════════════════════════════════════════════
  //  SLIDE 6 — Success Metrics + Team  (dark)
  //
  //  Layout math:
  //  Metrics lbl y=1.02 h=0.22 → 1.24
  //  Metrics tbl y=1.24  9r×0.245=2.205 → 3.445
  //  Legend      y=3.52 h=0.15 → 3.67
  //  Squad lbl   y=3.74 h=0.22 → 3.96
  //  Metrics tbl y=1.24  9r×0.228=2.052 → 3.292
  //  Legend      y=3.38 h=0.14 → 3.52
  //  Squad lbl   y=3.58 h=0.20 → 3.78
  //  Squad tbl   y=3.78  7r×0.205=1.435 → 5.215
  //  BM          5.625-5.215 = 0.410" ✓
  // ══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    hdr(s, "Success Metrics & Target Team Structure",
      "Before/after DORA targets · 6-month and 12-month checkpoints · Squad ownership · You build it, you run it", true);

    s.addText("Before / After — DORA & Engineering Health Targets", {
      x: LM, y: 1.02, w: CW, h: 0.22,
      fontSize: 9.5, bold: true, color: C.tealLt, fontFace: "Calibri", margin: 0
    });

    // Custom-coloured metrics table — 4 cols: Metric, Current, Target (12-mo), Measurement
    // colW sums to 9.54=CW; 9 rows × 0.228 = 2.052, starts 1.24, ends 3.292
    const mColW = [2.10, 1.70, 2.14, 3.60];
    const mRows = [
      ["Metric",               "Current (Atlas)",   "12-Month Target",      "Measurement"],
      ["Deploy Frequency",     "1.8 / week",        "1+ / day per service", "CI/CD pipeline metrics"],
      ["Lead Time (→ prod)",   "12.3 days",         "< 2 days",             "PR merge → deploy timestamp"],
      ["Change Failure Rate",  "18.2%",             "< 8%",                 "Hotfix deploys / total"],
      ["MTTR",                 "4.2 hours",         "< 30 minutes",         "PagerDuty incident duration"],
      ["P99 API Latency",      "820 ms",            "< 200 ms",             "Prometheus histograms"],
      ["CI Build Time",        "47 minutes",        "< 8 min / service",    "GitHub Actions duration"],
      ["Test Flakiness",       "6.8%",              "< 1%",                 "CI failure rate (excl. bugs)"],
      ["Dev Satisfaction",     "5.8 / 10",          "8.0 / 10",             "Quarterly eng survey"],
    ];
    const bColor = "263238";
    const mData = mRows.map((row, ri) => row.map((cell, ci) => {
      const isH = ri === 0;
      let bg = isH ? C.teal : C.darkCard;
      let fg = isH ? C.white : C.white;
      let bold = isH;
      if (!isH) {
        if (ci === 1) { bg = "3B1010"; fg = "FF8A80"; bold = true; }
        else if (ci === 2) { bg = "082210"; fg = "B9F6CA"; }
        else if (ci === 3) { bg = ri % 2 === 0 ? C.darkCard : C.darkCard2; fg = "90A4AE"; }
        if (ci === 0 && ri % 2 !== 0) bg = C.darkCard2;
      }
      return {
        text: String(cell),
        options: {
          fill: { color: bg }, color: fg, bold,
          fontSize: 9, fontFace: "Calibri",
          align: "left", valign: "middle",
          margin: [2, 4, 2, 4],
          border: [
            { type: "solid", pt: 0.5, color: bColor },
            { type: "solid", pt: 0.5, color: bColor },
            { type: "solid", pt: 0.5, color: bColor },
            { type: "solid", pt: 0.5, color: bColor },
          ]
        }
      };
    }));
    s.addTable(mData, { x: LM, y: 1.24, w: CW, colW: mColW, rowH: 0.228 });

    // Legend — table ends 3.292, legend at 3.36
    [
      { bg: "3B1010", fg: "FF8A80", lbl: "Current — problem state" },
      { bg: "082210", fg: "B9F6CA", lbl: "12-Month Target" },
    ].forEach((li, idx) => {
      const lx = LM + idx * 3.10;
      s.addShape("rect", { x: lx, y: 3.36, w: 0.16, h: 0.14,
        fill: { color: li.bg }, line: { color: li.fg, width: 0.5 } });
      s.addText(li.lbl, { x: lx+0.22, y: 3.36, w: 2.60, h: 0.14,
        fontSize: 7.5, color: li.fg, fontFace: "Calibri", margin: 0 });
    });

    // Squad table — y=3.72, 7 rows × 0.228 = 1.596, ends 5.316, BM=0.309"
    s.addText("Target Squad Structure — Service Ownership Model", {
      x: LM, y: 3.54, w: CW, h: 0.20,
      fontSize: 9.5, bold: true, color: C.tealLt, fontFace: "Calibri", margin: 0
    });
    tbl(s, [
      ["Squad",      "Services Owned",                                   "Size",   "On-Call",                  "Key Responsibility"],
      ["Identity",   "User Service, Auth",                         "4 eng",  "Yes — own pager",          "User CRUD, org mgmt, JWT / SSO"],
      ["Commerce",   "Order Service, Payment Service",               "6 eng",  "Yes — own pager",          "Checkout saga, Stripe, invoicing, pricing"],
      ["Discovery",  "Search Service",                               "3 eng",  "Shared w/ Platform",       "Elasticsearch, product index, query logs"],
      ["Platform",   "Notif. Svc, API Gateway, shared libs",         "4 eng",  "Yes — own pager",          "SQS templates, Kong routing, dev tooling"],
      ["Data",       "Analytics Service",                            "3 eng",  "Shared w/ Platform",       "ClickHouse, Kafka consumers, metrics"],
      ["SRE",        "Kafka, K8s, monitoring, incidents",            "3 eng",  "Yes — primary escalation", "Platform reliability, chaos eng., runbooks"],
    ], 3.72, [1.08, 2.56, 0.70, 1.60, 3.60], C.teal, 0.228, true, 9);
  }

  await pres.writeFile({ fileName: "outputs/software-qa.pptx" });
  console.log("✅  Wrote outputs/software-qa.pptx");
}

buildDeck().catch(err => { console.error(err); process.exit(1); });
