/**
 * Microservices Migration Deck Generator
 * Dark mode aesthetic: navy/charcoal background, light text, teal/cyan accents
 */
const fs = require("fs");
const path = require("path");
const pptxgen = require("../node_modules/pptxgenjs");

// ============================================================
// LAYER 1: Constants & Utilities
// ============================================================

const W = 10, H = 5.625;
const PAD = 0.5;
const TITLE_H = 0.5;
const BODY_TOP = TITLE_H + 0.25;
const BODY_W = W - PAD * 2;
const FOOTER_Y = 5.35;
const CONTENT_BOTTOM = FOOTER_Y - 0.12;
const SECTION_GAP = 0.12;
const MIN_FONT = 9;

// Dark mode palette
const C = {
  bg:         "0F1729",   // deep navy
  bgCard:     "1A2340",   // slightly lighter card bg
  bgCard2:    "1E2A4A",   // alt card bg
  text:       "E8ECF1",   // off-white text
  textDim:    "8899AA",   // dimmed text
  accent:     "00D4AA",   // teal accent
  accent2:    "00B4D8",   // cyan accent
  accent3:    "7C5CFC",   // purple accent
  warn:       "FF6B6B",   // red/warning
  warnAmber:  "FFB347",   // amber
  success:    "2ECC71",   // green
  headerLine: "00D4AA",   // teal line under headers
  white:      "FFFFFF",
};

const FONT = "Calibri";

// Utilities
function trimText(text, maxChars) {
  if (text.length <= maxChars) return text;
  const trimmed = text.substring(0, maxChars - 3).replace(/\s+\S*$/, "");
  return trimmed || text.substring(0, maxChars - 3);
}

function fitBullets(items, max, chars) {
  const result = items.slice(0, max).map(i => trimText(i, chars));
  if (items.length > max) {
    console.log(`  fitBullets: dropped ${items.length - max} items`);
  }
  return result;
}

function estimateLines(text, fontSize, boxW) {
  const charsPerLine = Math.floor(boxW * 72 / (fontSize * 0.6));
  return Math.ceil(text.length / charsPerLine);
}

function checkFit(label, text, fontSize, boxW, boxH, lineSpacing) {
  lineSpacing = lineSpacing || 1.15;
  const lines = estimateLines(text, fontSize, boxW);
  const needed = lines * (fontSize / 72) * lineSpacing;
  if (needed > boxH) {
    console.warn(`  OVERFLOW: "${label}" ~${needed.toFixed(2)}" exceeds box ${boxH}"`);
  }
}

// ============================================================
// LAYER 2: Helpers
// ============================================================

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Migrating from Monolith to Microservices";

pres.defineSlideMaster({
  title: "DARK_MASTER",
  background: { color: C.bg },
  objects: [],
});

function addHeader(slide, title) {
  slide.addText(title, {
    x: PAD, y: 0.18, w: BODY_W, h: TITLE_H,
    fontSize: 22, fontFace: FONT, color: C.white, bold: true,
    fit: "shrink", margin: 0,
  });
  slide.addShape(pres.shapes.LINE, {
    x: PAD, y: 0.72, w: BODY_W, h: 0,
    line: { color: C.accent, width: 1.5 },
  });
}

function addFooter(slide, pageNum) {
  slide.addShape(pres.shapes.LINE, {
    x: PAD, y: FOOTER_Y, w: BODY_W, h: 0,
    line: { color: C.accent, width: 0.5, dashType: "dash" },
  });
  if (pageNum) {
    slide.addText(String(pageNum), {
      x: W - 1, y: FOOTER_Y + 0.02, w: 0.5, h: 0.2,
      fontSize: 8, fontFace: FONT, color: C.textDim, align: "right", margin: 0,
    });
  }
}

function addSectionLabel(slide, text, y, opts) {
  const color = (opts && opts.color) || C.accent;
  slide.addText(text.toUpperCase(), {
    x: PAD, y, w: BODY_W, h: 0.22,
    fontSize: 10, fontFace: FONT, color, bold: true,
    charSpacing: 3, margin: 0,
  });
  slide.addShape(pres.shapes.LINE, {
    x: PAD, y: y + 0.22, w: 1.5, h: 0,
    line: { color, width: 1 },
  });
  return y + 0.32;
}

function addBullets(slide, items, x, y, w, h, fs) {
  fs = Math.max(fs || 11, MIN_FONT);
  const textArr = items.map((item, i) => ({
    text: item,
    options: {
      bullet: { code: "2022", color: C.accent },
      breakLine: i < items.length - 1,
      fontSize: fs, fontFace: FONT, color: C.text,
    },
  }));
  slide.addText(textArr, { x, y, w, h, valign: "top", margin: [0, 0, 0, 4] });
}

function addSubHeader(slide, text, x, y, w) {
  slide.addText(text, {
    x, y, w, h: 0.25,
    fontSize: 12, fontFace: FONT, color: C.accent2, bold: true, margin: 0,
  });
  return y + 0.28;
}

function twoColumnLayout(gap) {
  gap = gap || 0.3;
  const colW = (BODY_W - gap) / 2;
  return { leftX: PAD, rightX: PAD + colW + gap, colW };
}

function cardShadow() {
  return { type: "outer", blur: 4, offset: 2, color: "000000", opacity: 0.3, angle: 135 };
}

function addCard(slide, x, y, w, h, opts) {
  const fill = (opts && opts.fill) || C.bgCard;
  const accentColor = (opts && opts.accent) || null;
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h,
    fill: { color: fill },
    rectRadius: 0.08,
    shadow: cardShadow(),
  });
  if (accentColor) {
    slide.addShape(pres.shapes.RECTANGLE, {
      x: x + 0.05, y, w: w - 0.1, h: 0.04,
      fill: { color: accentColor },
    });
  }
}

function addBox(slide, label, x, y, w, h, fillColor, textColor, opts) {
  textColor = textColor || C.white;
  const borderColor = (opts && opts.border) || null;
  const shapeOpts = {
    x, y, w, h,
    fill: { color: fillColor },
    rectRadius: 0.06,
  };
  if (borderColor) {
    shapeOpts.line = { color: borderColor, width: 1 };
  }
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, shapeOpts);
  slide.addText(label, {
    x, y, w, h,
    fontSize: 10, fontFace: FONT, color: textColor,
    align: "center", valign: "middle", bold: true, margin: 0,
  });
}

function addArrow(slide, x, y, w, h, color) {
  color = color || C.accent;
  // Horizontal arrow using a line + triangle
  slide.addShape(pres.shapes.LINE, {
    x, y: y + h / 2, w, h: 0,
    line: { color, width: 2 },
  });
  // Arrow head (small triangle text)
  slide.addText("\u25B6", {
    x: x + w - 0.15, y: y + h / 2 - 0.1, w: 0.2, h: 0.2,
    fontSize: 10, color, align: "center", valign: "middle", margin: 0,
  });
}

function addDownArrow(slide, x, y, color) {
  color = color || C.accent;
  slide.addText("\u25BC", {
    x: x - 0.1, y, w: 0.2, h: 0.2,
    fontSize: 10, color, align: "center", valign: "middle", margin: 0,
  });
}

function addChipLabel(slide, text, x, y, color) {
  color = color || C.accent;
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w: text.length * 0.08 + 0.3, h: 0.22,
    fill: { color },
    rectRadius: 0.11,
  });
  slide.addText(text, {
    x, y, w: text.length * 0.08 + 0.3, h: 0.22,
    fontSize: 8, fontFace: FONT, color: C.bg, bold: true,
    align: "center", valign: "middle", margin: 0,
  });
}

// ============================================================
// LAYER 3: Slide Definitions
// ============================================================

let slideNum = 0;

function newSlide() {
  slideNum++;
  const s = pres.addSlide({ masterName: "DARK_MASTER" });
  return s;
}

// --------------------------------------------------
// SLIDE 1: Cover
// --------------------------------------------------
{
  const s = newSlide();
  // Decorative accent bar at top
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: W, h: 0.06,
    fill: { color: C.accent },
  });
  // Decorative geometric element
  s.addShape(pres.shapes.RECTANGLE, {
    x: 7.5, y: 1.5, w: 2.5, h: 2.5,
    fill: { color: C.accent, transparency: 90 },
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 8.0, y: 2.0, w: 2.0, h: 2.0,
    fill: { color: C.accent2, transparency: 85 },
  });

  let y = 1.2;
  s.addText("MIGRATING FROM MONOLITH\nTO MICROSERVICES", {
    x: PAD, y, w: 7, h: 1.2,
    fontSize: 32, fontFace: FONT, color: C.white, bold: true,
    lineSpacingMultiple: 1.1, margin: 0,
  });
  y += 1.4;

  s.addShape(pres.shapes.LINE, {
    x: PAD, y, w: 2, h: 0,
    line: { color: C.accent, width: 3 },
  });
  y += 0.25;

  s.addText("A Practical Guide for Backend Engineering Teams", {
    x: PAD, y, w: 7, h: 0.35,
    fontSize: 14, fontFace: FONT, color: C.textDim, margin: 0,
  });
  y += 0.55;

  s.addText("Presented by Alex Chen, Principal Engineer", {
    x: PAD, y, w: 7, h: 0.25,
    fontSize: 11, fontFace: FONT, color: C.text, margin: 0,
  });
  y += 0.3;
  s.addText("Platform Architecture Team  |  March 2026", {
    x: PAD, y, w: 7, h: 0.25,
    fontSize: 10, fontFace: FONT, color: C.textDim, margin: 0,
  });

  addFooter(s);
}

// --------------------------------------------------
// SLIDE 2: Current State - Monolith Pain Points
// --------------------------------------------------
{
  const s = newSlide();
  addHeader(s, "Current State: The Monolith Problem");
  addFooter(s, slideNum);

  let y = BODY_TOP + 0.55;

  // Three pain point cards
  const cards = [
    {
      icon: "\u26A0",
      title: "Deployment Bottleneck",
      color: C.warn,
      bullets: [
        "Single deployable artifact (~45 min build)",
        "Release train blocks all teams on slowest feature",
        "Rollbacks require full redeploy of entire app",
      ],
    },
    {
      icon: "\uD83D\uDCA5",
      title: "Blast Radius",
      color: C.warnAmber,
      bullets: [
        "OOM in billing module takes down entire platform",
        "No fault isolation between subsystems",
        "One bad deploy = full outage for all customers",
      ],
    },
    {
      icon: "\uD83D\uDD17",
      title: "Team Coupling",
      color: C.accent3,
      bullets: [
        "20 engineers contending on single codebase",
        "Merge conflicts average 3.2/day across teams",
        "Shared DB schema makes independent changes impossible",
      ],
    },
  ];

  const cardW = (BODY_W - 0.3) / 3;
  const cardH = 2.8;
  cards.forEach((card, i) => {
    const cx = PAD + i * (cardW + 0.15);
    addCard(s, cx, y, cardW, cardH, { accent: card.color });

    // Icon + Title
    s.addText(card.title, {
      x: cx + 0.15, y: y + 0.2, w: cardW - 0.3, h: 0.3,
      fontSize: 12, fontFace: FONT, color: card.color, bold: true, margin: 0,
    });

    // Bullets
    addBullets(s, card.bullets, cx + 0.15, y + 0.6, cardW - 0.3, cardH - 0.9, 10);
  });

  // Bottom stat bar
  const statY = y + cardH + 0.2;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: PAD, y: statY, w: BODY_W, h: 0.35,
    fill: { color: C.bgCard },
    rectRadius: 0.04,
  });
  s.addText("Current stats:   Deploy frequency: 1x/week   |   MTTR: 4.2 hours   |   Lead time: 12 days", {
    x: PAD + 0.2, y: statY, w: BODY_W - 0.4, h: 0.35,
    fontSize: 9, fontFace: FONT, color: C.textDim, align: "center", valign: "middle", margin: 0,
  });
}

// --------------------------------------------------
// SLIDE 3: Target Architecture
// --------------------------------------------------
{
  const s = newSlide();
  addHeader(s, "Target Architecture");
  addFooter(s, slideNum);

  let y = BODY_TOP + 0.55;

  // Client layer
  addBox(s, "Clients (Web / Mobile / API)", PAD + 1.5, y, BODY_W - 3, 0.35, C.bgCard2, C.text, { border: C.textDim });
  y += 0.45;
  addDownArrow(s, W / 2, y, C.accent);
  y += 0.25;

  // API Gateway
  addBox(s, "API Gateway  (Auth / Rate Limit / Routing)", PAD + 0.8, y, BODY_W - 1.6, 0.4, C.accent, C.bg);
  y += 0.52;
  addDownArrow(s, W / 2, y, C.accent);
  y += 0.25;

  // Service mesh row
  const svcW = 1.6;
  const svcGap = 0.2;
  const services = ["User\nService", "Order\nService", "Billing\nService", "Inventory\nService", "Notification\nService"];
  const totalSvcW = services.length * svcW + (services.length - 1) * svcGap;
  const svcStartX = (W - totalSvcW) / 2;

  // Background for service mesh
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: svcStartX - 0.15, y: y - 0.05, w: totalSvcW + 0.3, h: 0.85,
    fill: { color: C.bgCard },
    rectRadius: 0.06,
    line: { color: C.accent2, width: 0.5, dashType: "dash" },
  });
  s.addText("Service Mesh", {
    x: svcStartX - 0.1, y: y - 0.22, w: 1.2, h: 0.18,
    fontSize: 8, fontFace: FONT, color: C.accent2, margin: 0,
  });

  services.forEach((svc, i) => {
    const sx = svcStartX + i * (svcW + svcGap);
    addBox(s, svc, sx, y + 0.08, svcW, 0.6, C.bgCard2, C.text, { border: C.accent2 });
  });
  y += 0.95;
  addDownArrow(s, W / 2, y, C.accent);
  y += 0.25;

  // Event Bus
  addBox(s, "Event Bus  (Kafka / RabbitMQ)", PAD + 0.8, y, BODY_W - 1.6, 0.4, C.accent3, C.white);
  y += 0.52;
  addDownArrow(s, W / 2, y, C.accent);
  y += 0.25;

  // Data stores
  const dbW = 1.8;
  const dbGap = 0.3;
  const dbs = ["Users DB\n(Postgres)", "Orders DB\n(Postgres)", "Billing DB\n(Postgres)", "Cache\n(Redis)"];
  const totalDbW = dbs.length * dbW + (dbs.length - 1) * dbGap;
  const dbStartX = (W - totalDbW) / 2;
  dbs.forEach((db, i) => {
    const dx = dbStartX + i * (dbW + dbGap);
    addBox(s, db, dx, y, dbW, 0.5, C.bgCard2, C.textDim, { border: C.textDim });
  });
}

// --------------------------------------------------
// SLIDE 4: Migration Strategy - Strangler Fig Pattern
// --------------------------------------------------
{
  const s = newSlide();
  addHeader(s, "Migration Strategy: Strangler Fig Pattern");
  addFooter(s, slideNum);

  let y = BODY_TOP + 0.5;

  const phases = [
    {
      label: "Phase 1",
      title: "Edge Proxy",
      color: C.accent2,
      desc: "Deploy API gateway in front of monolith. Route all traffic through it. Zero behavior change.",
    },
    {
      label: "Phase 2",
      title: "Extract Read Paths",
      color: C.accent,
      desc: "Move read-only endpoints to new services. Dual-read validation ensures consistency.",
    },
    {
      label: "Phase 3",
      title: "Extract Write Paths",
      color: C.warnAmber,
      desc: "Migrate write operations with event sourcing. Use change data capture for sync during transition.",
    },
    {
      label: "Phase 4",
      title: "Decommission",
      color: C.success,
      desc: "Redirect remaining routes. Archive monolith codebase. Remove legacy DB dependencies.",
    },
  ];

  const phaseW = 2.0;
  const phaseH = 2.6;
  const arrowW = 0.25;
  const totalW = phases.length * phaseW + (phases.length - 1) * arrowW;
  const startX = (W - totalW) / 2;

  phases.forEach((phase, i) => {
    const px = startX + i * (phaseW + arrowW);

    // Phase card
    addCard(s, px, y, phaseW, phaseH, { accent: phase.color });

    // Chip label
    addChipLabel(s, phase.label, px + 0.15, y + 0.2, phase.color);

    // Title
    s.addText(phase.title, {
      x: px + 0.15, y: y + 0.55, w: phaseW - 0.3, h: 0.3,
      fontSize: 13, fontFace: FONT, color: C.white, bold: true, margin: 0,
    });

    // Description
    s.addText(phase.desc, {
      x: px + 0.15, y: y + 0.95, w: phaseW - 0.3, h: 1.4,
      fontSize: 10, fontFace: FONT, color: C.textDim, valign: "top", margin: 0,
    });

    // Arrow between phases
    if (i < phases.length - 1) {
      const ax = px + phaseW + 0.02;
      s.addText("\u25B6", {
        x: ax, y: y + phaseH / 2 - 0.15, w: arrowW, h: 0.3,
        fontSize: 14, color: C.accent, align: "center", valign: "middle", margin: 0,
      });
    }
  });

  // Bottom note
  s.addText("Each phase includes rollback capability and feature flags for gradual cutover", {
    x: PAD, y: y + phaseH + 0.15, w: BODY_W, h: 0.2,
    fontSize: 9, fontFace: FONT, color: C.textDim, align: "center", margin: 0, italic: true,
  });
}

// --------------------------------------------------
// SLIDE 5: Risk Matrix
// --------------------------------------------------
{
  const s = newSlide();
  addHeader(s, "Risk Matrix");
  addFooter(s, slideNum);

  let y = BODY_TOP + 0.5;

  const headerRow = [
    { text: "Risk", options: { fill: { color: C.bgCard }, color: C.accent, bold: true, fontSize: 10, fontFace: FONT } },
    { text: "Likelihood", options: { fill: { color: C.bgCard }, color: C.accent, bold: true, fontSize: 10, fontFace: FONT, align: "center" } },
    { text: "Impact", options: { fill: { color: C.bgCard }, color: C.accent, bold: true, fontSize: 10, fontFace: FONT, align: "center" } },
    { text: "Mitigation", options: { fill: { color: C.bgCard }, color: C.accent, bold: true, fontSize: 10, fontFace: FONT } },
  ];

  const cellOpts = (color) => ({ fill: { color: C.bg }, color: color || C.text, fontSize: 9, fontFace: FONT });
  const riskColor = (level) => level === "High" ? C.warn : level === "Medium" ? C.warnAmber : C.success;

  const risks = [
    ["Data inconsistency during dual-write phase", "High", "High", "Event sourcing + CDC with reconciliation jobs"],
    ["Service discovery failures", "Medium", "High", "Consul mesh with health checks + circuit breakers"],
    ["Cascading failures across services", "Medium", "High", "Bulkhead pattern + timeout policies + fallback caching"],
    ["Team skill gaps with distributed systems", "High", "Medium", "Pairing program + internal training + embedded SRE"],
    ["Performance regression from network calls", "Medium", "Medium", "gRPC for internal comms + connection pooling + caching"],
    ["Increased operational complexity", "High", "Medium", "GitOps + unified observability platform from day one"],
  ];

  const rows = [headerRow];
  risks.forEach((risk) => {
    rows.push([
      { text: risk[0], options: cellOpts(C.text) },
      { text: risk[1], options: { ...cellOpts(riskColor(risk[1])), align: "center", bold: true } },
      { text: risk[2], options: { ...cellOpts(riskColor(risk[2])), align: "center", bold: true } },
      { text: risk[3], options: cellOpts(C.textDim) },
    ]);
  });

  s.addTable(rows, {
    x: PAD, y, w: BODY_W,
    border: { pt: 0.5, color: C.bgCard2 },
    colW: [2.5, 1.0, 1.0, 4.5],
    margin: [4, 6, 4, 6],
  });
}

// --------------------------------------------------
// SLIDE 6: Timeline - 6-Month Phased Rollout
// --------------------------------------------------
{
  const s = newSlide();
  addHeader(s, "Timeline: 6-Month Phased Rollout");
  addFooter(s, slideNum);

  let y = BODY_TOP + 0.6;

  // Timeline bar
  const barX = PAD + 0.3;
  const barW = BODY_W - 0.6;
  const barH = 0.12;

  // Background bar
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: barX, y: y + 0.15, w: barW, h: barH,
    fill: { color: C.bgCard2 },
    rectRadius: 0.06,
  });

  // Phase segments on bar
  const phases = [
    { label: "Phase 1: Edge Proxy", w: 0.17, color: C.accent2, months: "Month 1" },
    { label: "Phase 2: Read Paths", w: 0.25, color: C.accent, months: "Months 2-3" },
    { label: "Phase 3: Write Paths", w: 0.33, color: C.warnAmber, months: "Months 3-5" },
    { label: "Phase 4: Decommission", w: 0.25, color: C.success, months: "Months 5-6" },
  ];

  let segX = barX;
  phases.forEach((phase) => {
    const segW = barW * phase.w;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: segX, y: y + 0.15, w: segW, h: barH,
      fill: { color: phase.color },
      rectRadius: 0.04,
    });
    segX += segW;
  });

  // Month markers
  const months = ["M1", "M2", "M3", "M4", "M5", "M6"];
  months.forEach((m, i) => {
    const mx = barX + (barW / 6) * i + (barW / 12);
    s.addText(m, {
      x: mx - 0.15, y: y - 0.15, w: 0.3, h: 0.15,
      fontSize: 8, fontFace: FONT, color: C.textDim, align: "center", margin: 0,
    });
  });

  y += 0.6;

  // Phase detail cards
  const cardW = (BODY_W - 0.45) / 4;
  const cardH = 2.6;
  const details = [
    {
      title: "Phase 1", subtitle: "Edge Proxy", color: C.accent2,
      items: ["Deploy API gateway", "Configure routing rules", "Instrument baseline metrics", "Zero behavior change"],
    },
    {
      title: "Phase 2", subtitle: "Read Paths", color: C.accent,
      items: ["Extract user profile reads", "Extract catalog queries", "Dual-read validation", "Shadow traffic testing"],
    },
    {
      title: "Phase 3", subtitle: "Write Paths", color: C.warnAmber,
      items: ["Migrate order writes", "CDC for data sync", "Event sourcing setup", "Feature flag cutover"],
    },
    {
      title: "Phase 4", subtitle: "Decommission", color: C.success,
      items: ["Redirect final routes", "Archive monolith repo", "Remove legacy schemas", "Post-migration audit"],
    },
  ];

  details.forEach((d, i) => {
    const cx = PAD + i * (cardW + 0.15);
    addCard(s, cx, y, cardW, cardH, { accent: d.color });

    s.addText(d.title, {
      x: cx + 0.12, y: y + 0.15, w: cardW - 0.24, h: 0.2,
      fontSize: 9, fontFace: FONT, color: d.color, bold: true, margin: 0,
    });
    s.addText(d.subtitle, {
      x: cx + 0.12, y: y + 0.35, w: cardW - 0.24, h: 0.25,
      fontSize: 12, fontFace: FONT, color: C.white, bold: true, margin: 0,
    });

    addBullets(s, d.items, cx + 0.12, y + 0.7, cardW - 0.24, cardH - 0.9, 9);
  });
}

// --------------------------------------------------
// SLIDE 7: Monitoring & Observability
// --------------------------------------------------
{
  const s = newSlide();
  addHeader(s, "Monitoring & Observability: Day One Setup");
  addFooter(s, slideNum);

  let y = BODY_TOP + 0.55;
  const { leftX, rightX, colW } = twoColumnLayout(0.3);

  // Left column: Three pillars
  y = addSectionLabel(s, "The Three Pillars", y);

  const pillars = [
    {
      name: "Metrics (Prometheus + Grafana)",
      color: C.accent,
      items: ["Request rate, error rate, duration (RED)", "Service-level SLIs/SLOs per endpoint", "Resource utilization per container"],
    },
    {
      name: "Logging (ELK Stack)",
      color: C.accent2,
      items: ["Structured JSON logs with correlation IDs", "Centralized aggregation with log levels", "Automated alerting on error spikes"],
    },
    {
      name: "Tracing (Jaeger / OpenTelemetry)",
      color: C.accent3,
      items: ["Distributed trace propagation", "Latency breakdown per service hop", "Dependency mapping auto-discovery"],
    },
  ];

  pillars.forEach((pillar) => {
    addCard(s, leftX, y, colW, 1.05, { accent: pillar.color });
    s.addText(pillar.name, {
      x: leftX + 0.12, y: y + 0.12, w: colW - 0.24, h: 0.2,
      fontSize: 10, fontFace: FONT, color: pillar.color, bold: true, margin: 0,
    });
    addBullets(s, pillar.items, leftX + 0.12, y + 0.35, colW - 0.24, 0.65, 9);
    y += 1.15;
  });

  // Right column: Alerting strategy
  let ry = BODY_TOP + 0.55;
  ry = addSectionLabel(s, "Alerting Strategy", ry, { color: C.accent2 });

  const alerts = [
    { level: "P1 - Critical", color: C.warn, desc: "Service down, SLO breach > 5min\nPagerDuty + auto-escalation" },
    { level: "P2 - Warning", color: C.warnAmber, desc: "Error rate > 1%, latency p99 > 500ms\nSlack alert + on-call review" },
    { level: "P3 - Info", color: C.accent, desc: "Deployment events, scaling triggers\nDashboard annotation only" },
  ];

  alerts.forEach((alert) => {
    addCard(s, rightX, ry, colW, 0.7, { accent: alert.color });
    s.addText(alert.level, {
      x: rightX + 0.12, y: ry + 0.12, w: colW - 0.24, h: 0.18,
      fontSize: 10, fontFace: FONT, color: alert.color, bold: true, margin: 0,
    });
    s.addText(alert.desc, {
      x: rightX + 0.12, y: ry + 0.32, w: colW - 0.24, h: 0.3,
      fontSize: 9, fontFace: FONT, color: C.textDim, margin: 0,
    });
    ry += 0.8;
  });

  // Key dashboards card
  ry += 0.1;
  addCard(s, rightX, ry, colW, 1.3, { fill: C.bgCard2 });
  s.addText("Key Dashboards", {
    x: rightX + 0.12, y: ry + 0.08, w: colW - 0.24, h: 0.2,
    fontSize: 10, fontFace: FONT, color: C.accent, bold: true, margin: 0,
  });
  addBullets(s, [
    "Service Health Overview (RED metrics)",
    "Infrastructure Resource Utilization",
    "Business KPI Real-time Feed",
    "Deployment & Rollback Tracker",
  ], rightX + 0.12, ry + 0.32, colW - 0.24, 0.9, 9);
}

// --------------------------------------------------
// SLIDE 8: Team Structure
// --------------------------------------------------
{
  const s = newSlide();
  addHeader(s, "Team Structure: Service-Oriented Squads");
  addFooter(s, slideNum);

  let y = BODY_TOP + 0.55;

  // Platform team at top
  addBox(s, "Platform Team (4 engineers)\nAPI Gateway  |  Service Mesh  |  CI/CD  |  Observability", PAD + 1.5, y, BODY_W - 3, 0.6, C.accent3, C.white);
  y += 0.75;

  // Arrow down
  addDownArrow(s, W / 2, y, C.accent3);
  y += 0.25;

  // Shared services label
  s.addText("Provides shared infrastructure to all squads", {
    x: PAD, y, w: BODY_W, h: 0.18,
    fontSize: 9, fontFace: FONT, color: C.textDim, align: "center", margin: 0, italic: true,
  });
  y += 0.3;

  // Service squads row
  const squads = [
    { name: "User Squad", engineers: "4 eng", services: "User Service\nAuth Service", color: C.accent },
    { name: "Commerce Squad", engineers: "5 eng", services: "Order Service\nBilling Service", color: C.accent2 },
    { name: "Inventory Squad", engineers: "4 eng", services: "Inventory Service\nCatalog Service", color: C.warnAmber },
    { name: "Comms Squad", engineers: "3 eng", services: "Notification Svc\nEmail Service", color: C.success },
  ];

  const squadW = (BODY_W - 0.45) / 4;
  const squadH = 1.6;
  squads.forEach((sq, i) => {
    const sx = PAD + i * (squadW + 0.15);
    addCard(s, sx, y, squadW, squadH, { accent: sq.color });

    s.addText(sq.name, {
      x: sx + 0.1, y: y + 0.15, w: squadW - 0.2, h: 0.22,
      fontSize: 12, fontFace: FONT, color: sq.color, bold: true, margin: 0,
    });
    addChipLabel(s, sq.engineers, sx + 0.1, y + 0.42, sq.color);
    s.addText(sq.services, {
      x: sx + 0.1, y: y + 0.75, w: squadW - 0.2, h: 0.7,
      fontSize: 10, fontFace: FONT, color: C.textDim, margin: 0,
    });
  });

  // Principles at bottom
  const py = y + squadH + 0.2;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: PAD, y: py, w: BODY_W, h: 0.35,
    fill: { color: C.bgCard },
    rectRadius: 0.04,
  });
  s.addText("Principles:   You build it, you run it   |   Two-pizza teams   |   Service ownership = on-call ownership", {
    x: PAD + 0.15, y: py, w: BODY_W - 0.3, h: 0.35,
    fontSize: 9, fontFace: FONT, color: C.textDim, align: "center", valign: "middle", margin: 0,
  });
}

// --------------------------------------------------
// SLIDE 9: Success Metrics
// --------------------------------------------------
{
  const s = newSlide();
  addHeader(s, "Success Metrics");
  addFooter(s, slideNum);

  let y = BODY_TOP + 0.5;

  // Three big metric cards at top
  const metrics = [
    { label: "Deploy Frequency", before: "1x / week", after: "Multiple / day", icon: "\uD83D\uDE80", color: C.accent },
    { label: "Mean Time to Recovery", before: "4.2 hours", after: "< 15 minutes", icon: "\u23F1", color: C.accent2 },
    { label: "p99 Latency", before: "1200 ms", after: "< 200 ms", icon: "\u26A1", color: C.warnAmber },
  ];

  const metricW = (BODY_W - 0.3) / 3;
  const metricH = 1.5;
  metrics.forEach((m, i) => {
    const mx = PAD + i * (metricW + 0.15);
    addCard(s, mx, y, metricW, metricH, { accent: m.color });

    s.addText(m.label, {
      x: mx + 0.15, y: y + 0.18, w: metricW - 0.3, h: 0.22,
      fontSize: 11, fontFace: FONT, color: m.color, bold: true, margin: 0,
    });

    // Before
    s.addText("BEFORE", {
      x: mx + 0.15, y: y + 0.5, w: (metricW - 0.3) / 2, h: 0.15,
      fontSize: 8, fontFace: FONT, color: C.textDim, margin: 0,
    });
    s.addText(m.before, {
      x: mx + 0.15, y: y + 0.65, w: (metricW - 0.3) / 2, h: 0.25,
      fontSize: 14, fontFace: FONT, color: C.warn, bold: true, margin: 0,
    });

    // Arrow
    s.addText("\u2192", {
      x: mx + metricW / 2 - 0.1, y: y + 0.65, w: 0.2, h: 0.25,
      fontSize: 14, color: C.textDim, align: "center", valign: "middle", margin: 0,
    });

    // After
    s.addText("AFTER", {
      x: mx + metricW / 2 + 0.1, y: y + 0.5, w: (metricW - 0.3) / 2, h: 0.15,
      fontSize: 8, fontFace: FONT, color: C.textDim, margin: 0,
    });
    s.addText(m.after, {
      x: mx + metricW / 2 + 0.1, y: y + 0.65, w: (metricW - 0.3) / 2, h: 0.25,
      fontSize: 14, fontFace: FONT, color: C.success, bold: true, margin: 0,
    });

    // Bottom detail
    s.addShape(pres.shapes.LINE, {
      x: mx + 0.15, y: y + 1.0, w: metricW - 0.3, h: 0,
      line: { color: C.bgCard2, width: 0.5 },
    });
    s.addText("Target at 6-month mark", {
      x: mx + 0.15, y: y + 1.08, w: metricW - 0.3, h: 0.2,
      fontSize: 8, fontFace: FONT, color: C.textDim, align: "center", margin: 0,
    });
  });

  y += metricH + 0.25;

  // Additional DORA metrics
  y = addSectionLabel(s, "DORA Metrics Targets", y);

  const doraMetrics = [
    ["Change Lead Time", "12 days", "< 1 day", "Time from commit to production"],
    ["Change Failure Rate", "23%", "< 5%", "Percentage of deploys causing incidents"],
    ["Service Availability", "99.5%", "99.95%", "Measured per-service, not monolith-wide"],
  ];

  const doraHeader = [
    { text: "Metric", options: { fill: { color: C.bgCard }, color: C.accent, bold: true, fontSize: 9, fontFace: FONT } },
    { text: "Current", options: { fill: { color: C.bgCard }, color: C.warn, bold: true, fontSize: 9, fontFace: FONT, align: "center" } },
    { text: "Target", options: { fill: { color: C.bgCard }, color: C.success, bold: true, fontSize: 9, fontFace: FONT, align: "center" } },
    { text: "Notes", options: { fill: { color: C.bgCard }, color: C.textDim, bold: true, fontSize: 9, fontFace: FONT } },
  ];

  const doraRows = [doraHeader];
  doraMetrics.forEach((row) => {
    doraRows.push([
      { text: row[0], options: { fill: { color: C.bg }, color: C.text, fontSize: 9, fontFace: FONT } },
      { text: row[1], options: { fill: { color: C.bg }, color: C.warn, fontSize: 9, fontFace: FONT, align: "center" } },
      { text: row[2], options: { fill: { color: C.bg }, color: C.success, fontSize: 9, fontFace: FONT, align: "center" } },
      { text: row[3], options: { fill: { color: C.bg }, color: C.textDim, fontSize: 9, fontFace: FONT } },
    ]);
  });

  s.addTable(doraRows, {
    x: PAD, y, w: BODY_W,
    border: { pt: 0.5, color: C.bgCard2 },
    colW: [2.0, 1.2, 1.2, 4.6],
    margin: [3, 6, 3, 6],
  });
}

// --------------------------------------------------
// SLIDE 10: Q&A / Closing
// --------------------------------------------------
{
  const s = newSlide();

  // Decorative accent bar at bottom
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: H - 0.06, w: W, h: 0.06,
    fill: { color: C.accent },
  });

  // Decorative geometric elements (mirror of cover)
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 2.5, h: 2.5,
    fill: { color: C.accent, transparency: 92 },
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0.5, w: 2.0, h: 2.0,
    fill: { color: C.accent2, transparency: 88 },
  });

  let y = 1.5;
  s.addText("Questions?", {
    x: PAD, y, w: BODY_W, h: 0.9,
    fontSize: 36, fontFace: FONT, color: C.white, bold: true,
    align: "center", margin: 0,
  });
  y += 1.0;

  s.addShape(pres.shapes.LINE, {
    x: W / 2 - 1, y, w: 2, h: 0,
    line: { color: C.accent, width: 2 },
  });
  y += 0.3;

  s.addText("Let's discuss architecture decisions,\nmigration concerns, and team readiness.", {
    x: PAD + 1.5, y, w: BODY_W - 3, h: 0.6,
    fontSize: 14, fontFace: FONT, color: C.textDim,
    align: "center", margin: 0,
  });
  y += 0.9;

  s.addText("Alex Chen  |  alex.chen@company.dev  |  #platform-arch", {
    x: PAD, y, w: BODY_W, h: 0.25,
    fontSize: 11, fontFace: FONT, color: C.accent,
    align: "center", margin: 0,
  });

  y += 0.45;
  s.addText("Deck + reference materials: wiki.internal/microservices-migration", {
    x: PAD, y, w: BODY_W, h: 0.2,
    fontSize: 9, fontFace: FONT, color: C.textDim,
    align: "center", margin: 0,
  });
}

// ============================================================
// WRITE FILE
// ============================================================

const fs2 = require("fs");
const outDir = path.join(__dirname, "output");
if (!fs2.existsSync(outDir)) {
  fs2.mkdirSync(outDir, { recursive: true });
}

pres.writeFile({ fileName: path.join(outDir, "microservices-deck.pptx") })
  .then(() => {
    console.log(`Done. ${slideNum} slides written to output/microservices-deck.pptx`);
  })
  .catch((err) => {
    console.error("Error writing file:", err);
    process.exit(1);
  });
