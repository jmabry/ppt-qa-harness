const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Mountain Biking Trail Guide — Pisgah & DuPont";

// ── Layer 1: Constants ────────────────────────────────────────────────────
const W = 10, H = 5.625;
const PAD = 0.5;
const TITLE_H = 0.5;
const BODY_TOP = TITLE_H + 0.12;
const BODY_W = W - PAD * 2;
const FOOTER_Y = 5.35;
const CONTENT_BOTTOM = FOOTER_Y - 0.12;
const SECTION_GAP = 0.12;
const MIN_FONT = 9;

// Forest palette
const FOREST = "2C5F2D";
const MOSS = "97BC62";
const BARK = "6B4226";
const DIRT = "8B7355";
const CREAM = "F5F0E8";
const WHITE = "FFFFFF";
const BLACK = "1A1A1A";
const DGRAY = "444444";
const SGRAY = "777777";
const MGRAY = "CCCCCC";
const LGRAY = "F2F2F2";
const RED = "CC3333";
const BLUE = "2D6A9F";

let slideNum = 0;

// ── Layer 1: Utilities ────────────────────────────────────────────────────

function trimText(text, maxChars) {
  if (text.length <= maxChars) return text;
  let cut = text.lastIndexOf(" ", maxChars);
  if (cut < maxChars * 0.6) cut = maxChars;
  return text.slice(0, cut).replace(/[,;:\s]+$/, "");
}

function fitBullets(items, maxItems, maxChars) {
  const result = items.slice(0, maxItems).map(b => trimText(b, maxChars));
  if (items.length > maxItems) {
    console.log(`Content fitting: dropped ${items.length - maxItems} bullets`);
  }
  return result;
}

function checkFit(label, text, fontSize, boxW, boxH, lineSpacing) {
  const charsPerLine = Math.floor(boxW * 72 / (fontSize * 0.6));
  const lines = Math.ceil(text.length / charsPerLine);
  const estH = lines * (fontSize / 72) * (lineSpacing || 1.2);
  if (estH > boxH) {
    console.warn(`⚠ OVERFLOW: "${label}" ~${estH.toFixed(2)}" exceeds box ${boxH}" (${text.length} chars at ${fontSize}pt)`);
  }
}

function estimateLines(text, fontSize, boxW) {
  const charsPerLine = Math.floor(boxW * 72 / (fontSize * 0.6));
  return Math.ceil(text.length / charsPerLine);
}

// ── Slide master ──────────────────────────────────────────────────────────
pres.defineSlideMaster({
  title: "MASTER",
  background: { color: WHITE },
  objects: []
});

// ── Layer 2: Helpers ──────────────────────────────────────────────────────

function addHeader(slide, title) {
  slide.addText(title, {
    x: PAD, y: 0.06, w: BODY_W - 2.5, h: TITLE_H,
    fontSize: 20, fontFace: "Georgia", bold: true, color: FOREST,
    valign: "bottom", margin: 0
  });
  slide.addText("Trail Guide", {
    x: W - PAD - 1.5, y: 0.12, w: 1.5, h: 0.22,
    fontSize: 9, fontFace: "Calibri", color: SGRAY,
    align: "right", valign: "middle", margin: 0
  });
}

function addFooter(slide) {
  slideNum++;
  slide.addShape(pres.shapes.LINE, {
    x: 0, y: FOOTER_Y, w: W, h: 0,
    line: { color: MOSS, width: 1.5 }
  });
  slide.addText(String(slideNum), {
    x: W - PAD - 0.5, y: FOOTER_Y + 0.02, w: 0.5, h: 0.22,
    fontSize: 8, fontFace: "Calibri", color: SGRAY,
    align: "right", valign: "middle", margin: 0
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: FOOTER_Y, w: 0.08, h: H - FOOTER_Y,
    fill: { color: FOREST }, line: { color: FOREST }
  });
}

function addSectionLabel(slide, text, y, opts = {}) {
  const x = opts.x !== undefined ? opts.x : PAD;
  const w = opts.w || BODY_W;
  const color = opts.color || DIRT;
  slide.addText(text.toUpperCase(), {
    x, y, w, h: 0.2,
    fontSize: 9, fontFace: "Calibri", bold: true, color,
    charSpacing: 1, margin: 0
  });
  slide.addShape(pres.shapes.LINE, {
    x, y: y + 0.2, w, h: 0,
    line: { color: opts.borderColor || MGRAY, width: 0.75 }
  });
  return y + 0.28;
}

function addBullets(slide, items, x, y, w, h, fontSize) {
  const fs = Math.max(fontSize || 10, MIN_FONT);
  const text = items.map((item, i) => ({
    text: item,
    options: { bullet: true, breakLine: i < items.length - 1 }
  }));
  slide.addText(text, {
    x, y, w, h,
    fontSize: fs, fontFace: "Calibri", color: BLACK,
    valign: "top", lineSpacingMultiple: 1.2, margin: 0
  });
}

function twoColumnLayout(gap) {
  const g = gap || 0.3;
  const colW = (BODY_W - g) / 2;
  return { leftX: PAD, rightX: PAD + colW + g, colW, gap: g };
}

function addTrailCard(slide, trail, x, y, w, h) {
  const diffColors = { easy: MOSS, moderate: BLUE, hard: RED, expert: BLACK };
  const accent = diffColors[trail.difficulty] || MOSS;

  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w, h,
    fill: { color: CREAM }, line: { color: MGRAY }
  });
  // Difficulty accent bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w, h: 0.06,
    fill: { color: accent }, line: { color: accent }
  });
  // Trail name
  slide.addText(trail.name, {
    x: x + 0.12, y: y + 0.12, w: w - 0.24, h: 0.22,
    fontSize: 12, fontFace: "Georgia", bold: true, color: BLACK, margin: 0
  });
  // Stats line
  slide.addText(`${trail.distance} mi  |  ${trail.elevation} ft gain  |  ${trail.difficulty}`, {
    x: x + 0.12, y: y + 0.34, w: w - 0.24, h: 0.18,
    fontSize: 9, fontFace: "Calibri", color: accent, bold: true, margin: 0
  });
  // Description
  slide.addText(trail.desc, {
    x: x + 0.12, y: y + 0.56, w: w - 0.24, h: h - 0.68,
    fontSize: 9, fontFace: "Calibri", color: DGRAY, margin: 0,
    lineSpacingMultiple: 1.2, valign: "top"
  });
}

// ── Layer 3: Slides ───────────────────────────────────────────────────────

// ── Slide 1: Cover ────────────────────────────────────────────────────────
{
  slideNum++;
  const s = pres.addSlide();
  s.background = { color: FOREST };

  // Decorative mountain shapes
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 4.2, w: W, h: 1.425,
    fill: { color: "1A3F1A" }
  });

  s.addText("Mountain Biking\nTrail Guide", {
    x: 0.8, y: 0.8, w: 8, h: 1.8,
    fontSize: 44, fontFace: "Georgia", bold: true, color: WHITE,
    lineSpacingMultiple: 1.1, margin: 0
  });

  s.addText("Pisgah National Forest & DuPont State Forest", {
    x: 0.8, y: 2.7, w: 7, h: 0.4,
    fontSize: 18, fontFace: "Calibri", color: MOSS,
    margin: 0
  });

  s.addText("Western North Carolina  |  2026 Season", {
    x: 0.8, y: 3.3, w: 7, h: 0.3,
    fontSize: 12, fontFace: "Calibri", color: "88AA88",
    margin: 0
  });
}

// ── Slide 2: Overview ─────────────────────────────────────────────────────
{
  const s = pres.addSlide({ masterName: "MASTER" });
  addHeader(s, "Why These Trails");
  addFooter(s);

  let y = BODY_TOP;
  y = addSectionLabel(s, "Two forests, endless singletrack", y);

  const { leftX, rightX, colW } = twoColumnLayout();

  addBullets(s, [
    "Pisgah: old-growth forest with technical rock gardens and root-laced descents",
    "DuPont: smooth flow trails and waterfall views — great for building skills",
    "Both within 30 minutes of Brevard, NC (aka the cycling capital of the east)",
    "Rideable 9+ months of the year with mild winters"
  ], leftX, y, colW, 2.5, 10);

  // Stats callout on right
  const stats = [
    { num: "500+", label: "Miles of trail" },
    { num: "4,000'", label: "Max elevation" },
    { num: "12", label: "Featured trails" },
    { num: "4", label: "Difficulty levels" }
  ];

  stats.forEach((st, i) => {
    const row = Math.floor(i / 2);
    const col = i % 2;
    const sx = rightX + col * (colW / 2);
    const sy = y + row * 1.2;

    s.addText(st.num, {
      x: sx, y: sy, w: colW / 2, h: 0.7,
      fontSize: 36, fontFace: "Georgia", bold: true, color: FOREST,
      align: "center", valign: "bottom", margin: 0
    });
    s.addText(st.label, {
      x: sx, y: sy + 0.7, w: colW / 2, h: 0.3,
      fontSize: 10, fontFace: "Calibri", color: SGRAY,
      align: "center", valign: "top", margin: 0
    });
  });
}

// ── Slide 3: Difficulty Guide ─────────────────────────────────────────────
{
  const s = pres.addSlide({ masterName: "MASTER" });
  addHeader(s, "Trail Difficulty Guide");
  addFooter(s);

  let y = BODY_TOP;
  y = addSectionLabel(s, "Know before you go", y);

  const levels = [
    { label: "EASY", color: MOSS, desc: "Smooth fire roads and gentle singletrack. Minimal elevation. Good for beginners and families.", terrain: "Packed dirt, gravel, gentle grades" },
    { label: "MODERATE", color: BLUE, desc: "Rolling singletrack with some roots and rocks. Moderate climbing. Intermediate fitness needed.", terrain: "Roots, small rocks, moderate grades, some switchbacks" },
    { label: "HARD", color: RED, desc: "Technical terrain with rock gardens, steep climbs, and fast descents. Strong bike handling required.", terrain: "Rock gardens, log crossings, steep grades, exposure" },
    { label: "EXPERT", color: BLACK, desc: "The real deal. Sustained technical features, hike-a-bike sections, serious consequences for mistakes.", terrain: "Mandatory drops, large rocks, steep chutes, route-finding" }
  ];

  const cardH = 0.85;
  levels.forEach((lv, i) => {
    const cy = y + i * (cardH + 0.1);
    s.addShape(pres.shapes.RECTANGLE, {
      x: PAD, y: cy, w: BODY_W, h: cardH,
      fill: { color: LGRAY }, line: { color: MGRAY }
    });
    // Color accent
    s.addShape(pres.shapes.RECTANGLE, {
      x: PAD, y: cy, w: 0.06, h: cardH,
      fill: { color: lv.color }, line: { color: lv.color }
    });
    // Label chip
    s.addShape(pres.shapes.RECTANGLE, {
      x: PAD + 0.18, y: cy + 0.12, w: 1.0, h: 0.25,
      fill: { color: lv.color }
    });
    s.addText(lv.label, {
      x: PAD + 0.18, y: cy + 0.12, w: 1.0, h: 0.25,
      fontSize: 10, fontFace: "Calibri", bold: true, color: WHITE,
      align: "center", valign: "middle", margin: 0
    });
    // Description
    s.addText(lv.desc, {
      x: PAD + 1.35, y: cy + 0.06, w: 5.0, h: 0.35,
      fontSize: 10, fontFace: "Calibri", color: BLACK, margin: 0, valign: "middle"
    });
    // Terrain
    s.addText("Terrain: " + lv.terrain, {
      x: PAD + 1.35, y: cy + 0.44, w: 5.0, h: 0.3,
      fontSize: 9, fontFace: "Calibri", color: SGRAY, margin: 0, valign: "top"
    });
  });
}

// ── Slides 4-5: Trail Cards (config-driven template) ─────────────────────

function addTrailSlide(title, trails) {
  const s = pres.addSlide({ masterName: "MASTER" });
  addHeader(s, title);
  addFooter(s);

  let y = BODY_TOP;
  const cols = trails.length <= 4 ? 2 : 3;
  const gap = 0.15;
  const cardW = (BODY_W - gap * (cols - 1)) / cols;
  const rows = Math.ceil(trails.length / cols);
  const cardH = (CONTENT_BOTTOM - y - gap * (rows - 1)) / rows;

  trails.forEach((trail, i) => {
    const col = i % cols;
    const row = Math.floor(i / cols);
    const cx = PAD + col * (cardW + gap);
    const cy = y + row * (cardH + gap);
    addTrailCard(s, trail, cx, cy, cardW, cardH);
  });
}

addTrailSlide("Pisgah Favorites", [
  { name: "Bennett Gap", distance: 8.2, elevation: 1400, difficulty: "hard", desc: "Classic Pisgah. Rocky singletrack descent through old-growth forest with multiple creek crossings." },
  { name: "Farlow Gap", distance: 5.0, elevation: 800, difficulty: "expert", desc: "Notorious rock garden descent. Steep, loose, and technical — the benchmark for Pisgah gnar." },
  { name: "Pilot Rock", distance: 7.5, elevation: 1200, difficulty: "moderate", desc: "Flowing ridgeline trail with panoramic views. Good climbing workout with rewarding descent." },
  { name: "Butter Gap", distance: 4.8, elevation: 600, difficulty: "moderate", desc: "Smooth and flowy for Pisgah standards. Gentle grades make this a solid introduction to the forest." },
  { name: "Black Mountain", distance: 12.0, elevation: 2400, difficulty: "hard", desc: "Long loop with sustained climbing to 5,800 ft. Remote feel with dense rhododendron tunnels." },
  { name: "Spencer Branch", distance: 3.5, elevation: 500, difficulty: "moderate", desc: "Short but sweet connector trail. Playful terrain with small drops and bermed corners." }
]);

addTrailSlide("DuPont Favorites", [
  { name: "Ridgeline Trail", distance: 5.5, elevation: 600, difficulty: "easy", desc: "Gentle singletrack along the ridge. Smooth surfaces, filtered views, perfect for warming up." },
  { name: "Jim Branch", distance: 3.2, elevation: 400, difficulty: "moderate", desc: "Fast and flowy with bermed turns. Ends near Triple Falls — bring a camera." },
  { name: "Big Rock Trail", distance: 2.8, elevation: 350, difficulty: "moderate", desc: "Technical rock features with optional lines. Great for practicing skills in a short loop." },
  { name: "Hooker Creek", distance: 4.0, elevation: 500, difficulty: "easy", desc: "Smooth fire road to singletrack. Passes Hooker Falls — a popular swimming hole in summer." },
  { name: "Cedar Rock Trail", distance: 3.0, elevation: 700, difficulty: "hard", desc: "Exposed granite slabs and tight switchbacks. Stunning views from the rock face." },
  { name: "Corn Mill Shoals", distance: 2.5, elevation: 200, difficulty: "easy", desc: "Flat riverside trail with gentle rollers. Family-friendly and great for a quick after-work ride." }
]);

// ── Slide 6: Gear Checklist ───────────────────────────────────────────────
{
  const s = pres.addSlide({ masterName: "MASTER" });
  addHeader(s, "Essential Gear");
  addFooter(s);

  let y = BODY_TOP;
  y = addSectionLabel(s, "What to bring on every ride", y);

  const { leftX, rightX, colW } = twoColumnLayout();

  const leftItems = [
    "Full-face or half-shell helmet (full-face for Pisgah tech)",
    "Hydration pack — 2L minimum, 3L for long rides",
    "Tubeless tire setup with sealant + spare tube as backup",
    "Multi-tool with chain breaker",
    "First aid kit with tick removal tool"
  ];

  const rightItems = [
    "Knee pads for rocky descents (highly recommended)",
    "Dropper seatpost — non-negotiable on steep terrain",
    "Trail map or GPS device (cell service is spotty in Pisgah)",
    "Bear spray if riding deep in Pisgah (rare but real)",
    "Snacks and electrolytes for rides over 2 hours"
  ];

  y = addSectionLabel(s, "Bike & safety", y, { x: leftX, w: colW, color: FOREST, borderColor: MOSS });
  addBullets(s, leftItems, leftX, y, colW, 2.5, 10);

  let ry = BODY_TOP + 0.28;
  ry = addSectionLabel(s, "Accessories & provisions", ry, { x: rightX, w: colW, color: FOREST, borderColor: MOSS });
  addBullets(s, rightItems, rightX, ry, colW, 2.5, 10);
}

// ── Slide 7: Season Guide ─────────────────────────────────────────────────
{
  const s = pres.addSlide({ masterName: "MASTER" });
  addHeader(s, "When to Ride");
  addFooter(s);

  let y = BODY_TOP;
  y = addSectionLabel(s, "Seasonal conditions", y);

  const seasons = [
    { name: "Spring (Mar-May)", conditions: "Wildflowers and waterfalls at peak. Trails can be muddy — check conditions after rain. Creek crossings may run high.", rating: "Great", ratingColor: MOSS },
    { name: "Summer (Jun-Aug)", conditions: "Hot and humid at lower elevations. Ride early or stick to high-altitude trails. Afternoon thunderstorms are common.", rating: "Good", ratingColor: BLUE },
    { name: "Fall (Sep-Nov)", conditions: "Peak season. Perfect temps, low humidity, spectacular foliage. Trails are dry and fast. Book lodging early.", rating: "Best", ratingColor: FOREST },
    { name: "Winter (Dec-Feb)", conditions: "Rideable on dry days. Higher trails may have ice. Shorter daylight hours — plan accordingly. Pisgah can get snow above 4,000 ft.", rating: "Fair", ratingColor: DIRT }
  ];

  const rowH = 0.85;
  seasons.forEach((sn, i) => {
    const sy = y + i * (rowH + 0.1);
    s.addShape(pres.shapes.RECTANGLE, {
      x: PAD, y: sy, w: BODY_W, h: rowH,
      fill: { color: i % 2 === 0 ? CREAM : WHITE }, line: { color: MGRAY }
    });
    // Season name
    s.addText(sn.name, {
      x: PAD + 0.12, y: sy + 0.08, w: 2.2, h: 0.25,
      fontSize: 12, fontFace: "Georgia", bold: true, color: BLACK, margin: 0
    });
    // Rating chip
    s.addShape(pres.shapes.RECTANGLE, {
      x: PAD + 0.12, y: sy + 0.4, w: 0.7, h: 0.22,
      fill: { color: sn.ratingColor }
    });
    s.addText(sn.rating, {
      x: PAD + 0.12, y: sy + 0.4, w: 0.7, h: 0.22,
      fontSize: 9, fontFace: "Calibri", bold: true, color: WHITE,
      align: "center", valign: "middle", margin: 0
    });
    // Conditions
    s.addText(sn.conditions, {
      x: PAD + 2.5, y: sy + 0.08, w: BODY_W - 2.74, h: rowH - 0.16,
      fontSize: 10, fontFace: "Calibri", color: DGRAY, margin: 0,
      valign: "middle", lineSpacingMultiple: 1.25
    });
  });
}

// ── Slide 8: Closing ──────────────────────────────────────────────────────
{
  slideNum++;
  const s = pres.addSlide();
  s.background = { color: FOREST };

  s.addText("Get Out\nand Ride", {
    x: 0.8, y: 1.2, w: 8, h: 1.8,
    fontSize: 48, fontFace: "Georgia", bold: true, color: WHITE,
    lineSpacingMultiple: 1.1, margin: 0
  });

  s.addText("The best trail is the one you're on today.", {
    x: 0.8, y: 3.2, w: 7, h: 0.4,
    fontSize: 16, fontFace: "Calibri", italic: true, color: MOSS,
    margin: 0
  });

  s.addText("trailforks.com/region/pisgah  |  dupontforest.com", {
    x: 0.8, y: 4.2, w: 7, h: 0.3,
    fontSize: 11, fontFace: "Calibri", color: "88AA88",
    margin: 0
  });
}

// ── Write output ──────────────────────────────────────────────────────────
pres.writeFile({ fileName: "output/mtb-trail-guide.pptx" })
  .then(() => console.log("Done: output/mtb-trail-guide.pptx"))
  .catch(e => console.error(e));
