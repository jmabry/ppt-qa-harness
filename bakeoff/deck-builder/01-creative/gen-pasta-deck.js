/**
 * gen-pasta-deck.js
 * Generates "The Art of Homemade Pasta" presentation for a community cooking night.
 * Output: output/pasta-deck.pptx
 */
const fs = require("fs");
const path = require("path");
const pptxgen = require("../node_modules/pptxgenjs");

// ============================================================
// LAYER 1: Constants & Design Tokens
// ============================================================

const W = 10, H = 5.625;
const PAD = 0.5;
const INNER_PAD = 0.35;
const BODY_W = W - PAD * 2;
const TITLE_H = 0.5;
const BODY_TOP = 1.05;
const FOOTER_Y = 5.35;
const CONTENT_BOTTOM = FOOTER_Y - 0.12;
const SECTION_GAP = 0.14;
const MIN_FONT = 9;

// Warm earthy palette
const C = {
  terracotta:   "C2552B",
  terracottaDk: "9B3F1E",
  cream:        "FDF6EC",
  creamDark:    "F5E6CC",
  sage:         "7A8B6F",
  sageDark:     "5C6B52",
  sageLight:    "D4DEC8",
  charcoal:     "3B3129",
  warmGray:     "8C7E72",
  linen:        "FAF3E8",
  white:        "FFFFFF",
  gold:         "D4A847",
  goldLight:    "F2E5C0",
  warmBrown:    "6B4E3D",
};

const FONT = {
  heading: "Georgia",
  body:    "Calibri",
};

// ============================================================
// LAYER 2: Utilities
// ============================================================

function trimText(text, maxChars) {
  if (text.length <= maxChars) return text;
  const trimmed = text.slice(0, maxChars - 1).replace(/[\s,.;:!?]+$/, "");
  const lastSentence = trimmed.lastIndexOf(". ");
  if (lastSentence > maxChars * 0.5) return trimmed.slice(0, lastSentence + 1);
  return trimmed;
}

function fitBullets(items, max, chars) {
  const result = items.slice(0, max).map(i => trimText(i, chars));
  if (items.length > max) {
    console.log(`  [fitBullets] dropped ${items.length - max} items`);
  }
  return result;
}

function estimateLines(text, fontSize, boxW) {
  const charsPerLine = Math.floor(boxW * 72 / (fontSize * 0.6));
  return Math.ceil(text.length / charsPerLine);
}

function checkFit(label, text, fontSize, boxW, boxH, lineSpacing = 1.15) {
  const lines = estimateLines(text, fontSize, boxW);
  const neededH = lines * (fontSize / 72) * lineSpacing;
  if (neededH > boxH) {
    console.log(`  [OVERFLOW] "${label}" ~${neededH.toFixed(2)}" exceeds box ${boxH.toFixed(2)}"`);
  }
}

// ============================================================
// LAYER 2: Helpers
// ============================================================

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "The Art of Homemade Pasta";
pres.author = "Community Cooking Night";

// Shadow factory (avoid mutation bug)
const cardShadow = () => ({
  type: "outer", blur: 6, offset: 2, angle: 135, color: "000000", opacity: 0.12
});

const subtleShadow = () => ({
  type: "outer", blur: 3, offset: 1, angle: 135, color: "000000", opacity: 0.08
});

// Slide master
pres.defineSlideMaster({
  title: "WARM",
  background: { color: C.cream },
  objects: [],
});

pres.defineSlideMaster({
  title: "LINEN",
  background: { color: C.linen },
  objects: [],
});

// --- Helpers ---

function addHeader(slide, title, opts = {}) {
  const color = opts.color || C.charcoal;
  const fontFace = opts.fontFace || FONT.heading;
  slide.addText(title, {
    x: PAD, y: 0.3, w: BODY_W, h: TITLE_H,
    fontSize: opts.fontSize || 22, fontFace, color,
    bold: true, margin: 0, fit: "shrink",
  });
  // accent underline
  slide.addShape(pres.shapes.LINE, {
    x: PAD, y: 0.82, w: 1.2, h: 0,
    line: { color: opts.accent || C.terracotta, width: 2.5 },
  });
  return BODY_TOP;
}

function addFooter(slide, pageNum) {
  slide.addShape(pres.shapes.LINE, {
    x: PAD, y: FOOTER_Y, w: BODY_W, h: 0,
    line: { color: C.creamDark, width: 0.5 },
  });
  if (pageNum) {
    slide.addText(String(pageNum), {
      x: W - 1.0, y: FOOTER_Y + 0.02, w: 0.5, h: 0.2,
      fontSize: 8, fontFace: FONT.body, color: C.warmGray,
      align: "right", margin: 0,
    });
  }
  slide.addText("The Art of Homemade Pasta", {
    x: PAD, y: FOOTER_Y + 0.02, w: 3, h: 0.2,
    fontSize: 8, fontFace: FONT.body, color: C.warmGray,
    italic: true, margin: 0,
  });
}

function addSectionLabel(slide, text, y, opts = {}) {
  slide.addText(text.toUpperCase(), {
    x: opts.x || PAD, y, w: opts.w || BODY_W, h: 0.22,
    fontSize: 10, fontFace: FONT.body, color: opts.color || C.terracotta,
    bold: true, charSpacing: 3, margin: 0,
  });
  return y + 0.28;
}

function addBullets(slide, items, x, y, w, h, fs = 11) {
  const fontSize = Math.max(fs, MIN_FONT);
  const textArr = items.map((item, i) => ({
    text: item,
    options: {
      bullet: true,
      breakLine: i < items.length - 1,
      fontSize,
      fontFace: FONT.body,
      color: C.charcoal,
      lineSpacingMultiple: 1.25,
    },
  }));
  slide.addText(textArr, { x, y, w, h, margin: [0, 0, 0, 4], valign: "top" });
}

function addCard(slide, title, body, x, y, w, h, opts = {}) {
  const fillColor = opts.fill || C.white;
  const accentColor = opts.accent || C.terracotta;
  const accentH = 0.06;

  // card background
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h, fill: { color: fillColor }, rectRadius: 0.08,
    shadow: cardShadow(),
  });
  // top accent strip
  slide.addShape(pres.shapes.RECTANGLE, {
    x: x + 0.01, y, w: w - 0.02, h: accentH,
    fill: { color: accentColor },
  });
  // title
  const titleH = 0.3;
  slide.addText(title, {
    x: x + INNER_PAD * 0.6, y: y + accentH + 0.06, w: w - INNER_PAD * 1.2, h: titleH,
    fontSize: 12, fontFace: FONT.heading, color: C.charcoal,
    bold: true, margin: 0, valign: "top",
  });
  // body
  if (typeof body === "string") {
    slide.addText(body, {
      x: x + INNER_PAD * 0.6, y: y + accentH + 0.06 + titleH, w: w - INNER_PAD * 1.2,
      h: h - accentH - 0.06 - titleH - 0.1,
      fontSize: opts.bodyFs || 10, fontFace: FONT.body, color: C.warmBrown,
      margin: 0, valign: "top", lineSpacingMultiple: 1.2,
    });
  } else if (Array.isArray(body)) {
    addBullets(slide, body,
      x + INNER_PAD * 0.6, y + accentH + 0.06 + titleH,
      w - INNER_PAD * 1.2, h - accentH - 0.06 - titleH - 0.1,
      opts.bodyFs || 10
    );
  }
}

function addCardRow(slide, cards, x, y, w, h, cols = 3) {
  const gap = 0.18;
  const cardW = (w - gap * (cols - 1)) / cols;
  cards.forEach((card, i) => {
    const cx = x + i * (cardW + gap);
    addCard(slide, card.title, card.body, cx, y, cardW, h, card.opts || {});
  });
}

function twoColumnLayout(gap = 0.4) {
  const colW = (BODY_W - gap) / 2;
  return { leftX: PAD, rightX: PAD + colW + gap, colW };
}

function addChipLabel(slide, text, x, y, opts = {}) {
  const chipW = text.length * 0.085 + 0.35;
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w: chipW, h: 0.26,
    fill: { color: opts.fill || C.sageLight },
    rectRadius: 0.13,
  });
  slide.addText(text, {
    x, y, w: chipW, h: 0.26,
    fontSize: 9, fontFace: FONT.body, color: opts.color || C.sageDark,
    bold: true, align: "center", valign: "middle", margin: 0,
  });
  return chipW;
}

function addIconCircle(slide, emoji, x, y, size, bgColor) {
  slide.addShape(pres.shapes.OVAL, {
    x, y, w: size, h: size,
    fill: { color: bgColor },
  });
  slide.addText(emoji, {
    x, y, w: size, h: size,
    fontSize: Math.round(size * 20),
    align: "center", valign: "middle", margin: 0,
  });
}

// ============================================================
// LAYER 3: Slide Definitions
// ============================================================

let pageNum = 0;

// --- SLIDE 1: Cover ---
{
  const s = pres.addSlide();
  s.background = { color: C.charcoal };

  // Large decorative background shape
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: W, h: H,
    fill: { color: C.terracottaDk, transparency: 85 },
  });

  // Accent bar left
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.12, h: H,
    fill: { color: C.terracotta },
  });

  // decorative horizontal line
  s.addShape(pres.shapes.LINE, {
    x: 0.8, y: 2.9, w: 2.5, h: 0,
    line: { color: C.gold, width: 1.5 },
  });

  let y = 1.2;
  // Overline
  s.addText("COMMUNITY COOKING NIGHT", {
    x: 0.8, y, w: 7, h: 0.3,
    fontSize: 11, fontFace: FONT.body, color: C.gold,
    charSpacing: 5, margin: 0,
  });
  y += 0.45;

  // Title
  s.addText("The Art of\nHomemade Pasta", {
    x: 0.8, y, w: 7, h: 1.2,
    fontSize: 40, fontFace: FONT.heading, color: C.cream,
    bold: true, margin: 0, lineSpacingMultiple: 1.05,
  });
  y += 1.35;

  // Subtitle
  s.addText("From flour to fork: master the craft that turns\nsimple ingredients into extraordinary meals", {
    x: 0.8, y: 3.1, w: 6, h: 0.6,
    fontSize: 13, fontFace: FONT.body, color: C.creamDark,
    margin: 0, lineSpacingMultiple: 1.3,
  });

  // Decorative card in bottom right
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 6.8, y: 3.6, w: 2.7, h: 1.5,
    fill: { color: C.terracotta, transparency: 70 },
    rectRadius: 0.1,
  });
  s.addText("8 slides\n3 shapes\n1 passion", {
    x: 6.8, y: 3.6, w: 2.7, h: 1.5,
    fontSize: 16, fontFace: FONT.heading, color: C.cream,
    align: "center", valign: "middle", margin: 0,
    lineSpacingMultiple: 1.5,
  });

  // Bottom bar
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: H - 0.08, w: W, h: 0.08,
    fill: { color: C.terracotta },
  });
}

// --- SLIDE 2: Why Fresh Pasta Beats Dried ---
{
  pageNum++;
  const s = pres.addSlide({ masterName: "WARM" });
  let y = addHeader(s, "Why Fresh Pasta Beats Dried");
  addFooter(s, pageNum);

  const { leftX, rightX, colW } = twoColumnLayout(0.35);

  // Left: Fresh pasta column
  y = addSectionLabel(s, "Fresh Pasta", y, { x: leftX, w: colW, color: C.sage });

  const freshPoints = [
    "Silky, tender texture that melts on the tongue",
    "Cooks in 2-3 minutes vs. 10-12 for dried",
    "Absorbs sauces better with its porous surface",
    "Richer flavor from fresh eggs and quality flour",
    "Impressive wow factor for guests",
  ];
  addBullets(s, freshPoints, leftX, y, colW, 2.8, 10);

  // Right: Dried pasta column
  let yR = addSectionLabel(s, "Dried Pasta", BODY_TOP, { x: rightX, w: colW, color: C.warmGray });

  const driedPoints = [
    "Firmer bite, holds up in baked dishes",
    "Months of shelf life, always available",
    "Consistent results every time",
    "Better for chunky, oil-based sauces",
    "A pantry staple, not a competitor",
  ];
  addBullets(s, driedPoints, rightX, yR, colW, 2.8, 10);

  // Bottom callout
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: PAD, y: 4.2, w: BODY_W, h: 0.7,
    fill: { color: C.sageLight }, rectRadius: 0.08,
  });
  s.addText("The takeaway: fresh pasta isn't better than dried, it's different. Tonight, you'll learn when each one shines.", {
    x: PAD + 0.3, y: 4.2, w: BODY_W - 0.6, h: 0.7,
    fontSize: 11, fontFace: FONT.body, color: C.sageDark,
    italic: true, valign: "middle", margin: 0,
  });
}

// --- SLIDE 3: Equipment Needed ---
{
  pageNum++;
  const s = pres.addSlide({ masterName: "LINEN" });
  let y = addHeader(s, "Equipment You'll Need");
  addFooter(s, pageNum);

  const { leftX, rightX, colW } = twoColumnLayout(0.35);

  // Essential column (card style)
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: leftX, y, w: colW, h: 3.8,
    fill: { color: C.white }, rectRadius: 0.1,
    shadow: cardShadow(),
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: leftX + 0.01, y, w: colW - 0.02, h: 0.06,
    fill: { color: C.terracotta },
  });

  let yL = y + 0.2;
  addChipLabel(s, "ESSENTIAL", leftX + 0.3, yL, { fill: C.terracotta, color: C.white });
  yL += 0.45;

  const essentials = [
    "Large wooden cutting board or clean countertop",
    "Rolling pin (longer is better, 18\"+ ideal)",
    "Sharp knife or bench scraper",
    "Fork for mixing dough in the well",
    "Large pot with plenty of salted water",
    "Digital kitchen scale for precise ratios",
  ];
  addBullets(s, essentials, leftX + 0.25, yL, colW - 0.5, 2.9, 10);

  // Nice-to-have column
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: rightX, y, w: colW, h: 3.8,
    fill: { color: C.white }, rectRadius: 0.1,
    shadow: cardShadow(),
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: rightX + 0.01, y, w: colW - 0.02, h: 0.06,
    fill: { color: C.sage },
  });

  let yR = y + 0.2;
  addChipLabel(s, "NICE TO HAVE", rightX + 0.3, yR, { fill: C.sage, color: C.white });
  yR += 0.45;

  const niceToHave = [
    "Pasta machine (hand-crank or KitchenAid attachment)",
    "Drying rack or wooden dowel for hanging",
    "Semolina flour for dusting (prevents sticking)",
    "Ravioli stamp or mold for uniform shapes",
    "Spray bottle for keeping dough moist",
  ];
  addBullets(s, niceToHave, rightX + 0.25, yR, colW - 0.5, 2.9, 10);
}

// --- SLIDE 4: Tagliatelle ---
{
  pageNum++;
  const s = pres.addSlide({ masterName: "WARM" });
  let y = addHeader(s, "Shape 1: Tagliatelle");
  addFooter(s, pageNum);

  // Left: description panel
  const { leftX, rightX, colW } = twoColumnLayout(0.3);

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: leftX, y, w: colW, h: 1.2,
    fill: { color: C.goldLight }, rectRadius: 0.08,
  });
  s.addText("The classic flat ribbon pasta. Wide enough to carry rich ragu, elegant enough for a dinner party. This is the shape that teaches you to roll evenly.", {
    x: leftX + 0.2, y, w: colW - 0.4, h: 1.2,
    fontSize: 10.5, fontFace: FONT.body, color: C.warmBrown,
    margin: 0, valign: "middle", lineSpacingMultiple: 1.3,
  });

  let yTech = y + 1.35;
  yTech = addSectionLabel(s, "Technique Tips", yTech, { x: leftX, w: colW });

  const tagliatelleTips = [
    "Roll dough to 1mm thickness, translucent when held up to light",
    "Flour generously, then roll into a loose log before cutting",
    "Cut ribbons 8-10mm wide with a sharp, decisive stroke",
    "Shake out nests immediately to prevent sticking",
    "Cook for just 90 seconds in rapidly boiling water",
  ];
  addBullets(s, tagliatelleTips, leftX, yTech, colW, 2.2, 10);

  // Right: visual info cards
  addCard(s, "Difficulty", "Beginner-friendly. If you can roll and cut, you can make tagliatelle.", rightX, y, colW, 1.1, { accent: C.sage, bodyFs: 10 });
  addCard(s, "Best Paired With", "Bolognese ragu, butter & sage, mushroom cream sauce, truffle oil", rightX, y + 1.25, colW, 1.0, { accent: C.gold, bodyFs: 10 });
  addCard(s, "Common Pitfall", "Rolling too thick. If you can't almost see your hand through it, keep rolling.", rightX, y + 2.4, colW, 1.0, { accent: C.terracotta, bodyFs: 10 });
}

// --- SLIDE 5: Ravioli ---
{
  pageNum++;
  const s = pres.addSlide({ masterName: "LINEN" });
  let y = addHeader(s, "Shape 2: Ravioli");
  addFooter(s, pageNum);

  const { leftX, rightX, colW } = twoColumnLayout(0.3);

  // Left side: description + tips
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: leftX, y, w: colW, h: 1.2,
    fill: { color: C.sageLight }, rectRadius: 0.08,
  });
  s.addText("The filled pasta that impresses everyone. A little packet of flavor that turns simple ricotta and spinach into something restaurant-worthy.", {
    x: leftX + 0.2, y, w: colW - 0.4, h: 1.2,
    fontSize: 10.5, fontFace: FONT.body, color: C.warmBrown,
    margin: 0, valign: "middle", lineSpacingMultiple: 1.3,
  });

  let yTech = y + 1.35;
  yTech = addSectionLabel(s, "Technique Tips", yTech, { x: leftX, w: colW });

  const ravioliTips = [
    "Keep dough thin: thick edges ruin the bite",
    "Chill filling for 30 min so it holds its shape",
    "Brush water around each mound for a tight seal",
    "Press out ALL air pockets before crimping edges",
    "Use a fork to crimp, not just pinch with fingers",
  ];
  addBullets(s, ravioliTips, leftX, yTech, colW, 2.2, 10);

  // Right: stacked cards
  addCard(s, "Difficulty", "Intermediate. The folding and sealing take practice, but mistakes taste great anyway.", rightX, y, colW, 1.1, { accent: C.terracotta, bodyFs: 10 });
  addCard(s, "Classic Fillings", "Ricotta & spinach, butternut squash & sage, mushroom & thyme, lobster & mascarpone", rightX, y + 1.25, colW, 1.0, { accent: C.gold, bodyFs: 10 });
  addCard(s, "Common Pitfall", "Overfilling. Less is more. A teaspoon per ravioli keeps seals tight.", rightX, y + 2.4, colW, 1.0, { accent: C.sage, bodyFs: 10 });
}

// --- SLIDE 6: Orecchiette ---
{
  pageNum++;
  const s = pres.addSlide({ masterName: "WARM" });
  let y = addHeader(s, "Shape 3: Orecchiette");
  addFooter(s, pageNum);

  const { leftX, rightX, colW } = twoColumnLayout(0.3);

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: leftX, y, w: colW, h: 1.2,
    fill: { color: C.goldLight }, rectRadius: 0.08,
  });
  s.addText("\"Little ears\" from Puglia. No rolling pin, no machine, just your thumb and a butter knife. The most tactile and meditative pasta shape to make.", {
    x: leftX + 0.2, y, w: colW - 0.4, h: 1.2,
    fontSize: 10.5, fontFace: FONT.body, color: C.warmBrown,
    margin: 0, valign: "middle", lineSpacingMultiple: 1.3,
  });

  let yTech = y + 1.35;
  yTech = addSectionLabel(s, "Technique Tips", yTech, { x: leftX, w: colW });

  const orecchietteTips = [
    "Use semolina + water dough (no egg) for the right chew",
    "Roll into a rope, cut small coins about 2cm wide",
    "Drag each coin toward you with a butter knife",
    "Flip over your thumb to create the ear shape",
    "Rough texture is a feature, not a bug: it grabs sauce",
  ];
  addBullets(s, orecchietteTips, leftX, yTech, colW, 2.2, 10);

  addCard(s, "Difficulty", "Easy technique, but slow. Perfect for a relaxing Sunday afternoon with music and wine.", rightX, y, colW, 1.1, { accent: C.sage, bodyFs: 10 });
  addCard(s, "Best Paired With", "Broccoli rabe & sausage, cherry tomato & ricotta salata, pesto Genovese", rightX, y + 1.25, colW, 1.0, { accent: C.gold, bodyFs: 10 });
  addCard(s, "Common Pitfall", "Making them too thick. Thin out the center so it cooks evenly with the rim.", rightX, y + 2.4, colW, 1.0, { accent: C.terracotta, bodyFs: 10 });
}

// --- SLIDE 7: Common Mistakes & Fixes ---
{
  pageNum++;
  const s = pres.addSlide({ masterName: "LINEN" });
  let y = addHeader(s, "Common Mistakes & How to Fix Them");
  addFooter(s, pageNum);

  const mistakes = [
    {
      title: "Dough Too Dry",
      body: "Add water a teaspoon at a time. Humidity affects flour, so never rely on the recipe amount alone.",
      opts: { accent: C.terracotta },
    },
    {
      title: "Dough Too Sticky",
      body: "Dust with semolina (not AP flour) and knead more. The gluten needs 10 full minutes to develop.",
      opts: { accent: C.warmBrown },
    },
    {
      title: "Pasta Tears When Rolling",
      body: "Rest the dough 30 minutes under a towel. Gluten relaxation is non-negotiable.",
      opts: { accent: C.sage },
    },
  ];
  addCardRow(s, mistakes, PAD, y, BODY_W, 1.5, 3);

  y += 1.65;
  const mistakes2 = [
    {
      title: "Ravioli Burst Open",
      body: "Seal edges completely and remove air. Drop gently into simmering (not boiling) water.",
      opts: { accent: C.gold },
    },
    {
      title: "Pasta Sticks Together",
      body: "Toss with semolina immediately after cutting. Use plenty of water when cooking: 1L per 100g.",
      opts: { accent: C.terracottaDk },
    },
    {
      title: "Bland Flavor",
      body: "Salt the dough AND the water generously. The water should taste like the sea. Save pasta water for sauce.",
      opts: { accent: C.sageDark },
    },
  ];
  addCardRow(s, mistakes2, PAD, y, BODY_W, 1.5, 3);
}

// --- SLIDE 8: Closing / Encouragement ---
{
  pageNum++;
  const s = pres.addSlide();
  s.background = { color: C.charcoal };

  // Accent bar
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.12, h: H,
    fill: { color: C.terracotta },
  });

  // Decorative line
  s.addShape(pres.shapes.LINE, {
    x: 0.8, y: 1.5, w: 2.5, h: 0,
    line: { color: C.gold, width: 1.5 },
  });

  // Main quote
  s.addText("Now, Let's Make Pasta.", {
    x: 0.8, y: 1.7, w: 8, h: 0.9,
    fontSize: 36, fontFace: FONT.heading, color: C.cream,
    bold: true, margin: 0,
  });

  s.addText("Every great pasta maker started exactly where you are tonight.\nThe dough will be imperfect. Some ravioli will burst. That's the point.\nPasta-making is a practice, not a performance.", {
    x: 0.8, y: 2.7, w: 7, h: 1.0,
    fontSize: 13, fontFace: FONT.body, color: C.creamDark,
    margin: 0, lineSpacingMultiple: 1.5,
  });

  // Three takeaway cards at the bottom
  const cardY = 4.0;
  const cardW = 2.5;
  const gap = 0.3;
  const startX = 0.8;

  const closingCards = [
    { label: "Touch", msg: "Feel the dough. Your hands will learn what your eyes can't.", accent: C.terracotta },
    { label: "Taste", msg: "Salt everything. Taste as you go. Adjust with confidence.", accent: C.sage },
    { label: "Share", msg: "The best pasta is the one you make for someone you love.", accent: C.gold },
  ];

  closingCards.forEach((c, i) => {
    const cx = startX + i * (cardW + gap);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: cx, y: cardY, w: cardW, h: 1.1,
      fill: { color: C.charcoal },
      line: { color: c.accent, width: 1 },
      rectRadius: 0.08,
    });
    s.addText(c.label.toUpperCase(), {
      x: cx, y: cardY + 0.12, w: cardW, h: 0.25,
      fontSize: 10, fontFace: FONT.body, color: c.accent,
      bold: true, align: "center", margin: 0, charSpacing: 4,
    });
    s.addText(c.msg, {
      x: cx + 0.15, y: cardY + 0.4, w: cardW - 0.3, h: 0.6,
      fontSize: 10, fontFace: FONT.body, color: C.creamDark,
      align: "center", valign: "top", margin: 0, lineSpacingMultiple: 1.2,
    });
  });

  // Bottom bar
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: H - 0.08, w: W, h: 0.08,
    fill: { color: C.terracotta },
  });
}

// ============================================================
// WRITE OUTPUT
// ============================================================

const outDir = path.join(__dirname, "output");
if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

pres.writeFile({ fileName: path.join(outDir, "pasta-deck.pptx") })
  .then(() => console.log("Created: output/pasta-deck.pptx"))
  .catch(err => { console.error("Error:", err); process.exit(1); });
