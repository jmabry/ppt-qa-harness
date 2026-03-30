const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Saturday Morning Cooking Class — Brunch Edition";

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

// Warm kitchen palette
const TERRA = "B85042";
const SAGE = "84B59F";
const CREAM = "FFF5E6";
const BUTTER = "F2D680";
const SPICE = "D4763C";
const LINEN = "FAF3E8";
const WHITE = "FFFFFF";
const BLACK = "1A1A1A";
const DGRAY = "444444";
const SGRAY = "777777";
const MGRAY = "CCCCCC";

let slideNum = 0;

// ── Layer 1: Utilities ────────────────────────────────────────────────────

function trimText(text, maxChars) {
  if (text.length <= maxChars) return text;
  let cut = text.lastIndexOf(" ", maxChars);
  if (cut < maxChars * 0.6) cut = maxChars;
  return text.slice(0, cut).replace(/[,;:\s]+$/, "");
}

function fitBullets(items, maxItems, maxChars) {
  return items.slice(0, maxItems).map(b => trimText(b, maxChars));
}

function estimateLines(text, fontSize, boxW) {
  const charsPerLine = Math.floor(boxW * 72 / (fontSize * 0.6));
  return Math.ceil(text.length / charsPerLine);
}

// ── Slide master ──────────────────────────────────────────────────────────
pres.defineSlideMaster({
  title: "MASTER",
  background: { color: LINEN },
  objects: []
});

// ── Layer 2: Helpers ──────────────────────────────────────────────────────

function addHeader(slide, title) {
  slide.addText(title, {
    x: PAD, y: 0.06, w: BODY_W - 2.5, h: TITLE_H,
    fontSize: 22, fontFace: "Georgia", bold: true, color: TERRA,
    valign: "bottom", margin: 0
  });
  slide.addText("Brunch Edition", {
    x: W - PAD - 1.5, y: 0.14, w: 1.5, h: 0.2,
    fontSize: 9, fontFace: "Calibri", italic: true, color: SGRAY,
    align: "right", valign: "middle", margin: 0
  });
}

function addFooter(slide) {
  slideNum++;
  slide.addShape(pres.shapes.LINE, {
    x: PAD, y: FOOTER_Y, w: BODY_W, h: 0,
    line: { color: BUTTER, width: 1.5 }
  });
  slide.addText(String(slideNum), {
    x: W - PAD - 0.5, y: FOOTER_Y + 0.02, w: 0.5, h: 0.22,
    fontSize: 8, fontFace: "Calibri", color: SGRAY,
    align: "right", valign: "middle", margin: 0
  });
}

function addSectionLabel(slide, text, y, opts = {}) {
  const x = opts.x !== undefined ? opts.x : PAD;
  const w = opts.w || BODY_W;
  slide.addText(text.toUpperCase(), {
    x, y, w, h: 0.2,
    fontSize: 9, fontFace: "Calibri", bold: true, color: opts.color || SPICE,
    charSpacing: 1, margin: 0
  });
  slide.addShape(pres.shapes.LINE, {
    x, y: y + 0.2, w, h: 0,
    line: { color: opts.borderColor || BUTTER, width: 0.75 }
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

// ── Config-driven template: recipe slide ──────────────────────────────────

function addRecipeSlide(recipe) {
  const s = pres.addSlide({ masterName: "MASTER" });
  addHeader(s, recipe.name);
  addFooter(s);

  let y = BODY_TOP;

  // Difficulty + time badges
  const badges = [
    { label: recipe.difficulty, color: recipe.difficulty === "Beginner" ? SAGE : recipe.difficulty === "Intermediate" ? SPICE : TERRA },
    { label: recipe.time, color: SGRAY },
    { label: "Serves " + recipe.serves, color: SAGE }
  ];

  let bx = PAD;
  badges.forEach(b => {
    const bw = b.label.length * 0.07 + 0.3;
    s.addShape(pres.shapes.RECTANGLE, {
      x: bx, y, w: bw, h: 0.24,
      fill: { color: b.color }
    });
    s.addText(b.label, {
      x: bx, y, w: bw, h: 0.24,
      fontSize: 9, fontFace: "Calibri", bold: true, color: WHITE,
      align: "center", valign: "middle", margin: 0
    });
    bx += bw + 0.08;
  });
  y += 0.38;

  // Description
  s.addText(recipe.description, {
    x: PAD, y, w: BODY_W, h: 0.35,
    fontSize: 10, fontFace: "Calibri", italic: true, color: DGRAY,
    margin: 0, valign: "top", lineSpacingMultiple: 1.2
  });
  y += 0.42;

  const { leftX, rightX, colW } = twoColumnLayout();

  // Left: ingredients
  y = addSectionLabel(s, "Ingredients", y, { x: leftX, w: colW });
  addBullets(s, recipe.ingredients, leftX, y, colW, 3.0, 10);

  // Right: steps
  let ry = y - 0.28;
  ry = addSectionLabel(s, "Method", ry, { x: rightX, w: colW, color: TERRA });

  const steps = recipe.steps.map((step, i) => ({
    text: `${i + 1}.  ${step}`,
    options: { breakLine: i < recipe.steps.length - 1 }
  }));
  s.addText(steps, {
    x: rightX, y: ry, w: colW, h: 3.0,
    fontSize: 10, fontFace: "Calibri", color: BLACK,
    valign: "top", lineSpacingMultiple: 1.3, margin: 0
  });

  // Chef's tip
  if (recipe.tip) {
    const tipY = CONTENT_BOTTOM - 0.5;
    s.addShape(pres.shapes.RECTANGLE, {
      x: PAD, y: tipY, w: BODY_W, h: 0.45,
      fill: { color: CREAM }, line: { color: BUTTER }
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: PAD, y: tipY, w: 0.06, h: 0.45,
      fill: { color: SPICE }, line: { color: SPICE }
    });
    s.addText("Chef's Tip:  " + recipe.tip, {
      x: PAD + 0.16, y: tipY, w: BODY_W - 0.28, h: 0.45,
      fontSize: 9, fontFace: "Calibri", italic: true, color: DGRAY,
      valign: "middle", margin: 0
    });
  }
}

// ── Slide 1: Cover ────────────────────────────────────────────────────────
{
  slideNum++;
  const s = pres.addSlide();
  s.background = { color: TERRA };

  s.addText("Saturday Morning\nCooking Class", {
    x: 0.8, y: 0.8, w: 8, h: 1.8,
    fontSize: 44, fontFace: "Georgia", bold: true, color: WHITE,
    lineSpacingMultiple: 1.1, margin: 0
  });

  s.addText("Brunch Edition — Five Recipes to Master", {
    x: 0.8, y: 2.7, w: 7, h: 0.4,
    fontSize: 18, fontFace: "Calibri", color: BUTTER,
    margin: 0
  });

  s.addText("Community Kitchen  |  10:00 AM - 1:00 PM", {
    x: 0.8, y: 3.3, w: 7, h: 0.3,
    fontSize: 12, fontFace: "Calibri", color: "DDAA99",
    margin: 0
  });
}

// ── Slide 2: Class Overview ───────────────────────────────────────────────
{
  const s = pres.addSlide({ masterName: "MASTER" });
  addHeader(s, "Today's Menu");
  addFooter(s);

  let y = BODY_TOP;
  y = addSectionLabel(s, "Five dishes, three hours, one delicious brunch", y);

  const dishes = [
    { name: "Shakshuka", desc: "Eggs poached in spiced tomato sauce", diff: "Beginner", time: "25 min", color: TERRA },
    { name: "Buttermilk Biscuits", desc: "Flaky, buttery, and still warm", diff: "Intermediate", time: "35 min", color: SPICE },
    { name: "Eggs Benedict", desc: "Poached eggs and foolproof hollandaise", diff: "Advanced", time: "40 min", color: TERRA },
    { name: "Lemon Ricotta Pancakes", desc: "Light, fluffy, and bright with citrus", diff: "Beginner", time: "20 min", color: SAGE },
    { name: "Smoked Salmon Board", desc: "Assembly, presentation, and garnish art", diff: "Beginner", time: "15 min", color: SAGE }
  ];

  const cardW = (BODY_W - 0.12 * 2) / 3;
  const cardH = 1.1;

  dishes.forEach((d, i) => {
    const col = i < 3 ? i : i - 3;
    const row = i < 3 ? 0 : 1;
    const cx = PAD + col * (cardW + 0.12);
    const cy = y + row * (cardH + 0.12);

    s.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: cy, w: cardW, h: cardH,
      fill: { color: WHITE }, line: { color: MGRAY }
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: cy, w: cardW, h: 0.05,
      fill: { color: d.color }, line: { color: d.color }
    });
    s.addText(d.name, {
      x: cx + 0.12, y: cy + 0.12, w: cardW - 0.24, h: 0.25,
      fontSize: 13, fontFace: "Georgia", bold: true, color: BLACK, margin: 0
    });
    s.addText(d.desc, {
      x: cx + 0.12, y: cy + 0.4, w: cardW - 0.24, h: 0.3,
      fontSize: 10, fontFace: "Calibri", color: DGRAY, margin: 0
    });
    s.addText(`${d.diff}  |  ${d.time}`, {
      x: cx + 0.12, y: cy + 0.75, w: cardW - 0.24, h: 0.2,
      fontSize: 9, fontFace: "Calibri", bold: true, color: d.color, margin: 0
    });
  });
}

// ── Slides 3-7: Individual recipes (config-driven) ───────────────────────

addRecipeSlide({
  name: "Shakshuka",
  difficulty: "Beginner",
  time: "25 min",
  serves: "4",
  description: "North African comfort food — eggs gently poached in a cumin-spiced tomato sauce. Serve straight from the skillet with crusty bread.",
  ingredients: [
    "2 tbsp olive oil",
    "1 onion, diced",
    "1 red bell pepper, diced",
    "3 cloves garlic, minced",
    "1 tsp cumin, 1 tsp paprika, pinch cayenne",
    "1 can (28 oz) crushed tomatoes",
    "6 large eggs",
    "Fresh cilantro and crumbled feta for topping",
    "Crusty bread for serving"
  ],
  steps: [
    "Heat olive oil in a 12-inch skillet over medium heat",
    "Cook onion and pepper until soft, about 5 minutes",
    "Add garlic and spices, stir 30 seconds until fragrant",
    "Pour in tomatoes, simmer 10 minutes until thickened",
    "Make 6 wells, crack an egg into each",
    "Cover and cook 5-7 min until whites set, yolks runny",
    "Top with cilantro and feta, serve immediately"
  ],
  tip: "Don't stir after adding eggs. Low heat keeps yolks runny — check at 5 minutes."
});

addRecipeSlide({
  name: "Buttermilk Biscuits",
  difficulty: "Intermediate",
  time: "35 min",
  serves: "8-10",
  description: "Tall, flaky, and golden. The secret is cold butter, minimal handling, and a hot oven. These disappear fast.",
  ingredients: [
    "2 cups all-purpose flour, plus more for dusting",
    "1 tbsp baking powder",
    "1/4 tsp baking soda",
    "1 tsp salt, 1 tsp sugar",
    "6 tbsp cold unsalted butter, cubed",
    "3/4 cup cold buttermilk",
    "2 tbsp melted butter for brushing"
  ],
  steps: [
    "Preheat oven to 450F. Line a baking sheet with parchment",
    "Whisk flour, baking powder, soda, salt, and sugar",
    "Cut in cold butter until pea-sized pieces remain",
    "Add buttermilk, stir just until dough comes together",
    "Pat to 1-inch thick on floured surface (don't roll!)",
    "Cut with 2.5-inch cutter — press straight down, don't twist",
    "Bake 12-14 minutes until golden, brush with melted butter"
  ],
  tip: "Freeze butter cubes 10 minutes before cutting in. Warm butter = flat biscuits."
});

addRecipeSlide({
  name: "Eggs Benedict",
  difficulty: "Advanced",
  time: "40 min",
  serves: "4",
  description: "The brunch classic, demystified. We'll nail poached eggs and build a foolproof blender hollandaise that won't break.",
  ingredients: [
    "4 English muffins, split and toasted",
    "8 slices Canadian bacon or ham",
    "8 large eggs (the freshest you can find)",
    "White vinegar for poaching water",
    "3 egg yolks (for hollandaise)",
    "1 tbsp lemon juice",
    "1/2 cup melted butter, still hot",
    "Pinch of cayenne, salt to taste",
    "Chives for garnish"
  ],
  steps: [
    "Blender hollandaise: blend yolks + lemon 5 sec, stream in hot butter",
    "Season with cayenne and salt, keep blender jar in warm water",
    "Bring a wide pot of water to bare simmer, add splash of vinegar",
    "Crack each egg into a small cup, slide gently into water",
    "Poach 3-4 minutes for runny yolks, remove with slotted spoon",
    "Toast muffins, warm Canadian bacon in a skillet",
    "Stack: muffin, bacon, egg, hollandaise, chives"
  ],
  tip: "Swirl the water before dropping each egg — the vortex wraps the white around the yolk."
});

addRecipeSlide({
  name: "Lemon Ricotta Pancakes",
  difficulty: "Beginner",
  time: "20 min",
  serves: "4",
  description: "Light and pillowy with a bright citrus note. The ricotta adds richness without heaviness. Top with fresh berries and a dusting of powdered sugar.",
  ingredients: [
    "1 cup ricotta cheese",
    "2 large eggs, separated",
    "3/4 cup milk",
    "Zest of 1 lemon + 1 tbsp juice",
    "1 cup flour",
    "2 tsp baking powder",
    "2 tbsp sugar, pinch of salt",
    "Butter for the griddle",
    "Fresh berries and powdered sugar for serving"
  ],
  steps: [
    "Whisk ricotta, egg yolks, milk, lemon zest, and juice",
    "Add flour, baking powder, sugar, and salt — stir gently",
    "Beat egg whites to soft peaks, fold into batter",
    "Heat griddle to 325F, butter lightly",
    "Pour 1/4 cup batter per pancake, cook until bubbles form",
    "Flip once, cook 1-2 more minutes",
    "Serve with berries and a light dusting of powdered sugar"
  ],
  tip: "Folding in whipped egg whites is what makes these cloud-like. Don't skip it, and don't over-mix."
});

addRecipeSlide({
  name: "Smoked Salmon Board",
  difficulty: "Beginner",
  time: "15 min",
  serves: "6-8",
  description: "No cooking required — just thoughtful assembly. A beautiful board is all about variety, contrast, and giving people choices.",
  ingredients: [
    "8 oz smoked salmon (lox-style)",
    "8 oz cream cheese, softened",
    "Everything bagels or bagel chips",
    "Capers, thinly sliced red onion",
    "Fresh dill, lemon wedges",
    "Cucumber slices, cherry tomatoes",
    "Cornichons or pickled anything",
    "Hard-boiled eggs, quartered",
    "Good olive oil and flaky salt"
  ],
  steps: [
    "Choose a large board or platter — wood looks great",
    "Place cream cheese in a small bowl off-center on the board",
    "Fan salmon slices in loose rosettes (don't flatten them)",
    "Cluster bagels/chips in one area — don't scatter",
    "Group garnishes: capers + onion together, herbs together",
    "Fill gaps with cucumbers, tomatoes, eggs, cornichons",
    "Drizzle olive oil on cream cheese, finish with dill and flaky salt"
  ],
  tip: "Board building rule: group similar items, vary colors side-by-side, and leave some breathing room."
});

// ── Slide 8: Closing ──────────────────────────────────────────────────────
{
  slideNum++;
  const s = pres.addSlide();
  s.background = { color: TERRA };

  s.addText("Now Let's\nCook.", {
    x: 0.8, y: 1.0, w: 8, h: 1.8,
    fontSize: 48, fontFace: "Georgia", bold: true, color: WHITE,
    lineSpacingMultiple: 1.1, margin: 0
  });

  s.addText("Taste everything. Season as you go. Have fun.", {
    x: 0.8, y: 3.0, w: 7, h: 0.4,
    fontSize: 16, fontFace: "Calibri", italic: true, color: BUTTER,
    margin: 0
  });

  s.addText("Recipes and notes at: cookingclass.example.com/brunch", {
    x: 0.8, y: 4.2, w: 7, h: 0.3,
    fontSize: 11, fontFace: "Calibri", color: "DDAA99",
    margin: 0
  });
}

// ── Write output ──────────────────────────────────────────────────────────
pres.writeFile({ fileName: "output/brunch-cooking-class.pptx" })
  .then(() => console.log("Done: output/brunch-cooking-class.pptx"))
  .catch(e => console.error(e));
