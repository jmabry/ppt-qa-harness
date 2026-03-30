const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Pack 42 Spring Campout — Lake Powhatan";

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

// Campfire palette
const NAVY = "1E3A5F";
const SKY = "5B9BD5";
const FLAME = "E87D2F";
const GOLD = "F2C14E";
const PINE = "3D7A54";
const CREAM = "FFF8EE";
const WHITE = "FFFFFF";
const BLACK = "1A1A1A";
const DGRAY = "444444";
const SGRAY = "777777";
const MGRAY = "CCCCCC";
const LGRAY = "F5F5F5";

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
  background: { color: WHITE },
  objects: []
});

// ── Layer 2: Helpers ──────────────────────────────────────────────────────

function addHeader(slide, title) {
  slide.addText(title, {
    x: PAD, y: 0.06, w: BODY_W - 2.5, h: TITLE_H,
    fontSize: 20, fontFace: "Georgia", bold: true, color: NAVY,
    valign: "bottom", margin: 0
  });
  slide.addText("Pack 42 Campout", {
    x: W - PAD - 1.8, y: 0.12, w: 1.8, h: 0.22,
    fontSize: 9, fontFace: "Calibri", color: SGRAY,
    align: "right", valign: "middle", margin: 0
  });
}

function addFooter(slide) {
  slideNum++;
  slide.addShape(pres.shapes.LINE, {
    x: 0, y: FOOTER_Y, w: W, h: 0,
    line: { color: NAVY, width: 1 }
  });
  slide.addText(String(slideNum), {
    x: W - PAD - 0.5, y: FOOTER_Y + 0.02, w: 0.5, h: 0.22,
    fontSize: 8, fontFace: "Calibri", color: SGRAY,
    align: "right", valign: "middle", margin: 0
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: FOOTER_Y, w: 0.08, h: H - FOOTER_Y,
    fill: { color: FLAME }, line: { color: FLAME }
  });
}

function addSectionLabel(slide, text, y, opts = {}) {
  const x = opts.x !== undefined ? opts.x : PAD;
  const w = opts.w || BODY_W;
  slide.addText(text.toUpperCase(), {
    x, y, w, h: 0.2,
    fontSize: 9, fontFace: "Calibri", bold: true, color: opts.color || NAVY,
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

// ── Config-driven template: activity slide ────────────────────────────────

function addActivitySlide(cfg) {
  const s = pres.addSlide({ masterName: "MASTER" });
  addHeader(s, cfg.title);
  addFooter(s);

  let y = BODY_TOP;

  // Time and location badges
  const badges = [
    { label: cfg.time, color: SKY },
    { label: cfg.location, color: PINE }
  ];
  if (cfg.leader) badges.push({ label: "Leader: " + cfg.leader, color: FLAME });

  let bx = PAD;
  badges.forEach(b => {
    const bw = b.label.length * 0.075 + 0.3;
    s.addShape(pres.shapes.RECTANGLE, {
      x: bx, y, w: bw, h: 0.26,
      fill: { color: b.color }
    });
    s.addText(b.label, {
      x: bx, y, w: bw, h: 0.26,
      fontSize: 9, fontFace: "Calibri", bold: true, color: WHITE,
      align: "center", valign: "middle", margin: 0
    });
    bx += bw + 0.1;
  });
  y += 0.4;

  const { leftX, rightX, colW } = twoColumnLayout();

  // Left: description + what to expect
  y = addSectionLabel(s, "What we're doing", y, { x: leftX, w: colW });
  s.addText(cfg.description, {
    x: leftX, y, w: colW, h: 0.6,
    fontSize: 10, fontFace: "Calibri", color: DGRAY, margin: 0,
    lineSpacingMultiple: 1.25, valign: "top"
  });
  y += 0.65;

  if (cfg.activities) {
    y = addSectionLabel(s, "Activities", y, { x: leftX, w: colW });
    addBullets(s, fitBullets(cfg.activities, 5, 100), leftX, y, colW, 2.0, 10);
  }

  // Right: what to bring + safety
  let ry = BODY_TOP + 0.4;
  if (cfg.bring) {
    ry = addSectionLabel(s, "What to bring", ry, { x: rightX, w: colW, color: FLAME, borderColor: GOLD });
    addBullets(s, fitBullets(cfg.bring, 5, 100), rightX, ry, colW, 1.5, 10);
    ry += 1.6;
  }

  if (cfg.safety) {
    ry = addSectionLabel(s, "Safety notes", ry, { x: rightX, w: colW, color: "CC3333", borderColor: "CC3333" });
    addBullets(s, cfg.safety, rightX, ry, colW, 1.2, 10);
  }
}

// ── Slide 1: Cover ────────────────────────────────────────────────────────
{
  slideNum++;
  const s = pres.addSlide();
  s.background = { color: NAVY };

  s.addText("Pack 42\nSpring Campout", {
    x: 0.8, y: 0.8, w: 8, h: 1.8,
    fontSize: 44, fontFace: "Georgia", bold: true, color: WHITE,
    lineSpacingMultiple: 1.1, margin: 0
  });

  s.addText("Lake Powhatan Campground, Bent Creek", {
    x: 0.8, y: 2.7, w: 7, h: 0.4,
    fontSize: 18, fontFace: "Calibri", color: GOLD,
    margin: 0
  });

  s.addText("April 18-20, 2026  |  Bears, Wolves & Tigers", {
    x: 0.8, y: 3.3, w: 7, h: 0.3,
    fontSize: 12, fontFace: "Calibri", color: SKY,
    margin: 0
  });

  // Campfire graphic (simple shapes)
  s.addShape(pres.shapes.OVAL, {
    x: 7.5, y: 3.8, w: 1.8, h: 0.4,
    fill: { color: "3A2A1A" }
  });
  s.addShape(pres.shapes.OVAL, {
    x: 7.9, y: 3.0, w: 1.0, h: 1.0,
    fill: { color: FLAME, transparency: 40 }
  });
  s.addShape(pres.shapes.OVAL, {
    x: 8.05, y: 3.15, w: 0.7, h: 0.7,
    fill: { color: GOLD, transparency: 30 }
  });
}

// ── Slide 2: Overview ─────────────────────────────────────────────────────
{
  const s = pres.addSlide({ masterName: "MASTER" });
  addHeader(s, "Weekend Overview");
  addFooter(s);

  let y = BODY_TOP;
  y = addSectionLabel(s, "Schedule at a glance", y);

  const rows = [
    [
      { text: "When", options: { bold: true, color: WHITE, fill: { color: NAVY } } },
      { text: "What", options: { bold: true, color: WHITE, fill: { color: NAVY } } },
      { text: "Who", options: { bold: true, color: WHITE, fill: { color: NAVY } } }
    ],
    ["Friday 5:00 PM", "Arrival & campsite setup", "All families"],
    ["Friday 7:00 PM", "Campfire & s'mores", "Everyone"],
    ["Saturday 8:00 AM", "Breakfast (provided by Pack)", "All scouts + siblings"],
    ["Saturday 9:30 AM", "Nature hike — waterfall loop", "All dens"],
    ["Saturday 12:00 PM", "Lunch & free swim", "All families"],
    ["Saturday 2:00 PM", "Outdoor skills stations", "Bears, Wolves, Tigers separately"],
    ["Saturday 5:30 PM", "Dutch oven cookoff", "Adults (scouts help)"],
    ["Saturday 7:30 PM", "Campfire program & skits", "Everyone"],
    ["Sunday 8:00 AM", "Breakfast & pack-down", "All families"],
    ["Sunday 10:00 AM", "Awards & depart", "All"],
  ];

  s.addTable(rows, {
    x: PAD, y, w: BODY_W,
    fontSize: 9, fontFace: "Calibri", color: BLACK,
    border: { pt: 0.5, color: MGRAY },
    colW: [1.8, 4.2, 3.0],
    autoPage: false
  });
}

// ── Slides 3-5: Activity details (config-driven) ─────────────────────────

addActivitySlide({
  title: "Nature Hike — Waterfall Loop",
  time: "Saturday 9:30 AM",
  location: "Bent Creek Trail",
  leader: "Mr. Rodriguez",
  description: "2.5-mile loop through the forest to a 30-foot waterfall. Mostly flat with one short climb. We'll stop for nature identification along the way — bring your adventure journals!",
  activities: [
    "Identify 5 native trees using field guides",
    "Stream crossing practice (shallow, supervised)",
    "Waterfall observation and sketch time",
    "Leave No Trace principles review",
    "Trail marker navigation exercise"
  ],
  bring: [
    "Sturdy closed-toe shoes (no sandals!)",
    "Water bottle — at least 16 oz",
    "Adventure journal and pencil",
    "Rain jacket (check forecast)",
    "Snack for the trail"
  ],
  safety: [
    "Buddy system at all times — no solo wandering",
    "Stay on marked trail",
    "Tell a leader before leaving the group for any reason"
  ]
});

addActivitySlide({
  title: "Outdoor Skills Stations",
  time: "Saturday 2:00 PM",
  location: "Campsite meadow",
  leader: "Den leaders",
  description: "Three rotating stations, 30 minutes each. Scouts rotate by den. Each station teaches a core outdoor skill that builds toward their next rank advancement.",
  activities: [
    "Station 1: Knot tying — bowline, clove hitch, square knot",
    "Station 2: Fire safety — match lighting, fire lay building (supervised demo)",
    "Station 3: Compass navigation — follow bearings to find hidden markers",
    "Bonus challenge: combine all three skills in a relay race"
  ],
  bring: [
    "Scout handbook (knot reference pages)",
    "Compass (loaners available)",
    "Water bottle",
    "Sunscreen"
  ],
  safety: [
    "Fire station is demonstration only — adults handle matches",
    "Knives stay closed unless a leader gives permission",
    "Hydration breaks every 15 minutes — it's hot out there"
  ]
});

addActivitySlide({
  title: "Dutch Oven Cookoff",
  time: "Saturday 5:30 PM",
  location: "Group fire ring",
  leader: "Mr. Chen & Ms. Patel",
  description: "Teams of 2-3 families cook a main dish or dessert in Dutch ovens. Scouts help with prep and stirring. Judging categories: taste, creativity, and teamwork. Everyone eats everything!",
  activities: [
    "Teams draw recipe themes from a hat (comfort food, international, dessert)",
    "30 minutes to plan and prep ingredients",
    "Cook time: 45 minutes over coals",
    "Scouts present their dish to the judges",
    "Awards: Golden Spatula, Most Creative, Best Teamwork"
  ],
  bring: [
    "Dutch oven (12-inch recommended) — or borrow one from the pack",
    "Charcoal briquettes (20 count per team)",
    "Ingredients for your recipe (coordinate with your team!)",
    "Heat-resistant gloves and lid lifter",
    "Serving utensils and plates"
  ],
  safety: [
    "Hot coals and Dutch ovens — scouts help with prep, adults handle the fire",
    "Keep a bucket of water near each cooking station",
    "Allergen check: ask about nut and dairy allergies before sharing food"
  ]
});

// ── Slide 6: Packing List ─────────────────────────────────────────────────
{
  const s = pres.addSlide({ masterName: "MASTER" });
  addHeader(s, "Packing Checklist");
  addFooter(s);

  let y = BODY_TOP;

  const colW = (BODY_W - 0.4) / 3;
  const gap = 0.2;

  const columns = [
    {
      title: "Sleeping",
      items: ["Tent (or share — coordinate!)", "Sleeping bag (rated to 40F)", "Sleeping pad or air mattress", "Pillow", "Extra blanket (nights get cold)"]
    },
    {
      title: "Clothing",
      items: ["Layers! Warm days, cool nights", "Rain jacket", "Sturdy shoes + camp shoes", "Hat and sunglasses", "Extra socks (always extra socks)"]
    },
    {
      title: "Essentials",
      items: ["Water bottle (labeled with name)", "Flashlight/headlamp + batteries", "Sunscreen and bug spray", "Camp chair", "Positive attitude (required)"]
    }
  ];

  columns.forEach((col, i) => {
    const cx = PAD + i * (colW + gap);
    y = BODY_TOP;
    y = addSectionLabel(s, col.title, y, { x: cx, w: colW, color: FLAME, borderColor: GOLD });
    addBullets(s, col.items, cx, y, colW, 3.5, 10);
  });
}

// ── Slide 7: Contact & Logistics ──────────────────────────────────────────
{
  const s = pres.addSlide({ masterName: "MASTER" });
  addHeader(s, "Logistics & Contacts");
  addFooter(s);

  let y = BODY_TOP;
  const { leftX, rightX, colW } = twoColumnLayout();

  y = addSectionLabel(s, "Getting there", y, { x: leftX, w: colW });
  addBullets(s, [
    "Lake Powhatan Campground, 375 Wesley Branch Rd, Asheville NC",
    "Group sites C3-C5 reserved under \"Pack 42\"",
    "Check-in opens at 4:00 PM Friday",
    "GPS coordinates: 35.4735 N, 82.6320 W",
    "Cell service is weak — download offline maps"
  ], leftX, y, colW, 2.0, 10);

  let ry = BODY_TOP;
  ry = addSectionLabel(s, "Emergency contacts", ry, { x: rightX, w: colW, color: "CC3333", borderColor: "CC3333" });

  const contacts = [
    ["Cubmaster", "Jake Mabry", "(828) 555-0142"],
    ["Asst. Cubmaster", "Maria Rodriguez", "(828) 555-0198"],
    ["First Aid", "Dr. Sarah Chen", "(828) 555-0167"],
    ["Campground Office", "", "(828) 670-5627"],
    ["Nearest ER", "Mission Hospital", "15 min drive"]
  ];

  const contactRows = [
    [
      { text: "Role", options: { bold: true, color: WHITE, fill: { color: NAVY } } },
      { text: "Name", options: { bold: true, color: WHITE, fill: { color: NAVY } } },
      { text: "Phone", options: { bold: true, color: WHITE, fill: { color: NAVY } } }
    ],
    ...contacts
  ];

  s.addTable(contactRows, {
    x: rightX, y: ry, w: colW,
    fontSize: 9, fontFace: "Calibri", color: BLACK,
    border: { pt: 0.5, color: MGRAY },
    colW: [1.2, 1.5, 1.6]
  });
}

// ── Slide 8: Closing ──────────────────────────────────────────────────────
{
  slideNum++;
  const s = pres.addSlide();
  s.background = { color: NAVY };

  s.addText("See You\nat the Campfire!", {
    x: 0.8, y: 1.0, w: 8, h: 1.8,
    fontSize: 44, fontFace: "Georgia", bold: true, color: WHITE,
    lineSpacingMultiple: 1.1, margin: 0
  });

  s.addText("Do Your Best", {
    x: 0.8, y: 3.0, w: 7, h: 0.5,
    fontSize: 20, fontFace: "Georgia", italic: true, color: GOLD,
    margin: 0
  });

  s.addText("RSVP by April 11  |  pack42campout@gmail.com", {
    x: 0.8, y: 4.2, w: 7, h: 0.3,
    fontSize: 12, fontFace: "Calibri", color: SKY,
    margin: 0
  });
}

// ── Write output ──────────────────────────────────────────────────────────
pres.writeFile({ fileName: "output/pack42-spring-campout.pptx" })
  .then(() => console.log("Done: output/pack42-spring-campout.pptx"))
  .catch(e => console.error(e));
