# pptxgenjs Guide

API reference and patterns for building presentation generators. Based on the [pptxgenjs](https://github.com/gitbrent/PptxGenJS) library (MIT licensed).

## Setup

```javascript
const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";  // 10" x 5.625"
pres.title = "My Deck";

const slide = pres.addSlide();
slide.addText("Hello", { x: 0.5, y: 0.5, fontSize: 24, color: "333333" });

pres.writeFile({ fileName: "output.pptx" });
```

### Layouts

| Layout | Width | Height |
|--------|-------|--------|
| `LAYOUT_16x9` | 10" | 5.625" |
| `LAYOUT_16x10` | 10" | 6.25" |
| `LAYOUT_4x3` | 10" | 7.5" |
| `LAYOUT_WIDE` | 13.3" | 7.5" |

All coordinates are in inches.

---

## Text

```javascript
slide.addText("Basic text", {
  x: 1, y: 1, w: 8, h: 1,
  fontSize: 18, fontFace: "Calibri", color: "333333",
  bold: true, align: "left", valign: "middle", margin: 0
});
```

### Rich text (mixed formatting)

```javascript
slide.addText([
  { text: "Bold part ", options: { bold: true } },
  { text: "normal part", options: { italic: true } }
], { x: 1, y: 2, w: 8, h: 1 });
```

### Multi-line text

Each segment needs `breakLine: true` to start a new line (except the last):

```javascript
slide.addText([
  { text: "Line 1", options: { breakLine: true } },
  { text: "Line 2", options: { breakLine: true } },
  { text: "Line 3" }
], { x: 0.5, y: 0.5, w: 8, h: 2 });
```

### Bullets

```javascript
slide.addText([
  { text: "First item", options: { bullet: true, breakLine: true } },
  { text: "Second item", options: { bullet: true, breakLine: true } },
  { text: "Sub-item", options: { bullet: true, indentLevel: 1, breakLine: true } },
  { text: "Third item", options: { bullet: true } }
], { x: 0.5, y: 0.5, w: 8, h: 3 });
```

**Bullet gotcha:** if you type `"• First item"` with a literal bullet character, pptxgenjs adds its own bullet too — you get a double bullet. Always use `bullet: true` in the options object.

### Character spacing

pptxgenjs has two spacing-related properties in its type definitions: `charSpacing` and `letterSpacing`. Only `charSpacing` actually works — `letterSpacing` compiles but produces no visible change:

```javascript
slide.addText("SPACED", { charSpacing: 6, ... });
```

### Text box padding

By default, `addText` adds internal padding inside the text box. This means the visible text position doesn't match the `x` coordinate you specified. To eliminate this offset — particularly when lining up text with shapes or icons — pass `margin: 0`.

---

## Shapes

```javascript
// Rectangle
slide.addShape(pres.shapes.RECTANGLE, {
  x: 1, y: 1, w: 3, h: 2,
  fill: { color: "0088CC" }, line: { color: "005588", width: 1 }
});

// Oval
slide.addShape(pres.shapes.OVAL, { x: 5, y: 1, w: 2, h: 2, fill: { color: "CC4444" } });

// Line
slide.addShape(pres.shapes.LINE, {
  x: 1, y: 4, w: 5, h: 0,
  line: { color: "CCCCCC", width: 0.75, dashType: "dash" }
});

// Rounded rectangle
slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 1, y: 1, w: 3, h: 2,
  fill: { color: "FFFFFF" }, rectRadius: 0.1
});

// With transparency
slide.addShape(pres.shapes.RECTANGLE, {
  x: 1, y: 1, w: 3, h: 2,
  fill: { color: "0088CC", transparency: 50 }
});
```

### Shadows

```javascript
slide.addShape(pres.shapes.RECTANGLE, {
  x: 1, y: 1, w: 3, h: 2,
  fill: { color: "FFFFFF" },
  shadow: { type: "outer", color: "000000", blur: 6, offset: 2, angle: 135, opacity: 0.15 }
});
```

Shadow properties:
- **`type`**: `"outer"` or `"inner"`
- **`color`**: 6-digit hex string without `#` (e.g., `"000000"`). Do not use 8-digit RGBA hex — it corrupts the output file. Control transparency with `opacity` instead.
- **`blur`**: blur radius in points (0-100)
- **`offset`**: distance in points (0-200). Negative values produce a corrupt file — for upward shadows, use `angle: 270` with a positive offset.
- **`angle`**: direction in degrees (0-359). Common values: 135 for lower-right, 315 for upper-left.
- **`opacity`**: 0.0 (invisible) to 1.0 (fully opaque)

**Gradients:** pptxgenjs only supports solid `fill.color` on shapes. For gradient effects, render the gradient as a PNG and use it as an image background on the shape or slide.

---

## Images

```javascript
// From file
slide.addImage({ path: "photo.png", x: 1, y: 1, w: 5, h: 3 });

// From base64
slide.addImage({ data: "image/png;base64,iVBOR...", x: 1, y: 1, w: 5, h: 3 });

// With options
slide.addImage({
  path: "photo.jpg",
  x: 1, y: 1, w: 5, h: 3,
  rounding: true,        // circular crop
  transparency: 30,      // 0-100
  sizing: { type: "cover", w: 5, h: 3 }  // or "contain", "crop"
});
```

### Preserve aspect ratio

```javascript
const origW = 1920, origH = 1080, maxH = 3.0;
const calcW = maxH * (origW / origH);
const centerX = (10 - calcW) / 2;
slide.addImage({ path: "wide.jpg", x: centerX, y: 1, w: calcW, h: maxH });
```

Supported formats: PNG, JPG, GIF, SVG (modern PowerPoint).

---

## Backgrounds

```javascript
slide.background = { color: "F5F5F5" };
slide.background = { color: "1A1A2E", transparency: 50 };
slide.background = { data: "image/png;base64,..." };
```

---

## Tables

```javascript
slide.addTable([
  ["Header 1", "Header 2", "Header 3"],
  ["Cell A", "Cell B", "Cell C"],
  ["Cell D", "Cell E", "Cell F"]
], {
  x: 0.5, y: 1, w: 9,
  border: { pt: 0.5, color: "CCCCCC" },
  colW: [3, 3, 3]
});
```

### Styled cells

```javascript
const rows = [
  [
    { text: "Header", options: { fill: { color: "333333" }, color: "FFFFFF", bold: true } },
    { text: "Value", options: { fill: { color: "333333" }, color: "FFFFFF", bold: true } }
  ],
  ["Data 1", "Data 2"],
  [{ text: "Merged", options: { colspan: 2 } }]
];
slide.addTable(rows, { x: 0.5, y: 1, w: 9 });
```

**Do not guess `rowH`** for variable content — pptxgenjs auto-sizes rows when `rowH` is omitted. Guessing causes overflow that renders past cell boundaries (pptxgenjs does not clip).

---

## Charts

```javascript
// Bar chart
slide.addChart(pres.charts.BAR, [{
  name: "Revenue", labels: ["Q1", "Q2", "Q3", "Q4"], values: [120, 145, 160, 180]
}], {
  x: 0.5, y: 1, w: 6, h: 3.5, barDir: "col",
  chartColors: ["2563EB", "3B82F6", "93C5FD"],
  valGridLine: { color: "D1D5DB", size: 0.5 },
  catGridLine: { style: "none" },
  showValue: true, dataLabelPosition: "outEnd",
  showLegend: false
});

// Line chart
slide.addChart(pres.charts.LINE, [{
  name: "Trend", labels: ["Jan", "Feb", "Mar"], values: [10, 25, 18]
}], { x: 0.5, y: 1, w: 6, h: 3, lineSmooth: true });

// Pie chart
slide.addChart(pres.charts.PIE, [{
  name: "Share", labels: ["A", "B", "C"], values: [45, 35, 20]
}], { x: 6, y: 1, w: 4, h: 3, showPercent: true });
```

Available chart types: `BAR`, `LINE`, `PIE`, `DOUGHNUT`, `SCATTER`, `BUBBLE`, `RADAR`.

---

## Slide Masters

```javascript
pres.defineSlideMaster({
  title: "MASTER",
  background: { color: "FFFFFF" },
  objects: []  // keep minimal — helpers give more control
});

const slide = pres.addSlide({ masterName: "MASTER" });
```

---

## Slide-Fit Rules

pptxgenjs does **not** clip overflow. Text renders past box boundaries onto whatever is below. You must handle this.

### MIN_FONT = 9

No body text below 9pt. When content doesn't fit at 9pt, **split the slide** into two rather than shrinking the font. One dense slide at 7pt is worse than two readable slides at 9pt.

### Automated fitting pipeline

1. **`fitBullets(items, maxItems, maxChars)`** — cap bullet count and length
2. **`trimText(text, maxChars)`** — shorten to a complete thought (never trailing ellipsis)
3. **`checkFit(label, text, fontSize, boxW, boxH)`** — warn at build time if text won't fit
4. **`estimateLines(text, fontSize, boxW)`** — wrapping-aware line count for height planning

```javascript
function estimateLines(text, fontSize, boxW) {
  const charsPerLine = Math.floor(boxW * 72 / (fontSize * 0.6));
  return Math.ceil(text.length / charsPerLine);
}
```

### Content prioritization

When source content exceeds capacity:
- **Status reports**: blockers > risks > next steps > accomplishments
- **Pitch decks**: value prop > differentiation > proof points > team
- General: which section drives a decision if the audience can only read one?

---

## Things That Will Bite You

These are bugs we hit while building real decks. Most produce corrupt files or invisible layout breakage.

### Color strings must be exactly 6 hex digits

pptxgenjs passes color values straight into the XML. A `#` prefix or an 8-digit RGBA string (like `"FF000080"`) produces invalid markup. The file may open in some viewers but fail in others.

```javascript
color: "E05030"     // correct
color: "#E05030"    // breaks
color: "E0503080"   // breaks (use opacity property instead)
```

### Bullet items need `breakLine: true` between them

Without it, all items render on a single line. The last item in the array doesn't need it.

### `lineSpacing` and bullets don't mix well

Setting `lineSpacing` on a bulleted text block adds spacing both between and within items, making the list look double-spaced. For vertical breathing room between bullets, use `paraSpaceAfter` on each item's options instead.

### One `pptxgen()` instance per deck

If you create two presentations in the same script, each needs its own `new pptxgen()`. Reusing an instance carries over slide masters and internal state from the previous deck.

### Option objects get mutated after use

When you pass an options object (shadow, fill, etc.) to `addShape` or `addText`, the library modifies that object's properties during rendering. If you pass the same object to a second call, it receives the already-transformed values and produces wrong output.

Fix: use a factory function that returns a fresh object each time.

```javascript
const cardShadow = () => ({ type: "outer", blur: 4, offset: 2, color: "333333", opacity: 0.2 });

slide.addShape(pres.shapes.RECTANGLE, { ..., shadow: cardShadow() });
slide.addShape(pres.shapes.RECTANGLE, { ..., shadow: cardShadow() });
```

### Accent bars on rounded rectangles leave visible corners

If you layer a thin rectangular accent strip over a `ROUNDED_RECTANGLE`, the strip's square corners poke out. Either use `RECTANGLE` for the card shape, or skip the accent overlay pattern entirely.

### Tables with fixed `rowH` overflow silently

pptxgenjs does not clip table cell content. If you set `rowH: 0.4` and the text needs 0.6", it renders past the cell boundary onto whatever is below. For variable-length content, omit `rowH` and let the engine auto-size rows.

### Font shrinking causes deck-wide inconsistency

When one slide has too much content, the temptation is to drop the font size. But slide 7 at 8pt next to slide 6 at 11pt looks broken. Enforce a MIN_FONT floor across the entire deck and split overloaded slides into two instead of shrinking.
