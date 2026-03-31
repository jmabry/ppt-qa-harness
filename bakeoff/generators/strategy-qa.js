"use strict";
const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Q3 2026 Strategic Review — NovaCrest";

// ─── PALETTE ─────────────────────────────────────────────────────
const C = {
  navy:    "0B1D35",
  blue:    "152C50",
  accent:  "1D6FD8",
  teal:    "0D9488",
  amber:   "D97706",
  red:     "C8192B",
  white:   "FFFFFF",
  offwhite:"F3F6FA",
  ltgray:  "DDE4EF",
  midgray: "8496AF",
  darkgray:"2D3E54",
  text:    "1A2B3E",
  gold:    "F59E0B",
};
const W = 10, H = 5.625;
const FOOTER_H = 0.26;
const HDR_H    = 0.50;

// ─── HELPERS ─────────────────────────────────────────────────────

/** Standard slide chrome: header bar + footer bar */
function chrome(slide, title, sub) {
  // Header
  slide.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:W, h:HDR_H, fill:{color:C.navy}, line:{color:C.navy} });
  slide.addShape(pres.shapes.RECTANGLE, { x:0, y:HDR_H, w:W, h:0.025, fill:{color:C.accent}, line:{color:C.accent} });
  slide.addText(title, { x:0.32, y:0, w:7.2, h:HDR_H, fontSize:15, bold:true, color:C.white, valign:"middle", margin:0 });
  if (sub) slide.addText(sub, { x:0, y:0, w:W-0.28, h:HDR_H, fontSize:9, color:"7A9ABF", valign:"middle", align:"right", margin:0 });
  // Footer
  slide.addShape(pres.shapes.RECTANGLE, { x:0, y:H-FOOTER_H, w:W, h:FOOTER_H, fill:{color:C.navy}, line:{color:C.navy} });
  slide.addText("NovaCrest  ·  Q3 2026 Board Review  ·  CONFIDENTIAL", {
    x:0.28, y:H-FOOTER_H, w:6, h:FOOTER_H, fontSize:7.5, color:"7A9ABF", valign:"middle", margin:0
  });
  slide.addText("March 31, 2026", {
    x:0, y:H-FOOTER_H, w:W-0.28, h:FOOTER_H, fontSize:7.5, color:"7A9ABF", align:"right", valign:"middle", margin:0
  });
}

/** KPI card with left accent bar */
function kpi(slide, x, y, w, h, label, value, sub, vc, bg) {
  bg = bg || C.white;
  vc = vc || C.accent;
  slide.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill:{color:bg}, line:{color:C.ltgray, width:0.7} });
  slide.addShape(pres.shapes.RECTANGLE, { x, y, w:0.06, h, fill:{color:vc}, line:{color:vc} });
  const vFontSize = value.length > 7 ? 16 : 20;
  slide.addText(value, { x:x+0.12, y:y+0.05, w:w-0.16, h:h*0.50, fontSize:vFontSize, bold:true, color:vc, valign:"middle", margin:0 });
  slide.addText(label, { x:x+0.12, y:y+h*0.52, w:w-0.16, h:h*0.26, fontSize:8, bold:true, color:C.darkgray, margin:0 });
  if (sub) slide.addText(sub, { x:x+0.12, y:y+h*0.76, w:w-0.16, h:h*0.22, fontSize:7, color:C.midgray, margin:0 });
}

/** Colored section label bar */
function slab(slide, x, y, w, text, color) {
  color = color || C.accent;
  slide.addShape(pres.shapes.RECTANGLE, { x, y, w, h:0.21, fill:{color}, line:{color} });
  slide.addText(text, { x:x+0.08, y, w:w-0.1, h:0.21, fontSize:7.5, bold:true, color:C.white, valign:"middle", margin:0 });
}

/** Lightweight table renderer */
function tbl(slide, rows, opts) {
  const { x, y, colW, hdrFill, fontSize } = opts;
  const rh = opts.rh || 0.265;
  const hf = hdrFill || C.navy;
  const altFill = opts.altFill || C.offwhite;

  rows.forEach((row, ri) => {
    let cx = x;
    row.forEach((cell, ci) => {
      const cw = colW[ci];
      const isHdr = ri === 0;
      const bg = isHdr ? hf : (ri % 2 === 0 ? C.white : altFill);
      slide.addShape(pres.shapes.RECTANGLE, { x:cx, y:y+ri*rh, w:cw, h:rh, fill:{color:bg}, line:{color:C.ltgray, width:0.5} });
      const raw  = typeof cell === "object" ? cell : { text: String(cell) };
      const txt  = raw.text !== undefined ? String(raw.text) : String(cell);
      const align= raw.align || (ci === 0 ? "left" : "center");
      const fc   = isHdr ? C.white : (raw.color || C.text);
      const fs   = raw.fontSize || fontSize || 7.6;
      slide.addText(txt, { x:cx+0.06, y:y+ri*rh, w:cw-0.08, h:rh, fontSize:fs, color:fc, bold:isHdr||raw.bold, align, valign:"middle", margin:0 });
      cx += cw;
    });
  });
}

/** Mini chart config factory */
function chartStyle(colors) {
  return {
    chartColors: colors || [C.accent],
    chartArea: { fill:{color:C.white}, border:{color:C.ltgray, width:0.5} },
    catAxisLabelColor: C.midgray, valAxisLabelColor: C.midgray,
    catAxisLabelFontSize: 9, valAxisLabelFontSize: 9,
    valGridLine: { color:C.ltgray, size:0.5 }, catGridLine: { style:"none" },
  };
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 1 — TITLE + EXECUTIVE SUMMARY
// ═══════════════════════════════════════════════════════════════
{
  const slide = pres.addSlide();

  // Full dark background
  slide.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:W, h:H, fill:{color:C.navy}, line:{color:C.navy} });
  // Left accent stripe
  slide.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:0.10, h:H, fill:{color:C.accent}, line:{color:C.accent} });
  // Right panel
  slide.addShape(pres.shapes.RECTANGLE, { x:5.6, y:0, w:4.4, h:H, fill:{color:C.blue}, line:{color:C.blue} });

  // ── Left: Title block ──
  slide.addText("Q3 2026", { x:0.28, y:0.38, w:5.1, h:0.52, fontSize:36, bold:true, color:C.accent, charSpacing:2, margin:0 });
  slide.addText("STRATEGIC REVIEW", { x:0.28, y:0.88, w:5.1, h:0.38, fontSize:22, bold:true, color:C.white, charSpacing:3, margin:0 });
  slide.addText("NovaCrest  ·  Board of Directors  ·  March 31, 2026  ·  Confidential", {
    x:0.28, y:1.3, w:5.1, h:0.24, fontSize:9, color:"5C7A9B", margin:0
  });

  // Separator
  slide.addShape(pres.shapes.RECTANGLE, { x:0.28, y:1.62, w:5.0, h:0.025, fill:{color:C.accent}, line:{color:C.accent} });

  // Company snapshot cards
  const snaps = [
    ["$15.6M", "Q3 ARR"], ["Series B", "$32M · Mar 2025"], ["112", "Headcount"], ["2021", "San Francisco"],
  ];
  snaps.forEach(([val, lbl], i) => {
    const sx = 0.28 + i * 1.27;
    slide.addShape(pres.shapes.RECTANGLE, { x:sx, y:1.70, w:1.22, h:0.68, fill:{color:"0A1828"}, line:{color:C.accent, width:0.7} });
    slide.addText(val, { x:sx+0.08, y:1.73, w:1.08, h:0.30, fontSize:14, bold:true, color:C.accent, margin:0 });
    slide.addText(lbl, { x:sx+0.08, y:2.02, w:1.08, h:0.28, fontSize:7.5, color:C.midgray, margin:0 });
  });

  // Team line
  slide.addText("CEO Sarah Chen  ·  CTO Marcus Williams  ·  CFO David Park  ·  VP Sales Rachel Torres", {
    x:0.28, y:2.48, w:5.1, h:0.20, fontSize:7.5, color:"4E6A88", margin:0
  });
  slide.addText("Board: Jeff Blackwell (Insight Partners)  ·  Priya Sharma (Accel)  ·  Tom Nguyen (Independent)", {
    x:0.28, y:2.67, w:5.1, h:0.20, fontSize:7.5, color:"4E6A88", margin:0
  });

  // ── Right panel sections ──
  const rpX = 5.76, rpW = 3.98;

  // WINS
  slide.addShape(pres.shapes.RECTANGLE, { x:rpX, y:0.22, w:rpW, h:0.22, fill:{color:C.teal}, line:{color:C.teal} });
  slide.addText("✓  KEY WINS — Q3 2026", { x:rpX+0.08, y:0.22, w:rpW-0.1, h:0.22, fontSize:8, bold:true, color:C.white, valign:"middle", margin:0 });
  const wins = [
    "ARR $15.6M  (+10.6% QoQ  ·  +4.2% vs plan)",
    "Pipeline $8.6M — best quarter ever; win rate 28.4%",
    "VP CS + Head of Partnerships + 2 ML PhD engineers hired",
  ];
  wins.forEach((w, i) => {
    slide.addText("→  " + w, { x:rpX+0.1, y:0.46+i*0.30, w:rpW-0.28, h:0.27, fontSize:8, color:"CCEAD4", margin:0 });
  });

  // CONCERNS
  slide.addShape(pres.shapes.RECTANGLE, { x:rpX, y:1.41, w:rpW, h:0.22, fill:{color:C.amber}, line:{color:C.amber} });
  slide.addText("⚠  CONCERNS", { x:rpX+0.08, y:1.41, w:rpW-0.1, h:0.22, fontSize:8, bold:true, color:C.white, valign:"middle", margin:0 });
  const concerns = [
    "Churn: $500K in Q3 (+66.7% QoQ) — 58% preventable",
    "Competitive: DataForge SMB module; Zenith A-round + hiring",
  ];
  concerns.forEach((c, i) => {
    slide.addText("→  " + c, { x:rpX+0.1, y:1.65+i*0.30, w:rpW-0.28, h:0.27, fontSize:8, color:"FFE0A0", margin:0 });
  });

  // BOARD DECISIONS
  slide.addShape(pres.shapes.RECTANGLE, { x:rpX, y:2.30, w:rpW, h:0.22, fill:{color:C.red}, line:{color:C.red} });
  slide.addText("●  BOARD DECISIONS REQUIRED", { x:rpX+0.08, y:2.30, w:rpW-0.1, h:0.22, fontSize:8, bold:true, color:C.white, valign:"middle", margin:0 });
  const decisions = [
    "APPROVE: $2.5M CS investment — 14-mo payback; save $700K ARR/yr",
    "APPROVE: Usage-based SMB tier ($499/mo entry) — +$1.2M ARR Y1",
    "DISCUSS: Series C timing — Q2 2027 from strength vs $30M ARR milestone",
  ];
  decisions.forEach((d, i) => {
    slide.addText("→  " + d, { x:rpX+0.1, y:2.54+i*0.30, w:rpW-0.28, h:0.27, fontSize:8, color:"FCA5A5", margin:0 });
  });

  // AGENDA
  slide.addShape(pres.shapes.RECTANGLE, { x:rpX, y:3.50, w:rpW, h:0.22, fill:{color:"0D1E3A"}, line:{color:C.accent, width:0.7} });
  slide.addText("AGENDA", { x:rpX+0.08, y:3.50, w:rpW-0.1, h:0.22, fontSize:7.5, bold:true, color:C.accent, valign:"middle", margin:0 });
  const agenda = ["2  Revenue Dashboard", "3  GTM + Customer Health", "4  Product + Team", "5  Financial Outlook + Competitive", "6  Board Asks"];
  agenda.forEach((a, i) => {
    slide.addText(a, { x:rpX+0.1, y:3.75+i*0.24, w:rpW-0.12, h:0.22, fontSize:8, color:C.midgray, margin:0 });
  });

  // Footer
  slide.addShape(pres.shapes.RECTANGLE, { x:0, y:H-FOOTER_H, w:W, h:FOOTER_H, fill:{color:"060D18"}, line:{color:"060D18"} });
  slide.addText("NovaCrest  ·  Q3 2026 Board Review  ·  CONFIDENTIAL", {
    x:0.28, y:H-FOOTER_H, w:9.4, h:FOOTER_H, fontSize:7.5, color:"7A9ABF", valign:"middle", margin:0
  });
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 2 — REVENUE DASHBOARD
// ═══════════════════════════════════════════════════════════════
{
  const slide = pres.addSlide();
  slide.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:W, h:H, fill:{color:C.offwhite}, line:{color:C.offwhite} });
  chrome(slide, "SLIDE 2  ·  Revenue Dashboard", "Q3 2026 Actuals vs Plan");

  const TOP = HDR_H + 0.025 + 0.10; // below chrome line

  // ── 5 KPI Cards ──
  const cards = [
    { label:"ARR",                value:"$15.6M",   sub:"+10.6% QoQ  ·  +4.2% vs plan",     vc:C.accent },
    { label:"Net Revenue Ret.",   value:"118%",     sub:"Target 120%+  ·  ↓ from 121% (churn)",vc:C.amber  },
    { label:"Gross Margin",       value:"78.2%",    sub:"Target 80%+  ·  Improving ↑",        vc:C.teal   },
    { label:"Magic Number",       value:"1.12x",    sub:"Q2: 0.98  ·  Target >0.75  ·  Strong",vc:C.teal  },
    { label:"CAC Payback",        value:"13.2 mo",  sub:"Target <16 mo  ·  Improving ↑",      vc:C.teal   },
  ];
  const cw = (W - 0.28) / cards.length;
  cards.forEach((c, i) => kpi(slide, 0.14+i*cw, TOP, cw-0.07, 0.80, c.label, c.value, c.sub, c.vc));

  const ROW2Y = TOP + 0.88;

  // ── ARR Trend Chart ──  (h=2.0 to leave room for bridge table below)
  slab(slide, 0.14, ROW2Y, 3.6, "ARR TREND  ($M)  ·  Q1 2025 – Q3 2026");
  const arrData = [{ name:"ARR ($M)", labels:["Q1'25","Q2'25","Q3'25","Q4'25","Q1'26","Q2'26","Q3'26"], values:[7.2,8.5,9.8,11.4,12.8,14.1,15.6] }];
  slide.addChart(pres.charts.LINE, arrData, Object.assign(chartStyle([C.accent]), {
    x:0.14, y:ROW2Y+0.23, w:3.6, h:2.0,
    lineSize:2.5, lineSmooth:false,
    valAxisMaxVal:20,   // headroom so the 15.6 top label isn't clipped
    showValue:true, dataLabelFontSize:8, dataLabelColor:C.darkgray,
    showLegend:false,
  }));

  // ── Revenue Breakdown Chart ──
  slab(slide, 3.85, ROW2Y, 2.82, "QUARTERLY REVENUE  ($M)");
  const revData = [
    { name:"Subscription", labels:["Q1'26","Q2'26","Q3'26"], values:[3.1,3.4,3.8] },
    { name:"Services",     labels:["Q1'26","Q2'26","Q3'26"], values:[0.3,0.3,0.3] },
  ];
  slide.addChart(pres.charts.BAR, revData, Object.assign(chartStyle([C.accent,"93C5FD"]), {
    x:3.85, y:ROW2Y+0.23, w:2.82, h:2.0,
    barDir:"col", barGrouping:"stacked",
    valAxisMaxVal:5.0,  // headroom for top label
    showValue:true, dataLabelFontSize:8, dataLabelColor:C.white,
    showLegend:true, legendPos:"b", legendFontSize:8,
  }));

  // ── Unit Economics Table  (rh=0.24 keeps it within chart height) ──
  slab(slide, 6.80, ROW2Y, 3.06, "UNIT ECONOMICS");
  tbl(slide, [
    ["Metric",         "Q3 2026", "Q2 2026", "Target",   "Trend"],
    ["Gross Margin",   "78.2%",   "77.5%",   "80%+",     {text:"↑ Improving", color:C.teal}],
    ["NRR",            "118%",    "121%",    "120%+",    {text:"↓ Churn risk", color:C.amber}],
    ["Logo Retention", "92.4%",   "94.1%",   "95%+",     {text:"↓ Behind",    color:C.red}],
    ["LTV:CAC",        "4.8x",    "4.6x",    "4.0x+",    {text:"↑ Healthy",   color:C.teal}],
    ["CAC Payback",    "13.2 mo", "14.1 mo", "<16 mo",   {text:"↑ Improving", color:C.teal}],
    ["Magic Number",   "1.12",    "0.98",    ">0.75",    {text:"↑ Strong",    color:C.teal}],
  // colW must sum to 3.06 (= slab width); table right edge = 6.80+3.06 = 9.86 (0.14" from slide edge)
  ], { x:6.80, y:ROW2Y+0.23, colW:[0.88,0.46,0.46,0.46,0.80], hdrFill:C.navy, rh:0.24 });
  // Unit econ ends at ROW2Y + 0.23 + 8*0.24 = ROW2Y + 2.15

  // ── ARR Bridge table  (positioned after charts end at ROW2Y+2.23) ──
  // BROW = ROW2Y+2.36 ensures 0.13" clear gap below both charts and unit-econ table.
  const BROW = ROW2Y + 2.36;
  slab(slide, 0.14, BROW, 6.59, "ARR BRIDGE  —  Q3 2026");
  // Bridge table colW sums to 6.59 (= slab width) so all columns have room.
  // 6 rows at rh=0.205 → table ends at BROW+0.23+6*0.205 = BROW+1.46.
  // Total y: TOP+0.88+2.36+0.23+1.23 = TOP+4.70 = 5.325 → 0.04" above footer. ✓
  tbl(slide, [
    ["Component",        "Q1 2026", "Q2 2026", "Q3 2026",                             "QoQ Δ",                    "vs Plan"],
    ["Beginning ARR",    "$11.4M",  "$12.8M",  "$14.1M",                              "—",                        "—"],
    [{text:"+ New ARR",     color:C.teal}, "$1.4M", "$1.6M", {text:"$1.9M",  color:C.teal}, "+18.8%",             {text:"+12% ↑", color:C.teal}],
    [{text:"+ Expansion",   color:C.teal}, "$0.6M", "$0.7M", {text:"$0.8M",  color:C.teal}, "+14.3%",             "On plan"],
    [{text:"− Churn ARR",   color:C.red},  "($0.3M)","($0.3M)",{text:"($0.5M)",color:C.red},{text:"+66.7% ⚠",color:C.red},{text:"Behind ↓",color:C.red}],
    [{text:"= Ending ARR",  bold:true},    "$12.8M","$14.1M", {text:"$15.6M",bold:true,color:C.accent},"+10.6%", {text:"+4.2% ↑",color:C.teal}],
  ], { x:0.14, y:BROW+0.23, colW:[1.69,0.98,0.98,0.98,0.98,0.98], hdrFill:C.navy, rh:0.205 });
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 3 — GTM + CUSTOMER HEALTH
// ═══════════════════════════════════════════════════════════════
{
  const slide = pres.addSlide();
  slide.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:W, h:H, fill:{color:C.offwhite}, line:{color:C.offwhite} });
  chrome(slide, "SLIDE 3  ·  GTM Performance & Customer Health", "Pipeline · Channels · Churn · At-Risk");

  const TOP = HDR_H + 0.025 + 0.10;

  // ── 5 GTM KPI cards ──
  const gtmCards = [
    { label:"Pipeline Generated", value:"$8.6M",    sub:"+21.1% QoQ  ·  Best quarter ever", vc:C.accent },
    { label:"Win Rate",           value:"28.4%",    sub:"+2.3pp QoQ  ·  Enterprise improving",vc:C.teal  },
    { label:"Avg Deal Size (ACV)", value:"$42K",    sub:"+10.5%  ·  Moving upmarket ✓",       vc:C.teal  },
    { label:"Quota Attainment",   value:"112%",     sub:"+18pp vs Q2 94%  ·  Rebounded",      vc:C.teal  },
    { label:"Sales Cycle",        value:"68 days",  sub:"-5.6% QoQ  ·  New demo flow",        vc:C.teal  },
  ];
  const cw = (W - 0.28) / gtmCards.length;
  gtmCards.forEach((c, i) => kpi(slide, 0.14+i*cw, TOP, cw-0.07, 0.75, c.label, c.value, c.sub, c.vc));

  const ROW2Y = TOP + 0.84;

  // ── Channel Performance table (left) ──
  slab(slide, 0.14, ROW2Y, 4.68, "CHANNEL PERFORMANCE  —  Q3 2026");
  tbl(slide, [
    ["Channel",           "Pipeline", "Won", "Avg ACV", "CAC",    "Notes"],
    ["Outbound SDR",      "$3.8M",    "18",  "$52K",    "$24.1K", "Scaled 4→6 SDRs in Q2"],
    ["Inbound Marketing", "$2.4M",    "14",  "$34K",    "$12.8K", "Content + paid search"],
    ["PLG / Self-Serve",  "$0.9M",    "8",   "$18K",    "$6.2K",  "4.8% trial→paid conv."],
    [{text:"Partner/Referral",color:C.teal},"$1.5M","5",{text:"$68K",color:C.teal},"$8.4K",{text:"SAP + Siemens ramping ↑",color:C.teal}],
    [{text:"TOTAL",bold:true},{text:"$8.6M",bold:true},{text:"45",bold:true},{text:"$42K",bold:true},"$15.8K",""],
  // colW sums to 4.68 — Notes column widened for legibility
  ], { x:0.14, y:ROW2Y+0.23, colW:[1.00,0.58,0.34,0.64,0.64,1.48], hdrFill:C.navy });

  // ── CAC Trend chart (right) ──
  slab(slide, 4.90, ROW2Y, 4.96, "BLENDED CAC TREND  ($K)  ·  Q1 2025 – Q3 2026");
  const cacData = [
    { name:"Blended",    labels:["Q1'25","Q2'25","Q3'25","Q4'25","Q1'26","Q2'26","Q3'26"], values:[22.4,20.1,19.8,18.2,17.6,16.4,15.8] },
    { name:"Enterprise", labels:["Q1'25","Q2'25","Q3'25","Q4'25","Q1'26","Q2'26","Q3'26"], values:[38.2,34.6,32.1,28.4,26.8,24.2,22.6] },
    { name:"Mid-Market", labels:["Q1'25","Q2'25","Q3'25","Q4'25","Q1'26","Q2'26","Q3'26"], values:[14.8,13.2,13.6,12.4,12.1,11.8,11.2] },
  ];
  slide.addChart(pres.charts.LINE, cacData, Object.assign(chartStyle([C.accent,C.teal,C.amber]), {
    x:4.90, y:ROW2Y+0.23, w:4.96, h:1.78,
    lineSize:2, showValue:false,
    showLegend:true, legendPos:"r", legendFontSize:7.5,
  }));

  const CHURN_Y = ROW2Y + 2.10;

  // ── Churn Deep Dive (left ~60%) ──
  slab(slide, 0.14, CHURN_Y, 5.88, "CHURN DEEP DIVE  —  $500K IN Q3  (+66% QoQ)  ·  58% PREVENTABLE", C.red);
  tbl(slide, [
    ["Account",            "Seg",        "Lost",   "Root Cause",                                   "Prevent?"],
    ["Apex Manufacturing", "Enterprise", "$145K",  "Acquired; contract lapsed post-M&A",           {text:"No",              color:C.midgray}],
    ["Precision Dynamics", "Mid-Market", "$62K",   "In-house analytics; felt 'overkill'",          {text:"Partial—down-tier",color:C.amber}],
    ["4 SMB accounts",     "SMB",        "$118K",  "Price; 3 → DataForge, 1 went OOB",             {text:"Yes — 2-3 saveable",color:C.red}],
    ["TechFab Solutions",  "Mid-Market", "$48K",   "Poor impl.; never hit time-to-value",          {text:"Yes — CS failure", color:C.red}],
    ["3 SMB accounts",     "SMB",        "$82K",   "Low usage (<10%); never onboarded",            {text:"Yes — onboarding", color:C.red}],
    ["Consolidated Parts", "Mid-Market", "$45K",   "Budget cut; eliminated analytics",             {text:"No",              color:C.midgray}],
    [{text:"TOTAL / PREVENTABLE",bold:true},"",{text:"$500K",bold:true,color:C.red},{text:"$290K of $500K (58%) preventable — root: onboarding + no down-tier",bold:true},{text:"$290K at risk",bold:true,color:C.red}],
  // rh=0.193 → 8 rows=1.544; ends at CHURN_Y+0.23+1.544=5.339 (0.026" above footer). ✓
  ], { x:0.14, y:CHURN_Y+0.23, colW:[1.30,0.72,0.56,2.40,1.05], hdrFill:"7F1D1D", rh:0.193 });

  // ── At-Risk Watch List (right 1/3) ──
  slab(slide, 6.14, CHURN_Y, 3.72, "AT-RISK  —  Q4 WATCH LIST", C.amber);
  tbl(slide, [
    ["Account",          "ARR",    "Signal",                                  "Owner"],
    [{text:"Sterling Inds.",color:C.red},{text:"$210K",color:C.red},"Champion left; new VP evaluating alternatives","R.Torres"],
    ["ClearPath Systems","$72K",   "Renewal 60 days; competitor POC running",  "D.Park"],
    ["Midwest Components","$58K",  "Usage -40% Aug; support tickets 3x",       "J.Lin (CS)"],
  // colW sums to 3.72 (= RX slab width 6.14→9.86)
  ], { x:6.14, y:CHURN_Y+0.23, colW:[0.94,0.50,1.76,0.52], hdrFill:C.amber, rh:0.31 });
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 4 — PRODUCT + TEAM
// ═══════════════════════════════════════════════════════════════
{
  const slide = pres.addSlide();
  slide.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:W, h:H, fill:{color:C.offwhite}, line:{color:C.offwhite} });
  chrome(slide, "SLIDE 4  ·  Product & Team", "Q3 Shipped  ·  Q4 Roadmap  ·  Headcount  ·  Key Hires");

  const TOP = HDR_H + 0.025 + 0.08;

  // ── LEFT COLUMN: Product ──
  const LW = 5.55;

  // Q3 Shipped — rh=0.28, 5 rows
  slab(slide, 0.14, TOP, LW, "Q3 2026 SHIPPED FEATURES");
  tbl(slide, [
    ["Feature",                        "30-Day Adoption",                    "Revenue Impact"],
    [{text:"Predictive Maint. v2",bold:true}, "67% of enterprise accounts",  {text:"$1.2M pipeline; 3 deals cite this",color:C.teal}],
    ["Self-Serve Dashboard Builder",   "340 dashboards · 89 accounts",       "CS ticket volume -18%"],
    [{text:"SAP Integration (native)",color:C.accent}, "12 connected; 8 in progress", {text:"Partner channel accelerant ↑",color:C.accent}],
    [{text:"SOC 2 Type II ✓ (8/15)",color:C.teal}, "Compliance milestone",   {text:"Unblocked 4 deals ($380K pipeline)",color:C.teal}],
  ], { x:0.14, y:TOP+0.23, colW:[1.74,1.97,1.79], hdrFill:C.navy, rh:0.28 });
  // Shipped ends at TOP + 0.23 + 5*0.28 = TOP + 1.63

  const MID_Y = TOP + 1.72;  // 0.09" gap

  // Q4 Roadmap — rh=0.255, header + 4 rows (drop P2 to save space)
  slab(slide, 0.14, MID_Y, LW, "Q4 ROADMAP PRIORITIES");
  tbl(slide, [
    ["Pri", "Feature",                                   "Status",                  "Expected Impact"],
    [{text:"P0",bold:true,color:C.red},   "Multi-tenant analytics (enterprise isolation)", "In dev — 60%",    {text:"Required for 2 deals ($200K+ each)",color:C.red}],
    [{text:"P0",bold:true,color:C.red},   "Siemens MindSphere integration",                "Design done; dev starting",{text:"Opens $4M TAM in industrial IoT",color:C.teal}],
    [{text:"P1",bold:true,color:C.amber}, "Usage-based SMB pricing tier ($499/mo)",        "Spec complete",   {text:"Addresses 58% of preventable churn",color:C.amber}],
    [{text:"P1",bold:true,color:C.amber}, "Customer health score (internal tool)",         "Prototype built", "Early warning for at-risk accounts"],
  ], { x:0.14, y:MID_Y+0.23, colW:[0.34,2.06,1.36,1.74], hdrFill:C.navy, rh:0.255 });
  // Roadmap ends at MID_Y + 0.23 + 5*0.255 = MID_Y + 1.505

  const HIRE_Y = MID_Y + 1.61;  // 0.10" gap

  // Key hires — rh=0.27, header + 3 rows
  slab(slide, 0.14, HIRE_Y, LW, "KEY HIRES — Q3 2026     ⚠ 6 departures (3 eng → FAANG); comp bands +12%");
  tbl(slide, [
    ["Role",                         "Background",                                "Focus"],
    [{text:"VP Customer Success (9/1)",bold:true},  "Datadog — CS org 20→80",    {text:"Own churn reduction; build SMB onboarding",color:C.teal}],
    [{text:"Head of Partnerships (8/15)",bold:true},"Siemens — SAP/Siemens channel",{text:"Scale partner pipeline → $5M target",color:C.accent}],
    ["2 Senior ML Engineers",                       "PhD, Georgia Tech",          "Predictive Maintenance v3 accuracy roadmap"],
  ], { x:0.14, y:HIRE_Y+0.23, colW:[1.74,1.88,1.88], hdrFill:C.navy, rh:0.268 });
  // Hires ends at HIRE_Y + 0.23 + 4*0.268 = HIRE_Y + 1.302

  // ── RIGHT COLUMN: Team ──
  const RX = 5.82, RW = 4.04;

  slab(slide, RX, TOP, RW, "HEADCOUNT BY FUNCTION");
  tbl(slide, [
    ["Function",           "Q2", "Q3", "Open",                               "Q4 Target"],
    ["Engineering",        "42", "46", {text:"4",color:C.amber},             {text:"50",color:C.teal}],
    ["Product & Design",   "8",  "9",  {text:"1",color:C.amber},             {text:"10",color:C.teal}],
    ["Sales",              "18", "22", {text:"3",color:C.amber},             {text:"25",color:C.teal}],
    ["Customer Success",   "12", "14", {text:"2",color:C.amber},             {text:"16",color:C.teal}],
    ["Marketing",          "8",  "9",  {text:"1",color:C.amber},             {text:"10",color:C.teal}],
    ["G&A",                "10", "12", {text:"1",color:C.amber},             {text:"13",color:C.teal}],
    [{text:"TOTAL",bold:true},{text:"98",bold:true},{text:"112",bold:true},{text:"12 open",bold:true,color:C.amber},{text:"124",bold:true,color:C.teal}],
  ], { x:RX, y:TOP+0.23, colW:[1.40,0.50,0.50,0.58,1.01], hdrFill:C.navy, rh:0.278 });

  // Headcount bar chart — height capped to avoid footer overlap
  // Headcount table ends at TOP+0.23+8*0.278=TOP+2.454; chart starts at TOP+2.52
  // Footer at 5.365; chart bottom must be ≤ 5.365; y = 0.605+2.52 = 3.125; h ≤ 2.24
  slide.addChart(pres.charts.BAR, [
    { name:"Q3 2026",  labels:["Eng","Prod","Sales","CS","Mktg","G&A"], values:[46,9,22,14,9,12] },
    { name:"Q4 Target",labels:["Eng","Prod","Sales","CS","Mktg","G&A"], values:[50,10,25,16,10,13] },
  ], Object.assign(chartStyle([C.accent,"93C5FD"]), {
    x:RX, y:TOP+2.52, w:RW, h:2.18,
    barDir:"col", barGrouping:"clustered",
    valAxisMaxVal:65,
    showValue:true, dataLabelFontSize:8,
    dataLabelPosition:"outEnd",
    showLegend:true, legendPos:"b", legendFontSize:8,
  }));
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 5 — FINANCIAL OUTLOOK + COMPETITIVE
// ═══════════════════════════════════════════════════════════════
{
  const slide = pres.addSlide();
  slide.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:W, h:H, fill:{color:C.offwhite}, line:{color:C.offwhite} });
  chrome(slide, "SLIDE 5  ·  Financial Outlook & Competitive Landscape", "P&L  ·  Cash Position  ·  Market Dynamics  ·  FY2026E");

  const TOP = HDR_H + 0.025 + 0.08;

  // ── LEFT: P&L ──
  slab(slide, 0.14, TOP, 5.82, "P&L SUMMARY  ($K)  ·  FY2026E");
  tbl(slide, [
    ["Line Item",          "Q1 2026", "Q2 2026", "Q3 2026",                               "Q3 YoY",                 "FY2026E"],
    ["Revenue",            "$3,400",  "$3,700",  "$4,100",                                "+64%",                   "$15,400"],
    ["COGS",               "($740)",  "($830)",  "($890)",                                "+52%",                   "($3,340)"],
    [{text:"Gross Profit",bold:true},  {text:"$2,660",bold:true},{text:"$2,870",bold:true},{text:"$3,210",bold:true,color:C.teal},{text:"+68%",color:C.teal},{text:"$12,060",bold:true}],
    ["Gross Margin",       "78.2%",   "77.6%",   "78.3%",                                "+2.1pp",                 "78.3%"],
    ["S&M",                "($1,420)","($1,580)","($1,720)",                              "+48%",                   "($6,480)"],
    ["R&D",                "($1,340)","($1,480)","($1,620)",                              "+55%",                   "($6,040)"],
    ["G&A",                "($480)",  "($510)",  "($540)",                                "+38%",                   "($2,060)"],
    [{text:"Total OpEx",bold:true},   "($3,240)","($3,570)","($3,880)",                  "+50%",                   "($14,580)"],
    [{text:"Net Income",bold:true},   {text:"($580)",color:C.red},{text:"($700)",color:C.red},{text:"($670)",color:C.amber},{text:"Improving ↑",color:C.teal},{text:"($2,520)",color:C.amber}],
    ["Monthly Burn",       "$193K",   "$233K",   "$223K",                                "—",                      "—"],
  ], { x:0.14, y:TOP+0.23, colW:[1.56,0.80,0.80,0.80,0.68,0.84], hdrFill:C.navy, rh:0.258 });

  // ── Cash KPI cards ──
  const CASH_Y = TOP + 3.28;
  slab(slide, 0.14, CASH_Y, 5.82, "CASH POSITION & EFFICIENCY");
  const cashCards = [
    { label:"Cash on Hand",   value:"$18.2M",   sub:"Series B $32M · Mar 2025",         vc:C.teal  },
    { label:"Monthly Burn",   value:"$223K",    sub:"Q3 avg — declining QoQ",            vc:C.accent},
    { label:"Runway",         value:"81 mo",    sub:"Comfortable — not near-term factor",vc:C.teal  },
    { label:"Burn Multiple",  value:"0.36x",    sub:"Excellent (<1x = efficient)",        vc:C.teal  },
  ];
  const ccw = 5.82 / cashCards.length;
  cashCards.forEach((c, i) => kpi(slide, 0.14+i*ccw, CASH_Y+0.23, ccw-0.07, 0.80, c.label, c.value, c.sub, c.vc));

  // ── RIGHT: Burn chart + Competitive table ──
  const RX = 6.08, RW = 3.78;

  slab(slide, RX, TOP, RW, "MONTHLY BURN  ($K)  ·  Q1–Q3 2026");
  slide.addChart(pres.charts.BAR, [
    { name:"Monthly Burn ($K)", labels:["Q1'26","Q2'26","Q3'26"], values:[193,233,223] },
  ], Object.assign(chartStyle(["C8192B"]), {
    x:RX, y:TOP+0.23, w:RW, h:1.45,
    barDir:"col", showValue:true, dataLabelFontSize:9,
    dataLabelColor:C.white, dataLabelPosition:"inEnd", valAxisMaxVal:260,
    showLegend:false,
  }));

  slab(slide, RX, TOP+1.80, RW, "COMPETITIVE LANDSCAPE");
  tbl(slide, [
    ["Dimension",   "NovaCrest",      "DataForge",      "Acme Analytics", "Zenith AI"],
    ["Stage",       "Series B $32M",  "Series C $85M",  "Public $2.1B",   "Series A $18M"],
    ["Est. ARR",    "$15.6M",         "~$45M",          "~$180M",         "~$4M"],
    ["Market",      "Mid-mkt mfg",    "SMB-Mid horiz.", "Enterprise all", "Mid-mkt mfg"],
    ["Avg ACV",     "$30-250K",       "$2.4-48K",       "$100K-500K+",    "$20-80K"],
    ["Win vs. us",  "—",              {text:"62% ✓",color:C.teal},{text:"34% ✗",color:C.red},{text:"55% ✓",color:C.teal}],
    ["Key Threat",  "—",              {text:"$199/mo mfg module",color:C.amber},{text:"Acme Lite Q1'27",color:C.amber},{text:"VC hiring blitz",color:C.red}],
  ], { x:RX, y:TOP+2.03, colW:[0.76,0.76,0.76,0.76,0.74], hdrFill:C.navy, rh:0.255 });
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 6 — BOARD ASKS
// ═══════════════════════════════════════════════════════════════
{
  const slide = pres.addSlide();
  slide.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:W, h:H, fill:{color:C.navy}, line:{color:C.navy} });
  // Top accent band
  slide.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:W, h:0.52, fill:{color:"060D18"}, line:{color:"060D18"} });
  slide.addShape(pres.shapes.RECTANGLE, { x:0, y:0.52, w:W, h:0.025, fill:{color:C.accent}, line:{color:C.accent} });
  slide.addText("BOARD ASKS  ·  Q3 2026 REVIEW  ·  THREE ITEMS REQUIRING ACTION", {
    x:0.28, y:0, w:9.44, h:0.52, fontSize:14, bold:true, color:C.white, charSpacing:1.5, valign:"middle", margin:0
  });

  // Card layout: 3 cards with 0.3" gutters
  // Total usable width: 10 - 0.14 - 0.14 - 2*0.30 = 9.12 → each card = 3.04
  const asks = [
    {
      num:"01", actionLabel:"APPROVE", actionColor:C.red,
      title:"$2.5M Incremental Customer Success Investment",
      x:0.14, w:3.04,
      context:[
        "58% of Q3 churn ($290K of $500K) preventable: onboarding gap, no down-tier, one CS failure",
        "Q4 at-risk: Sterling ($210K) + ClearPath ($72K) + Midwest ($58K) = $340K exposure right now",
        "Hire 8 FTEs: 4 onboarding specialists, 2 renewal mgrs, 1 SMB CS lead, 1 CS ops/analytics",
      ],
      metrics:[["ARR Saved/yr","~$700K"],["Payback","14 months"],["3-yr ROI","~280%"],["If delayed","Q4 churn ≥$600K"]],
      urgency:"URGENT: $340K at-risk Q4 pipeline — delay compounds churn.",
    },
    {
      num:"02", actionLabel:"APPROVE", actionColor:C.amber,
      title:"Usage-Based SMB Pricing ($499/mo vs $2.5K/mo min)",
      x:3.48, w:3.04,
      context:[
        "60+ SMB prospects/quarter lost on price — DataForge at $199/mo mfg. module",
        "Down-tier option: Precision Dynamics ($62K) churn was partially preventable",
        "Spec complete; 6-week dev est.; Q1 2027 GA launch target",
      ],
      metrics:[["New ARR Y1","+$1.2M"],["SMB churn cut","40-50%"],["Down-tier risk","($200K) ARR"],["Net impact","+$1.0M ARR"]],
      urgency:"DataForge accelerating — delay risks permanent SMB floor loss.",
    },
    {
      num:"03", actionLabel:"DISCUSS", actionColor:C.accent,
      title:"Series C Timing — Q2 2027 vs $30M ARR Milestone",
      x:6.82, w:3.04,
      context:[
        "81-mo runway = no urgency; raising from strength maximizes valuation",
        "Q2 2027: ~$22-25M ARR, improving NRR, burn 0.3x — strong story",
        "$30M: ~Q4 2027; 2 extra quarters of metrics; higher valuation ceiling",
      ],
      metrics:[["Cash on hand","$18.2M"],["Runway","81 months"],["Burn multiple","0.36x"],["Target raise","Board input req."]],
      urgency:"Input needed: timing, raise size ($20-40M?), investor targets.",
    },
  ];

  asks.forEach(ask => {
    const { x, w, num, actionLabel, actionColor, title, context, metrics, urgency } = ask;
    const BY = 0.60;

    // Card background
    slide.addShape(pres.shapes.RECTANGLE, { x, y:BY, w, h:H-BY-FOOTER_H, fill:{color:"0D1E38"}, line:{color:C.accent, width:0.7} });
    // Top color bar
    slide.addShape(pres.shapes.RECTANGLE, { x, y:BY, w, h:0.28, fill:{color:actionColor}, line:{color:actionColor} });
    slide.addText(`${num}  ${actionLabel}`, { x:x+0.1, y:BY, w:w-0.12, h:0.28, fontSize:10, bold:true, color:C.white, valign:"middle", margin:0 });

    // Title
    slide.addText(title, { x:x+0.1, y:BY+0.31, w:w-0.15, h:0.40, fontSize:9, bold:true, color:C.white, margin:0 });

    // Context bullets
    context.forEach((b, i) => {
      slide.addText("→  " + b, { x:x+0.1, y:BY+0.74+i*0.31, w:w-0.16, h:0.29, fontSize:8, color:"BCD4EE", margin:0 });
    });

    // Mini metrics table
    const MY = BY + 1.70;
    slide.addShape(pres.shapes.RECTANGLE, { x:x+0.08, y:MY, w:w-0.16, h:0.22, fill:{color:actionColor}, line:{color:actionColor} });
    slide.addText("KEY METRICS", { x:x+0.12, y:MY, w:w-0.2, h:0.22, fontSize:8, bold:true, color:C.white, valign:"middle", margin:0 });
    metrics.forEach(([label, val], i) => {
      const isBg = i % 2 === 0 ? "112236" : "0D1C30";
      slide.addShape(pres.shapes.RECTANGLE, { x:x+0.08, y:MY+0.22+i*0.26, w:w-0.16, h:0.26, fill:{color:isBg}, line:{color:C.navy, width:0.5} });
      slide.addText(label, { x:x+0.12, y:MY+0.22+i*0.26, w:(w-0.22)*0.60, h:0.26, fontSize:8.5, color:"8AB4D8", valign:"middle", margin:0 });
      slide.addText(val,   { x:x+0.12+(w-0.22)*0.60, y:MY+0.22+i*0.26, w:(w-0.22)*0.40, h:0.26, fontSize:8.5, bold:true, color:C.white, valign:"middle", align:"right", margin:0 });
    });

    // Urgency strip — explicit height to fit one clear line
    const UY = MY + 0.22 + metrics.length*0.26 + 0.08;
    slide.addShape(pres.shapes.RECTANGLE, { x:x+0.08, y:UY, w:w-0.16, h:0.34, fill:{color:"1A0A0A"}, line:{color:actionColor, width:0.8} });
    slide.addText(urgency, { x:x+0.12, y:UY, w:w-0.24, h:0.34, fontSize:9, color:"FCA5A5", valign:"middle", margin:0 });
  });

  // Footer
  slide.addShape(pres.shapes.RECTANGLE, { x:0, y:H-FOOTER_H, w:W, h:FOOTER_H, fill:{color:"060D18"}, line:{color:"060D18"} });
  slide.addText("NovaCrest  ·  Q3 2026 Board Review  ·  CONFIDENTIAL", {
    x:0.28, y:H-FOOTER_H, w:9.44, h:FOOTER_H, fontSize:7.5, color:"7A9ABF", valign:"middle", margin:0
  });
}

// ─── WRITE FILE ───────────────────────────────────────────────
pres.writeFile({ fileName: "outputs/strategy-qa.pptx" })
  .then(() => console.log("✅  outputs/strategy-qa.pptx written"))
  .catch(e => { console.error(e); process.exit(1); });
