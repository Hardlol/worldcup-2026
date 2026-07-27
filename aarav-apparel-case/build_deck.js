const pptxgen = require("pptxgenjs");
const T = require("./theme");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";          // 13.333 x 7.5 in
pres.author = "Aarav Apparel — Commercial Strategy";
pres.title = "Aarav Apparel: Ending the Channel War";

const W = 13.333, H = 7.5;
const M = 0.62;                        // left/right margin
const CW = W - 2 * M;                  // content width

const HEAD = "Cambria";
const BODY = "Calibri";

const C = {
  ink: T.INK, indigo: T.INDIGO, indigoL: T.INDIGO_L, saffron: T.SAFFRON,
  teal: T.TEAL, grey: T.GREY, band: T.BAND, band2: T.BAND_2, rule: T.RULE,
  white: "FFFFFF", crimson: T.CRIMSON,
};

/* ---------------- helpers ---------------- */

// Light content slide with eyebrow + title
function contentSlide(eyebrow, title) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  s.addText(eyebrow.toUpperCase(), {
    x: M, y: 0.34, w: CW, h: 0.26, margin: 0,
    fontFace: BODY, fontSize: 11.5, bold: true, color: C.saffron, charSpacing: 2,
  });
  s.addText(title, {
    x: M, y: 0.62, w: CW, h: 0.62, margin: 0,
    fontFace: HEAD, fontSize: 31, bold: true, color: C.indigo, valign: "top",
  });
  return s;
}

// Answer slide: numbered chip in the eyebrow row
function answerSlide(n, eyebrow, title) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  s.addShape(pres.ShapeType.ellipse, {
    x: M, y: 0.3, w: 0.44, h: 0.44, fill: { color: C.indigo },
  });
  s.addText(String(n), {
    x: M, y: 0.3, w: 0.44, h: 0.44, margin: 0,
    fontFace: HEAD, fontSize: 18, bold: true, color: C.white,
    align: "center", valign: "middle",
  });
  s.addText(eyebrow.toUpperCase(), {
    x: M + 0.62, y: 0.33, w: CW - 0.62, h: 0.2, margin: 0,
    fontFace: BODY, fontSize: 11, bold: true, color: C.saffron, charSpacing: 2,
  });
  s.addText(title, {
    x: M + 0.62, y: 0.54, w: CW - 0.62, h: 0.46, margin: 0,
    fontFace: HEAD, fontSize: 27, bold: true, color: C.indigo, valign: "top",
  });
  return s;
}

function card(s, x, y, w, h, fill, opts = {}) {
  s.addShape(pres.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.06,
    fill: { color: fill },
    line: opts.line ? { color: opts.line, width: 1 } : { color: fill, width: 0 },
    shadow: opts.shadow ? { type: "outer", angle: 90, blur: 8, offset: 1.5, color: "9A9A9A", opacity: 0.18 } : undefined,
  });
}

function stat(s, x, y, w, value, label, color) {
  s.addText(value, {
    x, y, w, h: 0.72, margin: 0,
    fontFace: HEAD, fontSize: 40, bold: true, color, align: "left", valign: "middle",
  });
  s.addText(label, {
    x, y: y + 0.7, w, h: 0.5, margin: 0,
    fontFace: BODY, fontSize: 11.5, color: C.grey, align: "left", valign: "top",
  });
}

function footNote(s, text) {
  s.addText(text, {
    x: M, y: H - 0.52, w: CW, h: 0.28, margin: 0,
    fontFace: BODY, fontSize: 9.5, italic: true, color: C.grey,
  });
}

/* =========================================================
   1. TITLE
   ========================================================= */
{
  const s = pres.addSlide();
  s.background = { color: C.indigo };
  s.addShape(pres.ShapeType.ellipse, {
    x: 9.5, y: -1.9, w: 6.2, h: 6.2, fill: { color: "27476E" }, line: { color: "27476E", width: 0 },
  });
  s.addShape(pres.ShapeType.ellipse, {
    x: 10.9, y: 3.55, w: 3.9, h: 3.9, fill: { color: C.indigo }, line: { color: T.SAFFRON, width: 1.75 },
  });

  s.addText("SALES & DISTRIBUTION MANAGEMENT  ·  CASE ANALYSIS", {
    x: M, y: 1.62, w: 9.6, h: 0.3, margin: 0,
    fontFace: BODY, fontSize: 12, bold: true, color: T.SAFFRON, charSpacing: 3,
  });
  s.addText("Ending the Channel War", {
    x: M, y: 2.0, w: 9.9, h: 1.15, margin: 0,
    fontFace: HEAD, fontSize: 52, bold: true, color: C.white,
  });
  s.addText("An omni-channel sales organisation, attribution logic, budget split\nand alignment mechanism for Aarav Apparel, FY27", {
    x: M, y: 3.22, w: 9.4, h: 0.95, margin: 0,
    fontFace: HEAD, fontSize: 18, italic: true, color: "CBD6E4", lineSpacing: 26,
  });
  s.addShape(pres.ShapeType.line, {
    x: M, y: 4.4, w: 3.1, h: 0, line: { color: C.saffron, width: 2.5 },
  });
  s.addText([
    { text: "Prepared for the Chief Financial Officer", options: { bold: true, color: C.white, breakLine: true } },
    { text: "120 stores  ·  D2C site & app  ·  Myntra, Ajio, Flipkart, Amazon  ·  July 2026", options: { color: "9FB2C9" } },
  ], {
    x: M, y: 4.66, w: 9.4, h: 0.8, margin: 0, fontFace: BODY, fontSize: 13, lineSpacing: 20,
  });
  s.addNotes("Opening line: 'You asked me two questions — how much should digital get, and how do I stop the field revolt. They are the same question, and the answer starts with how we count.' Set up: 12 minutes, four answers, one assumption set. Flag early that every number is assumed and stated.");
}

/* =========================================================
   2. SITUATION
   ========================================================= */
{
  const s = contentSlide("The situation", "A healthy P&L sitting on a broken measurement system");

  // vignette card
  card(s, M, 1.42, 5.55, 2.05, C.band, { shadow: true });
  s.addText("Indore, 6:40 pm", {
    x: M + 0.28, y: 1.6, w: 5.0, h: 0.28, margin: 0,
    fontFace: BODY, fontSize: 11, bold: true, color: C.saffron, charSpacing: 1,
  });
  s.addText("An associate fits a customer for twenty minutes. The customer photographs the tag and buys the same style that night on Myntra at 26% off.\n\nThe store books nothing. The associate earns nothing. Digital books 100%.", {
    x: M + 0.28, y: 1.9, w: 5.0, h: 1.45, margin: 0,
    fontFace: BODY, fontSize: 13, color: C.ink, lineSpacing: 18,
  });

  // stats
  const sx = 6.6, sw = 3.1;
  stat(s, sx, 1.45, sw, "₹1,150 cr", "FY26 net revenue", C.indigo);
  stat(s, sx + 3.2, 1.45, sw, "77 / 23", "Store vs digital revenue mix", C.indigo);
  stat(s, sx, 2.62, sw, "14 pts", "Self-inflicted online-vs-store discount gap", C.crimson);
  stat(s, sx + 3.2, 2.62, sw, "39%", "Store bills with a customer ID — six in ten journeys are invisible", C.crimson);

  // three forces
  s.addText("Three forces make this urgent, not merely unfair", {
    x: M, y: 4.08, w: CW, h: 0.3, margin: 0,
    fontFace: BODY, fontSize: 14, bold: true, color: C.indigoL,
  });
  const forces = [
    ["Quick commerce reached fashion", "Myntra's M-Now runs 30-minute delivery from 87+ dark stores across ten markets. The store's last operational edge — immediacy — is gone."],
    ["Marketplace margin is tightening", "Effective deductions of 30–40% of selling price. Marketplace growth buys revenue at two-thirds of a store rupee's margin."],
    ["D2C brands are buying what we own", "Snitch targeting 100 stores, The Souled Store 200. We already have 120 — and we account for them as a cost centre."],
  ];
  forces.forEach(([t, d], i) => {
    const x = M + i * (CW / 3 + 0.02);
    const w = CW / 3 - 0.28;
    card(s, x, 4.4, w, 1.8, C.band2, { line: C.rule });
    s.addShape(pres.ShapeType.ellipse, { x: x + 0.24, y: 4.6, w: 0.3, h: 0.3, fill: { color: C.teal } });
    s.addText(String(i + 1), { x: x + 0.24, y: 4.6, w: 0.3, h: 0.3, margin: 0, fontFace: BODY, fontSize: 12, bold: true, color: C.white, align: "center", valign: "middle" });
    s.addText(t, { x: x + 0.64, y: 4.58, w: w - 0.88, h: 0.34, margin: 0, fontFace: BODY, fontSize: 12.5, bold: true, color: C.indigo });
    s.addText(d, { x: x + 0.24, y: 4.98, w: w - 0.48, h: 1.14, margin: 0, fontFace: BODY, fontSize: 11, color: C.ink, lineSpacing: 15 });
  });
  footNote(s, "Sources: Indian Retailer (2025); TechCrunch (2024); Terra Insight (2025); Inc42 (2025); Exchange4media (2025).");
  s.addNotes("Land the vignette hard — it is the whole case in four sentences. Then the numbers: the discount gap is OURS, not Myntra's. We set it. And 39% capture means we are arguing about attribution while blind to six in ten journeys. The three forces answer 'why now'. Do not linger past 90 seconds.");
}

/* =========================================================
   3. ASSUMPTIONS
   ========================================================= */
{
  const s = contentSlide("On the record", "Every number in this deck is an assumption — here they are");

  const rows = [
    ["FY26 net revenue", "₹1,150 cr", "Stores ₹885 cr · Marketplace ₹184 cr · D2C ₹81 cr"],
    ["Network", "120 stores", "2,300 sq ft average · ₹7.38 cr revenue per door"],
    ["Gross margin", "52 / 55 / 37%", "Store · D2C · Marketplace (net of commission and returns)"],
    ["Store contribution", "17.5%", "After rent 9.5% and payroll 6.8% of store net sales"],
    ["Return rate", "3 / 17 / 29%", "Store · D2C · Marketplace — blended 14.2%"],
    ["Discount to MRP", "12 / 18 / 26%", "Store · D2C · Marketplace"],
    ["People", "1,080 + 62 + 45", "Store-facing · Digital · Omni ops, CRM and planning"],
    ["Digital near a store", "68%", "Share of digital revenue landing within 8 km of a door"],
    ["FY27 plan", "₹1,392 cr (+21%)", "18 net new stores · digital to ₹391 cr"],
    ["S&D budget in scope", "₹139 cr", "10.0% of planned revenue · FY26 actual ₹118 cr"],
  ];

  let y = 1.44;
  const rh = 0.545;
  rows.forEach(([a, b, c], i) => {
    if (i % 2 === 0) card(s, M, y, CW, rh, C.band2);
    s.addText(a, { x: M + 0.22, y, w: 2.85, h: rh, margin: 0, fontFace: BODY, fontSize: 12.5, bold: true, color: C.ink, valign: "middle" });
    s.addText(b, { x: M + 3.1, y, w: 2.35, h: rh, margin: 0, fontFace: HEAD, fontSize: 14, bold: true, color: C.indigo, valign: "middle" });
    s.addText(c, { x: M + 5.6, y, w: CW - 5.85, h: rh, margin: 0, fontFace: BODY, fontSize: 11.5, color: C.grey, valign: "middle" });
    y += rh;
  });

  footNote(s, "1 crore = 10 million. Figures net of GST. Full assumption set and sensitivities in Section 2 and 8.1 of the report.");
  s.addNotes("Twenty seconds only. The point is not the numbers, it is that they are declared. Say: 'None of this was given to me. If you disagree with a number, the model recalculates — Section 8.1 gives the three that actually move the answer.' Then move on quickly.");
}

/* =========================================================
   4. ANSWER 1 — ORGANISATION
   ========================================================= */
{
  const s = answerSlide(1, "Organisation", "One commercial P&L above the channels");

  // CEO / CCO spine
  card(s, 5.35, 1.28, 2.6, 0.5, C.grey);
  s.addText("CEO", { x: 5.35, y: 1.28, w: 2.6, h: 0.5, margin: 0, fontFace: BODY, fontSize: 13, bold: true, color: C.white, align: "center", valign: "middle" });

  card(s, 4.35, 1.95, 4.6, 0.62, C.indigo, { shadow: true });
  s.addText("Chief Commercial Officer  (new)", { x: 4.35, y: 1.98, w: 4.6, h: 0.32, margin: 0, fontFace: HEAD, fontSize: 15, bold: true, color: C.white, align: "center" });
  s.addText("The only role carrying a revenue number", { x: 4.35, y: 2.26, w: 4.6, h: 0.26, margin: 0, fontFace: BODY, fontSize: 10.5, color: "BFD0E2", align: "center" });

  s.addShape(pres.ShapeType.line, { x: 6.65, y: 1.78, w: 0, h: 0.17, line: { color: C.rule, width: 1.5 } });
  s.addShape(pres.ShapeType.line, { x: 6.65, y: 2.57, w: 0, h: 0.28, line: { color: C.rule, width: 1.5 } });
  s.addShape(pres.ShapeType.line, { x: 1.85, y: 2.85, w: 9.6, h: 0, line: { color: C.rule, width: 1.5 } });

  const dirs = [
    ["Retail Sales", "4 zonal heads\n16 area managers\n138 stores", C.teal, "CHANNEL"],
    ["Digital Commerce", "D2C · Marketplace\nQuick commerce\n16 catchment analysts", C.teal, "CHANNEL"],
    ["Omni Operations", "Availability-to-promise\nShip-from-store, C&C\nReturns network", C.saffron, "SHARED"],
    ["Customer & CRM", "Single customer ID\nLoyalty, attribution\nCatchment intelligence", C.saffron, "SHARED"],
  ];
  const cw = 2.42, gap = 0.24;
  const startX = (W - (4 * cw + 3 * gap)) / 2;
  dirs.forEach(([t, d, col, tag], i) => {
    const x = startX + i * (cw + gap);
    s.addShape(pres.ShapeType.line, { x: x + cw / 2, y: 2.85, w: 0, h: 0.25, line: { color: C.rule, width: 1.5 } });
    card(s, x, 3.1, cw, 1.86, C.white, { line: col, shadow: true });
    s.addText(tag, { x: x + 0.16, y: 3.2, w: cw - 0.32, h: 0.22, margin: 0, fontFace: BODY, fontSize: 8.5, bold: true, color: col, charSpacing: 1.5 });
    s.addText(t, { x: x + 0.16, y: 3.42, w: cw - 0.32, h: 0.34, margin: 0, fontFace: BODY, fontSize: 13, bold: true, color: C.indigo });
    s.addText(d, { x: x + 0.16, y: 3.8, w: cw - 0.32, h: 1.0, margin: 0, fontFace: BODY, fontSize: 10.5, color: C.grey, lineSpacing: 14 });
  });

  // three moves
  const moves = [
    ["Channels lose their revenue targets", "They keep contribution reporting for cost control. Only the CCO carries revenue — so neither director can win a transfer at the other's expense."],
    ["Shared services are peers, not support", "Report them into either channel and the data and the inventory get rationed in that channel's favour."],
    ["16 Catchment Councils bind both channels", "Chaired by the Area Manager: one forecast, one stock plan, one promotion calendar per cluster. Four outcomes carry two accountable owners, deliberately."],
  ];
  card(s, M, 5.14, CW, 1.66, C.band);
  moves.forEach(([t, d], i) => {
    const x = M + 0.26 + i * (CW / 3);
    const w = CW / 3 - 0.4;
    s.addText(t, { x, y: 5.3, w, h: 0.5, margin: 0, fontFace: BODY, fontSize: 12, bold: true, color: C.indigo, lineSpacing: 15 });
    s.addText(d, { x, y: 5.82, w, h: 0.9, margin: 0, fontFace: BODY, fontSize: 10.5, color: C.ink, lineSpacing: 14 });
  });
  s.addNotes("The design rule, say it verbatim: 'No unit below the CCO owns a customer. Units own channels; the CCO owns customers.' Two P&Ls chasing one customer in one pincode WILL compete — that is not a culture problem, it is arithmetic. Flag the Digital Catchment Analyst as the hinge role: it gives digital a named human being who wins or loses alongside the stores.");
}

/* =========================================================
   5. ANSWER 2 — THREE LEDGERS
   ========================================================= */
{
  const s = answerSlide(2, "Attribution", "Stop using one number for three jobs");

  const ledgers = [
    ["Revenue credit", "Where did the cash land?", "100% to the channel that took payment.\nSums to exactly 100%.", "Statutory accounts, GST, channel P&L", C.grey],
    ["Influence credit", "Who touched this journey?", "100% to every channel with a qualifying touch.\nSums to more than 100% by design.", "Budget allocation, catchment planning", C.teal],
    ["Incentive credit", "Who gets paid for this?", "Fractional split by journey type.\nSums to exactly 100%.", "Variable pay, bonus pool, league tables", C.saffron],
  ];
  const cw = 3.94, gap = 0.3;
  ledgers.forEach(([t, q, r, u, col], i) => {
    const x = M + i * (cw + gap);
    card(s, x, 1.34, cw, 2.6, C.white, { line: col, shadow: true });
    s.addText(t, { x: x + 0.24, y: 1.54, w: cw - 0.48, h: 0.34, margin: 0, fontFace: HEAD, fontSize: 17, bold: true, color: col });
    s.addText(q, { x: x + 0.24, y: 1.9, w: cw - 0.48, h: 0.3, margin: 0, fontFace: BODY, fontSize: 12, italic: true, color: C.ink });
    s.addText(r, { x: x + 0.24, y: 2.24, w: cw - 0.48, h: 0.86, margin: 0, fontFace: BODY, fontSize: 11.5, color: C.grey, lineSpacing: 15 });
    s.addShape(pres.ShapeType.line, { x: x + 0.24, y: 3.28, w: cw - 0.48, h: 0, line: { color: C.rule, width: 1 } });
    s.addText(u, { x: x + 0.24, y: 3.38, w: cw - 0.48, h: 0.48, margin: 0, fontFace: BODY, fontSize: 11, bold: true, color: C.indigo, lineSpacing: 14 });
  });

  s.addText("Incentive-credit split rules — matched on verified mobile or loyalty ID", {
    x: M, y: 4.18, w: CW, h: 0.3, margin: 0, fontFace: BODY, fontSize: 14, bold: true, color: C.indigoL,
  });

  const rules = [
    ["Pure walk-in", "100 / 0", "₹604 cr"],
    ["Webrooming → store (30-day)", "70 / 30", "₹281 cr"],
    ["Showrooming → online (14-day)", "45 / 55", "₹71 cr"],
    ["Endless aisle, associate-raised", "100 / 0", "₹9 cr"],
    ["Click & collect / ship-from-store", "40 / 60", "₹13 cr"],
    ["Marketplace / D2C, no store link", "0 / 100", "₹172 cr"],
  ];
  const rw = (CW - 0.4) / 3, rh = 0.56;
  rules.forEach(([n, split, val], i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const x = M + col * (rw + 0.2), y = 4.56 + row * (rh + 0.16);
    card(s, x, y, rw, rh, C.band2, { line: C.rule });
    s.addText(n, { x: x + 0.16, y, w: rw - 1.75, h: rh, margin: 0, fontFace: BODY, fontSize: 11, color: C.ink, valign: "middle" });
    s.addText(split, { x: x + rw - 1.62, y, w: 0.82, h: rh, margin: 0, fontFace: HEAD, fontSize: 13, bold: true, color: C.saffron, align: "right", valign: "middle" });
    s.addText(val, { x: x + rw - 0.76, y, w: 0.6, h: rh, margin: 0, fontFace: BODY, fontSize: 10.5, color: C.grey, align: "right", valign: "middle" });
  });

  s.addText("Splits track who bore the decisive cost and risk. In fashion that is fit and trial — which is why webrooming favours the store.", {
    x: M, y: 6.02, w: CW, h: 0.3, margin: 0, fontFace: BODY, fontSize: 12, italic: true, color: C.indigoL,
  });
  footNote(s, "Store % / Digital %. Only OTP-verified matches earn incentive credit; unmatched volume is sized by geo-holdout and used for budgeting only.");
  s.addNotes("The whole argument is here: we have been using booked revenue for three incompatible purposes. Separating them costs the accounts nothing — digital keeps every rupee it books. What changes is that booked revenue stops doing two jobs it was never fit for. If challenged on the split percentages: they are a judgement, they are published, and they are frozen for twelve months.");
}

/* =========================================================
   6. THE RESTATEMENT
   ========================================================= */
{
  const s = answerSlide(2, "Attribution — the test", "The model costs stores ₹38 cr. That is why it is credible.");

  s.addChart(pres.ChartType.bar, [
    { name: "Stores", labels: ["As booked today", "On incentive credit"], values: [885, 847] },
    { name: "Digital", labels: ["As booked today", "On incentive credit"], values: [265, 303] },
  ], {
    x: M, y: 1.4, w: 7.5, h: 3.55,
    barDir: "col", barGrouping: "stacked",
    chartColors: [C.indigo, C.saffron],
    showValue: true, dataLabelPosition: "ctr",
    dataLabelColor: "FFFFFF", dataLabelFontFace: BODY, dataLabelFontSize: 13, dataLabelFontBold: true,
    showLegend: true, legendPos: "b", legendFontFace: BODY, legendFontSize: 12, legendColor: C.ink,
    catAxisLabelFontFace: BODY, catAxisLabelFontSize: 13, catAxisLabelColor: C.ink,
    valAxisLabelFontFace: BODY, valAxisLabelFontSize: 10, valAxisLabelColor: C.grey,
    valAxisMinVal: 0, valAxisMaxVal: 1150, valAxisMajorUnit: 250,
    valGridLine: { color: "E6E6E6", size: 1 }, catGridLine: { style: "none" },
    showTitle: true, title: "FY26 revenue credit, ₹ crore", titleFontFace: BODY,
    titleFontSize: 12, titleColor: C.grey, titleAlign: "left",
    barGapWidthPct: 110,
  });

  const box = [
    ["Stores give up", "₹84.3 cr", "of webroomed revenue they bank in full today", C.crimson],
    ["Stores recover", "₹46.2 cr", "of digital-booked revenue they see none of today", C.teal],
    ["Net effect on stores", "−₹38 cr", "on the credited base — a 4.3% reduction", C.indigo],
  ];
  let y = 1.62;
  box.forEach(([l, v, d, col]) => {
    s.addText(l, { x: 8.6, y, w: 4.1, h: 0.26, margin: 0, fontFace: BODY, fontSize: 11, bold: true, color: C.grey, charSpacing: 0.5 });
    s.addText(v, { x: 8.6, y: y + 0.22, w: 4.1, h: 0.5, margin: 0, fontFace: HEAD, fontSize: 28, bold: true, color: col });
    s.addText(d, { x: 8.6, y: y + 0.72, w: 4.1, h: 0.44, margin: 0, fontFace: BODY, fontSize: 11, color: C.ink, lineSpacing: 14 });
    y += 1.28;
  });

  card(s, M, 5.24, CW, 1.44, C.band);
  s.addText("A model that only ever moved credit towards the aggrieved party would be a political settlement, not a measurement system — and the digital team would be right to refuse it.", {
    x: M + 0.3, y: 5.4, w: CW - 0.6, h: 0.36, margin: 0, fontFace: HEAD, fontSize: 14, italic: true, color: C.indigo,
  });
  s.addText("Fix store morale in the rate, not the base:  re-base commission from 1.20% of booked sales to 1.254% of credited sales. Target earnings held whole in year one (₹885 cr × 1.20% = ₹847 cr × 1.254% = ₹10.6 cr) — but on a base stores can now grow through digital rather than despite it.", {
    x: M + 0.3, y: 5.84, w: CW - 0.6, h: 0.66, margin: 0, fontFace: BODY, fontSize: 12, color: C.ink, lineSpacing: 16,
  });
  s.addNotes("This is the slide that wins the room. Expect the field to assume the model is a gift to them — it is not, and saying so first is what makes it survive contact with the digital team. Then deliver the resolution: the morale fix is the rate, not the base, and it costs nothing in year one because we hold target earnings whole. Pause here for questions.");
}

/* =========================================================
   7. ANSWER 3 — BUDGET
   ========================================================= */
{
  const s = answerSlide(3, "Budget", "₹139 cr for FY27: 43 / 40 / 17");

  const buckets = [
    ["Store sales support", "₹60.0 cr", "43.2%", "FY26: ₹44 cr", "+36%", C.indigo],
    ["Digital growth", "₹55.0 cr", "39.6%", "FY26: ₹68 cr", "−19%", C.saffron],
    ["Omni-shared, ring-fenced", "₹24.0 cr", "17.3%", "FY26: ₹6 cr", "+300%", C.teal],
  ];
  const cw = 3.94, gap = 0.3;
  buckets.forEach(([t, v, sh, prev, chg, col], i) => {
    const x = M + i * (cw + gap);
    card(s, x, 1.34, cw, 1.72, C.white, { line: col, shadow: true });
    s.addText(t, { x: x + 0.24, y: 1.46, w: cw - 0.48, h: 0.3, margin: 0, fontFace: BODY, fontSize: 12.5, bold: true, color: C.grey });
    s.addText(v, { x: x + 0.24, y: 1.76, w: cw - 1.2, h: 0.6, margin: 0, fontFace: HEAD, fontSize: 32, bold: true, color: col, valign: "middle" });
    s.addText(sh, { x: x + cw - 1.3, y: 1.86, w: 1.06, h: 0.42, margin: 0, fontFace: HEAD, fontSize: 17, bold: true, color: C.grey, align: "right", valign: "middle" });
    s.addText(`${prev}   →   ${chg}`, { x: x + 0.24, y: 2.42, w: cw - 0.48, h: 0.3, margin: 0, fontFace: BODY, fontSize: 11.5, color: C.ink });
  });

  // criteria
  s.addText("Six weighted criteria — the weights are yours to contest, the scores follow from the model", {
    x: M, y: 3.24, w: CW, h: 0.28, margin: 0, fontFace: BODY, fontSize: 13.5, bold: true, color: C.indigoL,
  });
  const crit = [
    ["Contribution per rupee of spend", "30%", "8", "5"],
    ["Incremental growth headroom", "20%", "5", "9"],
    ["Acquisition efficiency and CLV", "15%", "6", "8"],
    ["Committed fixed-cost coverage", "15%", "9", "3"],
    ["Strategic capability, irreversibility", "10%", "4", "9"],
    ["Attributed revenue share", "10%", "7", "3"],
  ];
  const tW = 7.1;
  s.addText([
    { text: "Criterion", options: { bold: true } },
  ], { x: M + 0.2, y: 3.58, w: 3.9, h: 0.26, margin: 0, fontFace: BODY, fontSize: 10.5, color: C.grey });
  ["Weight", "Store", "Digital"].forEach((hd, j) => {
    s.addText(hd, { x: M + 4.2 + j * 0.98, y: 3.58, w: 0.9, h: 0.26, margin: 0, fontFace: BODY, fontSize: 10.5, bold: true, color: C.grey, align: "center" });
  });
  crit.forEach(([n, w8, a, b], i) => {
    const y = 3.86 + i * 0.4;
    if (i % 2 === 0) card(s, M, y, tW, 0.4, C.band2);
    s.addText(n, { x: M + 0.2, y, w: 3.9, h: 0.4, margin: 0, fontFace: BODY, fontSize: 11, color: C.ink, valign: "middle" });
    s.addText(w8, { x: M + 4.2, y, w: 0.9, h: 0.4, margin: 0, fontFace: BODY, fontSize: 11, color: C.grey, align: "center", valign: "middle" });
    s.addText(a, { x: M + 5.18, y, w: 0.9, h: 0.4, margin: 0, fontFace: HEAD, fontSize: 12.5, bold: true, color: C.indigo, align: "center", valign: "middle" });
    s.addText(b, { x: M + 6.16, y, w: 0.9, h: 0.4, margin: 0, fontFace: HEAD, fontSize: 12.5, bold: true, color: C.saffron, align: "center", valign: "middle" });
  });
  card(s, M, 6.26, tW, 0.42, C.band);
  s.addText("Weighted score  →  52.3% / 47.7% of the ₹115 cr that is left after ring-fencing", {
    x: M + 0.2, y: 6.26, w: tW - 0.4, h: 0.42, margin: 0, fontFace: BODY, fontSize: 11, bold: true, color: C.indigo, valign: "middle",
  });

  // reconciliation callout
  card(s, M + tW + 0.3, 3.58, CW - tW - 0.3, 3.1, C.indigo);
  s.addText("Digital's budget falls 19% while its revenue grows 47%.\nIs that defensible?", {
    x: M + tW + 0.55, y: 3.74, w: CW - tW - 0.8, h: 0.8, margin: 0, fontFace: HEAD, fontSize: 13.5, bold: true, color: C.white, lineSpacing: 18,
  });
  s.addText([
    { text: "FY26 digital", options: { color: "BFD0E2", breakLine: true } },
    { text: "₹68.0 cr", options: { color: "FFFFFF", bold: true, breakLine: true } },
    { text: "less non-recurring launch and platform build", options: { color: "BFD0E2", breakLine: true } },
    { text: "− ₹11.0 cr", options: { color: "FFFFFF", bold: true, breakLine: true } },
    { text: "less spend reclassified to Omni-Shared", options: { color: "BFD0E2", breakLine: true } },
    { text: "− ₹14.0 cr", options: { color: "FFFFFF", bold: true, breakLine: true } },
    { text: "Continuing base", options: { color: "BFD0E2", breakLine: true } },
    { text: "₹43.0 cr", options: { color: T.SAFFRON, bold: true } },
  ], { x: M + tW + 0.55, y: 4.64, w: CW - tW - 0.8, h: 1.4, margin: 0, fontFace: BODY, fontSize: 10.5, lineSpacing: 13 });
  s.addText("FY27 of ₹55.0 cr is +28% on the continuing base. Digital's working spend rises — only its ownership of shared spend falls.", {
    x: M + tW + 0.55, y: 6.0, w: CW - tW - 0.8, h: 0.56, margin: 0, fontFace: BODY, fontSize: 10.5, italic: true, color: "E8D9C4", lineSpacing: 13,
  });
  s.addNotes("Lead with the ring-fence: 17% of the budget belongs to neither channel, and pretending otherwise is what started the argument. Then the criteria — invite the CFO to move the weights, the model recalculates. Finish on the reconciliation box: show this to the digital team BEFORE announcing the headline, or you will lose them over a number that is actually a 28% increase.");
}

/* =========================================================
   8. ANSWER 4 — ALIGNMENT
   ========================================================= */
{
  const s = answerSlide(4, "Alignment", "One scorecard, one pool, one gate");

  // variable pay
  s.addText("Rebuilt variable pay — same target cost, different earning basis", {
    x: M, y: 1.3, w: 7.3, h: 0.28, margin: 0, fontFace: BODY, fontSize: 13.5, bold: true, color: C.indigoL,
  });
  const pay = [
    ["Own-channel result", "55%", "45%", "Store: credited sales at 1.254%. Digital: net contribution, not GMV"],
    ["Omni catchment revenue", "25%", "25%", "Literally the same number for both teams — 68% of digital lands within 8 km of a door"],
    ["Shared behaviour scorecard", "20%", "30%", "Capture rate, endless aisle, fulfilment SLA, returns, price parity, NPS"],
  ];
  s.addText("Store", { x: M + 2.6, y: 1.62, w: 0.9, h: 0.24, margin: 0, fontFace: BODY, fontSize: 10, bold: true, color: C.grey, align: "center" });
  s.addText("Digital", { x: M + 3.5, y: 1.62, w: 0.9, h: 0.24, margin: 0, fontFace: BODY, fontSize: 10, bold: true, color: C.grey, align: "center" });
  pay.forEach(([n, a, b, d], i) => {
    const y = 1.9 + i * 0.86;
    card(s, M, y, 7.3, 0.78, i % 2 === 0 ? C.band2 : C.white, { line: i % 2 === 0 ? C.band2 : C.rule });
    s.addText(n, { x: M + 0.2, y: y + 0.06, w: 2.4, h: 0.3, margin: 0, fontFace: BODY, fontSize: 12, bold: true, color: C.indigo });
    s.addText(d, { x: M + 0.2, y: y + 0.34, w: 6.9, h: 0.4, margin: 0, fontFace: BODY, fontSize: 10, color: C.grey, lineSpacing: 13 });
    s.addText(a, { x: M + 2.6, y: y + 0.02, w: 0.9, h: 0.34, margin: 0, fontFace: HEAD, fontSize: 15, bold: true, color: C.indigo, align: "center" });
    s.addText(b, { x: M + 3.5, y: y + 0.02, w: 0.9, h: 0.34, margin: 0, fontFace: HEAD, fontSize: 15, bold: true, color: C.saffron, align: "center" });
  });

  // dual gate
  card(s, 8.32, 1.3, CW - 7.7, 3.22, C.indigo);
  s.addText("₹4.5 cr cross-channel pool", {
    x: 8.56, y: 1.46, w: 4.2, h: 0.34, margin: 0, fontFace: HEAD, fontSize: 17, bold: true, color: C.white,
  });
  s.addText("0.32% of planned net revenue, capped at 10% of any individual's fixed CTC", {
    x: 8.56, y: 1.8, w: 4.2, h: 0.4, margin: 0, fontFace: BODY, fontSize: 10.5, color: "BFD0E2", lineSpacing: 13,
  });
  const gates = [
    ["GATE A", "Store credited revenue\n≥ 97% of plan"],
    ["GATE B", "Digital net contribution\n≥ 100% of plan"],
  ];
  gates.forEach(([g, txt], i) => {
    const x = 8.56 + i * 2.22;
    card(s, x, 2.28, 1.8, 0.94, C.indigoL);
    s.addText(g, { x, y: 2.36, w: 1.8, h: 0.22, margin: 0, fontFace: BODY, fontSize: 9, bold: true, color: T.SAFFRON, align: "center", charSpacing: 1 });
    s.addText(txt, { x: x + 0.06, y: 2.58, w: 1.68, h: 0.56, margin: 0, fontFace: BODY, fontSize: 10, color: C.white, align: "center", lineSpacing: 13 });
  });
  s.addText("AND", { x: 10.36, y: 2.62, w: 0.42, h: 0.3, margin: 0, fontFace: BODY, fontSize: 9, bold: true, color: T.SAFFRON, align: "center" });
  s.addText("Miss either gate and the pool funds at 40% — for everyone. Neither team can win while the other loses. That is the entire mechanism.", {
    x: 8.56, y: 3.34, w: 4.2, h: 0.7, margin: 0, fontFace: BODY, fontSize: 11, color: C.white, lineSpacing: 15,
  });
  s.addText("62% store field  ·  20% digital  ·  18% omni ops and CRM", {
    x: 8.56, y: 4.06, w: 4.2, h: 0.3, margin: 0, fontFace: BODY, fontSize: 10, italic: true, color: "BFD0E2",
  });

  // three more instruments
  const inst = [
    ["Price-parity corridor", "Hard 5-point cap on the store-vs-own-digital gap, plus a store price-match wallet capped at 1.2% of net sales. We cannot control Myntra's discount. We set our own 14-point gap and can close it tomorrow."],
    ["₹90 per store-fulfilled order", "An internal transfer price into the store P&L for ship-from-store and click & collect. Not a bonus — payment for a service the store renders to digital."],
    ["Ten unified KPIs", "Omni catchment revenue, credited revenue, net omni contribution, capture rate, cross-channel share, returns, store-fulfilled share, parity compliance, ATP fill, omni NPS. Four are unmeasured today."],
  ];
  inst.forEach(([t, d], i) => {
    const x = M + i * (CW / 3 + 0.02);
    const w = CW / 3 - 0.28;
    card(s, x, 4.72, w, 1.98, C.band2, { line: C.rule });
    s.addText(t, { x: x + 0.22, y: 4.9, w: w - 0.44, h: 0.3, margin: 0, fontFace: BODY, fontSize: 12.5, bold: true, color: C.teal });
    s.addText(d, { x: x + 0.22, y: 5.22, w: w - 0.44, h: 1.32, margin: 0, fontFace: BODY, fontSize: 10.5, color: C.ink, lineSpacing: 14 });
  });
  s.addNotes("The catchment component is the innovation — it puts ₹303 cr of digital revenue inside a store team's line of sight for the first time, because 68% of it lands within 8 km of a door. Both teams are paid on the identical number. On the dual gate: the 40% floor is deliberate, not punitive — it must hurt enough that each side has a reason to help the other. And the parity corridor is the root-cause fix: the 14-point gap is self-inflicted.");
}

/* =========================================================
   9. RECOMMENDATIONS + ROADMAP
   ========================================================= */
{
  const s = contentSlide("The ask", "Nine moves, sequenced over four quarters");

  const recs = [
    ["Appoint a Chief Commercial Officer", "Remove independent revenue targets from Retail and Digital. Everything else depends on this."],
    ["Adopt the three-ledger credit model", "Freeze the rules for twelve months."],
    ["Restate FY26 on credited revenue and publish it", "₹847 cr store, ₹303 cr digital. The −₹38 cr is what makes it believable."],
    ["Move the budget to ₹60 / ₹55 / ₹24 cr", "Show digital the like-for-like reconciliation before the headline."],
    ["Re-base store variable pay to 1.254% of credited sales", "Written no-loss guarantee for FY27."],
    ["Fund the ₹4.5 cr dual-gated pool", "Converts a truce into cooperation."],
    ["Close the self-inflicted 14-point price gap", "5-point corridor plus a capped price-match wallet."],
    ["Pay ₹90 per ship-from-store or click & collect order", "Tells the field that serving a digital customer is paid work."],
    ["Instrument before you incentivise", "Capture rate to 55% before the pilot, 70% before national rollout."],
  ];
  const cw = (CW - 0.34) / 2;
  recs.forEach(([t, d], i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = M + col * (cw + 0.34), y = 1.42 + row * 0.79;
    s.addShape(pres.ShapeType.ellipse, { x, y: y + 0.06, w: 0.34, h: 0.34, fill: { color: i === 8 ? C.saffron : C.indigo } });
    s.addText(String(i + 1), { x, y: y + 0.06, w: 0.34, h: 0.34, margin: 0, fontFace: BODY, fontSize: 11, bold: true, color: C.white, align: "center", valign: "middle" });
    s.addText(t, { x: x + 0.46, y: y + 0.02, w: cw - 0.46, h: 0.3, margin: 0, fontFace: BODY, fontSize: 12.5, bold: true, color: C.indigo });
    s.addText(d, { x: x + 0.46, y: y + 0.32, w: cw - 0.46, h: 0.42, margin: 0, fontFace: BODY, fontSize: 10.5, color: C.grey, lineSpacing: 13 });
  });

  // roadmap strip
  const qs = [
    ["Q1 FY27", "CCO appointed · data platform and attribution engine built · consent capture live", "Gate: capture ≥ 55%"],
    ["Q2 FY27", "Parallel run in 24 stores · Catchment Councils stood up · parity corridor enforced", "Gate: variance < 3%"],
    ["Q3 FY27", "National rollout of scorecard and re-based pay · ship-from-store in 60 doors", "Gate: capture ≥ 70%"],
    ["Q4 FY27", "Pool pays out on the dual gate · FY28 budget rebuilt on credited revenue", "Gate: audited by Finance"],
  ];
  const qw = (CW - 0.36) / 4;
  qs.forEach(([q, d, g], i) => {
    const x = M + i * (qw + 0.12);
    card(s, x, 5.42, qw, 1.32, i === 0 ? C.indigo : C.band2, { line: i === 0 ? C.indigo : C.rule });
    const fg = i === 0 ? C.white : C.indigo;
    s.addText(q, { x: x + 0.18, y: 5.54, w: qw - 0.36, h: 0.26, margin: 0, fontFace: BODY, fontSize: 11.5, bold: true, color: i === 0 ? T.SAFFRON : C.saffron });
    s.addText(d, { x: x + 0.18, y: 5.8, w: qw - 0.36, h: 0.68, margin: 0, fontFace: BODY, fontSize: 10, color: i === 0 ? "DCE5EF" : C.ink, lineSpacing: 13 });
    s.addText(g, { x: x + 0.18, y: 6.44, w: qw - 0.36, h: 0.24, margin: 0, fontFace: BODY, fontSize: 9.5, bold: true, italic: true, color: i === 0 ? "BFD0E2" : C.grey });
  });
  s.addNotes("Read only the bold lines — the detail is in the report. Emphasise number nine: it is a gate, not an aspiration. If capture stalls below 55% we defer the whole incentive design, because launching it on thin data destroys more trust than the current system does. Then the roadmap: nothing goes national until a 24-store parallel run shows no associate loses money.");
}

/* =========================================================
   10. CLOSE
   ========================================================= */
{
  const s = pres.addSlide();
  s.background = { color: C.indigo };
  s.addText("WHAT SUCCESS LOOKS LIKE AT FY27 EXIT", {
    x: M, y: 0.86, w: CW, h: 0.3, margin: 0,
    fontFace: BODY, fontSize: 12, bold: true, color: T.SAFFRON, charSpacing: 3,
  });

  const kpis = [
    ["39% → 75%", "Store bills carrying a\nverified customer ID"],
    ["9% → 18%", "Active customers buying\nin two or more channels"],
    ["4% → 22%", "Digital orders fulfilled\nby a store"],
    ["6 → 10", "Omni Scorecard KPIs\nactually instrumented"],
  ];
  const kw = (CW - 0.9) / 4;
  kpis.forEach(([v, l], i) => {
    const x = M + i * (kw + 0.3);
    s.addText(v, { x, y: 1.44, w: kw, h: 0.78, margin: 0, fontFace: HEAD, fontSize: 30, bold: true, color: C.white, valign: "middle" });
    s.addShape(pres.ShapeType.line, { x, y: 2.3, w: 0.9, h: 0, line: { color: T.SAFFRON, width: 2 } });
    s.addText(l, { x, y: 2.42, w: kw, h: 0.8, margin: 0, fontFace: BODY, fontSize: 12, color: "BFD0E2", lineSpacing: 17 });
  });

  s.addShape(pres.ShapeType.line, { x: M, y: 3.7, w: CW, h: 0, line: { color: "3A5578", width: 1 } });

  s.addText("And the sentence this deck opened with —", {
    x: M, y: 4.0, w: CW, h: 0.34, margin: 0, fontFace: BODY, fontSize: 14, color: "9FB2C9",
  });
  s.addText("“an associate in Indore fits a customer for twenty minutes and books nothing”", {
    x: M, y: 4.34, w: CW - 0.5, h: 0.78, margin: 0, fontFace: HEAD, fontSize: 25, italic: true, bold: true, color: C.white, lineSpacing: 32,
  });
  s.addText("— is no longer true.", {
    x: M, y: 5.2, w: CW, h: 0.44, margin: 0, fontFace: HEAD, fontSize: 21, color: T.SAFFRON,
  });

  s.addText("Aarav Apparel  ·  Sales & Distribution Management case analysis  ·  July 2026  ·  Full analysis, assumptions and Harvard-style references in the accompanying report", {
    x: M, y: 6.6, w: CW, h: 0.3, margin: 0, fontFace: BODY, fontSize: 10, color: "7E93AE",
  });
  s.addNotes("Close on the vignette that opened. Do not add anything after it. If time remains, take questions on the split percentages and the ring-fence — those are the two things a CFO will push on.");
}

pres.writeFile({ fileName: "Aarav_Apparel_Omnichannel_Executive_Summary.pptx" })
  .then((f) => console.log("written:", f));
