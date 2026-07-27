const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle,
  PageNumber, Footer, Header, PageBreak, PageOrientation, LevelFormat,
  convertInchesToTwip, TabStopType, TabStopPosition,
} = require("docx");
const T = require("./theme");

const W = 9740;            // content width in DXA (A4, 0.75in margins)
const BODY = T.BODY_FONT;
const HEAD = T.HEAD_FONT;

/* ---------- helpers ---------- */
const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const thin = (c) => ({ style: BorderStyle.SINGLE, size: 4, color: c });

function p(text, o = {}) {
  return new Paragraph({
    alignment: o.align || AlignmentType.JUSTIFIED,
    spacing: { after: o.after === undefined ? 105 : o.after, line: o.line || 248 },
    indent: o.indent,
    children: [new TextRun({
      text, font: o.font || BODY, size: o.size || 19,
      color: o.color || T.INK, bold: o.bold, italics: o.italics,
    })],
  });
}

// rich paragraph: array of [text, {bold,italics,color}]
function rp(runs, o = {}) {
  return new Paragraph({
    alignment: o.align || AlignmentType.JUSTIFIED,
    spacing: { after: o.after === undefined ? 105 : o.after, line: o.line || 248 },
    indent: o.indent,
    children: runs.map(([t, s = {}]) => new TextRun({
      text: t, font: s.font || BODY, size: s.size || o.size || 19,
      color: s.color || o.color || T.INK, bold: s.bold, italics: s.italics,
    })),
  });
}

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 230, after: 110 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 10, color: T.SAFFRON, space: 5 } },
    children: [new TextRun({ text, font: HEAD, size: 27, bold: true, color: T.INDIGO })],
  });
}
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 165, after: 70 },
    children: [new TextRun({ text, font: HEAD, size: 22, bold: true, color: T.INDIGO_L })],
  });
}
function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 120, after: 55 },
    children: [new TextRun({ text, font: BODY, size: 20, bold: true, color: T.TEAL })],
  });
}
function bullet(text, lvl = 0) {
  return new Paragraph({
    numbering: { reference: "bul", level: lvl },
    spacing: { after: 52, line: 244 },
    alignment: AlignmentType.JUSTIFIED,
    children: [new TextRun({ text, font: BODY, size: 20, color: T.INK })],
  });
}
function bulletR(runs, lvl = 0) {
  return new Paragraph({
    numbering: { reference: "bul", level: lvl },
    spacing: { after: 52, line: 244 },
    alignment: AlignmentType.JUSTIFIED,
    children: runs.map(([t, s = {}]) => new TextRun({
      text: t, font: BODY, size: 19, color: s.color || T.INK, bold: s.bold, italics: s.italics })),
  });
}
function numbered(runs, lvl = 0, ref = "num") {
  return new Paragraph({
    numbering: { reference: ref, level: lvl },
    spacing: { after: 58, line: 244 },
    alignment: AlignmentType.JUSTIFIED,
    children: runs.map(([t, s = {}]) => new TextRun({
      text: t, font: BODY, size: 19, color: s.color || T.INK, bold: s.bold, italics: s.italics })),
  });
}
function caption(text) {
  return new Paragraph({
    spacing: { before: 45, after: 130 },
    children: [new TextRun({ text, font: BODY, size: 16, italics: true, color: T.GREY })],
  });
}
function exhibitTitle(text) {
  return new Paragraph({
    spacing: { before: 140, after: 55 },
    children: [new TextRun({ text, font: BODY, size: 17, bold: true, color: T.SAFFRON, allCaps: false })],
  });
}

function cell(text, w, o = {}) {
  const runs = Array.isArray(text) ? text : [[text, {}]];
  return new TableCell({
    width: { size: w, type: WidthType.DXA },
    shading: o.fill ? { type: ShadingType.CLEAR, fill: o.fill, color: "auto" } : undefined,
    margins: { top: 44, bottom: 44, left: 85, right: 85 },
    verticalAlign: o.vAlign || "center",
    columnSpan: o.span,
    borders: {
      top: thin(T.RULE), bottom: thin(T.RULE), left: thin(T.RULE), right: thin(T.RULE),
    },
    children: [new Paragraph({
      alignment: o.align || AlignmentType.LEFT,
      spacing: { after: 0, line: 228 },
      children: runs.map(([t, s = {}]) => new TextRun({
        text: t, font: BODY, size: o.size || 16,
        bold: o.bold || s.bold, italics: s.italics,
        color: o.color || s.color || T.INK,
      })),
    })],
  });
}

function table(cols, headers, rows, opts = {}) {
  const hdrFill = opts.headFill || T.INDIGO;
  const trs = [];
  if (headers) {
    trs.push(new TableRow({
      tableHeader: true,
      children: headers.map((t, i) =>
        cell(t, cols[i], { fill: hdrFill, bold: true, color: "FFFFFF", size: 15,
          align: opts.headAlign ? opts.headAlign[i] : (i === 0 ? AlignmentType.LEFT : AlignmentType.CENTER) })),
    }));
  }
  rows.forEach((r, ri) => {
    const isTotal = opts.totalRows && opts.totalRows.includes(ri);
    trs.push(new TableRow({
      children: r.map((t, i) => cell(t, cols[i], {
        fill: isTotal ? T.BAND : (ri % 2 === 1 ? T.BAND_2 : undefined),
        bold: isTotal,
        align: opts.align ? opts.align[i] : (i === 0 ? AlignmentType.LEFT : AlignmentType.CENTER),
      })),
    }));
  });
  return new Table({ columnWidths: cols, width: { size: W, type: WidthType.DXA }, rows: trs });
}

// Callout / box
function box(titleText, paras, accent = T.SAFFRON, fill = T.BAND_2) {
  return new Table({
    columnWidths: [W],
    width: { size: W, type: WidthType.DXA },
    rows: [new TableRow({
      children: [new TableCell({
        width: { size: W, type: WidthType.DXA },
        shading: { type: ShadingType.CLEAR, fill, color: "auto" },
        margins: { top: 105, bottom: 105, left: 165, right: 165 },
        borders: {
          top: thin(fill), bottom: thin(fill), right: thin(fill),
          left: { style: BorderStyle.SINGLE, size: 22, color: accent },
        },
        children: [
          ...(titleText ? [new Paragraph({
            spacing: { after: 70 },
            children: [new TextRun({ text: titleText, font: HEAD, size: 18, bold: true, color: T.INDIGO })],
          })] : []),
          ...paras,
        ],
      })],
    })],
  });
}

const spacer = (n = 120) => new Paragraph({ spacing: { after: n }, children: [] });

/* =======================================================================
   CONTENT
   ======================================================================= */
const body = [];

/* ---------- Title block ---------- */
body.push(new Paragraph({
  spacing: { after: 40 },
  children: [new TextRun({ text: "SALES & DISTRIBUTION MANAGEMENT  |  CASE ANALYSIS", font: BODY, size: 17, bold: true, color: T.SAFFRON, characterSpacing: 30 })],
}));
body.push(new Paragraph({
  spacing: { after: 60 },
  children: [new TextRun({ text: "Aarav Apparel: Ending the Channel War", font: HEAD, size: 40, bold: true, color: T.INDIGO })],
}));
body.push(new Paragraph({
  spacing: { after: 140 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 14, color: T.INDIGO, space: 8 } },
  children: [new TextRun({ text: "An omni-channel sales organisation, attribution logic, budget split and alignment mechanism for FY27", font: HEAD, size: 22, italics: true, color: T.GREY })],
}));
body.push(new Paragraph({
  spacing: { after: 240 },
  children: [new TextRun({ text: "Prepared for the Chief Financial Officer  ·  120 stores  ·  D2C and marketplace  ·  July 2026", font: BODY, size: 17, color: T.GREY })],
}));

/* ---------- 1. Situation ---------- */
body.push(h1("1. Situation Summary"));

body.push(rp([
  ["Aarav Apparel enters FY27 with a commercial system that is working and an organisation that is not. Booked revenue of ", {}],
  ["₹1,150 crore", { bold: true }],
  [" in FY26 came 77% from 120 own stores and 23% from a digital business that did not exist three years ago. Every headline number is healthy. Every internal number that matters — capture rate, cross-channel customer share, field engagement, price integrity — is deteriorating.", {}],
]));

body.push(p("The presenting complaint is showrooming. A customer walks into an Aarav store in Indore, is fitted by an associate for twenty minutes, photographs the tag, and buys the same style that evening on Myntra at 26% off. The store books nothing. The associate's variable pay records nothing. The digital P&L books the full sale and the digital team books the full credit. Repeated a few thousand times a month across 120 catchments, this is no longer a grievance about a single transaction; it is a structural mis-measurement that is quietly repricing the effort of 1,080 store-facing employees to zero."));

body.push(p("The resentment is rational rather than emotional. Store staff are held to a 1.20% commission on own-store till sales while the company deliberately runs an average online discount of 26% against 12% in store — a 14-point gap the field cannot match and did not set. Digital is credited with 100% of any journey that closes online, which is the same as saying that discovery, fitting, trial and trust are worth nothing. On the credit model in Section 5, stores materially influenced 85% of FY26 revenue, not 77%."));

body.push(p("Three market forces make this urgent rather than merely unfair. First, quick commerce has arrived in fashion. Myntra's M-Now now runs a 30-minute proposition from more than 87 dark stores across ten markets including Patna, Jaipur, Lucknow and Ahmedabad (Indian Retailer, 2025; TechCrunch, 2024), which removes the last operational advantage a store held over an app — immediacy. Second, marketplace economics are tightening: effective deductions of 30–40% of selling price once commission, logistics and returns are counted (Terra Insight, 2025) mean marketplace growth buys revenue at roughly two-thirds the gross margin of a store sale. Third, the traffic is moving in the other direction too — India's leading D2C brands, having exhausted cheap online acquisition, are opening physical stores at speed, with Snitch targeting 100 stores and The Souled Store 200 (Inc42, 2025; Exchange4media, 2025). Aarav already owns the 120-store asset those brands are paying to build. It is currently accounting for that asset as a cost centre."));

body.push(box("The CFO's real question", [
  p("Not \"how much should digital get?\" but \"what is the store network actually worth to the enterprise once we measure it honestly, and what organisation, credit system, budget and incentive structure follow from that number?\" The four answers in this report are internally consistent: change one and the others must move.", { after: 0, italics: true }),
]));

/* ---------- 2. Assumptions ---------- */
body.push(h1("2. Assumptions"));
body.push(p("No figure below was supplied in the brief. All are stated assumptions, chosen to be internally consistent and plausible for a mid-size Indian apparel retailer of this footprint. Every number used anywhere in this report derives from this box. 1 crore = 10 million; ₹ figures are net of GST. Sensitivities to the three assumptions that most move the conclusion are given in Section 9.",
  { italics: true, color: T.GREY }));

body.push(exhibitTitle("Exhibit 1 — Assumptions box"));
body.push(table([2960, 2180, 4600],
  ["Parameter", "Assumed value (FY26)", "Basis / note"],
  [
    ["Net revenue", "₹1,150 cr", "Base year FY26 (ended Mar-26)"],
    ["Channel mix — own stores", "₹885 cr (77.0%)", "Exclusive brand outlets, company-operated"],
    ["Channel mix — marketplace", "₹184 cr (16.0%)", "Myntra ~46%, Ajio ~24%, Flipkart/Amazon ~30% of this"],
    ["Channel mix — D2C site & app", "₹81 cr (7.0%)", "Launched FY25"],
    ["Store count / net adds FY27", "120 / +18", "58 metro-mall, 44 tier-2 high street, 18 tier-3"],
    ["Average store carpet area", "2,300 sq ft", "Format average"],
    ["Revenue per store p.a.", "₹7.38 cr", "₹885 cr ÷ 120"],
    ["Average selling price (net)", "₹1,780", "64.6 lakh units FY26"],
    ["Gross margin by channel", "Store 52% · D2C 55% · MP 37%", "Marketplace net of commission, ads, returns"],
    ["Store contribution margin", "17.5% of store net sales", "After rent 9.5%, payroll 6.8%, utilities/other"],
    ["Return rate", "Store 3% · D2C 17% · MP 29%", "Blended 14.2%"],
    ["Average discount to MRP", "Store 12% · D2C 18% · MP 26%", "The 14-point store-vs-marketplace gap"],
    ["Headcount — store-facing", "1,080 (avg 9/store)", "+16 area managers, +4 zonal heads"],
    ["Headcount — digital", "62", "D2C 34, marketplace 20, quick commerce 8"],
    ["Headcount — omni ops / CRM / planning", "45", "Currently split across functions"],
    ["Identified-customer capture (ICR)", "39% of store bills", "100% online by definition"],
    ["Loyalty base", "4.2 mn members, 1.6 mn active", "Active = transacted in last 12 months"],
    ["Cross-channel customer share", "9% of active customers", "Transacted in 2 or more channels"],
    ["Blended new-customer CAC", "D2C ₹690 · MP ₹255", "Media only; excludes commission"],
    ["Digital revenue landing within 8 km of a store", "68%", "Catchment overlap; drives the catchment KPI"],
    ["Store-fulfilled share of digital orders", "4%", "Ship-from-store and click & collect"],
    ["Average fixed CTC", "Associate ₹2.9 L · Digital ₹14 L", "Used only to size the incentive pool"],
    ["FY27 plan — net revenue", "₹1,392 cr (+21%)", "Stores ₹1,001 cr · MP ₹258 cr · D2C ₹133 cr"],
    ["S&D budget under CFO review, FY27", "₹139 cr (10.0% of plan)", "FY26 actual ₹118 cr"],
    ["FY26 split of that budget", "Store ₹44 cr · Digital ₹68 cr · Shared ₹6 cr", "37% / 58% / 5% — the imbalance in question"],
    ["Cost of capital (WACC)", "12.5%", "Hurdle for budget lines: 18% 3-yr IRR"],
  ],
  { align: [AlignmentType.LEFT, AlignmentType.CENTER, AlignmentType.LEFT] }));
body.push(caption("All values assumed for analysis. Percentages may not sum exactly owing to rounding."));

/* ---------- 3. Analytical approach ---------- */
body.push(h1("3. Analytical Approach"));
body.push(p("Four established lenses are applied to Aarav's facts. They are used as tools, not described as theory."));

body.push(exhibitTitle("Exhibit 2 — Lens applied to question"));
body.push(table([3060, 2180, 4500],
  ["Lens", "Applied to", "What it forces us to conclude"],
  [
    ["Channel conflict and channel power (Stern and El-Ansary, 1992; Ailawadi and Farris, 2017)", "Q1 — organisation", "Conflict is a symptom of two units optimising different objective functions. Fix the objective function above the units, not the behaviour inside them."],
    ["Responsibility-centre design and internal transfer pricing (Anthony and Govindarajan, 2007)", "Q2, Q4 — credit and fulfilment", "Credit and cash must be allowed to diverge. Ship-from-store is an internal service and must carry an internal price."],
    ["Fractional credit allocation in the spirit of Shapley (1953), operationalised as rule-based multi-touch attribution", "Q2 — attribution", "Credit should track marginal contribution to the journey, not the location of the till. Symmetry is the test of fairness."],
    ["Expectancy-based compensation and scorecard cascade (Churchill, Ford and Walker, 2000; Zoltners, Sinha and Lorimer, 2008; Kaplan and Norton, 1996)", "Q4 — alignment", "Effort follows measurable, attainable, personally consequential metrics. A metric an associate cannot influence produces cynicism, not effort."],
  ],
  { align: [AlignmentType.LEFT, AlignmentType.LEFT, AlignmentType.LEFT] }));

body.push(p("The empirical anchor is that omni-channel customers are worth more than single-channel ones: the 46,000-shopper study by Sopadjieva, Dholakia and Benjamin (2017) found multi-channel shoppers spent roughly 10% more online and were worth about 30% more over their lifetime, and Indian retailers such as Shoppers Stop have built click-and-collect, endless-aisle and ship-from-store capability on precisely that logic (Indian Apparel, n.d.). If that holds at Aarav, then any system which taxes the store for creating a digital customer is destroying value, not merely allocating it badly."));


/* ---------- 4. Q1 Org ---------- */
body.push(h1("4. Question 1 — The Omni-Channel Sales Organisation"));

body.push(p("Aarav today runs two sales organisations that meet only in the CEO's review. Retail reports to a Retail Director; D2C and marketplace report to a Chief Digital Officer with an independent P&L and growth target. Two P&Ls chasing the same customer in the same pincode will compete for that customer, because each is rewarded for doing so."));

body.push(h2("4.1 Recommended structure"));
body.push(p("Collapse the two into one commercial organisation under a new Chief Commercial Officer (CCO) who owns total demand — every rupee, every channel. Below the CCO sit four directors: two channel-facing, two shared. The design rule is explicit and should be written into the mandate: no unit below the CCO owns a customer; units own channels, the CCO owns customers."));

body.push(exhibitTitle("Exhibit 3 — Reporting lines"));
body.push(table([700, 2680, 2880, 3480],
  ["Level", "Role", "Solid line to", "Dotted line to (and for what)"],
  [
    ["1", "Chief Commercial Officer (new)", "CEO", "CFO — owns the single demand plan and the S&D budget"],
    ["2", "Director, Retail Sales (120→138 stores)", "CCO", "Dir. Omni Operations — fulfilment SLA and stock accuracy"],
    ["2", "Director, Digital Commerce", "CCO", "Dir. Retail Sales — catchment demand plan sign-off"],
    ["2", "Director, Omni Operations (shared)", "CCO", "COO — supply chain execution"],
    ["2", "Director, Customer & CRM (shared)", "CCO", "CMO — brand and loyalty proposition"],
    ["3", "4 × Zonal Retail Head (N/S/E/W)", "Dir. Retail Sales", "Digital Catchment Leads — joint zone forecast"],
    ["3", "D2C Lead · Marketplace Lead · Quick-Commerce Lead", "Dir. Digital Commerce", "Zonal Retail Heads — catchment targets"],
    ["4", "16 × Area Manager (7–9 stores)", "Zonal Retail Head", "Chairs the Catchment Council (see 4.3)"],
    ["4", "16 × Digital Catchment Analyst (new, redeployed)", "Digital Commerce", "Area Manager — sits in the Catchment Council"],
    ["5", "138 × Store Manager", "Area Manager", "—"],
    ["6", "~1,240 × Store associates / stylists", "Store Manager", "—"],
  ],
  { align: [AlignmentType.CENTER, AlignmentType.LEFT, AlignmentType.LEFT, AlignmentType.LEFT] }));

body.push(h2("4.2 The three structural moves that matter"));
body.push(numbered([["One P&L above the channels. ", { bold: true }],
  ["Retail and Digital keep contribution reporting for cost control but lose independent revenue targets. Only the CCO carries a revenue number. This removes the incentive for either director to win a transfer at the other's expense."]]));
body.push(numbered([["Shared services are peers, not support. ", { bold: true }],
  ["Omni Operations (availability-to-promise, ship-from-store, click and collect, returns network) and Customer & CRM (single customer ID, loyalty, catchment intelligence) report at the same level as the channels. If they report into either channel, the data and the inventory will be rationed in that channel's favour."]]));
body.push(numbered([["The Digital Catchment Analyst is the hinge role. ", { bold: true }],
  ["Sixteen analysts, one per area, redeployed from the existing performance-marketing pod. Each is accountable for digital revenue inside their area's pincode clusters and sits in that area's Catchment Council. This gives digital a named human being who wins or loses alongside the stores — the single most effective de-escalation available."]]));

body.push(h2("4.3 Shared accountabilities"));
body.push(p("Reporting lines allocate authority; shared accountabilities allocate blame, which is what actually changes behaviour. Sixteen Catchment Councils meet monthly — Area Manager chairs, Digital Catchment Analyst and an Omni Ops planner are standing members — and jointly own one forecast, one stock plan and one promotion calendar for their cluster. Their decisions bind both channels."));

body.push(exhibitTitle("Exhibit 4 — Shared accountability matrix (A = accountable, R = responsible, C = consulted)"));
body.push(table([3380, 1590, 1590, 1590, 1590],
  ["Decision / outcome", "Retail Sales", "Digital Commerce", "Omni Ops", "Customer & CRM"],
  [
    ["Catchment revenue (all channels, 8 km radius)", "A", "A", "C", "C"],
    ["Price and promotion within the parity corridor", "C", "R", "C", "A"],
    ["Assortment and buy plan by catchment", "R", "R", "A", "C"],
    ["Availability-to-promise / stock accuracy", "R", "C", "A", "—"],
    ["Ship-from-store and click & collect SLA", "R", "C", "A", "—"],
    ["Identified-customer capture rate (ICR)", "A", "R", "—", "A"],
    ["Blended return rate", "R", "A", "R", "C"],
    ["Net omni contribution per catchment", "A", "A", "R", "C"],
    ["Attribution rule changes", "C", "C", "C", "A"],
  ],
  { align: [AlignmentType.LEFT, AlignmentType.CENTER, AlignmentType.CENTER, AlignmentType.CENTER, AlignmentType.CENTER] }));
body.push(caption("Deliberate design choice: four outcomes carry two accountable owners. Dual accountability is normally poor practice, but here it is the point — neither party can deliver these alone, and single ownership is what produced the current conflict."));

/* ---------- 5. Q2 Attribution ---------- */
body.push(h1("5. Question 2 — Attribution and Budget-Allocation Logic"));

body.push(p("The root error is treating one number as three. Aarav currently uses booked revenue for statutory reporting, for performance measurement and for incentive payout. These have different purposes and must be allowed to differ."));

body.push(h2("5.1 A three-ledger model"));
body.push(exhibitTitle("Exhibit 5 — Three ledgers, one transaction"));
body.push(table([1700, 2280, 3060, 2700],
  ["Ledger", "Question it answers", "Rule", "Used for"],
  [
    ["Revenue credit", "Where did the cash land?", "100% to the channel that took payment. Sums to exactly 100%.", "Statutory accounts, GST, channel P&L"],
    ["Influence credit", "Who touched this journey?", "100% to every channel with a qualifying touch. Sums to more than 100% by design.", "Budget allocation, catchment planning, capability cases"],
    ["Incentive credit", "Who gets paid for this?", "Fractional split by the rules in Exhibit 6. Sums to exactly 100%.", "Variable pay, bonus pool, league tables"],
  ],
  { align: [AlignmentType.LEFT, AlignmentType.LEFT, AlignmentType.LEFT, AlignmentType.LEFT] }));

body.push(p("Separating these three resolves the argument immediately. Digital keeps every rupee of revenue credit it books today — nothing changes in the audited accounts. What changes is that revenue credit stops being used for the two purposes it was never fit for."));

body.push(h2("5.2 Incentive-credit split rules"));
body.push(p("Matching key is the verified mobile number or loyalty ID. Lookback windows: 14 days for a store visit preceding an online purchase, 30 days for a digital touch preceding a store purchase. Splits are set by which party bore the decisive cost and risk of the journey — in fashion, fit and trial carry more weight than impression, which is why webrooming favours the store."));

body.push(exhibitTitle("Exhibit 6 — Journey credit rules and FY26 restatement (₹ crore)"));
body.push(table([2560, 1180, 1000, 1000, 1000, 1500, 1500],
  ["Journey type", "Booked in", "Revenue ₹cr", "Store %", "Digital %", "Store credit", "Digital credit"],
  [
    ["Pure walk-in, no digital touch", "Store", "604", "100", "0", "604.0", "0.0"],
    ["Webrooming: online research → store purchase (30 d)", "Store", "281", "70", "30", "196.7", "84.3"],
    ["Showrooming: store visit → online purchase (14 d)", "Digital", "71", "45", "55", "32.0", "39.0"],
    ["Endless aisle / associate-initiated online order", "Digital", "9", "100", "0", "9.0", "0.0"],
    ["Click & collect and ship-from-store", "Digital", "13", "40", "60", "5.2", "7.8"],
    ["Marketplace, no store linkage", "Digital", "139", "0", "100", "0.0", "139.0"],
    ["D2C pure online, no store linkage", "Digital", "33", "0", "100", "0.0", "33.0"],
    ["Total FY26", "", "1,150", "", "", "846.9", "303.1"],
    ["Memo: as booked today", "", "1,150", "", "", "885.0", "265.0"],
    ["Memo: influence view (sums >100%)", "", "1,150", "", "", "978 (85.0%)", "546 (47.5%)"],
  ],
  { totalRows: [7],
    align: [AlignmentType.LEFT, AlignmentType.CENTER, AlignmentType.RIGHT, AlignmentType.CENTER, AlignmentType.CENTER, AlignmentType.RIGHT, AlignmentType.RIGHT] }));

body.push(box("The result is deliberately not a giveaway to stores", [
  rp([["On incentive credit, stores move from ₹885 cr to ₹847 cr — ", {}],
      ["down ₹38 cr, not up", { bold: true }],
      [". They surrender ₹84.3 cr of webroomed revenue they currently bank in full, and recover ₹46.2 cr of digital-booked revenue they currently see none of. Digital moves from ₹265 cr to ₹303 cr, ", {}],
      ["up ₹38 cr", { bold: true }],
      [". This symmetry is the credibility test. A model that only ever moved credit towards the aggrieved party would be a political settlement, not a measurement system, and the digital team would correctly refuse it.", {}]], { after: 90 }),
  rp([["The store morale problem is then solved where it belongs — in the ", {}],
      ["rate", { italics: true }],
      [", not the base. Re-basing store commission from 1.20% of booked sales to 1.254% of credited sales holds target earnings whole in year one (₹885 cr × 1.20% = ₹847 cr × 1.254% = ₹10.6 cr), while the new base is one stores can grow through digital rather than despite it.", {}]], { after: 0 }),
], T.TEAL));

body.push(h2("5.3 Unmatched journeys and audit"));
body.push(bullet("At 39% store-bill capture, roughly six in ten store visits are invisible. The ₹71 cr showrooming figure is therefore an identified floor, not the true total; grossed up at the observed capture rate the real number is closer to ₹150–180 cr. This is stated as an estimate and must not be paid on."));
body.push(bullet("Rule: only verified matches earn incentive credit. Unmatched volume is sized quarterly by geo-holdout testing and reported in a separate \"dark journeys\" line used for budgeting only. This keeps the payout system auditable while stopping the company from believing the leakage is small."));
body.push(bullet("Anti-gaming: capture quality is scored on OTP-verified matches only; invalid or duplicate numbers are stripped and repeat offenders lose the behaviour component of variable pay for the quarter. ICR is the metric most likely to be gamed and must be audited monthly."));
body.push(bullet("Consent and privacy: identity resolution across store and digital requires explicit consent under the Digital Personal Data Protection Act, 2023. Consent capture must be built into the associate app before the model goes live, not after."));

body.push(h2("5.4 From credit to budget"));
body.push(p("Budget follows influence credit, not revenue credit — because budget buys future journeys, and a journey is bought wherever it starts. On the influence view, stores touch 85% of revenue and digital 47.5%; normalised, that is a 64:36 weighting of demand-generating effort. That weighting is one of six inputs to the budget in Section 6, not the answer on its own, because influence measures reach and not marginal return."));

/* ---------- 6. Q3 Budget ---------- */
body.push(h1("6. Question 3 — FY27 Budget Split"));

body.push(p("The budget under review is ₹139 crore, 10.0% of planned FY27 net revenue of ₹1,392 crore. The FY26 split — store ₹44 cr, digital ₹68 cr, shared ₹6 cr — put 58% of discretionary commercial spend behind 23% of booked revenue and 26% of credited revenue. Some of that skew was correct: a new channel needs disproportionate investment. Most of it is now stale."));

body.push(h2("6.1 Allocation criteria and weights"));
body.push(p("Six criteria, weighted and scored 1–10. Weights are the CFO's judgement call and are stated so they can be contested; the scores follow from Sections 2 and 5."));

body.push(exhibitTitle("Exhibit 7 — Weighted allocation criteria"));
body.push(table([3050, 880, 1180, 1180, 3450],
  ["Criterion", "Weight", "Store score", "Digital score", "Why scored this way"],
  [
    ["Contribution margin per rupee of spend", "30%", "8", "5", "Store CM 17.5% vs marketplace GM 37% before a 30–40% effective platform deduction"],
    ["Incremental growth headroom", "20%", "5", "9", "Digital plan +47%; store LFL assumed +6–8% plus 18 new doors"],
    ["Acquisition efficiency and lifetime-value build", "15%", "6", "8", "D2C CAC ₹690 but multi-channel customers worth ~30% more (Sopadjieva et al., 2017)"],
    ["Coverage of committed fixed cost and asset utilisation", "15%", "9", "3", "₹84 cr of annual store rent and payroll is already committed; digital cost is variable"],
    ["Strategic capability and irreversibility", "10%", "4", "9", "Quick-commerce readiness, marketplace ranking and first-party data compound; shelf space does not"],
    ["Attributed revenue share (Exhibit 6)", "10%", "7", "3", "73.7% / 26.3% credited"],
    ["Weighted score", "100%", "6.75", "6.15", "Normalised: 52.3% / 47.7%"],
  ],
  { totalRows: [6],
    align: [AlignmentType.LEFT, AlignmentType.CENTER, AlignmentType.CENTER, AlignmentType.CENTER, AlignmentType.LEFT] }));

body.push(h2("6.2 The split"));
body.push(p("Ring-fence the genuinely joint spend first, then apply the weighted score to the remainder. ₹24 cr of the ₹139 cr is spend that cannot honestly be assigned to either channel — the customer data platform, geo-fenced catchment media, the omni inventory and returns network, the price-parity corridor and the cross-channel incentive pool. Assigning it to either team is what created the dispute in the first place. The residual ₹115 cr splits 52.3 : 47.7."));

body.push(exhibitTitle("Exhibit 8 — FY27 budget allocation (₹ crore)"));
body.push(table([3900, 1350, 1350, 1550, 1590],
  ["Bucket", "FY26 actual", "FY27 plan", "Change", "Share of FY27"],
  [
    ["Store sales support", "44.0", "60.0", "+16.0  (+36%)", "43.2%"],
    ["Digital growth", "68.0", "55.0", "−13.0  (−19%)", "39.6%"],
    ["Omni-shared (ring-fenced)", "6.0", "24.0", "+18.0  (+300%)", "17.3%"],
    ["Total", "118.0", "139.0", "+21.0  (+18%)", "100.0%"],
  ],
  { totalRows: [3],
    align: [AlignmentType.LEFT, AlignmentType.RIGHT, AlignmentType.RIGHT, AlignmentType.RIGHT, AlignmentType.CENTER] }));

body.push(box("Digital's budget falls 19% while its revenue grows 47% — is that defensible?", [
  p("Yes, on a like-for-like basis, and the reconciliation must be shown to the digital team before the number is announced:", { after: 90 }),
  rp([["FY26 digital ₹68.0 cr, less ₹11.0 cr of non-recurring launch and platform build, less ₹14.0 cr reclassified into Omni-Shared (geo-fenced catchment media and CRM, which serve both channels), gives a continuing base of ", {}],
      ["₹43.0 cr", { bold: true }],
      [". FY27 digital of ₹55.0 cr is therefore ", {}],
      ["+28% on the continuing base", { bold: true }],
      [", not −19%. Digital's working spend rises; only its ownership of shared spend falls.", {}]], { after: 0 }),
], T.INDIGO_L, T.BAND));

body.push(exhibitTitle("Exhibit 9 — What the money buys"));
body.push(table([3900, 900, 4940],
  ["Line item", "₹ cr", "Test it must pass"],
  [
    ["STORE SALES SUPPORT — ₹60.0 cr", "", ""],
    ["18 new-store launch, staffing and opening activation", "5.0", "New-door payback within 22 months at ₹7.4 cr average revenue"],
    ["Store digital toolkit: endless-aisle tablets, associate app, clienteling", "4.0", "Endless-aisle orders from ₹9 cr to ₹34 cr; 138 doors at ₹2.9 L per door"],
    ["Ship-from-store and click & collect enablement", "3.0", "Store-fulfilled share of digital orders 4% → 22%"],
    ["Field capability, stylist certification, fit training", "2.0", "Store conversion 22% → 25%; measured by mystery audit"],
    ["Catchment-level BTL and local activation", "2.0", "Footfall per door +8% in the 44 tier-2 catchments"],
    ["Continuing store support base (VM, trade schemes, field ops)", "44.0", "Held flat in absolute terms; −7% as a share of store sales"],
    ["DIGITAL GROWTH — ₹55.0 cr", "", ""],
    ["Marketplace commission support, ads and platform events", "24.0", "Contribution per order, not GMV; exit ROAS floor of 4.2×"],
    ["D2C performance media and retention", "14.0", "Blended D2C CAC held at ₹690 despite +64% revenue"],
    ["Quick-commerce readiness (M-Now / Flipkart Minutes / Blinkit pilots)", "7.0", "Pilot in 6 cities; no scale-up without positive contribution"],
    ["Platform, content, catalogue and site operations", "10.0", "Cost per SKU listed down 15%"],
    ["OMNI-SHARED (ring-fenced) — ₹24.0 cr", "", ""],
    ["Cross-channel incentive pool (Section 7)", "4.5", "Dual-gated; funds at 40% if either gate is missed"],
    ["Customer data platform, identity resolution, attribution engine", "6.5", "ICR 39% → 75%; without this nothing else in this report works"],
    ["Geo-fenced catchment performance media (store-credited)", "6.5", "Omni catchment revenue per door +18%"],
    ["Omni inventory ATP, ship-from-store network, returns", "4.0", "Blended return rate 14.2% → 12.0%"],
    ["Price-parity corridor and store price-match wallet", "2.5", "Capped at 1.2% of store net sales; manager-authorised"],
    ["TOTAL", "139.0", "Every line: 18% three-year IRR or 15-month contribution payback"],
  ],
  { totalRows: [0, 7, 12, 18],
    align: [AlignmentType.LEFT, AlignmentType.RIGHT, AlignmentType.LEFT] }));

body.push(h2("6.3 Stated decision criteria in plain terms"));
body.push(bulletR([["Spend follows contribution, not revenue. ", { bold: true }], ["A marketplace rupee carries a 30–40% effective platform deduction (Terra Insight, 2025). Growing it is a choice to buy revenue at roughly two-thirds the margin of a store rupee, and must be argued on strategic grounds rather than assumed."]]));
body.push(bulletR([["Committed cost gets protected first. ", { bold: true }], ["₹84 cr of store rent and payroll is contracted for FY27 whether or not anyone walks in. Underfunding traffic into a fixed-cost asset is the most expensive saving available to this company."]]));
body.push(bulletR([["Joint spend is ring-fenced, not allocated. ", { bold: true }], ["17.3% of the budget belongs to neither channel. Ring-fencing it removes the single largest source of the current argument."]]));
body.push(bulletR([["Capability spend clears a lower bar than volume spend. ", { bold: true }], ["The customer data platform and quick-commerce pilots are optioned against irreversibility, not this year's return."]]));
body.push(bulletR([["Nothing enters the budget without a named test. ", { bold: true }], ["Every line in Exhibit 9 carries a measurable exit condition and is reviewed at the half-year."]]));

/* ---------- 7. Q4 Alignment ---------- */
body.push(h1("7. Question 4 — The Alignment Mechanism"));

body.push(p("Structure and attribution set the conditions for alignment; they do not create it. Alignment happens when a store associate in Coimbatore and a performance marketer in Bengaluru are looking at the same number and both of their payslips move with it. Two instruments deliver that: one scorecard, and one pool."));

body.push(h2("7.1 The Omni Scorecard — one metric set, both teams"));
body.push(exhibitTitle("Exhibit 10 — Unified KPIs, FY26 baseline and FY27 exit target"));
body.push(table([2450, 3400, 1150, 1150, 1590],
  ["KPI", "Definition", "FY26", "FY27 target", "Owner"],
  [
    ["Omni Catchment Revenue", "All-channel revenue in a store's 8 km pincode cluster", "₹8.88 cr/door", "₹9.83 cr/door", "Joint"],
    ["Credited revenue", "Exhibit 6 incentive credit", "847 / 303", "1,010 / 382", "Joint"],
    ["Net Omni Contribution", "Contribution after channel cost-to-serve incl. returns and commission", "₹196 cr", "₹247 cr", "Joint"],
    ["Identified Customer Capture", "% of store bills with an OTP-verified ID", "39%", "75%", "Retail + CRM"],
    ["Cross-channel customer share", "% of active customers buying in 2+ channels", "9%", "18%", "CRM"],
    ["Blended return rate", "Returns as % of gross sales, all channels", "14.2%", "12.0%", "Digital + Omni Ops"],
    ["Store-fulfilled digital order share", "SFS + C&C as % of digital orders", "4%", "22%", "Omni Ops"],
    ["Price-parity compliance", "% SKU-days inside the 5-point corridor", "not measured", "≥95%", "CRM + Digital"],
    ["Availability-to-promise fill rate", "Promised stock actually shippable", "not measured", "≥97%", "Omni Ops"],
    ["Omni NPS", "Post-journey, channel-agnostic", "not measured", "≥52", "Joint"],
  ],
  { align: [AlignmentType.LEFT, AlignmentType.LEFT, AlignmentType.CENTER, AlignmentType.CENTER, AlignmentType.CENTER] }));
body.push(caption("Four of ten KPIs are currently unmeasured. That is itself a finding: Aarav cannot presently answer whether its channels are cooperating."));

body.push(h2("7.2 Rebuilt variable pay"));
body.push(p("Total target variable cost is held constant. What changes is what it is earned on. The catchment component is the critical innovation: it puts the ₹303 cr of digital-credited revenue partly inside a store team's line of sight for the first time, because roughly 68% of digital revenue lands within 8 km of a store."));

body.push(exhibitTitle("Exhibit 11 — Variable pay composition"));
body.push(table([2450, 1250, 1250, 4790],
  ["Component", "Store field", "Digital team", "Note"],
  [
    ["Own-channel result", "55%", "45%", "Store: credited own-store sales at a re-based 1.254%. Digital: net contribution, not GMV — this removes the discount-to-grow reflex"],
    ["Omni catchment revenue", "25%", "25%", "Literally the same number for both teams, at 0.35% for the store team — approx. ₹67,500 per door p.a. at FY27 plan"],
    ["Shared behaviour scorecard", "20%", "30%", "ICR, endless-aisle orders raised, SFS/C&C SLA, blended return rate, price-parity compliance, omni NPS"],
  ],
  { align: [AlignmentType.LEFT, AlignmentType.CENTER, AlignmentType.CENTER, AlignmentType.LEFT] }));

body.push(h2("7.3 The dual-gated cross-channel pool"));
body.push(bulletR([["Size: ", { bold: true }], ["₹4.5 cr, funded at 0.32% of planned net revenue, capped at 10% of any individual's fixed CTC."]]));
body.push(bulletR([["Dual gate: ", { bold: true }], ["the pool funds in full only if BOTH store credited revenue reaches 97% of plan AND digital net contribution reaches 100% of plan. If either gate is missed, the pool funds at 40% — for everyone. Neither team can win while the other loses, which is the entire mechanism."]]));
body.push(bulletR([["Distribution: ", { bold: true }], ["62% store field (≈₹2,540 average per associate, ≈₹6,340 in top-quartile catchments), 20% digital (≈₹1.45 lakh average), 18% omni ops, CRM and planning."]]));
body.push(bulletR([["Internal transfer price: ", { bold: true }], ["stores receive ₹90 per ship-from-store or click-and-collect order into the store P&L. This is not a bonus, it is payment for a service the store renders to the digital channel, and it should be booked as such."]]));
body.push(bulletR([["Price-parity corridor: ", { bold: true }], ["a hard 5-point cap on the net-of-promotion price gap between store and own digital channels on identical SKU-seasons, plus a manager-authorised price-match wallet capped at 1.2% of store net sales. This attacks the root cause. Marketplace pricing cannot be fully controlled, but Aarav's own 14-point self-inflicted gap can be closed tomorrow."]]));
body.push(bulletR([["Recognition: ", { bold: true }], ["a quarterly Catchment of the Quarter award judged on omni catchment revenue and omni NPS, presented by the CCO. Cheap, and in Indian retail field organisations consistently among the highest-leverage instruments available."]]));

/* ---------- 8. Risks & roadmap ---------- */
body.push(h1("8. Risks and Implementation"));

body.push(exhibitTitle("Exhibit 12 — Principal risks"));
body.push(table([2620, 1160, 1000, 4960],
  ["Risk", "Likelihood", "Impact", "Mitigation"],
  [
    ["ICR gaming — associates harvest invalid numbers", "High", "Medium", "OTP verification only; monthly audit; forfeiture of behaviour component on repeat breach"],
    ["Field rejects the smaller credited base", "High", "High", "90-day parallel run showing both numbers; written no-loss guarantee for FY27; re-based rate announced with the model, not after"],
    ["Marketplace concentration — 46% of digital on one platform", "Medium", "High", "Cap any single platform at 40% of digital revenue by FY28; build D2C and quick commerce as counterweights"],
    ["Quick commerce cannibalises store impulse categories", "Medium", "Medium", "Pilot in 6 cities only; measure catchment-level net effect, not channel-level"],
    ["Price-match wallet abuse", "Medium", "Low", "1.2% cap, manager authorisation, monthly exception report to the CCO"],
    ["DPDP Act 2023 consent gaps in identity resolution", "Medium", "High", "Consent capture built into the associate app before go-live; legal sign-off gates the CDP release"],
    ["Attribution disputes become the new argument", "Medium", "Medium", "Rules frozen for 12 months; changes need CRM sign-off (Exhibit 4) and take effect only at the next financial year"],
  ],
  { align: [AlignmentType.LEFT, AlignmentType.CENTER, AlignmentType.CENTER, AlignmentType.LEFT] }));

body.push(exhibitTitle("Exhibit 13 — Twelve-month sequence"));
body.push(table([1150, 3500, 5090],
  ["Quarter", "Milestone", "Gate to proceed"],
  [
    ["Q1 FY27", "CCO appointed; CDP and attribution engine built; consent capture live; rules published", "ICR above 55%; legal sign-off on DPDP consent"],
    ["Q2 FY27", "Parallel run in 24 stores across 2 zones; Catchment Councils stood up; price-parity corridor enforced", "Attribution variance under 3%; no store loses earnings in the parallel run"],
    ["Q3 FY27", "National rollout of scorecard and re-based variable pay; SFS and C&C live in 60 doors", "ICR above 70%; store-fulfilled share above 12%"],
    ["Q4 FY27", "Cross-channel pool pays out on the dual gate; FY28 budget rebuilt on credited revenue", "Both gates independently audited by Finance"],
  ],
  { align: [AlignmentType.CENTER, AlignmentType.LEFT, AlignmentType.LEFT] }));

body.push(h2("8.1 Sensitivity"));
body.push(p("Three assumptions move the conclusion materially. If the store-influenced share of digital revenue is 25% rather than the 35% assumed, the credited split shifts to 75.6 : 24.4 and the budget split moves roughly two points towards digital. If marketplace effective deduction is 30% rather than 35%, marketplace contribution scoring rises and digital's weighted score gains around 0.4 points. If ICR stalls below 55%, the entire incentive design must be deferred — the model cannot pay on journeys it cannot see, and launching it on thin data would destroy more trust than the current system does."));

/* ---------- 9. Recommendations ---------- */
body.push(h1("9. Recommendations"));

const recs = [
  ["Appoint a Chief Commercial Officer and remove independent revenue targets from Retail and Digital. ", "One P&L above the channels. No unit below the CCO owns a customer. Everything else in this report depends on this one move."],
  ["Adopt the three-ledger credit model. ", "Revenue credit unchanged for the accounts; influence credit for budgeting; fractional incentive credit for pay. Freeze the rules for twelve months."],
  ["Restate FY26 on credited revenue and publish it. ", "₹847 cr store, ₹303 cr digital. Publishing a restatement that costs stores ₹38 cr is what makes the model believable to the digital team."],
  ["Move the FY27 budget to ₹60 cr store, ₹55 cr digital, ₹24 cr ring-fenced omni-shared. ", "Show the digital team the like-for-like reconciliation — their working spend rises 28% — before announcing the headline."],
  ["Re-base store variable pay to 1.254% of credited sales and add catchment and behaviour components. ", "Target earnings held whole in year one, with a written no-loss guarantee for FY27."],
  ["Fund a ₹4.5 cr dual-gated cross-channel pool. ", "Neither team is paid in full unless both clear their gate. This is the instrument that converts a truce into cooperation."],
  ["Close the self-inflicted price gap. ", "A hard 5-point parity corridor on own channels plus a capped store price-match wallet. Aarav cannot control Myntra's discount, but it set its own 14-point gap and can close it immediately."],
  ["Pay stores ₹90 per ship-from-store or click-and-collect order. ", "An internal transfer price, booked into the store P&L. It funds the capability and, more importantly, tells the field that serving a digital customer is paid work."],
  ["Instrument before you incentivise. ", "Identified-customer capture must reach 55% before the parallel run and 70% before national rollout. Without identity there is no attribution, and without attribution this is merely a differently-argued opinion."],
];
recs.forEach((r) => body.push(numbered([[r[0], { bold: true }], [r[1]]], 0, "num2")));

body.push(spacer(160));
body.push(box("What success looks like at FY27 exit", [
  p("Store-influenced digital revenue is visible and paid. Identified capture is above 70%. Cross-channel customers have doubled to 18% of the active base. Nine of ten Omni Scorecard KPIs are instrumented, against six today. And the sentence that opened this report — that an associate in Indore fits a customer for twenty minutes and books nothing — is no longer true.", { after: 0 }),
], T.SAFFRON));

/* ---------- References ---------- */
body.push(h1("References"));
const refs = [
  "Ailawadi, K.L. and Farris, P.W. (2017) 'Managing multi- and omni-channel distribution: metrics and research directions', Journal of Retailing, 93(1), pp. 120–135.",
  "Anthony, R.N. and Govindarajan, V. (2007) Management Control Systems. 12th edn. New York: McGraw-Hill.",
  "Bell, D.R., Gallino, S. and Moreno, A. (2014) 'How to win in an omnichannel world', MIT Sloan Management Review, 56(1), pp. 45–53.",
  "Business Standard (2025) 'From baskets to blink: quick commerce reshapes retail landscape in 2025', Business Standard, 29 December. Available at: https://www.business-standard.com (Accessed: 27 July 2026).",
  "Churchill, G.A., Ford, N.M. and Walker, O.C. (2000) Sales Force Management. 6th edn. Boston: Irwin/McGraw-Hill.",
  "Exchange4media (2025) 'The great D2C reset: how 2025 merged offline and online commerce'. Available at: https://www.exchange4media.com (Accessed: 27 July 2026).",
  "Gallino, S. and Moreno, A. (2014) 'Integration of online and offline channels in retail: the impact of sharing reliable inventory availability information', Management Science, 60(6), pp. 1434–1451.",
  "Gensler, S., Neslin, S.A. and Verhoef, P.C. (2017) 'The showrooming phenomenon: it's more than just about price', Journal of Interactive Marketing, 38, pp. 29–43.",
  "Inc42 (2025) 'D2C fashion brand Snitch to open 10 new offline stores in January'. Available at: https://inc42.com (Accessed: 27 July 2026).",
  "Indian Apparel (n.d.) 'Omni-channel strategy by Shoppers Stop'. Available at: https://www.indian-apparel.com (Accessed: 27 July 2026).",
  "Indian Retailer (2025) 'Myntra brings 30-minute fashion delivery to tier II cities with M-Now'. Available at: https://www.indianretailer.com (Accessed: 27 July 2026).",
  "Kaplan, R.S. and Norton, D.P. (1996) The Balanced Scorecard: Translating Strategy into Action. Boston: Harvard Business School Press.",
  "Redseer Strategy Consultants (2025) India Quick Commerce Year-End Review 2025. Bengaluru: Redseer.",
  "Rigby, D. (2011) 'The future of shopping', Harvard Business Review, 89(12), pp. 65–76.",
  "Shapley, L.S. (1953) 'A value for n-person games', in Kuhn, H.W. and Tucker, A.W. (eds.) Contributions to the Theory of Games II. Princeton: Princeton University Press, pp. 307–317.",
  "Sopadjieva, E., Dholakia, U.M. and Benjamin, B. (2017) 'A study of 46,000 shoppers shows that omnichannel retailing works', Harvard Business Review, 3 January.",
  "Stern, L.W. and El-Ansary, A.I. (1992) Marketing Channels. 4th edn. Englewood Cliffs, NJ: Prentice Hall.",
  "TechCrunch (2024) 'Myntra pushes into India's quick commerce race with 30-minute fashion delivery', 5 December. Available at: https://techcrunch.com (Accessed: 27 July 2026).",
  "Terra Insight (2025) Myntra, Ajio and Flipkart Fashion Apparel Settlement Reconciliation. Available at: https://www.terra-insight.com (Accessed: 27 July 2026).",
  "Trent Limited (2025) Q4 and FY25 Results Press Release: Portfolio of 1,043 Fashion Stores. Mumbai: Trent Limited, 29 April.",
  "Verhoef, P.C., Kannan, P.K. and Inman, J.J. (2015) 'From multi-channel retailing to omni-channel retailing: introduction to the special issue on multi-channel retailing', Journal of Retailing, 91(2), pp. 174–181.",
  "Wazir Advisors (2025) Indian Textile and Apparel Industry: Annual Report 2025. Gurugram: Wazir Advisors.",
  "Wazir Advisors (2025) India Apparel Market Review: Dynamics and Direction 2020–2025. Gurugram: Wazir Advisors.",
  "Zoltners, A.A., Sinha, P. and Lorimer, S.E. (2008) 'Sales force effectiveness: a framework for researchers and practitioners', Journal of Personal Selling & Sales Management, 28(2), pp. 115–131.",
];
refs.forEach(r => body.push(new Paragraph({
  spacing: { after: 62, line: 236 },
  indent: { left: 360, hanging: 360 },
  alignment: AlignmentType.LEFT,
  children: [new TextRun({ text: r, font: BODY, size: 16, color: T.INK })],
})));

/* =======================================================================
   DOCUMENT
   ======================================================================= */
const doc = new Document({
  creator: "Aarav Apparel — Commercial Strategy",
  title: "Aarav Apparel: Ending the Channel War",
  description: "Sales & Distribution Management case analysis",
  numbering: {
    config: [
      { reference: "bul", levels: [
        { level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 360, hanging: 240 } },
                   run: { color: T.SAFFRON, font: BODY } } },
        { level: 1, format: LevelFormat.BULLET, text: "–", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 240 } } } },
      ]},
      { reference: "num", levels: [
        { level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 400, hanging: 400 } },
                   run: { color: T.INDIGO, bold: true, font: BODY } } },
      ]},
      { reference: "num2", levels: [
        { level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 430, hanging: 430 } },
                   run: { color: T.INDIGO, bold: true, font: BODY } } },
      ]},
    ],
  },
  styles: {
    default: {
      document: { run: { font: BODY, size: 20, color: T.INK } },
    },
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838, orientation: PageOrientation.PORTRAIT },
        margin: { top: 1080, right: 1080, bottom: 1080, left: 1080, header: 560, footer: 560 },
      },
    },
    headers: {
      default: new Header({ children: [new Paragraph({
        alignment: AlignmentType.RIGHT,
        spacing: { after: 0 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: T.RULE, space: 4 } },
        children: [new TextRun({ text: "Aarav Apparel  ·  Sales & Distribution Management Case Analysis", font: BODY, size: 15, color: T.GREY })],
      })] }),
    },
    footers: {
      default: new Footer({ children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0 },
        children: [
          new TextRun({ text: "Page ", font: BODY, size: 15, color: T.GREY }),
          new TextRun({ children: [PageNumber.CURRENT], font: BODY, size: 15, color: T.GREY }),
          new TextRun({ text: " of ", font: BODY, size: 15, color: T.GREY }),
          new TextRun({ children: [PageNumber.TOTAL_PAGES], font: BODY, size: 15, color: T.GREY }),
        ],
      })] }),
    },
    children: body,
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync("Aarav_Apparel_Omnichannel_Case_Analysis.docx", buf);
  console.log("written:", buf.length, "bytes");
});
