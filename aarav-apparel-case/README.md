# Aarav Apparel — Omni-Channel Sales & Distribution Case Analysis

Deliverables for the CFO of Aarav Apparel (Indian fashion retail; 120 stores, D2C + marketplace).

| File | What it is |
|---|---|
| `Aarav_Apparel_Omnichannel_Case_Analysis.docx` | 10-page report (+ references) |
| `Aarav_Apparel_Omnichannel_Case_Analysis.pdf` | PDF of the report |
| `Aarav_Apparel_Omnichannel_Executive_Summary.pptx` | 10-slide executive summary for a 10–12 minute talk, with speaker notes |
| `Aarav_Apparel_Omnichannel_Executive_Summary.pdf` | PDF of the deck |

## Regenerating

```bash
npm install                       # docx + pptxgenjs
node build_report.js              # -> .docx
node build_deck.js                # -> .pptx
soffice --headless --convert-to pdf <file>
```

`theme.js` holds the shared palette and type stack (indigo / saffron / teal; Cambria headings,
Calibri body) used by both deliverables.

## Contents

The report answers four questions, each grounded in a stated assumption set:

1. **Organisation** — one commercial P&L above the channels under a new Chief Commercial Officer,
   with explicit reporting lines, dotted lines and a shared-accountability (RACI) matrix.
2. **Attribution** — a three-ledger model (revenue / influence / incentive credit) with
   journey-level fractional split rules and a full FY26 restatement.
3. **Budget** — FY27 split of ₹139 cr across store support, digital growth and a ring-fenced
   omni-shared bucket, from six weighted criteria.
4. **Alignment** — a ten-KPI unified scorecard, rebuilt variable pay, and a dual-gated
   cross-channel incentive pool.

Every figure is an assumption, declared in the Assumptions box (report Section 2, deck slide 3).
Market grounding uses Indian apparel sources (quick commerce, Myntra/Ajio marketplace economics,
D2C offline expansion), cited Harvard-style.
