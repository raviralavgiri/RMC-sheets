# RMC-sheets

A pharmaceutical **Raw Material Cost (RMC) Sheet** generator for API / Intermediate products. Produces a fully-formatted, formula-driven Excel workbook from a Bill of Materials (BOM) / Master Formula Record.

---

## What's inside

| File | Purpose |
|---|---|
| `RMC_Template.xlsx` | Ready-to-use Excel workbook (3 sheets) |
| `build_rmc.js` | Node.js script that regenerates the workbook |
| `package.json` | Dependencies (`exceljs`) |

---

## Excel workbook – sheet overview

### Sheet 1 · RMC Sheet
The main cost sheet. Covers a sample 3-stage API synthesis.

**Columns**

| # | Column | Description |
|---|---|---|
| A | S. No. | Serial number |
| B | Material Code | Item code from BOM / SAP |
| C | Raw Material Name | API, KSM, Intermediate, Solvent, Reagent, Catalyst, Processing Aid |
| D | Category | Dropdown classification |
| E | Grade / Specification | IP / BP / USP / EP / Technical / GMP |
| F | Theoretical Qty per Batch | From MFR (kg or L) |
| G | Input Factor / Yield (%) | In-process yield or loss factor |
| H | Actual Qty Required | `= F / G` (formula) |
| I | UOM | kg / L / g / mL |
| J | Unit Rate (₹) | From approved vendor quotation |
| K | Source | Indigenous / Imported / In-house |
| L | Total Cost per Batch (₹) | `= H × J` (formula) |
| M | Cost per kg of Output (₹) | `= L / Batch Output` (formula) |
| N | % Contribution to Total RM | `= L / Gross RM Cost` (formula) |
| O | Basis of Estimate | Commercial / Pilot / Lab |
| P | Vendor Approval Status | Approved / Provisional / Alternate |
| Q | Remarks | Recovery notes, DMF ref, alternate vendor |

**Sections**

- **Stage 1** – Starting material conversion (6 line items)
- **Stage 2** – Intermediate synthesis (4 line items)
- **Final Stage** – API crystallisation & purification (4 line items)
- **Solvent Recovery Credits** – 4 solvents credited at plant recovery efficiency (negative cost, deducted from gross)
- **Gross RM Cost → Recovery Credits → Net RM Cost** totals
- **Yield Summary** – Stage-wise and overall yield chain (`= S1 × S2 × Final`)
- **Cost Summary** – Net RM cost, batch output, RM cost per kg of API
- **Top 3 Cost Drivers**
- **Assumptions & Notes** – 10 documented assumptions (exchange rate, yield basis, price validity, etc.)

### Sheet 2 · Variance Analysis
Standard Cost vs Actual Cost comparison for a production batch.

Columns: Std Qty | Actual Qty | Qty Variance | Std Rate | Actual Rate | Rate Variance | Std Cost | Actual Cost | Total Variance (₹) | Total Variance (%)

### Sheet 3 · Stage Summary
Stage-wise rollup showing gross cost, recovery credits, net cost, theoretical yield, actual yield, output qty, and cost per kg of intermediate — with a final API totals row.

---

## How to use the template

1. Open `RMC_Template.xlsx` in Excel.
2. Fill in the **header block** (rows 2–7): product name, batch size, document number, approval fields.
3. Update **blue cells** (hardcoded inputs): theoretical quantities, yield %, unit rates, source.
4. All formula cells (black text) recalculate automatically — Actual Qty, Total Cost, % Contribution, Cost/kg.
5. Update the **Solvent Recovery Credits** section with actual plant recovery factors.
6. Review the **Assumptions** section (rows 60–69) and update exchange rate, yield basis, and price validity date.

### Color convention

| Color | Meaning |
|---|---|
| Blue text | Hardcoded input — update for each product / batch |
| Black text | Formula — do not overwrite |
| Red italic | Recovery credit (negative cost) |
| Green row | Net RM Cost (key output) |
| Gold box | Summary / Cost/kg output |

---

## Regenerating the workbook

Requires Node.js ≥ 18.

```bash
npm install
node build_rmc.js
```

This overwrites `RMC_Template.xlsx` with a fresh copy. Edit `build_rmc.js` to change materials, rates, or layout.

---

## Adapting for your product

| Use case | What to change in `build_rmc.js` |
|---|---|
| Different number of stages | Add/remove `stageData` entries and row layout constants |
| Additional materials | Append rows to `stageData` / `creditData` arrays |
| Multi-currency | Add a currency column and exchange-rate lookup row |
| Regulatory / audit columns | Add DMF reference, COA reference, vendor approval columns |
| Variance analysis | Populate `vaData` in Sheet 2 with actual batch figures |

---

## Assumptions in the template (sample values)

1. Yield % based on average of 3 pilot batches
2. Solvent recovery: Methanol 85%, Ethyl Acetate 80%, DCM 75%, IPA 82%
3. Exchange rate: 1 USD = ₹83.50 (update each revision cycle)
4. Unit rates from approved vendor quotations; validity 6 months
5. Catalyst Pd/C amortised over 5 reuse cycles
6. Intermediate costs rolled stage-to-stage
7. Purified water at ₹15/L (utility cost basis)
8. Theoretical batch output: 75 kg API

---

## Dependencies

- [exceljs](https://github.com/exceljs/exceljs) — Excel workbook generation with full formatting and formula support
