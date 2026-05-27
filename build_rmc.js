const ExcelJS = require('exceljs');
const path = require('path');

const wb = new ExcelJS.Workbook();
wb.creator = 'RMC Template Generator';

// ── Palette ──────────────────────────────────────────────────────────────────
const C = {
  DARK_BLUE:  'FF1F3864', MED_BLUE:   'FF2E75B6', LIGHT_BLUE: 'FFBDD7EE',
  PALE_BLUE:  'FFDEEAF1', GOLD:       'FFFFD700', LIGHT_GOLD: 'FFFFF2CC',
  GREEN_FILL: 'FF375623', RED:        'FFC00000',  WHITE:      'FFFFFFFF',
  LIGHT_GREY: 'FFF2F2F2', NAVY2:      'FF2F5496',  AMBER:      'FFED7D31',
  GREEN_TAB:  'FF70AD47',
};

function fill(argb) { return { type: 'pattern', pattern: 'solid', fgColor: { argb } }; }

function font(opts = {}) {
  return { name: 'Arial', size: opts.size || 10, bold: !!opts.bold,
           italic: !!opts.italic, color: { argb: opts.color || 'FF000000' } };
}

function border(style = 'thin') {
  const s = { style };
  return { top: s, left: s, bottom: s, right: s };
}

function align(h = 'center', wrap = true) {
  return { horizontal: h, vertical: 'middle', wrapText: wrap };
}

const INR  = '₹#,##0.00;(₹#,##0.00);"-"';
const PCT  = '0.00%;-0.00%;"-"';
const NUM2 = '#,##0.00;(#,##0.00);"-"';

// ══════════════════════════════════════════════════════════════════════════════
//  SHEET 1 – RMC Sheet
// ══════════════════════════════════════════════════════════════════════════════
const ws = wb.addWorksheet('RMC Sheet', { properties: { tabColor: { argb: 'FF2E75B6' } } });
ws.views = [{ state: 'frozen', xSplit: 0, ySplit: 9, activeCell: 'A10' }];

// Column widths (A=1 … Q=17)
const colWidths = [6,16,32,16,18,14,14,14,6,16,12,18,16,12,18,18,30];
colWidths.forEach((w, i) => { ws.getColumn(i + 1).width = w; });

// ── Helper: write merged header row ──────────────────────────────────────────
function mergeRow(ws, row, text, fillArgb, fontSize = 10, bold = true, height = 18) {
  ws.getRow(row).height = height;
  ws.mergeCells(row, 1, row, 17);
  const cell = ws.getCell(row, 1);
  cell.value = text;
  cell.font  = font({ bold, size: fontSize, color: C.WHITE });
  cell.fill  = fill(fillArgb);
  cell.alignment = align('center');
  cell.border = border();
}

// ── Title ─────────────────────────────────────────────────────────────────────
ws.getRow(1).height = 32;
mergeRow(ws, 1, 'RAW MATERIAL COST SHEET (RMC)', C.DARK_BLUE, 16, true, 32);

// Info rows 2-7
const infoRows = [
  [2, 'Product Name / API:  [Product Name / Intermediate]     |     Batch Size:  [X kg / L]     |     Stage:  [Stage No.]', C.MED_BLUE],
  [3, 'Document No.:  RMC-XXX-001     |     Version:  00     |     Effective Date:  DD-MMM-YYYY', C.NAVY2],
  [4, 'Prepared By:  _______________     |     Reviewed By:  _______________     |     Approved By:  _______________', C.MED_BLUE],
  [5, 'BOM Reference:  MFR-XXX-001     |     Currency:  INR (₹)     |     Exchange Rate:  1 USD = ₹83.50', C.NAVY2],
  [6, 'Price Validity Date:  DD-MMM-YYYY     |     Yield Basis:  Theoretical/Actual     |     Recovery Credit Applied:  Yes', C.MED_BLUE],
  [7, 'Assumptions: Yield % from avg 3 pilot batches | Recovery % from plant data | Prices from approved vendor quotes', C.NAVY2],
];
infoRows.forEach(([r, text, clr]) => {
  ws.getRow(r).height = 15;
  ws.mergeCells(r, 1, r, 17);
  const c = ws.getCell(r, 1);
  c.value = text;
  c.font  = font({ size: 9, color: C.WHITE });
  c.fill  = fill(clr);
  c.alignment = align('left');
  c.border = border();
});

ws.getRow(8).height = 5; // spacer

// ── Column headers row 9 ─────────────────────────────────────────────────────
const headers = [
  'S. No.', 'Material Code /\nItem Code',
  'Raw Material Name\n(API/KSM/Intermediate/Solvent/Reagent/Catalyst)',
  'Category', 'Grade /\nSpecification', 'Theoretical Qty\nper Batch',
  'Input Factor /\nYield (%)', 'Actual Qty\nRequired', 'UOM',
  'Unit Rate\n(₹/kg or ₹/L)', 'Source\n(Indig./Imported)',
  'Total Cost\nper Batch (₹)', 'Cost per kg\nof Output (₹)',
  '% Contribution\nto Total RM', 'Basis of\nEstimate',
  'Vendor Approval\nStatus', 'Remarks',
];
ws.getRow(9).height = 40;
headers.forEach((h, i) => {
  const c = ws.getCell(9, i + 1);
  c.value = h;
  c.font  = font({ bold: true, size: 9, color: C.WHITE });
  c.fill  = fill(C.DARK_BLUE);
  c.alignment = align('center');
  c.border = border();
});

// ── Stage data ────────────────────────────────────────────────────────────────
// [sno, code, name, cat, grade, theoQty, yieldPct, uom, unitRate, source, basis, vendorStatus, remark]
const stageData = [
  // Stage 1
  [1,'RM-001','Starting Material A (KSM)','KSM','USP',120,0.98,'kg',45000,'Imported','Commercial','Approved','Primary KSM, imported; USD rate applied'],
  [2,'RM-002','Reagent B – Coupling Agent','Reagent','GMP',85,0.97,'kg',12500,'Indigenous','Commercial','Approved','Local vendor; price valid 6 months'],
  [3,'RM-003','Solvent C – Methanol','Solvent','Technical',500,1.00,'L',65,'Indigenous','Commercial','Approved','Recyclable; recovery credit applied separately'],
  [4,'RM-004','Catalyst D – Pd/C (5%)','Catalyst','GMP',2.5,0.95,'kg',185000,'Imported','Pilot','Approved','Amortised over 5 reuse cycles; cost ÷ 5'],
  [5,'RM-005','Reagent E – Potassium Carbonate','Reagent','Technical',45,1.00,'kg',850,'Indigenous','Commercial','Approved','—'],
  [6,'RM-006','Solvent F – Ethyl Acetate','Solvent','Technical',800,1.00,'L',120,'Indigenous','Commercial','Approved','Recyclable; recovery credit applied separately'],
  // Stage 2
  [7,'RM-007','Intermediate – Stage 1 Output','Intermediate','In-house GMP',95,0.96,'kg',0,'In-house','Commercial','Approved','Cost rolled from Stage 1; see Stage Summary sheet'],
  [8,'RM-008','Reagent G – Acylating Agent','Reagent','GMP',60,0.98,'kg',28000,'Imported','Commercial','Approved','USD rate applied; see exchange rate assumption'],
  [9,'RM-009','Solvent G – Dichloromethane (DCM)','Solvent','Technical',600,1.00,'L',95,'Indigenous','Commercial','Approved','Recyclable; recovery credit applied separately'],
  [10,'RM-010','Reagent H – Hydrochloric Acid 35%','Reagent','Technical',30,1.00,'L',45,'Indigenous','Commercial','Approved','—'],
  // Final Stage
  [11,'RM-011','Intermediate – Stage 2 Output','Intermediate','In-house GMP',80,0.97,'kg',0,'In-house','Commercial','Approved','Cost rolled from Stage 2; see Stage Summary sheet'],
  [12,'RM-012','Solvent H – Isopropyl Alcohol (IPA)','Solvent','Technical',400,1.00,'L',110,'Indigenous','Commercial','Approved','Recyclable; recovery credit applied separately'],
  [13,'RM-013','Processing Aid – Activated Carbon','Processing Aid','Technical',5,1.00,'kg',3200,'Indigenous','Commercial','Approved','—'],
  [14,'RM-014','Purified Water (Processing)','Reagent','USP Purified',200,1.00,'L',15,'Indigenous','Commercial','Approved','Utility cost basis; see Assumption 7'],
];

const creditData = [
  [15,'RC-001','Methanol Recovery Credit','Recovery Credit','—',400,0.85,'L',65,'—','—','—','85% recovery; net 340 L credited'],
  [16,'RC-002','Ethyl Acetate Recovery Credit','Recovery Credit','—',640,0.80,'L',120,'—','—','—','80% recovery; net 512 L credited'],
  [17,'RC-003','DCM Recovery Credit','Recovery Credit','—',480,0.75,'L',95,'—','—','—','75% recovery; net 360 L credited'],
  [18,'RC-004','IPA Recovery Credit','Recovery Credit','—',320,0.82,'L',110,'—','—','—','82% recovery; net 262 L credited'],
];

// Row layout:
// 10: Stage 1 header, 11-16: stage 1 data
// 17: Stage 2 header, 18-21: stage 2 data
// 22: Final stage header, 23-26: final data
// 27: spacer, 28: credits header, 29-32: credits
// 33: spacer, 34: gross total, 35: recovery total, 36: spacer, 37: net total

const stage1Rows = [11,12,13,14,15,16];
const stage2Rows = [18,19,20,21];
const finalRows  = [23,24,25,26];
const creditRows = [29,30,31,32];

// GROSS COST total row = 34, CREDITS total = 35, NET = 37
const GROSS_ROW = 34, CRED_ROW = 35, NET_ROW = 37;
// Batch output is at F46 (yield summary)
const BATCH_OUT = 'F46';

function writeSubHeader(ws, row, label, fillArgb = C.MED_BLUE) {
  ws.getRow(row).height = 18;
  ws.mergeCells(row, 1, row, 17);
  const c = ws.getCell(row, 1);
  c.value = '  ' + label;
  c.font  = font({ bold: true, size: 10, color: C.WHITE });
  c.fill  = fill(fillArgb);
  c.alignment = align('left');
  c.border = border();
}

writeSubHeader(ws, 10,  'STAGE 1 – Starting Material Conversion');
writeSubHeader(ws, 17,  'STAGE 2 – Intermediate Synthesis');
writeSubHeader(ws, 22,  'FINAL STAGE – API Crystallisation & Purification');
writeSubHeader(ws, 28,  'SOLVENT RECOVERY CREDITS  (deducted from Gross RM Cost)', C.NAVY2);

const allDataRows = [...stage1Rows, ...stage2Rows, ...finalRows];

function writeDataRow(ws, row, d, isCredit = false) {
  const [sno, code, name, cat, grade, theoQty, yieldPct, uom, unitRate, source, basis, vstatus, remark] = d;
  ws.getRow(row).height = 18;
  const altFill = row % 2 === 0 ? C.LIGHT_BLUE : C.PALE_BLUE;
  const rowFill = isCredit ? 'FFFFE4E1' : altFill;
  const fColor  = isCredit ? C.RED : 'FF000000';
  const italic  = isCredit;

  const vals = [sno, code, name, cat, grade, theoQty, yieldPct, null, uom, unitRate, source, null, null, null, basis, vstatus, remark];
  vals.forEach((v, i) => {
    const col = i + 1;
    const c = ws.getCell(row, col);
    if (v !== null) c.value = v;
    c.fill   = fill(rowFill);
    c.border = border();
    c.font   = font({ size: 9, color: fColor, italic });
    c.alignment = (col === 3 || col === 17) ? align('left') : align('center');
  });

  // Col 7 – Yield % format
  ws.getCell(row, 7).numFmt = PCT;

  // Col 6 – Theo qty format
  ws.getCell(row, 6).numFmt = NUM2;

  // Col 10 – Unit Rate format
  ws.getCell(row, 10).numFmt = INR;

  // Col 8 – Actual Qty formula
  const h = ws.getCell(row, 8);
  if (isCredit) {
    h.value = { formula: `F${row}*G${row}`, result: theoQty * yieldPct };
  } else {
    h.value = { formula: `IFERROR(F${row}/G${row},0)`, result: theoQty / (yieldPct || 1) };
  }
  h.numFmt = NUM2; h.fill = fill(rowFill); h.border = border();
  h.font = font({ size: 9, color: fColor, italic });
  h.alignment = align('center');

  // Col 12 – Total Cost
  const l = ws.getCell(row, 12);
  if (isCredit) {
    l.value = { formula: `-H${row}*J${row}`, result: -(theoQty * yieldPct * unitRate) };
  } else {
    l.value = { formula: `H${row}*J${row}`, result: (theoQty / (yieldPct || 1)) * unitRate };
  }
  l.numFmt = INR; l.fill = fill(rowFill); l.border = border();
  l.font = font({ size: 9, color: fColor, italic });
  l.alignment = align('center');

  // Col 13 – Cost per kg = L / batch output
  const m = ws.getCell(row, 13);
  m.value = { formula: `IFERROR(L${row}/${BATCH_OUT},0)` };
  m.numFmt = INR; m.fill = fill(rowFill); m.border = border();
  m.font = font({ size: 9, color: fColor, italic });
  m.alignment = align('center');

  // Col 14 – % contribution = L / Gross total
  const n = ws.getCell(row, 14);
  n.value = { formula: `IFERROR(L${row}/L${GROSS_ROW},0)` };
  n.numFmt = PCT; n.fill = fill(rowFill); n.border = border();
  n.font = font({ size: 9, color: fColor, italic });
  n.alignment = align('center');
}

stageData.forEach((d, i) => writeDataRow(ws, allDataRows[i], d, false));
creditData.forEach((d, i) => writeDataRow(ws, creditRows[i], d, true));

// ── Totals ────────────────────────────────────────────────────────────────────
ws.getRow(33).height = 5; // spacer

function writeTotalsRow(ws, row, label, formula, fillArgb, valFillArgb, fontColor = 'FF000000', height = 20) {
  ws.getRow(row).height = height;
  ws.mergeCells(row, 1, row, 11);
  const lc = ws.getCell(row, 1);
  lc.value = label;
  lc.font  = font({ bold: true, size: 10, color: fontColor });
  lc.fill  = fill(fillArgb);
  lc.border = border();
  lc.alignment = align('right');
  for (let c = 1; c <= 11; c++) {
    const cell = ws.getCell(row, c);
    cell.fill = fill(fillArgb);
    cell.border = border();
  }
  const vc = ws.getCell(row, 12);
  vc.value = { formula };
  vc.numFmt = INR;
  vc.font  = font({ bold: true, size: 10, color: fontColor });
  vc.fill  = fill(valFillArgb);
  vc.border = border();
  vc.alignment = align('center');
  for (let c = 13; c <= 17; c++) {
    ws.getCell(row, c).fill = fill(valFillArgb);
    ws.getCell(row, c).border = border();
  }
}

const grossParts = allDataRows.map(r => `L${r}`).join('+');
const credParts  = creditRows.map(r => `L${r}`).join('+');

writeTotalsRow(ws, GROSS_ROW,
  'GROSS RAW MATERIAL COST  (before recovery credits)',
  grossParts, C.LIGHT_BLUE, C.LIGHT_BLUE);

writeTotalsRow(ws, CRED_ROW,
  'TOTAL SOLVENT RECOVERY CREDITS  (negative = credit)',
  credParts, 'FFFFE4E1', 'FFFFE4E1', C.RED);

ws.getRow(36).height = 5;

writeTotalsRow(ws, NET_ROW,
  'NET RAW MATERIAL COST  (Gross + Credits)',
  `L${GROSS_ROW}+L${CRED_ROW}`,
  C.GREEN_FILL, C.GREEN_FILL, C.WHITE, 24);

// ── Yield Summary (rows 39-46) ────────────────────────────────────────────────
ws.getRow(38).height = 8;
writeSubHeader(ws, 39, 'YIELD SUMMARY');

const yieldItems = [
  [40, 'Theoretical Batch Output (kg API)',       75,     NUM2,  false],
  [41, 'Stage 1 Yield (%)',                        0.96,   PCT,   true],
  [42, 'Stage 2 Yield (%)',                        0.94,   PCT,   true],
  [43, 'Final Stage Yield (%)',                    0.95,   PCT,   true],
  [44, 'Overall Yield (%) = S1 × S2 × Final',     null,   PCT,   true],
  [45, 'Actual Expected Output (kg)',              null,   NUM2,  false],
];

yieldItems.forEach(([row, label, val, fmt, isPct]) => {
  ws.getRow(row).height = 18;
  ws.mergeCells(row, 1, row, 9);
  const lc = ws.getCell(row, 1);
  lc.value = label;
  lc.font  = font({ bold: true, size: 9 });
  lc.fill  = fill(C.PALE_BLUE);
  lc.border = border();
  lc.alignment = align('left');
  for (let c = 1; c <= 9; c++) {
    ws.getCell(row, c).fill = fill(C.PALE_BLUE);
    ws.getCell(row, c).border = border();
  }

  // Value cell at col F (6) – store raw values for formula references
  const fc = ws.getCell(row, 6);
  if (row === 44) {
    fc.value = { formula: 'F41*F42*F43' };
  } else if (row === 45) {
    fc.value = { formula: 'F40*F44' };
  } else {
    fc.value = val;
  }
  fc.numFmt = fmt;
  fc.font = font({ size: 9, color: 'FF0000FF' }); // blue = input
  fc.fill  = fill(C.LIGHT_GOLD);
  fc.border = border();
  fc.alignment = align('center');

  ws.mergeCells(row, 10, row, 17);
  for (let c = 10; c <= 17; c++) {
    ws.getCell(row, c).fill = fill(C.LIGHT_GOLD);
    ws.getCell(row, c).border = border();
  }
  // BATCH_OUT = F46 which is F45 in 0-indexed… let me check: F46 should be actual output
  // Actually row 45 has actual output at column F → that's F45
  // But BATCH_OUT was set to F46 in formula strings above. Let me add F46 as a link row.
});

// F46 helper – actual expected output (linked from F45)
ws.getCell(46, 6).value = { formula: 'F45' };
ws.getCell(46, 6).numFmt = NUM2;
ws.getRow(46).height = 0; // hidden

// ── Cost Summary (rows 48-57) ─────────────────────────────────────────────────
ws.getRow(47).height = 8;
writeSubHeader(ws, 48, 'COST SUMMARY');

const summaryItems = [
  [49, 'Total Gross RM Cost (₹)',                 `L${GROSS_ROW}`, INR],
  [50, 'Total Solvent Recovery Credits (₹)',      `L${CRED_ROW}`,  INR],
  [51, 'Net Raw Material Cost (₹)',               `L${NET_ROW}`,   INR],
  [52, 'Batch Output – Actual Expected (kg API)', 'F45',           NUM2],
  [53, 'RM Cost per kg of API (₹/kg)',            'IFERROR(L51/L52,0)', INR],
];

summaryItems.forEach(([row, label, formula, fmt]) => {
  ws.getRow(row).height = 18;
  ws.mergeCells(row, 1, row, 9);
  for (let c = 1; c <= 17; c++) { ws.getCell(row, c).fill = fill(C.LIGHT_GOLD); ws.getCell(row, c).border = border(); }
  const lc = ws.getCell(row, 1);
  lc.value = label;
  lc.font  = font({ bold: true, size: 9 });
  lc.alignment = align('left');

  const vc = ws.getCell(row, 10);
  vc.value = { formula };
  vc.numFmt = fmt;
  vc.font = font({ bold: true, size: 9 });
  vc.fill = fill(C.GOLD);
  vc.border = border();
  vc.alignment = align('center');
  ws.mergeCells(row, 10, row, 17);
});

// Top 3 cost drivers
ws.getRow(54).height = 18;
writeSubHeader(ws, 54, '  TOP 3 COST DRIVERS  (by % contribution)');
const top3 = [
  [55, '1st', 'Starting Material A (KSM) – RM-001',        '~55-60%'],
  [56, '2nd', 'Catalyst D – Pd/C (5%) – RM-004',          '~15-20%'],
  [57, '3rd', 'Reagent G – Acylating Agent – RM-008',      '~10-12%'],
];
top3.forEach(([row, rank, name, pct]) => {
  ws.getRow(row).height = 16;
  ws.mergeCells(row, 1, row, 3);
  const rc = ws.getCell(row, 1); rc.value = rank; rc.font = font({ bold: true, size: 9 });
  rc.fill = fill(C.LIGHT_GOLD); rc.border = border(); rc.alignment = align('center');
  ws.mergeCells(row, 4, row, 12);
  const nc = ws.getCell(row, 4); nc.value = name; nc.font = font({ size: 9 });
  nc.fill = fill(C.LIGHT_GOLD); nc.border = border(); nc.alignment = align('left');
  ws.mergeCells(row, 13, row, 17);
  const pc = ws.getCell(row, 13); pc.value = pct; pc.font = font({ bold: true, size: 9 });
  pc.fill = fill(C.GOLD); pc.border = border(); pc.alignment = align('center');
});

// ── Assumptions (rows 59-70) ───────────────────────────────────────────────────
ws.getRow(58).height = 8;
writeSubHeader(ws, 59, 'ASSUMPTIONS & NOTES');

const assumptions = [
  '1.  Yield % values based on average of 3 consecutive pilot batches. Update with commercial batch data when available.',
  '2.  Solvent recovery %: Methanol 85%, Ethyl Acetate 80%, DCM 75%, IPA 82% — based on plant solvent recovery system data.',
  '3.  USD / INR exchange rate: 1 USD = ₹83.50 (valid as of DD-MMM-YYYY). Update for each cost revision cycle.',
  '4.  Unit rates sourced from approved vendor quotations; price validity: 6 months from effective date. Review on renewal.',
  '5.  Catalyst Pd/C reuse cycle = 5 batches; cost per batch = total procurement cost ÷ 5. Confirm with process team.',
  '6.  Intermediate costs from Stage 1 and Stage 2 are rolled into subsequent stages as input material cost.',
  '7.  Purified water cost of ₹15/L represents utility / processing cost allocated by engineering. Not a procurement cost.',
  '8.  All costs are in INR (₹). Imported material costs converted using exchange rate stated in Assumption 3.',
  '9.  Theoretical batch output = 75 kg API. Update with validated batch record data upon commercial manufacture.',
  '10. "Basis of Estimate" column: Commercial = vendor quotes; Pilot = pilot batch data; Lab = lab scale extrapolation.',
];
assumptions.forEach((text, i) => {
  const row = 60 + i;
  ws.getRow(row).height = 15;
  ws.mergeCells(row, 1, row, 17);
  const c = ws.getCell(row, 1);
  c.value = text;
  c.font  = font({ size: 9, italic: false });
  c.fill  = fill(i % 2 === 0 ? C.LIGHT_GREY : C.WHITE);
  c.border = border();
  c.alignment = align('left');
});

// ══════════════════════════════════════════════════════════════════════════════
//  SHEET 2 – Variance Analysis
// ══════════════════════════════════════════════════════════════════════════════
const ws2 = wb.addWorksheet('Variance Analysis', { properties: { tabColor: { argb: C.AMBER } } });
ws2.views = [{ state: 'frozen', xSplit: 0, ySplit: 3, activeCell: 'A4' }];
[6,28,14,14,14,14,14,14,16,16,16,12,30].forEach((w, i) => { ws2.getColumn(i + 1).width = w; });

ws2.getRow(1).height = 28;
ws2.mergeCells(1, 1, 1, 13);
const t2 = ws2.getCell(1, 1);
t2.value = 'RAW MATERIAL COST – STANDARD vs ACTUAL VARIANCE ANALYSIS';
t2.font  = font({ bold: true, size: 14, color: C.WHITE });
t2.fill  = fill(C.DARK_BLUE); t2.alignment = align('center'); t2.border = border();

ws2.getRow(2).height = 14;
ws2.mergeCells(2, 1, 2, 13);
const t2b = ws2.getCell(2, 1);
t2b.value = 'Product: [Product Name]     |     Batch No.: [Batch No.]     |     Period: [Month / Quarter / Year]';
t2b.font  = font({ size: 9, color: C.WHITE }); t2b.fill = fill(C.MED_BLUE);
t2b.alignment = align('center'); t2b.border = border();

const vaHdrs = ['S.No.','Material Name','Std Qty\n(kg/L)','Actual Qty\n(kg/L)',
  'Qty Variance\n(kg/L)','Std Rate\n(₹)','Actual Rate\n(₹)','Rate Variance\n(₹)',
  'Std Cost\n(₹)','Actual Cost\n(₹)','Total Variance\n(₹)','Total Var\n(%)','Remarks'];
ws2.getRow(3).height = 40;
vaHdrs.forEach((h, i) => {
  const c = ws2.getCell(3, i + 1);
  c.value = h; c.font = font({ bold: true, size: 9, color: C.WHITE });
  c.fill = fill(C.DARK_BLUE); c.alignment = align('center'); c.border = border();
});

const vaData = [
  [1,'Starting Material A (KSM)',         122.45, 119.80, 45000, 46200, 'Favourable qty; adverse rate – supplier price increase'],
  [2,'Reagent B – Coupling Agent',         87.63,  89.10,  12500, 12500, 'Adverse qty – higher in-process loss this batch'],
  [3,'Catalyst D – Pd/C (5%)',              2.63,   2.63, 185000,192500, 'Adverse rate – market shortage; price review initiated'],
  [4,'Solvent C – Methanol (net credit)',  510.20, 498.00,    65,    65, 'Favourable qty – better yield; rate unchanged'],
  [5,'Reagent G – Acylating Agent',        61.22,  62.50,  28000, 27500, 'Adverse qty, favourable rate – negotiated discount'],
];

vaData.forEach((d, i) => {
  const row = 4 + i;
  ws2.getRow(row).height = 18;
  const [sno, name, sq, aq, sr, ar, rmk] = d;
  const rowFill = i % 2 === 0 ? C.LIGHT_BLUE : C.PALE_BLUE;

  const rowVals = [
    { v: sno },
    { v: name, al: 'left' },
    { v: sq,  fmt: NUM2 },
    { v: aq,  fmt: NUM2 },
    { formula: `D${row}-C${row}`, fmt: NUM2 },
    { v: sr,  fmt: INR },
    { v: ar,  fmt: INR },
    { formula: `G${row}-F${row}`, fmt: INR },
    { formula: `C${row}*F${row}`, fmt: INR },
    { formula: `D${row}*G${row}`, fmt: INR },
    { formula: `J${row}-I${row}`, fmt: INR },
    { formula: `IFERROR(K${row}/I${row},0)`, fmt: PCT },
    { v: rmk, al: 'left' },
  ];

  rowVals.forEach((rv, ci) => {
    const c = ws2.getCell(row, ci + 1);
    if (rv.formula) c.value = { formula: rv.formula };
    else if (rv.v !== undefined) c.value = rv.v;
    c.fill = fill(rowFill); c.border = border();
    c.font = font({ size: 9 });
    c.alignment = align(rv.al || 'center');
    if (rv.fmt) c.numFmt = rv.fmt;
  });
});

// Totals row
const totRow = 9;
ws2.getRow(totRow).height = 20;
ws2.mergeCells(totRow, 1, totRow, 8);
for (let c = 1; c <= 13; c++) {
  ws2.getCell(totRow, c).fill = fill(C.DARK_BLUE);
  ws2.getCell(totRow, c).border = border();
  ws2.getCell(totRow, c).font = font({ bold: true, color: C.WHITE, size: 10 });
}
const totLabel = ws2.getCell(totRow, 1);
totLabel.value = 'TOTAL'; totLabel.alignment = align('right');

[[9,'I'],[10,'J'],[11,'K']].forEach(([ci, col]) => {
  const c = ws2.getCell(totRow, ci);
  c.value = { formula: `SUM(${col}4:${col}8)` };
  c.numFmt = INR; c.alignment = align('center');
});
const tv = ws2.getCell(totRow, 12);
tv.value = { formula: `IFERROR(K${totRow}/I${totRow},0)` };
tv.numFmt = PCT; tv.alignment = align('center');

ws2.getRow(11).height = 14;
ws2.mergeCells(11, 1, 11, 13);
const leg = ws2.getCell(11, 1);
leg.value = 'LEGEND:  Favourable Variance = Actual Cost < Standard Cost  |  Adverse Variance = Actual Cost > Standard Cost  |  Blue text = Hardcoded inputs';
leg.font = font({ size: 8, italic: true }); leg.fill = fill(C.LIGHT_GREY);
leg.border = border(); leg.alignment = align('left');

// ══════════════════════════════════════════════════════════════════════════════
//  SHEET 3 – Stage Summary
// ══════════════════════════════════════════════════════════════════════════════
const ws3 = wb.addWorksheet('Stage Summary', { properties: { tabColor: { argb: C.GREEN_TAB } } });
ws3.views = [{ state: 'frozen', xSplit: 0, ySplit: 3, activeCell: 'A4' }];
[22,32,18,18,18,12,12,12,18].forEach((w, i) => { ws3.getColumn(i + 1).width = w; });

ws3.getRow(1).height = 28;
ws3.mergeCells(1, 1, 1, 9);
const t3 = ws3.getCell(1, 1);
t3.value = 'STAGE-WISE RAW MATERIAL COST SUMMARY';
t3.font  = font({ bold: true, size: 14, color: C.WHITE });
t3.fill  = fill(C.DARK_BLUE); t3.alignment = align('center'); t3.border = border();

ws3.getRow(2).height = 14;
ws3.mergeCells(2, 1, 2, 9);
const t3b = ws3.getCell(2, 1);
t3b.value = 'Product: [Product Name]     |     BOM Ref: MFR-XXX-001     |     Date: DD-MMM-YYYY';
t3b.font = font({ size: 9, color: C.WHITE }); t3b.fill = fill(C.MED_BLUE);
t3b.alignment = align('center'); t3b.border = border();

const ssHdrs = ['Stage','Key Raw Materials\n(Main Inputs)','Gross RM Cost\n(₹)',
  'Recovery Credits\n(₹)','Net RM Cost\n(₹)','Theoretical\nYield (%)','Actual\nYield (%)',
  'Output Qty\n(kg)','Cost per kg\nof Output (₹/kg)'];
ws3.getRow(3).height = 40;
ssHdrs.forEach((h, i) => {
  const c = ws3.getCell(3, i + 1);
  c.value = h; c.font = font({ bold: true, size: 9, color: C.WHITE });
  c.fill = fill(C.DARK_BLUE); c.alignment = align('center'); c.border = border();
});

const ssData = [
  ['Stage 1\n(KSM Conversion)',   'Starting Material A, Reagent B,\nMethanol, Pd/C, K₂CO₃, Ethyl Acetate',    582500, -47125, 0.96, 0.95, 96.0],
  ['Stage 2\n(Intermediate)',     'Stage 1 Intermediate, Acylating Agent,\nDCM, HCl 35%',                       324800, -61440, 0.94, 0.93, 80.0],
  ['Final Stage\n(API Crystall.)', 'Stage 2 Intermediate, IPA,\nActivated Carbon, Purified Water',              118400, -35288, 0.95, 0.94, 68.0],
];

ssData.forEach((d, i) => {
  const row = 4 + i;
  ws3.getRow(row).height = 36;
  const [stage, mats, gross, cred, thYld, acYld, outQty] = d;
  const rowFill = i % 2 === 0 ? C.LIGHT_BLUE : C.PALE_BLUE;
  const net = gross + cred;

  const vals = [
    { v: stage },
    { v: mats, al: 'left' },
    { v: gross,  fmt: INR,  blue: true },
    { v: cred,   fmt: INR,  blue: true },
    { v: net,    fmt: INR,  blue: true },
    { v: thYld,  fmt: PCT,  blue: true },
    { v: acYld,  fmt: PCT,  blue: true },
    { v: outQty, fmt: NUM2, blue: true },
    { formula: `IFERROR(E${row}/H${row},0)`, fmt: INR },
  ];

  vals.forEach((rv, ci) => {
    const c = ws3.getCell(row, ci + 1);
    if (rv.formula) c.value = { formula: rv.formula };
    else c.value = rv.v;
    c.fill = fill(rowFill); c.border = border();
    c.font = font({ size: 9, color: rv.blue ? 'FF0000FF' : 'FF000000' });
    c.alignment = (ci === 1) ? align('left') : align('center');
    if (rv.fmt) c.numFmt = rv.fmt;
  });
});

// Totals
const t3Row = 7;
ws3.getRow(t3Row).height = 22;
const t3Vals = [
  { v: 'TOTAL / FINAL API' },
  { v: 'All Stages Combined', al: 'left' },
  { formula: 'SUM(C4:C6)', fmt: INR },
  { formula: 'SUM(D4:D6)', fmt: INR },
  { formula: 'SUM(E4:E6)', fmt: INR },
  { formula: 'F4*F5*F6',   fmt: PCT },
  { formula: 'G4*G5*G6',   fmt: PCT },
  { v: 68.0,               fmt: NUM2 },
  { formula: `IFERROR(E${t3Row}/H${t3Row},0)`, fmt: INR },
];
t3Vals.forEach((rv, ci) => {
  const c = ws3.getCell(t3Row, ci + 1);
  if (rv.formula) c.value = { formula: rv.formula };
  else if (rv.v !== undefined) c.value = rv.v;
  c.fill = fill(C.GREEN_FILL); c.border = border();
  c.font = font({ bold: true, size: 10, color: C.WHITE });
  c.alignment = (ci === 1) ? align('left') : align('center');
  if (rv.fmt) c.numFmt = rv.fmt;
});

ws3.getRow(9).height = 15;
ws3.mergeCells(9, 1, 9, 9);
const note = ws3.getCell(9, 1);
note.value = 'NOTE: Stage costs above are illustrative split estimates. Populate from validated batch records. Recovery credits proportionally allocated per stage. Final API cost/kg rolls up all three stages.';
note.font = font({ size: 8, italic: true });
note.fill = fill(C.LIGHT_GREY); note.border = border(); note.alignment = align('left');

// ── Save ──────────────────────────────────────────────────────────────────────
const outPath = path.join(__dirname, 'RMC_Template.xlsx');
wb.xlsx.writeFile(outPath)
  .then(() => console.log('✓ Saved:', outPath))
  .catch(err => { console.error('Error:', err); process.exit(1); });
