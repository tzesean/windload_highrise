import { Workbook } from './evaluator.js';
import { WORKBOOK_BUNDLE } from './workbook-data.js';

const wb = new Workbook(WORKBOOK_BUNDLE);

const OUTPUT_NAMES = {
  Mz_cat_y: 'Mz_cat_y', Mh_y: 'Mh_y', Vsit_y: 'Vsit_y', Vdes_y: 'Vdes_y',
  Cdyn_y_SLS: 'Cdyn_y_SLS', Wy_W_SLS: 'Wy_W_SLS', Wy_L_SLS: 'Wy_L_SLS', Wy_SLS: 'Wy_SLS',
  a_max_y: 'a_max_y', Cdyn_y_ULS: 'Cdyn_y_ULS', Wy_W_ULS: 'Wy_W_ULS', Wy_L_ULS: 'Wy_L_ULS', Wy_ULS: 'Wy_ULS',
  Mz_cat_x: 'Mz_cat_x', Mh_x: 'Mh_x', Vsit_x: 'Vsit_x', Vdes_x: 'Vdes_x',
  Cdyn_x_SLS: 'Cdyn_x_SLS', Wx_W_SLS: 'Wx_W_SLS', Wx_L_SLS: 'Wx_L_SLS', Wx_SLS: 'Wx_SLS',
  a_max_x: 'a_max_x', Cdyn_x_ULS: 'Cdyn_x_ULS', Wx_W_ULS: 'Wx_W_ULS', Wx_L_ULS: 'Wx_L_ULS', Wx_ULS: 'Wx_ULS',
};

const DEFAULT_ROWS = [
  ["L1","3.00","50","50","5000"],
  ["L2","3.00","50","50","5000"],
  ["L3","3.00","50","50","5000"],
  ["L4","3.00","50","50","5000"],
  ["L5","3.00","50","50","5000"],
  ["L6","3.00","50","50","5000"],
  ["L7","3.00","50","50","5000"],
  ["L8","3.00","50","50","5000"],
  ["L9","3.00","50","50","5000"],
  ["L10","3.00","50","50","5000"],
];

export const state = {
  jobNo: '0000', jobTitle: 'None', designedBy: 'Designer', checkedBy: 'Checker', date: '01/07/2026',
  alpha: 0,
  structureCategory: 2, terrainCategory: 3, vbasic: 33.5,
  cdyn: 'Considered',
  naY: 0.2, naX: 0.2, zetaSLS: 0.005, zetaULS: 0.05,
  cpiPos: 0.6, cpiNeg: -0.3,
  Ka: 1.0, Kc: 1.0, Kl: 1.0, Kp: 1.0,
  msMode: 'Default', hs: 1, bs: 4, ns: 3,
  hillType: 'None', H: 30, Lu: 34, x: 48,
  rows: DEFAULT_ROWS.map(r => ({
    story: r[0], floorHt: parseFloat(r[1]), ly: parseFloat(r[2]), lx: parseFloat(r[3]), deadLoad: parseFloat(r[4]),
    results: null,
  })),
  hasRun: false,
};

function computeZSequence(rows) {
  const n = rows.length;
  const z = new Array(n);
  z[n - 1] = rows[n - 1].floorHt;
  for (let i = n - 2; i >= 0; i--) z[i] = z[i + 1] + rows[i].floorHt;
  return z;
}

function ensureSummaryRowFormulas(startRow, endRow) {
  for (let row = startRow; row <= endRow; row++) {
    const S = WORKBOOK_BUNDLE.sheets.Summary;
    if (!S[`M${row}`]) S[`M${row}`] = `=L${row}*H${row}*(C${row}+OFFSET(C${row},-1,0))/2`;
    if (!S[`S${row}`]) S[`S${row}`] = `=R${row}*H${row}*(C${row}+OFFSET(C${row},-1,0))/2`;
    if (!S[`AC${row}`]) S[`AC${row}`] = `=AB${row}*X${row}*(C${row}+OFFSET(C${row},-1,0))/2`;
    if (!S[`AI${row}`]) S[`AI${row}`] = `=AH${row}*X${row}*(C${row}+OFFSET(C${row},-1,0))/2`;
  }
}

export function runAll() {
  wb.clearOverrides();
  wb.reset();

  // ---- Global parameters (apply once, persist across all rows) ----
  wb.setOverride('Summary', 'J7', `Category ${state.structureCategory}`);
  wb.setOverride('Summary', 'J8', `Category ${state.terrainCategory}`);
  wb.setOverride('Summary', 'M4', state.vbasic);

  const hillStr = state.hillType === 'Escarpment' ? 'Escarpements'
    : state.hillType === 'None' ? 'None' : 'Hills and Ridges';
  for (const sheet of ['Wy', 'Wx']) {
    wb.setOverride(sheet, 'AV6', hillStr);
    wb.setOverride(sheet, 'AL17', state.H); wb.setOverride(sheet, 'AL19', state.Lu); wb.setOverride(sheet, 'AL21', state.x);
    wb.setOverride(sheet, 'AL36', state.H); wb.setOverride(sheet, 'AL38', state.Lu); wb.setOverride(sheet, 'AL40', state.x);
  }

  wb.setOverride('Summary', 'BA1', state.Ka);
  wb.setOverride('Summary', 'BA2', state.Kc);
  wb.setOverride('Summary', 'BA3', state.Kl);
  wb.setOverride('Summary', 'BA4', state.Kp);

  wb.setOverride('Summary', 'BB1', state.msMode);
  wb.setOverride('Summary', 'BB2', state.hs);
  wb.setOverride('Summary', 'BB3', state.bs);
  wb.setOverride('Summary', 'BB4', state.ns);

  const n = state.rows.length;
  const startRow = 16, endRow = 16 + n - 1;
  wb.setOverride('Summary', 'AJ15', endRow);

  // Rewrite the group-average formulas to span the live row count
  WORKBOOK_BUNDLE.sheets.Wy.CX21 = `=AVERAGE(Summary!H${startRow}:H${endRow})`;
  WORKBOOK_BUNDLE.sheets.Wx.CX21 = `=AVERAGE(Summary!X${startRow}:X${endRow})`;

  ensureSummaryRowFormulas(startRow, endRow);

  const z = computeZSequence(state.rows);

  for (let i = 0; i < n; i++) {
    const row = state.rows[i];
    const summaryRow = startRow + i;

    wb.setOverride('Summary', `B${summaryRow}`, z[i]);
    wb.setOverride('Summary', `C${summaryRow}`, row.floorHt);
    wb.setOverride('Summary', `H${summaryRow}`, row.ly);
    wb.setOverride('Summary', `X${summaryRow}`, row.lx);

    wb.setOverride('Wy', 'J17', z[i]); wb.setOverride('Wy', 'J13', row.ly); wb.setOverride('Wy', 'J15', row.lx);
    wb.setOverride('Wx', 'J17', z[i]); wb.setOverride('Wx', 'J15', row.ly); wb.setOverride('Wx', 'J13', row.lx);
    wb.reset();

    const out = {};
    for (const key in OUTPUT_NAMES) {
      const nm = WORKBOOK_BUNDLE.names[OUTPUT_NAMES[key]];
      out[key] = wb.getCellValue(nm.sheet, nm.ref);
    }

    wb.setOverride('Summary', `D${summaryRow}`, out.Mz_cat_y);
    wb.setOverride('Summary', `E${summaryRow}`, out.Mh_y);
    wb.setOverride('Summary', `F${summaryRow}`, out.Vsit_y);
    wb.setOverride('Summary', `G${summaryRow}`, out.Vdes_y);
    wb.setOverride('Summary', `I${summaryRow}`, out.Cdyn_y_SLS);
    wb.setOverride('Summary', `J${summaryRow}`, out.Wy_W_SLS);
    wb.setOverride('Summary', `K${summaryRow}`, out.Wy_L_SLS);
    wb.setOverride('Summary', `L${summaryRow}`, out.Wy_SLS);
    wb.setOverride('Summary', `N${summaryRow}`, out.a_max_y);
    wb.setOverride('Summary', `O${summaryRow}`, out.Cdyn_y_ULS);
    wb.setOverride('Summary', `P${summaryRow}`, out.Wy_W_ULS);
    wb.setOverride('Summary', `Q${summaryRow}`, out.Wy_L_ULS);
    wb.setOverride('Summary', `R${summaryRow}`, out.Wy_ULS);
    wb.setOverride('Summary', `T${summaryRow}`, out.Mz_cat_x);
    wb.setOverride('Summary', `U${summaryRow}`, out.Mh_x);
    wb.setOverride('Summary', `V${summaryRow}`, out.Vsit_x);
    wb.setOverride('Summary', `W${summaryRow}`, out.Vdes_x);
    wb.setOverride('Summary', `Y${summaryRow}`, out.Cdyn_x_SLS);
    wb.setOverride('Summary', `Z${summaryRow}`, out.Wx_W_SLS);
    wb.setOverride('Summary', `AA${summaryRow}`, out.Wx_L_SLS);
    wb.setOverride('Summary', `AB${summaryRow}`, out.Wx_SLS);
    wb.setOverride('Summary', `AD${summaryRow}`, out.a_max_x);
    wb.setOverride('Summary', `AE${summaryRow}`, out.Cdyn_x_ULS);
    wb.setOverride('Summary', `AF${summaryRow}`, out.Wx_W_ULS);
    wb.setOverride('Summary', `AG${summaryRow}`, out.Wx_L_ULS);
    wb.setOverride('Summary', `AH${summaryRow}`, out.Wx_ULS);
    wb.reset();

    row.results = {
      z: z[i],
      Vdes_y: out.Vdes_y, Wy_SLS: out.Wy_SLS,
      Vdes_x: out.Vdes_x, Wx_SLS: out.Wx_SLS,
      Fy_SLS: wb.getCellValue('Summary', `M${summaryRow}`),
      Fx_SLS: wb.getCellValue('Summary', `AC${summaryRow}`),
      // keep full result set for the export report
      full: out,
      Fy_ULS: wb.getCellValue('Summary', `S${summaryRow}`),
      Fx_ULS: wb.getCellValue('Summary', `AI${summaryRow}`),
    };
    wb.reset();
  }

  state.hasRun = true;
  return state.rows;
}

export function addRow() {
  if (state.rows.length >= 99) return false;
  state.rows.push({ story: `New Storey ${state.rows.length + 1}`, floorHt: 3.0, ly: 200, lx: 50, deadLoad: 0, results: null });
  return true;
}

export function removeRow(index) {
  if (state.rows.length <= 1) return false;
  state.rows.splice(index, 1);
  return true;
}
