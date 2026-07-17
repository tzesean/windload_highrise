// Builds the Wind Load report as a real .xlsx file in-browser using ExcelJS,
// from the live app state + computed results (no server involved).

const COL_WIDTH_A = (15 - 5) / 7;
const COL_WIDTH_REST = (60 - 5) / 7;

function styleHeaderCell(cell, text, opts = {}) {
  cell.value = text;
  cell.font = { name: 'Arial', bold: true, size: opts.size || 10 };
  cell.alignment = { vertical: 'middle', wrapText: true, horizontal: opts.align || 'left' };
}

function fullBorder(cell) {
  const thin = { style: 'thin', color: { argb: 'FF000000' } };
  cell.border = { left: thin, right: thin, top: thin, bottom: thin };
}

async function loadImageBuffer(path) {
  const res = await fetch(path);
  return await res.arrayBuffer();
}

function fieldRow(ws, row, label, value, unit = '', ref = '') {
  ws.getCell(row, 2).value = label;
  ws.getCell(row, 2).font = { name: 'Arial', size: 10 };
  ws.getCell(row, 4).value = '=';
  ws.getCell(row, 4).font = { name: 'Arial', size: 10 };
  const vc = ws.getCell(row, 5);
  vc.value = typeof value === 'number' ? Math.round(value * 100000) / 100000 : value;
  vc.font = { name: 'Arial', size: 10, bold: true };
  vc.alignment = { horizontal: 'right' };
  if (unit) { ws.getCell(row, 6).value = unit; ws.getCell(row, 6).font = { name: 'Arial', size: 9, italic: true, color: { argb: 'FF555555' } }; }
  if (ref) { ws.getCell(row, 11).value = ref; ws.getCell(row, 11).font = { name: 'Arial', size: 9, color: { argb: 'FF0B6623' } }; }
  return row + 2;
}

function sectionHeader(ws, row, text) {
  const c = ws.getCell(row, 2);
  c.value = text;
  c.font = { name: 'Arial', bold: true, underline: true, size: 10.5 };
  return row + 2;
}

function storeyTable(ws, row, headers, units, dataRows, wb) {
  const startRow = row;
  headers.forEach((h, j) => {
    const c = ws.getCell(row, 2 + j);
    c.value = h;
    c.font = { name: 'Arial', bold: true, size: 9 };
    c.alignment = { horizontal: 'center', wrapText: true };
    fullBorder(c);
  });
  row++;
  units.forEach((u, j) => {
    const c = ws.getCell(row, 2 + j);
    c.value = u;
    c.font = { name: 'Arial', italic: true, size: 8, color: { argb: 'FF555555' } };
    c.alignment = { horizontal: 'center' };
    fullBorder(c);
  });
  row++;
  for (const data of dataRows) {
    data.forEach((val, j) => {
      const c = ws.getCell(row, 2 + j);
      c.value = val;
      c.font = { name: 'Arial', size: 9 };
      c.alignment = { horizontal: 'center' };
      fullBorder(c);
    });
    row++;
  }
  return row + 1;
}

export async function exportExcelReport(state) {
  const ExcelJS = window.ExcelJS;
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Output', { pageSetup: { orientation: 'portrait' } });

  ws.getColumn(1).width = COL_WIDTH_A;
  for (let c = 2; c <= 12; c++) ws.getColumn(c).width = COL_WIDTH_REST;

  const results = state.rows.map(r => r.results).filter(Boolean);
  const structHeight = Math.max(...state.rows.map((r, i) => r.results ? r.results.z : 0));
  const showHill = state.hillType !== 'None';
  const figPath = state.hillType === 'Escarpment' ? 'assets/Figure_4_3.png' : 'assets/Figure_4_2.png';
  const figLabel = state.hillType === 'Escarpment' ? 'Figure 4.3' : 'Figure 4.2';

  let imageId = null;
  if (showHill) {
    try {
      const buf = await loadImageBuffer(figPath);
      imageId = wb.addImage({ buffer: buf, extension: 'png' });
    } catch (e) { console.warn('Could not load figure for export', e); }
  }

  function buildDirectionSheet(row, direction, sheetNo) {
    // Job header
    ws.mergeCells(row, 2, row + 1, 4);
    styleHeaderCell(ws.getCell(row, 2), 'Wind Load to MS 1553:2002 with AMD 1:2013', { size: 11 });
    ws.getCell(row, 6).value = 'Job No.'; ws.getCell(row, 6).font = { bold: true, size: 9 };
    ws.getCell(row, 8).value = state.jobNo; ws.getCell(row, 8).font = { bold: true, size: 9, color: { argb: 'FF0B6623' } };
    ws.getCell(row, 11).value = 'Sheet'; ws.getCell(row, 11).font = { bold: true, size: 9 };
    ws.getCell(row, 12).value = sheetNo; ws.getCell(row, 12).font = { bold: true, size: 9 };
    row++;
    ws.getCell(row, 6).value = 'Job Title'; ws.getCell(row, 6).font = { bold: true, size: 9 };
    ws.getCell(row, 8).value = state.jobTitle; ws.getCell(row, 8).font = { bold: true, size: 9, color: { argb: 'FF0B6623' } };
    row += 2;

    ws.mergeCells(row, 2, row + 1, 8);
    styleHeaderCell(ws.getCell(row, 2), `${direction} Direction Calculation in accordance with MS 1553:2002 with AMD 1:2013`, { size: 10.5 });
    ws.getCell(row, 11).value = 'Reference'; ws.getCell(row, 11).font = { bold: true, underline: true, size: 9 };
    row += 3;

    row = sectionHeader(ws, row, 'A) Dimensions');
    row = fieldRow(ws, row, 'Structure Height, h', structHeight, 'm');
    row = fieldRow(ws, row, 'z considered', structHeight, 'm');
    row++;

    let sectionLetter = 'C';
    if (showHill) {
      row = sectionHeader(ws, row, 'B) Hill Shape Multiplier');
      const boxTopRow = row;
      const boxRows = 7, boxCols = 8;
      ws.mergeCells(row, 2, row + boxRows - 1, 1 + boxCols);
      const boxCell = ws.getCell(row, 2);
      fullBorder(boxCell);
      if (imageId !== null) {
        ws.addImage(imageId, {
          tl: { col: 1.1, row: row - 1 + 0.15 },
          ext: { width: 420, height: 100 },
        });
      }
      ws.getCell(row, 11).value = figLabel;
      ws.getCell(row, 11).font = { bold: true, size: 9, color: { argb: 'FF0B6623' } };
      row += boxRows + 1;
      row = fieldRow(ws, row, 'H', state.H, 'm');
      row = fieldRow(ws, row, 'Lu', state.Lu, 'm');
      row = fieldRow(ws, row, 'x', state.x, 'm');
      row = fieldRow(ws, row, 'z', structHeight, 'm');
      row++;
    } else {
      sectionLetter = 'B';
    }

    const sample = results[0] || {};
    const key = direction === 'WX' ? 'x' : 'y';
    row = sectionHeader(ws, row, `${sectionLetter}) Design Wind Speed`);
    row = fieldRow(ws, row, 'Vs', state.vbasic, 'm/s', 'Table 3.1');
    row = fieldRow(ws, row, 'Md', 1.00, '', 'Sec 2.2');
    row = fieldRow(ws, row, 'Mz,cat', `${sample[`Mz_cat_${key}`] ?? ''}  (Cat ${state.terrainCategory})`, '', 'Table 4.1');
    row = fieldRow(ws, row, 'Ms', state.msMode === 'Default' ? 1 : (state.msMode === 'Suburban' ? 0.85 : 'calculated'), '', 'Sec 4.3');
    row = fieldRow(ws, row, 'Mh', sample[`Mh_${key}`] ?? '', '', 'Sec 4.4');
    row = fieldRow(ws, row, 'Vsit', sample[`Vsit_${key}`] ?? '', 'm/s', 'Sec 2.2');
    row = fieldRow(ws, row, 'I', 1.00, '', 'Table 3.2');
    row = fieldRow(ws, row, 'Vdes', sample[`Vdes_${key}`] ?? '', 'm/s', 'Sec 2.4.1');
    row += 2;
    return row;
  }

  let row = 1;
  row = buildDirectionSheet(row, 'WX', 1);
  row += 4;
  row = buildDirectionSheet(row, 'WY', 2);
  row += 4;

  // Storey tables (WY then WX), full row count
  const wyHeaders = ['Storey', 'Reduced Level, z [m]', 'Floor Height [m]', 'Mz,cat', 'Mh', 'Vsit [m/s]', 'Vdes [m/s]', 'Ly [m]'];
  const wyUnits = ['', '', '', 'Cat ' + state.terrainCategory, '', '', '', ''];
  const wyRows = state.rows.map(r => [
    r.story, r.results?.z ?? '', r.floorHt, r.results?.full?.Mz_cat_y ?? '', r.results?.full?.Mh_y ?? '',
    r.results?.Vdes_y ?? '', r.results?.Vdes_y ?? '', r.ly,
  ]);
  row = sectionHeader(ws, row, 'Storey Data — WY Direction');
  row = storeyTable(ws, row, wyHeaders, wyUnits, wyRows, wb);

  const wxHeaders = ['Storey', 'Reduced Level, z [m]', 'Floor Height [m]', 'Mz,cat', 'Mh', 'Vsit [m/s]', 'Vdes [m/s]', 'Lx [m]'];
  const wxRows = state.rows.map(r => [
    r.story, r.results?.z ?? '', r.floorHt, r.results?.full?.Mz_cat_x ?? '', r.results?.full?.Mh_x ?? '',
    r.results?.Vdes_x ?? '', r.results?.Vdes_x ?? '', r.lx,
  ]);
  row = sectionHeader(ws, row, 'Storey Data — WX Direction');
  row = storeyTable(ws, row, wxHeaders, wyUnits, wxRows, wb);

  // WY-SLS / WY-ULS / WX-SLS / WX-ULS tables
  function forceTable(title, headers, rows) {
    row = sectionHeader(ws, row, title);
    row = storeyTable(ws, row, headers, headers.map(() => ''), rows, wb);
  }
  forceTable('WY - SLS', ['Storey', 'Cdyn', 'Wy(W) [kPa]', 'Wy(L) [kPa]', 'Wy [kPa]', 'Fy [kN]', 'ay [mg]'],
    state.rows.map(r => [r.story, r.results?.full?.Cdyn_y_SLS ?? '', r.results?.full?.Wy_W_SLS ?? '', r.results?.full?.Wy_L_SLS ?? '', r.results?.Wy_SLS ?? '', r.results?.Fy_SLS ?? '', r.results?.full?.a_max_y ?? '']));
  forceTable('WY - ULS', ['Storey', 'Cdyn', 'Wy(W) [kPa]', 'Wy(L) [kPa]', 'WLy [kPa]', 'Fy [kN]'],
    state.rows.map(r => [r.story, r.results?.full?.Cdyn_y_ULS ?? '', r.results?.full?.Wy_W_ULS ?? '', r.results?.full?.Wy_L_ULS ?? '', r.results?.full?.Wy_ULS ?? '', r.results?.Fy_ULS ?? '']));
  forceTable('WX - SLS', ['Storey', 'Cdyn', 'Wx(W) [kPa]', 'Wx(L) [kPa]', 'Wx [kPa]', 'Fx [kN]', 'ax [mg]'],
    state.rows.map(r => [r.story, r.results?.full?.Cdyn_x_SLS ?? '', r.results?.full?.Wx_W_SLS ?? '', r.results?.full?.Wx_L_SLS ?? '', r.results?.Wx_SLS ?? '', r.results?.Fx_SLS ?? '', r.results?.full?.a_max_x ?? '']));
  forceTable('WX - ULS', ['Storey', 'Cdyn', 'Wx(W) [kPa]', 'Wx(L) [kPa]', 'Wx [kPa]', 'Fx [kN]'],
    state.rows.map(r => [r.story, r.results?.full?.Cdyn_x_ULS ?? '', r.results?.full?.Wx_W_ULS ?? '', r.results?.full?.Wx_L_ULS ?? '', r.results?.full?.Wx_ULS ?? '', r.results?.Fx_ULS ?? '']));

  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], { type: 'application/octet-stream' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `Wind_Load_Report_${state.jobNo || 'export'}.xlsx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
