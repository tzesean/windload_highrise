import { state, runAll, addRow, removeRow } from './app.js';
import { exportExcelReport } from './export.js';

// ---- Bind simple text/number inputs via data-bind attribute ----
function bindSimpleInputs() {
  document.querySelectorAll('[data-bind]').forEach(el => {
    const key = el.dataset.bind;
    if (state[key] !== undefined) {
      el.value = state[key];
    }
    el.addEventListener('input', () => {
      const v = el.type === 'number' ? parseFloat(el.value) : el.value;
      state[key] = v;
    });
  });
}

function toggleButtons(groupSelector, activeValue, stateKey) {
  document.querySelectorAll(groupSelector + ' button').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.value === String(state[stateKey]));
    btn.addEventListener('click', () => {
      state[stateKey] = btn.dataset.value;
      document.querySelectorAll(groupSelector + ' button').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
    });
  });
}

function updateHillVisibility() {
  const fields = document.getElementById('hillFields');
  const fig = document.getElementById('hillFig');
  const cap = document.getElementById('hillCap');
  if (state.hillType === 'None') { fields.style.display = 'none'; return; }
  fields.style.display = '';
  if (state.hillType === 'Hills and Ridges') { fig.src = 'assets/Figure_4_2_input.png'; cap.textContent = 'Figure 4.2 — Hills and Ridges'; }
  else { fig.src = 'assets/Figure_4_3_input.png'; cap.textContent = 'Figure 4.3 — Escarpments'; }
}

function updateMsVisibility() {
  document.getElementById('msFields').style.display = (state.msMode === 'Calculate') ? '' : 'none';
}

function renderRowCount() {
  document.getElementById('rowCountLabel').textContent = state.rows.length;
  document.getElementById('rowCountBar').style.width = Math.min(100, (state.rows.length / 99) * 100) + '%';
}

function fmt(v) {
  if (v === null || v === undefined || v === '' || (typeof v === 'object')) return '—';
  if (typeof v === 'number') return (Math.round(v * 1000) / 1000).toString();
  return v;
}

function renderTable() {
  const tbody = document.getElementById('rows');
  tbody.innerHTML = '';
  state.rows.forEach((row, i) => {
    const tr = document.createElement('tr');
    const r = row.results;
    tr.innerHTML = `
      <td class="sticky-col"><input class="cellinput" data-row="${i}" data-field="story" value="${row.story}" style="width:90px;background:transparent;border:none;color:inherit;font:inherit;"></td>
      <td class="input-cell"><input class="cellinput" data-row="${i}" data-field="floorHt" type="number" step="0.01" value="${row.floorHt}"></td>
      <td class="input-cell"><input class="cellinput" data-row="${i}" data-field="ly" type="number" step="1" value="${row.ly}"></td>
      <td class="input-cell"><input class="cellinput" data-row="${i}" data-field="lx" type="number" step="1" value="${row.lx}"></td>
      <td class="input-cell"><input class="cellinput" data-row="${i}" data-field="deadLoad" type="number" step="0.01" value="${row.deadLoad}"></td>
      <td class="calc-cell">${r ? fmt(r.z) : '—'}</td>
      <td class="calc-cell">${r ? fmt(r.Vdes_y) : '—'}</td>
      <td class="calc-cell">${r ? fmt(r.Wy_SLS) : '—'}</td>
      <td class="calc-cell">${r ? fmt(r.Fy_SLS) : '—'}</td>
      <td class="calc-cell">${r ? fmt(r.Vdes_x) : '—'}</td>
      <td class="calc-cell">${r ? fmt(r.Wx_SLS) : '—'}</td>
      <td class="calc-cell">${r ? fmt(r.Fx_SLS) : '—'}</td>
      <td><button class="rowdel" data-row="${i}" title="Remove row">✕</button></td>`;
    tbody.appendChild(tr);
  });

  tbody.querySelectorAll('.cellinput').forEach(inp => {
    inp.addEventListener('input', () => {
      const i = parseInt(inp.dataset.row, 10);
      const field = inp.dataset.field;
      state.rows[i][field] = field === 'story' ? inp.value : parseFloat(inp.value);
    });
  });
  tbody.querySelectorAll('.rowdel').forEach(btn => {
    btn.addEventListener('click', () => {
      removeRow(parseInt(btn.dataset.row, 10));
      renderTable();
      renderRowCount();
    });
  });

  const addRowTr = document.createElement('tr');
  addRowTr.className = 'addrow';
  addRowTr.innerHTML = `<td colspan="13">+ Add storey row (${state.rows.length} / 99 used)</td>`;
  addRowTr.addEventListener('click', () => {
    if (addRow()) { renderTable(); renderRowCount(); }
  });
  tbody.appendChild(addRowTr);
}

function init() {
  bindSimpleInputs();

  document.getElementById('hillSelect').value = state.hillType;
  document.getElementById('hillSelect').addEventListener('change', e => { state.hillType = e.target.value; updateHillVisibility(); });
  updateHillVisibility();

  document.getElementById('msSelect').value = state.msMode;
  document.getElementById('msSelect').addEventListener('change', e => { state.msMode = e.target.value; updateMsVisibility(); });
  updateMsVisibility();

  document.querySelectorAll('#cdynToggle button').forEach(btn => {
    btn.addEventListener('click', () => {
      state.cdyn = btn.dataset.value;
      document.querySelectorAll('#cdynToggle button').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
    });
  });

  document.getElementById('addRowBtn').addEventListener('click', () => {
    if (addRow()) { renderTable(); renderRowCount(); }
  });

  document.getElementById('runAllBtn').addEventListener('click', () => {
    const btn = document.getElementById('runAllBtn');
    btn.disabled = true; btn.textContent = 'Running…';
    setTimeout(() => {
      try {
        runAll();
        renderTable();
      } catch (e) {
        console.error(e);
        alert('Calculation error: ' + e.message);
      } finally {
        btn.disabled = false; btn.textContent = '▶ Run All';
      }
    }, 10);
  });

  document.getElementById('exportBtn').addEventListener('click', async () => {
    if (!state.hasRun) { alert('Run All first so the report has computed results.'); return; }
    const btn = document.getElementById('exportBtn');
    btn.disabled = true; btn.textContent = 'Generating…';
    try {
      await exportExcelReport(state);
    } catch (e) {
      console.error(e);
      alert('Export error: ' + e.message);
    } finally {
      btn.disabled = false; btn.textContent = '⤓ Generate Excel Report';
    }
  });

  renderTable();
  renderRowCount();
}

document.addEventListener('DOMContentLoaded', init);
