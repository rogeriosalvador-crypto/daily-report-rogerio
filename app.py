"""
Daily Report — Carteira Rogério Salvador
Flask dashboard completo: 6 períodos, filtro por loja, download Excel.
"""
from pathlib import Path
from flask import Flask, render_template_string, jsonify, request, send_file
import json, os, io

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, numbers
    OPENPYXL_OK = True
except ImportError:
    OPENPYXL_OK = False

BASE_DIR = Path(__file__).parent.absolute()
app = Flask(__name__)

# ---------------------------------------------------------------------------
# HTML TEMPLATE (single-file, no templates/ folder needed)
# ---------------------------------------------------------------------------
HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Relatório Diário — Carteira Rogério Salvador</title>
<style>
  :root {
    --bg-dark:#0F0F1A; --bg-card:#1A1A2E; --bg-card2:#16213E;
    --ifood-red:#EA1D2C; --ifood-red-dark:#B71C2B;
    --green-ok:#27AE60; --yellow-warn:#F39C12; --red-alert:#E74C3C;
    --text-primary:#FFFFFF; --text-secondary:#B0BEC5;
    --border:#2C3E50; --hover:#22304a;
  }
  *{box-sizing:border-box;margin:0;padding:0}
  body{background:var(--bg-dark);color:var(--text-primary);
       font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;
       font-size:14px;min-height:100vh;padding:16px}
  .container{max-width:1280px;margin:0 auto}

  /* HEADER */
  .header{background:var(--bg-card);border:1px solid var(--border);
          border-radius:12px;padding:18px;margin-bottom:16px}
  .header-top{display:flex;align-items:center;justify-content:space-between;
              flex-wrap:wrap;gap:10px;margin-bottom:10px}
  .logo{font-size:22px;font-weight:800;color:var(--ifood-red);letter-spacing:.5px}
  .update-badge{background:var(--ifood-red-dark);color:#fff;padding:5px 13px;
                border-radius:20px;font-size:12px;font-weight:600}
  .header-subtitle{color:var(--text-secondary);font-size:13px}
  .ref-date{font-weight:600;color:var(--text-primary)}
  .header-actions{display:flex;gap:8px;flex-wrap:wrap}
  .btn{display:inline-flex;align-items:center;gap:6px;padding:8px 14px;
       border:none;border-radius:8px;font-size:13px;font-weight:600;
       cursor:pointer;text-decoration:none;transition:opacity .15s}
  .btn-red{background:var(--ifood-red);color:#fff}
  .btn-outline{background:transparent;color:var(--text-secondary);
               border:1px solid var(--border)}
  .btn:hover{opacity:.85}

  /* GROUP TABS */
  .group-tabs{display:flex;gap:8px;margin-bottom:14px;flex-wrap:wrap}
  .group-tab{padding:10px 22px;border-radius:10px;border:2px solid var(--border);
             background:var(--bg-card);color:var(--text-secondary);
             cursor:pointer;font-weight:700;font-size:15px;transition:all .15s}
  .group-tab.active{border-color:var(--ifood-red);color:#fff;
                    background:var(--ifood-red-dark)}

  /* PERIOD PILLS */
  .period-bar{display:flex;gap:6px;flex-wrap:wrap;margin-bottom:16px}
  .period-pill{padding:6px 16px;border-radius:20px;border:1px solid var(--border);
               background:var(--bg-card);color:var(--text-secondary);
               cursor:pointer;font-size:13px;font-weight:600;transition:all .15s}
  .period-pill.active{background:var(--ifood-red);color:#fff;border-color:var(--ifood-red)}

  /* KPI CARDS */
  .kpi-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));
            gap:12px;margin-bottom:20px}
  .kpi-card{background:var(--bg-card);border:1px solid var(--border);
            border-radius:10px;padding:14px}
  .kpi-label{font-size:11px;color:var(--text-secondary);text-transform:uppercase;
             letter-spacing:.5px;margin-bottom:6px}
  .kpi-value{font-size:20px;font-weight:800;color:var(--text-primary)}
  .kpi-value.big{font-size:26px;color:var(--ifood-red)}
  .kpi-badge{display:inline-block;padding:3px 9px;border-radius:5px;
             font-weight:700;font-size:14px;margin-left:6px}
  .kpi-badge.verde{background:var(--green-ok);color:#fff}
  .kpi-badge.amarelo{background:var(--yellow-warn);color:#fff}
  .kpi-badge.vermelho{background:var(--red-alert);color:#fff}
  .kpi-delta{margin-top:4px;font-size:12px}
  .kpi-delta.up{color:var(--green-ok)}
  .kpi-delta.down{color:var(--red-alert)}
  .kpi-delta.neutral{color:var(--text-secondary)}

  /* STORES TABLE */
  .stores-section{background:var(--bg-card);border:1px solid var(--border);
                  border-radius:12px;padding:18px;margin-bottom:20px}
  .section-title{font-size:16px;font-weight:700;color:var(--ifood-red);margin-bottom:14px}
  .table-wrap{overflow-x:auto}
  table{width:100%;border-collapse:collapse}
  thead{background:var(--bg-card2)}
  th{text-align:left;padding:10px 12px;font-size:11px;text-transform:uppercase;
     letter-spacing:.5px;color:var(--text-secondary);font-weight:600;white-space:nowrap}
  td{padding:10px 12px;border-top:1px solid var(--border);font-size:13px}
  tbody tr:hover{background:var(--hover)}
  .val-pos{color:var(--green-ok);font-weight:600}
  .val-neg{color:var(--red-alert);font-weight:600}
  .val-warn{color:var(--yellow-warn);font-weight:600}
  .semaforo{width:10px;height:10px;border-radius:50%;display:inline-block;margin-right:6px}
  .sem-v{background:var(--green-ok);box-shadow:0 0 6px var(--green-ok)}
  .sem-a{background:var(--yellow-warn);box-shadow:0 0 6px var(--yellow-warn)}
  .sem-r{background:var(--red-alert);box-shadow:0 0 6px var(--red-alert)}

  /* PRINT */
  @media print{
    .header-actions,.group-tabs,.period-bar{display:none!important}
    body{background:#fff;color:#000;padding:0}
    .kpi-card,.stores-section{border:1px solid #ccc;background:#fff}
    .kpi-value,.section-title{color:#000}
  }

  /* MOBILE */
  @media(max-width:600px){
    .kpi-grid{grid-template-columns:repeat(2,1fr)}
    .header-top{flex-direction:column;align-items:flex-start}
    th,td{padding:8px 8px}
  }
</style>
</head>
<body>
<div class="container">

  <!-- HEADER -->
  <div class="header">
    <div class="header-top">
      <div class="logo">iFood</div>
      <div id="updateBadge" class="update-badge">Carregando…</div>
    </div>
    <div class="header-subtitle">
      Relatório Diário — Carteira Rogério Salvador &nbsp;|&nbsp;
      Referência: <span id="refDate" class="ref-date">…</span>
    </div>
    <div class="header-actions" style="margin-top:12px">
      <a class="btn btn-red" href="/download/excel" download>⬇ Download Excel</a>
      <button class="btn btn-outline" onclick="window.print()">🖨 Imprimir</button>
      <button class="btn btn-outline" onclick="loadData()">↺ Atualizar</button>
    </div>
  </div>

  <!-- GROUP TABS -->
  <div class="group-tabs" id="groupTabs"></div>

  <!-- PERIOD PILLS -->
  <div class="period-bar" id="periodBar"></div>

  <!-- KPI CARDS -->
  <div class="kpi-grid" id="kpiGrid"></div>

  <!-- Gráficos -->
  <div class="stores-section">
    <div class="section-title">📊 Evolução GMV por Período</div>
    <div id="chart-gmv-daily" style="padding:10px 0"></div>
  </div>
  <div class="stores-section">
    <div class="section-title">🎯 Atingimento% por Período</div>
    <div id="chart-ating-daily" style="padding:10px 0"></div>
  </div>
  <div class="stores-section">
    <div class="section-title">🏆 Top Lojas por GMV (Período Ativo)</div>
    <div id="chart-topLojas-daily" style="padding:10px 0"></div>
  </div>

  <!-- STORES TABLE -->
  <div class="stores-section" id="storesSection">
    <div class="section-title">🏪 Lojas do Grupo</div>
    <div class="table-wrap">
      <table>
        <thead id="storesThead"></thead>
        <tbody id="storesTbody"></tbody>
      </table>
    </div>
  </div>

</div>

<script>
// ─── State ────────────────────────────────────────────────────────────────
let DATA = null;
const PERIODS = ['D-1','D-7','D-15','D-21','MTD','CONSOLIDADO'];
const PERIOD_LABELS = {'D-1':'D-1','D-7':'D-7','D-15':'D-15','D-21':'D-21','MTD':'MTD','CONSOLIDADO':'Consolidado'};
let currentGroup = null;
let currentPeriod = 'D-1';

// ─── Formatters ───────────────────────────────────────────────────────────
const fmtBRL = v => v == null || v === 0 ? 'R$ —' :
  'R$ ' + Number(v).toLocaleString('pt-BR', {minimumFractionDigits:2,maximumFractionDigits:2});
const fmtPct = v => v == null || v === 0 ? '—' : Number(v).toLocaleString('pt-BR',{minimumFractionDigits:1,maximumFractionDigits:1}) + '%';
const fmtNum = v => v == null || v === 0 ? '—' : Number(v).toLocaleString('pt-BR');

function semClass(ating) {
  if(ating >= 100) return 'sem-v';
  if(ating >= 85)  return 'sem-a';
  return 'sem-r';
}
function badgeClass(ating) {
  if(ating >= 100) return 'verde';
  if(ating >= 85)  return 'amarelo';
  return 'vermelho';
}
function deltaHtml(val, prev, invertBad) {
  if(!prev || prev === 0 || !val) return '';
  const pct = ((val / prev) - 1) * 100;
  const up = pct >= 0;
  const cls = invertBad ? (up ? 'down' : 'up') : (up ? 'up' : 'down');
  const arrow = up ? '▲' : '▼';
  return `<div class="kpi-delta ${cls}">${arrow} ${Math.abs(pct).toLocaleString('pt-BR',{minimumFractionDigits:1,maximumFractionDigits:1})}% vs período ant.</div>`;
}

// ─── Load data ────────────────────────────────────────────────────────────
// ─── Group tabs ───────────────────────────────────────────────────────────
function renderGroupTabs(groups) {
  const el = document.getElementById('groupTabs');
  el.innerHTML = groups.map(g =>
    `<div class="group-tab ${g===currentGroup?'active':''}" onclick="selectGroup('${g}')">${g}</div>`
  ).join('');
}
function selectGroup(g) {
  currentGroup = g;
  renderGroupTabs(Object.keys(DATA.periodos));
  renderKPIs();
  renderStores();
  renderAllCharts();
}

// ─── Period pills ─────────────────────────────────────────────────────────
function renderPeriodBar() {
  const el = document.getElementById('periodBar');
  el.innerHTML = PERIODS.map(p =>
    `<div class="period-pill ${p===currentPeriod?'active':''}" onclick="selectPeriod('${p}')">${PERIOD_LABELS[p]}</div>`
  ).join('');
}
function selectPeriod(p) {
  currentPeriod = p;
  renderPeriodBar();
  renderKPIs();
  renderStores();
  renderAllCharts();
}

// ─── KPI Cards ────────────────────────────────────────────────────────────
function prevPeriod(p) {
  const idx = PERIODS.indexOf(p);
  return idx > 0 ? PERIODS[idx-1] : null;
}

function renderKPIs() {
  const periodos = (DATA.periodos || {})[currentGroup] || {};
  const cur = periodos[currentPeriod] || {};
  const prev = periodos[prevPeriod(currentPeriod)] || {};
  const grid = document.getElementById('kpiGrid');

  const isMTD = currentPeriod === 'MTD' || currentPeriod === 'CONSOLIDADO';

  const cards = [];

  // GMV + ating%
  const atingPct = cur.ating_pct || 0;
  cards.push(`<div class="kpi-card">
    <div class="kpi-label">GMV</div>
    <div class="kpi-value big">${fmtBRL(cur.gmv)}
      ${cur.ating_pct != null ? `<span class="kpi-badge ${badgeClass(atingPct)}">${fmtPct(atingPct)}</span>` : ''}
    </div>
    ${deltaHtml(cur.gmv, prev.gmv, false)}
  </div>`);

  // Meta GMV
  cards.push(`<div class="kpi-card">
    <div class="kpi-label">Meta GMV</div>
    <div class="kpi-value">${fmtBRL(cur.meta_gmv)}</div>
  </div>`);

  // Pedidos
  cards.push(`<div class="kpi-card">
    <div class="kpi-label">Pedidos</div>
    <div class="kpi-value">${fmtNum(cur.pedidos)}</div>
    ${deltaHtml(cur.pedidos, prev.pedidos, false)}
  </div>`);

  if(!isMTD) {
    // AOV
    cards.push(`<div class="kpi-card">
      <div class="kpi-label">AOV</div>
      <div class="kpi-value">${cur.aov ? fmtBRL(cur.aov) : '—'}</div>
      ${deltaHtml(cur.aov, prev.aov, false)}
    </div>`);

    // CAP/FAT
    cards.push(`<div class="kpi-card">
      <div class="kpi-label">CAP / FAT</div>
      <div class="kpi-value">${fmtPct(cur.cap_fat)}</div>
    </div>`);

    // Cancelamento
    const cancelClass = cur.cancel_pct > 7 ? 'val-neg' : cur.cancel_pct > 5 ? 'val-warn' : '';
    cards.push(`<div class="kpi-card">
      <div class="kpi-label">Cancelamento</div>
      <div class="kpi-value ${cancelClass}">${fmtPct(cur.cancel_pct)}</div>
      ${deltaHtml(cur.cancel_pct, prev.cancel_pct, true)}
    </div>`);

    // Ruptura
    const ruptClass = cur.ruptura_pct > 3 ? 'val-neg' : cur.ruptura_pct > 1 ? 'val-warn' : '';
    cards.push(`<div class="kpi-card">
      <div class="kpi-label">Ruptura</div>
      <div class="kpi-value ${ruptClass}">${fmtPct(cur.ruptura_pct)}</div>
      ${deltaHtml(cur.ruptura_pct, prev.ruptura_pct, true)}
    </div>`);

    // ER%
    cards.push(`<div class="kpi-card">
      <div class="kpi-label">ER</div>
      <div class="kpi-value">${fmtPct(cur.er_pct)}</div>
    </div>`);

    // NPS
    const npsClass = cur.nps >= 70 ? 'val-pos' : cur.nps >= 50 ? 'val-warn' : 'val-neg';
    cards.push(`<div class="kpi-card">
      <div class="kpi-label">NPS</div>
      <div class="kpi-value ${npsClass}">${cur.nps != null ? Number(cur.nps).toLocaleString('pt-BR',{minimumFractionDigits:1}) : '—'}</div>
      ${deltaHtml(cur.nps, prev.nps, false)}
    </div>`);

    // Inv. Merchant
    if(cur.inv_merchant != null) {
      cards.push(`<div class="kpi-card">
        <div class="kpi-label">Inv. Merchant</div>
        <div class="kpi-value">${fmtBRL(cur.inv_merchant)}</div>
      </div>`);
      cards.push(`<div class="kpi-card">
        <div class="kpi-label">% Inv. Merchant</div>
        <div class="kpi-value">${fmtPct(cur.pct_inv_merchant)}</div>
      </div>`);
    }
  }

  grid.innerHTML = cards.join('');
}

// ─── Stores table ─────────────────────────────────────────────────────────
function getLojasParaPeriodo(grupo, periodo) {
  const grupoData = (DATA.lojas || {})[grupo];
  if (!grupoData) return [];
  // Formato novo: {periodo: [lojas...]}
  if (Array.isArray(grupoData)) {
    // Formato legado: lista plana -> disponível apenas em D-1
    return periodo === 'D-1' ? grupoData : [];
  }
  return grupoData[periodo] || [];
}

function renderStores() {
  const lojas = getLojasParaPeriodo(currentGroup, currentPeriod);
  const thead = document.getElementById('storesThead');
  const tbody = document.getElementById('storesTbody');

  if(!lojas.length) {
    thead.innerHTML = '';
    tbody.innerHTML = '<tr><td colspan="7" style="color:var(--text-secondary);text-align:center;padding:20px">Sem dados de lojas para este período</td></tr>';
    return;
  }

  thead.innerHTML = `<tr>
    <th>Loja</th>
    <th style="text-align:right">GMV</th>
    <th style="text-align:right">Pedidos</th>
    <th style="text-align:right">Ating%</th>
    <th style="text-align:right">Cancel%</th>
    <th style="text-align:right">Ruptura%</th>
    <th style="text-align:right">NPS</th>
  </tr>`;

  tbody.innerHTML = lojas.map(l => {
    const ating = l.ating_pct || 0;
    const sem = semClass(ating);
    const cancelCls = l.cancel_pct > 7 ? 'val-neg' : l.cancel_pct > 5 ? 'val-warn' : '';
    const ruptCls = l.ruptura_pct > 3 ? 'val-neg' : l.ruptura_pct > 1 ? 'val-warn' : '';
    const npsCls = l.nps >= 70 ? 'val-pos' : l.nps >= 50 ? 'val-warn' : 'val-neg';
    return `<tr>
      <td><span class="semaforo ${sem}"></span>${l.nome}</td>
      <td style="text-align:right">${fmtBRL(l.gmv)}</td>
      <td style="text-align:right">${fmtNum(l.pedidos)}</td>
      <td style="text-align:right"><span class="kpi-badge ${badgeClass(ating)}" style="font-size:11px">${fmtPct(ating)}</span></td>
      <td style="text-align:right" class="${cancelCls}">${fmtPct(l.cancel_pct)}</td>
      <td style="text-align:right" class="${ruptCls}">${fmtPct(l.ruptura_pct)}</td>
      <td style="text-align:right" class="${npsCls}">${l.nps != null ? Number(l.nps).toLocaleString('pt-BR',{minimumFractionDigits:1}) : '—'}</td>
    </tr>`;
  }).join('');
}

// ─── Gráficos CSS/JS inline ────────────────────────────────────────────────
function fmtBRLjs(v) {
  v = parseFloat(v) || 0;
  return 'R$ ' + v.toLocaleString('pt-BR', {minimumFractionDigits:2, maximumFractionDigits:2});
}

function renderChartsGMV() {
  const el = document.getElementById('chart-gmv-daily');
  if (!el) return;
  const periodos = (DATA.periodos || {})[currentGroup] || {};
  const max = Math.max(...PERIODS.map(p => parseFloat((periodos[p]||{}).gmv)||0), 1);
  el.innerHTML = PERIODS.map(p => {
    const val = parseFloat((periodos[p]||{}).gmv)||0;
    const pct = (val/max*100).toFixed(1);
    const isAtivo = p === currentPeriod;
    return `<div style="display:flex;align-items:center;margin:6px 0;gap:10px">
      <span style="width:110px;font-size:12px;color:${isAtivo?'#fff':'#B0BEC5'};font-weight:${isAtivo?'700':'400'}">${PERIOD_LABELS[p]||p}</span>
      <div style="flex:1;background:#1A1A2E;border-radius:4px;height:24px;position:relative;overflow:hidden">
        <div style="width:${pct}%;background:${isAtivo?'#EA1D2C':'#4a4a6a'};height:100%;border-radius:4px;transition:width 0.3s"></div>
      </div>
      <span style="width:150px;text-align:right;font-size:12px;color:#FFF">${fmtBRLjs(val)}</span>
    </div>`;
  }).join('');
}

function renderChartsAting() {
  const el = document.getElementById('chart-ating-daily');
  if (!el) return;
  const periodos = (DATA.periodos || {})[currentGroup] || {};
  el.innerHTML = PERIODS.map(p => {
    const val = parseFloat((periodos[p]||{}).ating_pct)||0;
    const cor = val >= 100 ? '#27AE60' : val >= 85 ? '#F39C12' : '#E74C3C';
    const isAtivo = p === currentPeriod;
    return `<div style="display:flex;align-items:center;margin:6px 0;gap:10px">
      <span style="width:110px;font-size:12px;color:${isAtivo?'#fff':'#B0BEC5'};font-weight:${isAtivo?'700':'400'}">${PERIOD_LABELS[p]||p}</span>
      <div style="flex:1;background:#1A1A2E;border-radius:4px;height:24px;position:relative;overflow:hidden">
        <div style="width:${Math.min(val/150*100,100).toFixed(1)}%;background:${isAtivo?cor:cor+'99'};height:100%;border-radius:4px"></div>
        <div style="position:absolute;left:0;top:0;height:100%;width:${(100/150*100).toFixed(1)}%;border-right:2px dashed #555"></div>
      </div>
      <span style="width:80px;text-align:right;font-size:12px;color:${cor};font-weight:700">${val.toFixed(1)}%</span>
    </div>`;
  }).join('');
}

function renderChartTopLojas() {
  const el = document.getElementById('chart-topLojas-daily');
  if (!el) return;
  const lojas = getLojasParaPeriodo(currentGroup, currentPeriod);
  if (!lojas || !lojas.length) {
    el.innerHTML = '<p style="color:#666;text-align:center;padding:10px">Sem dados para este período</p>';
    return;
  }
  const top10 = [...lojas].sort((a,b) => (parseFloat(b.gmv)||0)-(parseFloat(a.gmv)||0)).slice(0,10);
  const max = Math.max(...top10.map(l => parseFloat(l.gmv)||0), 1);
  el.innerHTML = top10.map(l => {
    const val = parseFloat(l.gmv)||0;
    const pct = (val/max*100).toFixed(1);
    const ating = parseFloat(l.ating_pct)||0;
    const cor = ating >= 100 ? '#27AE60' : ating >= 85 ? '#F39C12' : '#E74C3C';
    return `<div style="display:flex;align-items:center;margin:5px 0;gap:10px">
      <span style="width:200px;font-size:11px;color:#B0BEC5;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${l.nome}">${l.nome}</span>
      <div style="flex:1;background:#1A1A2E;border-radius:4px;height:22px;position:relative;overflow:hidden">
        <div style="width:${pct}%;background:${cor};height:100%;border-radius:4px;opacity:0.85"></div>
      </div>
      <span style="width:150px;text-align:right;font-size:11px;color:#FFF">${fmtBRLjs(val)}</span>
    </div>`;
  }).join('');
}

function renderAllCharts() {
  renderChartsGMV();
  renderChartsAting();
  renderChartTopLojas();
}

// ─── Boot ─────────────────────────────────────────────────────────────────
async function loadData() {
  const res = await fetch('/api/dados');
  DATA = await res.json();
  const groups = Object.keys(DATA.periodos || {});
  if(!currentGroup || !groups.includes(currentGroup)) currentGroup = groups[0];
  renderGroupTabs(groups);
  renderPeriodBar();
  renderKPIs();
  renderStores();
  renderAllCharts();
  document.getElementById('updateBadge').textContent = 'Atualizado: ' + (DATA.data_atualizacao || '');
  document.getElementById('refDate').textContent = DATA.data_referencia || '';
}
loadData();
</script>
</body>
</html>
"""


# ---------------------------------------------------------------------------
# DATA LOADING
# ---------------------------------------------------------------------------
SAMPLE_DATA = {
    "data_referencia": "28/05/2026",
    "data_atualizacao": "28/05/2026 09:00",
    "periodos": {
        "NAGUMO": {
            "D-1":   {"gmv": 232355.98, "meta_gmv": 221977.84, "ating_pct": 104.7, "pedidos": 1746, "aov": 133.08, "cap_fat": 93.9, "cancel_pct": 4.2, "ruptura_pct": 0.1, "er_pct": 7.1, "inv_merchant": 0.0, "pct_inv_merchant": 0.0, "nps": 82.2},
            "D-7":   {"gmv": 167103.94, "meta_gmv": 210000.00, "ating_pct": 79.6, "pedidos": 1532, "aov": 109.07, "cap_fat": 91.2, "cancel_pct": 5.1, "ruptura_pct": 0.3, "er_pct": 7.8, "inv_merchant": 0.0, "pct_inv_merchant": 0.0, "nps": 80.5},
            "D-15":  {"gmv": 181400.00, "meta_gmv": 218000.00, "ating_pct": 83.2, "pedidos": 1590, "aov": 114.09, "cap_fat": 92.5, "cancel_pct": 4.8, "ruptura_pct": 0.2, "er_pct": 7.4, "inv_merchant": 0.0, "pct_inv_merchant": 0.0, "nps": 81.0},
            "D-21":  {"gmv": 194200.00, "meta_gmv": 220000.00, "ating_pct": 88.3, "pedidos": 1640, "aov": 118.41, "cap_fat": 93.0, "cancel_pct": 4.5, "ruptura_pct": 0.2, "er_pct": 7.2, "inv_merchant": 0.0, "pct_inv_merchant": 0.0, "nps": 81.7},
            "MTD":         {"gmv": 5663591.55, "meta_gmv": 6420000.00, "ating_pct": 88.2, "pedidos": 43200},
            "CONSOLIDADO": {"gmv": 6250000.00, "meta_gmv": 7100000.00, "ating_pct": 88.0, "pedidos": 47500}
        },
        "FESTVAL": {
            "D-1":   {"gmv": 78436.13, "meta_gmv": 65098.79, "ating_pct": 120.5, "pedidos": 493, "aov": 159.10, "cap_fat": 87.8, "cancel_pct": 5.0, "ruptura_pct": 2.2, "er_pct": 6.9, "inv_merchant": 5672.29, "pct_inv_merchant": 7.2, "nps": 70.9},
            "D-7":   {"gmv": 77209.27, "meta_gmv": 64000.00, "ating_pct": 120.6, "pedidos": 480, "aov": 160.85, "cap_fat": 87.0, "cancel_pct": 5.3, "ruptura_pct": 2.5, "er_pct": 7.1, "inv_merchant": 5100.00, "pct_inv_merchant": 6.6, "nps": 69.5},
            "D-15":  {"gmv": 74500.00, "meta_gmv": 63500.00, "ating_pct": 117.3, "pedidos": 468, "aov": 159.19, "cap_fat": 86.5, "cancel_pct": 5.5, "ruptura_pct": 2.8, "er_pct": 7.3, "inv_merchant": 4800.00, "pct_inv_merchant": 6.4, "nps": 68.8},
            "D-21":  {"gmv": 71000.00, "meta_gmv": 63000.00, "ating_pct": 112.7, "pedidos": 455, "aov": 156.04, "cap_fat": 86.0, "cancel_pct": 5.8, "ruptura_pct": 3.0, "er_pct": 7.5, "inv_merchant": 4500.00, "pct_inv_merchant": 6.3, "nps": 68.0},
            "MTD":         {"gmv": 1921225.61, "meta_gmv": 1880000.00, "ating_pct": 102.1, "pedidos": 12100},
            "CONSOLIDADO": {"gmv": 2050000.00, "meta_gmv": 2000000.00, "ating_pct": 102.5, "pedidos": 12900}
        },
        "JACOMAR": {
            "D-1":   {"gmv": 49435.12, "meta_gmv": 43813.55, "ating_pct": 112.8, "pedidos": 395, "aov": 125.15, "cap_fat": 91.8, "cancel_pct": 3.9, "ruptura_pct": 0.0, "er_pct": 9.6, "inv_merchant": 3092.63, "pct_inv_merchant": 6.3, "nps": 78.8},
            "D-7":   {"gmv": 46018.93, "meta_gmv": 43000.00, "ating_pct": 107.0, "pedidos": 378, "aov": 121.74, "cap_fat": 91.2, "cancel_pct": 4.1, "ruptura_pct": 0.1, "er_pct": 9.8, "inv_merchant": 2900.00, "pct_inv_merchant": 6.3, "nps": 77.5},
            "D-15":  {"gmv": 44200.00, "meta_gmv": 42500.00, "ating_pct": 104.0, "pedidos": 365, "aov": 121.10, "cap_fat": 90.5, "cancel_pct": 4.3, "ruptura_pct": 0.2, "er_pct": 10.0, "inv_merchant": 2700.00, "pct_inv_merchant": 6.1, "nps": 76.8},
            "D-21":  {"gmv": 43000.00, "meta_gmv": 42000.00, "ating_pct": 102.4, "pedidos": 355, "aov": 121.13, "cap_fat": 90.0, "cancel_pct": 4.5, "ruptura_pct": 0.2, "er_pct": 10.2, "inv_merchant": 2600.00, "pct_inv_merchant": 6.0, "nps": 76.0},
            "MTD":         {"gmv": 1253872.10, "meta_gmv": 1265000.00, "ating_pct": 99.0, "pedidos": 9850},
            "CONSOLIDADO": {"gmv": 1320000.00, "meta_gmv": 1300000.00, "ating_pct": 101.5, "pedidos": 10400}
        }
    },
    "lojas": {
        "NAGUMO": [
            {"nome": "NAGUMO EXPRESS ATIBAIA",       "gmv": 28450.00, "pedidos": 214, "ating_pct": 106.2, "cancel_pct": 3.8, "ruptura_pct": 0.0, "nps": 84.1},
            {"nome": "NAGUMO BRAGANÇA PAULISTA",     "gmv": 35600.00, "pedidos": 268, "ating_pct": 103.1, "cancel_pct": 4.0, "ruptura_pct": 0.2, "nps": 83.0},
            {"nome": "NAGUMO CAMPO LIMPO PAULISTA",  "gmv": 22100.00, "pedidos": 196, "ating_pct": 101.5, "cancel_pct": 4.5, "ruptura_pct": 0.1, "nps": 81.5},
            {"nome": "NAGUMO ITATIBA",               "gmv": 18900.00, "pedidos": 171, "ating_pct": 98.7, "cancel_pct": 5.1, "ruptura_pct": 0.0, "nps": 80.2},
            {"nome": "NAGUMO JUNDIAÍ",               "gmv": 41200.00, "pedidos": 305, "ating_pct": 110.3, "cancel_pct": 3.9, "ruptura_pct": 0.0, "nps": 85.0}
        ],
        "FESTVAL": [
            {"nome": "FESTVAL CURITIBA - BATEL",   "gmv": 31500.00, "pedidos": 198, "ating_pct": 122.0, "cancel_pct": 4.8, "ruptura_pct": 2.0, "nps": 72.1},
            {"nome": "FESTVAL CURITIBA - CABRAL",  "gmv": 25400.00, "pedidos": 162, "ating_pct": 118.5, "cancel_pct": 5.2, "ruptura_pct": 2.5, "nps": 70.5},
            {"nome": "FESTVAL CURITIBA - ÁGUA VERDE","gmv": 21536.13, "pedidos": 133, "ating_pct": 119.3, "cancel_pct": 5.3, "ruptura_pct": 2.1, "nps": 69.8}
        ],
        "JACOMAR": [
            {"nome": "JACOMAR CURITIBA - CIC",      "gmv": 16800.00, "pedidos": 134, "ating_pct": 114.0, "cancel_pct": 3.7, "ruptura_pct": 0.0, "nps": 79.5},
            {"nome": "JACOMAR CURITIBA - CAMPO COMPRIDO","gmv": 14200.00, "pedidos": 112, "ating_pct": 110.5, "cancel_pct": 4.1, "ruptura_pct": 0.1, "nps": 78.2},
            {"nome": "JACOMAR CURITIBA - HAUER",    "gmv": 11500.00, "pedidos": 89,  "ating_pct": 109.0, "cancel_pct": 4.0, "ruptura_pct": 0.0, "nps": 77.8},
            {"nome": "JACOMAR CURITIBA - CAJURU",   "gmv": 6935.12,  "pedidos": 60,  "ating_pct": 115.2, "cancel_pct": 3.6, "ruptura_pct": 0.0, "nps": 80.1}
        ]
    }
}


def load_data():
    """Carrega relatorio.json ou retorna dados de exemplo."""
    path = BASE_DIR / 'data' / 'relatorio.json'
    if path.exists():
        try:
            raw = json.loads(path.read_text(encoding='utf-8'))
            # Suporte ao formato antigo (migração transparente)
            if 'grupos' in raw and 'periodos' not in raw:
                raw = _migrate_old_format(raw)
            return raw
        except Exception as e:
            print(f"[WARN] Erro ao carregar relatorio.json: {e}")
    return SAMPLE_DATA


def _migrate_old_format(old):
    """Converte formato antigo (grupos/mtd) para novo (periodos/lojas)."""
    periodos = {}
    for grupo, kpis in old.get('grupos', {}).items():
        d1 = {k: v for k, v in kpis.items() if k != 'gmv_d7'}
        d7 = {k: v for k, v in d1.items()}
        if 'gmv_d7' in kpis:
            d7['gmv'] = kpis['gmv_d7']
        mtd_data = old.get('mtd', {}).get(grupo, {})
        periodos[grupo] = {
            'D-1': d1, 'D-7': d7,
            'D-15': {}, 'D-21': {},
            'MTD': {
                'gmv': mtd_data.get('gmv_atual', 0),
                'meta_gmv': 0,
                'ating_pct': mtd_data.get('ating_pct', 0),
                'pedidos': 0
            },
            'CONSOLIDADO': {}
        }
    return {
        'data_referencia': old.get('data_referencia', ''),
        'data_atualizacao': old.get('data_atualizacao', ''),
        'periodos': periodos,
        'lojas': {}
    }


# ---------------------------------------------------------------------------
# EXCEL GENERATION
# ---------------------------------------------------------------------------
def build_excel(data):
    wb = openpyxl.Workbook()

    HDR_FILL = PatternFill("solid", fgColor="EA1D2C")
    HDR_FONT = Font(bold=True, color="FFFFFF")
    SUB_FILL = PatternFill("solid", fgColor="1A1A2E")
    SUB_FONT = Font(bold=True, color="FFFFFF")
    center = Alignment(horizontal="center")
    right  = Alignment(horizontal="right")

    def style_header(cell, fill=HDR_FILL, font=HDR_FONT):
        cell.fill = fill
        cell.font = font
        cell.alignment = center

    periodos_list = ['D-1', 'D-7', 'D-15', 'D-21', 'MTD', 'CONSOLIDADO']
    grupos = list(data.get('periodos', {}).keys())

    # ── Aba Comparativo ────────────────────────────────────────────────
    ws1 = wb.active
    ws1.title = "Comparativo"
    ws1.column_dimensions['A'].width = 14
    for c in 'BCDEFG': ws1.column_dimensions[c].width = 18

    # Header row
    ws1.append(['Grupo / Período'] + periodos_list)
    for cell in ws1[1]: style_header(cell)

    for grupo in grupos:
        row_gmv     = [f"{grupo} — GMV"]
        row_pedidos = [f"{grupo} — Pedidos"]
        row_ating   = [f"{grupo} — Ating%"]
        for p in periodos_list:
            kpis = data['periodos'].get(grupo, {}).get(p, {})
            row_gmv.append(kpis.get('gmv', 0) or 0)
            row_pedidos.append(kpis.get('pedidos', 0) or 0)
            row_ating.append(kpis.get('ating_pct', 0) or 0)
        ws1.append(row_gmv)
        ws1.append(row_pedidos)
        ws1.append(row_ating)

    # ── Aba Por Loja ───────────────────────────────────────────────────
    ws2 = wb.create_sheet("Por Loja")
    ws2.column_dimensions['A'].width = 14
    ws2.column_dimensions['B'].width = 40
    for c in 'CDEFGH': ws2.column_dimensions[c].width = 16

    hdrs = ['Grupo','Loja','GMV','Pedidos','Ating%','Cancel%','Ruptura%','NPS']
    ws2.append(hdrs)
    for cell in ws2[1]: style_header(cell)

    for grupo, lojas in data.get('lojas', {}).items():
        for l in lojas:
            ws2.append([
                grupo, l.get('nome',''),
                l.get('gmv',0) or 0, l.get('pedidos',0) or 0,
                l.get('ating_pct',0) or 0, l.get('cancel_pct',0) or 0,
                l.get('ruptura_pct',0) or 0, l.get('nps',0) or 0
            ])

    # ── Aba MTD ────────────────────────────────────────────────────────
    ws3 = wb.create_sheet("MTD")
    ws3.column_dimensions['A'].width = 14
    for c in 'BCD': ws3.column_dimensions[c].width = 20

    ws3.append(['Grupo','GMV Acumulado','Meta GMV','Ating%'])
    for cell in ws3[1]: style_header(cell)

    for grupo in grupos:
        mtd = data.get('periodos', {}).get(grupo, {}).get('MTD', {})
        ws3.append([
            grupo,
            mtd.get('gmv', 0) or 0,
            mtd.get('meta_gmv', 0) or 0,
            mtd.get('ating_pct', 0) or 0
        ])

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ---------------------------------------------------------------------------
# FLASK ROUTES
# ---------------------------------------------------------------------------
@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)


@app.route('/api/dados')
def api_dados():
    return jsonify(load_data())


@app.route('/atualizar', methods=['POST'])
def atualizar():
    try:
        novo = request.get_json(force=True)
        if not novo:
            return jsonify({"error": "JSON inválido"}), 400
        path = BASE_DIR / 'data' / 'relatorio.json'
        path.parent.mkdir(exist_ok=True)
        path.write_text(json.dumps(novo, ensure_ascii=False, indent=2), encoding='utf-8')
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route('/download/excel')
def download_excel():
    if not OPENPYXL_OK:
        return jsonify({"error": "openpyxl não instalado"}), 500
    data = load_data()
    buf = build_excel(data)
    ref = data.get('data_referencia', 'relatorio').replace('/', '-')
    filename = f"relatorio_{ref}.xlsx"
    return send_file(
        buf,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename
    )


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    app.run(host='0.0.0.0', port=port, debug=False)
