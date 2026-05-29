"""
Daily Report — Carteira Rogério Salvador
Flask dashboard: 6 períodos, filtro por grupo, download Excel.
Dados reais via /api/dados → relatorio.json
"""
from pathlib import Path
from flask import Flask, render_template_string, jsonify, request, send_file
import json, os, io

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    OPENPYXL_OK = True
except ImportError:
    OPENPYXL_OK = False

BASE_DIR = Path(__file__).parent.absolute()
app = Flask(__name__)

# ---------------------------------------------------------------------------
# SAMPLE DATA (used when relatorio.json not present)
# ---------------------------------------------------------------------------
SAMPLE_DATA = {
    "data_referencia": "27/05/2026",
    "data_atualizacao": "28/05/2026 00:00",
    "periodos": {
        "NAGUMO": {
            "D-1":  {"gmv": 222983.96, "meta_gmv": 221977.84, "ating_pct": 100.5, "pedidos": 1747, "aov": 127.64, "cancel_pct": 6.1, "ruptura_pct": 21.0},
            "D-7":  {"gmv": 165039.56, "meta_gmv": 210000.00, "ating_pct": 78.6,  "pedidos": 1201, "aov": 137.42, "cancel_pct": 4.5, "ruptura_pct": 20.6},
            "D-15": {"gmv": 159816.04, "meta_gmv": 218000.00, "ating_pct": 73.3,  "pedidos": 1193, "aov": 133.96, "cancel_pct": 3.7, "ruptura_pct": 23.8},
            "D-21": {"gmv": 168792.19, "meta_gmv": 220000.00, "ating_pct": 76.7,  "pedidos": 1168, "aov": 144.51, "cancel_pct": 4.9, "ruptura_pct": 25.9},
            "MTD":  {"gmv": 5866597.87, "meta_gmv": 6420000.00, "ating_pct": 91.4, "pedidos": 41154},
            "CONSOLIDADO": {"gmv": 5866597.87, "meta_gmv": 6420000.00, "ating_pct": 91.4, "pedidos": 41154}
        },
        "FESTVAL": {
            "D-1":  {"gmv": 65828.35, "meta_gmv": 65098.79, "ating_pct": 101.1, "pedidos": 412, "aov": 159.78, "cancel_pct": 4.2, "ruptura_pct": 41.7},
            "D-7":  {"gmv": 64344.78, "meta_gmv": 64000.00, "ating_pct": 100.5, "pedidos": 372, "aov": 172.97, "cancel_pct": 2.9, "ruptura_pct": 39.0},
            "D-15": {"gmv": 63794.95, "meta_gmv": 63500.00, "ating_pct": 100.5, "pedidos": 394, "aov": 161.92, "cancel_pct": 3.9, "ruptura_pct": 41.9},
            "D-21": {"gmv": 68500.80, "meta_gmv": 63000.00, "ating_pct": 108.7, "pedidos": 384, "aov": 178.39, "cancel_pct": 3.0, "ruptura_pct": 42.7},
            "MTD":  {"gmv": 1988824.75, "meta_gmv": 1880000.00, "ating_pct": 105.8, "pedidos": 12316},
            "CONSOLIDADO": {"gmv": 1988824.75, "meta_gmv": 1880000.00, "ating_pct": 105.8, "pedidos": 12316}
        },
        "JACOMAR": {
            "D-1":  {"gmv": 41827.61, "meta_gmv": 43813.55, "ating_pct": 95.5, "pedidos": 349, "aov": 119.85, "cancel_pct": 2.0, "ruptura_pct": 23.8},
            "D-7":  {"gmv": 43891.72, "meta_gmv": 43000.00, "ating_pct": 102.1, "pedidos": 371, "aov": 118.31, "cancel_pct": 2.4, "ruptura_pct": 32.3},
            "D-15": {"gmv": 40656.83, "meta_gmv": 42500.00, "ating_pct": 95.7, "pedidos": 358, "aov": 113.57, "cancel_pct": 2.2, "ruptura_pct": 34.9},
            "D-21": {"gmv": 46077.00, "meta_gmv": 42000.00, "ating_pct": 109.7, "pedidos": 379, "aov": 121.58, "cancel_pct": 3.8, "ruptura_pct": 30.3},
            "MTD":  {"gmv": 1295642.49, "meta_gmv": 1265000.00, "ating_pct": 102.4, "pedidos": 10605},
            "CONSOLIDADO": {"gmv": 1295642.49, "meta_gmv": 1265000.00, "ating_pct": 102.4, "pedidos": 10605}
        }
    },
    "lojas": {"NAGUMO": {"D-1":[],"D-7":[],"D-15":[],"D-21":[],"MTD":[],"CONSOLIDADO":[]},
              "FESTVAL":{"D-1":[],"D-7":[],"D-15":[],"D-21":[],"MTD":[],"CONSOLIDADO":[]},
              "JACOMAR":{"D-1":[],"D-7":[],"D-15":[],"D-21":[],"MTD":[],"CONSOLIDADO":[]}}
}

def load_data():
    path = BASE_DIR / 'data' / 'relatorio.json'
    if path.exists():
        try:
            return json.loads(path.read_text(encoding='utf-8'))
        except Exception as e:
            print(f"[WARN] relatorio.json error: {e}")
    return SAMPLE_DATA

# ---------------------------------------------------------------------------
# HTML TEMPLATE
# ---------------------------------------------------------------------------
HTML = r"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Relatório Diário — Carteira Rogério Salvador</title>
<style>
  :root {
    --bg:#0F0F1A; --card:#1A1A2E; --card2:#16213E;
    --red:#EA1D2C; --red-dk:#B71C2B;
    --green:#27AE60; --yellow:#F39C12; --alert:#E74C3C;
    --txt:#FFFFFF; --txt2:#B0BEC5; --border:#2C3E50; --hover:#22304a;
  }
  *{box-sizing:border-box;margin:0;padding:0}
  body{background:var(--bg);color:var(--txt);
       font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;
       font-size:14px;min-height:100vh;padding:16px}
  .container{max-width:1280px;margin:0 auto}

  /* HEADER */
  .header{background:var(--card);border:1px solid var(--border);
          border-radius:12px;padding:18px;margin-bottom:16px}
  .header-top{display:flex;align-items:center;justify-content:space-between;
              flex-wrap:wrap;gap:10px;margin-bottom:10px}
  .logo{font-size:22px;font-weight:800;color:var(--red);letter-spacing:.5px}
  .update-badge{background:var(--red-dk);color:#fff;padding:5px 13px;
                border-radius:20px;font-size:12px;font-weight:600}
  .header-subtitle{color:var(--txt2);font-size:13px}
  .ref-date{font-weight:600;color:var(--txt)}
  .header-actions{display:flex;gap:8px;flex-wrap:wrap;margin-top:12px}
  .btn{display:inline-flex;align-items:center;gap:6px;padding:8px 14px;
       border:none;border-radius:8px;font-size:13px;font-weight:600;
       cursor:pointer;text-decoration:none;transition:opacity .15s}
  .btn-red{background:var(--red);color:#fff}
  .btn-outline{background:transparent;color:var(--txt2);border:1px solid var(--border)}
  .btn:hover{opacity:.85}

  /* TABS */
  .group-tabs{display:flex;gap:8px;margin-bottom:14px;flex-wrap:wrap}
  .group-tab{padding:10px 22px;border-radius:10px;border:2px solid var(--border);
             background:var(--card);color:var(--txt2);
             cursor:pointer;font-weight:700;font-size:15px;transition:all .15s}
  .group-tab.active{border-color:var(--red);color:#fff;background:var(--red-dk)}

  /* PERIOD PILLS */
  .period-bar{display:flex;gap:6px;flex-wrap:wrap;margin-bottom:16px}
  .period-pill{padding:6px 16px;border-radius:20px;border:1px solid var(--border);
               background:var(--card);color:var(--txt2);
               cursor:pointer;font-size:13px;font-weight:600;transition:all .15s}
  .period-pill.active{background:var(--red);color:#fff;border-color:var(--red)}

  /* KPI CARDS */
  .kpi-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));
            gap:12px;margin-bottom:20px}
  .kpi-card{background:var(--card);border:1px solid var(--border);
            border-radius:10px;padding:14px}
  .kpi-label{font-size:11px;color:var(--txt2);text-transform:uppercase;
             letter-spacing:.5px;margin-bottom:6px}
  .kpi-value{font-size:20px;font-weight:800;color:var(--txt)}
  .kpi-value.big{font-size:26px;color:var(--red)}
  .kpi-badge{display:inline-block;padding:3px 9px;border-radius:5px;
             font-weight:700;font-size:14px;margin-left:6px}
  .kpi-badge.verde{background:var(--green);color:#fff}
  .kpi-badge.amarelo{background:var(--yellow);color:#fff}
  .kpi-badge.vermelho{background:var(--alert);color:#fff}
  .kpi-delta{margin-top:4px;font-size:12px}
  .kpi-delta.up{color:var(--green)}.kpi-delta.down{color:var(--alert)}.kpi-delta.neutral{color:var(--txt2)}

  /* SECTION / TABLE */
  .section{background:var(--card);border:1px solid var(--border);
           border-radius:12px;padding:18px;margin-bottom:20px}
  .section-title{font-size:16px;font-weight:700;color:var(--red);margin-bottom:14px}
  .table-wrap{overflow-x:auto}
  table{width:100%;border-collapse:collapse}
  thead{background:var(--card2)}
  th{text-align:left;padding:10px 12px;font-size:11px;text-transform:uppercase;
     letter-spacing:.5px;color:var(--txt2);font-weight:600;white-space:nowrap}
  td{padding:10px 12px;border-top:1px solid var(--border);font-size:13px}
  tbody tr:hover{background:var(--hover)}
  .val-pos{color:var(--green);font-weight:600}
  .val-neg{color:var(--alert);font-weight:600}
  .val-warn{color:var(--yellow);font-weight:600}
  .semaforo{width:10px;height:10px;border-radius:50%;display:inline-block;margin-right:6px}
  .sem-v{background:var(--green);box-shadow:0 0 6px var(--green)}
  .sem-a{background:var(--yellow);box-shadow:0 0 6px var(--yellow)}
  .sem-r{background:var(--alert);box-shadow:0 0 6px var(--alert)}

  @media print{
    .header-actions,.group-tabs,.period-bar{display:none!important}
    body{background:#fff;color:#000;padding:0}
    .kpi-card,.section{border:1px solid #ccc;background:#fff}
    .kpi-value,.section-title{color:#000}
  }
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
    <div class="header-actions">
      <a class="btn btn-red" href="/download/excel" download>⬇ Download Excel</a>
      <button class="btn btn-outline" onclick="window.print()">🖨 Imprimir</button>
      <button class="btn btn-outline" onclick="init()">↺ Atualizar</button>
    </div>
  </div>

  <!-- GROUP TABS -->
  <div class="group-tabs" id="groupTabs"></div>

  <!-- PERIOD PILLS -->
  <div class="period-bar" id="periodBar"></div>

  <!-- KPI CARDS -->
  <div class="kpi-grid" id="kpiGrid"></div>

  <!-- Charts -->
  <div class="section">
    <div class="section-title">📊 Evolução GMV por Período</div>
    <div id="chartGmv"></div>
  </div>
  <div class="section">
    <div class="section-title">🎯 Atingimento% por Período</div>
    <div id="chartAting"></div>
  </div>
  <div class="section">
    <div class="section-title">🏆 Top Lojas por GMV (Período Ativo)</div>
    <div id="chartTop"></div>
  </div>

  <!-- STORES TABLE -->
  <div class="section">
    <div class="section-title">🏪 Lojas do Grupo — <span id="storesPeriodLabel"></span></div>
    <div class="table-wrap">
      <table>
        <thead id="storesThead"></thead>
        <tbody id="storesTbody"></tbody>
      </table>
    </div>
  </div>

</div>
<script>
// ── State ───────────────────────────────────────────────────────────────────
let DATA = null;
const PERIODS = ['D-1','D-7','D-15','D-21','MTD','CONSOLIDADO'];
const PERIOD_LABELS = {'D-1':'D-1','D-7':'D-7','D-15':'D-15','D-21':'D-21','MTD':'MTD','CONSOLIDADO':'Consolidado'};
let currentGroup = null;
let currentPeriod = 'D-1';

// ── Formatters ──────────────────────────────────────────────────────────────
const fmtBRL = v => v == null ? 'R$ —' :
  'R$ ' + Number(v).toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2});
const fmtPct = v => v == null ? '—' : Number(v).toLocaleString('pt-BR',{minimumFractionDigits:1,maximumFractionDigits:1})+'%';
const fmtNum = v => v == null ? '—' : Number(v).toLocaleString('pt-BR');
const semClass  = a => a >= 100 ? 'sem-v' : a >= 85 ? 'sem-a' : 'sem-r';
const badgeCls  = a => a >= 100 ? 'verde' : a >= 85 ? 'amarelo' : 'vermelho';

function delta(val, prev, invertBad) {
  if (!prev || prev === 0 || val == null) return '';
  const pct = ((val / prev) - 1) * 100;
  const up = pct >= 0;
  const cls = invertBad ? (up?'down':'up') : (up?'up':'down');
  const arrow = up ? '▲' : '▼';
  return `<div class="kpi-delta ${cls}">${arrow} ${Math.abs(pct).toLocaleString('pt-BR',{minimumFractionDigits:1,maximumFractionDigits:1})}% vs período ant.</div>`;
}

// ── Group tabs ──────────────────────────────────────────────────────────────
function renderGroupTabs() {
  const groups = Object.keys(DATA.periodos || {});
  document.getElementById('groupTabs').innerHTML = groups.map(g =>
    `<div class="group-tab ${g===currentGroup?'active':''}" onclick="selectGroup('${g}')">${g}</div>`
  ).join('');
}
function selectGroup(g) {
  currentGroup = g;
  renderGroupTabs();
  renderKPIs();
  renderStores();   // ← tabela muda para dados do grupo + período atual
  renderAllCharts();
}

// ── Period pills ────────────────────────────────────────────────────────────
function renderPeriodBar() {
  document.getElementById('periodBar').innerHTML = PERIODS.map(p =>
    `<div class="period-pill ${p===currentPeriod?'active':''}" onclick="selectPeriod('${p}')">${PERIOD_LABELS[p]}</div>`
  ).join('');
}
function selectPeriod(p) {
  currentPeriod = p;
  renderPeriodBar();
  renderKPIs();
  renderStores();   // ← tabela muda para dados de DATA.lojas[grupo][p]
  renderAllCharts();
}

// ── KPI Cards ───────────────────────────────────────────────────────────────
function prevPeriod(p) {
  const idx = PERIODS.indexOf(p);
  return idx > 0 ? PERIODS[idx-1] : null;
}
function renderKPIs() {
  const periodos = (DATA.periodos||{})[currentGroup]||{};
  const cur  = periodos[currentPeriod]||{};
  const prev = periodos[prevPeriod(currentPeriod)]||{};
  const isMTD = currentPeriod==='MTD'||currentPeriod==='CONSOLIDADO';
  const ating = cur.ating_pct||0;
  const cards = [];

  cards.push(`<div class="kpi-card">
    <div class="kpi-label">GMV</div>
    <div class="kpi-value big">${fmtBRL(cur.gmv)}
      ${cur.ating_pct!=null?`<span class="kpi-badge ${badgeCls(ating)}">${fmtPct(ating)}</span>`:''}
    </div>
    ${delta(cur.gmv, prev.gmv, false)}
  </div>`);

  cards.push(`<div class="kpi-card">
    <div class="kpi-label">Meta GMV</div>
    <div class="kpi-value">${fmtBRL(cur.meta_gmv)}</div>
  </div>`);

  cards.push(`<div class="kpi-card">
    <div class="kpi-label">Pedidos</div>
    <div class="kpi-value">${fmtNum(cur.pedidos)}</div>
    ${delta(cur.pedidos, prev.pedidos, false)}
  </div>`);

  if (!isMTD) {
    cards.push(`<div class="kpi-card">
      <div class="kpi-label">AOV</div>
      <div class="kpi-value">${cur.aov?fmtBRL(cur.aov):'—'}</div>
      ${delta(cur.aov, prev.aov, false)}
    </div>`);

    const cancelCls = cur.cancel_pct>7?'val-neg':cur.cancel_pct>5?'val-warn':'';
    cards.push(`<div class="kpi-card">
      <div class="kpi-label">Cancelamento</div>
      <div class="kpi-value ${cancelCls}">${fmtPct(cur.cancel_pct)}</div>
      ${delta(cur.cancel_pct, prev.cancel_pct, true)}
    </div>`);

    const ruptCls = cur.ruptura_pct>3?'val-neg':cur.ruptura_pct>1?'val-warn':'';
    cards.push(`<div class="kpi-card">
      <div class="kpi-label">Ruptura</div>
      <div class="kpi-value ${ruptCls}">${fmtPct(cur.ruptura_pct)}</div>
      ${delta(cur.ruptura_pct, prev.ruptura_pct, true)}
    </div>`);

    if (cur.nps != null) {
      const npsCls = cur.nps>=70?'val-pos':cur.nps>=50?'val-warn':'val-neg';
      cards.push(`<div class="kpi-card">
        <div class="kpi-label">NPS</div>
        <div class="kpi-value ${npsCls}">${Number(cur.nps).toLocaleString('pt-BR',{minimumFractionDigits:1})}</div>
      </div>`);
    }
  }
  document.getElementById('kpiGrid').innerHTML = cards.join('');
}

// ── Stores table — USES DATA.lojas[currentGroup][currentPeriod] ─────────────
function renderStores() {
  // This is the KEY fix: each period shows its own distinct store data
  const lojas = ((DATA.lojas||{})[currentGroup]||{})[currentPeriod]||[];
  const label = PERIOD_LABELS[currentPeriod]||currentPeriod;
  document.getElementById('storesPeriodLabel').textContent = label;

  const thead = document.getElementById('storesThead');
  const tbody = document.getElementById('storesTbody');

  if (!lojas.length) {
    thead.innerHTML = '';
    tbody.innerHTML = '<tr><td colspan="7" style="color:var(--txt2);text-align:center;padding:20px">Sem dados de lojas para este período</td></tr>';
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
    const ating = l.ating_pct||0;
    const sem = semClass(ating);
    const cCls = l.cancel_pct>7?'val-neg':l.cancel_pct>5?'val-warn':'';
    const rCls = l.ruptura_pct>3?'val-neg':l.ruptura_pct>1?'val-warn':'';
    const nCls = l.nps!=null?(l.nps>=70?'val-pos':l.nps>=50?'val-warn':'val-neg'):'';
    return `<tr>
      <td><span class="semaforo ${sem}"></span>${l.nome}</td>
      <td style="text-align:right">${fmtBRL(l.gmv)}</td>
      <td style="text-align:right">${fmtNum(l.pedidos)}</td>
      <td style="text-align:right"><span class="kpi-badge ${badgeCls(ating)}" style="font-size:11px">${fmtPct(ating)}</span></td>
      <td style="text-align:right" class="${cCls}">${l.cancel_pct!=null?fmtPct(l.cancel_pct):'—'}</td>
      <td style="text-align:right" class="${rCls}">${l.ruptura_pct!=null?fmtPct(l.ruptura_pct):'—'}</td>
      <td style="text-align:right" class="${nCls}">${l.nps!=null?Number(l.nps).toLocaleString('pt-BR',{minimumFractionDigits:1}):'—'}</td>
    </tr>`;
  }).join('');
}

// ── Charts ──────────────────────────────────────────────────────────────────
function renderAllCharts() {
  const periodos = (DATA.periodos||{})[currentGroup]||{};
  const max = Math.max(...PERIODS.map(p => parseFloat((periodos[p]||{}).gmv)||0), 1);

  // GMV chart
  document.getElementById('chartGmv').innerHTML = PERIODS.map(p => {
    const val = parseFloat((periodos[p]||{}).gmv)||0;
    const pct = (val/max*100).toFixed(1);
    const isActive = p===currentPeriod;
    return `<div style="display:flex;align-items:center;margin:6px 0;gap:10px">
      <span style="width:110px;font-size:12px;color:${isActive?'#fff':'#B0BEC5'};font-weight:${isActive?700:400}">${PERIOD_LABELS[p]}</span>
      <div style="flex:1;background:#1A1A2E;border-radius:4px;height:24px;overflow:hidden">
        <div style="width:${pct}%;background:${isActive?'#EA1D2C':'#4a4a6a'};height:100%;border-radius:4px;transition:width .3s"></div>
      </div>
      <span style="width:150px;text-align:right;font-size:12px;color:#FFF">R$ ${val.toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2})}</span>
    </div>`;
  }).join('');

  // Ating chart
  document.getElementById('chartAting').innerHTML = PERIODS.map(p => {
    const val = parseFloat((periodos[p]||{}).ating_pct)||0;
    const cor = val>=100?'#27AE60':val>=85?'#F39C12':'#E74C3C';
    const isActive = p===currentPeriod;
    return `<div style="display:flex;align-items:center;margin:6px 0;gap:10px">
      <span style="width:110px;font-size:12px;color:${isActive?'#fff':'#B0BEC5'};font-weight:${isActive?700:400}">${PERIOD_LABELS[p]}</span>
      <div style="flex:1;background:#1A1A2E;border-radius:4px;height:24px;position:relative;overflow:hidden">
        <div style="width:${Math.min(val/150*100,100).toFixed(1)}%;background:${isActive?cor:cor+'99'};height:100%;border-radius:4px"></div>
        <div style="position:absolute;left:0;top:0;height:100%;width:${(100/150*100).toFixed(1)}%;border-right:2px dashed #555"></div>
      </div>
      <span style="width:80px;text-align:right;font-size:12px;color:${cor};font-weight:700">${val.toFixed(1)}%</span>
    </div>`;
  }).join('');

  // Top lojas chart — uses current period store data
  const lojas = ((DATA.lojas||{})[currentGroup]||{})[currentPeriod]||[];
  const el = document.getElementById('chartTop');
  if (!lojas.length) {
    el.innerHTML = '<p style="color:#666;text-align:center;padding:10px">Sem dados para este período</p>';
    return;
  }
  const top10 = [...lojas].sort((a,b)=>(b.gmv||0)-(a.gmv||0)).slice(0,10);
  const maxGmv = Math.max(...top10.map(l=>l.gmv||0),1);
  el.innerHTML = top10.map(l => {
    const val = l.gmv||0;
    const pct = (val/maxGmv*100).toFixed(1);
    const ating = l.ating_pct||0;
    const cor = ating>=100?'#27AE60':ating>=85?'#F39C12':'#E74C3C';
    return `<div style="display:flex;align-items:center;margin:5px 0;gap:10px">
      <span style="width:200px;font-size:11px;color:#B0BEC5;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${l.nome}">${l.nome}</span>
      <div style="flex:1;background:#1A1A2E;border-radius:4px;height:22px;overflow:hidden">
        <div style="width:${pct}%;background:${cor};height:100%;border-radius:4px;opacity:.85"></div>
      </div>
      <span style="width:150px;text-align:right;font-size:11px;color:#FFF">R$ ${val.toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2})}</span>
    </div>`;
  }).join('');
}

// ── Boot ─────────────────────────────────────────────────────────────────────
async function init() {
  const res = await fetch('/api/dados');
  DATA = await res.json();
  const groups = Object.keys(DATA.periodos||{});
  if (!currentGroup || !groups.includes(currentGroup)) currentGroup = groups[0];
  renderGroupTabs();
  renderPeriodBar();
  renderKPIs();
  renderStores();
  renderAllCharts();
  document.getElementById('updateBadge').textContent = 'Atualizado: '+(DATA.data_atualizacao||'');
  document.getElementById('refDate').textContent = DATA.data_referencia||'';
}
init();
</script>
</body>
</html>"""

# ---------------------------------------------------------------------------
# EXCEL
# ---------------------------------------------------------------------------
def build_excel(data):
    wb = openpyxl.Workbook()
    HDR = PatternFill("solid", fgColor="EA1D2C")
    HFONT = Font(bold=True, color="FFFFFF")
    center = Alignment(horizontal="center")
    periodos_list = ['D-1','D-7','D-15','D-21','MTD','CONSOLIDADO']
    grupos = list(data.get('periodos',{}).keys())

    ws1 = wb.active; ws1.title = "Comparativo"
    ws1.column_dimensions['A'].width = 24
    for c in 'BCDEFG': ws1.column_dimensions[c].width = 18
    ws1.append(['Grupo / Período']+periodos_list)
    for cell in ws1[1]:
        cell.fill=HDR; cell.font=HFONT; cell.alignment=center

    for g in grupos:
        for metric, label in [('gmv','GMV'),('pedidos','Pedidos'),('ating_pct','Ating%')]:
            row = [f"{g} — {label}"]
            for p in periodos_list:
                row.append(data['periodos'].get(g,{}).get(p,{}).get(metric,0) or 0)
            ws1.append(row)

    ws2 = wb.create_sheet("Por Loja — D-1")
    ws2.column_dimensions['A'].width = 14; ws2.column_dimensions['B'].width = 40
    for c in 'CDEFGH': ws2.column_dimensions[c].width = 16
    ws2.append(['Grupo','Loja','GMV','Pedidos','Cancel%','Ruptura%','NPS'])
    for cell in ws2[1]:
        cell.fill=HDR; cell.font=HFONT; cell.alignment=center

    for g, pdata in data.get('lojas',{}).items():
        for l in (pdata.get('D-1') or []):
            ws2.append([g,l.get('nome',''),l.get('gmv',0),l.get('pedidos',0),
                        l.get('cancel_pct'),l.get('ruptura_pct'),l.get('nps')])

    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf

# ---------------------------------------------------------------------------
# ROUTES
# ---------------------------------------------------------------------------
@app.route('/')
def index():
    return render_template_string(HTML)

@app.route('/api/dados')
def api_dados():
    return jsonify(load_data())

@app.route('/atualizar', methods=['POST'])
def atualizar():
    try:
        novo = request.get_json(force=True)
        if not novo:
            return jsonify({"error":"JSON inválido"}), 400
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
    ref = data.get('data_referencia','relatorio').replace('/','-')
    return send_file(buf,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True, download_name=f"relatorio_{ref}.xlsx")

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    app.run(host='0.0.0.0', port=port, debug=False)
