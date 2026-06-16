"""Daily Report Rogério Salvador — Dashboard inline completo (render_template_string)"""
import json, os
from pathlib import Path
from datetime import datetime, date, timedelta
from flask import Flask, jsonify, request, render_template_string

app = Flask(__name__)
BASE_DIR = Path(__file__).parent
# Try multiple locations for data file
DATA_FILE = BASE_DIR / "data" / "relatorio.json"
DATA_FILE_ALT = BASE_DIR / "relatorio.json"


def load_data():
    for f in [DATA_FILE, DATA_FILE_ALT]:
        if f.exists():
            return json.loads(f.read_text(encoding="utf-8"))
    return {}


GRUPOS = ["NAGUMO", "FESTVAL", "JACOMAR"]
PERIODOS = ["D-1", "D-7", "D-15", "D-21", "MTD", "CONSOLIDADO"]


def semaforo(kpi, val):
    if val is None:
        return "gray"
    try:
        v = float(val)
    except Exception:
        return "gray"
    bad = {"cancel_pct": (5, 8), "ruptura_pct": (10, 25), "er_pct": (10, 20)}
    good = {"ating_pct": (90, 80), "nps": (70, 50), "sla_pct": (90, 70), "online_pct": (95, 80)}
    if kpi in bad:
        lo, hi = bad[kpi]
        return "green" if v < lo else ("yellow" if v < hi else "red")
    if kpi in good:
        hi, lo = good[kpi]
        return "green" if v >= hi else ("yellow" if v >= lo else "red")
    return "gray"


def fmt_brl(v):
    try:
        return f"R$ {float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "—"


def fmt_pct(v, d=1):
    try:
        return f"{float(v):.{d}f}%"
    except Exception:
        return "—"



def dados_desatualizados(data: dict) -> bool:
    """Gate 3: Retorna True se data_referencia for anterior a ontem."""
    try:
        ref = datetime.strptime(data.get('data_referencia', ''), '%d/%m/%Y').date()
        ontem = date.today() - timedelta(days=1)
        return ref < ontem
    except Exception:
        return True

TEMPLATE = r"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Carteira Rogério | Daily</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{background:#0F0F1A;color:#e0e0e0;font-family:system-ui,sans-serif;min-height:100vh}
.header{background:#1A1A2E;border-bottom:3px solid #EA1D2C;padding:1rem 1.5rem;display:flex;align-items:center;gap:1rem;flex-wrap:wrap}
.header h1{font-size:1.25rem;color:#fff;font-weight:700}
.badge{background:#EA1D2C;color:#fff;padding:.25rem .65rem;border-radius:20px;font-size:.78rem;font-weight:600}
.badge-ok{background:#16a34a}
.group-bar{display:flex;gap:.6rem;padding:.8rem 1.5rem;background:#12122a;flex-wrap:wrap;align-items:center}
.group-btn{padding:.45rem 1.1rem;border-radius:20px;cursor:pointer;border:2px solid #EA1D2C;font-size:.88rem;font-weight:700;color:#EA1D2C;background:transparent;transition:all .2s}
.group-btn.active{background:#EA1D2C;color:#fff}
.group-btn:hover:not(.active){background:rgba(234,29,44,.15)}
.container{padding:1.5rem;max-width:1400px;margin:0 auto}
.section-title{color:#EA1D2C;font-size:.95rem;font-weight:700;margin:1.5rem 0 .75rem;text-transform:uppercase;letter-spacing:.06em}
.cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(290px,1fr));gap:1rem;margin-bottom:1.5rem}
.card{background:#1A1A2E;border-radius:10px;padding:1.25rem;border:1px solid #2a2a4a}
.card-title{font-size:1rem;font-weight:700;color:#fff;margin-bottom:.75rem;border-bottom:2px solid #EA1D2C;padding-bottom:.4rem}
.card-row{display:flex;justify-content:space-between;align-items:center;padding:.22rem 0;font-size:.85rem}
.card-row .label{color:#888}
.card-row .val{font-weight:600;color:#e0e0e0}
.gmv-main{font-size:1.35rem;font-weight:700;color:#EA1D2C}
.pill{display:inline-block;padding:.18rem .55rem;border-radius:10px;font-size:.78rem;font-weight:700}
.pill-green{background:#15803d;color:#d1fae5}
.pill-yellow{background:#92400e;color:#fef3c7}
.pill-red{background:#991b1b;color:#fee2e2}
.pill-gray{background:#374151;color:#d1d5db}
.dot{display:inline-block;width:9px;height:9px;border-radius:50%;margin-right:4px;vertical-align:middle}
.dot-green{background:#16a34a}
.dot-yellow{background:#ca8a04}
.dot-red{background:#dc2626}
.dot-gray{background:#6b7280}
.table-wrap{overflow-x:auto;border-radius:8px;margin-bottom:1.5rem}
table{width:100%;border-collapse:collapse;font-size:.82rem}
th{background:#1e1e3a;color:#aaa;padding:.5rem .7rem;text-align:left;white-space:nowrap;border-bottom:2px solid #EA1D2C;font-size:.8rem;text-transform:uppercase}
td{padding:.4rem .7rem;border-bottom:1px solid #1a1a2e;white-space:nowrap;color:#d1d5db}
tr:hover td{background:#1a1a2e}
.selects{display:flex;gap:.75rem;margin-bottom:.75rem;flex-wrap:wrap;align-items:center}
select{background:#1e1e3a;color:#e0e0e0;border:1px solid #333;padding:.4rem .7rem;border-radius:6px;font-size:.85rem;cursor:pointer}
.no-data{color:#f87171;margin-top:2rem;text-align:center;padding:2rem}
.comp-section{display:none}
.comp-section.active{display:block}
</style>
</head>
<body>
{% if desatualizado %}
<div style="background:#B71C1C;color:#fff;text-align:center;padding:8px;font-size:13px;font-weight:600;">
  ⚠️ Dados desatualizados — última referência: {{ data.get('data_referencia','N/A') }}. Atualização prevista para hoje.
</div>
{% endif %}
<div class="header">
  <h1>📊 Carteira Rogério Salvador</h1>
  {% if data %}
  <span class="badge">D-1: {{ data.get('data_referencia','—') }}</span>
  <span class="badge badge-ok">✅ Atualizado: {{ data.get('data_atualizacao','—') }}</span>
  {% else %}
  <span class="badge">Sem dados</span>
  {% endif %}
</div>

{% if not data %}
<div class="no-data">⚠️ Dados não disponíveis. POST /atualizar para inserir dados.</div>
{% else %}

<div class="group-bar">
  <span style="color:#888;font-size:.82rem;margin-right:.3rem">GRUPO:</span>
  {% for g in grupos %}
  <button class="group-btn {% if loop.first %}active{% endif %}" onclick="selectGroup('{{ g }}')">{{ g }}</button>
  {% endfor %}
</div>

<div class="container">

<div class="section-title">Resumo D-1 — <span id="grupo-label">{{ grupos[0] }}</span></div>
<div class="cards" id="resumo-cards">
{% for g in grupos %}
{% set gd = periodos_data.get(g, {}).get('D-1', {}) %}
<div class="card" data-grupo="{{ g }}" {% if not loop.first %}style="display:none"{% endif %}>
  <div class="card-title">{{ g }}</div>
  <div class="card-row"><span class="label">GMV D-1</span><span class="gmv-main">{{ fmt_brl(gd.get('gmv')) }}</span></div>
  <div class="card-row"><span class="label">Meta GMV</span><span class="val">{{ fmt_brl(gd.get('meta_gmv')) }}</span></div>
  <div class="card-row"><span class="label">Atingimento</span><span class="pill pill-{{ semaforo('ating_pct', gd.get('ating_pct')) }}">{{ fmt_pct(gd.get('ating_pct')) }}</span></div>
  <div class="card-row"><span class="label">Pedidos</span><span class="val">{{ gd.get('pedidos','—') }}</span></div>
  <div class="card-row"><span class="label">AOV</span><span class="val">{{ fmt_brl(gd.get('aov')) }}</span></div>
  <div class="card-row"><span class="label">Cancel%</span><span class="val"><span class="dot dot-{{ semaforo('cancel_pct', gd.get('cancel_pct')) }}"></span>{{ fmt_pct(gd.get('cancel_pct')) }}</span></div>
  <div class="card-row"><span class="label">Ruptura%</span><span class="val"><span class="dot dot-{{ semaforo('ruptura_pct', gd.get('ruptura_pct')) }}"></span>{{ fmt_pct(gd.get('ruptura_pct')) }}</span></div>
  <div class="card-row"><span class="label">NSU%</span><span class="val"><span class="dot dot-gray"></span>{{ fmt_pct(gd.get('nsu_pct')) }}</span></div>
  <div class="card-row"><span class="label">SLA%</span><span class="val"><span class="dot dot-{{ semaforo('sla_pct', gd.get('sla_pct')) }}"></span>{{ fmt_pct(gd.get('sla_pct')) }}</span></div>
  <div class="card-row"><span class="label">Online%</span><span class="val"><span class="dot dot-{{ semaforo('online_pct', gd.get('online_pct')) }}"></span>{{ fmt_pct(gd.get('online_pct')) }}</span></div>
  <div class="card-row"><span class="label">ER%</span><span class="val">{{ fmt_pct(gd.get('er_pct')) }}</span></div>
</div>
{% endfor %}
</div>

<div class="section-title" style="margin-top:1.5rem">MTD — Comparação MoM (Jun vs Mai mesmo período)</div>
<div class="cards" id="mtd-cards">
{% for g in grupos %}
{% set mtd = periodos_data.get(g, {}).get('MTD', {}) %}
<div class="card" data-grupo="{{ g }}" {% if not loop.first %}style="display:none"{% endif %}>
  <div class="card-title">{{ g }} — MTD Jun/25 vs Mai/25</div>
  <div class="card-row"><span class="label">GMV Jun</span><span class="gmv-main">{{ fmt_brl(mtd.get('gmv')) }}</span></div>
  <div class="card-row"><span class="label">GMV Mai (prior)</span><span class="val" style="color:#888">{{ fmt_brl(mtd.get('gmv_prior')) }}</span></div>
  <div class="card-row"><span class="label">Var% GMV MoM</span>
    {% set var = mtd.get('gmv_var_pct') %}
    {% if var is not none %}
      <span class="pill {% if var >= 0 %}pill-green{% else %}pill-red{% endif %}">{{ '+' if var > 0 else '' }}{{ var }}%</span>
    {% else %}<span class="val">—</span>{% endif %}
  </div>
  <div class="card-row"><span class="label">Pedidos Jun</span><span class="val">{{ mtd.get('pedidos', '—') }}</span></div>
  <div class="card-row"><span class="label">Pedidos Mai (prior)</span><span class="val" style="color:#888">{{ mtd.get('pedidos_prior', '—') }}</span></div>
  <div class="card-row"><span class="label">Var% Pedidos MoM</span>
    {% set varp = mtd.get('pedidos_var_pct') %}
    {% if varp is not none %}
      <span class="pill {% if varp >= 0 %}pill-green{% else %}pill-red{% endif %}">{{ '+' if varp > 0 else '' }}{{ varp }}%</span>
    {% else %}<span class="val">—</span>{% endif %}
  </div>
  <div class="card-row"><span class="label">NPS Jun</span><span class="val">{{ mtd.get('nps', '—') }}</span></div>
  <div class="card-row"><span class="label">NPS Mai (prior)</span><span class="val" style="color:#888">{{ mtd.get('nps_prior', '—') }}</span></div>
  <div class="card-row"><span class="label">SLA% Jun</span><span class="val">{{ fmt_pct(mtd.get('sla_pct')) }}</span></div>
  <div class="card-row"><span class="label">SLA% Mai (prior)</span><span class="val" style="color:#888">{{ fmt_pct(mtd.get('sla_pct_prior')) }}</span></div>
</div>
{% endfor %}
</div>

<div class="section-title" style="margin-top:1.5rem">CONSOLIDADO — Atingimento vs Meta Mês Inteiro (Jun/26)</div>
<div class="cards" id="consol-cards">
{% for g in grupos %}
{% set consol = periodos_data.get(g, {}).get('CONSOLIDADO', {}) %}
<div class="card" data-grupo="{{ g }}" {% if not loop.first %}style="display:none"{% endif %}>
  <div class="card-title">{{ g }} — CONSOLIDADO</div>
  <div class="card-row"><span class="label">GMV Real (jun 1-14)</span><span class="gmv-main">{{ fmt_brl(consol.get('gmv')) }}</span></div>
  <div class="card-row"><span class="label">Meta Mês Inteiro (jun 1-30)</span><span class="val" style="color:#f59e0b">{{ fmt_brl(consol.get('meta_gmv')) }}</span></div>
  <div class="card-row"><span class="label">Atingimento vs Meta</span>
    <span class="pill pill-{{ semaforo('ating_pct', consol.get('ating_pct')) }}" style="font-size:.92rem">{{ fmt_pct(consol.get('ating_pct')) }}</span>
  </div>
  <div class="card-row"><span class="label">Pedidos</span><span class="val">{{ consol.get('pedidos', '—') }}</span></div>
  <div class="card-row"><span class="label">AOV</span><span class="val">{{ fmt_brl(consol.get('aov')) }}</span></div>
  <div class="card-row"><span class="label">Cancel%</span><span class="val"><span class="dot dot-{{ semaforo('cancel_pct', consol.get('cancel_pct')) }}"></span>{{ fmt_pct(consol.get('cancel_pct')) }}</span></div>
  <div class="card-row"><span class="label">NPS</span><span class="val"><span class="dot dot-{{ semaforo('nps', consol.get('nps')) }}"></span>{{ consol.get('nps', '—') }}</span></div>
  <div class="card-row"><span class="label">SLA%</span><span class="val"><span class="dot dot-{{ semaforo('sla_pct', consol.get('sla_pct')) }}"></span>{{ fmt_pct(consol.get('sla_pct')) }}</span></div>
</div>
{% endfor %}
</div>

<div class="section-title">Comparativo de Períodos</div>
{% for g in grupos %}
<div id="comp-{{ g }}" class="comp-section {% if loop.first %}active{% endif %}">
<div class="table-wrap">
<table>
<thead><tr><th>Indicador</th><th>D-1</th><th>D-7</th><th>D-15</th><th>D-21</th><th>Var D1→D7</th></tr></thead>
<tbody>
{% set p = periodos_data.get(g, {}) %}
{% set d1 = p.get('D-1',{}) %}{% set d7 = p.get('D-7',{}) %}{% set d15 = p.get('D-15',{}) %}{% set d21 = p.get('D-21',{}) %}
{% for ik, il in [('gmv','GMV'),('pedidos','Pedidos'),('aov','AOV'),('cancel_pct','Cancel%'),('ruptura_pct','Ruptura%'),('er_pct','ER%'),('nps','NPS'),('sla_pct','SLA%'),('online_pct','Online%'),('gmv_rupt','GMV Rupt'),('gmv_recup','GMV Recup'),('nsu_pct','NSU%')] %}
{% set v1=d1.get(ik) %}{% set v7=d7.get(ik) %}{% set v15=d15.get(ik) %}{% set v21=d21.get(ik) %}
<tr>
<td><strong>{{ il }}</strong></td>
{% for vv in [v1, v7, v15, v21] %}
<td>{% if ik in ['gmv','gmv_rupt','gmv_recup','aov'] %}{{ fmt_brl(vv) }}{% elif ik in ['cancel_pct','ruptura_pct','er_pct','sla_pct','online_pct','nsu_pct'] %}<span class="dot dot-{{ semaforo(ik,vv) }}"></span>{{ fmt_pct(vv) }}{% elif ik == 'nps' %}<span class="dot dot-{{ semaforo('nps',vv) }}"></span>{{ vv if vv is not none else '—' }}{% else %}{{ vv if vv is not none else '—' }}{% endif %}</td>
{% endfor %}
<td>{% if v1 is not none and v7 is not none and v7 != 0 %}{% set var=((v1-v7)/v7*100)|round(1) %}<span style="color:{% if var>=0 %}#4ade80{% else %}#f87171{% endif %}">{{ '+' if var>0 else '' }}{{ var }}%</span>{% else %}—{% endif %}</td>
</tr>
{% endfor %}
</tbody>
</table>
</div>
</div>
{% endfor %}

<div class="section-title" style="margin-top:1.5rem">Lojas por Grupo / Período</div>
<div class="selects">
  <select id="selGrupoLojas" onchange="renderLojas()">{% for g in grupos %}<option value="{{ g }}">{{ g }}</option>{% endfor %}</select>
  <select id="selPeriodoLojas" onchange="renderLojas()">{% for p in periodos %}<option value="{{ p }}">{{ p }}</option>{% endfor %}</select>
</div>
<div class="table-wrap" id="lojasTable"></div>

</div>
{% endif %}

<script>
var LOJAS = {{ lojas_json | safe }};
var PERIODOS_DATA = {{ periodos_json | safe }};
var GRUPOS = {{ grupos_json | safe }};
var grupoAtivo = GRUPOS[0];

var COLS = ['loja','gmv','pedidos','ating_pct','aov','cancel_pct','ruptura_pct','er_pct','nps','sla_pct','online_pct','gmv_rupt','gmv_recup','nsu_pct'];
var LABELS = ['Loja','GMV','Pedidos','Ating%','AOV','Cancel%','Ruptura%','ER%','NPS','SLA%','Online%','GMV Rupt','GMV Recup','NSU%'];
var MONEY = ['gmv','aov','gmv_rupt','gmv_recup'];
var PCTS = ['cancel_pct','ruptura_pct','er_pct','sla_pct','online_pct','nsu_pct','ating_pct'];

function sem(k,v){
  if(v==null||v===undefined)return'gray';
  var fv=parseFloat(v);if(isNaN(fv))return'gray';
  var bad={cancel_pct:[5,8],ruptura_pct:[10,25],er_pct:[10,20]};
  var good={ating_pct:[90,80],nps:[70,50],sla_pct:[90,70],online_pct:[95,80]};
  if(k in bad){var lo=bad[k][0],hi=bad[k][1];return fv<lo?'green':fv<hi?'yellow':'red';}
  if(k in good){var hi2=good[k][0],lo2=good[k][1];return fv>=hi2?'green':fv>=lo2?'yellow':'red';}
  return'gray';
}
function brl(v){if(v==null)return'—';try{return'R$ '+parseFloat(v).toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2});}catch(e){return'—';}}
function pct(v){if(v==null)return'—';try{return parseFloat(v).toFixed(1)+'%';}catch(e){return'—';}}

function selectGroup(g){
  grupoAtivo=g;
  document.querySelectorAll('.group-btn').forEach(function(b){b.classList.toggle('active',b.textContent.trim()===g);});
  var lbl=document.getElementById('grupo-label');if(lbl)lbl.textContent=g;
  document.querySelectorAll('#resumo-cards .card').forEach(function(c){c.style.display=(c.dataset.grupo===g)?'':'none';});
  document.querySelectorAll('#mtd-cards .card').forEach(function(c){c.style.display=(c.dataset.grupo===g)?'':'none';});
  document.querySelectorAll('#consol-cards .card').forEach(function(c){c.style.display=(c.dataset.grupo===g)?'':'none';});
  GRUPOS.forEach(function(grp){var el=document.getElementById('comp-'+grp);if(el)el.classList.toggle('active',grp===g);});
  var sel=document.getElementById('selGrupoLojas');if(sel){sel.value=g;renderLojas();}
}

function renderLojas(){
  var g=document.getElementById('selGrupoLojas').value;
  var p=document.getElementById('selPeriodoLojas').value;
  var rows=(LOJAS[g]&&LOJAS[g][p])?LOJAS[g][p]:[];
  if(!rows.length){document.getElementById('lojasTable').innerHTML='<p style="color:#888;padding:.75rem">Sem dados para '+g+'/'+p+'</p>';return;}
  var h='<table><thead><tr>';
  LABELS.forEach(function(l){h+='<th>'+l+'</th>';});
  h+='</tr></thead><tbody>';
  rows.forEach(function(r){
    h+='<tr>';
    COLS.forEach(function(c){
      var v=r[c],cell='';
      if(MONEY.indexOf(c)>=0)cell=brl(v);
      else if(PCTS.indexOf(c)>=0)cell='<span class="dot dot-'+sem(c,v)+'"></span>'+pct(v);
      else if(c==='nps')cell='<span class="dot dot-'+sem('nps',v)+'"></span>'+(v!=null?parseFloat(v).toFixed(1):'—');
      else if(c==='pedidos')cell=(v!=null?parseInt(v):'—');
      else cell=(v!=null?String(v):'—');
      h+='<td>'+cell+'</td>';
    });
    h+='</tr>';
  });
  h+='</tbody></table>';
  document.getElementById('lojasTable').innerHTML=h;
}
renderLojas();
</script>
</body>
</html>"""


@app.route("/")
def index():
    data = load_data()
    periodos_data = data.get("periodos", data.get("grupos", {})) if data else {}
    lojas = data.get("lojas", {}) if data else {}
    return render_template_string(
        TEMPLATE,
        data=data,
        grupos=GRUPOS,
        periodos=PERIODOS,
        periodos_data=periodos_data,
        lojas_json=json.dumps(lojas, ensure_ascii=False),
        periodos_json=json.dumps(periodos_data, ensure_ascii=False),
        grupos_json=json.dumps(GRUPOS),
        fmt_brl=fmt_brl,
        fmt_pct=fmt_pct,
        semaforo=semaforo,
        desatualizado=dados_desatualizados(data) if data else False,
    )


@app.route("/atualizar", methods=["POST"])
def atualizar():
    try:
        payload = request.get_json(force=True)
        if not payload:
            return jsonify({"error": "JSON inválido"}), 400
        DATA_FILE.parent.mkdir(parents=True, exist_ok=True)
        DATA_FILE.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
        return jsonify({"status": "ok", "message": "Dados atualizados com sucesso"})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/dados")
def api_dados():
    data = load_data()
    if not data:
        return jsonify({"error": "Sem dados disponíveis"}), 404
    return jsonify(data)


@app.route("/health")
def health():
    return jsonify({"status": "ok", "timestamp": datetime.utcnow().isoformat()})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
