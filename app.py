"""
AVD — Configuração e Geração de Bases
Fluxo: Upload → Mapeamento → Edição → Validação → Config → Geração
"""
import io, uuid, json, zipfile, re
import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from difflib import SequenceMatcher
 
st.set_page_config(page_title="AVD — Configuração", page_icon="📋",
                   layout="wide", initial_sidebar_state="collapsed")
 
# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
 
html,body,
[data-testid="stAppViewContainer"],
[data-testid="stHeader"],
[data-testid="stSidebar"],
.main .block-container,
section[data-testid="stMain"] { background:#080b11 !important; }
 
html,body,[class*="css"] { font-family:'Inter',sans-serif; }
.block-container { padding:2.5rem 3rem 5rem !important; max-width:1140px; }
 
/* ── Steps ── */
.snav { display:flex; gap:0; border-radius:8px; overflow:hidden;
        border:1px solid #1c2033; margin-bottom:2.5rem; }
.sn   { flex:1; padding:12px 6px; text-align:center; font-size:11px; font-weight:500;
        background:#0d1018; color:#5a6070; border-right:1px solid #1c2033; cursor:default; }
.sn:last-child { border-right:none; }
.sn.cur  { background:#13163a; color:#a5b4fc; font-weight:700; }
.sn.done { background:#0d1a14; color:#4ade80; cursor:pointer; }
.sni { display:inline-flex; width:18px; height:18px; border-radius:50%;
       border:1.5px solid currentColor; font-size:9px;
       align-items:center; justify-content:center; margin-right:5px; }
.sn.done .sni { background:#4ade80; color:#052e16; border-color:#4ade80; }
 
/* ── Typography ── */
h1 { font-size:1.6rem !important; font-weight:700 !important; color:#f1f0ed !important; margin-bottom:.25rem !important; }
h2 { font-size:1.15rem !important; font-weight:600 !important; color:#e2e0db !important; }
h3 { font-size:.95rem !important; font-weight:600 !important; color:#ccc9c2 !important; }
p, li { color:#ccc9c2; font-size:13px; }
.caption { font-size:11.5px; color:#6b7280; }
.sh { font-size:10px; font-weight:700; letter-spacing:1.5px; text-transform:uppercase;
      color:#4b5260; margin:1.8rem 0 .6rem; border-bottom:1px solid #1c2033;
      padding-bottom:6px; }
 
/* ── Cards / containers ── */
.card { background:#0d1018; border:1px solid #1c2033; border-radius:10px;
        padding:18px 22px; margin-bottom:10px; }
.card-hi { border-color:#2d3155; background:#10142a; }
 
/* ── Map table ── */
.maptable { width:100%; border-collapse:collapse; }
.maptable th { font-size:10px; font-weight:700; letter-spacing:1px; text-transform:uppercase;
               color:#4b5260; padding:8px 12px; text-align:left;
               border-bottom:1px solid #1c2033; }
.maptable td { padding:6px 12px; border-bottom:1px solid #141720; vertical-align:middle; }
.maptable tr:last-child td { border-bottom:none; }
.col-source { font-family:'Courier New',monospace; font-size:11.5px; color:#a5b4fc;
              background:#0d1018; border:1px solid #1c2033; border-radius:5px;
              padding:4px 10px; display:inline-block; max-width:280px;
              overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
.col-req { color:#f87171; font-size:10px; margin-left:4px; }
.col-opt { color:#4b5260; font-size:10px; margin-left:4px; }
.col-ok  { color:#4ade80; font-size:13px; }
.col-miss{ color:#f87171; font-size:13px; }
 
/* ── Validation ── */
.vsec { border-radius:10px; padding:16px 20px; margin-bottom:12px; }
.v-ok  { background:rgba(74,222,128,.06);  border:1px solid rgba(74,222,128,.2); }
.v-err { background:rgba(248,113,113,.06); border:1px solid rgba(248,113,113,.2); }
.v-wrn { background:rgba(252,211,77,.06);  border:1px solid rgba(252,211,77,.2); }
.vtitle { font-size:13px; font-weight:700; margin-bottom:10px; }
.v-ok  .vtitle { color:#4ade80; }
.v-err .vtitle { color:#f87171; }
.v-wrn .vtitle { color:#fcd34d; }
.vitem { font-size:12.5px; padding:5px 0; color:#b8b4ad; border-bottom:1px solid #1c2033; }
.vitem:last-child { border-bottom:none; }
 
/* ── Pergunta table ── */
.ptbl { width:100%; border-collapse:collapse; font-size:12.5px; }
.ptbl th { background:#0d1018; color:#4b5260; font-size:10px; letter-spacing:.8px;
           text-transform:uppercase; padding:9px 12px; text-align:left;
           border-bottom:1px solid #1c2033; }
.ptbl td { padding:9px 12px; border-bottom:1px solid #141720; color:#ccc9c2; vertical-align:top; }
.ptbl tr:hover td { background:rgba(255,255,255,.015); }
.ptbl tr.r-err td { background:rgba(248,113,113,.05); }
.ptbl tr.r-wrn td { background:rgba(252,211,77,.04); }
 
/* ── Escala preview ── */
.erow { display:flex; gap:5px; margin:10px 0; }
.ept  { flex:1; background:#0d1018; border:1px solid #1c2033; border-radius:6px;
        padding:10px 4px; text-align:center; font-size:11px; color:#b8b4ad; line-height:1.5; }
.en   { font-size:17px; font-weight:700; color:#fff; display:block; margin-bottom:3px; }
 
/* ── Badges ── */
.bdg { font-size:10px; padding:2px 9px; border-radius:99px; font-weight:600;
       border:1px solid; margin-right:4px; display:inline-block; }
.b-ok   { color:#4ade80; border-color:rgba(74,222,128,.3); background:rgba(74,222,128,.08); }
.b-err  { color:#f87171; border-color:rgba(248,113,113,.3);background:rgba(248,113,113,.08);}
.b-wrn  { color:#fcd34d; border-color:rgba(252,211,77,.3); background:rgba(252,211,77,.08);}
.b-blue { color:#a5b4fc; border-color:rgba(165,180,252,.3);background:rgba(165,180,252,.08);}
.b-gray { color:#6b7280; border-color:#1c2033; background:#0d1018; }
 
/* ── Log ── */
.logbox { background:#060810; border:1px solid #1c2033; border-radius:8px;
          padding:14px 16px; font-family:'Courier New',monospace; font-size:12px;
          max-height:380px; overflow-y:auto; line-height:2; color:#9ca3af; }
.lok { color:#4ade80; } .lerr { color:#f87171; }
.linf { color:#a5b4fc; } .lwrn { color:#fcd34d; }
 
/* ── Streamlit overrides ── */
div[data-testid="stMetric"] {
  background:#0d1018 !important; border:1px solid #1c2033 !important;
  border-radius:10px; padding:14px 18px; }
div[data-testid="stMetricValue"] { font-size:1.8rem !important; font-weight:700; color:#f1f0ed !important; }
div[data-testid="stMetricLabel"] { color:#6b7280 !important; font-size:12px !important; }
 
div[data-testid="stTabs"] button {
  background:#0d1018 !important; color:#6b7280 !important; border-color:#1c2033 !important; }
div[data-testid="stTabs"] button[aria-selected="true"] {
  color:#a5b4fc !important; border-bottom-color:#a5b4fc !important; background:#0d1018 !important; }
 
.stTextInput input, .stTextArea textarea {
  background:#0d1018 !important; color:#e2e0db !important; border-color:#2a2f45 !important; }
.stTextInput input:focus, .stTextArea textarea:focus {
  border-color:#a5b4fc !important; box-shadow:0 0 0 2px rgba(165,180,252,.12) !important; }
.stTextInput label, .stTextArea label, .stSelectbox label,
.stNumberInput label, .stMultiSelect label, .stCheckbox label { color:#b8b4ad !important; font-size:12.5px !important; }
.stSelectbox [data-baseweb="select"]>div,
.stMultiSelect [data-baseweb="select"]>div {
  background:#0d1018 !important; border-color:#2a2f45 !important; color:#e2e0db !important; }
[data-baseweb="menu"], [data-baseweb="popover"] {
  background:#0d1018 !important; border-color:#2a2f45 !important; }
[data-baseweb="option"] { background:#0d1018 !important; color:#ccc9c2 !important; }
[data-baseweb="option"]:hover { background:#13163a !important; }
 
div[data-testid="stExpander"] {
  background:#0d1018 !important; border-color:#1c2033 !important; border-radius:8px !important; }
div[data-testid="stExpander"] summary { color:#b8b4ad !important; }
div[data-testid="stExpander"] summary:hover { color:#e2e0db !important; }
 
div[data-testid="stDataFrame"] { background:#0d1018 !important; }
[data-testid="stFileUploader"] {
  background:#0d1018 !important; border-color:#2a2f45 !important; border-radius:8px !important; }
[data-testid="stFileUploader"] p,[data-testid="stFileUploader"] span { color:#6b7280 !important; }
 
button[kind="primary"] {
  background:#4f46e5 !important; color:#fff !important; border:none !important;
  font-weight:600 !important; border-radius:7px !important;
  box-shadow:0 2px 10px rgba(79,70,229,.35) !important; }
button[kind="primary"]:hover { background:#4338ca !important; }
button[kind="secondary"] {
  background:#0d1018 !important; color:#b8b4ad !important;
  border:1px solid #2a2f45 !important; border-radius:7px !important; }
button[kind="secondary"]:hover { background:#13163a !important; border-color:#3d4460 !important; color:#e2e0db !important; }
 
div[data-testid="stAlert"] { border-radius:8px !important; }
hr { border-color:#1c2033 !important; margin:1.5rem 0 !important; }
[data-testid="stMarkdownContainer"] p { color:#b8b4ad; }
[data-testid="stMarkdownContainer"] strong { color:#e2e0db; }
</style>
""", unsafe_allow_html=True)
 
# ── Constants ─────────────────────────────────────────────────────────────────
# Campos alvo e seus sinônimos para auto-detecção
TARGET_FIELDS = [
    ("question_set",      "Bloco (question_set)",       True,
     ["nome do bloco","bloco","question_set","categoria"]),
    ("pergunta",          "Pergunta",                   True,
     ["pergunta","question","enunciado","texto"]),
    ("aberta",            "Aberta? (sim/não)",          True,
     ["é uma pergunta aberta","aberta","open","tipo"]),
    ("opcional",          "Opcional? (sim/não)",        True,
     ["essa pergunta é opcional","opcional","optional"]),
    ("escala",            "Nº da Escala",               False,
     ["qual o número do modelo da escala","escala","scale","número da escala"]),
    ("competencia",       "Competência",                False,
     ["nome da competência","competência","competencia","skill","tema"]),
    ("desc_question_set", "Descrição do bloco",         False,
     ["descrição do bloco","desc_question_set"]),
    ("min_caracters",     "Mín. caracteres",            False,
     ["número mínimo de caracteres","min caracters"]),
    ("max_caracters",     "Máx. caracteres",            False,
     ["número máximo de caracteres","max caracters"]),
    ("definicao",         "Definição da pergunta",      False,
     ["descrição geral sobre a pergunta","definição","definicao"]),
    ("g_auto",            "Autoavaliação (sim/não)",    False,
     ["autoavaliação","autoavaliacao"]),
    ("g_lider",           "Líder→Liderado (sim/não)",  False,
     ["líder avaliando liderados","lider avaliando liderados"]),
    ("g_liderado",        "Liderado→Líder (sim/não)",  False,
     ["liderados avaliando líder","liderados avaliando lider"]),
    ("g_par",             "Pares (sim/não)",            False,
     ["pares","par"]),
    ("g_stakeholder",     "Stakeholders (sim/não)",     False,
     ["stakeholders","stakeholder"]),
]
FIELD_KEY   = {f[0]: f for f in TARGET_FIELDS}
GRUPOS_OPTS = ["Autoavaliação","Time","Gestor","Par","Stakeholder"]
DICT_CORES  = {
    2:['#A8002E','#006499'], 3:['#A8002E','#FFC252','#006499'],
    4:['#A8002E','#FFC252','#35AEFF','#033D17'],
    5:['#A8002E','#FFC252','#35AEFF','#006499','#033D17'],
    6:['#A8002E','#FA6F26','#FFC252','#35AEFF','#006499','#033D17'],
}
OUTPUT_COLS = ["question_set","desc_question_set","competencia","pergunta",
               "opcional","aberta","escala","min caracters","max caracters","definição"]
 
# ── Helpers ───────────────────────────────────────────────────────────────────
def uid(): return str(uuid.uuid4())[:8]
def norm_bool(v):
    v = str(v or "").strip().lower()
    return "sim" if v in ("sim","yes","true","1","s","y") else "não"
def sim_r(a,b): return SequenceMatcher(None,a.lower().strip(),b.lower().strip()).ratio()
 
def auto_detect(headers):
    """Detecta mapeamento automático: field → header original."""
    hl = [str(h or "").lower().strip().rstrip("?").strip() for h in headers]
    mapping = {}
    for key, label, req, syns in TARGET_FIELDS:
        for i,h in enumerate(hl):
            if any(s in h or h in s for s in syns):
                mapping[key] = headers[i]
                break
    return mapping
 
def make_p(**kw):
    return {"id":uid(),"texto":kw.get("texto",""),"competencia":kw.get("competencia",""),
            "aberta":kw.get("aberta","não"),"escala":str(kw.get("escala","")),"opcional":kw.get("opcional","não"),
            "min_caracters":str(kw.get("min_c","")),"max_caracters":str(kw.get("max_c","")),"definicao":kw.get("definicao",""),
            "grupos":kw.get("grupos",list(GRUPOS_OPTS))}
def make_b(nome="",desc="",perguntas=None,origem="manual"):
    return {"id":uid(),"nome":nome,"desc":desc,"perguntas":perguntas or [],"origem":origem}
def make_e(num=None,pontos=None):
    if num is None: num = max((e["num"] for e in ss.escalas),default=0)+1
    try: num = int(float(str(num)))
    except: pass
    return {"id":uid(),"num":num,"pontos":pontos or ["",""]}
 
# ── Session state ─────────────────────────────────────────────────────────────
DEFS = {
    "step":1,
    # upload / mapping
    "raw_sheets":{},      # sheet_name → list of raw rows (including header)
    "perg_sheet":"",      # aba de perguntas selecionada
    "esc_sheet":"",       # aba de escalas selecionada
    "col_map":{},         # field → header_string
    "imported_file":"",
    # data
    "escalas":[], "blocos":[], "bloco_sel":None,
    # config / results
    "cfg":{"cliente":"","token":"","id_rodada":"","id_qs":"","id_cq":"","id_oq":"","val_min_nan":"0","sso":True},
    "resultados":None,"log":[],
}
for k,v in DEFS.items():
    if k not in st.session_state: st.session_state[k]=v
ss = st.session_state
 
# ── Nav ───────────────────────────────────────────────────────────────────────
STEPS = ["1. Upload","2. Mapeamento","3. Edição","4. Validação","5. Config","6. Geração"]
def steps_nav():
    pills = ""
    for i,l in enumerate(STEPS,1):
        cls = "cur" if i==ss.step else ("done" if i<ss.step else "")
        n   = "✓" if i<ss.step else str(i)
        pills += f'<div class="sn {cls}"><span class="sni">{n}</span>{l}</div>'
    st.markdown(f'<div class="snav">{pills}</div>', unsafe_allow_html=True)
def go(n): ss.step=n; st.rerun()
 
def nav_row(back=None, fwd=None, fwd_label="Continuar →", fwd_dis=False):
    c1,_,c2 = st.columns([1,5,1])
    if back and c1.button("← Voltar", use_container_width=True): go(back)
    if fwd  and c2.button(fwd_label, type="primary", use_container_width=True, disabled=fwd_dis): go(fwd)
 
# ── Scale reader ──────────────────────────────────────────────────────────────
def read_escalas_from_rows(rows):
    """
    Lê escalas em dois formatos:
    - Horizontal (Implementação_Performance): blocos de 3 colunas
    - Vertical (Base_Perguntas-2): linhas com colunas descrição_1, descrição_2...
    """
    if not rows: return []
    header = [str(c or "").strip() for c in rows[0]]
    header_l = [h.lower() for h in header]
 
    # Detecta formato horizontal: cabeçalho tem "modelo" ou "escala" em posição >0
    is_horizontal = any(
        ("modelo" in h or ("escala" in h and h != "escala"))
        for h in header_l if h
    )
 
    escalas = []
    if is_horizontal:
        # Monta pares (col_val, col_desc) por bloco de 3
        pair_cols = []
        i = 0
        while i < len(header):
            if header[i]:
                pair_cols.append((i, i+1))
                i += 3
            else:
                i += 1
        for ci_v, ci_d in pair_cols:
            nome = header[ci_v]
            vals, pontos = [], []
            for row in rows[2:]:   # linha 1 é sub-cabeçalho
                v = row[ci_v] if ci_v < len(row) else None
                d = row[ci_d] if ci_d < len(row) else None
                if v is not None:
                    try: vals.append(float(v))
                    except: pass
                if d and str(d).strip(): pontos.append(str(d).strip())
            if pontos:
                num = int(min(vals)) if vals else len(escalas)+1
                escalas.append(make_e(num=num, pontos=pontos))
    else:
        # Formato vertical
        desc_cols = [i for i,h in enumerate(header_l) if "descrição_" in h]
        for row in rows[1:]:
            if not row[0]: continue
            try: num = int(float(str(row[0])))
            except: num = str(row[0])
            pontos = [str(row[i] or "").strip() for i in desc_cols if i<len(row) and row[i]]
            if pontos: escalas.append(make_e(num=num, pontos=pontos))
 
    return escalas
 
# ── Apply mapping & build blocks ──────────────────────────────────────────────
def apply_mapping(rows, col_map):
    """Converte linhas brutas em blocos usando o mapeamento confirmado."""
    if not rows: return [], []
    headers = [str(c or "").strip() for c in rows[0]]
    h_idx   = {h: i for i,h in enumerate(headers)}
 
    def get(row, field):
        col = col_map.get(field,"")
        if not col or col not in h_idx: return ""
        i = h_idx[col]
        return str(row[i] or "").strip() if i < len(row) else ""
 
    def get_bool(row, field):
        v = get(row, field)
        return norm_bool(v) == "sim"
 
    bmap = {}
    for row in rows[1:]:
        if not any(c not in (None,"") for c in row): continue
        bn = get(row,"question_set")
        pt = get(row,"pergunta")
        if not bn and not pt: continue
 
        grupos = [gl for gf,gl in [("g_auto","Autoavaliação"),("g_lider","Time"),
                                    ("g_liderado","Gestor"),("g_par","Par"),("g_stakeholder","Stakeholder")]
                  if get_bool(row,gf)]
        if not grupos: grupos = list(GRUPOS_OPTS)
 
        p = make_p(texto=pt, competencia=get(row,"competencia"),
                   aberta=norm_bool(get(row,"aberta") or "não"),
                   escala=get(row,"escala"),
                   opcional=norm_bool(get(row,"opcional") or "não"),
                   min_c=get(row,"min_caracters"), max_c=get(row,"max_caracters"),
                   definicao=get(row,"definicao"), grupos=grupos)
        bmap.setdefault(bn, []).append(p)
 
    msgs = []
    nomes = list(bmap.keys())
    for i,n1 in enumerate(nomes):
        for n2 in nomes[i+1:]:
            if sim_r(n1,n2)>0.85: msgs.append(f"⚠ Blocos parecidos: **{n1}** e **{n2}**")
 
    blocos = [make_b(nome=bn, perguntas=ps, origem="import") for bn,ps in bmap.items()]
    return blocos, msgs
 
# ── Validation ────────────────────────────────────────────────────────────────
def validar():
    erros, avisos = [], []
    enum = {str(e["num"]) for e in ss.escalas}
    if not ss.escalas: erros.append("Nenhuma escala cadastrada.")
    if not ss.blocos:  erros.append("Nenhum bloco cadastrado.")
    seen = set()
    for b in ss.blocos:
        n = b["nome"].strip()
        if not n: erros.append("Bloco sem nome.")
        if n in seen: erros.append(f'Bloco duplicado: "{n}"')
        seen.add(n)
        if not b["perguntas"]: avisos.append(f'"{n}" sem perguntas.')
        for p in b["perguntas"]:
            if not p["texto"].strip(): erros.append(f'Pergunta vazia em "{n}".')
            if p["aberta"]=="não":
                if not p["escala"]: avisos.append(f'Fechada sem escala em "{n}".')
                elif p["escala"] not in enum: erros.append(f'Escala {p["escala"]} inexistente em "{n}".')
            if p["aberta"]=="sim" and not p["max_caracters"]: avisos.append(f'Aberta sem máx chars em "{n}".')
    total = sum(len(b["perguntas"]) for b in ss.blocos)
    return erros, avisos, total
 
def val_p(p, enum):
    e,a=[],[]
    if not p["texto"].strip(): e.append("Texto vazio")
    if p["aberta"]=="não":
        if not p["escala"]: a.append("Sem escala")
        elif p["escala"] not in enum: e.append(f"Escala {p['escala']} inexistente")
    if p["aberta"]=="sim" and not p["max_caracters"]: a.append("Sem máx chars")
    return e,a
 
# ── Script ────────────────────────────────────────────────────────────────────
def run_script(cfg):
    log = []
    def L(msg,t="inf"): log.append((t,msg))
    try:
        id_rodada=int(cfg["id_rodada"]); id_qs=int(cfg["id_qs"])
        id_cq=int(cfg["id_cq"]); id_oq=int(cfg["id_oq"])
        val_min=float(cfg.get("val_min_nan") or "0")
    except Exception as e: L(f"IDs inválidos: {e}","err"); return log,{}
 
    rows_p=[]
    for b in ss.blocos:
        for p in b["perguntas"]:
            rows_p.append({"question_set":b["nome"],"desc_question_set":b["desc"],
                "competencia":p["competencia"],"pergunta":p["texto"],"opcional":p["opcional"],"aberta":p["aberta"],
                "escala":int(float(p["escala"])) if p["aberta"]=="não" and p["escala"] else np.nan,
                "min caracters":int(p["min_caracters"]) if p["min_caracters"] else np.nan,
                "max caracters":int(p["max_caracters"]) if p["max_caracters"] else np.nan,
                "definição":p["definicao"] or np.nan})
    bp = pd.DataFrame(rows_p)
 
    rows_e=[]
    for e in ss.escalas:
        r={"escala":e["num"],
           "min":float(e["pontos"][0]) if e["pontos"] and e["pontos"][0] else np.nan,
           "max":float(e["pontos"][-1]) if e["pontos"] and e["pontos"][-1] else np.nan}
        for i,pt in enumerate(e["pontos"],1): r[f"descrição_{i}"]=pt or np.nan
        rows_e.append(r)
    be = pd.DataFrame(rows_e)
    L(f"Perguntas: {len(bp)} | Escalas: {len(be)}")
 
    # Question Sets
    L("Gerando Question Sets...")
    bqs=pd.DataFrame(columns=['name','description','id','order','evaluation_rounds'])
    first=True;cnt=0;criados=[]
    for _,row in bp.iterrows():
        if row['question_set'] not in criados:
            desc="" if pd.isna(row.get('desc_question_set','')) else str(row.get('desc_question_set',''))
            reg={'id':id_qs if first else id_qs+cnt,'order':cnt+1,'name':row['question_set'],
                 'description':desc,'evaluation_rounds':id_rodada}
            if first: first=False
            bqs.loc[len(bqs)]=reg; criados.append(row['question_set']); cnt+=1
    L(f"✓ {len(bqs)} question set(s)","ok")
 
    # Open Questions
    L("Gerando Open Questions...")
    boq=pd.DataFrame(columns=['info','description','id','question_set','is_optional',
                               'order','key','min_answer_length','max_answer_length','evaluation_rounds'])
    first=True;cnt_oq=0;cnt_p=1
    for _,row in bp.iterrows():
        if str(row['aberta']).lower()=='não': cnt_p+=1; continue
        desc=("" if pd.isna(row.get('competencia')) else str(row['competencia']).upper()+" | ")+row['pergunta']
        opc=0 if str(row['opcional']).lower()=='não' else 1
        min_c=0 if pd.isna(row.get('min caracters')) else int(row['min caracters'])
        max_c=int(row['max caracters'])
        qs=bqs[bqs['name']==row['question_set']]
        reg={'id':id_oq if first else id_oq+cnt_oq,'order':cnt_p,'key':id_oq if first else id_oq+cnt_oq,
             'description':desc,'info':'','is_optional':opc,'min_answer_length':min_c,
             'max_answer_length':max_c,'question_set':qs['id'].iloc[0],'evaluation_rounds':id_rodada}
        if first: first=False
        boq.loc[len(boq)]=reg; cnt_oq+=1; cnt_p+=1
    L(f"✓ {len(boq)} open question(s)","ok")
 
    # Choice Questions
    L("Gerando Choice Questions...")
    bcq=pd.DataFrame(columns=['info','description','id','question_set','is_optional','order','key','evaluation_rounds','escala'])
    first=True;cnt_cq=0;cnt_p=1
    for _,row in bp.iterrows():
        if str(row['aberta']).lower()=='sim': cnt_p+=1; continue
        desc=("" if pd.isna(row.get('competencia')) else str(row['competencia']).upper()+" | ")+row['pergunta']
        opc=0 if str(row['opcional']).lower()=='não' else 1
        qs=bqs[bqs['name']==row['question_set']]
        reg={'id':id_cq if first else id_cq+cnt_cq,'order':cnt_p,'key':id_cq if first else id_cq+cnt_cq,
             'description':desc,'info':'','is_optional':opc,'question_set':qs['id'].iloc[0],
             'escala':int(row['escala']),'evaluation_rounds':id_rodada}
        if first: first=False
        bcq.loc[len(bcq)]=reg; cnt_cq+=1; cnt_p+=1
    L(f"✓ {len(bcq)} choice question(s)","ok")
 
    # Escalas
    L("Gerando itens de escala...")
    besc=pd.DataFrame(columns=['content','id','question','value','color'])
    for _,row in bcq.iterrows():
        bt=be[be['escala']==row['escala']]
        if len(bt)==0: L(f"Escala {row['escala']} não encontrada","err"); continue
        for _,re_ in bt.iterrows():
            cols_d=[c for c in bt.columns if c not in['escala','min','max']]
            pts=[c for c in cols_d if not pd.isna(re_.get(c))]
            cnt_pts=len(pts)
            minimo=re_.get('min'); maximo=float(re_['max'])
            flag_nan=pd.isna(minimo)
            if flag_nan: minimo=val_min; div=(maximo-float(minimo))/max(cnt_pts-2,1)
            else: minimo=float(minimo); div=(maximo-minimo)/max(cnt_pts-1,1)
            cores=DICT_CORES.get(cnt_pts,['#888888']*cnt_pts)
            ci=0
            for col in cols_d:
                if pd.isna(re_.get(col)): continue
                if flag_nan:
                    reg={'id':'','question':row['id'],'value':np.nan,'content':str(re_[col]),'color':'#000000'}
                    flag_nan=False
                else:
                    reg={'id':'','question':row['id'],'value':round(float(minimo)+ci*div,2),
                         'content':str(re_[col]),'color':cores[ci] if ci<len(cores) else '#888'}
                    ci+=1
                besc.loc[len(besc)]=reg
    L(f"✓ {len(besc)} itens de escala","ok")
 
    cliente=cfg["cliente"]
    res={
        f"Question Sets {cliente}.json":    bqs.to_json(orient='records',force_ascii=False),
        f"Open Questions {cliente}.json":   boq.to_json(orient='records',force_ascii=False),
        f"Choice Questions {cliente}.json": bcq.drop(columns=['escala'],errors='ignore').to_json(orient='records',force_ascii=False),
        f"Escalas {cliente}.json":          besc.to_json(orient='records',force_ascii=False),
    }
    L("✓ Todos os JSONs prontos","ok")
    return log, res
 
def build_zip(res):
    buf=io.BytesIO()
    with zipfile.ZipFile(buf,'w',zipfile.ZIP_DEFLATED) as z:
        for fn,c in res.items(): z.writestr(fn,c.encode('utf-8'))
    return buf.getvalue()
 
 
# ═══════════════════════════════════════════════════════════════════════════════
# STEP 1 — Upload
# ═══════════════════════════════════════════════════════════════════════════════
if ss.step == 1:
    steps_nav()
    st.markdown("# Configuração de Avaliação")
    st.markdown("Importe a planilha do cliente para começar, ou pule direto para a edição manual.")
 
    tab_imp, tab_man = st.tabs(["📥  Importar planilha", "✏️  Criar do zero"])
 
    with tab_imp:
        st.markdown("")
        up = st.file_uploader("Arraste o arquivo .xlsx do cliente aqui",
                               type=["xlsx","xls"], label_visibility="collapsed")
        if up:
            wb = load_workbook(io.BytesIO(up.read()), read_only=True, data_only=True)
            ss.raw_sheets = {}
            for sn in wb.sheetnames:
                ws = wb[sn]
                rows = [list(r) for r in ws.iter_rows(values_only=True)
                        if any(c is not None for c in r)]
                if rows: ss.raw_sheets[sn] = rows
            wb.close()
            ss.imported_file = up.name
 
            sheet_names = list(ss.raw_sheets.keys())
 
            # Separa automaticamente: "escala" no nome → escalas; resto → perguntas
            esc_sheets  = [s for s in sheet_names if "escal" in s.lower()]
            perg_sheets = [s for s in sheet_names if "escal" not in s.lower()]
            ss.esc_sheet  = esc_sheets[0]  if esc_sheets  else ""
            ss.perg_sheet = perg_sheets[0] if perg_sheets else sheet_names[0]
 
            # Mostra o que foi detectado
            st.success(f"✓ **{up.name}** lido com sucesso")
 
            c1, c2 = st.columns(2)
            with c1:
                st.markdown('<div class="sh">Abas de perguntas</div>', unsafe_allow_html=True)
                for s in perg_sheets:
                    n_rows = max(0, len(ss.raw_sheets[s]) - 1)
                    st.markdown(f"📋 **{s}** — {n_rows} linha(s)")
                if not perg_sheets:
                    st.warning("Nenhuma aba de perguntas detectada.")
 
            with c2:
                st.markdown('<div class="sh">Aba de escalas</div>', unsafe_allow_html=True)
                if esc_sheets:
                    for s in esc_sheets:
                        st.markdown(f"⚖️ **{s}**")
                else:
                    st.markdown("<span style='color:#6b7280;font-size:12px'>Nenhuma aba de escalas — adicione manualmente na edição.</span>",
                                unsafe_allow_html=True)
 
            # Auto-detect colunas da primeira aba de perguntas
            if perg_sheets and perg_sheets[0] in ss.raw_sheets:
                headers = [str(c or "").strip() for c in ss.raw_sheets[perg_sheets[0]][0]]
                ss.col_map = auto_detect(headers)
                detected = len([k for k in ss.col_map if ss.col_map[k]])
                st.markdown("")
                st.markdown(f"<span class='caption'>🔍 {detected} de {len(TARGET_FIELDS)} campos mapeados automaticamente a partir de **{perg_sheets[0]}** — revise no próximo passo.</span>",
                            unsafe_allow_html=True)
 
            if st.button("Continuar para mapeamento →", type="primary", key="go_map"):
                if esc_sheets:
                    ss.escalas = read_escalas_from_rows(ss.raw_sheets[esc_sheets[0]])
                go(2)
 
    with tab_man:
        st.markdown("")
        st.info("Você vai cadastrar escalas, blocos e perguntas diretamente na etapa de Edição.")
        if st.button("Ir para Edição →", type="primary", key="go_edit_direct"):
            go(3)
 
 
# ═══════════════════════════════════════════════════════════════════════════════
# STEP 2 — Mapeamento de colunas
# ═══════════════════════════════════════════════════════════════════════════════
elif ss.step == 2:
    steps_nav()
    st.markdown("# Mapeamento de Colunas")
    perg_sheets_2 = [s for s in ss.raw_sheets if "escal" not in s.lower()]
    perg_ref = perg_sheets_2[0] if perg_sheets_2 else ""
    abas_str = ", ".join(f"**{s}**" for s in perg_sheets_2)
    st.markdown(f"Arquivo: **{ss.imported_file}** · Abas de perguntas: {abas_str}")
    st.markdown("Confirme quais colunas correspondem a cada campo. "
                "O mapeamento vale para todas as abas de perguntas (compartilham o mesmo cabeçalho). "
                "Campos marcados com <span style='color:#f87171'>🔴</span> são obrigatórios.",
                unsafe_allow_html=True)
 
    if not perg_ref or perg_ref not in ss.raw_sheets:
        st.error("Nenhuma aba de perguntas encontrada. Volte ao Upload.")
        nav_row(back=1)
        st.stop()
 
    headers = [str(c or "").strip() for c in ss.raw_sheets[perg_ref][0]]
    headers_clean = [h for h in headers if h]
    col_opts = ["— ignorar —"] + headers_clean
 
    st.markdown('<div class="sh">Mapeamento</div>', unsafe_allow_html=True)
 
    # Tabela de mapeamento — dois campos por linha
    field_pairs = []
    fields = TARGET_FIELDS
    for i in range(0, len(fields), 2):
        field_pairs.append(fields[i:i+2])
 
    updated_map = dict(ss.col_map)
 
    for pair in field_pairs:
        cols = st.columns(2)
        for ci, (key, label, req, _) in enumerate(pair):
            with cols[ci]:
                current = updated_map.get(key, "")
                idx = col_opts.index(current) if current in col_opts else 0
                chosen = st.selectbox(
                    f"{'🔴 ' if req else '⚪ '}{label}",
                    col_opts, index=idx, key=f"cm_{key}",
                    help="Obrigatório" if req else "Opcional"
                )
                updated_map[key] = "" if chosen == "— ignorar —" else chosen
 
    ss.col_map = updated_map
 
    # Status do mapeamento
    st.markdown("")
    missing_req = [FIELD_KEY[k][1] for k,_,req,_ in TARGET_FIELDS if req and not updated_map.get(k)]
    if missing_req:
        st.error(f"Campos obrigatórios não mapeados: **{', '.join(missing_req)}**")
 
    # Preview das primeiras linhas com o mapeamento atual
    with st.expander("👁 Preview — primeiras 5 linhas com o mapeamento atual"):
        rows = ss.raw_sheets[perg_ref]
        h_idx = {h: i for i,h in enumerate(headers)}
        preview_rows = []
        for row in rows[1:6]:
            r = {}
            for key, label, req, _ in TARGET_FIELDS:
                col = updated_map.get(key,"")
                if col and col in h_idx:
                    i = h_idx[col]
                    r[label] = str(row[i] or "").strip()[:60] if i < len(row) else ""
            preview_rows.append(r)
        if preview_rows:
            st.dataframe(pd.DataFrame(preview_rows), use_container_width=True, hide_index=True)
 
    nav_row(back=1, fwd=3, fwd_label="Confirmar e editar →", fwd_dis=bool(missing_req))
 
    # Ao avançar, processa os dados
    if ss.step == 3:  # triggered by nav_row
        blocos, msgs = apply_mapping(ss.raw_sheets[ss.perg_sheet], ss.col_map)
        ss.blocos = blocos
        ss.bloco_sel = None
 
 
# ═══════════════════════════════════════════════════════════════════════════════
# STEP 3 — Edição
# ═══════════════════════════════════════════════════════════════════════════════
elif ss.step == 3:
    # Aplica mapeamento se veio do step 2 — mescla todas as abas de perguntas
    perg_sheets_3 = [s for s in ss.raw_sheets if "escal" not in s.lower()]
    if ss.raw_sheets and perg_sheets_3 and ss.col_map and not ss.blocos:
        all_blocos, all_msgs = [], []
        for sheet in perg_sheets_3:
            rows = ss.raw_sheets[sheet]
            if len(rows) < 2: continue
            bs, ms = apply_mapping(rows, ss.col_map)
            all_blocos.extend(bs)
            all_msgs.extend(ms)
        # Mescla blocos com mesmo nome e mesmas perguntas entre abas
        bmap = {}
        for b in all_blocos:
            if b["nome"] not in bmap:
                bmap[b["nome"]] = b
            else:
                existing_txts = {p["texto"] for p in bmap[b["nome"]]["perguntas"]}
                for p in b["perguntas"]:
                    if p["texto"] not in existing_txts:
                        bmap[b["nome"]]["perguntas"].append(p)
        ss.blocos = list(bmap.values())
 
    steps_nav()
    st.markdown("# Edição")
 
    escala_opts = [""] + [str(e["num"]) for e in ss.escalas]
 
    tab_esc, tab_blocos = st.tabs(["⚖️  Escalas", "📋  Blocos & Perguntas"])
 
    # ── Escalas ──
    with tab_esc:
        st.markdown("")
        st.markdown("Defina as escalas e seus pontos de ancoragem.")
 
        for ei, esc in enumerate(ss.escalas):
            with st.container(border=True):
                h1,h2 = st.columns([8,1])
                with h1:
                    nv = st.number_input("Nº da escala", min_value=1,
                                          value=int(esc["num"]) if str(esc["num"]).replace(".","").isdigit() else ei+1,
                                          key=f"en_{esc['id']}")
                    esc["num"] = nv
                    if any(str(p).strip() for p in esc["pontos"]):
                        prev = "".join(f'<div class="ept"><span class="en">{i+1}</span>{p or "—"}</div>'
                                       for i,p in enumerate(esc["pontos"]))
                        st.markdown(f'<div class="erow">{prev}</div>', unsafe_allow_html=True)
                    with st.expander("✏ Editar pontos"):
                        n_pts = st.number_input("Nº de pontos", 2, 10, len(esc["pontos"]), key=f"np_{esc['id']}")
                        while len(esc["pontos"]) < n_pts: esc["pontos"].append("")
                        esc["pontos"] = esc["pontos"][:n_pts]
                        pc = st.columns(min(n_pts,5))
                        for pi in range(n_pts):
                            esc["pontos"][pi] = pc[pi%len(pc)].text_input(
                                f"Ponto {pi+1}", value=esc["pontos"][pi], key=f"pt_{esc['id']}_{pi}")
                with h2:
                    st.markdown("<br><br>", unsafe_allow_html=True)
                    if st.button("✕", key=f"de_{esc['id']}"): ss.escalas.pop(ei); st.rerun()
 
        if st.button("＋ Adicionar escala", key="ae"): ss.escalas.append(make_e()); st.rerun()
        if not ss.escalas:
            st.info("Nenhuma escala ainda. Adicione manualmente ou volte ao Upload com um arquivo que contenha a aba Escalas.")
 
    # ── Blocos & Perguntas ──
    with tab_blocos:
        st.markdown("")
        col_bl, col_pr = st.columns([5,7], gap="large")
 
        with col_bl:
            st.markdown(f"**{len(ss.blocos)} bloco(s)** · {sum(len(b['perguntas']) for b in ss.blocos)} pergunta(s)")
            for bi, bloco in enumerate(ss.blocos):
                np_ = len(bloco["perguntas"])
                ne_ = sum(1 for p in bloco["perguntas"] if not p["texto"] or (p["aberta"]=="não" and not p["escala"]))
                is_s = ss.bloco_sel == bloco["id"]
                with st.container(border=is_s):
                    st.markdown(
                        f"**{bloco['nome'] or '(sem nome)'}**  \n"
                        f"<span style='font-size:11px;color:#6b7280'>{np_}p "
                        f"{'· <span style=\"color:#f87171\">'+str(ne_)+' erro(s)</span>' if ne_ else '· <span style=\"color:#4ade80\">OK</span>'}</span>",
                        unsafe_allow_html=True)
                    b1,b2 = st.columns([3,1])
                    if b1.button("Editar", key=f"sb_{bloco['id']}", use_container_width=True):
                        ss.bloco_sel = bloco["id"] if not is_s else None; st.rerun()
                    if b2.button("✕", key=f"db_{bloco['id']}"):
                        ss.blocos = [b for b in ss.blocos if b["id"]!=bloco["id"]]
                        if ss.bloco_sel==bloco["id"]: ss.bloco_sel=None
                        st.rerun()
            st.markdown("")
            if st.button("＋ Novo bloco", use_container_width=True, key="nb"):
                nb = make_b(nome="Novo bloco"); ss.blocos.append(nb); ss.bloco_sel=nb["id"]; st.rerun()
 
        with col_pr:
            bloco = next((b for b in ss.blocos if b["id"]==ss.bloco_sel), None)
            if bloco is None:
                st.info("👈 Selecione um bloco para editar suas perguntas.")
            else:
                bloco["nome"] = st.text_input("Nome do bloco *", value=bloco["nome"], key=f"bn_{bloco['id']}")
                bloco["desc"] = st.text_input("Descrição", value=bloco["desc"], key=f"bd_{bloco['id']}")
 
                if bloco["perguntas"]:
                    st.markdown(f"**{len(bloco['perguntas'])} pergunta(s)**")
                    df_t = pd.DataFrame([{
                        "#": i+1,
                        "Pergunta": p["texto"][:65]+("…" if len(p["texto"])>65 else ""),
                        "Tipo": "Aberta" if p["aberta"]=="sim" else "Fechada",
                        "Escala": p["escala"] if p["aberta"]=="não" else "—",
                        "Opcional": "Sim" if p["opcional"]=="sim" else "Não",
                    } for i,p in enumerate(bloco["perguntas"])])
                    st.dataframe(df_t, use_container_width=True, hide_index=True)
                    rc = st.columns(min(len(bloco["perguntas"]),6))
                    for i,p in enumerate(bloco["perguntas"]):
                        if rc[i%len(rc)].button(f"✕{i+1}", key=f"dp_{p['id']}", use_container_width=True):
                            bloco["perguntas"].pop(i); st.rerun()
                    st.divider()
 
                st.markdown("**Adicionar pergunta:**")
                with st.container(border=True):
                    ntx = st.text_area("Texto *", key=f"ntx_{bloco['id']}", placeholder="Texto da pergunta...", height=75, label_visibility="collapsed")
                    fa,fb = st.columns(2)
                    nco = fa.text_input("Competência", key=f"nc_{bloco['id']}", placeholder="Ex: Liderança")
                    nty = fb.selectbox("Tipo", ["Fechada","Aberta"], key=f"nty_{bloco['id']}")
                    nab = "não" if nty=="Fechada" else "sim"
                    fc,fd = st.columns(2)
                    if nab=="não":
                        nesc = fc.selectbox("Escala *", escala_opts, key=f"ne_{bloco['id']}")
                        nmin = ""; nmax = ""
                    else:
                        nesc = ""
                        nmin = fc.text_input("Mín chars", key=f"nmi_{bloco['id']}", placeholder="0")
                        nmax = fd.text_input("Máx chars", key=f"nma_{bloco['id']}", placeholder="10000")
                    ngr = st.multiselect("Grupos avaliativos *", GRUPOS_OPTS, default=GRUPOS_OPTS, key=f"ngr_{bloco['id']}")
                    fe,ff = st.columns(2)
                    nop = fe.checkbox("Opcional?", key=f"nop_{bloco['id']}")
                    ndf = ff.text_input("Definição", key=f"ndf_{bloco['id']}")
                    if st.button("＋ Adicionar ao bloco", type="primary", use_container_width=True, key=f"add_{bloco['id']}"):
                        if not ntx.strip(): st.error("Texto obrigatório.")
                        elif nab=="não" and not nesc: st.error("Selecione uma escala.")
                        else:
                            bloco["perguntas"].append(make_p(texto=ntx,competencia=nco,aberta=nab,
                                escala=nesc,opcional="sim" if nop else "não",min_c=nmin,max_c=nmax,
                                definicao=ndf,grupos=ngr))
                            st.rerun()
 
    nav_row(back=2 if ss.raw_sheets else None, fwd=4, fwd_label="Ir para Validação →", fwd_dis=not ss.blocos)
 
 
# ═══════════════════════════════════════════════════════════════════════════════
# STEP 4 — Validação
# ═══════════════════════════════════════════════════════════════════════════════
elif ss.step == 4:
    steps_nav()
    st.markdown("# Validação")
    erros, avisos, total_p = validar()
    enum = {str(e["num"]) for e in ss.escalas}
 
    m1,m2,m3,m4 = st.columns(4)
    m1.metric("Blocos", len(ss.blocos))
    m2.metric("Perguntas", total_p)
    m3.metric("Escalas", len(ss.escalas))
    m4.metric("Erros", len(erros), delta=str(len(erros)) if erros else None,
              delta_color="inverse" if erros else "off")
 
    st.markdown("")
 
    # Painel de status
    if not erros and not avisos:
        st.markdown('<div class="vsec v-ok"><div class="vtitle">✅ Tudo certo — pronto para exportar</div></div>', unsafe_allow_html=True)
    else:
        ce, cw = st.columns(2)
        with ce:
            if erros:
                items = "".join(f'<div class="vitem">✕ {e}</div>' for e in erros)
                st.markdown(f'<div class="vsec v-err"><div class="vtitle">❌ {len(erros)} erro(s) — obrigatório corrigir</div>{items}</div>', unsafe_allow_html=True)
            else:
                st.markdown('<div class="vsec v-ok"><div class="vtitle">✅ Sem erros</div></div>', unsafe_allow_html=True)
        with cw:
            if avisos:
                items = "".join(f'<div class="vitem">⚠ {a}</div>' for a in avisos)
                st.markdown(f'<div class="vsec v-wrn"><div class="vtitle">⚠️ {len(avisos)} aviso(s)</div>{items}</div>', unsafe_allow_html=True)
            else:
                st.markdown('<div class="vsec v-ok"><div class="vtitle">✅ Sem avisos</div></div>', unsafe_allow_html=True)
 
    st.divider()
    st.markdown('<div class="sh">Prévia por bloco</div>', unsafe_allow_html=True)
 
    for b in ss.blocos:
        pergs = b["perguntas"]
        eb, ab_ = [], []
        for p in pergs:
            pe,pa = val_p(p,enum); eb+=pe; ab_+=pa
        badge = f'<span class="bdg b-err">{len(eb)} erro(s)</span>' if eb else \
                (f'<span class="bdg b-wrn">{len(ab_)} aviso(s)</span>' if ab_ else '<span class="bdg b-ok">OK</span>')
 
        with st.expander(f"{b['nome']}  ·  {len(pergs)}p  {badge}", expanded=bool(eb)):
            rows_html = ""
            for i,p in enumerate(pergs):
                pe,pa = val_p(p,enum)
                cls = "r-err" if pe else ("r-wrn" if pa else "")
                st_ = "✕ "+"<br>".join(pe) if pe else ("⚠ "+"<br>".join(pa) if pa else "✓")
                rows_html += (f'<tr class="{cls}">'
                    f'<td style="color:#4b5260;font-family:monospace">{i+1}</td>'
                    f'<td>{p["texto"][:80]+("…" if len(p["texto"])>80 else "")}</td>'
                    f'<td>{p["competencia"] or "—"}</td>'
                    f'<td>{"Aberta" if p["aberta"]=="sim" else "Fechada"}</td>'
                    f'<td>{p["escala"] if p["aberta"]=="não" else "—"}</td>'
                    f'<td>{"Sim" if p["opcional"]=="sim" else "Não"}</td>'
                    f'<td style="font-size:11px">{st_}</td></tr>')
            st.markdown(
                '<table class="ptbl"><thead><tr>'
                '<th>#</th><th>Pergunta</th><th>Competência</th>'
                '<th>Tipo</th><th>Escala</th><th>Opcional</th><th>Status</th>'
                '</tr></thead><tbody>'+rows_html+'</tbody></table>',
                unsafe_allow_html=True)
 
    st.divider()
    st.markdown('<div class="sh">Escalas cadastradas</div>', unsafe_allow_html=True)
    for e in ss.escalas:
        tem = any(str(p).strip() for p in e["pontos"])
        with st.expander(f"Escala {e['num']} — {len(e['pontos'])} ponto(s)" + (" ⚠ sem descrições" if not tem else "")):
            if tem:
                prev = "".join(f'<div class="ept"><span class="en">{i+1}</span>{p or "—"}</div>'
                               for i,p in enumerate(e["pontos"]))
                st.markdown(f'<div class="erow">{prev}</div>', unsafe_allow_html=True)
            else:
                st.warning("Pontos de ancoragem não preenchidos.")
 
    nav_row(back=3, fwd=5, fwd_label="Configurar rodada →", fwd_dis=bool(erros))
 
 
# ═══════════════════════════════════════════════════════════════════════════════
# STEP 5 — Configuração
# ═══════════════════════════════════════════════════════════════════════════════
elif ss.step == 5:
    steps_nav()
    st.markdown("# Configuração da Rodada")
    st.markdown("Preencha os dados da rodada criada na plataforma.")
    cfg = ss.cfg
 
    st.markdown('<div class="sh">Acesso</div>', unsafe_allow_html=True)
    c1,c2 = st.columns(2)
    cfg["cliente"] = c1.text_input("Tenant do cliente *", value=cfg["cliente"], placeholder="ex: mindsight")
    cfg["token"]   = c2.text_input("Token do Performance *", value=cfg["token"], type="password")
 
    st.markdown('<div class="sh">IDs da rodada</div>', unsafe_allow_html=True)
    st.caption("Crie a rodada e os grupos na plataforma antes de preencher.")
    r1,r2 = st.columns(2)
    cfg["id_rodada"] = r1.text_input("ID da Rodada *",           value=cfg["id_rodada"])
    cfg["id_qs"]     = r2.text_input("ID do 1º Question Set *",  value=cfg["id_qs"])
    r3,r4 = st.columns(2)
    cfg["id_cq"] = r3.text_input("ID da 1ª Choice Question *", value=cfg["id_cq"])
    cfg["id_oq"] = r4.text_input("ID da 1ª Open Question *",   value=cfg["id_oq"])
 
    st.markdown('<div class="sh">Configurações adicionais</div>', unsafe_allow_html=True)
    a1,a2 = st.columns(2)
    cfg["val_min_nan"] = a1.text_input("Valor mín. p/ escala NaN", value=cfg["val_min_nan"])
    cfg["sso"]         = a2.checkbox("SSO?", value=cfg["sso"])
 
    ids_ok = all(cfg[k].strip() for k in ["id_rodada","id_qs","id_cq","id_oq"])
    if not cfg["cliente"].strip(): st.warning("Preencha o tenant.")
    if not ids_ok: st.warning("Preencha todos os IDs.")
 
    nav_row(back=4, fwd=6, fwd_label="Gerar bases →", fwd_dis=not(ids_ok and cfg["cliente"].strip()))
 
 
# ═══════════════════════════════════════════════════════════════════════════════
# STEP 6 — Geração
# ═══════════════════════════════════════════════════════════════════════════════
elif ss.step == 6:
    steps_nav()
    st.markdown("# Geração e Export")
 
    if ss.resultados is None:
        st.info("Clique para processar e gerar os arquivos JSON.")
        if st.button("▶  Gerar arquivos", type="primary", use_container_width=True):
            with st.spinner("Processando..."): log,res = run_script(ss.cfg)
            ss.log=log; ss.resultados=res; st.rerun()
    else:
        st.markdown('<div class="sh">Log de execução</div>', unsafe_allow_html=True)
        log_html = "".join(
            f'<div class="{"lok" if t=="ok" else "lerr" if t=="err" else "linf"}">'
            f'{"✓" if t=="ok" else "✕" if t=="err" else "→"} {msg}</div>'
            for t,msg in ss.log)
        st.markdown(f'<div class="logbox">{log_html}</div>', unsafe_allow_html=True)
 
        erros_exec = [m for t,m in ss.log if t=="err"]
        if erros_exec:
            st.error(f"{len(erros_exec)} erro(s). Volte e corrija os dados.")
        else:
            st.success("✅ Arquivos gerados com sucesso!")
            st.divider()
            st.markdown('<div class="sh">Downloads</div>', unsafe_allow_html=True)
            for fname, content in ss.resultados.items():
                c1,c2 = st.columns([5,1])
                c1.markdown(f"**{fname}**  \n<span style='font-size:11px;color:#6b7280'>{len(json.loads(content))} registro(s)</span>",
                            unsafe_allow_html=True)
                c2.download_button("⬇", data=content.encode("utf-8"), file_name=fname,
                                   mime="application/json", key=f"dl_{fname}", use_container_width=True)
            st.divider()
            st.download_button("⬇  Baixar todos (.zip)", data=build_zip(ss.resultados),
                               file_name=f"AVD_{ss.cfg['cliente']}.zip", mime="application/zip",
                               type="primary", use_container_width=True)
 
        st.divider()
        c1,_ = st.columns([1,5])
        with c1:
            if st.button("← Voltar"): go(5)
        if st.button("↺ Regerar"): ss.resultados=None; ss.log=[]; st.rerun()
