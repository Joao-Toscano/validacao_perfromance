
"""
AVD — Configuração e Geração de Bases
"""
import io, uuid, json, zipfile
import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from difflib import SequenceMatcher
 
st.set_page_config(page_title="AVD — Configuração", page_icon="📋",
                   layout="wide", initial_sidebar_state="collapsed")
 
# ── Força dark mode via query_params trick + CSS completo ─────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
 
/* Força fundo escuro em todos os contextos */
html, body { background-color: #0e1117 !important; }
[data-testid="stAppViewContainer"] { background-color: #0e1117 !important; }
[data-testid="stHeader"] { background-color: #0e1117 !important; }
[data-testid="stSidebar"] { background-color: #0e1117 !important; }
.main .block-container { background-color: #0e1117 !important; }
section[data-testid="stMain"] { background-color: #0e1117 !important; }
 
html,body,[class*="css"]{ font-family:'Inter',sans-serif; color:#f0ede8; }
.block-container{ padding:2rem 3rem 4rem !important; max-width:1100px; }
 
/* Steps nav */
.steps-nav{ display:flex; gap:0; border-radius:10px; overflow:hidden;
            border:1px solid rgba(255,255,255,.1); margin-bottom:2rem; }
.sn{ flex:1; padding:13px 6px; text-align:center; font-size:12px; font-weight:500;
     background:#1a1d27; color:rgba(255,255,255,.3);
     border-right:1px solid rgba(255,255,255,.08); }
.sn:last-child{ border-right:none; }
.sn.active{ background:rgba(129,140,248,.2); color:#818cf8; font-weight:700; }
.sn.done  { background:rgba(16,185,129,.1); color:#34d399; }
.sn-num{ display:inline-flex; width:19px; height:19px; border-radius:50%;
         border:1.5px solid currentColor; font-size:10px; align-items:center;
         justify-content:center; margin-right:5px; }
.sn.done .sn-num{ background:#34d399; color:#064e3b; border-color:#34d399; }
 
/* Section header */
.sh{ font-size:11px; font-weight:600; letter-spacing:1px; text-transform:uppercase;
     color:rgba(255,255,255,.3); margin:1.4rem 0 .5rem; }
 
/* Escala preview */
.esc-row{ display:flex; gap:4px; margin:8px 0; }
.esc-pt{ flex:1; background:#1e2130; border:1px solid rgba(255,255,255,.1);
         border-radius:6px; padding:8px 4px; text-align:center;
         font-size:10px; color:rgba(255,255,255,.5); line-height:1.4; }
.esc-n{ font-size:15px; font-weight:700; color:#fff; display:block; }
 
/* Validation cards */
.val-section{ border-radius:10px; padding:16px 20px; margin-bottom:12px; }
.val-ok  { background:rgba(52,211,153,.08); border:1px solid rgba(52,211,153,.2); }
.val-err { background:rgba(248,113,113,.08); border:1px solid rgba(248,113,113,.2); }
.val-wrn { background:rgba(251,191,36,.08);  border:1px solid rgba(251,191,36,.2); }
.val-title{ font-size:13px; font-weight:700; margin-bottom:10px; }
.val-ok   .val-title{ color:#34d399; }
.val-err  .val-title{ color:#f87171; }
.val-wrn  .val-title{ color:#fbbf24; }
.val-item{ font-size:12px; padding:4px 0; color:rgba(255,255,255,.7);
           border-bottom:1px solid rgba(255,255,255,.05); }
.val-item:last-child{ border-bottom:none; }
 
/* Bloco card na validação */
.bloco-val{ background:#1a1d27; border:1px solid rgba(255,255,255,.08);
            border-radius:10px; margin-bottom:10px; overflow:hidden; }
.bloco-val-header{ display:flex; align-items:center; gap:10px;
                   padding:12px 16px; border-bottom:1px solid rgba(255,255,255,.06); }
.bloco-val-nome{ font-size:13px; font-weight:600; flex:1; }
.bloco-val-body{ padding:12px 16px; }
 
/* Table */
.ptable{ width:100%; border-collapse:collapse; font-size:12px; }
.ptable th{ background:#1e2130; color:rgba(255,255,255,.35); font-size:10px;
            letter-spacing:.5px; text-transform:uppercase; padding:8px 12px;
            text-align:left; border-bottom:1px solid rgba(255,255,255,.08); }
.ptable td{ padding:8px 12px; border-bottom:1px solid rgba(255,255,255,.05);
            color:rgba(255,255,255,.75); vertical-align:top; }
.ptable tr.row-err td{ background:rgba(248,113,113,.05); }
.ptable tr.row-wrn td{ background:rgba(251,191,36,.04); }
 
/* Badges */
.badge{ font-size:10px; padding:2px 8px; border-radius:99px; font-weight:500;
        border:1px solid; margin-right:3px; display:inline-block; }
.b-ok  { color:#34d399; border-color:rgba(52,211,153,.3);  background:rgba(52,211,153,.08);  }
.b-err { color:#f87171; border-color:rgba(248,113,113,.3); background:rgba(248,113,113,.08); }
.b-wrn { color:#fbbf24; border-color:rgba(251,191,36,.3);  background:rgba(251,191,36,.08);  }
.b-blue{ color:#93c5fd; border-color:rgba(147,197,253,.3); background:rgba(147,197,253,.08); }
.b-gray{ color:rgba(255,255,255,.35); border-color:rgba(255,255,255,.1); background:rgba(255,255,255,.04); }
 
/* Log */
.logbox{ background:#0a0c12; border:1px solid rgba(255,255,255,.08); border-radius:8px;
         padding:14px 16px; font-family:monospace; font-size:12px;
         max-height:360px; overflow-y:auto; line-height:1.9; }
.log-ok { color:#34d399; } .log-err{ color:#f87171; }
.log-inf{ color:#93c5fd; } .log-wrn{ color:#fbbf24; }
 
/* Streamlit overrides para dark */
div[data-testid="stMetric"]{
  background:#1a1d27 !important;
  border:1px solid rgba(255,255,255,.1) !important;
  border-radius:10px; padding:14px 18px; }
div[data-testid="stMetricValue"]{ font-size:1.7rem !important; font-weight:700; }
div[data-testid="stTabs"] button{ background:#1a1d27 !important; color:rgba(255,255,255,.5) !important; }
div[data-testid="stTabs"] button[aria-selected="true"]{ color:#818cf8 !important; }
.stTextInput input, .stTextArea textarea, .stSelectbox select{
  background:#1a1d27 !important; color:#f0ede8 !important;
  border-color:rgba(255,255,255,.12) !important; }
.stCheckbox label{ color:#f0ede8 !important; }
.stMultiSelect [data-baseweb="select"]{ background:#1a1d27 !important; }
div[data-testid="stExpander"]{ background:#1a1d27 !important; border-color:rgba(255,255,255,.08) !important; }
div[data-testid="stDataFrame"]{ background:#1a1d27 !important; }
[data-testid="stFileUploader"]{ background:#1a1d27 !important; border-color:rgba(255,255,255,.12) !important; }
button[kind="primary"]{ background:#818cf8 !important; color:#fff !important; border:none !important; }
button[kind="secondary"]{ background:#1a1d27 !important; color:#f0ede8 !important;
                           border:1px solid rgba(255,255,255,.12) !important; }
</style>
""", unsafe_allow_html=True)
 
# ── Constants ─────────────────────────────────────────────────────────────────
COL_MAP = {
    "question_set":      ["nome do bloco de perguntas","nome do bloco","bloco","question_set","categoria"],
    "desc_question_set": ["descrição do bloco de perguntas","descrição do bloco","desc_question_set"],
    "competencia":       ["nome da competência","competência","competencia","skill","tema"],
    "pergunta":          ["pergunta","question","enunciado","texto","item"],
    "opcional":          ["essa pergunta é opcional","essa pergunta é opcional?","opcional","optional"],
    "aberta":            ["é uma pergunta aberta","é uma pergunta aberta?","aberta","open","tipo"],
    "escala":            ["qual o número do modelo da escala","qual o número do modelo da escala? (pergunta fechada)",
                          "número do modelo da escala","escala","scale"],
    "min_caracters":     ["número mínimo de caracteres","número mínimo de caracteres (opcional)","min caracters"],
    "max_caracters":     ["número máximo de caracteres","número máximo de caracteres (pergunta aberta)","max caracters"],
    "definicao":         ["descrição geral sobre a pergunta","definição","definicao","definition"],
    "g_auto":            ["autoavaliação","autoavaliacao"],
    "g_lider":           ["líder avaliando liderados","lider avaliando liderados","time"],
    "g_liderado":        ["liderados avaliando líder","liderados avaliando lider","gestor"],
    "g_par":             ["pares","par"],
    "g_stakeholder":     ["stakeholders","stakeholder"],
}
DESC_SCALE_SYNS = ["descrição_","descrição do primeiro","descrição do segundo","descrição do terceiro",
                   "descrição do quarto","descrição do quinto","duplique"]
GRUPOS_OPTS = ["Autoavaliação","Time","Gestor","Par","Stakeholder"]
DICT_CORES = {
    2:['#A8002E','#006499'], 3:['#A8002E','#FFC252','#006499'],
    4:['#A8002E','#FFC252','#35AEFF','#033D17'], 5:['#A8002E','#FFC252','#35AEFF','#006499','#033D17'],
    6:['#A8002E','#FA6F26','#FFC252','#35AEFF','#006499','#033D17'],
}
OUTPUT_COLS = ["question_set","desc_question_set","competencia","pergunta",
               "opcional","aberta","escala","min caracters","max caracters","definição"]
 
# ── Helpers ───────────────────────────────────────────────────────────────────
def uid(): return str(uuid.uuid4())[:8]
def norm_bool(v):
    v=str(v or "").strip().lower()
    return "sim" if v in("sim","yes","true","1","s","y") else "não"
def detect_col(hl, field):
    syns=COL_MAP.get(field,[])
    for i,h in enumerate(hl):
        hc=h.strip().rstrip("?").strip()
        if any(s in hc or hc in s for s in syns): return i
    return None
def sim_r(a,b): return SequenceMatcher(None,a.lower().strip(),b.lower().strip()).ratio()
def make_p(**kw):
    return {"id":uid(),"texto":kw.get("texto",""),"competencia":kw.get("competencia",""),
            "aberta":kw.get("aberta","não"),"escala":str(kw.get("escala","")),"opcional":kw.get("opcional","não"),
            "min_caracters":str(kw.get("min_c","")),"max_caracters":str(kw.get("max_c","")),"definicao":kw.get("definicao",""),
            "grupos":kw.get("grupos",list(GRUPOS_OPTS))}
def make_b(nome="",desc="",perguntas=None,origem="manual"):
    return {"id":uid(),"nome":nome,"desc":desc,"perguntas":perguntas or [],"origem":origem}
def make_e(num=None,pontos=None):
    if num is None: num=max((e["num"] for e in ss.escalas),default=0)+1
    try: num=int(float(str(num)))
    except: pass
    return {"id":uid(),"num":num,"pontos":pontos or ["",""]}
 
# ── Session state ─────────────────────────────────────────────────────────────
DEFS={"step":1,"escalas":[],"blocos":[],"bloco_sel":None,
      "cfg":{"cliente":"","token":"","id_rodada":"","id_qs":"","id_cq":"","id_oq":"",
             "val_min_nan":"0","autonomia_bp":True,"sso":True},
      "resultados":None,"log":[]}
for k,v in DEFS.items():
    if k not in st.session_state: st.session_state[k]=v
ss=st.session_state
 
STEPS=["1. Entrada","2. Validação","3. Configuração","4. Geração"]
def steps_nav():
    pills="".join(
        f'<div class="sn {"active" if i==ss.step else "done" if i<ss.step else ""}"><span class="sn-num">{"✓" if i<ss.step else i}</span>{l}</div>'
        for i,l in enumerate(STEPS,1))
    st.markdown(f'<div class="steps-nav">{pills}</div>',unsafe_allow_html=True)
def go(n): ss.step=n; st.rerun()
def nav_btns(back=None,fwd=None,fwd_label="Continuar →",fwd_disabled=False):
    c1,_,c2=st.columns([1,5,1])
    if back and c1.button("← Voltar",use_container_width=True): go(back)
    if fwd  and c2.button(fwd_label,type="primary",use_container_width=True,disabled=fwd_disabled): go(fwd)
 
# ── Import ────────────────────────────────────────────────────────────────────
def read_escalas_sheet(ws):
    """
    Lê aba Escalas em dois formatos:
    1. Formato Implementação_Performance (horizontal): blocos de 3 colunas (valor, desc, sep)
    2. Formato Base_Perguntas-2 (vertical): colunas descrição_1, descrição_2...
    """
    import re as _re
    data = list(ws.iter_rows(values_only=True))
    if len(data) < 2: return []
 
    header = data[0]
    header_str = [str(c or "").strip() for c in header]
 
    # ── Formato 1: horizontal (Implementação_Performance) ──
    # Detecta se há células com "modelo" ou "escala" no cabeçalho horizontal
    has_horizontal = any("modelo" in h.lower() or ("escala" in h.lower() and h.lower() != "escala")
                         for h in header_str if h)
 
    if has_horizontal:
        # Encontra pares (col_valor, col_desc) por bloco de 3 colunas
        pair_cols = []
        i = 0
        while i < len(header_str):
            h = header_str[i].lower()
            if h:  # célula não vazia = início de um bloco
                pair_cols.append((i, i+1))
                i += 3  # pula bloco de 3
            else:
                i += 1
 
        escalas = []
        for ci_val, ci_desc in pair_cols:
            nome = header_str[ci_val] if ci_val < len(header_str) else ""
            if not nome: continue
            vals, pontos = [], []
            for row in data[2:]:  # dados a partir da linha 2 (linha 1 é subheader)
                v    = row[ci_val]  if ci_val  < len(row) else None
                desc = row[ci_desc] if ci_desc < len(row) else None
                if v is not None:
                    try: vals.append(float(v))
                    except: pass
                if desc and str(desc).strip():
                    pontos.append(str(desc).strip())
            if pontos:
                num = int(min(vals)) if vals else len(escalas)+1
                max_val = int(max(vals)) if vals else len(pontos)
                escalas.append(make_e(num=num, pontos=pontos))
        return escalas
 
    # ── Formato 2: vertical (Base_Perguntas-2) ──
    headers_l = [h.lower() for h in header_str]
    desc_cols = [i for i,h in enumerate(headers_l) if "descrição_" in h]
    escalas = []
    for row in data[1:]:
        if not row[0]: continue
        try: num = int(float(str(row[0])))
        except: num = str(row[0])
        pontos = [str(row[i] or "").strip() for i in desc_cols if i < len(row) and row[i]]
        if pontos: escalas.append(make_e(num=num, pontos=pontos))
    return escalas
 
def import_xlsx(file_bytes):
    wb=load_workbook(io.BytesIO(file_bytes),read_only=True,data_only=True)
    msgs=[]
 
    # 1. Tenta ler aba dedicada de Escalas (Base_Perguntas-2 ou similar)
    escalas_out=[]
    for sn in wb.sheetnames:
        if "escala" in sn.lower():
            escalas_out=read_escalas_sheet(wb[sn])
            if escalas_out:
                msgs.append(f"✓ Escalas lidas da aba **{sn}** ({len(escalas_out)} escala(s)).")
                break
 
    # 2. Se não achou aba de escalas, tenta extrair das colunas de descrição nas abas de perguntas
    if not escalas_out:
        emap={}
        for sn in wb.sheetnames:
            ws=wb[sn]; data=list(ws.iter_rows(values_only=True))
            if not data: continue
            hl=[str(c or "").lower().strip() for c in data[0]]
            si=detect_col(hl,"escala")
            di=[i for i,h in enumerate(hl) if any(s in h for s in DESC_SCALE_SYNS)]
            for row in data[1:]:
                if si is None or si>=len(row): continue
                num=str(row[si] or "").strip()
                if not num: continue
                descs=[str(row[i] or "").strip() for i in di if i<len(row) and row[i]]
                if num not in emap: emap[num]=descs
                elif descs and not emap[num]: emap[num]=descs
        for ns,descs in emap.items():
            try: n=int(float(ns))
            except: n=ns
            escalas_out.append(make_e(num=n,pontos=descs if descs else ["",""]))
        if escalas_out:
            nums=[str(e["num"]) for e in escalas_out]
            msgs.append(f"⚠ Escalas extraídas das colunas de descrição: {', '.join(nums)} — **verifique os pontos de ancoragem**.")
        else:
            # 3. Só os números referenciados, sem descrições
            nums_ref=set()
            for sn in wb.sheetnames:
                ws=wb[sn]; data=list(ws.iter_rows(values_only=True))
                if not data: continue
                hl=[str(c or "").lower().strip() for c in data[0]]
                si=detect_col(hl,"escala")
                if si is None: continue
                for row in data[1:]:
                    v=str(row[si] or "").strip() if si<len(row) else ""
                    if v and v not in ("none","nan"): nums_ref.add(v)
            for ns in sorted(nums_ref):
                try: n=int(float(ns))
                except: n=ns
                escalas_out.append(make_e(num=n,pontos=["",""]))
            if escalas_out:
                msgs.append(f"⚠ Encontradas referências às escalas: **{', '.join(str(e['num']) for e in escalas_out)}** — sem descrições. Preencha os pontos de ancoragem manualmente ou importe o arquivo Base_Perguntas-2.xlsx.")
 
    # Blocos e perguntas
    bmap={}
    for sn in wb.sheetnames:
        if "escala" in sn.lower(): continue   # pula aba de escalas
        ws=wb[sn]; data=list(ws.iter_rows(values_only=True))
        if len(data)<2: continue
        hl=[str(c or "").lower().strip() for c in data[0]]
        idx={f:detect_col(hl,f) for f in COL_MAP}
        def get(row,f):
            i=idx.get(f)
            return str(row[i] or "").strip() if(i is not None and i<len(row)) else ""
        def get_bool(row,f):
            i=idx.get(f)
            if i is None or i>=len(row): return True
            return norm_bool(str(row[i] or ""))=="sim"
        for row in data[1:]:
            if not any(c not in(None,"") for c in row): continue
            bn=get(row,"question_set"); pt=get(row,"pergunta")
            if not bn and not pt: continue
            grupos=[gl for gf,gl in [("g_auto","Autoavaliação"),("g_lider","Time"),
                                      ("g_liderado","Gestor"),("g_par","Par"),("g_stakeholder","Stakeholder")]
                    if get_bool(row,gf)]
            if not grupos: grupos=list(GRUPOS_OPTS)
            p=make_p(texto=pt,competencia=get(row,"competencia"),
                     aberta=norm_bool(get(row,"aberta") or "não"),escala=get(row,"escala"),
                     opcional=norm_bool(get(row,"opcional") or "não"),
                     min_c=get(row,"min_caracters"),max_c=get(row,"max_caracters"),
                     definicao=get(row,"definicao"),grupos=grupos)
            bmap.setdefault(bn,{}).setdefault(sn,[]).append(p)
    wb.close()
 
    blocos_out=[]
    for bn,abas in bmap.items():
        al=list(abas.keys())
        if len(al)==1:
            blocos_out.append(make_b(nome=bn,perguntas=abas[al[0]],origem="import")); continue
        sets=[set(p["texto"] for p in abas[a]) for a in al]
        if all(s==sets[0] for s in sets[1:]):
            blocos_out.append(make_b(nome=bn,perguntas=abas[al[0]],origem="import"))
            msgs.append(f"✓ **{bn}** idêntico em {len(al)} abas → mesclado.")
        else:
            for a in al:
                blocos_out.append(make_b(nome=f"{bn} — {a}",perguntas=abas[a],origem="import"))
            msgs.append(f"⚠ **{bn}** diferente em {len(al)} abas → renomeado com sufixo.")
    nomes=[b["nome"] for b in blocos_out]
    for i,n1 in enumerate(nomes):
        for n2 in nomes[i+1:]:
            if sim_r(n1,n2)>0.85: msgs.append(f"⚠ Blocos parecidos: **{n1}** e **{n2}**")
    return blocos_out,escalas_out,msgs
 
# ── Validação ─────────────────────────────────────────────────────────────────
def validar():
    erros,avisos=[],[]
    enum={str(e["num"]) for e in ss.escalas}
    if not ss.escalas: erros.append("Nenhuma escala cadastrada.")
    if not ss.blocos:  erros.append("Nenhum bloco cadastrado.")
    seen=set()
    nomes=[b["nome"].strip() for b in ss.blocos]
    for n in nomes:
        if n in seen: erros.append(f'Bloco duplicado: "{n}"')
        seen.add(n)
    for i,n1 in enumerate(nomes):
        for n2 in nomes[i+1:]:
            if n1!=n2 and sim_r(n1,n2)>0.85: avisos.append(f'Nomes parecidos: "{n1}" × "{n2}"')
    total=0
    for b in ss.blocos:
        if not b["nome"].strip(): erros.append("Bloco sem nome.")
        if not b["perguntas"]:   avisos.append(f'"{b["nome"]}" sem perguntas.')
        for p in b["perguntas"]:
            total+=1
            if not p["texto"].strip(): erros.append(f'Pergunta vazia em "{b["nome"]}".')
            if p["aberta"]=="não":
                if not p["escala"]: avisos.append(f'Fechada sem escala em "{b["nome"]}".')
                elif p["escala"] not in enum: erros.append(f'Escala {p["escala"]} não existe em "{b["nome"]}".')
            if p["aberta"]=="sim" and not p["max_caracters"]: avisos.append(f'Aberta sem máx chars em "{b["nome"]}".')
    return erros,avisos,total
 
# ── Validação por pergunta (para tabela colorida) ─────────────────────────────
def val_pergunta(p, enum):
    erros,avisos=[],[]
    if not p["texto"].strip(): erros.append("Texto vazio")
    if p["aberta"]=="não":
        if not p["escala"]: avisos.append("Sem escala")
        elif p["escala"] not in enum: erros.append(f"Escala {p['escala']} inexistente")
    if p["aberta"]=="sim" and not p["max_caracters"]: avisos.append("Sem máx chars")
    return erros,avisos
 
# ── Script ────────────────────────────────────────────────────────────────────
def run_script(cfg):
    log=[]
    def L(msg,t="inf"): log.append((t,msg))
    try:
        id_rodada=int(cfg["id_rodada"]); id_qs=int(cfg["id_qs"])
        id_cq=int(cfg["id_cq"]); id_oq=int(cfg["id_oq"])
        val_min_nan=float(cfg.get("val_min_nan") or "0")
    except Exception as e: L(f"IDs inválidos: {e}","err"); return log,{}
 
    rows_p=[]
    for b in ss.blocos:
        for p in b["perguntas"]:
            rows_p.append({"question_set":b["nome"],"desc_question_set":b["desc"],
                "competencia":p["competencia"],"pergunta":p["texto"],"opcional":p["opcional"],"aberta":p["aberta"],
                "escala":int(float(p["escala"])) if p["aberta"]=="não" and p["escala"] else np.nan,
                "min caracters":int(p["min_caracters"]) if p["min_caracters"] else np.nan,
                "max caracters":int(p["max_caracters"]) if p["max_caracters"] else np.nan,
                "definição":p["definicao"] if p["definicao"] else np.nan})
    bp=pd.DataFrame(rows_p)
 
    rows_e=[]
    for e in ss.escalas:
        r={"escala":e["num"],
           "min":float(e["pontos"][0]) if e["pontos"] and e["pontos"][0] else np.nan,
           "max":float(e["pontos"][-1]) if e["pontos"] and e["pontos"][-1] else np.nan}
        for i,pt in enumerate(e["pontos"],1): r[f"descrição_{i}"]=pt if pt else np.nan
        rows_e.append(r)
    be=pd.DataFrame(rows_e)
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
        opcional=0 if str(row['opcional']).lower()=='não' else 1
        min_c=0 if pd.isna(row.get('min caracters')) else int(row['min caracters'])
        max_c=int(row['max caracters'])
        qs=bqs[bqs['name']==row['question_set']]
        reg={'id':id_oq if first else id_oq+cnt_oq,'order':cnt_p,'key':id_oq if first else id_oq+cnt_oq,
             'description':desc,'info':'','is_optional':opcional,'min_answer_length':min_c,
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
        opcional=0 if str(row['opcional']).lower()=='não' else 1
        qs=bqs[bqs['name']==row['question_set']]
        reg={'id':id_cq if first else id_cq+cnt_cq,'order':cnt_p,'key':id_cq if first else id_cq+cnt_cq,
             'description':desc,'info':'','is_optional':opcional,'question_set':qs['id'].iloc[0],
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
            if flag_nan:
                minimo=val_min_nan; div=(maximo-float(minimo))/max(cnt_pts-2,1)
            else:
                minimo=float(minimo); div=(maximo-minimo)/max(cnt_pts-1,1)
            cores=DICT_CORES.get(cnt_pts,['#000000']*cnt_pts)
            ci=0
            for col in cols_d:
                if pd.isna(re_.get(col)): continue
                if flag_nan:
                    reg={'id':'','question':row['id'],'value':np.nan,'content':str(re_[col]),'color':'#000000'}
                    flag_nan=False
                else:
                    reg={'id':'','question':row['id'],'value':round(float(minimo)+ci*div,2),
                         'content':str(re_[col]),'color':cores[ci] if ci<len(cores) else '#000000'}
                    ci+=1
                besc.loc[len(besc)]=reg
    L(f"✓ {len(besc)} itens de escala","ok")
 
    cliente=cfg["cliente"]
    resultados={
        f"Question Sets {cliente}.json":    bqs.to_json(orient='records',force_ascii=False),
        f"Open Questions {cliente}.json":   boq.to_json(orient='records',force_ascii=False),
        f"Choice Questions {cliente}.json": bcq.drop(columns=['escala'],errors='ignore').to_json(orient='records',force_ascii=False),
        f"Escalas {cliente}.json":          besc.to_json(orient='records',force_ascii=False),
    }
    L("✓ Todos os JSONs gerados","ok")
    return log,resultados
 
def build_zip(res):
    buf=io.BytesIO()
    with zipfile.ZipFile(buf,'w',zipfile.ZIP_DEFLATED) as z:
        [z.writestr(fn,c.encode('utf-8')) for fn,c in res.items()]
    return buf.getvalue()
 
 
# ═══════════════════════════════════════════════════════════════════════════════
# STEP 1 — Entrada
# ═══════════════════════════════════════════════════════════════════════════════
if ss.step==1:
    steps_nav()
    st.markdown("# Configuração de Avaliação")
 
    tab_imp,tab_man=st.tabs(["📥  Importar planilha","✏️  Preencher manualmente"])
 
    with tab_imp:
        st.caption("O sistema aceita o arquivo do cliente (qualquer número de abas) "
                   "**ou** o arquivo Base_Perguntas-2.xlsx que já tem a aba Escalas separada.")
        up=st.file_uploader("Arquivo .xlsx",type=["xlsx","xls"],label_visibility="collapsed")
        if up:
            with st.spinner("Lendo arquivo..."):
                bi,ei,mi=import_xlsx(up.read())
            c1,c2,c3=st.columns(3)
            c1.metric("Blocos",len(bi))
            c2.metric("Perguntas",sum(len(b["perguntas"]) for b in bi))
            c3.metric("Escalas",len(ei))
            if mi:
                with st.expander("📋 O que foi detectado e ajustado"):
                    for m in mi: st.markdown(m)
            if ei and any(len(e["pontos"])>=2 and all(e["pontos"]) for e in ei):
                st.success("✅ Escalas com pontos de ancoragem encontradas.")
            elif ei:
                st.warning("⚠️ Escalas detectadas sem pontos de ancoragem. Preencha-os na aba Manual ou importe o Base_Perguntas-2.xlsx.")
            ca,cb=st.columns(2)
            if ca.button("✓ Usar esses dados",type="primary",key="ir"):
                ss.blocos=bi; ss.escalas=ei; ss.bloco_sel=None; go(2)
            if ss.blocos and cb.button("＋ Adicionar aos dados existentes",key="ia"):
                ex={b["nome"]:b for b in ss.blocos}
                for nb in bi:
                    if nb["nome"] in ex: ex[nb["nome"]]["perguntas"].extend(nb["perguntas"])
                    else: ss.blocos.append(nb)
                for ne in ei:
                    if not any(e["num"]==ne["num"] for e in ss.escalas): ss.escalas.append(ne)
                go(2)
 
    with tab_man:
        escala_opts=[""]+[str(e["num"]) for e in ss.escalas]
        st.markdown('<div class="sh">Escalas</div>',unsafe_allow_html=True)
        st.caption("Defina as escalas e seus pontos de ancoragem.")
        for ei_idx,esc in enumerate(ss.escalas):
            with st.container(border=True):
                h1,h2=st.columns([8,1])
                with h1:
                    nv=st.number_input("Nº",min_value=1,
                                       value=int(esc["num"]) if str(esc["num"]).replace(".","").isdigit() else ei_idx+1,
                                       key=f"en_{esc['id']}")
                    esc["num"]=nv
                    if any(str(p).strip() for p in esc["pontos"]):
                        prev="".join(f'<div class="esc-pt"><span class="esc-n">{i+1}</span>{p or "—"}</div>'
                                     for i,p in enumerate(esc["pontos"]))
                        st.markdown(f'<div class="esc-row">{prev}</div>',unsafe_allow_html=True)
                    with st.expander("✏ Editar pontos"):
                        n_pts=st.number_input("Pontos",2,10,len(esc["pontos"]),key=f"np_{esc['id']}")
                        while len(esc["pontos"])<n_pts: esc["pontos"].append("")
                        esc["pontos"]=esc["pontos"][:n_pts]
                        pc=st.columns(min(n_pts,5))
                        for pi in range(n_pts):
                            esc["pontos"][pi]=pc[pi%len(pc)].text_input(
                                f"Ponto {pi+1}",value=esc["pontos"][pi],key=f"pt_{esc['id']}_{pi}",placeholder=f"Nível {pi+1}")
                with h2:
                    st.markdown("<br><br>",unsafe_allow_html=True)
                    if st.button("✕",key=f"de_{esc['id']}"): ss.escalas.pop(ei_idx); st.rerun()
        if st.button("＋ Escala",key="ae"): ss.escalas.append(make_e()); st.rerun()
 
        st.divider()
        st.markdown('<div class="sh">Blocos & Perguntas</div>',unsafe_allow_html=True)
        col_bl,col_pr=st.columns([5,7],gap="large")
        with col_bl:
            st.caption(f"{len(ss.blocos)} bloco(s)")
            for bi_idx,bloco in enumerate(ss.blocos):
                np_=len(bloco["perguntas"])
                ne_=sum(1 for p in bloco["perguntas"] if not p["texto"] or(p["aberta"]=="não" and not p["escala"]))
                is_s=ss.bloco_sel==bloco["id"]
                with st.container(border=is_s):
                    st.markdown(f"**{bloco['nome'] or '(sem nome)'}**  \n"
                                f"<span style='font-size:11px;color:rgba(255,255,255,.4)'>"
                                f"{np_}p {'· '+str(ne_)+' erro(s)' if ne_ else '· OK'}</span>",
                                unsafe_allow_html=True)
                    b1,b2=st.columns([3,1])
                    if b1.button("Editar",key=f"sb_{bloco['id']}",use_container_width=True):
                        ss.bloco_sel=bloco["id"] if not is_s else None; st.rerun()
                    if b2.button("✕",key=f"db_{bloco['id']}"):
                        ss.blocos=[b for b in ss.blocos if b["id"]!=bloco["id"]]
                        if ss.bloco_sel==bloco["id"]: ss.bloco_sel=None
                        st.rerun()
            st.markdown("")
            if st.button("＋ Novo bloco",use_container_width=True,key="nb"):
                nb=make_b(nome="Novo bloco"); ss.blocos.append(nb); ss.bloco_sel=nb["id"]; st.rerun()
 
        with col_pr:
            bloco=next((b for b in ss.blocos if b["id"]==ss.bloco_sel),None)
            if bloco is None:
                st.info("👈 Selecione ou crie um bloco para editar suas perguntas.")
            else:
                bloco["nome"]=st.text_input("Nome do bloco *",value=bloco["nome"],key=f"bn_{bloco['id']}")
                bloco["desc"]=st.text_input("Descrição (opcional)",value=bloco["desc"],key=f"bd_{bloco['id']}")
 
                # Tabela de perguntas existentes
                if bloco["perguntas"]:
                    st.markdown(f"**{len(bloco['perguntas'])} pergunta(s):**")
                    df_tbl=pd.DataFrame([{
                        "#":i+1,"Pergunta":p["texto"][:65]+("…" if len(p["texto"])>65 else ""),
                        "Tipo":"Aberta" if p["aberta"]=="sim" else "Fechada",
                        "Escala":p["escala"] if p["aberta"]=="não" else "—",
                        "Opcional":"Sim" if p["opcional"]=="sim" else "Não",
                    } for i,p in enumerate(bloco["perguntas"])])
                    st.dataframe(df_tbl,use_container_width=True,hide_index=True)
                    rem_cols=st.columns(min(len(bloco["perguntas"]),6))
                    for i,p in enumerate(bloco["perguntas"]):
                        if rem_cols[i%len(rem_cols)].button(f"✕{i+1}",key=f"dp_{p['id']}",use_container_width=True):
                            bloco["perguntas"].pop(i); st.rerun()
                    st.divider()
 
                # Formulário adicionar
                st.markdown("**Adicionar pergunta:**")
                with st.container(border=True):
                    novo_txt=st.text_area("Texto *",key=f"ntx_{bloco['id']}",
                                          placeholder="Texto da pergunta...",height=75,label_visibility="collapsed")
                    fa,fb=st.columns(2)
                    novo_comp=fa.text_input("Competência",key=f"nc_{bloco['id']}",placeholder="Ex: Liderança")
                    novo_tipo=fb.selectbox("Tipo",["Fechada","Aberta"],key=f"nty_{bloco['id']}")
                    novo_aberta="não" if novo_tipo=="Fechada" else "sim"
                    fc,fd=st.columns(2)
                    if novo_aberta=="não":
                        novo_escala=fc.selectbox("Escala *",escala_opts,key=f"ne_{bloco['id']}")
                        novo_min=""; novo_max=""
                    else:
                        novo_escala=""
                        novo_min=fc.text_input("Mín chars",key=f"nmi_{bloco['id']}",placeholder="0")
                        novo_max=fd.text_input("Máx chars",key=f"nma_{bloco['id']}",placeholder="10000")
                    novo_grupos=st.multiselect("Grupos avaliativos *",GRUPOS_OPTS,default=GRUPOS_OPTS,key=f"ngr_{bloco['id']}")
                    fe,ff=st.columns(2)
                    novo_opc=fe.checkbox("Opcional?",key=f"nop_{bloco['id']}")
                    novo_def=ff.text_input("Definição",key=f"ndf_{bloco['id']}")
                    if st.button("＋ Adicionar ao bloco",type="primary",use_container_width=True,key=f"add_{bloco['id']}"):
                        if not novo_txt.strip(): st.error("Texto obrigatório.")
                        elif novo_aberta=="não" and not novo_escala: st.error("Selecione uma escala.")
                        else:
                            bloco["perguntas"].append(make_p(texto=novo_txt,competencia=novo_comp,aberta=novo_aberta,
                                escala=novo_escala,opcional="sim" if novo_opc else "não",
                                min_c=novo_min,max_c=novo_max,definicao=novo_def,grupos=novo_grupos))
                            st.rerun()
 
    st.divider()
    erros,_,_=validar()
    _,c2=st.columns([5,1])
    with c2:
        if st.button("Ir para Validação →",type="primary",use_container_width=True,disabled=not ss.blocos): go(2)
 
 
# ═══════════════════════════════════════════════════════════════════════════════
# STEP 2 — Validação (redesenhada)
# ═══════════════════════════════════════════════════════════════════════════════
elif ss.step==2:
    steps_nav()
    st.markdown("# Validação")
    erros,avisos,total_p=validar()
    enum={str(e["num"]) for e in ss.escalas}
 
    # Métricas
    m1,m2,m3,m4=st.columns(4)
    m1.metric("Blocos",len(ss.blocos))
    m2.metric("Perguntas",total_p)
    m3.metric("Escalas",len(ss.escalas))
    m4.metric("Erros",len(erros),delta=str(len(erros)) if erros else None,delta_color="inverse" if erros else "off")
 
    st.markdown("")
 
    # Painel de status geral
    if not erros and not avisos:
        st.markdown('<div class="val-section val-ok"><div class="val-title">✅ Tudo certo — pronto para exportar</div></div>',unsafe_allow_html=True)
    else:
        col_e,col_w=st.columns(2)
        with col_e:
            if erros:
                items="".join(f'<div class="val-item">✕ {e}</div>' for e in erros)
                st.markdown(f'<div class="val-section val-err"><div class="val-title">❌ {len(erros)} erro(s) — obrigatório corrigir</div>{items}</div>',unsafe_allow_html=True)
            else:
                st.markdown('<div class="val-section val-ok"><div class="val-title">✅ Sem erros</div></div>',unsafe_allow_html=True)
        with col_w:
            if avisos:
                items="".join(f'<div class="val-item">⚠ {a}</div>' for a in avisos)
                st.markdown(f'<div class="val-section val-wrn"><div class="val-title">⚠️ {len(avisos)} aviso(s)</div>{items}</div>',unsafe_allow_html=True)
            else:
                st.markdown('<div class="val-section val-ok"><div class="val-title">✅ Sem avisos</div></div>',unsafe_allow_html=True)
 
    st.divider()
 
    # Prévia por bloco — tabela colorida por status
    st.markdown('<div class="sh">Prévia por bloco</div>',unsafe_allow_html=True)
    for b in ss.blocos:
        pergs=b["perguntas"]
        erros_b=[]; avisos_b=[]
        for p in pergs:
            pe,pa=val_pergunta(p,enum)
            erros_b+=pe; avisos_b+=pa
        if erros_b:    status_badge=f'<span class="badge b-err">{len(erros_b)} erro(s)</span>'
        elif avisos_b: status_badge=f'<span class="badge b-wrn">{len(avisos_b)} aviso(s)</span>'
        else:          status_badge='<span class="badge b-ok">OK</span>'
        count_badge=f'<span class="badge b-blue">{len(pergs)}p</span>'
 
        with st.expander(f"{b['nome']}  —  {len(pergs)} pergunta(s)",expanded=bool(erros_b)):
            # Tabela com linha colorida por status
            rows_html=""
            for i,p in enumerate(pergs):
                pe,pa=val_pergunta(p,enum)
                cls="row-err" if pe else("row-wrn" if pa else "")
                status="✕ "+"<br>".join(pe) if pe else("⚠ "+"<br>".join(pa) if pa else "✓")
                rows_html+=(f'<tr class="{cls}">'
                    f'<td style="color:rgba(255,255,255,.3);font-family:monospace">{i+1}</td>'
                    f'<td>{p["texto"][:80]+("…" if len(p["texto"])>80 else "")}</td>'
                    f'<td>{p["competencia"] or "—"}</td>'
                    f'<td>{"Aberta" if p["aberta"]=="sim" else "Fechada"}</td>'
                    f'<td>{p["escala"] if p["aberta"]=="não" else "—"}</td>'
                    f'<td>{"Sim" if p["opcional"]=="sim" else "Não"}</td>'
                    f'<td>{", ".join(p.get("grupos",[]))}</td>'
                    f'<td style="font-size:11px">{status}</td>'
                    f'</tr>')
            st.markdown(
                f'<table class="ptable"><thead><tr>'
                f'<th>#</th><th>Pergunta</th><th>Competência</th>'
                f'<th>Tipo</th><th>Escala</th><th>Opcional</th><th>Grupos</th><th>Status</th>'
                f'</tr></thead><tbody>{rows_html}</tbody></table>',
                unsafe_allow_html=True)
 
    # Escalas
    st.divider()
    st.markdown('<div class="sh">Escalas</div>',unsafe_allow_html=True)
    if ss.escalas:
        for e in ss.escalas:
            tem_pts=any(str(p).strip() for p in e["pontos"])
            with st.expander(f"Escala {e['num']} — {len(e['pontos'])} ponto(s)" + (" ⚠ sem descrições" if not tem_pts else "")):
                if tem_pts:
                    prev="".join(f'<div class="esc-pt"><span class="esc-n">{i+1}</span>{p or "—"}</div>'
                                 for i,p in enumerate(e["pontos"]))
                    st.markdown(f'<div class="esc-row">{prev}</div>',unsafe_allow_html=True)
                else:
                    st.warning("Pontos de ancoragem não preenchidos. Volte à Entrada para preencher.")
    else:
        st.warning("Nenhuma escala cadastrada.")
 
    nav_btns(back=1,fwd=3,fwd_label="Configurar rodada →",fwd_disabled=bool(erros))
 
 
# ═══════════════════════════════════════════════════════════════════════════════
# STEP 3 — Configuração
# ═══════════════════════════════════════════════════════════════════════════════
elif ss.step==3:
    steps_nav()
    st.markdown("# Configuração da Rodada")
    cfg=ss.cfg
    st.markdown('<div class="sh">Acesso</div>',unsafe_allow_html=True)
    c1,c2=st.columns(2)
    cfg["cliente"]=c1.text_input("Tenant *",value=cfg["cliente"],placeholder="ex: mindsight")
    cfg["token"]=c2.text_input("Token *",value=cfg["token"],type="password")
    st.markdown('<div class="sh">IDs da rodada</div>',unsafe_allow_html=True)
    st.caption("Crie a rodada e os grupos na plataforma antes de preencher os IDs.")
    r1,r2=st.columns(2)
    cfg["id_rodada"]=r1.text_input("ID da Rodada *",value=cfg["id_rodada"])
    cfg["id_qs"]=r2.text_input("ID do 1º Question Set *",value=cfg["id_qs"])
    r3,r4=st.columns(2)
    cfg["id_cq"]=r3.text_input("ID da 1ª Choice Question *",value=cfg["id_cq"])
    cfg["id_oq"]=r4.text_input("ID da 1ª Open Question *",value=cfg["id_oq"])
    st.markdown('<div class="sh">Configurações adicionais</div>',unsafe_allow_html=True)
    a1,a2,a3=st.columns(3)
    cfg["val_min_nan"]=a1.text_input("Valor mín. p/ escala NaN",value=cfg["val_min_nan"])
    cfg["autonomia_bp"]=a2.checkbox("Autonomia de BPs?",value=cfg["autonomia_bp"])
    cfg["sso"]=a3.checkbox("SSO?",value=cfg["sso"])
    ids_ok=all(cfg[k].strip() for k in["id_rodada","id_qs","id_cq","id_oq"])
    if not cfg["cliente"].strip(): st.warning("Preencha o tenant.")
    if not ids_ok: st.warning("Preencha todos os IDs.")
    nav_btns(back=2,fwd=4,fwd_label="Gerar bases →",fwd_disabled=not(ids_ok and cfg["cliente"].strip()))
 
 
# ═══════════════════════════════════════════════════════════════════════════════
# STEP 4 — Geração
# ═══════════════════════════════════════════════════════════════════════════════
elif ss.step==4:
    steps_nav()
    st.markdown("# Geração e Export")
    if ss.resultados is None:
        st.info("Clique para processar e gerar os arquivos JSON.")
        if st.button("▶  Gerar arquivos",type="primary",use_container_width=True):
            with st.spinner("Processando..."): log,res=run_script(ss.cfg)
            ss.log=log; ss.resultados=res; st.rerun()
    else:
        st.markdown('<div class="sh">Log</div>',unsafe_allow_html=True)
        log_html="".join(f'<div class="log-{t}">{"✓" if t=="ok" else "✕" if t=="err" else "→"} {msg}</div>'
                         for t,msg in ss.log)
        st.markdown(f'<div class="logbox">{log_html}</div>',unsafe_allow_html=True)
        erros_exec=[m for t,m in ss.log if t=="err"]
        if erros_exec:
            st.error(f"{len(erros_exec)} erro(s). Volte e corrija.")
        else:
            st.success("✅ Arquivos gerados!")
            st.divider()
            for fname,content in ss.resultados.items():
                c1,c2=st.columns([5,1])
                c1.markdown(f"**{fname}**  \n<span style='font-size:11px;color:rgba(255,255,255,.4)'>{len(json.loads(content))} registro(s)</span>",unsafe_allow_html=True)
                c2.download_button("⬇",data=content.encode("utf-8"),file_name=fname,
                                   mime="application/json",key=f"dl_{fname}",use_container_width=True)
            st.divider()
            st.download_button("⬇  Baixar todos (.zip)",data=build_zip(ss.resultados),
                               file_name=f"AVD_{ss.cfg['cliente']}.zip",mime="application/zip",
                               type="primary",use_container_width=True)
        st.divider()
        c1,_=st.columns([1,5])
        with c1:
            if st.button("← Voltar"): go(3)
        if st.button("↺ Regerar"): ss.resultados=None; ss.log=[]; st.rerun()
