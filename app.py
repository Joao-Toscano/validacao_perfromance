
"""
AVD — Configuração e Geração de Bases
Fluxo: Entrada → Visualização/Validação → Configuração da Rodada → Geração e Export
"""
 
import io, uuid, json, zipfile, re
import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from difflib import SequenceMatcher
 
st.set_page_config(page_title="AVD — Configuração", page_icon="📋",
                   layout="wide", initial_sidebar_state="collapsed")
 
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
html,body,[class*="css"]{ font-family:'Inter',sans-serif; }
.block-container{ padding:2rem 3rem 4rem !important; max-width:1100px; }
 
/* Steps nav */
.steps-nav{ display:flex; gap:0; border-radius:10px; overflow:hidden;
            border:1px solid rgba(255,255,255,.07); margin-bottom:2rem; }
.sn{ flex:1; padding:13px 6px; text-align:center; font-size:12px; font-weight:500;
     background:rgba(255,255,255,.03); color:rgba(255,255,255,.3);
     border-right:1px solid rgba(255,255,255,.07); cursor:default; }
.sn:last-child{ border-right:none; }
.sn.active{ background:rgba(99,102,241,.15); color:#818cf8; font-weight:700; }
.sn.done  { background:rgba(16,185,129,.07); color:#34d399; cursor:pointer; }
.sn-num{ display:inline-flex; width:19px; height:19px; border-radius:50%;
         border:1.5px solid currentColor; font-size:10px; align-items:center;
         justify-content:center; margin-right:5px; }
.sn.done .sn-num{ background:#34d399; color:#064e3b; border-color:#34d399; }
 
/* Cards */
.card{ background:rgba(255,255,255,.04); border:1px solid rgba(255,255,255,.08);
       border-radius:10px; padding:16px 20px; margin-bottom:8px; }
.card.sel{ border-color:#818cf8; background:rgba(99,102,241,.08); }
.card-title{ font-size:13px; font-weight:600; margin-bottom:4px; }
.card-sub{ font-size:11px; color:rgba(255,255,255,.4); }
 
/* Badges */
.badge{ font-size:10px; padding:2px 9px; border-radius:99px;
        font-weight:500; border:1px solid; margin-right:3px; }
.b-blue { color:#93c5fd; border-color:rgba(147,197,253,.3); background:rgba(147,197,253,.07); }
.b-green{ color:#86efac; border-color:rgba(134,239,172,.3); background:rgba(134,239,172,.07); }
.b-red  { color:#fca5a5; border-color:rgba(252,165,165,.3); background:rgba(252,165,165,.07); }
.b-gray { color:rgba(255,255,255,.35); border-color:rgba(255,255,255,.1); background:rgba(255,255,255,.03); }
.b-purp { color:#c4b5fd; border-color:rgba(196,181,253,.3); background:rgba(196,181,253,.07); }
 
/* Escala preview */
.esc-row{ display:flex; gap:4px; margin:8px 0; }
.esc-pt { flex:1; background:rgba(255,255,255,.06); border:1px solid rgba(255,255,255,.08);
           border-radius:6px; padding:8px 4px; text-align:center;
           font-size:10px; color:rgba(255,255,255,.5); line-height:1.4; }
.esc-n  { font-size:15px; font-weight:700; color:#fff; display:block; }
 
/* Section header */
.sh{ font-size:11px; font-weight:600; letter-spacing:1px; text-transform:uppercase;
     color:rgba(255,255,255,.3); margin:1.5rem 0 .6rem; }
 
/* Status items */
.sok { color:#34d399; font-size:13px; }
.serr{ color:#f87171; font-size:13px; }
.swrn{ color:#fbbf24; font-size:13px; }
 
/* Log box */
.logbox{ background:rgba(0,0,0,.4); border:1px solid rgba(255,255,255,.08);
         border-radius:8px; padding:14px 16px; font-family:monospace;
         font-size:12px; color:rgba(255,255,255,.7); max-height:360px;
         overflow-y:auto; line-height:1.8; }
.log-ok  { color:#34d399; }
.log-err { color:#f87171; }
.log-inf { color:#93c5fd; }
.log-wrn { color:#fbbf24; }
 
/* Metrics */
div[data-testid="stMetric"]{ background:rgba(255,255,255,.04);
  border:1px solid rgba(255,255,255,.08); border-radius:10px; padding:14px 18px; }
div[data-testid="stMetricValue"]{ font-size:1.7rem !important; font-weight:700; }
 
/* Download card */
.dl-card{ background:rgba(255,255,255,.04); border:1px solid rgba(255,255,255,.08);
          border-radius:10px; padding:14px 20px; display:flex;
          align-items:center; gap:14px; margin-bottom:8px; }
.dl-icon{ font-size:22px; }
.dl-info{ flex:1; }
.dl-name{ font-size:13px; font-weight:600; }
.dl-desc{ font-size:11px; color:rgba(255,255,255,.4); }
</style>
""", unsafe_allow_html=True)
 
# ── Constants ─────────────────────────────────────────────────────────────────
SYNONYMS = {
    "question_set":      ["nome do bloco","bloco","question_set","block","grupo","seção","secao","categoria"],
    "desc_question_set": ["descrição do bloco","desc_question_set","desc bloco"],
    "competencia":       ["competência","competencia","nome da competência","skill","tema"],
    "pergunta":          ["pergunta","question","enunciado","texto","item"],
    "opcional":          ["essa pergunta é opcional","opcional","optional"],
    "aberta":            ["é uma pergunta aberta","aberta","open","tipo","type"],
    "escala":            ["qual o número do modelo da escala","escala","scale","número da escala","modelo da escala"],
    "min_caracters":     ["número mínimo de caracteres","min caracters","min_caracters"],
    "max_caracters":     ["número máximo de caracteres","max caracters","max_caracters"],
    "definicao":         ["descrição geral sobre a pergunta","definição","definicao","definition"],
}
DESC_SCALE_SYNS = ["descrição do primeiro","descrição do segundo","descrição do terceiro",
                   "descrição do quarto","descrição do quinto","duplique"]
GRUPOS_OPTS = ["Autoavaliação","Time","Gestor","Par","Stakeholder"]
DICT_CORES = {
    2:['#A8002E','#006499'], 3:['#A8002E','#FFC252','#006499'],
    4:['#A8002E','#FFC252','#35AEFF','#033D17'], 5:['#A8002E','#FFC252','#35AEFF','#006499','#033D17'],
    6:['#A8002E','#FA6F26','#FFC252','#35AEFF','#006499','#033D17'],
}
 
# ── Helpers ───────────────────────────────────────────────────────────────────
def uid(): return str(uuid.uuid4())[:8]
 
def norm_bool(val):
    v = str(val).strip().lower()
    return "sim" if v in ("sim","yes","true","1","s","y") else "não"
 
def detect_col(hl, field):
    for i,h in enumerate(hl):
        if any(s in h or h in s for s in SYNONYMS.get(field,[])):
            return i
    return None
 
def sim_r(a,b):
    return SequenceMatcher(None,a.lower().strip(),b.lower().strip()).ratio()
 
def make_p(**kw):
    return {"id":uid(),"texto":kw.get("texto",""),"competencia":kw.get("competencia",""),
            "aberta":kw.get("aberta","não"),"escala":str(kw.get("escala","")),"opcional":kw.get("opcional","não"),
            "min_caracters":str(kw.get("min_c","")),"max_caracters":str(kw.get("max_c","")),"definicao":kw.get("definicao",""),
            "grupos":kw.get("grupos",list(GRUPOS_OPTS))}
 
def make_b(nome="",desc="",perguntas=None,origem="manual"):
    return {"id":uid(),"nome":nome,"desc":desc,"perguntas":perguntas or [],"origem":origem}
 
def make_e(num=None,pontos=None):
    if num is None:
        num = max((e["num"] for e in ss.escalas),default=0)+1
    return {"id":uid(),"num":int(num) if str(num).replace(".","").isdigit() else num,"pontos":pontos or ["",""]}
 
# ── Session state ─────────────────────────────────────────────────────────────
DEFAULTS = {
    "step":1, "escalas":[], "blocos":[], "bloco_sel":None,
    "cfg":{"cliente":"","token":"","id_rodada":"","id_qs":"","id_cq":"","id_oq":"",
           "val_min_nan":"0","autonomia_bp":True,"sso":True},
    "resultados":None, "log":[]
}
for k,v in DEFAULTS.items():
    if k not in st.session_state: st.session_state[k]=v
ss = st.session_state
 
# ── Steps nav ─────────────────────────────────────────────────────────────────
STEPS = ["1. Entrada","2. Validação","3. Configuração","4. Geração"]
 
def steps_nav():
    pills=""
    for i,label in enumerate(STEPS,1):
        cls = "active" if i==ss.step else ("done" if i<ss.step else "")
        num = "✓" if i<ss.step else str(i)
        click = f'onclick="void(0)"' if cls!="done" else ""
        pills+=f'<div class="sn {cls}"><span class="sn-num">{num}</span>{label}</div>'
    st.markdown(f'<div class="steps-nav">{pills}</div>',unsafe_allow_html=True)
 
def go(n): ss.step=n; st.rerun()
 
def nav_btns(back=None,fwd=None,fwd_label="Continuar →",fwd_disabled=False):
    c1,_,c2=st.columns([1,5,1])
    if back and c1.button("← Voltar",use_container_width=True): go(back)
    if fwd  and c2.button(fwd_label,type="primary",use_container_width=True,disabled=fwd_disabled): go(fwd)
 
# ── Import xlsx ───────────────────────────────────────────────────────────────
def import_xlsx(file_bytes):
    wb=load_workbook(io.BytesIO(file_bytes),read_only=True,data_only=True)
    msgs=[]
    # Escalas
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
    escalas_out=[]
    for ns,descs in emap.items():
        try: n=int(float(ns))
        except: n=ns
        escalas_out.append(make_e(num=n,pontos=descs if descs else ["",""]))
    # Blocos
    bmap={}
    for sn in wb.sheetnames:
        ws=wb[sn]; data=list(ws.iter_rows(values_only=True))
        if len(data)<2: continue
        hl=[str(c or "").lower().strip() for c in data[0]]
        idx={f:detect_col(hl,f) for f in SYNONYMS}
        def get(row,f):
            i=idx.get(f)
            return str(row[i] or "").strip() if (i is not None and i<len(row)) else ""
        for row in data[1:]:
            if not any(c not in (None,"") for c in row): continue
            bn=get(row,"question_set"); pt=get(row,"pergunta")
            if not bn and not pt: continue
            p=make_p(texto=pt,competencia=get(row,"competencia"),
                     aberta=norm_bool(get(row,"aberta") or "não"),
                     escala=get(row,"escala"),opcional=norm_bool(get(row,"opcional") or "não"),
                     min_c=get(row,"min_caracters"),max_c=get(row,"max_caracters"),
                     definicao=get(row,"definicao"))
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
            if n1!=n2 and sim_r(n1,n2)>0.85: avisos.append(f'Blocos parecidos: "{n1}" e "{n2}"')
    total=0
    for b in ss.blocos:
        if not b["nome"].strip(): erros.append("Bloco sem nome.")
        if not b["perguntas"]: avisos.append(f'Bloco "{b["nome"]}" sem perguntas.')
        for p in b["perguntas"]:
            total+=1
            if not p["texto"].strip(): erros.append(f'Pergunta vazia em "{b["nome"]}".')
            if p["aberta"]=="não":
                if not p["escala"]: avisos.append(f'Fechada sem escala em "{b["nome"]}".')
                elif p["escala"] not in enum: erros.append(f'Escala {p["escala"]} não existe em "{b["nome"]}".')
            if p["aberta"]=="sim" and not p["max_caracters"]: avisos.append(f'Aberta sem máx chars em "{b["nome"]}".')
    return erros,avisos,total
 
# ── Script logic (portado do notebook) ───────────────────────────────────────
def run_script(cfg):
    log=[]
    def L(msg,tipo="inf"): log.append((tipo,msg))
 
    try:
        id_rodada      = int(cfg["id_rodada"])
        id_qs          = int(cfg["id_qs"])
        id_cq          = int(cfg["id_cq"])
        id_oq          = int(cfg["id_oq"])
        val_min_nan    = float(cfg["val_min_nan"] or "0")
    except Exception as e:
        log.append(("err",f"IDs inválidos: {e}")); return log,{}
 
    # Montar base_perguntas e base_escalas como DataFrames
    rows_p=[]
    for b in ss.blocos:
        for p in b["perguntas"]:
            row={"question_set":b["nome"],"desc_question_set":b["desc"],
                 "competencia":p["competencia"],"pergunta":p["texto"],
                 "opcional":p["opcional"],"aberta":p["aberta"],
                 "escala":int(float(p["escala"])) if p["aberta"]=="não" and p["escala"] else np.nan,
                 "min caracters":int(p["min_caracters"]) if p["min_caracters"] else np.nan,
                 "max caracters":int(p["max_caracters"]) if p["max_caracters"] else np.nan,
                 "definição":p["definicao"] if p["definicao"] else np.nan}
            rows_p.append(row)
    base_perguntas=pd.DataFrame(rows_p)
 
    # base_escalas
    rows_e=[]
    for e in ss.escalas:
        r={"escala":e["num"],
           "min":float(e["pontos"][0]) if e["pontos"] and e["pontos"][0] else np.nan,
           "max":float(e["pontos"][-1]) if e["pontos"] and e["pontos"][-1] else np.nan}
        for i,pt in enumerate(e["pontos"],1): r[f"descrição_{i}"]=pt if pt else np.nan
        rows_e.append(r)
    base_escalas=pd.DataFrame(rows_e)
 
    L(f"Base de perguntas: {len(base_perguntas)} linhas","inf")
    L(f"Base de escalas: {len(base_escalas)} linhas","inf")
 
    # ── Question Sets ──
    L("Gerando Question Sets...","inf")
    base_question_sets=pd.DataFrame(columns=['name','description','id','order','evaluation_rounds'])
    first=True; contador=0; criados=[]
    for _,row in base_perguntas.iterrows():
        if row['question_set'] not in criados:
            desc="" if pd.isna(row.get('desc_question_set','')) else row.get('desc_question_set','')
            reg={'id':id_qs if first else id_qs+contador,
                 'order':contador+1,'name':row['question_set'],
                 'description':desc,'evaluation_rounds':id_rodada}
            if first: first=False
            base_question_sets.loc[len(base_question_sets)]=reg
            criados.append(row['question_set']); contador+=1
    L(f"✓ {len(base_question_sets)} question set(s) gerado(s)","ok")
 
    # ── Open Questions ──
    L("Gerando Open Questions...","inf")
    base_open_questions=pd.DataFrame(columns=['info','description','id','question_set',
                                               'is_optional','order','key','min_answer_length',
                                               'max_answer_length','evaluation_rounds'])
    first=True; cnt_oq=0; cnt_p=1
    for _,row in base_perguntas.iterrows():
        if str(row['aberta']).lower()=='não': cnt_p+=1; continue
        pergunta=("" if pd.isna(row.get('competencia')) else str(row['competencia']).upper()+" | ")+row['pergunta']
        opcional=0 if str(row['opcional']).lower()=='não' else 1
        min_c=0 if pd.isna(row.get('min caracters')) else int(row['min caracters'])
        max_c=int(row['max caracters'])
        qs=base_question_sets[base_question_sets['name']==row['question_set']]
        descricao="" if pd.isna(row.get('definição')) else str(row['definição'])
        reg={'id':id_oq if first else id_oq+cnt_oq,'order':cnt_p,'key':id_oq if first else id_oq+cnt_oq,
             'description':pergunta,'info':descricao,'is_optional':opcional,
             'min_answer_length':min_c,'max_answer_length':max_c,
             'question_set':qs['id'].iloc[0],'evaluation_rounds':id_rodada}
        if first: first=False
        base_open_questions.loc[len(base_open_questions)]=reg
        cnt_oq+=1; cnt_p+=1
    L(f"✓ {len(base_open_questions)} open question(s) gerada(s)","ok")
 
    # ── Choice Questions ──
    L("Gerando Choice Questions...","inf")
    base_choice_questions=pd.DataFrame(columns=['info','description','id','question_set',
                                                  'is_optional','order','key','evaluation_rounds','escala'])
    first=True; cnt_cq=0; cnt_p=1
    for _,row in base_perguntas.iterrows():
        if str(row['aberta']).lower()=='sim': cnt_p+=1; continue
        pergunta=("" if pd.isna(row.get('competencia')) else str(row['competencia']).upper()+" | ")+row['pergunta']
        opcional=0 if str(row['opcional']).lower()=='não' else 1
        qs=base_question_sets[base_question_sets['name']==row['question_set']]
        desc_escalas=[]
        for col in base_perguntas.columns:
            if 'descrição_' in col and not pd.isna(row.get(col)):
                desc_escalas.append(row[col])
        descricao=""
        if not pd.isna(row.get('definição')) or desc_escalas:
            # simplified info_construct
            d=""
            if not pd.isna(row.get('definição')): d=f"<p><strong>{row['definição']}</strong></p>\n\n"
            descricao=d
        reg={'id':id_cq if first else id_cq+cnt_cq,'order':cnt_p,'key':id_cq if first else id_cq+cnt_cq,
             'description':pergunta,'info':descricao,'is_optional':opcional,
             'question_set':qs['id'].iloc[0],'escala':int(row['escala']),'evaluation_rounds':id_rodada}
        if first: first=False
        base_choice_questions.loc[len(base_choice_questions)]=reg
        cnt_cq+=1; cnt_p+=1
    L(f"✓ {len(base_choice_questions)} choice question(s) gerada(s)","ok")
 
    # ── Escalas ──
    L("Gerando Escalas...","inf")
    base_escalas_import=pd.DataFrame(columns=['content','id','question','value','color'])
    for _,row in base_choice_questions.iterrows():
        base_temp=base_escalas[base_escalas['escala']==row['escala']]
        if len(base_temp)==0:
            L(f"✕ Escala {row['escala']} não encontrada na base de escalas","err"); continue
        for _,row_e in base_temp.iterrows():
            colunas=[c for c in base_temp.columns if c not in ['escala','min','max']]
            minimo=row_e.get('min'); maximo=float(row_e['max'])
            flag_nan=False
            pts_validos=[c for c in colunas if not pd.isna(row_e.get(c))]
            contador=len(pts_validos)
            if pd.isna(minimo):
                flag_nan=True; minimo=val_min_nan
                divisao=(maximo-float(minimo))/(contador-2) if contador>2 else 1
            else:
                minimo=float(minimo)
                divisao=(maximo-minimo)/(contador-1) if contador>1 else 1
            cores=DICT_CORES.get(contador,['#000000']*contador)
            cnt_i=0
            for col in colunas:
                if pd.isna(row_e.get(col)): continue
                if flag_nan:
                    reg={'id':'','question':row['id'],'value':np.nan,'content':str(row_e[col]),'color':'#000000'}
                    flag_nan=False
                else:
                    cor=cores[cnt_i] if cnt_i<len(cores) else '#000000'
                    reg={'id':'','question':row['id'],'value':round(float(minimo)+(cnt_i*divisao),2),
                         'content':str(row_e[col]),'color':cor}
                    cnt_i+=1
                base_escalas_import.loc[len(base_escalas_import)]=reg
    L(f"✓ {len(base_escalas_import)} item(ns) de escala gerado(s)","ok")
 
    # ── Serializar JSONs ──
    L("Serializando JSONs...","inf")
    cliente=cfg["cliente"]
 
    def to_json(df):
        return df.to_json(orient='records',force_ascii=False)
 
    cq_export=base_choice_questions.drop(columns=['escala'],errors='ignore').copy()
 
    resultados={
        f"Question Sets {cliente}.json":    to_json(base_question_sets),
        f"Open Questions {cliente}.json":   to_json(base_open_questions),
        f"Choice Questions {cliente}.json": to_json(cq_export),
        f"Escalas {cliente}.json":          to_json(base_escalas_import),
    }
    L("✓ JSONs prontos para download","ok")
    L(f"Resumo: {len(base_question_sets)} QS · {len(base_open_questions)} OQ · {len(base_choice_questions)} CQ · {len(base_escalas_import)} itens de escala","ok")
 
    return log, resultados
 
# ── Build zip ─────────────────────────────────────────────────────────────────
def build_zip(resultados):
    buf=io.BytesIO()
    with zipfile.ZipFile(buf,'w',zipfile.ZIP_DEFLATED) as z:
        for fname,content in resultados.items():
            z.writestr(fname,content.encode('utf-8'))
    return buf.getvalue()
 
 
# ═══════════════════════════════════════════════════════════════════════════════
# STEP 1 — Entrada
# ═══════════════════════════════════════════════════════════════════════════════
if ss.step==1:
    steps_nav()
    st.markdown("# Configuração de Avaliação")
    st.caption("Importe a planilha do cliente ou preencha as informações manualmente.")
 
    modo=st.radio("Como deseja começar?",
                  ["📥 Importar planilha existente","✏️ Preencher manualmente"],
                  horizontal=True, label_visibility="collapsed")
 
    st.divider()
 
    # ── MODO IMPORT ──────────────────────────────────────────────────────────
    if "Importar" in modo:
        up=st.file_uploader("Arraste o arquivo .xlsx do cliente",type=["xlsx","xls"],label_visibility="collapsed")
        if up:
            with st.spinner("Lendo arquivo..."):
                bi,ei,mi=import_xlsx(up.read())
            c1,c2,c3=st.columns(3)
            c1.metric("Blocos",len(bi))
            c2.metric("Perguntas",sum(len(b["perguntas"]) for b in bi))
            c3.metric("Escalas",len(ei))
            if mi:
                with st.expander("📋 Ajustes automáticos"):
                    for m in mi: st.markdown(m)
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
 
    # ── MODO MANUAL ──────────────────────────────────────────────────────────
    else:
        escala_opts=[""]+[str(e["num"]) for e in ss.escalas]
 
        # Seção Escalas
        st.markdown('<div class="sh">Escalas de avaliação</div>',unsafe_allow_html=True)
        for ei_idx,esc in enumerate(ss.escalas):
            with st.container(border=True):
                hc1,hc2=st.columns([7,1])
                with hc1:
                    nv=st.number_input("Escala nº",value=int(esc["num"]) if str(esc["num"]).replace(".","").isdigit() else ei_idx+1,
                                       min_value=1,key=f"en_{esc['id']}",label_visibility="collapsed")
                    esc["num"]=nv
                    if any(str(p).strip() for p in esc["pontos"]):
                        prev="".join(f'<div class="esc-pt"><span class="esc-n">{i+1}</span>{p or "—"}</div>'
                                     for i,p in enumerate(esc["pontos"]))
                        st.markdown(f'<div class="esc-row">{prev}</div>',unsafe_allow_html=True)
                    with st.expander(f"✏ Editar — Escala {esc['num']}"):
                        n_pts=st.number_input("Pontos",2,10,len(esc["pontos"]),key=f"np_{esc['id']}")
                        while len(esc["pontos"])<n_pts: esc["pontos"].append("")
                        esc["pontos"]=esc["pontos"][:n_pts]
                        pc=st.columns(min(n_pts,5))
                        for pi in range(n_pts):
                            esc["pontos"][pi]=pc[pi%len(pc)].text_input(f"Ponto {pi+1}",value=esc["pontos"][pi],key=f"pt_{esc['id']}_{pi}")
                with hc2:
                    st.markdown("<br><br>",unsafe_allow_html=True)
                    if st.button("✕",key=f"de_{esc['id']}"): ss.escalas.pop(ei_idx); st.rerun()
 
        if st.button("＋ Escala",key="ae"): ss.escalas.append(make_e()); st.rerun()
        if not ss.escalas: st.info("Adicione pelo menos uma escala para perguntas fechadas.")
 
        st.divider()
 
        # Seção Blocos & Perguntas
        st.markdown('<div class="sh">Blocos & Perguntas</div>',unsafe_allow_html=True)
        col_bl,col_pr=st.columns([2,3],gap="large")
 
        with col_bl:
            st.markdown(f"**{len(ss.blocos)} bloco(s)**")
            for bi_idx,bloco in enumerate(ss.blocos):
                np_=len(bloco["perguntas"])
                ne_=sum(1 for p in bloco["perguntas"] if not p["texto"] or (p["aberta"]=="não" and not p["escala"]))
                is_s=ss.bloco_sel==bloco["id"]
                badges=f'<span class="badge b-blue">{np_}p</span>'+(f'<span class="badge b-red">{ne_}✕</span>' if ne_ else '<span class="badge b-green">OK</span>')
                st.markdown(f'<div class="card {"sel" if is_s else ""}"><div class="card-title">{bloco["nome"] or "(sem nome)"}</div><div style="margin-top:5px">{badges}</div></div>',unsafe_allow_html=True)
                b1,b2=st.columns([4,1])
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
                st.markdown(f"**{len(bloco['perguntas'])} pergunta(s)**")
 
                for pi,p in enumerate(bloco["perguntas"]):
                    with st.container(border=True):
                        r1,r2=st.columns([9,1])
                        p["texto"]=r1.text_area(f"#{pi+1}",value=p["texto"],key=f"ptx_{p['id']}",
                                                 label_visibility="collapsed",placeholder="Texto da pergunta...",height=70)
                        r2.markdown("<br>",unsafe_allow_html=True)
                        if r2.button("✕",key=f"dp_{p['id']}"): bloco["perguntas"].pop(pi); st.rerun()
 
                        fa,fb,fc=st.columns([2,1,2])
                        p["competencia"]=fa.text_input("Competência",value=p["competencia"],key=f"pc_{p['id']}",placeholder="Ex: Liderança")
                        tipo=fb.selectbox("Tipo",["Fechada","Aberta"],index=0 if p["aberta"]=="não" else 1,key=f"pty_{p['id']}")
                        p["aberta"]="não" if tipo=="Fechada" else "sim"
                        if p["aberta"]=="não":
                            p["escala"]=fc.selectbox("Escala *",escala_opts,
                                                      index=escala_opts.index(p["escala"]) if p["escala"] in escala_opts else 0,
                                                      key=f"pesc_{p['id']}")
                        else:
                            m1,m2=fc.columns(2)
                            p["min_caracters"]=m1.text_input("Mín chars",value=p["min_caracters"],key=f"pmi_{p['id']}",placeholder="0")
                            p["max_caracters"]=m2.text_input("Máx chars",value=p["max_caracters"],key=f"pma_{p['id']}",placeholder="10000")
 
                        # Grupos avaliativos
                        p["grupos"]=st.multiselect("Grupos avaliativos",GRUPOS_OPTS,
                                                    default=p.get("grupos",GRUPOS_OPTS),key=f"pgr_{p['id']}")
 
                        with st.expander("Campos opcionais"):
                            oa,ob=st.columns(2)
                            p["opcional"]="sim" if oa.checkbox("Opcional?",value=p["opcional"]=="sim",key=f"popc_{p['id']}") else "não"
                            p["definicao"]=ob.text_input("Definição",value=p["definicao"],key=f"pdef_{p['id']}")
 
                st.markdown("")
                if st.button("＋ Adicionar pergunta",key=f"ap_{bloco['id']}",use_container_width=True):
                    bloco["perguntas"].append(make_p()); st.rerun()
 
    st.divider()
    erros,_,tp=validar()
    _,c2=st.columns([5,1])
    with c2:
        if st.button("Ir para Validação →",type="primary",use_container_width=True,disabled=bool(erros) or not ss.blocos):
            go(2)
    if erros and ss.blocos:
        st.error(f"{len(erros)} erro(s) encontrado(s) — corrija antes de continuar.")
 
 
# ═══════════════════════════════════════════════════════════════════════════════
# STEP 2 — Visualização e Validação
# ═══════════════════════════════════════════════════════════════════════════════
elif ss.step==2:
    steps_nav()
    st.markdown("# Visualização e Validação")
    erros,avisos,total_p=validar()
 
    m1,m2,m3,m4=st.columns(4)
    m1.metric("Blocos",len(ss.blocos))
    m2.metric("Perguntas",total_p)
    m3.metric("Escalas",len(ss.escalas))
    m4.metric("Erros",len(erros),delta=str(len(erros)) if erros else None,delta_color="inverse" if erros else "off")
 
    st.divider()
 
    # Checklist
    col_e,col_w=st.columns(2)
    with col_e:
        st.markdown("**Erros (obrigatório corrigir)**" if erros else "**✅ Sem erros**")
        for e in erros: st.markdown(f'<div class="serr">✕ {e}</div>',unsafe_allow_html=True)
    with col_w:
        st.markdown("**Avisos (recomendado verificar)**" if avisos else "**✅ Sem avisos**")
        for a in avisos: st.markdown(f'<div class="swrn">⚠ {a}</div>',unsafe_allow_html=True)
 
    st.divider()
 
    # Preview por bloco
    st.markdown('<div class="sh">Prévia dos blocos e perguntas</div>',unsafe_allow_html=True)
    for b in ss.blocos:
        np_=len(b["perguntas"]); nab=sum(1 for p in b["perguntas"] if p["aberta"]=="sim")
        eu=sorted(set(p["escala"] for p in b["perguntas"] if p["escala"]))
        ne_=sum(1 for p in b["perguntas"] if not p["texto"] or (p["aberta"]=="não" and not p["escala"]))
        badges=(f'<span class="badge b-blue">{np_}p</span>'
                +f'<span class="badge b-gray">{np_-nab} fechada(s)</span>'
                +f'<span class="badge b-gray">{nab} aberta(s)</span>'
                +(f'<span class="badge b-red">{ne_} erro(s)</span>' if ne_ else '<span class="badge b-green">OK</span>'))
        with st.expander(f"{b['nome']}  —  {badges}", expanded=ne_>0):
            st.markdown(f"**Descrição:** {b['desc'] or '—'}")
            if eu: st.markdown(f"**Escalas usadas:** {', '.join(eu)}")
            df_prev=pd.DataFrame([{
                "#":i+1,"Pergunta":p["texto"][:80]+("…" if len(p["texto"])>80 else ""),
                "Competência":p["competencia"],"Tipo":"Aberta" if p["aberta"]=="sim" else "Fechada",
                "Escala":p["escala"] if p["aberta"]=="não" else "—",
                "Opcional":"Sim" if p["opcional"]=="sim" else "Não",
                "Grupos":", ".join(p.get("grupos",[])),
            } for i,p in enumerate(b["perguntas"])])
            st.dataframe(df_prev,use_container_width=True,hide_index=True)
 
    # Preview escalas
    st.divider()
    st.markdown('<div class="sh">Escalas</div>',unsafe_allow_html=True)
    for e in ss.escalas:
        with st.expander(f"Escala {e['num']} — {len(e['pontos'])} pontos"):
            prev="".join(f'<div class="esc-pt"><span class="esc-n">{i+1}</span>{p or "—"}</div>'
                          for i,p in enumerate(e["pontos"]))
            st.markdown(f'<div class="esc-row">{prev}</div>',unsafe_allow_html=True)
 
    nav_btns(back=1,fwd=3,fwd_label="Configurar rodada →",fwd_disabled=bool(erros))
 
 
# ═══════════════════════════════════════════════════════════════════════════════
# STEP 3 — Configuração da Rodada
# ═══════════════════════════════════════════════════════════════════════════════
elif ss.step==3:
    steps_nav()
    st.markdown("# Configuração da Rodada")
    st.caption("Preencha as informações da rodada e os IDs gerados na plataforma.")
 
    cfg=ss.cfg
 
    st.markdown('<div class="sh">Acesso à plataforma</div>',unsafe_allow_html=True)
    c1,c2=st.columns(2)
    cfg["cliente"]=c1.text_input("Tenant do cliente",value=cfg["cliente"],placeholder="ex: mindsight")
    cfg["token"]=c2.text_input("Token do Performance",value=cfg["token"],type="password",placeholder="Token de acesso")
 
    st.markdown('<div class="sh">IDs da rodada</div>',unsafe_allow_html=True)
    st.caption("Crie a rodada e os grupos avaliativos na plataforma antes de preencher.")
    r1,r2=st.columns(2)
    cfg["id_rodada"]=r1.text_input("ID da Rodada *",value=cfg["id_rodada"])
    cfg["id_qs"]    =r2.text_input("ID do 1º Question Set *",value=cfg["id_qs"],help="ID genérico criado na plataforma")
 
    r3,r4=st.columns(2)
    cfg["id_cq"]=r3.text_input("ID da 1ª Choice Question *",value=cfg["id_cq"])
    cfg["id_oq"]=r4.text_input("ID da 1ª Open Question *",value=cfg["id_oq"])
 
    st.markdown('<div class="sh">Configurações adicionais</div>',unsafe_allow_html=True)
    a1,a2,a3=st.columns(3)
    cfg["val_min_nan"]=a1.text_input("Valor mínimo p/ escalas com NaN",value=cfg["val_min_nan"])
    cfg["autonomia_bp"]=a2.checkbox("Autonomia de BPs ativa?",value=cfg["autonomia_bp"])
    cfg["sso"]=a3.checkbox("SSO?",value=cfg["sso"])
 
    # Validação dos IDs
    ids_ok=all(cfg[k].strip() for k in ["id_rodada","id_qs","id_cq","id_oq"])
    tenant_ok=bool(cfg["cliente"].strip())
    if not tenant_ok: st.warning("Preencha o tenant do cliente.")
    if not ids_ok:    st.warning("Preencha todos os IDs obrigatórios.")
 
    nav_btns(back=2,fwd=4,fwd_label="Gerar bases →",fwd_disabled=not(ids_ok and tenant_ok))
 
 
# ═══════════════════════════════════════════════════════════════════════════════
# STEP 4 — Geração e Export
# ═══════════════════════════════════════════════════════════════════════════════
elif ss.step==4:
    steps_nav()
    st.markdown("# Geração e Export")
 
    if ss.resultados is None:
        if st.button("▶ Gerar arquivos agora",type="primary",use_container_width=True):
            with st.spinner("Gerando bases..."):
                log,resultados=run_script(ss.cfg)
            ss.log=log; ss.resultados=resultados; st.rerun()
    else:
        # Log de execução
        st.markdown('<div class="sh">Log de execução</div>',unsafe_allow_html=True)
        log_html="".join(
            f'<div class="log-{t}">'
            +("✓ " if t=="ok" else "✕ " if t=="err" else "⚠ " if t=="wrn" else "→ ")
            +msg+"</div>"
            for t,msg in ss.log)
        st.markdown(f'<div class="logbox">{log_html}</div>',unsafe_allow_html=True)
 
        erros_exec=[m for t,m in ss.log if t=="err"]
        if erros_exec:
            st.error(f"{len(erros_exec)} erro(s) na geração. Volte e corrija os dados.")
        else:
            st.success("✅ Todos os arquivos foram gerados com sucesso!")
            st.divider()
            st.markdown('<div class="sh">Downloads</div>',unsafe_allow_html=True)
 
            # Download individual
            for fname,content in ss.resultados.items():
                dc1,dc2=st.columns([5,1])
                with dc1:
                    rows=len(json.loads(content))
                    st.markdown(f'<div class="dl-card"><div class="dl-icon">📄</div><div class="dl-info"><div class="dl-name">{fname}</div><div class="dl-desc">{rows} registro(s)</div></div></div>',unsafe_allow_html=True)
                with dc2:
                    st.markdown("<br>",unsafe_allow_html=True)
                    st.download_button("⬇",data=content.encode("utf-8"),file_name=fname,
                                       mime="application/json",key=f"dl_{fname}",use_container_width=True)
 
            st.divider()
            # Download zip
            st.download_button(
                "⬇ Baixar todos (.zip)",
                data=build_zip(ss.resultados),
                file_name=f"AVD_{ss.cfg['cliente']}.zip",
                mime="application/zip",
                type="primary",
                use_container_width=True,
            )
 
        st.divider()
        col_b,col_r=st.columns([1,5])
        with col_b:
            if st.button("← Voltar"): go(3)
        with col_r:
            if st.button("↺ Regerar",key="regen"):
                ss.resultados=None; ss.log=[]; st.rerun()
