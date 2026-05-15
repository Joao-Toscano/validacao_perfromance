"""
AVD — Configuração de Perguntas
App para clientes configurarem blocos de perguntas e escalas,
gerando o Base_Perguntas-2.xlsx pronto para importação.
 
Rodar: streamlit run app.py
"""
 
import io
import uuid
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from difflib import SequenceMatcher
 
st.set_page_config(
    page_title="Configuração de Avaliação",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="collapsed",
)
 
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
.block-container { padding: 2.5rem 3rem 3rem !important; max-width: 1100px; }
.steps-bar { display:flex; gap:0; margin-bottom:2.5rem; border-radius:10px; overflow:hidden; border:1px solid rgba(255,255,255,0.08); }
.step-pill { flex:1; padding:12px 8px; text-align:center; font-size:12px; font-weight:500; background:rgba(255,255,255,0.03); color:rgba(255,255,255,0.35); border-right:1px solid rgba(255,255,255,0.08); }
.step-pill:last-child { border-right:none; }
.step-pill.active { background:rgba(99,102,241,0.15); color:#818cf8; font-weight:600; }
.step-pill.done { background:rgba(16,185,129,0.08); color:#34d399; }
.step-num { display:inline-block; width:20px; height:20px; border-radius:50%; border:1.5px solid currentColor; font-size:10px; line-height:18px; margin-right:6px; text-align:center; }
.step-pill.done .step-num { background:#34d399; color:#064e3b; border-color:#34d399; }
.escala-preview { display:flex; gap:4px; margin-top:10px; margin-bottom:4px; }
.escala-ponto { flex:1; background:rgba(255,255,255,0.06); border-radius:6px; padding:8px 4px; text-align:center; font-size:10px; color:rgba(255,255,255,0.5); border:1px solid rgba(255,255,255,0.08); line-height:1.3; }
.escala-num { font-size:16px; font-weight:700; color:white; display:block; }
.bloco-card { background:rgba(255,255,255,0.04); border:1px solid rgba(255,255,255,0.08); border-radius:10px; padding:14px 18px; margin-bottom:4px; }
.bloco-card.sel { border-color:#818cf8; background:rgba(99,102,241,0.1); }
.bloco-title { font-size:14px; font-weight:600; margin-bottom:4px; }
.bloco-meta { font-size:11px; color:rgba(255,255,255,0.4); }
.badge { font-size:10px; padding:2px 8px; border-radius:99px; font-weight:500; border:1px solid; margin-right:4px; }
.badge-blue  { color:#93c5fd; border-color:rgba(147,197,253,0.3); background:rgba(147,197,253,0.08); }
.badge-green { color:#86efac; border-color:rgba(134,239,172,0.3); background:rgba(134,239,172,0.08); }
.badge-red   { color:#fca5a5; border-color:rgba(252,165,165,0.3); background:rgba(252,165,165,0.08); }
.badge-gray  { color:rgba(255,255,255,0.4); border-color:rgba(255,255,255,0.1); background:rgba(255,255,255,0.04); }
.sec-title { font-size:11px; font-weight:600; letter-spacing:1px; text-transform:uppercase; color:rgba(255,255,255,0.3); margin-bottom:12px; margin-top:4px; }
.val-ok   { color:#34d399; font-size:13px; padding:4px 0; }
.val-err  { color:#f87171; font-size:13px; padding:4px 0; }
.val-warn { color:#fbbf24; font-size:13px; padding:4px 0; }
div[data-testid="stMetric"] { background:rgba(255,255,255,0.04); border:1px solid rgba(255,255,255,0.08); border-radius:10px; padding:14px 18px; }
div[data-testid="stMetricValue"] { font-size:1.8rem !important; font-weight:700; }
</style>
""", unsafe_allow_html=True)
 
# ── Constantes ────────────────────────────────────────────────────────────────
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
OUTPUT_COLS = ["question_set","desc_question_set","competencia","pergunta",
               "opcional","aberta","escala","min caracters","max caracters","definição"]
 
# ── Helpers ───────────────────────────────────────────────────────────────────
def uid(): return str(uuid.uuid4())[:8]
 
def norm_bool(val):
    v = str(val).strip().lower()
    return "sim" if v in ("sim","yes","true","1","s","y") else "não"
 
def detect_col(headers_l, field):
    syns = SYNONYMS.get(field, [])
    for i, h in enumerate(headers_l):
        if any(s in h or h in s for s in syns):
            return i
    return None
 
def sim(a, b):
    return SequenceMatcher(None, a.lower().strip(), b.lower().strip()).ratio()
 
def make_p(texto="", competencia="", aberta="não", escala="", opcional="não", min_c="", max_c="", definicao=""):
    return {"id":uid(),"texto":texto,"competencia":competencia,"aberta":aberta,
            "escala":str(escala),"opcional":opcional,"min_caracters":str(min_c),
            "max_caracters":str(max_c),"definicao":definicao}
 
def make_b(nome="", desc="", perguntas=None, origem="manual"):
    return {"id":uid(),"nome":nome,"desc":desc,"perguntas":perguntas or [],"origem":origem}
 
def make_e(num=None, pontos=None):
    if num is None:
        num = max((e["num"] for e in ss.escalas), default=0) + 1
    return {"id":uid(),"num":num,"pontos":pontos or ["",""]}
 
# ── Session state ─────────────────────────────────────────────────────────────
def init():
    defs = {"step":1,"escalas":[],"blocos":[],"bloco_sel":None,"imported":False}
    for k,v in defs.items():
        if k not in st.session_state:
            st.session_state[k] = v
init()
ss = st.session_state
 
# ── Steps bar ─────────────────────────────────────────────────────────────────
def steps_bar():
    labels = ["Escalas","Blocos & Perguntas","Revisão","Exportar"]
    pills = ""
    for i,label in enumerate(labels,1):
        cls = "active" if i==ss.step else ("done" if i<ss.step else "")
        num = "✓" if i<ss.step else str(i)
        pills += f'<div class="step-pill {cls}"><span class="step-num">{num}</span>{label}</div>'
    st.markdown(f'<div class="steps-bar">{pills}</div>', unsafe_allow_html=True)
 
def nav(back=None, fwd=None, fwd_label="Continuar →", fwd_disabled=False):
    ca, _, cb = st.columns([1,4,1])
    with ca:
        if back and st.button("← Voltar", use_container_width=True):
            ss.step=back; st.rerun()
    with cb:
        if fwd and st.button(fwd_label, type="primary", use_container_width=True, disabled=fwd_disabled):
            ss.step=fwd; st.rerun()
 
# ── Import ────────────────────────────────────────────────────────────────────
def import_xlsx(file_bytes):
    wb = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    msgs = []
 
    # Escalas
    escalas_map = {}
    for sn in wb.sheetnames:
        ws = wb[sn]
        data = list(ws.iter_rows(values_only=True))
        if not data: continue
        hl = [str(c or "").lower().strip() for c in data[0]]
        si = detect_col(hl, "escala")
        di = [i for i,h in enumerate(hl) if any(s in h for s in DESC_SCALE_SYNS)]
        for row in data[1:]:
            if si is None or si >= len(row): continue
            num = str(row[si] or "").strip()
            if not num: continue
            descs = [str(row[i] or "").strip() for i in di if i < len(row) and row[i]]
            if num not in escalas_map:
                escalas_map[num] = descs
            elif descs and not escalas_map[num]:
                escalas_map[num] = descs
 
    escalas_out = []
    for ns, descs in escalas_map.items():
        try: n = int(float(ns))
        except: n = ns
        escalas_out.append(make_e(num=n, pontos=descs if descs else ["",""]))
 
    # Blocos
    bloco_abas = {}
    for sn in wb.sheetnames:
        ws = wb[sn]
        data = list(ws.iter_rows(values_only=True))
        if len(data) < 2: continue
        hl = [str(c or "").lower().strip() for c in data[0]]
        idx = {f: detect_col(hl, f) for f in SYNONYMS}
 
        def get(row, f):
            i = idx.get(f)
            return str(row[i] or "").strip() if (i is not None and i < len(row)) else ""
 
        for row in data[1:]:
            if not any(c not in (None,"") for c in row): continue
            bn = get(row,"question_set")
            pt = get(row,"pergunta")
            if not bn and not pt: continue
            p = make_p(texto=pt, competencia=get(row,"competencia"),
                       aberta=norm_bool(get(row,"aberta") or "não"),
                       escala=get(row,"escala"),
                       opcional=norm_bool(get(row,"opcional") or "não"),
                       min_c=get(row,"min_caracters"), max_c=get(row,"max_caracters"),
                       definicao=get(row,"definicao"))
            bloco_abas.setdefault(bn, {}).setdefault(sn, []).append(p)
 
    wb.close()
 
    blocos_out = []
    for bn, abas in bloco_abas.items():
        al = list(abas.keys())
        if len(al) == 1:
            blocos_out.append(make_b(nome=bn, perguntas=abas[al[0]], origem="import"))
            continue
        sets = [set(p["texto"] for p in abas[a]) for a in al]
        all_eq = all(s == sets[0] for s in sets[1:])
        if all_eq:
            blocos_out.append(make_b(nome=bn, perguntas=abas[al[0]], origem="import"))
            msgs.append(f"✓ **{bn}** idêntico em {len(al)} abas → mesclado em 1.")
        else:
            for a in al:
                blocos_out.append(make_b(nome=f"{bn} — {a}", perguntas=abas[a], origem="import"))
            msgs.append(f"⚠ **{bn}** diferente em {len(al)} abas → renomeado com sufixo.")
 
    nomes = [b["nome"] for b in blocos_out]
    for i,n1 in enumerate(nomes):
        for n2 in nomes[i+1:]:
            if sim(n1,n2) > 0.85:
                msgs.append(f"⚠ Blocos parecidos: **{n1}** e **{n2}**")
 
    return blocos_out, escalas_out, msgs
 
# ── Validação ─────────────────────────────────────────────────────────────────
def validar():
    erros, avisos = [], []
    enum = {str(e["num"]) for e in ss.escalas}
    if not ss.escalas: erros.append("Nenhuma escala cadastrada.")
    if not ss.blocos:  erros.append("Nenhum bloco cadastrado.")
    seen = set()
    nomes = [b["nome"].strip() for b in ss.blocos]
    for n in nomes:
        if n in seen: erros.append(f'Bloco duplicado: "{n}"')
        seen.add(n)
    for i,n1 in enumerate(nomes):
        for n2 in nomes[i+1:]:
            if n1 != n2 and sim(n1,n2) > 0.85:
                avisos.append(f'Blocos muito parecidos: "{n1}" e "{n2}"')
    total = 0
    for b in ss.blocos:
        if not b["nome"].strip(): erros.append(f'Bloco sem nome (id {b["id"]})')
        if not b["perguntas"]: avisos.append(f'Bloco "{b["nome"]}" sem perguntas.')
        for p in b["perguntas"]:
            total += 1
            if not p["texto"].strip(): erros.append(f'Pergunta vazia em "{b["nome"]}".')
            if p["aberta"] == "não":
                if not p["escala"]: avisos.append(f'Pergunta fechada sem escala em "{b["nome"]}".')
                elif p["escala"] not in enum: erros.append(f'Escala {p["escala"]} inexistente em "{b["nome"]}".')
            if p["aberta"] == "sim" and not p["max_caracters"]:
                avisos.append(f'Aberta sem máx. caracteres em "{b["nome"]}".')
    return erros, avisos, total
 
# ── Export ────────────────────────────────────────────────────────────────────
def build_export():
    out = io.BytesIO()
    rows = []
    for b in ss.blocos:
        for p in b["perguntas"]:
            rows.append({
                "question_set": b["nome"], "desc_question_set": b["desc"],
                "competencia": p["competencia"], "pergunta": p["texto"],
                "opcional": p["opcional"], "aberta": p["aberta"],
                "escala": p["escala"] if p["aberta"]=="não" else "",
                "min caracters": p["min_caracters"], "max caracters": p["max_caracters"],
                "definição": p["definicao"],
            })
    df = pd.DataFrame(rows, columns=OUTPUT_COLS)
 
    max_pts = max((len(e["pontos"]) for e in ss.escalas), default=2)
    esc_rows = []
    for e in ss.escalas:
        r = {"escala": e["num"],
             "min": e["pontos"][0] if e["pontos"] else "",
             "max": e["pontos"][-1] if e["pontos"] else ""}
        for i,p in enumerate(e["pontos"],1): r[f"descrição_{i}"] = p
        esc_rows.append(r)
    esc_cols = ["escala","min","max"] + [f"descrição_{i}" for i in range(1,max_pts+1)]
    df_esc = pd.DataFrame(esc_rows)
    for c in esc_cols:
        if c not in df_esc.columns: df_esc[c] = ""
    df_esc = df_esc[[c for c in esc_cols if c in df_esc.columns]]
 
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Competências e Perguntas", index=False)
        df_esc.to_excel(writer, sheet_name="Escalas", index=False)
        ws = writer.sheets["Competências e Perguntas"]
        for i,w in enumerate([28,20,22,60,8,8,8,12,12,40],1):
            ws.column_dimensions[ws.cell(1,i).column_letter].width = w
    return out.getvalue()
 
 
# ═══════════════════════════════════════════════════════════════════════════════
# STEP 1 — Escalas
# ═══════════════════════════════════════════════════════════════════════════════
if ss.step == 1:
    steps_bar()
    st.markdown("## Escalas de avaliação")
    st.caption("Defina as escalas usadas nas perguntas fechadas. Cada escala tem um número e pontos de ancoragem.")
 
    with st.expander("📥 Importar de uma planilha existente", expanded=not ss.imported):
        up = st.file_uploader("Arquivo .xlsx do cliente", type=["xlsx","xls"],
                              key="up1", label_visibility="collapsed")
        if up:
            with st.spinner("Lendo..."):
                bi, ei, mi = import_xlsx(up.read())
            st.success(f"✓ {len(bi)} bloco(s) e {len(ei)} escala(s) encontrados.")
            if mi:
                with st.expander("O que foi ajustado"):
                    for m in mi: st.markdown(m)
            if st.button("✓ Usar esses dados", type="primary", key="use_import"):
                ss.blocos = bi
                ss.escalas = ei
                ss.imported = True
                st.rerun()
 
    st.divider()
    st.markdown('<div class="sec-title">Escalas cadastradas</div>', unsafe_allow_html=True)
 
    for ei_idx, esc in enumerate(ss.escalas):
        with st.container(border=True):
            c1, c2 = st.columns([6,1])
            with c1:
                new_num = st.number_input("Número da escala", value=int(esc["num"]) if str(esc["num"]).isdigit() else ei_idx+1,
                                          min_value=1, key=f"enum_{esc['id']}")
                esc["num"] = new_num
 
                # Preview
                if any(p.strip() for p in esc["pontos"]):
                    prev = "".join(f'<div class="escala-ponto"><span class="escala-num">{i+1}</span>{p or "—"}</div>'
                                   for i,p in enumerate(esc["pontos"]))
                    st.markdown(f'<div class="escala-preview">{prev}</div>', unsafe_allow_html=True)
                    st.markdown("")
 
                with st.expander("✏ Editar pontos de ancoragem"):
                    n_pts = st.number_input("Nº de pontos", min_value=2, max_value=10,
                                            value=len(esc["pontos"]), key=f"npts_{esc['id']}")
                    while len(esc["pontos"]) < n_pts: esc["pontos"].append("")
                    esc["pontos"] = esc["pontos"][:n_pts]
                    cols = st.columns(min(n_pts,5))
                    for pi in range(n_pts):
                        esc["pontos"][pi] = cols[pi % len(cols)].text_input(
                            f"Ponto {pi+1}", value=esc["pontos"][pi],
                            key=f"pt_{esc['id']}_{pi}", placeholder=f"Ex: Nível {pi+1}")
            with c2:
                st.markdown("<br><br>", unsafe_allow_html=True)
                if st.button("✕", key=f"del_e_{esc['id']}"):
                    ss.escalas.pop(ei_idx); st.rerun()
 
    if st.button("＋ Adicionar escala", key="add_e"):
        ss.escalas.append(make_e()); st.rerun()
 
    if not ss.escalas:
        st.info("Importe uma planilha ou adicione uma escala manualmente.")
 
    nav(fwd=2, fwd_label="Blocos & Perguntas →", fwd_disabled=not ss.escalas)
 
 
# ═══════════════════════════════════════════════════════════════════════════════
# STEP 2 — Blocos & Perguntas
# ═══════════════════════════════════════════════════════════════════════════════
elif ss.step == 2:
    steps_bar()
    escala_opts = [""] + [str(e["num"]) for e in ss.escalas]
 
    col_list, col_detail = st.columns([2,3], gap="large")
 
    # ── Lista de blocos ──
    with col_list:
        st.markdown("## Blocos")
        total_p = sum(len(b["perguntas"]) for b in ss.blocos)
        st.caption(f"{len(ss.blocos)} bloco(s) · {total_p} pergunta(s)")
 
        with st.expander("📥 Importar mais perguntas"):
            up2 = st.file_uploader("Arquivo .xlsx", type=["xlsx","xls"],
                                   key="up2", label_visibility="collapsed")
            if up2:
                with st.spinner("Importando..."):
                    bi2, _, mi2 = import_xlsx(up2.read())
                if mi2:
                    for m in mi2: st.caption(m)
                if st.button("Adicionar aos blocos existentes", key="btn_imp2"):
                    existing = {b["nome"]: b for b in ss.blocos}
                    for nb in bi2:
                        if nb["nome"] in existing:
                            existing[nb["nome"]]["perguntas"].extend(nb["perguntas"])
                        else:
                            ss.blocos.append(nb)
                    st.rerun()
 
        st.markdown("")
 
        for bi_idx, bloco in enumerate(ss.blocos):
            n_p = len(bloco["perguntas"])
            n_err = sum(1 for p in bloco["perguntas"]
                        if not p["texto"] or (p["aberta"]=="não" and not p["escala"]))
            is_sel = ss.bloco_sel == bloco["id"]
 
            badges = f'<span class="badge badge-blue">{n_p}p</span>'
            badges += f' <span class="badge badge-red">{n_err} erro{"s" if n_err!=1 else ""}</span>' if n_err \
                      else ' <span class="badge badge-green">OK</span>'
            if bloco["origem"] == "import":
                badges += ' <span class="badge badge-gray">importado</span>'
 
            card_cls = "bloco-card sel" if is_sel else "bloco-card"
            st.markdown(f"""
            <div class="{card_cls}">
                <div class="bloco-title">{bloco["nome"] or "(sem nome)"}</div>
                <div class="bloco-meta">{badges}</div>
            </div>
            """, unsafe_allow_html=True)
 
            bc1, bc2 = st.columns([3,1])
            with bc1:
                if st.button("Editar", key=f"sel_{bloco['id']}", use_container_width=True):
                    ss.bloco_sel = bloco["id"]; st.rerun()
            with bc2:
                if st.button("✕", key=f"del_{bloco['id']}"):
                    ss.blocos = [b for b in ss.blocos if b["id"] != bloco["id"]]
                    if ss.bloco_sel == bloco["id"]: ss.bloco_sel = None
                    st.rerun()
 
        st.markdown("")
        if st.button("＋ Novo bloco", use_container_width=True, key="new_b"):
            nb = make_b(nome="Novo bloco")
            ss.blocos.append(nb)
            ss.bloco_sel = nb["id"]
            st.rerun()
 
    # ── Painel de detalhe ──
    with col_detail:
        bloco = next((b for b in ss.blocos if b["id"] == ss.bloco_sel), None)
 
        if bloco is None:
            st.markdown("## Perguntas")
            st.info("👈 Selecione um bloco ao lado para editar suas perguntas.")
        else:
            st.markdown(f"## Editando bloco")
 
            bloco["nome"] = st.text_input("Nome do bloco *", value=bloco["nome"], key=f"bn_{bloco['id']}")
            bloco["desc"] = st.text_input("Descrição do bloco (opcional)", value=bloco["desc"],
                                          key=f"bd_{bloco['id']}", placeholder="Breve descrição...")
            st.divider()
            st.markdown(f'<div class="sec-title">Perguntas — {len(bloco["perguntas"])}</div>', unsafe_allow_html=True)
 
            for pi, p in enumerate(bloco["perguntas"]):
                with st.container(border=True):
                    r1c1, r1c2 = st.columns([9,1])
                    p["texto"] = r1c1.text_area(f"Pergunta {pi+1}", value=p["texto"],
                                                key=f"pt_{p['id']}", label_visibility="collapsed",
                                                placeholder="Texto da pergunta...", height=75)
                    r1c2.markdown("<br>", unsafe_allow_html=True)
                    if r1c2.button("✕", key=f"dp_{p['id']}"):
                        bloco["perguntas"].pop(pi); st.rerun()
 
                    ra, rb, rc = st.columns([2,1,2])
                    p["competencia"] = ra.text_input("Competência", value=p["competencia"],
                                                     key=f"pc_{p['id']}", placeholder="Ex: Trabalho em equipe")
                    tipo = rb.selectbox("Tipo", ["Fechada","Aberta"],
                                        index=0 if p["aberta"]=="não" else 1, key=f"pt2_{p['id']}")
                    p["aberta"] = "não" if tipo=="Fechada" else "sim"
 
                    if p["aberta"] == "não":
                        p["escala"] = rc.selectbox("Escala *", escala_opts,
                                                    index=escala_opts.index(p["escala"]) if p["escala"] in escala_opts else 0,
                                                    key=f"pe_{p['id']}")
                    else:
                        rm1, rm2 = rc.columns(2)
                        p["min_caracters"] = rm1.text_input("Mín. chars", value=p["min_caracters"],
                                                             key=f"pmin_{p['id']}", placeholder="0")
                        p["max_caracters"] = rm2.text_input("Máx. chars", value=p["max_caracters"],
                                                             key=f"pmax_{p['id']}", placeholder="10000")
 
                    with st.expander("Campos opcionais"):
                        oc1, oc2 = st.columns(2)
                        p["opcional"] = "sim" if oc1.checkbox("Pergunta opcional?",
                                                               value=p["opcional"]=="sim",
                                                               key=f"popc_{p['id']}") else "não"
                        p["definicao"] = oc2.text_input("Definição", value=p["definicao"], key=f"pdef_{p['id']}")
 
            st.markdown("")
            if st.button("＋ Adicionar pergunta", key=f"ap_{bloco['id']}", use_container_width=True):
                bloco["perguntas"].append(make_p()); st.rerun()
 
    nav(back=1, fwd=3, fwd_label="Revisar →", fwd_disabled=not ss.blocos)
 
 
# ═══════════════════════════════════════════════════════════════════════════════
# STEP 3 — Revisão
# ═══════════════════════════════════════════════════════════════════════════════
elif ss.step == 3:
    steps_bar()
    st.markdown("## Revisão")
    erros, avisos, total_p = validar()
 
    m1,m2,m3,m4 = st.columns(4)
    m1.metric("Blocos", len(ss.blocos))
    m2.metric("Perguntas", total_p)
    m3.metric("Escalas", len(ss.escalas))
    m4.metric("Erros", len(erros), delta=str(len(erros)) if erros else None, delta_color="inverse" if erros else "off")
 
    st.divider()
 
    if not erros and not avisos:
        st.success("✅ Tudo certo! Pronto para exportar.")
    else:
        if erros:
            st.markdown("**❌ Erros (obrigatório corrigir):**")
            for e in erros: st.markdown(f'<div class="val-err">✕ {e}</div>', unsafe_allow_html=True)
            st.markdown("")
        if avisos:
            st.markdown("**⚠️ Avisos (recomendado verificar):**")
            for a in avisos: st.markdown(f'<div class="val-warn">⚠ {a}</div>', unsafe_allow_html=True)
 
    st.divider()
    st.markdown('<div class="sec-title">Resumo por bloco</div>', unsafe_allow_html=True)
    for b in ss.blocos:
        np_ = len(b["perguntas"])
        nab = sum(1 for p in b["perguntas"] if p["aberta"]=="sim")
        esc_u = sorted(set(p["escala"] for p in b["perguntas"] if p["escala"]))
        st.markdown(
            f"**{b['nome']}** &nbsp;"
            f"<span style='color:rgba(255,255,255,.4);font-size:12px;'>"
            f"{np_}p · {np_-nab} fechada(s) · {nab} aberta(s) · escala(s): {', '.join(esc_u) or '—'}"
            f"</span>", unsafe_allow_html=True)
 
    nav(back=2, fwd=4, fwd_label="Exportar →", fwd_disabled=bool(erros))
 
 
# ═══════════════════════════════════════════════════════════════════════════════
# STEP 4 — Exportar
# ═══════════════════════════════════════════════════════════════════════════════
elif ss.step == 4:
    steps_bar()
    st.markdown("## Exportar")
    erros, avisos, total_p = validar()
 
    if avisos:
        st.warning(f"{len(avisos)} aviso(s). O arquivo será gerado mesmo assim.")
 
    c1,c2,c3 = st.columns(3)
    c1.metric("Blocos", len(ss.blocos))
    c2.metric("Perguntas", total_p)
    c3.metric("Escalas", len(ss.escalas))
 
    st.divider()
    st.download_button(
        "⬇ Baixar Base_Perguntas-2.xlsx",
        data=build_export(),
        file_name="Base_Perguntas-2.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )
    st.caption("Abas geradas: **Competências e Perguntas** e **Escalas** — formato esperado pelo script de importação.")
 
    st.divider()
    cb, _, cr = st.columns([1,4,1])
    with cb:
        if st.button("← Voltar"): ss.step=3; st.rerun()
    with cr:
        if st.button("↺ Nova configuração"):
            ss.escalas=[]; ss.blocos=[]; ss.bloco_sel=None; ss.imported=False; ss.step=1
            st.rerun()
