"""
Conversor de Perguntas AVD
Transforma o arquivo xlsx do cliente (N abas) em Base_Perguntas-2.xlsx
pronto para o script Bases_AVD.ipynb.
 
Rodar: streamlit run app.py
"""
 
import io
import re
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
 
# ─── Página ──────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Conversor de Perguntas AVD",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)
 
# ─── CSS custom ──────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .block-container { padding-top: 4rem !important; padding-bottom: 2rem; }
    section[data-testid="stSidebar"] .stMarkdown p { color: inherit !important; }
    .stTabs [data-baseweb="tab"] { font-size: 0.85rem; }
    div[data-testid="stMetric"] {
        background: rgba(255,255,255,0.05);
        border-radius: 8px;
        padding: 12px 16px;
        border: 1px solid rgba(255,255,255,0.12);
    }
    div[data-testid="stMetricValue"] { font-size: 1.6rem !important; }
    div[data-testid="stMetricLabel"] { opacity: 0.7; }
    .badge-err  { background:rgba(192,57,43,0.2); color:#ff7675; border-radius:4px; padding:2px 8px; font-size:11px; font-weight:600; }
    .badge-warn { background:rgba(230,126,34,0.2); color:#fdcb6e; border-radius:4px; padding:2px 8px; font-size:11px; font-weight:600; }
    .badge-ok   { background:rgba(39,174,96,0.2); color:#55efc4; border-radius:4px; padding:2px 8px; font-size:11px; font-weight:600; }
    .step-header { font-size:1.35rem; font-weight:700; margin-bottom:0.2rem; }
    .step-sub { opacity:0.6; font-size:0.9rem; margin-bottom:1.2rem; }
    hr { margin: 1.2rem 0; }
</style>
""", unsafe_allow_html=True)
 
# ─── Constantes ──────────────────────────────────────────────────────────────
REQUIRED_COLS = ["question_set", "pergunta", "opcional", "aberta"]
 
SYNONYMS = {
    "question_set":      ["nome do bloco", "bloco", "question_set", "block", "grupo", "section", "seção", "categoria"],
    "desc_question_set": ["descrição do bloco", "desc_question_set", "desc bloco"],
    "competencia":       ["competência", "competencia", "nome da competência", "skill", "tema"],
    "pergunta":          ["pergunta", "question", "enunciado", "texto", "item"],
    "opcional":          ["essa pergunta é opcional", "opcional", "optional"],
    "aberta":            ["é uma pergunta aberta", "aberta", "open", "tipo", "type"],
    "escala":            ["qual o número do modelo da escala", "escala", "scale", "número da escala", "modelo da escala"],
    "min_caracters":     ["número mínimo de caracteres", "min caracters", "min_caracters", "min caracteres"],
    "max_caracters":     ["número máximo de caracteres", "max caracters", "max_caracters", "max caracteres"],
    "definicao":         ["descrição geral sobre a pergunta", "definição", "definicao", "definition"],
}
 
OUTPUT_COLS = [
    "question_set", "desc_question_set", "competencia", "pergunta",
    "opcional", "aberta", "escala", "min caracters", "max caracters", "definição",
]
 
# ─── Helpers ─────────────────────────────────────────────────────────────────
def detect_mapping(headers: list[str]) -> dict:
    """Auto-detecta mapeamento de colunas por sinônimos."""
    mapping = {}
    headers_lower = [h.lower().strip() for h in headers]
    for field, syns in SYNONYMS.items():
        for i, h in enumerate(headers_lower):
            if any(s in h or h in s for s in syns):
                mapping[field] = headers[i]
                break
    return mapping
 
 
def norm_bool(val: str) -> str:
    v = str(val).strip().lower()
    if v in ("sim", "yes", "true", "1", "s", "y"):
        return "sim"
    if v in ("não", "nao", "no", "false", "0", "n"):
        return "não"
    return str(val).strip()
 
 
def validate_row(row: pd.Series) -> tuple[list, list]:
    errors, warnings = [], []
    if not str(row.get("question_set", "")).strip():
        errors.append("question_set vazio")
    if not str(row.get("pergunta", "")).strip():
        errors.append("pergunta vazia")
    for f in ("opcional", "aberta"):
        v = str(row.get(f, "")).strip()
        if v not in ("sim", "não"):
            errors.append(f"{f} inválido: '{v}'")
    if str(row.get("aberta", "")) == "não" and not str(row.get("escala", "")).strip():
        warnings.append("fechada sem escala")
    if str(row.get("aberta", "")) == "sim" and not str(row.get("max_caracters", "")).strip():
        warnings.append("aberta sem max_caracters")
    if 0 < len(str(row.get("pergunta", ""))) < 5:
        warnings.append("pergunta muito curta")
    return errors, warnings
 
 
def revalidate(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["_errors"] = ""
    df["_warnings"] = ""
    df["_status"] = "ok"
    for i, row in df.iterrows():
        errs, warns = validate_row(row)
        df.at[i, "_errors"] = " | ".join(errs)
        df.at[i, "_warnings"] = " | ".join(warns)
        df.at[i, "_status"] = "erro" if errs else ("aviso" if warns else "ok")
    return df
 
 
def load_sheets(file_bytes: bytes) -> dict[str, pd.DataFrame]:
    """Lê todas as abas do xlsx, retorna dict nome→df."""
    wb = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    sheets = {}
    for name in wb.sheetnames:
        ws = wb[name]
        data = list(ws.iter_rows(values_only=True))
        if len(data) < 2:
            continue
        headers = [str(c).strip() if c is not None else "" for c in data[0]]
        rows = [
            [str(c).strip() if c is not None else "" for c in row]
            for row in data[1:]
            if any(c not in (None, "") for c in row)
        ]
        if rows:
            sheets[name] = pd.DataFrame(rows, columns=headers)
    wb.close()
    return sheets
 
 
def extract_scales(sheets: dict, selected: list, col_map: dict) -> pd.DataFrame:
    """Extrai escalas únicas do arquivo do cliente."""
    scale_col = col_map.get("escala", "")
    # Cols de descrição de pontos da escala (índices fixos após o número)
    DESC_SYNONYMS = ["descrição do primeiro", "descrição do segundo", "descrição do terceiro",
                     "descrição do quarto", "descrição do quinto", "duplique"]
    scale_data = {}  # num -> {min, max, desc_1..N}
 
    for name in selected:
        df_raw = sheets[name]
        headers = list(df_raw.columns)
        headers_l = [h.lower() for h in headers]
 
        # Achar colunas de descrição de pontos
        desc_cols = [headers[i] for i, h in enumerate(headers_l)
                     if any(s in h for s in DESC_SYNONYMS)]
 
        for _, row in df_raw.iterrows():
            num = str(row.get(scale_col, "") if scale_col and scale_col in df_raw.columns else "").strip()
            if not num:
                continue
            descs = [str(row.get(c, "")).strip() for c in desc_cols if str(row.get(c, "")).strip()]
            if num not in scale_data:
                scale_data[num] = {"descs": descs or []}
            elif descs and not scale_data[num]["descs"]:
                scale_data[num]["descs"] = descs
 
    if not scale_data:
        return pd.DataFrame(columns=["escala", "min", "max"] + [f"descrição_{i}" for i in range(1, 6)])
 
    rows = []
    for num, data in sorted(scale_data.items(), key=lambda x: (float(x[0]) if x[0].replace(".", "").isdigit() else x[0])):
        descs = data["descs"]
        n = len(descs)
        row = {"escala": num, "min": descs[0] if n > 0 else "", "max": descs[-1] if n > 0 else ""}
        for i, d in enumerate(descs, 1):
            row[f"descrição_{i}"] = d
        rows.append(row)
    return pd.DataFrame(rows)
 
 
def dedup_blocks(df: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    """
    Resolve blocos duplicados:
    - Mesmo nome + mesmas perguntas → mantém um só (mescla, remove duplicatas)
    - Mesmo nome + perguntas diferentes → renomeia com sufixo da aba de origem
    Retorna df limpo e lista de mensagens do que foi feito.
    """
    messages = []
    from difflib import SequenceMatcher
 
    # Agrupa por nome de bloco
    blocos = df["question_set"].unique()
    name_groups = {}
    for b in blocos:
        mask = df["question_set"] == b
        abas = df[mask]["_sheet"].unique()
        if len(abas) <= 1:
            continue
        # Pega perguntas por aba
        pergs_por_aba = {}
        for aba in abas:
            pergs = set(df[(df["question_set"] == b) & (df["_sheet"] == aba)]["pergunta"].str.strip())
            pergs_por_aba[aba] = pergs
        name_groups[b] = pergs_por_aba
 
    rows_to_drop = []
    renames = {}  # (bloco, aba) -> novo nome
 
    for bloco, pergs_por_aba in name_groups.items():
        abas = list(pergs_por_aba.keys())
        # Verificar se todas as abas têm perguntas idênticas
        base = pergs_por_aba[abas[0]]
        all_identical = all(pergs_por_aba[a] == base for a in abas[1:])
 
        if all_identical:
            # Mescla: mantém só da primeira aba, dropa o resto
            keep_aba = abas[0]
            for aba in abas[1:]:
                idxs = df[(df["question_set"] == bloco) & (df["_sheet"] == aba)].index.tolist()
                rows_to_drop.extend(idxs)
            messages.append(f'✓ **"{bloco}"** era idêntico em {len(abas)} abas → mesclado em um único bloco.')
        else:
            # Perguntas diferentes → renomeia com sufixo da aba
            for aba in abas:
                novo = f"{bloco} ({aba})"
                renames[(bloco, aba)] = novo
            messages.append(f'⚠ **"{bloco}"** tinha perguntas diferentes em {len(abas)} abas → renomeado com sufixo da aba.')
 
    # Aplicar drops
    df = df.drop(index=rows_to_drop).reset_index(drop=True)
 
    # Aplicar renames
    def apply_rename(row):
        key = (row["question_set"], row["_sheet"])
        return renames.get(key, row["question_set"])
    if renames:
        df["question_set"] = df.apply(apply_rename, axis=1)
 
    # Detectar nomes muito parecidos (fuzzy) entre blocos diferentes
    from difflib import SequenceMatcher
    blocos_finais = list(df["question_set"].unique())
    similares = []
    for i, b1 in enumerate(blocos_finais):
        for b2 in blocos_finais[i+1:]:
            ratio = SequenceMatcher(None, b1.lower(), b2.lower()).ratio()
            if ratio > 0.85:
                similares.append((b1, b2, round(ratio * 100)))
    if similares:
        for b1, b2, pct in similares:
            messages.append(f'⚠ Blocos muito parecidos ({pct}% similares): **"{b1}"** e **"{b2}"** — verifique se deveriam ser o mesmo.')
 
    return df, messages
 
 
def merge_sheets(sheets: dict, selected: list, col_map: dict) -> tuple[pd.DataFrame, pd.DataFrame, list[str]]:
    """Mescla abas, deduplica blocos e extrai escalas. Retorna (df_perguntas, df_escalas, mensagens)."""
    parts = []
    for name in selected:
        df_raw = sheets[name]
        rec = {"_sheet": name}
        for field, src_col in col_map.items():
            if src_col and src_col in df_raw.columns:
                rec[field] = df_raw[src_col]
            else:
                rec[field] = ""
        part = pd.DataFrame(rec)
        for f in ("opcional", "aberta"):
            if f in part.columns:
                part[f] = part[f].apply(norm_bool)
        parts.append(part)
 
    if not parts:
        return pd.DataFrame(), pd.DataFrame(), []
 
    df = pd.concat(parts, ignore_index=True)
    key_cols = [c for c in ["question_set", "pergunta"] if c in df.columns]
    df = df[df[key_cols].apply(lambda r: any(str(v).strip() for v in r), axis=1)].copy()
 
    # Deduplicar blocos
    df, messages = dedup_blocks(df)
 
    # Extrair escalas
    df_escalas = extract_scales(sheets, selected, col_map)
 
    return df, df_escalas, messages
 
 
def build_export(df: pd.DataFrame, df_escalas: pd.DataFrame | None = None) -> bytes:
    out = io.BytesIO()
    export_df = df.copy()
    col_rename = {"min_caracters": "min caracters", "max_caracters": "max caracters", "definicao": "definição"}
    export_df = export_df.rename(columns=col_rename)
    keep = [c for c in OUTPUT_COLS if c in export_df.columns]
    export_df = export_df[keep]
 
    # Escalas: usa as extraídas do arquivo, ou gera placeholder
    if df_escalas is not None and not df_escalas.empty:
        esc_df = df_escalas.copy()
    else:
        escalas = sorted(set(
            v for v in df.get("escala", pd.Series()).dropna().unique()
            if str(v).strip()
        ), key=lambda x: (float(x) if str(x).replace(".","",1).isdigit() else x))
        escala_rows = [[str(e), "", ""] + [""] * 5 for e in escalas]
        esc_df = pd.DataFrame(
            escala_rows,
            columns=["escala", "min", "max"] + [f"descrição_{i}" for i in range(1, 6)],
        )
 
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        export_df.to_excel(writer, sheet_name="Competências e Perguntas", index=False)
        esc_df.to_excel(writer, sheet_name="Escalas", index=False)
 
        # Formatação básica das colunas
        ws = writer.sheets["Competências e Perguntas"]
        widths = [28, 20, 22, 60, 8, 8, 8, 12, 12, 40]
        for i, w in enumerate(widths[:ws.max_column], 1):
            ws.column_dimensions[ws.cell(1, i).column_letter].width = w
 
    return out.getvalue()
 
 
# ─── Session state ───────────────────────────────────────────────────────────
def init_state():
    defaults = {
        "step": 1,
        "file_bytes": None,
        "file_name": "",
        "sheets": {},          # nome → df raw
        "selected_sheets": [],
        "col_map": {},
        "df": None,            # df processado + validado
        "df_escalas": None,    # escalas extraídas
        "dedup_messages": [],  # mensagens de deduplicação
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v
 
init_state()
ss = st.session_state
 
 
# ─── Sidebar ─────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 📋 Conversor AVD")
    st.caption("Base_Perguntas-2 · v1.0")
    st.divider()
 
    steps = ["Upload", "Abas & Colunas", "Validação", "Exportar"]
    for i, label in enumerate(steps, 1):
        if i < ss.step:
            st.markdown(f"✓ ~~{i}. {label}~~")
        elif i == ss.step:
            st.markdown(f"**→ {i}. {label}**")
        else:
            st.markdown(f"&nbsp;&nbsp;&nbsp;{i}. {label}", unsafe_allow_html=True)
 
    st.divider()
    st.caption("Arquivos processados localmente.\nNenhum dado sai da sua máquina.")
 
    if ss.step > 1:
        st.divider()
        if st.button("↺ Recomeçar", use_container_width=True):
            for k in ["file_bytes","file_name","sheets","selected_sheets","col_map","df","df_escalas","dedup_messages"]:
                ss[k] = [] if k in ("selected_sheets","dedup_messages") else ({} if k in ("sheets","col_map") else None)
            ss.step = 1
            st.rerun()
 
 
# ═══════════════════════════════════════════════════════════════════════════════
# STEP 1 — Upload
# ═══════════════════════════════════════════════════════════════════════════════
if ss.step == 1:
    st.markdown('<div class="step-header">📂 Upload do arquivo</div>', unsafe_allow_html=True)
    st.markdown('<div class="step-sub">Carregue o xlsx do cliente — qualquer número de abas.</div>', unsafe_allow_html=True)
 
    uploaded = st.file_uploader("Selecione o arquivo", type=["xlsx", "xls"], label_visibility="collapsed")
 
    if uploaded:
        with st.spinner("Lendo arquivo..."):
            try:
                file_bytes = uploaded.read()
                sheets = load_sheets(file_bytes)
                ss.file_bytes = file_bytes
                ss.file_name = uploaded.name
                ss.sheets = sheets
 
                total_rows = sum(len(df) for df in sheets.values())
                col1, col2, col3 = st.columns(3)
                col1.metric("Abas encontradas", len(sheets))
                col2.metric("Linhas totais", total_rows)
                col3.metric("Arquivo", uploaded.name)
 
                st.success(f"✓ Arquivo lido com sucesso — {len(sheets)} abas, {total_rows} linhas.")
 
                if st.button("Continuar →", type="primary"):
                    # Auto-selecionar todas as abas e detectar mapeamento
                    ss.selected_sheets = list(sheets.keys())
                    first_df = next(iter(sheets.values()))
                    ss.col_map = detect_mapping(list(first_df.columns))
                    ss.step = 2
                    st.rerun()
 
            except Exception as e:
                st.error(f"Não foi possível ler o arquivo: {e}")
 
 
# ═══════════════════════════════════════════════════════════════════════════════
# STEP 2 — Abas & Colunas
# ═══════════════════════════════════════════════════════════════════════════════
elif ss.step == 2:
    st.markdown('<div class="step-header">🗂 Abas & Colunas</div>', unsafe_allow_html=True)
    st.markdown('<div class="step-sub">Selecione as abas e confirme o mapeamento de colunas.</div>', unsafe_allow_html=True)
 
    tab_sheets, tab_cols = st.tabs(["Abas", "Mapeamento de colunas"])
 
    # ── Abas ──
    with tab_sheets:
        st.caption("Desmarque abas que não são de perguntas (instruções, índice, etc.).")
        selected = []
        cols = st.columns(3)
        for i, (name, df_raw) in enumerate(ss.sheets.items()):
            checked = st.session_state.get(f"sheet_{i}", name in ss.selected_sheets)
            val = cols[i % 3].checkbox(f"{name} ({len(df_raw)} linhas)", value=checked, key=f"sheet_{i}")
            if val:
                selected.append(name)
        ss.selected_sheets = selected
 
        total_sel = sum(len(ss.sheets[n]) for n in selected)
        st.caption(f"**{len(selected)}** abas selecionadas · **{total_sel}** linhas estimadas")
 
    # ── Mapeamento ──
    with tab_cols:
        st.caption("Detectado automaticamente pelos nomes das colunas. Ajuste se necessário.")
        if ss.selected_sheets:
            first_df = ss.sheets[ss.selected_sheets[0]]
            available_cols = [""] + list(first_df.columns)
 
            FIELD_LABELS = {
                "question_set":      ("question_set *",      "Nome do bloco"),
                "desc_question_set": ("desc_question_set",   "Descrição do bloco (opcional)"),
                "competencia":       ("competencia",         "Nome da competência (opcional)"),
                "pergunta":          ("pergunta *",          "Texto da pergunta"),
                "opcional":          ("opcional *",          "'sim' ou 'não'"),
                "aberta":            ("aberta *",            "'sim' = aberta, 'não' = fechada"),
                "escala":            ("escala",              "Número do modelo de escala"),
                "min_caracters":     ("min caracters",       "Mínimo de caracteres (abertas)"),
                "max_caracters":     ("max caracters",       "Máximo de caracteres (abertas)"),
                "definicao":         ("definição",           "Descrição geral da pergunta"),
            }
 
            updated_map = {}
            for field, (label, hint) in FIELD_LABELS.items():
                current = ss.col_map.get(field, "")
                idx = available_cols.index(current) if current in available_cols else 0
                chosen = st.selectbox(
                    label, available_cols, index=idx,
                    help=hint, key=f"colmap_{field}"
                )
                updated_map[field] = chosen
            ss.col_map = updated_map
        else:
            st.warning("Selecione pelo menos uma aba na aba anterior.")
 
    st.divider()
    missing_req = [f for f in REQUIRED_COLS if not ss.col_map.get(f)]
    if missing_req:
        st.error(f"Colunas obrigatórias não mapeadas: **{', '.join(missing_req)}**")
    
    col_back, col_fwd = st.columns([1, 5])
    with col_back:
        if st.button("← Voltar"):
            ss.step = 1; st.rerun()
    with col_fwd:
        if st.button("Processar e validar →", type="primary", disabled=bool(missing_req)):
            with st.spinner("Mesclando abas e validando..."):
                df, df_escalas, dedup_msgs = merge_sheets(ss.sheets, ss.selected_sheets, ss.col_map)
                df = revalidate(df)
                ss.df = df
                ss.df_escalas = df_escalas
                ss.dedup_messages = dedup_msgs
            ss.step = 3
            st.rerun()
 
 
# ═══════════════════════════════════════════════════════════════════════════════
# STEP 3 — Validação
# ═══════════════════════════════════════════════════════════════════════════════
elif ss.step == 3:
    df = ss.df.copy()
 
    st.markdown('<div class="step-header">🔍 Validação</div>', unsafe_allow_html=True)
    st.markdown('<div class="step-sub">Revise e corrija erros antes de exportar. Use as ferramentas abaixo para edições em massa.</div>', unsafe_allow_html=True)
 
    # Métricas
    n_total = len(df)
    n_err   = (df["_status"] == "erro").sum()
    n_warn  = (df["_status"] == "aviso").sum()
    n_ok    = (df["_status"] == "ok").sum()
 
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total", n_total)
    m2.metric("❌ Erros", n_err, delta=f"-{n_err}" if n_err else None, delta_color="inverse")
    m3.metric("⚠️ Avisos", n_warn)
    m4.metric("✅ OK", n_ok)
 
    # Mensagens de deduplicação de blocos
    if ss.dedup_messages:
        with st.expander(f"🔀 Blocos processados ({len(ss.dedup_messages)} ocorrências)", expanded=True):
            for msg in ss.dedup_messages:
                st.markdown(msg)
 
    # Escalas extraídas
    if ss.df_escalas is not None and not ss.df_escalas.empty:
        with st.expander(f"⚖️ {len(ss.df_escalas)} escala(s) extraída(s) do arquivo"):
            st.dataframe(ss.df_escalas, use_container_width=True, hide_index=True)
            st.caption("As descrições dos pontos serão incluídas no export se estiverem preenchidas no arquivo de origem.")
    else:
        with st.expander("⚖️ Escalas não encontradas"):
            st.warning("Nenhuma escala detectada automaticamente. Verifique se a coluna 'escala' foi mapeada corretamente no passo 2.")
 
    st.divider()
 
    # ── Ferramentas de correção em massa ──────────────────────────────────────
    with st.expander("🔧 Correções em massa", expanded=(n_err > 0 or n_warn > 0)):
 
        massa_tab1, massa_tab2, massa_tab3 = st.tabs([
            "Aplicar valor em lote",
            "Buscar & substituir",
            "Preencher vazios",
        ])
 
        # ── Tab 1: Aplicar valor a linhas filtradas ──
        with massa_tab1:
            st.caption("Filtre linhas por critério e aplique o mesmo valor a um campo de uma vez.")
 
            c1, c2 = st.columns(2)
            with c1:
                filtro_campo = st.selectbox("Filtrar por campo", ["(todas as linhas)", "status", "aba", "question_set", "aberta", "escala"], key="lote_filtro_campo")
                filtro_valor = ""
                if filtro_campo == "status":
                    filtro_valor = st.selectbox("Valor do filtro", ["erro", "aviso", "ok"], key="lote_filtro_val")
                elif filtro_campo == "aba":
                    filtro_valor = st.selectbox("Aba", ss.selected_sheets, key="lote_filtro_aba")
                elif filtro_campo != "(todas as linhas)":
                    filtro_valor = st.text_input("Valor do filtro", key="lote_filtro_txt")
 
            with c2:
                campo_alvo = st.selectbox("Campo a alterar", [
                    "question_set", "competencia", "opcional", "aberta", "escala",
                    "min_caracters", "max_caracters", "desc_question_set",
                ], key="lote_campo_alvo")
 
                if campo_alvo in ("opcional", "aberta"):
                    novo_valor = st.selectbox("Novo valor", ["não", "sim"], key="lote_novo_val")
                else:
                    novo_valor = st.text_input("Novo valor", key="lote_novo_txt")
 
            # Preview quantas linhas serão afetadas
            if filtro_campo == "(todas as linhas)":
                mask = pd.Series([True] * len(df))
            elif filtro_campo == "status":
                mask = df["_status"] == filtro_valor
            elif filtro_campo == "aba":
                mask = df["_sheet"] == filtro_valor
            else:
                mask = df[filtro_campo].astype(str).str.lower().str.contains(str(filtro_valor).lower(), na=False)
 
            n_afetadas = mask.sum()
            st.caption(f"**{n_afetadas}** linhas serão afetadas.")
 
            if st.button(f"✎ Aplicar em {n_afetadas} linhas", type="primary", key="btn_lote", disabled=(n_afetadas == 0)):
                df.loc[mask, campo_alvo] = novo_valor
                df = revalidate(df)
                ss.df = df
                st.success(f"✓ {n_afetadas} linhas atualizadas.")
                st.rerun()
 
        # ── Tab 2: Buscar & substituir ──
        with massa_tab2:
            st.caption("Localiza ocorrências exatas ou parciais em um campo e substitui.")
            c1, c2, c3 = st.columns(3)
            with c1:
                br_campo = st.selectbox("Campo", [
                    "question_set", "competencia", "pergunta", "escala",
                    "opcional", "aberta", "definicao",
                ], key="br_campo")
            with c2:
                br_busca = st.text_input("Buscar", key="br_busca")
            with c3:
                br_subst = st.text_input("Substituir por", key="br_subst")
 
            br_exato = st.checkbox("Correspondência exata (não parcial)", key="br_exato")
            br_case  = st.checkbox("Case sensitive", key="br_case")
 
            if br_busca:
                if br_exato:
                    if br_case:
                        mask_br = df[br_campo].astype(str) == br_busca
                    else:
                        mask_br = df[br_campo].astype(str).str.lower() == br_busca.lower()
                else:
                    flags = 0 if br_case else re.IGNORECASE
                    mask_br = df[br_campo].astype(str).str.contains(re.escape(br_busca), flags=flags, na=False)
 
                n_br = mask_br.sum()
                st.caption(f"**{n_br}** ocorrências encontradas.")
 
                if n_br > 0:
                    with st.expander(f"Prévia das {min(n_br,10)} primeiras"):
                        st.dataframe(
                            df[mask_br][[br_campo, "_sheet", "question_set"]].head(10),
                            use_container_width=True, hide_index=True,
                        )
 
                if st.button(f"🔁 Substituir {n_br} ocorrências", type="primary", key="btn_br", disabled=(n_br == 0)):
                    if br_exato:
                        df.loc[mask_br, br_campo] = br_subst
                    else:
                        flags = 0 if br_case else re.IGNORECASE
                        df.loc[mask_br, br_campo] = df.loc[mask_br, br_campo].astype(str).str.replace(
                            re.escape(br_busca), br_subst, flags=flags, regex=True
                        )
                    df = revalidate(df)
                    ss.df = df
                    st.success(f"✓ {n_br} substituições realizadas.")
                    st.rerun()
 
        # ── Tab 3: Preencher vazios ──
        with massa_tab3:
            st.caption("Preenche células vazias num campo com um valor fixo ou propagando o valor anterior (útil para question_set).")
 
            c1, c2 = st.columns(2)
            with c1:
                pv_campo = st.selectbox("Campo", [
                    "question_set", "competencia", "escala", "opcional", "aberta",
                ], key="pv_campo")
            with c2:
                pv_modo = st.radio("Modo", ["Valor fixo", "Propagar valor anterior (fill-down)"], key="pv_modo", horizontal=True)
 
            pv_valor = ""
            if pv_modo == "Valor fixo":
                if pv_campo in ("opcional", "aberta"):
                    pv_valor = st.selectbox("Valor", ["não", "sim"], key="pv_val_bool")
                else:
                    pv_valor = st.text_input("Valor", key="pv_val_txt")
 
            mask_vazio = df[pv_campo].astype(str).str.strip() == ""
            n_vazio = mask_vazio.sum()
            st.caption(f"**{n_vazio}** células vazias em '{pv_campo}'.")
 
            if st.button(f"↓ Preencher {n_vazio} células", type="primary", key="btn_pv", disabled=(n_vazio == 0)):
                if pv_modo == "Valor fixo":
                    df.loc[mask_vazio, pv_campo] = pv_valor
                else:
                    # fill-down: substitui "" por None para que ffill funcione
                    df[pv_campo] = df[pv_campo].replace("", None).ffill().fillna("")
                df = revalidate(df)
                ss.df = df
                st.success(f"✓ Células preenchidas.")
                st.rerun()
 
    st.divider()
 
    # ── Filtros e tabela ─────────────────────────────────────────────────────
    f1, f2, f3, f4 = st.columns([1, 1, 2, 2])
    filtro_status = f1.selectbox("Status", ["todos", "erro", "aviso", "ok"], key="fil_status")
    filtro_aba    = f2.selectbox("Aba", ["todas"] + ss.selected_sheets, key="fil_aba")
    filtro_bloco  = f3.text_input("Buscar bloco / competência", key="fil_bloco")
    filtro_perg   = f4.text_input("Buscar texto da pergunta", key="fil_perg")
 
    mask_view = pd.Series([True] * len(df))
    if filtro_status != "todos":
        mask_view &= df["_status"] == filtro_status
    if filtro_aba != "todas":
        mask_view &= df["_sheet"] == filtro_aba
    if filtro_bloco:
        mask_view &= (
            df["question_set"].str.contains(filtro_bloco, case=False, na=False) |
            df.get("competencia", pd.Series([""] * len(df))).str.contains(filtro_bloco, case=False, na=False)
        )
    if filtro_perg:
        mask_view &= df["pergunta"].str.contains(filtro_perg, case=False, na=False)
 
    df_view = df[mask_view].copy()
    st.caption(f"{len(df_view)} linhas exibidas de {len(df)} totais.")
 
    # Colunas visíveis na tabela
    display_cols = ["_sheet", "question_set", "competencia", "pergunta", "aberta", "escala", "opcional", "_errors", "_warnings"]
    rename_map = {
        "_sheet": "Aba", "question_set": "Bloco", "competencia": "Competência",
        "pergunta": "Pergunta", "aberta": "Tipo", "escala": "Escala",
        "opcional": "Opcional", "_errors": "Erros", "_warnings": "Avisos",
    }
 
    edited = st.data_editor(
        df_view[[c for c in display_cols if c in df_view.columns]].rename(columns=rename_map),
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        column_config={
            "Tipo":     st.column_config.SelectboxColumn("Tipo",     options=["não", "sim"]),
            "Opcional": st.column_config.SelectboxColumn("Opcional", options=["não", "sim"]),
            "Erros":    st.column_config.TextColumn("Erros",    disabled=True),
            "Avisos":   st.column_config.TextColumn("Avisos",   disabled=True),
            "Aba":      st.column_config.TextColumn("Aba",      disabled=True),
        },
        key="data_editor",
    )
 
    if st.button("💾 Salvar edições da tabela", key="btn_save_edit"):
        edited_back = edited.rename(columns={v: k for k, v in rename_map.items()})
        for col in edited_back.columns:
            if col in df.columns and col not in ("_errors", "_warnings", "_status", "_sheet"):
                df.loc[mask_view, col] = edited_back[col].values
        df = revalidate(df)
        ss.df = df
        st.success("✓ Edições salvas.")
        st.rerun()
 
    st.divider()
    col_back, col_fwd = st.columns([1, 5])
    with col_back:
        if st.button("← Voltar"):
            ss.step = 2; st.rerun()
    with col_fwd:
        btn_label = "Continuar →" if n_err == 0 else f"Continuar com {n_err} erro(s) →"
        if st.button(btn_label, type="primary"):
            ss.step = 4; st.rerun()
 
 
# ═══════════════════════════════════════════════════════════════════════════════
# STEP 4 — Exportar
# ═══════════════════════════════════════════════════════════════════════════════
elif ss.step == 4:
    df = ss.df
 
    st.markdown('<div class="step-header">⬇️ Exportar</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="step-sub">Arquivo pronto para o script <code>Bases_AVD.ipynb</code>.</div>',
        unsafe_allow_html=True,
    )
 
    n_err  = (df["_status"] == "erro").sum()
    n_warn = (df["_status"] == "aviso").sum()
    blocos = df["question_set"].dropna().unique()
    escalas = df["escala"].replace("", None).dropna().unique()
 
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Perguntas", len(df))
    c2.metric("Blocos", len(blocos))
    c3.metric("Erros restantes", n_err)
    c4.metric("Avisos", n_warn)
 
    if n_err > 0:
        st.warning(f"⚠️ {n_err} linhas com erro serão exportadas — o script pode falhar nelas. Recomendado voltar e corrigir.")
 
    st.divider()
 
    # Resumo por bloco
    with st.expander("📊 Resumo por bloco"):
        resumo = df.groupby("question_set").agg(
            Perguntas=("pergunta", "count"),
            Abertas=("aberta", lambda x: (x == "sim").sum()),
            Fechadas=("aberta", lambda x: (x == "não").sum()),
            Erros=("_status", lambda x: (x == "erro").sum()),
        ).reset_index().rename(columns={"question_set": "Bloco"})
        st.dataframe(resumo, use_container_width=True, hide_index=True)
 
    with st.expander("⚖️ Escalas referenciadas"):
        if len(escalas):
            st.info(f"Escalas encontradas: **{', '.join(sorted(str(e) for e in escalas))}**\n\nA aba *Escalas* no arquivo exportado contém uma linha placeholder para cada uma — preencha os pontos de ancoragem.")
        else:
            st.warning("Nenhuma escala referenciada. Verifique se o mapeamento da coluna 'escala' está correto.")
 
    st.divider()
 
    xlsx_bytes = build_export(df, ss.df_escalas)
    st.download_button(
        label="⬇️ Baixar Base_Perguntas-2.xlsx",
        data=xlsx_bytes,
        file_name="Base_Perguntas-2.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )
 
    st.divider()
    col_back, col_novo = st.columns([1, 5])
    with col_back:
        if st.button("← Revisar"):
            ss.step = 3; st.rerun()
    with col_novo:
        if st.button("↺ Novo arquivo"):
            for k in ["file_bytes","file_name","sheets","selected_sheets","col_map","df","df_escalas","dedup_messages"]:
                ss[k] = [] if k in ("selected_sheets","dedup_messages") else ({} if k in ("sheets","col_map") else None)
            ss.step = 1
            st.rerun()
