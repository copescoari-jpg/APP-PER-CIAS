"""
SISTEMA ARI — GeraLaudo
Interface Streamlit: geração de laudos periciais e impugnações.
"""

import os
import re
import threading
from pathlib import Path

# Prefixo Windows para caminhos > 260 chars (MAX_PATH)
_LONG_PFX = '\\\\?\\'


def _walk_safe(folder: Path):
    """Varre recursivamente com suporte a caminhos longos no Windows."""
    root = str(folder.resolve())
    scan = _LONG_PFX + root if len(root) > 240 else root
    for dirpath, _, filenames in os.walk(scan):
        base = dirpath.replace(_LONG_PFX, '', 1)
        for fname in filenames:
            yield Path(os.path.join(base, fname))

import streamlit as st

from core.pdf_reader  import extrair_texto_pdf, extrair_texto_docx
from core.prompts     import montar_prompt_laudo, montar_prompt_impugnacao, SYSTEM_PROMPT_LAUDO, SYSTEM_PROMPT_IMPUGNACAO
from core.claude_runner import chamar_claude
from core.vault       import detectar_agentes, carregar_contexto_nr
from core.docx_builder import salvar_docx, construir_nome_arquivo, caminho_saida
from db.database      import inicializar, upsert_processo, salvar_laudo, listar_processos

# ── constantes ────────────────────────────────────────────────────────────────
_RE_PROC  = re.compile(r'\b(\d{4,7}-\d{2}\.\d{4}\.\d\.\d{2}\.\d{4})\b')
_RE_NOME  = re.compile(r'Reclamante[:\s]+([A-ZÁÉÍÓÚÃÕÂÊÎÔÛÇÀÜ][a-záéíóúãõâêîôûçàü\s]+)', re.I)
_RE_RECDA = re.compile(r'Reclamada[:\s]+([A-ZÁÉÍÓÚÃÕÂÊÎÔÛÇÀÜ][a-záéíóúãõâêîôûçàü\s&.\-/]+)', re.I)
_RE_FUNC  = re.compile(r'[Ff]un[cç][aã]o[:\s]+([A-ZÁÉÍÓÚÃÕÂÊÎÔÛÇÀÜ][\wÀ-ú\s]+)', re.I)

FOTO_EXTS   = {".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".webp"}
DOC_EXTS    = {".docx", ".pdf", ".doc", ".dotx", ".dotm"}
EVAL_TOKENS = {"avalia", "medic", "relat", "nho", "nr-", "laudo_avaliacao", "dosim"}
LAUDO_TOKENS = {"laudo", "periç", "peric"}
IMP_TOKENS   = {"impugn", "quesit", "complement", "esclarec"}

MAX_FOTOS = 20


# ── inicialização do banco ─────────────────────────────────────────────────────
inicializar()


# ── helpers ───────────────────────────────────────────────────────────────────

def _nome_lower(p: Path) -> str:
    return p.stem.lower()


def _is_avaliacao(p: Path) -> bool:
    return any(t in _nome_lower(p) for t in EVAL_TOKENS)


def _is_laudo(p: Path) -> bool:
    return any(t in _nome_lower(p) for t in LAUDO_TOKENS)


def _is_impugnacao(p: Path) -> bool:
    return any(t in _nome_lower(p) for t in IMP_TOKENS)


def _extrair_texto_doc(p: Path) -> str:
    ext = p.suffix.lower()
    if ext in {".docx", ".doc", ".dotx", ".dotm"}:
        txt = extrair_texto_docx(p)
        if not txt.startswith("[Erro"):
            return txt
    if ext == ".pdf":
        return extrair_texto_pdf(p, max_chars=40_000)
    return ""


def auto_detect(pasta: str) -> dict:
    """Varre a pasta e classifica arquivos automaticamente."""
    folder = Path(pasta)
    if not folder.is_dir():
        return {}

    docs        = []
    avaliacoes  = []
    imp_doc     = None
    fotos       = []

    for p in _walk_safe(folder):
        ext = p.suffix.lower()
        if ext in FOTO_EXTS:
            fotos.append(p)
            continue
        if ext not in DOC_EXTS:
            continue
        if _is_avaliacao(p):
            avaliacoes.append(p)
        elif _is_impugnacao(p):
            imp_doc = p
        elif not _is_laudo(p):
            docs.append(p)

    # documento principal: preferir .docx, depois mais recente
    def _mtime(p: Path) -> float:
        try:
            lp = _LONG_PFX + str(p.resolve())
            return os.stat(lp).st_mtime
        except OSError:
            return 0.0
    docs.sort(key=lambda p: (0 if p.suffix.lower() == ".docx" else 1, -_mtime(p)))
    main_doc = docs[0] if docs else None
    fotos    = sorted(fotos, key=lambda p: p.name)[:MAX_FOTOS]

    # extrair meta do doc principal
    texto_base = _extrair_texto_doc(main_doc) if main_doc else ""
    processo_n = ""
    for src in [texto_base, pasta]:
        m = _RE_PROC.search(src)
        if m:
            processo_n = m.group(1)
            break

    reclamante = ""
    m = _RE_NOME.search(texto_base)
    if m:
        reclamante = m.group(1).strip().title()

    reclamada = ""
    m = _RE_RECDA.search(texto_base)
    if m:
        reclamada = m.group(1).strip().title()

    funcao = ""
    m = _RE_FUNC.search(texto_base)
    if m:
        funcao = m.group(1).strip().title()

    insalubr   = bool(re.search(r'insalubr', texto_base, re.I))
    periculos  = bool(re.search(r'periculosid|periculos', texto_base, re.I))

    return {
        "main_doc":       str(main_doc) if main_doc else None,
        "avaliacoes":     avaliacoes,
        "imp_doc":        str(imp_doc) if imp_doc else None,
        "fotos":          fotos,
        "texto_base":     texto_base,
        "processo_numero": processo_n,
        "reclamante":     reclamante,
        "reclamada":      reclamada,
        "funcao":         funcao,
        "insalubridade":  insalubr,
        "periculosidade": periculos,
    }


def _gerar_laudo_thread(doc_text, obs, numero_processo, reclamante, reclamada,
                         funcao, insalubridade, periculosidade,
                         avaliacoes_paths, pasta_saida, result_store):
    """Executa geração em thread separada para não travar a UI."""
    try:
        agentes   = detectar_agentes(doc_text)
        nr_ctx    = carregar_contexto_nr(agentes)
        prompt    = montar_prompt_laudo(doc_text, obs=obs,
                                        nr_contexto=nr_ctx,
                                        numero_processo=numero_processo)
        texto     = chamar_claude(prompt, SYSTEM_PROMPT_LAUDO, model="sonnet", timeout=360)

        nome_arq  = construir_nome_arquivo(reclamante, reclamada,
                                           insalubridade, periculosidade, funcao)
        caminho   = caminho_saida(pasta_saida, nome_arq)
        salvar_docx(texto, caminho,
                    avaliacoes_paths=[Path(p) for p in avaliacoes_paths])

        pid = upsert_processo(numero_processo or nome_arq,
                              reclamante=reclamante, reclamada=reclamada,
                              funcao=funcao, pasta=pasta_saida,
                              insalubridade=int(insalubridade),
                              periculosidade=int(periculosidade))
        salvar_laudo(pid, texto, caminho)

        result_store["texto"]   = texto
        result_store["caminho"] = caminho
        result_store["ok"]      = True
    except Exception as exc:
        result_store["erro"] = str(exc)
        result_store["ok"]   = False


def _gerar_impugnacao_thread(doc_text, laudo_text, obs, pasta_saida, result_store):
    try:
        prompt = montar_prompt_impugnacao(doc_text, meu_laudo=laudo_text, obs=obs)
        texto  = chamar_claude(prompt, SYSTEM_PROMPT_IMPUGNACAO, model="sonnet", timeout=300)
        caminho = caminho_saida(pasta_saida, "Esclarecimentos")
        salvar_docx(texto, caminho)
        result_store["texto"]   = texto
        result_store["caminho"] = caminho
        result_store["ok"]      = True
    except Exception as exc:
        result_store["erro"] = str(exc)
        result_store["ok"]   = False


# ── página ────────────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="SISTEMA ARI — GeraLaudo",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
    /* cabeçalho principal */
    .ari-header {
        background: linear-gradient(135deg, #1a3a6b 0%, #2a5298 100%);
        border-radius: 12px;
        padding: 20px 32px 16px;
        margin-bottom: 28px;
    }
    .ari-header h1 { color: #fff; margin: 0; font-size: 1.6rem; font-weight: 700; }
    .ari-header p  { color: #b8c8e8; margin: 4px 0 0; font-size: 0.9rem; }

    /* cartão azul — informações detectadas */
    .card-info {
        background: #f0f6ff;
        border: 2px solid #8ab0d4;
        border-radius: 12px;
        padding: 16px 20px;
        margin-bottom: 16px;
    }

    /* cartão âmbar — observações */
    .card-obs {
        background: #fffbf0;
        border: 2px solid #c8a020;
        border-radius: 12px;
        padding: 16px 20px;
        margin-bottom: 16px;
    }
    .card-obs .obs-label {
        color: #c8a020; font-weight: 700; font-size: 0.85rem;
        letter-spacing: 0.5px; margin-bottom: 6px;
    }

    /* badge de tipo */
    .badge {
        display: inline-block;
        background: #1a3a6b; color: #fff;
        border-radius: 6px; padding: 2px 10px;
        font-size: 0.78rem; font-weight: 600;
        margin-right: 4px;
    }
    .badge-green { background: #1a6b3a; }
    .badge-amber { background: #8b6200; }

    /* botão primário */
    div.stButton > button[kind="primary"] {
        background: #1a3a6b !important;
        color: #fff !important;
        border: none !important;
        border-radius: 8px !important;
        font-size: 1rem !important;
        padding: 10px 28px !important;
        font-weight: 700 !important;
    }
    div.stButton > button[kind="primary"]:hover {
        background: #2a5298 !important;
    }

    /* rodapé */
    .footer-txt { color: #888; font-size: 0.78rem; text-align: center; margin-top: 40px; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="ari-header">
  <h1>SISTEMA ARI — GeraLaudo</h1>
  <p>Ari Vladimir Copesco Júnior · Engenheiro de Segurança do Trabalho · CREA 060097553-3 · Perito Judicial TRT-15</p>
</div>
""", unsafe_allow_html=True)

aba_laudo, aba_impug, aba_hist = st.tabs(["Laudo Pericial", "Impugnação / Esclarecimentos", "Histórico"])


# ══════════════════════════════════════════════════════════════════════════════
# ABA 1 — LAUDO PERICIAL
# ══════════════════════════════════════════════════════════════════════════════
with aba_laudo:

    # ── Passo 1: pasta ────────────────────────────────────────────────────────
    st.markdown("#### 1. Selecione a pasta do processo")
    st.caption("Deve conter o documento base (pré-laudo), fotos e avaliações técnicas.")
    pasta_input = st.text_input(
        "Caminho da pasta", key="pasta_laudo",
        placeholder=r"Ex: C:\Processos\João Silva x Empresa XYZ",
    )

    detected = {}
    if pasta_input and Path(pasta_input).is_dir():
        with st.spinner("Lendo pasta..."):
            detected = auto_detect(pasta_input)

        if not detected.get("main_doc"):
            st.warning("Nenhum documento principal encontrado na pasta.")
        else:
            # ── Passo 2: card de informações ──────────────────────────────────
            st.markdown("#### 2. Confirme as informações detectadas")

            n_fotos = len(detected.get("fotos", []))
            n_aval  = len(detected.get("avaliacoes", []))

            tipo_badges = ""
            if detected.get("insalubridade"):
                tipo_badges += '<span class="badge">Insalubridade</span>'
            if detected.get("periculosidade"):
                tipo_badges += '<span class="badge badge-amber">Periculosidade</span>'
            if not tipo_badges:
                tipo_badges = '<span style="color:#888;font-size:0.85rem">Não detectado</span>'

            st.markdown(f"""
<div class="card-info">
  <table style="width:100%;border-collapse:collapse">
    <tr>
      <td style="width:50%;padding:4px 8px;color:#555;font-size:0.85rem">Documento principal</td>
      <td style="padding:4px 8px;font-size:0.85rem"><b>{Path(detected['main_doc']).name}</b></td>
    </tr>
    <tr>
      <td style="padding:4px 8px;color:#555;font-size:0.85rem">Fotos encontradas</td>
      <td style="padding:4px 8px;font-size:0.85rem"><b>{n_fotos}</b></td>
    </tr>
    <tr>
      <td style="padding:4px 8px;color:#555;font-size:0.85rem">Avaliações técnicas</td>
      <td style="padding:4px 8px;font-size:0.85rem"><b>{n_aval}</b></td>
    </tr>
    <tr>
      <td style="padding:4px 8px;color:#555;font-size:0.85rem">Tipo de perícia</td>
      <td style="padding:4px 8px">{tipo_badges}</td>
    </tr>
  </table>
</div>
""", unsafe_allow_html=True)

            col1, col2 = st.columns(2)
            with col1:
                num_proc = st.text_input(
                    "Nº do processo (CNJ)",
                    value=detected.get("processo_numero", ""),
                    key="num_proc_laudo",
                )
                reclamante = st.text_input(
                    "Reclamante",
                    value=detected.get("reclamante", ""),
                    key="reclamante_laudo",
                )
            with col2:
                reclamada = st.text_input(
                    "Reclamada",
                    value=detected.get("reclamada", ""),
                    key="reclamada_laudo",
                )
                funcao = st.text_input(
                    "Função",
                    value=detected.get("funcao", ""),
                    key="funcao_laudo",
                )

            col3, col4 = st.columns(2)
            with col3:
                insalubridade = st.checkbox("Insalubridade",
                                            value=detected.get("insalubridade", False),
                                            key="insalubridade_laudo")
            with col4:
                periculosidade = st.checkbox("Periculosidade",
                                             value=detected.get("periculosidade", False),
                                             key="periculosidade_laudo")

            # ── Passo 3: observações ──────────────────────────────────────────
            st.markdown("#### 3. Observações do perito")
            st.markdown("""
<div class="card-obs">
  <div class="obs-label">★ PRIORIDADE MÁXIMA</div>
  <p style="margin:0;font-size:0.82rem;color:#7a6000">
    Tudo escrito aqui prevalece sobre qualquer informação dos documentos.
    Use para correções, instruções específicas ou informações de campo.
  </p>
</div>
""", unsafe_allow_html=True)
            obs_laudo = st.text_area(
                "Observações (opcional)",
                key="obs_laudo",
                height=120,
                placeholder="Ex: Ruído medido foi 91 dB(A), exposição de 8h diárias sem EPI adequado...",
            )

            # ── Passo 4: gerar ────────────────────────────────────────────────
            st.markdown("#### 4. Gerar laudo")
            gerar_btn = st.button("Gerar Laudo", type="primary", key="btn_gerar_laudo",
                                  use_container_width=True)

            if gerar_btn:
                if not detected.get("texto_base"):
                    st.error("Documento principal está vazio ou ilegível.")
                else:
                    result_store: dict = {}
                    prog = st.progress(0, text="Iniciando geração...")

                    t = threading.Thread(
                        target=_gerar_laudo_thread,
                        args=(
                            detected["texto_base"],
                            obs_laudo,
                            num_proc,
                            reclamante,
                            reclamada,
                            funcao,
                            insalubridade,
                            periculosidade,
                            [str(p) for p in detected.get("avaliacoes", [])],
                            pasta_input,
                            result_store,
                        ),
                        daemon=True,
                    )
                    t.start()

                    import time
                    step = 0
                    msgs = [
                        "Analisando documento...",
                        "Carregando referências técnicas (NRs)...",
                        "Claude gerando laudo — aguarde...",
                        "Formatando .docx...",
                    ]
                    while t.is_alive():
                        frac = min(0.9, step / 40)
                        prog.progress(frac, text=msgs[min(step // 10, len(msgs) - 1)])
                        step += 1
                        time.sleep(1)

                    t.join()
                    prog.progress(1.0, text="Concluído!")

                    if result_store.get("ok"):
                        st.success(f"Laudo salvo em: `{result_store['caminho']}`")
                        with st.expander("Pré-visualização do texto"):
                            st.text(result_store["texto"][:8000] + (
                                "\n\n[...texto truncado na pré-visualização...]"
                                if len(result_store["texto"]) > 8000 else ""
                            ))
                    else:
                        st.error(f"Erro: {result_store.get('erro', 'desconhecido')}")

    elif pasta_input:
        st.error("Pasta não encontrada. Verifique o caminho digitado.")


# ══════════════════════════════════════════════════════════════════════════════
# ABA 2 — IMPUGNAÇÃO / ESCLARECIMENTOS
# ══════════════════════════════════════════════════════════════════════════════
with aba_impug:

    st.markdown("#### 1. Selecione a pasta do processo")
    pasta_imp = st.text_input(
        "Caminho da pasta",
        key="pasta_impug",
        placeholder=r"Ex: C:\Processos\João Silva x Empresa XYZ",
    )

    detected_imp = {}
    if pasta_imp and Path(pasta_imp).is_dir():
        with st.spinner("Lendo pasta..."):
            detected_imp = auto_detect(pasta_imp)

        imp_doc_path = detected_imp.get("imp_doc")
        main_doc_path = detected_imp.get("main_doc")

        st.markdown("#### 2. Documentos detectados")

        col_a, col_b = st.columns(2)
        with col_a:
            st.markdown("**Impugnação / Quesitos recebidos:**")
            imp_override = st.text_input(
                "Caminho (ou deixe o detectado)",
                value=imp_doc_path or "",
                key="imp_doc_path",
            )
        with col_b:
            st.markdown("**Meu laudo original (referência):**")
            laudo_override = st.text_input(
                "Caminho (opcional)",
                value=main_doc_path or "",
                key="laudo_ref_path",
            )

        st.markdown("#### 3. Observações do perito")
        st.markdown("""
<div class="card-obs">
  <div class="obs-label">★ PRIORIDADE MÁXIMA</div>
  <p style="margin:0;font-size:0.82rem;color:#7a6000">
    Instruções específicas para esta resposta — prevalece sobre qualquer documento.
  </p>
</div>
""", unsafe_allow_html=True)
        obs_imp = st.text_area(
            "Observações (opcional)",
            key="obs_impug",
            height=100,
        )

        st.markdown("#### 4. Gerar esclarecimentos")
        gerar_imp_btn = st.button("Gerar Esclarecimentos", type="primary",
                                  key="btn_gerar_imp", use_container_width=True)

        if gerar_imp_btn:
            imp_path = Path(imp_override) if imp_override else None
            if not imp_path or not imp_path.exists():
                st.error("Informe o caminho do documento de impugnação/quesitos.")
            else:
                doc_text_imp   = _extrair_texto_doc(imp_path)
                laudo_text_imp = ""
                lp = Path(laudo_override) if laudo_override else None
                if lp and lp.exists():
                    laudo_text_imp = _extrair_texto_doc(lp)

                result_imp: dict = {}
                prog2 = st.progress(0, text="Iniciando...")

                t2 = threading.Thread(
                    target=_gerar_impugnacao_thread,
                    args=(doc_text_imp, laudo_text_imp, obs_imp, pasta_imp, result_imp),
                    daemon=True,
                )
                t2.start()

                import time
                step2 = 0
                while t2.is_alive():
                    prog2.progress(min(0.9, step2 / 30), text="Claude redigindo esclarecimentos...")
                    step2 += 1
                    time.sleep(1)
                t2.join()
                prog2.progress(1.0, text="Concluído!")

                if result_imp.get("ok"):
                    st.success(f"Salvo em: `{result_imp['caminho']}`")
                    with st.expander("Pré-visualização"):
                        st.text(result_imp["texto"][:6000])
                else:
                    st.error(f"Erro: {result_imp.get('erro', 'desconhecido')}")

    elif pasta_imp:
        st.error("Pasta não encontrada.")


# ══════════════════════════════════════════════════════════════════════════════
# ABA 3 — HISTÓRICO
# ══════════════════════════════════════════════════════════════════════════════
with aba_hist:
    st.markdown("#### Últimos processos gerados")

    processos = listar_processos(50)
    if not processos:
        st.info("Nenhum processo no banco ainda.")
    else:
        for proc in processos:
            tipos = []
            if proc.get("insalubridade"):
                tipos.append("Insalubr.")
            if proc.get("periculosidade"):
                tipos.append("Periculos.")
            tipo_str = " + ".join(tipos) or "—"

            with st.expander(f"{proc['numero'] or '(sem nº)'} — {proc['reclamante'] or '?'} x {proc['reclamada'] or '?'}"):
                st.markdown(f"""
- **Função:** {proc['funcao'] or '—'}
- **Tipo:** {tipo_str}
- **Pasta:** `{proc['pasta'] or '—'}`
- **Atualizado:** {proc['atualizado_em']}
""")


# ── rodapé ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="footer-txt">
  SISTEMA ARI · Ari Vladimir Copesco Júnior · CREA 060097553-3 · TRT-15
</div>
""", unsafe_allow_html=True)
