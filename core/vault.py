"""
Carrega trechos relevantes das NRs/NHOs do vault local.
Vault: C:\\Users\\ari\\OneDrive\\Documentos\\SISTEMA ARI\\parte técnica\\
"""

import re
from pathlib import Path
from core.pdf_reader import extrair_texto_pdf, extrair_texto_docx

VAULT_PATH = Path(r"C:\Users\ari\OneDrive\Documentos\SISTEMA ARI\parte técnica")

MAX_CHARS_POR_DOC = 3000

# Mapeamento agente → arquivos do vault
_VAULT_MAP: dict[str, list[str]] = {
    "ruído":          ["NHO01.pdf"],
    "calor":          ["NHO6-051_000012204.pdf"],
    "vibração":       ["NHO08_000002517.pdf", "NHO_09_subst_2_000002292.pdf"],
    "rni":            ["NHO-11_f_4_000001530.pdf"],
    "hidrocarboneto": ["nr-15-atualizada-2025.pdf"],
    "frio":           ["nr-15-atualizada-2025.pdf"],
    "geral":          ["nr-15-atualizada-2025.pdf"],
}

_RE_AGENTES = {
    "ruído":    re.compile(r"ru[íi]do|dosimetri|nho.?01|db\(a\)|dba|nps |press[aã]o sonora", re.I),
    "calor":    re.compile(r"calor|ibutg|nho.?06|temperatura|tgbh", re.I),
    "vibração": re.compile(r"vibra[cç][aã]o|vibracao|nho.?08|nho.?09", re.I),
    "rni":      re.compile(r"rni|radia[cç][aã]o n[aã]o ionizante|solda|ultravioleta|infravermelho|nho.?11", re.I),
    "hidrocarboneto": re.compile(r"hidrocarboneto|benzeno|tolueno|qu[íi]mico|agente qu[íi]mico", re.I),
    "frio":     re.compile(r"c[âa]mara fria|ambiente frio|frio", re.I),
}


def detectar_agentes(texto: str) -> list[str]:
    """Detecta tipos de agentes nocivos a partir de texto livre."""
    agentes = [ag for ag, rx in _RE_AGENTES.items() if rx.search(texto)]
    return agentes if agentes else ["geral"]


def _ler_arquivo_vault(p: Path) -> str:
    """Lê arquivo do vault retornando texto truncado."""
    ext = p.suffix.lower()
    if ext == ".pdf":
        return extrair_texto_pdf(p, max_chars=MAX_CHARS_POR_DOC)
    if ext in (".docx", ".doc"):
        return extrair_texto_docx(p)[:MAX_CHARS_POR_DOC]
    if ext == ".md":
        return p.read_text(encoding="utf-8", errors="ignore")[:MAX_CHARS_POR_DOC]
    return ""


def carregar_contexto_nr(agentes: list[str], max_docs: int = 3) -> str:
    """
    Carrega trechos das NRs/NHOs relevantes para os agentes detectados.
    Retorna string formatada pronta para injeção no prompt.
    """
    if not VAULT_PATH.exists():
        return ""

    arquivos: list[Path] = []
    for ag in agentes:
        for fname in _VAULT_MAP.get(ag, []):
            p = VAULT_PATH / fname
            if p.exists() and p not in arquivos:
                arquivos.append(p)

    partes = []
    for p in arquivos[:max_docs]:
        txt = _ler_arquivo_vault(p)
        if txt and not txt.startswith("[Erro"):
            partes.append(f"=== {p.stem} ===\n{txt}")

    return "\n\n".join(partes)


def listar_arquivos_vault() -> list[str]:
    """Lista todos os arquivos disponíveis no vault."""
    if not VAULT_PATH.exists():
        return []
    return [f.name for f in sorted(VAULT_PATH.iterdir()) if f.is_file()]
