"""
Acesso SQLite — sem ORM, sqlite3 puro.
Banco criado em ~/.sistema_ari/geralaud.db
"""

import sqlite3
from pathlib import Path

_DB_DIR  = Path.home() / ".sistema_ari"
_DB_PATH = _DB_DIR / "geralaud.db"
_SCHEMA  = Path(__file__).parent / "schema.sql"


def _conn() -> sqlite3.Connection:
    _DB_DIR.mkdir(exist_ok=True)
    con = sqlite3.connect(str(_DB_PATH))
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA foreign_keys = ON")
    return con


def inicializar():
    """Cria as tabelas se ainda não existirem."""
    sql = _SCHEMA.read_text(encoding="utf-8")
    with _conn() as con:
        con.executescript(sql)


# ── processos ─────────────────────────────────────────────────────────────────

def upsert_processo(numero: str, **campos) -> int:
    """Insere ou atualiza processo pelo número CNJ. Retorna o id."""
    with _conn() as con:
        row = con.execute(
            "SELECT id FROM processos WHERE numero = ?", (numero,)
        ).fetchone()
        if row:
            pid = row["id"]
            sets  = ", ".join(f"{k} = ?" for k in campos)
            vals  = list(campos.values()) + [pid]
            con.execute(
                f"UPDATE processos SET {sets}, atualizado_em = datetime('now','localtime') WHERE id = ?",
                vals,
            )
        else:
            cols = ", ".join(["numero"] + list(campos))
            phs  = ", ".join(["?"] * (1 + len(campos)))
            vals = [numero] + list(campos.values())
            cur  = con.execute(
                f"INSERT INTO processos ({cols}) VALUES ({phs})", vals
            )
            pid = cur.lastrowid
    return pid


def get_processo(numero: str) -> dict | None:
    with _conn() as con:
        row = con.execute(
            "SELECT * FROM processos WHERE numero = ?", (numero,)
        ).fetchone()
    return dict(row) if row else None


def listar_processos(limit: int = 50) -> list[dict]:
    with _conn() as con:
        rows = con.execute(
            "SELECT * FROM processos ORDER BY atualizado_em DESC LIMIT ?", (limit,)
        ).fetchall()
    return [dict(r) for r in rows]


# ── laudos ────────────────────────────────────────────────────────────────────

def salvar_laudo(processo_id: int, texto: str, arquivo_docx: str,
                 modelo: str = "sonnet") -> int:
    with _conn() as con:
        cur = con.execute(
            "INSERT INTO laudos (processo_id, texto, arquivo_docx, modelo) VALUES (?,?,?,?)",
            (processo_id, texto, arquivo_docx, modelo),
        )
    return cur.lastrowid


def laudos_do_processo(processo_id: int) -> list[dict]:
    with _conn() as con:
        rows = con.execute(
            "SELECT * FROM laudos WHERE processo_id = ? ORDER BY gerado_em DESC",
            (processo_id,),
        ).fetchall()
    return [dict(r) for r in rows]
