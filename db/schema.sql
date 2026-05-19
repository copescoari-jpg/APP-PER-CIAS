-- SISTEMA ARI — GeraLaudo
-- Schema SQLite

CREATE TABLE IF NOT EXISTS processos (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    numero          TEXT NOT NULL,
    reclamante      TEXT NOT NULL DEFAULT '',
    reclamada       TEXT NOT NULL DEFAULT '',
    funcao          TEXT NOT NULL DEFAULT '',
    vara            TEXT NOT NULL DEFAULT '',
    insalubridade   INTEGER NOT NULL DEFAULT 0,
    periculosidade  INTEGER NOT NULL DEFAULT 0,
    pasta           TEXT NOT NULL DEFAULT '',
    criado_em       TEXT NOT NULL DEFAULT (datetime('now','localtime')),
    atualizado_em   TEXT NOT NULL DEFAULT (datetime('now','localtime'))
);

CREATE TABLE IF NOT EXISTS laudos (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    processo_id     INTEGER NOT NULL REFERENCES processos(id) ON DELETE CASCADE,
    texto           TEXT NOT NULL DEFAULT '',
    arquivo_docx    TEXT NOT NULL DEFAULT '',
    modelo          TEXT NOT NULL DEFAULT 'sonnet',
    gerado_em       TEXT NOT NULL DEFAULT (datetime('now','localtime'))
);

CREATE TABLE IF NOT EXISTS diligencias (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    processo_id     INTEGER NOT NULL REFERENCES processos(id) ON DELETE CASCADE,
    data_dilig      TEXT NOT NULL DEFAULT '',
    local           TEXT NOT NULL DEFAULT '',
    observacoes     TEXT NOT NULL DEFAULT '',
    registrado_em   TEXT NOT NULL DEFAULT (datetime('now','localtime'))
);

CREATE INDEX IF NOT EXISTS idx_laudos_processo ON laudos(processo_id);
CREATE INDEX IF NOT EXISTS idx_processos_numero ON processos(numero);
