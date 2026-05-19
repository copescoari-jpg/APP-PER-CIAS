"""
Wrapper para chamar o Claude Code CLI via subprocess.
Usa `claude --print` para saída não-interativa.
"""

import subprocess
import tempfile
from pathlib import Path

CLAUDE_EXE = str(Path.home() / ".local" / "bin" / "claude.exe")


def chamar_claude(
    prompt: str,
    system_prompt: str,
    model: str = "sonnet",
    timeout: int = 300,
    progress_cb=None,
) -> str:
    """
    Chama o Claude Code CLI em modo não-interativo.

    Args:
        prompt: Mensagem do usuário (pode ser muito longa — enviada via stdin).
        system_prompt: System prompt completo (substitui o padrão do CLI).
        model: 'sonnet', 'opus', 'haiku' ou nome completo do modelo.
        timeout: Timeout em segundos (padrão: 5 min).
        progress_cb: Callback(str) para atualizar status na UI.

    Returns:
        Texto gerado pelo Claude.
    """
    if progress_cb:
        progress_cb("Chamando Claude CLI — aguarde...")

    cmd = [
        CLAUDE_EXE,
        "--print",
        "--system-prompt", system_prompt,
        "--model", model,
        "--output-format", "text",
    ]

    result = subprocess.run(
        cmd,
        input=prompt,
        capture_output=True,
        text=True,
        encoding="utf-8",
        timeout=timeout,
        cwd=tempfile.gettempdir(),  # evita CLAUDE.md auto-discovery
    )

    if result.returncode != 0:
        stderr = result.stderr.strip()
        raise RuntimeError(f"Claude CLI falhou (código {result.returncode}):\n{stderr}")

    texto = result.stdout.strip()
    if not texto:
        raise RuntimeError("Claude CLI retornou resposta vazia.")

    return texto


def testar_claude() -> bool:
    """Verifica se o Claude CLI está acessível e funcionando."""
    try:
        r = subprocess.run(
            [CLAUDE_EXE, "--print", "Responda apenas: OK"],
            capture_output=True, text=True, encoding="utf-8", timeout=30,
            cwd=tempfile.gettempdir(),
        )
        return r.returncode == 0 and "OK" in r.stdout
    except Exception:
        return False
