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
        prompt: Mensagem do usuário.
        system_prompt: System prompt completo.
        model: 'sonnet', 'opus', 'haiku' ou nome completo do modelo.
        timeout: Timeout em segundos (padrão: 5 min).
        progress_cb: Callback(str) para atualizar status na UI.

    Returns:
        Texto gerado pelo Claude.
    """
    if progress_cb:
        progress_cb("Chamando Claude CLI — aguarde...")

    # System prompt via arquivo — evita limite de tamanho de argumento do Windows
    sp_file = Path(tempfile.mktemp(suffix="_sp.txt"))
    sp_file.write_text(system_prompt, encoding="utf-8")

    try:
        cmd = [
            CLAUDE_EXE,
            "--print",
            "--no-session-persistence",
            "--system-prompt-file", str(sp_file),
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
            cwd=tempfile.gettempdir(),
        )

    finally:
        sp_file.unlink(missing_ok=True)

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
            [CLAUDE_EXE, "--print", "--no-session-persistence", "Responda apenas: OK"],
            capture_output=True, timeout=60,
            cwd=tempfile.gettempdir(),
        )
        return r.returncode == 0 and b"OK" in r.stdout
    except Exception:
        return False
