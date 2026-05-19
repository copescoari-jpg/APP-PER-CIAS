#!/usr/bin/env python3
"""
Cria o ícone do SISTEMA ARI e o atalho na área de trabalho.
Execute uma vez: python3.12 criar_atalho.py
"""

import subprocess
import tempfile
from pathlib import Path
from PIL import Image, ImageDraw

SCRIPT_DIR = Path(__file__).parent.resolve()
ICON_OUT   = SCRIPT_DIR / "icone_ari.ico"
VBS_OUT    = SCRIPT_DIR / "iniciar_app.vbs"
MAIN_PY    = SCRIPT_DIR / "main.py"


# ── Geração do ícone ──────────────────────────────────────────────────────────

def gerar_icone() -> None:
    S = 256
    img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)

    NAVY  = (27,  52,  85)
    GOLD  = (196, 160,  60)
    WHITE = (255, 255, 255)
    LGRAY = (210, 220, 235)
    LBLUE = ( 80, 110, 155)

    d.rounded_rectangle([(0, 0), (S-1, S-1)], radius=52, fill=NAVY)

    M, TOP, BOT, FOLD = 46, 44, 202, 32
    d.polygon(
        [(M, TOP), (S-M-FOLD, TOP), (S-M, TOP+FOLD), (S-M, BOT), (M, BOT)],
        fill=WHITE,
    )
    d.polygon(
        [(S-M-FOLD, TOP), (S-M, TOP+FOLD), (S-M-FOLD, TOP+FOLD)],
        fill=LGRAY,
    )
    d.rectangle([(M, TOP), (S-M-FOLD, TOP+8)], fill=GOLD)

    try:
        from PIL import ImageFont
        fb = ImageFont.truetype("C:/Windows/Fonts/arialbd.ttf", 64)
        fs = ImageFont.truetype("C:/Windows/Fonts/arial.ttf",   15)
    except Exception:
        fb = fs = ImageFont.load_default()

    bb = d.textbbox((0, 0), "ARI", font=fb)
    tx = (S - (bb[2] - bb[0])) // 2 - bb[0]
    ty = TOP + 28 - bb[1]
    d.text((tx, ty), "ARI", fill=NAVY, font=fb)

    ly = ty + (bb[3] - bb[1]) + 12
    d.line([(M + 18, ly), (S - M - 18, ly)], fill=GOLD, width=2)

    try:
        bb2 = d.textbbox((0, 0), "LAUDOS PERICIAIS", font=fs)
        sx  = (S - (bb2[2] - bb2[0])) // 2 - bb2[0]
        d.text((sx, ly + 8), "LAUDOS PERICIAIS", fill=LBLUE, font=fs)
    except Exception:
        pass

    img.save(
        str(ICON_OUT), format="ICO",
        sizes=[(16,16),(32,32),(48,48),(64,64),(128,128),(256,256)],
    )
    print(f"  Icone: {ICON_OUT}")


# ── Launcher VBScript (sem janela de console) ─────────────────────────────────

def criar_vbs() -> None:
    """Lança o Streamlit com python3.12 sem abrir janela de console."""
    cmd = (
        f'python3.12 -m streamlit run "{MAIN_PY}" '
        '--server.headless false --browser.gatherUsageStats false'
    )
    vbs = f'CreateObject("WScript.Shell").Run "{cmd}", 0, False\n'
    VBS_OUT.write_text(vbs, encoding="utf-8")
    print(f"  Launcher: {VBS_OUT}")


# ── Atalho na área de trabalho ────────────────────────────────────────────────

def criar_atalho() -> None:
    r = subprocess.run(
        ["powershell", "-NoProfile", "-Command",
         "[Environment]::GetFolderPath('Desktop')"],
        capture_output=True, text=True,
    )
    desktop = Path(r.stdout.strip())
    lnk     = desktop / "SISTEMA ARI.lnk"
    wscript = r"C:\Windows\System32\wscript.exe"

    ps = "\n".join([
        '$ws = New-Object -ComObject WScript.Shell',
        f'$s = $ws.CreateShortcut("{lnk}")',
        f'$s.TargetPath = "{wscript}"',
        f'$s.Arguments = \'"{VBS_OUT}"\'',
        f'$s.WorkingDirectory = "{SCRIPT_DIR}"',
        f'$s.IconLocation = "{ICON_OUT}"',
        '$s.Description = "SISTEMA ARI - Laudos Periciais"',
        '$s.WindowStyle = 1',
        '$s.Save()',
    ])

    tmp = Path(tempfile.mktemp(suffix=".ps1"))
    tmp.write_text(ps, encoding="utf-8")
    res = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
         "-File", str(tmp)],
        capture_output=True, text=True,
    )
    tmp.unlink(missing_ok=True)

    if res.returncode == 0:
        print(f"  Atalho: {lnk}")
    else:
        print(f"  ERRO ao criar atalho:\n{res.stderr}")


# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    print("Atualizando icone e atalho do SISTEMA ARI...\n")
    gerar_icone()
    criar_vbs()
    criar_atalho()
    print("\nConcluido! Clique em 'SISTEMA ARI' na area de trabalho.")
