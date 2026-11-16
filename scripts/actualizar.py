"""Actualiza la memoria y el resumen desde la planilla, en tu PC.

Uso (desde la carpeta del repo):
    python scripts/actualizar.py

1. Toma planilla/Rullo_PSV_estatico_fatiga.xlsx.
   - Si encuentra LibreOffice, la recalcula en una copia temporal (no toca tu archivo).
   - Si no, usa los valores guardados por Excel (guardala en Excel antes de correr esto).
2. Escribe memoria/datos.tex y las figuras (scripts/exportar_datos.py).
3. Compila memoria/memoria.tex y memoria/resumen.tex (latexmk, o pdflatex dos veces).

Requisitos: Python con openpyxl y matplotlib; una distribución LaTeX (MiKTeX o TeX Live).
"""
import shutil, subprocess, sys, tempfile
from pathlib import Path

raiz = Path(__file__).resolve().parent.parent
planilla = raiz / "planilla" / "Rullo_PSV_estatico_fatiga.xlsx"

def buscar_soffice():
    for c in ("soffice", "libreoffice",
              r"C:\Program Files\LibreOffice\program\soffice.exe",
              r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"):
        p = shutil.which(c) or (c if Path(c).exists() else None)
        if p:
            return p
    return None

entrada = planilla
soffice = buscar_soffice()
if soffice:
    tmp = Path(tempfile.mkdtemp())
    perfil = tmp / "perfil" / "user"; perfil.mkdir(parents=True)
    (perfil / "registrymodifications.xcu").write_text(
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\n'
        '<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse">'
        '<value>0</value></prop></item>\n</oor:items>\n', encoding="utf-8")
    subprocess.run([soffice, f"-env:UserInstallation={(tmp / 'perfil').as_uri()}", "--headless",
                    "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", str(tmp), str(planilla)],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    entrada = tmp / planilla.name
    print("Planilla recalculada con LibreOffice (copia temporal).")
else:
    print("LibreOffice no encontrado: se usan los valores guardados por Excel.")

r = subprocess.run([sys.executable, str(raiz / "scripts" / "exportar_datos.py"), str(entrada)])
if r.returncode != 0:
    sys.exit("Faltan valores calculados: abrí la planilla en Excel, guardala y volvé a correr.")

mem = raiz / "memoria"
latexmk = shutil.which("latexmk")
for tex in ("memoria.tex", "resumen.tex"):
    if latexmk:
        cmd = [[latexmk, "-pdf", "-interaction=nonstopmode", "-halt-on-error", tex]]
    else:
        cmd = [["pdflatex", "-interaction=nonstopmode", "-halt-on-error", tex]] * 2
    for c in cmd:
        if subprocess.run(c, cwd=mem, stdout=subprocess.DEVNULL).returncode != 0:
            sys.exit(f"Error al compilar {tex}: ver memoria/{tex[:-4]}.log")
    print(f"OK: memoria/{tex[:-4]}.pdf")
