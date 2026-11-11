"""Exporta los valores de la planilla a la memoria LaTeX.

Uso:  python3 scripts/exportar_datos.py [planilla.xlsx]
Lee los nombres definidos (v...) de la planilla y escribe:
  memoria/datos.tex              macros \\vNombre con cada valor
  memoria/figuras/diagramas.pdf  My, Mz, M, flecha y CS a lo largo del rullo
La planilla tiene que estar guardada desde Excel/LibreOffice (valores calculados).
"""
import sys, datetime, math
from pathlib import Path
from openpyxl import load_workbook

raiz = Path(__file__).resolve().parent.parent
xlsx = Path(sys.argv[1]) if len(sys.argv) > 1 else raiz / "planilla" / "Rullo_PSV_estatico_fatiga.xlsx"
wb = load_workbook(xlsx, data_only=True)

CRUDOS = {"vpp", "vpm", "vlado", "vswfl", "vkonuno", "vkondos", "vkontres"}  # enteros para \ifnum

def esc(t):
    for a, b in (("\\", r"\textbackslash{}"), ("&", r"\&"), ("%", r"\%"), ("#", r"\#"), ("_", r"\_"), ("^", r"\^{}")):
        t = t.replace(a, b)
    return t

def fmt(nombre, v):
    if v is None:
        return r"\textbf{??}"
    if nombre in CRUDOS:
        return str(int(round(float(v))))
    if isinstance(v, str):
        return esc(v)
    v = float(v)
    escs = nombre in ("vcsb", "vcsg", "vflb", "vflg") or nombre.startswith(("vkcs", "vkfl"))
    if escs and v >= 9999:
        return r"$>$\num{9999}"
    if v == 0:
        return r"\num{0}"
    dig = 4 - int(math.floor(math.log10(abs(v)))) - 1
    v = round(v, dig)
    if abs(v) >= 1000:
        v = round(v)
    s = repr(int(v)) if float(v).is_integer() else repr(v)
    return r"\num{" + s + "}"

faltan = []
lineas = ["% Generado por scripts/exportar_datos.py -- no editar a mano",
          f"% Planilla: {xlsx.name}"]
for nombre, dn in sorted(wb.defined_names.items()):
    if not nombre.startswith("v"):
        continue
    hoja, ref = next(iter(dn.destinations))
    val = wb[hoja][ref.replace("$", "")].value
    if val is None:
        faltan.append(nombre)
    lineas.append(f"\\newcommand{{\\{nombre}}}{{{fmt(nombre, val)}}}")
filas = []
for i, rom in enumerate(("uno", "dos", "tres"), 1):
    if f"\\newcommand{{\\vkon{rom}}}{{1}}" in lineas:
        c = lambda k: f"\\vk{k}{rom}"
        filas.append(f"{i} & {c('x')} & {c('Kt')} & {c('Kts')} & {c('r')} & {c('Kf')} & {c('sa')} & {c('sm')} & {c('cs')} & {c('fl')} \\\\")
if not filas:
    filas.append("\\multicolumn{10}{c}{Sin secciones activas} \\\\")
lineas.append("\\newcommand{\\vFilasConc}{" + " ".join(filas) + "}")
hoy = datetime.date.today().strftime("%d/%m/%Y")
lineas.append(f"\\newcommand{{\\vFechaDatos}}{{{hoy}}}")
lineas.append(f"\\newcommand{{\\vArchivo}}{{{esc(xlsx.name)}}}")
(raiz / "memoria" / "datos.tex").write_text("\n".join(lineas) + "\n", encoding="utf-8")

# diagramas --------------------------------------------------------------
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
cx = wb["Cálculo x"]
col = lambda c: [cx[f"{c}{r}"].value for r in range(4, 205)]
x, My, Mz, M, fl, cs = col("A"), col("E"), col("F"), col("G"), col("AC"), col("R")
if None in x + M:
    faltan.append("Cálculo x")
else:
    plt.rcParams.update({"font.size": 9, "font.family": "serif"})
    fig, ax = plt.subplots(3, 1, figsize=(6.3, 6.6), sharex=True)
    ax[0].plot(x, [m / 1e3 for m in My], "k--", lw=1, label="$M_y$ (manchón)")
    ax[0].plot(x, [m / 1e3 for m in Mz], "k:", lw=1, label="$M_z$ (gravedad)")
    ax[0].plot(x, [m / 1e3 for m in M], "k-", lw=1.4, label=r"$M=\sqrt{M_y^2+M_z^2}$")
    ax[0].set_ylabel("M [N·m]"); ax[0].legend(frameon=False)
    ax[1].plot(x, fl, "k-", lw=1.4); ax[1].set_ylabel("Flecha total [mm]")
    ax[2].plot(x, [min(c, 50) for c in cs], "k-", lw=1.4); ax[2].axhline(1, color="k", lw=0.6)
    ax[2].set_ylabel("CS fatiga cuerpo"); ax[2].set_xlabel("x [mm]")
    for a in ax:
        a.grid(lw=0.3, color="0.8")
    fig.tight_layout()
    fig.savefig(raiz / "memoria" / "figuras" / "diagramas.pdf")

# diagrama S-N ------------------------------------------------------------
def val(n):
    hoja, ref = next(iter(wb.defined_names[n].destinations))
    return wb[hoja][ref.replace("$", "")].value
Va, Vc, Se, Su, srev = (val(k) for k in ("vVa", "vVc", "vSe", "vSu", "vsrev"))
if None in (Va, Vc, Se, Su, srev):
    faltan.append("diagrama S-N")
else:
    import numpy as np
    Nn = np.logspace(3, 8, 200)
    S = np.where(Nn <= 1e6, Va * Nn ** (-1 / Vc), Se)
    fig, ax = plt.subplots(figsize=(6.0, 3.6))
    ax.loglog(Nn, S, "k-", lw=1.4, label="S-N pieza")
    ax.axhline(srev, color="k", ls="--", lw=1, label=r"$\sigma_a$ equivalente sección crítica")
    ax.plot([1e3, 1e6], [0.9 * Su, Se], "ko", ms=3)
    ax.annotate(r"$0{,}9\,S_u$", (1e3, 0.9 * Su), textcoords="offset points", xytext=(6, 4))
    ax.annotate(r"$S_e$", (1e6, Se), textcoords="offset points", xytext=(4, 6))
    ax.set_xlabel("N [ciclos]"); ax.set_ylabel("S [MPa]"); ax.set_xlim(1e3, 1e8)
    ax.grid(which="both", lw=0.3, color="0.8"); ax.legend(frameon=False, loc="upper right")
    fig.tight_layout(); fig.savefig(raiz / "memoria" / "figuras" / "diagrama_sn.pdf")

if faltan:
    print("ATENCIÓN: sin valor calculado (abrir y guardar la planilla en Excel):", ", ".join(faltan))
    sys.exit(1)
print("datos.tex y diagramas.pdf actualizados")
