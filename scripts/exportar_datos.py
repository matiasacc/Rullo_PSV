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

CRUDOS = {"vpp", "vpm", "vub", "vswp", "vlado", "vswfl", "vkonuno", "vkondos", "vkontres"}  # enteros para \ifnum

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
    escs = nombre in ("vcsb", "vcsg", "vflb", "vflg") or nombre.startswith(("vkcs", "vkfl", "vscs", "vsfl"))
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
_hb, _rb = next(iter(wb.defined_names["vbfin"].destinations))
_bfin = wb[_hb][_rb.replace("$", "")].value
_ha, _ra = next(iter(wb.defined_names["vainicio"].destinations))
_aini = wb[_ha][_ra.replace("$", "")].value
_ok = _bfin is not None and _aini is not None and _bfin >= 0 and _aini >= 0
lineas.append("\\newcommand{\\vgeomok}{" + ("1" if _ok else "0") + "}")
hoy = datetime.date.today().strftime("%d/%m/%Y")
lineas.append(f"\\newcommand{{\\vFechaDatos}}{{{hoy}}}")
lineas.append(f"\\newcommand{{\\vArchivo}}{{{esc(xlsx.name)}}}")
(raiz / "memoria" / "datos.tex").write_text("\n".join(lineas) + "\n", encoding="utf-8")

# diagramas --------------------------------------------------------------
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle, Polygon, FancyArrowPatch
def val(n):
    hoja, ref = next(iter(wb.defined_names[n].destinations))
    return wb[hoja][ref.replace("$", "")].value
g = {k: val(k) for k in ("vL", "vxsI", "vxsD", "vxgcero", "vLg", "vainicio", "vB", "vD", "vDextmI", "vDextmD", "veg", "vlado", "vLmI", "vLmD", "vswp", "vcp", "vDp")}
cx = wb["Cálculo x"]
col = lambda c: [cx[f"{c}{r}"].value for r in range(4, 205)]
x, My, Mz, M, fl, cs = col("A"), col("E"), col("F"), col("G"), col("AC"), col("R")
plt.rcParams.update({"font.size": 9, "font.family": "serif"})

def marcas(a):
    """Zonas: muñones (gris), manchón (rojo), apoyos."""
    a.axvspan(0, g["vxsI"], color="0.88", lw=0); a.axvspan(g["vxsD"], g["vL"], color="0.88", lw=0)
    a.axvline(g["vxsI"], color="0.5", lw=0.6); a.axvline(g["vxsD"], color="0.5", lw=0.6)
    a.axvline(g["vainicio"], color="#C00000", lw=0.6, ls="--"); a.axvline(g["vainicio"] + g["vB"], color="#C00000", lw=0.6, ls="--")

if None in x + M or None in g.values():
    faltan.append("Cálculo x / geometría")
else:
    # esquema a escala ------------------------------------------------------
    L, xsI, xsD, D = g["vL"], g["vxsI"], g["vxsD"], g["vD"]
    fig, ax = plt.subplots(figsize=(6.6, 2.6))
    ext = max(L * 0.12, 60)
    hm_I, hm_D, eg = g["vDextmI"] / 2, g["vDextmD"] / 2, g["veg"]
    ax.add_patch(Rectangle((-min(g["vLmI"] - xsI, ext), -hm_I), xsI + min(g["vLmI"] - xsI, ext), 2 * hm_I, fc="0.85", ec="k", lw=0.6))
    ax.add_patch(Rectangle((xsD, -hm_D), (L - xsD) + min(g["vLmD"] - (L - xsD), ext), 2 * hm_D, fc="0.85", ec="k", lw=0.6))
    ax.add_patch(Rectangle((xsI, -D / 2), xsD - xsI, D, fc="0.7", ec="k", lw=0.8))
    x0g, Lg = g["vxgcero"], g["vLg"]
    for sgn in (1, -1):
        ax.add_patch(Rectangle((x0g, sgn * D / 2 if sgn > 0 else -D / 2 - eg), Lg, eg, fc="0.35", ec="k", lw=0.4))
    a0, B = g["vainicio"], g["vB"]
    ax.add_patch(Rectangle((a0, D / 2 + eg), B, D * 0.18, fc="none", ec="#C00000", hatch="////", lw=0.8))
    ax.annotate("", xy=(a0 + B, D / 2 + eg + D * 0.3), xytext=(a0, D / 2 + eg + D * 0.3),
                arrowprops=dict(arrowstyle="->", color="#C00000", lw=0.8))
    ax.text(a0 + B / 2, D / 2 + eg + D * 0.36, "manchón (tope a la derecha)", ha="center", color="#C00000", fontsize=8)
    for xa in (0, L):
        h = (hm_I if xa == 0 else hm_D)
        ax.add_patch(Polygon([[xa, -h], [xa - L * 0.02, -h - D * 0.25], [xa + L * 0.02, -h - D * 0.25]], closed=True, fc="w", ec="k", lw=0.8))
    ax.text(0, -hm_I - D * 0.42, "apoyo izq.", ha="center", fontsize=8); ax.text(L, -hm_D - D * 0.42, "apoyo der.", ha="center", fontsize=8)
    if g["vswp"] == 1:
        xp_, rp = L + g["vcp"], g["vDp"] / 2
        ax.add_patch(Rectangle((xp_ - L * 0.012, -rp), L * 0.024, 2 * rp, fc="w", ec="k", lw=0.8))
        ax.text(xp_, rp + D * 0.06, "polea", ha="center", fontsize=8)
        ax.annotate("", xy=(xp_ - L * 0.06, -rp * 0.2), xytext=(xp_ - L * 0.06, rp * 0.2),
                    arrowprops=dict(arrowstyle="<-", lw=0.6))
        ext = max(ext, g["vcp"] + L * 0.04)
    else:
        xm = L + min(g["vLmD"] - (L - xsD), ext)
        ax.annotate("T (acople)", xy=(xm, 0), xytext=(xm + L * 0.08, D * 0.45), ha="center", fontsize=8,
                    arrowprops=dict(arrowstyle="->", lw=0.8))
    for xv, t in ((xsI, "sold."), (xsD, "sold.")):
        ax.text(xv, -D / 2 - eg - D * 0.12, t, ha="center", fontsize=7, color="0.3")
    ax.annotate("", xy=(L, -D * 1.05), xytext=(0, -D * 1.05), arrowprops=dict(arrowstyle="<->", lw=0.6))
    ax.text(L / 2, -D * 1.0, f"L = {L:.0f} mm", ha="center", fontsize=8)
    ax.text(-ext, D * 0.9, "muñón izq.", fontsize=8); ax.text(L - ext * 0.2, D * 0.9, "muñón der.", fontsize=8, ha="center")
    ax.set_xlim(-ext * 1.6, L + ext * 1.6); ax.set_ylim(-D * 1.15, D * 1.1); ax.set_aspect("auto"); ax.axis("off")
    fig.tight_layout(); fig.savefig(raiz / "memoria" / "figuras" / "esquema.pdf")

    fig, ax = plt.subplots(3, 1, figsize=(6.3, 6.6), sharex=True)
    ax[0].plot(x, [m / 1e3 for m in My], "k--", lw=1, label="$M_y$ (manchón)")
    ax[0].plot(x, [m / 1e3 for m in Mz], "k:", lw=1, label="$M_z$ (gravedad)")
    ax[0].plot(x, [m / 1e3 for m in M], "k-", lw=1.4, label=r"$M=\sqrt{M_y^2+M_z^2}$")
    ax[0].set_ylabel("M [N·m]"); ax[0].legend(frameon=False, fontsize=8)
    ax[1].plot(x, fl, "k-", lw=1.4); ax[1].set_ylabel("Flecha total [mm]")
    ax[2].plot(x, [min(c, 50) for c in cs], "k-", lw=1.4); ax[2].axhline(1, color="k", lw=0.6)
    ax[2].set_ylabel("CS fatiga eje"); ax[2].set_xlabel("x [mm]  (gris: muñones; rojo: manchón)")
    for a in ax:
        marcas(a); a.grid(lw=0.3, color="0.8"); a.set_xlim(0, L)
    fig.tight_layout()
    fig.savefig(raiz / "memoria" / "figuras" / "diagramas.pdf")

# diagrama S-N ------------------------------------------------------------
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
