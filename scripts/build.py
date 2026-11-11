from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.drawing.image import Image
from openpyxl.chart import ScatterChart, Reference, Series
from openpyxl.comments import Comment

F = "Arial"
fb = Font(name=F, size=10); fbold = Font(name=F, size=10, bold=True)
fin = Font(name=F, size=10, color="0000FF"); fh = Font(name=F, size=12, bold=True)
fgreen = Font(name=F, size=10, color="008000"); fsmall = Font(name=F, size=9, italic=True, color="555555")
yel = PatternFill("solid", fgColor="FFFF00"); grey = PatternFill("solid", fgColor="E7E6E6")
thin = Side(style="thin", color="999999"); box = Border(left=thin, right=thin, top=thin, bottom=thin)

wb = Workbook()

# ------------------------------------------------------------------ DATOS
ds = wb.active; ds.title = "Datos"
ds.column_dimensions["A"].width = 46; ds.column_dimensions["B"].width = 16
ds.column_dimensions["C"].width = 10; ds.column_dimensions["D"].width = 70
ds["A1"] = "RULLO PSV - DATOS DE ENTRADA"; ds["A1"].font = fh
ds["A2"] = ("Editar solo las celdas amarillas (texto azul). Los valores cargados son un EJEMPLO, "
            "no son datos del rullo real."); ds["A2"].font = fsmall
N = {}  # name -> absolute ref
r = 4
def sec(t):
    global r
    ds.cell(r, 1, t).font = fbold
    for c in range(1, 5): ds.cell(r, c).fill = grey
    r += 1
def inp(name, label, val, unit, note=""):
    global r
    ds.cell(r, 1, label).font = fb
    c = ds.cell(r, 2, val); c.font = fin; c.fill = yel; c.border = box
    ds.cell(r, 3, unit).font = fb; ds.cell(r, 4, note).font = fsmall
    N[name] = f"Datos!$B${r}"; r += 1
def der(name, label, formula, unit, note="", fmt="0.000"):
    global r
    ds.cell(r, 1, label).font = fb
    c = ds.cell(r, 2, formula); c.font = fb; c.number_format = fmt; c.border = box
    ds.cell(r, 3, unit).font = fb; ds.cell(r, 4, note).font = fsmall
    N[name] = f"Datos!$B${r}"; r += 1

sec("Geometría del tubo")
inp("D", "Diámetro exterior D", 100, "mm")
inp("t", "Espesor de pared t", 10, "mm")
inp("L", "Distancia entre centros de rodamientos L", 1200, "mm", "Modelo: viga simplemente apoyada (ver hoja Fórmulas)")
sec("Engomado (recubrimiento del tubo)")
inp("eg", "Espesor del engomado", 10, "mm", "* Solo suma peso y cambia el diámetro de apoyo del manchón; su rigidez se desprecia (E goma << E acero)")
inp("rhog", "Densidad de la goma", 1.3e-6, "kg/mm³", "* Valor típico de caucho; reemplazar por el dato del compuesto")
sec("Manchón / correas (plano y: dirección de la tensión)")
inp("nb", "Cantidad de correas", 20, "-")
inp("bw", "Ancho de cada correa", 20, "mm", "Ancho del manchón B = cantidad × ancho")
inp("mb", "Masa de una correa", 0.15, "kg", "Pesada")
inp("Lp", "Desarrollo de la correa", 1200, "mm")
inp("fr", "Frecuencia medida (por correa)", 40, "Hz", "Ramal libre vibrando")
inp("Ltin", "Largo libre del ramal Lt (0 = calcular)", 0, "mm",
    "* Si es 0: Lt = (desarrollo - π·D)/2, ramales paralelos y rullos de igual diámetro (línea primitiva ≈ D exterior)")
inp("a", "Inicio del manchón, desde el apoyo izquierdo (a)", 400, "mm", "El fin queda en a + B")
sec("Plano z (gravedad)")
inp("pp", "Incluir peso propio del tubo (1 = sí, 0 = no)", 1, "-", "Carga uniforme w = ρ·g·A en todo L")
inp("pm", "Incluir peso del manchón sobre este rullo (1 = sí, 0 = no)", 1, "-",
    "* Supuesto: la mitad del peso del manchón apoya en cada rullo, repartida en B")
sec("Torque")
inp("P", "Potencia transmitida P", 2.0, "kW")
inp("n", "Velocidad del rullo n", 150, "rpm", "También se usa para pasar ciclos a horas (1 ciclo por vuelta)")
inp("lado", "Lado por donde entra el torque (1 = izquierdo, 0 = derecho)", 1, "-",
    "* Supuesto: el torque entra por un extremo y las correas lo toman de forma pareja en el tramo cargado")
sec("Balanceo (fuera del compendio)")
inp("G", "Grado de calidad de balanceo G (ISO 21940-11)", 6.3, "mm/s", "* Desbalanceo residual admisible: e·ω = G")
sec("Material")
inp("Su", "Resistencia a rotura Su", 630, "MPa", "Ejemplo: AISI 1045 estirado en frío")
inp("Sy", "Tensión de fluencia Sy", 530, "MPa")
inp("E", "Módulo de elasticidad E", 207000, "MPa")
inp("rho", "Densidad ρ", 7.85e-6, "kg/mm³")
sec("Fatiga")
inp("term", "Terminación superficial", "Maquinado o estirado en frío", "-", "Lista: ec. (26) del compendio")
term_row = r - 1
inp("f1", "Otros factores f1 (temperatura, confiabilidad, etc.)", 1.0, "-", "Ec. (20). 1 = sin reducción adicional")
inp("swfl", "¿Aplicar Kf en la verificación a fluencia? (1 = sí, 0 = no)", 1, "-",
    "* A DEFINIR entre ambos: ec. (3) dice que en dúctil estático no se usa Kt; ec. (77) toma σm+σa. 1 es lo conservativo")

sec("Secciones con concentrador de tensiones (hasta 3)")
hdr = ["", "Sección 1", "Sección 2", "Sección 3"]
ds.cell(r, 1, "Dato").font = fbold
for j in range(1, 4): ds.cell(r, 1 + j, hdr[j]).font = fbold
ds.cell(r, 5, "Nota").font = fbold
r += 1
conc_rows = {}
conc_def = [("kx", "Posición x desde el apoyo izquierdo", [60, 1140, 600], "mm", "Ej.: escalón del asiento de rodamiento"),
            ("kon", "Activa (1 = sí, 0 = no)", [1, 1, 0], "-", ""),
            ("kKt", "Kt flexión (Anexo 2)", [1.7, 1.7, 1.0], "-", "Leer del gráfico con r/d y D/d"),
            ("kKts", "Kts torsión (Anexo 2)", [1.4, 1.4, 1.0], "-", ""),
            ("kr", "Radio de entalla r", [2.0, 2.0, 2.0], "mm", "Para la sensibilidad q (Neuber, ec. 9)")]
ds.column_dimensions["E"].width = 40
for key, lab, vals, unit, note in conc_def:
    ds.cell(r, 1, f"{lab} [{unit}]").font = fb
    for j, v in enumerate(vals):
        c = ds.cell(r, 2 + j, v); c.font = fin; c.fill = yel; c.border = box
    ds.cell(r, 5, note).font = fsmall
    conc_rows[key] = r; r += 1

r += 1
sec("Propiedades derivadas (no editar)")
der("d", "Diámetro interior d = D - 2t", f"={N['D']}-2*{N['t']}", "mm", fmt="0.0")
der("A", "Área A = π/4·(D² - d²)", f"=PI()/4*({N['D']}^2-B{r-1}^2)", "mm²", fmt="0.0")
N_d = N["d"]
ds.cell(r - 1, 2).value = f"=PI()/4*({N['D']}^2-{N_d}^2)"
der("I", "Momento de inercia I = π/64·(D⁴ - d⁴)", f"=PI()/64*({N['D']}^4-{N_d}^4)", "mm⁴", fmt="#,##0")
der("J", "Momento polar J = 2I", f"=2*{N['I']}", "mm⁴", fmt="#,##0")
der("c", "Fibra extrema c = D/2", f"={N['D']}/2", "mm", fmt="0.0")
der("Q", "Momento estático Q = ⅔·(R³ - r³)", f"=2/3*(({N['D']}/2)^3-({N_d}/2)^3)", "mm³", "Corte transversal en el eje neutro", fmt="#,##0")
der("bs", "Ancho resistente al corte b = 2t", f"=2*{N['t']}", "mm", fmt="0.0")
der("Ag", "Área del engomado Ag = π/4·((D+2e)² - D²)", f"=PI()/4*(({N['D']}+2*{N['eg']})^2-{N['D']}^2)", "mm²", fmt="0.0")
der("w", "Peso propio w = g·(ρ·A + ρg·Ag) (si se incluye)", f"={N['pp']}*9.81*({N['rho']}*{N['A']}+{N['rhog']}*{N['Ag']})", "N/mm",
    "Tubo + engomado", fmt="0.0000")
der("cl", "Ancho del manchón B = cantidad × ancho", f"={N['nb']}*{N['bw']}", "mm", fmt="0.0")
der("b", "Distancia del fin del manchón al apoyo derecho b = L - a - B", f"={N['L']}-{N['a']}-{N['cl']}", "mm",
    "Debe ser ≥ 0 (si da negativo el manchón no entra entre apoyos)", fmt="0.0")
der("mlin", "Masa por unidad de largo m = masa/desarrollo", f"={N['mb']}/({N['Lp']}/1000)", "kg/m", fmt="0.0000")
der("Lt", "Largo libre del ramal Lt", f"=IF({N['Ltin']}>0,{N['Ltin']},({N['Lp']}-PI()*({N['D']}+2*{N['eg']}))/2)", "mm", fmt="0.0")
der("Tb", "Tensión por correa T = 4·m·Lt²·f²", f"=4*{N['mlin']}*({N['Lt']}/1000)^2*{N['fr']}^2", "N",
    "* Cuerda vibrante (tensiómetro sónico). No está en el compendio: a validar", fmt="0.0")
der("Ttot", "Tensión total del manchón = T × cantidad", f"={N['Tb']}*{N['nb']}", "N", fmt="#,##0")
der("W", "Fuerza sobre el rullo (plano y) F = 2·T total", f"=2*{N['Ttot']}", "N", "Ramales paralelos (180° abrazados): F = 2T", fmt="#,##0")
der("qy", "Carga distribuida qy = F/B", f"={N['W']}/{N['cl']}", "N/mm", fmt="0.00")
der("Wz", "Peso del manchón sobre el rullo (plano z)", f"={N['pm']}*{N['nb']}*{N['mb']}*9.81/2", "N", fmt="0.0")
der("qz", "Carga distribuida qz = peso manchón/B", f"={N['Wz']}/{N['cl']}", "N/mm", fmt="0.0000")
der("T0", "Torque T = P/ω", f"={N['P']}*1000/(2*PI()*{N['n']}/60)*1000", "N·mm", fmt="#,##0")
der("R1y", "Reacción R1y (plano y)", f"={N['W']}*({N['b']}+{N['cl']}/2)/{N['L']}", "N", "Equilibrio de momentos respecto de R2", fmt="#,##0")
der("R2y", "Reacción R2y (plano y)", f"={N['W']}*({N['a']}+{N['cl']}/2)/{N['L']}", "N", fmt="#,##0")
der("R1z", "Reacción R1z (plano z)", f"={N['Wz']}*({N['b']}+{N['cl']}/2)/{N['L']}+{N['w']}*{N['L']}/2", "N", fmt="#,##0")
der("R2z", "Reacción R2z (plano z)", f"={N['Wz']}*({N['a']}+{N['cl']}/2)/{N['L']}+{N['w']}*{N['L']}/2", "N", fmt="#,##0")
der("R1", "Reacción resultante en rodamiento izquierdo R1 = √(R1y² + R1z²)", f"=SQRT({N['R1y']}^2+{N['R1z']}^2)", "N", fmt="#,##0")
der("R2", "Reacción resultante en rodamiento derecho R2", f"=SQRT({N['R2y']}^2+{N['R2z']}^2)", "N", fmt="#,##0")
der("chk", "Control: ΣRy - F + ΣRz - pesos (debe dar 0)",
    f"={N['R1y']}+{N['R2y']}-{N['W']}+{N['R1z']}+{N['R2z']}-{N['Wz']}-{N['w']}*{N['L']}", "N", fmt="0.00")
der("fS_a", "Factor de superficie: a", f"=INDEX(Fórmulas!$B$5:$B$8,MATCH({N['term']},Fórmulas!$A$5:$A$8,0))", "-", "Ec. (26), unidades MPa")
der("fS_b", "Factor de superficie: b", f"=INDEX(Fórmulas!$C$5:$C$8,MATCH({N['term']},Fórmulas!$A$5:$A$8,0))", "-")
der("fS", "fS = a·Su^b", f"={N['fS_a']}*{N['Su']}^{N['fS_b']}", "-", "Ec. (25)")
der("fT", "fT, factor de tamaño (eje rotativo, d = D exterior*)",
    f'=IF({N["D"]}<=51,1.24*{N["D"]}^-0.107,IF({N["D"]}<=254,1.51*{N["D"]}^-0.157,1.51*254^-0.157))', "-",
    "Ec. (21). * Para tubo se tomó D exterior. Si D > 254 mm se limita al valor de 254 mm")
der("fC", "fC, factor de carga", 1, "-", "Ec. (24): fC = 1 con carga combinada (la torsión entra por Von Mises)")
ds.cell(r - 1, 2).font = fb
der("Sep", "Se' = mín(0,5·Su ; 690)", f"=MIN(0.5*{N['Su']},690)", "MPa", "Ec. (17)", fmt="0.0")
der("Se", "Se = fS·fT·fC·f1·Se'", f"={N['fS']}*{N['fT']}*{N['fC']}*{N['f1']}*{N['Sep']}", "MPa", "Ec. (20)", fmt="0.0")
der("Va", "Vida limitada: a = (0,9·Su)²/Se", f"=(0.9*{N['Su']})^2/{N['Se']}", "MPa", "Ec. (31)", fmt="0.0")
der("Vb", "Vida limitada: b = -⅓·log(0,9·Su/Se)", f"=-1/3*LOG10(0.9*{N['Su']}/{N['Se']})", "-", "Ec. (31)", fmt="0.0000")
der("Vc", "Vida limitada: c = -1/b", f"=-1/{N['Vb']}", "-", "Ec. (31)", fmt="0.000")
der("sa_f", "√α Neuber flexión", f"=1.24-0.00225*{N['Su']}+1.6E-6*{N['Su']}^2-4.1E-10*{N['Su']}^3", "√mm", "Ec. (10)-a (Su en MPa)", fmt="0.0000")
der("sa_t", "√α Neuber torsión", f"=0.958-0.001834*{N['Su']}+1.43E-6*{N['Su']}^2-4.1E-10*{N['Su']}^3", "√mm",
    "Ec. (10)-b. * Coeficientes leídos del texto extraído del PDF: verificar contra el compendio impreso", fmt="0.0000")
der("mu", "Masa por unidad de largo que gira (tubo + goma)", f"=({N['rho']}*{N['A']}+{N['rhog']}*{N['Ag']})*1000", "kg/m", fmt="0.00")
der("Mrot", "Masa que gira M", f"={N['mu']}*{N['L']}/1000", "kg", fmt="0.0")
der("nc", "1ª velocidad crítica de flexión n_c = (30/π)·(π/L)²·√(EI/μ)", 
    f"=30/PI()*(PI()/({N['L']}/1000))^2*SQRT({N['E']}*1E6*{N['I']}*1E-12/{N['mu']})", "rpm",
    "* Viga uniforme simplemente apoyada (teoría de vigas, fuera del compendio). No incluye la masa del manchón", fmt="#,##0")
der("rnc", "Relación n / n_c", f"={N['n']}/{N['nc']}", "-", "Lejos de 1 = sin resonancia de flexión", fmt="0.000")
der("Fu", "Fuerza centrífuga por desbalanceo F_u = M·G·ω", f"={N['Mrot']}*{N['G']}/1000*2*PI()*{N['n']}/60", "N",
    "* Gira con el eje: para el material es una carga fija (tensión media), no alternante", fmt="0.0")
der("rFu", "Relación F_u / F manchón", f"={N['Fu']}/{N['W']}", "-", fmt="0.0000")

dv = DataValidation(type="list", formula1='"Rectificado,Maquinado o estirado en frío,Laminado en caliente,Forjado"', allow_blank=False)
ds.add_data_validation(dv); dv.add(f"B{term_row}")
img = Image("dcl.png"); img.width, img.height = 620, 434
ds.add_image(img, "G3")
ds["G26"] = "Diagrama de cuerpo libre (modelo simplemente apoyado). x se mide desde el apoyo izquierdo."; ds["G26"].font = fsmall

# helper formula pieces (x given as a cell ref) -----------------------------
L, a, b, w, cl, T0 = (N[k] for k in ("L", "a", "b", "w", "cl", "T0"))
def Vg(x, R, qq, ww):
    return f"({R}-{ww}*{x}-{qq}*MAX(0,MIN({x},{L}-{b})-{a}))"
def Mg(x, R, qq, ww):
    return (f"({R}*{x}-{ww}*{x}^2/2-{qq}*IF({x}<={a},0,IF({x}<={L}-{b},({x}-{a})^2/2,{cl}*({x}-{a}-{cl}/2))))")
def Vyf(x): return Vg(x, N['R1y'], N['qy'], "0")
def Vzf(x): return Vg(x, N['R1z'], N['qz'], w)
def Myf(x): return Mg(x, N['R1y'], N['qy'], "0")
def Mzf(x): return Mg(x, N['R1z'], N['qz'], w)
def Vf(x): return f"SQRT({Vyf(x)}^2+{Vzf(x)}^2)"
def Mf(x): return f"SQRT({Myf(x)}^2+{Mzf(x)}^2)"
def Tside(x, aa, bb):
    return f"IF({x}<={aa},{T0},IF({x}>={L}-{bb},0,{T0}*({L}-{bb}-{x})/{cl}))"
def Tf(x):
    return f"IF({N['lado']}=1,{Tside(x, a, b)},{Tside(f'({L}-{x})', b, a)})"

# ------------------------------------------------------------------ CALCULO X
from openpyxl.utils import get_column_letter, column_index_from_string
cs = wb.create_sheet("Cálculo x")
heads = ["x [mm]", "Vy [N]", "Vz [N]", "V = √(Vy²+Vz²) [N]", "My [N·mm]", "Mz [N·mm]", "M = √(My²+Mz²) [N·mm]", "T [N·mm]",
         "σ0 = M·c/I [MPa]", "τT = T·c/J [MPa]", "τV = V·Q/(I·b) [MPa]",
         "A: σ'a [MPa]", "A: σ'm [MPa]", "A: CS Goodman", "B: σ'a [MPa]", "B: σ'm [MPa]", "B: CS Goodman",
         "CS fatiga (mín A,B)", "σ'máx [MPa]", "CS fluencia",
         "My/EI [1/mm]", "∫ My/EI", "∫∫ My/EI", "Flecha y [mm]",
         "Mz/EI [1/mm]", "∫ Mz/EI", "∫∫ Mz/EI", "Flecha z [mm]", "Flecha total [mm]"]
cs["A1"] = "CUERPO DEL TUBO (Kf = 1). Concentradores en hoja 'Concentradores'. A = fibra extrema, B = eje neutro. y: tensión del manchón, z: gravedad"; cs["A1"].font = fbold
for j, h in enumerate(heads):
    c = cs.cell(3, 1 + j, h); c.font = fbold; c.fill = grey; c.alignment = Alignment(wrap_text=True, horizontal="center")
    cs.column_dimensions[c.column_letter].width = 13
cs.row_dimensions[3].height = 42
NSEG = 200; r0 = 4; rN = r0 + NSEG
Se, Su, Sy, I, J, c_, Q, bs, E = (N[k] for k in ("Se", "Su", "Sy", "I", "J", "c", "Q", "bs", "E"))
fmts = ["0.0", "#,##0", "#,##0", "#,##0", "#,##0", "#,##0", "#,##0", "#,##0"] + ["0.00"] * 12 + ["0.00E+00"] * 3 + ["0.0000"] + ["0.00E+00"] * 3 + ["0.0000", "0.0000"]
for i in range(NSEG + 1):
    rr = r0 + i; x = f"A{rr}"
    F_ = {}
    F_["A"] = f"={L}*{i}/{NSEG}"
    F_["B"] = "=" + Vyf(x); F_["C"] = "=" + Vzf(x); F_["D"] = f"=SQRT(B{rr}^2+C{rr}^2)"
    F_["E"] = "=" + Myf(x); F_["F"] = "=" + Mzf(x); F_["G"] = f"=SQRT(E{rr}^2+F{rr}^2)"
    F_["H"] = "=" + Tf(x)
    F_["I"] = f"=G{rr}*{c_}/{I}"
    F_["J"] = f"=H{rr}*{c_}/{J}"
    F_["K"] = f"=D{rr}*{Q}/({I}*{bs})"
    F_["L"] = f"=ABS(I{rr})"
    F_["M"] = f"=SQRT(3)*J{rr}"
    F_["N"] = f"=IFERROR(1/(L{rr}/{Se}+M{rr}/{Su}),9999)"
    F_["O"] = f"=SQRT(3)*K{rr}"
    F_["P"] = f"=SQRT(3)*J{rr}"
    F_["Q"] = f"=IFERROR(1/(O{rr}/{Se}+P{rr}/{Su}),9999)"
    F_["R"] = f"=MIN(N{rr},Q{rr})"
    F_["S"] = f"=MAX(SQRT(I{rr}^2+3*J{rr}^2),SQRT(3)*(K{rr}+J{rr}))"
    F_["T"] = f"=IFERROR({Sy}/S{rr},9999)"
    for (mc, kc, i1, i2, fl) in (("E", "U", "V", "W", "X"), ("F", "Y", "Z", "AA", "AB")):
        F_[kc] = f"={mc}{rr}/({E}*{I})"
        if i == 0:
            F_[i1] = 0; F_[i2] = 0
        else:
            F_[i1] = f"={i1}{rr-1}+({kc}{rr-1}+{kc}{rr})/2*(A{rr}-A{rr-1})"
            F_[i2] = f"={i2}{rr-1}+({i1}{rr-1}+{i1}{rr})/2*(A{rr}-A{rr-1})"
        F_[fl] = f"=-({i2}{rr}-A{rr}/{L}*${i2}${rN})"
    F_["AC"] = f"=SQRT(X{rr}^2+AB{rr}^2)"
    for j in range(len(heads)):
        col = get_column_letter(j + 1)
        cell = cs[f"{col}{rr}"]; cell.value = F_[col]; cell.font = fb; cell.number_format = fmts[j]
cs.freeze_panes = "B4"
rng = lambda col: f"'Cálculo x'!${col}${r0}:${col}${rN}"

# ------------------------------------------------------------------ CONCENTRADORES
cc = wb.create_sheet("Concentradores")
cc.column_dimensions["A"].width = 44
for col in "BCD": cc.column_dimensions[col].width = 14
cc.column_dimensions["E"].width = 60
cc["A1"] = "SECCIONES CON CONCENTRADOR (datos en hoja Datos)"; cc["A1"].font = fh
cc["A3"] = "Magnitud"; cc["E3"] = "Ecuación / nota"
for j in range(3): cc.cell(3, 2 + j, f"Sección {j+1}")
for col in "ABCDE": cc[f"{col}3"].font = fbold; cc[f"{col}3"].fill = grey
rows = [
 ("x [mm]", lambda k: f"=Datos!{k}{conc_rows['kx']}", "Dato"),
 ("Activa", lambda k: f"=Datos!{k}{conc_rows['kon']}", "Dato"),
 ("V [N]", None, "Equilibrio (DCL)"),
 ("M [N·mm]", None, "Equilibrio (DCL)"),
 ("T [N·mm]", None, "* Distribución de T supuesta (ver Datos)"),
 ("σ0 = M·c/I [MPa]", None, "Tensión nominal de flexión"),
 ("τT = T·c/J [MPa]", None, "Tensión nominal de torsión"),
 ("τV = V·Q/(I·b) [MPa]", None, "Corte transversal en el eje neutro"),
 ("q flexión = 1/(1+√α/√r)", None, "Ec. (9) Neuber"),
 ("q torsión", None, "Ec. (9) Neuber"),
 ("Kf = 1 + q·(Kt - 1)", None, "Ec. (6)"),
 ("Kfs = 1 + qs·(Kts - 1)", None, "Ec. (6)"),
 ("A: σa = Kf·|σ0| [MPa]", None, "Tabla (14): Kf solo sobre la alterna (dúctil). Eje rotativo: flexión alternante, σm = 0"),
 ("A: τm = τT [MPa]", None, "Tabla (14): media sin Kf (dúctil). * Torque constante: τa = 0"),
 ("A: σ'a [MPa]", None, "Ec. (72)"),
 ("A: σ'm [MPa]", None, "Ec. (72)"),
 ("B: τa = Kfs·τV [MPa]", None, "* Kfs aplicado al corte transversal (conservativo)"),
 ("B: σ'a [MPa]", None, "Ec. (73)"),
 ("B: σ'm [MPa]", None, "Ec. (73)"),
 ("A: CS Goodman", None, "Ec. (44)-b"),
 ("A: CS Soderberg", None, "Ec. (44)-a"),
 ("A: CS Gerber", None, "Ec. (44)-c"),
 ("B: CS Goodman", None, "Ec. (44)-b"),
 ("CS fatiga Goodman (mín A, B)", None, ""),
 ("σ'máx A (fluencia) [MPa]", None, "Ec. (77): tensor medio + alternante, Von Mises"),
 ("σ'máx B (fluencia) [MPa]", None, "Ec. (77)"),
 ("CS fluencia", None, "Ec. (77)"),
 ("Vida: SN = σ'a/(1 - σ'm/Su) [MPa]", None, "Ec. (48)-b con CS = 1 (solo si CS Goodman < 1)"),
 ("Vida N [ciclos]", None, "Ec. (32): N = (a/SN)^c. Válida si Se < SN ≤ 0,9·Su"),
 ("Vida [horas]", None, "N / (n·60)"),
]
for i, (lab, f, note) in enumerate(rows):
    rr = 4 + i
    cc.cell(rr, 1, lab).font = fb; cc.cell(rr, 5, note).font = fsmall
R = {lab: 4 + i for i, (lab, f, note) in enumerate(rows)}
for j, k in enumerate("BCD"):
    col = k
    g = lambda lab: f"{col}{R[lab]}"
    x = g("x [mm]")
    kt = f"Datos!{k}{conc_rows['kKt']}"; kts = f"Datos!{k}{conc_rows['kKts']}"; kr = f"Datos!{k}{conc_rows['kr']}"
    F_ = {
     "x [mm]": f"=Datos!{k}{conc_rows['kx']}",
     "Activa": f"=Datos!{k}{conc_rows['kon']}",
     "V [N]": "=" + Vf(x),
     "M [N·mm]": "=" + Mf(x),
     "T [N·mm]": "=" + Tf(x),
     "σ0 = M·c/I [MPa]": f"={g('M [N·mm]')}*{c_}/{I}",
     "τT = T·c/J [MPa]": f"={g('T [N·mm]')}*{c_}/{J}",
     "τV = V·Q/(I·b) [MPa]": f"=ABS({g('V [N]')})*{Q}/({I}*{bs})",
     "q flexión = 1/(1+√α/√r)": f"=1/(1+{N['sa_f']}/SQRT({kr}))",
     "q torsión": f"=1/(1+{N['sa_t']}/SQRT({kr}))",
     "Kf = 1 + q·(Kt - 1)": f"=1+{g('q flexión = 1/(1+√α/√r)')}*({kt}-1)",
     "Kfs = 1 + qs·(Kts - 1)": f"=1+{g('q torsión')}*({kts}-1)",
     "A: σa = Kf·|σ0| [MPa]": f"={g('Kf = 1 + q·(Kt - 1)')}*ABS({g('σ0 = M·c/I [MPa]')})",
     "A: τm = τT [MPa]": f"={g('τT = T·c/J [MPa]')}",
     "A: σ'a [MPa]": f"={g('A: σa = Kf·|σ0| [MPa]')}",
     "A: σ'm [MPa]": f"=SQRT(3)*{g('A: τm = τT [MPa]')}",
     "B: τa = Kfs·τV [MPa]": f"={g('Kfs = 1 + qs·(Kts - 1)')}*{g('τV = V·Q/(I·b) [MPa]')}",
     "B: σ'a [MPa]": f"=SQRT(3)*{g('B: τa = Kfs·τV [MPa]')}",
     "B: σ'm [MPa]": f"=SQRT(3)*{g('A: τm = τT [MPa]')}",
    }
    ea = f"({g(chr(65)+': σ'+chr(39)+'a [MPa]')}/{Se})"
    sA_a = g("A: σ'a [MPa]"); sA_m = g("A: σ'm [MPa]"); sB_a = g("B: σ'a [MPa]"); sB_m = g("B: σ'm [MPa]")
    F_["A: CS Goodman"] = f"=IFERROR(1/({sA_a}/{Se}+{sA_m}/{Su}),9999)"
    F_["A: CS Soderberg"] = f"=IFERROR(1/({sA_a}/{Se}+{sA_m}/{Sy}),9999)"
    F_["A: CS Gerber"] = (f"=IFERROR(IF({sA_m}=0,{Se}/{sA_a},(-({sA_a}/{Se})+SQRT(({sA_a}/{Se})^2+4*({sA_m}/{Su})^2))"
                          f"/(2*({sA_m}/{Su})^2)),9999)")
    F_["B: CS Goodman"] = f"=IFERROR(1/({sB_a}/{Se}+{sB_m}/{Su}),9999)"
    F_["CS fatiga Goodman (mín A, B)"] = f"=MIN({g('A: CS Goodman')},{g('B: CS Goodman')})"
    kfl = f"IF({N['swfl']}=1,{g('Kf = 1 + q·(Kt - 1)')},1)"; kfsl = f"IF({N['swfl']}=1,{g('Kfs = 1 + qs·(Kts - 1)')},1)"
    F_["σ'máx A (fluencia) [MPa]"] = f"=SQRT(({kfl}*{g('σ0 = M·c/I [MPa]')})^2+3*{g('τT = T·c/J [MPa]')}^2)"
    F_["σ'máx B (fluencia) [MPa]"] = f"=SQRT(3)*({kfsl}*{g('τV = V·Q/(I·b) [MPa]')}+{g('τT = T·c/J [MPa]')})"
    F_["CS fluencia"] = f"=IFERROR({Sy}/MAX({g('σ' + chr(39) + 'máx A (fluencia) [MPa]')},{g('σ' + chr(39) + 'máx B (fluencia) [MPa]')}),9999)"
    sn = g("Vida: SN = σ'a/(1 - σ'm/Su) [MPa]")
    F_["Vida: SN = σ'a/(1 - σ'm/Su) [MPa]"] = (f"=IF({g('CS fatiga Goodman (mín A, B)')}>=1,\"Vida infinita\","
        f"IF({g('A: CS Goodman')}<={g('B: CS Goodman')},{sA_a}/(1-{sA_m}/{Su}),{sB_a}/(1-{sB_m}/{Su})))")
    F_["Vida N [ciclos]"] = (f"=IF(ISNUMBER({sn}),IF({sn}>0.9*{Su},\"< 10^3: fuera de rango\",({N['Va']}/{sn})^{N['Vc']}),\"Vida infinita\")")
    vn = g("Vida N [ciclos]")
    F_["Vida [horas]"] = f"=IF(ISNUMBER({vn}),{vn}/({N['n']}*60),\"-\")"
    for lab, fml in F_.items():
        cell = cc[g(lab)]; cell.value = fml; cell.font = fb; cell.number_format = "0.00"
    for lab in ("V [N]", "M [N·mm]", "T [N·mm]", "Vida N [ciclos]"):
        cc[g(lab)].number_format = "#,##0"
    for lab in ("q flexión = 1/(1+√α/√r)", "q torsión", "Kf = 1 + q·(Kt - 1)", "Kfs = 1 + qs·(Kts - 1)"):
        cc[g(lab)].number_format = "0.000"
    cc[g("x [mm]")].font = fgreen; cc[g("Activa")].font = fgreen

# ------------------------------------------------------------------ RESULTADOS
rs = wb.create_sheet("Resultados", 1)
rs.column_dimensions["A"].width = 50; rs.column_dimensions["B"].width = 18
rs.column_dimensions["C"].width = 12; rs.column_dimensions["D"].width = 60
rs["A1"] = "RULLO PSV - RESUMEN ESTÁTICO Y FATIGA"; rs["A1"].font = fh
rs["A2"] = "Modelo: viga simplemente apoyada en los rodamientos, eje rotativo, carga de correas de dirección fija."; rs["A2"].font = fsmall
rr = 4
def rsec(t):
    global rr
    rs.cell(rr, 1, t).font = fbold
    for c in range(1, 5): rs.cell(rr, c).fill = grey
    rr += 1
RS = {}
def rline(key, lab, fml, unit="", note="", fmt="0.00"):
    global rr
    rs.cell(rr, 1, lab).font = fb
    c = rs.cell(rr, 2, fml); c.font = fb; c.number_format = fmt; c.border = box
    rs.cell(rr, 3, unit).font = fb; rs.cell(rr, 4, note).font = fsmall
    RS[key] = f"B{rr}"; rr += 1

rsec("Manchón")
rline("Tb", "Tensión por correa (frecuencia)", f"={N['Tb']}", "N", "* T = 4·m·Lt²·f², a validar", fmt="0.0")
rline("F", "Fuerza del manchón sobre el rullo (y)", f"={N['W']}", "N", fmt="#,##0")
rline("qy", "qy", f"={N['qy']}", "N/mm")
rline("B", "Ancho del manchón B / posición a → a+B", f"={N['cl']}", "mm", fmt="0.0")
rline("bchk", "b = L - a - B (debe ser ≥ 0)", f"={N['b']}", "mm", fmt="0.0")
rsec("Cargas y esfuerzos")
rline("R1y", "Reacción R1y / R1z", f"={N['R1y']}", "N", fmt="#,##0")
rs.cell(rr - 1, 3).value = f"={N['R1z']}"; rs.cell(rr - 1, 3).number_format = "#,##0"; rs.cell(rr - 1, 4).value = "Columna B: plano y, columna C: plano z [N]"
rline("R2y", "Reacción R2y / R2z", f"={N['R2y']}", "N", fmt="#,##0")
rs.cell(rr - 1, 3).value = f"={N['R2z']}"; rs.cell(rr - 1, 3).number_format = "#,##0"
rline("R1", "Reacción resultante rodamiento izquierdo", f"={N['R1']}", "N", "Dato para verificar el rodamiento", fmt="#,##0")
rline("R2", "Reacción resultante rodamiento derecho", f"={N['R2']}", "N", fmt="#,##0")
rline("T", "Torque", f"={N['T0']}/1000", "N·m", fmt="0.0")
rline("Mmax", "Momento flector máximo", f"=MAX({rng('G')})", "N·mm", fmt="#,##0")
rline("xM", "Posición del M máximo", f"=INDEX({rng('A')},MATCH(MAX({rng('G')}),{rng('G')},0))", "mm", fmt="0.0")
rline("smax", "σ0 máximo (cuerpo, nominal)", f"=MAX({rng('I')})", "MPa")
rline("fl", "Flecha máxima", f"=MAX({rng('AC')})", "mm", "Integración numérica de M/EI (200 tramos)", fmt="0.000")
rline("xfl", "Posición de la flecha máxima", f"=INDEX({rng('A')},MATCH(MAX({rng('AC')}),{rng('AC')},0))", "mm", fmt="0.0")
rsec("Resistencia a fatiga de la pieza")
rline("fS", "fS (superficie)", f"={N['fS']}", "-", "Ec. (25)-(26)", "0.000")
rline("fT", "fT (tamaño)", f"={N['fT']}", "-", "Ec. (21)", "0.000")
rline("fC", "fC (carga)", f"={N['fC']}", "-", "Ec. (24)", "0.000")
rline("f1", "f1 (otros)", f"={N['f1']}", "-", "Ec. (20)", "0.000")
rline("Sep", "Se' probeta", f"={N['Sep']}", "MPa", "Ec. (17)", "0.0")
rline("Se", "Se pieza", f"={N['Se']}", "MPa", "Ec. (20)", "0.0")
rsec("Cuerpo del tubo (sin concentrador)")
rline("csb", "CS fatiga Goodman mínimo", f"=MIN({rng('R')})", "-", "Ec. (44)-b")
rline("xcsb", "Posición", f"=INDEX({rng('A')},MATCH(MIN({rng('R')}),{rng('R')},0))", "mm", fmt="0.0")
rline("flb", "CS fluencia mínimo", f"=MIN({rng('T')})", "-", "Ec. (77)")
rsec("Secciones con concentrador")
for j, k in enumerate("BCD"):
    rline(f"cs{j}", f"Sección {j+1}: CS fatiga Goodman",
          f"=IF(Concentradores!{k}{R['Activa']}=1,Concentradores!{k}{R['CS fatiga Goodman (mín A, B)']},\"inactiva\")", "-")
    rline(f"fl{j}", f"Sección {j+1}: CS fluencia",
          f"=IF(Concentradores!{k}{R['Activa']}=1,Concentradores!{k}{R['CS fluencia']},\"inactiva\")", "-")
    rline(f"v{j}", f"Sección {j+1}: vida",
          f"=IF(Concentradores!{k}{R['Activa']}=1,Concentradores!{k}{R['Vida N [ciclos]']},\"inactiva\")", "ciclos", fmt="#,##0")
rsec("Gobierna")
rline("csg", "CS fatiga mínimo (cuerpo y concentradores)",
      f"=MIN({RS['csb']},{RS['cs0']},{RS['cs1']},{RS['cs2']})", "-", "* Si CS < 1: vida limitada, ver hoja Concentradores / Tensores")
rline("flg", "CS fluencia mínimo", f"=MIN({RS['flb']},{RS['fl0']},{RS['fl1']},{RS['fl2']})", "-")
def srev(sa, sm): return f"{sa}/(1-{sm}/{Su})"
cands = [srev("Tensores!$G$11", "Tensores!$C$11"), srev("Tensores!$S$11", "Tensores!$C$11")]
for k in "BCD":
    on = f"Concentradores!{k}{R['Activa']}"
    cands.append(f"{on}*" + srev(f"Concentradores!{k}{R[chr(65)+': σ'+chr(39)+'a [MPa]']}", f"Concentradores!{k}{R[chr(65)+': σ'+chr(39)+'m [MPa]']}"))
    cands.append(f"{on}*" + srev(f"Concentradores!{k}{R['B: σ'+chr(39)+'a [MPa]']}", f"Concentradores!{k}{R['B: σ'+chr(39)+'m [MPa]']}"))
rline("srev", "Tensión alternante equivalente máxima (σm = 0)", "=MAX(" + ",".join(cands) + ")", "MPa",
      "σa/(1 - σ'm/Su), ec. (48)-b. Es la que se ubica en el diagrama S-N")
rs[RS['csg']].font = fbold; rs[RS['flg']].font = fbold
rr += 1
rs.cell(rr, 1, "Los valores cargados son un ejemplo. CS: coeficiente de seguridad (notación del compendio).").font = fsmall

# vida cuerpo (en Tensores)
# ------------------------------------------------------------------ TENSORES
ts = wb.create_sheet("Tensores", 2)
ts["A1"] = "TENSORES DE TENSIONES [MPa]  (x axial, y radial, z tangente; punto en la superficie)"; ts["A1"].font = fh
ts["A2"] = ("A = fibra extrema (flexión máxima, τV = 0). B = eje neutro (σ flexión = 0, τV máximo). "
            "Medio y alternante según el procedimiento de la sección 7.1 del compendio."); ts["A2"].font = fsmall
for col in "ABCDEFGHIJKLMNOPQRST": ts.column_dimensions[col].width = 9
def mat(r0_, c0, title, sxx, txz, vm_label, vm_formula):
    ts.cell(r0_, c0, title).font = fbold
    vals = [[sxx, "0", txz], ["0", "0", "0"], [txz, "0", "0"]]
    for i in range(3):
        for j in range(3):
            cell = ts.cell(r0_ + 1 + i, c0 + j, "=" + vals[i][j] if vals[i][j] != "0" else 0)
            cell.font = fb; cell.number_format = "0.00"; cell.border = box; cell.alignment = Alignment(horizontal="center")
    ts.cell(r0_ + 4, c0, vm_label).font = fsmall
    v = ts.cell(r0_ + 4, c0 + 2, "=" + vm_formula); v.font = fbold; v.number_format = "0.00"

def block(r0_, title, sig_a, tau_m, tau_aB, kfl_sig, kfl_tauB, note):
    ts.cell(r0_, 1, title).font = Font(name=F, size=11, bold=True)
    ts.cell(r0_ + 1, 1, note).font = fsmall
    base = r0_ + 2
    mat(base, 1, "A medio", "0", tau_m, "σ'm =", f"SQRT(3)*ABS({tau_m})")
    mat(base, 5, "A alternante", sig_a, "0", "σ'a =", f"ABS({sig_a})")
    mat(base, 9, "A máx (fluencia)", kfl_sig, tau_m, "σ'máx =", f"SQRT(({kfl_sig})^2+3*({tau_m})^2)")
    mat(base, 13, "B medio", "0", tau_m, "σ'm =", f"SQRT(3)*ABS({tau_m})")
    mat(base, 17, "B alternante", "0", tau_aB, "σ'a =", f"SQRT(3)*ABS({tau_aB})")
    return base + 6

# body critical row
idx = f"MATCH(MIN({rng('R')}),{rng('R')},0)"
pick = lambda col: f"INDEX({rng(col)},{idx})"
r_ = 4
ts.cell(r_, 1, "Cuerpo: posición crítica x [mm] =").font = fb
ts.cell(r_, 5, "=" + pick("A")).number_format = "0.0"
r_ = block(r_ + 1, "Cuerpo del tubo, sección crítica a fatiga (Kf = 1)", pick("I"), pick("J"), pick("K"),
           pick("I"), pick("K"), "Kf = 1 (tubo entero). Valores tomados de la fila de CS mínimo de la hoja 'Cálculo x'")
# body life
ts.cell(r_, 1, "Vida del cuerpo:").font = fb
sa_b = f"MAX(ABS({pick('I')}),SQRT(3)*{pick('K')})"
ts.cell(r_, 5, f"=IF(MIN({rng('R')})>=1,\"Vida infinita\",IFERROR(({N['Va']}/(ABS({pick('I')})/(1-SQRT(3)*{pick('J')}/{Su})))^{N['Vc']},\"revisar\"))").number_format = "#,##0"
ts.cell(r_, 8, "Ec. (48)-b y (32), punto A").font = fsmall
r_ += 2
for j, k in enumerate("BCD"):
    gc = lambda lab: f"Concentradores!{k}{R[lab]}"
    kfl_s = f"IF({N['swfl']}=1,{gc('Kf = 1 + q·(Kt - 1)')},1)*{gc('σ0 = M·c/I [MPa]')}"
    r_ = block(r_, f"Sección {j+1} con concentrador", gc("A: σa = Kf·|σ0| [MPa]"), gc("A: τm = τT [MPa]"),
               gc("B: τa = Kfs·τV [MPa]"), kfl_s, gc("B: τa = Kfs·τV [MPa]"),
               f"x según Datos. Kf solo en la alternante (Tabla 14). Si la sección está inactiva, ignorar este bloque.")
    ts.cell(r_ - 6 - 1, 12, "x [mm] =").font = fsmall
    ts.cell(r_ - 6 - 1, 14, "=" + gc("x [mm]")).number_format = "0.0"
    r_ += 1

# ------------------------------------------------------------------ GRAFICOS
sn = wb.create_sheet("Diagrama SN")
sn.column_dimensions["A"].width = 14; sn.column_dimensions["B"].width = 16; sn.column_dimensions["C"].width = 16
sn["A1"] = "DIAGRAMA S-N DE LA PIEZA (ecs. 29-32 y 15). Escala log-log"; sn["A1"].font = fbold
sn["A2"] = "Recta entre (10³ ; 0,9·Su) y (10⁶ ; Se); horizontal después de 10⁶ (acero, vida ilimitada)."; sn["A2"].font = fsmall
for j, h in enumerate(["N [ciclos]", "S_N pieza [MPa]", "σa equivalente [MPa]"]):
    c = sn.cell(4, 1 + j, h); c.font = fbold; c.fill = grey
import math as _m
Ns = [10 ** (3 + i * 0.2) for i in range(26)]
for i, Nv in enumerate(Ns):
    rr_ = 5 + i
    sn.cell(rr_, 1, round(Nv, 3)).font = fb; sn.cell(rr_, 1).number_format = "0.00E+00"
    sn.cell(rr_, 2, f"=IF(A{rr_}<=1E6,{N['Va']}*A{rr_}^(-1/{N['Vc']}),{N['Se']})").number_format = "0.0"
    sn.cell(rr_, 3, f"=Resultados!{RS['srev']}").number_format = "0.0"
    sn.cell(rr_, 2).font = fb; sn.cell(rr_, 3).font = fgreen
last = 5 + len(Ns) - 1
snc = ScatterChart(); snc.title = "Diagrama S-N"; snc.style = 13; snc.height = 10; snc.width = 18
snc.x_axis.title = "N [ciclos]"; snc.y_axis.title = "S [MPa]"
snc.x_axis.scaling.logBase = 10; snc.y_axis.scaling.logBase = 10
snc.x_axis.scaling.min = 1000; snc.x_axis.scaling.max = 1E8
snc.x_axis.delete = False; snc.y_axis.delete = False
xs_ = Reference(sn, min_col=1, min_row=5, max_row=last)
for colj, tit in ((2, "S-N pieza"), (3, "σa equivalente (sección crítica)")):
    se_ = Series(Reference(sn, min_col=colj, min_row=5, max_row=last), xs_, title=tit)
    se_.marker.symbol = "none"; se_.graphicalProperties.line.width = 19000
    snc.series.append(se_)
snc.series[1].graphicalProperties.line.dashStyle = "dash"
sn.add_chart(snc, "E4")
sn["E26"] = "Si la línea de σa equivalente queda debajo de Se: vida ilimitada. Si la corta en la recta: vida N (ec. 32)."; sn["E26"].font = fsmall

gs = wb.create_sheet("Gráficos")
def chart(col, title, ytitle, anchor):
    ch = ScatterChart(); ch.title = title; ch.style = 13
    ch.x_axis.title = "x [mm]"; ch.y_axis.title = ytitle; ch.legend = None
    ch.height = 7.5; ch.width = 16
    xs = Reference(cs, min_col=1, min_row=r0, max_row=rN)
    colidx = column_index_from_string(col)
    ys = Reference(cs, min_col=colidx, min_row=r0, max_row=rN)
    s = Series(ys, xs); s.marker.symbol = "none"; s.graphicalProperties.line.width = 15000
    ch.series.append(s)
    ch.x_axis.delete = False; ch.y_axis.delete = False
    gs.add_chart(ch, anchor)
chart("D", "Corte resultante V(x)", "N", "A1"); chart("G", "Momento flector resultante M(x)", "N·mm", "K1")
chart("H", "Torque T(x)", "N·mm", "A17"); chart("AC", "Flecha total", "mm", "K17")
chart("I", "σ0 nominal (fibra extrema)", "MPa", "A33"); chart("R", "CS fatiga del cuerpo (Kf = 1)", "-", "K33")
chart("E", "My(x) plano y (manchón)", "N·mm", "A49"); chart("F", "Mz(x) plano z (gravedad)", "N·mm", "K49")

# ------------------------------------------------------------------ FORMULAS
fs = wb.create_sheet("Fórmulas")
fs.column_dimensions["A"].width = 34; fs.column_dimensions["B"].width = 12; fs.column_dimensions["C"].width = 12
fs.column_dimensions["D"].width = 95
fs["A1"] = "MEMORIA DE CÁLCULO - fuente: Compendio de Cálculo Estructural II, FCEFyN-UNC 2024, cap. 14"; fs["A1"].font = fh
fs["A3"] = "Tabla ec. (26): factor de superficie fS = a·Su^b (Su en MPa)"; fs["A3"].font = fbold
for j, h in enumerate(["Terminación", "a", "b"]):
    fs.cell(4, 1 + j, h).font = fbold
for i, (t_, aa, bb) in enumerate([("Rectificado", 1.58, -0.085), ("Maquinado o estirado en frío", 4.51, -0.265),
                                   ("Laminado en caliente", 57.7, -0.718), ("Forjado", 272, -0.995)]):
    fs.cell(5 + i, 1, t_).font = fb
    fs.cell(5 + i, 2, aa).font = fin; fs.cell(5 + i, 3, bb).font = fin

lines = [
("ESQUEMA DE USO", ""),
("1. Datos", "Cargar en la hoja Datos: geometría (D, t, L), manchón (cantidad, ancho, masa, desarrollo, frecuencia, inicio a), pesos (z), torque (P, n, lado), material (Su, Sy, E, ρ), terminación y f1."),
("2. Concentradores", "Para cada escalón, chavetero o ranura: posición x, Kt y Kts del Anexo 2 y radio r. Desactivar las secciones que no se usen."),
("3. Revisar DCL", "Controlar que el control de equilibrio dé 0 y que b ≥ 0 (hoja Datos), y los diagramas My, Mz, M, T (hoja Gráficos)."),
("4. Resultados", "Hoja Resultados: CS de fatiga y de fluencia del cuerpo y de cada sección, y flecha máxima."),
("5. Tensores", "Hoja Tensores: tensor medio, alternante y máximo en los puntos A y B de cada sección."),
("6. Vida", "Si algún CS de fatiga es < 1: vida limitada en ciclos y horas (hoja Concentradores)."),
("", ""),
("HIPÓTESIS", ""),
("Modelo de viga", "Simplemente apoyada en los centros de los rodamientos (los rodamientos permiten el giro de la sección)."),
("Ejes", "x a lo largo del rullo (desde el apoyo izquierdo), y dirección de la tensión del manchón, z dirección de la gravedad."),
("Carga y *", "Manchón: qy = F/B uniforme entre x = a y x = a + B. F = 2·T total (ramales paralelos). T por correa = 4·m·Lt²·f² (cuerda vibrante, no está en el compendio)."),
("Carga z *", "Peso propio del tubo (todo L) y mitad del peso del manchón repartida en B (opcionales)."),
("Engomado *", "Suma peso (plano z) y define el diámetro de apoyo del manchón (D + 2e). Su rigidez no se cuenta (E goma << E acero)."),
("Desbalanceo *", "F_u = M·e·ω² con e·ω = G (ISO 21940-11). Gira con el eje: tensión media, no alternante. Se informa la relación F_u/F."),
("Velocidad crítica *", "n_c = (30/π)·(π/L)²·√(EI/μ), viga uniforme simplemente apoyada. Fuera del compendio."),
("Combinación", "Cada plano se resuelve por separado; M = √(My² + Mz²) y V = √(Vy² + Vz²) (sección circular). Flecha total = √(fy² + fz²)."),
("Giro", "El eje gira y la carga no: la flexión es completamente alternante en cada punto (σm = 0, σa = σ0). 1 ciclo por vuelta."),
("Torque *", "Entra por un extremo y las correas lo toman de forma pareja en el tramo cargado: T(x) constante hasta el inicio de la carga, lineal hasta 0 al final. Torque constante en el tiempo: τa = 0."),
("Punto A", "Fibra extrema: σ0 máximo y τV = 0. Tensor: σx (alternante), τxz = τT (media)."),
("Punto B", "Eje neutro: σ0 = 0 y τV máximo (alternante por el giro). τxz = τT (media) ± τV (alternante)."),
("Kf", "Tabla (14), material dúctil: Kf solo sobre la componente alternante, la media va sin concentrador."),
("Kfs sobre τV *", "Se aplicó Kfs al corte transversal en el punto B (conservativo; el compendio no lo trata explícitamente)."),
("fT *", "Ec. (21) para ejes rotativos usando D exterior del tubo."),
("Kf en fluencia *", "A definir: switch en Datos. Ec. (3): en dúctil estático no se usa Kt. 1 = se aplica (conservativo)."),
("", ""),
("ECUACIONES (compendio, cap. 14)", ""),
("Equilibrio (por plano)", "R1 = W·(b + B/2)/L + w·L/2 ; R2 = W·(a + B/2)/L + w·L/2 ; W = q·B ; b = L - a - B (w solo en z)"),
("Tensión del manchón *", "m = masa/desarrollo ; Lt = (desarrollo - π·D)/2 si no se carga ; T = 4·m·Lt²·f² ; F = 2·n·T ; qy = F/B"),
("M(x) (por plano)", "R1·x - w·x²/2 - q·(x-a)²/2 en a ≤ x ≤ L-b (0 antes de a; W·(x - a - B/2) después de L-b)"),
("Flecha", "Doble integración numérica de My/EI y Mz/EI (trapecios, 200 tramos) con flecha nula en los apoyos"),
("Tensiones nominales", "σ0 = M·c/I ; τT = T·c/J ; τV = V·Q/(I·b), Q = ⅔(R³ - r³), b = 2t"),
("(6)", "Kf = 1 + q·(Kt - 1)"),
("(9), (10)", "q = 1/(1 + √α/√r). Flexión: √α = 1,24 - 0,00225·Su + 1,6E-6·Su² - 4,1E-10·Su³. Torsión: 0,958 - 0,001834·Su + 1,43E-6·Su² - 4,1E-10·Su³ (Su MPa, r mm). * Verificar coeficientes contra el impreso"),
("(17)", "Se' = 0,5·Su ≤ 690 MPa (aceros)"),
("(20)", "Se = fT·fC·fS·f1·Se'"),
("(21)", "fT = 1,24·d^-0,107 (3 ≤ d ≤ 51 mm) ; 1,51·d^-0,157 (51 < d ≤ 254 mm)"),
("(24)", "fC = 1 con carga combinada (flexión + torsión)"),
("(25), (26)", "fS = a·Su^b (tabla arriba)"),
("(63)-(65)", "Cargas medias y alternas → tensor medio y tensor alternante"),
("(72)", "σ'm = √(σm² + 3τm²) ; σ'a = √((Kf·σa)² + 3(Kf·τa)²)"),
("(44)-a Soderberg", "CS = 1/(σ'm/Sy + σ'a/Se)"),
("(44)-b Goodman", "CS = 1/(σ'm/Su + σ'a/Se)"),
("(44)-c Gerber", "CS = [-ηa + √(ηa² + 4ηm²)]/(2ηm²), ηa = σ'a/Se, ηm = σ'm/Su"),
("(77) Fluencia", "CS = Sy/σ'máx, σ'máx = Von Mises del tensor (medio + alternante)"),
("(31)", "a = (0,9·Su)²/Se ; b = -⅓·log(0,9·Su/Se) ; c = -1/b"),
("(48)-b", "SN = σ'a/(1 - σ'm/Su) (Goodman con CS = 1)"),
("(32)", "N = (a/SN)^c, válida para 10³ ≤ N ≤ 10⁶"),
("", ""),
("LIMITACIONES", ""),
("", "El procedimiento 7.1 supone que todas las cargas varían con la misma ley en el tiempo; si no es así, el resultado queda del lado de la seguridad."),
("", "Kt de tablas solo para geometrías estándar (Anexo 2). Las uniones soldadas no están cubiertas por esta planilla."),
("", "Fatiga de ciclos altos (N > 10³). Los arranques y paradas no se consideran (daño acumulado: sección 9 del compendio)."),
]
rr = 11
for k, v in lines:
    fs.cell(rr, 1, k).font = fbold if v == "" or k.isupper() else fb
    fs.cell(rr, 4, v).font = fb; fs.cell(rr, 4).alignment = Alignment(wrap_text=True)
    rr += 1

for wsx in wb.worksheets:
    wsx.sheet_view.showGridLines = wsx.title in ("Cálculo x",)

# ------------------------------------------------------------------ NOMBRES DEFINIDOS (para la memoria LaTeX)
from openpyxl.workbook.defined_name import DefinedName
def nm(name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)
dmap = [("vD","D"),("vt","t"),("vL","L"),("vnb","nb"),("vbw","bw"),("vmb","mb"),("vLp","Lp"),("vfr","fr"),("vLtin","Ltin"),
        ("vainicio","a"),("vpp","pp"),("vpm","pm"),("vP","P"),("vn","n"),("vlado","lado"),("vSu","Su"),("vSy","Sy"),("vE","E"),
        ("vrho","rho"),("vterm","term"),("vfuno","f1"),("vswfl","swfl"),("vdint","d"),("vArea","A"),("vI","I"),("vJ","J"),
        ("vB","cl"),("vbfin","b"),("vmlin","mlin"),("vLt","Lt"),("vTb","Tb"),("vTtot","Ttot"),("vF","W"),("vqy","qy"),
        ("vWz","Wz"),("vqz","qz"),("vw","w"),("vTcero","T0"),("vRunoy","R1y"),("vRdosy","R2y"),("vRunoz","R1z"),
        ("vRdosz","R2z"),("vRuno","R1"),("vRdos","R2"),("vchk","chk"),("vfS","fS"),("vfT","fT"),("vfC","fC"),
        ("vSep","Sep"),("vSe","Se"),("vVa","Va"),("vVc","Vc"),
        ("veg","eg"),("vrhog","rhog"),("vAg","Ag"),("vG","G"),("vmu","mu"),("vMrot","Mrot"),("vnc","nc"),("vrnc","rnc"),
        ("vFu","Fu"),("vrFu","rFu")]
for ln, k in dmap:
    nm(ln, N[k])
for ln, k in [("vsrev","srev"),("vMmax","Mmax"),("vxM","xM"),("vsmax","smax"),("vfl","fl"),("vxfl","xfl"),("vcsb","csb"),("vxcsb","xcsb"),
              ("vflb","flb"),("vcsg","csg"),("vflg","flg")]:
    nm(ln, f"Resultados!${RS[k][0]}${RS[k][1:]}")
for ln, cell in [("vtxcrit","E4"),("vttaum","C8"),("vtsiga","E8"),("vtvma","G11"),("vtvmm","C11"),("vtvmmax","K11"),
                 ("vttauaB","S8"),("vtvmaB","S11")]:
    nm(ln, f"Tensores!${cell[0]}${cell[1:]}")
cmap = [("x","x [mm]"),("on","Activa"),("Kt",None),("Kts",None),("r",None),("Kf","Kf = 1 + q·(Kt - 1)"),("Kfs","Kfs = 1 + qs·(Kts - 1)"),
        ("sa","A: σ'a [MPa]"),("sm","A: σ'm [MPa]"),("cs","CS fatiga Goodman (mín A, B)"),("fl","CS fluencia"),("vida","Vida N [ciclos]")]
for j, (k, roman) in enumerate(zip("BCD", ("uno", "dos", "tres"))):
    for key, lab in cmap:
        if lab is None:
            dkey = {"Kt": "kKt", "Kts": "kKts", "r": "kr"}[key]
            nm(f"vk{key}{roman}", f"Datos!${k}${conc_rows[dkey]}")
        else:
            nm(f"vk{key}{roman}", f"Concentradores!${k}${R[lab]}")

wb.save("Rullo_PSV_estatico_fatiga.xlsx")
print("saved")
