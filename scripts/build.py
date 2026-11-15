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
inp("xtope", "Tope del manchón, lado derecho (motriz): x (0 = fin del engomado)", 0, "mm",
    "Las correas entran desde la izquierda y hacen tope a la derecha: el manchón va de x_tope - B a x_tope")
sec("Plano z (gravedad)")
inp("pp", "Incluir peso propio (tubo, muñones y engomado) (1 = sí, 0 = no)", 1, "-", "Cada peso repartido en su propio tramo")
inp("pm", "Incluir peso del manchón sobre este rullo (1 = sí, 0 = no)", 1, "-",
    "* Supuesto: la mitad del peso del manchón apoya en cada rullo, repartida en B")
sec("Torque")
inp("P", "Potencia transmitida P", 2.0, "kW")
inp("n", "Velocidad del rullo n", 150, "rpm", "También se usa para pasar ciclos a horas (1 ciclo por vuelta)")
inp("lado", "Lado por donde entra el torque (1 = izquierdo, 0 = derecho = motriz)", 0, "-",
    "* Supuesto: el torque entra por un extremo y las correas lo toman de forma pareja en el tramo cargado")
sec("Balanceo (fuera del compendio)")
inp("ub", "Incluir desbalanceo por rotación en las tensiones (1 = sí, 0 = no)", 1, "-", "Se suma como tensión media (gira con el eje)")
inp("G", "Grado de calidad de balanceo G (ISO 21940-11)", 6.3, "mm/s", "* Desbalanceo residual admisible: e·ω = G")
inp("eu", "Excentricidad de masa medida e (0 = usar e_per de G)", 0, "mm", "Para cargar una excentricidad medida (goma ovalizada, tubo descentrado)")
sec("Masa que gira: eje (acero, densidad ρ de Material)")
inp("DmI", "Muñón izquierdo: diámetro exterior", 60, "mm", "Tubo: entre las soldaduras (D, t). Muñón: desde la soldadura hacia afuera")
inp("dmI", "Muñón izquierdo: diámetro interior (0 = macizo)", 0, "mm")
inp("LmI", "Muñón izquierdo: largo total", 150, "mm", "Desde la soldadura izquierda hasta el extremo (incluye lo que queda fuera del apoyo)")
inp("DmD", "Muñón derecho: diámetro exterior", 60, "mm")
inp("dmD", "Muñón derecho: diámetro interior (0 = macizo)", 0, "mm")
inp("LmD", "Muñón derecho: largo total", 150, "mm", "Desde la soldadura derecha hasta el extremo")
sec("Masa que gira: engomado (espesor y densidad en Engomado)")
inp("xg0", "Inicio del engomado, desde el apoyo izquierdo", 100, "mm")
inp("Lg", "Largo del engomado", 1000, "mm", "Debe quedar sobre el tubo")
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

sec("Soldaduras a tope tubo-muñón (extremos del tubo)")
inp("xsI", "Posición x de la soldadura izquierda", 100, "mm", "Desde el apoyo izquierdo (centro del rodamiento). Entre 0 y L")
inp("xsD", "Posición x de la soldadura derecha", 1100, "mm", "Desde el apoyo izquierdo. Entre 0 y L")
inp("SuM", "Muñón: resistencia a rotura Su", 630, "MPa", "Material del muñón (ejemplo: igual al tubo)")
inp("SyM", "Muñón: tensión de fluencia Sy", 530, "MPa")
inp("termM", "Muñones: terminación superficial", "Maquinado o estirado en frío", "-", "Lista: ec. (26). Se usa en las zonas de muñón")
termM_row = r - 1
inp("elec", "Electrodo (AWS)", "E70xx", "-", "Lista: Shigley Tabla 9-3 (tabla en hoja Fórmulas)")
elec_row = r - 1
inp("sobre", "Sobremonta del cordón", "Con sobremonta", "-", "Lista. Define Kf y fS de la soldadura (tabla en hoja Fórmulas)")
sobre_row = r - 1

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
der("cl", "Ancho del manchón B = cantidad × ancho", f"={N['nb']}*{N['bw']}", "mm", fmt="0.0")
der("xtp", "Tope efectivo del manchón x_tope", f"=IF({N['xtope']}>0,{N['xtope']},{N['xg0']}+{N['Lg']})", "mm",
    "Si no se carga: fin del engomado", fmt="0.0")
der("a", "Inicio del manchón a = x_tope - B", f"={N['xtp']}-{N['cl']}", "mm", "Debe ser ≥ 0 y quedar sobre el engomado", fmt="0.0")
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
der("mu", "Masa por unidad de largo tubo + goma (para n_c)", f"=({N['rho']}*{N['A']}+{N['rhog']}*{N['Ag']})*1000", "kg/m",
    "* n_c y peso propio w toman tubo + goma en todo L (sin muñones)", fmt="0.00")
der("AmI", "Muñón izquierdo: área π/4·(Dm² - dm²)", f"=PI()/4*({N['DmI']}^2-{N['dmI']}^2)", "mm²", fmt="0.0")
der("AmD", "Muñón derecho: área π/4·(Dm² - dm²)", f"=PI()/4*({N['DmD']}^2-{N['dmD']}^2)", "mm²", fmt="0.0")
# pesos por tramo (plano z)
der("wmI", "Peso muñón izquierdo ρ·g·A_mI (tramo 0 → x sold. izq.)", f"={N['pp']}*{N['rho']}*9.81*{N['AmI']}", "N/mm", fmt="0.0000")
der("w", "Peso tubo ρ·g·A (tramo entre soldaduras)", f"={N['pp']}*{N['rho']}*9.81*{N['A']}", "N/mm", fmt="0.0000")
der("wmD", "Peso muñón derecho ρ·g·A_mD (tramo x sold. der. → L)", f"={N['pp']}*{N['rho']}*9.81*{N['AmD']}", "N/mm", fmt="0.0000")
der("wg", "Peso engomado ρg·g·Ag (tramo del engomado)", f"={N['pp']}*{N['rhog']}*9.81*{N['Ag']}", "N/mm",
    "* Lo que queda fuera de los apoyos (punta de los muñones) no se considera", fmt="0.0000")
ZSEG = [(N['wmI'], "0", N['xsI']), (N['w'], N['xsI'], N['xsD']), (N['wmD'], N['xsD'], N['L']),
        (N['wg'], N['xg0'], f"({N['xg0']}+{N['Lg']})"), (N['qz'], N['a'], f"({N['a']}+{N['cl']})")]
Lr = N['L']
der("R1z", "Reacción R1z (plano z) = Σ q·(e-s)·(L - (s+e)/2)/L",
    "=" + "+".join(f"{q_}*({e_}-{s_})*({Lr}-({s_}+{e_})/2)/{Lr}" for q_, s_, e_ in ZSEG), "N", "Suma de los tramos cargados", fmt="#,##0")
der("R2z", "Reacción R2z (plano z) = Σ q·(e-s)·((s+e)/2)/L",
    "=" + "+".join(f"{q_}*({e_}-{s_})*(({s_}+{e_})/2)/{Lr}" for q_, s_, e_ in ZSEG), "N", fmt="#,##0")
der("R1", "Reacción resultante en rodamiento izquierdo R1 = √(R1y² + R1z²)", f"=SQRT({N['R1y']}^2+{N['R1z']}^2)", "N", fmt="#,##0")
der("R2", "Reacción resultante en rodamiento derecho R2", f"=SQRT({N['R2y']}^2+{N['R2z']}^2)", "N", fmt="#,##0")
der("chk", "Control: ΣRy - F + ΣRz - Σ pesos (debe dar 0)",
    f"={N['R1y']}+{N['R2y']}-{N['W']}+{N['R1z']}+{N['R2z']}-(" + "+".join(f"{q_}*({e_}-{s_})" for q_, s_, e_ in ZSEG) + ")", "N", fmt="0.00")
# secciones y resistencia de los muñones
def _ft(D_): return f"IF({D_}<=51,1.24*{D_}^-0.107,IF({D_}<=254,1.51*{D_}^-0.157,1.51*254^-0.157))"
for sfx, Dm_, dm_ in (("I", N['DmI'], N['dmI']), ("D", N['DmD'], N['dmD'])):
    lado_ = "izquierdo" if sfx == "I" else "derecho"
    der(f"Im{sfx}", f"Muñón {lado_}: I = π/64·(Dm⁴ - dm⁴)", f"=PI()/64*({Dm_}^4-{dm_}^4)", "mm⁴", fmt="#,##0")
    der(f"cm{sfx}", f"Muñón {lado_}: c = Dm/2", f"={Dm_}/2", "mm", fmt="0.0")
    der(f"Qm{sfx}", f"Muñón {lado_}: Q = ⅔·(R³ - r³)", f"=2/3*(({Dm_}/2)^3-({dm_}/2)^3)", "mm³", fmt="#,##0")
    der(f"bm{sfx}", f"Muñón {lado_}: b = Dm - dm", f"={Dm_}-{dm_}", "mm", fmt="0.0")
    der(f"fTm{sfx}", f"Muñón {lado_}: fT", "=" + _ft(Dm_), "-", "Ec. (21) con Dm", fmt="0.000")
der("fSm", "Muñones: fS = a·SuM^b", f"=INDEX(Fórmulas!$B$5:$B$8,MATCH({N['termM']},Fórmulas!$A$5:$A$8,0))*{N['SuM']}"
    f"^INDEX(Fórmulas!$C$5:$C$8,MATCH({N['termM']},Fórmulas!$A$5:$A$8,0))", "-", "Ec. (25)-(26) con la terminación de los muñones", fmt="0.000")
der("SemI", "Muñón izquierdo: Se = fS·fT·fC·f1·mín(0,5·SuM ; 690)", f"={N['fSm']}*{N['fTmI']}*{N['f1']}*MIN(0.5*{N['SuM']},690)", "MPa", "Ec. (20), fC = 1", fmt="0.0")
der("SemD", "Muñón derecho: Se = fS·fT·fC·f1·mín(0,5·SuM ; 690)", f"={N['fSm']}*{N['fTmD']}*{N['f1']}*MIN(0.5*{N['SuM']},690)", "MPa", "Ec. (20), fC = 1", fmt="0.0")
der("safM", "Muñones: √α Neuber flexión (SuM)", f"=1.24-0.00225*{N['SuM']}+1.6E-6*{N['SuM']}^2-4.1E-10*{N['SuM']}^3", "√mm", "Ec. (10)-a", fmt="0.0000")
der("satM", "Muñones: √α Neuber torsión (SuM)", f"=0.958-0.001834*{N['SuM']}+1.43E-6*{N['SuM']}^2-4.1E-10*{N['SuM']}^3", "√mm", "Ec. (10)-b", fmt="0.0000")
der("Ltu", "Largo del tubo L_tubo = x sold. der. - x sold. izq.", f"={N['xsD']}-{N['xsI']}", "mm", fmt="0.0")
der("Meje", "Masa del eje M_eje = ρ·(A·L_tubo + A_mI·L_mI + A_mD·L_mD)",
    f"={N['rho']}*({N['A']}*{N['Ltu']}+{N['AmI']}*{N['LmI']}+{N['AmD']}*{N['LmD']})", "kg", fmt="0.00")
der("Mgoma", "Masa del engomado M_goma = ρg·Ag·L_g", f"={N['rhog']}*{N['Ag']}*{N['Lg']}", "kg", fmt="0.00")
der("Mrot", "Masa que gira M = M_eje + M_goma", f"={N['Meje']}+{N['Mgoma']}", "kg", fmt="0.00")
der("xG", "Centro de masa x_G (desde el apoyo izquierdo)",
    f"=({N['rho']}*({N['A']}*{N['Ltu']}*({N['xsI']}+{N['xsD']})/2+{N['AmI']}*{N['LmI']}*({N['xsI']}-{N['LmI']}/2)"
    f"+{N['AmD']}*{N['LmD']}*({N['xsD']}+{N['LmD']}/2))+{N['Mgoma']}*({N['xg0']}+{N['Lg']}/2))/{N['Mrot']}", "mm",
    "Σ(mᵢ·xᵢ)/M de muñones, tubo y engomado", fmt="0.0")
der("om", "Velocidad angular ω = 2π·n/60", f"=2*PI()*{N['n']}/60", "rad/s", fmt="0.00")
der("eper", "Excentricidad admisible e_per = 1000·G/ω", f"=1000*{N['G']}/{N['om']}", "g·mm/kg",
    "ISO 21940-11. g·mm/kg = µm", fmt="0.0")
der("Uper", "Desbalance residual admisible U_per = e_per·M", f"={N['eper']}*{N['Mrot']}", "g·mm", "Rotor completo (eje + engomado)", fmt="#,##0")
der("UperI", "U_per en el apoyo izquierdo = U_per·(L - x_G)/L", f"={N['Uper']}*({N['L']}-{N['xG']})/{N['L']}", "g·mm",
    "* Reparto por palanca respecto del centro de masa: verificar contra ISO 21940-11", fmt="#,##0")
der("UperD", "U_per en el apoyo derecho = U_per·x_G/L", f"={N['Uper']}*{N['xG']}/{N['L']}", "g·mm", "*", fmt="#,##0")
der("nc", "1ª velocidad crítica de flexión n_c = (30/π)·(π/L)²·√(EI/μ)", 
    f"=30/PI()*(PI()/({N['L']}/1000))^2*SQRT({N['E']}*1E6*{N['I']}*1E-12/{N['mu']})", "rpm",
    "* Viga uniforme simplemente apoyada (teoría de vigas, fuera del compendio). No incluye la masa del manchón", fmt="#,##0")
der("rnc", "Relación n / n_c", f"={N['n']}/{N['nc']}", "-", "Lejos de 1 = sin resonancia de flexión", fmt="0.000")
der("ee", "Excentricidad usada e", f"=IF({N['eu']}>0,{N['eu']},{N['eper']}/1000)", "mm", "e_per [µm]/1000 si no se carga e medida", fmt="0.0000")
der("Fu", "Fuerza centrífuga por desbalanceo F_u = M·e·ω²", f"={N['Mrot']}*{N['ee']}/1000*{N['om']}^2", "N",
    "* Gira con el eje: para el material es una carga fija (tensión media), no alternante. Con e = e_per: F_u = M·G·ω", fmt="0.0")
der("wu", "Desbalanceo en las tensiones wu = F_u/L (si se incluye)", f"={N['ub']}*{N['Fu']}/{N['L']}", "N/mm",
    "* Supuesto: desbalanceo uniforme a lo largo de L", fmt="0.00000")
der("rFu", "Relación F_u / F manchón", f"={N['Fu']}/{N['W']}", "-", fmt="0.0000")

dv = DataValidation(type="list", formula1='"Rectificado,Maquinado o estirado en frío,Laminado en caliente,Forjado"', allow_blank=False)
ds.add_data_validation(dv); dv.add(f"B{term_row}"); dv.add(f"B{termM_row}")
dve = DataValidation(type="list", formula1="=Fórmulas!$F$5:$F$10", allow_blank=False)
ds.add_data_validation(dve); dve.add(f"B{elec_row}")
dvs = DataValidation(type="list", formula1='"Con sobremonta,Maquinada al ras"', allow_blank=False)
ds.add_data_validation(dvs); dvs.add(f"B{sobre_row}")
img = Image("dcl.png"); img.width, img.height = 620, 434
ds.add_image(img, "G3")
ds["G26"] = "Diagrama de cuerpo libre (modelo simplemente apoyado). x se mide desde el apoyo izquierdo."; ds["G26"].font = fsmall

# helper formula pieces (x given as a cell ref) -----------------------------
L, a, b, w, cl, T0 = (N[k] for k in ("L", "a", "b", "w", "cl", "T0"))
def segV(x, q_, s_, e_): return f"{q_}*MAX(0,MIN({x},{e_})-{s_})"
def segM(x, q_, s_, e_): return f"{q_}*IF({x}<={s_},0,IF({x}<={e_},({x}-{s_})^2/2,({e_}-{s_})*({x}-({s_}+{e_})/2)))"
YSEG = [(N['qy'], N['a'], f"({N['a']}+{N['cl']})")]
def Vyf(x): return f"({N['R1y']}-" + "-".join(segV(x, *sg) for sg in YSEG) + ")"
def Myf(x): return f"({N['R1y']}*{x}-" + "-".join(segM(x, *sg) for sg in YSEG) + ")"
def Vzf(x): return f"({N['R1z']}-" + "-".join(segV(x, *sg) for sg in ZSEG) + ")"
def Mzf(x): return f"({N['R1z']}*{x}-" + "-".join(segM(x, *sg) for sg in ZSEG) + ")"
def Vf(x): return f"SQRT({Vyf(x)}^2+{Vzf(x)}^2)"
def Mf(x): return f"SQRT({Myf(x)}^2+{Mzf(x)}^2)"
def zsel(x, vI, vT, vD): return f"IF({x}<{N['xsI']},{vI},IF({x}<={N['xsD']},{vT},{vD}))"
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
         "Mz/EI [1/mm]", "∫ Mz/EI", "∫∫ Mz/EI", "Flecha z [mm]", "Flecha total [mm]",
         "Vu desbalanceo [N]", "Mu desbalanceo [N·mm]", "σm desbalanceo [MPa]", "τm desbalanceo (corte) [MPa]",
         "Zona (1 muñón izq, 2 tubo, 3 muñón der)", "I [mm⁴]", "J [mm⁴]", "c [mm]", "Q [mm³]", "b [mm]", "Se [MPa]", "Su [MPa]", "Sy [MPa]"]
cs["A1"] = ("EJE COMPLETO ENTRE APOYOS (Kf = 1): muñón izq. | tubo | muñón der. Sección y material según la zona. "
            "Concentradores en hoja 'Concentradores'. A = fibra extrema, B = eje neutro. y: tensión del manchón, z: gravedad"); cs["A1"].font = fbold
for j, h in enumerate(heads):
    c = cs.cell(3, 1 + j, h); c.font = fbold; c.fill = grey; c.alignment = Alignment(wrap_text=True, horizontal="center")
    cs.column_dimensions[c.column_letter].width = 13
cs.row_dimensions[3].height = 42
NSEG = 200; r0 = 4; rN = r0 + NSEG
Se, Su, Sy, I, J, c_, Q, bs, E = (N[k] for k in ("Se", "Su", "Sy", "I", "J", "c", "Q", "bs", "E"))
fmts = ["0.0", "#,##0", "#,##0", "#,##0", "#,##0", "#,##0", "#,##0", "#,##0"] + ["0.00"] * 12 + ["0.00E+00"] * 3 + ["0.0000"] + ["0.00E+00"] * 3 + ["0.0000", "0.0000", "0.00", "0.0", "0.000", "0.000", "0", "#,##0", "#,##0", "0.0", "#,##0", "0.0", "0.0", "0", "0"]
for i in range(NSEG + 1):
    rr = r0 + i; x = f"A{rr}"
    F_ = {}
    F_["A"] = f"={L}*{i}/{NSEG}"
    F_["B"] = "=" + Vyf(x); F_["C"] = "=" + Vzf(x); F_["D"] = f"=SQRT(B{rr}^2+C{rr}^2)"
    F_["E"] = "=" + Myf(x); F_["F"] = "=" + Mzf(x); F_["G"] = f"=SQRT(E{rr}^2+F{rr}^2)"
    F_["H"] = "=" + Tf(x)
    F_["I"] = f"=G{rr}*AK{rr}/AI{rr}"
    F_["J"] = f"=H{rr}*AK{rr}/AJ{rr}"
    F_["K"] = f"=D{rr}*AL{rr}/(AI{rr}*AM{rr})"
    F_["L"] = f"=ABS(I{rr})"
    F_["M"] = f"=SQRT(AF{rr}^2+3*J{rr}^2)"
    F_["N"] = f"=IFERROR(1/(L{rr}/AN{rr}+M{rr}/AO{rr}),9999)"
    F_["O"] = f"=SQRT(3)*K{rr}"
    F_["P"] = f"=SQRT(3)*(J{rr}+AG{rr})"
    F_["Q"] = f"=IFERROR(1/(O{rr}/AN{rr}+P{rr}/AO{rr}),9999)"
    F_["R"] = f"=MIN(N{rr},Q{rr})"
    F_["S"] = f"=MAX(SQRT((ABS(I{rr})+AF{rr})^2+3*J{rr}^2),SQRT(3)*(K{rr}+J{rr}+AG{rr}))"
    F_["T"] = f"=IFERROR(AP{rr}/S{rr},9999)"
    for (mc, kc, i1, i2, fl) in (("E", "U", "V", "W", "X"), ("F", "Y", "Z", "AA", "AB")):
        F_[kc] = f"={mc}{rr}/({E}*AI{rr})"
        if i == 0:
            F_[i1] = 0; F_[i2] = 0
        else:
            F_[i1] = f"={i1}{rr-1}+({kc}{rr-1}+{kc}{rr})/2*(A{rr}-A{rr-1})"
            F_[i2] = f"={i2}{rr-1}+({i1}{rr-1}+{i1}{rr})/2*(A{rr}-A{rr-1})"
        F_[fl] = f"=-({i2}{rr}-A{rr}/{L}*${i2}${rN})"
    F_["AC"] = f"=SQRT(X{rr}^2+AB{rr}^2)"
    F_["AD"] = f"={N['wu']}*({L}/2-A{rr})"
    F_["AE"] = f"={N['wu']}*A{rr}*({L}-A{rr})/2"
    F_["AF"] = f"=AE{rr}*AK{rr}/AI{rr}"
    F_["AG"] = f"=ABS(AD{rr})*AL{rr}/(AI{rr}*AM{rr})"
    F_["AH"] = "=" + zsel(x, 1, 2, 3)
    F_["AI"] = "=" + zsel(x, N['ImI'], I, N['ImD'])
    F_["AJ"] = f"=2*AI{rr}"
    F_["AK"] = "=" + zsel(x, N['cmI'], c_, N['cmD'])
    F_["AL"] = "=" + zsel(x, N['QmI'], Q, N['QmD'])
    F_["AM"] = "=" + zsel(x, N['bmI'], bs, N['bmD'])
    F_["AN"] = "=" + zsel(x, N['SemI'], Se, N['SemD'])
    F_["AO"] = "=" + zsel(x, N['SuM'], Su, N['SuM'])
    F_["AP"] = "=" + zsel(x, N['SyM'], Sy, N['SyM'])
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
 ("Zona (1 muñón izq, 2 tubo, 3 muñón der)", None, "Según x frente a las soldaduras"),
 ("I zona [mm⁴]", None, "Sección de la zona (tubo o muñón)"),
 ("J zona [mm⁴]", None, "J = 2I"),
 ("c zona [mm]", None, ""),
 ("Q zona [mm³]", None, ""),
 ("b zona [mm]", None, ""),
 ("Su zona [MPa]", None, "Tubo: Su; muñón: Su del muñón"),
 ("Sy zona [MPa]", None, ""),
 ("Se zona [MPa]", None, "Ec. (20) con el material, la terminación y el diámetro de la zona"),
 ("√α flexión zona", None, "Ec. (10)-a con Su de la zona"),
 ("√α torsión zona", None, "Ec. (10)-b con Su de la zona"),
 ("V [N]", None, "Equilibrio (DCL)"),
 ("M [N·mm]", None, "Equilibrio (DCL)"),
 ("T [N·mm]", None, "* Distribución de T supuesta (ver Datos)"),
 ("σ0 = M·c/I [MPa]", None, "Tensión nominal de flexión"),
 ("τT = T·c/J [MPa]", None, "Tensión nominal de torsión"),
 ("τV = V·Q/(I·b) [MPa]", None, "Corte transversal en el eje neutro"),
 ("σm desbalanceo [MPa]", None, "* Mu·c/I, Mu = wu·x·(L-x)/2. Media sin Kf (Tabla 14)"),
 ("τm desbalanceo (corte) [MPa]", None, "* Vu·Q/(I·b) en el eje neutro"),
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
     "Zona (1 muñón izq, 2 tubo, 3 muñón der)": "=" + zsel(x, 1, 2, 3),
     "I zona [mm⁴]": "=" + zsel(x, N['ImI'], I, N['ImD']),
     "J zona [mm⁴]": f"=2*{g('I zona [mm⁴]')}",
     "c zona [mm]": "=" + zsel(x, N['cmI'], c_, N['cmD']),
     "Q zona [mm³]": "=" + zsel(x, N['QmI'], Q, N['QmD']),
     "b zona [mm]": "=" + zsel(x, N['bmI'], bs, N['bmD']),
     "Su zona [MPa]": "=" + zsel(x, N['SuM'], Su, N['SuM']),
     "Sy zona [MPa]": "=" + zsel(x, N['SyM'], Sy, N['SyM']),
     "Se zona [MPa]": "=" + zsel(x, N['SemI'], Se, N['SemD']),
     "√α flexión zona": "=" + zsel(x, N['safM'], N['sa_f'], N['safM']),
     "√α torsión zona": "=" + zsel(x, N['satM'], N['sa_t'], N['satM']),
     "V [N]": "=" + Vf(x),
     "M [N·mm]": "=" + Mf(x),
     "T [N·mm]": "=" + Tf(x),
     "σ0 = M·c/I [MPa]": f"={g('M [N·mm]')}*{g('c zona [mm]')}/{g('I zona [mm⁴]')}",
     "τT = T·c/J [MPa]": f"={g('T [N·mm]')}*{g('c zona [mm]')}/{g('J zona [mm⁴]')}",
     "τV = V·Q/(I·b) [MPa]": f"=ABS({g('V [N]')})*{g('Q zona [mm³]')}/({g('I zona [mm⁴]')}*{g('b zona [mm]')})",
     "q flexión = 1/(1+√α/√r)": f"=1/(1+{g('√α flexión zona')}/SQRT({kr}))",
     "q torsión": f"=1/(1+{g('√α torsión zona')}/SQRT({kr}))",
     "Kf = 1 + q·(Kt - 1)": f"=1+{g('q flexión = 1/(1+√α/√r)')}*({kt}-1)",
     "Kfs = 1 + qs·(Kts - 1)": f"=1+{g('q torsión')}*({kts}-1)",
     "A: σa = Kf·|σ0| [MPa]": f"={g('Kf = 1 + q·(Kt - 1)')}*ABS({g('σ0 = M·c/I [MPa]')})",
     "A: τm = τT [MPa]": f"={g('τT = T·c/J [MPa]')}",
     "A: σ'a [MPa]": f"={g('A: σa = Kf·|σ0| [MPa]')}",
     "σm desbalanceo [MPa]": f"={N['wu']}*{x}*({L}-{x})/2*{g('c zona [mm]')}/{g('I zona [mm⁴]')}",
     "τm desbalanceo (corte) [MPa]": f"=ABS({N['wu']}*({L}/2-{x}))*{g('Q zona [mm³]')}/({g('I zona [mm⁴]')}*{g('b zona [mm]')})",
     "A: σ'm [MPa]": f"=SQRT({g('σm desbalanceo [MPa]')}^2+3*{g('A: τm = τT [MPa]')}^2)",
     "B: τa = Kfs·τV [MPa]": f"={g('Kfs = 1 + qs·(Kts - 1)')}*{g('τV = V·Q/(I·b) [MPa]')}",
     "B: σ'a [MPa]": f"=SQRT(3)*{g('B: τa = Kfs·τV [MPa]')}",
     "B: σ'm [MPa]": f"=SQRT(3)*({g('A: τm = τT [MPa]')}+{g('τm desbalanceo (corte) [MPa]')})",
    }
    ea = f"({g(chr(65)+': σ'+chr(39)+'a [MPa]')}/{g('Se zona [MPa]')})"
    sA_a = g("A: σ'a [MPa]"); sA_m = g("A: σ'm [MPa]"); sB_a = g("B: σ'a [MPa]"); sB_m = g("B: σ'm [MPa]")
    F_["A: CS Goodman"] = f"=IFERROR(1/({sA_a}/{g('Se zona [MPa]')}+{sA_m}/{g('Su zona [MPa]')}),9999)"
    F_["A: CS Soderberg"] = f"=IFERROR(1/({sA_a}/{g('Se zona [MPa]')}+{sA_m}/{g('Sy zona [MPa]')}),9999)"
    F_["A: CS Gerber"] = (f"=IFERROR(IF({sA_m}=0,{g('Se zona [MPa]')}/{sA_a},(-({sA_a}/{g('Se zona [MPa]')})+SQRT(({sA_a}/{g('Se zona [MPa]')})^2+4*({sA_m}/{g('Su zona [MPa]')})^2))"
                          f"/(2*({sA_m}/{g('Su zona [MPa]')})^2)),9999)")
    F_["B: CS Goodman"] = f"=IFERROR(1/({sB_a}/{g('Se zona [MPa]')}+{sB_m}/{g('Su zona [MPa]')}),9999)"
    F_["CS fatiga Goodman (mín A, B)"] = f"=MIN({g('A: CS Goodman')},{g('B: CS Goodman')})"
    kfl = f"IF({N['swfl']}=1,{g('Kf = 1 + q·(Kt - 1)')},1)"; kfsl = f"IF({N['swfl']}=1,{g('Kfs = 1 + qs·(Kts - 1)')},1)"
    F_["σ'máx A (fluencia) [MPa]"] = f"=SQRT(({kfl}*ABS({g('σ0 = M·c/I [MPa]')})+{g('σm desbalanceo [MPa]')})^2+3*{g('τT = T·c/J [MPa]')}^2)"
    F_["σ'máx B (fluencia) [MPa]"] = f"=SQRT(3)*({kfsl}*{g('τV = V·Q/(I·b) [MPa]')}+{g('τT = T·c/J [MPa]')}+{g('τm desbalanceo (corte) [MPa]')})"
    F_["CS fluencia"] = f"=IFERROR({g('Sy zona [MPa]')}/MAX({g('σ' + chr(39) + 'máx A (fluencia) [MPa]')},{g('σ' + chr(39) + 'máx B (fluencia) [MPa]')}),9999)"
    sn = g("Vida: SN = σ'a/(1 - σ'm/Su) [MPa]")
    F_["Vida: SN = σ'a/(1 - σ'm/Su) [MPa]"] = (f"=IF({g('CS fatiga Goodman (mín A, B)')}>=1,\"Vida infinita\","
        f"IF({g('A: CS Goodman')}<={g('B: CS Goodman')},{sA_a}/(1-{sA_m}/{g('Su zona [MPa]')}),{sB_a}/(1-{sB_m}/{g('Su zona [MPa]')})))")
    F_["Vida N [ciclos]"] = (f"=IF(ISNUMBER({sn}),IF({sn}>0.9*{g('Su zona [MPa]')},\"< 10^3: fuera de rango\",((0.9*{g('Su zona [MPa]')})^2/{g('Se zona [MPa]')}/{sn})^(3/LOG10(0.9*{g('Su zona [MPa]')}/{g('Se zona [MPa]')}))),\"Vida infinita\")")
    vn = g("Vida N [ciclos]")
    F_["Vida [horas]"] = f"=IF(ISNUMBER({vn}),{vn}/({N['n']}*60),\"-\")"
    for lab, fml in F_.items():
        cell = cc[g(lab)]; cell.value = fml; cell.font = fb; cell.number_format = "0.00"
    for lab in ("V [N]", "M [N·mm]", "T [N·mm]", "Vida N [ciclos]", "I zona [mm⁴]", "J zona [mm⁴]", "Q zona [mm³]"):
        cc[g(lab)].number_format = "#,##0"
    for lab in ("q flexión = 1/(1+√α/√r)", "q torsión", "Kf = 1 + q·(Kt - 1)", "Kfs = 1 + qs·(Kts - 1)"):
        cc[g(lab)].number_format = "0.000"
    cc[g("x [mm]")].font = fgreen; cc[g("Activa")].font = fgreen

# ------------------------------------------------------------------ SOLDADURA (a tope tubo-muñón)
sw = wb.create_sheet("Soldadura")
sw.column_dimensions["A"].width = 44
for col in "BC": sw.column_dimensions[col].width = 14
sw.column_dimensions["D"].width = 95
sw["A1"] = "SOLDADURAS A TOPE TUBO-MUÑÓN (extremos del tubo)"; sw["A1"].font = fh
sw["A2"] = ("Uso: 1) Datos > 'Soldaduras a tope tubo-muñón': posición x de cada soldadura, material del muñón, "
            "electrodo y sobremonta.  2) Esta hoja calcula sola, no se edita.  3) Resumen en Resultados > Soldaduras."); sw["A2"].font = fsmall
sw["A3"] = ("Hipótesis: soldadura circunferencial a tope con penetración completa: la sección resistente es la corona del tubo "
            "(garganta = espesor t, sin sobremonta). Fuera de alcance: penetración parcial, filete, flexión local y tensiones "
            "residuales (ver Fórmulas)."); sw["A3"].font = fsmall
sw["A5"] = "Magnitud"; sw["B5"] = "Izquierda"; sw["C5"] = "Derecha"; sw["D5"] = "Ecuación / fuente"
for col in "ABCD": sw[f"{col}5"].font = fbold; sw[f"{col}5"].fill = grey
srows = [
 ("x [mm]", "Dato (hoja Datos)"),
 ("V [N]", "Equilibrio (DCL), mismo modelo que el resto de la planilla"),
 ("M [N·mm]", "Equilibrio (DCL)"),
 ("T [N·mm]", "* Distribución de T supuesta (ver Datos)"),
 ("Garganta t = espesor del tubo [mm]", "Compendio cap. 18, ec. (1): t = espesor de las partes a unir; no incluye la sobremonta"),
 ("I = π/64·(D⁴ - d⁴) [mm⁴]", "Penetración completa: la sección de la soldadura es la corona del tubo (hoja Datos)"),
 ("J = 2I [mm⁴]", "Hoja Datos"),
 ("σ0 = M·c/I [MPa]", "Compendio cap. 18, ec. (1) y Tabla 4 (σ = M/S); c = D/2"),
 ("τT = T·c/J [MPa]", "Compendio cap. 18, ec. (2) y Tabla 4 (τ = T·c/J)"),
 ("τV = V·Q/(I·b) [MPa]", "Corte transversal en el eje neutro, mismo criterio que la hoja Cálculo x"),
 ("Su tubo [MPa]", "Hoja Datos"),
 ("Su muñón [MPa]", "Hoja Datos"),
 ("Su electrodo [MPa]", "Shigley 10ª ed., Tabla 9-3 (mínimos del metal de aporte); tabla en hoja Fórmulas"),
 ("Su mínimo [MPa]", "Compendio cap. 18: verificar contra el material de aporte y los materiales base. Se toma el menor"),
 ("Sy tubo [MPa]", "Hoja Datos"),
 ("Sy muñón [MPa]", "Hoja Datos"),
 ("Sy electrodo [MPa]", "Shigley 10ª ed., Tabla 9-3"),
 ("Sy mínimo [MPa]", "Compendio cap. 18: se toma el menor"),
 ("Kf soldadura", "Con sobremonta: 1,2 (Shigley Tabla 9-5, soldadura a tope reforzada). * Maquinada al ras: 1,0 (supuesto)"),
 ("Terminación para fS", "Con sobremonta: Forjado (Shigley sec. 9-7). Maquinada al ras: Maquinado (tabla en hoja Fórmulas)"),
 ("fS = a·(Su mín)^b", "Ec. (25)-(26), con Su mínimo"),
 ("fT", "Ec. (21), igual que el tubo (* D exterior)"),
 ("fC", "Ec. (24): fC = 1 con carga combinada"),
 ("f1", "Ec. (20), dato en hoja Datos"),
 ("Se' = mín(0,5·Su mín ; 690) [MPa]", "Ec. (17)"),
 ("Se = fS·fT·fC·f1·Se' [MPa]", "Ec. (20)"),
 ("A: σa = Kf·|σ0| [MPa]", "Tabla (14): Kf solo sobre la alterna (dúctil). Eje rotativo: flexión alternante, σm = 0"),
 ("A: τm = τT [MPa]", "Tabla (14): media sin Kf (dúctil). * Torque constante: τa = 0"),
 ("A: σ'a [MPa]", "Ec. (72)"),
 ("A: σ'm [MPa]", "Ec. (72): σ'm = √3·τm"),
 ("B: τa = Kf·τV [MPa]", "* Kf de la soldadura aplicado al corte transversal (mismo criterio conservativo que en Concentradores)"),
 ("B: σ'a [MPa]", "Ec. (73)"),
 ("B: σ'm [MPa]", "Ec. (73)"),
 ("A: CS Goodman", "Ec. (44)-b, con Su mínimo"),
 ("A: CS Soderberg", "Ec. (44)-a, con Sy mínimo"),
 ("A: CS Gerber", "Ec. (44)-c, con Su mínimo"),
 ("B: CS Goodman", "Ec. (44)-b"),
 ("CS fatiga Goodman (mín A, B)", ""),
 ("σ'máx A (fluencia) [MPa]", "Ec. (77): tensor medio + alternante, Von Mises. Kf según el selector de Datos"),
 ("σ'máx B (fluencia) [MPa]", "Ec. (77)"),
 ("CS fluencia = Sy mín / σ'máx", "Ec. (77)"),
 ("Vida: a = (0,9·Su mín)²/Se [MPa]", "Ec. (31)"),
 ("Vida: c = -1/b = 3/log(0,9·Su mín/Se)", "Ec. (31)"),
 ("Vida: SN = σ'a/(1 - σ'm/Su mín) [MPa]", "Ec. (48)-b con CS = 1 (solo si CS Goodman < 1)"),
 ("Vida N [ciclos]", "Ec. (32): N = (a/SN)^c. Válida si Se < SN ≤ 0,9·Su"),
 ("Vida [horas]", "N / (n·60)"),
]
SR = {lab: 6 + i for i, (lab, note) in enumerate(srows)}
for lab, note in srows:
    sw.cell(SR[lab], 1, lab).font = fb; sw.cell(SR[lab], 4, note).font = fsmall
FTAB = "Fórmulas!$A$5:$A$8"
for k, xs in (("B", N["xsI"]), ("C", N["xsD"])):
    g = lambda lab, k=k: f"{k}{SR[lab]}"
    x = g("x [mm]")
    sumin = g("Su mínimo [MPa]"); symin = g("Sy mínimo [MPa]"); se = g("Se = fS·fT·fC·f1·Se' [MPa]")
    sAa, sAm = g("A: σ'a [MPa]"), g("A: σ'm [MPa]"); sBa, sBm = g("B: σ'a [MPa]"), g("B: σ'm [MPa]")
    kf = g("Kf soldadura")
    term = g("Terminación para fS")
    F_ = {
     "x [mm]": f"={xs}",
     "V [N]": "=" + Vf(x),
     "M [N·mm]": "=" + Mf(x),
     "T [N·mm]": "=" + Tf(x),
     "Garganta t = espesor del tubo [mm]": f"={N['t']}",
     "I = π/64·(D⁴ - d⁴) [mm⁴]": f"={I}",
     "J = 2I [mm⁴]": f"={J}",
     "σ0 = M·c/I [MPa]": f"={g('M [N·mm]')}*{c_}/{g('I = π/64·(D⁴ - d⁴) [mm⁴]')}",
     "τT = T·c/J [MPa]": f"={g('T [N·mm]')}*{c_}/{g('J = 2I [mm⁴]')}",
     "τV = V·Q/(I·b) [MPa]": f"=ABS({g('V [N]')})*{Q}/({g('I = π/64·(D⁴ - d⁴) [mm⁴]')}*{bs})",
     "Su tubo [MPa]": f"={N['Su']}",
     "Su muñón [MPa]": f"={N['SuM']}",
     "Su electrodo [MPa]": f"=INDEX(Fórmulas!$G$5:$G$10,MATCH({N['elec']},Fórmulas!$F$5:$F$10,0))",
     "Su mínimo [MPa]": f"=MIN({g('Su tubo [MPa]')},{g('Su muñón [MPa]')},{g('Su electrodo [MPa]')})",
     "Sy tubo [MPa]": f"={N['Sy']}",
     "Sy muñón [MPa]": f"={N['SyM']}",
     "Sy electrodo [MPa]": f"=INDEX(Fórmulas!$H$5:$H$10,MATCH({N['elec']},Fórmulas!$F$5:$F$10,0))",
     "Sy mínimo [MPa]": f"=MIN({g('Sy tubo [MPa]')},{g('Sy muñón [MPa]')},{g('Sy electrodo [MPa]')})",
     "Kf soldadura": f"=INDEX(Fórmulas!$G$15:$G$16,MATCH({N['sobre']},Fórmulas!$F$15:$F$16,0))",
     "Terminación para fS": f"=INDEX(Fórmulas!$H$15:$H$16,MATCH({N['sobre']},Fórmulas!$F$15:$F$16,0))",
     "fS = a·(Su mín)^b": (f"=INDEX(Fórmulas!$B$5:$B$8,MATCH({term},{FTAB},0))*{sumin}"
                           f"^INDEX(Fórmulas!$C$5:$C$8,MATCH({term},{FTAB},0))"),
     "fT": f"={N['fT']}",
     "fC": 1,
     "f1": f"={N['f1']}",
     "Se' = mín(0,5·Su mín ; 690) [MPa]": f"=MIN(0.5*{sumin},690)",
     "Se = fS·fT·fC·f1·Se' [MPa]": (f"={g('fS = a·(Su mín)^b')}*{g('fT')}*{g('fC')}*{g('f1')}"
                                    f"*{g(chr(83)+'e'+chr(39)+' = mín(0,5·Su mín ; 690) [MPa]')}"),
     "A: σa = Kf·|σ0| [MPa]": f"={kf}*ABS({g('σ0 = M·c/I [MPa]')})",
     "A: τm = τT [MPa]": f"={g('τT = T·c/J [MPa]')}",
     "A: σ'a [MPa]": f"={g('A: σa = Kf·|σ0| [MPa]')}",
     "A: σ'm [MPa]": f"=SQRT(({N['wu']}*{x}*({L}-{x})/2*{c_}/{I})^2+3*{g('A: τm = τT [MPa]')}^2)",
     "B: τa = Kf·τV [MPa]": f"={kf}*{g('τV = V·Q/(I·b) [MPa]')}",
     "B: σ'a [MPa]": f"=SQRT(3)*{g('B: τa = Kf·τV [MPa]')}",
     "B: σ'm [MPa]": f"=SQRT(3)*(ABS({g('A: τm = τT [MPa]')})+ABS({N['wu']}*({L}/2-{x}))*{Q}/({I}*{bs}))",
     "A: CS Goodman": f"=IFERROR(1/({sAa}/{se}+{sAm}/{sumin}),9999)",
     "A: CS Soderberg": f"=IFERROR(1/({sAa}/{se}+{sAm}/{symin}),9999)",
     "A: CS Gerber": (f"=IFERROR(IF({sAm}=0,{se}/{sAa},(-({sAa}/{se})+SQRT(({sAa}/{se})^2+4*({sAm}/{sumin})^2))"
                      f"/(2*({sAm}/{sumin})^2)),9999)"),
     "B: CS Goodman": f"=IFERROR(1/({sBa}/{se}+{sBm}/{sumin}),9999)",
     "CS fatiga Goodman (mín A, B)": f"=MIN({g('A: CS Goodman')},{g('B: CS Goodman')})",
    }
    kfl = f"IF({N['swfl']}=1,{kf},1)"
    F_["σ'máx A (fluencia) [MPa]"] = f"=SQRT(({kfl}*ABS({g('σ0 = M·c/I [MPa]')})+{N['wu']}*{x}*({L}-{x})/2*{c_}/{I})^2+3*{g('τT = T·c/J [MPa]')}^2)"
    F_["σ'máx B (fluencia) [MPa]"] = f"=SQRT(3)*({kfl}*{g('τV = V·Q/(I·b) [MPa]')}+ABS({g('τT = T·c/J [MPa]')})+ABS({N['wu']}*({L}/2-{x}))*{Q}/({I}*{bs}))"
    F_["CS fluencia = Sy mín / σ'máx"] = (f"=IFERROR({symin}/MAX({g(chr(963)+chr(39)+'máx A (fluencia) [MPa]')},"
                                          f"{g(chr(963)+chr(39)+'máx B (fluencia) [MPa]')}),9999)")
    va = g("Vida: a = (0,9·Su mín)²/Se [MPa]"); vc = g("Vida: c = -1/b = 3/log(0,9·Su mín/Se)")
    F_["Vida: a = (0,9·Su mín)²/Se [MPa]"] = f"=(0.9*{sumin})^2/{se}"
    F_["Vida: c = -1/b = 3/log(0,9·Su mín/Se)"] = f"=3/LOG10(0.9*{sumin}/{se})"
    sn_ = g("Vida: SN = σ'a/(1 - σ'm/Su mín) [MPa]")
    F_["Vida: SN = σ'a/(1 - σ'm/Su mín) [MPa]"] = (f"=IF({g('CS fatiga Goodman (mín A, B)')}>=1,\"Vida infinita\","
        f"IF({g('A: CS Goodman')}<={g('B: CS Goodman')},{sAa}/(1-{sAm}/{sumin}),{sBa}/(1-{sBm}/{sumin})))")
    F_["Vida N [ciclos]"] = (f"=IF(ISNUMBER({sn_}),IF({sn_}>0.9*{sumin},\"< 10^3: fuera de rango\",({va}/{sn_})^{vc}),\"Vida infinita\")")
    vn_ = g("Vida N [ciclos]")
    F_["Vida [horas]"] = f"=IF(ISNUMBER({vn_}),{vn_}/({N['n']}*60),\"-\")"
    for lab, fml in F_.items():
        cell = sw[g(lab)]; cell.value = fml; cell.font = fb; cell.number_format = "0.00"
    for lab in ("V [N]", "M [N·mm]", "T [N·mm]", "I = π/64·(D⁴ - d⁴) [mm⁴]", "J = 2I [mm⁴]", "Vida N [ciclos]"):
        sw[g(lab)].number_format = "#,##0"
    for lab in ("Su tubo [MPa]", "Su muñón [MPa]", "Su electrodo [MPa]", "Su mínimo [MPa]",
                "Sy tubo [MPa]", "Sy muñón [MPa]", "Sy electrodo [MPa]", "Sy mínimo [MPa]", "x [mm]"):
        sw[g(lab)].number_format = "0"
    for lab in ("Kf soldadura", "fS = a·(Su mín)^b", "fT", "fC", "f1", "Vida: c = -1/b = 3/log(0,9·Su mín/Se)"):
        sw[g(lab)].number_format = "0.000"
    for lab in ("x [mm]", "Garganta t = espesor del tubo [mm]", "I = π/64·(D⁴ - d⁴) [mm⁴]", "J = 2I [mm⁴]",
                "Su tubo [MPa]", "Su muñón [MPa]", "Sy tubo [MPa]", "Sy muñón [MPa]", "fT", "f1"):
        sw[g(lab)].font = fgreen
sw.freeze_panes = "B6"

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
rline("achk", "a = x_tope - B (debe ser ≥ 0)", f"={N['a']}", "mm", fmt="0.0")
rline("bchk", "b = L - a - B (debe ser ≥ 0)", f"={N['b']}", "mm", fmt="0.0")
rline("geo", "Control: el manchón entra entre apoyos", f'=IF(OR({N["a"]}<0,{N["b"]}<0),"ERROR: el manchón no entra entre apoyos (a < 0 o b < 0)",IF({N["a"]}<{N["xg0"]},"AVISO: el manchón empieza antes del engomado","OK"))', "", "Si da ERROR, todos los resultados carecen de sentido")
rs[RS["geo"]].font = Font(name=F, size=10, bold=True, color="C00000")
rsec("Cargas y esfuerzos")
rline("R1y", "Reacción R1y / R1z", f"={N['R1y']}", "N", fmt="#,##0")
rs.cell(rr - 1, 3).value = f"={N['R1z']}"; rs.cell(rr - 1, 3).number_format = "#,##0"; rs.cell(rr - 1, 4).value = "Columna B: plano y, columna C: plano z [N]"
rline("R2y", "Reacción R2y / R2z", f"={N['R2y']}", "N", fmt="#,##0")
rs.cell(rr - 1, 3).value = f"={N['R2z']}"; rs.cell(rr - 1, 3).number_format = "#,##0"
rline("R1", "Reacción resultante rodamiento izquierdo", f"={N['R1']}", "N", "Dato para verificar el rodamiento", fmt="#,##0")
rline("R2", "Reacción resultante rodamiento derecho", f"={N['R2']}", "N", fmt="#,##0")
rline("Fu2", "Fuerza rotante por desbalanceo en cada rodamiento", f"={N['Fu']}/2", "N", "* F_u/2, gira con el eje (carga rotante para el rodamiento)", fmt="0.0")
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
rsec("Soldaduras a tope tubo-muñón (hoja Soldadura)")
for k, lado in (("B", "izquierda"), ("C", "derecha")):
    rline(f"sx{k}", f"Soldadura {lado}: posición x", f"=Soldadura!{k}{SR['x [mm]']}", "mm", fmt="0.0")
    rline(f"scs{k}", f"Soldadura {lado}: CS fatiga Goodman", f"=Soldadura!{k}{SR['CS fatiga Goodman (mín A, B)']}", "-",
          "Ec. (44)-b con Su mínimo (tubo, muñón, electrodo) y Kf de la soldadura")
    rline(f"sfl{k}", f"Soldadura {lado}: CS fluencia", f"=Soldadura!{k}{SR[chr(67)+'S fluencia = Sy mín / '+chr(963)+chr(39)+'máx']}", "-",
          "Ec. (77) con Sy mínimo")
    rline(f"sv{k}", f"Soldadura {lado}: vida", f"=Soldadura!{k}{SR['Vida N [ciclos]']}", "ciclos", fmt="#,##0")
rsec("Gobierna")
rline("csg", "CS fatiga mínimo (cuerpo, concentradores y soldaduras)",
      f"=MIN({RS['csb']},{RS['cs0']},{RS['cs1']},{RS['cs2']},{RS['scsB']},{RS['scsC']})", "-",
      "* Si CS < 1: vida limitada, ver hojas Concentradores / Tensores / Soldadura")
rline("flg", "CS fluencia mínimo", f"=MIN({RS['flb']},{RS['fl0']},{RS['fl1']},{RS['fl2']},{RS['sflB']},{RS['sflC']})", "-")
def srev(sa, sm): return f"{sa}/(1-{sm}/{Su})"
cands = [srev("Tensores!$G$11", "Tensores!$C$11"), srev("Tensores!$S$11", "Tensores!$O$11")]
for k in "BCD":
    on = f"Concentradores!{k}{R['Activa']}"
    cands.append(f"{on}*" + srev(f"Concentradores!{k}{R[chr(65)+': σ'+chr(39)+'a [MPa]']}", f"Concentradores!{k}{R[chr(65)+': σ'+chr(39)+'m [MPa]']}"))
    cands.append(f"{on}*" + srev(f"Concentradores!{k}{R['B: σ'+chr(39)+'a [MPa]']}", f"Concentradores!{k}{R['B: σ'+chr(39)+'m [MPa]']}"))
rline("srev", "Tensión alternante equivalente máxima (σm = 0)", "=MAX(" + ",".join(cands) + ")", "MPa",
      "σa/(1 - σ'm/Su), ec. (48)-b. Es la que se ubica en el diagrama S-N. No incluye las soldaduras (tienen otro Se)")
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

def block(r0_, title, sig_a, tau_m, tau_aB, kfl_sig, kfl_tauB, note, sig_m="0", tau_mB=None):
    tau_mB = tau_mB or tau_m
    ts.cell(r0_, 1, title).font = Font(name=F, size=11, bold=True)
    ts.cell(r0_ + 1, 1, note).font = fsmall
    base = r0_ + 2
    mat(base, 1, "A medio", sig_m, tau_m, "σ'm =", f"SQRT(({sig_m})^2+3*({tau_m})^2)")
    mat(base, 5, "A alternante", sig_a, "0", "σ'a =", f"ABS({sig_a})")
    smax = f"(ABS({kfl_sig})+{sig_m})"
    mat(base, 9, "A máx (fluencia)", smax, tau_m, "σ'máx =", f"SQRT({smax}^2+3*({tau_m})^2)")
    mat(base, 13, "B medio", "0", tau_mB, "σ'm =", f"SQRT(3)*ABS({tau_mB})")
    mat(base, 17, "B alternante", "0", tau_aB, "σ'a =", f"SQRT(3)*ABS({tau_aB})")
    return base + 6

# body critical row
idx = f"MATCH(MIN({rng('R')}),{rng('R')},0)"
pick = lambda col: f"INDEX({rng(col)},{idx})"
r_ = 4
ts.cell(r_, 1, "Cuerpo: posición crítica x [mm] =").font = fb
ts.cell(r_, 5, "=" + pick("A")).number_format = "0.0"
r_ = block(r_ + 1, "Cuerpo del tubo, sección crítica a fatiga (Kf = 1)", pick("I"), pick("J"), pick("K"),
           pick("I"), pick("K"), "Kf = 1 (tubo entero). Valores tomados de la fila de CS mínimo de la hoja 'Cálculo x'",
           pick("AF"), f"({pick('J')}+{pick('AG')})")
# body life
ts.cell(r_, 1, "Vida del cuerpo:").font = fb
sa_b = f"MAX(ABS({pick('I')}),SQRT(3)*{pick('K')})"
ts.cell(r_, 5, f"=IF(MIN({rng('R')})>=1,\"Vida infinita\",IFERROR(((0.9*{pick('AO')})^2/{pick('AN')}/(ABS({pick('I')})/(1-Tensores!$C$11/{pick('AO')})))^(3/LOG10(0.9*{pick('AO')}/{pick('AN')})),\"revisar\"))").number_format = "#,##0"
ts.cell(r_, 8, "Ec. (48)-b y (32), punto A").font = fsmall
r_ += 2
for j, k in enumerate("BCD"):
    gc = lambda lab: f"Concentradores!{k}{R[lab]}"
    kfl_s = f"IF({N['swfl']}=1,{gc('Kf = 1 + q·(Kt - 1)')},1)*{gc('σ0 = M·c/I [MPa]')}"
    r_ = block(r_, f"Sección {j+1} con concentrador", gc("A: σa = Kf·|σ0| [MPa]"), gc("A: τm = τT [MPa]"),
               gc("B: τa = Kfs·τV [MPa]"), kfl_s, gc("B: τa = Kfs·τV [MPa]"),
               f"x según Datos. Kf solo en la alternante (Tabla 14). Si la sección está inactiva, ignorar este bloque.",
               gc("σm desbalanceo [MPa]"), f"({gc('A: τm = τT [MPa]')}+{gc('τm desbalanceo (corte) [MPa]')})")
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
gs["W1"] = "Datos auxiliares de los gráficos (marcas de zonas): no editar"; gs["W1"].font = fsmall
_aux = [0]
def chart(col, title, ytitle, anchor):
    ch = ScatterChart(); ch.title = title; ch.style = 13
    ch.x_axis.title = "x [mm]"; ch.y_axis.title = ytitle
    ch.height = 8.5; ch.width = 16
    xs = Reference(cs, min_col=1, min_row=r0, max_row=rN)
    colidx = column_index_from_string(col)
    ys = Reference(cs, min_col=colidx, min_row=r0, max_row=rN)
    s = Series(ys, xs, title=title); s.marker.symbol = "none"; s.graphicalProperties.line.width = 15000
    s.graphicalProperties.line.solidFill = "000000"
    ch.series.append(s)
    # marcas: límites tubo/muñones, manchón y apoyos (se actualizan con los datos)
    b0 = 3 + _aux[0] * 14; _aux[0] += 1
    ymin = f"MIN(0,{rng(col)})"; ymax = f"MAX(0,{rng(col)})"
    gs.cell(b0, 23, title).font = fsmall
    mf = f"{N['a']}+{N['cl']}"
    pts = [("si", [(N['xsI'], ymin), (N['xsI'], ymax)]), ("sd", [(N['xsD'], ymin), (N['xsD'], ymax)]),
           ("mi", [(N['a'], ymin), (N['a'], ymax)]), ("mf", [(mf, ymin), (mf, ymax)]),
           ("ap", [("0", "0"), (N['L'], "0")])]
    rowp = b0 + 1; refs = {}
    for key, lst in pts:
        r_start = rowp
        for p in lst:
            gs.cell(rowp, 23, "=" + p[0]); gs.cell(rowp, 24, "=" + p[1]); rowp += 1
        refs[key] = (r_start, rowp - 1)
    for key, tit, color, dash, mark in (("si", "Soldadura izq.", "7F7F7F", None, None),
                                        ("sd", "Soldadura der.", "7F7F7F", None, None),
                                        ("mi", "Manchón inicio", "C00000", "dash", None),
                                        ("mf", "Manchón tope", "C00000", "dash", None),
                                        ("ap", "Apoyos", "000000", None, "triangle")):
        r1_, r2_ = refs[key]
        se_ = Series(Reference(gs, min_col=24, min_row=r1_, max_row=r2_), Reference(gs, min_col=23, min_row=r1_, max_row=r2_), title=tit)
        if mark:
            se_.marker.symbol = mark; se_.marker.size = 9
            se_.marker.graphicalProperties.solidFill = color; se_.marker.graphicalProperties.line.solidFill = color
            se_.graphicalProperties.line.noFill = True
        else:
            se_.marker.symbol = "none"; se_.graphicalProperties.line.solidFill = color
            se_.graphicalProperties.line.width = 12000
            if dash: se_.graphicalProperties.line.dashStyle = dash
        ch.series.append(se_)
    ch.legend.position = "b"
    ch.x_axis.delete = False; ch.y_axis.delete = False
    gs.add_chart(ch, anchor)
chart("D", "Corte resultante V(x)", "N", "A1"); chart("G", "Momento flector resultante M(x)", "N·mm", "K1")
chart("H", "Torque T(x)", "N·mm", "A19"); chart("AC", "Flecha total", "mm", "K19")
chart("I", "σ0 nominal (fibra extrema)", "MPa", "A37"); chart("R", "CS fatiga del eje (Kf = 1)", "-", "K37")
chart("E", "My(x) plano y (manchón)", "N·mm", "A55"); chart("F", "Mz(x) plano z (gravedad)", "N·mm", "K55")
gs["A73"] = ("Línea gris: soldaduras (a la izquierda y derecha de ellas están los muñones). Línea roja: extensión del manchón "
             "(entra por la izquierda y hace tope a la derecha). Triángulos: apoyos (centros de rodamientos).")
gs["A73"].font = fsmall

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

for col, wdt in (("E", 3), ("F", 18), ("G", 12), ("H", 30)):
    fs.column_dimensions[col].width = wdt
fs["F3"] = "Electrodos: metal de aporte, valores mínimos"; fs["F3"].font = fbold
for j, h in enumerate(["Electrodo AWS", "Su [MPa]", "Sy [MPa]"]):
    fs.cell(4, 6 + j, h).font = fbold
for i, (e_, su_, sy_) in enumerate([("E60xx", 427, 345), ("E70xx", 482, 393), ("E80xx", 551, 462),
                                     ("E90xx", 620, 531), ("E100xx", 689, 600), ("E120xx", 827, 737)]):
    fs.cell(5 + i, 6, e_).font = fb
    fs.cell(5 + i, 7, su_).font = fin; fs.cell(5 + i, 8, sy_).font = fin
fs["F11"] = "Fuente: Shigley 10ª ed., Tabla 9-3 (el compendio, cap. 18 Anexo 2, cita la misma fuente)"; fs["F11"].font = fsmall
fs["F13"] = "Soldadura a tope: Kf y terminación para fS"; fs["F13"].font = fbold
for j, h in enumerate(["Sobremonta", "Kf", "Terminación para fS"]):
    fs.cell(14, 6 + j, h).font = fbold
for i, (s_, k_, t_) in enumerate([("Con sobremonta", 1.2, "Forjado"), ("Maquinada al ras", 1.0, "Maquinado o estirado en frío")]):
    fs.cell(15 + i, 6, s_).font = fb
    fs.cell(15 + i, 7, k_).font = fin; fs.cell(15 + i, 8, t_).font = fin
fs["F17"] = "Kf 1,2: Shigley Tabla 9-5 (soldadura a tope reforzada). Forjado: Shigley sec. 9-7"; fs["F17"].font = fsmall
fs["F18"] = "* Maquinada al ras: Kf = 1 es supuesto (el Shigley no da valor; el compendio cap. 18 indica maquinar la sobremonta)"; fs["F18"].font = fsmall

lines = [
("ESQUEMA DE USO", ""),
("1. Datos", "Cargar en la hoja Datos: geometría (D, t, L), manchón (cantidad, ancho, masa, desarrollo, frecuencia, inicio a), pesos (z), torque (P, n, lado), material (Su, Sy, E, ρ), terminación y f1."),
("2. Concentradores", "Para cada escalón, chavetero o ranura: posición x, Kt y Kts del Anexo 2 y radio r. Desactivar las secciones que no se usen."),
("3. Revisar DCL", "Controlar que el control de equilibrio dé 0 y que b ≥ 0 (hoja Datos), y los diagramas My, Mz, M, T (hoja Gráficos)."),
("4. Resultados", "Hoja Resultados: CS de fatiga y de fluencia del cuerpo y de cada sección, y flecha máxima."),
("5. Tensores", "Hoja Tensores: tensor medio, alternante y máximo en los puntos A y B de cada sección."),
("6. Vida", "Si algún CS de fatiga es < 1: vida limitada en ciclos y horas (hoja Concentradores)."),
("7. Soldaduras", "Datos > 'Soldaduras a tope tubo-muñón': posición x de cada soldadura, Su y Sy del muñón, electrodo y sobremonta. Resultado en hoja Soldadura y en Resultados."),
("8. Balanceo", "Datos: grado G, muñones (diámetros y largo) y engomado (inicio y largo). Resultado en Datos: M_eje, M_goma, x_G, e_per, U_per total y por apoyo, F_u."),
("", ""),
("HIPÓTESIS", ""),
("Modelo de viga", "Simplemente apoyada en los centros de los rodamientos (los rodamientos permiten el giro de la sección)."),
("Geometría", "Vista de frente: a la izquierda el muñón que entra en el apoyo, en el medio el tubo (entre soldaduras) y a la derecha el muñón motriz. L es la distancia entre centros de rodamientos; entre 0 y la soldadura izquierda, y entre la soldadura derecha y L, la sección es la del muñón."),
("Manchón", "Las correas entran desde la izquierda y hacen tope a la derecha: el manchón va de x_tope - B a x_tope (x_tope = fin del engomado si no se carga)."),
("Sección por zona", "I, J, c, Q, b, Su, Sy y Se se toman según la zona de cada x (muñón izq., tubo, muñón der.). Muñones: material y terminación propios; fT con su diámetro."),
("Pesos por tramo", "Muñones, tubo y engomado se cargan cada uno en su tramo (plano z). La punta de los muñones fuera de los apoyos no se considera *."),
("Ejes", "x a lo largo del rullo (desde el apoyo izquierdo), y dirección de la tensión del manchón, z dirección de la gravedad."),
("Carga y *", "Manchón: qy = F/B uniforme entre x = a y x = a + B. F = 2·T total (ramales paralelos). T por correa = 4·m·Lt²·f² (cuerda vibrante, no está en el compendio)."),
("Carga z *", "Peso propio del tubo (todo L) y mitad del peso del manchón repartida en B (opcionales)."),
("Engomado *", "Suma peso (plano z) y define el diámetro de apoyo del manchón (D + 2e). Su rigidez no se cuenta (E goma << E acero)."),
("Desbalanceo *", "F_u = M·e·ω², e = e_per (ISO 21940-11) o excentricidad medida. Repartido uniforme en L: Mu = wu·x·(L-x)/2. Gira con el eje: entra como tensión media (σm en A, τm en B), sin Kf, en cuerpo, concentradores y soldaduras. En fluencia: σmáx = Kf·σa + σm."),
("Masa que gira", "M = M_eje + M_goma. M_eje = ρ·(A·L_tubo + A_mI·L_mI + A_mD·L_mD), tubo entre soldaduras y muñones hacia afuera. M_goma = ρg·Ag·L_g."),
("Desbalance admisible", "e_per = 1000·G/ω [g·mm/kg] ; U_per = e_per·M [g·mm] (ISO 21940-11, rotor rígido). Un solo G para todo el rotor."),
("Reparto por apoyo *", "U_per,izq = U_per·(L - x_G)/L ; U_per,der = U_per·x_G/L (palanca respecto del centro de masa). Verificar contra la norma."),
("n_c y w *", "La velocidad crítica y el peso propio toman tubo + goma en todo L (sin muñones ni largo real del engomado)."),
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
("Soldadura a tope", "Circunferencial, penetración completa: la sección resistente es la corona del tubo (compendio cap. 18, ec. (1): garganta = espesor de las partes, sin sobremonta)."),
("Soldadura: materiales", "Se compara con el menor Su y Sy entre tubo, muñón y electrodo (compendio cap. 18: verificar aporte y materiales base)."),
("Soldadura: Kf y fS", "Con sobremonta: Kf = 1,2 (Shigley Tabla 9-5) y fS de forjado (Shigley sec. 9-7). Maquinada al ras: Kf = 1 * (supuesto) y fS de maquinado."),
("Soldadura: fatiga", "Mismo procedimiento que el tubo (sección 7.1 del compendio): Kf solo en la alterna (Tabla 14), Von Mises (72), Goodman (44)-b."),
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
("", "Kt de tablas solo para geometrías estándar (Anexo 2)."),
("", "Soldaduras: solo a tope tubo-muñón con penetración completa. No cubre penetración parcial (raíz sin fundir), filete, flexión local por cambio de rigidez ni tensiones residuales (FEA: PrePoMax)."),
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
        ("vFu","Fu"),("vrFu","rFu"),("vub","ub"),("vexc","ee"),("vwu","wu"),("veu","eu"),
        ("vDextmI","DmI"),("vDintmI","dmI"),("vLmI","LmI"),("vDextmD","DmD"),("vDintmD","dmD"),("vLmD","LmD"),
        ("vxgcero","xg0"),("vLg","Lg"),("vAmI","AmI"),("vAmD","AmD"),("vLtubo","Ltu"),("vMeje","Meje"),("vMgoma","Mgoma"),
        ("vxG","xG"),("vxtp","xtp"),("vxtope","xtope"),("vwmI","wmI"),("vwmD","wmD"),("vwg","wg"),("vSemI","SemI"),("vSemD","SemD"),("vtermM","termM"),("vomega","om"),("veper","eper"),("vUper","Uper"),("vUperI","UperI"),("vUperD","UperD")]
for ln, k in dmap:
    nm(ln, N[k])
for ln, k in [("vsrev","srev"),("vMmax","Mmax"),("vxM","xM"),("vsmax","smax"),("vfl","fl"),("vxfl","xfl"),("vcsb","csb"),("vxcsb","xcsb"),
              ("vflb","flb"),("vcsg","csg"),("vflg","flg")]:
    nm(ln, f"Resultados!${RS[k][0]}${RS[k][1:]}")
for ln, k in [("vxsI","xsI"),("vxsD","xsD"),("vSuM","SuM"),("vSyM","SyM"),("velec","elec"),("vsobre","sobre")]:
    nm(ln, N[k])
for k, lado in (("B", "izq"), ("C", "der")):
    for ln, lab in [("sKf", "Kf soldadura"), ("sSe", "Se = fS·fT·fC·f1·Se' [MPa]"), ("ssa", "A: σ'a [MPa]"),
                    ("ssm", "A: σ'm [MPa]"), ("scs", "CS fatiga Goodman (mín A, B)"),
                    ("sfl", "CS fluencia = Sy mín / σ'máx"), ("svida", "Vida N [ciclos]")]:
        nm(f"v{ln}{lado}", f"Soldadura!${k}${SR[lab]}")
for ln, cell in [("vtsigm","A8"),("vttaumB","O8"),("vtvmmB","O11"),("vtxcrit","E4"),("vttaum","C8"),("vtsiga","E8"),("vtvma","G11"),("vtvmm","C11"),("vtvmmax","K11"),
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
