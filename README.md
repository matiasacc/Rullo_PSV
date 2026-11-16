# Rullo PSV

Verificación estructural (estática y fatiga) del rullo delantero (motriz) del torno PSV, Dayco.

Fuente de cálculo: *Compendio de Cálculo Estructural II*, FCEFyN-UNC 2024, cap. 14 (Falla por fatiga).

## Contenido

| Carpeta | Archivo | Qué es |
|---|---|---|
| `planilla/` | `Rullo_PSV_estatico_fatiga.xlsx` | Planilla de cálculo (estático, fatiga, flecha, tensores) |
| `scripts/` | `build.py` | Genera la planilla (openpyxl) |
| `scripts/` | `dcl.py`, `dcl.png` | Diagrama de cuerpo libre insertado en la planilla |

| `memoria/` | `memoria.tex`, `memoria.pdf` | Memoria de cálculo en LaTeX, toma los valores de la planilla |
| `memoria/` | `revisiones.tex` | Registro de revisiones (se completa a mano) |
| `scripts/` | `exportar_datos.py` | Lee la planilla y escribe `memoria/datos.tex` y `memoria/figuras/diagramas.pdf` |

## Actualizar la memoria

1. Cambiar datos en la planilla y **guardarla desde Excel** (así quedan los valores calculados).
2. Agregar una fila en `memoria/revisiones.tex` si el cambio es relevante.
3. Subir los cambios (`git push`). La acción de GitHub *Memoria de cálculo* exporta los datos, compila el LaTeX y sube `memoria/memoria.pdf` actualizado.

En local: `python3 scripts/exportar_datos.py` y después `cd memoria && latexmk -pdf memoria.tex`.

Logo: colocar `memoria/figuras/logo_dayco.png` (si no está, el encabezado muestra el texto DAYCO).

Para regenerar la planilla desde cero (borra los datos cargados): `cd scripts && python3 dcl.py && python3 build.py`, mover el archivo a `planilla/` y abrirlo en Excel. Los valores de la planilla se exportan por nombres definidos (`v...`): no borrarlos.

## Modelo

- Viga simplemente apoyada, apoyos en los extremos (conservativo).
- Geometría longitudinal en un solo bloque de Datos (x = 0 en el extremo izquierdo):
  `X1` eje izq. Ø30 | `X2` cono (misma pieza) | soldadura filete | tubo | soldadura al ras = tope de las correas | `X3` disco | `X4` eje motriz.
  Se cargan **Largo de la pieza (X1+X2+X_tubo+X3), X1, X2, X3, X4**; `X_tubo = Largo − (X1+X2+X3)`, `L = Dist. total = Largo + X4` y las posiciones de soldaduras y secciones salen solas.
- Ejes: **x** a lo largo del rullo (desde el extremo izquierdo), **y** dirección de la tensión del manchón, **z** gravedad.
- Plano y: manchón como carga distribuida `qy = 2·n·T / B` entre `x = a` y `x = a + B`, con `B = n · ancho de correa`.
- Plano z: peso propio del tubo (todo L) y mitad del peso del manchón en B (opcionales).
- Combinación: `M = √(My² + Mz²)`, `V = √(Vy² + Vz²)` (sección circular).
- El eje gira y las cargas tienen dirección fija: la flexión es completamente alternante (1 ciclo por vuelta).

## Esquema de uso

1. **Datos**: geometría (Dist. total, X1…X4; el gráfico de segmentos muestra tramos, secciones S1–S3, soldaduras y ▲/◆ en el punto más crítico a fatiga/fluencia), manchón (cantidad, ancho, masa, desarrollo, frecuencia, inicio a), pesos, torque, material, terminación.
2. **Concentradores** (en Datos): posición x, Kt y Kts (Anexo 2) y radio r de cada escalón o chavetero.
3. Revisar el equilibrio (control = 0) y que `b ≥ 0`.
4. **Resultados**: CS de fatiga y de fluencia (eje completo y mínimo con su x en cada tramo X1, X2, tubo, X3, X4), reacciones en rodamientos, flecha.
5. **Tensores**: tensor medio, alternante y máximo en los puntos A (fibra extrema) y B (eje neutro).
6. Si CS de fatiga < 1: vida limitada en ciclos y horas (hoja Concentradores).
7. **Soldaduras** (en Datos): tipo de cada lado (a tope / filete), posición x, electrodo (lista); a tope: garganta y sobremonta; filete: cateto w y radio del cordón. Resultado en las filas *adoptado* de la hoja **Soldadura** y en Resultados.
8. **Balanceo** (en Datos): grado G, muñones (diámetro exterior, interior y largo) y engomado (inicio y largo). Resultado en Datos: M_eje, M_goma, x_G, e_per, U_per total y por apoyo, F_u.

## Procedimiento de fatiga (compendio, sección 7.1)

1. Cargas medias y alternas de cada esfuerzo (63).
2. Tensor medio y tensor alternante (64), (65). Kf solo sobre la alternante en material dúctil (Tabla 14).
3. Von Mises de cada tensor (72).
4. Goodman (44)-b; también Soderberg (44)-a y Gerber (44)-c.
5. Fluencia por separado (77).
6. Vida limitada: (48)-b y (32).

`Se = fT·fC·fS·f1·Se'` (20), con `Se' = 0,5·Su ≤ 690 MPa` (17), fT (21), fC = 1 (24), fS (25)-(26).

## Supuestos a validar (*)

- Tensión por correa `T = 4·m·Lt²·f²` (cuerda vibrante / tensiómetro sónico). No está en el compendio.
- `Lt = (desarrollo − π·D)/2` si no se carga a mano: ramales paralelos, rullos de igual diámetro, línea primitiva ≈ D exterior.
- Mitad del peso del manchón sobre cada rullo.
- Torque: entra por un extremo y lo toman las correas de forma pareja en B; constante en el tiempo.
- Apoyo izquierdo en x = 0, borde exterior del asiento (real ≈ X1/2): conservativo.
- Apoyo derecho al final de X4. **X4 pendiente de medición**: se carga la cota superior (más corta que la real es inseguro).
- Disco (X3) macizo, Ø como dato (0 = D del tubo), con material y terminación del muñón derecho; eje derecho (X4) con su Ø.
- Muñón izquierdo (Ø30 de 0 a 40 mm + cono de 40 a 80 mm) modelado como cilindro Ø30 en todo el largo: conservativo. Cerca de x = 80 el CS del cuerpo sale bajo por este supuesto.
- Kt de los escalones Ø30: Shigley Tabla 7-1, escalón agudo r/d = 0,02 (Kt = 2,7; Kts = 2,2). Fuera del compendio: confirmar con el Anexo 2. Aplicado sobre una transición cónica: conservativo.
- Kfs aplicado al corte transversal en el eje neutro (conservativo).
- fT con el diámetro exterior del tubo.
- Kf en la verificación a fluencia: selector en Datos (ec. 3 frente a ec. 77). Pendiente de definir.
- Coeficientes de Neuber (10)-a/b leídos del texto extraído del PDF: verificar contra el impreso.
- Soldadura maquinada al ras: Kf = 1 (el Shigley no da valor; el compendio, cap. 18, indica maquinar la sobremonta).
- Soldadura: Kf aplicado al corte transversal en el eje neutro (mismo criterio que en Concentradores).
- Filete: Kf = 1,5 (pie de filete transversal, cordón continuo); τM y τV con su máximo a la vez; radio del cordón = d/2 si no se carga.
- Balanceo: un solo G para todo el rotor (eje + engomado); reparto de U_per en los apoyos por palanca respecto del centro de masa (verificar contra ISO 21940-11).
- Velocidad crítica y peso propio: tubo + goma en todo L (sin muñones ni largo real del engomado).

## Soldaduras tubo-muñón (hoja Soldadura)

Cada lado se elige a tope o filete. A tope:


- Circunferencial, penetración completa: sección resistente = corona del tubo; garganta = espesor t, sin sobremonta (compendio cap. 18, ec. 1-2).
- Tensiones: σ0 = M·c/I, τT = T·c/J, τV = V·Q/(I·b), con M, V, T del mismo modelo de viga en la posición x de cada soldadura.
- Materiales: se toma el menor Su y Sy entre tubo, muñón y electrodo (compendio cap. 18). Electrodos: Shigley Tabla 9-3.
- Con sobremonta: Kf = 1,2 (Shigley Tabla 9-5) y fS de forjado (Shigley sec. 9-7). Maquinada al ras: Kf = 1 (*) y fS de maquinado.
- Fatiga con el mismo procedimiento del compendio (sección 7.1) que el resto de la planilla.

Filete (método simplificado del compendio, cap. 18):

- Todas las cargas como corte en la garganta t = 0,707·w (ec. 3); soldadura como línea circular de radio r: A = 1,414·π·w·r, I = 0,707·w·π·r³, J = 0,707·w·2π·r³ (Shigley Tablas 9-1 y 9-2).
- τM = M·r/I y τV = V/A alternantes; τT = T·r/J y desbalanceo como media; composición vectorial (ec. 14).
- Estático: τadm = mín(0,3·Su ; 0,4·Sy) del electrodo (AISC, Anexo 2), CS = τadm/τtotal (ec. 15); y Von Mises √3·τ frente a Sy mínimo.
- Fatiga: Kf = 1,5 (Shigley T9-5), fS de forjado, σ' = √3·τ, Goodman (44)-b.

## Limitaciones

- No cubre uniones atornilladas (muñones) ni la raíz sin fundir de la soldadura a tope con penetración parcial.
- Soldaduras: no incluye la flexión local por el cambio de rigidez ni las tensiones residuales (verificar con FEA).
- Fatiga de ciclos altos (N > 10³); sin daño acumulado por arranques y paradas.
- Kt solo para geometrías estándar del Anexo 2.
- Los valores cargados en la planilla son un **ejemplo**, no son datos del rullo real.
