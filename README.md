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

- Viga simplemente apoyada en los centros de los rodamientos.
- Ejes: **x** a lo largo del rullo (desde el apoyo izquierdo), **y** dirección de la tensión del manchón, **z** gravedad.
- Plano y: manchón como carga distribuida `qy = 2·n·T / B` entre `x = a` y `x = a + B`, con `B = n · ancho de correa`.
- Plano z: peso propio del tubo (todo L) y mitad del peso del manchón en B (opcionales).
- Combinación: `M = √(My² + Mz²)`, `V = √(Vy² + Vz²)` (sección circular).
- El eje gira y las cargas tienen dirección fija: la flexión es completamente alternante (1 ciclo por vuelta).

## Esquema de uso

1. **Datos**: geometría, manchón (cantidad, ancho, masa, desarrollo, frecuencia, inicio a), pesos, torque, material, terminación.
2. **Concentradores** (en Datos): posición x, Kt y Kts (Anexo 2) y radio r de cada escalón o chavetero.
3. Revisar el equilibrio (control = 0) y que `b ≥ 0`.
4. **Resultados**: CS de fatiga y de fluencia, reacciones en rodamientos, flecha.
5. **Tensores**: tensor medio, alternante y máximo en los puntos A (fibra extrema) y B (eje neutro).
6. Si CS de fatiga < 1: vida limitada en ciclos y horas (hoja Concentradores).
7. **Soldaduras** (en Datos): posición x de cada soldadura a tope tubo-muñón, Su y Sy del muñón, electrodo (lista) y sobremonta (con sobremonta / maquinada al ras). Resultado en la hoja **Soldadura** y en Resultados.

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
- Kfs aplicado al corte transversal en el eje neutro (conservativo).
- fT con el diámetro exterior del tubo.
- Kf en la verificación a fluencia: selector en Datos (ec. 3 frente a ec. 77). Pendiente de definir.
- Coeficientes de Neuber (10)-a/b leídos del texto extraído del PDF: verificar contra el impreso.
- Soldadura maquinada al ras: Kf = 1 (el Shigley no da valor; el compendio, cap. 18, indica maquinar la sobremonta).
- Soldadura: Kf aplicado al corte transversal en el eje neutro (mismo criterio que en Concentradores).

## Soldaduras a tope tubo-muñón (hoja Soldadura)

- Circunferencial, penetración completa: sección resistente = corona del tubo; garganta = espesor t, sin sobremonta (compendio cap. 18, ec. 1-2).
- Tensiones: σ0 = M·c/I, τT = T·c/J, τV = V·Q/(I·b), con M, V, T del mismo modelo de viga en la posición x de cada soldadura.
- Materiales: se toma el menor Su y Sy entre tubo, muñón y electrodo (compendio cap. 18). Electrodos: Shigley Tabla 9-3.
- Con sobremonta: Kf = 1,2 (Shigley Tabla 9-5) y fS de forjado (Shigley sec. 9-7). Maquinada al ras: Kf = 1 (*) y fS de maquinado.
- Fatiga con el mismo procedimiento del compendio (sección 7.1) que el resto de la planilla.

## Limitaciones

- No cubre uniones atornilladas (muñones), soldaduras de filete ni a tope con penetración parcial.
- Soldaduras: no incluye la flexión local por el cambio de rigidez ni las tensiones residuales (verificar con FEA).
- Fatiga de ciclos altos (N > 10³); sin daño acumulado por arranques y paradas.
- Kt solo para geometrías estándar del Anexo 2.
- Los valores cargados en la planilla son un **ejemplo**, no son datos del rullo real.
