# Rullo PSV

Verificación estructural (estática y fatiga) del rullo delantero (motriz) del torno PSV, Dayco.

Fuente de cálculo: *Compendio de Cálculo Estructural II*, FCEFyN-UNC 2024, cap. 14 (Falla por fatiga).

## Contenido

| Carpeta | Archivo | Qué es |
|---|---|---|
| `planilla/` | `Rullo_PSV_estatico_fatiga.xlsx` | Planilla de cálculo (estático, fatiga, flecha, tensores) |
| `scripts/` | `build.py` | Genera la planilla (openpyxl) |
| `scripts/` | `dcl.py`, `dcl.png` | Diagrama de cuerpo libre insertado en la planilla |

Para regenerar la planilla: `cd scripts && python3 dcl.py && python3 build.py`. Después hay que abrirla en Excel o recalcularla en LibreOffice para que aparezcan los valores.

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

## Limitaciones

- No cubre uniones soldadas ni atornilladas (muñones).
- Fatiga de ciclos altos (N > 10³); sin daño acumulado por arranques y paradas.
- Kt solo para geometrías estándar del Anexo 2.
- Los valores cargados en la planilla son un **ejemplo**, no son datos del rullo real.
