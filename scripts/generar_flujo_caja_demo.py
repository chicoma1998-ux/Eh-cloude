"""Flujo de caja mensual – Proyecto «Lote demo» (Anexo A, cabida numérica Rev. A).

Genera entregables/Flujo_Caja_Lote_Demo.xlsx con:
  Resumen · Supuestos · Flujo_Mensual · Terreno_Estrategias
Todas las cifras de cálculo son fórmulas; solo se editan las celdas amarillas (texto azul).
Uso:  python scripts/generar_flujo_caja_demo.py
"""
import json
import os
import subprocess
import sys

from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "entregables", "Flujo_Caja_Lote_Demo.xlsx")

ARIAL = "Arial"
F = Font(name=ARIAL, size=10)
F_IN = Font(name=ARIAL, size=10, color="0000FF")
F_LINK = Font(name=ARIAL, size=10, color="008000")
F_B = Font(name=ARIAL, size=10, bold=True)
F_TITLE = Font(name=ARIAL, size=14, bold=True, color="FFFFFF")
F_SEC = Font(name=ARIAL, size=10, bold=True, color="FFFFFF")
F_NOTE = Font(name=ARIAL, size=9, italic=True, color="595959")
F_RED = Font(name=ARIAL, size=10, bold=True, color="C00000")
FILL_TITLE = PatternFill("solid", fgColor="1F3864")
FILL_SEC = PatternFill("solid", fgColor="2E5496")
FILL_IN = PatternFill("solid", fgColor="FFF2CC")
FILL_TOT = PatternFill("solid", fgColor="D9E1F2")
FILL_PH = {"Inicio": "E2EFDA", "Preventa": "FCE4D6", "Construcción": "DDEBF7", "Cierre": "EDEDED"}
THIN = Side(style="thin", color="BFBFBF")
TOP = Border(top=Side(style="thin", color="000000"))

USD = '#,##0;(#,##0);"-"'
PCT = '0.0%;(0.0%);"-"'
PCT2 = '0.00%;(0.00%);"-"'
INT = '0;(0);"-"'

N_MONTHS = 23            # meses 0..22
FIRST_COL = 3            # columna C = mes 0
LAST_COL = FIRST_COL + N_MONTHS - 1
LC = get_column_letter(LAST_COL)

wb = Workbook()
names = {}


def name(ws, ref, nm):
    col, row = ref.rstrip("0123456789"), ref[len(ref.rstrip("0123456789")):]
    wb.defined_names[nm] = DefinedName(nm, attr_text=f"'{ws.title}'!${col}${row}")
    names[nm] = f"{ws.title}!{ref}"


def title(ws, text, sub, width=8):
    ws["A1"] = text
    ws["A1"].font = F_TITLE
    for c in range(1, width + 1):
        ws.cell(row=1, column=c).fill = FILL_TITLE
    ws["A2"] = sub
    ws["A2"].font = F_NOTE


def section(ws, row, text, width):
    ws.cell(row=row, column=1, value=text).font = F_SEC
    for c in range(1, width + 1):
        ws.cell(row=row, column=c).fill = FILL_SEC


def inp(cell, value, fmt=USD, note=None):
    cell.value = value
    cell.font = F_IN
    cell.fill = FILL_IN
    cell.number_format = fmt
    if note:
        cell.comment = Comment(note, "Modelo")


def calc(cell, formula, fmt=USD, bold=False):
    cell.value = formula
    cell.font = F_B if bold else F
    cell.number_format = fmt


# ════════════════════════════════════════════════════════════════════ Supuestos
S = wb.active
S.title = "Supuestos"
title(S, "SUPUESTOS · Lote demo (Escenario «CPU») · US$",
      "Fuente base: Anexo A – Cabida numérica Rev. A (19-nov-2026), «Perfil económico». "
      "Celdas amarillas con texto azul = editables. Todo lo demás se calcula.", 7)
for col, w in zip("ABCDEFG", [46, 14, 14, 14, 14, 14, 52]):
    S.column_dimensions[col].width = w

# 1. Ingresos ------------------------------------------------------------------
section(S, 4, "1. INGRESOS POR VENTAS (precios de lista, Anexo A)", 7)
for i, h in enumerate(["Producto", "Unidades", "Venta sin IGV", "IGV", "Total con IGV", "% ventas"]):
    c = S.cell(row=5, column=1 + i, value=h)
    c.font = F_B
ventas = [
    ("C · 1D 50 m²", 8, 2685183, 241667),
    ("A · 3D 85 m²", 8, 4511141, 406003),
    ("B · 3D 75 m²", 7, 3716043, 334444),
    ("Azotea dúplex A", None, 239654, 21569),
    ("Azotea dúplex C", None, 142650, 12839),
    ("Azotea dúplex B", None, 225617, 20306),
    ("Estacionamientos (16)", 16, 880734, 79266),
]
r = 6
for lbl, u, s, igv in ventas:
    S.cell(row=r, column=1, value=lbl).font = F
    if u is not None:
        inp(S.cell(row=r, column=2), u, INT)
    inp(S.cell(row=r, column=3), s)
    calc(S.cell(row=r, column=4), f"=C{r}*TasaIGV*(1-PctTerrIGV)")
    calc(S.cell(row=r, column=5), f"=C{r}+D{r}")
    calc(S.cell(row=r, column=6), f"=IFERROR(E{r}/$E$13,0)", PCT)
    r += 1
S.cell(row=13, column=1, value="Total ingresos").font = F_B
calc(S["B13"], "=SUM(B6:B8)", INT, True)
for col in "CDEF":
    calc(S[f"{col}13"], f"=SUM({col}6:{col}12)", PCT if col == "F" else USD, True)
for col in "ABCDEF":
    S[f"{col}13"].fill = FILL_TOT
S["G13"] = "Cuadra con Anexo A: US$ 13,517,115 (23 deptos, 1,744 m² vendibles)."
S["G13"].font = F_NOTE
S.cell(row=14, column=1, value="Primera venta: % del precio atribuido al terreno (no paga IGV)").font = F
inp(S["B14"], 0.50, PCT, "Reglamento de la Ley del IGV: en la primera venta de inmuebles por el constructor, "
    "la base imponible excluye el valor del terreno, que se considera el 50 % del valor de venta. "
    "Por eso el IGV efectivo es 18 % x 50 % = 9 % del precio sin IGV.")
name(S, "B14", "PctTerrIGV")
S["C14"] = "Tasa IGV"
S["C14"].font = F_B
S["C14"].alignment = Alignment(horizontal="right")
inp(S["D14"], 0.18, PCT, "Tasa general: 16 % IGV + 2 % IPM.")
name(S, "D14", "TasaIGV")
S["G14"] = "IGV de ventas = precio sin IGV x 18 % x (1 - 50 %) = 9 % efectivo (igual al Anexo A)."
S["G14"].font = F_NOTE
name(S, "E13", "VentasTot")
name(S, "C13", "VentasSinIGV")
name(S, "D13", "IGVVentas")

# 2. Egresos --------------------------------------------------------------------
section(S, 15, "2. EGRESOS DEL PROYECTO (Anexo A)", 7)
for i, h in enumerate(["Concepto", "Sin IGV", "IGV", "Total con IGV", "% IGV/total", "Afecto IGV (1/0)"]):
    S.cell(row=16, column=1 + i, value=h).font = F_B
egresos = [
    ("Terreno – precio base acordado", 950000, 0, "Terreno", "Precio del Anexo A. El precio final depende del mecanismo (sección 5)."),
    ("Construcción (sobre rasante, azotea, sótanos, amenities, equipamiento)", 3805297, 684954, "Construccion", None),
    ("Desarrollo del proyecto (arquitectura, especialidades, estudios)", 148407, 26713, "Desarrollo", None),
    ("Gerencia de proyectos", 434036, 78126, "GerProy", None),
    ("Gastos administrativos", 62005, 11161, "GastAdm", None),
    ("Publicidad", 372031, 66966, "Publicidad", None),
    ("Gerencia de ventas / comisiones", 248020, 44644, "GerVentas", "Se devenga en proporción a las ventas a terceros."),
]
r = 17
for lbl, s, igv, nm, note in egresos:
    S.cell(row=r, column=1, value=lbl).font = F
    inp(S.cell(row=r, column=2), s)
    inp(S.cell(row=r, column=6), 0 if nm == "Terreno" else 1, INT,
        "La venta de terreno no está gravada con IGV." if nm == "Terreno" else None)
    calc(S.cell(row=r, column=3), f"=B{r}*TasaIGV*F{r}")
    calc(S.cell(row=r, column=4), f"=B{r}+C{r}")
    calc(S.cell(row=r, column=5), f"=IFERROR(C{r}/D{r},0)", PCT2)
    if note:
        S.cell(row=r, column=7, value=note).font = F_NOTE
    name(S, f"D{r}", nm)
    name(S, f"E{r}", "r" + nm)
    r += 1
# r == 24
S.cell(row=24, column=1, value="Alcabala – tasa").font = F
inp(S["B24"], 0.03, PCT)
name(S, "B24", "TasaAlc")
S.cell(row=25, column=1, value="Alcabala – tramo inafecto (10 UIT, en US$)").font = F
inp(S["B25"], 15933, USD, "Calibrado para reproducir la alcabala del Anexo A (US$ 28,022 sobre US$ 950,000). Actualizar con UIT y tipo de cambio vigentes.")
name(S, "B25", "ExonAlc")
S.cell(row=26, column=1, value="Gastos notariales y registrales (SUNARP) del terreno").font = F
inp(S["B26"], 6000, USD, "Supuesto nuevo (no figura en el Anexo A): escritura, partes notariales, inscripción de transferencia e hipotecas.")
name(S, "B26", "Notariales")
S["G26"] = "Supuesto del modelo, no está en el Anexo A."
S["G26"].font = F_NOTE

# 3. Cronograma -----------------------------------------------------------------
section(S, 28, "3. CRONOGRAMA (mes 0 = firma del acuerdo de terreno)", 7)
crono = [
    ("Mes de transferencia del terreno (firma)", 0, "MesTerreno", None),
    ("Inicio de preventa", 1, "PrevIni", "Supuesto del cliente: 3 meses de preventa."),
    ("Fin de preventa", 3, "PrevFin", None),
    ("Inicio de construcción", 4, "ObraIni", "Supuesto del cliente: 15 meses de obra."),
    ("Fin de construcción / entrega de unidades", 18, "MesEntrega", "Total proyecto 18 meses."),
    ("Inicio de cobro de saldos (desembolsos hipotecarios)", 19, "SaldoIni", "Los bancos de los compradores desembolsan tras conformidad de obra / independización."),
    ("N.º de meses para cobrar los saldos", 2, "SaldoN", None),
    ("Desarrollo del proyecto – mes inicial", 0, "DesIni", None),
    ("Desarrollo del proyecto – mes final", 3, "DesFin", "Anteproyecto y licencia deben estar listos al iniciar obra."),
    ("Gerencia y administración – mes inicial", 0, "GerIni", None),
    ("Gerencia y administración – mes final", 20, "GerFin", "Incluye cierre: conformidad, independización, liquidación."),
    ("Mes de pago del impuesto a la renta", 22, "MesIR", "Simplificación: un solo pago al cierre (no modela pagos a cuenta mensuales)."),
]
r = 29
for lbl, v, nm, note in crono:
    S.cell(row=r, column=1, value=lbl).font = F
    inp(S.cell(row=r, column=2), v, INT)
    name(S, f"B{r}", nm)
    if note:
        S.cell(row=r, column=7, value=note).font = F_NOTE
    r += 1
# r == 41
S.cell(row=41, column=1, value="Publicidad: % del presupuesto gastado en la preventa").font = F
inp(S["B41"], 0.40, PCT)
name(S, "B41", "PubPrev")

# 4. Ventas y cobranza ---------------------------------------------------------
section(S, 43, "4. FORMA DE PAGO DE LOS COMPRADORES", 7)
S.cell(row=44, column=1, value="Cuota inicial (al firmar la compraventa)").font = F
inp(S["B44"], 0.20, PCT, "Recomendado 20 %: con 10 % el préstamo puente sube a ~US$ 1.6 M y dura hasta el mes 19; con 20 % se paga en el mes 4 con la preventa.")
name(S, "B44", "PctIni")
S.cell(row=45, column=1, value="Cuotas durante la obra (repartidas hasta la entrega)").font = F
inp(S["B45"], 0.10, PCT)
name(S, "B45", "PctCuo")
S.cell(row=46, column=1, value="Saldo con crédito hipotecario (a la entrega)").font = F
calc(S["B46"], "=1-B44-B45", PCT)
name(S, "B46", "PctSal")
S.cell(row=47, column=1, value="Meta de preventa exigida por el banco (% de ventas)").font = F
inp(S["B47"], 0.30, PCT, "Los bancos suelen pedir entre 20 % y 40 % de preventa (punto de equilibrio) para levantar condiciones del crédito promotor.")
name(S, "B47", "MetaPrev")
S["G44"] = "La velocidad de ventas mes a mes se edita en Flujo_Mensual, fila 7."
S["G44"].font = F_NOTE

# 5. Terreno --------------------------------------------------------------------
section(S, 49, "5. MECANISMO DE ADQUISICIÓN DEL TERRENO (escenarios)", 7)
S.cell(row=50, column=1, value="Escenario elegido (1 a 4) →").font = F_B
inp(S["B50"], 4, INT, "1 Contado · 2 Compraventa con pago diferido · 3 Permuta total · 4 Mixto (permuta parcial + pago diferido)")
name(S, "B50", "EscTerr")
calc(S["C50"], "=INDEX(B52:E52,EscTerr)", "@", True)
heads = ["Parámetro", "1 · Contado", "2 · Pago diferido", "3 · Permuta total", "4 · Mixto", "Elegido"]
for i, h in enumerate(heads):
    S.cell(row=51, column=1 + i, value=h).font = F_B
S.cell(row=52, column=1, value="Nombre").font = F
for i, n in enumerate(["Contado", "Pago diferido", "Permuta total", "Mixto"]):
    S.cell(row=52, column=2 + i, value=n).font = F_IN
terr = [
    ("Prima sobre precio base (pago por esperar / riesgo)", [0, 0.05, 0.20, 0.12], PCT, "TPrima"),
    ("% del precio pagado en departamentos (permuta)", [0, 0, 1.0, 0.50], PCT, "TPerm"),
    ("% pagado en efectivo a la firma (mes 0)", [1.0, 0.10, 0, 0.10], PCT, "T0"),
    ("% pagado al activar el crédito promotor", [0, 0.40, 0, 0.20], PCT, "TAct"),
    ("% pagado en efectivo en el mes final", [0, 0.50, 0, 0.20], PCT, "TFin"),
    ("Mes del pago final", [0, 18, 18, 18], INT, "TMesFin"),
]
r = 53
for lbl, vals, fmt, nm in terr:
    S.cell(row=r, column=1, value=lbl).font = F
    for i, v in enumerate(vals):
        inp(S.cell(row=r, column=2 + i), v, fmt)
    calc(S.cell(row=r, column=6), f"=INDEX(B{r}:E{r},EscTerr)", fmt, True)
    name(S, f"F{r}", nm)
    r += 1
# r == 59
S.cell(row=59, column=1, value="Control: la suma de % debe ser 100 %").font = F
for i in range(5):
    col = get_column_letter(2 + i)
    calc(S[f"{col}59"], f'=IF(ABS({col}54+{col}55+{col}56+{col}57-1)<0.0001,"OK","REVISAR")', "@")
S.cell(row=60, column=1, value="Precio total del terreno (con prima)").font = F_B
calc(S["F60"], "=Terreno*(1+TPrima)", USD, True)
name(S, "F60", "PrecioTerr")
S.cell(row=61, column=1, value="Valor entregado en departamentos al propietario").font = F
calc(S["F61"], "=PrecioTerr*TPerm")
name(S, "F61", "PermutaVal")
S.cell(row=62, column=1, value="Ventas disponibles para terceros (con IGV)").font = F
calc(S["F62"], "=VentasTot-PermutaVal")
name(S, "F62", "VentasNetas")
S.cell(row=63, column=1, value="Equivalente en m² vendibles entregados al propietario").font = F
calc(S["F63"], "=PermutaVal/(VentasTot/1744)", '#,##0.0')
S["G63"] = "Referencia: precio promedio US$ 7,751/m² con estacionamientos."
S["G63"].font = F_NOTE
S["G53"] = "Ver hoja Terreno_Estrategias para el detalle legal y SUNARP de cada mecanismo."
S["G53"].font = F_NOTE

# 6. Financiamiento ----------------------------------------------------------------
section(S, 65, "6. FINANCIAMIENTO", 7)
fin = [
    ("Aporte de capital de la desarrolladora", 0, USD, "Aporte", "Sin aporte propio: todo lo previo al crédito promotor (incluida la cuota inicial del terreno) lo cubre el préstamo puente."),
    ("Mes del aporte", 0, INT, "MesAporte", None),
    ("Caja mínima operativa", 20000, USD, "CajaMin", None),
    ("Crédito promotor – línea máxima (= % x costo de construcción, ver fila 82)", 0, USD, "Linea", "Se usa solo lo necesario (no revolvente). El banco financia obra, no terreno ni el repago de otras deudas."),
    ("Crédito promotor – mes de activación", 4, INT, "MesAct", "Requiere: terreno inscrito a nombre de la desarrolladora (o fideicomiso), licencia y meta de preventa."),
    ("Crédito promotor – TEA", 0.11, PCT, "TEABanco", "Referencial en US$; confirmar con cotización del banco."),
    ("Crédito promotor – comisión de estructuración (% línea)", 0.01, PCT, "ComBanco", None),
    ("Supervisión bancaria / fideicomiso (US$ por mes de obra)", 1500, USD, "SupBanco", None),
    ("Préstamo puente / inversionista – monto máximo", 2500000, USD, "PuenteMax", "Cubre los meses previos al crédito promotor. Fondo privado, socio o financiamiento participativo."),
    ("Préstamo puente – TEA (intereses capitalizados)", 0.18, PCT, "TEAPuente", "Referencial para deuda privada subordinada en US$. Intereses capitalizados: se pagan juntos al final."),
    ("Retener cobranzas en escrow hasta activar el crédito (1 = sí, 0 = no)", 1, INT, "Escrow", "Práctica usual: las cuotas de preventa quedan en cuenta recaudadora hasta alcanzar el punto de equilibrio."),
    ("Tasa de descuento para VAN (TEA)", 0.15, PCT, "TDesc", None),
    ("Impuesto a la renta", 0.295, PCT, "TasaIR", None),
]
r = 66
for lbl, v, fmt, nm, note in fin:
    S.cell(row=r, column=1, value=lbl).font = F
    inp(S.cell(row=r, column=2), v, fmt)
    name(S, f"B{r}", nm)
    if note:
        S.cell(row=r, column=7, value=note).font = F_NOTE
    r += 1
# r == 79
S.cell(row=79, column=1, value="Tasa mensual crédito promotor").font = F
calc(S["B79"], "=(1+TEABanco)^(1/12)-1", PCT2)
name(S, "B79", "TMBanco")
S.cell(row=80, column=1, value="Tasa mensual préstamo puente").font = F
calc(S["B80"], "=(1+TEAPuente)^(1/12)-1", PCT2)
name(S, "B80", "TMPuente")
S.cell(row=82, column=1, value="Crédito promotor – % del costo de construcción que financia el banco").font = F
inp(S["B82"], 0.60, PCT, "Rango usual en bancos peruanos: 50 % a 70 % del costo de obra. 60 % es de aprobación holgada si hay preventa y terreno inscrito.")
name(S, "B82", "PctLinea")
calc(S["B69"], "=PctLinea*Construccion")
S.cell(row=83, column=1, value="Pago del préstamo puente: 1 = al activar el crédito · 0 = al final, después del banco").font = F
inp(S["B83"], 0, INT, "0 recomendado: el banco no deja que su crédito pague otras deudas; exige que el puente quede subordinado hasta cancelar el crédito promotor.")
name(S, "B83", "PagoPuente")
S["G82"] = "Supuesto conservador recomendado: 60 %."
S["G82"].font = F_NOTE
S.cell(row=81, column=1, value="Tasa mensual de descuento").font = F
calc(S["B81"], "=(1+TDesc)^(1/12)-1", PCT2)
name(S, "B81", "TMDesc")
S.freeze_panes = "B4"

# ════════════════════════════════════════════════════════════════ Flujo_Mensual
FM = wb.create_sheet("Flujo_Mensual")
title(FM, "FLUJO DE CAJA MENSUAL · Lote demo · US$ (con IGV)",
      "Mes 0 = firma del terreno · 1–3 preventa · 4–18 construcción · 19–22 cobro de saldos y cierre. "
      "Filas 7 y 8 (amarillas) son curvas editables.", LAST_COL)
FM.column_dimensions["A"].width = 50
FM.column_dimensions["B"].width = 13
for c in range(FIRST_COL, LAST_COL + 1):
    FM.column_dimensions[get_column_letter(c)].width = 11
FM.cell(row=3, column=1, value="Mes").font = F_B
FM.cell(row=3, column=2, value="TOTAL").font = F_B
FM.cell(row=4, column=1, value="Fase").font = F_B
for i in range(N_MONTHS):
    c = FIRST_COL + i
    col = get_column_letter(c)
    cell = FM.cell(row=3, column=c, value=i)
    cell.font = F_B
    cell.alignment = Alignment(horizontal="center")
    ph = FM.cell(row=4, column=c,
                 value=f'=IF({col}3<PrevIni,"Inicio",IF({col}3<=PrevFin,"Preventa",IF({col}3<=MesEntrega,"Construcción","Cierre")))')
    ph.font = Font(name=ARIAL, size=8)
    ph.alignment = Alignment(horizontal="center")

sales_curve = [0, .10, .10, .10] + [round(.70 / 15, 6)] * 14 + [0.70 - round(.70 / 15, 6) * 14] + [0] * 4
cons_curve = [0, 0, 0, 0] + [.03, .04, .05, .06, .07, .08, .08, .09, .09, .09, .08, .08, .07, .05, .04] + [0] * 4
assert len(sales_curve) == N_MONTHS and len(cons_curve) == N_MONTHS

ROWS = {}
labels = []  # (row, label, kind)


def row_formula(rw, label, fn, fmt=USD, bold=False, total=True, first=None, kind="calc"):
    """fn(col, prevcol) -> formula string. first: alt formula for month-0 column."""
    FM.cell(row=rw, column=1, value=label).font = F_B if bold else F
    for i in range(N_MONTHS):
        c = FIRST_COL + i
        col, prev = get_column_letter(c), get_column_letter(c - 1)
        f = first(col) if (i == 0 and first) else fn(col, prev)
        calc(FM.cell(row=rw, column=c), f, fmt, bold)
    if total:
        calc(FM.cell(row=rw, column=2), f"=SUM(C{rw}:{LC}{rw})", fmt, True)
    if bold:
        for c in range(1, LAST_COL + 1):
            FM.cell(row=rw, column=c).fill = FILL_TOT
            FM.cell(row=rw, column=c).border = TOP


def sec(rw, text):
    section(FM, rw, text, LAST_COL)


sec(6, "CURVAS DE AVANCE (editables · cada fila debe sumar 100 %)")
for rw, lbl, curve in [(7, "Velocidad de ventas (% de unidades vendidas en el mes)", sales_curve),
                       (8, "Avance de obra – curva S (% del costo de construcción)", cons_curve)]:
    FM.cell(row=rw, column=1, value=lbl).font = F
    for i, v in enumerate(curve):
        inp(FM.cell(row=rw, column=FIRST_COL + i), v, PCT)
    calc(FM.cell(row=rw, column=2), f"=SUM(C{rw}:{LC}{rw})", PCT, True)

sec(10, "INGRESOS")
row_formula(11, "Ventas del mes (contratos firmados con terceros)", lambda c, p: f"=VentasNetas*{c}7")
row_formula(12, "  % vendido acumulado", lambda c, p: f"=IFERROR(SUM($C$11:{c}11)/VentasNetas,0)", PCT, total=False)
row_formula(13, "Cobro de cuotas iniciales", lambda c, p: f"={c}11*PctIni")
row_formula(14, "  Cuotas por cobrar – saldo inicial", lambda c, p: f"={p}17", first=lambda c: "=0", total=False)
row_formula(15, "  Cuotas nuevas generadas", lambda c, p: f"={c}11*PctCuo")
row_formula(16, "Cobro de cuotas durante la obra",
            lambda c, p: f"=IF({c}3<=MesEntrega,({c}14+{c}15)/(MesEntrega-{c}3+1),{c}14+{c}15)")
row_formula(17, "  Cuotas por cobrar – saldo final", lambda c, p: f"={c}14+{c}15-{c}16", total=False)
row_formula(18, "Cobro de saldos (desembolsos hipotecarios)",
            lambda c, p: f"=IF(AND({c}3>=SaldoIni,{c}3<SaldoIni+SaldoN),SUMIF($C$3:${LC}$3,\"<\"&SaldoIni,$C$11:${LC}$11)*PctSal/SaldoN,0)+IF({c}3>=SaldoIni,{c}11*PctSal,0)")
row_formula(19, "TOTAL COBRANZAS", lambda c, p: f"={c}13+{c}16+{c}18", bold=True)
row_formula(20, "  (–) Depósito en cuenta escrow (antes del crédito)",
            lambda c, p: f"=IF(AND(Escrow=1,{c}3<MesAct),-{c}19,0)")
row_formula(21, "  (+) Liberación del escrow al activar el crédito",
            lambda c, p: f"=IF(AND(Escrow=1,{c}3=MesAct),-SUMIF($C$3:${LC}$3,\"<\"&MesAct,$C$20:${LC}$20),0)")
row_formula(22, "Cobranzas disponibles", lambda c, p: f"={c}19+{c}20+{c}21", bold=True)

sec(24, "EGRESOS")
row_formula(25, "Terreno – pagos en efectivo al propietario",
            lambda c, p: f"=PrecioTerr*(IF({c}3=MesTerreno,T0,0)+IF({c}3=MesAct,TAct,0)+IF({c}3=TMesFin,TFin,0))")
row_formula(26, "Alcabala + gastos notariales y registrales",
            lambda c, p: f"=IF({c}3=MesTerreno,TasaAlc*MAX(0,PrecioTerr-ExonAlc)+Notariales,0)")
row_formula(27, "Desarrollo del proyecto",
            lambda c, p: f"=IF(AND({c}3>=DesIni,{c}3<=DesFin),Desarrollo/(DesFin-DesIni+1),0)")
row_formula(28, "Construcción", lambda c, p: f"=Construccion*{c}8")
row_formula(29, "Gerencia de proyectos",
            lambda c, p: f"=IF(AND({c}3>=GerIni,{c}3<=GerFin),GerProy/(GerFin-GerIni+1),0)")
row_formula(30, "Gastos administrativos",
            lambda c, p: f"=IF(AND({c}3>=GerIni,{c}3<=GerFin),GastAdm/(GerFin-GerIni+1),0)")
row_formula(31, "Publicidad",
            lambda c, p: f"=Publicidad*(IF(AND({c}3>=PrevIni,{c}3<=PrevFin),PubPrev/(PrevFin-PrevIni+1),0)"
                         f"+IF(AND({c}3>=ObraIni,{c}3<=MesEntrega),(1-PubPrev)/(MesEntrega-ObraIni+1),0))")
row_formula(32, "Gerencia de ventas / comisiones", lambda c, p: f"=GerVentas*{c}11/VentasTot")
row_formula(33, "IGV neto pagado a SUNAT", lambda c, p: f"={c}37-{p}37", first=lambda c: f"={c}37")
row_formula(34, "  IGV débito (sobre cobranzas)", lambda c, p: f"={c}19*IGVVentas/VentasTot")
row_formula(35, "  IGV crédito fiscal (compras)",
            lambda c, p: f"={c}27*rDesarrollo+{c}28*rConstruccion+{c}29*rGerProy+{c}30*rGastAdm+{c}31*rPublicidad+{c}32*rGerVentas")
row_formula(36, "  IGV neto acumulado por pagar", lambda c, p: f"=MAX(0,SUM($C$34:{c}34)-SUM($C$35:{c}35))", total=False)
row_formula(37, "  Máximo acumulado pagado", lambda c, p: f"=MAX($C$36:{c}36)", total=False)
row_formula(38, "TOTAL EGRESOS", lambda c, p: f"=SUM({c}25:{c}33)", bold=True)

sec(40, "FLUJO ECONÓMICO (antes de financiamiento e impuesto a la renta)")
row_formula(41, "Flujo económico del mes", lambda c, p: f"={c}19-{c}38", bold=True)
row_formula(42, "  Flujo económico acumulado", lambda c, p: f"=SUM($C$41:{c}41)", total=False)
row_formula(43, "Flujo disponible (considera el escrow)", lambda c, p: f"={c}22-{c}38")

sec(45, "FINANCIAMIENTO · cascada: aporte → crédito promotor → préstamo puente")
row_formula(46, "Caja inicial", lambda c, p: f"={p}72", first=lambda c: "=0", total=False)
row_formula(47, "Aporte de capital de la desarrolladora", lambda c, p: f"=IF({c}3=MesAporte,Aporte,0)")
FM.cell(row=49, column=1, value="Crédito promotor (banco)").font = F_B
row_formula(50, "  Saldo inicial", lambda c, p: f"={p}53", first=lambda c: "=0", total=False)
row_formula(51, "  (+) Desembolso",
            lambda c, p: f"=IF(AND({c}3>=MesAct,{c}3<=MesEntrega),MIN({c}68+IF(AND(PagoPuente=1,{c}3=MesAct),MAX(0,{c}60+{c}61-MAX(0,{c}67-CajaMin)),0),Linea-{p}57),0)",
            first=lambda c: f"=IF(AND({c}3>=MesAct,{c}3<=MesEntrega),MIN({c}68+IF(AND(PagoPuente=1,{c}3=MesAct),MAX(0,{c}60+{c}61-MAX(0,{c}67-CajaMin)),0),Linea),0)")
row_formula(52, "  (–) Amortización (con cobranzas)", lambda c, p: f"=MIN({c}69,{c}50)")
row_formula(53, "  Saldo final", lambda c, p: f"={c}50+{c}51-{c}52", total=False)
row_formula(54, "  Intereses pagados", lambda c, p: f"={c}50*TMBanco")
row_formula(55, "  Comisión de estructuración", lambda c, p: f"=IF({c}3=MesAct,Linea*ComBanco,0)")
row_formula(56, "  Supervisión bancaria / fideicomiso",
            lambda c, p: f"=IF(AND({c}3>=MesAct,{c}3<=MesEntrega),SupBanco,0)")
row_formula(57, "  Desembolsos acumulados", lambda c, p: f"=SUM($C$51:{c}51)", total=False)
FM.cell(row=59, column=1, value="Préstamo puente / inversionista").font = F_B
row_formula(60, "  Saldo inicial", lambda c, p: f"={p}64", first=lambda c: "=0", total=False)
row_formula(61, "  Intereses capitalizados", lambda c, p: f"={c}60*TMPuente")
row_formula(62, "  (+) Desembolso", lambda c, p: f"=MAX(0,MIN({c}68-{c}51,PuenteMax-{p}65))",
            first=lambda c: f"=MAX(0,MIN({c}68-{c}51,PuenteMax))")
row_formula(63, "  (–) Repago (subordinado: después de cancelar el banco)", lambda c, p: f"=MIN({c}69-{c}52,{c}60+{c}61)")
row_formula(64, "  Saldo final", lambda c, p: f"={c}60+{c}61+{c}62-{c}63", total=False)
row_formula(65, "  Desembolsos acumulados", lambda c, p: f"=SUM($C$62:{c}62)", total=False)
FM.cell(row=66, column=1, value="Cascada de caja").font = F_B
row_formula(67, "  Caja antes de financiamiento",
            lambda c, p: f"={c}46+{c}43+{c}47-{c}54-{c}55-{c}56", total=False)
row_formula(68, "  Necesidad de fondos", lambda c, p: f"=MAX(0,CajaMin-{c}67)")
row_formula(69, "  Excedente para repagar deuda", lambda c, p: f"=MAX(0,{c}67+{c}51+{c}62-CajaMin)", total=False)
row_formula(70, "  DÉFICIT NO CUBIERTO (debe ser cero)", lambda c, p: f"=MAX(0,{c}68-{c}51-{c}62)")
row_formula(72, "CAJA FINAL", lambda c, p: f"={c}67+{c}51+{c}62-{c}52-{c}63", bold=True, total=False)

sec(74, "FLUJO DEL ACCIONISTA (desarrolladora)")
row_formula(75, "Aporte de capital", lambda c, p: f"=-{c}47")
row_formula(76, "Impuesto a la renta", lambda c, p: f"=IF({c}3=MesIR,-Resumen!$C$21,0)")
row_formula(77, "Retiro de caja al cierre", lambda c, p: f"=IF({c}3=MAX($C$3:${LC}$3),{c}72,0)")
row_formula(78, "FLUJO NETO DEL ACCIONISTA", lambda c, p: f"={c}75+{c}76+{c}77", bold=True)
for rw in (76,):
    for c in range(FIRST_COL, LAST_COL + 1):
        FM.cell(row=rw, column=c).font = F_LINK
FM.freeze_panes = "C5"

# phase coloring on header
for i in range(N_MONTHS):
    m = i
    ph = "Inicio" if m < 1 else "Preventa" if m <= 3 else "Construcción" if m <= 18 else "Cierre"
    FM.cell(row=3, column=FIRST_COL + i).fill = PatternFill("solid", fgColor=FILL_PH[ph])
    FM.cell(row=4, column=FIRST_COL + i).fill = PatternFill("solid", fgColor=FILL_PH[ph])

# ═══════════════════════════════════════════════════════════════════════ Resumen
R = wb.create_sheet("Resumen", 0)
title(R, "RESUMEN · Flujo de caja Lote demo (Huidobro Desarrolladora)",
      "18 meses: 3 de preventa + 15 de obra. Cifras en US$. Cambie supuestos en la hoja Supuestos.", 5)
for col, w in zip("ABCDE", [4, 52, 16, 16, 50]):
    R.column_dimensions[col].width = w
FMn = "Flujo_Mensual"
items = [
    (4, "sec", "ESCENARIO DE TERRENO"),
    (5, "Mecanismo elegido", "=Supuestos!C50", "@", "Cambiar en Supuestos!B50 (1 a 4)."),
    (6, "Precio total del terreno (incluye prima)", "=PrecioTerr", USD, None),
    (7, "  pagado en departamentos (permuta)", "=PermutaVal", USD, None),
    (8, "  pagado en efectivo", "=PrecioTerr-PermutaVal", USD, None),
    (10, "sec", "RESULTADOS"),
    (11, "Ventas totales con IGV (incluye unidades de permuta)", "=VentasTot", USD, None),
    (12, "Ventas a terceros con IGV", "=VentasNetas", USD, None),
    (13, "Cobranzas totales", f"={FMn}!B19", USD, None),
    (14, "Egresos operativos totales (con IGV neto)", f"={FMn}!B38", USD, None),
    (15, "Costos financieros (intereses, comisiones, supervisión)",
     f"={FMn}!B54+{FMn}!B55+{FMn}!B56+{FMn}!B61", USD, None),
    (16, "Utilidad antes de IR (contable, sin IGV)",
     f"=VentasSinIGV-({FMn}!B27+{FMn}!B28+{FMn}!B29+{FMn}!B30+{FMn}!B31+{FMn}!B32-{FMn}!B35)-PrecioTerr-{FMn}!B26-C15", USD,
     "Ventas sin IGV – costos sin IGV – terreno – alcabala – costos financieros."),
    (17, "Margen antes de IR / ventas sin IGV", "=IFERROR(C16/VentasSinIGV,0)", PCT, None),
    (21, "Impuesto a la renta", "=TasaIR*MAX(0,C16)", USD, None),
    (22, "Utilidad neta", "=C16-C21", USD, None),
    (23, "Margen neto / ventas sin IGV", "=IFERROR(C22/VentasSinIGV,0)", PCT, None),
    (25, "sec", "NECESIDADES DE FINANCIAMIENTO"),
    (26, "Exposición máxima del proyecto (flujo económico acumulado mínimo)", f"=-MIN(0,MIN({FMn}!C42:{LC}42))", USD,
     "Lo que habría que financiar si no hubiera escrow ni crédito."),
    (27, "Aporte de capital de la desarrolladora", "=Aporte", USD, "Supuesto: cero."),
    (28, "Saldo máximo del crédito promotor", f"=MAX({FMn}!C53:{LC}53)", USD, None),
    (29, "  uso de la línea aprobada", "=IFERROR(C28/Linea,0)", PCT, None),
    (30, "Saldo máximo del préstamo puente", f"=MAX({FMn}!C64:{LC}64)", USD, None),
    (31, "Mes en que se cancela el préstamo puente",
     f"=IFERROR(INDEX({FMn}!C3:{LC}3,MATCH(1,INDEX(({FMn}!C64:{LC}64<1)*({FMn}!C3:{LC}3>0)*({FMn}!C65:{LC}65>0),0),0)),\"-\")", INT, None),
    (32, "Déficit no cubierto (debe ser cero)", f"={FMn}!B70", USD, None),
    (33, "Intereses totales pagados al inversionista puente", f"={FMn}!B61", USD, None),
    (34, "sec", "INDICADORES"),
    (35, "TIR económica (anual)", f"=(1+IRR({FMn}!C41:{LC}41,0.02))^12-1", PCT, "Flujo del proyecto antes de financiamiento e IR."),
    (36, "VAN económico @ tasa de descuento", f"=NPV(TMDesc,{FMn}!D41:{LC}41)+{FMn}!C41", USD, None),
    (37, "TIR del accionista (anual, después de IR)", f'=IF(Aporte=0,"n.a.",(1+IRR({FMn}!C78:{LC}78,0.05))^12-1)', PCT,
     "Sin aporte propio la TIR no se puede calcular (no hay inversión): mirar la utilidad neta."),
    (38, "Múltiplo sobre el capital aportado", '=IF(Aporte=0,"n.a.",(C22+Aporte)/Aporte)', '0.0"x"', None),
    (39, "Cobertura del puente: utilidad antes de IR / saldo máximo del puente", "=IFERROR(C16/C30,0)", '0.0"x"',
     "Cuántas veces el margen del proyecto cubre la deuda puente (colchón para el inversionista)."),
    (40, "sec", "CONTROLES"),
    (41, "Preventa lograda al fin de la preventa (% de ventas a terceros)",
     f"=INDEX({FMn}!C12:{LC}12,MATCH(PrevFin,{FMn}!C3:{LC}3,0))", PCT, None),
    (42, "¿Cumple la meta de preventa del banco?", '=IF(C41>=MetaPrev,"Cumple","No cumple")', "@", None),
    (43, "Curva de ventas suma 100 % y termina antes de la entrega",
     f'=IF(AND(ABS({FMn}!B7-1)<0.0001,SUMIF({FMn}!C3:{LC}3,">"&MesEntrega,{FMn}!C7:{LC}7)=0),"OK","REVISAR")', "@", None),
    (44, "Curva de obra suma 100 %", f'=IF(ABS({FMn}!B8-1)<0.0001,"OK","REVISAR")', "@", None),
    (45, "Estructura de pago del terreno suma 100 %", "=Supuestos!F59", "@", None),
    (46, "Caja final ≥ impuesto a la renta",
     f'=IF(INDEX({FMn}!C72:{LC}72,MATCH(MesIR,{FMn}!C3:{LC}3,0))>=C21,"OK","REVISAR")', "@", None),
]
for it in items:
    rw = it[0]
    if it[1] == "sec":
        section(R, rw, it[2], 5)
        continue
    _, lbl, f, fmt, note = it
    R.cell(row=rw, column=2, value=lbl).font = F
    calc(R.cell(row=rw, column=3), f, fmt, True)
    if note:
        R.cell(row=rw, column=5, value=note).font = F_NOTE
R.cell(row=18, column=2, value="Referencia Anexo A: utilidad antes de IR US$ 6,353,205 (47.0 %), sin costos financieros.").font = F_NOTE

section(R, 48, "FONDOS A CONSEGUIR ANTES DEL CRÉDITO PROMOTOR (meses previos a la activación)", 5)
pre = lambda rw: f'=SUMIF({FMn}!$C$3:${LC}$3,"<"&MesAct,{FMn}!$C${rw}:${LC}${rw})'
fondos = [
    (49, "Cuota inicial del terreno (pago a la firma)", pre(25)),
    (50, "Alcabala y gastos notariales/registrales", pre(26)),
    (51, "Desarrollo del proyecto (arquitectura, licencia)", pre(27)),
    (52, "Gerencia de proyectos y gastos administrativos", pre(29) + "+" + pre(30)[1:]),
    (53, "Publicidad de lanzamiento", pre(31)),
    (54, "Comisiones de ventas", pre(32)),
    (55, "IGV neto pagado", pre(33)),
    (56, "Caja mínima al activar el crédito", "=CajaMin"),
]
for rw, lbl, f in fondos:
    R.cell(row=rw, column=2, value=lbl).font = F
    calc(R.cell(row=rw, column=3), f)
R.cell(row=57, column=2, value="TOTAL A CONSEGUIR").font = F_B
calc(R["C57"], "=SUM(C49:C56)", USD, True)
R.cell(row=58, column=2, value="Cubierto con: aporte propio").font = F
calc(R["C58"], "=Aporte")
R.cell(row=59, column=2, value="Cubierto con: préstamo puente / inversionista").font = F
calc(R["C59"], pre(62))
R.cell(row=60, column=2, value="Saldo del puente al activar el crédito (incluye intereses)").font = F
calc(R["C60"], f"=INDEX({FMn}!C60:{LC}60,MATCH(MesAct,{FMn}!C3:{LC}3,0))")
R["E57"] = "Las cobranzas de preventa quedan en escrow: no ayudan a pagar estos montos."
R["E57"].font = F_NOTE
R["E59"] = "Deuda subordinada: se cobra al final, cuando ya se canceló el crédito promotor."
R["E59"].font = F_NOTE
for col in "BC":
    R[f"{col}57"].fill = FILL_TOT

section(R, 62, "COMPARATIVO DE MECANISMOS DE TERRENO (valores calculados con este modelo)", 5)
R.cell(row=63, column=2, value="Tabla de valores fijos: se obtuvo corriendo el modelo con cada escenario. "
                               "Si cambia supuestos, vuelva a correr scripts/generar_flujo_caja_demo.py.").font = F_NOTE
R.freeze_panes = "A4"

# ═══════════════════════════════════════════════════════════ Terreno_Estrategias
T = wb.create_sheet("Terreno_Estrategias")
title(T, "¿CÓMO LOGRAR QUE EL TERRENO FIGURE EN SUNARP A NOMBRE DE LA DESARROLLADORA?",
      "Opciones legales y financieras con poco capital propio. Validar con abogado y notario antes de firmar.", 6)
for col, w in zip("ABCDEF", [24, 42, 34, 34, 34, 34]):
    T.column_dimensions[col].width = w
heads = ["Mecanismo", "Cómo funciona", "Cómo queda en SUNARP", "Ventajas", "Riesgos / cuidados", "Compatibilidad con crédito promotor"]
for i, h in enumerate(heads):
    c = T.cell(row=4, column=1 + i, value=h)
    c.font = F_SEC
    c.fill = FILL_SEC
rows = [
    ("1. Compra al contado (escenario 1)",
     "Se paga el 100 % a la firma de la escritura.",
     "Transferencia inscrita a nombre de la desarrolladora, sin cargas.",
     "Lo más simple; el banco lo acepta de inmediato; mejor precio.",
     "Exige ~US$ 1 M que hoy no se tiene: requeriría deuda puente cara o un socio de capital.",
     "Total."),
    ("2. Compraventa con saldo de precio (escenario 2)",
     "Escritura de compraventa con pago inicial (10 %), segundo pago al activar el crédito (40 %) y saldo a la entrega (50 %). El propietario cobra una prima por esperar.",
     "Se inscribe la transferencia; por ley queda hipoteca legal a favor del vendedor por el saldo (art. 1118 CC), salvo renuncia.",
     "Titularidad inmediata con poco efectivo; el pago grande lo cubre el banco y la preventa.",
     "El banco exige primera hipoteca: el propietario debe ceder rango o levantar la hipoteca legal al recibir el 2.º pago. Pactar cláusula resolutoria con cuidado.",
     "Alta, si el vendedor acepta ceder rango o cobrar al primer desembolso."),
    ("3. Permuta de bien futuro (escenario 3)",
     "El propietario transfiere el terreno a cambio de departamentos terminados (valor + prima de 15–25 %).",
     "Transferencia inscrita a nombre de la desarrolladora. El propietario suele pedir garantía: hipoteca, carta fianza o fideicomiso.",
     "Cero efectivo por el terreno. El propietario comparte el éxito del proyecto.",
     "Se entregan unidades (menos ventas a terceros). La permuta genera IGV en primera venta y alcabala. El propietario asume riesgo de obra: pedirá garantías.",
     "Media-alta: el banco acepta si la garantía del propietario queda subordinada o dentro de un fideicomiso."),
    ("4. Mixto: permuta parcial + pago diferido (escenario 4, recomendado)",
     "Parte del precio en departamentos (≈50 %) y el resto en efectivo: 10 % a la firma, 20 % al activar el crédito, 20 % a la entrega.",
     "Transferencia inscrita a nombre de la desarrolladora desde el mes 0.",
     "Equilibra caja y atractivo para el propietario (cobra algo en efectivo y una prima menor que en la permuta total).",
     "Mismo cuidado de rangos de garantía. Definir unidades a entregar (tipo, piso) en la escritura.",
     "Alta."),
    ("5. Fideicomiso inmobiliario (complemento)",
     "El terreno entra a un patrimonio fideicometido administrado por una fiduciaria; protege al propietario, al banco y a los compradores de preventa.",
     "Figura a nombre de la fiduciaria (dominio fiduciario), no de la desarrolladora; al final se transfieren las unidades.",
     "Resuelve el conflicto de garantías entre propietario y banco; da confianza a compradores.",
     "Costo de la fiduciaria. No cumple al pie de la letra que el terreno figure a nombre de la desarrolladora.",
     "Muy alta: es la estructura preferida por muchos bancos."),
    ("6. Asociación en participación / socio terrenista",
     "El propietario aporta el terreno como socio y recibe un % de la utilidad en lugar de un precio fijo.",
     "Si aporta en propiedad a una sociedad del proyecto (SPV), el terreno queda a nombre de la SPV.",
     "Cero efectivo; alinea intereses.",
     "Se reparte la utilidad (a 47 % de margen puede costar más que comprar). Requiere contrato de gobierno claro.",
     "Alta si el terreno se aporta en propiedad a la SPV."),
    ("7. Opción de compra inscrita (puente de 3 meses)",
     "Se paga una prima pequeña por el derecho exclusivo a comprar en un plazo; se ejerce cuando la preventa y el crédito estén listos.",
     "Se inscribe el contrato de opción, no la propiedad.",
     "Asegura el terreno con muy poco dinero mientras se valida la preventa.",
     "No da titularidad: muchos bancos y compradores piden propiedad inscrita antes de la preventa.",
     "Baja hasta ejercer la opción."),
]
for i, rw in enumerate(rows):
    for j, v in enumerate(rw):
        c = T.cell(row=5 + i, column=1 + j, value=v)
        c.font = F_B if j == 0 else F
        c.alignment = Alignment(wrap_text=True, vertical="top")
        c.border = Border(bottom=THIN)
    T.row_dimensions[5 + i].height = 105

r0 = 5 + len(rows) + 1
section(T, r0, "FUENTES DE FINANCIAMIENTO EXTERNO", 6)
fuentes = [
    ("Crédito promotor bancario", "Financia la obra contra preventa (20–40 %), terreno inscrito y licencia. Desembolsa por avance valorizado. No suele financiar la compra del terreno."),
    ("Préstamo puente de fondo privado o inversionista", "Cubre los meses 0–4 (terreno, licencia, preventa). Tasas mayores en US$; se paga con la liberación del escrow y el primer desembolso del banco."),
    ("Socio de capital (equity)", "Aporta el dinero del terreno a cambio de % de utilidad. Reduce deuda; diluye la ganancia."),
    ("Financiamiento participativo regulado (SMV)", "Plataformas autorizadas que levantan deuda o capital de muchos inversionistas para proyectos inmobiliarios."),
    ("Preventa en cuenta escrow", "No es financiamiento directo, pero al alcanzar la meta el banco libera las cuotas de los compradores."),
]
for i, (a, b) in enumerate(fuentes):
    T.cell(row=r0 + 1 + i, column=1, value=a).font = F_B
    c = T.cell(row=r0 + 1 + i, column=2, value=b)
    c.font = F
    c.alignment = Alignment(wrap_text=True, vertical="top")
    T.merge_cells(start_row=r0 + 1 + i, start_column=2, end_row=r0 + 1 + i, end_column=6)
    T.row_dimensions[r0 + 1 + i].height = 30
r1 = r0 + 1 + len(fuentes) + 1
section(T, r1, "SECUENCIA RECOMENDADA", 6)
pasos = [
    "Mes −1: negociar con el propietario el esquema mixto (escenario 4). Firmar arras o compromiso con plazo de 30–45 días.",
    "Mes −1: obtener carta de intención del banco (crédito promotor) y del inversionista puente, con estos flujos.",
    "Mes 0: escritura de permuta parcial + compraventa con saldo de precio; propietario cede rango a favor del banco. Inscribir en SUNARP.",
    "Mes 0: constituir fideicomiso o cuenta escrow para las cuotas de preventa; aportar capital y desembolsar puente.",
    "Meses 1–3: preventa hasta la meta del banco; licencia de edificación.",
    "Mes 4: activar crédito promotor, liberar escrow, pagar 2.º tramo al propietario y cancelar el puente.",
    "Mes 18: entrega; pagar saldo al propietario y entregar sus unidades. Meses 19–20: cobro de hipotecarios y cancelación del banco.",
]
for i, p in enumerate(pasos):
    c = T.cell(row=r1 + 1 + i, column=1, value=p)
    c.font = F
    T.merge_cells(start_row=r1 + 1 + i, start_column=1, end_row=r1 + 1 + i, end_column=6)
T.freeze_panes = "A5"

# settings
wb.calculation.fullCalcOnLoad = True
os.makedirs(os.path.dirname(OUT), exist_ok=True)
wb.save(OUT)

# ═══════════════════════════ comparativo: correr cada escenario y guardar valores
RECALC = sys.argv[1] if len(sys.argv) > 1 else None
if RECALC:
    metrics = [("Efectivo pagado por el terreno", 8, USD), ("Fondos a conseguir antes del crédito", 57, USD), ("Utilidad neta", 22, USD), ("Margen neto", 23, PCT),
               ("Saldo máx. crédito promotor", 28, USD),
               ("Saldo máx. préstamo puente", 30, USD), ("Mes de pago del puente", 31, INT), ("Costos financieros", 15, USD),
               ("Déficit no cubierto", 32, USD)]
    results = {}
    tmp = OUT.replace(".xlsx", "_tmp.xlsx")
    for esc in range(1, 5):
        w = load_workbook(OUT)
        w["Supuestos"]["B50"] = esc
        w.save(tmp)
        out = subprocess.run([sys.executable, RECALC, tmp, "90"], capture_output=True, text=True)
        info = json.loads(out.stdout)
        assert info.get("status") == "success", info
        v = load_workbook(tmp, data_only=True)["Resumen"]
        results[esc] = [v.cell(row=m[1], column=3).value for m in metrics]
    os.remove(tmp)
    w = load_workbook(OUT)
    ws = w["Resumen"]
    hdr = ["Indicador", "1 Contado", "2 Pago diferido", "3 Permuta total", "4 Mixto"]
    ws.column_dimensions["F"].width = 16
    for i, h in enumerate(hdr):
        c = ws.cell(row=64, column=2 + i, value=h)
        c.font = F_B
    for k, (lbl, _, fmt) in enumerate(metrics):
        ws.cell(row=65 + k, column=2, value=lbl).font = F
        for esc in range(1, 5):
            c = ws.cell(row=65 + k, column=2 + esc, value=results[esc][k])
            c.font = F
            c.number_format = fmt
    w.calculation.fullCalcOnLoad = True
    w.save(OUT)
    print(json.dumps(results, indent=1, default=str))
print("OK", OUT)
