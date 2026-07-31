# Cotización y modelo financiero – Proyecto inmobiliario

Paquete para que **Enrique Huidobro** (economista, consultor independiente) cotice a
**Inversiones Estratégicas Integrales S.A.C.** la elaboración del **flujo de caja**,
el **cronograma valorizado** y el **estado de resultados** del proyecto de vivienda
**Villa El Milagro** (modalidad AVN – Programa Techo Propio), como sustento
técnico-financiero para obtener una **carta fianza ante empresas aseguradoras**.

> Honorarios: **S/ 10,000** (Recibo por Honorarios, con suspensión de 4.ª categoría →
> sin IGV ni retención), pagaderos por etapas **20 % / 30 % / 50 %**
> (S/ 2,000 / 3,000 / 5,000). Honorario dimensionado sobre la obra real
> (~S/ 3.3/m² ≈ 0.25 % del costo de obra). **La cotización está completa.**
>
> Datos reales cargados en el modelo: área techada **3,009.395 m²**, costo total de
> obra **S/ 3,998,621.74**. Pendiente para afinar el modelo: N° de unidades, precio
> de venta, costo del terreno y % de obra respaldado por la carta fianza.

## Entregables (`entregables/`)

| Archivo | Descripción |
|---|---|
| `Cotizacion_Servicios_Profesionales.docx` | Cotización editable (Word). Reemplace los textos en **rojo/[corchetes]**. |
| `Cotizacion_Servicios_Profesionales.pdf` | Versión final de la cotización, lista para enviar/imprimir. |
| `Modelo_Financiero_Proyecto_Inmobiliario.xlsx` | Modelo financiero dinámico (el trabajo que entregará a la empresa). |

## El modelo financiero (Excel)

Libro con 6 hojas; **solo se editan las celdas amarillas** (inputs), el resto se
recalcula solo:

1. **Portada** – identificación del proyecto y guía de uso.
2. **Supuestos** – parámetros: áreas, precios, costos, % de financiamiento, tasas.
3. **Cronograma_Valorizado** – avance de obra (perfil S) y **desembolsos bancarios**
   atados a la valorización mensual (base para las solicitudes de desembolso al banco).
4. **Flujo_Caja** – ingresos, egresos y servicio de deuda por mes (24 meses); flujo
   económico y financiero, con acumulados.
5. **Estado_Resultados** – P&L proyectado, utilidad y márgenes.
6. **Resumen** – indicadores: inversión, apalancamiento, ROI, **VAN y TIR**
   (económica y financiera).

El archivo tiene activado el recálculo al abrir, por lo que Excel/Google Sheets
muestran los valores automáticamente.

### Resultados con los datos de ejemplo

Proyecto de referencia: 20 unidades, 3 500 m², venta S/ 11.2 MM, obra S/ 5.78 MM,
línea del banco 60 % de la construcción.

| Indicador | Valor |
|---|---|
| Inversión total | S/ 9 405 500 |
| Línea máxima del banco | S/ 3 465 000 |
| Utilidad neta | S/ 1 016 529 (margen 9.1 %) |
| VAN financiero | S/ 259 039 |
| TIR financiera anual | 19.4 % |
| TIR económica anual | 17.4 % |

> Cifras referenciales; sustituya los supuestos por los del proyecto real.

## Nota tributaria

El consultor emite **Recibo por Honorarios** (renta de 4.ª categoría) con
**suspensión de retenciones** vigente: el servicio **no está afecto a IGV ni a
retención de renta**, por lo que el importe cotizado es el monto neto a transferir.

## Regenerar los documentos (`scripts/`)

```bash
# Modelo financiero (Excel)
pip install openpyxl
python scripts/generar_modelo_financiero.py

# Cotización (Word)
npm install docx
node scripts/generar_cotizacion_docx.js

# Cotización (PDF) — se genera desde cotizacion_plantilla.html con Chromium headless:
#   chrome --headless --print-to-pdf=entregables/Cotizacion_Servicios_Profesionales.pdf \
#          --no-pdf-header-footer scripts/cotizacion_plantilla.html

# Verificación numérica independiente del modelo
python scripts/verificar_modelo.py
```
