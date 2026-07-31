# Cotización y modelo financiero – Proyecto inmobiliario

Paquete para que una **persona natural** (consultor independiente) cotice a una
**persona jurídica** (empresa desarrolladora inmobiliaria) la elaboración del
**flujo de caja**, el **cronograma valorizado** y el **estado de resultados**
de un proyecto, con el fin de sustentar una **línea de financiamiento** ante un banco.

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

## Nota tributaria (importante)

El impuesto en la cotización depende del régimen del consultor:

- **Recibo por Honorarios** (renta 4.ª categoría): **sin IGV**; la empresa puede
  **retener 8 %** de renta.
- **Factura** (persona natural con negocio, 3.ª categoría): **IGV 18 %**.

La plantilla muestra el escenario con IGV; ajuste la línea que corresponda a su caso.

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
