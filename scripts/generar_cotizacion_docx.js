// Cotización de servicios profesionales - persona natural -> persona jurídica
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType,
  PageBreak, Footer, Header, PositionalTab, PositionalTabAlignment,
  PositionalTabLeader, VerticalAlign,
} = require("docx");

const OUT = "/home/user/Eh-cloude/entregables/Cotizacion_Servicios_Profesionales.docx";

// ---- paleta ----
const NAVY = "1F3864";
const BLUE = "2E5496";
const GREY = "808080";
const PH = "C00000"; // placeholders (rojo) para que resalten

const FONT = "Calibri";

// helper: placeholder run (rojo, negrita)
const ph = (t) => new TextRun({ text: t, bold: true, color: PH, font: FONT });
const t = (text, opts = {}) => new TextRun({ text, font: FONT, ...opts });

function h(text, level) {
  return new Paragraph({
    heading: level,
    spacing: { before: 220, after: 100 },
    children: [new TextRun({ text, font: FONT, bold: true, color: NAVY })],
  });
}

function p(children, opts = {}) {
  return new Paragraph({
    spacing: { after: 120, line: 276 },
    alignment: opts.align || AlignmentType.JUSTIFIED,
    children: Array.isArray(children) ? children : [t(children)],
    ...opts,
  });
}

function bullet(children) {
  return new Paragraph({
    bullet: { level: 0 },
    spacing: { after: 60, line: 264 },
    alignment: AlignmentType.JUSTIFIED,
    children: Array.isArray(children) ? children : [t(children)],
  });
}

// ---- celdas de tabla ----
function cell(text, { bold, shade, color, align, width, size } = {}) {
  return new TableCell({
    width: width ? { size: width, type: WidthType.DXA } : undefined,
    verticalAlign: VerticalAlign.CENTER,
    shading: shade ? { type: ShadingType.CLEAR, fill: shade, color: "auto" } : undefined,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: [
      new Paragraph({
        alignment: align || AlignmentType.LEFT,
        children: Array.isArray(text)
          ? text
          : [new TextRun({ text, font: FONT, bold: !!bold, color: color || "000000", size: size || 20 })],
      }),
    ],
  });
}

const NB = { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" };
const borders = { top: NB, bottom: NB, left: NB, right: NB, insideHorizontal: NB, insideVertical: NB };

// =============================================================
// Contenido
// =============================================================
const children = [];

// Aviso de uso
children.push(new Paragraph({
  spacing: { after: 160 },
  shading: { type: ShadingType.CLEAR, fill: "FFF2CC", color: "auto" },
  border: { top: NB, bottom: NB, left: NB, right: NB },
  children: [
    new TextRun({ text: "  Instrucciones: ", bold: true, font: FONT, color: PH, size: 18 }),
    new TextRun({ text: "reemplace los textos en ", font: FONT, size: 18, color: GREY }),
    new TextRun({ text: "rojo y entre corchetes [ ]", font: FONT, bold: true, size: 18, color: PH }),
    new TextRun({ text: " con sus datos reales. Los montos son referenciales y editables.  ", font: FONT, size: 18, color: GREY }),
  ],
}));

// Encabezado profesional
children.push(new Paragraph({
  alignment: AlignmentType.LEFT, spacing: { after: 20 },
  children: [ph("[TU NOMBRE COMPLETO]")],
}));
children.push(new Paragraph({
  spacing: { after: 0 },
  children: [t("Consultor(a) independiente – ", { color: GREY, size: 18 }), ph("[Profesión / Especialidad]")],
}));
children.push(new Paragraph({
  spacing: { after: 0 },
  children: [t("RUC / DNI: ", { color: GREY, size: 18 }), ph("[N° RUC o DNI]"),
             t("   ·   Tel: ", { color: GREY, size: 18 }), ph("[teléfono]"),
             t("   ·   Email: ", { color: GREY, size: 18 }), ph("[correo]")],
}));
children.push(new Paragraph({
  border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: NAVY } },
  spacing: { after: 200 }, children: [t("")],
}));

// Título
children.push(new Paragraph({
  alignment: AlignmentType.CENTER, spacing: { after: 40 },
  children: [new TextRun({ text: "COTIZACIÓN DE SERVICIOS PROFESIONALES", bold: true, font: FONT, color: NAVY, size: 30 })],
}));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER, spacing: { after: 200 },
  children: [t("N° ", { color: GREY }), ph("[COT-2026-001]"),
             t("        Fecha: ", { color: GREY }), ph("[dd/mm/aaaa]"),
             t("        Ciudad: ", { color: GREY }), ph("[Lima]")],
}));

// Datos del cliente (tabla)
children.push(h("Señores", HeadingLevel.HEADING_2));
const clientTable = new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [2600, 6760],
  borders,
  rows: [
    new TableRow({ children: [cell("Razón social", { bold: true, shade: "D9E1F2", width: 2600 }), cell([ph("[RAZÓN SOCIAL DE LA EMPRESA DESARROLLADORA S.A.C.]")], { width: 6760 })] }),
    new TableRow({ children: [cell("RUC", { bold: true, shade: "D9E1F2", width: 2600 }), cell([ph("[N° RUC de la empresa]")], { width: 6760 })] }),
    new TableRow({ children: [cell("Atención", { bold: true, shade: "D9E1F2", width: 2600 }), cell([ph("[Nombre – cargo, p. ej. Gerente General]")], { width: 6760 })] }),
    new TableRow({ children: [cell("Proyecto", { bold: true, shade: "D9E1F2", width: 2600 }), cell([ph("[Nombre del proyecto inmobiliario]")], { width: 6760 })] }),
  ],
});
children.push(clientTable);
children.push(p([t("", {})]));

children.push(new Paragraph({
  spacing: { after: 200 },
  children: [t("Ref.: ", { bold: true }),
    t("Elaboración de flujo de caja, cronograma valorizado y estado de resultados proyectado del proyecto, para sustento de la línea de financiamiento ante la entidad bancaria.")],
}));

// 1. Presentación
children.push(h("1.  Presentación", HeadingLevel.HEADING_1));
children.push(p([
  t("Estimados señores:"),
]));
children.push(p([
  t("Es grato dirigirme a ustedes para presentar mi propuesta de servicios profesionales de consultoría financiera, orientada a estructurar la información económico-financiera del proyecto "),
  ph("[nombre del proyecto]"),
  t(", con el objetivo de sustentar técnicamente la solicitud de una línea de financiamiento ante "),
  ph("[nombre del banco]"),
  t(". Los entregables han sido diseñados conforme a los requisitos habituales que las áreas de riesgos y negocios inmobiliarios de la banca solicitan para la evaluación crediticia de proyectos de este tipo."),
]));

// 2. Objeto
children.push(h("2.  Objeto de la propuesta", HeadingLevel.HEADING_1));
children.push(p("Proveer a la empresa desarrolladora de un modelo financiero integral y de la documentación de sustento necesaria para presentar el proyecto al banco y gestionar los desembolsos asociados al avance de obra."));

// 3. Alcance
children.push(h("3.  Alcance de los servicios", HeadingLevel.HEADING_1));
children.push(p("La consultoría comprende los siguientes componentes:"));

children.push(h("3.1  Flujo de caja del proyecto", HeadingLevel.HEADING_2));
children.push(bullet("Proyección mensual de ingresos y egresos durante todo el horizonte del proyecto."));
children.push(bullet("Ingresos por cuota inicial de clientes (preventas), desembolsos de la línea bancaria y cobro de saldos vía créditos hipotecarios."));
children.push(bullet("Egresos por terreno, construcción, licencias y proyectos, comercialización, gastos administrativos y servicio de la deuda."));
children.push(bullet("Flujo de caja económico (sin financiamiento) y financiero (con financiamiento), con saldos acumulados."));

children.push(h("3.2  Cronograma valorizado de obra", HeadingLevel.HEADING_2));
children.push(bullet("Valorización mensual de la obra según curva de avance físico (perfil S)."));
children.push(bullet("Determinación del calendario de desembolsos bancarios atados a las valorizaciones, base para las solicitudes de desembolso ante el banco."));
children.push(bullet("Avance acumulado y saldo por ejecutar por período."));

children.push(h("3.3  Estado de resultados proyectado", HeadingLevel.HEADING_2));
children.push(bullet("Estado de resultados del proyecto: ventas, costo de ventas, utilidad bruta, gastos operativos, resultado financiero e impuestos."));
children.push(bullet("Indicadores de rentabilidad: márgenes bruto y neto, ROI, y evaluación con VAN y TIR (económica y financiera)."));

children.push(h("3.4  Informe / memoria de sustento para el banco", HeadingLevel.HEADING_2));
children.push(bullet("Documento explicativo que integra los supuestos, la metodología y los resultados del modelo financiero."));
children.push(bullet("Análisis de sensibilidad básico y conclusiones sobre la viabilidad y capacidad de repago del proyecto."));

// 4. Metodología
children.push(h("4.  Metodología de trabajo", HeadingLevel.HEADING_1));
children.push(bullet("Reunión de levantamiento de información y definición de supuestos con el equipo de la empresa."));
children.push(bullet("Construcción del modelo en hoja de cálculo dinámica y parametrizable (los resultados se recalculan al variar los supuestos)."));
children.push(bullet("Presentación de un borrador para revisión y ajustes conjuntos."));
children.push(bullet("Entrega final de los modelos e informe, con acompañamiento para su presentación al banco."));

// 5. Entregables (tabla)
children.push(h("5.  Entregables", HeadingLevel.HEADING_1));
const entTable = new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [600, 5160, 3600],
  borders,
  rows: [
    new TableRow({ tableHeader: true, children: [
      cell("N°", { bold: true, shade: NAVY, color: "FFFFFF", align: AlignmentType.CENTER, width: 600 }),
      cell("Entregable", { bold: true, shade: NAVY, color: "FFFFFF", width: 5160 }),
      cell("Formato", { bold: true, shade: NAVY, color: "FFFFFF", width: 3600 }),
    ]}),
    new TableRow({ children: [cell("1", { align: AlignmentType.CENTER, width: 600 }), cell("Flujo de caja del proyecto (mensual)", { width: 5160 }), cell("Excel dinámico", { width: 3600 })] }),
    new TableRow({ children: [cell("2", { align: AlignmentType.CENTER, width: 600 }), cell("Cronograma valorizado y calendario de desembolsos", { width: 5160 }), cell("Excel dinámico", { width: 3600 })] }),
    new TableRow({ children: [cell("3", { align: AlignmentType.CENTER, width: 600 }), cell("Estado de resultados proyectado e indicadores (VAN/TIR)", { width: 5160 }), cell("Excel dinámico", { width: 3600 })] }),
    new TableRow({ children: [cell("4", { align: AlignmentType.CENTER, width: 600 }), cell("Informe / memoria de sustento para el banco", { width: 5160 }), cell("Word y PDF", { width: 3600 })] }),
  ],
});
children.push(entTable);
children.push(p([t("")]));

// 6. Plazo
children.push(h("6.  Plazo de ejecución", HeadingLevel.HEADING_1));
children.push(p([
  t("El plazo total estimado es de "), ph("[15]"), t(" días calendario"),
  t(", contados desde la firma de la propuesta y la entrega completa de la información base por parte de la empresa, según el siguiente cronograma:"),
]));
const plazoTable = new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [5760, 3600],
  borders,
  rows: [
    new TableRow({ tableHeader: true, children: [
      cell("Etapa", { bold: true, shade: NAVY, color: "FFFFFF", width: 5760 }),
      cell("Plazo", { bold: true, shade: NAVY, color: "FFFFFF", align: AlignmentType.CENTER, width: 3600 }),
    ]}),
    new TableRow({ children: [cell("Levantamiento de información y supuestos", { width: 5760 }), cell([ph("[3] días"),], { align: AlignmentType.CENTER, width: 3600 })] }),
    new TableRow({ children: [cell("Construcción del modelo y borrador", { width: 5760 }), cell([ph("[7] días")], { align: AlignmentType.CENTER, width: 3600 })] }),
    new TableRow({ children: [cell("Revisión, ajustes y entrega final", { width: 5760 }), cell([ph("[5] días")], { align: AlignmentType.CENTER, width: 3600 })] }),
  ],
});
children.push(plazoTable);
children.push(p([t("")]));

// 7. Honorarios
children.push(h("7.  Honorarios profesionales", HeadingLevel.HEADING_1));
children.push(p([
  t("Los honorarios profesionales por el servicio integral ascienden a "),
  new TextRun({ text: "S/ ", bold: true, font: FONT }), ph("[15,000.00]"),
  t(" (más impuestos según corresponda), pagaderos "),
  new TextRun({ text: "por etapas", bold: true, font: FONT }),
  t(" conforme al siguiente detalle:"),
]));

const money = (v) => v; // placeholders
const honTable = new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [4260, 1500, 1800, 1800],
  borders,
  rows: [
    new TableRow({ tableHeader: true, children: [
      cell("Etapa / hito de pago", { bold: true, shade: NAVY, color: "FFFFFF", width: 4260 }),
      cell("%", { bold: true, shade: NAVY, color: "FFFFFF", align: AlignmentType.CENTER, width: 1500 }),
      cell("Subtotal (S/)", { bold: true, shade: NAVY, color: "FFFFFF", align: AlignmentType.RIGHT, width: 1800 }),
      cell("IGV 18% (S/)", { bold: true, shade: NAVY, color: "FFFFFF", align: AlignmentType.RIGHT, width: 1800 }),
    ]}),
    new TableRow({ children: [
      cell("Adelanto (a la firma de la propuesta)", { width: 4260 }),
      cell("40%", { align: AlignmentType.CENTER, width: 1500 }),
      cell([ph("[6,000.00]")], { align: AlignmentType.RIGHT, width: 1800 }),
      cell([ph("[1,080.00]")], { align: AlignmentType.RIGHT, width: 1800 }),
    ]}),
    new TableRow({ children: [
      cell("Entrega del borrador para revisión", { width: 4260 }),
      cell("30%", { align: AlignmentType.CENTER, width: 1500 }),
      cell([ph("[4,500.00]")], { align: AlignmentType.RIGHT, width: 1800 }),
      cell([ph("[810.00]")], { align: AlignmentType.RIGHT, width: 1800 }),
    ]}),
    new TableRow({ children: [
      cell("Entrega final aprobada", { width: 4260 }),
      cell("30%", { align: AlignmentType.CENTER, width: 1500 }),
      cell([ph("[4,500.00]")], { align: AlignmentType.RIGHT, width: 1800 }),
      cell([ph("[810.00]")], { align: AlignmentType.RIGHT, width: 1800 }),
    ]}),
    new TableRow({ children: [
      cell("Subtotal", { bold: true, shade: "D9E1F2", width: 4260 }),
      cell("100%", { bold: true, shade: "D9E1F2", align: AlignmentType.CENTER, width: 1500 }),
      cell([ph("[15,000.00]")], { bold: true, shade: "D9E1F2", align: AlignmentType.RIGHT, width: 1800 }),
      cell([ph("[2,700.00]")], { bold: true, shade: "D9E1F2", align: AlignmentType.RIGHT, width: 1800 }),
    ]}),
    new TableRow({ children: [
      cell("TOTAL A PAGAR (incluye IGV)", { bold: true, shade: NAVY, color: "FFFFFF", width: 4260 }),
      cell("", { shade: NAVY, width: 1500 }),
      cell([new TextRun({ text: "S/ ", bold: true, color: "FFFFFF", font: FONT }), new TextRun({ text: "[17,700.00]", bold: true, color: "FFFFFF", font: FONT })], { shade: NAVY, align: AlignmentType.RIGHT, width: 1800 }),
      cell("", { shade: NAVY, width: 1800 }),
    ]}),
  ],
});
children.push(honTable);

// Nota tributaria
children.push(new Paragraph({
  spacing: { before: 160, after: 160 },
  shading: { type: ShadingType.CLEAR, fill: "FFF2CC", color: "auto" },
  border: { top: NB, bottom: NB, left: NB, right: NB },
  children: [
    new TextRun({ text: "Nota tributaria (importante): ", bold: true, font: FONT, color: PH, size: 18 }),
    new TextRun({ text: "el tratamiento del impuesto depende de cómo emita sus comprobantes. ", font: FONT, size: 18 }),
    new TextRun({ text: "(a) Si emite Recibo por Honorarios (renta de 4.ª categoría), el servicio NO está afecto a IGV; en su lugar, la empresa podría retener el 8% de renta. ", font: FONT, size: 18 }),
    new TextRun({ text: "(b) Si emite Factura (persona natural con negocio / 3.ª categoría), aplica IGV 18%. ", font: FONT, size: 18 }),
    new TextRun({ text: "Elimine la columna/línea que no corresponda a su régimen antes de enviar la cotización.", font: FONT, size: 18, italics: true, color: GREY }),
  ],
}));

// 8. Forma de pago
children.push(h("8.  Forma y condiciones de pago", HeadingLevel.HEADING_1));
children.push(bullet([t("Moneda: Soles (S/). Pagos mediante "), ph("[transferencia / depósito]"), t(" a la cuenta que se indicará en cada comprobante.")]));
children.push(bullet("Cada hito se factura contra su cumplimiento; la entrega final se realiza una vez cancelado el saldo pendiente."));
children.push(bullet([t("Plazo de pago: "), ph("[a la presentación del comprobante / a X días]"), t(".")]));

// 9. Condiciones generales
children.push(h("9.  Condiciones generales", HeadingLevel.HEADING_1));
children.push(bullet("La calidad de los resultados depende de la veracidad y oportunidad de la información proporcionada por la empresa."));
children.push(bullet("El modelo financiero constituye una proyección basada en supuestos; no garantiza la aprobación del financiamiento, la cual es potestad exclusiva del banco."));
children.push(bullet("La información del proyecto será tratada con confidencialidad y utilizada únicamente para los fines de esta consultoría."));
children.push(bullet([t("Se incluye "), ph("[una (1)]"), t(" ronda de ajustes sobre el borrador. Rondas o alcances adicionales se cotizan por separado.")]));

// 10. Exclusiones
children.push(h("10.  Exclusiones", HeadingLevel.HEADING_1));
children.push(bullet("Elaboración de expedientes técnicos, tasaciones, estudios de mercado o valorizaciones de obra ejecutadas en campo."));
children.push(bullet("Gestión o trámite directo ante la entidad bancaria y gastos notariales, registrales o de tasación que el banco requiera."));

// 11. Validez
children.push(h("11.  Validez de la oferta", HeadingLevel.HEADING_1));
children.push(p([t("La presente cotización tiene una validez de "), ph("[15]"), t(" días calendario a partir de su fecha de emisión.")]));

// Aceptación / firmas
children.push(new Paragraph({ children: [new PageBreak()] }));
children.push(h("Aceptación", HeadingLevel.HEADING_1));
children.push(p("En señal de conformidad con los términos, alcances y condiciones económicas de la presente propuesta, las partes suscriben:"));
children.push(p([t("")]));
children.push(p([t("")]));
children.push(p([t("")]));

const firmaTable = new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [4680, 4680],
  borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE }, insideHorizontal: { style: BorderStyle.NONE }, insideVertical: { style: BorderStyle.NONE } },
  rows: [
    new TableRow({ children: [
      cell([new TextRun({ text: "_______________________________", font: FONT })], { width: 4680, align: AlignmentType.CENTER }),
      cell([new TextRun({ text: "_______________________________", font: FONT })], { width: 4680, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      cell([ph("[TU NOMBRE COMPLETO]")], { width: 4680, align: AlignmentType.CENTER }),
      cell([ph("[REPRESENTANTE DE LA EMPRESA]")], { width: 4680, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      cell([new TextRun({ text: "El Consultor", font: FONT, color: GREY, size: 18 })], { width: 4680, align: AlignmentType.CENTER }),
      cell([new TextRun({ text: "La Empresa", font: FONT, color: GREY, size: 18 })], { width: 4680, align: AlignmentType.CENTER }),
    ]}),
  ],
});
children.push(firmaTable);

// =============================================================
const doc = new Document({
  creator: "Consultor",
  title: "Cotización de Servicios Profesionales",
  styles: {
    default: {
      document: { run: { font: FONT, size: 21 } },
    },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 24, bold: true, color: NAVY }, paragraph: { spacing: { before: 240, after: 120 } } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 22, bold: true, color: BLUE }, paragraph: { spacing: { before: 160, after: 80 } } },
    ],
  },
  sections: [{
    properties: { page: { margin: { top: 1000, bottom: 1000, left: 1080, right: 1080 } } },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          border: { top: { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" } },
          children: [
            new TextRun({ text: "Cotización de servicios profesionales   ·   ", font: FONT, size: 16, color: GREY }),
            ph("[Tu nombre]"),
            new TextRun({ text: "   ·   Página ", font: FONT, size: 16, color: GREY }),
          ],
        })],
      }),
    },
    children,
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(OUT, buf);
  console.log("Guardado:", OUT);
});
