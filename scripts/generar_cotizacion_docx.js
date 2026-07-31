// Cotización de servicios profesionales - Enrique Huidobro -> Inversiones Estratégicas Integrales S.A.C.
// Proyecto Villa El Milagro (AVN - Techo Propio) - sustento para carta fianza ante aseguradoras.
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType,
  PageBreak, Footer, VerticalAlign,
} = require("docx");

const OUT = "/home/user/Eh-cloude/entregables/Cotizacion_Servicios_Profesionales.docx";

const NAVY = "1F3864";
const BLUE = "2E5496";
const GREY = "808080";
const PH = "C00000";
const FONT = "Calibri";

const ph = (t) => new TextRun({ text: t, bold: true, color: PH, font: FONT });
const t = (text, opts = {}) => new TextRun({ text, font: FONT, ...opts });

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
function cell(text, { bold, shade, color, align, width, size } = {}) {
  return new TableCell({
    width: width ? { size: width, type: WidthType.DXA } : undefined,
    verticalAlign: VerticalAlign.CENTER,
    shading: shade ? { type: ShadingType.CLEAR, fill: shade, color: "auto" } : undefined,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: [new Paragraph({
      alignment: align || AlignmentType.LEFT,
      children: Array.isArray(text) ? text
        : [new TextRun({ text, font: FONT, bold: !!bold, color: color || "000000", size: size || 20 })],
    })],
  });
}
const NB = { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" };
const borders = { top: NB, bottom: NB, left: NB, right: NB, insideHorizontal: NB, insideVertical: NB };

const children = [];

// Encabezado profesional
children.push(new Paragraph({ spacing: { after: 20 },
  children: [new TextRun({ text: "ENRIQUE HUIDOBRO", bold: true, font: FONT, color: NAVY, size: 26 })] }));
children.push(new Paragraph({ spacing: { after: 0 },
  children: [t("Economista  ·  Consultor financiero independiente", { color: GREY, size: 18 })] }));
children.push(new Paragraph({ spacing: { after: 0 },
  children: [t("RUC: 10723712039   ·   Tel: 974 293 131   ·   Email: chicoma.1998@gmail.com", { color: GREY, size: 18 })] }));
children.push(new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: NAVY } },
  spacing: { after: 200 }, children: [t("")] }));

// Título
children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 },
  children: [new TextRun({ text: "COTIZACIÓN DE SERVICIOS PROFESIONALES", bold: true, font: FONT, color: NAVY, size: 30 })] }));
children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 },
  children: [t("N° COT-2026-001        Fecha: 31/07/2026        Lima", { color: GREY })] }));

// Señores
children.push(new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 120, after: 100 },
  children: [new TextRun({ text: "Señores", font: FONT, bold: true, color: NAVY })] }));
children.push(new Table({
  width: { size: 9360, type: WidthType.DXA }, columnWidths: [2600, 6760], borders,
  rows: [
    new TableRow({ children: [cell("Razón social", { bold: true, shade: "D9E1F2", width: 2600 }), cell("INVERSIONES ESTRATÉGICAS INTEGRALES S.A.C.", { width: 6760 })] }),
    new TableRow({ children: [cell("RUC", { bold: true, shade: "D9E1F2", width: 2600 }), cell("20503680174", { width: 6760 })] }),
    new TableRow({ children: [cell("Atención", { bold: true, shade: "D9E1F2", width: 2600 }), cell("Sr. José Antonio Navarro Merea – Gerente General (Tel. 998 107 048)", { width: 6760 })] }),
    new TableRow({ children: [cell("Proyecto", { bold: true, shade: "D9E1F2", width: 2600 }), cell("VILLA EL MILAGRO – Adquisición de Vivienda Nueva (AVN), Programa Techo Propio", { width: 6760 })] }),
  ],
}));
children.push(p([t("")]));
children.push(new Paragraph({ spacing: { after: 200 }, children: [
  t("Ref.: ", { bold: true }),
  t("Elaboración del flujo de caja, cronograma valorizado y estado de resultados proyectado del proyecto de vivienda "),
  t("VILLA EL MILAGRO", { bold: true }),
  t(" (modalidad AVN – Programa Techo Propio), como sustento técnico-financiero para la obtención de una "),
  t("carta fianza ante empresas aseguradoras", { bold: true }), t("."),
]}));

function H1(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 240, after: 100 },
    children: [new TextRun({ text, font: FONT, bold: true, color: NAVY })] });
}
function H3(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 140, after: 60 },
    children: [new TextRun({ text, font: FONT, bold: true, color: BLUE })] });
}

children.push(H1("1.  Presentación"));
children.push(p("Estimados señores:"));
children.push(p([
  t("Es grato dirigirme a ustedes para presentar mi propuesta de servicios profesionales de consultoría financiera, orientada a estructurar la información económico-financiera del proyecto "),
  t("VILLA EL MILAGRO", { bold: true }),
  t(" (modalidad Adquisición de Vivienda Nueva – AVN del Programa Techo Propio), con el objetivo de sustentar técnicamente la solicitud de una "),
  t("carta fianza ante empresas aseguradoras", { bold: true }),
  t(", requerida para el desarrollo y la garantía del proyecto. Los entregables han sido diseñados conforme a la información que las áreas de riesgos y suscripción de las aseguradoras evalúan para la emisión de este tipo de garantías."),
]));

children.push(H1("2.  Objeto de la propuesta"));
children.push(p("Proveer a la empresa de un modelo financiero integral y de la documentación de sustento necesaria para gestionar la carta fianza ante la aseguradora, respaldando la viabilidad del proyecto y su capacidad de cumplimiento."));

children.push(H1("3.  Alcance de los servicios"));
children.push(p("La consultoría comprende los siguientes componentes:"));
children.push(H3("3.1  Flujo de caja del proyecto"));
children.push(bullet("Proyección mensual de ingresos y egresos durante todo el horizonte del proyecto."));
children.push(bullet("Ingresos por ventas —Bono Familiar Habitacional (Techo Propio), ahorro del beneficiario y crédito complementario, según corresponda— y por los desembolsos asociados al financiamiento del proyecto."));
children.push(bullet("Egresos por terreno, construcción, licencias y proyectos, comercialización, gastos administrativos y servicio de la deuda."));
children.push(bullet("Flujo de caja económico (sin financiamiento) y financiero (con financiamiento), con saldos acumulados."));
children.push(H3("3.2  Cronograma valorizado de obra"));
children.push(bullet("Valorización mensual de la obra según curva de avance físico (perfil S)."));
children.push(bullet("Determinación del calendario de desembolsos y del dimensionamiento de la carta fianza en función del avance del proyecto."));
children.push(bullet("Avance acumulado y saldo por ejecutar por período."));
children.push(H3("3.3  Estado de resultados proyectado"));
children.push(bullet("Estado de resultados del proyecto: ventas, costo de ventas, utilidad bruta, gastos operativos, resultado financiero e impuestos."));
children.push(bullet("Indicadores de rentabilidad: márgenes bruto y neto, ROI, y evaluación con VAN y TIR (económica y financiera)."));
children.push(H3("3.4  Informe / memoria de sustento para la aseguradora"));
children.push(bullet("Documento explicativo que integra los supuestos, la metodología y los resultados del modelo financiero."));
children.push(bullet("Análisis de sensibilidad básico y conclusiones sobre la viabilidad y la capacidad de cumplimiento del proyecto, orientado a la entidad emisora de la carta fianza."));

children.push(H1("4.  Metodología de trabajo"));
children.push(bullet("Reunión de levantamiento de información y definición de supuestos con el equipo de la empresa."));
children.push(bullet("Construcción del modelo en hoja de cálculo dinámica y parametrizable (los resultados se recalculan al variar los supuestos)."));
children.push(bullet("Presentación de un borrador para revisión y ajustes conjuntos."));
children.push(bullet("Entrega final de los modelos e informe, con acompañamiento para su presentación ante la aseguradora."));

children.push(H1("5.  Entregables"));
children.push(new Table({
  width: { size: 9360, type: WidthType.DXA }, columnWidths: [600, 5160, 3600], borders,
  rows: [
    new TableRow({ tableHeader: true, children: [
      cell("N°", { bold: true, shade: NAVY, color: "FFFFFF", align: AlignmentType.CENTER, width: 600 }),
      cell("Entregable", { bold: true, shade: NAVY, color: "FFFFFF", width: 5160 }),
      cell("Formato", { bold: true, shade: NAVY, color: "FFFFFF", width: 3600 }),
    ]}),
    new TableRow({ children: [cell("1", { align: AlignmentType.CENTER, width: 600 }), cell("Flujo de caja del proyecto (mensual)", { width: 5160 }), cell("Excel dinámico", { width: 3600 })] }),
    new TableRow({ children: [cell("2", { align: AlignmentType.CENTER, width: 600 }), cell("Cronograma valorizado y calendario de desembolsos", { width: 5160 }), cell("Excel dinámico", { width: 3600 })] }),
    new TableRow({ children: [cell("3", { align: AlignmentType.CENTER, width: 600 }), cell("Estado de resultados proyectado e indicadores (VAN/TIR)", { width: 5160 }), cell("Excel dinámico", { width: 3600 })] }),
    new TableRow({ children: [cell("4", { align: AlignmentType.CENTER, width: 600 }), cell("Informe / memoria de sustento para la aseguradora", { width: 5160 }), cell("Word y PDF", { width: 3600 })] }),
  ],
}));
children.push(p([t("")]));

children.push(H1("6.  Plazo de ejecución"));
children.push(p([t("El plazo total estimado es de "), t("15 días calendario", { bold: true }),
  t(", contados desde la firma de la propuesta y la entrega completa de la información base por parte de la empresa, según el siguiente cronograma:")]));
children.push(new Table({
  width: { size: 9360, type: WidthType.DXA }, columnWidths: [5760, 3600], borders,
  rows: [
    new TableRow({ tableHeader: true, children: [
      cell("Etapa", { bold: true, shade: NAVY, color: "FFFFFF", width: 5760 }),
      cell("Plazo", { bold: true, shade: NAVY, color: "FFFFFF", align: AlignmentType.CENTER, width: 3600 }),
    ]}),
    new TableRow({ children: [cell("Levantamiento de información y supuestos", { width: 5760 }), cell("3 días", { align: AlignmentType.CENTER, width: 3600 })] }),
    new TableRow({ children: [cell("Construcción del modelo y borrador", { width: 5760 }), cell("10 días", { align: AlignmentType.CENTER, width: 3600 })] }),
    new TableRow({ children: [cell("Revisión, ajustes y entrega final", { width: 5760 }), cell("5 días", { align: AlignmentType.CENTER, width: 3600 })] }),
  ],
}));
children.push(p([t("")]));

children.push(H1("7.  Honorarios profesionales"));
children.push(p([
  t("Los honorarios profesionales por el servicio integral ascienden a "),
  t("S/ 10,000.00", { bold: true }), t(" (diez mil y 00/100 soles), pagaderos "),
  t("por etapas", { bold: true }), t(" conforme al siguiente detalle:"),
]));
children.push(new Table({
  width: { size: 9360, type: WidthType.DXA }, columnWidths: [5460, 1300, 2600], borders,
  rows: [
    new TableRow({ tableHeader: true, children: [
      cell("Etapa / hito de pago", { bold: true, shade: NAVY, color: "FFFFFF", width: 5460 }),
      cell("%", { bold: true, shade: NAVY, color: "FFFFFF", align: AlignmentType.CENTER, width: 1300 }),
      cell("Importe (S/)", { bold: true, shade: NAVY, color: "FFFFFF", align: AlignmentType.RIGHT, width: 2600 }),
    ]}),
    new TableRow({ children: [cell("Adelanto (a la firma de la propuesta)", { width: 5460 }), cell("20%", { align: AlignmentType.CENTER, width: 1300 }), cell("2,000.00", { align: AlignmentType.RIGHT, width: 2600 })] }),
    new TableRow({ children: [cell("Entrega del borrador para revisión", { width: 5460 }), cell("30%", { align: AlignmentType.CENTER, width: 1300 }), cell("3,000.00", { align: AlignmentType.RIGHT, width: 2600 })] }),
    new TableRow({ children: [cell("Entrega final aprobada", { width: 5460 }), cell("50%", { align: AlignmentType.CENTER, width: 1300 }), cell("5,000.00", { align: AlignmentType.RIGHT, width: 2600 })] }),
    new TableRow({ children: [
      cell("TOTAL", { bold: true, shade: NAVY, color: "FFFFFF", width: 5460 }),
      cell("100%", { bold: true, shade: NAVY, color: "FFFFFF", align: AlignmentType.CENTER, width: 1300 }),
      cell([new TextRun({ text: "S/ 10,000.00", bold: true, color: "FFFFFF", font: FONT })], { shade: NAVY, align: AlignmentType.RIGHT, width: 2600 }),
    ]}),
  ],
}));
children.push(new Paragraph({
  spacing: { before: 160, after: 160 },
  shading: { type: ShadingType.CLEAR, fill: "FFF2CC", color: "auto" },
  border: { top: NB, bottom: NB, left: NB, right: NB },
  children: [
    new TextRun({ text: "Condición tributaria: ", bold: true, font: FONT, color: NAVY, size: 18 }),
    new TextRun({ text: "el consultor emite Recibo por Honorarios (renta de 4.ª categoría) y cuenta con suspensión de retenciones vigente; por lo tanto, el servicio no está afecto a IGV ni a retención de renta. El importe indicado es el monto neto a transferir.", font: FONT, size: 18 }),
  ],
}));

children.push(H1("8.  Forma y condiciones de pago"));
children.push(bullet([t("Moneda: Soles (S/). Pagos mediante "), t("transferencia bancaria", { bold: true }), t(" a la cuenta que se indicará en cada Recibo por Honorarios.")]));
children.push(bullet("Cada hito se cobra contra su cumplimiento; la entrega final se realiza una vez cancelado el saldo pendiente."));
children.push(bullet("Plazo de pago: a la presentación del comprobante."));

children.push(H1("9.  Condiciones generales"));
children.push(bullet("La calidad de los resultados depende de la veracidad y oportunidad de la información proporcionada por la empresa."));
children.push(bullet("El modelo financiero constituye una proyección basada en supuestos; no garantiza la emisión de la carta fianza, la cual es potestad exclusiva de la aseguradora."));
children.push(bullet("La información del proyecto será tratada con confidencialidad y utilizada únicamente para los fines de esta consultoría."));
children.push(bullet("Se incluye una (1) ronda de ajustes sobre el borrador. Rondas o alcances adicionales se cotizan por separado."));

children.push(H1("10.  Exclusiones"));
children.push(bullet("Elaboración de expedientes técnicos, tasaciones, estudios de mercado o el registro del proyecto ante el Fondo MIVIVIENDA / Techo Propio."));
children.push(bullet("Gestión o trámite directo ante la aseguradora, y los costos de emisión y comisión de la carta fianza."));

children.push(H1("11.  Validez de la oferta"));
children.push(p([t("La presente cotización tiene una validez de "), t("15 días calendario", { bold: true }), t(" a partir de su fecha de emisión.")]));

children.push(new Paragraph({ children: [new PageBreak()] }));
children.push(H1("Aceptación"));
children.push(p("En señal de conformidad con los términos, alcances y condiciones económicas de la presente propuesta, las partes suscriben:"));
children.push(p([t("")])); children.push(p([t("")])); children.push(p([t("")]));
children.push(new Table({
  width: { size: 9360, type: WidthType.DXA }, columnWidths: [4680, 4680],
  borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE }, insideHorizontal: { style: BorderStyle.NONE }, insideVertical: { style: BorderStyle.NONE } },
  rows: [
    new TableRow({ children: [
      cell([new TextRun({ text: "_______________________________", font: FONT })], { width: 4680, align: AlignmentType.CENTER }),
      cell([new TextRun({ text: "_______________________________", font: FONT })], { width: 4680, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      cell([new TextRun({ text: "ENRIQUE HUIDOBRO", bold: true, font: FONT })], { width: 4680, align: AlignmentType.CENTER }),
      cell([new TextRun({ text: "José Antonio Navarro Merea", bold: true, font: FONT })], { width: 4680, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      cell([new TextRun({ text: "Economista · RUC 10723712039 · El Consultor", font: FONT, color: GREY, size: 16 })], { width: 4680, align: AlignmentType.CENTER }),
      cell([new TextRun({ text: "Gerente General · Inversiones Estratégicas Integrales S.A.C.", font: FONT, color: GREY, size: 16 })], { width: 4680, align: AlignmentType.CENTER }),
    ]}),
  ],
}));

const doc = new Document({
  creator: "Enrique Huidobro",
  title: "Cotización de Servicios Profesionales - Villa El Milagro",
  styles: { default: { document: { run: { font: FONT, size: 21 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 24, bold: true, color: NAVY }, paragraph: { spacing: { before: 240, after: 120 } } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 22, bold: true, color: BLUE }, paragraph: { spacing: { before: 160, after: 80 } } },
    ],
  },
  sections: [{
    properties: { page: { margin: { top: 1000, bottom: 1000, left: 1080, right: 1080 } } },
    footers: { default: new Footer({ children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      border: { top: { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" } },
      children: [ new TextRun({ text: "Cotización COT-2026-001   ·   Enrique Huidobro   ·   Villa El Milagro", font: FONT, size: 16, color: GREY }) ],
    })] }) },
    children,
  }],
});

Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(OUT, buf); console.log("Guardado:", OUT); });
