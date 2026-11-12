const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, AlignmentType, Header, PageNumber, PageBreak,
  Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, LevelFormat,
  HeadingLevel, TabStopType, TabStopPosition,
} = require("docx");
const C = require("./contenido.js");

C.tablas.cualidades.nota = "Elaboración propia.";

const OUT = process.argv[2] || path.join(__dirname, "Trabajo_Entrenador_Deportivo.docx");
const PAGES = fs.existsSync(path.join(__dirname, "paginas.json"))
  ? JSON.parse(fs.readFileSync(path.join(__dirname, "paginas.json"), "utf8")) : {};

const FONT = "Times New Roman";
const LINE = 360; // 1.5
const IND = 709;  // 1.25 cm

// *cursiva* -> runs
function runs(s, base = {}) {
  return s.split(/(\*[^*]+\*)/).filter(Boolean).map(t =>
    t.startsWith("*") ? new TextRun({ text: t.slice(1, -1), italics: true, ...base }) : new TextRun({ text: t, ...base }));
}
const para = (s, o = {}) => new Paragraph({
  children: runs(s), alignment: AlignmentType.JUSTIFIED,
  indent: { firstLine: IND }, spacing: { line: LINE, after: 0 }, ...o,
});
const center = (s, o = {}, r = {}) => new Paragraph({ alignment: AlignmentType.CENTER, spacing: { line: LINE, after: 0 }, ...o, children: runs(s, r) });
const blank = (n = 1) => Array.from({ length: n }, () => new Paragraph({ spacing: { line: LINE, after: 0 }, children: [] }));
const H1 = (s, pb = false) => new Paragraph({ heading: HeadingLevel.HEADING_1, pageBreakBefore: pb, children: [new TextRun(s)] });
const H2 = s => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(s)] });
const H3 = s => new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun(s)] });

const border = { style: BorderStyle.SINGLE, size: 6, color: "000000" };
const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
function table(t) {
  const cell = (txt, w, head, last) => new TableCell({
    width: { size: w, type: WidthType.DXA },
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    shading: head ? { type: ShadingType.CLEAR, color: "auto", fill: "E7E6E6" } : undefined,
    borders: { top: head ? border : none, bottom: head || last ? border : none, left: none, right: none },
    children: [new Paragraph({ spacing: { line: 240, after: 0 }, children: [new TextRun({ text: txt, bold: head, size: 22 })] })],
  });
  const total = t.widths.reduce((a, b) => a + b, 0);
  return [
    new Paragraph({ spacing: { before: 240, line: LINE, after: 0 }, keepNext: true, children: [new TextRun({ text: `Tabla ${t.num}`, bold: true })] }),
    new Paragraph({ spacing: { line: LINE, after: 60 }, keepNext: true, children: [new TextRun({ text: t.titulo, italics: true })] }),
    new Table({
      width: { size: total, type: WidthType.DXA }, columnWidths: t.widths,
      rows: [
        new TableRow({ tableHeader: true, children: t.head.map((h, i) => cell(h, t.widths[i], true, false)) }),
        ...t.rows.map((r, ri) => new TableRow({ cantSplit: true, children: r.map((c, i) => cell(c, t.widths[i], false, ri === t.rows.length - 1)) })),
      ],
    }),
    new Paragraph({ spacing: { before: 60, line: 240, after: 240 }, children: [new TextRun({ text: "Nota. ", italics: true, size: 20 }), new TextRun({ text: t.nota, size: 20 })] }),
  ];
}

// ---------- Portada ----------
const P = C.portada;
const portada = [
  center(P.institucion, {}, { bold: true, size: 28 }),
  center(P.facultad, {}, { size: 26 }),
  ...blank(4),
  center(P.titulo, {}, { bold: true, size: 30 }),
  ...blank(2),
  center("Trabajo de investigación documental", {}, { italics: true }),
  ...blank(5),
  center(`Materia: ${P.materia}`),
  center(`Presenta: ${P.alumno}`),
  center(`Docente: ${P.docente}`),
  ...blank(4),
  center(P.lugar),
];

// ---------- Índice ----------
const idx = [
  ["Introducción", 0], ["Desarrollo", 0],
  ...C.desarrollo.filter(x => x.h || x.h2).map(x => [x.h || x.h2, x.h ? 1 : 2]),
  ["Conclusiones", 0], ["Referencias", 0],
];
const indice = [
  new Paragraph({ pageBreakBefore: true, alignment: AlignmentType.CENTER, spacing: { after: 240 }, children: [new TextRun({ text: "Índice", bold: true, size: 24 })] }),
  ...idx.map(([t, lvl]) => new Paragraph({
    spacing: { line: 300, after: 40 }, indent: { left: lvl * 400 },
    tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX, leader: "dot" }],
    children: [new TextRun({ text: t, bold: lvl === 0 }), new TextRun({ text: "\t" + String(PAGES[t] ?? "") })],
  })),
];

// ---------- Cuerpo ----------
const body = [H1("Introducción", true), ...C.introduccion.map(s => para(s))];
body.push(H1("Desarrollo", true));
for (const x of C.desarrollo) {
  if (typeof x === "string") body.push(para(x));
  else if (x.h) body.push(H2(x.h));
  else if (x.h2) body.push(H3(x.h2));
  else if (x.li) body.push(new Paragraph({ numbering: { reference: "bul", level: 0 }, alignment: AlignmentType.JUSTIFIED, spacing: { line: LINE, after: 0 }, children: runs(x.li) }));
  else if (x.table) body.push(...table(C.tablas[x.table]));
}
body.push(H1("Conclusiones", true), ...C.conclusiones.map(s => para(s)));
body.push(H1("Referencias", true), ...C.referencias.map(r => new Paragraph({
  children: runs(r), spacing: { line: LINE, after: 0 }, indent: { left: IND, hanging: IND },
})));

const head = (size, before, after, extra = {}) => ({
  run: { font: FONT, size, bold: true, color: "000000", ...extra },
  paragraph: { spacing: { before, after, line: LINE }, keepNext: true, keepLines: true },
});

const doc = new Document({
  creator: "", title: P.titulo,
  styles: {
    default: { document: { run: { font: FONT, size: 24 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, ...head(24, 0, 240), paragraph: { alignment: AlignmentType.CENTER, spacing: { before: 0, after: 240, line: LINE }, keepNext: true, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, ...head(24, 240, 60), paragraph: { spacing: { before: 240, after: 60, line: LINE }, keepNext: true, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: FONT, size: 24, bold: true, italics: true, color: "000000" }, paragraph: { spacing: { before: 120, after: 0, line: LINE }, keepNext: true, outlineLevel: 2 } },
    ],
  },
  numbering: { config: [{ reference: "bul", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] }] },
  sections: [{
    properties: {
      titlePage: true,
      page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, bottom: 1440, left: 1440, right: 1440, header: 708 } },
    },
    headers: {
      default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ children: [PageNumber.CURRENT] })] })] }),
      first: new Header({ children: [new Paragraph({ children: [] })] }),
    },
    children: [...portada, ...indice, ...body],
  }],
});

Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("ok", OUT); });
