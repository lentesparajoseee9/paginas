const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, AlignmentType, Footer, PageNumber, LevelFormat,
} = require("docx");
const C = require("./contenido_simple.js");

const OUT = path.join(__dirname, "Trabajo_Entrenador_Deportivo.docx");
const FONT = "Arial";
const LINE = 360; // 1.5

function runs(s, base = {}) {
  return s.split(/(\*[^*]+\*)/).filter(Boolean).map(t =>
    t.startsWith("*") ? new TextRun({ text: t.slice(1, -1), bold: true, ...base }) : new TextRun({ text: t, ...base }));
}
const para = s => new Paragraph({ children: runs(s), alignment: AlignmentType.JUSTIFIED, spacing: { line: LINE, after: 160 } });
const li = s => new Paragraph({ children: runs(s), numbering: { reference: "bul", level: 0 }, alignment: AlignmentType.JUSTIFIED, spacing: { line: LINE, after: 80 } });
const center = (s, r = {}, after = 0) => new Paragraph({ alignment: AlignmentType.CENTER, spacing: { line: LINE, after }, children: [new TextRun({ text: s, ...r })] });
const blank = n => Array.from({ length: n }, () => new Paragraph({ children: [] }));
const title = (s, pb) => new Paragraph({ pageBreakBefore: pb, alignment: AlignmentType.CENTER, spacing: { after: 240 }, keepNext: true, children: [new TextRun({ text: s.toUpperCase(), bold: true, size: 28 })] });
const sub = s => new Paragraph({ spacing: { before: 200, after: 120 }, keepNext: true, children: [new TextRun({ text: s, bold: true, size: 26, color: "1F3864" })] });
const sub2 = s => new Paragraph({ spacing: { before: 120, after: 60 }, keepNext: true, children: [new TextRun({ text: s, bold: true, italics: true })] });

const P = C.portada;
const portada = [
  center(P.escuela, { bold: true, size: 28 }),
  center(P.materia, { size: 24 }),
  ...blank(8),
  center(P.titulo, { bold: true, size: 44 }, 0),
  ...blank(9),
  center(`Alumno(a): ${P.alumno}`),
  center(`Grupo: ${P.grupo}`),
  center(`Profesor(a): ${P.docente}`),
  center(`Fecha: ${P.fecha}`),
];

const body = [title("Introducción", true), ...C.introduccion.map(para), title("Desarrollo", true)];
for (const x of C.desarrollo) {
  if (typeof x === "string") body.push(para(x));
  else if (x.h) body.push(sub(x.h));
  else if (x.h2) body.push(sub2(x.h2));
  else if (x.li) body.push(li(x.li));
}
body.push(title("Conclusión", true), ...C.conclusion.map(para));
body.push(title("Bibliografía", true), ...C.bibliografia.map(r => new Paragraph({
  children: r.split(/(\*[^*]+\*)/).filter(Boolean).map(t => t.startsWith("*") ? new TextRun({ text: t.slice(1, -1), italics: true }) : new TextRun(t)),
  spacing: { line: LINE, after: 120 }, indent: { left: 709, hanging: 709 },
})));

const doc = new Document({
  creator: "", title: P.titulo,
  styles: { default: { document: { run: { font: FONT, size: 24 } } } },
  numbering: { config: [{ reference: "bul", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] }] },
  sections: [{
    properties: { titlePage: true, page: { size: { width: 12240, height: 15840 }, margin: { top: 1418, bottom: 1418, left: 1418, right: 1418 } } },
    footers: {
      default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT] })] })] }),
      first: new Footer({ children: [new Paragraph({ children: [] })] }),
    },
    children: [...portada, ...body],
  }],
});
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("ok"); });
