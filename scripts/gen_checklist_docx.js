// 產出「公開發行○○股份有限公司114年度財務報告公告檢查表—管區意見」Word 檔
// 用法：node gen_checklist_docx.js <content.json> <out.docx>
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, AlignmentType, VerticalAlign, HeadingLevel, PageBreak,
} = require("docx");

const [, , contentPath, outPath] = process.argv;
const C = JSON.parse(fs.readFileSync(contentPath, "utf-8"));

const FONT = "PMingLiU"; // 新細明體，貼近原模板
const sz = (n) => n * 2;  // pt → half-pt

function runs(text, opts = {}) {
  return new TextRun({ text, font: { name: FONT, eastAsia: FONT }, size: sz(opts.size || 12), bold: !!opts.bold, color: opts.color });
}
function para(text, opts = {}) {
  return new Paragraph({
    alignment: opts.align,
    spacing: { after: opts.after ?? 80, line: opts.line },
    indent: opts.indent,
    children: Array.isArray(text) ? text : [runs(text, opts)],
  });
}

// ---- 檢查表 ----
const W = { g: 1300, t: 5200, y: 900, n: 900, b: 1740 };
const TOTAL = W.g + W.t + W.y + W.n + W.b;

function cell(children, width, opts = {}) {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    verticalAlign: VerticalAlign.CENTER,
    columnSpan: opts.span,
    rowSpan: opts.rowSpan,
    margins: { top: 60, bottom: 60, left: 80, right: 80 },
    children,
  });
}

const headerRows = [
  new TableRow({
    tableHeader: true,
    children: [
      cell([para("內容", { align: AlignmentType.CENTER, bold: true, after: 0 })], W.g, { rowSpan: 2 }),
      cell([para("檢查內容", { align: AlignmentType.CENTER, bold: true, after: 0 })], W.t, { rowSpan: 2 }),
      cell([para("填報項目", { align: AlignmentType.CENTER, bold: true, after: 0 })], W.y + W.n + W.b, { span: 3 }),
    ],
  }),
  new TableRow({
    tableHeader: true,
    children: [
      cell([para("是", { align: AlignmentType.CENTER, bold: true, after: 0 }), para("(正常)", { align: AlignmentType.CENTER, size: 10, after: 0 })], W.y),
      cell([para("否", { align: AlignmentType.CENTER, bold: true, after: 0 }), para("(異常)", { align: AlignmentType.CENTER, size: 10, after: 0 })], W.n),
      cell([para("備註", { align: AlignmentType.CENTER, bold: true, after: 0 })], W.b),
    ],
  }),
];

const bodyRows = [];
for (const g of C.groups) {
  g.items.forEach((it, idx) => {
    const cells = [];
    if (idx === 0) cells.push(cell([para(g.group, { after: 0 })], W.g, { rowSpan: g.items.length }));
    cells.push(cell(it.text.split("\n").map((t) => para(t, { after: 0, size: 11 })), W.t));
    cells.push(cell([para(it.mark === "yes" ? "V" : "", { align: AlignmentType.CENTER, after: 0 })], W.y));
    cells.push(cell([para(it.mark === "no" ? "V" : "", { align: AlignmentType.CENTER, after: 0 })], W.n));
    cells.push(cell((it.note || "").split("\n").filter(Boolean).map((t) => para(t, { after: 0, size: 10 })), W.b));
    bodyRows.push(new TableRow({ children: cells }));
  });
}

const table = new Table({
  width: { size: TOTAL, type: WidthType.DXA },
  columnWidths: [W.g, W.t, W.y, W.n, W.b],
  rows: [...headerRows, ...bodyRows],
});

// ---- 後段說明 ----
function miniTable(rows) {
  const n = rows[0].length;
  const total = 9200;
  const widths = rows[0].map((_, i) => (i === n - 1 ? total - Math.floor(total / (n + 1)) * (n - 1) : Math.floor(total / (n + 1))));
  return new Table({
    width: { size: total, type: WidthType.DXA },
    columnWidths: widths,
    rows: rows.map((r, ri) => new TableRow({
      children: r.map((t, i) => cell([para(t, { after: 0, size: 10, bold: ri === 0, align: i < n - 1 && ri > 0 ? undefined : AlignmentType.CENTER })], widths[i])),
    })),
  });
}

const sectionParas = [];
for (const s of C.sections) {
  sectionParas.push(para(s.title, { bold: true, size: 13, after: 120 }));
  for (const b of s.body) {
    if (b.h) sectionParas.push(para(b.h, { bold: true, indent: { left: 360 }, after: 60 }));
    let tblRows = [];
    const flush = () => {
      if (tblRows.length) { sectionParas.push(miniTable(tblRows)); sectionParas.push(para("", { after: 40 })); tblRows = []; }
    };
    for (const p of b.paras || []) {
      if (p.startsWith("【表】")) { tblRows.push(p.slice(3).split("｜")); continue; }
      flush();
      sectionParas.push(para(p, { indent: { left: b.h ? 720 : 480 }, after: 80, line: 320 }));
    }
    flush();
  }
  sectionParas.push(para("", { after: 60 }));
}

const notes = (C.footnotes || []).map((t) => para(t, { size: 10, after: 40 }));

const doc = new Document({
  sections: [{
    properties: { page: { margin: { top: 1080, bottom: 1080, left: 1080, right: 1080 } } },
    children: [
      para(C.title, { align: AlignmentType.CENTER, bold: true, size: 15, after: 200 }),
      table,
      ...notes,
      new Paragraph({ children: [new PageBreak()] }),
      ...sectionParas,
    ],
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(outPath, buf);
  console.log("寫出", outPath, buf.length, "bytes");
});
