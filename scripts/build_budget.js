const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, PageBreak,
} = require("docx");
const fs = require("fs");

const FONT = "Arial";
const ORANGE = "E07C16";
const ORANGE_L = "FEF3E6";
const SOFT = "FBF3E9";
const GREY = "6D6E70";
const INK = "26262A";
const LINE = "D9D7D3";
const HEAD_BG = "F3EFEA";
const GREEN = "1E7E45";
const RED = "B23A34";
const WARN = "9A6A12";

function R(text, o = {}) {
  return new TextRun({ text, font: FONT, rightToLeft: true, size: o.size || 22,
    bold: o.bold, italics: o.italics, color: o.color || INK });
}
function P(runs, o = {}) {
  if (typeof runs === "string") runs = [R(runs, o)];
  return new Paragraph({ bidirectional: true, alignment: o.align || AlignmentType.RIGHT,
    spacing: { after: o.after != null ? o.after : 120, before: o.before || 0, line: o.line || 300 },
    ...(o.shading ? { shading: o.shading } : {}), ...(o.border ? { border: o.border } : {}), children: runs });
}
function H1(num, text) {
  return new Paragraph({ bidirectional: true, alignment: AlignmentType.RIGHT, heading: HeadingLevel.HEADING_1,
    spacing: { before: 300, after: 130 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: ORANGE, space: 4 } },
    children: [ R(num + "  ", { bold: true, size: 26, color: ORANGE }), R(text, { bold: true, size: 26, color: INK }) ] });
}
function cellP(text, o = {}) {
  return new TableCell({ width: { size: o.w, type: WidthType.DXA },
    shading: o.shade ? { type: ShadingType.CLEAR, color: "auto", fill: o.shade } : undefined,
    margins: { top: 72, bottom: 72, left: 100, right: 100 }, verticalAlign: "center",
    children: [P([R(text, { size: o.size || 20, bold: o.bold, color: o.color || INK })],
      { after: 0, align: o.align || AlignmentType.RIGHT, line: 265 })] });
}
function hc(text, w) {
  return new TableCell({ width: { size: w, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, color: "auto", fill: HEAD_BG },
    margins: { top: 72, bottom: 72, left: 100, right: 100 }, verticalAlign: "center",
    children: [P([R(text, { bold: true, size: 20 })], { after: 0, align: AlignmentType.CENTER })] });
}
function table(widths, header, rows) {
  const trs = [ new TableRow({ tableHeader: true, children: header.map((h, i) => hc(h, widths[i])) }) ];
  for (const r of rows) trs.push(new TableRow({ children: r.map((c, i) =>
    cellP(c.t, { w: widths[i], align: c.align, bold: c.bold, color: c.color, shade: c.shade, size: c.size })) }));
  return new Table({ width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    columnWidths: widths, visuallyRightToLeft: true,
    borders: { top: { style: BorderStyle.SINGLE, size: 4, color: LINE }, bottom: { style: BorderStyle.SINGLE, size: 4, color: LINE },
      left: { style: BorderStyle.SINGLE, size: 4, color: LINE }, right: { style: BorderStyle.SINGLE, size: 4, color: LINE },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: LINE }, insideVertical: { style: BorderStyle.SINGLE, size: 4, color: LINE } },
    rows: trs });
}
const C = (t, o = {}) => ({ t, ...o }); // cell shorthand

const kids = [];

// ---- Header ----
kids.push(new Paragraph({ spacing: { after: 0 }, alignment: AlignmentType.RIGHT, bidirectional: true,
  children: [ new TextRun({ text: "درب لخدمات المحطات", font: FONT, rightToLeft: true, bold: true, size: 22, color: GREY }) ] }));
kids.push(new Paragraph({ spacing: { before: 60, after: 0 }, alignment: AlignmentType.RIGHT, bidirectional: true,
  children: [ new TextRun({ text: "ميزانيات ملف الشراكات — برنامج تانكي للولاء", font: FONT, rightToLeft: true, bold: true, size: 34, color: INK }) ] }));
kids.push(new Paragraph({ spacing: { before: 40, after: 0 }, alignment: AlignmentType.RIGHT, bidirectional: true,
  children: [ new TextRun({ text: "وثيقة رسمية للإدارة التنفيذية · تقديرات تخطيطية للسنة الأولى", font: FONT, rightToLeft: true, size: 22, color: ORANGE, bold: true }) ] }));
kids.push(new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { before: 120, after: 160 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 18, color: ORANGE, space: 8 } }, children: [] }));

// ---- Disclaimer box ----
kids.push(P([ R("تنبيه: ", { bold: true, color: WARN }),
  R("جميع الأرقام تقديرات تخطيطية أولية للسوق السعودي بالريال، تُضبط بعروض فعلية ورواتب الموارد البشرية الحقيقية. الغرض منها تحديد حجم الاستثمار المطلوب لا الالتزام برقم نهائي.") ],
  { shading: { type: ShadingType.CLEAR, color: "auto", fill: SOFT },
    border: { left: { style: BorderStyle.SINGLE, size: 18, color: WARN, space: 8 } }, after: 200, line: 290 }));

// ---- 1. HR ----
kids.push(H1("١.", "ميزانية الموارد البشرية (سنوية)"));
kids.push(P("الرواتب تقديرية شهرية قبل التحميلات؛ يُضاف نحو ٢٥٪ تحميلات (تأمينات ومزايا).", { size: 20, color: GREY, after: 120 }));
kids.push(table([2650, 1150, 2650, 2760],
  ["الدور", "العدد", "الراتب الشهري (﷼)", "السنوي التقريبي (﷼)"],
  [
    [C("مدير الشراكات", { bold: true }), C("١", { align: AlignmentType.CENTER }), C("٢٥٬٠٠٠ – ٤٠٬٠٠٠", { align: AlignmentType.CENTER }), C("٣٠٠ – ٤٨٠ ألف", { align: AlignmentType.CENTER })],
    [C("مدير تطوير أعمال (BD)", { bold: true }), C("٢", { align: AlignmentType.CENTER }), C("١٢٬٠٠٠ – ٢٠٬٠٠٠ /فرد", { align: AlignmentType.CENTER }), C("٢٨٨ – ٤٨٠ ألف", { align: AlignmentType.CENTER })],
    [C("محلل تجاري", { bold: true }), C("١", { align: AlignmentType.CENTER }), C("١٠٬٠٠٠ – ١٥٬٠٠٠", { align: AlignmentType.CENTER }), C("١٢٠ – ١٨٠ ألف", { align: AlignmentType.CENTER })],
    [C("مختص تسويات مالية", { bold: true }), C("١", { align: AlignmentType.CENTER }), C("٩٬٠٠٠ – ١٤٬٠٠٠", { align: AlignmentType.CENTER }), C("١٠٨ – ١٦٨ ألف", { align: AlignmentType.CENTER })],
    [C("الفريق الأساسي (٥ أفراد)", { bold: true, shade: ORANGE_L }), C("", { shade: ORANGE_L }), C("", { shade: ORANGE_L }), C("٠٫٩ – ١٫٣ مليون", { align: AlignmentType.CENTER, bold: true, shade: ORANGE_L })],
    [C("+ التحميلات (~٢٥٪)", { bold: true, shade: ORANGE_L }), C("", { shade: ORANGE_L }), C("", { shade: ORANGE_L }), C("≈ ١٫١ – ١٫٦ مليون", { align: AlignmentType.CENTER, bold: true, color: ORANGE, shade: ORANGE_L })],
  ]));
kids.push(P([ R("التوسّع لاحقاً ", { bold: true }),
  R("(مدير حسابات ×٢، تسويق شراكات، منسّق عمليات، دعم تكامل) يضيف نحو ٥٥٠ – ٩٠٠ ألف ﷼ سنوياً.") ], { before: 100 }));

// ---- 2. Opex/Capex ----
kids.push(H1("٢.", "الميزانية التشغيلية والرأسمالية (السنة الأولى)"));
kids.push(table([2550, 2350, 4310],
  ["البند", "القيمة التقديرية (﷼)", "ملاحظة"],
  [
    [C("الرصيد المقدّم / رأس المال العامل (Float)", { bold: true }), C("٢٠٠ – ٥٠٠ ألف+", { align: AlignmentType.CENTER }), C("لمخزون القسائم — قابل للاسترداد، ليس مصروفاً؛ يكبر مع الحجم.")],
    [C("التسويق والتمويل المشترك (Co-funding)", { bold: true }), C("٣٠٠ – ٨٠٠ ألف", { align: AlignmentType.CENTER }), C("حملة الإطلاق + تسويق مشترك مع التجار.")],
    [C("التقنية والأنظمة", { bold: true }), C("١٥٠ – ٤٥٠ ألف", { align: AlignmentType.CENTER }), C("تكامل Gift Card API والمحفظة (لمرة واحدة) + صيانة. أداة إدارة التجار جاهزة في تانكي.")],
    [C("الامتثال والترخيص والقانوني", { bold: true }), C("١٥٠ – ٤٠٠ ألف", { align: AlignmentType.CENTER }), C("إعداد شراكة ساما + صياغة العقود + الرقابة الشرعية + امتثال PDPL.")],
  ]));

// ---- 3. Totals ----
kids.push(H1("٣.", "الإجمالي التقديري للسنة الأولى"));
kids.push(table([5400, 3810],
  ["الفئة", "المدى التقديري (﷼)"],
  [
    [C("موارد بشرية (فريق أساسي، محمّل)"), C("١٫١ – ١٫٦ مليون", { align: AlignmentType.CENTER })],
    [C("تسويق وتمويل مشترك"), C("٠٫٣ – ٠٫٨ مليون", { align: AlignmentType.CENTER })],
    [C("تقنية وأنظمة"), C("٠٫١٥ – ٠٫٤٥ مليون", { align: AlignmentType.CENTER })],
    [C("امتثال وترخيص وقانوني"), C("٠٫١٥ – ٠٫٤ مليون", { align: AlignmentType.CENTER })],
    [C("الإجمالي التشغيلي (السنة ١)", { bold: true, shade: ORANGE_L }), C("≈ ١٫٦ – ٣٫٣ مليون", { align: AlignmentType.CENTER, bold: true, color: ORANGE, shade: ORANGE_L })],
    [C("+ رأس مال عامل (Float — قابل للاسترداد)", { bold: true, shade: SOFT }), C("≈ ٠٫٢ – ٠٫٥ مليون", { align: AlignmentType.CENTER, bold: true, shade: SOFT })],
  ]));
kids.push(P([ R("الحد الأدنى الواقعي للانطلاق: ", { bold: true, color: ORANGE }),
  R("نحو ١٫٦ مليون ﷼ تشغيلي + ٢٠٠ ألف ﷼ رأس مال عامل — بفريق لين وتسويق محافظ.") ], { before: 120 }));

// ---- 4. Notes ----
kids.push(H1("٤.", "ملاحظات مهمة"));
[
  ["متكرّر مقابل لمرة واحدة:", "الرواتب والصيانة والتسويق بنود متكرّرة سنوياً؛ تكامل التقنية والترخيص غالباً لمرة واحدة."],
  ["الرصيد المقدّم ليس مصروفاً:", "رأس مال عامل يعود إلى درب لكنه يشغل سيولة، فيُدرَج منفصلاً عن المصروفات."],
  ["تكلفة النقاط بند منفصل:", "التزام البرنامج (قيمة النقطة نصف هللة لكل ريال) محسوب في نموذج تانكي الاقتصادي، وليس ضمن تشغيل ملف الشراكات."],
  ["أكبر متغيّرين:", "حجم الفريق (توسّع مبكّر أو لين) وطموح التسويق والتمويل المشترك — هما اللذان يحرّكان الرقم صعوداً أو هبوطاً."],
].forEach(([b, t]) => kids.push(new Paragraph({ bidirectional: true, alignment: AlignmentType.RIGHT,
  spacing: { after: 90, line: 290 }, bullet: { level: 0 },
  children: [ R(b + " ", { bold: true }), R(t) ] })));

kids.push(P([ R("وثيقة داخلية للمناقشة التنفيذية — تُضبط الأرقام النهائية عند التعاقد والتوظيف الفعلي.", { italics: true, size: 18, color: GREY }) ],
  { before: 240, align: AlignmentType.CENTER }));

const doc = new Document({
  creator: "Darb", title: "ميزانيات ملف الشراكات - تانكي",
  styles: { default: { document: { run: { font: FONT, size: 22, color: INK } } },
    paragraphStyles: [ { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
      run: { font: FONT, bold: true, size: 26, color: INK } } ] },
  sections: [ { properties: { page: { size: { width: 12240, height: 15840 },
    margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } }, children: kids } ],
});
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(process.argv[2] || "budget.docx", buf); console.log("written", buf.length); });
