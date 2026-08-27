const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, PageBreak,
} = require("docx");
const fs = require("fs");

const FONT = "Arial";
const ORANGE = "E07C16", ORANGE_L = "FEF3E6", SOFT = "FBF3E9";
const GREY = "6D6E70", INK = "26262A", LINE = "D9D7D3", HEAD_BG = "F3EFEA";
const WARN = "9A6A12", WARN_L = "FBF1DD";

function R(t, o = {}) { return new TextRun({ text: t, font: FONT, rightToLeft: true, size: o.size || 22, bold: o.bold, italics: o.italics, color: o.color || INK }); }
function P(runs, o = {}) {
  if (typeof runs === "string") runs = [R(runs, o)];
  return new Paragraph({ bidirectional: true, alignment: o.align || AlignmentType.RIGHT,
    spacing: { after: o.after != null ? o.after : 110, before: o.before || 0, line: o.line || 300 },
    ...(o.shading ? { shading: o.shading } : {}), ...(o.border ? { border: o.border } : {}), children: runs });
}
function H1(num, text) {
  return new Paragraph({ bidirectional: true, alignment: AlignmentType.RIGHT, heading: HeadingLevel.HEADING_1,
    spacing: { before: 280, after: 120 }, border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: ORANGE, space: 4 } },
    children: [ R(num + "  ", { bold: true, size: 26, color: ORANGE }), R(text, { bold: true, size: 26, color: INK }) ] });
}
function cellP(text, o = {}) {
  return new TableCell({ width: { size: o.w, type: WidthType.DXA },
    shading: o.shade ? { type: ShadingType.CLEAR, color: "auto", fill: o.shade } : undefined,
    margins: { top: 64, bottom: 64, left: 90, right: 90 }, verticalAlign: "center",
    children: [P([R(text, { size: o.size || 18, bold: o.bold, color: o.color || INK })], { after: 0, align: o.align || AlignmentType.RIGHT, line: 255 })] });
}
function headRow(labels, widths, fill) {
  return new TableRow({ tableHeader: true, children: labels.map((l, i) =>
    new TableCell({ width: { size: widths[i], type: WidthType.DXA }, shading: { type: ShadingType.CLEAR, color: "auto", fill: fill || HEAD_BG }, margins: { top: 64, bottom: 64, left: 90, right: 90 }, verticalAlign: "center",
      children: [P([R(l, { bold: true, size: 18 })], { after: 0, align: AlignmentType.CENTER })] })) });
}
const C = (t, o = {}) => ({ t, ...o });
function table(widths, header, rows) {
  const trs = [ headRow(header, widths) ];
  for (const r of rows) trs.push(new TableRow({ children: r.map((c, i) => cellP(c.t, { w: widths[i], align: c.align, bold: c.bold, color: c.color, shade: c.shade, size: c.size })) }));
  return new Table({ width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths, visuallyRightToLeft: true,
    borders: { top: { style: BorderStyle.SINGLE, size: 4, color: LINE }, bottom: { style: BorderStyle.SINGLE, size: 4, color: LINE }, left: { style: BorderStyle.SINGLE, size: 4, color: LINE }, right: { style: BorderStyle.SINGLE, size: 4, color: LINE }, insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: LINE }, insideVertical: { style: BorderStyle.SINGLE, size: 4, color: LINE } }, rows: trs });
}
function qbullet(text) {
  return new Paragraph({ bidirectional: true, alignment: AlignmentType.RIGHT, spacing: { after: 70, line: 285 }, bullet: { level: 0 }, children: [R(text)] });
}

const kids = [];

// Header
kids.push(new Paragraph({ spacing: { after: 0 }, alignment: AlignmentType.RIGHT, bidirectional: true, children: [ new TextRun({ text: "درب لخدمات المحطات", font: FONT, rightToLeft: true, bold: true, size: 22, color: GREY }) ] }));
kids.push(new Paragraph({ spacing: { before: 60, after: 0 }, alignment: AlignmentType.RIGHT, bidirectional: true, children: [ new TextRun({ text: "مقارنة مزوّدي بوابات الدفع وقائمة الأسئلة — محفظة تانكي", font: FONT, rightToLeft: true, bold: true, size: 32, color: INK }) ] }));
kids.push(new Paragraph({ spacing: { before: 40, after: 0 }, alignment: AlignmentType.RIGHT, bidirectional: true, children: [ new TextRun({ text: "أداة تقييم موحّدة لإدارة المالية · تُعبّأ لكل مزوّد", font: FONT, rightToLeft: true, size: 22, color: ORANGE, bold: true }) ] }));
kids.push(new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { before: 120, after: 150 }, border: { bottom: { style: BorderStyle.SINGLE, size: 18, color: ORANGE, space: 8 } }, children: [] }));
kids.push(P("تُستخدم هذه الوثيقة لمقارنة عروض مزوّدي بوابات الدفع بنفس المعايير. تُعبّأ الأعمدة من عروض كل مزوّد (الراجحي / HyperPay / جيديا …)، وتُرفق قائمة الأسئلة عند طلب العروض.", { after: 180 }));

// ===== TABLE 1: الرسوم =====
kids.push(H1("١.", "جدول مقارنة الرسوم"));
const W = [3060, 2100, 2100, 2100];
kids.push(table(W,
  ["البند", "الراجحي (مرجع)", "HyperPay", "مزوّد ٣"],
  [
    [C("رسوم التأسيس (لمرة)", { bold: true, shade: HEAD_BG }), C("٣,٠٠٠ ﷼", { align: AlignmentType.CENTER }), C("……", { align: AlignmentType.CENTER, shade: ORANGE_L }), C("……", { align: AlignmentType.CENTER })],
    [C("الاشتراك الشهري", { bold: true, shade: HEAD_BG }), C("٣٠٠ ﷼", { align: AlignmentType.CENTER }), C("……", { align: AlignmentType.CENTER, shade: ORANGE_L }), C("……", { align: AlignmentType.CENTER })],
    [C("عمولة مدى (MDR)", { bold: true, shade: HEAD_BG }), C("١٪ (بحد ٢٠٠ ﷼)", { align: AlignmentType.CENTER }), C("……", { align: AlignmentType.CENTER, shade: ORANGE_L }), C("……", { align: AlignmentType.CENTER })],
    [C("عمولة البطاقات الائتمانية", { bold: true, shade: HEAD_BG }), C("٢٫٥٠٪ + ١ ﷼", { align: AlignmentType.CENTER }), C("……", { align: AlignmentType.CENTER, shade: ORANGE_L }), C("……", { align: AlignmentType.CENTER })],
    [C("رسم شحن المحفظة عبر مدى", { bold: true, shade: WARN_L, color: WARN }), C("١٫٥ ﷼ ثابت", { align: AlignmentType.CENTER, shade: WARN_L, bold: true }), C("……", { align: AlignmentType.CENTER, shade: WARN_L }), C("……", { align: AlignmentType.CENTER, shade: WARN_L })],
    [C("رسوم الاسترجاع (Refund)", { bold: true, shade: HEAD_BG }), C("١ ﷼", { align: AlignmentType.CENTER }), C("……", { align: AlignmentType.CENTER, shade: ORANGE_L }), C("……", { align: AlignmentType.CENTER })],
    [C("رسوم الاعتراض (Chargeback)", { bold: true, shade: HEAD_BG }), C("على التاجر", { align: AlignmentType.CENTER }), C("……", { align: AlignmentType.CENTER, shade: ORANGE_L }), C("……", { align: AlignmentType.CENTER })],
    [C("ضريبة القيمة المضافة", { bold: true, shade: HEAD_BG }), C("تُضاف", { align: AlignmentType.CENTER }), C("……", { align: AlignmentType.CENTER, shade: ORANGE_L }), C("……", { align: AlignmentType.CENTER })],
  ]));

// ===== TABLE 2: القدرات =====
kids.push(H1("٢.", "جدول مقارنة القدرات والشروط"));
kids.push(table(W,
  ["المعيار", "الراجحي (مرجع)", "HyperPay", "مزوّد ٣"],
  [
    [C("منتج محفظة/قيمة مخزّنة", { bold: true, shade: HEAD_BG }), C("لا (قبول فقط)", { align: AlignmentType.CENTER }), C("……", { align: AlignmentType.CENTER, shade: ORANGE_L }), C("……", { align: AlignmentType.CENTER })],
    [C("Apple Pay / STC Pay", { bold: true, shade: HEAD_BG }), C("……", { align: AlignmentType.CENTER }), C("……", { align: AlignmentType.CENTER, shade: ORANGE_L }), C("……", { align: AlignmentType.CENTER })],
    [C("التقسيط (تابي/تمارا)", { bold: true, shade: HEAD_BG }), C("……", { align: AlignmentType.CENTER }), C("……", { align: AlignmentType.CENTER, shade: ORANGE_L }), C("……", { align: AlignmentType.CENTER })],
    [C("صرف للتجار (Payout)", { bold: true, shade: WARN_L, color: WARN }), C("……", { align: AlignmentType.CENTER, shade: WARN_L }), C("……", { align: AlignmentType.CENTER, shade: WARN_L }), C("……", { align: AlignmentType.CENTER, shade: WARN_L })],
    [C("دعم KYC / نفاذ", { bold: true, shade: HEAD_BG }), C("……", { align: AlignmentType.CENTER }), C("……", { align: AlignmentType.CENTER, shade: ORANGE_L }), C("……", { align: AlignmentType.CENTER })],
    [C("تقليل نطاق PCI (Hosted)", { bold: true, shade: HEAD_BG }), C("……", { align: AlignmentType.CENTER }), C("……", { align: AlignmentType.CENTER, shade: ORANGE_L }), C("……", { align: AlignmentType.CENTER })],
    [C("مدة العقد / الخروج", { bold: true, shade: HEAD_BG }), C("سنتان · تلقائي", { align: AlignmentType.CENTER }), C("……", { align: AlignmentType.CENTER, shade: ORANGE_L }), C("……", { align: AlignmentType.CENTER })],
    [C("مدة التسوية للحساب", { bold: true, shade: HEAD_BG }), C("……", { align: AlignmentType.CENTER }), C("……", { align: AlignmentType.CENTER, shade: ORANGE_L }), C("……", { align: AlignmentType.CENTER })],
    [C("زمن التكامل التقني", { bold: true, shade: HEAD_BG }), C("……", { align: AlignmentType.CENTER }), C("……", { align: AlignmentType.CENTER, shade: ORANGE_L }), C("……", { align: AlignmentType.CENTER })],
  ]));

// ===== QUESTIONS =====
kids.push(new Paragraph({ children: [new PageBreak()] }));
kids.push(H1("٣.", "قائمة الأسئلة (تُرسل لكل مزوّد)"));

kids.push(P([R("أ) الرسوم والتكاليف", { bold: true, size: 23, color: ORANGE })], { after: 80 }));
["ما رسوم التأسيس (لمرة واحدة) والاشتراك الشهري/السنوي؟",
 "ما نسبة العمولة (MDR) لكل عملية — لمدى وللبطاقات الائتمانية منفصلين؟",
 "هل لديكم سعر خاص لشحن المحفظة الإلكترونية عبر مدى؟ (مرجع: الراجحي ١٫٥ ﷼ ثابت للعملية.)",
 "ما رسوم الاسترجاع (Refund) والاعتراض (Chargeback) ومَن يتحمّلها؟",
 "هل الرسوم شاملة أم تُضاف عليها ضريبة القيمة المضافة؟",
 "هل تنخفض النسب مع زيادة حجم المعاملات؟ وما الشرائح؟"].forEach(q => kids.push(qbullet(q)));

kids.push(P([R("ب) القدرات التقنية", { bold: true, size: 23, color: ORANGE })], { before: 120, after: 80 }));
["هل تقدّمون منتج محفظة/قيمة مخزّنة، أم قبول الدفع فقط؟",
 "هل تدعمون Apple Pay و STC Pay والتقسيط (تابي/تمارا)؟",
 "هل تدعمون الصرف للتجار (Payout/Disbursement) لتسوية شركاء الولاء؟",
 "هل تدعمون التحقق من الهوية (نفاذ/Nafath) أو نحتاجه من مزوّد خارجي؟",
 "هل يقلّل حلّكم من نطاق امتثال PCI-DSS لدينا (Hosted fields/Tokenization)؟",
 "كم مدة التكامل التقني المتوقعة؟ وما مستوى الدعم أثناءه؟"].forEach(q => kids.push(qbullet(q)));

kids.push(P([R("ج) الشروط التعاقدية والتشغيلية", { bold: true, size: 23, color: ORANGE })], { before: 120, after: 80 }));
["ما مدة العقد وآلية التجديد وشروط الخروج؟",
 "هل يمكنكم تعديل الأسعار خلال مدة العقد؟ وبأي إشعار؟",
 "ما مدة تسوية الأموال إلى حسابنا البنكي؟ وهل هناك رسم تسوية؟",
 "ما اتفاقية مستوى الخدمة (SLA) للتوفّر والدعم؟",
 "كيف تُدار حالات الاحتيال والاعتراضات إجرائياً؟",
 "هل هناك حدّ أدنى شهري للمعاملات أو رسوم عدم استخدام؟"].forEach(q => kids.push(qbullet(q)));

// Note box
kids.push(P([ R("تنبيه مهم:  ", { bold: true, color: WARN }),
  R("بوابة الدفع تغطّي «قبول الشحن» فقط. الاحتفاظ برصيد المحفظة نشاط منظّم من ساما يتطلب ترخيصاً خاصاً أو شراكة مع محفظة مرخّصة (STC Pay / urpay / بنك) — يُقيَّم بالتوازي مع البوابة.") ],
  { before: 200, shading: { type: ShadingType.CLEAR, color: "auto", fill: WARN_L }, border: { left: { style: BorderStyle.SINGLE, size: 18, color: WARN, space: 8 } }, line: 290 }));

kids.push(P([ R("وثيقة داخلية — أداة تقييم لإدارة المالية. الأرقام المرجعية للراجحي من عرضهم القياسي وقابلة للتفاوض.", { italics: true, size: 18, color: GREY }) ], { before: 200, align: AlignmentType.CENTER }));

const doc = new Document({
  creator: "Darb", title: "مقارنة مزوّدي بوابات الدفع - تانكي",
  numbering: { config: [ { reference: "default-bullet", levels: [ { level: 0, format: "bullet", text: "•", alignment: AlignmentType.RIGHT, style: { run: { color: ORANGE, font: FONT }, paragraph: { indent: { right: 360, hanging: 220 } } } } ] } ] },
  styles: { default: { document: { run: { font: FONT, size: 22, color: INK } } },
    paragraphStyles: [ { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: FONT, bold: true, size: 26, color: INK } } ] },
  sections: [ { properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } }, children: kids } ],
});
Packer.toBuffer(doc).then((b) => { fs.writeFileSync(process.argv[2] || "psp.docx", b); console.log("written", b.length); });
