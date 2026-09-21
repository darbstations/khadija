const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, BorderStyle, ShadingType,
} = require("docx");
const fs = require("fs");

const FONT = "Arial";
const ORANGE = "E07C16", ORANGE_L = "FEF3E6", SOFT = "FBF3E9";
const GREY = "6D6E70", INK = "26262A", LINE = "D9D7D3", RED = "B23A34", RED_L = "FBE9E8";

function R(t, o = {}) { return new TextRun({ text: t, font: FONT, rightToLeft: true, size: o.size || 22, bold: o.bold, italics: o.italics, color: o.color || INK }); }
function P(runs, o = {}) {
  if (typeof runs === "string") runs = [R(runs, o)];
  return new Paragraph({ bidirectional: true, alignment: o.align || AlignmentType.RIGHT,
    spacing: { after: o.after != null ? o.after : 90, before: o.before || 0, line: o.line || 300 },
    ...(o.shading ? { shading: o.shading } : {}), ...(o.border ? { border: o.border } : {}), children: runs });
}
function section(num, title, critical) {
  return new Paragraph({ bidirectional: true, alignment: AlignmentType.RIGHT, heading: HeadingLevel.HEADING_1,
    spacing: { before: 260, after: 120 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: critical ? RED : ORANGE, space: 4 } },
    shading: critical ? { type: ShadingType.CLEAR, color: "auto", fill: RED_L } : undefined,
    children: [ R("القسم " + num + " — ", { bold: true, size: 25, color: critical ? RED : ORANGE }),
      R(title, { bold: true, size: 25, color: INK }),
      ...(critical ? [R("   🔴 استراتيجي", { bold: true, size: 18, color: RED })] : []) ] });
}
// question with options
let qn = 0;
function Q(text, opts, isOpen) {
  qn++;
  const out = [ new Paragraph({ bidirectional: true, alignment: AlignmentType.RIGHT, spacing: { before: 120, after: 50, line: 290 },
    children: [ R("س" + qn + ". ", { bold: true, color: ORANGE }), R(text, { bold: true }) ] }) ];
  if (isOpen) {
    out.push(P([R("……………………………………………………………………………………………", { color: GREY })], { after: 60 }));
  } else if (opts) {
    for (const o of opts) out.push(new Paragraph({ bidirectional: true, alignment: AlignmentType.RIGHT,
      spacing: { after: 30, line: 270 }, indent: { right: 260 },
      children: [ R("☐  ", { color: ORANGE, size: 22 }), R(o) ] }));
  }
  return out;
}

const kids = [];

// Header
kids.push(new Paragraph({ spacing: { after: 0 }, alignment: AlignmentType.RIGHT, bidirectional: true, children: [ new TextRun({ text: "درب لخدمات المحطات", font: FONT, rightToLeft: true, bold: true, size: 22, color: GREY }) ] }));
kids.push(new Paragraph({ spacing: { before: 60, after: 0 }, alignment: AlignmentType.RIGHT, bidirectional: true, children: [ new TextRun({ text: "استبيان تجربة العميل — قبل إطلاق تطبيق درب (تانكي)", font: FONT, rightToLeft: true, bold: true, size: 32, color: INK }) ] }));
kids.push(new Paragraph({ spacing: { before: 40, after: 0 }, alignment: AlignmentType.RIGHT, bidirectional: true, children: [ new TextRun({ text: "نموذج جاهز · ٨ أقسام · يُعبّأ من العميل", font: FONT, rightToLeft: true, size: 22, color: ORANGE, bold: true }) ] }));
kids.push(new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { before: 120, after: 140 }, border: { bottom: { style: BorderStyle.SINGLE, size: 18, color: ORANGE, space: 8 } }, children: [] }));
kids.push(P("شكراً لمشاركتك — إجاباتك تساعدنا نبني تجربة أفضل لك. الاستبيان يأخذ ٥–٧ دقائق، وكل الإجابات سرّية. (أكمل الاستبيان وادخل السحب المجاني 🎁)", { after: 160, shading: { type: ShadingType.CLEAR, color: "auto", fill: SOFT } }));

// Section 1
kids.push(section("١", "السلوك الحالي"));
Q("كم مرة تعبّي بنزين في الأسبوع؟", ["مرة أو أقل", "٢–٣ مرات", "٤ مرات فأكثر"]).forEach(k => kids.push(k));
Q("متوسط مبلغ التعبئة الواحدة؟", ["أقل من ١٠٠ ﷼", "١٠٠–٢٠٠ ﷼", "٢٠٠–٤٠٠ ﷼", "أكثر من ٤٠٠ ﷼"]).forEach(k => kids.push(k));
Q("أي محطات تفضّل حالياً؟ ولماذا؟", null, true).forEach(k => kids.push(k));
Q("كيف تدفع الآن غالباً؟", ["نقد", "مدى", "Apple Pay", "بطاقة ائتمانية"]).forEach(k => kids.push(k));
Q("كم سيارة لديك؟", ["واحدة", "اثنتان", "٣ فأكثر", "أسطول/شركة"]).forEach(k => kids.push(k));

// Section 2
kids.push(section("٢", "برامج الولاء الحالية"));
Q("هل تستخدم أي برنامج ولاء لمحطات حالياً؟", ["نعم — ساسكو", "نعم — أدنوك", "نعم — غيره", "لا أستخدم أي برنامج"]).forEach(k => kids.push(k));
Q("ما الذي يعجبك أو لا يعجبك في برنامج الولاء الذي تستخدمه؟", null, true).forEach(k => kids.push(k));
Q("هل تستبدل نقاطك فعلاً؟", ["نعم دائماً", "أحياناً", "نادراً", "تتراكم بلا استخدام"]).forEach(k => kids.push(k));

// Section 3 (critical)
kids.push(section("٣", "المحفظة والدفع", true));
Q("هل أنت مستعد لشحن رصيد مقدّم في محفظة التطبيق؟", ["نعم", "ربما", "لا"]).forEach(k => kids.push(k));
Q("كم مبلغ مستعد تشحن في المرة الواحدة؟", ["٥٠ ﷼", "١٠٠ ﷼", "٢٠٠ ﷼", "٥٠٠ ﷼ فأكثر"]).forEach(k => kids.push(k));
Q("ما الذي يمنعك من شحن المحفظة؟", ["لا يوجد مانع", "لا أثق بحفظ رصيدي", "أخاف ألا أستطيع استرجاعه", "لا أحتاجها", "أراها معقّدة"]).forEach(k => kids.push(k));
Q("لو شحنت ولم تستخدم الرصيد، هل تتوقّع أن تستطيع استرجاعه؟", ["نعم أتوقّع ذلك", "لا أدري", "لا أتوقّع"]).forEach(k => kids.push(k));
Q("تفضّل شحن المحفظة بـ:", ["مدى", "بطاقة ائتمانية", "Apple Pay", "لا يهم"]).forEach(k => kids.push(k));

// Section 4 (critical)
kids.push(section("٤", "قيمة المكافأة", true));
Q("كم نسبة المكافأة (كاش باك/نقاط) التي تجعلك تغيّر محطتك المعتادة؟", ["١٪", "٣٪", "٥٪", "أكثر من ٥٪"]).forEach(k => kids.push(k));
Q("أي مكافأة تحمّسك أكثر؟", ["خصم بنزين", "قهوة/سناك مجاني", "غسيل سيارة", "بطاقات هدايا (أمازون/جرير)", "تحويل لرصيد المحفظة"]).forEach(k => kids.push(k));
Q("تفضّل نقاطاً تتجمّع، أم خصماً فورياً؟", ["نقاط تتجمّع", "خصم فوري", "لا يهم"]).forEach(k => kids.push(k));

// Section 5
kids.push(section("٥", "الباقات والعروض"));
Q("هل يهمك شراء باقة مسبقة الدفع (بنزين + غسيل + قهوة بسعر أوفر)؟", ["نعم جداً", "ربما", "لا"]).forEach(k => kids.push(k));
Q("ما المبلغ المناسب لباقة كهذه؟", ["أقل من ٥٠٠ ﷼", "٥٠٠–١٠٠٠ ﷼", "أكثر من ١٠٠٠ ﷼"]).forEach(k => kids.push(k));
Q("هل تهمك عروض التجار داخل التطبيق؟ (مطاعم · كافيهات · مغاسل)", ["نعم", "قليلاً", "لا"]).forEach(k => kids.push(k));

// Section 6
kids.push(section("٦", "التعبئة الذاتية و QR"));
Q("هل أنت مستعد للتعبئة الذاتية عبر التطبيق (بدون موظف)؟", ["نعم", "ربما", "لا"]).forEach(k => kids.push(k));
Q("ما الذي يقلقك في الدفع عبر QR عند المضخة؟", null, true).forEach(k => kids.push(k));

// Section 7
kids.push(section("٧", "الثقة والمخاوف"));
Q("ما مدى ثقتك بأن شركة محطات تحفظ رصيدك؟ (١ ضعيفة – ٥ عالية)", ["١", "٢", "٣", "٤", "٥"]).forEach(k => kids.push(k));
Q("هل تقلقك خصوصية بياناتك في التطبيق؟", ["نعم كثيراً", "قليلاً", "لا"]).forEach(k => kids.push(k));

// Section 8
kids.push(section("٨", "بيانات عامة (اختياري)"));
Q("العمر:", ["أقل من ٢٥", "٢٥–٣٤", "٣٥–٤٤", "٤٥ فأكثر"]).forEach(k => kids.push(k));
Q("الجنس:", ["ذكر", "أنثى", "أفضّل عدم الإفصاح"]).forEach(k => kids.push(k));
Q("المنطقة:", null, true).forEach(k => kids.push(k));

// Closing open question
kids.push(section("★", "سؤال ختامي"));
Q("ما الشيء الوحيد الذي يجعلك تستخدم تطبيق درب كل يوم؟", null, true).forEach(k => kids.push(k));

kids.push(P([ R("شكراً لوقتك — رأيك يصنع الفرق. 🌟", { bold: true, color: ORANGE }) ], { before: 200, align: AlignmentType.CENTER }));
kids.push(P([ R("نموذج داخلي · درب لخدمات المحطات · الأسئلة المميّزة بالأحمر (٣ و٤) هي الأهم استراتيجياً.", { italics: true, size: 18, color: GREY }) ], { before: 120, align: AlignmentType.CENTER }));

const doc = new Document({
  creator: "Darb", title: "استبيان تجربة العميل - تانكي",
  styles: { default: { document: { run: { font: FONT, size: 22, color: INK } } },
    paragraphStyles: [ { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: FONT, bold: true, size: 25, color: INK } } ] },
  sections: [ { properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } }, children: kids } ],
});
Packer.toBuffer(doc).then((b) => { fs.writeFileSync(process.argv[2] || "survey.docx", b); console.log("written", b.length, "questions", qn); });
