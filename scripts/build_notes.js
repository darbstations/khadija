const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, PageBreak,
} = require("docx");
const fs = require("fs");

const FONT = "Arial";
const ORANGE = "E07C16", ORANGE_L = "FEF3E6", SOFT = "FBF3E9";
const GREY = "6D6E70", INK = "26262A", LINE = "D9D7D3", HEAD_BG = "F3EFEA";
const GREEN = "1E7E45", GREEN_L = "E4F2E7";
const RED = "B23A34", RED_L = "FBE9E8", WARN = "9A6A12", WARN_L = "FBF1DD";

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
function H2(text, color) {
  return new Paragraph({ bidirectional: true, alignment: AlignmentType.RIGHT, spacing: { before: 160, after: 80 },
    children: [ R(text, { bold: true, size: 22, color: color || ORANGE }) ] });
}
function bullet(lead, text, o = {}) {
  const runs = [];
  if (lead) runs.push(R(lead + " ", { bold: true, color: o.leadColor || INK }));
  runs.push(R(text));
  return new Paragraph({ bidirectional: true, alignment: AlignmentType.RIGHT, spacing: { after: 70, line: 285 }, bullet: { level: 0 }, children: runs });
}
function box(label, text, fill, bar) {
  return P([ R(label + "  ", { bold: true, color: bar }), R(text) ],
    { shading: { type: ShadingType.CLEAR, color: "auto", fill }, border: { left: { style: BorderStyle.SINGLE, size: 18, color: bar, space: 8 } }, after: 130, line: 288 });
}

const kids = [];

// ---- Header ----
kids.push(new Paragraph({ spacing: { after: 0 }, alignment: AlignmentType.RIGHT, bidirectional: true, children: [ new TextRun({ text: "درب لخدمات المحطات", font: FONT, rightToLeft: true, bold: true, size: 22, color: GREY }) ] }));
kids.push(new Paragraph({ spacing: { before: 60, after: 0 }, alignment: AlignmentType.RIGHT, bidirectional: true, children: [ new TextRun({ text: "مراجعة تصميم تطبيق درب (مع برنامج تانكي) — التقرير الكامل", font: FONT, rightToLeft: true, bold: true, size: 33, color: INK }) ] }));
kids.push(new Paragraph({ spacing: { before: 40, after: 0 }, alignment: AlignmentType.RIGHT, bidirectional: true, children: [ new TextRun({ text: "مراجعة تجربة المستخدم وربط المكافآت · Darb Platform V1 (Figma)", font: FONT, rightToLeft: true, size: 22, color: ORANGE, bold: true }) ] }));
kids.push(new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { before: 120, after: 150 }, border: { bottom: { style: BorderStyle.SINGLE, size: 18, color: ORANGE, space: 8 } }, children: [] }));

// ---- CRITICAL ISSUES ----
kids.push(H1("★", "الملاحظات الحرجة (تحتاج قراراً قبل الإطلاق)"));
kids.push(box("🔴 ١. السحب (Raffle) = مراهنة/ميسر:",
  "شاشات السحب تتيح الاشتراك بنقاط/محاولات مدفوعة للفوز بجائزة (سيارة). دفع قيمة مقابل فرصة فوز = ميسر = محرّم شرعاً (سبق الاتفاق على ذلك). الحل: سحب مجاني بلا مقابل، أو دخول تلقائي مرتبط بالسلوك (كل تعبئة = دخول مجاني) لا بالدفع. راجعوا «Share to Win» بالمنطق نفسه.", RED_L, RED));
kids.push(box("🔴 ٢. معدّل الكسب لا يطابق النموذج الاقتصادي:",
  "شاشة التعبئة تعرض «+٢٠ نقطة» على ~١١٥ ﷼، بينما النموذج المعتمد = ١ نقطة لكل ريال (~١١٥ نقطة). فرق كبير يغيّر كامل الاقتصاد. يجب توحيد معدّل الكسب بين التصميم والنموذج (النقطة = نصف هللة).", RED_L, RED));
kids.push(box("🔴 ٣. تعارض سياسة انتهاء النقاط:",
  "شاشات تعرض «صالح لمدة ٢٤ شهراً» (انتهاء النقاط)، بينما تصميم المستويات يَعِد «النقاط لا تنتهي». يجب قرار موحّد: انتهاء (يقلّل الالتزام المالي ويحسّن الـ Breakage) أو عدم انتهاء (الوعد المعلَن) — ورسالة واحدة للعميل.", RED_L, RED));

// ---- PART 1 ----
kids.push(H1("١.", "الدخول والتسجيل (KYC) والمحفظة والسيارة"));
kids.push(H2("نقاط القوة", GREEN));
kids.push(bullet("المحفظة مدمجة:", "شحن الرصيد يخدم تجميع الـ Float داخل درب.", { leadColor: GREEN }));
kids.push(bullet("نوع الوقود عند إضافة السيارة:", "أساس للكسب المتغيّر حسب نوع الوقود.", { leadColor: GREEN }));
kids.push(bullet("سقف المحفظة (٩,٠٠٠ ﷼):", "وعي بحدود القيمة المخزّنة (ساما).", { leadColor: GREEN }));
kids.push(H2("ملاحظات"));
kids.push(box("🟠 طلب الشحن أثناء التسجيل:", "يرفع هجر التسجيل — يُنصح بجعله اختيارياً/مؤجّلاً وطلبه عند الحاجة. (طُبّق لاحقاً في التصميم عبر Wallet-Popup السياقي — جيد.)", SOFT, WARN));
kids.push(box("🟠 هوية المكافآت في التسجيل:", "يُنصح بإبراز تانكي وهدية الترحيب (١٠٠٠ نقطة) كمحفّز للتسجيل.", SOFT, WARN));
kids.push(box("🟠 اكتمال KYC للمحفظة:", "المحفظة النقدية تتطلب تحقق هوية نظامي (نفاذ) وفق ساما، لا الاكتفاء بالاسم.", SOFT, WARN));

// ---- TANKI LAYER ----
kids.push(H1("٢.", "تانكي داخل درب — طبقة لا شاشة"));
kids.push(P("درب هو التطبيق الأم؛ وتانكي يعمل كطبقة مكافآت فوق كل فعل. لا يُبنى تسجيل منفصل، بل تُرشّ إشارة المكافأة على تدفق درب القائم: شريط ترحيبي عند إدخال الجوال، سطر هدية على شاشة النجاح، شريحة تعريف واحدة، إشعار نقاط بعد كل فعل، وودجت رصيد ومستوى ثابت في الرئيسية. القاعدة: درب يعطي الخدمة · تانكي يكافئ عليها."));

// ---- PART 2 ----
kids.push(H1("٣.", "الرئيسية والخدمات والخريطة والإشعارات"));
kids.push(H2("نقاط القوة", GREEN));
kids.push(bullet("المستوى ظاهر في الرئيسية،", "والإشعارات مدفوعة بتانكي ومبوّبة (العروض/الرصيد/الباقات/النقاط)، وزر QR مركزي، وخدمات شاملة.", { leadColor: GREEN }));
kids.push(H2("ملاحظات"));
kids.push(box("🔴 أين رصيد النقاط؟", "المستوى ظاهر، لكن رقم النقاط (الذي يراقبه العميل) يجب أن يكون بطل الشاشة.", WARN_L, RED));
kids.push(box("🟠 ازدحام الرئيسية:", "يلزم مدخل واضح «نقاطي/تانكي» وسط الوحدات الكثيرة.", SOFT, WARN));
kids.push(box("🟠 وظيفة زر QR:", "توضيح دور الزر (كسب/دفع) — أهم نقطة تماس.", SOFT, WARN));
kids.push(box("🟡 الخريطة كمحرّك ولاء:", "تمييز المحطات ذات العروض/مضاعفة النقاط بدل محدّد محطات فقط.", WARN_L, WARN));
kids.push(box("🟡 إرهاق الإشعارات:", "التبويب جيد؛ يلزم حوكمة التكرار وفصل المعاملاتي عن التسويقي.", WARN_L, WARN));

// ---- PART 3 ----
kids.push(new Paragraph({ children: [new PageBreak()] }));
kids.push(H1("٤.", "العمليات التشغيلية — التعبئة الذاتية والاستبدال واستلام الهدية"));
kids.push(H2("نقاط القوة", GREEN));
kids.push(bullet("الكسب ظاهر لحظة الفعل", "(+نقاط أثناء التعبئة)، وحجز مؤقت للتعبئة الذاتية، و QR بعدّاد ٣٠ ثانية، وتحقق حضور «هل أنت أمام الكاشير؟» يمنع الاستبدال الوهمي، وربط العملية بالسيارة.", { leadColor: GREEN }));
kids.push(H2("ملاحظات"));
kids.push(box("🟢 عولجت سابقاً:", "«انتهت الصلاحية + إنشاء QR جديد» يعالج قصر مدة الكود — جيد.", GREEN_L, GREEN));
kids.push(box("🟡 الحجز المؤقت:", "توضيح متى يُرجّع الفرق بعد التعبئة الفعلية لتفادي التباس الخصم.", WARN_L, WARN));
kids.push(box("🟡 كشف نقاط موحّد:", "سجل الفواتير يعرض لكل عملية؛ يلزم صفحة «نقاطي» تجمع المكتسب/المستبدل/الرصيد/القريب من الانتهاء. (متوفّر في قسم Points Management — جيد.)", WARN_L, WARN));

// ---- PART 4 ----
kids.push(H1("٥.", "توسّعات الولاء — النقاط والإحالة والكوبونات والباقات والسحب"));
kids.push(H2("نقاط القوة", GREEN));
kids.push(bullet("كشف حساب النقاط", "(Points Management / Loyalty Points) مع تحويل نقاط وتاريخ، والإحالة (Referral) كمحرّك نمو، والكوبونات باستبدال QR، وإدارة مركبات شاملة (ميزانية لكل سيارة).", { leadColor: GREEN }));
kids.push(H2("ملاحظات"));
kids.push(box("🔴 السحب (Raffle):", "أهم ملاحظة — مراهنة إن كان الاشتراك بمقابل مدفوع (راجع القسم الحرج ★١). يجب تحويله لسحب مجاني بلا مقابل.", RED_L, RED));
kids.push(box("🟠 «Share to Win»:", "لو المكافأة إحالة مضمونة = سليم؛ لو فوز بالحظ = ميسر. يجب التأكد.", SOFT, WARN));
kids.push(box("🟡 تحويل النقاط:", "يبقى ضمن عالم درب المغلق (لا سحب نقدي) ليظل حلالاً وأخفّ تنظيماً.", WARN_L, WARN));

// ---- PART 5 ----
kids.push(H1("٦.", "الحساب والمحفظة والملف الشخصي والدعم"));
kids.push(bullet("شاشات مساندة", "(الملف الشخصي، المحفظة، الدعم/التذاكر، تحديث المعلومات، حالات نقص/انعدام الرصيد) — منفّذة جيداً.", { leadColor: GREEN }));
kids.push(bullet("إيجابي للامتثال:", "وجود «الخصوصية والشروط» و«حذف الحساب» يخدم نظام حماية البيانات (PDPL) وحقوق العميل.", { leadColor: GREEN }));

// ---- CONCLUSION ----
kids.push(H1("٧.", "الخلاصة"));
kids.push(bullet("", "التصميم قوي في الهوية والتنفيذ، وتانكي حاضر جيداً في المستوى والإشعارات والعمليات."));
kids.push(bullet("", "قبل الإطلاق يجب حسم ثلاث نقاط حرجة: السحب (المراهنة) · معدّل الكسب · انتهاء النقاط."));
kids.push(bullet("", "المبدأ الحاكم: تانكي طبقة تكافئ كل فعل داخل درب — والاقتصاد والشرعية يُوحَّدان مع النموذج المعتمد."));
kids.push(P([ R("وثيقة داخلية لمراجعة التصميم — مبنية على لقطات Darb Platform V1.", { italics: true, size: 18, color: GREY }) ], { before: 220, align: AlignmentType.CENTER }));

const doc = new Document({
  creator: "Darb", title: "مراجعة تصميم تطبيق درب - تانكي",
  styles: { default: { document: { run: { font: FONT, size: 22, color: INK } } },
    paragraphStyles: [ { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: FONT, bold: true, size: 26, color: INK } } ] },
  sections: [ { properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } }, children: kids } ],
});
Packer.toBuffer(doc).then((b) => { fs.writeFileSync(process.argv[2] || "notes.docx", b); console.log("written", b.length); });
