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
const GREEN_L = "E4F2E7";
const RED = "B23A34";
const WARN = "9A6A12";
const WARN_L = "FBF1DD";

function R(text, o = {}) {
  return new TextRun({ text, font: FONT, rightToLeft: true, size: o.size || 22,
    bold: o.bold, italics: o.italics, color: o.color || INK });
}
function P(runs, o = {}) {
  if (typeof runs === "string") runs = [R(runs, o)];
  return new Paragraph({ bidirectional: true, alignment: o.align || AlignmentType.RIGHT,
    spacing: { after: o.after != null ? o.after : 110, before: o.before || 0, line: o.line || 300 },
    ...(o.shading ? { shading: o.shading } : {}), ...(o.border ? { border: o.border } : {}), children: runs });
}
function H1(num, text) {
  return new Paragraph({ bidirectional: true, alignment: AlignmentType.RIGHT, heading: HeadingLevel.HEADING_1,
    spacing: { before: 300, after: 130 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: ORANGE, space: 4 } },
    children: [ R(num + "  ", { bold: true, size: 26, color: ORANGE }), R(text, { bold: true, size: 26, color: INK }) ] });
}
function H2(text, color) {
  return new Paragraph({ bidirectional: true, alignment: AlignmentType.RIGHT,
    spacing: { before: 180, after: 90 },
    children: [ R(text, { bold: true, size: 23, color: color || ORANGE }) ] });
}
function bullet(lead, text, o = {}) {
  const runs = [];
  if (lead) runs.push(R(lead + " ", { bold: true, color: o.leadColor || INK }));
  runs.push(R(text));
  return new Paragraph({ bidirectional: true, alignment: AlignmentType.RIGHT,
    spacing: { after: 80, line: 290 }, bullet: { level: 0 }, children: runs });
}
function noteBox(label, text, fill, bar) {
  return P([ R(label + "  ", { bold: true, color: bar }), R(text) ],
    { shading: { type: ShadingType.CLEAR, color: "auto", fill },
      border: { left: { style: BorderStyle.SINGLE, size: 18, color: bar, space: 8 } }, after: 140, line: 290 });
}

const kids = [];

// ---- Cover header ----
kids.push(new Paragraph({ spacing: { after: 0 }, alignment: AlignmentType.RIGHT, bidirectional: true,
  children: [ new TextRun({ text: "درب لخدمات المحطات", font: FONT, rightToLeft: true, bold: true, size: 22, color: GREY }) ] }));
kids.push(new Paragraph({ spacing: { before: 60, after: 0 }, alignment: AlignmentType.RIGHT, bidirectional: true,
  children: [ new TextRun({ text: "مراجعة تصميم تطبيق درب (مع برنامج تانكي) — الملاحظات", font: FONT, rightToLeft: true, bold: true, size: 34, color: INK }) ] }));
kids.push(new Paragraph({ spacing: { before: 40, after: 0 }, alignment: AlignmentType.RIGHT, bidirectional: true,
  children: [ new TextRun({ text: "مراجعة تجربة المستخدم وربط برنامج المكافآت · Darb Platform V1 (Figma)", font: FONT, rightToLeft: true, size: 22, color: ORANGE, bold: true }) ] }));
kids.push(new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { before: 120, after: 160 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 18, color: ORANGE, space: 8 } }, children: [] }));
kids.push(P("تجمع هذه الوثيقة ملاحظات مراجعة تصميم تطبيق درب على جزأين: (١) الدخول والتسجيل والمحفظة والسيارة، (٢) الرئيسية والخدمات والخريطة والإشعارات — مع التركيز على تجربة المستخدم وربط برنامج المكافآت «تانكي».", { after: 200 }));

// ============ PART 1 ============
kids.push(H1("١.", "الجزء الأول — الدخول والتسجيل (KYC) والمحفظة والسيارة"));
kids.push(H2("نقاط القوة", GREEN));
kids.push(bullet("المحفظة مدمجة من التسجيل:", "شحن الرصيد قبل الرئيسية يخدم هدف تجميع الرصيد (Float) داخل درب.", { leadColor: GREEN }));
kids.push(bullet("التقاط نوع الوقود عند إضافة السيارة:", "أساس ممتاز للكسب المتغيّر حسب نوع الوقود (٩١/٩٥/٩٨/ديزل/كهرباء).", { leadColor: GREEN }));
kids.push(bullet("سقف المحفظة (٩,٠٠٠ ﷼):", "وعي بحدود القيمة المخزّنة — يخدم متطلبات تنظيم ساما.", { leadColor: GREEN }));
kids.push(bullet("الهوية والتنفيذ:", "برتقالي درب متسق، RTL محترف، وحالات خطأ واضحة (رسالة + لون).", { leadColor: GREEN }));

kids.push(H2("ملاحظات مهمة (مرتّبة بالأهم)"));
kids.push(noteBox("🔴 ١. طلب الشحن أثناء التسجيل = خطر تسرّب عالٍ:",
  "طلب شحن المحفظة قبل أن يرى العميل أي قيمة أكبر مصدر لهجر التسجيل. الأفضل جعل الشحن اختيارياً/مؤجّلاً («تخطّى الآن») وطلبه لحظة الحاجة (أول تعبئة أو دفع).", WARN_L, RED));
kids.push(noteBox("🟠 ٢. غياب هوية المكافآت من التسجيل:",
  "التدفق يبدو تطبيق محفظة/دفع لا مدفوعاً بالولاء. لا ذكر لتانكي ولا لهدية الترحيب (١٠٠٠ نقطة). أقوى محفّز للتسجيل هو المكافأة — يُنصح بإضافة لحظة «سجّل واستلم ١٠٠٠ نقطة ترحيبية».", SOFT, WARN));
kids.push(noteBox("🟠 ٣. اكتمال KYC للمحفظة:",
  "شاشة «خلّنا نتعرّف عليك» تكتفي بالاسم؛ بينما المحفظة النقدية تتطلب تحقق هوية نظامي (نفاذ/الهوية) وفق متطلبات ساما. يجب التأكد من وجود هذا التحقق.", SOFT, WARN));
kids.push(noteBox("🟡 ٤. وضوح الإلزامي مقابل الاختياري:",
  "غير واضح إن كانت إضافة السيارة والشحن إلزاميين؛ إن كانا كذلك يصبح التسجيل ثقيلاً. يجب تعليم ما يمكن تخطّيه بوضوح.", WARN_L, WARN));
kids.push(noteBox("🟡 ٥. ثلاث طبقات تحقق:",
  "OTP + رمز مرور + Face ID ضرورية لمحفظة، لكنها تضيف خطوات — يجب إبقاؤها متتابعة وسلسة دون تكرار.", WARN_L, WARN));

kids.push(H2("الترتيب الأمثل المقترح"));
kids.push(P([ R("جوال → OTP → معلومات بسيطة + هدية ترحيب → الرئيسية (يرى القيمة) → ثم طلب الشحن وإضافة السيارة سياقياً عند الحاجة. القاعدة: تأجيل طلب المال والبيانات حتى يرى العميل الفائدة — فيكتمل التسجيل أسرع ويرتفع التفعيل.", { }) ],
  { shading: { type: ShadingType.CLEAR, color: "auto", fill: ORANGE_L } }));

// ============ TANKI-INSIDE-DARB ============
kids.push(H1("٢.", "كيف يظهر تانكي داخل تطبيق درب (طبقة لا شاشة)"));
kids.push(P("درب هو التطبيق الأم (وقود، محفظة، سيارة، دفع)، وتانكي برنامج المكافآت الذي يعمل كطبقة فوق كل شيء. لا يُبنى تسجيل منفصل لتانكي، بل تُرشّ إشارة المكافأة على تدفق درب القائم بأقل تغيير:"));
const placements = [
  ["شاشة إدخال الجوال (الأولى)", "شريط علوي: «سجّل واحصل على ١٠٠٠ نقطة تانكي 🎁»", "صفر"],
  ["شاشة «مبروك، حسابك جاهز»", "سطر إضافي: «وأضفنا لك ١٠٠٠ نقطة تانكي ترحيبية»", "صفر"],
  ["شرائح الترحيب", "شريحة واحدة: «مع تانكي، كل تعبئة وكل عملية تكسبك نقاط»", "شريحة واحدة"],
  ["بعد إضافة السيارة / الشحن", "إشعار صغير (Toast): «+٥٠ نقطة تانكي»", "صفر"],
  ["الصفحة الرئيسية", "ودجت ثابت لرصيد النقاط والمستوى", "عنصر بالرئيسية"],
];
const PW = [2500, 4560, 1300];
const prows = [ new TableRow({ tableHeader: true, children: [
  new TableCell({ width: { size: PW[0], type: WidthType.DXA }, shading: { type: ShadingType.CLEAR, color: "auto", fill: HEAD_BG }, margins: {top:70,bottom:70,left:90,right:90}, children: [P([R("الموضع (شاشة موجودة)", {bold:true,size:20})],{after:0,align:AlignmentType.CENTER})] }),
  new TableCell({ width: { size: PW[1], type: WidthType.DXA }, shading: { type: ShadingType.CLEAR, color: "auto", fill: HEAD_BG }, margins: {top:70,bottom:70,left:90,right:90}, children: [P([R("كيف يُطبّق", {bold:true,size:20})],{after:0,align:AlignmentType.CENTER})] }),
  new TableCell({ width: { size: PW[2], type: WidthType.DXA }, shading: { type: ShadingType.CLEAR, color: "auto", fill: HEAD_BG }, margins: {top:70,bottom:70,left:90,right:90}, children: [P([R("خطوات جديدة", {bold:true,size:20})],{after:0,align:AlignmentType.CENTER})] }),
]}) ];
for (const r of placements) prows.push(new TableRow({ children: [
  new TableCell({ width: { size: PW[0], type: WidthType.DXA }, margins: {top:70,bottom:70,left:90,right:90}, children: [P([R(r[0],{bold:true,size:19})],{after:0,line:260})] }),
  new TableCell({ width: { size: PW[1], type: WidthType.DXA }, margins: {top:70,bottom:70,left:90,right:90}, children: [P([R(r[1],{size:19})],{after:0,line:260})] }),
  new TableCell({ width: { size: PW[2], type: WidthType.DXA }, shading: { type: ShadingType.CLEAR, color: "auto", fill: r[2]==="صفر"?GREEN_L:SOFT }, margins: {top:70,bottom:70,left:90,right:90}, children: [P([R(r[2],{size:19,bold:true,color:r[2]==="صفر"?GREEN:WARN})],{after:0,align:AlignmentType.CENTER})] }),
]}));
kids.push(new Table({ width: { size: PW[0]+PW[1]+PW[2], type: WidthType.DXA }, columnWidths: PW, visuallyRightToLeft: true,
  borders: { top:{style:BorderStyle.SINGLE,size:4,color:LINE}, bottom:{style:BorderStyle.SINGLE,size:4,color:LINE}, left:{style:BorderStyle.SINGLE,size:4,color:LINE}, right:{style:BorderStyle.SINGLE,size:4,color:LINE}, insideHorizontal:{style:BorderStyle.SINGLE,size:4,color:LINE}, insideVertical:{style:BorderStyle.SINGLE,size:4,color:LINE} }, rows: prows }));
kids.push(P([ R("القاعدة الذهبية: ", { bold: true, color: ORANGE }), R("درب يعطي الخدمة · تانكي يكافئ عليها — هدية عند التسجيل، نقاط عند كل فعل، ورصيد ظاهر دائماً في الرئيسية.") ], { before: 120 }));

// ============ PART 2 ============
kids.push(new Paragraph({ children: [new PageBreak()] }));
kids.push(H1("٣.", "الجزء الثاني — الرئيسية والخدمات والخريطة والإشعارات"));
kids.push(H2("نقاط القوة", GREEN));
kids.push(bullet("المستوى ظاهر في الرئيسية:", "«المستوى البرتقالي» تحت الاسم — تانكي حاضر كطبقة من أول لحظة.", { leadColor: GREEN }));
kids.push(bullet("الإشعارات مدفوعة بتانكي:", "تأكيد كسب، تذكير استبدال، حالة الباقة، السحب، عروض — الولاء في كل نقطة تماس.", { leadColor: GREEN }));
kids.push(bullet("تبويب الإشعارات:", "(العروض/الرصيد/الباقات/النقاط/الكل) يمنع الزحمة ويرتّب الرسائل.", { leadColor: GREEN }));
kids.push(bullet("زر QR المركزي:", "نقطة الكسب/الدفع الأساسية عند المضخة والتجار في مكان بارز.", { leadColor: GREEN }));
kids.push(bullet("الخدمات شاملة:", "وقود، شحن كهربائي، مطاعم، غسيل، صراف — كل خدمة فرصة كسب.", { leadColor: GREEN }));

kids.push(H2("ملاحظات مهمة (مرتّبة بالأهم)"));
kids.push(noteBox("🔴 ١. المستوى ظاهر لكن أين رصيد النقاط؟",
  "المستوى واضح، لكن الرقم الذي يراقبه العميل (رصيد نقاطه) يجب أن يكون بطل الشاشة — رقم كبير يكبر ويحمّسه. المستوى وحده لا يكفي.", WARN_L, RED));
kids.push(noteBox("🟠 ٢. ازدحام الرئيسية — هوية تانكي قد تضيع:",
  "عرض فلاش + محطة + سيارة + محفظة + عروض سريعة + خدمات… كثير. يلزم مدخل واضح «نقاطي/تانكي» (كرت مميّز) وإلا صار الولاء عنصراً مبعثراً.", SOFT, WARN));
kids.push(noteBox("🟠 ٣. وظيفة زر QR غير محدّدة:",
  "دفع؟ كسب؟ الاثنان؟ هذه أهم نقطة تماس في البرنامج ويجب أن تكون واضحة («امسح للكسب والدفع»).", SOFT, WARN));
kids.push(noteBox("🟡 ٤. الخريطة فرصة ولاء ضائعة:",
  "الخريطة حالياً محدّد محطات فقط. يمكن تحويلها لمحرّك ولاء بتمييز المحطات ذات العروض/مضاعفة النقاط («×٢ نقاط هنا») لتوجيه العميل لمحطات درب.", WARN_L, WARN));
kids.push(noteBox("🟡 ٥. خطر إرهاق الإشعارات:",
  "عدد كبير من دفعات الولاء (كسب، استبدال، انتهاء باقة، سحب، عرض). التبويب جيد، لكن يجب حوكمة التكرار وفصل المعاملاتي المهم عن التسويقي.", WARN_L, WARN));

kids.push(H2("فكرة تربط الكل"));
kids.push(P([ R("إظهار شارة الكسب («اكسب نقاط تانكي») على كل خدمة في شاشة الخدمات — فيفهم العميل أن درب كله يكافئ، ويتحوّل تانكي من تبويب إلى طبقة تغطّي كل فعل، وهو جوهر نجاح البرنامج.", {}) ],
  { shading: { type: ShadingType.CLEAR, color: "auto", fill: ORANGE_L } }));

// ---- Conclusion ----
kids.push(H1("٤.", "الخلاصة"));
kids.push(bullet("", "التصميم قوي في الهوية والتنفيذ، وبرنامج تانكي حاضر بشكل جيد في المستوى والإشعارات."));
kids.push(bullet("", "أهم تحسينين: (١) تأجيل شحن المحفظة وطلب البيانات إلى ما بعد إظهار القيمة، (٢) إبراز رصيد النقاط وهدية الترحيب كبطل للتجربة."));
kids.push(bullet("", "المبدأ الحاكم: تانكي طبقة تكافئ كل فعل داخل درب — لا شاشة ولا تطبيق منفصل."));
kids.push(P([ R("وثيقة داخلية لمراجعة التصميم — مبنية على لقطات Darb Platform V1.", { italics: true, size: 18, color: GREY }) ],
  { before: 240, align: AlignmentType.CENTER }));

const doc = new Document({
  creator: "Darb", title: "ملاحظات مراجعة تصميم تطبيق درب - تانكي",
  styles: { default: { document: { run: { font: FONT, size: 22, color: INK } } },
    paragraphStyles: [ { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
      run: { font: FONT, bold: true, size: 26, color: INK } } ] },
  sections: [ { properties: { page: { size: { width: 12240, height: 15840 },
    margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } }, children: kids } ],
});
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(process.argv[2] || "notes.docx", buf); console.log("written", buf.length); });
