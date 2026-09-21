const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType,
  PageBreak, LevelFormat, PositionalTab, PositionalTabAlignment, PositionalTabLeader,
} = require("docx");
const fs = require("fs");

const FONT = "Arial";
const ORANGE = "E07C16";
const ORANGE_L = "FEF3E6";
const GREY = "6D6E70";
const INK = "26262A";
const LINE = "D9D7D3";
const GREEN = "1E7E45";
const RED = "B23A34";
const WARN = "9A6A12";
const HEAD_BG = "F3EFEA";

// ---------- helpers ----------
function R(text, o = {}) {
  return new TextRun({ text, font: FONT, rightToLeft: true, size: o.size || 22,
    bold: o.bold, italics: o.italics, color: o.color || INK, break: o.break });
}
function P(runs, o = {}) {
  if (typeof runs === "string") runs = [R(runs, o)];
  return new Paragraph({
    bidirectional: true, alignment: o.align || AlignmentType.RIGHT,
    spacing: { after: o.after != null ? o.after : 120, before: o.before || 0, line: o.line || 300 },
    children: runs, ...(o.border ? { border: o.border } : {}),
    ...(o.bullet ? { numbering: { reference: o.bullet.ref, level: 0 } } : {}),
    ...(o.heading ? { heading: o.heading } : {}),
    ...(o.shading ? { shading: o.shading } : {}),
  });
}
function H1(num, text) {
  return new Paragraph({
    bidirectional: true, alignment: AlignmentType.RIGHT, heading: HeadingLevel.HEADING_1,
    spacing: { before: 320, after: 140 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: ORANGE, space: 4 } },
    children: [ R(num + "  ", { bold: true, size: 26, color: ORANGE }), R(text, { bold: true, size: 26, color: INK }) ],
  });
}
function H2(text) {
  return new Paragraph({
    bidirectional: true, alignment: AlignmentType.RIGHT, heading: HeadingLevel.HEADING_2,
    spacing: { before: 200, after: 100 },
    children: [ R(text, { bold: true, size: 23, color: ORANGE }) ],
  });
}
function bullet(text, boldLead) {
  const runs = [];
  if (boldLead) { runs.push(R(boldLead + " ", { bold: true })); runs.push(R(text)); }
  else runs.push(R(text));
  return P(runs, { bullet: { ref: "b" }, after: 60, line: 290 });
}
function cell(children, o = {}) {
  if (typeof children === "string") children = [P(children, { after: 0, size: o.size || 20, align: o.align })];
  return new TableCell({
    width: { size: o.w, type: WidthType.DXA },
    shading: o.shade ? { type: ShadingType.CLEAR, color: "auto", fill: o.shade } : undefined,
    margins: { top: 70, bottom: 70, left: 90, right: 90 },
    verticalAlign: "center",
    children,
  });
}
function hcell(text, w, shade) {
  return new TableCell({
    width: { size: w, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, color: "auto", fill: shade || HEAD_BG },
    margins: { top: 70, bottom: 70, left: 90, right: 90 }, verticalAlign: "center",
    children: [P([R(text, { bold: true, size: 20, color: INK })], { after: 0, align: AlignmentType.CENTER })],
  });
}
function tbl(widths, rows) {
  return new Table({
    width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    columnWidths: widths, visuallyRightToLeft: true,
    borders: {
      top: { style: BorderStyle.SINGLE, size: 4, color: LINE },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: LINE },
      left: { style: BorderStyle.SINGLE, size: 4, color: LINE },
      right: { style: BorderStyle.SINGLE, size: 4, color: LINE },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: LINE },
      insideVertical: { style: BorderStyle.SINGLE, size: 4, color: LINE },
    },
    rows,
  });
}
function tc(text, w, o = {}) { // text cell with color/size
  return cell([P([R(text, { size: o.size || 19, color: o.color || INK, bold: o.bold })],
    { after: 0, align: o.align || AlignmentType.RIGHT, line: 260 })], { w, shade: o.shade });
}

const TW = 9360; // usable text width ~ Letter minus 1440 margins

// ================= DOCUMENT =================
const kids = [];

// ---------- COVER ----------
kids.push(new Paragraph({ spacing: { before: 1400, after: 0 }, alignment: AlignmentType.CENTER, bidirectional: true,
  children: [ new TextRun({ text: "درب لخدمات المحطات", font: FONT, rightToLeft: true, bold: true, size: 30, color: GREY }) ] }));
kids.push(new Paragraph({ spacing: { before: 80, after: 40 }, alignment: AlignmentType.CENTER, bidirectional: true,
  children: [ new TextRun({ text: "برنامج تانكي للولاء", font: FONT, rightToLeft: true, bold: true, size: 24, color: ORANGE }) ] }));
kids.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 320, after: 0 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 18, color: ORANGE, space: 10 } }, children: [] }));
kids.push(new Paragraph({ spacing: { before: 420, after: 0 }, alignment: AlignmentType.CENTER, bidirectional: true,
  children: [ new TextRun({ text: "تقرير استراتيجي للإدارة التنفيذية", font: FONT, rightToLeft: true, bold: true, size: 40, color: INK }) ] }));
kids.push(new Paragraph({ spacing: { before: 160, after: 0 }, alignment: AlignmentType.CENTER, bidirectional: true,
  children: [ new TextRun({ text: "تقييم عرض شركة ديسكوير والتوصية بمسار التنفيذ", font: FONT, rightToLeft: true, size: 26, color: GREY }) ] }));

kids.push(new Paragraph({ spacing: { before: 900 }, children: [] }));
const meta = [
  ["مقدَّم إلى", "الإدارة التنفيذية — شركة درب"],
  ["الموضوع", "بناء برنامج الولاء داخلياً مقابل التعاقد مع مشغّل خارجي"],
  ["التاريخ", "يوليو 2026"],
  ["التصنيف", "سري — للتداول الداخلي"],
];
kids.push(new Table({
  width: { size: 6600, type: WidthType.DXA }, columnWidths: [2000, 4600], visuallyRightToLeft: true,
  alignment: AlignmentType.CENTER,
  borders: { top: {style:BorderStyle.NONE}, bottom:{style:BorderStyle.NONE}, left:{style:BorderStyle.NONE}, right:{style:BorderStyle.NONE}, insideHorizontal:{style:BorderStyle.SINGLE,size:4,color:LINE}, insideVertical:{style:BorderStyle.NONE} },
  rows: meta.map(([k, v]) => new TableRow({ children: [
    cell([P([R(k, { bold: true, color: ORANGE, size: 20 })], { after: 0 })], { w: 2000 }),
    cell([P([R(v, { size: 20 })], { after: 0 })], { w: 4600 }),
  ]})),
}));
kids.push(new Paragraph({ children: [new PageBreak()] }));

// ---------- 1. EXECUTIVE SUMMARY ----------
kids.push(H1("١.", "الملخص التنفيذي"));
kids.push(P([
  R("تعمل درب على إطلاق برنامج ولاء «تانكي» يربط تعبئة الوقود بالتجار والمحفظة الرقمية. تقدّمت شركة ديسكوير بعرض يتضمّن خارطة طريق لأربعة أشهر وقوائم شركاء مقترحة. يقيّم هذا التقرير العرض، ويوضّح أثره على "),
  R("التشغيل والوضع المالي والتسويق والأصول الاستراتيجية", { bold: true }),
  R("، ويقدّم توصية واضحة بمسار التنفيذ."),
]));
kids.push(P([ R("الخلاصة: ", { bold: true, color: ORANGE }),
  R("يوصى بأن تحتفظ درب بملكية "), R("نواة البرنامج", { bold: true }),
  R(" (محرّك النقاط + بيانات العملاء + المحفظة + العلامة)، وأن تستأجر "),
  R("شركاء متخصّصين", { bold: true }),
  R(" للأجزاء الصعبة فقط — تماماً كما فعل المنافس المباشر «ساسكو». تسليم المنصة والبيانات بالكامل لمشغّل خارجي يكلّف درب أثمن أصولها دون أن يحلّ أصعب بند تقني، الذي يبقى مسؤولية درب في كل السيناريوهات."),
]));
kids.push(H2("أبرز النتائج"));
kids.push(bullet("قوائم شركاء ديسكوير عامة (أزياء وتجزئة وعطور) ولا تخدم عميل محطة الوقود ذا الصرف اليومي المتكرر.", "١)"));
kids.push(bullet("في نموذج ديسكوير، تصبح بيانات العملاء والحسابات على منصتهم — أي تفقد درب أثمن أصولها وتتحوّل إلى واجهة.", "٢)"));
kids.push(bullet("أصعب بند (المحفظة + تكامل مضخات الوقود) يبقى مسؤولية درب في كل السيناريوهات، بإقرار خارطة ديسكوير نفسها.", "٣)"));
kids.push(bullet("المحفظة النقدية تستلزم ترخيص/شراكة تنظيمية من ساما؛ النموذج المغلق يخفّف العبء، والشراكة المرخّصة تختصره.", "٤)"));
kids.push(bullet("شبكة استبدال ديسكوير قابلة للاستبدال بمزوّد Gift Card API أبسط تقنياً مع احتفاظ درب بالملكية الكاملة.", "٥)"));
kids.push(bullet("المنافس المباشر «ساسكو» اعتمد المسار الموصى به: برنامج مغلق يملكه، مع شريك تمويل مرخّص (إمكان).", "٦)"));

// ---------- 2. BACKGROUND ----------
kids.push(H1("٢.", "الخلفية والسياق"));
kids.push(P("«تانكي» هو برنامج ولاء تطبيق درب: يجمع العميل نقاطاً من تعبئة الوقود ومن التجار المتعاقدين، ويستبدلها داخل شبكة درب (وقود + باقات + عروض شركاء)، عبر محفظة رقمية مغلقة. صُمّم النموذج ليكون النظام محميّاً مالياً عبر آلية ضمان (Escrow)، وليخضع لتنظيم أخف عبر إبقاء المحفظة مغلقة."));
kids.push(P("قدّمت ديسكوير مستندين: خارطة طريق للانطلاق على أربعة أشهر، وقائمة شركاء مقترحة (شركاء كسب وشركاء استبدال). يقيّم هذا التقرير الجدوى الاستراتيجية للعرض مقارنةً ببناء البرنامج داخلياً."));

// ---------- 3. EVALUATION OF DSQUARES OFFER ----------
kids.push(H1("٣.", "تقييم عرض ديسكوير"));
kids.push(H2("٣.١  خارطة الطريق"));
kids.push(P("مقترح على مرحلتين خلال أربعة أشهر من الانطلاق (M0):"));
kids.push(bullet("كسب واستبدال على الوقود + المستويات + كتالوج مكافآت (٢٠–٣٠ تاجراً) + استبدال قسائم في المتاجر عبر شبكتهم (١٥٠٠ فرع / ٣٠٠ علامة) + تسليم واجهات برمجية + اختبار وإطلاق.", "المرحلة ١ (M0–M2) — الحد الأدنى القابل للتشغيل:"));
kids.push(bullet("كسب من المتاجر على إجمالي الفاتورة + شحن جوال + تبرعات + الحلقة المفتوحة (Open Loop) على محفظة نقدية.", "المرحلة ٢ (M2–M4) — التوسّع:"));
kids.push(P([ R("ملاحظة جوهرية: ", { bold: true, color: WARN }),
  R("خارطة الطريق لا تتضمّن اقتصاد النقطة (قيمة النقطة، معدل الكسب، جهة التمويل، الرسوم)، وتضع «جاهزية المحفظة النقدية» و«تكامل مضخات الوقود» على المسار الحرج وبمسؤولية درب.") ]));

kids.push(H2("٣.٢  قوائم الشركاء"));
kids.push(tbl([2600, 2000, 4760], [
  new TableRow({ tableHeader: true, children: [ hcell("القائمة", 2600), hcell("العدد المقترح", 2000), hcell("الملاحظة", 4760) ] }),
  new TableRow({ children: [ tc("شركاء الاستبدال", 2600, { bold: true }), tc("٢٠٢ تاجراً", 2000, { align: AlignmentType.CENTER }), tc("علامات كبيرة (أمازون، جرير، نون…)؛ استبدال بطاقات هدايا، هامش تفاوضي منخفض.", 4760) ] }),
  new TableRow({ children: [ tc("شركاء الكسب", 2600, { bold: true }), tc("٦٥ تاجراً", 2000, { align: AlignmentType.CENTER }), tc("علامات محلية أصغر (مطاعم ولايف ستايل)؛ ملائمة للتكرار لكن ضعيفة في الفئات اليومية.", 4760) ] }),
  new TableRow({ children: [ tc("نسبة النجاح الفعلية", 2600, { bold: true, shade: ORANGE_L }), tc("٢٥٪ – ٣٠٪", 2000, { align: AlignmentType.CENTER, shade: ORANGE_L, color: RED, bold: true }), tc("بإقرار ديسكوير: من ٢٠٢ تاجر استبدال يُتوقّع تعاقد ~٥٠–٦٠ فعلياً.", 4760, { shade: ORANGE_L }) ] }),
]));
kids.push(P([ R("تنبيهات من إقرار ديسكوير: ", { bold: true }),
  R("بعض التجار يشترطون تمويلاً مشتركاً أو دعماً تسويقياً؛ نماذج الإصدار/التعبئة قد تتطلب رصيداً مقدّماً (Float) للحفاظ على مخزون القسائم — وهو بند سيولة نقدية يجب احتسابه.") ], { before: 80 }));

// ---------- 4. STRATEGIC CONSIDERATIONS ----------
kids.push(H1("٤.", "الاعتبارات الاستراتيجية الثلاثة"));
kids.push(bullet("قائمة ديسكوير مبنية لبرنامج بنك/بطاقة ائتمانية لا لمحطة وقود؛ يجمع العميل نقاطاً من الوقود دون قنوات استبدال تخدم يومه، فينخفض التفاعل.", "ملاءمة القائمة لعميل المحطات:"));
kids.push(bullet("من يملك البيانات يملك القدرة على التسويق والفهم والبناء المستقبلي. في نموذج ديسكوير تكون البيانات لديهم.", "ملكية بيانات العميل:"));
kids.push(bullet("إجراء كل نقطة ورصيد ومعاملة على منصتهم يجعل درب مستأجِرة للنظام لا مالكة له، مع ارتهان تعاقدي وخروج مكلف.", "منصة الحسابات:"));

// ---------- 5. WALLET & LICENSING ----------
kids.push(H1("٥.", "تحليل المحفظة والترخيص التنظيمي"));
kids.push(P([ R("القاعدة الأساسية: ", { bold: true }), R("النقاط مكافأة ولا تحتاج ترخيصاً؛ أما الرصيد النقدي الذي يشحنه العميل فهو أموال جمهور، ومسكها نشاط تنظّمه مؤسسة النقد (ساما) حمايةً للعميل ومكافحةً لغسل الأموال وضماناً للاستقرار المالي.") ]));
kids.push(P([ R("موقع درب: ", { bold: true, color: ORANGE }),
  R("هدف تجميع رصيد المحافظ (Float) يعني حكماً مسك أموال الجمهور. النموذج المغلق (الرصيد يشتري وقوداً وباقات درب فقط، بلا سحب نقدي أو دفع لطرف ثالث) يخفّف العبء التنظيمي مقارنةً بالحلقة المفتوحة التي تقترب من ترخيص الأموال الإلكترونية الكامل.") ]));
kids.push(H2("المخرجان"));
kids.push(bullet("ترخيص خاص بدرب: سيطرة كاملة، لكن بطيء ومكلّف وبمتطلبات رأس مال وامتثال ثقيلة.", "١)"));
kids.push(bullet("شراكة محفظة/تمويل مرخّص (موصى به): يحمل الشريك عبء الترخيص والامتثال، وتحتفظ درب بالبرنامج والبيانات والعلاقة — ويختصر أثقل بند تنظيمي.", "٢)"));

// ---------- 6. DARB vs DSQUARES ----------
kids.push(new Paragraph({ children: [new PageBreak()] }));
kids.push(H1("٦.", "المقارنة: درب مالكة للمنصة مقابل ديسكوير مشغّل كامل"));
const cmpRows = [
  ["⚙️ التشغيل", "تحكّم كامل بالمحرّك والقواعد؛ إطلاق أبطأ لكن مستقل.", "منصة جاهزة وإطلاق أسرع؛ أي تعديل يخضع لأولوياتهم."],
  ["💰 المالي", "هامش النقطة كامل؛ الـ Float يبقى داخل درب؛ لا رسوم متكررة.", "Setup + رسوم شهرية + رسوم لكل معاملة؛ تسرّب الـ Float."],
  ["📣 التسويق", "ملكية البيانات وحملات دقيقة؛ قائمة مخصّصة لعميل المحطات.", "قائمة عامة غير ملائمة؛ قدرة تسويق محدودة لأن البيانات لديهم."],
  ["🏛️ الأصول", "البيانات والمحفظة والعلامة ملك درب؛ أصل متراكم بلا ارتهان.", "البيانات والحسابات لديهم؛ ارتهان تعاقدي وخروج مكلّف."],
];
kids.push(tbl([1700, 3830, 3830], [
  new TableRow({ tableHeader: true, children: [ hcell("البُعد", 1700), hcell("🟠 درب تملك المنصة", 1700===0?0:3830, ORANGE_L), hcell("⚪ ديسكوير مشغّل كامل", 3830) ] }),
  ...cmpRows.map(r => new TableRow({ children: [
    tc(r[0], 1700, { bold: true, shade: HEAD_BG }),
    tc(r[1], 3830, { shade: ORANGE_L }),
    tc(r[2], 3830),
  ]})),
]));
kids.push(P([ R("الحكم: ", { bold: true, color: ORANGE }), R("متقارب في التشغيل، ولصالح درب في المالي والتسويق والأصول.") ], { before: 100 }));

// ---------- 7. SASCO BENCHMARK ----------
kids.push(H1("٧.", "المقارنة المرجعية — ساسكو (المنافس المباشر)"));
kids.push(P("يوفّر تطبيق «ساسكو Control» محفظة رقمية (شحن عبر مدى/فيزا/ماستركارد)، ودفعاً بـ QR، ونقاط كاش باك تُستبدل داخل مطاعم وكافيهات ساسكو، وخدمة «اشحن وادفع لاحقاً» بالشراكة مع شركة التمويل المرخّصة «إمكان»."));
kids.push(H2("الدلالة"));
kids.push(bullet("نقاط ساسكو مغلقة (تُستبدل داخل منافذهم) — نفس اختيار درب.", "•"));
kids.push(bullet("لم يبنِ ساسكو الجزء المالي المرخّص بنفسه، بل استأجر شريكاً مرخّصاً (إمكان) — وهو المخرج الموصى به لدرب.", "•"));
kids.push(bullet("احتفظ ساسكو بالمنصة والبيانات والعميل، ولم يسلّمها لمشغّل خارجي.", "•"));
kids.push(P([ R("الخلاصة: ", { bold: true }), R("اختار المنافس المباشر مسار الملكية + الشركاء المرخّصين، لا مسار التسليم الكامل — وهو ذاته المسار الموصى به في هذا التقرير.") ]));

// ---------- 8. ALTERNATIVES ----------
kids.push(H1("٨.", "بدائل شبكة ديسكوير"));
kids.push(P("يجمع عرض ديسكوير خدمتين قابلتين للفصل: شبكة الاستبدال (بطاقات هدايا) وهي سهلة الاستبدال، وشركاء الكسب (تجار محليون) وهي قيمة درب الأساسية التي لا يُستحسن إسنادها لطرف خارجي."));
kids.push(tbl([1900, 2490, 2490, 2480], [
  new TableRow({ tableHeader: true, children: [
    hcell("المعيار", 1900), hcell("ديسكوير — مشغّل كامل", 2490),
    hcell("البديل الموصى", 2490, ORANGE_L), hcell("بناء كامل من الصفر", 2480) ] }),
  ...[
    ["سرعة الإطلاق", "سريعة", "سريعة", "بطيئة"],
    ["التكلفة", "رسوم متكررة", "لكل قسيمة فقط", "Capex عالٍ"],
    ["الجهد التقني", "منخفض", "متوسط", "مرتفع جداً"],
    ["الملكية والبيانات", "لديهم", "كاملة لدرب", "كاملة لدرب"],
    ["الارتهان", "مرتفع", "شبه معدوم", "لا يوجد"],
  ].map(r => new TableRow({ children: [
    tc(r[0], 1900, { bold: true, shade: HEAD_BG }),
    tc(r[1], 2490, { align: AlignmentType.CENTER }),
    tc(r[2], 2490, { align: AlignmentType.CENTER, shade: ORANGE_L, bold: true }),
    tc(r[3], 2480, { align: AlignmentType.CENTER }),
  ]})),
]));
kids.push(H2("البديل الموصى — ثلاث قطع"));
kids.push(bullet("مزوّد Gift Card API نقي (مثل YouGotaGift أو Tillo) يوفّر كتالوج بطاقات هدايا عبر تكامل واحد، مع بقاء البيانات وحساب النقاط لدى درب. (تُثبَّت الأسماء والأسعار عند التعاقد.)", "الاستبدال:"));
kids.push(bullet("توقيع التجار المحليين مباشرةً عبر فريق تطوير أعمال — مصدر الهامش والعلاقة بالعميل؛ والأداة (قاعدة التجار + نموذج Earn/Burn) مبنية في تطبيق تانكي.", "الكسب:"));
kids.push(bullet("التركيب فوق محفظة/تمويل مرخّص من ساما (كما فعل ساسكو مع إمكان) بدل الترخيص من الصفر.", "المحفظة:"));

// ---------- 9. RECOMMENDATIONS ----------
kids.push(H1("٩.", "التوصيات"));
kids.push(P([ R("المسار الوسط: ", { bold: true, color: ORANGE }), R("لا بناء كل شيء، ولا تسليم كل شيء.") ]));
kids.push(bullet("تملك درب النواة: محرّك النقاط + بيانات العملاء + المحفظة + العلامة.", "١)"));
kids.push(bullet("شراكة محفظة/تمويل مرخّص لاختصار أثقل بند تنظيمي.", "٢)"));
kids.push(bullet("استبدال شبكة ديسكوير بمزوّد Gift Card API لقناة الاستبدال، مع بقاء البيانات لدى درب.", "٣)"));
kids.push(bullet("بناء شبكة شركاء الكسب داخلياً باستخدام أدوات تانكي الجاهزة.", "٤)"));
kids.push(bullet("حصر أي تعامل مع ديسكوير في دور مورّد قناة استبدال ثانوية عند الحاجة — لا مشغّل للمنصة.", "٥)"));

// ---------- 10. PRE-SIGNATURE QUESTIONS ----------
kids.push(H1("١٠.", "أسئلة ما قبل التوقيع (لأي مورّد)"));
kids.push(bullet("ورقة اقتصاد كاملة: قيمة النقطة، معدل الكسب، جهة تمويل الاستبدال، وهيكل الرسوم (Setup + شهري + لكل معاملة).", "•"));
kids.push(bullet("مَن يملك بيانات العملاء والمحفظة تعاقدياً؟ وهل يُضمن التصدير الكامل عند الخروج؟", "•"));
kids.push(bullet("موقف الترخيص للمحفظة النقدية والحلقة المفتوحة، ومَن يتحمّل تبعة ساما.", "•"));
kids.push(bullet("حجم الرصيد المقدّم (Float) المطلوب للقسائم، وشروط الدفع للتجار.", "•"));
kids.push(bullet("قائمة الكسب المضمونة فعلاً (لا المقترحة) بأسعار مثبّتة.", "•"));

// ---------- 11. CONCLUSION ----------
kids.push(H1("١١.", "الخلاصة"));
kids.push(P("بما أن أصعب بند (المحفظة وتكامل المضخات) يبقى مسؤولية درب في كل السيناريوهات، فإن التنازل عن ملكية المنصة والبيانات لمشغّل خارجي لا يحلّ هذا البند، بل يكلّف درب أثمن أصولها دون مقابل حقيقي. التوصية أن تحتفظ درب بملكية نواة برنامج تانكي، وتستأجر شركاء متخصّصين مرخّصين للأجزاء الصعبة — وهو المسار الذي أثبت المنافس المباشر «ساسكو» جدواه."));
kids.push(P([ R("وثيقة داخلية للمناقشة التنفيذية — تُثبَّت الأرقام والبنود التجارية وأسماء المورّدين عند التفاوض النهائي.", { italics: true, size: 18, color: GREY }) ], { before: 240, align: AlignmentType.CENTER }));

// ================= BUILD =================
const doc = new Document({
  creator: "Darb", title: "تقرير تانكي التنفيذي",
  numbering: { config: [ { reference: "b", levels: [ { level: 0, format: LevelFormat.BULLET, text: "•",
    alignment: AlignmentType.RIGHT, style: { run: { color: ORANGE, font: FONT }, paragraph: { indent: { right: 360, hanging: 220 } } } } ] } ] },
  styles: { default: { document: { run: { font: FONT, size: 22, color: INK } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: FONT, bold: true, size: 26, color: INK } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: FONT, bold: true, size: 23, color: ORANGE } },
    ] },
  sections: [ {
    properties: { page: {
      size: { width: 12240, height: 15840 },
      margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 },
    } },
    children: kids,
  } ],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(process.argv[2] || "report.docx", buf);
  console.log("written", (process.argv[2] || "report.docx"), buf.length, "bytes");
});
