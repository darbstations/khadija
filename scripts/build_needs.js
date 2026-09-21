const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, PageBreak,
} = require("docx");
const fs = require("fs");

const FONT = "Arial";
const ORANGE = "E07C16";
const ORANGE_L = "FEF3E6";
const WHY_BG = "FBF3E9";
const GREY = "6D6E70";
const INK = "26262A";
const LINE = "D9D7D3";
const HEAD_BG = "F3EFEA";
const GREEN = "1E7E45";

function R(text, o = {}) {
  return new TextRun({ text, font: FONT, rightToLeft: true, size: o.size || 22,
    bold: o.bold, italics: o.italics, color: o.color || INK });
}
function P(runs, o = {}) {
  if (typeof runs === "string") runs = [R(runs, o)];
  return new Paragraph({ bidirectional: true, alignment: o.align || AlignmentType.RIGHT,
    spacing: { after: o.after != null ? o.after : 120, before: o.before || 0, line: o.line || 300 },
    children: runs });
}
function H1(num, text) {
  return new Paragraph({ bidirectional: true, alignment: AlignmentType.RIGHT, heading: HeadingLevel.HEADING_1,
    spacing: { before: 300, after: 130 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: ORANGE, space: 4 } },
    children: [ R(num + "  ", { bold: true, size: 26, color: ORANGE }), R(text, { bold: true, size: 26, color: INK }) ] });
}
function hcell(text, w) {
  return new TableCell({ width: { size: w, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, color: "auto", fill: HEAD_BG },
    margins: { top: 70, bottom: 70, left: 100, right: 100 }, verticalAlign: "center",
    children: [P([R(text, { bold: true, size: 21, color: INK })], { after: 0, align: AlignmentType.RIGHT })] });
}
function needCell(text, w, flag) {
  const runs = [R(text, { bold: true, size: 20 })];
  if (flag) runs.push(R("  [" + flag.t + "]", { bold: true, size: 15, color: flag.c }));
  return new TableCell({ width: { size: w, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, color: "auto", fill: HEAD_BG },
    margins: { top: 80, bottom: 80, left: 100, right: 100 }, verticalAlign: "center",
    children: [P(runs, { after: 0, line: 270 })] });
}
function whyCell(text, w) {
  return new TableCell({ width: { size: w, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, color: "auto", fill: WHY_BG },
    margins: { top: 80, bottom: 80, left: 100, right: 100 }, verticalAlign: "center",
    children: [P([R("↳ ", { bold: true, color: ORANGE }), R(text, { size: 20 })], { after: 0, line: 270 })] });
}
function section(num, title, rows) {
  const NW = 3150, WW = 6210;
  const trs = [ new TableRow({ tableHeader: true, children: [ hcell("الاحتياج", NW), hcell("السبب", WW) ] }) ];
  for (const r of rows) trs.push(new TableRow({ children: [ needCell(r[0], NW, r[2]), whyCell(r[1], WW) ] }));
  return [ H1(num, title), new Table({
    width: { size: NW + WW, type: WidthType.DXA }, columnWidths: [NW, WW], visuallyRightToLeft: true,
    borders: {
      top: { style: BorderStyle.SINGLE, size: 4, color: LINE }, bottom: { style: BorderStyle.SINGLE, size: 4, color: LINE },
      left: { style: BorderStyle.SINGLE, size: 4, color: LINE }, right: { style: BorderStyle.SINGLE, size: 4, color: LINE },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: LINE }, insideVertical: { style: BorderStyle.SINGLE, size: 4, color: LINE },
    }, rows: trs }) ];
}

const kids = [];

// ---- Header block ----
kids.push(new Paragraph({ spacing: { after: 0 }, alignment: AlignmentType.RIGHT, bidirectional: true,
  children: [ new TextRun({ text: "درب لخدمات المحطات", font: FONT, rightToLeft: true, bold: true, size: 22, color: GREY }) ] }));
kids.push(new Paragraph({ spacing: { before: 60, after: 0 }, alignment: AlignmentType.RIGHT, bidirectional: true,
  children: [ new TextRun({ text: "احتياجات ملف الشراكات — برنامج تانكي للولاء", font: FONT, rightToLeft: true, bold: true, size: 36, color: INK }) ] }));
kids.push(new Paragraph({ spacing: { before: 40, after: 0 }, alignment: AlignmentType.RIGHT, bidirectional: true,
  children: [ new TextRun({ text: "وثيقة رسمية للإدارة التنفيذية", font: FONT, rightToLeft: true, size: 22, color: ORANGE, bold: true }) ] }));
kids.push(new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { before: 120, after: 160 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 18, color: ORANGE, space: 8 } }, children: [] }));
kids.push(P("تعرض هذه الوثيقة الاحتياجات اللازمة لإنجاح ملف الشراكات في برنامج «تانكي»، موزّعة على أربع ركائز، مع بيان سبب كل احتياج. لا تتضمن هذه النسخة التكاليف؛ تُعرض الميزانيات في وثيقة منفصلة.", { after: 200 }));

// ---- 1 HR ----
section("١.", "الموارد البشرية", [
  ["مدير الشراكات", "يملك الملف كاملاً ويقود الاستراتيجية والتفاوض الكبير؛ بدون قائد تتشتت الجهود ولا تُسنَد المسؤولية."],
  ["مديرو تطوير الأعمال (BD)", "يوقّعون التجار ويبنون قاعدة الشركاء؛ بدونهم لا شركاء، ويبقى البرنامج وقوداً بلا قنوات استبدال."],
  ["محلل تجاري", "يحسب ربحية كل صفقة قبل توقيعها ويحمي هامش البرنامج؛ صمّام الأمان الذي يمنع الصفقات الخاسرة."],
  ["مديرو حسابات التجار", "يتابعون التاجر بعد التوقيع (تفعيل وأداء وتجديد)؛ بدونهم يوقّع التاجر ثم يُهمَل فيتسرّب."],
  ["أخصائي تسويق الشراكات", "يصمّم العروض والحملات المشتركة؛ بدونه يوجد التجار دون أن يعرف عنهم العميل فينخفض الاستبدال."],
  ["منسّق عمليات التفعيل (Onboarding)", "يفعّل التاجر تقنياً وتشغيلياً ويحل عوائقه اليومية؛ بدونه فجوة بين «وقّعنا» و«اشتغل فعلاً»."],
  ["مختص التسويات المالية", "يسوّي مستحقات التجار ويدير الرصيد المقدّم (Float)؛ بدونه فوضى مالية وخلافات على المستحقات."],
  ["دعم التكامل التقني", "يربط أنظمة التجار (POS/QR) بالبرنامج؛ بدونه لا يستطيع العميل الكسب أو الاستبدال لدى التاجر."],
]).forEach(k => kids.push(k));

// ---- 2 Departments ----
section("٢.", "الأقسام والوظائف الداعمة", [
  ["قسم الشراكات وتطوير الأعمال", "القسم المالك للملف؛ بدونه لا جهة مسؤولة عن نجاح الشراكات."],
  ["القانوني والعقود", "يصوغ ويحمي كل اتفاقية؛ بدونه تتعرّض درب لالتزامات ونزاعات غير محسوبة."],
  ["المالية", "تدير التسويات والرصيد المقدّم والفوترة والضريبة؛ بدونها يتحرّك المال دون ضبط."],
  ["التقنية والتكامل", "يبني الواجهات البرمجية ويربط المحفظة والمضخات؛ بدونه لا منتج قابل للتشغيل."],
  ["التسويق والاتصال", "يطلق البرنامج ويوصّل العروض للعميل والتاجر؛ بدونه يبقى أفضل برنامج مجهولاً."],
  ["الامتثال والمخاطر", "يضمن ترخيص ساما وحماية البيانات (PDPL) والالتزام الشرعي؛ بدونه لا ينطلق البرنامج قانونياً.", { t: "حرج", c: "B23A34" }],
  ["خدمة العملاء", "ترد على استفسارات النقاط والاستبدال؛ بدونها يقتل إحباط العميل الثقة بالبرنامج."],
  ["البيانات والتحليلات", "يقيس أداء الشركاء ومعدلات الكسب والحرق والاسترداد؛ بدونه تُقاد البرنامج دون مؤشرات."],
]).forEach(k => kids.push(k));

// ---- 3 Contracts ----
kids.push(new Paragraph({ children: [new PageBreak()] }));
section("٣.", "التعاقدات والاتفاقيات", [
  ["اتفاقية التاجر الرئيسية", "تحدّد نسب الكسب والحرق والرسوم والتسويات والالتزامات؛ بدونها كل تاجر بلا مرجع ملزم."],
  ["اتفاقية شبكة الاستبدال (Gift Card API)", "تؤمّن كتالوج الاستبدال مع بقاء البيانات لدى درب؛ بدونها لا قناة استبدال جاهزة."],
  ["اتفاقية شريك المحفظة المرخّص", "تتيح تشغيل المحفظة قانونياً دون ترخيص ساما من الصفر؛ بدونها لا محفظة نقدية نظامية."],
  ["اتفاقية بوابة/معالج الدفع", "تمكّن شحن المحفظة عبر مدى/فيزا؛ بدونها لا يستطيع العميل شحن رصيده."],
  ["اتفاقية مستوى الخدمة (SLA)", "تضمن توفّر واستجابة الشركاء التقنيين؛ بدونها تبقى الأعطال دون التزام بالإصلاح."],
  ["اتفاقية معالجة البيانات (DPA)", "تحمي بيانات العملاء وفق نظام PDPL مع كل شريك؛ بدونها مخالفة نظامية ومخاطر تسريب."],
  ["اتفاقية عدم الإفصاح (NDA)", "تحمي أسرار درب التجارية أثناء التفاوض؛ بدونها تنكشف المعلومات الحساسة."],
  ["اتفاقية التمويل المشترك (Co-funding)", "تنظّم تقاسم تكلفة العروض مع التاجر؛ بدونها خلاف على من يتحمّل تكلفة كل حملة."],
  ["الموافقة التنظيمية + الفتوى الشرعية", "ترخّص المحفظة وتجيز النموذج شرعياً؛ بدونها يتعطّل البرنامج قانونياً وشرعياً."],
]).forEach(k => kids.push(k));

// ---- 4 Systems ----
section("٤.", "الأنظمة والممكّنات", [
  ["نظام إدارة التجار + حاسبة الكسب/الحرق", "لتوثيق كل تاجر وحساب ربحيته؛ بدونه يجري التفاوض دون أداة. (مبنيّ بالفعل في تطبيق تانكي.)", { t: "جاهز", c: GREEN }],
  ["آلية الرصيد المقدّم / الضمان (Escrow)", "لتغطية قيمة القسائم وضمان مبدأ «النظام لا يخسر»؛ بدونها مخاطرة مالية مباشرة."],
  ["لوحة مؤشرات الأداء (KPIs)", "لمتابعة عدد التجار والتفعيل والاسترداد وتكلفة الاكتساب؛ بدونها لا يُعرف مسار النجاح."],
  ["قناة تواصل مع التجار", "لإدارة العلاقة والعروض والتحديثات؛ بدونها يشعر الشركاء بالإهمال."],
]).forEach(k => kids.push(k));

// ---- Conclusion ----
kids.push(new Paragraph({ bidirectional: true, alignment: AlignmentType.RIGHT,
  spacing: { before: 260, after: 120, line: 300 },
  border: { top: { style: BorderStyle.SINGLE, size: 4, color: LINE, space: 6 },
    left: { style: BorderStyle.SINGLE, size: 18, color: ORANGE, space: 8 } },
  shading: { type: ShadingType.CLEAR, color: "auto", fill: WHY_BG },
  children: [ R("الخلاصة:  ", { bold: true, color: ORANGE }),
    R("ينجح ملف الشراكات باكتمال أربع ركائز متوازية — فريق متخصّص يقوده مدير الشراكات ومحلل تجاري، وأقسام داعمة أحرجها الامتثال والمخاطر، وحزمة عقود أهمها اتفاقية التاجر وشريك المحفظة المرخّص، وأنظمة تشغيل جوهرها جاهز في تطبيق تانكي. غياب أي ركيزة يُضعف الملف بأكمله.") ] }));
kids.push(P([ R("وثيقة داخلية للمناقشة التنفيذية — تُعرض الميزانيات والجداول الزمنية في وثائق منفصلة.", { italics: true, size: 18, color: GREY }) ], { before: 220, align: AlignmentType.CENTER }));

const doc = new Document({
  creator: "Darb", title: "احتياجات ملف الشراكات - تانكي",
  styles: { default: { document: { run: { font: FONT, size: 22, color: INK } } },
    paragraphStyles: [ { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
      run: { font: FONT, bold: true, size: 26, color: INK } } ] },
  sections: [ { properties: { page: { size: { width: 12240, height: 15840 },
    margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } }, children: kids } ],
});
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(process.argv[2] || "needs.docx", buf); console.log("written", buf.length); });
