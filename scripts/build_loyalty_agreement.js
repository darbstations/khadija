const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, BorderStyle, ShadingType,
  Table, TableRow, TableCell, WidthType,
} = require("docx");
const fs = require("fs");

const FONT = "Arial";
const ORANGE = "E07C16", ORANGE_L = "FEF3E6", SOFT = "FBF3E9";
const GREY = "6D6E70", INK = "26262A", LINE = "C9C7C3", HEAD_BG = "F3EFEA";

function R(t, o = {}) { return new TextRun({ text: t, font: FONT, rightToLeft: true, size: o.size || 22, bold: o.bold, italics: o.italics, color: o.color || INK }); }
function P(runs, o = {}) {
  if (typeof runs === "string") runs = [R(runs, o)];
  return new Paragraph({ bidirectional: true, alignment: o.align || AlignmentType.RIGHT,
    spacing: { after: o.after != null ? o.after : 100, before: o.before || 0, line: o.line || 300 },
    ...(o.shading ? { shading: o.shading } : {}), children: runs });
}
function article(num, title) {
  return new Paragraph({ bidirectional: true, alignment: AlignmentType.RIGHT, heading: HeadingLevel.HEADING_1,
    spacing: { before: 240, after: 100 },
    children: [ R("المادة (" + num + "): ", { bold: true, size: 23, color: ORANGE }), R(title, { bold: true, size: 23, color: INK }) ] });
}
function item(text, sub) {
  return new Paragraph({ bidirectional: true, alignment: AlignmentType.RIGHT, spacing: { after: 60, line: 290 },
    bullet: { level: sub ? 1 : 0 }, children: [R(text)] });
}
function cell(text, w, o = {}) {
  return new TableCell({ width: { size: w, type: WidthType.DXA }, shading: o.shade ? { type: ShadingType.CLEAR, color: "auto", fill: o.shade } : undefined,
    margins: { top: 60, bottom: 60, left: 90, right: 90 }, verticalAlign: "center",
    children: [P([R(text, { size: o.size || 19, bold: o.bold, color: o.color || INK })], { after: 0, align: o.align || AlignmentType.RIGHT, line: 260 })] });
}

const kids = [];

// ===== TITLE =====
kids.push(P([R("شركة درب لخدمات المحطات", { bold: true, size: 22, color: GREY })], { after: 20 }));
kids.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 80, after: 0 }, bidirectional: true,
  children: [R("اتفاقية الانضمام لبرنامج الولاء «تانكي»", { bold: true, size: 34, color: INK })] }));
kids.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 40, after: 0 }, bidirectional: true,
  children: [R("(قالب موحّد للمستأجرين / التجّار)", { size: 22, color: ORANGE, bold: true })] }));
kids.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 100, after: 160 }, border: { bottom: { style: BorderStyle.SINGLE, size: 18, color: ORANGE, space: 8 } }, children: [] }));

// ===== PARTIES =====
kids.push(P([R("أُبرمت هذه الاتفاقية في يوم [   ] بتاريخ [   /   /   14 هـ] الموافق [   /   /   20 م]، بين كلٍ من:", {})], { after: 100 }));
kids.push(P([R("الطرف الأول: ", { bold: true }), R("شركة درب لخدمات المحطات، سجل تجاري رقم [   ]، وعنوانها [   ]، ويُشار إليها بـ «درب» أو «مشغّل البرنامج».")], { after: 80 }));
kids.push(P([R("الطرف الثاني: ", { bold: true }), R("[اسم المستأجر/التاجر]، سجل تجاري رقم [   ]، وعنوانه [   ]، ويُشار إليه بـ «التاجر».")], { after: 100 }));
kids.push(P([R("تمهيد: ", { bold: true }), R("حيث إن درب تُشغّل برنامج ولاء باسم «تانكي» يمنح عملاءها نقاطاً قابلة للاستبدال، ويرغب التاجر في الانضمام للبرنامج ضمن فروعه في مواقع درب؛ فقد اتفق الطرفان — بكامل أهليتهما المعتبرة شرعاً ونظاماً — على ما يلي:")], { after: 60 }));
kids.push(P([R("تُعدّ التعريفات والملاحق جزءاً لا يتجزأ من هذه الاتفاقية.", { italics: true, size: 20, color: GREY })], { after: 120 }));

// ===== ARTICLES =====
kids.push(article("1", "التعريفات"));
[
  "النقطة: وحدة المكافأة في برنامج تانكي، وقيمتها الاستبدالية محدّدة في الجدول التجاري (الملحق أ).",
  "الكسب (Earn): منح العميل نقاطاً عند شرائه من فروع التاجر المشمولة.",
  "الحرق/الاستبدال (Burn): استخدام العميل نقاطه للحصول على خصم أو منتج لدى التاجر.",
  "المعاملة المؤهِّلة: عملية شراء تنطبق عليها شروط الكسب أو الاستبدال وفق هذه الاتفاقية.",
  "العرض: تخفيض أو ميزة يقدّمها التاجر حصرياً لعملاء تطبيق درب.",
  "التسوية: احتساب وتبادل المستحقات المالية بين الطرفين دورياً.",
].forEach(t => kids.push(item(t)));

kids.push(article("2", "نطاق الاتفاقية"));
kids.push(item("تسري هذه الاتفاقية على فروع التاجر المحدّدة في الجدول التجاري (الملحق أ) داخل مواقع درب."));
kids.push(item("يجوز إضافة أو حذف فروع بموجب تحديث للملحق (أ) موقّع من الطرفين."));

kids.push(article("3", "آلية الكسب (Earn)"));
kids.push(item("يمنح العميل نقاطاً عند الشراء من التاجر بالنسبة المحدّدة في الملحق (أ)."));
kids.push(item("يساهم التاجر في تمويل الكسب بالقيمة المحدّدة في الملحق (أ) (هللة عن كل ريال)، وتتحمّل درب الباقي إن وُجد."));
kids.push(item("لا تُحتسب النقاط على المعاملات غير المؤهِّلة (مثل الاستردادات والمعاملات الملغاة)."));

kids.push(article("4", "آلية الاستبدال (Burn)"));
kids.push(item("يجوز للعميل استبدال نقاطه لدى التاجر وفق قيمة النقطة المحدّدة في الملحق (أ)."));
kids.push(item("تُخصم قيمة الاستبدال من مستحقات التاجر ضمن التسوية، ما لم يُتفق على تمويل مشترك في الملحق (أ)."));
kids.push(item("يُطبَّق حدّ أدنى للاستبدال إن وُرد في الملحق (أ)."));

kids.push(article("5", "العروض الحصرية"));
kids.push(item("يلتزم التاجر بتقديم العروض المتفق عليها في الملحق (أ) حصرياً لعملاء تطبيق درب."));
kids.push(item("يُعتمد كل عرض من الطرفين كتابياً قبل نشره بمدة لا تقل عن [7] أيام."));

kids.push(article("6", "التسوية المالية والفوترة"));
kids.push(item("تُحتسب مستحقات كل طرف (مساهمة الكسب مقابل قيمة الاستبدال) وتُسوّى بشكل [أسبوعي/شهري] حسب الملحق (أ)."));
kids.push(item("تُصدر درب كشفاً تفصيلياً بالمعاملات، وللتاجر حق الاعتراض خلال [7] أيام من استلامه."));
kids.push(item("تُضاف ضريبة القيمة المضافة وفق النظام."));

kids.push(article("7", "التكامل التقني والتشغيل"));
kids.push(item("يمكّن التاجر درب من ربط أنظمته (نقاط البيع/QR) بالبرنامج، وتوفّر درب الدعم الفني اللازم."));
kids.push(item("يلتزم التاجر بتدريب موظفيه على آلية الكسب والاستبدال."));

kids.push(article("8", "البيانات وحمايتها"));
kids.push(item("تظل بيانات معاملات العملاء عبر البرنامج مملوكةً لدرب."));
kids.push(item("يلتزم الطرفان بحماية البيانات وسرّيتها وفق نظام حماية البيانات الشخصية (PDPL) والأنظمة ذات العلاقة."));

kids.push(article("9", "التسويق والعلامات"));
kids.push(item("يجوز إطلاق حملات تسويقية مشتركة، وتقاسم تكاليف العروض (Co-funding) وفق الملحق (أ)."));
kids.push(item("لا يجوز لأي طرف استخدام علامة الطرف الآخر إلا بموافقة كتابية مسبقة."));

kids.push(article("10", "الحصرية"));
kids.push(item("[اختياري] يمنح التاجر درب حصرية في فئة [   ] ضمن مواقع درب طوال مدة الاتفاقية، وفق ما يُحدَّد في الملحق (أ)."));

kids.push(article("11", "مؤشرات الأداء والمراجعة"));
kids.push(item("يراجع الطرفان أداء البرنامج دورياً وفق المؤشرات المحدّدة في الملحق (أ)."));

kids.push(article("12", "المدة والتجديد والإنهاء"));
kids.push(item("مدة هذه الاتفاقية [   ] من تاريخ توقيعها، وتُجدّد تلقائياً لمدد مماثلة ما لم يُشعِر أحد الطرفين الآخر كتابياً برغبته في عدم التجديد قبل [30] يوماً."));
kids.push(item("لأي طرف إنهاء الاتفاقية بإشعار كتابي مسبق بمدة [30] يوماً، مع تسوية المستحقات القائمة."));

kids.push(article("13", "المسؤولية والتعويض والسرّية"));
kids.push(item("يتحمّل كل طرف مسؤولية إخلاله بالتزاماته، وتبقى المعلومات المتبادلة سرّية أثناء الاتفاقية ولمدة [سنتين] بعد انتهائها."));

kids.push(article("14", "النظام واجب التطبيق والاختصاص"));
kids.push(item("تخضع هذه الاتفاقية لأنظمة المملكة العربية السعودية، وتختص محاكم [   ] بأي نزاع يتعذّر حلّه ودياً خلال [30] يوماً."));

kids.push(article("15", "أحكام عامة"));
kids.push(item("لا يُعدّل هذا الاتفاق إلا بملحق كتابي موقّع من الطرفين."));
kids.push(item("تُوجّه الإشعارات كتابياً على العناوين المدوّنة، ولا يُعدّ تساهل أي طرف تنازلاً عن حقوقه."));
kids.push(item("حُرّرت من نسختين، بيد كل طرف نسخة للعمل بموجبها."));

// ===== SIGNATURES =====
kids.push(new Paragraph({ spacing: { before: 260, after: 100 }, alignment: AlignmentType.RIGHT, bidirectional: true, children: [R("التوقيعات", { bold: true, size: 23, color: ORANGE })] }));
kids.push(new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [4680, 4680], visuallyRightToLeft: true,
  borders: { top: { style: BorderStyle.SINGLE, size: 4, color: LINE }, bottom: { style: BorderStyle.SINGLE, size: 4, color: LINE }, left: { style: BorderStyle.SINGLE, size: 4, color: LINE }, right: { style: BorderStyle.SINGLE, size: 4, color: LINE }, insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: LINE }, insideVertical: { style: BorderStyle.SINGLE, size: 4, color: LINE } },
  rows: [
    new TableRow({ children: [ cell("الطرف الأول (درب)", 4680, { bold: true, align: AlignmentType.CENTER, shade: HEAD_BG }), cell("الطرف الثاني (التاجر)", 4680, { bold: true, align: AlignmentType.CENTER, shade: HEAD_BG }) ] }),
    new TableRow({ children: [ cell("الاسم: ..........................\nالصفة: ..........................\nالتوقيع: ..........................\nالتاريخ:      /      /      ", 4680), cell("الاسم: ..........................\nالصفة: ..........................\nالتوقيع: ..........................\nالتاريخ:      /      /      ", 4680) ] }),
  ] }));

// ===== SCHEDULE A =====
kids.push(new Paragraph({ children: [ new (require("docx").PageBreak)() ] }));
kids.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, bidirectional: true, children: [R("الملحق (أ) — الجدول التجاري", { bold: true, size: 28, color: ORANGE })] }));
kids.push(P([R("يُعبّأ هذا الجدول لكل تاجر على حدة، ويُعدّ جزءاً لا يتجزأ من الاتفاقية.", { size: 20, color: GREY, italics: true })], { after: 140, align: AlignmentType.CENTER }));

const SW = [3400, 5960];
const schedRows = [
  ["البند", "القيمة المتفق عليها"],
  ["اسم التاجر / الفئة", "……………………………………"],
  ["الفروع المشمولة", "……………………………………"],
  ["نسبة الكسب للعميل (Earn)", "…… % من قيمة الشراء"],
  ["مساهمة التاجر في تمويل الكسب", "…… هللة عن كل ريال"],
  ["قيمة النقطة عند الاستبدال (Burn)", "…… نقطة = 1 ريال"],
  ["من يتحمّل قيمة الاستبدال", "☐ التاجر   ☐ مناصفة   ☐ درب"],
  ["الحد الأدنى للاستبدال", "…… نقطة"],
  ["العروض الحصرية لعملاء التطبيق", "……………………………………"],
  ["دورية التسوية", "☐ أسبوعي   ☐ شهري"],
  ["تمويل الحملات (Co-funding)", "……………………………………"],
  ["الحصرية (الفئة)", "☐ نعم — فئة ……   ☐ لا"],
  ["مؤشرات الأداء المتفق عليها", "……………………………………"],
  ["مدة الاتفاقية", "……………………………………"],
];
kids.push(new Table({ width: { size: SW[0] + SW[1], type: WidthType.DXA }, columnWidths: SW, visuallyRightToLeft: true,
  borders: { top: { style: BorderStyle.SINGLE, size: 4, color: LINE }, bottom: { style: BorderStyle.SINGLE, size: 4, color: LINE }, left: { style: BorderStyle.SINGLE, size: 4, color: LINE }, right: { style: BorderStyle.SINGLE, size: 4, color: LINE }, insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: LINE }, insideVertical: { style: BorderStyle.SINGLE, size: 4, color: LINE } },
  rows: schedRows.map((r, i) => new TableRow({ tableHeader: i === 0, children: [
    cell(r[0], SW[0], { bold: true, shade: i === 0 ? ORANGE_L : HEAD_BG, align: i === 0 ? AlignmentType.CENTER : AlignmentType.RIGHT }),
    cell(r[1], SW[1], { shade: i === 0 ? ORANGE_L : undefined, bold: i === 0, align: i === 0 ? AlignmentType.CENTER : AlignmentType.RIGHT }),
  ] })) }));

kids.push(P([R("توقيع الطرفين على الملحق (أ):  درب: ..................    التاجر: ..................", { size: 20 })], { before: 160 }));
kids.push(P([R("قالب داخلي — تُملأ الأقواس [   ] والجدول لكل تاجر. يُراجع من الإدارة القانونية قبل الاعتماد.", { italics: true, size: 18, color: GREY })], { before: 200, align: AlignmentType.CENTER }));

const doc = new Document({
  creator: "Darb", title: "اتفاقية الانضمام لبرنامج الولاء تانكي - قالب المستأجرين",
  styles: { default: { document: { run: { font: FONT, size: 22, color: INK } } },
    paragraphStyles: [ { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: FONT, bold: true, size: 23, color: INK } } ] },
  sections: [ { properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } }, children: kids } ],
});
Packer.toBuffer(doc).then((b) => { fs.writeFileSync(process.argv[2] || "loyalty.docx", b); console.log("written", b.length); });
