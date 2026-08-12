const pptx = require("pptxgenjs");
const p = new pptx();
p.layout = "LAYOUT_WIDE";           // 13.3 × 7.5
p.rtlMode = true;

const NAVY = "1E2761", ICE = "CADCFC", W = "FFFFFF", INK = "18202F",
      MUT = "6B7A90", RED = "B3261E", GRN = "2C6E49", GOLD = "C79A2E",
      LGT = "F4F6FA";
const HF = "Cambria", BF = "Arial";
const SW = 13.3, SH = 7.5, M = 0.7;

const rtl = { align: "right", rtlMode: true };

function dark(t, sub, kicker) {
  const s = p.addSlide();
  s.background = { color: NAVY };
  if (kicker) s.addText(kicker, { x: M, y: 2.15, w: SW - 2 * M, h: 0.35,
    fontFace: BF, fontSize: 13, color: ICE, charSpacing: 3, ...rtl });
  s.addText(t, { x: M, y: 2.55, w: SW - 2 * M, h: 1.2, fontFace: HF,
    fontSize: 42, bold: true, color: W, ...rtl });
  if (sub) s.addText(sub, { x: M, y: 3.85, w: SW - 2 * M, h: 0.9,
    fontFace: BF, fontSize: 16, color: ICE, ...rtl });
  return s;
}

function page(t, sub) {
  const s = p.addSlide();
  s.background = { color: W };
  s.addText(t, { x: M, y: 0.45, w: SW - 2 * M, h: 0.6, fontFace: HF,
    fontSize: 32, bold: true, color: NAVY, ...rtl });
  if (sub) s.addText(sub, { x: M, y: 1.05, w: SW - 2 * M, h: 0.38,
    fontFace: BF, fontSize: 13, color: MUT, ...rtl });
  return s;
}

function card(s, x, y, w, h, fill) {
  s.addShape(p.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.06,
    fill: { color: fill || LGT }, line: { color: "E3E8F0", width: 1 },
    shadow: { type: "outer", color: "9AA7BC", blur: 8, offset: 1, angle: 90, opacity: 0.18 } });
}

function stat(s, x, y, w, num, lbl, col) {
  s.addText(num, { x, y, w, h: 0.72, fontFace: HF, fontSize: 38, bold: true,
    color: col || NAVY, align: "center", margin: 0 });
  s.addText(lbl, { x, y: y + 0.72, w, h: 0.5, fontFace: BF, fontSize: 11,
    color: MUT, align: "center", margin: 0 });
}

/* ══ 1 · الغلاف ══ */
{
  const s = dark("الخطة التجارية", "ثلاثة منافذ بيع · أربع قنوات · نماذج تنفيذ مبنية على بيانات الشبكة",
                 "الإدارة التجارية — درب");
  s.addText("عرض على الإدارة التنفيذية   ·   النصف الأول 2026", { x: M, y: 6.3,
    w: SW - 2 * M, h: 0.4, fontFace: BF, fontSize: 12, color: MUT, ...rtl });
  s.addNotes("الخطة مبنية على بيانات 55 محطة للنصف الأول 2026، وقوائم دخل فعلية، وبيانات توريد وعمالة وشكاوى.");
}

/* ══ 2 · الوضع ══ */
{
  const s = page("أين نقف", "بيانات 55 محطة · 1 يناير – 30 يونيو 2026");
  const d = [
    ["1.27", "مليار ريال\nإيراد سنوي مُطبَّع", NAVY],
    ["588", "مليون لتر\nحجم سنوي", NAVY],
    ["−43.8", "مليون ريال\nصافي الميزان", RED],
    ["25%", "إشغال الوحدات\n96 وحدة شاغرة", RED],
  ];
  d.forEach((v, i) => {
    const x = M + i * 3.05;
    card(s, x, 1.75, 2.85, 1.85);
    stat(s, x, 1.95, 2.85, v[0], v[1], v[2]);
  });
  s.addText("ثلاث حقائق تحكم الخطة", { x: M, y: 4.0, w: SW - 2 * M, h: 0.4,
    fontFace: HF, fontSize: 19, bold: true, color: NAVY, ...rtl });
  const f = [
    "سعر الوقود مقنَّن — فالربح يأتي من الحجم والخدمة لا من السعر",
    "المصاريف النقدية 0.94 هللة لكل لتر — لا مجال للتحسين عبر خفض التكلفة",
    "التمويل والإهلاك على الوحدات مدفوعان — فالوحدة الشاغرة تكلفة لا فرصة ضائعة",
  ];
  f.forEach((t, i) => {
    const y = 4.5 + i * 0.62;
    s.addShape(p.ShapeType.ellipse, { x: SW - M - 0.34, y: y + 0.06, w: 0.24, h: 0.24,
      fill: { color: NAVY } });
    s.addText(t, { x: M, y, w: SW - 2 * M - 0.55, h: 0.5, fontFace: BF,
      fontSize: 14, color: INK, margin: 0, ...rtl });
  });
  s.addNotes("الميزان السالب محسوب على 41 محطة لها بيانات نمو: 13 هابطة تفوق أثر 17 صاعدة.");
}

/* ══ 3 · الإطار ══ */
{
  const s = page("الإطار", "ثلاثة منافذ — والوقود قناتان مختلفتان تماماً");
  const cols = [
    { t: "الوقود · أفراد", m: "كل مشغّلي المحطات", c: NAVY,
      p: ["تدريب العمالة على سرعة الخدمة", "مستهدف لكل محطة", "حافز لكل عامل", "خفض مصاريف المحطة"] },
    { t: "الوقود · شركات", m: "بترو آب · واعي · المنصات", c: "2A3F7A",
      p: ["مندوب أساطيل الحج والعمرة", "خصم هللة لكل لتر", "الرفع على سيارة آب", "الاشتراك في كل منصة"] },
    { t: "العقار والتأجير", m: "الوسطاء العقاريون", c: GRN,
      p: ["شبكة وسطاء بعمولة", "مغسلة وسوبرماركت مع كل افتتاح", "حصر الوحدات وتصنيفها"] },
    { t: "الإكسسوارات", m: "أصحاب محلات التجزئة", c: GOLD,
      p: ["التعاقد مع التجزئة", "مساحة بيع ومساحة في التطبيق", "نسبة من المبيعات"] },
  ];
  cols.forEach((c, i) => {
    const x = M + i * 3.05, w = 2.85;
    card(s, x, 1.7, w, 4.6, W);
    s.addShape(p.ShapeType.rect, { x, y: 1.7, w, h: 0.62, fill: { color: c.c } });
    s.addText(c.t, { x: x + 0.12, y: 1.78, w: w - 0.24, h: 0.46, fontFace: HF,
      fontSize: 15, bold: true, color: W, align: "center", margin: 0 });
    s.addText("المسيطر: " + c.m, { x: x + 0.15, y: 2.45, w: w - 0.3, h: 0.55,
      fontFace: BF, fontSize: 11, color: MUT, ...rtl });
    s.addText(c.p.map((t, k) => ({ text: t, options: { bullet: true, breakLine: k < c.p.length - 1 } })),
      { x: x + 0.15, y: 3.05, w: w - 0.3, h: 3.1, fontFace: BF, fontSize: 12,
        color: INK, paraSpaceAfter: 8, ...rtl });
  });
  s.addNotes("في العقار والإكسسوارات لا سلسلة مسيطرة — وهما أضعف منافذنا وأقلّها منافسة.");
}

/* ══ 4 · المنافسون ══ */
{
  const s = page("المنافسون حسب المنفذ", "من يسيطر · وبماذا · وردّنا");
  const rows = [
    ["الدريس", "وقود · أفراد", "1,383 محطة · 18.95%", "الانتشار وسرعة الضمّ", "المنافسة على الخدمة لا السعر"],
    ["ساسكو", "وقود · أفراد", "~683 محطة", "النمو بالاستحواذ", "تحسين تجربة المحطة"],
    ["بترو آب", "وقود · شركات", "500 ألف مركبة", "ملكية العلاقة بلا أصول", "اشتراك كنقطة قبول + تعاقد مباشر"],
    ["واعي — الدريس", "وقود · شركات", "شبكة الدريس", "شريحة في فوهة المضخة", "خصم هللة يفوق الميزة التقنية"],
    ["سيارة آب", "وقود · شركات", "2,500+ محطة قبول", "منصة وشبكة معاً", "شراكة قائمة — رفع أساطيلنا"],
    ["الوسطاء العقاريون", "عقار", "سوق مجزّأ", "قائمة المستأجرين", "التحالف لا المنافسة"],
    ["محلات التجزئة", "إكسسوارات", "سوق مجزّأ", "المخزون والخبرة", "التعاقد — مساحة مقابل نسبة"],
  ];
  const head = ["المنافس", "المنفذ", "الحجم", "مصدر السيطرة", "الردّ"];
  const tbl = [head.map(h => ({ text: h, options: { bold: true, color: W, fill: { color: NAVY },
    fontFace: BF, fontSize: 12, align: "center", valign: "middle" } }))];
  rows.forEach((r, i) => tbl.push(r.map((c, j) => ({ text: c, options: {
    fontFace: BF, fontSize: 11.5, color: INK, align: j === 0 ? "right" : (j <= 2 ? "center" : "right"),
    valign: "middle", fill: { color: i % 2 ? LGT : W }, bold: j === 0 } }))));
  s.addTable(tbl, { x: M, y: 1.7, w: SW - 2 * M, colW: [2.3, 1.7, 2.2, 2.6, 3.1],
    rowH: 0.52, border: { type: "solid", color: "E3E8F0", pt: 1 }, rtlMode: true });
  s.addText("في قناة الأفراد المنافس مشغّل محطات والسلاح تجربة الخدمة · وفي الشركات المنافس منصة والسلاح السعر والعقد",
    { x: M, y: 6.35, w: SW - 2 * M, h: 0.5, fontFace: BF, fontSize: 13,
      color: NAVY, bold: true, ...rtl });
}

/* ══ 5 · الوقود أفراد ══ */
{
  const s = page("الوقود · قناة الأفراد", "العميل يبحث عن سرعة الخدمة — والأداة مستهدف وحافز");
  card(s, M, 1.7, 6.0, 2.5, LGT);
  s.addText("ما تقوله الشكاوى", { x: M + 0.25, y: 1.9, w: 5.5, h: 0.4,
    fontFace: HF, fontSize: 18, bold: true, color: NAVY, ...rtl });
  [["56%", "من الشكاوى عن العامل مباشرة"],
   ["74%", "من شكاوى أكبر محطة: بطء وسلوك"],
   ["صفر", "استبدال أو إيقاف من 58 شكوى"]].forEach((v, i) => {
    const y = 2.4 + i * 0.55;
    s.addText(v[0], { x: M + 5.3, y, w: 1.15, h: 0.42, fontFace: HF,
      fontSize: 20, bold: true, color: RED, align: "right", margin: 0 });
    s.addText(v[1], { x: M + 0.25, y: y + 0.05, w: 4.9, h: 0.42, fontFace: BF,
      fontSize: 13, color: INK, margin: 0, ...rtl });
  });
  card(s, M + 6.35, 1.7, SW - 2 * M - 6.35, 2.5, W);
  s.addText("أثر خطة الحافز", { x: M + 6.6, y: 1.9, w: 4.6, h: 0.4,
    fontFace: HF, fontSize: 18, bold: true, color: NAVY, ...rtl });
  [["نمو مستهدف", "5%"], ["لترات إضافية", "+23.9 مليون"],
   ["الحافز الموزّع", "480 ألف ريال"], ["الربح الصافي للشركة", "≈ 2.7 مليون ريال"]].forEach((v, i) => {
    const y = 2.4 + i * 0.42;
    s.addText(v[0], { x: M + 6.6, y, w: 2.6, h: 0.38, fontFace: BF, fontSize: 12.5,
      color: MUT, margin: 0, ...rtl });
    s.addText(v[1], { x: M + 6.6, y, w: 4.6, h: 0.38, fontFace: BF, fontSize: 13,
      bold: true, color: i === 3 ? GRN : INK, align: "left", margin: 0 });
  });
  s.addText("آلية الحافز — كما اعتُمدت", { x: M, y: 4.45, w: SW - 2 * M, h: 0.4, fontFace: HF,
    fontSize: 19, bold: true, color: NAVY, ...rtl });
  ["15% من هامش اللتر الإضافي بسقف 100 ريال شهرياً لكل عامل — لا مبلغ ثابت",
   "خط الأساس يُشتق من حصة وردية العامل من الزيارات ÷ عدد عمالها — لا يُقدَّر يدوياً",
   "بوابة جودة تحجب الحافز كاملاً: شكوى مثبتة أو زمن تعبئة أو غياب في الذروة"].forEach((t, i) => {
    const y = 4.95 + i * 0.55;
    s.addShape(p.ShapeType.ellipse, { x: SW - M - 0.32, y: y + 0.05, w: 0.22, h: 0.22,
      fill: { color: NAVY } });
    s.addText(t, { x: M, y, w: SW - 2 * M - 0.5, h: 0.45, fontFace: BF,
      fontSize: 13.5, color: INK, margin: 0, ...rtl });
  });
  s.addNotes("خفض المصاريف محدود: 88% من المصاريف إهلاك غير نقدي، والنقدي 0.94 هللة لكل لتر. " +
             "سقف الـ100 ريال لا يقيّد على مستوى الشبكة (578 ألف سقفاً مقابل 479 ألف مستحقة) — " +
             "يقيّد المتميزين أفراداً فقط.");
}

/* ══ 5d · مطابقة الوردية مع الطلب ══ */
{
  const s = page("مطابقة الوردية مع الطلب",
                 "العمالة موزّعة نصفين — والطلب ليس نصفين · بيانات 212 يوماً فعلياً");
  const rows = [
    ["MK019 عرفات الشرايع", "56.1%", "50.0%", "97", "124", "+28%"],
    ["MK007 النورية", "51.0%", "45.0%", "190", "242", "+27%"],
    ["MK017 عرفات الشوقية", "55.3%", "50.0%", "189", "234", "+24%"],
    ["MK023 بن درويش", "54.9%", "50.0%", "178", "217", "+22%"],
    ["MK002 المعيصم", "50.4%", "50.0%", "214", "217", "+1%"],
  ];
  const head = ["المحطة", "زيارات المساء", "عمالة المساء", "حِمل عامل الصباح",
                "حِمل عامل المساء", "الفارق"];
  const tbl = [head.map(h => ({ text: h, options: { bold: true, color: W, fill: { color: NAVY },
    fontFace: BF, fontSize: 11.5, align: "center", valign: "middle" } }))];
  rows.forEach((r, i) => tbl.push(r.map((c, j) => ({ text: c, options: {
    fontFace: BF, fontSize: 12, color: j === 5 ? (i === 4 ? MUT : RED) : INK,
    bold: j === 0 || j === 5, align: j === 0 ? "right" : "center", valign: "middle",
    fill: { color: i % 2 ? LGT : W } } }))));
  s.addTable(tbl, { x: M, y: 1.65, w: SW - 2 * M, colW: [3.0, 1.9, 1.9, 2.0, 2.0, 1.1],
    rowH: 0.45, border: { type: "solid", color: "E3E8F0", pt: 1 }, rtlMode: true });

  card(s, M, 4.5, 5.9, 2.0, LGT);
  s.addText("ما يعنيه هذا للحافز", { x: M + 0.28, y: 4.62, w: 5.3, h: 0.35,
    fontFace: HF, fontSize: 17, bold: true, color: NAVY, ...rtl });
  s.addText([{ text: "خط أساس موحّد يظلم عامل المساء ويكافئ عامل الصباح على جهد أقل", options: { bullet: true, breakLine: true } },
             { text: "لذلك يُشتق خط الأساس من حصة الوردية — لا من متوسط المحطة", options: { bullet: true, breakLine: true } },
             { text: "الاشتقاق يجعل 5% نمو تعني الجهد نفسه على الورديتين", options: { bullet: true } }],
    { x: M + 0.28, y: 5.02, w: 5.3, h: 1.4, fontFace: BF, fontSize: 12.5,
      color: INK, paraSpaceAfter: 5, ...rtl });

  card(s, M + 6.25, 4.5, SW - 2 * M - 6.25, 2.0, W);
  s.addText("ما يمكن تنفيذه فعلاً", { x: M + 6.53, y: 4.62, w: 5.3, h: 0.35,
    fontFace: HF, fontSize: 17, bold: true, color: NAVY, ...rtl });
  s.addText([{ text: "النورية وحدها فيها عمالة تكفي لنقل عامل كامل (20 عاملاً)", options: { bullet: true, breakLine: true } },
             { text: "البقية فجوتها أقل من عامل — الحل مواعيد الدخول لا العدد", options: { bullet: true, breakLine: true } },
             { text: "تكديس الدخول بين 8 و10 مساءً يغطي الذروة بلا تكلفة", options: { bullet: true } }],
    { x: M + 6.53, y: 5.02, w: 5.3, h: 1.4, fontFace: BF, fontSize: 12.5,
      color: INK, paraSpaceAfter: 5, ...rtl });

  s.addText("ذروة الزيارات 21:00 و22:00 في أربع محطات — والعمالة تُوزَّع كأن الطلب متساوٍ",
    { x: M, y: 6.7, w: SW - 2 * M, h: 0.4, fontFace: BF, fontSize: 13,
      bold: true, color: NAVY, ...rtl });
  s.addNotes("المعيصم وحدها متوازنة لأن ذروتها 5 إلى 7 مساءً لا بعد التاسعة. " +
             "الوردية الصباحية 6 ص إلى 6 م والمسائية 6 م إلى 6 ص، والحصص من عدد الزيارات الفعلي.");
}


/* ══ 5b · تشخيص فقدان الحركة ══ */
{
  const s = page("تشخيص فقدان الحركة — مكة", "بيانات يناير – 4 أغسطس 2026 · السبب مؤكَّد ميدانياً");
  const rows = [
    ["MK002 المعيصم", "2,163", "601", "−72%", "20.5 ← 20.9", "تغيّر مسار الطريق", RED],
    ["MK019 الشرايع", "830", "380", "−54%", "22.1 ← 23.2", "منافس فتح قبلنا", RED],
    ["MK072 العزيزية", "2,778", "2,488", "−10%", "20.3", "موسمي — تعافت", GOLD],
    ["MK007 العمرة", "4,322", "4,266", "−1%", "31.7 ← 29.0", "موسمي — عادت لمستواها", GRN],
    ["MK017 عرفات", "1,210", "1,391", "+15%", "20.4 ← 20.8", "نمو — المرجع", GRN],
  ];
  const head = ["المحطة", "زيارات يناير", "يونيو–يوليو", "التغيّر", "لتر/زيارة", "السبب"];
  const tbl = [head.map(h => ({ text: h, options: { bold: true, color: W, fill: { color: NAVY },
    fontFace: BF, fontSize: 12, align: "center", valign: "middle" } }))];
  rows.forEach((r, i) => tbl.push(r.slice(0, 6).map((c, j) => ({ text: c, options: {
    fontFace: BF, fontSize: 12, color: j === 3 ? r[6] : INK, bold: j === 0 || j === 3,
    align: j === 0 || j === 5 ? "right" : "center", valign: "middle",
    fill: { color: i % 2 ? LGT : W } } }))));
  s.addTable(tbl, { x: M, y: 1.65, w: SW - 2 * M, colW: [2.5, 1.8, 1.8, 1.4, 1.9, 2.5],
    rowH: 0.5, border: { type: "solid", color: "E3E8F0", pt: 1 }, rtlMode: true });

  card(s, M, 4.9, 5.9, 1.9, LGT);
  s.addText("ما تثبته البيانات", { x: M + 0.28, y: 5.02, w: 5.3, h: 0.35,
    fontFace: HF, fontSize: 17, bold: true, color: NAVY, ...rtl });
  s.addText([{ text: "سلة العميل ثابتة — الفقد في عدد العملاء لا سلوكهم", options: { bullet: true, breakLine: true } },
             { text: "سقوط 60% و53% في شهر واحد = حدث خارجي بتاريخ", options: { bullet: true, breakLine: true } },
             { text: "الخدمة والمضخات والتسعير — كلها مستبعَدة", options: { bullet: true } }],
    { x: M + 0.28, y: 5.42, w: 5.3, h: 1.3, fontFace: BF, fontSize: 12.5,
      color: INK, paraSpaceAfter: 5, ...rtl });

  card(s, M + 6.25, 4.9, SW - 2 * M - 6.25, 1.9, W);
  s.addText("الأثر", { x: M + 6.5, y: 5.02, w: 5.3, h: 0.35, fontFace: HF,
    fontSize: 17, bold: true, color: NAVY, ...rtl });
  [["2,013", "زيارة مفقودة يومياً", RED], ["15.7", "مليون لتر سنوياً", RED],
   ["2.1", "مليون ريال هامش سنوياً", RED]].forEach((v, i) => {
    stat(s, M + 6.5 + i * 1.8, 5.42, 1.7, v[0], v[1], v[2]);
  });
  s.addNotes("الحدثان اكتُشفا بعد شهور من وقوعهما — بتحليل لاحق لا بنظام إنذار. هذه الفجوة أهم من الحدثين نفسيهما.");
}

/* ══ 5c · الردّ ══ */
{
  const s = page("الردّ على فقدان الحركة", "كل حالة تستدعي ردّاً مختلفاً — والفرق بينهما جوهري");
  const cases = [
    { t: "المعيصم — تغيّر المسار", k: "فقد دائم", c: RED,
      d: ["العميل لم يعد يمرّ — ولا يُسترد بسعر ولا خدمة",
          "التحويل من ممرّ إلى وجهة: مغسلة وسوبرماركت ومطعم",
          "إعادة تحجيم العمالة والمصاريف على الحركة الجديدة"],
      kpi: "المؤشر: إيراد غير وقودي لكل زيارة" },
    { t: "الشرايع — منافس قبلنا", k: "قابل للاسترداد", c: GOLD,
      d: ["المنافس يعترض العميل قبل وصوله إلينا",
          "سبب للتجاوز: عرض أو خدمة لا توجد عنده",
          "لافتات مبكرة قبل نقطة اعتراضه · وتعاقد مع الأساطيل المارّة"],
      kpi: "المؤشر: الزيارات اليومية مقابل خط الأساس" },
    { t: "الشبكة — الدرس العام", k: "مخاطرة قائمة", c: NAVY,
      d: ["حدثان كلّفانا 2.1 مليون ريال واكتُشفا بعد شهور",
          "تنبيه آلي عند انخفاض الزيارات الأسبوعية أكثر من 15%",
          "مسح ربعي للمنافسين وأعمال الطرق حول كل محطة"],
      kpi: "المؤشر: زمن اكتشاف الحدث" },
  ];
  cases.forEach((c, i) => {
    const x = M + i * 4.05;
    card(s, x, 1.7, 3.85, 4.7, W);
    s.addShape(p.ShapeType.rect, { x, y: 1.7, w: 3.85, h: 0.58, fill: { color: c.c } });
    s.addText(c.t, { x: x + 0.1, y: 1.77, w: 3.65, h: 0.44, fontFace: HF,
      fontSize: 14, bold: true, color: W, align: "center", margin: 0 });
    s.addText(c.k, { x: x + 0.15, y: 2.42, w: 3.55, h: 0.38, fontFace: BF,
      fontSize: 13, bold: true, color: c.c, align: "center", margin: 0 });
    s.addText(c.d.map((t, k) => ({ text: t, options: { bullet: true, breakLine: k < c.d.length - 1 } })),
      { x: x + 0.2, y: 2.9, w: 3.45, h: 2.5, fontFace: BF, fontSize: 12,
        color: INK, paraSpaceAfter: 7, ...rtl });
    s.addText(c.kpi, { x: x + 0.2, y: 5.75, w: 3.45, h: 0.5, fontFace: BF,
      fontSize: 11, color: MUT, italic: true, ...rtl });
  });
  s.addText("الأولوية: نظام الإنذار المبكر قبل أي مبادرة أخرى — فما لا يُكتشف لا يُعالَج",
    { x: M, y: 6.6, w: SW - 2 * M, h: 0.4, fontFace: BF, fontSize: 13,
      bold: true, color: NAVY, ...rtl });
}

/* ══ 6 · الوقود شركات ══ */
{
  const s = page("الوقود · قناة الشركات", "مندوب أساطيل · خصم هللة لكل لتر · الاشتراك في كل منصة");
  const steps = [["١", "مندوب أساطيل", "حصر أساطيل الحج والعمرة"],
                 ["٢", "خصم هللة", "سعر تفضيلي مقابل التزام حجم"],
                 ["٣", "سيارة آب", "رفع أساطيلنا على الشراكة القائمة"],
                 ["٤", "كل منصة", "بترو آب · Fleet Plus كنقطة قبول"]];
  steps.forEach((st, i) => {
    const x = M + i * 3.05;
    card(s, x, 1.7, 2.85, 1.75, W);
    s.addShape(p.ShapeType.ellipse, { x: x + 1.2, y: 1.88, w: 0.46, h: 0.46,
      fill: { color: NAVY } });
    s.addText(st[0], { x: x + 1.2, y: 1.88, w: 0.46, h: 0.46, fontFace: HF,
      fontSize: 17, bold: true, color: W, align: "center", valign: "middle", margin: 0 });
    s.addText(st[1], { x: x + 0.12, y: 2.45, w: 2.6, h: 0.35, fontFace: HF,
      fontSize: 14.5, bold: true, color: NAVY, align: "center", margin: 0 });
    s.addText(st[2], { x: x + 0.12, y: 2.8, w: 2.6, h: 0.55, fontFace: BF,
      fontSize: 11.5, color: MUT, align: "center", margin: 0 });
  });
  card(s, M, 3.75, SW - 2 * M, 2.6, LGT);
  s.addText("اقتصاديات خصم الهللة", { x: M + 0.3, y: 3.95, w: 6, h: 0.4,
    fontFace: HF, fontSize: 19, bold: true, color: NAVY, ...rtl });
  [["13.36", "الهامش قبل الخصم\n(هللة/لتر)", NAVY],
   ["12.36", "الهامش بعد الخصم", NAVY],
   ["8%", "نمو التعادل المطلوب", GOLD],
   ["109", "مليون لتر ديزل\nسوق الأساطيل المتاح", GRN]].forEach((v, i) => {
    stat(s, M + 0.3 + i * 2.9, 4.5, 2.7, v[0], v[1], v[2]);
  });
  s.addText("الحجم يمرّ في محطاتنا اليوم بلا عقد — فالتعاقد تثبيت لا اكتساب",
    { x: M, y: 6.5, w: SW - 2 * M, h: 0.4, fontFace: BF, fontSize: 13,
      bold: true, color: NAVY, ...rtl });
  s.addNotes("المعيار الحاكم: عمولة المنصة لكل لتر مقابل تكلفة المندوب لكل لتر — يُراجع كل ربع.");
}

/* ══ 7 · العقار ══ */
{
  const s = page("العقار والتأجير", "الوحدة الشاغرة تكلفة لا فرصة ضائعة");
  card(s, M, 1.7, 6.2, 2.9, LGT);
  s.addText("العمولة مقابل الشغور", { x: M + 0.28, y: 1.9, w: 5.6, h: 0.4,
    fontFace: HF, fontSize: 19, bold: true, color: NAVY, ...rtl });
  s.addText("مثال: إيجار سنوي 120 ألف · تكلفة ثابتة 45 ألف · شغور 8 أشهر",
    { x: M + 0.28, y: 2.3, w: 5.6, h: 0.35, fontFace: BF, fontSize: 11.5,
      color: MUT, ...rtl });
  [["110,000", "خسارة الشغور", RED], ["6,000", "عمولة الوسيط 5%", GRN],
   ["92%", "الحد المبرَّر للعمولة", NAVY]].forEach((v, i) => {
    stat(s, M + 0.28 + i * 1.95, 2.75, 1.85, v[0], v[1], v[2]);
  });
  card(s, M + 6.55, 1.7, SW - 2 * M - 6.55, 2.9, W);
  s.addText("معيار افتتاح المحطة", { x: M + 6.8, y: 1.9, w: 4.4, h: 0.4,
    fontFace: HF, fontSize: 19, bold: true, color: NAVY, ...rtl });
  [["مغسلة", "إلزامي", GRN], ["سوبرماركت", "إلزامي", GRN],
   ["قهوة", "موصى به", MUT], ["خدمات سيارات", "موصى به", MUT]].forEach((v, i) => {
    const y = 2.4 + i * 0.5;
    s.addText(v[0], { x: M + 6.8, y, w: 2.4, h: 0.4, fontFace: BF, fontSize: 13.5,
      color: INK, bold: i < 2, margin: 0, ...rtl });
    s.addText(v[1], { x: M + 6.8, y, w: 4.4, h: 0.4, fontFace: BF, fontSize: 12,
      color: v[2], bold: i < 2, align: "left", margin: 0 });
  });
  s.addText("لماذا الوسطاء", { x: M, y: 4.85, w: SW - 2 * M, h: 0.4, fontFace: HF,
    fontSize: 19, bold: true, color: NAVY, ...rtl });
  ["الوسيط يملك قائمة المستأجرين الباحثين — ونحن لا نملكها",
   "العمولة تكلفة متغيرة تُدفع عند الإشغال فقط",
   "52 محطة تحت الإنشاء — التعاقد قبل الافتتاح أرخص من ملء محطة عاملة"].forEach((t, i) => {
    const y = 5.35 + i * 0.52;
    s.addShape(p.ShapeType.ellipse, { x: SW - M - 0.32, y: y + 0.05, w: 0.22, h: 0.22,
      fill: { color: GRN } });
    s.addText(t, { x: M, y, w: SW - 2 * M - 0.5, h: 0.42, fontFace: BF,
      fontSize: 13.5, color: INK, margin: 0, ...rtl });
  });
  s.addNotes("التمويل وإهلاك حق الاستخدام يلتهمان 63% من دخل الإيجار في العمرة و106% في الشرايع — وهما ثابتان.");
}

/* ══ 8 · الإكسسوارات ══ */
{
  const s = page("الإكسسوارات", "التعاقد مع التجزئة — مساحة بيع ومساحة في التطبيق مقابل نسبة");
  const b = [
    { t: "لا تشغيل ذاتي", d: "المخزون البطيء والمرتجعات على الشريك — والتجزئة تملك الخبرة والمورّد" },
    { t: "مساحتان لا واحدة", d: "مساحة بيع داخل المحطة · ومساحة في التطبيق تُوسّع سوق الشريك خارجها" },
    { t: "نسبة لا إيجار ثابت", d: "حد أدنى مضمون يحمي عائدنا · والنسبة ترفع سقفه مع نمو الشريك" },
    { t: "شرط البيانات", d: "مشاركة بيانات المبيعات شرط أساسي — بدونها لا يمكن التحقق من النسبة" },
  ];
  b.forEach((c, i) => {
    const x = M + (i % 2) * 6.15, y = 1.75 + Math.floor(i / 2) * 2.3;
    card(s, x, y, 5.85, 2.0, i < 2 ? LGT : W);
    s.addShape(p.ShapeType.ellipse, { x: x + 5.85 - 0.72, y: y + 0.28, w: 0.44, h: 0.44,
      fill: { color: GOLD } });
    s.addText(String(i + 1), { x: x + 5.85 - 0.72, y: y + 0.28, w: 0.44, h: 0.44,
      fontFace: HF, fontSize: 16, bold: true, color: W, align: "center",
      valign: "middle", margin: 0 });
    s.addText(c.t, { x: x + 0.28, y: y + 0.3, w: 4.8, h: 0.42, fontFace: HF,
      fontSize: 17, bold: true, color: NAVY, ...rtl });
    s.addText(c.d, { x: x + 0.28, y: y + 0.82, w: 5.3, h: 1.0, fontFace: BF,
      fontSize: 13, color: INK, ...rtl });
  });
  s.addText("السوق مجزّأ بلا سلسلة مسيطرة — وهذا المنفذ لا يحتاج استثماراً بل ملء وحدات شاغرة قائمة",
    { x: M, y: 6.5, w: SW - 2 * M, h: 0.4, fontFace: BF, fontSize: 13,
      bold: true, color: NAVY, ...rtl });
}

/* ══ 9 · التجريبي ══ */
{
  const s = page("النموذج التجريبي", "ثلاث محطات · قياس قبل التعميم · متوسط 212 يوماً");
  const st = [
    { n: "عرفات الشوقية", c: "MK017", k: "الوقود — المرجع", g: "+6.9%", col: GRN,
      d: ["1,269 زيارة يومياً", "9 مضخات · 3+3 على المضخات", "ماذا تفعل المحطة الناجحة"] },
    { n: "بن درويش", c: "MK023", k: "الوقود — الهدف", g: "−12.1%", col: RED,
      d: ["790 زيارة يومياً", "4 مضخات · 2+2 على المضخات", "198 زيارة لكل مضخة"] },
    { n: "النورية", c: "MK007", k: "المستأجرون", g: "أعلى حركة", col: NAVY,
      d: ["4,261 زيارة يومياً", "إيجار 3,887 ألف ريال", "حملات وقياس أثرها"] },
  ];
  st.forEach((v, i) => {
    const x = M + i * 4.05;
    card(s, x, 1.75, 3.85, 3.5, W);
    s.addShape(p.ShapeType.rect, { x, y: 1.75, w: 3.85, h: 0.55, fill: { color: v.col } });
    s.addText(v.k, { x: x + 0.12, y: 1.82, w: 3.6, h: 0.4, fontFace: BF, fontSize: 12,
      bold: true, color: W, align: "center", margin: 0 });
    s.addText(v.n, { x: x + 0.15, y: 2.45, w: 3.55, h: 0.42, fontFace: HF,
      fontSize: 19, bold: true, color: NAVY, align: "center", margin: 0 });
    s.addText(v.c, { x: x + 0.15, y: 2.87, w: 3.55, h: 0.3, fontFace: BF,
      fontSize: 11, color: MUT, align: "center", margin: 0 });
    s.addText(v.g, { x: x + 0.15, y: 3.2, w: 3.55, h: 0.5, fontFace: HF,
      fontSize: 24, bold: true, color: v.col, align: "center", margin: 0 });
    s.addText(v.d.map((t, k) => ({ text: t, options: { bullet: true, breakLine: k < v.d.length - 1 } })),
      { x: x + 0.22, y: 3.85, w: 3.4, h: 1.25, fontFace: BF, fontSize: 12,
        color: INK, paraSpaceAfter: 6, ...rtl });
  });
  card(s, M, 5.5, SW - 2 * M, 1.35, LGT);
  s.addText("لماذا هذا التصميم", { x: M + 0.28, y: 5.62, w: 5, h: 0.35,
    fontFace: HF, fontSize: 16, bold: true, color: NAVY, ...rtl });
  s.addText("محطتا الوقود في نفس المنطقة والنمط والمزيج وطريقة الدفع — فالفارق في النمو لا يُفسَّر بالموقع ولا المنتج ولا نوع العميل، بل بما نستطيع تغييره: الخدمة والطاقة.",
    { x: M + 0.28, y: 5.98, w: SW - 2 * M - 0.56, h: 0.75, fontFace: BF,
      fontSize: 13, color: INK, ...rtl });
}

/* ══ 10 · القرارات ══ */
{
  const s = p.addSlide();
  s.background = { color: NAVY };
  s.addText("المطلوب من الإدارة التنفيذية", { x: M, y: 0.6, w: SW - 2 * M, h: 0.7,
    fontFace: HF, fontSize: 34, bold: true, color: W, ...rtl });
  s.addText("قرار واحد اعتُمد — وستة ما زالت مطلوبة", { x: M, y: 1.28, w: SW - 2 * M,
    h: 0.32, fontFace: BF, fontSize: 12, color: ICE, ...rtl });
  const dec = [
    ["✓", "الحافز — معتمد", "15% من هامش اللتر الإضافي بسقف 100 ريال شهرياً · نمو مستهدف 5%", true],
    ["١", "خصم الهللة", "الحد الأدنى للحجم الذي يستحق الخصم وآلية الائتمان والفوترة"],
    ["٢", "تعيين مندوب الأساطيل", "وظيفة مخصصة لأساطيل الحج والعمرة"],
    ["٣", "عمولة الوسطاء", "نسبة العمولة وسقفها — الحد المبرَّر يصل 92% من الإيجار السنوي"],
    ["٤", "معيار الافتتاح", "إلزام المغسلة والسوبرماركت مع كل محطة جديدة — 52 محطة تحت الإنشاء"],
    ["٥", "إعادة توزيع الورديات", "نقل عامل للمساء في النورية · وتعديل مواعيد الدخول في الباقي"],
    ["٦", "نظام الإنذار المبكر", "تنبيه أسبوعي عند تراجع الزيارات — حدثان كلّفانا 2.1 مليون ريال"],
  ];
  dec.forEach((d, i) => {
    const y = 1.72 + i * 0.75, done = d[3];
    s.addShape(p.ShapeType.roundRect, { x: M, y, w: SW - 2 * M, h: 0.64,
      rectRadius: 0.05, fill: { color: done ? "1F4D3A" : "2A3A6B" },
      line: { color: done ? "35785C" : "3C4E84", width: 1 } });
    s.addShape(p.ShapeType.ellipse, { x: SW - M - 0.66, y: y + 0.13, w: 0.38, h: 0.38,
      fill: { color: done ? "7BD4A8" : ICE } });
    s.addText(d[0], { x: SW - M - 0.66, y: y + 0.13, w: 0.38, h: 0.38, fontFace: HF,
      fontSize: 15, bold: true, color: done ? "0F3325" : NAVY, align: "center",
      valign: "middle", margin: 0 });
    s.addText(d[1], { x: SW - M - 3.4, y: y + 0.05, w: 2.55, h: 0.54, fontFace: HF,
      fontSize: 15, bold: true, color: W, align: "right", valign: "middle", margin: 0 });
    s.addText(d[2], { x: M + 0.2, y: y + 0.05, w: SW - 2 * M - 3.8, h: 0.54,
      fontFace: BF, fontSize: 11.5, color: done ? "C9EDDA" : ICE, valign: "middle",
      margin: 0, ...rtl });
  });
  s.addNotes("الحافز اعتُمد بسقف 100 ريال — والسقف لا يقيّد على مستوى الشبكة بل يقيّد المتميزين أفراداً. " +
             "إعادة توزيع الورديات بند جديد نشأ من بيانات المطابقة التشغيلية.");
}

p.writeFile({ fileName: "/home/user/khadija/docs/الخطة-التجارية-عرض-تنفيذي.pptx" })
 .then(f => console.log("saved:", f));
