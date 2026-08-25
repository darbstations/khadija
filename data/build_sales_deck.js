/* عرض الإدارة التنفيذية — تحليل المنافذ والمنافسين وخطط المبيعات
   يُشغَّل من جذر المستودع:  node data/build_sales_deck.js
   المصدر: data/sales-plan.json (يُبنى من PYTHONPATH=data python3 data/sales_plan.py) */
const fs = require("fs");
const pptx = require("pptxgenjs");
const D = JSON.parse(fs.readFileSync("data/sales-plan.json", "utf8"));

const p = new pptx();
p.layout = "LAYOUT_WIDE";                 // 13.3 × 7.5
p.rtlMode = true;
p.author = "الإدارة التجارية — درب";
p.title = "تحليل المنافذ وخطة المبيعات";

/* ── هوية درب ── */
const ORANGE = "F5831F", GOLD = "F7A94B", BGRAY = "55565A",
      INK = "3D3D3D", INK2 = "6E6A64", INK3 = "9B968E",
      BG = "F7F4EF", W = "FFFFFF", LINE = "ECE6DD", LINE2 = "E3DCD1",
      GOOD = "2E8B6F", BAD = "C0503A", BLUE = "3E6E8E",
      T_OR = "FBEEE0", T_GOLD = "FCE7C8", T_GOOD = "D6E9E1",
      T_BAD = "F2DAD4", T_NEU = "F5F2ED", T_BAND = "EDE7DE",
      D_GOOD = "1F5E4A", D_BAD = "9A3E2C";
const F = "DIN Next Arabic";
const SW = 13.3, SH = 7.5, M = 0.62, CW = SW - 2 * M;
const rtl = { align: "right", rtlMode: true };
const ctr = { align: "center", rtlMode: true };

const rtlx = (i, cwid, gap) => M + CW - cwid - i * (cwid + gap);
const ar = n => n.toLocaleString("ar-EG", { maximumFractionDigits: 0 });
const ar1 = n => n.toLocaleString("ar-EG", { minimumFractionDigits: 1, maximumFractionDigits: 1 });
const pc = n => ar1(n * 100) + "٪";
const arn = t => String(t).replace(/[0-9]/g, d => "٠١٢٣٤٥٦٧٨٩"[d]).replace(/,/g, "٬");

/* ── هياكل الشرائح ── */
function cover(kicker, title, sub, foot) {
  const s = p.addSlide();
  s.background = { color: BGRAY };
  s.addShape(p.ShapeType.rect, { x: 0, y: 0, w: SW, h: 0.11, fill: { color: ORANGE } });
  s.addText("درب", { x: SW - M - 2.2, y: 0.5, w: 2.2, h: 0.5, fontFace: F,
    fontSize: 26, bold: true, color: W, align: "right", margin: 0 });
  s.addText("DARB FUEL", { x: SW - M - 2.2, y: 1.0, w: 2.2, h: 0.28, fontFace: F,
    fontSize: 10, color: ORANGE, align: "right", charSpacing: 2, margin: 0 });
  if (kicker) s.addText(kicker, { x: M, y: 2.75, w: CW, h: 0.35, fontFace: F,
    fontSize: 13, color: GOLD, charSpacing: 3, ...rtl });
  s.addText(title, { x: M, y: 3.15, w: CW, h: 1.1, fontFace: F, fontSize: 40,
    bold: true, color: W, ...rtl });
  if (sub) s.addText(sub, { x: M, y: 4.3, w: CW, h: 0.8, fontFace: F,
    fontSize: 15, color: "D6D2CC", ...rtl });
  s.addShape(p.ShapeType.rect, { x: SW - M - 3.4, y: 5.5, w: 3.4, h: 0.035, fill: { color: ORANGE } });
  if (foot) s.addText(foot, { x: M, y: 5.7, w: CW, h: 0.4, fontFace: F,
    fontSize: 11, color: INK3, ...rtl });
  return s;
}

function divider(no, title, sub) {
  const s = p.addSlide();
  s.background = { color: BG };
  s.addShape(p.ShapeType.rect, { x: SW - M - 0.06, y: 2.4, w: 0.06, h: 2.4, fill: { color: ORANGE } });
  s.addText(no, { x: M, y: 2.35, w: CW - 0.3, h: 0.6, fontFace: F, fontSize: 15,
    bold: true, color: ORANGE, charSpacing: 3, ...rtl });
  s.addText(title, { x: M, y: 2.95, w: CW - 0.3, h: 1.0, fontFace: F, fontSize: 36,
    bold: true, color: BGRAY, ...rtl });
  if (sub) s.addText(sub, { x: M, y: 3.95, w: CW - 0.3, h: 0.7, fontFace: F,
    fontSize: 14, color: INK2, ...rtl });
  return s;
}

let PNO = 0;
function page(title, sub) {
  const s = p.addSlide();
  PNO++;
  s.background = { color: W };
  s.addShape(p.ShapeType.rect, { x: 0, y: 0, w: SW, h: 0.075, fill: { color: ORANGE } });
  s.addText(title, { x: M, y: 0.32, w: CW - 1.0, h: 0.58, fontFace: F, fontSize: 27,
    bold: true, color: BGRAY, ...rtl });
  if (sub) s.addText(sub, { x: M, y: 0.92, w: CW - 1.0, h: 0.34, fontFace: F,
    fontSize: 12, color: INK2, ...rtl });
  s.addShape(p.ShapeType.rect, { x: M, y: 1.34, w: CW, h: 0.02, fill: { color: LINE2 } });
  s.addText("درب · الإدارة التجارية", { x: M, y: 7.02, w: 4, h: 0.3, fontFace: F,
    fontSize: 9, color: INK3, align: "left", margin: 0 });
  s.addText(String(PNO), { x: SW - M - 0.6, y: 7.02, w: 0.6, h: 0.3, fontFace: F,
    fontSize: 9, color: INK3, align: "right", margin: 0 });
  return s;
}

/* جدول يُرسم يدوياً — الأعمدة من اليمين لليسار */
function table(s, x, y, w, cols, rows, opt) {
  const o = Object.assign({ hh: 0.42, rh: 0.34, fs: 11, hfs: 11, zebra: true }, opt || {});
  const tot = cols.reduce((a, c) => a + c.w, 0);
  const xs = []; let acc = 0;
  cols.forEach(c => { acc += c.w; xs.push(x + w - (acc / tot) * w); });   // يمين ← يسار
  const cw = cols.map(c => (c.w / tot) * w);
  s.addShape(p.ShapeType.rect, { x, y, w, h: o.hh, fill: { color: BGRAY } });
  cols.forEach((c, i) => s.addText(c.t, { x: xs[i], y, w: cw[i], h: o.hh,
    fontFace: F, fontSize: o.hfs, bold: true, color: W, valign: "middle",
    align: c.a || "center", rtlMode: true, margin: [0, 0.06, 0, 0.06] }));
  rows.forEach((raw, ri) => {
    // المصفوفة تحمل fill كدالة مدمجة — فتُغلَّف أولاً
    const r = Array.isArray(raw) ? { c: raw } : raw;
    const yy = y + o.hh + ri * o.rh;
    const fill = r.fill || (o.zebra && ri % 2 ? T_NEU : W);
    s.addShape(p.ShapeType.rect, { x, y: yy, w, h: o.rh,
      fill: { color: fill }, line: { color: LINE2, width: 0.6 } });
    r.c.forEach((v, i) => {
      const cell = (v && typeof v === "object" && !Array.isArray(v)) ? v : { t: v };
      s.addText(String(cell.t), { x: xs[i], y: yy, w: cw[i], h: o.rh, fontFace: F,
        fontSize: cell.fs || o.fs, bold: !!cell.b || !!r.b, color: cell.c || (r.c2 || INK),
        valign: "middle", align: cell.a || cols[i].a || "center", rtlMode: true,
        margin: [0, 0.06, 0, 0.06] });
    });
  });
  return y + o.hh + rows.length * o.rh;
}

function stat(s, x, y, w, h, num, lbl, col, note) {
  s.addShape(p.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.05,
    fill: { color: BG }, line: { color: LINE2, width: 1 } });
  s.addShape(p.ShapeType.rect, { x, y, w, h: 0.05, fill: { color: col || ORANGE } });
  s.addText(num, { x: x + 0.08, y: y + 0.18, w: w - 0.16, h: 0.66, fontFace: F,
    fontSize: 30, bold: true, color: col || BGRAY, align: "center", margin: 0 });
  s.addText(lbl, { x: x + 0.1, y: y + 0.86, w: w - 0.2, h: 0.3, fontFace: F,
    fontSize: 11, color: INK, align: "center", margin: 0 });
  if (note) s.addText(note, { x: x + 0.1, y: y + 1.18, w: w - 0.2, h: 0.32,
    fontFace: F, fontSize: 9, color: INK3, align: "center", margin: 0 });
}

function bullets(s, x, y, w, items, col, size) {
  s.addText(items.map((t, i) => ({ text: t, options: { bullet: { code: "25AA" },
    breakLine: i < items.length - 1 } })), { x, y, w, h: 0.4 + items.length * 0.4,
    fontFace: F, fontSize: size || 12.5, color: col || INK, lineSpacing: 19,
    paraSpaceAfter: 6, ...rtl });
}

function note(s, t) {
  s.addShape(p.ShapeType.rect, { x: M, y: 6.55, w: CW, h: 0.02, fill: { color: LINE2 } });
  s.addText(t, { x: M, y: 6.6, w: CW, h: 0.36, fontFace: F, fontSize: 10, color: INK3, ...rtl });
}

const N = D.network, T = D.totals, SEG = D.segments, ST = D.stations;
const days = N.days;
const netLpd = N.volume / days, netVpd = N.visits / days;

/* ═══ ١ · الغلاف ═══ */
cover("الإدارة التجارية · أغسطس ٢٠٢٦",
      "تحليل المنافذ وخطة المبيعات",
      "٥٥ محطة · خمس شرائح سوقية · مستهدف لكل محطة ووردية وعامل · حملات مرتبطة بحالة كل محطة",
      "المصادر: كاش إن وناتج (١٠٫٩٦ مليون عملية) · أوامر التحميل · قائمة الدخل يوليو ٢٠٢٦ · سجل الوحدات · مسح المنافسين الميداني");

/* ═══ ٢ · الخلاصة التنفيذية ═══ */
{
  const s = page("الخلاصة التنفيذية", "ما تقوله البيانات — وما نطلب إقراره اليوم");
  const d = [
    [ar(Math.round(N.revenue / 1e6)), "مليون ريال", "مبيعات ٧ أشهر ، ٥٥ محطة", ORANGE],
    [ar(Math.round(N.volume / 1e6)), "مليون لتر", ar(Math.round(N.visits / 1e6)) + " مليون زيارة", BGRAY],
    ["+" + ar1(17.5) + "٪", "نمو حقيقي", "على قاعدة ثابتة ٢٥ محطة", GOOD],
    [ar1(T.sar / 1e6), "مليون ريال", "فرصة سنوية من رفع السلة", GOOD],
  ];
  const cwid = (CW - 3 * 0.18) / 4;
  d.forEach((v, i) => stat(s, rtlx(i, cwid, 0.18), 1.55, cwid, 1.72, v[0], v[1], v[3], v[2]));

  s.addText("ثلاثة قرارات مطلوبة", { x: M, y: 3.55, w: CW, h: 0.42, fontFace: F,
    fontSize: 18, bold: true, color: ORANGE, ...rtl });
  const dec = [
    ["اعتماد التقسيم والمستهدف", "خمس شرائح · معيار لتر/زيارة داخل كل شريحة ، ٣٩ محطة دون معيارها",
     ar(T.upl) + " لتر/يوم"],
    ["إقفال ثغرة البيانات قبل الإنفاق", "١٨ محطة لا تسجّل وسيلة الدفع — ١٨٣ مليون ريال بلا هوية عميل",
     "٢٩٪ من المبيعات"],
    ["تمويل الحملات من المعلن لا من هامش الوقود", "علبة المناديل بهامش ١١٫٣ هللة/لتر لا تسترد كلفتها",
     "شرط الاعتماد"],
  ];
  dec.forEach((v, i) => {
    const y = 4.05 + i * 0.86;
    s.addShape(p.ShapeType.roundRect, { x: M, y, w: CW, h: 0.76, rectRadius: 0.04,
      fill: { color: i === 0 ? T_OR : BG }, line: { color: LINE2, width: 1 } });
    s.addShape(p.ShapeType.rect, { x: SW - M - 0.055, y, w: 0.055, h: 0.76, fill: { color: ORANGE } });
    s.addText(String(i + 1), { x: SW - M - 0.62, y: y + 0.16, w: 0.5, h: 0.44,
      fontFace: F, fontSize: 17, bold: true, color: ORANGE, align: "center", margin: 0 });
    s.addText(v[0], { x: M + 2.5, y: y + 0.07, w: CW - 3.15, h: 0.34, fontFace: F,
      fontSize: 14, bold: true, color: BGRAY, ...rtl });
    s.addText(v[1], { x: M + 2.5, y: y + 0.4, w: CW - 3.15, h: 0.32, fontFace: F,
      fontSize: 11, color: INK2, ...rtl });
    s.addText(v[2], { x: M + 0.12, y: y + 0.19, w: 2.3, h: 0.4, fontFace: F,
      fontSize: 13, bold: true, color: ORANGE, align: "left", margin: 0 });
  });
  note(s, "الأرقام من بيانات نقاط البيع الفعلية من ١ يناير إلى ٣١ يوليو ٢٠٢٦، لا من تقديرات.");
  s.addNotes("النمو +17.5٪ محسوب على 25 محطة بتغطية كاملة 212 يوماً — الإجمالي الخام مضلِّل لأن عدد المحطات المغطاة ارتفع من 39 إلى 55.");
}

/* ═══ ٣ · القسم الأول ═══ */
divider("القسم الأول", "أين نقف", "الشبكة بالأرقام · الاتجاه الحقيقي · من يشتري منّا");

/* ═══ ٤ · الشبكة بالأرقام ═══ */
{
  const s = page("الشبكة بالأرقام", "من ١ يناير إلى ٣١ يوليو ٢٠٢٦ · " + ar(days) + " يوم-محطة مغطّى");
  const rows = [
    ["المبيعات (شامل الضريبة)", ar(N.revenue) + " ريال", "١٫٩٦٨ ريال/لتر صافي الضريبة"],
    ["اللترات", ar(N.volume), ar(netLpd) + " لتر يومياً للشبكة"],
    ["الزيارات", ar(N.visits), ar(netVpd) + " زيارة يومياً"],
    ["متوسط الفاتورة", ar1(N.inv) + " ريال", "٢٥٫٦ لتر لكل زيارة"],
    ["المحطات", ar(N.stations), "٨٣٪ تغطية زمنية ، ٢٥ محطة بسجل كامل"],
    ["حجم البيانات", ar(N.raw_rows) + " عملية", ar(N.files) + " ملفاً · مستبعَد ٠٫٠٣٪"],
  ];
  table(s, M, 1.6, CW * 0.62, [
    { t: "البند", w: 34, a: "right" }, { t: "القيمة", w: 30 }, { t: "القراءة", w: 36, a: "right" }],
    rows, { rh: 0.44, fs: 12 });

  const x2 = M + CW * 0.64, w2 = CW * 0.36;
  s.addText("مزيج الوقود", { x: x2, y: 1.6, w: w2, h: 0.34, fontFace: F,
    fontSize: 14, bold: true, color: ORANGE, ...rtl });
  const fu = [["بنزين ٩١", 0.438, 0.555], ["بنزين ٩٥", 0.374, 0.360],
              ["ديزل", 0.188, 0.085], ["٩٨ وكيروسين", 0.0003, 0.0002]];
  fu.forEach((f, i) => {
    const y = 2.0 + i * 0.62;
    s.addText(f[0], { x: x2 + w2 * 0.68, y, w: w2 * 0.32, h: 0.3, fontFace: F,
      fontSize: 11.5, bold: true, color: INK, ...rtl });
    s.addShape(p.ShapeType.rect, { x: x2, y: y + 0.32, w: w2, h: 0.16, fill: { color: T_NEU } });
    s.addShape(p.ShapeType.rect, { x: x2 + w2 * (1 - f[1]), y: y + 0.32, w: w2 * f[1], h: 0.16,
      fill: { color: i === 2 ? BLUE : ORANGE } });
    s.addText(pc(f[1]) + " من الإيراد · " + pc(f[2]) + " من الزيارات",
      { x: x2, y, w: w2 * 0.64, h: 0.3, fontFace: F, fontSize: 10, color: INK2,
        align: "left", margin: 0 });
  });
  s.addShape(p.ShapeType.roundRect, { x: x2, y: 4.6, w: w2, h: 1.05, rectRadius: 0.04,
    fill: { color: T_OR }, line: { color: ORANGE, width: 1 } });
  s.addText("الديزل ٨٫٥٪ من الزيارات و١٨٫٨٪ من الإيراد — فاتورته ١٢٦ ريالاً، ٢٫٨ ضعف المتوسط. هذا عميل مختلف تماماً وتُبنى له خطة مستقلة.",
    { x: x2 + 0.14, y: 4.72, w: w2 - 0.28, h: 0.82, fontFace: F, fontSize: 11, color: INK, ...rtl });
  note(s, "ثلاث محطات (MK029 · MK072 · RY075) لتراتها ناقصة في المصدر فاستُبعدت من كل حساب لتر.");
}

/* ═══ ٥ · الاتجاه الحقيقي ═══ */
{
  const s = page("الاتجاه الحقيقي", "الإجمالي الخام مضلِّل — عدد المحطات المغطّاة ارتفع من ٣٩ إلى ٥٥");
  const base = [633296, 624536, 654588, 635997, 689121, 713904, 743943];
  const mn = ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو"];
  const x0 = M + 0.2, w0 = CW * 0.60, y0 = 1.75, h0 = 3.5;
  const mx = 780000, mi = 580000;
  s.addText("لتر يومياً · قاعدة ثابتة من ٢٥ محطة بتغطية كاملة ٢١٢ يوماً",
    { x: x0, y: y0 - 0.32, w: w0, h: 0.3, fontFace: F, fontSize: 11, color: INK2, ...rtl });
  base.forEach((v, i) => {
    const bw = w0 / 7 * 0.62, bx = x0 + w0 - (i + 1) * (w0 / 7) + (w0 / 7 - bw) / 2;
    const bh = (v - mi) / (mx - mi) * h0;
    s.addShape(p.ShapeType.rect, { x: bx, y: y0 + h0 - bh, w: bw, h: bh,
      fill: { color: i === 6 ? ORANGE : (i >= 4 ? GOLD : BGRAY) } });
    s.addText(ar(v), { x: bx - 0.15, y: y0 + h0 - bh - 0.34, w: bw + 0.3, h: 0.3,
      fontFace: F, fontSize: 10, bold: true, color: i === 6 ? ORANGE : INK2, align: "center", margin: 0 });
    s.addText(mn[i], { x: bx - 0.15, y: y0 + h0 + 0.06, w: bw + 0.3, h: 0.3,
      fontFace: F, fontSize: 10.5, color: INK, align: "center", margin: 0 });
  });
  s.addShape(p.ShapeType.rect, { x: x0, y: y0 + h0, w: w0, h: 0.02, fill: { color: LINE2 } });

  const x2 = M + CW * 0.64, w2 = CW * 0.36;
  stat(s, x2, 1.62, w2, 1.62, "+١٧٫٥٪", "نمو يناير ← يوليو", GOOD, "على نفس المحطات");
  s.addText("لكن مكة تسير عكس الشبكة", { x: x2, y: 3.35, w: w2, h: 0.36, fontFace: F,
    fontSize: 14, bold: true, color: BAD, ...rtl });
  const mk = [["المعيصم MK002", "−٧٢٫٧٪", BAD], ["بن درويش MK023", "−١٨٫٥٪", BAD],
              ["كورنيش جيزان JA069", "−٩٫٢٪", BAD], ["كورنيش الخبر EP051", "+٤٨٫١٪", GOOD],
              ["الكر النازل MK084", "+٣٣٫٢٪", GOOD]];
  mk.forEach((v, i) => {
    const y = 3.78 + i * 0.42;
    s.addText(v[0], { x: x2 + 1.1, y, w: w2 - 1.1, h: 0.36, fontFace: F, fontSize: 11.5, color: INK, ...rtl });
    s.addText(v[1], { x: x2, y, w: 1.0, h: 0.36, fontFace: F, fontSize: 12,
      bold: true, color: v[2], align: "left", margin: 0 });
  });
  s.addShape(p.ShapeType.roundRect, { x: x2, y: 5.95, w: w2, h: 0.6, rectRadius: 0.04,
    fill: { color: T_BAD }, line: { color: BAD, width: 1 } });
  s.addText("٤ محطات هابطة ، ٨ مستقرة ، ٩ نامية", { x: x2 + 0.12, y: 6.06, w: w2 - 0.24, h: 0.4,
    fontFace: F, fontSize: 12, bold: true, color: D_BAD, align: "center", margin: 0 });
  note(s, "القاعدة الثابتة تعزل أثر دخول محطات جديدة على التقارير — وهي القراءة الوحيدة الصالحة للاتجاه.");
}

/* ═══ ٦ · تقسيم السوق ═══ */
{
  const s = page("تقسيم السوق — من يشتري منّا", "أربعة عملاء مختلفون داخل نفس المحطة، لكل منهم خطة مختلفة");
  const seg = [
    ["أفراد ٩١", "٥٥٫٥٪ من الزيارات", "٤٥٫١ ريال", "الأكبر عدداً والأصغر سلة — هنا يعمل رفع السلة", ORANGE],
    ["أفراد ٩٥", "٣٦٫٠٪ من الزيارات", "٥٩٫٥ ريال", "قدرة إنفاق أعلى — قناة الخدمات المضافة", GOLD],
    ["أساطيل الديزل", "٨٫٥٪ من الزيارات", "١٢٦٫٣ ريال", "١٨٫٨٪ من الإيراد — يُدار بالتعاقد لا بالحملة", BLUE],
    ["أساطيل رقمية", "٠٫٧٪ من الزيارات", "١٤٩ إلى ١٧٢ ريال", "سيارة وبترو — أعلى فاتورة وأقل حصة، فرصة نمو صافية", GOOD],
  ];
  const cwid = (CW - 3 * 0.2) / 4;
  seg.forEach((v, i) => {
    const x = rtlx(i, cwid, 0.2);
    s.addShape(p.ShapeType.roundRect, { x, y: 1.6, w: cwid, h: 3.0, rectRadius: 0.05,
      fill: { color: W }, line: { color: LINE2, width: 1 } });
    s.addShape(p.ShapeType.rect, { x, y: 1.6, w: cwid, h: 0.55, fill: { color: v[4] } });
    s.addText(v[0], { x: x + 0.08, y: 1.66, w: cwid - 0.16, h: 0.44, fontFace: F,
      fontSize: 15, bold: true, color: W, align: "center", margin: 0 });
    s.addText(v[1], { x: x + 0.1, y: 2.3, w: cwid - 0.2, h: 0.32, fontFace: F,
      fontSize: 11.5, color: INK2, align: "center", margin: 0 });
    s.addText(v[2], { x: x + 0.1, y: 2.68, w: cwid - 0.2, h: 0.5, fontFace: F,
      fontSize: 21, bold: true, color: v[4], align: "center", margin: 0 });
    s.addText("متوسط الفاتورة", { x: x + 0.1, y: 3.16, w: cwid - 0.2, h: 0.28, fontFace: F,
      fontSize: 9.5, color: INK3, align: "center", margin: 0 });
    s.addText(v[3], { x: x + 0.12, y: 3.5, w: cwid - 0.24, h: 1.0, fontFace: F,
      fontSize: 11, color: INK, ...rtl });
  });
  s.addShape(p.ShapeType.roundRect, { x: M, y: 4.82, w: CW, h: 1.55, rectRadius: 0.05,
    fill: { color: T_BAD }, line: { color: BAD, width: 1.2 } });
  s.addText("الثغرة التي تسبق كل خطة", { x: M + 0.2, y: 4.94, w: CW - 0.4, h: 0.36,
    fontFace: F, fontSize: 15, bold: true, color: D_BAD, ...rtl });
  s.addText("٣٦٫١٪ من الإيراد (٢٢٦ مليون ريال) مسجّل «وسيلة دفع غير محددة» — و١٨ محطة لا تسجّلها إطلاقاً بإيراد ١٨٣ مليون ريال. أي حملة ولاء أو استهداف عميل تُبنى على بيانات ناقصة بثلثها. التفعيل يسبق الإنفاق.",
    { x: M + 0.2, y: 5.32, w: CW - 0.4, h: 0.95, fontFace: F, fontSize: 12.5, color: INK, ...rtl });
  note(s, "محفظة كاش إن نفسها: ٣١٨ ريالاً و٨ عمليات في سبعة أشهر — قناة غير مفعّلة عملياً.");
}

/* ═══ ٧ · القسم الثاني ═══ */
divider("القسم الثاني", "تقسيم المحطات", "خمس شرائح تُحدَّد من سلوك العميل لا من الجغرافيا");

/* ═══ ٨ · الشرائح الخمس ═══ */
{
  const s = page("خمس شرائح — والقاعدة التي تصنّف بها أي محطة",
                 "التصنيف من البيانات: حصة الديزل ثم الدفع المؤسسي ثم الحجم");
  const rows = SEG.map((g, i) => ({
    c: [{ t: g.seg, b: true, a: "right", c: BGRAY }, g.rule, ar(g.n),
        ar(g.lpd), pc(g.lpd / T.lpd), ar1(g.lpv), pc(g.diesel), pc(g.fleet), ar1(g.inv)],
    fill: i % 2 ? T_NEU : W }));
  rows.push({ c: [{ t: "الإجمالي", b: true }, "", ar(T.n), ar(T.lpd), "١٠٠٪",
                  ar1(N.volume / N.visits), "٢٠٪", "٥٫٤٪", ar1(N.inv)],
              fill: T_BAND, b: true });
  table(s, M, 1.6, CW, [
    { t: "الشريحة", w: 17, a: "right" }, { t: "قاعدة التصنيف", w: 20, a: "right" },
    { t: "محطات", w: 7 }, { t: "لتر/يوم", w: 11 }, { t: "٪ الحجم", w: 9 },
    { t: "لتر/زيارة", w: 10 }, { t: "ديزل", w: 8 }, { t: "أسطول", w: 8 }, { t: "الفاتورة", w: 10 }],
    rows, { rh: 0.44, fs: 11.5 });

  s.addText("من هو العميل في كل شريحة", { x: M, y: 4.8, w: CW, h: 0.32, fontFace: F,
    fontSize: 15, bold: true, color: ORANGE, ...rtl });
  const cwid = (CW - 4 * 0.14) / 5;
  SEG.forEach((g, i) => {
    const x = rtlx(i, cwid, 0.14);
    s.addShape(p.ShapeType.roundRect, { x, y: 5.16, w: cwid, h: 1.3, rectRadius: 0.04,
      fill: { color: BG }, line: { color: LINE2, width: 1 } });
    s.addShape(p.ShapeType.rect, { x, y: 5.16, w: cwid, h: 0.04, fill: { color: ORANGE } });
    s.addText(g.seg, { x: x + 0.08, y: 5.25, w: cwid - 0.16, h: 0.3, fontFace: F,
      fontSize: 12, bold: true, color: BGRAY, align: "center", margin: 0 });
    s.addText(g.who, { x: x + 0.1, y: 5.58, w: cwid - 0.2, h: 0.84, fontFace: F,
      fontSize: 10.5, color: INK, ...rtl });
  });
  note(s, "٥٢ محطة من ٥٥ دخلت التقسيم — الثلاث المستبعدة لتراتها ناقصة في المصدر.");
}

/* ═══ ٩ · القمة والقاع ═══ */
{
  const s = page("تحليل المنافذ — القمة والقاع", "أعلى وأدنى المحطات بالمبيعات اليومية · الترتيب يقود توزيع الجهد");
  const top = [...ST].sort((a, b) => b.lpd - a.lpd);
  const mk = r => [{ t: r.name.slice(0, 19), a: "right", b: true }, r.code, r.seg,
                   ar(r.lpd), ar(r.vpd), ar1(r.lpv), ar1(r.inv)];
  const cols = [{ t: "المحطة", w: 30, a: "right" }, { t: "الكود", w: 10 },
                { t: "الشريحة", w: 19 }, { t: "لتر/يوم", w: 12 },
                { t: "زيارة/يوم", w: 10 }, { t: "لتر/زيارة", w: 9 }, { t: "الفاتورة", w: 10 }];
  s.addText("أعلى ٨", { x: M + CW * 0.515, y: 1.5, w: CW * 0.485, h: 0.3, fontFace: F,
    fontSize: 13, bold: true, color: GOOD, ...rtl });
  table(s, M + CW * 0.515, 1.84, CW * 0.485, cols, top.slice(0, 8).map(mk),
        { rh: 0.375, fs: 9.5, hfs: 9.5 });
  s.addText("أدنى ٨", { x: M, y: 1.5, w: CW * 0.485, h: 0.3, fontFace: F,
    fontSize: 13, bold: true, color: BAD, ...rtl });
  table(s, M, 1.84, CW * 0.485, cols, top.slice(-8).reverse().map(mk),
        { rh: 0.375, fs: 9.5, hfs: 9.5 });
  const box = (x, w, t, b, col) => {
    s.addShape(p.ShapeType.roundRect, { x, y: 5.45, w, h: 1.05, rectRadius: 0.04,
      fill: { color: col === GOOD ? T_GOOD : T_BAD }, line: { color: col, width: 1 } });
    s.addText(t, { x: x + 0.14, y: 5.54, w: w - 0.28, h: 0.3, fontFace: F,
      fontSize: 12.5, bold: true, color: col === GOOD ? D_GOOD : D_BAD, ...rtl });
    s.addText(b, { x: x + 0.14, y: 5.86, w: w - 0.28, h: 0.58, fontFace: F,
      fontSize: 11, color: INK, ...rtl });
  };
  box(M + CW * 0.515, CW * 0.485, "التركّز شديد",
      "أعلى ٥ محطات تصنع ٣٠٪ من حجم الشبكة. وMK007 وحدها ١٣٠ ألف لتر يومياً — ٤٫٤ أضعاف متوسط المحطة.", GOOD);
  box(M, CW * 0.485, "القاع ليس فشلاً بالضرورة",
      "٧ محطات دون ١٢ ألف لتر/يوم لا تصنع سوى ٣٫٩٪ من الحجم. الأولوية فيها ضبط التكلفة والتأجير لا الحملات.", BAD);
}

/* ═══ ١٠ · القسم الثالث ═══ */
divider("القسم الثالث", "المنافسون", "نخسر بالقرب لا بالخدمة");

/* ═══ ١١ · الموقف التنافسي ═══ */
{
  const s = page("الموقف التنافسي — مسح ميداني على خمسة مواقع",
                 "تقييمات خرائط جوجل ومسافات فعلية ، ٤٠ منافساً مرصوداً");
  const rows = D.five.map((f, i) => ({
    c: [{ t: f.name, a: "right", b: true }, f.code,
        { t: ar1(f.rating) + "★", b: true, c: GOOD },
        { t: ar1(f.avg) + "★", c: BAD }, ar(f.n), arn(f.near), f.who.slice(0, 16), f.density],
    fill: i % 2 ? T_NEU : W }));
  table(s, M, 1.6, CW, [
    { t: "المحطة", w: 18, a: "right" }, { t: "الكود", w: 8 }, { t: "تقييم درب", w: 11 },
    { t: "متوسط المنافس", w: 13 }, { t: "منافسون داخل ٥ كم", w: 13 },
    { t: "أقرب منافس", w: 11 }, { t: "من هو", w: 15 }, { t: "الكثافة", w: 12 }],
    rows, { rh: 0.42, fs: 11 });

  const cwid = (CW - 2 * 0.2) / 3;
  const cards = [
    ["نتفوّق في كل موقع", "٤٫٨ و٤٫٩ نجمة لنا مقابل ٣٫٦ إلى ٤٫١ للمنافسين. و٥٧٪ من المنافسين المرصودين دون أربع نجوم.", GOOD],
    ["ونخسر بالمسافة", "أقرب منافس على ١٨١ متراً في المعيصم و٤١١ في الفردوس و٤٨٨ في الشرائع — الاعتراض قبل الوصول.", BAD],
    ["الدريس هو المنافس", "١٥ من ٤٠ موقعاً مرصوداً — ٣٨٪ محلياً مقابل ١٨٫٩٥٪ وطنياً. تركيز مقصود على مكة.", BLUE],
  ];
  cards.forEach((c, i) => {
    const x = rtlx(i, cwid, 0.2);
    s.addShape(p.ShapeType.roundRect, { x, y: 4.28, w: cwid, h: 1.44, rectRadius: 0.05,
      fill: { color: c[2] === GOOD ? T_GOOD : c[2] === BAD ? T_BAD : T_NEU },
      line: { color: c[2], width: 1.2 } });
    s.addText(c[0], { x: x + 0.14, y: 4.38, w: cwid - 0.28, h: 0.32, fontFace: F,
      fontSize: 14, bold: true, color: c[2] === GOOD ? D_GOOD : c[2] === BAD ? D_BAD : BLUE, ...rtl });
    s.addText(c[1], { x: x + 0.14, y: 4.72, w: cwid - 0.28, h: 0.94, fontFace: F,
      fontSize: 11.5, color: INK, ...rtl });
  });
  s.addShape(p.ShapeType.roundRect, { x: M, y: 5.82, w: CW, h: 0.62, rectRadius: 0.04,
    fill: { color: BGRAY } });
  s.addText("النتيجة التشغيلية: لا تُبنى خطة على تحسين الخدمة — الخدمة أفضل أصلاً. تُبنى على الاعتراض: لافتة، مدخل، سرعة، وإبراز التقييم حيث يقرر السائق قبل ٥٠٠ متر.",
    { x: M + 0.2, y: 5.92, w: CW - 0.4, h: 0.44, fontFace: F, fontSize: 12.5, bold: true, color: W, ...rtl });
}

/* ═══ ١٢ · حالة المعيصم ═══ */
{
  const s = page("حالة تشرح كل شيء — المعيصم MK002", "أعلى تقييم · أقرب منافس · أسوأ انهيار في الشبكة");
  const bars = [44360, 17707, 15734, 14478, 9539, 12749, 12012];
  const mn = ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو"];
  const x0 = M, w0 = CW * 0.56, y0 = 1.9, h0 = 3.1, mx = 48000;
  s.addText("لتر يومياً — MK002 المعيصم", { x: x0, y: 1.55, w: w0, h: 0.3,
    fontFace: F, fontSize: 12, bold: true, color: BGRAY, ...rtl });
  bars.forEach((v, i) => {
    const bw = w0 / 7 * 0.6, bx = x0 + w0 - (i + 1) * (w0 / 7) + (w0 / 7 - bw) / 2;
    const bh = v / mx * h0;
    s.addShape(p.ShapeType.rect, { x: bx, y: y0 + h0 - bh, w: bw, h: bh,
      fill: { color: i === 0 ? BGRAY : BAD } });
    s.addText(ar(v), { x: bx - 0.16, y: y0 + h0 - bh - 0.32, w: bw + 0.32, h: 0.28,
      fontFace: F, fontSize: 9.5, bold: true, color: INK2, align: "center", margin: 0 });
    s.addText(mn[i], { x: bx - 0.16, y: y0 + h0 + 0.05, w: bw + 0.32, h: 0.28,
      fontFace: F, fontSize: 10, color: INK, align: "center", margin: 0 });
  });
  s.addShape(p.ShapeType.rect, { x: x0, y: y0 + h0, w: w0, h: 0.02, fill: { color: LINE2 } });

  const x2 = M + CW * 0.60, w2 = CW * 0.40;
  const facts = [["التقييم", "٤٫٨★ مقابل ٣٫٧★ للمنافسين", GOOD],
                 ["أقرب منافس", "١٨١ متراً — الأقرب في الشبكة", BAD],
                 ["منافسون داخل ٥ كم", "١٨ محطة، كثافة مرتفعة", BAD],
                 ["الانهيار", "−٧٢٫٧٪ من يناير إلى يوليو", BAD],
                 ["مطابقة المشتريات", "+٠٫٧١٪ — لا فاقد ولا خلل", GOOD]];
  facts.forEach((f, i) => {
    const y = 1.72 + i * 0.66;
    s.addShape(p.ShapeType.rect, { x: SW - M - 0.045, y, w: 0.045, h: 0.56, fill: { color: f[2] } });
    s.addText(f[0], { x: x2, y, w: w2 - 0.16, h: 0.3, fontFace: F, fontSize: 11,
      color: INK3, ...rtl });
    s.addText(f[1], { x: x2, y: y + 0.28, w: w2 - 0.16, h: 0.32, fontFace: F,
      fontSize: 12.5, bold: true, color: f[2], ...rtl });
  });
  s.addShape(p.ShapeType.roundRect, { x: x2, y: 5.15, w: w2, h: 1.25, rectRadius: 0.05,
    fill: { color: T_OR }, line: { color: ORANGE, width: 1.2 } });
  s.addText("التشخيص", { x: x2 + 0.14, y: 5.26, w: w2 - 0.28, h: 0.3, fontFace: F,
    fontSize: 13, bold: true, color: ORANGE, ...rtl });
  s.addText("لا مشكلة تشغيل ولا فاقد ولا خدمة. محطة اعترضت الطريق على ١٨١ متراً. العلاج اعتراضي: لافتة ومدخل وعرض على المسار — لا تدريب ولا تحسين خدمة.",
    { x: x2 + 0.14, y: 5.58, w: w2 - 0.28, h: 0.78, fontFace: F, fontSize: 11.5, color: INK, ...rtl });
  note(s, "نفس النمط في MK019 الشرائع (−٥٤٪ · أقرب منافس ٤٨٨ م) — الحالة ليست فردية.");
}

/* ═══ ١٣ · المنافذ داخل المحطة ═══ */
{
  const s = page("المنفذ الثاني — الوحدات التأجيرية داخل المحطة",
                 "سجل ١٨٦ محطة · الوحدة = كشك أو محل أو درايف ثرو أو مغسلة أو سوبرماركت");
  const U = D.units;
  const cwid = (CW - 3 * 0.18) / 4;
  const st = [[ar(U.total), "وحدة مسجّلة", ar(U.n) + " محطة", BGRAY],
              [ar(U.leased), "مؤجرة", pc(U.leased / U.total) + " إشغال", ORANGE],
              [ar(U.vacant), "شاغرة", "الفرصة المباشرة", BAD],
              [pc(U.shops_leased / U.shops), "إشغال المحلات", ar(U.shops) + " محلاً — المنتج العالق", BAD]];
  st.forEach((v, i) => stat(s, rtlx(i, cwid, 0.18), 1.55, cwid, 1.6, v[0], v[1], v[3], v[2]));
  const rows = U.top.map((u, i) => ({
    c: [{ t: u.name.slice(0, 26), a: "right", b: true }, u.code, u.cat,
        ar(u.n), ar(u.leased), { t: ar(u.vacant), b: true, c: D_BAD },
        u.n ? pc(u.leased / u.n) : "—"],
    fill: i % 2 ? T_NEU : W }));
  table(s, M, 3.45, CW * 0.60, [
    { t: "المحطة", w: 32, a: "right" }, { t: "الكود", w: 10 }, { t: "الفئة", w: 15 },
    { t: "وحدات", w: 11 }, { t: "مؤجرة", w: 11 }, { t: "شاغرة", w: 11 }, { t: "إشغال", w: 10 }],
    rows, { rh: 0.36, fs: 10.5, hfs: 10 });
  s.addText("أكبر ست فرص تأجير", { x: M, y: 3.12, w: CW * 0.6, h: 0.3, fontFace: F,
    fontSize: 13, bold: true, color: ORANGE, ...rtl });

  const x2 = M + CW * 0.63, w2 = CW * 0.37;
  s.addText("أين الشغور فعلاً", { x: x2, y: 3.12, w: w2, h: 0.3, fontFace: F,
    fontSize: 13, bold: true, color: ORANGE, ...rtl });
  bullets(s, x2, 3.45, w2, [
    "الشبكة العاملة إشغالها ٧٩٪ — لا أزمة شغور في المحطات المفتوحة",
    "٦٨٩ وحدة في ٤٣ محطة تحت التنفيذ — التأجير المسبق هو الفرصة الأكبر",
    "٢٦٩ وحدة في الامتياز ومحلاته ٥ مؤجرة من ١٤٦ (٣٪)",
    "السجل لا يحمل قيمة إيجارية — لا يمكن تسعير الفرصة بالريال منه وحده",
  ], INK, 11.5);
  note(s, "الوحدات مرتبطة بالوقود: مستأجر نشط يولّد زيارات، والزيارة تولّد لتراً — وهذا أساس حملات المستأجرين.");
}

/* ═══ ١٤ · القسم الرابع ═══ */
divider("القسم الرابع", "خطة المبيعات", "مستهدف لكل محطة ثم لكل وردية ثم لكل عامل");

/* ═══ ١٥ · منهج المستهدف ═══ */
{
  const s = page("منهج المستهدف — لماذا اللتر لكل زيارة", "لا نستهدف نمواً مطلقاً، بل إغلاق فجوة مثبتة داخل الشريحة");
  const steps = [
    ["١", "الحجم = زيارات × لتر لكل زيارة", "الزيارات تحكمها الجغرافيا والمنافسة — بطيئة التغيّر. اللتر لكل زيارة يحكمه سلوك العامل والعرض — يتحرك خلال أسابيع."],
    ["٢", "المعيار من داخل الشريحة", "الربيع الأعلى للتر/زيارة بين محطات نفس الشريحة. محطة طريق لا تُقاس بمحطة حي."],
    ["٣", "نغلق ٤٠٪ من الفجوة في ٦ أشهر", "لا نطلب الوصول للمعيار دفعة واحدة. الفارق مثبت أن غيرنا حقّقه في نفس الظروف."],
    ["٤", "القيمة بهامش الشبكة الفعلي", "١١٫٣ هللة لكل لتر — من قائمة دخل يوليو للشبكة المشغّلة، لا افتراض."],
  ];
  steps.forEach((v, i) => {
    const y = 1.6 + i * 1.02;
    s.addShape(p.ShapeType.roundRect, { x: M, y, w: CW * 0.63, h: 0.9, rectRadius: 0.04,
      fill: { color: i === 1 ? T_OR : BG }, line: { color: LINE2, width: 1 } });
    s.addShape(p.ShapeType.ellipse, { x: SW - M - CW * 0.37 - 0.72, y: y + 0.22, w: 0.46, h: 0.46,
      fill: { color: ORANGE } });
    s.addText(v[0], { x: SW - M - CW * 0.37 - 0.72, y: y + 0.26, w: 0.46, h: 0.38,
      fontFace: F, fontSize: 15, bold: true, color: W, align: "center", margin: 0 });
    s.addText(v[1], { x: M + 0.14, y: y + 0.1, w: CW * 0.63 - 0.95, h: 0.32, fontFace: F,
      fontSize: 13.5, bold: true, color: BGRAY, ...rtl });
    s.addText(v[2], { x: M + 0.14, y: y + 0.42, w: CW * 0.63 - 0.95, h: 0.44, fontFace: F,
      fontSize: 11, color: INK2, ...rtl });
  });
  const x2 = M + CW * 0.66, w2 = CW * 0.34;
  s.addShape(p.ShapeType.roundRect, { x: x2, y: 1.6, w: w2, h: 4.0, rectRadius: 0.05,
    fill: { color: BGRAY } });
  s.addText("النتيجة على الشبكة", { x: x2 + 0.16, y: 1.75, w: w2 - 0.32, h: 0.36,
    fontFace: F, fontSize: 15, bold: true, color: ORANGE, ...rtl });
  const res = [["القاعدة اليوم", ar(T.lpd) + " لتر/يوم"],
               ["محطات دون معيارها", ar(T.below) + " من " + ar(T.n)],
               ["الرفع المستهدف", "+" + ar(T.upl) + " لتر/يوم"],
               ["نسبة الرفع", "+" + ar1(T.upl / T.lpd * 100) + "٪"],
               ["القيمة السنوية", ar(T.sar) + " ريال"]];
  res.forEach((r, i) => {
    const y = 2.25 + i * 0.62;
    s.addText(r[0], { x: x2 + 0.16, y, w: w2 - 0.32, h: 0.28, fontFace: F,
      fontSize: 10.5, color: "C9C4BE", ...rtl });
    s.addText(r[1], { x: x2 + 0.16, y: y + 0.24, w: w2 - 0.32, h: 0.34, fontFace: F,
      fontSize: 15, bold: true, color: i === 4 ? GOLD : W, ...rtl });
  });
  note(s, "الرفع محسوب محطة بمحطة ثم جُمع — لا نسبة مسقطة على الشبكة.");
}

/* ═══ ١٦ · المستهدف بالشريحة ═══ */
{
  const s = page("المستهدف بالشريحة", "معيار لتر/زيارة داخل كل شريحة · والرفع المطلوب خلال ستة أشهر");
  const rows = SEG.map((g, i) => ({
    c: [{ t: g.seg, a: "right", b: true, c: BGRAY }, ar(g.n),
        ar1(g.lpv), { t: ar1(g.bench), b: true, c: ORANGE },
        { t: ar(g.below), c: g.below ? D_BAD : INK },
        { t: "+" + ar(g.upl), b: true, c: D_GOOD }, "+" + ar1(g.upl / g.lpd * 100) + "٪",
        { t: ar(g.sar), b: true }],
    fill: i % 2 ? T_NEU : W }));
  rows.push({ c: [{ t: "الإجمالي", b: true }, ar(T.n), ar1(N.volume / N.visits), "—",
                  ar(T.below), "+" + ar(T.upl), "+" + ar1(T.upl / T.lpd * 100) + "٪", ar(T.sar)],
              fill: T_BAND, b: true });
  table(s, M, 1.6, CW, [
    { t: "الشريحة", w: 20, a: "right" }, { t: "محطات", w: 9 },
    { t: "لتر/زيارة اليوم", w: 14 }, { t: "المعيار", w: 11 },
    { t: "دون المعيار", w: 12 }, { t: "الرفع لتر/يوم", w: 13 },
    { t: "٪", w: 9 }, { t: "القيمة السنوية (ر)", w: 16 }],
    rows, { rh: 0.46, fs: 11.5 });

  s.addShape(p.ShapeType.roundRect, { x: M, y: 5.05, w: CW, h: 1.32, rectRadius: 0.05,
    fill: { color: T_OR }, line: { color: ORANGE, width: 1.2 } });
  s.addText("كيف تُقرأ", { x: M + 0.2, y: 5.16, w: CW - 0.4, h: 0.3, fontFace: F,
    fontSize: 13.5, bold: true, color: ORANGE, ...rtl });
  s.addText("محاور الطريق تحمل نصف الفرصة (+٥١٬٧٨٣ لتراً يومياً) لأن سلتها الكبيرة تجعل كل نقطة تحسّن أثمن · الحضرية الكبرى تليها بحجمها لا بفجوتها · أما المحلية الصغيرة فقيمتها ١٠٤ آلاف ريال سنوياً — لا تستحق حملة، وخطتها ضبط التكلفة والتأجير.",
    { x: M + 0.2, y: 5.5, w: CW - 0.4, h: 0.78, fontFace: F, fontSize: 12, color: INK, ...rtl });
}

/* ═══ خطة كل شريحة ═══ */
{
  const s = page("خطة كل شريحة", "لكل شريحة محرّك واحد وفعل تجاري واحد وحملة واحدة — لا خطة عامة للشبكة");
  const rows2 = SEG.map((g, i) => ({
    c: [{ t: g.seg, a: "right", b: true, c: BGRAY }, { t: g.driver, a: "right", b: true, c: ORANGE },
        { t: g.action, a: "right" }, { t: g.camp, a: "right" }],
    fill: i % 2 ? T_NEU : W }));
  table(s, M, 1.6, CW, [
    { t: "الشريحة", w: 15, a: "right" }, { t: "المحرّك", w: 17, a: "right" },
    { t: "الفعل التجاري", w: 42, a: "right" }, { t: "الحملة الأنسب", w: 26, a: "right" }],
    rows2, { rh: 0.72, fs: 11, hfs: 11.5 });
  s.addShape(p.ShapeType.roundRect, { x: M, y: 5.65, w: CW, h: 1.25, rectRadius: 0.05,
    fill: { color: BGRAY } });
  s.addText("القاعدة التي تحكم كل ما سبق", { x: M + 0.2, y: 5.76, w: CW - 0.4, h: 0.3,
    fontFace: F, fontSize: 14, bold: true, color: ORANGE, ...rtl });
  s.addText("العميل يحدد الخطة لا الجغرافيا. محطتان متجاورتان في مكة قد تنتميان لشريحتين مختلفتين إن كانت إحداهما على مسار الشاحنات والأخرى داخل حي سكني — فتُدار كل واحدة بأدوات مختلفة تماماً.",
    { x: M + 0.2, y: 6.1, w: CW - 0.4, h: 0.7, fontFace: F, fontSize: 12.5, color: W, ...rtl });
}

/* ═══ ١٧ · أكبر الفرص ═══ */
{
  const s = page("أكبر ١٤ فرصة — بالمحطة", "مرتّبة بحجم الرفع اليومي · وهي قائمة الأولوية التنفيذية");
  const rows = ST.slice(0, 14).map((r, i) => ({
    c: [{ t: r.name.slice(0, 22), a: "right", b: true }, r.code, r.seg,
        ar1(r.lpv), ar1(r.bench), { t: ar1(r.tgt_lpv), b: true, c: ORANGE },
        { t: "+" + ar(r.upl_lpd), b: true, c: D_GOOD }, ar(r.upl_sar),
        { t: r.conds.slice(0, 2).join(" · ") || "—", a: "right", fs: 9.5 }],
    fill: i % 2 ? T_NEU : W }));
  table(s, M, 1.55, CW, [
    { t: "المحطة", w: 19, a: "right" }, { t: "الكود", w: 7 }, { t: "الشريحة", w: 14 },
    { t: "لتر/زيارة", w: 9 }, { t: "المعيار", w: 8 }, { t: "المستهدف", w: 9 },
    { t: "الرفع/يوم", w: 10 }, { t: "القيمة السنوية", w: 11 }, { t: "الحالة المرصودة", w: 20, a: "right" }],
    rows, { rh: 0.31, fs: 10, hfs: 10 });
  s.addShape(p.ShapeType.roundRect, { x: M, y: 6.42, w: CW, h: 0.52, rectRadius: 0.04,
    fill: { color: T_OR }, line: { color: ORANGE, width: 1 } });
  s.addText("هذه الأربع عشرة محطة تحمل " + pc(ST.slice(0, 14).reduce((a, r) => a + r.upl_lpd, 0) / T.upl) +
            " من الفرصة كلها — فالجهد يبدأ منها لا من الشبكة دفعة واحدة.",
    { x: M + 0.16, y: 6.5, w: CW - 0.32, h: 0.36, fontFace: F, fontSize: 12, bold: true, color: INK, ...rtl });
}

/* ═══ ١٨ · مستهدف الوردية والعامل ═══ */
{
  const s = page("من مستهدف المحطة إلى مستهدف العامل",
                 "ثلاث طبقات — والوحدة «معاملة» في الطبقتين الأوليين، والريال في الثالثة فقط");

  /* ① المثال المحسوب — سلسلة من اليمين لليسار */
  s.addText("① كيف يُشتقّ الرقم — العمرة الجديدة MK007", { x: M, y: 1.46, w: CW, h: 0.3,
    fontFace: F, fontSize: 13.5, bold: true, color: ORANGE, ...rtl });
  const bw = 2.62, gp = 0.34, y0 = 1.76, bh = 0.82;
  /* الصندوق الأول على اليمين */
  s.addShape(p.ShapeType.roundRect, { x: SW - M - bw, y: y0 + 0.42, w: bw, h: bh,
    rectRadius: 0.05, fill: { color: BGRAY } });
  s.addText("٤٬٢٦١", { x: SW - M - bw, y: y0 + 0.5, w: bw, h: 0.44, fontFace: F,
    fontSize: 25, bold: true, color: W, align: "center", margin: 0 });
  s.addText("معاملة يومياً — كل المحطة", { x: SW - M - bw, y: y0 + 0.93, w: bw, h: 0.3,
    fontFace: F, fontSize: 10.5, color: "D6D2CC", align: "center", margin: 0 });

  const step = (x, y, ttl, big, sub, col) => {
    s.addShape(p.ShapeType.roundRect, { x, y, w: bw, h: bh, rectRadius: 0.05,
      fill: { color: col === ORANGE ? T_OR : T_NEU }, line: { color: col, width: 1.2 } });
    s.addText(ttl, { x: x + 0.08, y: y + 0.06, w: bw - 0.16, h: 0.26, fontFace: F,
      fontSize: 10, color: INK2, align: "center", margin: 0 });
    s.addText(big, { x: x + 0.08, y: y + 0.28, w: bw - 0.16, h: 0.36, fontFace: F,
      fontSize: 11.5, bold: true, color: BGRAY, align: "center", margin: 0 });
    s.addText(sub, { x: x + 0.08, y: y + 0.62, w: bw - 0.16, h: 0.26, fontFace: F,
      fontSize: 12.5, bold: true, color: col, align: "center", margin: 0 });
  };
  const arrow = (x, y) => s.addText("◀", { x, y, w: gp, h: bh, fontFace: F,
    fontSize: 13, color: INK3, align: "center", valign: "middle", margin: 0 });

  const x1 = SW - M - bw - gp - bw, x2 = x1 - gp - bw, x3 = x2 - gp - bw;
  arrow(SW - M - bw - gp, y0 + 0.43);
  step(x1, y0, "الصباح ٦ص – ٦م", "٢٬٠٨٧ معاملة ÷ ١١ عاملاً", "= ١٩٠ معاملة للعامل", BGRAY);
  step(x1, y0 + 0.9, "المساء ٦م – ٦ص", "٢٬١٧٤ معاملة ÷ ٩ عمال", "= ٢٤٢ معاملة للعامل", ORANGE);
  arrow(x1 - gp, y0 + 0.43);
  step(x2, y0 + 0.43, "الفارق بين الورديتين", "المساء ٥١٪ من الطلب · ٤٥٪ من الطاقم",
       "المسائي يخدم +٢٧٪", BAD);
  arrow(x2 - gp, y0 + 0.43);
  step(x3, y0 + 0.43, "ما نطلبه فعلاً", "فاقد ذروة وردية المساء",
       "٥١ معاملة تُستردّ", GOOD);

  /* ② الجدول */
  s.addText("② الأرقام — كلها بالمعاملة", { x: M, y: 3.56, w: CW, h: 0.3,
    fontFace: F, fontSize: 13.5, bold: true, color: ORANGE, ...rtl });
  const SH = [["العمرة الجديدة", "MK007", 4261, 2087, 2174, "١١ / ٩", 190, 242, "+٢٧٪", 34, 51],
              ["عرفات الشوقية", "MK017", 1269, 567, 702, "٣ / ٣", 189, 234, "+٢٤٪", 1, 1],
              ["المعيصم", "MK002", 865, 429, 436, "٢ / ٢", 214, 218, "+٢٪", 1, 9],
              ["بن درويش", "MK023", 790, 357, 434, "٢ / ٢", 178, 217, "+٢٢٪", 0, 0],
              ["الشرايع", "MK019", 443, 194, 248, "٢ / ٢", 97, 124, "+٢٨٪", 1, 1]];
  const rows = SH.map((r, i) => ({
    c: [{ t: r[0], a: "right", b: true }, r[1], ar(r[2]), ar(r[3]), ar(r[4]), r[5],
        { t: ar(r[6]), b: true }, { t: ar(r[7]), b: true, c: ORANGE },
        { t: r[8], c: D_BAD }, { t: ar(r[9] + r[10]), b: true, c: D_GOOD }],
    fill: i % 2 ? T_NEU : W }));
  table(s, M, 3.86, CW, [
    { t: "المحطة", w: 15, a: "right" }, { t: "الكود", w: 8 },
    { t: "معاملات المحطة/يوم", w: 12 }, { t: "منها صباحاً", w: 10 },
    { t: "منها مساءً", w: 10 }, { t: "عمال ص / م", w: 9 },
    { t: "عبء الصباحي", w: 10 }, { t: "عبء المسائي", w: 10 },
    { t: "الفارق", w: 8 }, { t: "المستهدف اليومي", w: 11 }],
    rows, { rh: 0.335, fs: 10.5, hfs: 9.5 });

  /* ③ الطبقات الثلاث */
  const L = [["الطبقة ١ · عبء الوردية", "معاملة",
              "معاملات الوردية ÷ عمالها. حِمل لا مستهدف: العامل لا يصنع المعاملة.", BGRAY],
             ["الطبقة ٢ · مستهدف الوردية", "معاملة",
              "ما تفقده المحطة في ساعات ذروتها مقارنة بشريحتها. هذا وحده ما يُطالَب به.", ORANGE],
             ["الطبقة ٣ · الحافز", "ريال",
              "اللترات الإضافية × هامش المحطة × نسبة، ببوابة جودة وسقف ١٠٠ ريال.", GOOD]];
  const cwid = (CW - 2 * 0.2) / 3;
  L.forEach((k, i) => {
    const x = rtlx(i, cwid, 0.2);
    s.addShape(p.ShapeType.roundRect, { x, y: 6.06, w: cwid, h: 0.88, rectRadius: 0.05,
      fill: { color: BG }, line: { color: k[3], width: 1.2 } });
    s.addShape(p.ShapeType.rect, { x, y: 6.06, w: cwid, h: 0.05, fill: { color: k[3] } });
    s.addText(k[0], { x: x + 1.02, y: 6.14, w: cwid - 1.12, h: 0.28, fontFace: F,
      fontSize: 11.5, bold: true, color: BGRAY, ...rtl });
    s.addShape(p.ShapeType.roundRect, { x: x + 0.1, y: 6.15, w: 0.78, h: 0.25,
      rectRadius: 0.03, fill: { color: k[3] } });
    s.addText(k[1], { x: x + 0.1, y: 6.15, w: 0.78, h: 0.25, fontFace: F,
      fontSize: 9.5, bold: true, color: W, align: "center", valign: "middle", margin: 0 });
    s.addText(k[2], { x: x + 0.1, y: 6.42, w: cwid - 0.2, h: 0.5, fontFace: F,
      fontSize: 10, color: INK, ...rtl });
  });
}

/* ═══ ١٩ · نموذج الحافز ═══ */
{
  const s = page("نموذج الحافز — معتمد ومطبَّق", "الحافز يُدفع من هامش اللتر الإضافي، لا من ميزانية منفصلة");
  const flow = [["اللترات فوق الأساس", "الفارق بين إنتاج العامل وأساس ورديته"],
                ["× هامش المحطة", "هللة/لتر الفعلي لكل محطة — ١٠٫١٥ إلى ١٣٫١١"],
                ["× نسبة المشاركة", "حصة العامل من الهامش الإضافي"],
                ["بوابة الجودة", "غير مستوفاة ← صفر مهما بلغ الإنتاج"],
                ["سقف ١٠٠ ريال", "حد شهري لكل عامل — معتمد"]];
  flow.forEach((f, i) => {
    const y = 1.6 + i * 0.78;
    const last = i >= 3;
    s.addShape(p.ShapeType.roundRect, { x: M, y, w: CW * 0.55, h: 0.66, rectRadius: 0.04,
      fill: { color: last ? T_BAD : T_OR }, line: { color: last ? BAD : ORANGE, width: 1 } });
    s.addText(f[0], { x: M + 0.16, y: y + 0.05, w: CW * 0.55 - 0.32, h: 0.3, fontFace: F,
      fontSize: 13, bold: true, color: last ? D_BAD : BGRAY, ...rtl });
    s.addText(f[1], { x: M + 0.16, y: y + 0.34, w: CW * 0.55 - 0.32, h: 0.28, fontFace: F,
      fontSize: 10.5, color: INK2, ...rtl });
  });
  const x2 = M + CW * 0.58, w2 = CW * 0.42;
  s.addText("لماذا السقف والبوابات", { x: x2, y: 1.6, w: w2, h: 0.34, fontFace: F,
    fontSize: 15, bold: true, color: ORANGE, ...rtl });
  bullets(s, x2, 1.98, w2, [
    "الحافز يُحتسب لكل عامل بينما الشركة تجني الصافي — بلا سقف تجاوزت نسبة الحافز إلى الهامش ٤٥٪ في المحاكاة مقابل ١٥٪ المعتمدة",
    "بوابة الجودة تمنع شراء الحجم على حساب الخدمة — وخدمتنا هي ميزتنا الوحيدة على المنافس",
    "بوابة المحطة: لا حافز في محطة صافيها سالب مهما بلغ إنتاج الفرد",
    "المؤشر الحاكم في اللوحة: الحافز ÷ الهامش الإضافي — يُراجَع شهرياً",
  ], INK, 11.5);
  s.addShape(p.ShapeType.roundRect, { x: x2, y: 4.95, w: w2, h: 1.45, rectRadius: 0.05,
    fill: { color: BGRAY } });
  s.addText("الملف جاهز ومسلَّم", { x: x2 + 0.16, y: 5.06, w: w2 - 0.32, h: 0.32,
    fontFace: F, fontSize: 13, bold: true, color: ORANGE, ...rtl });
  s.addText("«نموذج العامل» — ورقة إدخال تفاعلية تحتسب الأساس والحافز لكل عامل، ونسخة ويب على الجوال. تحتاج فقط جدول عمالة كل محطة لتعميمها.",
    { x: x2 + 0.16, y: 5.42, w: w2 - 0.32, h: 0.9, fontFace: F, fontSize: 11.5, color: W, ...rtl });
}

/* ═══ ٢٠ · القسم الخامس ═══ */
divider("القسم الخامس", "الحملات", "لا حملة عامة — كل حملة تُطلق بحالة مرصودة في محطة بعينها");

/* ═══ ٢١ · حملات المستأجرين ═══ */
{
  const s = page("حملات المستأجرين", "المستأجر يولّد زيارة، والزيارة تولّد لتراً — القناتان تُداران معاً لا منفصلتين");
  const camps = [
    ["مساحة مقابل نسبة", "٥٧٢ محلاً شاغراً",
     "بدل إيجار ثابت مرتفع يعطّل الوحدة: إيجار أساسي منخفض + نسبة من المبيعات. يقلّل حاجز الدخول ويربط دخلنا بنجاح المستأجر.",
     ORANGE],
    ["التأجير المسبق", "٦٨٩ وحدة في ٤٣ محطة تحت التنفيذ",
     "التسويق يبدأ قبل الافتتاح بثلاثة أشهر بباقة «مستأجر مؤسِّس»: أفضلية موقع + إعفاء أول شهرين. الوحدة تفتح مؤجَّرة لا شاغرة.",
     BLUE],
    ["الكوبون المتبادل", "كل محطة بمستأجر نشط",
     "تعبئة تبلغ العتبة تعطي كوبون المستأجر · وفاتورة المستأجر تعطي خصم غسيل. حركة تتبادل بين القناتين بلا كلفة نقدية.",
     GOOD],
    ["شبكة الوسطاء", "الامتياز وتحت التنفيذ",
     "عمولة للوسيط المحلي على العقد المُوقَّع لا على العرض. تُوجَّه للامتياز حيث إشغال المحلات ٣٪ فقط، لا للمحطات العاملة حيث الإشغال ٧٩٪.",
     GOLD],
  ];
  const cwid = (CW - 3 * 0.18) / 4;
  camps.forEach((c, i) => {
    const x = rtlx(i, cwid, 0.18);
    s.addShape(p.ShapeType.roundRect, { x, y: 1.55, w: cwid, h: 3.5, rectRadius: 0.05,
      fill: { color: W }, line: { color: LINE2, width: 1 } });
    s.addShape(p.ShapeType.rect, { x, y: 1.55, w: cwid, h: 0.6, fill: { color: c[3] } });
    s.addText(c[0], { x: x + 0.08, y: 1.62, w: cwid - 0.16, h: 0.46, fontFace: F,
      fontSize: 14, bold: true, color: W, align: "center", margin: 0 });
    s.addText(c[1], { x: x + 0.1, y: 2.28, w: cwid - 0.2, h: 0.46, fontFace: F,
      fontSize: 11, bold: true, color: c[3], align: "center", margin: 0 });
    s.addShape(p.ShapeType.rect, { x: x + 0.35, y: 2.8, w: cwid - 0.7, h: 0.02, fill: { color: LINE2 } });
    s.addText(c[2], { x: x + 0.12, y: 2.92, w: cwid - 0.24, h: 2.0, fontFace: F,
      fontSize: 11, color: INK, ...rtl });
  });
  s.addShape(p.ShapeType.roundRect, { x: M, y: 5.25, w: CW, h: 1.15, rectRadius: 0.05,
    fill: { color: T_GOLD }, line: { color: GOLD, width: 1.2 } });
  s.addText("ما ينقص لتسعير هذه الحملات", { x: M + 0.2, y: 5.36, w: CW - 0.4, h: 0.32,
    fontFace: F, fontSize: 13.5, bold: true, color: "A8681A", ...rtl });
  s.addText("سجل الوحدات يحمل الأعداد لا القيم الإيجارية. لا يمكن قول «الفرصة تساوي كذا مليوناً» قبل جدول إيجارات لكل نوع وحدة ومدينة. المطلوب: جدول الإيجار الحالي والمستهدف لكل نوع — ثم تُسعَّر الـ٦٨٩ وحدة والـ٥٧٢ محلاً خلال أسبوع.",
    { x: M + 0.2, y: 5.7, w: CW - 0.4, h: 0.62, fontFace: F, fontSize: 12, color: INK, ...rtl });
}

/* ═══ ٢٢ · علبة المناديل ═══ */
{
  const s = page("هدايا علب المناديل — التصميم واقتصادياته",
                 "المكافأة يجب أن تسترد كلفتها من اللتر الإضافي، وإلا فهي إنفاق لا حملة");
  s.addText("الحساب الحاكم", { x: M, y: 1.5, w: CW * 0.44, h: 0.32, fontFace: F,
    fontSize: 14, bold: true, color: ORANGE, ...rtl });
  const be = [[2.0, 17.7], [2.5, 22.1], [3.0, 26.5], [3.5, 30.9]];
  const rows = be.map((b, i) => ({
    c: [ar1(b[0]) + " ريال", ar1(b[1]) + " لتر",
        { t: ar1(b[1] / 2.0) + " ضعفاً", b: true, c: D_BAD }],
    fill: i % 2 ? T_NEU : W }));
  table(s, M, 1.85, CW * 0.44, [
    { t: "كلفة العلبة", w: 26 }, { t: "اللترات الإضافية للتعادل", w: 38 },
    { t: "مقابل الرفع الواقعي", w: 36 }], rows, { rh: 0.38, fs: 11, hfs: 10.5 });
  s.addShape(p.ShapeType.roundRect, { x: M, y: 3.92, w: CW * 0.44, h: 1.5, rectRadius: 0.05,
    fill: { color: T_BAD }, line: { color: BAD, width: 1.2 } });
  s.addText("الخلاصة الصريحة", { x: M + 0.14, y: 4.03, w: CW * 0.44 - 0.28, h: 0.3,
    fontFace: F, fontSize: 13, bold: true, color: D_BAD, ...rtl });
  s.addText("الرفع المستهدف للسلة +١٫٢ إلى +٣٫١ لتر لكل زيارة. عند هامش ١١٫٣ هللة لا تسترد علبة بريالين كلفتها إلا بـ١٧٫٧ لتراً إضافياً — أي تسعة إلى خمسة عشر ضعف الرفع الواقعي. الهدية الممولة من هامش الوقود تخسر.",
    { x: M + 0.14, y: 4.35, w: CW * 0.44 - 0.28, h: 1.0, fontFace: F, fontSize: 11.5, color: INK, ...rtl });

  const x2 = M + CW * 0.47, w2 = CW * 0.53;
  s.addText("التصميم المعتمد — العلبة تموّل نفسها", { x: x2, y: 1.5, w: w2, h: 0.32,
    fontFace: F, fontSize: 14, bold: true, color: ORANGE, ...rtl });
  const dz = [
    ["١", "العلبة مساحة إعلانية", "تُطبع بعلامة معلن — مستأجر أو مورّد زيوت أو تأمين أو اتصالات. الكلفة تنتقل للمعلن وتتحول لدخل."],
    ["٢", "توزيع مشروط لا عام", "لا تُعطى لكل زيارة: فقط في ٣٩ محطة دون معيارها، وفقط عند بلوغ عتبة الشريحة: ٥٠ لتراً في محاور الطريق و٣٠ لتراً في الحضرية."],
    ["٣", "سقف يومي لكل محطة", "عدد محدد يومياً يمنع الانفلات ويجعل الهدية محدودة — والندرة نفسها محرّك سلوك."],
    ["٤", "القياس أسبوعي", "لتر/زيارة قبل وبعد في نفس المحطة. لا تمديد لحملة لم ترفع السلة خلال ٤ أسابيع."],
  ];
  dz.forEach((d, i) => {
    const y = 1.85 + i * 1.14;
    s.addShape(p.ShapeType.roundRect, { x: x2, y, w: w2, h: 1.02, rectRadius: 0.04,
      fill: { color: i === 0 ? T_GOOD : BG }, line: { color: i === 0 ? GOOD : LINE2, width: 1 } });
    s.addShape(p.ShapeType.ellipse, { x: x2 + w2 - 0.6, y: y + 0.28, w: 0.44, h: 0.44,
      fill: { color: i === 0 ? GOOD : ORANGE } });
    s.addText(d[0], { x: x2 + w2 - 0.6, y: y + 0.32, w: 0.44, h: 0.36, fontFace: F,
      fontSize: 14, bold: true, color: W, align: "center", margin: 0 });
    s.addText(d[1], { x: x2 + 0.14, y: y + 0.1, w: w2 - 0.82, h: 0.3, fontFace: F,
      fontSize: 12.5, bold: true, color: BGRAY, ...rtl });
    s.addText(d[2], { x: x2 + 0.14, y: y + 0.4, w: w2 - 0.82, h: 0.56, fontFace: F,
      fontSize: 10.5, color: INK2, ...rtl });
  });
  note(s, "المطلوب لإقرار الحملة: عرض سعر طباعة العلبة، وسعر المساحة الإعلانية عليها. عندها تُحسم جدواها برقم لا برأي.");
}

/* ═══ ٢٣ · مصفوفة الأحداث ═══ */
{
  const s = page("مصفوفة الأحداث — الحالة تختار الفعل",
                 "كل حالة تُرصد آلياً من بيانات المحطة، ولكل حالة فعل واحد محدد ومالك واحد");
  const cnt = { "فجوة مطابقة": 2, "انهيار حجم": 4, "بيانات دفع مفقودة": 18,
                "سلة دون المعيار": T.below, "ديزل مرتفع": 16, "وحدات شاغرة": 0 };
  ST.forEach(r => { if (r.conds.includes("وحدات شاغرة")) cnt["وحدات شاغرة"]++; });
  const key = ["فجوة مطابقة", "انهيار حجم", "بيانات دفع مفقودة", "سلة دون المعيار",
               "ديزل مرتفع", "تنافس", "وحدات شاغرة", "إمداد", "تأجير"];
  const rows = D.events.map((e, i) => {
    const n = cnt[Object.keys(cnt).find(k => e[0].includes(k.split(" ")[0])) ] ;
    const hi = e[4] === "hi";
    return { c: [{ t: e[0], a: "right", b: true, c: BGRAY }, { t: e[1], a: "right", c: hi ? D_BAD : INK },
                 { t: e[2], a: "right", fs: 10 }, { t: e[3], c: hi ? D_BAD : ORANGE, b: true }],
             fill: hi ? T_BAD : (i % 2 ? T_NEU : W) };
  });
  table(s, M, 1.55, CW, [
    { t: "الحالة المرصودة في البيانات", w: 22, a: "right" },
    { t: "الفعل", w: 20, a: "right" }, { t: "كيف يُنفَّذ", w: 42, a: "right" },
    { t: "التصنيف", w: 12 }], rows, { rh: 0.44, fs: 11, hfs: 11 });
  s.addShape(p.ShapeType.roundRect, { x: M, y: 6.05, w: CW, h: 0.85, rectRadius: 0.04,
    fill: { color: BGRAY } });
  s.addText("قاعدة الأولوية: الرقابة والبيانات قبل الحملة. محطة عليها فجوة مطابقة أو لا تسجّل وسيلة الدفع لا تدخل أي حملة قبل إقفال حالتها — وإلا أنفقنا على قياس لا نثق به.",
    { x: M + 0.2, y: 6.18, w: CW - 0.4, h: 0.6, fontFace: F, fontSize: 12.5, bold: true, color: W, ...rtl });
}

/* ═══ ٢٤ · التنفيذ ═══ */
{
  const s = page("خارطة ٩٠ يوماً", "ثلاث موجات — لا تبدأ الثانية قبل إقفال بوابة الأولى");
  const waves = [
    ["أول ٣٠ يوماً", "الإقفال", BAD, [
      "تحقيق MK008 وRY024 — جرد ومعايرة",
      "تفعيل تسجيل وسيلة الدفع في ١٨ محطة",
      "توحيد أسماء المحطات على الكود",
      "جدول عمالة لكل محطة ووردية",
      "جدول الإيجارات لكل نوع وحدة",
    ]],
    ["اليوم ٣١ إلى ٦٠", "الإطلاق المحدود", ORANGE, [
      "المستهدف في أعلى ١٤ محطة فقط",
      "علبة المناديل في ٤ محطات تجريبية",
      "عرض المعلن على العلبة — تعاقد أول",
      "تعاقد أساطيل في ٣ محاور طريق",
      "يوم مفتوح للمستأجرين في أعلى ٣ شواغر",
    ]],
    ["اليوم ٦١ إلى ٩٠", "التعميم المشروط", GOOD, [
      "تعميم ما أثبت رفع السلة ٤ أسابيع",
      "حافز العامل على الشبكة المؤهَّلة",
      "التأجير المسبق في تحت التنفيذ",
      "خطة الاعتراض في المعيصم والشرائع",
      "مراجعة المستهدف على الأداء الفعلي",
    ]],
  ];
  const cwid = (CW - 2 * 0.22) / 3;
  waves.forEach((w, i) => {
    const x = rtlx(i, cwid, 0.22);
    s.addShape(p.ShapeType.roundRect, { x, y: 1.55, w: cwid, h: 3.95, rectRadius: 0.05,
      fill: { color: W }, line: { color: w[2], width: 1.4 } });
    s.addShape(p.ShapeType.rect, { x, y: 1.55, w: cwid, h: 0.82, fill: { color: w[2] } });
    s.addText(w[0], { x: x + 0.1, y: 1.6, w: cwid - 0.2, h: 0.34, fontFace: F,
      fontSize: 11.5, color: W, align: "center", margin: 0 });
    s.addText(w[1], { x: x + 0.1, y: 1.94, w: cwid - 0.2, h: 0.38, fontFace: F,
      fontSize: 17, bold: true, color: W, align: "center", margin: 0 });
    bullets(s, x + 0.14, 2.55, cwid - 0.28, w[3], INK, 11.5);
  });
  s.addShape(p.ShapeType.roundRect, { x: M, y: 5.75, w: CW, h: 0.92, rectRadius: 0.04,
    fill: { color: T_OR }, line: { color: ORANGE, width: 1 } });
  s.addText("بوابة الموجة الثانية: لا حملة في محطة لم تُقفل حالتها الرقابية والبياناتية.   ·   بوابة الثالثة: لا تعميم لآلية لم ترفع لتر/زيارة قياساً بأربعة أسابيع سابقة في نفس المحطة.",
    { x: M + 0.2, y: 5.88, w: CW - 0.4, h: 0.66, fontFace: F, fontSize: 12.5, bold: true, color: INK, ...rtl });
}

/* ═══ ٢٥ · لوحة المتابعة ═══ */
{
  const s = page("لوحة المتابعة الشهرية", "ثمانية مؤشرات — تُقرأ في صفحة واحدة كل شهر");
  const kpis = [
    ["لتر لكل زيارة", "بالشريحة مقابل معيارها", "الأساسي"],
    ["الرفع المتحقق", "لتر/يوم مقابل +" + ar(T.upl), "الأساسي"],
    ["نسبة تسجيل وسيلة الدفع", "من ٦٤٪ إلى ٩٥٪", "بوابة"],
    ["فجوة المشتريات/المبيعات", "±١٪ لكل محطة", "بوابة"],
    ["الحافز ÷ الهامش الإضافي", "١٥٪ فأقل", "ضابط"],
    ["إشغال الوحدات", "بالفئة — مشغّلة وتحت تنفيذ وامتياز", "نمو"],
    ["حصة الأساطيل الرقمية", "من ١٫٩٪ صعوداً", "نمو"],
    ["تغطية الخزان", "أيام — تنبيه دون ٤", "مخاطر"],
  ];
  const cwid = (CW - 3 * 0.18) / 4;
  kpis.forEach((k, i) => {
    const x = rtlx(i % 4, cwid, 0.18), y = 1.6 + Math.floor(i / 4) * 1.55;
    const col = k[2] === "بوابة" ? BAD : k[2] === "ضابط" ? GOLD : k[2] === "مخاطر" ? BLUE : ORANGE;
    s.addShape(p.ShapeType.roundRect, { x, y, w: cwid, h: 1.32, rectRadius: 0.05,
      fill: { color: BG }, line: { color: LINE2, width: 1 } });
    s.addShape(p.ShapeType.rect, { x, y, w: cwid, h: 0.05, fill: { color: col } });
    s.addText(k[2], { x: x + 0.1, y: y + 0.12, w: cwid - 0.2, h: 0.26, fontFace: F,
      fontSize: 9.5, bold: true, color: col, align: "center", margin: 0 });
    s.addText(k[0], { x: x + 0.1, y: y + 0.42, w: cwid - 0.2, h: 0.4, fontFace: F,
      fontSize: 12.5, bold: true, color: BGRAY, align: "center", margin: 0 });
    s.addText(k[1], { x: x + 0.1, y: y + 0.86, w: cwid - 0.2, h: 0.38, fontFace: F,
      fontSize: 10, color: INK2, align: "center", margin: 0 });
  });
  s.addText("ما نحتاجه من الإدارات لإكمال الخطة", { x: M, y: 4.85, w: CW, h: 0.34,
    fontFace: F, fontSize: 15, bold: true, color: ORANGE, ...rtl });
  const need = [
    ["العمليات", "جدول عمالة كل محطة ووردية — لتعميم مستهدف العامل"],
    ["العقار", "جدول الإيجار الحالي والمستهدف لكل نوع وحدة — لتسعير ٦٨٩ وحدة و٥٧٢ محلاً"],
    ["المالية", "هامش المحطة الفعلي لكل محطة — الحافز يُحتسب عليه"],
    ["التسويق", "عرض طباعة علبة المناديل وسعر مساحتها الإعلانية"],
  ];
  need.forEach((n, i) => {
    const y = 5.25 + i * 0.42;
    s.addShape(p.ShapeType.rect, { x: SW - M - 0.04, y: y + 0.04, w: 0.04, h: 0.32, fill: { color: ORANGE } });
    s.addText(n[0], { x: SW - M - 1.85, y, w: 1.7, h: 0.36, fontFace: F, fontSize: 12,
      bold: true, color: BGRAY, ...rtl });
    s.addText(n[1], { x: M, y, w: CW - 2.0, h: 0.36, fontFace: F, fontSize: 11.5, color: INK, ...rtl });
  });
}

/* ═══ ٢٦ · الختام ═══ */
{
  const s = cover("الإدارة التجارية · درب", "القرار المطلوب",
    "اعتماد التقسيم والمستهدف · إقفال ثغرة البيانات قبل أي إنفاق · تمويل الحملات من المعلن لا من هامش الوقود",
    "الفرصة المحسوبة: " + ar(T.sar) + " ريال سنوياً من رفع اللتر لكل زيارة في " + ar(T.below) + " محطة");
}

p.writeFile({ fileName: "docs/عرض-تحليل-المنافذ-وخطة-المبيعات.pptx" })
 .then(f => console.log("✓ " + f + "  ·  " + p.slides.length + " شريحة"));
