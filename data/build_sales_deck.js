/* عرض الإدارة التنفيذية — تحليل المنافذ والمنافسين وخطة المبيعات
   يُشغَّل من جذر المستودع:  node data/build_sales_deck.js
   المصدر: data/sales-plan.json (PYTHONPATH=data python3 data/sales_plan.py)
   القاعدة التحريرية: رقم وسطر معنى — لا فقرات شارحة. */
const fs = require("fs");
const pptx = require("pptxgenjs");
const D = JSON.parse(fs.readFileSync("data/sales-plan.json", "utf8"));

const p = new pptx();
p.layout = "LAYOUT_WIDE";
p.rtlMode = true;
p.author = "الإدارة التجارية — درب";
p.title = "تحليل المنافذ وخطة المبيعات";

/* ── هوية درب ── */
const ORANGE = "F5831F", GOLD = "F7A94B", BGRAY = "55565A",
      INK = "3D3D3D", INK2 = "6E6A64", INK3 = "9B968E",
      BG = "F7F4EF", W = "FFFFFF", LINE2 = "E3DCD1",
      GOOD = "2E8B6F", BAD = "C0503A", BLUE = "3E6E8E",
      T_OR = "FBEEE0", T_GOOD = "D6E9E1", T_BAD = "F2DAD4",
      T_NEU = "F5F2ED", T_BAND = "EDE7DE",
      D_GOOD = "1F5E4A", D_BAD = "9A3E2C", D_GOLD = "A8681A";
const F = "DIN Next Arabic";
const SW = 13.3, SH = 7.5, M = 0.62, CW = SW - 2 * M;
const rtl = { align: "right", rtlMode: true };
const rtlx = (i, cwid, gap) => M + CW - cwid - i * (cwid + gap);

const ar = n => n.toLocaleString("ar-EG", { maximumFractionDigits: 0 });
const ar1 = n => n.toLocaleString("ar-EG", { minimumFractionDigits: 1, maximumFractionDigits: 1 });
const ar2 = n => n.toLocaleString("ar-EG", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const pc = n => ar1(n * 100) + "٪";
const pc0 = n => ar(Math.round(n * 100)) + "٪";
const arn = t => String(t).replace(/[0-9]/g, d => "٠١٢٣٤٥٦٧٨٩"[d]).replace(/,/g, "٬");

/* ── هياكل ── */
function cover(kicker, title, sub, foot) {
  const s = p.addSlide();
  s.background = { color: BGRAY };
  s.addShape(p.ShapeType.rect, { x: 0, y: 0, w: SW, h: 0.11, fill: { color: ORANGE } });
  s.addText("درب", { x: SW - M - 2.2, y: 0.5, w: 2.2, h: 0.5, fontFace: F,
    fontSize: 26, bold: true, color: W, align: "right", margin: 0 });
  s.addText("DARB FUEL", { x: SW - M - 2.2, y: 1.0, w: 2.2, h: 0.28, fontFace: F,
    fontSize: 10, color: ORANGE, align: "right", charSpacing: 2, margin: 0 });
  if (kicker) s.addText(kicker, { x: M, y: 2.85, w: CW, h: 0.35, fontFace: F,
    fontSize: 13, color: GOLD, charSpacing: 3, ...rtl });
  s.addText(title, { x: M, y: 3.25, w: CW, h: 1.05, fontFace: F, fontSize: 40,
    bold: true, color: W, ...rtl });
  if (sub) s.addText(sub, { x: M, y: 4.35, w: CW, h: 0.7, fontFace: F,
    fontSize: 15, color: "D6D2CC", ...rtl });
  s.addShape(p.ShapeType.rect, { x: SW - M - 3.4, y: 5.4, w: 3.4, h: 0.035, fill: { color: ORANGE } });
  if (foot) s.addText(foot, { x: M, y: 5.6, w: CW, h: 0.5, fontFace: F,
    fontSize: 11, color: INK3, ...rtl });
  return s;
}

function divider(no, title, sub) {
  const s = p.addSlide();
  s.background = { color: BG };
  s.addShape(p.ShapeType.rect, { x: SW - M - 0.06, y: 2.5, w: 0.06, h: 2.2, fill: { color: ORANGE } });
  s.addText(no, { x: M, y: 2.45, w: CW - 0.3, h: 0.5, fontFace: F, fontSize: 15,
    bold: true, color: ORANGE, charSpacing: 3, ...rtl });
  s.addText(title, { x: M, y: 3.0, w: CW - 0.3, h: 0.9, fontFace: F, fontSize: 36,
    bold: true, color: BGRAY, ...rtl });
  if (sub) s.addText(sub, { x: M, y: 3.95, w: CW - 0.3, h: 0.5, fontFace: F,
    fontSize: 14, color: INK2, ...rtl });
  return s;
}

let PNO = 0;
function page(title, sub) {
  const s = p.addSlide();
  PNO++;
  s.background = { color: W };
  s.addShape(p.ShapeType.rect, { x: 0, y: 0, w: SW, h: 0.075, fill: { color: ORANGE } });
  s.addText(title, { x: M, y: 0.3, w: CW - 1.0, h: 0.54, fontFace: F, fontSize: 26,
    bold: true, color: BGRAY, ...rtl });
  if (sub) s.addText(sub, { x: M, y: 0.86, w: CW - 1.0, h: 0.3, fontFace: F,
    fontSize: 12, color: INK2, ...rtl });
  s.addShape(p.ShapeType.rect, { x: M, y: 1.24, w: CW, h: 0.02, fill: { color: LINE2 } });
  s.addText("درب · الإدارة التجارية", { x: M, y: 7.04, w: 4, h: 0.28, fontFace: F,
    fontSize: 9, color: INK3, align: "left", margin: 0 });
  s.addText(String(PNO), { x: SW - M - 0.6, y: 7.04, w: 0.6, h: 0.28, fontFace: F,
    fontSize: 9, color: INK3, align: "right", margin: 0 });
  return s;
}

function table(s, x, y, w, cols, rows, opt) {
  const o = Object.assign({ hh: 0.4, rh: 0.34, fs: 11, hfs: 10.5, zebra: true }, opt || {});
  const tot = cols.reduce((a, c) => a + c.w, 0);
  const xs = []; let acc = 0;
  cols.forEach(c => { acc += c.w; xs.push(x + w - (acc / tot) * w); });
  const cw = cols.map(c => (c.w / tot) * w);
  s.addShape(p.ShapeType.rect, { x, y, w, h: o.hh, fill: { color: BGRAY } });
  cols.forEach((c, i) => s.addText(c.t, { x: xs[i], y, w: cw[i], h: o.hh,
    fontFace: F, fontSize: o.hfs, bold: true, color: W, valign: "middle",
    align: c.a || "center", rtlMode: true, margin: [0, 0.06, 0, 0.06] }));
  rows.forEach((raw, ri) => {
    const r = Array.isArray(raw) ? { c: raw } : raw;
    const yy = y + o.hh + ri * o.rh;
    s.addShape(p.ShapeType.rect, { x, y: yy, w, h: o.rh,
      fill: { color: r.fill || (o.zebra && ri % 2 ? T_NEU : W) },
      line: { color: LINE2, width: 0.6 } });
    r.c.forEach((v, i) => {
      const cell = (v && typeof v === "object" && !Array.isArray(v)) ? v : { t: v };
      s.addText(String(cell.t), { x: xs[i], y: yy, w: cw[i], h: o.rh, fontFace: F,
        fontSize: cell.fs || o.fs, bold: !!cell.b || !!r.b, color: cell.c || INK,
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
  s.addText(num, { x: x + 0.08, y: y + 0.16, w: w - 0.16, h: 0.62, fontFace: F,
    fontSize: 28, bold: true, color: col || BGRAY, align: "center", margin: 0 });
  s.addText(lbl, { x: x + 0.1, y: y + 0.8, w: w - 0.2, h: 0.28, fontFace: F,
    fontSize: 11, color: INK, align: "center", margin: 0 });
  if (note) s.addText(note, { x: x + 0.1, y: y + 1.1, w: w - 0.2, h: 0.3,
    fontFace: F, fontSize: 9, color: INK3, align: "center", margin: 0 });
}

function band(s, y, t, col) {
  s.addShape(p.ShapeType.roundRect, { x: M, y, w: CW, h: 0.56, rectRadius: 0.04,
    fill: { color: col || BGRAY } });
  s.addText(t, { x: M + 0.2, y: y + 0.05, w: CW - 0.4, h: 0.46, fontFace: F,
    fontSize: 12.5, bold: true, color: W, valign: "middle", ...rtl });
}

function head(s, y, t) {
  s.addText(t, { x: M, y, w: CW, h: 0.3, fontFace: F, fontSize: 13.5,
    bold: true, color: ORANGE, ...rtl });
}

function foot(s, t) {
  s.addText(t, { x: M, y: 6.72, w: CW, h: 0.28, fontFace: F, fontSize: 9.5,
    color: INK3, ...rtl });
}

const N = D.network, T = D.totals, SEG = D.segments, ST = D.stations, L = D.litre;

/* ═══ ١ · الغلاف ═══ */
cover("الإدارة التجارية · أغسطس ٢٠٢٦", "تحليل المنافذ وخطة المبيعات",
      "٥٥ محطة · ٥٥١ MLPA · خمس شرائح · مستهدف لكل محطة ووردية · باقات تأجير وحملات تُطلقها حالة مرصودة",
      "كاش إن وناتج ١٠٫٩٦ مليون معاملة · أوامر التحميل · قائمة الدخل يوليو ٢٠٢٦ · سجل الوحدات · مسح المنافسين الميداني");

/* ═══ ٢ · الخلاصة ═══ */
{
  const s = page("الخلاصة التنفيذية", "ما تقوله البيانات — وما نطلب إقراره");
  const d = [[ar(Math.round(N.revenue / 1e6)), "مليون ريال", "مبيعات ٧ أشهر", ORANGE],
             [ar(Math.round(T.mlpa)), "MLPA", "١٠٫٦ للمحطة", BGRAY],
             ["+١٧٫٥٪", "نمو حقيقي", "قاعدة ثابتة ٢٥ محطة", GOOD],
             [ar1(T.sar / 1e6), "مليون ريال", "الفرصة المؤكَّدة", GOOD]];
  const cw = (CW - 3 * 0.18) / 4;
  d.forEach((v, i) => stat(s, rtlx(i, cw, 0.18), 1.38, cw, 1.52, v[0], v[1], v[3], v[2]));
  head(s, 3.2, "ثلاثة قرارات");
  const dec = [["اعتماد المستهدف برافعتين",
                "المعاملات ثم السلة — لا نطلب من العميل أن يعبّي أكثر مما يحتاج",
                ar(Math.round(T.sar / 1000)) + " ألف ر"],
               ["إقفال ثغرة البيانات قبل الإنفاق", "١٨ محطة لا تسجّل وسيلة الدفع", "١٨٣ مليون ر"],
               ["مسح المنافسين في ثلاثة منتجات", "العقار والإكسسوارات والإعلان", "٣ من ٥"]];
  dec.forEach((v, i) => {
    const y = 3.6 + i * 0.88;
    s.addShape(p.ShapeType.roundRect, { x: M, y, w: CW, h: 0.78, rectRadius: 0.04,
      fill: { color: i === 0 ? T_OR : BG }, line: { color: LINE2, width: 1 } });
    s.addShape(p.ShapeType.rect, { x: SW - M - 0.055, y, w: 0.055, h: 0.78, fill: { color: ORANGE } });
    s.addText(String(i + 1), { x: SW - M - 0.62, y: y + 0.17, w: 0.5, h: 0.44,
      fontFace: F, fontSize: 17, bold: true, color: ORANGE, align: "center", margin: 0 });
    s.addText(v[0], { x: M + 2.3, y: y + 0.08, w: CW - 2.95, h: 0.32, fontFace: F,
      fontSize: 14, bold: true, color: BGRAY, ...rtl });
    s.addText(v[1], { x: M + 2.3, y: y + 0.41, w: CW - 2.95, h: 0.3, fontFace: F,
      fontSize: 11, color: INK2, ...rtl });
    s.addText(v[2], { x: M + 0.12, y: y + 0.2, w: 2.1, h: 0.4, fontFace: F,
      fontSize: 13, bold: true, color: ORANGE, align: "left", margin: 0 });
  });
  foot(s, "بيانات نقاط بيع فعلية من ١ يناير إلى ٣١ يوليو ٢٠٢٦");
}

/* ═══ ٣ · القسم الأول ═══ */
divider("القسم الأول", "أين نقف", "الشبكة · الاتجاه · قيمة اللتر · من يشتري");

/* ═══ ٤ · الشبكة ═══ */
{
  const s = page("الشبكة بالأرقام", "من ١ يناير إلى ٣١ يوليو ٢٠٢٦");
  const cw = (CW - 4 * 0.16) / 5;
  const d = [[ar(Math.round(N.revenue / 1e6)), "مليون ريال", "شامل الضريبة", ORANGE],
             [ar(Math.round(T.mlpa)), "MLPA", "١٠٫٦ للمحطة", BGRAY],
             [ar1(N.visits / 1e6), "مليون معاملة", "٢٥٫٦ لتراً لكل واحدة", BGRAY],
             [ar1(N.inv), "ريال", "قيمة المعاملة", BGRAY],
             [ar(N.stations), "محطة", "٨٣٪ تغطية زمنية", BGRAY]];
  d.forEach((v, i) => stat(s, rtlx(i, cw, 0.16), 1.36, cw, 1.5, v[0], v[1], v[3], v[2]));
  head(s, 3.1, "مزيج الوقود");
  const fu = [["بنزين ٩١", 0.438, 0.555, ORANGE], ["بنزين ٩٥", 0.374, 0.360, GOLD],
              ["ديزل", 0.188, 0.085, BLUE]];
  const bwd = CW * 0.58;
  fu.forEach((f, i) => {
    const y = 3.5 + i * 0.74;
    s.addText(f[0], { x: M + bwd + 0.18, y, w: 1.6, h: 0.3, fontFace: F,
      fontSize: 12, bold: true, color: INK, ...rtl });
    s.addShape(p.ShapeType.rect, { x: M, y: y + 0.32, w: bwd, h: 0.2, fill: { color: T_NEU } });
    s.addShape(p.ShapeType.rect, { x: M + bwd * (1 - f[1]), y: y + 0.32, w: bwd * f[1],
      h: 0.2, fill: { color: f[3] } });
    s.addText(pc(f[1]) + " من الإيراد · " + pc(f[2]) + " من المعاملات",
      { x: M, y, w: bwd * 0.75, h: 0.3, fontFace: F, fontSize: 10.5, color: INK2,
        align: "left", margin: 0 });
  });
  const x2 = M + CW * 0.66, w2 = CW * 0.34;
  s.addShape(p.ShapeType.roundRect, { x: x2, y: 3.46, w: w2, h: 2.2, rectRadius: 0.05,
    fill: { color: T_OR }, line: { color: ORANGE, width: 1.2 } });
  s.addText("الديزل عميل مختلف", { x: x2 + 0.16, y: 3.58, w: w2 - 0.32, h: 0.32,
    fontFace: F, fontSize: 14, bold: true, color: ORANGE, ...rtl });
  [["٨٫٥٪", "من المعاملات"], ["١٨٫٨٪", "من الإيراد"], ["١٢٦ ريالاً", "قيمة المعاملة — ٢٫٨ ضعف"]]
    .forEach((r, i) => {
      const y = 4.05 + i * 0.5;
      s.addText(r[0], { x: x2 + w2 - 1.55, y, w: 1.4, h: 0.34, fontFace: F,
        fontSize: 14, bold: true, color: BGRAY, ...rtl });
      s.addText(r[1], { x: x2 + 0.16, y: y + 0.02, w: w2 - 1.8, h: 0.32, fontFace: F,
        fontSize: 11, color: INK, align: "left", margin: 0 });
    });
  foot(s, "٣ محطات لتراتها ناقصة في المصدر فاستُبعدت من حسابات اللتر");
}

/* ═══ ٥ · الاتجاه ═══ */
{
  const s = page("الاتجاه الحقيقي", "على قاعدة ثابتة من ٢٥ محطة بتغطية كاملة ٢١٢ يوماً");
  const base = [633296, 624536, 654588, 635997, 689121, 713904, 743943];
  const mn = ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو"];
  const w0 = CW * 0.62, y0 = 1.72, h0 = 3.4, mx = 780000, mi = 580000;
  base.forEach((v, i) => {
    const bw = w0 / 7 * 0.62, bx = M + w0 - (i + 1) * (w0 / 7) + (w0 / 7 - bw) / 2;
    const bh = (v - mi) / (mx - mi) * h0;
    s.addShape(p.ShapeType.rect, { x: bx, y: y0 + h0 - bh, w: bw, h: bh,
      fill: { color: i === 6 ? ORANGE : (i >= 4 ? GOLD : BGRAY) } });
    s.addText(ar(v), { x: bx - 0.15, y: y0 + h0 - bh - 0.3, w: bw + 0.3, h: 0.26,
      fontFace: F, fontSize: 10, bold: true, color: i === 6 ? ORANGE : INK2,
      align: "center", margin: 0 });
    s.addText(mn[i], { x: bx - 0.15, y: y0 + h0 + 0.05, w: bw + 0.3, h: 0.26,
      fontFace: F, fontSize: 10.5, color: INK, align: "center", margin: 0 });
  });
  s.addShape(p.ShapeType.rect, { x: M, y: y0 + h0, w: w0, h: 0.02, fill: { color: LINE2 } });
  s.addText("لتراً يومياً", { x: M, y: 1.38, w: w0, h: 0.28, fontFace: F,
    fontSize: 11, color: INK2, ...rtl });
  const x2 = M + CW * 0.67, w2 = CW * 0.33;
  stat(s, x2, 1.4, w2, 1.5, "+١٧٫٥٪", "يناير ← يوليو", GOOD, "على نفس المحطات");
  s.addText("ومحطتان تسيران عكسها", { x: x2, y: 3.12, w: w2, h: 0.3, fontFace: F,
    fontSize: 13, bold: true, color: BAD, ...rtl });
  const mk = [["المعيصم MK002", "−٧٠٪", "تغيير مسار ٧–٨ فبراير"],
              ["الشرايع MK019", "−٥٢٪", "تغيير مسار ٤–٥ فبراير"]];
  mk.forEach((v, i) => {
    const y = 3.5 + i * 0.74;
    s.addShape(p.ShapeType.roundRect, { x: x2, y, w: w2, h: 0.66, rectRadius: 0.04,
      fill: { color: T_BAD }, line: { color: BAD, width: 1 } });
    s.addText(v[0], { x: x2 + 0.12, y: y + 0.04, w: w2 - 1.1, h: 0.3, fontFace: F,
      fontSize: 12, bold: true, color: BGRAY, ...rtl });
    s.addText(v[1], { x: x2 + w2 - 1.0, y: y + 0.04, w: 0.88, h: 0.3, fontFace: F,
      fontSize: 13, bold: true, color: D_BAD, align: "left", margin: 0 });
    s.addText(v[2], { x: x2 + 0.12, y: y + 0.34, w: w2 - 0.24, h: 0.28, fontFace: F,
      fontSize: 10, color: INK2, ...rtl });
  });
  s.addShape(p.ShapeType.roundRect, { x: x2, y: 5.06, w: w2, h: 0.8, rectRadius: 0.04,
    fill: { color: BG }, line: { color: LINE2, width: 1 } });
  s.addText("٤ هابطة · ٨ مستقرة · ٩ نامية", { x: x2 + 0.12, y: 5.16, w: w2 - 0.24, h: 0.3,
    fontFace: F, fontSize: 12, bold: true, color: BGRAY, align: "center", margin: 0 });
  s.addText("من ٢١ محطة بتغطية كاملة", { x: x2 + 0.12, y: 5.48, w: w2 - 0.24, h: 0.28,
    fontFace: F, fontSize: 10, color: INK3, align: "center", margin: 0 });
  foot(s, "الإجمالي الخام مضلِّل — المحطات المغطّاة ارتفعت من ٣٩ إلى ٥٥ خلال الفترة");
}

/* ═══ ٦ · قيمة اللتر ═══ */
{
  const s = page("قيمة اللتر الواحد", "من سعر المضخة إلى الصافي — وما تحتمله الحملة");
  band(s, 1.34, "سعر المضخة  ←  ناقص الضريبة ١٥٪  ←  هامش المساهمة ١١٫٤٤ هللة  ←  ناقص التشغيل ٤٫٣٤  =  صافي ٧٫١٠ هللة لكل لتر");
  head(s, 2.1, "تعبئة بخمسين ريالاً — العتبة المعتمدة");
  const rows = L.rows.map((r, i) => ({
    c: [{ t: r.fuel, a: "right", b: true }, ar2(r.price), ar2(r.net_price),
        { t: ar2(r.litres), b: true }, ar2(r.margin), ar2(r.opex),
        { t: ar2(r.net), b: true, c: D_GOOD }, ar2(r.box),
        { t: ar2(r.after), b: true, c: D_GOLD }, pc0(r.box / r.net)],
    fill: i % 2 ? T_NEU : W }));
  rows.push({ c: [{ t: "المرجّح", b: true }, ar2(50 / L.w_litres), ar2(50 / L.w_litres / 1.15),
                  ar2(L.w_litres), ar2(L.w_litres * D.margin), ar2(L.w_litres * D.opex),
                  ar2(L.w_net), ar2(L.box), ar2(L.w_net - L.box), pc0(L.box / L.w_net)],
              fill: T_BAND, b: true });
  table(s, M, 2.42, CW, [
    { t: "المنتج", w: 12, a: "right" }, { t: "سعر المضخة", w: 10 }, { t: "صافي الضريبة", w: 10 },
    { t: "لتر بـ٥٠ ريالاً", w: 11 }, { t: "هامش المساهمة", w: 11 }, { t: "خصم التشغيل", w: 10 },
    { t: "الصافي", w: 9 }, { t: "علبة المناديل", w: 10 }, { t: "بعد الهدية", w: 9 },
    { t: "الهدية من الصافي", w: 12 }], rows, { rh: 0.4, fs: 11, hfs: 9.5 });
  const cw = (CW - 2 * 0.2) / 3;
  const cards = [["٤٥ هللة", "كلفة العلبة", "مؤكَّدة من الإدارة", ORANGE],
                 ["٦٫٣ لتر", "تعادل الهدية", "١٣٫٩ ريالاً إنفاقاً إضافياً", GOOD],
                 ["٥٧٫١٨ ريالاً", "قيمة المعاملة اليوم", "فوق العتبة أصلاً", BAD]];
  cards.forEach((c, i) => {
    const x = rtlx(i, cw, 0.2);
    s.addShape(p.ShapeType.roundRect, { x, y: 4.75, w: cw, h: 1.3, rectRadius: 0.05,
      fill: { color: c[3] === BAD ? T_BAD : BG }, line: { color: c[3], width: 1.2 } });
    s.addText(c[0], { x: x + 0.1, y: 4.87, w: cw - 0.2, h: 0.44, fontFace: F,
      fontSize: 21, bold: true, color: c[3], align: "center", margin: 0 });
    s.addText(c[1], { x: x + 0.1, y: 5.32, w: cw - 0.2, h: 0.3, fontFace: F,
      fontSize: 12, bold: true, color: BGRAY, align: "center", margin: 0 });
    s.addText(c[2], { x: x + 0.1, y: 5.63, w: cw - 0.2, h: 0.34, fontFace: F,
      fontSize: 10, color: INK2, align: "center", margin: 0 });
  });
  foot(s, "خصم التشغيل على أساس الشبكة كلها · لو كان على المشغّلة وحدها فهو ٩٫٢٤ هللة والصافي ٢٫٢٠ — سؤال معلَّق للمالية");
}

/* ═══ ٧ · تقسيم السوق ═══ */
{
  const s = page("من يشتري منّا", "أربعة عملاء داخل نفس المحطة");
  const seg = [["أفراد ٩١", "٥٥٫٥٪ من المعاملات", "٤٥٫١ ريالاً", "الأكبر عدداً والأصغر سلة", ORANGE],
               ["أفراد ٩٥", "٣٦٫٠٪", "٥٩٫٥ ريالاً", "قدرة إنفاق أعلى", GOLD],
               ["أساطيل الديزل", "٨٫٥٪", "١٢٦٫٣ ريالاً", "١٩٪ من الإيراد — بالتعاقد لا بالحملة", BLUE],
               ["أساطيل رقمية", "٠٫٧٪", "١٤٩–١٧٢ ريالاً", "أعلى فاتورة وأقل حصة", GOOD]];
  const cw = (CW - 3 * 0.2) / 4;
  seg.forEach((v, i) => {
    const x = rtlx(i, cw, 0.2);
    s.addShape(p.ShapeType.roundRect, { x, y: 1.36, w: cw, h: 2.7, rectRadius: 0.05,
      fill: { color: W }, line: { color: LINE2, width: 1 } });
    s.addShape(p.ShapeType.rect, { x, y: 1.36, w: cw, h: 0.54, fill: { color: v[4] } });
    s.addText(v[0], { x: x + 0.08, y: 1.42, w: cw - 0.16, h: 0.42, fontFace: F,
      fontSize: 15, bold: true, color: W, align: "center", margin: 0 });
    s.addText(v[1], { x: x + 0.1, y: 2.0, w: cw - 0.2, h: 0.28, fontFace: F,
      fontSize: 11.5, color: INK2, align: "center", margin: 0 });
    s.addText(v[2], { x: x + 0.1, y: 2.34, w: cw - 0.2, h: 0.46, fontFace: F,
      fontSize: 20, bold: true, color: v[4], align: "center", margin: 0 });
    s.addText("قيمة المعاملة", { x: x + 0.1, y: 2.8, w: cw - 0.2, h: 0.26, fontFace: F,
      fontSize: 9.5, color: INK3, align: "center", margin: 0 });
    s.addText(v[3], { x: x + 0.12, y: 3.14, w: cw - 0.24, h: 0.8, fontFace: F,
      fontSize: 11, color: INK, ...rtl });
  });
  s.addShape(p.ShapeType.roundRect, { x: M, y: 4.3, w: CW, h: 1.3, rectRadius: 0.05,
    fill: { color: T_BAD }, line: { color: BAD, width: 1.2 } });
  s.addText("٣٦٪ من الإيراد بلا هوية عميل", { x: M + 0.2, y: 4.42, w: CW - 0.4, h: 0.34,
    fontFace: F, fontSize: 15, bold: true, color: D_BAD, ...rtl });
  s.addText("٢٢٦ مليون ريال «وسيلة دفع غير محددة» · ١٨ محطة لا تسجّلها إطلاقاً بإيراد ١٨٣ مليوناً · التفعيل يسبق الإنفاق",
    { x: M + 0.2, y: 4.8, w: CW - 0.4, h: 0.7, fontFace: F, fontSize: 12.5, color: INK, ...rtl });
  foot(s, "محفظة كاش إن: ٣١٨ ريالاً و٨ عمليات في سبعة أشهر — قناة غير مفعّلة");
}

/* ═══ وقود الشركات — خطة التعاقد ═══ */
{
  const s = page("وقود الشركات — خطة التعاقد", "لكل محطة قائمة صيد · وعملتان للدفع: هللة أو خدمة");
  const FL = D.fleet;
  band(s, 1.32, "٨١ مليون ريال ديزل يُشترى نقداً أو بالبطاقة — ٤٥٫٢ مليون لتر بلا تعاقد · هامشها ٥٬١٦٩٬٠٦٠ ريال");
  s.addText("① قائمة الصيد — أعلى ١٠ محطات ديزلاً", { x: M, y: 2.06, w: CW * 0.56,
    h: 0.3, fontFace: F, fontSize: 13.5, bold: true, color: ORANGE, ...rtl });
  const rows = FL.targets.map((r, i) => ({
    c: [{ t: r.name.slice(0, 19), a: "right", b: true }, r.code,
        { t: ar(r.dl), b: true }, pc0(r.dsh), pc(r.fl),
        { t: r.svc ? ar(r.svc) : "لا يوجد", b: !r.svc, c: r.svc ? INK : D_BAD },
        { t: r.svc ? (r.fl < 0.05 ? "خدمات" : "خدمات + هللة") : "هللة فقط",
          b: true, c: r.svc ? D_GOOD : D_BAD }],
    fill: r.svc ? (i % 2 ? T_NEU : W) : T_BAD }));
  table(s, M, 2.38, CW * 0.56, [
    { t: "المحطة", w: 26, a: "right" }, { t: "الكود", w: 11 },
    { t: "ديزل لتر/يوم", w: 15 }, { t: "حصته", w: 10 },
    { t: "الدفع المؤسسي", w: 13 }, { t: "وحدات خدمة", w: 12 },
    { t: "العملة الممكنة", w: 17 }], rows, { rh: 0.335, fs: 10, hfs: 9.5 });

  const x2 = M + CW * 0.59, w2 = CW * 0.41;
  s.addText("② بأي عملة ندفع", { x: x2, y: 2.06, w: w2, h: 0.3, fontFace: F,
    fontSize: 13.5, bold: true, color: ORANGE, ...rtl });
  const lr = FL.ladder.map((l, i) => ({
    c: [{ t: ar(l.cut) + " هللة", a: "right", b: true },
        { t: ar2(l.left) + " هللة", b: true, c: l.share > 0.5 ? D_BAD : D_GOOD },
        { t: pc0(l.share), c: l.share > 0.5 ? D_BAD : INK }],
    fill: i % 2 ? T_NEU : W }));
  table(s, x2, 2.38, w2, [{ t: "الخصم", w: 34, a: "right" },
    { t: "يبقى من الصافي", w: 36 }, { t: "ما نتنازل عنه", w: 30 }],
    lr, { rh: 0.34, fs: 10.5, hfs: 10 });

  const S = FL.sample;
  s.addText("أسطول نموذجي: ٢٠ شاحنة · ١٦٬٠٠٠ لتر شهرياً · صافيه لنا ١٬١٣٦ ريالاً",
    { x: x2, y: 4.14, w: w2, h: 0.28, fontFace: F, fontSize: 10.5, color: INK2, ...rtl });
  const cur = [["خصم ٢ هللة", ar(S.disc_cost), ar(S.disc_face), "١٫٠٠", BGRAY],
               ["٢٠ غسلة", ar(S.svc_cost), ar(S.svc_face), "١٫٨٧", GOOD]];
  const cw = (w2 - 0.18) / 2;
  cur.forEach((c, i) => {
    const x = x2 + w2 - cw - i * (cw + 0.18);
    s.addShape(p.ShapeType.roundRect, { x, y: 4.46, w: cw, h: 1.52, rectRadius: 0.05,
      fill: { color: c[4] === GOOD ? T_GOOD : BG }, line: { color: c[4], width: 1.2 } });
    s.addText(c[0], { x: x + 0.08, y: 4.53, w: cw - 0.16, h: 0.3, fontFace: F,
      fontSize: 13, bold: true, color: c[4] === GOOD ? D_GOOD : BGRAY,
      align: "center", margin: 0 });
    [["كلفتها علينا", c[1] + " ر"], ["قيمتها للعميل", c[2] + " ر"],
     ["لكل ريال ننفقه", c[3] + " ر قيمة"]].forEach((r, k) => {
      const y = 4.86 + k * 0.36;
      s.addText(r[0], { x: x + 0.08, y, w: cw - 0.16, h: 0.18, fontFace: F,
        fontSize: 9, color: INK3, align: "center", margin: 0 });
      s.addText(r[1], { x: x + 0.08, y: y + 0.15, w: cw - 0.16, h: 0.22, fontFace: F,
        fontSize: 12, bold: true, color: c[4] === GOOD ? D_GOOD : BGRAY,
        align: "center", margin: 0 });
    });
  });
  band(s, 6.1, "الخدمة تعطي ١٫٨٧ ريال قيمة لكل ريال ننفقه — وتملأ وحدات الخدمة الشاغرة · لكن ٣ من أعلى ١٠ محطات ديزلاً بلا وحدة خدمة أصلاً", ORANGE);
  foot(s, "غسلة بكلفة ١٥ ريالاً تعادل ٢١١ لتراً من الصافي — فالسقف غسلة لكل ألف لتر");
}

/* ═══ ٨ · القسم الثاني ═══ */
divider("القسم الثاني", "المنافسون", "خمسة منتجات — ولكلٍّ مسيطر مختلف");

/* ═══ ٩ · المنافسون حسب المنتج ═══ */
{
  const s = page("المنافسون حسب المنتج", "من يسيطر على كل منتج — وأين لا نعرف");
  const rows = D.products.map((r, i) => {
    const gap = r[5] === "مفقود";
    return { c: [{ t: r[0], a: "right", b: true, c: BGRAY }, { t: r[1], b: true },
                 r[2], { t: r[3], b: true, c: gap ? D_BAD : BLUE },
                 { t: r[4], a: "right" }, { t: r[5], b: true, c: gap ? D_BAD : D_GOOD }],
             fill: gap ? T_BAD : (i % 2 ? T_NEU : W) };
  });
  table(s, M, 1.36, CW, [
    { t: "المنتج", w: 15, a: "right" }, { t: "حجمنا", w: 17 }, { t: "موقعنا", w: 14 },
    { t: "المسيطر", w: 13 }, { t: "ما نعرفه عنه", w: 27, a: "right" },
    { t: "حالة المسح", w: 14 }], rows, { rh: 0.56, fs: 11, hfs: 10.5 });
  const cw = (CW - 2 * 0.2) / 3;
  const cards = [["٢ من ٥", "منتجات تعمل", "الوقود بشقّيه — والباقي لم يبدأ", ORANGE],
                 ["١ من ٥", "مسح منافسين مكتمل", "وقود الأفراد على خمسة مواقع", GOLD],
                 ["٣ من ٥", "بلا معرفة تنافسية", "العقار والإكسسوارات والإعلان", BAD]];
  cards.forEach((c, i) => {
    const x = rtlx(i, cw, 0.2);
    s.addShape(p.ShapeType.roundRect, { x, y: 4.72, w: cw, h: 1.32, rectRadius: 0.05,
      fill: { color: c[3] === BAD ? T_BAD : BG }, line: { color: c[3], width: 1.2 } });
    s.addText(c[0], { x: x + 0.1, y: 4.84, w: cw - 0.2, h: 0.44, fontFace: F,
      fontSize: 22, bold: true, color: c[3], align: "center", margin: 0 });
    s.addText(c[1], { x: x + 0.1, y: 5.3, w: cw - 0.2, h: 0.3, fontFace: F,
      fontSize: 12, bold: true, color: BGRAY, align: "center", margin: 0 });
    s.addText(c[2], { x: x + 0.1, y: 5.62, w: cw - 0.2, h: 0.34, fontFace: F,
      fontSize: 10, color: INK2, align: "center", margin: 0 });
  });
  foot(s, "لم تُذكر أسماء منافسين في المنتجات الثلاثة لأننا لا نملك مسحاً — والفراغ مقصود لا سهو");
}

/* ═══ ١٠ · منافسة الوقود ═══ */
{
  const s = page("وقود الأفراد — الموقف الميداني", "خمسة مواقع · ٤٠ منافساً مرصوداً");
  const rows = D.five.map((f, i) => ({
    c: [{ t: f.name, a: "right", b: true }, f.code,
        { t: ar1(f.rating) + "★", b: true, c: D_GOOD }, { t: ar1(f.avg) + "★", c: D_BAD },
        ar(f.n), arn(f.near), f.who.slice(0, 16), f.density],
    fill: i % 2 ? T_NEU : W }));
  table(s, M, 1.36, CW, [
    { t: "المحطة", w: 18, a: "right" }, { t: "الكود", w: 8 }, { t: "تقييمنا", w: 10 },
    { t: "متوسط المنافس", w: 13 }, { t: "منافسون داخل ٥ كم", w: 13 },
    { t: "أقرب منافس", w: 11 }, { t: "من هو", w: 15 }, { t: "الكثافة", w: 12 }],
    rows, { rh: 0.42, fs: 11 });
  const cw = (CW - 2 * 0.2) / 3;
  const cards = [["نتفوّق في كل موقع", "٤٫٨ و٤٫٩ نجمة مقابل ٣٫٦ إلى ٤٫١ · و٥٧٪ من المنافسين دون أربع نجوم", GOOD],
                 ["ونخسر بالمسافة", "أقرب منافس على ١٨١ متراً في المعيصم و٤١١ في الفردوس", BAD],
                 ["الدريس هو المنافس", "١٥ من ٤٠ موقعاً — ٣٨٪ محلياً مقابل ١٨٫٩٥٪ وطنياً", BLUE]];
  cards.forEach((c, i) => {
    const x = rtlx(i, cw, 0.2);
    s.addShape(p.ShapeType.roundRect, { x, y: 3.9, w: cw, h: 1.4, rectRadius: 0.05,
      fill: { color: c[2] === GOOD ? T_GOOD : c[2] === BAD ? T_BAD : T_NEU },
      line: { color: c[2], width: 1.2 } });
    s.addText(c[0], { x: x + 0.14, y: 4.02, w: cw - 0.28, h: 0.32, fontFace: F,
      fontSize: 14, bold: true,
      color: c[2] === GOOD ? D_GOOD : c[2] === BAD ? D_BAD : BLUE, ...rtl });
    s.addText(c[1], { x: x + 0.14, y: 4.36, w: cw - 0.28, h: 0.86, fontFace: F,
      fontSize: 11.5, color: INK, ...rtl });
  });
  band(s, 5.55, "الخدمة أفضل أصلاً — فالخطة اعتراضية: لافتة ومدخل وسرعة وإبراز التقييم قبل نقطة القرار بخمسمئة متر");
}

/* ═══ ١١ · المعيصم ═══ */
{
  const s = page("المعيصم MK002 — تغيير مسار لا منافسة", "الحدث ٧–٨ فبراير ٢٠٢٦ · ليلة واحدة");
  const bars = [42641, 35095, 25185, 9879, 9840, 12813];
  const lbl = ["المتوسط قبل", "٦ فبراير", "٧ فبراير", "٨ فبراير", "٩ فبراير", "المتوسط بعد"];
  const w0 = CW * 0.54, y0 = 1.8, h0 = 3.0, mx = 46000;
  bars.forEach((v, i) => {
    const bw = w0 / 6 * 0.6, bx = M + w0 - (i + 1) * (w0 / 6) + (w0 / 6 - bw) / 2;
    const bh = v / mx * h0;
    s.addShape(p.ShapeType.rect, { x: bx, y: y0 + h0 - bh, w: bw, h: bh,
      fill: { color: i <= 1 ? BGRAY : (i === 2 ? GOLD : BAD) } });
    s.addText(ar(v), { x: bx - 0.2, y: y0 + h0 - bh - 0.28, w: bw + 0.4, h: 0.26,
      fontFace: F, fontSize: 9.5, bold: true, color: INK2, align: "center", margin: 0 });
    s.addText(lbl[i], { x: bx - 0.24, y: y0 + h0 + 0.05, w: bw + 0.48, h: 0.26,
      fontFace: F, fontSize: 9, color: INK, align: "center", margin: 0 });
  });
  s.addShape(p.ShapeType.rect, { x: M, y: y0 + h0, w: w0, h: 0.02, fill: { color: LINE2 } });
  s.addText("لتراً يومياً", { x: M, y: 1.44, w: w0, h: 0.28, fontFace: F,
    fontSize: 11, color: INK2, ...rtl });
  const x2 = M + CW * 0.58, w2 = CW * 0.42;
  s.addText("توقيع الوصول لا المنافسة", { x: x2, y: 1.4, w: w2, h: 0.3, fontFace: F,
    fontSize: 13.5, bold: true, color: ORANGE, ...rtl });
  const f = [["المعاملات", "٢٬٠٧٣ ← ٦٠١", "−٧١٪", BAD],
             ["اللترات", "٤٢٬٦٤١ ← ١٢٬٨١٣", "−٧٠٪", BAD],
             ["حجم التعبئة", "٢٠٫٦ ← ٢١٫٣", "+٣٫٦٪", GOOD],
             ["تقييمنا", "٤٫٨ مقابل ٣٫٧", "الأعلى", GOOD],
             ["مطابقة المشتريات", "+٠٫٧١٪", "سليمة", GOOD]];
  f.forEach((r, i) => {
    const y = 1.8 + i * 0.6;
    s.addShape(p.ShapeType.rect, { x: SW - M - 0.045, y, w: 0.045, h: 0.52, fill: { color: r[3] } });
    s.addText(r[0], { x: x2 + 2.5, y, w: w2 - 2.65, h: 0.26, fontFace: F,
      fontSize: 10.5, color: INK3, ...rtl });
    s.addText(r[1], { x: x2 + 2.5, y: y + 0.24, w: w2 - 2.65, h: 0.3, fontFace: F,
      fontSize: 12, bold: true, color: BGRAY, ...rtl });
    s.addText(r[2], { x: x2, y: y + 0.1, w: 2.3, h: 0.34, fontFace: F,
      fontSize: 14, bold: true, color: r[3], align: "left", margin: 0 });
  });
  s.addShape(p.ShapeType.roundRect, { x: x2, y: 4.88, w: w2, h: 1.1, rectRadius: 0.05,
    fill: { color: T_OR }, line: { color: ORANGE, width: 1.2 } });
  s.addText("المعاملات تسقط والتعبئة تثبت — المنافس يسحب العميل الحدّي، والمدخل المغلق يمنع الجميع بالتساوي.",
    { x: x2 + 0.16, y: 5.0, w: w2 - 0.32, h: 0.86, fontFace: F, fontSize: 12, color: INK, ...rtl });
  band(s, 6.15, "الخسارة ١٬٢٤٥٬٥٠٣ ريال هامش سنوياً · والشرائع نفس التوقيع في ٤–٥ فبراير بـ٣٩٢٬٩٠٢ · ولا ثالث لهما في ٥٥ محطة", BAD);
}

/* ═══ ١٢ · القسم الثالث ═══ */
divider("القسم الثالث", "تقسيم المحطات", "نوع الموقع وفئة الإنتاجية — بمصطلحات القطاع");

/* ═══ ١٣ · الشرائح ═══ */
{
  const s = page("خمس شرائح", "التصنيف من سلوك العميل: حصة الديزل ثم الدفع المؤسسي ثم الإنتاجية");
  const rows = SEG.map((g, i) => ({
    c: [{ t: g.seg, b: true, a: "right", c: BGRAY }, { t: g.en, fs: 9.5, c: INK3 },
        g.rule, ar(g.n), ar1(g.mlpa), pc0(g.mlpa / T.mlpa), ar1(g.lpv),
        pc0(g.diesel), ar1(g.inv), { t: g.driver, a: "right", b: true, c: ORANGE }],
    fill: i % 2 ? T_NEU : W }));
  rows.push({ c: [{ t: "الإجمالي", b: true }, "", "", ar(T.n), ar(Math.round(T.mlpa)), "١٠٠٪",
                  ar1(N.volume / N.visits), "٢٠٪", ar1(N.inv), ""], fill: T_BAND, b: true });
  table(s, M, 1.36, CW, [
    { t: "الشريحة", w: 15, a: "right" }, { t: "Site Type", w: 13 },
    { t: "قاعدة التصنيف", w: 17, a: "right" }, { t: "محطات", w: 7 },
    { t: "MLPA", w: 8 }, { t: "٪ الحجم", w: 8 }, { t: "متوسط التعبئة", w: 9 },
    { t: "ديزل", w: 7 }, { t: "قيمة المعاملة", w: 9 }, { t: "المحرّك", w: 17, a: "right" }],
    rows, { rh: 0.46, fs: 11, hfs: 9.5 });
  head(s, 4.4, "فئة الإنتاجية — بمقياس MLPA");
  const cls = [["فئة أ", "٩ MLPA فأكثر", "٢٣ محطة", "٣٩٢٫٤ MLPA", ORANGE],
               ["فئة ب", "٤٫٥ إلى ٩", "١٨ محطة", "١٢٣٫٥ MLPA", GOLD],
               ["فئة ج", "دون ٤٫٥", "١١ محطة", "٣٥٫٢ MLPA", BGRAY]];
  const cw = (CW - 2 * 0.2) / 3;
  cls.forEach((c, i) => {
    const x = rtlx(i, cw, 0.2);
    s.addShape(p.ShapeType.roundRect, { x, y: 4.76, w: cw, h: 1.05, rectRadius: 0.05,
      fill: { color: BG }, line: { color: c[4], width: 1.2 } });
    s.addShape(p.ShapeType.rect, { x, y: 4.76, w: cw, h: 0.05, fill: { color: c[4] } });
    s.addText(c[0], { x: x + cw - 1.3, y: 4.88, w: 1.18, h: 0.32, fontFace: F,
      fontSize: 15, bold: true, color: c[4], ...rtl });
    s.addText(c[1], { x: x + 0.12, y: 4.9, w: cw - 1.5, h: 0.3, fontFace: F,
      fontSize: 11, color: INK2, align: "left", margin: 0 });
    s.addText(c[2] + "  ·  " + c[3], { x: x + 0.12, y: 5.26, w: cw - 0.24, h: 0.4,
      fontFace: F, fontSize: 12.5, bold: true, color: BGRAY, align: "center", margin: 0 });
  });
  foot(s, "الشبكة ٥٥١ MLPA على ٥٢ محطة — متوسط ١٠٫٦ للمحطة");
}

/* ═══ ١٤ · القمة والقاع ═══ */
{
  const s = page("القمة والقاع", "الإنتاجية اليومية · الترتيب يقود توزيع الجهد");
  const top = [...ST].sort((a, b) => b.lpd - a.lpd);
  const mk = r => [{ t: r.name.slice(0, 19), a: "right", b: true }, r.code, r.seg,
                   ar1(r.mlpa), ar(r.vpd), ar1(r.lpv), ar1(r.inv)];
  const cols = [{ t: "المحطة", w: 29, a: "right" }, { t: "الكود", w: 10 },
                { t: "الشريحة", w: 20 }, { t: "MLPA", w: 11 },
                { t: "معاملات/يوم", w: 11 }, { t: "التعبئة", w: 9 }, { t: "المعاملة", w: 10 }];
  s.addText("أعلى ٨", { x: M + CW * 0.515, y: 1.34, w: CW * 0.485, h: 0.3, fontFace: F,
    fontSize: 13, bold: true, color: GOOD, ...rtl });
  table(s, M + CW * 0.515, 1.68, CW * 0.485, cols, top.slice(0, 8).map(mk),
        { rh: 0.375, fs: 9.5, hfs: 9.5 });
  s.addText("أدنى ٨", { x: M, y: 1.34, w: CW * 0.485, h: 0.3, fontFace: F,
    fontSize: 13, bold: true, color: BAD, ...rtl });
  table(s, M, 1.68, CW * 0.485, cols, top.slice(-8).reverse().map(mk),
        { rh: 0.375, fs: 9.5, hfs: 9.5 });
  const box = (x, w, t, b, col) => {
    s.addShape(p.ShapeType.roundRect, { x, y: 5.3, w, h: 1.0, rectRadius: 0.04,
      fill: { color: col === GOOD ? T_GOOD : T_BAD }, line: { color: col, width: 1 } });
    s.addText(t, { x: x + 0.14, y: 5.4, w: w - 0.28, h: 0.3, fontFace: F,
      fontSize: 12.5, bold: true, color: col === GOOD ? D_GOOD : D_BAD, ...rtl });
    s.addText(b, { x: x + 0.14, y: 5.72, w: w - 0.28, h: 0.5, fontFace: F,
      fontSize: 11, color: INK, ...rtl });
  };
  box(M + CW * 0.515, CW * 0.485, "التركّز شديد",
      "أعلى ٥ محطات تصنع ٣٠٪ من الحجم · MK007 وحدها ٤٧٫٥ MLPA", GOOD);
  box(M, CW * 0.485, "القاع ليس فشلاً",
      "٧ محطات دون ٤٫٥ MLPA لا تصنع سوى ٣٫٩٪ · أولويتها التكلفة والتأجير", BAD);
}

/* ═══ ١٥ · القسم الرابع ═══ */
divider("القسم الرابع", "خطة المبيعات", "رافعتان: المعاملات ثم السلة");

/* ═══ ١٦ · المنهج ═══ */
{
  const s = page("المنهج — لماذا رافعتان", "لا نطلب من العميل أن يعبّي أكثر مما يحتاج");
  band(s, 1.32, "الحجم  =  المعاملات  ×  متوسط حجم التعبئة        ←  الأولى تشغيلية، والثانية محكومة بمزيج الوقود");
  const cw = (CW - 0.24) / 2;
  const lev = [
    ["رافعة (أ) · المعاملات", ar(Math.round(T.gap_peak)) + " معاملة/يوم",
     ar(Math.round(T.sar_txn / 1000)) + " ألف ريال", GOOD,
     ["نقارن شكل ساعات المحطة بشكل شريحتها",
      "نحتسب النقص في ساعات ذروتها وحدها",
      "النقص في ساعات الهدوء اختلاف طلب لا فاقد",
      "العلاج: طاقم الذروة وتوفّر المضخات"]],
    ["رافعة (ب) · حجم التعبئة", ar(Math.round(T.upl_fill)) + " لتر/يوم",
     ar(Math.round(T.sar_fill / 1000)) + " ألف ريال", ORANGE,
     ["ثلثا تباين التعبئة مزيج وقود لا سلوك — R² ٠٫٦٦",
      "فالمعيار من البواقي: كم تبيع مقابل ما يفسّره مزيجك",
      "٣٤ محطة تحت ما يفسّره مزيجها",
      "نغلق ٤٠٪ من الفجوة في ستة أشهر"]]];
  lev.forEach((v, i) => {
    const x = rtlx(i, cw, 0.24);
    s.addShape(p.ShapeType.roundRect, { x, y: 2.05, w: cw, h: 3.2, rectRadius: 0.05,
      fill: { color: W }, line: { color: v[3], width: 1.4 } });
    s.addShape(p.ShapeType.rect, { x, y: 2.05, w: cw, h: 0.56, fill: { color: v[3] } });
    s.addText(v[0], { x: x + 0.12, y: 2.11, w: cw - 0.24, h: 0.44, fontFace: F,
      fontSize: 15, bold: true, color: W, align: "center", margin: 0 });
    s.addText(v[1], { x: x + 0.12, y: 2.72, w: cw / 2 - 0.18, h: 0.4, fontFace: F,
      fontSize: 16, bold: true, color: BGRAY, align: "center", margin: 0 });
    s.addText(v[2], { x: x + cw / 2 + 0.06, y: 2.76, w: cw / 2 - 0.18, h: 0.36, fontFace: F,
      fontSize: 14, bold: true, color: v[3], align: "center", margin: 0 });
    s.addShape(p.ShapeType.rect, { x: x + 0.4, y: 3.2, w: cw - 0.8, h: 0.02, fill: { color: LINE2 } });
    s.addText(v[4].map((t, k) => ({ text: t, options: { bullet: { code: "25AA" },
      breakLine: k < v[4].length - 1 } })), { x: x + 0.16, y: 3.32, w: cw - 0.32, h: 1.8,
      fontFace: F, fontSize: 11.5, color: INK, lineSpacing: 18, paraSpaceAfter: 6, ...rtl });
  });
  const cw2 = (CW - 2 * 0.2) / 3;
  [[ar1(T.sar / 1e6), "مليون ريال — الفرصة المؤكَّدة", GOOD],
   ["١١٫٤٤", "هللة لكل لتر — هامش الشبكة المشغّلة", BGRAY],
   [ar(Math.round(T.gap_quiet)), "معاملة/يوم غير مؤكَّدة — تحتاج جدول عمالة بالساعة", INK3]]
    .forEach((c, i) => {
      const x = rtlx(i, cw2, 0.2);
      s.addShape(p.ShapeType.roundRect, { x, y: 5.45, w: cw2, h: 0.95, rectRadius: 0.04,
        fill: { color: BG }, line: { color: LINE2, width: 1 } });
      s.addText(c[0], { x: x + 0.1, y: 5.54, w: cw2 - 0.2, h: 0.4, fontFace: F,
        fontSize: 20, bold: true, color: c[2], align: "center", margin: 0 });
      s.addText(c[1], { x: x + 0.1, y: 5.94, w: cw2 - 0.2, h: 0.4, fontFace: F,
        fontSize: 10, color: INK2, align: "center", margin: 0 });
    });
}

/* ═══ ١٧ · المستهدف بالشريحة ═══ */
{
  const s = page("المستهدف بالشريحة", "بالرافعتين · وستة أشهر");
  const rows = SEG.map((g, i) => ({
    c: [{ t: g.seg, a: "right", b: true, c: BGRAY }, ar(g.n), ar1(g.lpv),
        { t: ar2(g.r2), c: g.r2 > 0.4 ? D_BAD : INK3 },
        { t: ar(g.below), c: g.below ? D_BAD : INK },
        { t: "+" + ar(g.upl_fill), b: true }, { t: ar(g.sar_fill), b: true, c: D_GOLD },
        { t: ar(g.gap_peak), b: true }, { t: ar(g.sar_txn), b: true, c: D_GOOD },
        { t: ar(g.sar_fill + g.sar_txn), b: true }],
    fill: i % 2 ? T_NEU : W }));
  rows.push({ c: [{ t: "الإجمالي", b: true }, ar(T.n), ar1(N.volume / N.visits), "—",
                  ar(T.below), "+" + ar(Math.round(T.upl_fill)), ar(Math.round(T.sar_fill)),
                  ar(Math.round(T.gap_peak)), ar(Math.round(T.sar_txn)), ar(Math.round(T.sar))],
              fill: T_BAND, b: true });
  table(s, M, 1.36, CW, [
    { t: "الشريحة", w: 16, a: "right" }, { t: "محطات", w: 7 }, { t: "التعبئة", w: 8 },
    { t: "R² المزيج", w: 8 }, { t: "دون المتوقَّع", w: 9 },
    { t: "رفع السلة لتر/يوم", w: 11 }, { t: "قيمتها (ر)", w: 11 },
    { t: "فاقد الذروة معاملة/يوم", w: 11 }, { t: "قيمته (ر)", w: 10 },
    { t: "الإجمالي (ر)", w: 11 }], rows, { rh: 0.46, fs: 11, hfs: 9.5 });
  head(s, 4.4, "الفعل التجاري لكل شريحة");
  const rows2 = SEG.map((g, i) => ({
    c: [{ t: g.seg, a: "right", b: true, c: BGRAY },
        { t: g.driver, a: "right", b: true, c: ORANGE },
        { t: g.action, a: "right" }, { t: ar2(g.net_txn) + " ر", b: true }],
    fill: i % 2 ? T_NEU : W }));
  table(s, M, 4.74, CW, [
    { t: "الشريحة", w: 16, a: "right" }, { t: "المحرّك", w: 18, a: "right" },
    { t: "الفعل", w: 50, a: "right" }, { t: "صافي المعاملة", w: 16 }],
    rows2, { rh: 0.34, fs: 10.5, hfs: 10 });
  foot(s, "R² يقيس كم من تباين التعبئة يفسّره مزيج الوقود — كلما ارتفع قلّ ما يملكه التشغيل");
}

/* ═══ ١٨ · أكبر الفرص ═══ */
{
  const s = page("أكبر ١٢ فرصة", "مرتّبة بالقيمة السنوية · وهي قائمة الأولوية");
  const rows = ST.slice(0, 12).map((r, i) => ({
    c: [{ t: r.name.slice(0, 20), a: "right", b: true }, r.code, r.seg,
        ar1(r.lpv), ar1(r.pred), { t: ar1(r.resid), c: r.resid < 0 ? D_BAD : D_GOOD },
        { t: "+" + ar(r.upl_fill), b: true }, { t: ar(r.gap_peak), b: true },
        { t: ar(r.sar_total), b: true, c: D_GOOD },
        { t: r.conds.slice(0, 2).join(" · ") || "—", a: "right", fs: 9 }],
    fill: i % 2 ? T_NEU : W }));
  table(s, M, 1.36, CW, [
    { t: "المحطة", w: 18, a: "right" }, { t: "الكود", w: 7 }, { t: "الشريحة", w: 13 },
    { t: "التعبئة", w: 8 }, { t: "المتوقَّع من مزيجها", w: 9 }, { t: "الفارق", w: 7 },
    { t: "رفع السلة", w: 9 }, { t: "فاقد الذروة", w: 9 },
    { t: "القيمة السنوية (ر)", w: 11 }, { t: "الحالة المرصودة", w: 19, a: "right" }],
    rows, { rh: 0.375, fs: 10, hfs: 9.5 });
  band(s, 6.32, "هذه الاثنتا عشرة تحمل " +
       pc0(ST.slice(0, 12).reduce((a, r) => a + r.sar_total, 0) / T.sar) +
       " من الفرصة · والعمرة الجديدة ليست فيها: تعبئتها فوق ما يفسّره مزيجها، وتغطية خزانها ٢٫٥ يوم");
}

/* ═══ ١٩ · مستهدف الوردية والعامل ═══ */
{
  const s = page("من مستهدف المحطة إلى مستهدف العامل",
                 "ثلاث طبقات — الوحدة «معاملة» في الأوليين والريال في الثالثة");
  head(s, 1.32, "① الاشتقاق — العمرة الجديدة MK007");
  const bw = 2.62, gp = 0.34, y0 = 1.64, bh = 0.8;
  s.addShape(p.ShapeType.roundRect, { x: SW - M - bw, y: y0 + 0.42, w: bw, h: bh,
    rectRadius: 0.05, fill: { color: BGRAY } });
  s.addText("٤٬٢٦١", { x: SW - M - bw, y: y0 + 0.5, w: bw, h: 0.42, fontFace: F,
    fontSize: 24, bold: true, color: W, align: "center", margin: 0 });
  s.addText("معاملة يومياً — كل المحطة", { x: SW - M - bw, y: y0 + 0.9, w: bw, h: 0.28,
    fontFace: F, fontSize: 10.5, color: "D6D2CC", align: "center", margin: 0 });
  const step = (x, y, ttl, big, sub, col) => {
    s.addShape(p.ShapeType.roundRect, { x, y, w: bw, h: bh, rectRadius: 0.05,
      fill: { color: col === ORANGE ? T_OR : T_NEU }, line: { color: col, width: 1.2 } });
    s.addText(ttl, { x: x + 0.08, y: y + 0.05, w: bw - 0.16, h: 0.24, fontFace: F,
      fontSize: 10, color: INK2, align: "center", margin: 0 });
    s.addText(big, { x: x + 0.08, y: y + 0.27, w: bw - 0.16, h: 0.32, fontFace: F,
      fontSize: 11.5, bold: true, color: BGRAY, align: "center", margin: 0 });
    s.addText(sub, { x: x + 0.08, y: y + 0.56, w: bw - 0.16, h: 0.24, fontFace: F,
      fontSize: 12, bold: true, color: col, align: "center", margin: 0 });
  };
  const arrow = (x, y) => s.addText("◀", { x, y, w: gp, h: bh, fontFace: F,
    fontSize: 13, color: INK3, align: "center", valign: "middle", margin: 0 });
  const x1 = SW - M - bw - gp - bw, x2 = x1 - gp - bw, x3 = x2 - gp - bw;
  arrow(SW - M - bw - gp, y0 + 0.42);
  step(x1, y0, "الصباح ٦ص – ٦م", "٢٬٠٨٧ معاملة ÷ ١١ عاملاً", "= ١٩٠ للعامل", BGRAY);
  step(x1, y0 + 0.86, "المساء ٦م – ٦ص", "٢٬١٧٤ معاملة ÷ ٩ عمال", "= ٢٤٢ للعامل", ORANGE);
  arrow(x1 - gp, y0 + 0.42);
  step(x2, y0 + 0.42, "الفارق", "المساء ٥١٪ من الطلب · ٤٥٪ من الطاقم", "المسائي +٢٧٪", BAD);
  arrow(x2 - gp, y0 + 0.42);
  step(x3, y0 + 0.42, "ما نطلبه", "فاقد ذروة وردية المساء", "٥١ معاملة", GOOD);
  head(s, 3.24, "② الأرقام — كلها بالمعاملة");
  const rows = D.shifts.map((r, i) => ({
    c: [{ t: r.name.slice(0, 16), a: "right", b: true }, r.code, ar(r.vpd),
        ar(r.day), ar(r.eve), ar(r.dw) + " / " + ar(r.ew),
        { t: ar(r.load_d), b: true }, { t: ar(r.load_e), b: true, c: ORANGE },
        { t: "+" + pc0(r.load_e / r.load_d - 1), c: D_BAD },
        { t: ar(r.tgt_d + r.tgt_e), b: true, c: D_GOOD }],
    fill: i % 2 ? T_NEU : W }));
  table(s, M, 3.56, CW, [
    { t: "المحطة", w: 15, a: "right" }, { t: "الكود", w: 8 },
    { t: "معاملات المحطة", w: 12 }, { t: "منها صباحاً", w: 10 }, { t: "منها مساءً", w: 10 },
    { t: "عمال ص / م", w: 9 }, { t: "عبء الصباحي", w: 10 }, { t: "عبء المسائي", w: 10 },
    { t: "الفارق", w: 8 }, { t: "المستهدف اليومي", w: 11 }],
    rows, { rh: 0.335, fs: 10.5, hfs: 9.5 });
  const LY = [["الطبقة ١ · عبء الوردية", "معاملة",
               "معاملات الوردية ÷ عمالها. حِمل لا مستهدف — العامل لا يصنع المعاملة.", BGRAY],
              ["الطبقة ٢ · مستهدف الوردية", "معاملة",
               "ما تفقده المحطة في ساعات ذروتها. هذا وحده ما يُطالَب به.", ORANGE],
              ["الطبقة ٣ · الحافز", "ريال",
               "لترات إضافية × هامش المحطة × نسبة · بوابة جودة وسقف ١٠٠ ريال.", GOOD]];
  const cw = (CW - 2 * 0.2) / 3;
  LY.forEach((k, i) => {
    const x = rtlx(i, cw, 0.2);
    s.addShape(p.ShapeType.roundRect, { x, y: 5.94, w: cw, h: 0.9, rectRadius: 0.05,
      fill: { color: BG }, line: { color: k[3], width: 1.2 } });
    s.addShape(p.ShapeType.rect, { x, y: 5.94, w: cw, h: 0.05, fill: { color: k[3] } });
    s.addShape(p.ShapeType.roundRect, { x: x + 0.1, y: 6.05, w: 0.78, h: 0.25,
      rectRadius: 0.03, fill: { color: k[3] } });
    s.addText(k[1], { x: x + 0.1, y: 6.05, w: 0.78, h: 0.25, fontFace: F,
      fontSize: 9.5, bold: true, color: W, align: "center", valign: "middle", margin: 0 });
    s.addText(k[0], { x: x + 1.0, y: 6.04, w: cw - 1.1, h: 0.28, fontFace: F,
      fontSize: 11.5, bold: true, color: BGRAY, ...rtl });
    s.addText(k[2], { x: x + 0.1, y: 6.34, w: cw - 0.2, h: 0.46, fontFace: F,
      fontSize: 10, color: INK, ...rtl });
  });
}

/* ═══ ٢٠ · الحافز ═══ */
{
  const s = page("نموذج الحافز", "يُدفع من هامش اللتر الإضافي — لا من ميزانية منفصلة");
  const flow = [["اللترات فوق الأساس", "الفارق بين إنتاج العامل وأساس ورديته", ORANGE],
                ["× هامش المحطة", "١٠٫١٥ إلى ١٣٫١١ هللة — لكل محطة رقمها", ORANGE],
                ["× نسبة المشاركة", "حصة العامل من الهامش الإضافي", ORANGE],
                ["بوابة الجودة", "غير مستوفاة ← صفر", BAD],
                ["بوابة المحطة", "صافيها سالب ← صفر", BAD],
                ["سقف ١٠٠ ريال", "شهرياً لكل عامل — معتمد", BAD]];
  flow.forEach((f, i) => {
    const y = 1.36 + i * 0.72;
    s.addShape(p.ShapeType.roundRect, { x: M, y, w: CW * 0.52, h: 0.62, rectRadius: 0.04,
      fill: { color: f[2] === BAD ? T_BAD : T_OR }, line: { color: f[2], width: 1 } });
    s.addText(f[0], { x: M + 0.16, y: y + 0.04, w: CW * 0.52 - 0.32, h: 0.28, fontFace: F,
      fontSize: 13, bold: true, color: f[2] === BAD ? D_BAD : BGRAY, ...rtl });
    s.addText(f[1], { x: M + 0.16, y: y + 0.31, w: CW * 0.52 - 0.32, h: 0.26, fontFace: F,
      fontSize: 10.5, color: INK2, ...rtl });
  });
  const x2 = M + CW * 0.55, w2 = CW * 0.45;
  s.addText("سقف المكافأة في أي حملة", { x: x2, y: 1.32, w: w2, h: 0.3, fontFace: F,
    fontSize: 13.5, bold: true, color: ORANGE, ...rtl });
  const rows = SEG.map((g, i) => ({
    c: [{ t: g.seg, a: "right", b: true }, ar2(g.net_txn), ar2(g.net_txn * 2), ar2(g.net_txn * 3)],
    fill: i % 2 ? T_NEU : W }));
  table(s, x2, 1.68, w2, [
    { t: "الشريحة", w: 34, a: "right" }, { t: "+معاملة", w: 22 },
    { t: "+٢", w: 22 }, { t: "+٣", w: 22 }], rows, { rh: 0.36, fs: 10.5, hfs: 10 });
  const ok = [["علبة مناديل", "٠٫٤٥ ر", true], ["قارورة ماء", "١٫٥٠ ر", true],
              ["قهوة", "٤٫٠٠ ر", false], ["٥ لترات بنزين", "١١٫٠٠ ر", false],
              ["غسيل مجاني", "٢٥٫٠٠ ر", false]];
  s.addText("ما يمرّ وما لا يمرّ", { x: x2, y: 3.82, w: w2, h: 0.3, fontFace: F,
    fontSize: 12.5, bold: true, color: BGRAY, ...rtl });
  ok.forEach((k, i) => {
    const y = 4.18 + i * 0.44;
    s.addShape(p.ShapeType.roundRect, { x: x2, y, w: w2, h: 0.38, rectRadius: 0.03,
      fill: { color: k[2] ? T_GOOD : T_BAD } });
    s.addText(k[2] ? "✓" : "✕", { x: x2 + 0.1, y, w: 0.35, h: 0.38, fontFace: F,
      fontSize: 13, bold: true, color: k[2] ? D_GOOD : D_BAD, valign: "middle",
      align: "center", margin: 0 });
    s.addText(k[0], { x: x2 + 0.5, y, w: w2 - 1.6, h: 0.38, fontFace: F, fontSize: 11.5,
      color: INK, valign: "middle", ...rtl });
    s.addText(k[1], { x: x2 + w2 - 1.05, y, w: 0.95, h: 0.38, fontFace: F, fontSize: 11.5,
      bold: true, color: k[2] ? D_GOOD : D_BAD, valign: "middle", align: "left", margin: 0 });
  });
  band(s, 6.4, "القاعدة: أي مكافأة فوق ١٫٦ ريال لكل معاملة إضافية تُموَّل من معلن أو مستأجر — لا من هامش الوقود");
}

/* ═══ ٢١ · القسم الخامس ═══ */
divider("القسم الخامس", "التأجير والحملات", "باقات تُباع · وحملات تُطلقها حالة مرصودة");

/* ═══ ٢٢ · الوحدات ═══ */
{
  const s = page("الوحدات التأجيرية", "سجل ١٨٦ محطة · الوحدة كشك أو محل أو درايف ثرو أو مغسلة أو سوبرماركت");
  const U = D.units;
  const cw = (CW - 3 * 0.18) / 4;
  const st = [[ar(U.total), "وحدة مسجّلة", ar(U.n) + " محطة", BGRAY],
              [ar(U.leased), "مؤجَّرة", pc0(U.leased / U.total) + " إشغال", ORANGE],
              [ar(U.vacant), "شاغرة", "الفرصة المباشرة", BAD],
              [pc0(U.shops_leased / U.shops), "إشغال المحلات", ar(U.shops) + " محلاً — المنتج العالق", BAD]];
  st.forEach((v, i) => stat(s, rtlx(i, cw, 0.18), 1.36, cw, 1.5, v[0], v[1], v[3], v[2]));
  s.addText("الشغور في ثلاث فئات — ولكلٍّ باقتها", { x: M, y: 3.06, w: CW * 0.6, h: 0.3,
    fontFace: F, fontSize: 13.5, bold: true, color: ORANGE, ...rtl });
  const rows = U.cats.map((c, i) => ({
    c: [{ t: c.cat, a: "right", b: true, c: BGRAY }, ar(c.n),
        { t: ar(c.vacant), b: true, c: D_BAD },
        { t: ["نسبة من المبيعات", "التدريج", "المسوّق الخارجي"][i], b: true, c: ORANGE },
        { t: ["الحركة معلومة وقابلة للقياس", "الحركة تبدأ من صفر وتنمو",
              "لا فريق لنا هناك"][i], a: "right" }],
    fill: i % 2 ? T_NEU : W }));
  table(s, M, 3.4, CW * 0.6, [
    { t: "الفئة", w: 18, a: "right" }, { t: "محطات", w: 10 }, { t: "وحدات شاغرة", w: 14 },
    { t: "الباقة", w: 22 }, { t: "لماذا", w: 36, a: "right" }],
    rows, { rh: 0.42, fs: 11, hfs: 10 });
  const x2 = M + CW * 0.63, w2 = CW * 0.37;
  s.addText("أكبر الفرص", { x: x2, y: 3.06, w: w2, h: 0.3, fontFace: F,
    fontSize: 13.5, bold: true, color: ORANGE, ...rtl });
  const rows2 = U.top.slice(0, 5).map((u, i) => ({
    c: [{ t: u.name.slice(0, 18), a: "right", b: true }, u.code,
        { t: ar(u.vacant), b: true, c: D_BAD }, u.cat],
    fill: i % 2 ? T_NEU : W }));
  table(s, x2, 3.4, w2, [{ t: "المحطة", w: 36, a: "right" }, { t: "الكود", w: 16 },
    { t: "شاغرة", w: 16 }, { t: "الفئة", w: 22 }], rows2, { rh: 0.42, fs: 10.5, hfs: 10 });
  band(s, 6.0, "الشبكة العاملة إشغالها ٧٩٪ — لا أزمة شغور فيها · الشغور الحقيقي في تحت التنفيذ والامتياز");
  foot(s, "محطة حلي MK036: ٩٢ وحدة وصفر مؤجّر — لا سوق مستأجرين محلياً بهذا الحجم · والسجل بلا قيمة إيجارية");
}

/* ═══ ٢٣ · باقات التأجير ═══ */
{
  const s = page("باقات التأجير", "منتج يُباع — لا حملة تُطلق");
  const rows = D.packages.map((r, i) => {
    const live = r[1] === "قائمة";
    return { c: [{ t: r[0], a: "right", b: true, c: BGRAY },
                 { t: r[1], b: true, c: live ? D_GOOD : D_GOLD },
                 { t: r[2], a: "right" }, { t: r[3], a: "right" },
                 { t: r[4], a: "right", fs: 10 }],
             fill: live ? T_GOOD : (i % 2 ? T_NEU : W) };
  });
  table(s, M, 1.36, CW, [
    { t: "الباقة", w: 15, a: "right" }, { t: "الحالة", w: 9 },
    { t: "لمن", w: 20, a: "right" }, { t: "ما هي", w: 32, a: "right" },
    { t: "ما تحتاجه لتُطبَّق", w: 24, a: "right" }], rows, { rh: 0.52, fs: 11, hfs: 10.5 });
  band(s, 5.95, "الأخضر ثلاث باقات قائمة اليوم · والباقي مقترح يحتاج قرار تسعير أو نموذج عقد — لا موارد جديدة");
  foot(s, "جدول الإيجارات الحالية والمستهدفة لكل نوع وحدة ومدينة شرط لتسعير أي باقة");
}

/* ═══ التطبيق — قناة الأفراد ═══ */
{
  const s = page("التطبيق — قناة الأفراد ومحطات الأحياء", "يبدأ التشغيل أول ديسمبر ٢٠٢٦");
  const cw = (CW - 3 * 0.18) / 4;
  const st = [["٣٠٬٨٥٦", "معاملة يومياً", "الحضرية والأحياء — سوقه", ORANGE],
              ["١١٫٣", "مليون معاملة سنوياً", "قاعدته المحتملة", BGRAY],
              ["٣٦٪", "بلا هوية عميل اليوم", "٢٢٦ مليون ريال", BAD],
              ["٣٫٠٪", "حصة التطبيقات اليوم", "وكلها أساطيل لا أفراد", BAD]];
  st.forEach((v, i) => stat(s, rtlx(i, cw, 0.18), 1.36, cw, 1.5, v[0], v[1], v[3], v[2]));
  const cw2 = (CW - 0.24) / 2;
  const bx = [rtlx(0, cw2, 0.24), rtlx(1, cw2, 0.24)];
  s.addShape(p.ShapeType.roundRect, { x: bx[0], y: 3.06, w: cw2, h: 1.62, rectRadius: 0.05,
    fill: { color: T_BAD }, line: { color: BAD, width: 1.2 } });
  s.addText("ما لا نبني عليه القرار", { x: bx[0] + 0.16, y: 3.16, w: cw2 - 0.32, h: 0.3,
    fontFace: F, fontSize: 13.5, bold: true, color: D_BAD, ...rtl });
  s.addText("رفع التردد وحده يعطي ٢٩ إلى ٢٨٧ ألف ريال سنوياً حسب التبني — رقم متواضع أمام كلفة تطبيق وتشغيله.",
    { x: bx[0] + 0.16, y: 3.5, w: cw2 - 0.32, h: 1.04, fontFace: F, fontSize: 12, color: INK, ...rtl });
  s.addShape(p.ShapeType.roundRect, { x: bx[1], y: 3.06, w: cw2, h: 1.62, rectRadius: 0.05,
    fill: { color: T_GOOD }, line: { color: GOOD, width: 1.2 } });
  s.addText("وما نبنيه عليه", { x: bx[1] + 0.16, y: 3.16, w: cw2 - 0.32, h: 0.3,
    fontFace: F, fontSize: 13.5, bold: true, color: D_GOOD, ...rtl });
  s.addText("الهوية. التطبيق يحوّل ٣٦٪ من الإيراد من مجهول إلى عميل معروف — وهو ما يجعل كل حملة قابلة للاستهداف والقياس.",
    { x: bx[1] + 0.16, y: 3.5, w: cw2 - 0.32, h: 1.04, fontFace: F, fontSize: 12, color: INK, ...rtl });
  head(s, 4.88, "ما يفتحه — ولا يُفتح بغيره");
  const use = [["استهداف المناديل", "بدل كل من تجاوز ٥٠ ريالاً، فقط من تعبئته دون معدّله"],
               ["مكافأة المستأجر", "توصيلها بكلفة صفر — لا طباعة ولا كوبون ورقي"],
               ["قياس التردد", "المؤشر الوحيد الذي لا نقيسه اليوم إطلاقاً"],
               ["محرّك الأحياء", "التردد هو محرّك الشريحة — والتطبيق أداته الوحيدة الرخيصة"]];
  const cw3 = (CW - 3 * 0.18) / 4;
  use.forEach((u, i) => {
    const x = rtlx(i, cw3, 0.18);
    s.addShape(p.ShapeType.roundRect, { x, y: 5.22, w: cw3, h: 1.1, rectRadius: 0.04,
      fill: { color: BG }, line: { color: ORANGE, width: 1 } });
    s.addShape(p.ShapeType.rect, { x, y: 5.22, w: cw3, h: 0.05, fill: { color: ORANGE } });
    s.addText(u[0], { x: x + 0.1, y: 5.34, w: cw3 - 0.2, h: 0.3, fontFace: F,
      fontSize: 12.5, bold: true, color: BGRAY, align: "center", margin: 0 });
    s.addText(u[1], { x: x + 0.1, y: 5.66, w: cw3 - 0.2, h: 0.56, fontFace: F,
      fontSize: 10, color: INK, align: "center", margin: 0 });
  });
  foot(s, "يلزم تثبيته قبل الإطلاق: هل يتضمن الدفع · وبأي عملة يكافئ · وهل يربط المستأجرين · وفي كم محطة يبدأ");
}

/* ═══ ٢٤ · علبة المناديل ═══ */
{
  const s = page("هدية علبة المناديل", "٤٥ هللة للعلبة · عند تعبئة بخمسين ريالاً فأكثر");
  const cw = (CW - 3 * 0.18) / 4;
  const st = [["٤٥ هللة", "كلفة العلبة", "مؤكَّدة", ORANGE],
              ["٥٠ ريالاً", "العتبة", "قرار الإدارة", BGRAY],
              ["٦٫٣ لتر", "تعادل الهدية", "١٣٫٩ ريالاً إضافية", GOOD],
              ["٥٧٫١٨", "قيمة المعاملة اليوم", "فوق العتبة أصلاً", BAD]];
  st.forEach((v, i) => stat(s, rtlx(i, cw, 0.18), 1.36, cw, 1.5, v[0], v[1], v[3], v[2]));
  const cw2 = (CW - 0.24) / 2;
  const bx = [rtlx(0, cw2, 0.24), rtlx(1, cw2, 0.24)];
  s.addShape(p.ShapeType.roundRect, { x: bx[0], y: 3.06, w: cw2, h: 1.5, rectRadius: 0.05,
    fill: { color: T_BAD }, line: { color: BAD, width: 1.2 } });
  s.addText("المشكلة", { x: bx[0] + 0.16, y: 3.17, w: cw2 - 0.32, h: 0.3, fontFace: F,
    fontSize: 14, bold: true, color: D_BAD, ...rtl });
  s.addText("متوسط فاتورتنا ٥٧٫١٨ ريالاً — فوق العتبة. العلبة تذهب لمن كان سيتجاوزها بلا حملة، ولا ترفع سلة أحد.",
    { x: bx[0] + 0.16, y: 3.52, w: cw2 - 0.32, h: 0.92, fontFace: F, fontSize: 12, color: INK, ...rtl });
  s.addShape(p.ShapeType.roundRect, { x: bx[1], y: 3.06, w: cw2, h: 1.5, rectRadius: 0.05,
    fill: { color: T_GOOD }, line: { color: GOOD, width: 1.2 } });
  s.addText("الحل — من منتجنا الخامس", { x: bx[1] + 0.16, y: 3.17, w: cw2 - 0.32, h: 0.3,
    fontFace: F, fontSize: 14, bold: true, color: D_GOOD, ...rtl });
  s.addText("العلبة مساحة إعلانية. بريال للوجه عند ١٠٪ من المعاملات: دخل ١٥٦٬٥١٨ مقابل كلفة ٧٠٬٤٣٣ شهرياً = صافي ١٫٠٣ مليون سنوياً.",
    { x: bx[1] + 0.16, y: 3.52, w: cw2 - 0.32, h: 0.92, fontFace: F, fontSize: 12, color: INK, ...rtl });
  head(s, 4.76, "شروط الإطلاق");
  const rules = [["مشروطة", "٣٤ محطة دون المتوقَّع فقط — لا الشبكة"],
                 ["مسقوفة", "عدد يومي محدد لكل محطة"],
                 ["ممولة", "وجه العلبة مبيع لمعلن قبل الطباعة"],
                 ["مقيسة", "التعبئة قبل وبعد — ٤ أسابيع أو تُوقَف"]];
  const cw3 = (CW - 3 * 0.18) / 4;
  rules.forEach((r, i) => {
    const x = rtlx(i, cw3, 0.18);
    s.addShape(p.ShapeType.roundRect, { x, y: 5.1, w: cw3, h: 1.05, rectRadius: 0.04,
      fill: { color: BG }, line: { color: ORANGE, width: 1 } });
    s.addShape(p.ShapeType.rect, { x, y: 5.1, w: cw3, h: 0.05, fill: { color: ORANGE } });
    s.addText(r[0], { x: x + 0.1, y: 5.22, w: cw3 - 0.2, h: 0.32, fontFace: F,
      fontSize: 14, bold: true, color: ORANGE, align: "center", margin: 0 });
    s.addText(r[1], { x: x + 0.1, y: 5.56, w: cw3 - 0.2, h: 0.5, fontFace: F,
      fontSize: 10.5, color: INK, align: "center", margin: 0 });
  });
  foot(s, "المطلوب لإقرارها: سعر المساحة الإعلانية على العلبة");
}

/* ═══ ٢٥ · مصفوفة الأحداث ═══ */
{
  const s = page("مصفوفة الأحداث", "الحالة تُرصد من البيانات — ولكل حالة فعل واحد");
  const rows = D.events.map((e, i) => {
    const hi = e[4] === "hi";
    return { c: [{ t: e[0], a: "right", b: true, c: BGRAY },
                 { t: e[1], a: "right", b: true, c: hi ? D_BAD : ORANGE },
                 { t: e[2], a: "right", fs: 10.5 },
                 { t: e[3], b: true, c: hi ? D_BAD : ORANGE }],
             fill: hi ? T_BAD : (i % 2 ? T_NEU : W) };
  });
  table(s, M, 1.36, CW, [
    { t: "الحالة المرصودة", w: 22, a: "right" }, { t: "الفعل", w: 20, a: "right" },
    { t: "كيف يُنفَّذ", w: 44, a: "right" }, { t: "التصنيف", w: 12 }],
    rows, { rh: 0.52, fs: 11, hfs: 10.5 });
  band(s, 5.95, "الأحمر أربع بوابات: محطة عليها إحداها لا تدخل أي حملة قبل إقفالها — وإلا أنفقنا على قياس لا نثق به", BAD);
}

/* ═══ ٢٦ · القسم السادس ═══ */
divider("القسم السادس", "التنفيذ", "ثلاث موجات وبوابة بين كل اثنتين");

/* ═══ ٢٧ · الخارطة ═══ */
{
  const s = page("خارطة ٩٠ يوماً", "لا تبدأ موجة قبل إقفال بوابة سابقتها");
  const waves = [["أول ٣٠ يوماً", "الإقفال", BAD,
    ["تحقيق MK008 وRY024", "تفعيل تسجيل الدفع في ١٨ محطة",
     "تصحيح مداخل المعيصم والشرائع على الخرائط", "توحيد أسماء المحطات على الكود",
     "جدول عمالة بالساعة لكل محطة", "جدول إيجارات لكل نوع وحدة",
     "مسح منافسي العقار والإكسسوارات والإعلان"]],
    ["اليوم ٣١ إلى ٦٠", "الإطلاق المحدود", ORANGE,
     ["المستهدف في أعلى ١٢ محطة", "طاقم الذروة في محطات فاقد الذروة",
      "علبة المناديل في ٤ محطات بعد بيع المساحة", "تعاقد أساطيل في ٣ محاور طريق",
      "باقة التدريج في ٣ محطات تحت التنفيذ", "المستأجر الرئيسي لحلي"]],
    ["اليوم ٦١ إلى ٩٠", "التعميم المشروط", GOOD,
     ["تعميم ما أثبت أثره ٤ أسابيع", "حافز العامل على الشبكة المؤهَّلة",
      "باقات التأجير على كل الفئات", "إعادة تحجيم عمالة المعيصم",
      "تثبيت مؤشرات التطبيق قبل إطلاقه أول ديسمبر", "مراجعة المستهدف على الأداء الفعلي"]]];
  const cw = (CW - 2 * 0.22) / 3;
  waves.forEach((w, i) => {
    const x = rtlx(i, cw, 0.22);
    s.addShape(p.ShapeType.roundRect, { x, y: 1.34, w: cw, h: 4.4, rectRadius: 0.05,
      fill: { color: W }, line: { color: w[2], width: 1.4 } });
    s.addShape(p.ShapeType.rect, { x, y: 1.34, w: cw, h: 0.8, fill: { color: w[2] } });
    s.addText(w[0], { x: x + 0.1, y: 1.39, w: cw - 0.2, h: 0.3, fontFace: F,
      fontSize: 11.5, color: W, align: "center", margin: 0 });
    s.addText(w[1], { x: x + 0.1, y: 1.71, w: cw - 0.2, h: 0.38, fontFace: F,
      fontSize: 17, bold: true, color: W, align: "center", margin: 0 });
    s.addText(w[3].map((t, k) => ({ text: t, options: { bullet: { code: "25AA" },
      breakLine: k < w[3].length - 1 } })), { x: x + 0.14, y: 2.28, w: cw - 0.28, h: 3.3,
      fontFace: F, fontSize: 11, color: INK, lineSpacing: 17, paraSpaceAfter: 5, ...rtl });
  });
  band(s, 5.92, "بوابة الثانية: لا حملة في محطة لم تُقفل حالتها الرقابية · بوابة الثالثة: لا تعميم لآلية لم ترفع رقمها في ٤ أسابيع");
}

/* ═══ ٢٨ · لوحة المتابعة ═══ */
{
  const s = page("لوحة المتابعة الشهرية", "ثمانية مؤشرات في صفحة واحدة");
  const kpis = [["فاقد ساعات الذروة", "معاملة/يوم — الهدف صفر", "أساسي"],
                ["متوسط حجم التعبئة", "مقابل ما يفسّره المزيج", "أساسي"],
                ["تسجيل وسيلة الدفع", "من ٦٤٪ إلى ٩٥٪", "بوابة"],
                ["فجوة المشتريات والمبيعات", "±١٪ لكل محطة", "بوابة"],
                ["الحافز ÷ الهامش الإضافي", "١٥٪ فأقل", "ضابط"],
                ["إشغال الوحدات بالفئة", "مشغّلة · تحت تنفيذ · امتياز", "نمو"],
                ["حصة الأساطيل الرقمية", "من ١٫٩٪ صعوداً", "نمو"],
                ["تغطية الخزان", "أيام — تنبيه دون ٤", "مخاطر"]];
  const cw = (CW - 3 * 0.18) / 4;
  kpis.forEach((k, i) => {
    const x = rtlx(i % 4, cw, 0.18), y = 1.34 + Math.floor(i / 4) * 1.42;
    const col = k[2] === "بوابة" ? BAD : k[2] === "ضابط" ? GOLD : k[2] === "مخاطر" ? BLUE : ORANGE;
    s.addShape(p.ShapeType.roundRect, { x, y, w: cw, h: 1.24, rectRadius: 0.05,
      fill: { color: BG }, line: { color: LINE2, width: 1 } });
    s.addShape(p.ShapeType.rect, { x, y, w: cw, h: 0.05, fill: { color: col } });
    s.addText(k[2], { x: x + 0.1, y: y + 0.12, w: cw - 0.2, h: 0.26, fontFace: F,
      fontSize: 9.5, bold: true, color: col, align: "center", margin: 0 });
    s.addText(k[0], { x: x + 0.1, y: y + 0.42, w: cw - 0.2, h: 0.38, fontFace: F,
      fontSize: 12, bold: true, color: BGRAY, align: "center", margin: 0 });
    s.addText(k[1], { x: x + 0.1, y: y + 0.82, w: cw - 0.2, h: 0.34, fontFace: F,
      fontSize: 10, color: INK2, align: "center", margin: 0 });
  });
  head(s, 4.36, "ما نحتاجه من الإدارات");
  const need = [["العمليات", "جدول عمالة بالساعة لكل محطة — لتعميم مستهدف الوردية"],
                ["العقار", "جدول الإيجار لكل نوع وحدة — لتسعير الباقات"],
                ["المالية", "هل «مصاريف التشغيل» تغطي الامتياز؟ الجواب يقلب صافي اللتر"],
                ["التسويق", "سعر المساحة الإعلانية على علبة المناديل"],
                ["التطوير", "مسح منافسي العقار والإكسسوارات والإعلان"]];
  need.forEach((n, i) => {
    const y = 4.74 + i * 0.44;
    s.addShape(p.ShapeType.rect, { x: SW - M - 0.04, y: y + 0.05, w: 0.04, h: 0.32,
      fill: { color: ORANGE } });
    s.addText(n[0], { x: SW - M - 1.85, y, w: 1.7, h: 0.38, fontFace: F, fontSize: 12,
      bold: true, color: BGRAY, ...rtl });
    s.addText(n[1], { x: M, y, w: CW - 2.0, h: 0.38, fontFace: F, fontSize: 11.5,
      color: INK, ...rtl });
  });
}

/* ═══ ٢٩ · الختام ═══ */
cover("الإدارة التجارية · درب", "القرار المطلوب",
  "اعتماد المستهدف برافعتين · إقفال ثغرة البيانات قبل الإنفاق · مسح المنافسين في المنتجات الثلاثة المجهولة",
  "الفرصة المؤكَّدة " + ar(Math.round(T.sar)) + " ريال سنوياً — " +
  ar(Math.round(T.gap_peak)) + " معاملة و" + ar(Math.round(T.upl_fill)) + " لتراً يومياً");

p.writeFile({ fileName: "docs/عرض-تحليل-المنافذ-وخطة-المبيعات.pptx" })
 .then(f => console.log("✓ " + f + "  ·  " + p.slides.length + " شريحة"));
