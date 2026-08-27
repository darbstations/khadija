import { useScenario } from "../context/ScenarioContext";
import { fmtInt, fmtNum, fmtPct } from "../model/engine";
import { Card, Badge } from "./ui";

export default function ModelSummary() {
  const { inputs } = useScenario();
  const pv = inputs.pointValue;
  const fuelCash = inputs.earnWhite / pv;

  const earnRows: [string, boolean, string][] = [
    ["🔋 شحن المحفظة", true, `يكسب ${fmtPct(fuelCash)} (مرة واحدة)`],
    ["⛽ بنزين من المحفظة", false, "كُسب عند الشحن — لا تكرار"],
    ["⛽ بنزين خارج المحفظة + QR", true, "يكسب عند الضخ"],
    ["🛍️ شراء عادي من متجر", true, "1% للعميل (من 3%) — يموّله التاجر"],
    ["🪪 شراء/بنزين/قسائم الباقة", false, "مدفوعة مسبقاً — 0 نقطة"],
    ["💎 استبدال النقاط", false, "خصم من الرصيد"],
  ];

  return (
    <div className="space-y-4">
      <div>
        <h2 className="text-lg font-extrabold">📋 ملخص النموذج النهائي · تانكي</h2>
        <p className="text-xs text-darb-mut">مرجع واحد لكل قرارات النموذج (أرقام حيّة من السيناريوهات)</p>
      </div>

      <div className="grid md:grid-cols-2 gap-3">
        <Card title="💠 قيمة النقطة">
          <p className="text-sm text-darb-ink/90 leading-relaxed">
            النقطة = <b>{fmtNum((1 / pv) * 100, 2)} هللة</b> · <b>{fmtInt(pv)} نقطة = ريال</b>.
            القاعدة: عدد النقاط ÷ {fmtInt(pv)} = القيمة بالريال. الاستبدال عند التجار يعطي قيمة أعلى (×1.5–1.9).
          </p>
        </Card>
        <Card title="👛 المحفظة (مغلقة)">
          <p className="text-sm text-darb-ink/90 leading-relaxed">
            تشتري <b>منتجات درب فقط (بنزين + باقات)</b> · <b>بلا سحب نقدي</b>. فائدتها: سيولة مقدّمة + رسوم بطاقات
            أقل + تنظيم أخف (كبطاقة هدايا) + منع تلاعب.
          </p>
        </Card>
      </div>

      <Card title="📐 قواعد الكسب — كل ريال يكسب مرة واحدة">
        <div className="overflow-x-auto">
          <table className="w-full border-collapse">
            <thead>
              <tr><th className="th">الحدث</th><th className="th">نقاط؟</th><th className="th">السبب</th></tr>
            </thead>
            <tbody>
              {earnRows.map((r) => (
                <tr key={r[0]}>
                  <td className="td font-bold">{r[0]}</td>
                  <td className="td"><Badge tone={r[1] ? "good" : "bad"}>{r[1] ? "✅" : "❌"}</Badge></td>
                  <td className="td text-xs">{r[2]}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Card>

      <div className="grid md:grid-cols-2 gap-3">
        <Card title="🔁 قنوات الاستبدال">
          <ul className="text-sm space-y-1.5 text-darb-ink/90 list-disc pr-5">
            <li>🛍️ <b>عروض التجار</b> — الأرخص على درب (التاجر يموّل) والأعلى قيمة للعميل.</li>
            <li>⛽ <b>خصم بنزين</b> — قيمة اسمية.</li>
            <li>🪪 <b>باقات درب</b> — شراء بالفلوس أو استبدال بالنقاط.</li>
          </ul>
        </Card>
        <Card title="🎯 المستويات">
          <ul className="text-sm space-y-1.5 text-darb-ink/90 list-disc pr-5">
            <li>الكاش باك {fmtPct(fuelCash)} <b>للجميع</b> (عدالة).</li>
            <li>التمييز <b>بالمنافع</b>: فضي (×1.5 مناسبات، دعم أولوية)، برتقالي (×2، VIP، خصم باقات).</li>
            <li>مسار منفصل <b>«الأعمال/الأساطيل»</b> بالدعوة (مدير حساب).</li>
          </ul>
        </Card>
      </div>

      <Card title="🤝 نموذج التاجر">
        <p className="text-sm text-darb-ink/90 leading-relaxed">
          <b>3% على كل فاتورة عميل برنامج</b> = 1% للعميل (نقاط) + 2% لدرب. نقطة التعادل للتاجر = النسبة ÷ هامشه
          (~5%)، والولاء يعطي 15–30%. <b>الاستدامة:</b> التحوّل لاحقاً لـ«التسعير على النمو» (يدفع فقط على نموّه فمستحيل يخسر)،
          والمعدل يُضبط حسب هامشه.
        </p>
      </Card>

      <div className="rounded-2xl bg-darb-orange/10 border border-darb-orange/40 p-5">
        <h3 className="font-extrabold text-darb-orange mb-2">🔑 المبادئ الحاكمة</h3>
        <ul className="text-sm space-y-1.5 text-darb-ink/90">
          <li><b>رصيد موحّد للعميل</b> (بنزين + تجار) · التقسيم داخلي لدرب فقط.</li>
          <li><b>لا ازدواج:</b> كل ريال يكسب مرة واحدة (شحن أو شراء).</li>
          <li><b>لا خسارة:</b> النقاط مغطّاة بالإسكرو (كل ريال يُستبدل كان مودَعاً).</li>
          <li><b>القيمة من الشركاء:</b> الوقود يجذب، والشركاء يموّلون الكرم.</li>
        </ul>
      </div>
    </div>
  );
}
