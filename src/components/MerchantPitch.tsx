import { useState } from "react";
import { fmtSar, fmtPct, fmtNum } from "../model/engine";
import { Card, Stat, NumberInput, Badge } from "./ui";

export default function MerchantPitch() {
  const [sales, setSales] = useState(100000); // مبيعات البرنامج/شهر
  const [marginPct, setMarginPct] = useState(60); // هامش التاجر %
  const [totalPct, setTotalPct] = useState(3); // النسبة الكلية من الفاتورة %
  const [customerPct, setCustomerPct] = useState(1); // حصة العميل (نقاط) %
  const [upliftPct, setUpliftPct] = useState(15); // الزيادة المتوقعة في المبيعات %

  const margin = marginPct / 100;
  const platformPct = Math.max(0, totalPct - customerPct); // حصة درب (المنصة)
  const cost = (sales * totalPct) / 100;
  const customerShare = (sales * customerPct) / 100;
  const darbRevenue = (sales * platformPct) / 100;
  const breakEvenUplift = margin ? totalPct / margin : 0; // % الزيادة المطلوبة للتعادل = النسبة ÷ الهامش (3% ÷ 0.60 = 5%)
  const upliftRevenue = (sales * upliftPct) / 100;
  const extraMargin = upliftRevenue * margin;
  const netGain = extraMargin - cost;
  const merchantROI = cost ? netGain / cost : 0;
  const convincing = upliftPct >= breakEvenUplift;

  // جدول مرجعي: النسبة المقترحة لإبقاء التعادل ~5%
  const refMargins = [70, 60, 45, 30, 20];

  return (
    <div className="space-y-4">
      <div>
        <h2 className="text-lg font-extrabold">🤝 حاسبة إقناع التاجر</h2>
        <p className="text-xs text-darb-mut">
          أثبتي للتاجر رياضياً أنه رابح — النسبة تُقدَّم كـ% من الهامش لا من الإيراد
        </p>
      </div>

      <Card>
        <div className="grid sm:grid-cols-2 lg:grid-cols-3 gap-4">
          <NumberInput label="إجمالي فواتير عملاء البرنامج/شهر (﷼)" value={sales} onChange={setSales} step={10000} hint="النسبة تُحسب على كل فاتورة" />
          <NumberInput label="هامش التاجر (%)" value={marginPct} onChange={setMarginPct} step={5} suffix="%" />
          <NumberInput label="نمو إجمالي المبيعات المتوقع (%)" value={upliftPct} onChange={setUpliftPct} step={1} suffix="%" hint="الولاء يعطي 15–30%" />
          <NumberInput label="النسبة الكلية من الفاتورة (%)" value={totalPct} onChange={setTotalPct} step={0.5} suffix="%" />
          <NumberInput label="حصة العميل · نقاط (%)" value={customerPct} onChange={setCustomerPct} step={0.5} suffix="%" hint={`حصة درب (المنصة) = ${fmtNum(platformPct, 1)}%`} />
        </div>
      </Card>

      {/* النتيجة الحاسمة */}
      <div className={`rounded-2xl border p-5 ${convincing ? "bg-darb-good/10 border-darb-good/40" : "bg-darb-bad/10 border-darb-bad/40"}`}>
        <div className="flex items-center justify-between flex-wrap gap-2">
          <div>
            <div className="text-xs text-darb-mut">نقطة التعادل = النسبة ÷ الهامش</div>
            <div className="text-2xl font-extrabold">
              يكفي زيادة <span className="text-darb-orange">+{fmtNum(breakEvenUplift, 1)}%</span> مبيعات
            </div>
          </div>
          <Badge tone={convincing ? "good" : "bad"}>
            {convincing ? `✅ مقنع · المتوقع +${upliftPct}% أعلى من التعادل` : "⚠️ غير مقنع · خفّضي النسبة"}
          </Badge>
        </div>
      </div>

      <div className="grid sm:grid-cols-2 lg:grid-cols-4 gap-3">
        <Stat label="💸 تكلفة التاجر/شهر" value={fmtSar(cost)} hint={`${fmtSar(customerShare)} عميل + ${fmtSar(darbRevenue)} درب`} tone="warn" />
        <Stat label="📈 ربح إضافي (من الزيادة)" value={fmtSar(extraMargin)} tone="accent" />
        <Stat label="✅ صافي ربح التاجر" value={fmtSar(netGain)} tone={netGain > 0 ? "good" : "bad"} />
        <Stat label="🎯 ROI للتاجر" value={fmtPct(merchantROI, 0)} tone={merchantROI > 0 ? "good" : "bad"} />
      </div>

      <Card>
        <div className="flex items-center justify-between stat">
          <span className="text-sm">💰 إيراد درب (المنصة) من هذا التاجر</span>
          <span className="font-extrabold text-darb-orange">{fmtSar(darbRevenue)}/شهر · {fmtSar(darbRevenue * 12)}/سنة</span>
        </div>
        <p className="text-xs text-darb-mut mt-2">
          حصة درب = {fmtNum(platformPct, 1)}% من مبيعات البرنامج (النسبة الكلية {fmtNum(totalPct, 1)}% − حصة العميل {fmtNum(customerPct, 1)}%).
        </p>
      </Card>

      <Card title="📊 النسبة المقترحة حسب الهامش (لإبقاء التعادل ~5%)">
        <div className="overflow-x-auto">
          <table className="w-full border-collapse">
            <thead>
              <tr>
                <th className="th">هامش التاجر</th>
                <th className="th">نسبة مقترحة</th>
                <th className="th">= % من الهامش</th>
                <th className="th">التعادل المطلوب</th>
              </tr>
            </thead>
            <tbody>
              {refMargins.map((mg) => {
                const suggested = +(0.05 * mg).toFixed(1); // تبقي التعادل ~5%
                return (
                  <tr key={mg} className={mg === marginPct ? "bg-darb-orange/10" : ""}>
                    <td className="td font-bold">{mg}%</td>
                    <td className="td text-darb-orange font-bold">{fmtNum(suggested, 1)}%</td>
                    <td className="td">{fmtNum((suggested / mg) * 100, 0)}%</td>
                    <td className="td">+{fmtNum(suggested / mg * 100, 1)}%</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
        <p className="text-xs text-darb-mut mt-2">القاعدة: كل ما قلّ هامش التاجر، خفّضي نسبته ليبقى التعادل سهلاً (~5%).</p>
      </Card>

      <Card title="ℹ️ الأساس · النسبة على كل فاتورة من عميل البرنامج">
        <p className="text-sm text-darb-ink/90 leading-relaxed">
          الـ{fmtNum(totalPct, 1)}% تُحسب على <b>كل فاتورة من عميل درب</b> (عميل المنظومة). هؤلاء عملاء تقودهم
          <b> نقاط درب</b> لاختيار متجرك — فالنسبة على <b>مبيعات تقودها المنظومة</b>، مقابل وصولك لقاعدة عملاء درب كاملة
          + بيانات + تسويق. والـ<b>Pilot 3 شهور</b> يثبت الأثر قبل أي التزام.
        </p>
      </Card>

      <Card title="🗣️ سكربت الإقناع (بالترتيب)">
        <ol className="text-sm space-y-1.5 text-darb-ink/90 list-decimal pr-5">
          <li><b>ابدئي بالقيمة:</b> «عندنا 50 ألف عميل درب — هؤلاء عملاؤك المحتملون.»</li>
          <li><b>اقلبي النسبة لهامش:</b> «الـ{fmtNum(totalPct, 1)}% = {fmtNum((totalPct / marginPct) * 100, 0)}% من ربحك فقط.»</li>
          <li><b>نقطة التعادل:</b> «يكفي نمو مبيعاتك +{fmtNum(breakEvenUplift, 1)}% لتغطية التكلفة، والولاء يعطي 15–30%.»</li>
          <li><b>de-risk:</b> «جرّب Pilot 3 شهور، لو ما تحقق النمو انسحب بدون رسوم.»</li>
          <li><b>أغلقي:</b> «تختبر بلا مخاطرة، والنمو الصافي يغطّي التكلفة بأضعاف.»</li>
        </ol>
      </Card>
    </div>
  );
}
