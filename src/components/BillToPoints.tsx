import { useState } from "react";
import { useScenario } from "../context/ScenarioContext";
import { fmtSar, fmtInt, fmtNum } from "../model/engine";
import { Card, Stat, NumberInput, Badge } from "./ui";

export default function BillToPoints() {
  const { inputs } = useScenario();
  const pv = inputs.pointValue; // 200 نقطة = ريال
  const [bill, setBill] = useState(100); // فاتورة عميل واحد
  const [totalPct, setTotalPct] = useState(3);
  const [customerPct, setCustomerPct] = useState(1);
  const [monthly, setMonthly] = useState(100000); // إجمالي فواتير عملاء البرنامج/شهر
  const [terms, setTerms] = useState(15); // مهلة سداد (يوم)

  const platformPct = Math.max(0, totalPct - customerPct);

  // فاتورة واحدة
  const merchantPays = (bill * totalPct) / 100;
  const customerValue = (bill * customerPct) / 100;
  const customerPoints = customerValue * pv;
  const darbShare = (bill * platformPct) / 100;

  // شهرياً لتاجر واحد
  const billedM = (monthly * totalPct) / 100;
  const custValueM = (monthly * customerPct) / 100;
  const custPointsM = custValueM * pv;
  const darbRevM = (monthly * platformPct) / 100;

  return (
    <div className="space-y-4">
      <div>
        <h2 className="text-lg font-extrabold">🧾 من الفاتورة إلى النقاط والتحصيل</h2>
        <p className="text-xs text-darb-mut">
          كيف تنعكس نسبة التاجر على نقاط العميل ومحفظته وكيف تُحصَّل من التاجر
        </p>
      </div>

      <Card>
        <div className="grid sm:grid-cols-2 lg:grid-cols-4 gap-4">
          <NumberInput label="فاتورة عميل واحد (﷼)" value={bill} onChange={setBill} step={10} />
          <NumberInput label="النسبة الكلية (%)" value={totalPct} onChange={setTotalPct} step={0.5} suffix="%" />
          <NumberInput label="حصة العميل · نقاط (%)" value={customerPct} onChange={setCustomerPct} step={0.5} suffix="%" hint={`حصة درب = ${fmtNum(platformPct, 1)}%`} />
          <NumberInput label="قيمة النقطة (نقطة = ريال)" value={pv} onChange={() => {}} step={0} hint="من السيناريوهات" />
        </div>
      </Card>

      {/* فاتورة واحدة */}
      <Card title="① فاتورة واحدة — كيف تنقسم">
        <div className="grid sm:grid-cols-2 lg:grid-cols-4 gap-3">
          <Stat label="💸 يدفع التاجر" value={fmtSar(merchantPays)} tone="warn" hint={`${fmtNum(totalPct, 1)}% من الفاتورة`} />
          <Stat label="💎 نقاط للعميل" value={fmtInt(customerPoints)} tone="accent" hint={`${fmtNum(customerPct, 1)}% = ${fmtSar(customerValue)}`} />
          <Stat label="👛 تُودع في محفظته" value={fmtSar(customerValue)} tone="good" hint="مغطّاة بريال حقيقي" />
          <Stat label="💰 إيراد درب" value={fmtSar(darbShare)} hint={`${fmtNum(platformPct, 1)}%`} />
        </div>
        <p className="text-xs text-darb-mut mt-3">
          العميل يكسب <b>{fmtInt(customerPoints)} نقطة</b> (= {fmtSar(customerValue)}) لحظة الشراء، مغطّاة بالكامل — ودرب
          تحصّل قيمتها لاحقاً من التاجر.
        </p>
      </Card>

      {/* شهرياً */}
      <Card title="② شهرياً لتاجر واحد — التحصيل">
        <div className="mb-3 max-w-xs">
          <NumberInput label="إجمالي فواتير عملاء البرنامج/شهر (﷼)" value={monthly} onChange={setMonthly} step={10000} />
        </div>
        <div className="grid sm:grid-cols-2 lg:grid-cols-4 gap-3">
          <Stat label="🧾 فاتورة درب للتاجر" value={fmtSar(billedM)} tone="warn" />
          <Stat label="💎 نقاط صُرفت للعملاء" value={fmtInt(custPointsM)} hint={fmtSar(custValueM)} tone="accent" />
          <Stat label="👛 تمويل المحافظ (إسكرو)" value={fmtSar(custValueM)} tone="good" />
          <Stat label="💰 إيراد درب/شهر" value={fmtSar(darbRevM)} hint={`${fmtSar(darbRevM * 12)}/سنة`} tone="good" />
        </div>
      </Card>

      {/* آلية التحصيل */}
      <Card title="⚙️ آلية التحصيل من التاجر">
        <ol className="text-sm space-y-1.5 text-darb-ink/90 list-decimal pr-5">
          <li>لحظة الشراء: العميل يمسح QR → يكسب نقاطه فوراً في محفظته (درب تُقدّم القيمة مؤقتاً).</li>
          <li>نهاية الشهر: النظام يجمع <b>كل فواتير عملاء البرنامج</b> عند التاجر ويحسب {fmtNum(totalPct, 1)}%.</li>
          <li>فاتورة إلكترونية تلقائية للتاجر بقيمة <b>{fmtSar(billedM)}</b> (منها {fmtSar(custValueM)} تموّل المحافظ + {fmtSar(darbRevM)} إيراد درب).</li>
          <li>مهلة سداد <b>{terms} يوم</b> · خصم تلقائي أو تحويل · غرامة عند التأخير.</li>
          <li>التسوية تغذّي <b>إسكرو المحافظ</b> فتبقى النقاط مغطّاة دائماً.</li>
        </ol>
        <div className="mt-3 max-w-xs">
          <NumberInput label="مهلة السداد (يوم)" value={terms} onChange={setTerms} step={5} />
        </div>
      </Card>

      <Card>
        <p className="text-sm text-darb-ink/90 leading-relaxed">
          🔑 <b>الفصل الزمني:</b> النقاط تُمنح <b>لحظياً</b> للعميل، والتحصيل من التاجر <b>شهري</b>. درب تُقدّم
          حصة العميل مؤقتاً ثم تستردّها + عمولتها ({fmtNum(platformPct, 1)}%) من التاجر — فتبقى المحفظة مغطّاة والنظام
          رابح. <Badge tone="good">لا خسارة</Badge>
        </p>
      </Card>
    </div>
  );
}
