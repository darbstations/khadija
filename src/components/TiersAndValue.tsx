import { useScenario } from "../context/ScenarioContext";
import { tiers, pointMath, fmtInt, fmtNum, fmtPct } from "../model/engine";
import { Card, Stat } from "./ui";

const BENEFITS = [
  ["نقاط مضاعفة في الأعياد", "❌", "🎯 1.5×", "🎯 2×"],
  ["دعوة لفعاليات درب", "❌", "✅", "✅✅ VIP"],
  ["خصومات الشركاء", "5%", "10%", "15%"],
  ["أولوية الدعم الفني", "عادي", "سريع", "فوري"],
  ["إهداء النقاط للعائلة", "❌", "✅ محدود", "✅ بلا حدود"],
  ["تجميد الحساب", "❌", "❌", "✅ مرة/سنة"],
];

export default function TiersAndValue() {
  const { inputs } = useScenario();
  const t = tiers(inputs);
  const m = pointMath(inputs);
  const welcome = [inputs.welcomeGift, inputs.welcomeGift * 2, inputs.welcomeGift * 4];

  return (
    <div className="space-y-4">
      <div>
        <h2 className="text-lg font-extrabold">🎯 المستويات الثلاثة وقيمة النقطة</h2>
        <p className="text-xs text-darb-mut">الهيكل + الكسب + المنافع + المعادلات الأساسية</p>
      </div>

      <div className="grid sm:grid-cols-2 lg:grid-cols-4 gap-3">
        <Stat label="قيمة النقطة (ريال)" value={fmtNum(m.valuePerPointRiyal, 4)} />
        <Stat label="قيمة النقطة (هللة)" value={fmtNum(m.valuePerPointHalala, 3)} />
        <Stat label="نقاط ساسكو/ريال" value={fmtNum(m.sascoPointsPerRiyal, 2)} />
        <Stat label="الأبيض مقابل ساسكو" value={`${fmtInt(t[0].earnRate / m.sascoPointsPerRiyal)}x`} tone="accent" />
      </div>

      <Card title="🎯 هيكل المستويات">
        <div className="overflow-x-auto">
          <table className="w-full border-collapse">
            <thead>
              <tr>
                <th className="th">المستوى</th>
                <th className="th">معدل الكسب</th>
                <th className="th">الحد الأدنى (﷼/سنة)</th>
                <th className="th">الكاش باك</th>
                <th className="th">هدية الترحيب</th>
              </tr>
            </thead>
            <tbody>
              {t.map((tier, idx) => (
                <tr key={tier.key}>
                  <td className="td font-bold">
                    {tier.emoji} {tier.label}
                  </td>
                  <td className="td">{tier.earnRate} نقطة/﷼</td>
                  <td className="td">{fmtInt(tier.minSpend)}</td>
                  <td className="td text-darb-accent">{fmtPct(tier.cashbackPct)}</td>
                  <td className="td">{fmtInt(welcome[idx])} نقطة</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Card>

      <Card title="🎁 المنافع الإضافية">
        <div className="overflow-x-auto">
          <table className="w-full border-collapse">
            <thead>
              <tr>
                <th className="th">الميزة</th>
                <th className="th">🤍 الأبيض</th>
                <th className="th">🩶 الفضي</th>
                <th className="th">🟠 البرتقالي</th>
              </tr>
            </thead>
            <tbody>
              {BENEFITS.map((b) => (
                <tr key={b[0]}>
                  <td className="td font-bold">{b[0]}</td>
                  <td className="td">{b[1]}</td>
                  <td className="td">{b[2]}</td>
                  <td className="td">{b[3]}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Card>

      <Card title="⚙️ قواعد الترقية والتنزيل">
        <ul className="text-sm space-y-1.5 text-darb-ink/90">
          <li>🔼 الترقية: تلقائية فور تجاوز الحد الأدنى للمستوى التالي.</li>
          <li>🔽 التنزيل: بعد {inputs.spendPeriodMonths} شهر من عدم الوصول للحد الأدنى.</li>
          <li>📅 المراجعة: 31 ديسمبر · يُحتسب آخر {inputs.spendPeriodMonths} شهر.</li>
          <li>🎯 فترة سماح: {inputs.graceDays} يوم بعد التنزيل لاستعادة المستوى.</li>
          <li>💯 النقاط لا تنتهي · باقية حتى لو تغيّر المستوى.</li>
        </ul>
      </Card>
    </div>
  );
}
