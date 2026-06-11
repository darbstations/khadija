import { useScenario } from "../context/ScenarioContext";
import { tiers, pointMath, fmtInt, fmtPct, fmtNum } from "../model/engine";
import { Card, Stat, Badge } from "./ui";

export default function ExecutiveSummary() {
  const { inputs } = useScenario();
  const t = tiers(inputs);
  const m = pointMath(inputs);
  const sample = [5000, 20000, 50000]; // إنفاق نموذجي لكل مستوى (سنوي)

  return (
    <div className="space-y-4">
      <div>
        <h2 className="text-lg font-extrabold">📊 الملخص التنفيذي</h2>
        <p className="text-xs text-darb-mut">
          نقاط أكبر من المنافس · قيمة استبدال واضحة ({fmtInt(inputs.pointValue)} نقطة = ريال)
        </p>
      </div>

      <div className="grid sm:grid-cols-3 gap-3">
        {t.map((tier) => (
          <Stat
            key={tier.key}
            label={`${tier.emoji} الكاش باك (${tier.label})`}
            value={fmtPct(tier.cashbackPct)}
            hint={`${tier.earnRate} نقطة لكل ريال`}
            tone="accent"
          />
        ))}
      </div>

      <Card title="🎯 المستويات الثلاثة · تفاصيل شاملة">
        <div className="overflow-x-auto">
          <table className="w-full border-collapse">
            <thead>
              <tr>
                <th className="th">المستوى</th>
                <th className="th">الكسب (نقطة/ريال)</th>
                <th className="th">إنفاق نموذجي (سنوي)</th>
                <th className="th">النقاط/سنة</th>
                <th className="th">القيمة بالريال</th>
                <th className="th">الكاش باك</th>
                <th className="th">تكلفة درب/سنة</th>
              </tr>
            </thead>
            <tbody>
              {t.map((tier, idx) => {
                const spend = sample[idx];
                const pts = spend * tier.earnRate;
                const val = pts / inputs.pointValue;
                const cost = (spend * inputs.costPerRiyalHalala) / 100;
                return (
                  <tr key={tier.key}>
                    <td className="td font-bold">
                      {tier.emoji} {tier.label}
                    </td>
                    <td className="td">{tier.earnRate}</td>
                    <td className="td">{fmtInt(spend)} ﷼</td>
                    <td className="td">{fmtInt(pts)}</td>
                    <td className="td text-darb-accent font-bold">{fmtInt(val)} ﷼</td>
                    <td className="td">{fmtPct(tier.cashbackPct)}</td>
                    <td className="td">{fmtNum(cost, 0)} ﷼</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </Card>

      <Card title="🆚 مقارنة درب مع ساسكو · الميزة التنافسية">
        <div className="overflow-x-auto">
          <table className="w-full border-collapse">
            <thead>
              <tr>
                <th className="th">المعيار</th>
                <th className="th">ساسكو</th>
                <th className="th">🤍 الأبيض</th>
                <th className="th">🩶 الفضي</th>
                <th className="th">🟠 البرتقالي</th>
                <th className="th">ميزة درب</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td className="td font-bold">نقاط لكل ريال</td>
                <td className="td">{fmtNum(m.sascoPointsPerRiyal, 2)}</td>
                <td className="td">{t[0].earnRate}</td>
                <td className="td">{t[1].earnRate}</td>
                <td className="td">{t[2].earnRate}</td>
                <td className="td">
                  <Badge tone="accent">
                    {fmtInt(t[0].earnRate / m.sascoPointsPerRiyal)}x –{" "}
                    {fmtInt(t[2].earnRate / m.sascoPointsPerRiyal)}x
                  </Badge>
                </td>
              </tr>
              <tr>
                <td className="td font-bold">الكاش باك الفعلي</td>
                <td className="td">{fmtPct(m.sascoPointsPerRiyal / 100 / inputs.literPrice, 2)}</td>
                <td className="td">{fmtPct(t[0].cashbackPct)}</td>
                <td className="td">{fmtPct(t[1].cashbackPct)}</td>
                <td className="td">{fmtPct(t[2].cashbackPct)}</td>
                <td className="td">
                  <Badge tone="good">إحساس بالثراء</Badge>
                </td>
              </tr>
              <tr>
                <td className="td font-bold">نقاط لـ 100 ريال</td>
                <td className="td">{fmtInt((100 / inputs.literPrice))}</td>
                <td className="td">{fmtInt(100 * t[0].earnRate)}</td>
                <td className="td">{fmtInt(100 * t[1].earnRate)}</td>
                <td className="td">{fmtInt(100 * t[2].earnRate)}</td>
                <td className="td">
                  <Badge tone="accent">أرقام ضخمة</Badge>
                </td>
              </tr>
            </tbody>
          </table>
        </div>
      </Card>

      <Card title="🎯 التوصيات الاستراتيجية">
        <ul className="text-sm space-y-2 text-darb-ink/90 leading-relaxed list-disc pr-5">
          <li>نقاط أكبر من ساسكو ({fmtInt(inputs.pointValue)} نقطة = ريال) دون مبالغة تكشف ضعف القيمة.</li>
          <li>
            الوقود يجذب والشركاء يموّلون: كسب الوقود قليل (~0.5–1%)، والكرم (3–5%) عند الشركاء
            ومموّل منهم.
          </li>
          <li>هدية ترحيب {fmtInt(inputs.welcomeGift)} نقطة (تكلفتها {fmtNum(inputs.welcomeGift / inputs.pointValue, 0)} ﷼ فقط).</li>
          <li>وجّه الاستبدال نحو مكافآت الشركاء لا خصم الوقود (يحمي الهامش الرفيع).</li>
          <li>الديزل هامشه ضعيف ({fmtNum(inputs.dieselMarginHalala, 1)} هللة) → كسب منخفض جداً.</li>
        </ul>
      </Card>
    </div>
  );
}
