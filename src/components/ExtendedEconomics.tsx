import { useScenario } from "../context/ScenarioContext";
import { extendedEconomics, fmtNum, fmtPct } from "../model/engine";
import { Card, Stat, Badge } from "./ui";

const feasBadge = (f: "good" | "tight" | "loss") =>
  f === "good" ? (
    <Badge tone="good">✅ مربح</Badge>
  ) : f === "tight" ? (
    <Badge tone="warn">⚠️ هامش ضيق</Badge>
  ) : (
    <Badge tone="bad">🔴 خسارة</Badge>
  );

export default function ExtendedEconomics() {
  const { inputs } = useScenario();
  const e = extendedEconomics(inputs);

  return (
    <div className="space-y-4">
      <div>
        <h2 className="text-lg font-extrabold">🆕 الاقتصاد الموسّع — الديزل والشركاء</h2>
        <p className="text-xs text-darb-mut">
          مربوط فعلياً بالمدخلات · جوهر الفكرة: الوقود يجذب والشركاء يموّلون
        </p>
      </div>

      <div className="grid lg:grid-cols-2 gap-4">
        <Card title="⛽ الديزل — هامش رفيع جداً">
          <div className="grid grid-cols-2 gap-3">
            <Stat label="سعر اللتر" value={`${fmtNum(e.diesel.price, 2)} ﷼`} />
            <Stat
              label="هامش الديزل"
              value={`${fmtNum(e.diesel.marginHalala, 1)} هللة`}
              hint={fmtPct(e.diesel.marginPct)}
            />
            <Stat label="كسب الديزل" value={`${e.diesel.earn} نقطة/﷼`} hint={fmtPct(e.diesel.cashback)} />
            <Stat
              label="تكلفة درب"
              value={`${fmtNum(e.diesel.darbCostHalala, 2)} هللة/﷼`}
            />
          </div>
          <div className="mt-3 flex items-center justify-between stat">
            <span className="text-sm">صافي هامش درب بعد الولاء</span>
            <span className="flex items-center gap-2 font-bold">
              {fmtNum(e.diesel.netMarginHalala, 2)} هللة/﷼ {feasBadge(e.diesel.feasible)}
            </span>
          </div>
          <p className="text-xs text-darb-mut mt-2">
            💡 هامش الديزل {fmtNum(e.diesel.marginHalala, 1)} هللة فقط — لا يحتمل كسباً كبيراً.
            أبقِ كسب الديزل منخفضاً جداً أو استثنِه من مضاعفات المستويات.
          </p>
        </Card>

        <Card title="☕ الشركاء — مصدر الكرم والربح">
          <div className="grid grid-cols-2 gap-3">
            <Stat label="كسب العميل عند الشريك" value={`${e.partner.earn} نقطة/﷼`} hint={fmtPct(e.partner.cashback)} tone="accent" />
            <Stat label="تكلفة درب" value="≈ 0 ﷼" tone="good" hint="الشريك يموّل المكافأة" />
            <Stat
              label="مساهمة الشريك لدرب"
              value={`${fmtNum(e.partner.contributionHalala, 1)} هللة/﷼`}
              tone="good"
            />
            <Stat
              label="صافي لدرب من الشريك"
              value={`+${fmtNum(e.partner.netToDarbHalala, 1)} هللة/﷼`}
              tone="good"
            />
          </div>
          <p className="text-xs text-darb-mut mt-3">
            💡 العميل يشعر بكاش باك {fmtPct(e.partner.cashback)} (أضعاف الوقود)، بينما درب
            <b className="text-darb-good"> تربح</b> من مساهمة الشريك بدل أن تدفع.
          </p>
        </Card>
      </div>

      <Card title="🔬 مقارنة ريال الوقود مقابل ريال الشريك (أثره على درب)">
        <div className="overflow-x-auto">
          <table className="w-full border-collapse">
            <thead>
              <tr>
                <th className="th">البند</th>
                <th className="th">⛽ ريال على البنزين</th>
                <th className="th">☕ ريال عند الشريك</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td className="td font-bold">هامش/دخل درب</td>
                <td className="td">+{fmtNum(e.fuel.darbMarginHalala, 2)} هللة (ربح بيع)</td>
                <td className="td text-darb-good">+{fmtNum(e.partner.netToDarbHalala, 2)} هللة (مساهمة)</td>
              </tr>
              <tr>
                <td className="td font-bold">تكلفة الولاء على درب</td>
                <td className="td text-darb-bad">−{fmtNum(e.fuel.loyaltyCostHalala, 2)} هللة</td>
                <td className="td">0 هللة (الشريك يموّل)</td>
              </tr>
              <tr>
                <td className="td font-bold">الصافي على درب</td>
                <td className="td">
                  <Badge tone={e.fuel.netHalala > 0 ? "good" : "bad"}>
                    {fmtNum(e.fuel.netHalala, 2)} هللة/﷼
                  </Badge>
                </td>
                <td className="td">
                  <Badge tone="good">+{fmtNum(e.partner.netToDarbHalala, 2)} هللة/﷼</Badge>
                </td>
              </tr>
            </tbody>
          </table>
        </div>
        <p className="text-xs text-darb-mut mt-3 leading-relaxed">
          الخلاصة: على الوقود، البرنامج تكلفة احتفاظ (هامش رفيع). الربح الحقيقي يأتي من
          توسيع شبكة الشركاء وتوجيه الاستبدال نحوهم — لذلك أولوية درب الاستراتيجية = توقيع الشركاء.
        </p>
      </Card>
    </div>
  );
}
