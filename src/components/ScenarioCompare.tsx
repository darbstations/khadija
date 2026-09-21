import { useScenario } from "../context/ScenarioContext";
import { scenarioComparison, fmtNum, fmtPct } from "../model/engine";
import { Card } from "./ui";

export default function ScenarioCompare() {
  const { inputs } = useScenario();
  const list = scenarioComparison(inputs);

  return (
    <div className="space-y-4">
      <div>
        <h2 className="text-lg font-extrabold">🆚 مقارنة السيناريوهات</h2>
        <p className="text-xs text-darb-mut">
          العمود ⭐ الحالي يعكس مدخلاتك الحيّة · الباقي ثابت للمقارنة
        </p>
      </div>

      <Card>
        <div className="overflow-x-auto">
          <table className="w-full border-collapse">
            <thead>
              <tr>
                <th className="th">المعيار</th>
                {list.map((s) => (
                  <th
                    key={s.key}
                    className={`th ${s.key === "current" ? "text-darb-accent" : ""}`}
                  >
                    {s.label}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              <tr>
                <td className="td font-bold">قيمة النقطة</td>
                {list.map((s) => (
                  <td key={s.key} className="td">
                    {fmtNum(s.pv, 0)}
                  </td>
                ))}
              </tr>
              <tr>
                <td className="td font-bold">كسب الأبيض</td>
                {list.map((s) => (
                  <td key={s.key} className="td">
                    {fmtNum(s.w, 2)}
                  </td>
                ))}
              </tr>
              <tr>
                <td className="td font-bold">كسب الفضي</td>
                {list.map((s) => (
                  <td key={s.key} className="td">
                    {fmtNum(s.s, 2)}
                  </td>
                ))}
              </tr>
              <tr>
                <td className="td font-bold">كسب البرتقالي</td>
                {list.map((s) => (
                  <td key={s.key} className="td">
                    {fmtNum(s.o, 2)}
                  </td>
                ))}
              </tr>
              <tr>
                <td className="td font-bold">كاش باك الأبيض</td>
                {list.map((s) => (
                  <td key={s.key} className="td">
                    {fmtPct(s.cashbackWhite, 2)}
                  </td>
                ))}
              </tr>
              <tr>
                <td className="td font-bold">تكلفة درب (هللة/﷼)</td>
                {list.map((s) => (
                  <td key={s.key} className="td">
                    {fmtNum(s.cost, 2)}
                  </td>
                ))}
              </tr>
              <tr className="bg-darb-panel/60 font-bold">
                <td className="td">💰 تكلفة سنوية (م﷼)*</td>
                {list.map((s) => (
                  <td
                    key={s.key}
                    className={`td ${s.key === "current" ? "text-darb-accent" : "text-darb-warn"}`}
                  >
                    {fmtNum(s.annualCostM, 2)}
                  </td>
                ))}
              </tr>
            </tbody>
          </table>
        </div>
        <p className="text-xs text-darb-mut mt-3">
          * افتراض 230 محطة × 80,000 لتر × 12 شهر × سعر اللتر، بنسبة تنشيط 70%.
        </p>
      </Card>
    </div>
  );
}
