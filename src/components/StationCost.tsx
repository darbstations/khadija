import { useScenario } from "../context/ScenarioContext";
import { stationEconomics, fmtNum } from "../model/engine";
import { Card, Badge } from "./ui";

export default function StationCost() {
  const { inputs } = useScenario();
  const rows = stationEconomics(inputs);
  const feas = (f: "good" | "tight" | "loss") =>
    f === "good" ? (
      <Badge tone="good">✅ مربح</Badge>
    ) : f === "tight" ? (
      <Badge tone="warn">⚠️ هامش ضيق</Badge>
    ) : (
      <Badge tone="bad">🔴 خسارة</Badge>
    );

  return (
    <div className="space-y-4">
      <div>
        <h2 className="text-lg font-extrabold">🏗️ تكلفة الولاء على كل نوع محطة</h2>
        <p className="text-xs text-darb-mut">الجدوى الاقتصادية لكل نموذج تشغيل</p>
      </div>
      <Card>
        <div className="overflow-x-auto">
          <table className="w-full border-collapse">
            <thead>
              <tr>
                <th className="th">البند</th>
                {rows.map((r) => (
                  <th key={r.key} className="th">
                    {r.label}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              <tr>
                <td className="td font-bold">هامش/لتر (هللة)</td>
                {rows.map((r) => (
                  <td key={r.key} className="td">
                    {fmtNum(r.marginPerLiter, 1)}
                  </td>
                ))}
              </tr>
              <tr>
                <td className="td font-bold">هامش/ريال (هللة)</td>
                {rows.map((r) => (
                  <td key={r.key} className="td">
                    {fmtNum(r.marginPerRiyal, 2)}
                  </td>
                ))}
              </tr>
              <tr>
                <td className="td font-bold">تكلفة الولاء/ريال (هللة)</td>
                {rows.map((r) => (
                  <td key={r.key} className="td">
                    {fmtNum(r.loyaltyCostPerRiyal, 2)}
                  </td>
                ))}
              </tr>
              <tr>
                <td className="td font-bold">مساهمة درب %</td>
                {rows.map((r) => (
                  <td key={r.key} className="td">
                    {fmtNum(r.darbSharePct * 100, 0)}%
                  </td>
                ))}
              </tr>
              <tr>
                <td className="td font-bold">تكلفة درب/ريال (هللة)</td>
                {rows.map((r) => (
                  <td key={r.key} className="td">
                    {fmtNum(r.darbCostPerRiyal, 2)}
                  </td>
                ))}
              </tr>
              <tr>
                <td className="td font-bold">صافي هامش المالك/ريال</td>
                {rows.map((r) => (
                  <td key={r.key} className="td font-bold text-darb-accent">
                    {fmtNum(r.ownerNetPerRiyal, 2)}
                  </td>
                ))}
              </tr>
              <tr>
                <td className="td font-bold">الجدوى</td>
                {rows.map((r) => (
                  <td key={r.key} className="td">
                    {feas(r.feasible)}
                  </td>
                ))}
              </tr>
            </tbody>
          </table>
        </div>
        <p className="text-xs text-darb-mut mt-3">
          💡 محطة الامتياز هامشها ضيق ({fmtNum(rows[2].marginPerLiter, 1)} هللة/لتر) — لذلك
          التحصيل منها يكون 50% فقط من تكلفة النقاط عبر فاتورة شهرية تلقائية.
        </p>
      </Card>
    </div>
  );
}
