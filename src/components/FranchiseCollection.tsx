import { useScenario } from "../context/ScenarioContext";
import { franchiseCollection, fmtInt, fmtSar, fmtNum } from "../model/engine";
import { Card, Stat } from "./ui";

export default function FranchiseCollection() {
  const { inputs } = useScenario();
  const { rows, monthly, annual } = franchiseCollection(inputs);

  return (
    <div className="space-y-4">
      <div>
        <h2 className="text-lg font-extrabold">🤝 التحصيل من محطات الامتياز</h2>
        <p className="text-xs text-darb-mut">
          فاتورة شهرية تلقائية لكل محطة بناءً على لتراتها — اللتر × السعر × التكلفة × الحصة
        </p>
      </div>

      <div className="grid sm:grid-cols-3 gap-3">
        <Stat label="📊 إنفاق شهري (10 محطات)" value={fmtSar(monthly.spend)} />
        <Stat label="مساهمة درب/شهر" value={fmtSar(monthly.darbPay)} tone="warn" />
        <Stat label="مساهمة المالك/شهر" value={fmtSar(monthly.ownerPay)} tone="good" />
      </div>

      <Card>
        <div className="overflow-x-auto">
          <table className="w-full border-collapse">
            <thead>
              <tr>
                <th className="th">الكود</th>
                <th className="th">المدينة</th>
                <th className="th">لتر/شهر</th>
                <th className="th">الإنفاق (﷼)</th>
                <th className="th">مساهمة المالك</th>
                <th className="th">مساهمة درب</th>
              </tr>
            </thead>
            <tbody>
              {rows.map((r) => (
                <tr key={r.code}>
                  <td className="td">{r.code}</td>
                  <td className="td">{r.city}</td>
                  <td className="td">{fmtInt(r.liters)}</td>
                  <td className="td">{fmtInt(r.spend)}</td>
                  <td className="td text-darb-good">{fmtNum(r.ownerPay, 2)}</td>
                  <td className="td text-darb-warn">{fmtNum(r.darbPay, 2)}</td>
                </tr>
              ))}
              <tr className="bg-darb-panel/60 font-bold">
                <td className="td" colSpan={2}>
                  📊 الإجمالي الشهري
                </td>
                <td className="td">{fmtInt(monthly.liters)}</td>
                <td className="td">{fmtInt(monthly.spend)}</td>
                <td className="td text-darb-good">{fmtNum(monthly.ownerPay, 2)}</td>
                <td className="td text-darb-warn">{fmtNum(monthly.darbPay, 2)}</td>
              </tr>
              <tr className="font-bold">
                <td className="td" colSpan={2}>
                  📅 الإجمالي السنوي
                </td>
                <td className="td">{fmtInt(annual.liters)}</td>
                <td className="td">{fmtInt(annual.spend)}</td>
                <td className="td text-darb-good">{fmtInt(annual.ownerPay)}</td>
                <td className="td text-darb-warn">{fmtInt(annual.darbPay)}</td>
              </tr>
            </tbody>
          </table>
        </div>
      </Card>

      <Card title="⚙️ آلية التحصيل">
        <ol className="text-sm space-y-1 text-darb-ink/90 list-decimal pr-5">
          <li>نهاية كل شهر: النظام يحسب إجمالي اللترات لكل محطة.</li>
          <li>المساهمة = اللترات × سعر اللتر × {fmtNum(inputs.costPerRiyalHalala, 1)} هللة × {fmtNum(inputs.darbShareFranchise * 100, 0)}%.</li>
          <li>فاتورة إلكترونية تلقائية في اليوم الأول.</li>
          <li>مهلة سداد {inputs.graceDays} يوم · غرامة 1% أسبوعياً بعدها.</li>
        </ol>
      </Card>
    </div>
  );
}
