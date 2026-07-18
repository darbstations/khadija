import { useScenario } from "../context/ScenarioContext";
import { fmtSar, fmtInt, fmtNum, fmtPct } from "../model/engine";
import { Card, Stat, Badge } from "./ui";

export default function PointValueModel() {
  const { inputs } = useScenario();
  const pv = inputs.pointValue; // نقطة = ريال
  const valuePerPointRiyal = 1 / pv;
  const valuePerPointHalala = valuePerPointRiyal * 100;

  // قيمة الاستبدال حسب القناة (points = النقاط المطلوبة، value = ما يحصله العميل)
  const rows = [
    { icon: "⛽", label: "خصم بنزين", points: pv * 1, value: 1, funded: "درب", note: "قيمة اسمية 1:1" },
    { icon: "☕", label: "قهوة من تاجر", points: 1600, value: 15, funded: "التاجر", note: "التاجر يضيف من هامشه" },
    { icon: "🍔", label: "وجبة من تاجر", points: 3000, value: 25, funded: "التاجر", note: "التاجر يضيف من هامشه" },
    { icon: "🪪", label: "باقة ذيبان", points: pv * 500, value: 529, funded: "درب (باقة)", note: "قيمة تقديرية للمحتوى" },
    { icon: "🎁", label: "قسيمة جرير 10﷼", points: pv * 10, value: 9.7, funded: "درب (بخصم)", note: "شراء بخصم جملة" },
    { icon: "📱", label: "كرت شحن 10﷼", points: pv * 10, value: 9.5, funded: "درب (عمولة)", note: "عمولة المشغّل" },
  ];

  const multBadge = (m: number) =>
    m > 1.05 ? <Badge tone="good">×{fmtNum(m, 2)} أعلى</Badge> : m >= 0.99 ? <Badge tone="accent">×{fmtNum(m, 2)}</Badge> : <Badge tone="warn">×{fmtNum(m, 2)}</Badge>;

  return (
    <div className="space-y-4">
      <div>
        <h2 className="text-lg font-extrabold">💠 نموذج قيمة النقطة والاستبدال</h2>
        <p className="text-xs text-darb-mut">القيمة الأساسية · وقيمة الاستبدال حسب القناة</p>
      </div>

      {/* القيمة الأساسية */}
      <Card title="① القيمة الأساسية للنقطة">
        <div className="grid sm:grid-cols-2 lg:grid-cols-4 gap-3">
          <Stat label="قيمة النقطة" value={`${fmtNum(valuePerPointHalala, 2)} هللة`} hint={`${fmtNum(valuePerPointRiyal, 4)} ﷼`} tone="accent" />
          <Stat label="النقاط لكل ريال" value={`${fmtInt(pv)} نقطة`} />
          <Stat label="كسب البنزين" value={`${fmtInt(inputs.earnWhite)} نقطة/﷼`} hint={fmtPct(inputs.earnWhite / pv)} tone="warn" />
          <Stat label="كسب التجار" value="2 نقطة/﷼" hint="1% يموّله التاجر" tone="good" />
        </div>
        <p className="text-xs text-darb-mut mt-2">القاعدة للعميل: <b>عدد النقاط ÷ {fmtInt(pv)} = قيمتها بالريال</b>.</p>
      </Card>

      {/* قيمة الاستبدال حسب القناة */}
      <Card title="② قيمة الاستبدال حسب القناة">
        <div className="overflow-x-auto">
          <table className="w-full border-collapse">
            <thead>
              <tr>
                <th className="th">المكافأة</th>
                <th className="th">النقاط المطلوبة</th>
                <th className="th">قيمتها الأساسية</th>
                <th className="th">ما يحصله العميل</th>
                <th className="th">المضاعف</th>
                <th className="th">من يموّل</th>
              </tr>
            </thead>
            <tbody>
              {rows.map((r) => {
                const base = r.points / pv;
                const mult = base ? r.value / base : 0;
                return (
                  <tr key={r.label}>
                    <td className="td font-bold">{r.icon} {r.label}</td>
                    <td className="td">{fmtInt(r.points)}</td>
                    <td className="td text-darb-mut">{fmtSar(base)}</td>
                    <td className="td text-darb-accent font-bold">{fmtSar(r.value)}</td>
                    <td className="td">{multBadge(mult)}</td>
                    <td className="td text-xs">{r.funded}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
        <p className="text-xs text-darb-mut mt-3 leading-relaxed">
          💡 <b>عروض التجار تعطي أكثر من القيمة الأساسية</b> (×1.5–1.9) لأن التاجر يضيف من هامشه ليجذب العميل —
          فهي الأفضل للعميل والأرخص على درب. خصم البنزين قيمة اسمية (×1). الخارجي/الشحن أقل قليلاً (رسوم/عمولة).
        </p>
      </Card>

      {/* كم نقطة أحتاج */}
      <Card title="③ كم نقطة أحتاج لـ…">
        <div className="grid sm:grid-cols-2 lg:grid-cols-4 gap-3">
          {[1, 5, 10, 50].map((r) => (
            <Stat key={r} label={`${fmtSar(r)} قيمة`} value={`${fmtInt(r * pv)} نقطة`} />
          ))}
        </div>
        <p className="text-[11px] text-darb-mut mt-2">للاستبدال بقيمة اسمية (خصم بنزين). عند التجار تحصل قيمة أعلى بنفس النقاط.</p>
      </Card>
    </div>
  );
}
