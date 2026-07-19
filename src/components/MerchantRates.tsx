import { useScenario } from "../context/ScenarioContext";
import { fmtInt, fmtNum } from "../model/engine";
import { useMerchants, type Merchant } from "../model/merchants";
import { Card, Stat, Badge } from "./ui";

export default function MerchantRates() {
  const { inputs } = useScenario();
  const pv = inputs.pointValue;
  const { merchants, setMerchants } = useMerchants();

  const update = (id: number, patch: Partial<Merchant>) =>
    setMerchants((p) => p.map((m) => (m.id === id ? { ...m, ...patch } : m)));
  const remove = (id: number) => setMerchants((p) => p.filter((m) => m.id !== id));
  const add = () =>
    setMerchants((p) => [
      ...p,
      { id: Math.max(0, ...p.map((m) => m.id)) + 1, name: "تاجر جديد", type: "كافيه", margin: 0.5, earnPct: 3, customerPct: 1, burnPct: 5 },
    ]);

  const avgEarn = merchants.length ? merchants.reduce((a, m) => a + m.earnPct, 0) / merchants.length : 0;
  const avgBurn = merchants.length ? merchants.reduce((a, m) => a + m.burnPct, 0) / merchants.length : 0;

  const num = (v: number) => (Number.isFinite(v) ? v : 0);

  return (
    <div className="space-y-4">
      <div>
        <h2 className="text-lg font-extrabold">🎚️ معدلات التجار · Earn & Burn</h2>
        <p className="text-xs text-darb-mut">تتفاوضين مع كل تاجر على معدله الخاص — تُحفظ في المتصفح</p>
      </div>

      <div className="grid sm:grid-cols-3 gap-3">
        <Stat label="عدد التجار" value={fmtInt(merchants.length)} />
        <Stat label="متوسط Earn (شراء)" value={`${fmtNum(avgEarn, 1)}%`} tone="warn" />
        <Stat label="متوسط Burn (استبدال)" value={`${fmtNum(avgBurn, 1)}%`} tone="accent" />
      </div>

      <Card title="🟢 Earn (عند الشراء) · 🔴 Burn (عند الاستبدال)">
        <div className="overflow-x-auto">
          <table className="w-full border-collapse">
            <thead>
              <tr>
                <th className="th">التاجر</th>
                <th className="th">الهامش</th>
                <th className="th">🟢 Earn %</th>
                <th className="th">حصة العميل %</th>
                <th className="th">حصة درب (Earn) %</th>
                <th className="th">🔴 Burn % (رسوم تسويق)</th>
                <th className="th"></th>
              </tr>
            </thead>
            <tbody>
              {merchants.map((m) => {
                const darbEarn = Math.max(0, m.earnPct - m.customerPct);
                const beEarn = m.margin ? m.earnPct / (m.margin * 100) * 100 : 0; // تعادل الـEarn %
                return (
                  <tr key={m.id} className="hover:bg-darb-panel/40">
                    <td className="td">
                      <input value={m.name} onChange={(e) => update(m.id, { name: e.target.value })}
                        className="bg-transparent border-b border-darb-line/50 focus:border-darb-orange outline-none w-28 text-darb-ink font-bold" />
                    </td>
                    <td className="td">
                      <input type="number" value={Math.round(m.margin * 100)} onChange={(e) => update(m.id, { margin: (parseFloat(e.target.value) || 0) / 100 })}
                        className="bg-darb-panel border border-darb-line rounded px-2 py-1 w-14 text-left" dir="ltr" />%
                    </td>
                    <td className="td">
                      <input type="number" step={0.5} value={num(m.earnPct)} onChange={(e) => update(m.id, { earnPct: parseFloat(e.target.value) || 0 })}
                        className="bg-darb-yellow/10 border border-darb-yellow/30 rounded px-2 py-1 w-16 text-left" dir="ltr" />
                    </td>
                    <td className="td">
                      <input type="number" step={0.5} value={num(m.customerPct)} onChange={(e) => update(m.id, { customerPct: parseFloat(e.target.value) || 0 })}
                        className="bg-darb-yellow/10 border border-darb-yellow/30 rounded px-2 py-1 w-16 text-left" dir="ltr" />
                    </td>
                    <td className="td font-bold text-darb-good">{fmtNum(darbEarn, 1)}%</td>
                    <td className="td">
                      <input type="number" step={0.5} value={num(m.burnPct)} onChange={(e) => update(m.id, { burnPct: parseFloat(e.target.value) || 0 })}
                        className="bg-darb-yellow/10 border border-darb-yellow/30 rounded px-2 py-1 w-16 text-left" dir="ltr" />
                      {beEarn > 0 && <div className="text-[9px] text-darb-mut mt-0.5">تعادل ~{fmtNum(beEarn, 1)}%</div>}
                    </td>
                    <td className="td">
                      <button onClick={() => remove(m.id)} className="text-darb-bad hover:opacity-70" title="حذف">✕</button>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
        <button onClick={add} className="mt-3 text-xs font-bold px-3 py-1.5 rounded-lg border border-darb-line hover:border-darb-orange text-darb-mut hover:text-darb-orange transition">
          + إضافة تاجر
        </button>
      </Card>

      <Card title="📖 كيف تُقرأ">
        <ul className="text-sm space-y-1.5 text-darb-ink/90 list-disc pr-5">
          <li><b>🟢 Earn %</b> = ما يدفعه التاجر <b>عند شراء العميل منه</b> (MDR)؛ منها <b>حصة العميل</b> نقاطاً، والباقي لدرب.</li>
          <li><b>🔴 Burn %</b> = رسوم تسويق يدفعها التاجر <b>عند استبدال العميل عنده</b> (لأن درب وصّلت له عميلاً). قيمة المكافأة نفسها من الإسكرو.</li>
          <li>القاعدة: كل ما زاد <b>هامش التاجر</b>، تقدرين ترفعين Earn وBurn (يبقى التعادل سهلاً ~5%).</li>
        </ul>
        <p className="text-[11px] text-darb-mut mt-2">
          مثال (تاجر Earn 3% · حصة العميل 1%): على 100﷼ شراء → التاجر يدفع 3﷼ (العميل {fmtInt(pv)} نقطة = 1﷼ + درب 2﷼).
        </p>
      </Card>

      <Card>
        <div className="flex flex-wrap gap-2">
          <Badge tone="good">Earn = دخل درب من الشراء</Badge>
          <Badge tone="accent">Burn = دخل درب من الاستبدال (تسويق)</Badge>
          <Badge tone="warn">كلاهما متفاوض عليه لكل تاجر</Badge>
        </div>
      </Card>
    </div>
  );
}
