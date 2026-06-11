import { useState, useEffect } from "react";
import { useScenario } from "../context/ScenarioContext";
import { stationPointCalc, fmtNum, fmtPct, fmtSar, fmtInt } from "../model/engine";
import { FUEL_TYPES, STATION_SAMPLES } from "../model/defaults";
import { Card, Stat, Badge } from "./ui";

interface StationRow {
  id: number;
  name: string;
  city: string;
  fuel: string;
  margin: number; // هللة/لتر
  liters: number; // لتر/شهر
}

const STORAGE = "tanki.stations.v1";
const priceOf = (fuelKey: string) =>
  FUEL_TYPES.find((f) => f.key === fuelKey)?.price ?? 2.33;
const fuelLabel = (k: string) => FUEL_TYPES.find((f) => f.key === k)?.label ?? k;

export default function StationsDB() {
  const { inputs } = useScenario();
  const [budgetPct, setBudgetPct] = useState(15); // % من الهامش يُخصص للولاء
  const [stations, setStations] = useState<StationRow[]>(() => {
    try {
      const s = localStorage.getItem(STORAGE);
      if (s) return JSON.parse(s);
    } catch {
      /* تجاهل */
    }
    return STATION_SAMPLES.map((s, idx) => ({ id: idx + 1, ...s }));
  });

  useEffect(() => {
    try {
      localStorage.setItem(STORAGE, JSON.stringify(stations));
    } catch {
      /* تجاهل */
    }
  }, [stations]);

  const calc = (s: StationRow) =>
    stationPointCalc({
      literPrice: priceOf(s.fuel),
      marginHalalaPerLiter: s.margin,
      litersPerMonth: s.liters,
      pointCostHalala: 0, // سنستخدم الموصى به
      pointValue: inputs.pointValue,
      budgetPctOfMargin: budgetPct / 100,
    });

  // كل محطة عند تكلفة النقطة الموصى بها (نسبة من هامشها)
  const detailed = stations.map((s) => {
    const base = calc(s);
    const atRec = stationPointCalc({
      literPrice: priceOf(s.fuel),
      marginHalalaPerLiter: s.margin,
      litersPerMonth: s.liters,
      pointCostHalala: +base.recommendedHalala.toFixed(3),
      pointValue: inputs.pointValue,
      budgetPctOfMargin: budgetPct / 100,
    });
    return { s, r: atRec };
  });

  const totals = detailed.reduce(
    (a, { r }) => ({
      liters: a.liters + r.monthlyRevenue / r.literPrice,
      revenue: a.revenue + r.monthlyRevenue,
      monthlyCost: a.monthlyCost + r.monthlyCost,
    }),
    { liters: 0, revenue: 0, monthlyCost: 0 }
  );
  const blendedCashback = totals.revenue ? totals.monthlyCost / totals.revenue : 0;

  const update = (id: number, patch: Partial<StationRow>) =>
    setStations((p) => p.map((s) => (s.id === id ? { ...s, ...patch } : s)));
  const add = () =>
    setStations((p) => [
      ...p,
      {
        id: Math.max(0, ...p.map((s) => s.id)) + 1,
        name: "محطة جديدة",
        city: "",
        fuel: "p95",
        margin: 13,
        liters: 80000,
      },
    ]);
  const remove = (id: number) => setStations((p) => p.filter((s) => s.id !== id));

  const feas = (f: "good" | "tight" | "loss") =>
    f === "good" ? (
      <Badge tone="good">✅</Badge>
    ) : f === "tight" ? (
      <Badge tone="warn">⚠️</Badge>
    ) : (
      <Badge tone="bad">🔴</Badge>
    );

  return (
    <div className="space-y-4">
      <div>
        <h2 className="text-lg font-extrabold">⛽ محطاتي · تكلفة النقطة لكل محطة</h2>
        <p className="text-xs text-darb-mut">
          أدخل هامش كل محطة الفعلي · النقطة تُحسب حسب هامشها · تُحفظ في المتصفح
        </p>
      </div>

      <div className="grid sm:grid-cols-4 gap-3">
        <Stat label="عدد المحطات" value={fmtInt(stations.length)} />
        <Stat label="إجمالي المبيعات/شهر" value={fmtSar(totals.revenue)} />
        <Stat label="💸 تكلفة الولاء/شهر" value={fmtSar(totals.monthlyCost)} tone="warn" hint={`${fmtSar(totals.monthlyCost * 12)} سنوياً`} />
        <Stat label="متوسط الكاش باك" value={fmtPct(blendedCashback)} tone="accent" />
      </div>

      <Card>
        <div className="flex items-center gap-3 flex-wrap mb-3">
          <span className="text-xs font-bold text-darb-ink">نسبة الهامش المخصصة للولاء:</span>
          <input
            type="range"
            min={5}
            max={30}
            step={1}
            value={budgetPct}
            onChange={(e) => setBudgetPct(parseInt(e.target.value))}
            className="accent-darb-accent grow max-w-xs"
          />
          <Badge tone={budgetPct <= 20 ? "good" : "warn"}>{budgetPct}% من الهامش</Badge>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full border-collapse">
            <thead>
              <tr>
                <th className="th">المحطة</th>
                <th className="th">المدينة</th>
                <th className="th">الوقود</th>
                <th className="th">هامش (هللة/لتر)</th>
                <th className="th">لتر/شهر</th>
                <th className="th">🎯 تكلفة النقطة (هللة)</th>
                <th className="th">الكسب (نقطة/﷼)</th>
                <th className="th">تكلفة/شهر</th>
                <th className="th">جدوى</th>
                <th className="th"></th>
              </tr>
            </thead>
            <tbody>
              {detailed.map(({ s, r }) => (
                <tr key={s.id} className="hover:bg-darb-panel/50">
                  <td className="td">
                    <input
                      value={s.name}
                      onChange={(e) => update(s.id, { name: e.target.value })}
                      className="bg-transparent border-b border-darb-line/50 focus:border-darb-accent outline-none w-28 text-darb-ink font-bold"
                    />
                  </td>
                  <td className="td">
                    <input
                      value={s.city}
                      onChange={(e) => update(s.id, { city: e.target.value })}
                      className="bg-transparent border-b border-darb-line/50 focus:border-darb-accent outline-none w-16 text-darb-ink"
                    />
                  </td>
                  <td className="td">
                    <select
                      value={s.fuel}
                      onChange={(e) => update(s.id, { fuel: e.target.value })}
                      className="bg-darb-panel border border-darb-line rounded px-1.5 py-1 text-xs"
                    >
                      {FUEL_TYPES.map((f) => (
                        <option key={f.key} value={f.key}>
                          {fuelLabel(f.key)}
                        </option>
                      ))}
                    </select>
                  </td>
                  <td className="td">
                    <input
                      type="number"
                      value={s.margin}
                      onChange={(e) => update(s.id, { margin: parseFloat(e.target.value) || 0 })}
                      className="bg-darb-yellow/10 border border-darb-yellow/30 rounded px-2 py-1 w-16 text-left"
                      dir="ltr"
                    />
                  </td>
                  <td className="td">
                    <input
                      type="number"
                      value={s.liters}
                      onChange={(e) => update(s.id, { liters: parseFloat(e.target.value) || 0 })}
                      className="bg-darb-yellow/10 border border-darb-yellow/30 rounded px-2 py-1 w-20 text-left"
                      dir="ltr"
                    />
                  </td>
                  <td className="td text-darb-accent font-bold">{fmtNum(r.pointCostHalala, 2)}</td>
                  <td className="td">{fmtInt(r.earnRate)}</td>
                  <td className="td text-darb-warn">{fmtSar(r.monthlyCost)}</td>
                  <td className="td">{feas(r.feasible)}</td>
                  <td className="td">
                    <button onClick={() => remove(s.id)} className="text-darb-bad hover:opacity-70" title="حذف">
                      ✕
                    </button>
                  </td>
                </tr>
              ))}
              <tr className="bg-darb-panel/60 font-bold">
                <td className="td" colSpan={4}>
                  📊 الإجمالي
                </td>
                <td className="td">{fmtInt(totals.liters)}</td>
                <td className="td">—</td>
                <td className="td">—</td>
                <td className="td text-darb-warn">{fmtSar(totals.monthlyCost)}</td>
                <td className="td" colSpan={2}></td>
              </tr>
            </tbody>
          </table>
        </div>
        <button
          onClick={add}
          className="mt-3 text-xs font-bold px-3 py-1.5 rounded-lg border border-darb-line hover:border-darb-accent text-darb-mut hover:text-darb-accent transition"
        >
          + إضافة محطة
        </button>
      </Card>

      <Card>
        <p className="text-sm text-darb-ink/90 leading-relaxed">
          💡 لاحظي أن المحطات ذات الهامش العالي (١٣ هللة) تحسب النقطة بقيمة أكبر، بينما محطات
          <b> الديزل والامتياز</b> (هامش ٤-٦ هللات) تحسب النقطة بأقل قيمة تلقائياً — فلا تخسرين على أي محطة.
          الإجمالي أعلاه يجمع تكلفة الولاء الحقيقية على كل شبكتك.
        </p>
      </Card>
    </div>
  );
}
