import { useState } from "react";
import { useScenario } from "../context/ScenarioContext";
import { stationPointCalc, fmtNum, fmtPct, fmtSar, fmtInt } from "../model/engine";
import { Card, Stat, NumberInput, Badge } from "./ui";

const FUELS = [
  { key: "p95", label: "بنزين 95", price: 2.33, margin: 13 },
  { key: "p91", label: "بنزين 91", price: 2.18, margin: 13 },
  { key: "diesel", label: "ديزل", price: 1.79, margin: 4 },
];

const feas = (f: "good" | "tight" | "loss") =>
  f === "good" ? (
    <Badge tone="good">✅ مستدام</Badge>
  ) : f === "tight" ? (
    <Badge tone="warn">⚠️ هامش ضيق</Badge>
  ) : (
    <Badge tone="bad">🔴 غير مجدٍ</Badge>
  );

export default function StationCalculator() {
  const { inputs } = useScenario();
  const [fuelKey, setFuelKey] = useState("p95");
  const fuel = FUELS.find((f) => f.key === fuelKey)!;
  const [literPrice, setLiterPrice] = useState(fuel.price);
  const [margin, setMargin] = useState(13); // هللة/لتر — هامش هذي المحطة
  const [liters, setLiters] = useState(80000);
  const [pointCost, setPointCost] = useState(0.84); // هللة/ريال — القرار

  const r = stationPointCalc({
    literPrice,
    marginHalalaPerLiter: margin,
    litersPerMonth: liters,
    pointCostHalala: pointCost,
    pointValue: inputs.pointValue,
  });

  const pickFuel = (key: string) => {
    const f = FUELS.find((x) => x.key === key)!;
    setFuelKey(key);
    setLiterPrice(f.price);
    setMargin(f.margin);
  };
  const useRecommended = () => setPointCost(+r.recommendedHalala.toFixed(2));

  // جدول مرجعي: تكلفة النقطة الموصى بها لهوامش مختلفة (15% من الهامش)
  const refMargins = [4, 6, 8, 10, 13, 15];

  return (
    <div className="space-y-4">
      <div>
        <h2 className="text-lg font-extrabold">🧮 حاسبة النقطة لكل محطة</h2>
        <p className="text-xs text-darb-mut">
          كل محطة لها هامش مختلف — أدخل هامشها الفعلي لتعرف كم هللة تحسب النقطة دون أن تخسر
        </p>
      </div>

      <Card>
        <div className="grid lg:grid-cols-2 gap-4">
          <div className="space-y-3">
            <div>
              <span className="text-xs font-bold text-darb-ink">نوع الوقود</span>
              <div className="flex gap-1.5 mt-1.5">
                {FUELS.map((f) => (
                  <button
                    key={f.key}
                    onClick={() => pickFuel(f.key)}
                    className={`text-xs font-bold px-3 py-1.5 rounded-lg border transition ${
                      fuelKey === f.key
                        ? "border-darb-accent bg-darb-accent/15 text-darb-accent"
                        : "border-darb-line text-darb-mut hover:text-darb-ink"
                    }`}
                  >
                    {f.label}
                  </button>
                ))}
              </div>
            </div>
            <NumberInput label="سعر اللتر (ريال)" value={literPrice} onChange={setLiterPrice} step={0.01} />
            <NumberInput
              label="هامش هذي المحطة (هللة/لتر)"
              value={margin}
              onChange={setMargin}
              step={0.5}
              hint={`= ${fmtNum(r.marginPerRiyal, 2)} هللة لكل ريال مبيعات`}
            />
            <NumberInput label="مبيعات المحطة (لتر/شهر)" value={liters} onChange={setLiters} step={5000} />
            <div>
              <NumberInput
                label="تكلفة النقطة (هللة/ريال)"
                value={pointCost}
                onChange={setPointCost}
                step={0.05}
                status={
                  r.feasible === "good"
                    ? { text: `${fmtPct(r.pctOfMargin, 0)} من الهامش`, tone: "good" }
                    : r.feasible === "tight"
                    ? { text: `${fmtPct(r.pctOfMargin, 0)} من الهامش`, tone: "warn" }
                    : { text: "تتجاوز الهامش!", tone: "bad" }
                }
              />
              <button
                onClick={useRecommended}
                className="mt-2 text-xs font-bold px-3 py-1.5 rounded-lg border border-darb-line hover:border-darb-accent text-darb-mut hover:text-darb-accent transition"
              >
                ↺ استخدم الموصى به ({fmtNum(r.recommendedHalala, 2)} هللة · 15% من الهامش)
              </button>
            </div>
          </div>

          <div className="grid grid-cols-2 gap-3 content-start">
            <Stat label="🎯 كم هللة تحسب النقطة" value={`${fmtNum(r.pointCostHalala, 2)} هللة`} tone="accent" hint="لكل ريال مبيعات" />
            <Stat label="📈 معدل الكسب المكافئ" value={`${fmtInt(r.earnRate)} نقطة/﷼`} hint={`كاش باك ${fmtPct(r.cashbackPct)}`} />
            <Stat label="هامش المحطة/ريال" value={`${fmtNum(r.marginPerRiyal, 2)} هللة`} />
            <Stat
              label="صافي الهامش بعد الولاء"
              value={`${fmtNum(r.netMarginPerRiyal, 2)} هللة`}
              tone={r.netMarginPerRiyal > 0 ? "good" : "bad"}
            />
            <Stat label="💸 تكلفة شهرية على المحطة" value={fmtSar(r.monthlyCost)} tone="warn" />
            <Stat label="💸 تكلفة سنوية" value={fmtSar(r.annualCost)} tone="warn" />
          </div>
        </div>
        <div className="mt-3 flex items-center justify-between stat">
          <span className="text-sm">الجدوى لهذي المحطة عند {fmtNum(r.pointCostHalala, 2)} هللة</span>
          {feas(r.feasible)}
        </div>
        <p className="text-xs text-darb-mut mt-3 leading-relaxed">
          💡 القاعدة: لا تتجاوز تكلفة النقطة <b>~15-20%</b> من هامش المحطة. كلما زاد هامش المحطة،
          قدرت تعطي نقاطاً أكثر. المحطات ذات الهامش الضعيف (الديزل/الامتياز) تحسب النقطة بأقل قيمة.
        </p>
      </Card>

      <Card title="📊 جدول مرجعي · تكلفة النقطة الموصى بها حسب الهامش" subtitle="عند 15% من الهامش · سعر اللتر الحالي">
        <div className="overflow-x-auto">
          <table className="w-full border-collapse">
            <thead>
              <tr>
                <th className="th">هامش المحطة (هللة/لتر)</th>
                <th className="th">هامش/ريال (هللة)</th>
                <th className="th">تكلفة النقطة الموصى بها (هللة)</th>
                <th className="th">معدل الكسب المكافئ</th>
                <th className="th">الكاش باك</th>
              </tr>
            </thead>
            <tbody>
              {refMargins.map((mg) => {
                const rr = stationPointCalc({
                  literPrice,
                  marginHalalaPerLiter: mg,
                  litersPerMonth: liters,
                  pointCostHalala: 0,
                  pointValue: inputs.pointValue,
                });
                const rec = rr.recommendedHalala;
                return (
                  <tr key={mg} className={mg === margin ? "bg-darb-accent/10" : ""}>
                    <td className="td font-bold">{fmtNum(mg, 1)}</td>
                    <td className="td">{fmtNum(rr.marginPerRiyal, 2)}</td>
                    <td className="td text-darb-accent font-bold">{fmtNum(rec, 2)}</td>
                    <td className="td">{fmtInt((rec / 100) * inputs.pointValue)} نقطة/﷼</td>
                    <td className="td">{fmtPct(rec / 100)}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </Card>
    </div>
  );
}
