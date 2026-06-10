import { useState } from "react";
import { useScenario } from "../context/ScenarioContext";
import { simulateCustomer, fmtInt, fmtPct, fmtSar } from "../model/engine";
import { CUSTOMER_SAMPLES } from "../model/defaults";
import { Card, Stat, Bar, Badge } from "./ui";

export default function CustomerSimulator() {
  const { inputs } = useScenario();
  const [monthly, setMonthly] = useState(600);
  const r = simulateCustomer(inputs, monthly);

  return (
    <div className="space-y-4">
      <div>
        <h2 className="text-lg font-extrabold">🎮 محاكي العميل التفاعلي</h2>
        <p className="text-xs text-darb-mut">
          النقاط (للاستبدال) منفصلة تماماً عن المستوى (للحالة)
        </p>
      </div>

      <Card>
        <div className="flex flex-wrap items-end gap-4">
          <label className="block grow min-w-[220px]">
            <span className="text-xs font-bold text-darb-ink">
              💰 الإنفاق الشهري على الوقود (ريال)
            </span>
            <input
              type="range"
              min={100}
              max={10000}
              step={100}
              value={monthly}
              onChange={(e) => setMonthly(parseInt(e.target.value))}
              className="w-full mt-2 accent-darb-orange"
            />
          </label>
          <input
            type="number"
            value={monthly}
            onChange={(e) => setMonthly(parseFloat(e.target.value) || 0)}
            className="w-28 bg-darb-yellow/10 border border-darb-yellow/40 rounded-lg px-3 py-2 font-bold text-left"
            dir="ltr"
          />
        </div>
      </Card>

      <div className="grid sm:grid-cols-2 lg:grid-cols-4 gap-3">
        <Stat label="📅 الإنفاق السنوي" value={fmtSar(r.annualSpend)} />
        <Stat
          label="🎯 المستوى الحالي"
          value={`${r.tier.emoji} ${r.tier.label}`}
          tone="accent"
        />
        <Stat label="💎 نقاط/سنة (للاستبدال)" value={fmtInt(r.pointsPerYear)} />
        <Stat
          label="💰 القيمة المالية"
          value={fmtSar(r.valueRiyal)}
          hint={`كاش باك ${fmtPct(r.cashbackPct)}`}
          tone="good"
        />
      </div>

      <Card title="📊 تقدّمك للمستوى التالي">
        <div className="flex items-center justify-between text-sm mb-2">
          <span>
            المستوى التالي: <b>{r.nextTierLabel}</b>
          </span>
          <span className="text-darb-mut">
            {r.remainingSpend > 0
              ? `باقي ${fmtSar(r.remainingSpend)} إنفاق`
              : "وصلت أعلى مستوى ✨"}
          </span>
        </div>
        <Bar value={r.progress} tone="orange" />
        <div className="text-left text-xs text-darb-mut mt-1">{fmtPct(r.progress, 0)}</div>
      </Card>

      <Card title="🔍 الفصل بين النقاط والمستوى">
        <div className="grid sm:grid-cols-2 gap-3 text-sm">
          <div className="stat">
            <div className="font-bold text-darb-accent mb-1">💎 النقاط (للاستبدال)</div>
            <ul className="space-y-1 text-darb-ink/90">
              <li>عند الإنفاق: +{fmtInt(monthly * r.tier.earnRate)} نقطة شهرياً</li>
              <li>بعد سنة: {fmtInt(r.pointsPerYear)} نقطة (= {fmtSar(r.valueRiyal)})</li>
              <li>عند الاستبدال: تنقص بحسب المكافأة</li>
              <li>التجديد السنوي: الرصيد يبقى كما هو</li>
            </ul>
          </div>
          <div className="stat">
            <div className="font-bold text-darb-orange mb-1">🎯 المستوى (للحالة)</div>
            <ul className="space-y-1 text-darb-ink/90">
              <li>عند الإنفاق: +{fmtInt(monthly)} ﷼ للعدّاد</li>
              <li>بعد سنة: {fmtSar(r.annualSpend)} إنفاق</li>
              <li>عند الاستبدال: لا يتأثر · المستوى محفوظ</li>
              <li>التجديد: العدّاد من 0 · المستوى محفوظ 12 شهر</li>
            </ul>
          </div>
        </div>
      </Card>

      <Card title="🎲 عيّنات نموذجية" subtitle="اضغط أي عميل لتطبيق إنفاقه">
        <div className="overflow-x-auto">
          <table className="w-full border-collapse">
            <thead>
              <tr>
                <th className="th">العميل</th>
                <th className="th">إنفاق/شهر</th>
                <th className="th">إنفاق/سنة</th>
                <th className="th">المستوى</th>
                <th className="th">النقاط/سنة</th>
                <th className="th">القيمة</th>
                <th className="th">تكلفة درب</th>
              </tr>
            </thead>
            <tbody>
              {CUSTOMER_SAMPLES.map((c) => {
                const sr = simulateCustomer(inputs, c.monthly);
                const active = c.monthly === monthly;
                return (
                  <tr
                    key={c.label}
                    onClick={() => setMonthly(c.monthly)}
                    className={`cursor-pointer hover:bg-darb-panel/60 ${
                      active ? "bg-darb-accent/10" : ""
                    }`}
                  >
                    <td className="td font-bold">{c.label}</td>
                    <td className="td">{fmtSar(c.monthly)}</td>
                    <td className="td">{fmtSar(sr.annualSpend)}</td>
                    <td className="td">
                      {sr.tier.emoji} {sr.tier.label}
                    </td>
                    <td className="td">{fmtInt(sr.pointsPerYear)}</td>
                    <td className="td text-darb-accent">{fmtSar(sr.valueRiyal)}</td>
                    <td className="td">
                      <Badge tone={sr.darbCostRiyal > 150 ? "warn" : "good"}>
                        {fmtSar(sr.darbCostRiyal)}
                      </Badge>
                    </td>
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
