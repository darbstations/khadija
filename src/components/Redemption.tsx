import { useState } from "react";
import { redemptionBlend, fmtPct, fmtSar } from "../model/engine";
import { REDEMPTION_CHANNELS } from "../model/defaults";
import { Card, Stat, NumberInput, Badge } from "./ui";

interface ChannelState {
  key: string;
  label: string;
  mix: number;
  costFactor: number;
  fundedBy: string;
  note: string;
}

export default function Redemption() {
  const [totalValue, setTotalValue] = useState(1_000_000); // إجمالي القيمة المستبدلة سنوياً (ريال)
  const [channels, setChannels] = useState<ChannelState[]>(
    REDEMPTION_CHANNELS.map((c) => ({
      key: c.key,
      label: c.label,
      mix: c.defaultMix,
      costFactor: c.darbCostFactor,
      fundedBy: c.fundedBy,
      note: c.note,
    }))
  );

  const blend = redemptionBlend(totalValue, channels);
  const mixOk = Math.abs(blend.mixSum - 1) < 0.001;

  const setMix = (key: string, v: number) =>
    setChannels((p) => p.map((c) => (c.key === key ? { ...c, mix: v / 100 } : c)));
  const setCost = (key: string, v: number) =>
    setChannels((p) => p.map((c) => (c.key === key ? { ...c, costFactor: v / 100 } : c)));

  return (
    <div className="space-y-4">
      <div>
        <h2 className="text-lg font-extrabold">🔁 قنوات الاستبدال · تكلفتها على درب</h2>
        <p className="text-xs text-darb-mut">
          العميل يجمع نقاطاً من البنزين والمستأجرين · ويستبدلها عبر 4 قنوات لكل منها تكلفة مختلفة
        </p>
      </div>

      <div className="grid sm:grid-cols-4 gap-3">
        <Stat label="إجمالي القيمة المستبدلة" value={fmtSar(totalValue)} hint="سنوياً" />
        <Stat label="💸 تكلفة درب الفعلية" value={fmtSar(blend.totalDarbCost)} tone="warn" />
        <Stat label="متوسط تكلفة الريال المستبدل" value={fmtPct(blend.blendedFactor)} tone="accent" />
        <Stat label="💰 التوفير مقابل الكل-بنزين" value={fmtSar(blend.savings)} tone="good" />
      </div>

      <Card title="🎛️ مزيج قنوات الاستبدال" subtitle="عدّل نسبة كل قناة وتكلفتها على درب">
        <div className="mb-3">
          <NumberInput
            label="إجمالي القيمة المستبدلة سنوياً (ريال)"
            value={totalValue}
            onChange={setTotalValue}
            step={100000}
          />
        </div>

        <div className="space-y-4">
          {channels.map((c) => (
            <div key={c.key} className="stat">
              <div className="flex items-center justify-between mb-2 flex-wrap gap-2">
                <span className="font-bold text-darb-ink">{c.label}</span>
                <div className="flex items-center gap-2">
                  <Badge tone={c.costFactor <= 0.01 ? "good" : c.costFactor >= 1 ? "bad" : "warn"}>
                    تكلفة درب {fmtPct(c.costFactor, 0)}
                  </Badge>
                  <span className="text-[11px] text-darb-mut">تمويل: {c.fundedBy}</span>
                </div>
              </div>
              <div className="grid sm:grid-cols-2 gap-3 items-center">
                <div>
                  <div className="flex items-center justify-between text-xs mb-1">
                    <span className="text-darb-mut">النسبة من الاستبدال</span>
                    <b className="text-darb-ink">{fmtPct(c.mix, 0)}</b>
                  </div>
                  <input
                    type="range"
                    min={0}
                    max={100}
                    step={5}
                    value={Math.round(c.mix * 100)}
                    onChange={(e) => setMix(c.key, parseInt(e.target.value))}
                    className="w-full accent-darb-accent"
                  />
                </div>
                <div>
                  <div className="flex items-center justify-between text-xs mb-1">
                    <span className="text-darb-mut">تكلفة درب لكل ريال (%)</span>
                    <b className="text-darb-ink">{fmtPct(c.costFactor, 0)}</b>
                  </div>
                  <input
                    type="range"
                    min={0}
                    max={100}
                    step={1}
                    value={Math.round(c.costFactor * 100)}
                    onChange={(e) => setCost(c.key, parseInt(e.target.value))}
                    className="w-full accent-darb-warn"
                  />
                </div>
              </div>
              <p className="text-[11px] text-darb-mut mt-2">{c.note}</p>
            </div>
          ))}
        </div>

        {!mixOk && (
          <p className="text-xs text-darb-warn mt-3">
            ⚠️ مجموع النسب = {fmtPct(blend.mixSum, 0)} (يُفضّل 100%). النتائج تُحسب على المجموع الحالي.
          </p>
        )}
      </Card>

      <Card title="📊 تفصيل التكلفة لكل قناة">
        <div className="overflow-x-auto">
          <table className="w-full border-collapse">
            <thead>
              <tr>
                <th className="th">القناة</th>
                <th className="th">من يموّلها</th>
                <th className="th">القيمة المستبدلة</th>
                <th className="th">تكلفة درب/ريال</th>
                <th className="th">تكلفة درب الكلية</th>
              </tr>
            </thead>
            <tbody>
              {blend.rows.map((r) => (
                <tr key={r.key}>
                  <td className="td font-bold">{r.label}</td>
                  <td className="td text-xs">{r.fundedBy}</td>
                  <td className="td">{fmtSar(r.value)}</td>
                  <td className="td">{fmtPct(r.costFactor, 0)}</td>
                  <td className="td text-darb-warn font-bold">{fmtSar(r.darbCost)}</td>
                </tr>
              ))}
              <tr className="bg-darb-panel/60 font-bold">
                <td className="td" colSpan={2}>
                  📊 الإجمالي
                </td>
                <td className="td">{fmtSar(totalValue * blend.mixSum)}</td>
                <td className="td">{fmtPct(blend.blendedFactor, 0)}</td>
                <td className="td text-darb-warn">{fmtSar(blend.totalDarbCost)}</td>
              </tr>
            </tbody>
          </table>
        </div>
        <p className="text-xs text-darb-mut mt-3 leading-relaxed">
          💡 <b>الاستراتيجية:</b> وجّهي الاستبدال نحو <b>عروض المستأجرين</b> (تكلفتها ≈ صفر لأن المستأجر
          يموّلها) بدل <b>خصم البنزين</b> (يأكل هامشك الرفيع). الشركات الخارجية وكرت الشحن قناتان جذابتان
          بتكلفة معقولة (٩٥–٩٧٪) للتنويع. كل ١٠٪ تنقلينها من البنزين لعروض المستأجرين توفّر{" "}
          {fmtSar(totalValue * 0.1)} سنوياً.
        </p>
      </Card>
    </div>
  );
}
