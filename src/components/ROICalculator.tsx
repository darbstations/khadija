import { useState } from "react";
import { useScenario } from "../context/ScenarioContext";
import { projections, fmtNum, fmtPct } from "../model/engine";
import { Card, Stat, NumberInput, Badge } from "./ui";

/** حاسبة ROI تفاعلية — التكاليف تشتق من التوقعات، والإيرادات من افتراضات قابلة للتعديل */
export default function ROICalculator() {
  const { inputs } = useScenario();
  const p = projections(inputs);

  // افتراضات قابلة للتعديل (مليون ريال / نِسَب)
  const [appDev, setAppDev] = useState(1.45); // تطوير التطبيق (5 سنوات)
  const [team, setTeam] = useState(19.78); // فريق الولاء
  const [marketing, setMarketing] = useState(5.7); // تسويق
  const [partnerRevM, setPartnerRevM] = useState(13.7); // عمولة الشركاء (5 سنوات)
  const [fuelUpliftRevM, setFuelUpliftRevM] = useState(80); // زيادة مبيعات الوقود
  const [otherRevM, setOtherRevM] = useState(53.5); // جذب عملاء + توفير + نمو + بيانات

  const pointsCost5y = p.fiveYearTotal; // تكلفة النقاط من المحرك
  const totalCost = appDev + team + marketing + pointsCost5y;
  const totalRev = partnerRevM + fuelUpliftRevM + otherRevM;
  const netProfit = totalRev - totalCost;
  const roi = totalCost > 0 ? netProfit / totalCost : 0;

  const verdict =
    roi > 1
      ? { t: "⭐ مشروع ممتاز", tone: "good" as const }
      : roi > 0.5
      ? { t: "✅ مربح", tone: "good" as const }
      : roi > 0
      ? { t: "🟡 مقبول", tone: "warn" as const }
      : { t: "🔴 إعادة نظر", tone: "bad" as const };

  return (
    <div className="space-y-4">
      <div>
        <h2 className="text-lg font-extrabold">💎 حاسبة العائد على الاستثمار (ROI)</h2>
        <p className="text-xs text-darb-mut">
          تكلفة النقاط مشتقة من محرك التوقعات · عدّل بقية الافتراضات (مليون ريال / 5 سنوات)
        </p>
      </div>

      <div className="grid sm:grid-cols-2 lg:grid-cols-4 gap-3">
        <Stat label="💸 إجمالي التكاليف" value={`${fmtNum(totalCost, 1)} م﷼`} tone="warn" />
        <Stat label="💰 إجمالي الإيرادات" value={`${fmtNum(totalRev, 1)} م﷼`} tone="accent" />
        <Stat label="💎 صافي الربح" value={`${fmtNum(netProfit, 1)} م﷼`} tone={netProfit > 0 ? "good" : "bad"} />
        <Stat label="🎯 ROI" value={fmtPct(roi, 0)} tone={roi > 0 ? "good" : "bad"} />
      </div>

      <div className="flex items-center gap-3">
        <span className="text-sm text-darb-mut">التقييم النهائي:</span>
        <Badge tone={verdict.tone}>{verdict.t}</Badge>
      </div>

      <div className="grid lg:grid-cols-2 gap-4">
        <Card title="💸 التكاليف (5 سنوات · مليون ريال)">
          <div className="space-y-3">
            <div className="stat flex items-center justify-between">
              <span className="text-sm">🎁 كلفة النقاط (من المحرك)</span>
              <span className="font-bold text-darb-warn">{fmtNum(pointsCost5y, 2)} م﷼</span>
            </div>
            <NumberInput label="💻 تطوير التطبيق" value={appDev} onChange={setAppDev} step={0.1} />
            <NumberInput label="👥 فريق الولاء" value={team} onChange={setTeam} step={0.5} />
            <NumberInput label="📣 التسويق" value={marketing} onChange={setMarketing} step={0.5} />
          </div>
        </Card>

        <Card title="💰 الإيرادات (5 سنوات · مليون ريال)">
          <div className="space-y-3">
            <NumberInput label="⛽ زيادة مبيعات الوقود" value={fuelUpliftRevM} onChange={setFuelUpliftRevM} step={1} />
            <NumberInput label="🤝 عمولة الشركاء" value={partnerRevM} onChange={setPartnerRevM} step={0.5} />
            <NumberInput label="🎯 جذب/توفير/نمو/بيانات" value={otherRevM} onChange={setOtherRevM} step={1} />
          </div>
        </Card>
      </div>

      <Card>
        <p className="text-sm text-darb-ink/90 leading-relaxed">
          💡 لاحظ أن <b>تكلفة النقاط</b> ({fmtNum(pointsCost5y, 1)} م﷼) تتغيّر تلقائياً مع أي تعديل
          في معدلات الكسب أو عدد المحطات في لوحة السيناريوهات — فهي مشتقة من المحرك لا رقماً ثابتاً.
          الربحية تعتمد أساساً على <b>عمولة الشركاء وزيادة مبيعات الوقود</b>، لا على الوقود وحده.
        </p>
      </Card>
    </div>
  );
}
