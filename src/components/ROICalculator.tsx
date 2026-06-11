import { useState } from "react";
import { useScenario } from "../context/ScenarioContext";
import { projections, fmtNum, fmtPct, fmtSar } from "../model/engine";
import { useStations, stationsLoyaltyCost } from "../model/stations";
import { Card, Stat, NumberInput, Badge } from "./ui";

/** حاسبة ROI تفاعلية — التكاليف تشتق من التوقعات، والإيرادات من افتراضات قابلة للتعديل */
export default function ROICalculator() {
  const { inputs } = useScenario();
  const p = projections(inputs);
  const { stations } = useStations();
  const stationsCost = stationsLoyaltyCost(stations, inputs.pointValue);

  // مصدر تكلفة النقاط: من التوقعات أو من محطاتي الفعلية
  const [costSource, setCostSource] = useState<"projection" | "stations">("projection");

  // افتراضات قابلة للتعديل (مليون ريال / نِسَب)
  const [appDev, setAppDev] = useState(1.45); // تطوير التطبيق (5 سنوات)
  const [team, setTeam] = useState(19.78); // فريق الولاء
  const [marketing, setMarketing] = useState(5.7); // تسويق
  const [partnerRevM, setPartnerRevM] = useState(13.7); // عمولة الشركاء (5 سنوات)
  const [fuelUpliftRevM, setFuelUpliftRevM] = useState(80); // زيادة مبيعات الوقود
  const [otherRevM, setOtherRevM] = useState(53.5); // جذب عملاء + توفير + نمو + بيانات

  const stations5yM = (stationsCost.annualCost * 5) / 1_000_000; // محطاتي الفعلية × 5 سنوات
  const pointsCost5y = costSource === "stations" ? stations5yM : p.fiveYearTotal;
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

      <div className="flex items-center gap-3 flex-wrap">
        <span className="text-sm text-darb-mut">التقييم النهائي:</span>
        <Badge tone={verdict.tone}>{verdict.t}</Badge>
      </div>

      <Card title="🎁 مصدر تكلفة النقاط (5 سنوات)">
        <div className="flex flex-wrap gap-2">
          <button
            onClick={() => setCostSource("projection")}
            className={`text-xs font-bold px-3 py-2 rounded-lg border transition ${
              costSource === "projection"
                ? "border-darb-accent bg-darb-accent/15 text-darb-accent"
                : "border-darb-line text-darb-mut hover:text-darb-ink"
            }`}
          >
            📈 من التوقعات (نمو المحطات) · {fmtNum(p.fiveYearTotal, 1)} م﷼
          </button>
          <button
            onClick={() => setCostSource("stations")}
            className={`text-xs font-bold px-3 py-2 rounded-lg border transition ${
              costSource === "stations"
                ? "border-darb-accent bg-darb-accent/15 text-darb-accent"
                : "border-darb-line text-darb-mut hover:text-darb-ink"
            }`}
          >
            ⛽ من «محطاتي» الفعلية × 5 · {fmtNum(stations5yM, 1)} م﷼
          </button>
        </div>
        <p className="text-xs text-darb-mut mt-2">
          محطاتي: {stations.length} محطة · تكلفة ولاء سنوية {fmtSar(stationsCost.annualCost)} · متوسط
          كاش باك {fmtPct(stationsCost.blendedCashback)}. أضيفي/عدّلي محطاتك في شاشة «محطاتي» وتنعكس هنا.
        </p>
      </Card>

      <div className="grid lg:grid-cols-2 gap-4">
        <Card title="💸 التكاليف (5 سنوات · مليون ريال)">
          <div className="space-y-3">
            <div className="stat flex items-center justify-between">
              <span className="text-sm">
                🎁 كلفة النقاط ({costSource === "stations" ? "من محطاتي" : "من التوقعات"})
              </span>
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
