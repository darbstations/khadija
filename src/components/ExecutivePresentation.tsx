import { useScenario } from "../context/ScenarioContext";
import { tiers, pointMath, projections, fmtInt, fmtPct, fmtNum, fmtSar } from "../model/engine";
import { DarbLogo } from "./DarbLogo";
import { Card, Stat, Badge } from "./ui";

// أرقام النموذج المعتمدة (ورقة ROI · 5 سنوات)
const MODEL = { revenue: 147.2, cost: 53.78, net: 93.42, roi: 1.74, breakeven: "السنة 1" };

const PHASES = [
  { c: "🟢", t: "إطلاق الوقود", p: "شهر 1-3", g: "50K مستخدم" },
  { c: "🔵", t: "تفاوض المستأجرين", p: "شهر 3-6", g: "3-5 عقود + Pilot" },
  { c: "🟣", t: "التوسع التدريجي", p: "شهر 6-12", g: "100K مستخدم · 5 أقسام" },
  { c: "🟠", t: "النضج الكامل", p: "شهر 12-18", g: "200K مستخدم · بطاقات مشتركة" },
];

export default function ExecutivePresentation() {
  const { inputs } = useScenario();
  const t = tiers(inputs);
  const m = pointMath(inputs);
  const p = projections(inputs);
  const totalStations =
    inputs.stationsOperating + inputs.stationsInvestment + inputs.stationsFranchise;

  return (
    <div className="space-y-5">
      {/* الواجهة */}
      <div className="rounded-2xl bg-gradient-to-bl from-darb-orange/20 to-darb-card border border-darb-line p-6">
        <div className="flex items-start justify-between flex-wrap gap-4">
          <div>
            <DarbLogo className="h-12 w-auto mb-3" />
            <h1 className="text-2xl font-extrabold">تانكي · برنامج ولاء درب</h1>
            <p className="text-darb-mut mt-1">
              عرض تنفيذي · النموذج المالي والاستراتيجية — 2026
            </p>
          </div>
          <button
            onClick={() => window.print()}
            className="no-print text-xs font-bold px-3 py-2 rounded-lg border border-darb-orange text-darb-orange hover:bg-darb-orange/15 transition"
          >
            🖨️ طباعة / PDF
          </button>
        </div>
      </div>

      {/* الأطروحة */}
      <div className="grid md:grid-cols-3 gap-3">
        <Card title="⛽ الوقود يجذب">
          <p className="text-sm text-darb-ink/90 leading-relaxed">
            هامش الوقود رفيع (5-6%)، فالكسب عليه بسيط (~0.5-1%) — وظيفته جذب العميل للمحطة وبناء العادة.
          </p>
        </Card>
        <Card title="🛍️ الشركاء يموّلون">
          <p className="text-sm text-darb-ink/90 leading-relaxed">
            الكرم الحقيقي (3-5%) يأتي من الشركاء بهوامشهم العالية (50-70%)، فيتحملون كلفة المكافآت لا درب.
          </p>
        </Card>
        <Card title="👛 المحفظة تحمي الثقة">
          <p className="text-sm text-darb-ink/90 leading-relaxed">
            النقاط واجهة فوق محفظة مغطّاة بالريال — لا أحد يموّل وعود غيره، والنظام لا يخسر.
          </p>
        </Card>
      </div>

      {/* الأرقام المفتاحية */}
      <Card title="📊 الأرقام المفتاحية">
        <div className="grid sm:grid-cols-3 lg:grid-cols-6 gap-3">
          <Stat label="قيمة النقطة" value={`${fmtInt(inputs.pointValue)} = ﷼`} tone="accent" />
          <Stat label="الكاش باك" value={`${fmtPct(t[0].cashbackPct, 1)}–${fmtPct(t[2].cashbackPct, 0)}`} />
          <Stat label="مقابل ساسكو (نقاط)" value={`${fmtInt(t[0].earnRate / m.sascoPointsPerRiyal)}–${fmtInt(t[2].earnRate / m.sascoPointsPerRiyal)}x`} tone="accent" />
          <Stat label="عدد المحطات" value={fmtInt(totalStations)} />
          <Stat label="تكلفة النقاط 5 سنوات" value={`${fmtNum(p.fiveYearTotal, 1)} م﷼`} hint="حيّة من المحرك" tone="warn" />
          <Stat label="هدية الترحيب" value={`${fmtInt(inputs.welcomeGift)} نقطة`} hint={fmtSar(inputs.welcomeGift / inputs.pointValue)} />
        </div>
      </Card>

      {/* جدوى الاستثمار */}
      <Card title="💎 جدوى الاستثمار (5 سنوات)">
        <div className="grid sm:grid-cols-2 lg:grid-cols-5 gap-3">
          <Stat label="الإيرادات الإضافية" value={`${MODEL.revenue} م﷼`} tone="accent" />
          <Stat label="إجمالي التكاليف" value={`${MODEL.cost} م﷼`} tone="warn" />
          <Stat label="صافي الربح" value={`${MODEL.net} م﷼`} tone="good" />
          <Stat label="ROI" value={fmtPct(MODEL.roi, 0)} tone="good" />
          <Stat label="نقطة التعادل" value={MODEL.breakeven} tone="good" />
        </div>
        <div className="mt-3 flex items-center gap-2">
          <Badge tone="good">⭐ مشروع ممتاز</Badge>
          <span className="text-xs text-darb-mut">العائد أساساً من عمولة الشركاء وزيادة مبيعات الوقود</span>
        </div>
      </Card>

      {/* المستويات */}
      <Card title="🎯 المستويات الثلاثة">
        <div className="grid sm:grid-cols-3 gap-3">
          {t.map((tier) => (
            <div key={tier.key} className="stat text-center">
              <div className="text-3xl">{tier.emoji}</div>
              <div className="font-extrabold mt-1">{tier.label}</div>
              <div className="text-darb-orange font-bold text-lg mt-1">{fmtPct(tier.cashbackPct, 2)}</div>
              <div className="text-[11px] text-darb-mut">{tier.earnRate} نقطة/ريال · من {fmtInt(tier.minSpend)} ﷼/سنة</div>
            </div>
          ))}
        </div>
      </Card>

      {/* خطة التوسع */}
      <Card title="🚀 خطة التوسع · 18 شهراً">
        <div className="grid sm:grid-cols-2 lg:grid-cols-4 gap-3">
          {PHASES.map((ph) => (
            <div key={ph.t} className="stat">
              <div className="text-2xl">{ph.c}</div>
              <div className="font-bold mt-1">{ph.t}</div>
              <div className="text-[11px] text-darb-mut">{ph.p}</div>
              <div className="text-xs text-darb-accent mt-1">{ph.g}</div>
            </div>
          ))}
        </div>
      </Card>

      {/* التوصية */}
      <div className="rounded-2xl bg-darb-orange/10 border border-darb-orange/40 p-5">
        <h3 className="font-extrabold text-darb-orange mb-2">✅ التوصية الاستراتيجية</h3>
        <ul className="text-sm space-y-1.5 text-darb-ink/90">
          <li>اعتماد نموذج المحفظة المموّلة: نقاط واجهة + رصيد حقيقي بالريال (ثقة + لا خسارة).</li>
          <li>كسب الوقود بسيط (~0.5-1%)، والكرم من الشركاء عبر سوق عروض تنافسي.</li>
          <li>أولوية التنفيذ: توقيع الشركاء — لأن الربحية منهم لا من الوقود.</li>
          <li>قيمة نقطة معقولة ({fmtInt(inputs.pointValue)} = ريال) بلا تضخيم يكشف ضعف القيمة.</li>
        </ul>
      </div>

      <p className="text-center text-[11px] text-darb-mut">
        درب · إدارة الشراكات · النموذج المالي تانكي · 2026
      </p>
    </div>
  );
}
