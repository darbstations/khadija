import { useScenario } from "../context/ScenarioContext";
import { fmtInt, fmtPct } from "../model/engine";
import { Card, Badge } from "./ui";

export default function EarnRules() {
  const { inputs } = useScenario();
  const fuelPct = inputs.earnWhite / inputs.pointValue;
  const partnerPct = inputs.earnPartner / inputs.pointValue;

  const flow = [
    { emoji: "🔋", title: "شحن المحفظة", earn: false, note: "مجرد إيداع رصيد — ليس إنفاقاً" },
    { emoji: "⛽", title: "تعبئة بنزين + مسح QR", earn: true, note: `يكسب ${fmtPct(fuelPct)} (تموّله درب)` },
    { emoji: "🛍️", title: "شراء من متجر + مسح QR", earn: true, note: `يكسب ${fmtPct(partnerPct)} (يموّله الشريك)` },
  ];

  return (
    <div className="space-y-4">
      <div>
        <h2 className="text-lg font-extrabold">📐 قواعد كسب النقاط</h2>
        <p className="text-xs text-darb-mut">حدث كسب واحد فقط — لمنع الازدواج والخسارة</p>
      </div>

      {/* القاعدة الأساسية */}
      <div className="rounded-2xl bg-gradient-to-bl from-darb-orange/20 to-darb-card border border-darb-orange/40 p-5 text-center">
        <div className="text-2xl font-extrabold text-darb-orange">امسح QR لحظة الشراء = تكسب</div>
        <p className="text-sm text-darb-mut mt-1">
          الكسب مربوط بمسح الرمز عند البيع · بمعدل الفئة · مهما كانت طريقة الدفع (محفظة / بطاقة / كاش)
        </p>
      </div>

      {/* التدفّق */}
      <div className="grid sm:grid-cols-3 gap-3">
        {flow.map((f) => (
          <div key={f.title} className={`card text-center border ${f.earn ? "border-darb-good/40" : "border-darb-bad/40"}`}>
            <div className="text-4xl">{f.emoji}</div>
            <div className="font-bold mt-2">{f.title}</div>
            <div className="my-2">
              {f.earn ? <Badge tone="good">✅ يكسب نقاط</Badge> : <Badge tone="bad">❌ صفر نقاط</Badge>}
            </div>
            <div className="text-xs text-darb-mut">{f.note}</div>
          </div>
        ))}
      </div>

      {/* جدول يكسب / ما يكسب */}
      <Card title="✅ يكسب · ❌ ما يكسب">
        <div className="overflow-x-auto">
          <table className="w-full border-collapse">
            <thead>
              <tr><th className="th">الحدث</th><th className="th">نقاط؟</th><th className="th">السبب</th></tr>
            </thead>
            <tbody>
              <tr><td className="td font-bold">شحن المحفظة</td><td className="td"><Badge tone="bad">❌</Badge></td><td className="td text-xs">إيداع رصيد لا إنفاق</td></tr>
              <tr><td className="td font-bold">شراء بنزين (أي دفع) + QR</td><td className="td"><Badge tone="good">✅</Badge></td><td className="td text-xs">بيع فعلي · معدل البنزين</td></tr>
              <tr><td className="td font-bold">شراء متجر (أي دفع) + QR</td><td className="td"><Badge tone="good">✅</Badge></td><td className="td text-xs">بيع فعلي · معدل الشريك</td></tr>
              <tr><td className="td font-bold">الدفع من المحفظة لشراء</td><td className="td"><Badge tone="good">✅ مرة وحدة</Badge></td><td className="td text-xs">يكسب عند الشراء (لا عند الشحن)</td></tr>
              <tr><td className="td font-bold">بنزين ضمن باقة مدفوعة مسبقاً</td><td className="td"><Badge tone="bad">❌</Badge></td><td className="td text-xs">مدفوع بخصم مسبقاً — لا ازدواج خصم+نقاط</td></tr>
            </tbody>
          </table>
        </div>
      </Card>

      {/* مثال منع الازدواج */}
      <Card title="🔍 مثال · منع الازدواج">
        <div className="space-y-2 text-sm">
          <div className="stat flex items-center justify-between">
            <span>1. شحن المحفظة 600 ﷼</span>
            <Badge tone="bad">+0 نقطة</Badge>
          </div>
          <div className="stat flex items-center justify-between">
            <span>2. تعبئة بنزين 600 ﷼ (دفع من المحفظة) + مسح QR</span>
            <Badge tone="good">+{fmtInt(600 * inputs.earnWhite)} نقطة</Badge>
          </div>
          <div className="stat flex items-center justify-between font-bold">
            <span>الإجمالي على 600 ﷼</span>
            <span className="text-darb-good">{fmtInt(600 * inputs.earnWhite)} نقطة (مرة وحدة ✅)</span>
          </div>
          <p className="text-xs text-darb-bad mt-1">
            ❌ لو كسبت عند الشحن + الشراء = {fmtInt(600 * inputs.earnWhite * 2)} نقطة على 600 ﷼ = ضعف التكلفة = خسارة.
          </p>
        </div>
      </Card>

      <Card>
        <p className="text-sm text-darb-ink/90 leading-relaxed">
          💡 الـ QR هو <b>محرّك الكسب</b> عند البيع، والدفع منفصل تماماً. كذا: لا ازدواج · المعدل الصحيح
          لكل فئة · الشريك يتحاسب صح · واستخدام المحفظة ما يُعاقَب. (نفس آلية ساسكو/أدنوك/إمارات.)
        </p>
      </Card>
    </div>
  );
}
