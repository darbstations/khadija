import { useScenario } from "../context/ScenarioContext";
import { fmtInt } from "../model/engine";
import { Card, Badge } from "./ui";

export default function EarnRules() {
  const { inputs } = useScenario();
  const fuelRate = inputs.earnWhite; // نقاط/ريال على البنزين

  const flow = [
    { emoji: "🔋", title: "شحن محفظة البنزين", earn: true, note: `يكسب ${fmtInt(fuelRate)} نقطة/ريال — لأنها بنزين فقط` },
    { emoji: "⛽", title: "بنزين خارج المحفظة + QR", earn: true, note: "يكسب عند الضخ (بطاقة/كاش)" },
    { emoji: "🛍️", title: "شراء من متجر + QR", earn: true, note: "يكسب بمعدل المتجر (يموّله التاجر)" },
  ];

  const table: [string, boolean, string][] = [
    ["🔋 شحن محفظة البنزين", true, "يكسب بمعدل البنزين (المحفظة بنزين فقط)"],
    ["⛽ بنزين من المحفظة", false, "كُسب عند الشحن — لا تكرار"],
    ["⛽ بنزين بالبطاقة/كاش + QR", true, "استهلاك فعلي عند الضخ"],
    ["🛍️ شراء عادي من متجر", true, "بيع فعلي · التاجر يدفع نسبته"],
    ["🪪 شراء باقة", false, "منتج مدفوع مسبقاً بخصم — الخصم هو الفائدة"],
    ["⛽ بنزين ضمن الباقة", false, "مخصوم مسبقاً — لا خصم + نقاط"],
    ["☕ قسيمة قهوة/غسلة من الباقة", false, "مدفوعة مسبقاً · درب تسوّي مع التاجر (لا 3%)"],
    ["💎 استبدال نقاط (عرض/بنزين/باقة)", false, "خصم من الرصيد — ليس كسباً"],
  ];

  return (
    <div className="space-y-4">
      <div>
        <h2 className="text-lg font-extrabold">📐 قواعد كسب النقاط</h2>
        <p className="text-xs text-darb-mut">كل ريال يكسب مرة واحدة فقط — لمنع الازدواج والخسارة</p>
      </div>

      <div className="rounded-2xl bg-gradient-to-bl from-darb-orange/20 to-darb-card border border-darb-orange/40 p-5 text-center">
        <div className="text-2xl font-extrabold text-darb-orange">كل ريال يكسب نقاطاً مرة واحدة</div>
        <p className="text-sm text-darb-mut mt-1">
          عند <b>شحن محفظة البنزين</b> أو عند <b>الشراء (QR)</b> — لا يجتمعان على نفس الريال
        </p>
      </div>

      <div className="grid sm:grid-cols-3 gap-3">
        {flow.map((f) => (
          <div key={f.title} className="card text-center border border-darb-good/40">
            <div className="text-4xl">{f.emoji}</div>
            <div className="font-bold mt-2">{f.title}</div>
            <div className="my-2"><Badge tone="good">✅ يكسب نقاط</Badge></div>
            <div className="text-xs text-darb-mut">{f.note}</div>
          </div>
        ))}
      </div>

      <Card title="✅ يكسب · ❌ ما يكسب">
        <div className="overflow-x-auto">
          <table className="w-full border-collapse">
            <thead>
              <tr><th className="th">الحدث</th><th className="th">نقاط؟</th><th className="th">السبب</th></tr>
            </thead>
            <tbody>
              {table.map((r) => (
                <tr key={r[0]}>
                  <td className="td font-bold">{r[0]}</td>
                  <td className="td"><Badge tone={r[1] ? "good" : "bad"}>{r[1] ? "✅" : "❌"}</Badge></td>
                  <td className="td text-xs">{r[2]}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Card>

      <Card title="🔍 منع الازدواج — مثال">
        <div className="space-y-2 text-sm">
          <div className="stat flex items-center justify-between">
            <span>1. شحن محفظة البنزين 600 ﷼</span>
            <Badge tone="good">+{fmtInt(600 * fuelRate)} نقطة</Badge>
          </div>
          <div className="stat flex items-center justify-between">
            <span>2. تعبئة بنزين 600 ﷼ من المحفظة</span>
            <Badge tone="bad">+0 (كُسب عند الشحن)</Badge>
          </div>
          <div className="stat flex items-center justify-between font-bold">
            <span>الإجمالي على 600 ﷼</span>
            <span className="text-darb-good">{fmtInt(600 * fuelRate)} نقطة (مرة وحدة ✅)</span>
          </div>
        </div>
      </Card>

      <Card title="🧩 كيف لا تتلخبط عمليات الباقات؟">
        <p className="text-sm text-darb-ink/90 leading-relaxed">
          النظام يميّز <b>نوع القسيمة الممسوحة</b> عند التاجر:
        </p>
        <ul className="text-sm space-y-1.5 text-darb-ink/90 list-disc pr-5 mt-2">
          <li><b>شراء عادي</b> (يدفع بفلوسه) → ✅ نقاط + التاجر يدفع نسبته.</li>
          <li><b>قسيمة باقة</b> (قهوة/غسلة مدفوعة مسبقاً) → ❌ 0 نقطة · لا 3% · درب تسوّي مع التاجر بسعر الباقة.</li>
        </ul>
        <p className="text-xs text-darb-mut mt-2">القسيمة رمز مختلف عن عملية الشراء — فلا لبس ولا ازدواج.</p>
      </Card>

      <Card>
        <p className="text-sm text-darb-ink/90 leading-relaxed">
          🔑 <b>الخلاصة:</b> الكسب مرة واحدة — عند شحن محفظة البنزين أو الشراء العادي. <b>الباقة ومحتوياتها
          مدفوعة مسبقاً → صفر نقاط</b>. المحفظة <b>بنزين فقط وبلا سحب نقدي</b> (يمنع التلاعب). فلا ازدواج ولا خسارة.
        </p>
      </Card>
    </div>
  );
}
