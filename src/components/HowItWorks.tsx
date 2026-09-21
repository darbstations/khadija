import { useScenario } from "../context/ScenarioContext";
import { fmtInt } from "../model/engine";
import { Card } from "./ui";

const COMPARE = [
  ["الغرض", "الاستبدال بمكافآت", "تحديد المستوى"],
  ["متى يزيد؟", "مع كل تعبئة", "مع كل تعبئة (إنفاق)"],
  ["متى ينقص؟", "عند استبدال نقاط", "لا ينقص داخل السنة"],
  ["يتجدد؟", "لا · رصيد متراكم", "نعم · كل 12 شهر"],
  ["يؤثر على المستوى؟", "❌ لا", "✅ نعم"],
];

const PHASES = [
  {
    c: "🟢",
    period: "شهر 1-3",
    title: "إطلاق الوقود",
    goal: "قاعدة 50K مستخدم بالوقود فقط",
    cost: "2.3 م﷼/سنة",
  },
  {
    c: "🔵",
    period: "شهر 3-6",
    title: "تفاوض المستأجرين",
    goal: "توقيع 3-5 عقود + Pilot كافيه",
    cost: "+ تكامل",
  },
  {
    c: "🟣",
    period: "شهر 6-12",
    title: "التوسع التدريجي",
    goal: "5+ أقسام + 100K مستخدم",
    cost: "3.5 م﷼/سنة",
  },
  {
    c: "🟠",
    period: "شهر 12-18",
    title: "النضج الكامل",
    goal: "200K مستخدم + بطاقات مشتركة",
    cost: "5+ م﷼/سنة",
  },
];

const TENANT_MODELS = [
  ["🟢 المستأجر 100%", "2 هللة/﷼", "0", "هامش عالي >60% · مغاسل/كافيهات", "⭐⭐⭐⭐⭐"],
  ["🔵 70/30 (المستأجر)", "1.5", "0.5", "هامش متوسط 40-60% · مطاعم", "⭐⭐⭐⭐"],
  ["🟡 50/50", "1", "1", "سلاسل عالمية · شركاء كبار", "⭐⭐⭐"],
  ["🟠 30/70 (درب)", "0.5", "1.5", "هامش منخفض <25% · سوبرماركت", "⭐⭐"],
];

export default function HowItWorks() {
  const { inputs } = useScenario();
  return (
    <div className="space-y-4">
      <div>
        <h2 className="text-lg font-extrabold">📚 كيف يعمل النظام + خطة التوسع</h2>
        <p className="text-xs text-darb-mut">
          النقاط للاستبدال · المستوى للحالة · شيئان منفصلان تماماً
        </p>
      </div>

      <Card title="📊 الفرق بين العدّاديْن">
        <div className="overflow-x-auto">
          <table className="w-full border-collapse">
            <thead>
              <tr>
                <th className="th">المعيار</th>
                <th className="th">💎 رصيد النقاط</th>
                <th className="th">🎯 عدّاد المستوى</th>
              </tr>
            </thead>
            <tbody>
              {COMPARE.map((r) => (
                <tr key={r[0]}>
                  <td className="td font-bold">{r[0]}</td>
                  <td className="td">{r[1]}</td>
                  <td className="td">{r[2]}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Card>

      <Card title="🚀 خطة التوسع · 18 شهراً">
        <div className="grid sm:grid-cols-2 lg:grid-cols-4 gap-3">
          {PHASES.map((ph) => (
            <div key={ph.title} className="stat">
              <div className="text-2xl">{ph.c}</div>
              <div className="font-bold mt-1">{ph.title}</div>
              <div className="text-[11px] text-darb-mut">{ph.period}</div>
              <div className="text-xs mt-2 text-darb-ink/90">{ph.goal}</div>
              <div className="text-[11px] text-darb-accent mt-1">{ph.cost}</div>
            </div>
          ))}
        </div>
      </Card>

      <Card title="🤝 نماذج التفاوض مع المستأجرين">
        <div className="overflow-x-auto">
          <table className="w-full border-collapse">
            <thead>
              <tr>
                <th className="th">النموذج</th>
                <th className="th">مساهمة المستأجر</th>
                <th className="th">مساهمة درب</th>
                <th className="th">مناسب لمن؟</th>
                <th className="th">التقييم</th>
              </tr>
            </thead>
            <tbody>
              {TENANT_MODELS.map((r) => (
                <tr key={r[0]}>
                  <td className="td font-bold">{r[0]}</td>
                  <td className="td">{r[1]}</td>
                  <td className="td">{r[2]}</td>
                  <td className="td text-xs">{r[3]}</td>
                  <td className="td">{r[4]}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <p className="text-xs text-darb-mut mt-3">
          💬 رسالة التفاوض: «عميل درب يدخل محطتك مجاناً من التطبيق. تموّل مكافأة بسيطة من هامشك
          مقابل عميل متكرر + بيانات + تسويق داخل التطبيق. النقطة قيمتها ثابتة: {fmtInt(inputs.pointValue)} نقطة = ريال.»
        </p>
      </Card>

      <Card title="🌟 الرسائل المفتاحية للعميل">
        <ul className="text-sm space-y-1.5 text-darb-ink/90">
          <li>💯 استبدل نقاطك متى شئت · مستواك لن يتأثر أبداً.</li>
          <li>📈 كلما أنفقت أكثر · ارتقيت أعلى · لا علاقة لذلك بنقاطك.</li>
          <li>🎁 النقاط متعتك · المستوى إنجازك · شيئان منفصلان.</li>
          <li>🔒 مستواك محفوظ 12 شهراً بعد الترقية · حتى لو قلّ إنفاقك.</li>
        </ul>
      </Card>
    </div>
  );
}
