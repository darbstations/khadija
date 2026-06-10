import { useScenario } from "../context/ScenarioContext";
import { Card, NumberInput } from "./ui";
import type { ScenarioInputs } from "../model/types";

type Status = { text: string; tone: "good" | "warn" | "bad" } | undefined;

export default function ScenarioPanel() {
  const { inputs, set, reset } = useScenario();
  const i = inputs;

  // مؤشرات حيّة تطابق معادلات IF في ورقة السيناريوهات
  const ratio = (rate: number) => rate / i.pointValue;
  const stPV: Status =
    i.pointValue <= 5000
      ? { text: "🔴 سخي جداً", tone: "bad" }
      : i.pointValue <= 10000
      ? { text: "✅ موصى به", tone: "good" }
      : i.pointValue <= 15000
      ? { text: "⚠️ متحفظ", tone: "warn" }
      : { text: "🔴 بخيل", tone: "bad" };
  const stWhite: Status =
    ratio(i.earnWhite) < 0.005
      ? { text: "🔴 ضعيف", tone: "bad" }
      : ratio(i.earnWhite) <= 0.01
      ? { text: "✅ مثالي", tone: "good" }
      : ratio(i.earnWhite) <= 0.015
      ? { text: "⚠️ سخي", tone: "warn" }
      : { text: "🔴 مكلف", tone: "bad" };
  const stCost: Status =
    i.costPerRiyalHalala <= 0.75
      ? { text: "✅ مستدام", tone: "good" }
      : i.costPerRiyalHalala <= 1.25
      ? { text: "⚠️ مقبول", tone: "warn" }
      : { text: "🔴 ضغط", tone: "bad" };
  const stationSum = i.stationsOperating + i.stationsInvestment + i.stationsFranchise;
  const stStations: Status =
    stationSum >= 150 ? { text: "✅ منطقي", tone: "good" } : { text: "⚠️ تحقق", tone: "warn" };
  const stSales: Status =
    i.stationMonthlyLiters >= 50000
      ? { text: "✅ ممتاز", tone: "good" }
      : i.stationMonthlyLiters >= 30000
      ? { text: "⚠️ متوسط", tone: "warn" }
      : { text: "🔴 منخفض", tone: "bad" };

  const num = (k: keyof ScenarioInputs) => i[k] as number;

  return (
    <div className="space-y-4">
      <div className="flex items-center justify-between flex-wrap gap-2">
        <div>
          <h2 className="text-lg font-extrabold">⚙️ لوحة التحكم بالسيناريوهات</h2>
          <p className="text-xs text-darb-mut">
            غيّر القيم في الحقول الصفراء · كل الشاشات تتحدث فوراً ↻
          </p>
        </div>
        <button
          onClick={reset}
          className="text-xs font-bold px-3 py-1.5 rounded-lg border border-darb-line hover:border-darb-accent text-darb-mut hover:text-darb-accent transition"
        >
          ↺ إعادة القيم الموصى بها
        </button>
      </div>

      <Card title="1️⃣ قيمة النقطة ومعدلات الكسب" subtitle="النموذج المُضخّم">
        <div className="grid sm:grid-cols-2 lg:grid-cols-3 gap-4">
          <NumberInput
            label="قيمة النقطة (نقطة = 1 ريال)"
            value={num("pointValue")}
            onChange={(v) => set("pointValue", v)}
            step={1000}
            status={stPV}
            hint="كم نقطة تساوي ريالاً عند الاستبدال"
          />
          <NumberInput
            label="كسب الأبيض (نقطة/ريال)"
            value={num("earnWhite")}
            onChange={(v) => set("earnWhite", v)}
            step={5}
            status={stWhite}
            hint={`كاش باك ${(ratio(i.earnWhite) * 100).toFixed(2)}%`}
          />
          <NumberInput
            label="كسب الفضي (نقطة/ريال)"
            value={num("earnSilver")}
            onChange={(v) => set("earnSilver", v)}
            step={5}
            hint={`كاش باك ${(ratio(i.earnSilver) * 100).toFixed(2)}%`}
          />
          <NumberInput
            label="كسب البرتقالي (نقطة/ريال)"
            value={num("earnOrange")}
            onChange={(v) => set("earnOrange", v)}
            step={5}
            hint={`كاش باك ${(ratio(i.earnOrange) * 100).toFixed(2)}%`}
          />
          <NumberInput
            label="تكلفة الولاء على درب (هللة/ريال)"
            value={num("costPerRiyalHalala")}
            onChange={(v) => set("costPerRiyalHalala", v)}
            step={0.25}
            status={stCost}
          />
          <NumberInput
            label="سعر لتر البنزين 95 (ريال)"
            value={num("literPrice")}
            onChange={(v) => set("literPrice", v)}
            step={0.01}
            hint="يونيو 2026 · 2.33 ﷼"
          />
          <NumberInput
            label="هدية الترحيب (نقطة)"
            value={num("welcomeGift")}
            onChange={(v) => set("welcomeGift", v)}
            step={5000}
            hint={`= ${(i.welcomeGift / i.pointValue).toFixed(0)} ﷼ تكلفة فعلية`}
          />
        </div>
      </Card>

      <Card title="2️⃣ معايير الترقية بين المستويات" subtitle="إنفاق سنوي بالريال">
        <div className="grid sm:grid-cols-2 lg:grid-cols-4 gap-4">
          <NumberInput
            label="حد الترقية للفضي (﷼/سنة)"
            value={num("silverThreshold")}
            onChange={(v) => set("silverThreshold", v)}
            step={600}
            hint={`${(i.silverThreshold / 12).toFixed(0)} ﷼/شهر`}
          />
          <NumberInput
            label="حد الترقية للبرتقالي (﷼/سنة)"
            value={num("orangeThreshold")}
            onChange={(v) => set("orangeThreshold", v)}
            step={1200}
            hint={`${(i.orangeThreshold / 12).toFixed(0)} ﷼/شهر`}
          />
          <NumberInput
            label="فترة احتساب الإنفاق (شهر)"
            value={num("spendPeriodMonths")}
            onChange={(v) => set("spendPeriodMonths", v)}
          />
          <NumberInput
            label="فترة سماح بعد التنزيل (يوم)"
            value={num("graceDays")}
            onChange={(v) => set("graceDays", v)}
          />
        </div>
      </Card>

      <Card title="3️⃣ المحطات وهوامش الربح">
        <div className="grid sm:grid-cols-2 lg:grid-cols-3 gap-4">
          <NumberInput
            label="عدد محطات التشغيل"
            value={num("stationsOperating")}
            onChange={(v) => set("stationsOperating", v)}
            status={stStations}
          />
          <NumberInput
            label="عدد محطات الاستثمار"
            value={num("stationsInvestment")}
            onChange={(v) => set("stationsInvestment", v)}
          />
          <NumberInput
            label="عدد محطات الامتياز"
            value={num("stationsFranchise")}
            onChange={(v) => set("stationsFranchise", v)}
          />
          <NumberInput
            label="هامش التشغيل (هللة/لتر)"
            value={num("marginOperating")}
            onChange={(v) => set("marginOperating", v)}
            hint="هامش درب الفعلي ~13 هللة"
          />
          <NumberInput
            label="هامش الاستثمار (هللة/لتر)"
            value={num("marginInvestment")}
            onChange={(v) => set("marginInvestment", v)}
          />
          <NumberInput
            label="هامش الامتياز (هللة/لتر)"
            value={num("marginFranchise")}
            onChange={(v) => set("marginFranchise", v)}
            step={0.5}
          />
          <NumberInput
            label="مساهمة درب في الامتياز (%)"
            value={num("darbShareFranchise")}
            onChange={(v) => set("darbShareFranchise", v)}
            step={0.05}
            hint={`مساهمة صاحب الامتياز = ${((1 - i.darbShareFranchise) * 100).toFixed(0)}%`}
          />
          <NumberInput
            label="متوسط مبيعات المحطة (لتر/شهر)"
            value={num("stationMonthlyLiters")}
            onChange={(v) => set("stationMonthlyLiters", v)}
            step={5000}
            status={stSales}
          />
        </div>
      </Card>

      <Card title="4️⃣ الاقتصاد الموسّع — الديزل والشركاء" subtitle="مربوط فعلياً بالحسابات (شاشة الاقتصاد الموسّع)">
        <div className="grid sm:grid-cols-2 lg:grid-cols-3 gap-4">
          <NumberInput
            label="سعر لتر الديزل (ريال)"
            value={num("dieselPrice")}
            onChange={(v) => set("dieselPrice", v)}
            step={0.01}
            hint="يناير 2026 · 1.79 ﷼"
          />
          <NumberInput
            label="هامش الديزل (هللة/لتر)"
            value={num("dieselMarginHalala")}
            onChange={(v) => set("dieselMarginHalala", v)}
            step={0.5}
          />
          <NumberInput
            label="كسب الديزل (نقطة/ريال)"
            value={num("earnDiesel")}
            onChange={(v) => set("earnDiesel", v)}
            step={5}
            hint={`كاش باك ${(ratio(i.earnDiesel) * 100).toFixed(2)}%`}
          />
          <NumberInput
            label="كسب الشركاء (نقطة/ريال)"
            value={num("earnPartner")}
            onChange={(v) => set("earnPartner", v)}
            step={50}
            hint={`قيمة ${(ratio(i.earnPartner) * 100).toFixed(2)}% للعميل`}
          />
          <NumberInput
            label="مساهمة الشريك لدرب (هللة/ريال)"
            value={num("partnerContributionHalala")}
            onChange={(v) => set("partnerContributionHalala", v)}
            step={0.5}
            hint="ما يدفعه الشريك مقابل عملاء درب"
          />
        </div>
      </Card>
    </div>
  );
}
