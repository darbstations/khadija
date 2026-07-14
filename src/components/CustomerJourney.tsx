import { useState } from "react";
import { useScenario } from "../context/ScenarioContext";
import { simulateCustomer, fmtInt, fmtSar, fmtPct } from "../model/engine";
import { REDEEM_MIN_POINTS, TIER_PERKS } from "../model/defaults";
import { Card, Stat, Badge } from "./ui";

export default function CustomerJourney() {
  const { inputs } = useScenario();
  const [fuel, setFuel] = useState(600); // بنزين/شهر
  const [merchant, setMerchant] = useState(200); // متاجر/شهر

  const sim = simulateCustomer(inputs, fuel);
  const fuelPts = fuel * inputs.earnWhite; // كل ريال = نقطة
  const merchantPts = merchant * inputs.earnPartner; // الشركاء (ممول من الشريك)
  const monthlyPts = fuelPts + merchantPts;
  const monthlyValue = monthlyPts / inputs.pointValue;
  const annualPts = monthlyPts * 12;

  // متى يقدر يستبدل: فور تجاوز الحد الأدنى
  const cheapestOffer = 600; // أرخص عرض (نقطة)
  const daysToFirst = monthlyPts > 0 ? Math.max(1, Math.round((cheapestOffer / monthlyPts) * 30)) : 0;
  const tierPerks = TIER_PERKS.find((t) => t.label === sim.tier.label) ?? TIER_PERKS[0];

  // جدول تراكمي مبسّط
  const months = ["شهر 1", "شهر 3", "شهر 6", "شهر 12"];
  const mult = [1, 3, 6, 12];

  return (
    <div className="space-y-4">
      <div>
        <h2 className="text-lg font-extrabold">🧭 رحلة العميل · من الجمع إلى الاستبدال</h2>
        <p className="text-xs text-darb-mut">كيف يجمع من البنزين والمتاجر · ومتى يقدر يستبدل</p>
      </div>

      <Card>
        <div className="grid sm:grid-cols-2 gap-5">
          <label className="block">
            <div className="flex justify-between text-xs mb-1"><span className="font-bold">⛽ بنزين/شهر</span><b>{fmtSar(fuel)}</b></div>
            <input type="range" min={100} max={5000} step={50} value={fuel} onChange={(e) => setFuel(+e.target.value)} className="w-full accent-darb-orange" />
          </label>
          <label className="block">
            <div className="flex justify-between text-xs mb-1"><span className="font-bold">🛍️ متاجر/شهر</span><b>{fmtSar(merchant)}</b></div>
            <input type="range" min={0} max={3000} step={50} value={merchant} onChange={(e) => setMerchant(+e.target.value)} className="w-full accent-darb-accent" />
          </label>
        </div>
      </Card>

      {/* مصادر الجمع */}
      <div className="grid sm:grid-cols-2 lg:grid-cols-4 gap-3">
        <Stat label="⛽ نقاط البنزين/شهر" value={fmtInt(fuelPts)} hint={`${fmtPct(inputs.earnWhite / inputs.pointValue)} · تموّلها درب`} tone="warn" />
        <Stat label="🛍️ نقاط المتاجر/شهر" value={fmtInt(merchantPts)} hint={`${fmtPct(inputs.earnPartner / inputs.pointValue)} · يموّلها الشريك`} tone="good" />
        <Stat label="💎 إجمالي/شهر" value={fmtInt(monthlyPts)} hint={fmtSar(monthlyValue)} tone="accent" />
        <Stat label="📅 إجمالي/سنة" value={fmtInt(annualPts)} hint={fmtSar(annualPts / inputs.pointValue)} tone="accent" />
      </div>

      {/* متى يستبدل */}
      <Card title="⏱️ متى يقدر يستبدل؟">
        <ul className="text-sm space-y-1.5 text-darb-ink/90">
          <li>✅ <b>لا فترة انتظار</b> — يستبدل فور تجاوز رصيده الحد الأدنى ({REDEEM_MIN_POINTS} نقطة = {fmtSar(REDEEM_MIN_POINTS / inputs.pointValue)}).</li>
          <li>☕ يكفي لأرخص عرض (~{fmtInt(cheapestOffer)} نقطة) خلال <b>~{daysToFirst} يوم</b> بمعدل إنفاقه الحالي.</li>
          <li>💎 النقاط <b>متراكمة</b> وصالحة حتى 24 شهراً — يقدر يجمّع لعرض أكبر أو يستبدل فوراً.</li>
          <li>🎯 الاستبدال لا يؤثر على مستواه (النقاط للاستبدال · الإنفاق للمستوى).</li>
        </ul>
      </Card>

      {/* الخط الزمني */}
      <Card title="📈 تراكم النقاط عبر الوقت">
        <div className="overflow-x-auto">
          <table className="w-full border-collapse">
            <thead>
              <tr><th className="th">الفترة</th><th className="th">النقاط المتراكمة</th><th className="th">قيمتها</th><th className="th">ماذا يقدر يستبدل؟</th></tr>
            </thead>
            <tbody>
              {months.map((m, i) => {
                const pts = monthlyPts * mult[i];
                const val = pts / inputs.pointValue;
                const what = pts >= 4000 ? "خصم سوبرماركت 25 ﷼ + قهوة" : pts >= 1500 ? "وجبة بخصم / غسلة كاملة" : pts >= cheapestOffer ? "قهوة أو ساندويتش" : "يقترب من أول عرض";
                return (
                  <tr key={m}>
                    <td className="td font-bold">{m}</td>
                    <td className="td">{fmtInt(pts)}</td>
                    <td className="td text-darb-accent">{fmtSar(val)}</td>
                    <td className="td text-xs">{what}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </Card>

      {/* المستوى والمميزات */}
      <Card title="🎯 مستواه ومميزاته">
        <div className="flex items-center gap-2 mb-3">
          <span>مستوى العميل بناءً على إنفاق البنزين السنوي ({fmtSar(sim.annualSpend)}):</span>
          <Badge tone="accent">{sim.tier.emoji} {sim.tier.label}</Badge>
        </div>
        <div className="grid sm:grid-cols-3 gap-3">
          {TIER_PERKS.filter((t) => t.key !== "business").map((t) => (
            <div key={t.key} className={`stat ${t.label === sim.tier.label ? "border-darb-orange/50" : ""}`}>
              <div className="font-bold mb-1">{t.emoji} {t.label}</div>
              <ul className="text-xs space-y-1 text-darb-ink/90 list-disc pr-4">
                {t.perks.map((p, i) => <li key={i}>{p}</li>)}
              </ul>
            </div>
          ))}
        </div>
        <p className="text-[11px] text-darb-mut mt-2">مستواه الحالي: <b>{tierPerks.label}</b> — يرتقي بزيادة إنفاق البنزين السنوي.</p>

        {/* مسار منفصل: الأعمال/الأساطيل (B2B) */}
        {TIER_PERKS.filter((t) => t.key === "business").map((t) => (
          <div key={t.key} className="mt-3 rounded-xl border border-darb-accent/40 bg-darb-accent/5 p-3">
            <div className="font-bold mb-1">{t.emoji} {t.label} · مسار منفصل للشركات</div>
            <ul className="text-xs grid sm:grid-cols-2 gap-x-4 gap-y-1 text-darb-ink/90 list-disc pr-4">
              {t.perks.map((p, i) => <li key={i}>{p}</li>)}
            </ul>
            <p className="text-[10px] text-darb-mut mt-1">هنا فقط يوجد مدير حساب فعلي — لأنه عدد محدود من الشركات يبرّر التكلفة.</p>
          </div>
        ))}
      </Card>
    </div>
  );
}
