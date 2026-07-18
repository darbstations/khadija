import { useState } from "react";
import { useScenario } from "../context/ScenarioContext";
import { fmtSar, fmtInt, fmtPct } from "../model/engine";
import { Card, Stat, Badge } from "./ui";

type Source = "topup" | "fuel" | "merchant" | "redeem" | "none";
interface Step {
  icon: string;
  title: string;
  walletD: number;
  pointsD: number;
  spend: number;
  source: Source;
  note: string;
  earn: boolean;
}

const MERCHANT_CUST_PCT = 1; // حصة العميل عند المتجر %

export default function CustomerFlow() {
  const { inputs } = useScenario();
  const pv = inputs.pointValue;
  const fuelRate = inputs.earnWhite; // نقاط/ريال على البنزين

  const [wallet, setWallet] = useState(0);
  const [points, setPoints] = useState(0);
  const [amount, setAmount] = useState(200);
  const [pay, setPay] = useState<"wallet" | "card">("wallet");
  const [log, setLog] = useState<Step[]>([]);

  const push = (s: Step) => {
    setWallet((w) => w + s.walletD);
    setPoints((p) => Math.max(0, p + s.pointsD));
    setLog((l) => [s, ...l]);
  };
  const err = (title: string, note: string) =>
    setLog((l) => [{ icon: "⛔", title, walletD: 0, pointsD: 0, spend: 0, source: "none", note, earn: false }, ...l]);

  const topup = () =>
    push({ icon: "🔋", title: `شحن المحفظة ${fmtSar(amount)}`, walletD: amount, pointsD: 0, spend: 0, source: "topup", note: "لا كسب — إيداع رصيد", earn: false });

  const spend = (kind: "fuel" | "merchant") => {
    if (pay === "wallet" && amount > wallet) return err("رصيد المحفظة لا يكفي", "اشحني أو ادفعي بالبطاقة");
    const walletD = pay === "wallet" ? -amount : 0;
    const pts = kind === "fuel" ? amount * fuelRate : amount * (MERCHANT_CUST_PCT / 100) * pv;
    push({
      icon: kind === "fuel" ? "⛽" : "🛍️",
      title: `${kind === "fuel" ? "تعبئة بنزين" : "شراء من متجر"} ${fmtSar(amount)} (${pay === "wallet" ? "من المحفظة" : "بطاقة"})`,
      walletD, pointsD: pts, spend: amount, source: kind,
      note: `كسب عند الشراء (QR) · ${kind === "merchant" ? "التاجر يدفع 3%" : "تموّله درب"}`,
      earn: true,
    });
  };

  const redeem = (icon: string, title: string, cost: number) => {
    if (cost > points) return err(`${title} — نقاط غير كافية`, `تحتاج ${fmtInt(cost)} نقطة`);
    push({ icon, title: `استبدال: ${title}`, walletD: 0, pointsD: -cost, spend: 0, source: "redeem", note: "خصم من رصيد النقاط", earn: false });
  };

  const reset = () => { setWallet(0); setPoints(0); setLog([]); };

  // تجميعات الرحلة
  const sum = (f: (s: Step) => number) => log.reduce((a, s) => a + f(s), 0);
  const fuelSpend = sum((s) => (s.source === "fuel" ? s.spend : 0));
  const merchSpend = sum((s) => (s.source === "merchant" ? s.spend : 0));
  const fuelPts = sum((s) => (s.source === "fuel" ? s.pointsD : 0));
  const merchPts = sum((s) => (s.source === "merchant" ? s.pointsD : 0));
  const earned = fuelPts + merchPts;
  const redeemed = -sum((s) => (s.source === "redeem" ? s.pointsD : 0));
  const totalSpend = fuelSpend + merchSpend;
  const effCashback = totalSpend ? earned / pv / totalSpend : 0;

  return (
    <div className="space-y-4">
      <div>
        <h2 className="text-lg font-extrabold">🎬 محاكي رحلة العميل المتكامل</h2>
        <p className="text-xs text-darb-mut">يجمع نقاطاً من البنزين والتجار · ابدئي من الشحن أو التعبئة</p>
      </div>

      <div className="grid sm:grid-cols-3 gap-3">
        <Stat label="👛 محفظة الشحن (ريال)" value={fmtSar(wallet)} tone="accent" />
        <Stat label="💎 رصيد النقاط" value={fmtInt(points)} hint={`= ${fmtSar(points / pv)}`} tone="good" />
        <Stat label="📋 عدد الخطوات" value={fmtInt(log.length)} />
      </div>

      {/* بطاقة تجميعية: من البنزين والتجار */}
      <Card title="📊 من أين جمع نقاطه؟">
        <div className="grid sm:grid-cols-2 lg:grid-cols-4 gap-3">
          <Stat label="⛽ من البنزين" value={fmtInt(fuelPts)} hint={`إنفاق ${fmtSar(fuelSpend)} · تموّله درب`} tone="warn" />
          <Stat label="🛍️ من التجار" value={fmtInt(merchPts)} hint={`إنفاق ${fmtSar(merchSpend)} · يموّله التاجر`} tone="good" />
          <Stat label="💎 إجمالي المكتسب" value={fmtInt(earned)} hint={`= ${fmtSar(earned / pv)}`} tone="accent" />
          <Stat label="🎁 المستبدل / الرصيد" value={`${fmtInt(redeemed)} / ${fmtInt(points)}`} />
        </div>
        {totalSpend > 0 && (
          <div className="mt-3 flex items-center gap-2 text-sm">
            <span className="text-darb-mut">الكاش باك الفعلي للرحلة (بنزين + تجار):</span>
            <Badge tone="accent">{fmtPct(effCashback)}</Badge>
          </div>
        )}
      </Card>

      <Card>
        <div className="flex flex-wrap items-end gap-3 mb-3">
          <label className="block">
            <span className="text-xs font-bold text-darb-ink">المبلغ (ريال)</span>
            <input type="number" value={amount} onChange={(e) => setAmount(parseFloat(e.target.value) || 0)}
              className="block w-32 mt-1 bg-darb-yellow/10 border border-darb-yellow/40 rounded-lg px-3 py-2 font-bold text-left" dir="ltr" />
          </label>
          <div>
            <span className="text-xs font-bold text-darb-ink">طريقة الدفع للشراء</span>
            <div className="flex gap-1.5 mt-1">
              {(["wallet", "card"] as const).map((m) => (
                <button key={m} onClick={() => setPay(m)}
                  className={`text-xs font-bold px-3 py-2 rounded-lg border transition ${pay === m ? "border-darb-orange bg-darb-orange/15 text-darb-orange" : "border-darb-line text-darb-mut"}`}>
                  {m === "wallet" ? "👛 من المحفظة" : "💳 بطاقة"}
                </button>
              ))}
            </div>
          </div>
          <button onClick={reset} className="ms-auto text-xs font-bold px-3 py-2 rounded-lg border border-darb-line hover:border-darb-orange text-darb-mut hover:text-darb-orange">↺ إعادة</button>
        </div>

        <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-2">
          <FlowBtn onClick={topup} label="🔋 شحن المحفظة" />
          <FlowBtn onClick={() => spend("fuel")} label="⛽ تعبئة بنزين" />
          <FlowBtn onClick={() => spend("merchant")} label="🛍️ شراء متجر" />
          <FlowBtn onClick={() => redeem("🎁", "قهوة مجانية", 1600)} label="🎁 استبدل قهوة" />
          <FlowBtn onClick={() => redeem("⛽", "خصم بنزين 10﷼", 10 * pv)} label="⛽ استبدل بنزين" />
          <FlowBtn onClick={() => redeem("🪪", "باقة ذيبان", 500 * pv)} label="🪪 استبدل باقة" />
        </div>
      </Card>

      <Card title="🧾 سجل الرحلة">
        {log.length === 0 ? (
          <p className="text-sm text-darb-mut text-center py-6">ابدئي بالضغط على «🔋 شحن المحفظة» أو «⛽ تعبئة بنزين».</p>
        ) : (
          <div className="space-y-2">
            {log.map((s, i) => (
              <div key={i} className="stat flex items-center justify-between gap-3">
                <div className="flex items-center gap-2 min-w-0">
                  <span className="text-xl">{s.icon}</span>
                  <div className="min-w-0">
                    <div className="font-bold text-darb-ink truncate">{s.title}</div>
                    <div className="text-[11px] text-darb-mut">{s.note}</div>
                  </div>
                </div>
                <div className="text-left shrink-0">
                  {s.walletD !== 0 && <div className={`text-xs font-bold ${s.walletD > 0 ? "text-darb-good" : "text-darb-warn"}`}>{s.walletD > 0 ? "+" : ""}{fmtSar(s.walletD)}</div>}
                  {s.pointsD !== 0 ? (
                    <div className={`text-xs font-bold ${s.pointsD > 0 ? "text-darb-accent" : "text-darb-bad"}`}>{s.pointsD > 0 ? "+" : ""}{fmtInt(s.pointsD)} نقطة</div>
                  ) : (
                    s.source === "topup" && <Badge tone="warn">0 نقطة</Badge>
                  )}
                </div>
              </div>
            ))}
          </div>
        )}
      </Card>

      <Card>
        <p className="text-sm text-darb-ink/90 leading-relaxed">
          🔑 <b>القاعدة:</b> الشحن = 0 نقطة. الشراء (بنزين/متجر) = نقاط عند QR مهما كانت طريقة الدفع.
          البنزين {fmtInt(fuelRate)} نقطة/ريال (تموّله درب)، والمتجر {MERCHANT_CUST_PCT}% (يموّله التاجر) — والعميل يجمع من الاثنين في رصيد واحد.
        </p>
      </Card>
    </div>
  );
}

function FlowBtn({ onClick, label }: { onClick: () => void; label: string }) {
  return (
    <button onClick={onClick} className="text-xs font-bold px-2 py-2.5 rounded-lg border border-darb-line hover:border-darb-orange hover:bg-darb-orange/10 text-darb-ink transition">
      {label}
    </button>
  );
}
