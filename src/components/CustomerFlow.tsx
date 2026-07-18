import { useState } from "react";
import { useScenario } from "../context/ScenarioContext";
import { fmtSar, fmtInt } from "../model/engine";
import { Card, Stat, Badge } from "./ui";

interface Step {
  icon: string;
  title: string;
  walletD: number; // تغيّر محفظة الشحن (ريال)
  pointsD: number; // تغيّر النقاط
  note: string;
  earn: boolean;
}

const MERCHANT_CUST_PCT = 1; // حصة العميل عند المتجر % (من نموذج 3%)

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

  const topup = () =>
    push({ icon: "🔋", title: `شحن المحفظة ${fmtSar(amount)}`, walletD: amount, pointsD: 0, note: "لا كسب — إيداع رصيد", earn: false });

  const spend = (kind: "fuel" | "merchant") => {
    if (pay === "wallet" && amount > wallet) {
      setLog((l) => [{ icon: "⛔", title: "رصيد المحفظة لا يكفي", walletD: 0, pointsD: 0, note: "اشحني أو ادفعي بالبطاقة", earn: false }, ...l]);
      return;
    }
    const walletD = pay === "wallet" ? -amount : 0;
    const pts = kind === "fuel" ? amount * fuelRate : amount * (MERCHANT_CUST_PCT / 100) * pv;
    push({
      icon: kind === "fuel" ? "⛽" : "🛍️",
      title: `${kind === "fuel" ? "تعبئة بنزين" : "شراء من متجر"} ${fmtSar(amount)} (${pay === "wallet" ? "من المحفظة" : "بطاقة"})`,
      walletD,
      pointsD: pts,
      note: `كسب عند الشراء (QR) · ${kind === "merchant" ? "التاجر يدفع 3%" : "تموّله درب"}`,
      earn: true,
    });
  };

  const redeem = (icon: string, title: string, cost: number) => {
    if (cost > points) {
      setLog((l) => [{ icon: "⛔", title: `${title} — نقاط غير كافية`, walletD: 0, pointsD: 0, note: `تحتاج ${fmtInt(cost)} نقطة`, earn: false }, ...l]);
      return;
    }
    push({ icon, title: `استبدال: ${title}`, walletD: 0, pointsD: -cost, note: "خصم من رصيد النقاط", earn: false });
  };

  const reset = () => { setWallet(0); setPoints(0); setLog([]); };

  return (
    <div className="space-y-4">
      <div>
        <h2 className="text-lg font-extrabold">🎬 محاكي رحلة العميل المتكامل</h2>
        <p className="text-xs text-darb-mut">ابدئي من شحن المحفظة أو تعبئة البنزين · شوفي المحفظتين تتحدثان بقاعدة الكسب</p>
      </div>

      {/* المحفظتان */}
      <div className="grid sm:grid-cols-3 gap-3">
        <Stat label="👛 محفظة الشحن (ريال)" value={fmtSar(wallet)} tone="accent" />
        <Stat label="💎 رصيد النقاط" value={fmtInt(points)} hint={`= ${fmtSar(points / pv)}`} tone="good" />
        <Stat label="📋 عدد الخطوات" value={fmtInt(log.length)} />
      </div>

      {/* لوحة التحكم */}
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

      {/* الخط الزمني */}
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
                    !s.earn && s.walletD >= 0 && <Badge tone="warn">0 نقطة</Badge>
                  )}
                </div>
              </div>
            ))}
          </div>
        )}
      </Card>

      <Card>
        <p className="text-sm text-darb-ink/90 leading-relaxed">
          🔑 <b>القاعدة المطبّقة:</b> الشحن = <b>0 نقطة</b> (إيداع). الشراء (بنزين/متجر) = <b>نقاط عند مسح QR</b>
          مهما كانت طريقة الدفع. البنزين يكسب {fmtInt(fuelRate)} نقطة/ريال (تموّله درب)، والمتجر {MERCHANT_CUST_PCT}% (يموّله التاجر).
          فلا ازدواج ولا خسارة.
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
