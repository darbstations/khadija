import { useState } from "react";
import { useScenario } from "../context/ScenarioContext";
import { fmtSar, fmtInt } from "../model/engine";
import { BUNDLES } from "../model/defaults";
import { Card, Badge } from "./ui";

export default function Bundles() {
  const { inputs } = useScenario();
  const pv = inputs.pointValue; // نقطة = ريال
  const [wallet, setWallet] = useState(1500); // محفظة الشحن (ريال)
  const [points, setPoints] = useState(150000); // رصيد النقاط
  const [bought, setBought] = useState<string[]>([]);
  const [last, setLast] = useState<string | null>(null);

  const buyCash = (b: (typeof BUNDLES)[number]) => {
    if (b.price > wallet || bought.includes(b.id)) return;
    setWallet((w) => w - b.price);
    setBought((p) => [...p, b.id]);
    setLast(`✅ اشتريت «${b.name}» بـ ${fmtSar(b.price)} من محفظة الشحن`);
  };
  const redeemPoints = (b: (typeof BUNDLES)[number]) => {
    const cost = b.price * pv;
    if (cost > points || bought.includes(b.id)) return;
    setPoints((p) => p - cost);
    setBought((p) => [...p, b.id]);
    setLast(`✅ استبدلت «${b.name}» بـ ${fmtInt(cost)} نقطة`);
  };

  return (
    <div className="space-y-4">
      <div>
        <h2 className="text-lg font-extrabold">🪪 باقات درب · شراء أو استبدال بالنقاط</h2>
        <p className="text-xs text-darb-mut">
          العميل يستبدل نقاطه بـ 3 قنوات: 🛍️ عروض التجار · ⛽ بنزين · 🪪 باقات درب
        </p>
      </div>

      {/* الرصيدان */}
      <div className="grid sm:grid-cols-2 gap-3">
        <div className="rounded-2xl bg-gradient-to-bl from-darb-orange/25 to-darb-card border border-darb-line p-4 flex items-center justify-between">
          <div>
            <div className="text-xs text-darb-mut">👛 محفظة الشحن (ريال)</div>
            <div className="text-2xl font-extrabold text-darb-orange">{fmtSar(wallet)}</div>
          </div>
          <button onClick={() => setWallet((w) => w + 500)} className="text-xs font-bold px-3 py-2 rounded-lg border border-darb-line hover:border-darb-good text-darb-mut hover:text-darb-good">+ شحن 500</button>
        </div>
        <div className="rounded-2xl bg-darb-card border border-darb-line p-4 flex items-center justify-between">
          <div>
            <div className="text-xs text-darb-mut">💎 رصيد النقاط</div>
            <div className="text-2xl font-extrabold text-darb-accent">{fmtInt(points)} نقطة</div>
            <div className="text-[10px] text-darb-mut">= {fmtSar(points / pv)}</div>
          </div>
          <button onClick={() => { setWallet(1500); setPoints(150000); setBought([]); setLast(null); }} className="text-xs font-bold px-3 py-2 rounded-lg border border-darb-line hover:border-darb-orange text-darb-mut hover:text-darb-orange">↺ إعادة</button>
        </div>
      </div>

      {last && <div className="stat border-darb-good/40 text-sm text-darb-good">{last}</div>}

      <div className="grid md:grid-cols-3 gap-3">
        {BUNDLES.map((b) => {
          const worth = b.items.reduce((a, x) => a + x.worth, 0);
          const savings = worth - b.price;
          const pointsPrice = b.price * pv;
          const done = bought.includes(b.id);
          const canCash = b.price <= wallet;
          const canPoints = pointsPrice <= points;
          return (
            <div key={b.id} className={`card flex flex-col ${done ? "opacity-60" : ""}`}>
              <div className="flex items-center justify-between">
                <div className="text-3xl">{b.emoji}</div>
                <Badge tone="accent">{b.tag}</Badge>
              </div>
              <div className="mt-2 text-lg font-extrabold">{b.name}</div>
              <div className="flex items-baseline gap-2">
                <span className="text-2xl font-extrabold text-darb-orange">{fmtSar(b.price)}</span>
                <span className="text-xs text-darb-accent">أو {fmtInt(pointsPrice)} نقطة</span>
              </div>

              <ul className="mt-3 space-y-1.5 text-sm grow">
                {b.items.map((it, idx) => (
                  <li key={idx} className="flex items-center justify-between border-b border-darb-line/40 pb-1.5">
                    <span className="text-darb-ink/90">• {it.label}</span>
                    <span className="font-bold text-darb-ink">{it.qty}</span>
                  </li>
                ))}
              </ul>

              <div className="mt-3 flex items-center justify-between text-xs">
                <span className="text-darb-mut">قيمة تقديرية {fmtSar(worth)}</span>
                {savings > 0 && <Badge tone="good">توفير {fmtSar(savings)}</Badge>}
              </div>

              {done ? (
                <div className="mt-3"><Badge tone="good">✅ تم</Badge></div>
              ) : (
                <div className="mt-3 grid grid-cols-2 gap-2">
                  <button onClick={() => buyCash(b)} disabled={!canCash}
                    className={`text-xs font-bold px-2 py-2 rounded-lg border transition ${canCash ? "border-darb-orange text-darb-orange hover:bg-darb-orange/15" : "border-darb-line text-darb-mut/40 cursor-not-allowed"}`}>
                    👛 شراء
                  </button>
                  <button onClick={() => redeemPoints(b)} disabled={!canPoints}
                    className={`text-xs font-bold px-2 py-2 rounded-lg border transition ${canPoints ? "border-darb-accent text-darb-accent hover:bg-darb-accent/15" : "border-darb-line text-darb-mut/40 cursor-not-allowed"}`}>
                    💎 استبدال
                  </button>
                </div>
              )}
            </div>
          );
        })}
      </div>

      <Card>
        <p className="text-sm text-darb-ink/90 leading-relaxed">
          💡 الباقة لها طريقتان: <b>شراء بالفلوس</b> (محفظة الشحن) — تعطي درب سيولة، أو <b>استبدال بالنقاط</b>
          (مغطّاة بالإسكرو، فدرب تموّل الباقة من نفس الرصيد). والباقات الكبيرة تصير <b>هدفاً تحفيزياً</b> يجمّع
          له العميل نقاطه (أثر تدرّج الهدف).
        </p>
      </Card>
    </div>
  );
}
