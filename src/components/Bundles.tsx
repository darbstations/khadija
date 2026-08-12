import { useState } from "react";
import { useScenario } from "../context/ScenarioContext";
import { fmtSar, fmtInt } from "../model/engine";
import { BUNDLES } from "../model/defaults";
import { Card, Badge } from "./ui";

type Method = "wallet" | "card" | "tabby" | "tamara";
const METHODS: { key: Method; label: string; installments?: number }[] = [
  { key: "wallet", label: "👛 محفظة درب" },
  { key: "card", label: "💳 بطاقة" },
  { key: "tabby", label: "تابي (تقسيط)", installments: 4 },
  { key: "tamara", label: "تمارا (تقسيط)", installments: 4 },
];

export default function Bundles() {
  const { inputs } = useScenario();
  const pv = inputs.pointValue;
  const [wallet, setWallet] = useState(1500);
  const [points, setPoints] = useState(150000);
  const [method, setMethod] = useState<Method>("wallet");
  const [bought, setBought] = useState<string[]>([]);
  const [last, setLast] = useState<string | null>(null);

  const m = METHODS.find((x) => x.key === method)!;

  const buy = (b: (typeof BUNDLES)[number]) => {
    if (bought.includes(b.id)) return;
    if (method === "wallet") {
      if (b.price > wallet) return;
      setWallet((w) => w - b.price);
    }
    setBought((p) => [...p, b.id]);
    const inst = m.installments ? ` · ${m.installments} دفعات ${fmtSar(b.price / m.installments)}` : "";
    setLast(`✅ اشتريت «${b.name}» عبر ${m.label.replace(/^[^ ]+ /, "")}${inst}`);
  };
  const redeemPoints = (b: (typeof BUNDLES)[number]) => {
    const cost = b.price * pv;
    if (cost > points || bought.includes(b.id)) return;
    setPoints((p) => p - cost);
    setBought((p) => [...p, b.id]);
    setLast(`✅ استبدلت «${b.name}» بـ ${fmtInt(cost)} نقطة`);
  };

  const canBuy = (price: number) => method !== "wallet" || price <= wallet;

  return (
    <div className="space-y-4">
      <div>
        <h2 className="text-lg font-extrabold">🏪 متجر الباقات · درب</h2>
        <p className="text-xs text-darb-mut">اشترِ بأي وسيلة (محفظة/بطاقة/تقسيط) أو استبدل بالنقاط — بدون حاجة لشحن المحفظة</p>
      </div>

      {/* الرصيدان */}
      <div className="grid sm:grid-cols-2 gap-3">
        <div className="rounded-2xl bg-gradient-to-bl from-darb-orange/25 to-darb-card border border-darb-line p-4 flex items-center justify-between">
          <div>
            <div className="text-xs text-darb-mut">👛 محفظة درب (ريال)</div>
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

      {/* وسيلة الدفع */}
      <Card>
        <div className="flex items-center gap-2 flex-wrap">
          <span className="text-xs font-bold text-darb-ink">وسيلة الشراء:</span>
          {METHODS.map((x) => (
            <button key={x.key} onClick={() => setMethod(x.key)}
              className={`text-xs font-bold px-3 py-1.5 rounded-lg border transition ${method === x.key ? "border-darb-orange bg-darb-orange/15 text-darb-orange" : "border-darb-line text-darb-mut hover:text-darb-ink"}`}>
              {x.label}
            </button>
          ))}
        </div>
      </Card>

      {last && <div className="stat border-darb-good/40 text-sm text-darb-good">{last}</div>}

      <div className="grid md:grid-cols-3 gap-3">
        {BUNDLES.map((b) => {
          const worth = b.items.reduce((a, x) => a + x.worth, 0);
          const savings = worth - b.price;
          const pointsPrice = b.price * pv;
          const done = bought.includes(b.id);
          const affordBuy = canBuy(b.price);
          const affordPoints = pointsPrice <= points;
          return (
            <div key={b.id} className={`card flex flex-col ${done ? "opacity-60" : ""}`}>
              <div className="flex items-center justify-between">
                <div className="text-3xl">{b.emoji}</div>
                <Badge tone="accent">{b.tag}</Badge>
              </div>
              <div className="mt-2 text-lg font-extrabold">{b.name}</div>
              {b.desc && <div className="text-xs text-darb-mut mb-1">{b.desc}</div>}
              <div className="flex items-baseline gap-2">
                <span className="text-2xl font-extrabold text-darb-orange">{fmtSar(b.price)}</span>
                <span className="text-xs text-darb-accent">أو {fmtInt(pointsPrice)} نقطة</span>
              </div>
              {m.installments && (
                <div className="text-[11px] text-darb-good mt-0.5">أو {m.installments} دفعات {fmtSar(b.price / m.installments)} عبر {m.label.replace(/^[^ ]+ /, "").replace(" (تقسيط)", "")}</div>
              )}

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
                  <button onClick={() => buy(b)} disabled={!affordBuy}
                    className={`text-xs font-bold px-2 py-2 rounded-lg border transition ${affordBuy ? "border-darb-orange text-darb-orange hover:bg-darb-orange/15" : "border-darb-line text-darb-mut/40 cursor-not-allowed"}`}>
                    شراء
                  </button>
                  <button onClick={() => redeemPoints(b)} disabled={!affordPoints}
                    className={`text-xs font-bold px-2 py-2 rounded-lg border transition ${affordPoints ? "border-darb-accent text-darb-accent hover:bg-darb-accent/15" : "border-darb-line text-darb-mut/40 cursor-not-allowed"}`}>
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
          💡 الباقة تُشترى بـ<b>محفظة درب أو بطاقة أو تقسيط (تابي/تمارا)</b> — فلا تحتاج شحن المحفظة، ويقدر العميل
          يقسّطها. أو <b>يستبدلها بنقاطه</b>. التقسيط يرفع مبيعات الباقات الكبيرة (فزيع/الشاهين).
        </p>
      </Card>
    </div>
  );
}
