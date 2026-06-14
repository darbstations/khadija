import { useState } from "react";
import { fmtSar } from "../model/engine";
import { BUNDLES } from "../model/defaults";
import { Card, Badge } from "./ui";

export default function Bundles() {
  const [wallet, setWallet] = useState(1500); // رصيد محفظة درب (ريال)
  const [bought, setBought] = useState<string[]>([]);
  const [last, setLast] = useState<string | null>(null);

  const buy = (b: (typeof BUNDLES)[number]) => {
    if (b.price > wallet || bought.includes(b.id)) return;
    setWallet((w) => w - b.price);
    setBought((p) => [...p, b.id]);
    setLast(`✅ اشتريت «${b.name}» مقابل ${fmtSar(b.price)} من محفظة درب`);
  };

  return (
    <div className="space-y-4">
      <div>
        <h2 className="text-lg font-extrabold">🪪 باقات درب · بطاقات مدفوعة مسبقاً</h2>
        <p className="text-xs text-darb-mut">
          منتجات مستقلة عن النقاط · تُشترى من <b>محفظة درب</b> (رصيد بالريال)
        </p>
      </div>

      {/* محفظة درب */}
      <div className="rounded-2xl bg-gradient-to-bl from-darb-orange/25 to-darb-card border border-darb-line p-5 flex items-center justify-between flex-wrap gap-3">
        <div>
          <div className="text-xs text-darb-mut">رصيد محفظة درب</div>
          <div className="text-3xl font-extrabold text-darb-orange">{fmtSar(wallet)}</div>
        </div>
        <div className="flex items-center gap-2">
          <button onClick={() => setWallet((w) => w + 500)} className="text-xs font-bold px-3 py-2 rounded-lg border border-darb-line hover:border-darb-good text-darb-mut hover:text-darb-good transition">+ شحن 500 ﷼</button>
          <button onClick={() => { setWallet(1500); setBought([]); setLast(null); }} className="text-xs font-bold px-3 py-2 rounded-lg border border-darb-line hover:border-darb-orange text-darb-mut hover:text-darb-orange transition">↺ إعادة</button>
        </div>
      </div>

      {last && <div className="stat border-darb-good/40 text-sm text-darb-good">{last}</div>}

      <div className="grid md:grid-cols-3 gap-3">
        {BUNDLES.map((b) => {
          const worth = b.items.reduce((a, x) => a + x.worth, 0);
          const savings = worth - b.price;
          const done = bought.includes(b.id);
          const afford = b.price <= wallet;
          return (
            <div key={b.id} className={`card flex flex-col ${done ? "opacity-60" : ""}`}>
              <div className="flex items-center justify-between">
                <div className="text-3xl">{b.emoji}</div>
                <Badge tone="accent">{b.tag}</Badge>
              </div>
              <div className="mt-2 text-lg font-extrabold">{b.name}</div>
              <div className="text-2xl font-extrabold text-darb-orange">{fmtSar(b.price)}</div>

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
                <div className="mt-3"><Badge tone="good">✅ تم الشراء</Badge></div>
              ) : (
                <button
                  onClick={() => buy(b)}
                  disabled={!afford}
                  className={`mt-3 w-full text-sm font-bold px-3 py-2 rounded-lg border transition ${
                    afford ? "border-darb-orange text-darb-orange hover:bg-darb-orange/15" : "border-darb-line text-darb-mut/40 cursor-not-allowed"
                  }`}
                >
                  {afford ? "اشترِ من محفظة درب" : "رصيد المحفظة غير كافٍ"}
                </button>
              )}
            </div>
          );
        })}
      </div>

      <Card>
        <p className="text-sm text-darb-ink/90 leading-relaxed">
          💡 الباقات <b>منتجات مستقلة</b> لا تُكتسب بالنقاط — تُشترى نقداً من محفظة درب وتعطي العميل
          رصيد بنزين + خدمات شركاء بسعر مجمّع. فائدتها لدرب: <b>سيولة مقدّماً</b> + التزام إنفاق +
          بيع متقاطع لخدمات الشركاء.
        </p>
      </Card>
    </div>
  );
}
