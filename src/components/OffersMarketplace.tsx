import { useState } from "react";
import { useScenario } from "../context/ScenarioContext";
import { fmtInt, fmtSar } from "../model/engine";
import { OFFER_LOCATIONS, OFFER_CATEGORIES, DRAW_PRIZES } from "../model/defaults";
import { useOffers, visibleOffers, type Offer } from "../model/offers";
import { Card, Badge } from "./ui";

export default function OffersMarketplace() {
  const { inputs } = useScenario();
  const { offers: allOffers } = useOffers();
  const [location, setLocation] = useState(OFFER_LOCATIONS[1]);
  const [cat, setCat] = useState("all");
  const [balance, setBalance] = useState(8000); // رصيد العميل (نقطة)
  const [redeemed, setRedeemed] = useState<number[]>([]);
  const [last, setLast] = useState<string | null>(null);

  // العميل يرى فقط العروض المعتمدة والنشطة من درب
  const offers = visibleOffers(allOffers).filter(
    (o) => (location === "كل المواقع" || o.loc === location) && (cat === "all" || o.cat === cat)
  );

  const redeem = (o: Offer) => {
    if (o.points > balance || redeemed.includes(o.id)) return;
    setBalance((b) => b - o.points);
    setRedeemed((r) => [...r, o.id]);
    setLast(`✅ استبدلت «${o.title}» من ${o.merchant} مقابل ${fmtInt(o.points)} نقطة`);
  };

  const [entries, setEntries] = useState<Record<number, number>>({});
  const pv = inputs.pointValue;

  const spend = (label: string, cost: number) => {
    if (cost > balance) return;
    setBalance((b) => b - cost);
    setLast(`✅ ${label} مقابل ${fmtInt(cost)} نقطة`);
  };
  const enterDraw = (id: number, name: string) => {
    // دخول مجاني — بلا نقاط (حلال)
    setEntries((e) => ({ ...e, [id]: (e[id] || 0) + 1 }));
    setLast(`🎁 دخلت السحب المجاني على «${name}»`);
  };

  const quick = [
    { icon: "⛽", label: "خصم بنزين 10﷼", cost: 10 * pv },
    { icon: "📱", label: "شحن STC 10﷼", cost: 10 * pv },
    { icon: "📱", label: "شحن Mobily 10﷼", cost: 10 * pv },
    { icon: "📱", label: "شحن Zain 10﷼", cost: 10 * pv },
    { icon: "🎁", label: "قسيمة جرير 25﷼", cost: 25 * pv },
    { icon: "🎬", label: "اشتراك شاهد شهر", cost: 3000 },
    { icon: "❤️", label: "تبرّع 5﷼", cost: 5 * pv },
  ];

  const catColor = (k: string) =>
    ({ restaurant: "bg-darb-orange/15 text-darb-orange", cafe: "bg-amber-400/15 text-amber-300", wash: "bg-sky-400/15 text-sky-300", service: "bg-darb-mut/20 text-darb-ink", grocery: "bg-darb-good/15 text-darb-good" }[k] || "bg-darb-line text-darb-mut");

  return (
    <div className="space-y-4">
      <div>
        <h2 className="text-lg font-extrabold">🎁 سوق العروض · استبدل نقاطك</h2>
        <p className="text-xs text-darb-mut">عروض متاجر المحطة قربك — مطاعم · كافيهات · مغاسل · خدمة · بقالة</p>
      </div>

      {/* بطاقة الرصيد */}
      <div className="rounded-2xl bg-gradient-to-bl from-darb-orange/25 to-darb-card border border-darb-line p-5 flex items-center justify-between flex-wrap gap-3">
        <div>
          <div className="text-xs text-darb-mut">رصيد نقاطك</div>
          <div className="text-3xl font-extrabold text-darb-orange">{fmtInt(balance)} نقطة</div>
          <div className="text-xs text-darb-mut mt-1">≈ {fmtSar(balance / inputs.pointValue)} للاستبدال</div>
        </div>
        <button
          onClick={() => { setBalance(8000); setRedeemed([]); setLast(null); }}
          className="no-print text-xs font-bold px-3 py-2 rounded-lg border border-darb-line hover:border-darb-orange text-darb-mut hover:text-darb-orange transition"
        >
          ↺ إعادة التجربة
        </button>
      </div>

      {last && (
        <div className="stat border-darb-good/40 text-sm text-darb-good">{last}</div>
      )}

      {/* الموقع + الفئات */}
      <Card>
        <div className="flex items-center gap-2 flex-wrap mb-3">
          <span className="text-sm">📍</span>
          <select
            value={location}
            onChange={(e) => setLocation(e.target.value)}
            className="bg-darb-panel border border-darb-line rounded-lg px-3 py-2 text-sm font-bold"
          >
            {OFFER_LOCATIONS.map((l) => (
              <option key={l} value={l}>{l}</option>
            ))}
          </select>
          <span className="text-[11px] text-darb-mut">العروض تظهر حسب موقعك</span>
        </div>
        <div className="flex gap-1.5 flex-wrap">
          {OFFER_CATEGORIES.map((c) => (
            <button
              key={c.key}
              onClick={() => setCat(c.key)}
              className={`text-xs font-bold px-3 py-1.5 rounded-full border transition ${
                cat === c.key ? "border-darb-orange bg-darb-orange/15 text-darb-orange" : "border-darb-line text-darb-mut hover:text-darb-ink"
              }`}
            >
              {c.emoji} {c.label}
            </button>
          ))}
        </div>
      </Card>

      {/* شبكة العروض */}
      {offers.length === 0 ? (
        <Card><p className="text-sm text-darb-mut text-center py-6">لا توجد عروض في هذا الموقع/الفئة.</p></Card>
      ) : (
        <div className="grid sm:grid-cols-2 lg:grid-cols-3 gap-3">
          {offers.map((o) => {
            const done = redeemed.includes(o.id);
            const afford = o.points <= balance;
            return (
              <div key={o.id} className={`card flex flex-col ${done ? "opacity-60" : ""}`}>
                <div className="flex items-start justify-between">
                  <div className="text-3xl">{o.emoji}</div>
                  <span className={`pill ${catColor(o.cat)}`}>
                    {OFFER_CATEGORIES.find((c) => c.key === o.cat)?.label}
                  </span>
                </div>
                <div className="mt-2 font-extrabold text-darb-ink">{o.title}</div>
                <div className="text-xs text-darb-mut">{o.merchant}</div>
                <div className="text-[11px] text-darb-mut mt-1">📍 {o.loc}</div>
                <div className="flex items-center justify-between mt-3 pt-3 border-t border-darb-line">
                  <div>
                    <div className="font-extrabold text-darb-orange">{fmtInt(o.points)} نقطة</div>
                    <div className="text-[10px] text-darb-mut">القيمة {fmtSar(o.value)}</div>
                  </div>
                  {done ? (
                    <Badge tone="good">✅ مُستبدل</Badge>
                  ) : (
                    <button
                      onClick={() => redeem(o)}
                      disabled={!afford}
                      className={`text-xs font-bold px-3 py-1.5 rounded-lg border transition ${
                        afford ? "border-darb-orange text-darb-orange hover:bg-darb-orange/15" : "border-darb-line text-darb-mut/40 cursor-not-allowed"
                      }`}
                    >
                      {afford ? "استبدل" : "نقاط غير كافية"}
                    </button>
                  )}
                </div>
              </div>
            );
          })}
        </div>
      )}

      <Card>
        <p className="text-sm text-darb-ink/90 leading-relaxed">
          💡 العروض يصدرها كل متجر بنفسه ويسعّرها بالنقاط · تظهر للعميل <b>حسب موقعه</b> (محطته الأقرب) ·
        </p>
      </Card>

      {/* السحب المجاني (حلال) */}
      <Card title="🎁 السحب المجاني على جوائز كبرى" subtitle="دخول مجاني بلا نقاط — عن كل تعبئة/شهر">
        <div className="grid sm:grid-cols-2 lg:grid-cols-4 gap-3">
          {DRAW_PRIZES.map((p, idx) => (
            <div key={p.name} className="card text-center">
              <div className="text-4xl">{p.icon}</div>
              <div className="font-extrabold mt-1">{p.name}</div>
              <div className="text-[11px] text-darb-mut">قيمتها {fmtSar(p.value)}</div>
              {entries[idx] ? <div className="my-1"><Badge tone="good">دخولاتك: {entries[idx]}</Badge></div> : null}
              <button
                onClick={() => enterDraw(idx, p.name)}
                className="mt-2 w-full text-xs font-bold px-3 py-2 rounded-lg border border-darb-good text-darb-good hover:bg-darb-good/15 transition"
              >
                ادخل السحب (مجاناً)
              </button>
            </div>
          ))}
        </div>
        <p className="text-[11px] text-darb-mut mt-2">✅ الدخول مجاني بلا نقاط ولا رسوم — لا مقابل مدفوع (ليس ميساراً). تموّله درب كحملة.</p>
      </Card>

      {/* استبدال سريع */}
      <Card title="⚡ استبدال سريع — قنوات أخرى">
        <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-5 gap-2">
          {quick.map((q) => {
            const can = q.cost <= balance;
            return (
              <button
                key={q.label}
                onClick={() => spend(q.label, q.cost)}
                disabled={!can}
                className={`text-xs font-bold px-2 py-3 rounded-lg border transition text-center ${can ? "border-darb-line hover:border-darb-orange hover:bg-darb-orange/10 text-darb-ink" : "border-darb-line text-darb-mut/40 cursor-not-allowed"}`}
              >
                <div className="text-2xl">{q.icon}</div>
                <div className="mt-1">{q.label}</div>
                <div className="text-[10px] text-darb-mut">{fmtInt(q.cost)} نقطة</div>
              </button>
            );
          })}
        </div>
        <p className="text-[11px] text-darb-mut mt-2">وأيضاً: خدمات المحطة · تحويل أميال/قطاف · إهداء النقاط للعائلة · تحويل لرصيد المحفظة.</p>
      </Card>

      <Card>
        <p className="text-sm text-darb-ink/90 leading-relaxed">
          💡 كل قنوات الاستبدال في مكان واحد: عروض التجار · السحوبات · قنوات سريعة — والرصيد <b>موحّد</b>.
          والاستبدال يُخصم من رصيد نقاطه ويُموّل عبر <b>نموذج المحفظة</b> (التاجر يستلم من المحفظة + يجذب عميلاً).
        </p>
      </Card>
    </div>
  );
}
