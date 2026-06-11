import { useState } from "react";
import { useScenario } from "../context/ScenarioContext";
import { fmtSar, fmtInt, fmtNum } from "../model/engine";
import { Card, Stat, Badge } from "./ui";

interface Purchase {
  id: number;
  source: string;
  amount: number; // ريال
  rate: number; // % يودعها المُصدِر في المحفظة
}
interface Offer {
  id: number;
  merchant: string;
  walletDraw: number; // ريال يُسحب من المحفظة
  customerValue: number; // القيمة التي يحصلها العميل عند التاجر
}

export default function WalletRedemption() {
  const { inputs } = useScenario();
  const [purchases, setPurchases] = useState<Purchase[]>([
    { id: 1, source: "⛽ بنزين (درب)", amount: 600, rate: 0.5 },
    { id: 2, source: "☕ كافيه", amount: 200, rate: 4 },
    { id: 3, source: "🍔 مطعم", amount: 150, rate: 3 },
    { id: 4, source: "🛒 سوبرماركت", amount: 400, rate: 1 },
  ]);
  const [offers, setOffers] = useState<Offer[]>([
    { id: 1, merchant: "⛽ درب (خصم بنزين)", walletDraw: 10, customerValue: 10 },
    { id: 2, merchant: "☕ كافيه", walletDraw: 10, customerValue: 13 },
    { id: 3, merchant: "🍔 مطعم", walletDraw: 10, customerValue: 12 },
    { id: 4, merchant: "🎁 جرير (قسيمة)", walletDraw: 10, customerValue: 10.5 },
  ]);
  const [redeemed, setRedeemed] = useState(0); // إجمالي ما سُحب من المحفظة
  const [last, setLast] = useState<{ merchant: string; draw: number; bonus: number; value: number } | null>(null);

  const funded = purchases.reduce((a, p) => a + (p.amount * p.rate) / 100, 0);
  const balance = funded - redeemed;
  const points = balance * inputs.pointValue;

  const updateP = (id: number, patch: Partial<Purchase>) =>
    setPurchases((p) => p.map((x) => (x.id === id ? { ...x, ...patch } : x)));
  const updateO = (id: number, patch: Partial<Offer>) =>
    setOffers((o) => o.map((x) => (x.id === id ? { ...x, ...patch } : x)));

  const sortedOffers = [...offers].sort((a, b) => b.customerValue - a.customerValue);
  const best = sortedOffers[0]?.id;

  const doRedeem = (o: Offer) => {
    if (o.walletDraw > balance) return;
    setRedeemed((r) => r + o.walletDraw);
    setLast({ merchant: o.merchant, draw: o.walletDraw, bonus: o.customerValue - o.walletDraw, value: o.customerValue });
  };

  return (
    <div className="space-y-4">
      <div>
        <h2 className="text-lg font-extrabold">👛 المحفظة وسوق الاستبدال</h2>
        <p className="text-xs text-darb-mut">
          كل عملية تودع ريالاً حقيقياً في محفظة العميل · والتجار يتنافسون على سحبها عند الاستبدال
        </p>
      </div>

      <div className="grid sm:grid-cols-4 gap-3">
        <Stat label="💰 رصيد المحفظة (مغطّى بالريال)" value={fmtSar(balance, 2)} tone="good" />
        <Stat label="💎 يُعرض للعميل كنقاط" value={fmtInt(points)} tone="accent" />
        <Stat label="📥 إجمالي ما تغذّت به" value={fmtSar(funded, 2)} />
        <Stat label="📤 صُرف للتجار" value={fmtSar(redeemed, 2)} tone="warn" />
      </div>

      <Card title="📥 كيف تتغذّى المحفظة" subtitle="كل مُصدِر يودع نسبته على مبيعاته هو فقط — لا أحد يموّل غيره">
        <div className="overflow-x-auto">
          <table className="w-full border-collapse">
            <thead>
              <tr>
                <th className="th">المصدر</th>
                <th className="th">قيمة الشراء (﷼)</th>
                <th className="th">نسبة الإيداع %</th>
                <th className="th">المودَع في المحفظة</th>
              </tr>
            </thead>
            <tbody>
              {purchases.map((p) => (
                <tr key={p.id}>
                  <td className="td font-bold">{p.source}</td>
                  <td className="td">
                    <input
                      type="number"
                      value={p.amount}
                      onChange={(e) => updateP(p.id, { amount: parseFloat(e.target.value) || 0 })}
                      className="bg-darb-yellow/10 border border-darb-yellow/30 rounded px-2 py-1 w-24 text-left"
                      dir="ltr"
                    />
                  </td>
                  <td className="td">
                    <input
                      type="number"
                      step={0.5}
                      value={p.rate}
                      onChange={(e) => updateP(p.id, { rate: parseFloat(e.target.value) || 0 })}
                      className="bg-darb-yellow/10 border border-darb-yellow/30 rounded px-2 py-1 w-20 text-left"
                      dir="ltr"
                    />
                  </td>
                  <td className="td text-darb-good font-bold">{fmtSar((p.amount * p.rate) / 100, 2)}</td>
                </tr>
              ))}
              <tr className="bg-darb-panel/60 font-bold">
                <td className="td" colSpan={3}>
                  📊 إجمالي التغذية
                </td>
                <td className="td text-darb-good">{fmtSar(funded, 2)}</td>
              </tr>
            </tbody>
          </table>
        </div>
      </Card>

      <Card
        title="🏷️ سوق عروض الاستبدال (مناقصة التجار)"
        subtitle="كل تاجر ينزل عرضاً يجذب محفظة العميل · يتنافسون على القيمة"
      >
        {last && (
          <div className="stat mb-3 border-darb-accent/40">
            <div className="text-sm">
              ✅ آخر استبدال عند <b>{last.merchant}</b>: سُحب{" "}
              <b className="text-darb-warn">{fmtSar(last.draw, 2)}</b> من المحفظة، أضاف التاجر{" "}
              <b className="text-darb-accent">{fmtSar(last.bonus, 2)}</b> من جيبه، فحصل العميل على{" "}
              <b className="text-darb-good">{fmtSar(last.value, 2)}</b>.
            </div>
          </div>
        )}
        <div className="overflow-x-auto">
          <table className="w-full border-collapse">
            <thead>
              <tr>
                <th className="th">التاجر</th>
                <th className="th">يُسحب من المحفظة (﷼)</th>
                <th className="th">يحصل العميل (﷼)</th>
                <th className="th">يضيف التاجر (تنافس)</th>
                <th className="th">القيمة/﷼</th>
                <th className="th"></th>
              </tr>
            </thead>
            <tbody>
              {sortedOffers.map((o) => {
                const bonus = o.customerValue - o.walletDraw;
                const mult = o.walletDraw ? o.customerValue / o.walletDraw : 0;
                return (
                  <tr key={o.id} className={o.id === best ? "bg-darb-good/10" : ""}>
                    <td className="td font-bold">
                      {o.merchant} {o.id === best && <Badge tone="good">أفضل عرض</Badge>}
                    </td>
                    <td className="td">
                      <input
                        type="number"
                        value={o.walletDraw}
                        onChange={(e) => updateO(o.id, { walletDraw: parseFloat(e.target.value) || 0 })}
                        className="bg-darb-yellow/10 border border-darb-yellow/30 rounded px-2 py-1 w-20 text-left"
                        dir="ltr"
                      />
                    </td>
                    <td className="td">
                      <input
                        type="number"
                        step={0.5}
                        value={o.customerValue}
                        onChange={(e) => updateO(o.id, { customerValue: parseFloat(e.target.value) || 0 })}
                        className="bg-darb-yellow/10 border border-darb-yellow/30 rounded px-2 py-1 w-20 text-left"
                        dir="ltr"
                      />
                    </td>
                    <td className={`td ${bonus > 0 ? "text-darb-accent" : "text-darb-mut"}`}>
                      {bonus > 0 ? `+${fmtSar(bonus, 2)}` : "—"}
                    </td>
                    <td className="td">{fmtNum(mult, 2)}x</td>
                    <td className="td">
                      <button
                        onClick={() => doRedeem(o)}
                        disabled={o.walletDraw > balance}
                        className={`text-xs font-bold px-3 py-1.5 rounded-lg border transition ${
                          o.walletDraw > balance
                            ? "border-darb-line text-darb-mut/40 cursor-not-allowed"
                            : "border-darb-accent text-darb-accent hover:bg-darb-accent/15"
                        }`}
                      >
                        استبدل
                      </button>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
        <div className="flex items-center justify-between mt-3">
          <p className="text-xs text-darb-mut max-w-2xl">
            💡 الاستبدال يُسحب من <b>محفظة العميل</b> (فلوس مودعة مسبقاً) — التاجر يستلمها <b>ولا يموّل أحداً</b>،
            بل يضيف من جيبه فقط ليفوز بالعميل. الاستبدال صار <b>قناة طلب</b> يتنافسون عليها.
          </p>
          {redeemed > 0 && (
            <button
              onClick={() => {
                setRedeemed(0);
                setLast(null);
              }}
              className="text-xs font-bold px-3 py-1.5 rounded-lg border border-darb-line hover:border-darb-accent text-darb-mut hover:text-darb-accent transition shrink-0"
            >
              ↺ تصفير الاستبدال
            </button>
          )}
        </div>
      </Card>

      <Card>
        <p className="text-sm text-darb-ink/90 leading-relaxed">
          🔒 <b>النظام لا يخسر:</b> كل ريال يُستبدل كان مودَعاً مسبقاً في المحفظة، فلا يوجد التزام مكشوف.
          النقاط مجرد <b>واجهة عرض</b> فوق محفظة حقيقية بالريال — وهذا يحل الثقة والتضخيم معاً، ويزيل اعتراض
          التجار لأن لا أحد يموّل وعود غيره.
        </p>
      </Card>
    </div>
  );
}
