import { useState } from "react";
import { fmtInt } from "../model/engine";
import { OFFER_CATEGORIES, OFFER_LOCATIONS } from "../model/defaults";
import { useOffers, ACCOUNT_MANAGERS, emojiForCat, type Offer, type OfferStatus } from "../model/offers";
import { Card, Stat, Badge } from "./ui";

const CATS = OFFER_CATEGORIES.filter((c) => c.key !== "all");
const LOCS = OFFER_LOCATIONS.filter((l) => l !== "كل المواقع");
const statusBadge = (s: OfferStatus) =>
  s === "approved" ? <Badge tone="good">معتمد</Badge> : s === "pending" ? <Badge tone="warn">قيد المراجعة</Badge> : <Badge tone="bad">مرفوض</Badge>;

export default function OffersAdmin() {
  const { offers, setOffers } = useOffers();
  const [mgr, setMgr] = useState("all");
  const [status, setStatus] = useState("all");

  const update = (id: number, patch: Partial<Offer>) =>
    setOffers((p) => p.map((o) => (o.id === id ? { ...o, ...patch } : o)));
  const remove = (id: number) => setOffers((p) => p.filter((o) => o.id !== id));
  const add = () =>
    setOffers((p) => [
      {
        id: Math.max(0, ...p.map((o) => o.id)) + 1,
        merchant: "تاجر جديد",
        title: "عرض جديد",
        cat: "cafe",
        emoji: emojiForCat("cafe"),
        loc: LOCS[0],
        points: 1000,
        value: 10,
        status: "pending" as OfferStatus,
        active: false,
        manager: ACCOUNT_MANAGERS[0],
      },
      ...p,
    ]);

  const filtered = offers.filter(
    (o) => (mgr === "all" || o.manager === mgr) && (status === "all" || o.status === status)
  );
  const count = (f: (o: Offer) => boolean) => offers.filter(f).length;

  return (
    <div className="space-y-4">
      <div>
        <h2 className="text-lg font-extrabold">🛠️ إدارة العروض · درب</h2>
        <p className="text-xs text-darb-mut">
          التاجر يقترح · مدير الحساب في درب يعتمد ويفعّل · لا ينشر التاجر بنفسه
        </p>
      </div>

      <div className="grid sm:grid-cols-4 gap-3">
        <Stat label="إجمالي العروض" value={fmtInt(offers.length)} />
        <Stat label="معتمدة" value={fmtInt(count((o) => o.status === "approved"))} tone="good" />
        <Stat label="قيد المراجعة" value={fmtInt(count((o) => o.status === "pending"))} tone="warn" />
        <Stat label="ظاهرة للعميل" value={fmtInt(count((o) => o.status === "approved" && o.active))} tone="accent" />
      </div>

      <Card>
        <div className="flex items-center gap-2 flex-wrap mb-3">
          <span className="text-xs font-bold text-darb-ink">مدير الحساب:</span>
          <select value={mgr} onChange={(e) => setMgr(e.target.value)} className="bg-darb-panel border border-darb-line rounded-lg px-2 py-1.5 text-xs">
            <option value="all">الكل</option>
            {ACCOUNT_MANAGERS.map((m) => <option key={m} value={m}>{m}</option>)}
          </select>
          <span className="text-xs font-bold text-darb-ink">الحالة:</span>
          <select value={status} onChange={(e) => setStatus(e.target.value)} className="bg-darb-panel border border-darb-line rounded-lg px-2 py-1.5 text-xs">
            <option value="all">الكل</option>
            <option value="approved">معتمد</option>
            <option value="pending">قيد المراجعة</option>
            <option value="rejected">مرفوض</option>
          </select>
          <button onClick={add} className="ms-auto text-xs font-bold px-3 py-1.5 rounded-lg border border-darb-orange text-darb-orange hover:bg-darb-orange/15 transition">
            + عرض جديد (يقترحه التاجر)
          </button>
        </div>

        <div className="overflow-x-auto">
          <table className="w-full border-collapse">
            <thead>
              <tr>
                <th className="th">العرض</th>
                <th className="th">التاجر</th>
                <th className="th">الفئة</th>
                <th className="th">الموقع</th>
                <th className="th">مدير الحساب</th>
                <th className="th">نقاط</th>
                <th className="th">قيمة</th>
                <th className="th">الحالة</th>
                <th className="th">نشط</th>
                <th className="th">إجراءات</th>
              </tr>
            </thead>
            <tbody>
              {filtered.map((o) => (
                <tr key={o.id} className="hover:bg-darb-panel/40">
                  <td className="td">
                    <input value={o.title} onChange={(e) => update(o.id, { title: e.target.value })}
                      className="bg-transparent border-b border-darb-line/50 focus:border-darb-orange outline-none w-32 text-darb-ink font-bold" />
                  </td>
                  <td className="td">
                    <input value={o.merchant} onChange={(e) => update(o.id, { merchant: e.target.value })}
                      className="bg-transparent border-b border-darb-line/50 focus:border-darb-orange outline-none w-24 text-darb-ink" />
                  </td>
                  <td className="td">
                    <select value={o.cat} onChange={(e) => update(o.id, { cat: e.target.value, emoji: emojiForCat(e.target.value) })}
                      className="bg-darb-panel border border-darb-line rounded px-1.5 py-1 text-xs">
                      {CATS.map((c) => <option key={c.key} value={c.key}>{c.emoji} {c.label}</option>)}
                    </select>
                  </td>
                  <td className="td">
                    <select value={o.loc} onChange={(e) => update(o.id, { loc: e.target.value })}
                      className="bg-darb-panel border border-darb-line rounded px-1.5 py-1 text-xs">
                      {LOCS.map((l) => <option key={l} value={l}>{l}</option>)}
                    </select>
                  </td>
                  <td className="td">
                    <select value={o.manager} onChange={(e) => update(o.id, { manager: e.target.value })}
                      className="bg-darb-panel border border-darb-line rounded px-1.5 py-1 text-xs">
                      {ACCOUNT_MANAGERS.map((m) => <option key={m} value={m}>{m}</option>)}
                    </select>
                  </td>
                  <td className="td">
                    <input type="number" value={o.points} onChange={(e) => update(o.id, { points: parseInt(e.target.value) || 0 })}
                      className="bg-darb-yellow/10 border border-darb-yellow/30 rounded px-2 py-1 w-20 text-left" dir="ltr" />
                  </td>
                  <td className="td">
                    <input type="number" value={o.value} onChange={(e) => update(o.id, { value: parseFloat(e.target.value) || 0 })}
                      className="bg-darb-yellow/10 border border-darb-yellow/30 rounded px-2 py-1 w-16 text-left" dir="ltr" />
                  </td>
                  <td className="td">{statusBadge(o.status)}</td>
                  <td className="td">
                    <input type="checkbox" checked={o.active} disabled={o.status !== "approved"}
                      onChange={(e) => update(o.id, { active: e.target.checked })}
                      className="accent-darb-orange w-4 h-4 disabled:opacity-30" />
                  </td>
                  <td className="td">
                    <div className="flex items-center gap-1">
                      {o.status !== "approved" && (
                        <button onClick={() => update(o.id, { status: "approved", active: true })}
                          className="text-darb-good hover:opacity-70 text-xs font-bold" title="اعتماد">✓</button>
                      )}
                      {o.status !== "rejected" && (
                        <button onClick={() => update(o.id, { status: "rejected", active: false })}
                          className="text-darb-warn hover:opacity-70 text-xs font-bold" title="رفض">✕</button>
                      )}
                      <button onClick={() => remove(o.id)} className="text-darb-bad hover:opacity-70 text-sm" title="حذف">🗑</button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        {filtered.length === 0 && <p className="text-sm text-darb-mut text-center py-4">لا عروض بهذا التصفية.</p>}
      </Card>

      <Card>
        <p className="text-sm text-darb-ink/90 leading-relaxed">
          💡 <b>التحكم بيد درب:</b> التاجر يقترح العرض، و<b>مدير الحساب</b> المسؤول عن مجموعته يراجع القيمة
          والجدوى ثم يعتمده ويفعّله — فلا يظهر للعميل إلا ما اعتمدته درب. هذا يضمن جودة العروض وحماية الهوامش
          وتوحيد قيمة النقطة.
        </p>
      </Card>
    </div>
  );
}
