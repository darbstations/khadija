import { useState, useEffect } from "react";
import { tenantImpact, fmtSar, fmtPct, fmtNum } from "../model/engine";
import {
  NEGOTIATION_MODELS,
  TENANT_STATUSES,
  NEGOTIATION_SCRIPT,
  OBJECTIONS,
} from "../model/defaults";
import { Card, Stat, NumberInput } from "./ui";

interface TenantRow {
  id: number;
  name: string;
  type: string;
  margin: number;
  model: string;
  earnPct: number; // Earn · النسبة عند الشراء
  burnPct: number; // Burn · رسوم التسويق عند الاستبدال
  appOffers: string; // عروض خاصة لعملاء التطبيق
  redemptionPricing: string; // قيمة المنتجات بالنقاط عند الاستبدال
  status: string;
}

const STORAGE = "tanki.tenants.v3";
const stars = (n: number) => "⭐".repeat(n);
const modelByKey = (k: string) =>
  NEGOTIATION_MODELS.find((m) => m.key === k) ?? NEGOTIATION_MODELS[2];

export default function TenantTool() {
  const [tenants, setTenants] = useState<TenantRow[]>(() => {
    try {
      const s = localStorage.getItem(STORAGE);
      if (s) return JSON.parse(s);
    } catch {
      /* تجاهل */
    }
    return []; // تبدأ فاضية · إدخال يدوي
  });

  useEffect(() => {
    try {
      localStorage.setItem(STORAGE, JSON.stringify(tenants));
    } catch {
      /* تجاهل */
    }
  }, [tenants]);

  // الحاسبة
  const [name, setName] = useState("سلسلة كافيه افتراضية");
  const [monthly, setMonthly] = useState(150000);
  const [margin, setMargin] = useState(0.65);
  const [modelKey, setModelKey] = useState("split5050");
  const [uplift, setUplift] = useState(0.15);
  const model = modelByKey(modelKey);
  const contribution = model.tenantHalala;
  const r = tenantImpact(monthly, margin, contribution, uplift);

  const loadTenant = (t: TenantRow) => {
    setName(t.name);
    setMargin(t.margin);
    setModelKey(t.model);
  };

  const updateTenant = (id: number, patch: Partial<TenantRow>) =>
    setTenants((prev) => prev.map((t) => (t.id === id ? { ...t, ...patch } : t)));
  const addTenant = () =>
    setTenants((prev) => [
      ...prev,
      {
        id: Math.max(0, ...prev.map((t) => t.id)) + 1,
        name: "",
        type: "",
        margin: 0.5,
        model: "split5050",
        earnPct: 3,
        burnPct: 5,
        appOffers: "",
        redemptionPricing: "",
        status: "لم نتواصل",
      },
    ]);
  const removeTenant = (id: number) =>
    setTenants((prev) => prev.filter((t) => t.id !== id));

  return (
    <div className="space-y-4">
      <div>
        <h2 className="text-lg font-extrabold">🤝 أداة المستأجرين · التفاوض والمتابعة</h2>
        <p className="text-xs text-darb-mut">
          قاعدة بيانات + نماذج التفاوض + حاسبة الأثر المالي + سكربتات الإقناع
        </p>
      </div>

      {/* الحاسبة */}
      <Card title="💰 حاسبة الأثر المالي لكل مستأجر" subtitle="عدّل القيم أو اضغط مستأجراً من الجدول أدناه">
        <div className="grid lg:grid-cols-2 gap-4">
          <div className="space-y-3">
            <label className="block">
              <span className="text-xs font-bold text-darb-ink">اسم المستأجر</span>
              <input
                value={name}
                onChange={(e) => setName(e.target.value)}
                className="w-full mt-1 bg-darb-panel border border-darb-line rounded-lg px-3 py-2 text-darb-ink"
              />
            </label>
            <NumberInput
              label="الإيرادات الشهرية المقدّرة (ريال)"
              value={monthly}
              onChange={setMonthly}
              step={10000}
            />
            <NumberInput
              label="هامش الربح المقدّر (%)"
              value={+(margin * 100).toFixed(0)}
              onChange={(v) => setMargin(v / 100)}
              step={5}
              suffix="%"
            />
            <NumberInput
              label="الزيادة المتوقعة في المبيعات (%)"
              value={+(uplift * 100).toFixed(0)}
              onChange={(v) => setUplift(v / 100)}
              step={5}
              suffix="%"
              hint="تحفّظياً 10-20%"
            />
            <div>
              <span className="text-xs font-bold text-darb-ink">نموذج التفاوض</span>
              <div className="flex flex-wrap gap-1.5 mt-1.5">
                {NEGOTIATION_MODELS.map((m) => (
                  <button
                    key={m.key}
                    onClick={() => setModelKey(m.key)}
                    className={`text-xs font-bold px-2.5 py-1.5 rounded-lg border transition ${
                      modelKey === m.key
                        ? "border-darb-accent bg-darb-accent/15 text-darb-accent"
                        : "border-darb-line text-darb-mut hover:text-darb-ink"
                    }`}
                  >
                    {m.label}
                  </button>
                ))}
              </div>
              <p className="text-[11px] text-darb-mut mt-1.5">
                مساهمة المستأجر: <b>{fmtNum(contribution, 1)} هللة/ريال</b> · {model.fit}
              </p>
            </div>
          </div>

          <div className="grid grid-cols-2 gap-3 content-start">
            <Stat label="📅 الإيرادات السنوية" value={fmtSar(r.annualRevenue)} />
            <Stat label="💰 الربح السنوي" value={fmtSar(r.annualProfit)} />
            <Stat label="💸 تكلفة النقاط/سنة" value={fmtSar(r.pointsCost)} tone="warn" />
            <Stat label="📊 % من ربح المستأجر" value={fmtPct(r.pctOfProfit)} />
            <Stat label="📈 الزيادة المتوقعة" value={fmtSar(r.upliftRevenue)} />
            <Stat label="💎 الربح الإضافي" value={fmtSar(r.extraProfit)} tone="accent" />
            <Stat label="✅ صافي العائد" value={fmtSar(r.netReturn)} tone={r.netReturn > 0 ? "good" : "bad"} />
            <Stat label="🎯 ROI للمستأجر" value={fmtPct(r.roi, 0)} tone={r.roi > 1 ? "good" : "warn"} />
          </div>
        </div>
        <p className="text-xs text-darb-mut mt-3">
          💡 ROI &gt; 100% يعني المستأجر سيوافق بثقة — الربح الإضافي من زيادة المبيعات يفوق تكلفة النقاط بأضعاف.
        </p>
      </Card>

      {/* قاعدة البيانات */}
      <Card title="📋 قاعدة بيانات المستأجرين المحتملين" subtitle="قابلة للتعديل · تُحفظ في المتصفح">
        <div className="overflow-x-auto">
          <table className="w-full border-collapse">
            <thead>
              <tr>
                <th className="th">المستأجر</th>
                <th className="th">النوع</th>
                <th className="th">الهامش</th>
                <th className="th">النموذج</th>
                <th className="th">🟢 Earn %</th>
                <th className="th">🔴 Burn %</th>
                <th className="th">عروض خاصة لعملاء التطبيق</th>
                <th className="th">الاستبدال · قيمة المنتجات بالنقاط</th>
                <th className="th">الحالة</th>
                <th className="th"></th>
              </tr>
            </thead>
            <tbody>
              {tenants.map((t) => (
                <tr key={t.id} className="hover:bg-darb-panel/40 align-top">
                  <td className="td">
                    <div className="flex items-center gap-1">
                      <input value={t.name} onChange={(e) => updateTenant(t.id, { name: e.target.value })} placeholder="اسم التاجر"
                        className="bg-transparent border-b border-darb-line/50 focus:border-darb-accent outline-none w-28 text-darb-ink font-bold" />
                      <button onClick={() => loadTenant(t)} title="حمّل في الحاسبة أعلاه" className="text-darb-accent hover:opacity-70 text-xs shrink-0">🧮</button>
                    </div>
                  </td>
                  <td className="td">
                    <input value={t.type} onChange={(e) => updateTenant(t.id, { type: e.target.value })} placeholder="كافيه…"
                      className="bg-transparent border-b border-darb-line/50 focus:border-darb-accent outline-none w-20 text-xs text-darb-ink" />
                  </td>
                  <td className="td whitespace-nowrap">
                    <input type="number" value={Math.round(t.margin * 100)} onChange={(e) => updateTenant(t.id, { margin: (parseFloat(e.target.value) || 0) / 100 })}
                      className="bg-darb-panel border border-darb-line rounded px-2 py-1 w-14 text-left" dir="ltr" />%
                  </td>
                  <td className="td">
                    <select value={t.model} onChange={(e) => updateTenant(t.id, { model: e.target.value })}
                      className="bg-darb-panel border border-darb-line rounded px-1.5 py-1 text-xs">
                      {NEGOTIATION_MODELS.map((m) => (<option key={m.key} value={m.key}>{m.label}</option>))}
                    </select>
                  </td>
                  <td className="td">
                    <input type="number" step={0.5} value={t.earnPct} onChange={(e) => updateTenant(t.id, { earnPct: parseFloat(e.target.value) || 0 })}
                      className="bg-darb-yellow/10 border border-darb-yellow/30 rounded px-2 py-1 w-14 text-left" dir="ltr" />
                  </td>
                  <td className="td">
                    <input type="number" step={0.5} value={t.burnPct} onChange={(e) => updateTenant(t.id, { burnPct: parseFloat(e.target.value) || 0 })}
                      className="bg-darb-yellow/10 border border-darb-yellow/30 rounded px-2 py-1 w-14 text-left" dir="ltr" />
                  </td>
                  <td className="td">
                    <textarea value={t.appOffers} onChange={(e) => updateTenant(t.id, { appOffers: e.target.value })} rows={2}
                      placeholder="مثال: قهوة ثانية مجاناً · خصم 20% على الوجبات لعملاء التطبيق"
                      className="bg-transparent border border-darb-line/60 rounded px-2 py-1 w-56 text-xs text-darb-ink resize-y align-top" />
                  </td>
                  <td className="td">
                    <textarea value={t.redemptionPricing} onChange={(e) => updateTenant(t.id, { redemptionPricing: e.target.value })} rows={2}
                      placeholder="مثال: قهوة = 1,600 نقطة · وجبة = 3,000 نقطة"
                      className="bg-transparent border border-darb-line/60 rounded px-2 py-1 w-56 text-xs text-darb-ink resize-y align-top" />
                  </td>
                  <td className="td">
                    <select value={t.status} onChange={(e) => updateTenant(t.id, { status: e.target.value })}
                      className="bg-darb-panel border border-darb-line rounded px-1.5 py-1 text-xs">
                      {TENANT_STATUSES.map((s) => (<option key={s} value={s}>{s}</option>))}
                    </select>
                  </td>
                  <td className="td">
                    <button onClick={() => removeTenant(t.id)} className="text-darb-bad hover:opacity-70 text-sm" title="حذف">✕</button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <button
          onClick={addTenant}
          className="mt-3 text-xs font-bold px-3 py-1.5 rounded-lg border border-darb-line hover:border-darb-accent text-darb-mut hover:text-darb-accent transition"
        >
          + إضافة مستأجر
        </button>
      </Card>

      {/* نماذج التفاوض */}
      <Card title="🤝 نماذج التفاوض الأربعة">
        <div className="overflow-x-auto">
          <table className="w-full border-collapse">
            <thead>
              <tr>
                <th className="th">النموذج</th>
                <th className="th">المستأجر</th>
                <th className="th">درب</th>
                <th className="th">مناسب لمن؟</th>
                <th className="th">حجة التفاوض</th>
                <th className="th">التقييم</th>
              </tr>
            </thead>
            <tbody>
              {NEGOTIATION_MODELS.map((m) => (
                <tr key={m.key} className={modelKey === m.key ? "bg-darb-accent/10" : ""}>
                  <td className="td font-bold">{m.label}</td>
                  <td className="td">{fmtNum(m.tenantHalala, 1)} هللة</td>
                  <td className="td">{fmtNum(m.darbHalala, 1)} هللة</td>
                  <td className="td text-xs">{m.fit}</td>
                  <td className="td text-xs max-w-xs">{m.argument}</td>
                  <td className="td whitespace-nowrap">{stars(m.rating)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Card>

      {/* سكربت التفاوض */}
      <Card title="🗣️ سكربت محادثة التفاوض · جاهز للاستخدام">
        <div className="space-y-3">
          {NEGOTIATION_SCRIPT.map((s, idx) => (
            <div key={idx} className="stat">
              <div className="flex items-center justify-between mb-1">
                <span className="font-bold text-darb-ink">
                  {idx + 1}. {s.stage}
                </span>
                <span className="text-[11px] text-darb-mut">
                  {s.min} · {s.goal}
                </span>
              </div>
              <p className="text-sm text-darb-ink/90 leading-relaxed">{s.script}</p>
              <p className="text-[11px] text-darb-accent mt-1.5">💡 {s.tip}</p>
            </div>
          ))}
        </div>
      </Card>

      {/* الاعتراضات */}
      <Card title="⚠️ الاعتراضات الشائعة وكيف تردّ">
        <div className="overflow-x-auto">
          <table className="w-full border-collapse">
            <thead>
              <tr>
                <th className="th">🔴 الاعتراض</th>
                <th className="th">ردك الذكي</th>
                <th className="th">دعم بالأرقام</th>
              </tr>
            </thead>
            <tbody>
              {OBJECTIONS.map((o, idx) => (
                <tr key={idx}>
                  <td className="td font-bold">«{o.q}»</td>
                  <td className="td">{o.a}</td>
                  <td className="td text-xs text-darb-accent">{o.data}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Card>
    </div>
  );
}
