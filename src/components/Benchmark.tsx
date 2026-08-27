import { useScenario } from "../context/ScenarioContext";
import { tiers, pointMath, fmtPct, fmtInt } from "../model/engine";
import { BENCHMARKS, BENCHMARK_SOURCES } from "../model/defaults";
import { Card, Badge } from "./ui";

export default function Benchmark() {
  const { inputs } = useScenario();
  const t = tiers(inputs);
  const m = pointMath(inputs);

  return (
    <div className="space-y-4">
      <div>
        <h2 className="text-lg font-extrabold">🌍 معايير المنافسين · برامج ولاء الوقود</h2>
        <p className="text-xs text-darb-mut">كيف تحسب الشركات العالمية والإقليمية نقاطها — مقارنة بدرب</p>
      </div>

      {/* طرق الحساب الثلاث */}
      <div className="grid sm:grid-cols-3 gap-3">
        <Card title="① لكل لتر">
          <p className="text-sm text-darb-ink/90">مرتبطة بالكمية لا السعر — Shell · Emarat.</p>
        </Card>
        <Card title="② لكل وحدة عملة">
          <p className="text-sm text-darb-ink/90">نقاط لكل ريال/درهم — ADNOC · <b className="text-darb-orange">درب</b>.</p>
        </Card>
        <Card title="③ قائمة على القيمة">
          <p className="text-sm text-darb-ink/90">عائد صريح كنسبة من السعر — TotalEnergies.</p>
        </Card>
      </div>

      <Card title="📊 جدول المقارنة">
        <div className="overflow-x-auto">
          <table className="w-full border-collapse">
            <thead>
              <tr>
                <th className="th">البرنامج</th>
                <th className="th">المنطقة</th>
                <th className="th">طريقة الحساب</th>
                <th className="th">الكسب</th>
                <th className="th">الاستبدال</th>
                <th className="th">الفعلي</th>
              </tr>
            </thead>
            <tbody>
              {BENCHMARKS.map((b) => (
                <tr key={b.name}>
                  <td className="td font-bold">{b.name}</td>
                  <td className="td text-xs">{b.region}</td>
                  <td className="td text-xs">{b.method}</td>
                  <td className="td text-xs">{b.earn}</td>
                  <td className="td text-xs">{b.redeem}</td>
                  <td className="td">{b.eff}</td>
                </tr>
              ))}
              {/* صف درب الحيّ */}
              <tr className="bg-darb-orange/10 font-bold">
                <td className="td text-darb-orange">🟠 درب · تانكي</td>
                <td className="td text-xs">السعودية</td>
                <td className="td text-xs">متدرّج · نقاط/ريال</td>
                <td className="td text-xs">
                  {t[0].earnRate} / {t[1].earnRate} / {t[2].earnRate} نقطة/ريال
                </td>
                <td className="td text-xs">{fmtInt(inputs.pointValue)} نقطة = ريال</td>
                <td className="td">
                  {fmtPct(t[0].cashbackPct, 1)}–{fmtPct(t[2].cashbackPct, 0)}
                </td>
              </tr>
            </tbody>
          </table>
        </div>
      </Card>

      <Card title="🎯 موقع درب من المعيار">
        <ul className="text-sm space-y-2 text-darb-ink/90 list-disc pr-5">
          <li>
            <b>Emarat</b> يستخدم نفس فكرتك بالضبط: مستويات (برونزي/فضي/ذهبي = أبيض/فضي/برتقالي) بكسب متصاعد.
          </li>
          <li>
            <b>ADNOC</b> الأكرم ظاهرياً (3 نقاط/درهم + بونص) — نفس حيلة الأرقام الكبيرة، لكن قيمة نقطته غير معلنة.
          </li>
          <li>
            <b>TotalEnergies</b> الأصدق: يعلن القيمة صراحة (15 سنت/لتر) — وهو المنهج الذي اخترناه لدرب (قيمة واضحة بلا تضخيم).
          </li>
          <li>
            درب عند <b>{fmtPct(t[0].cashbackPct, 1)}–{fmtPct(t[2].cashbackPct, 0)}</b> متوافق مع المعيار العالمي للوقود
            (هوامش رفيعة) — والكرم الإضافي ({fmtInt(t[0].earnRate / m.sascoPointsPerRiyal)}x ساسكو) + مكافآت الشركاء.
          </li>
        </ul>
        <div className="mt-3 flex flex-wrap gap-2">
          <Badge tone="good">متوافق مع المعيار</Badge>
          <Badge tone="accent">قيمة واضحة بلا تضخيم</Badge>
        </div>
      </Card>

      <p className="text-[11px] text-darb-mut">
        المصادر:{" "}
        {BENCHMARK_SOURCES.map((s, i) => (
          <span key={s.url}>
            <a href={s.url} target="_blank" rel="noreferrer" className="text-darb-accent hover:underline">
              {s.label}
            </a>
            {i < BENCHMARK_SOURCES.length - 1 ? " · " : ""}
          </span>
        ))}
      </p>
    </div>
  );
}
