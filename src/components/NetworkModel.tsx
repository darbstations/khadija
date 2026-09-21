import { Card, Badge } from "./ui";

const COMPARE = [
  ["ملكية العلاقة", "درب مباشرة", "Dsquares"],
  ["الشروط والنسب", "درب تتفاوض (100/70-30/50-50/30-70)", "Dsquares تحدّد + عمولة"],
  ["البيانات", "ملك درب", "تمرّ عبر Dsquares"],
  ["التمويل", "كل تاجر يموّل مبيعاته هو", "حسب برنامج العميل"],
  ["السرعة/الانتشار", "تدريجي (جهد درب)", "فوري (+21,000 تاجر)"],
  ["التحكّم", "كامل", "أقل"],
];

export default function NetworkModel() {
  return (
    <div className="space-y-4">
      <div>
        <h2 className="text-lg font-extrabold">🕸️ شبكة التجار · درب مقابل Dsquares</h2>
        <p className="text-xs text-darb-mut">من يملك العلاقة والبيانات — ومن يموّل ماذا</p>
      </div>

      {/* محفظة واحدة */}
      <Card title="👛 محفظة واحدة · العميل يكسب من كل تجارك">
        <div className="flex flex-wrap items-center justify-center gap-3 text-center">
          {["⛽ بنزين", "☕ كافيه", "🍔 مطعم", "🚿 مغسلة", "🛒 بقالة"].map((s) => (
            <div key={s} className="stat px-4 py-2">{s}</div>
          ))}
          <div className="text-2xl text-darb-orange">←</div>
          <div className="rounded-xl bg-darb-orange/15 border border-darb-orange/40 px-5 py-3 font-extrabold text-darb-orange">
            💎 رصيد نقاط واحد
          </div>
        </div>
        <p className="text-xs text-darb-mut mt-3 text-center">
          كل تاجر يموّل نقاط مبيعاته هو فقط — فكسب العميل من عدة تجار <b>لا يخسّر درب</b> (نموذج المحفظة).
        </p>
      </Card>

      {/* المقارنة */}
      <Card title="⚖️ تجار درب المباشرون مقابل شبكة Dsquares">
        <div className="overflow-x-auto">
          <table className="w-full border-collapse">
            <thead>
              <tr>
                <th className="th">البند</th>
                <th className="th text-darb-orange">🟠 تجار تجيبهم درب</th>
                <th className="th">🔵 تجار عبر Dsquares</th>
              </tr>
            </thead>
            <tbody>
              {COMPARE.map((r) => (
                <tr key={r[0]}>
                  <td className="td font-bold">{r[0]}</td>
                  <td className="td text-darb-ink">{r[1]}</td>
                  <td className="td text-darb-mut">{r[2]}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Card>

      {/* وضع التاجر */}
      <div className="grid md:grid-cols-2 gap-3">
        <Card title="🟠 وضع التاجر اللي تجيبه درب">
          <ul className="text-sm space-y-1.5 text-darb-ink/90 list-disc pr-5">
            <li>يوقّع مع درب مباشرة ويصير ضمن تانكي.</li>
            <li>يموّل مكافآت عملائه من هامشه.</li>
            <li>يحصل: traffic + بيانات + تسويق داخل التطبيق.</li>
            <li>مدير حساب من درب يديره ويعتمد عروضه.</li>
            <li>قيمة النقطة موحّدة (تحكم درب) — لا شتات.</li>
          </ul>
          <div className="mt-2"><Badge tone="good">تحكّم + بيانات + هامش أفضل</Badge></div>
        </Card>
        <Card title="🔵 وضع التاجر عبر Dsquares">
          <ul className="text-sm space-y-1.5 text-darb-ink/90 list-disc pr-5">
            <li>ضمن شبكة Dsquares (21,000+ تاجر).</li>
            <li>Dsquares تملك العلاقة وتأخذ عمولة.</li>
            <li>انتشار فوري وخيارات استبدال أوسع.</li>
            <li>تحكّم وبيانات أقل لدرب + اعتمادية.</li>
          </ul>
          <div className="mt-2"><Badge tone="accent">سرعة + انتشار فوري</Badge></div>
        </Card>
      </div>

      {/* الاستبدال */}
      <Card title="🔁 الاستبدال عند العميل · من أين ومن يموّله">
        <div className="overflow-x-auto">
          <table className="w-full border-collapse">
            <thead>
              <tr>
                <th className="th">القناة</th>
                <th className="th">أين</th>
                <th className="th">من يموّل</th>
                <th className="th">تكلفة درب</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td className="td font-bold">🛍️ عروض تجار درب</td>
                <td className="td">داخل المحطة</td>
                <td className="td">التاجر (من محفظة العميل)</td>
                <td className="td"><Badge tone="good">≈ صفر</Badge></td>
              </tr>
              <tr>
                <td className="td font-bold">⛽ خصم بنزين</td>
                <td className="td">المحطة</td>
                <td className="td">درب</td>
                <td className="td"><Badge tone="bad">مرتفعة</Badge></td>
              </tr>
              <tr>
                <td className="td font-bold">🔵 شبكة Dsquares</td>
                <td className="td">+21,000 تاجر خارجي</td>
                <td className="td">تسوية عبر Dsquares</td>
                <td className="td"><Badge tone="warn">عمولة</Badge></td>
              </tr>
              <tr>
                <td className="td font-bold">🎁 خارجي / كرت شحن</td>
                <td className="td">جرير/أمازون/اتصالات</td>
                <td className="td">درب (بخصم جملة)</td>
                <td className="td"><Badge tone="warn">95–97%</Badge></td>
              </tr>
            </tbody>
          </table>
        </div>
        <p className="text-xs text-darb-mut mt-3 leading-relaxed">
          العميل يستبدل من <b>رصيد واحد</b> في أي قناة. الاستبدال <b>≤ رصيد المحفظة</b> (لا انكشاف ولا خسارة)،
          <b> لا استبدال نقدي</b>، وكل عملية تُعتمد عبر البوابة. وجّهي الاستبدال نحو <b>عروض تجار درب</b> (تكلفتها ≈ صفر).
        </p>
      </Card>

      {/* التوصية */}
      <div className="rounded-2xl bg-darb-orange/10 border border-darb-orange/40 p-5">
        <h3 className="font-extrabold text-darb-orange mb-2">🎯 التوصية · نموذج هجين</h3>
        <ul className="text-sm space-y-1.5 text-darb-ink/90">
          <li><b>داخل محطاتك</b> → جيبي التجار بنفسك (تحكّم + بيانات + هامش أفضل) عبر أداة المستأجرين.</li>
          <li><b>خارج محطاتك</b> → عبر شبكة Dsquares (توسيع خيارات الاستبدال بسرعة).</li>
          <li>النتيجة: <b>أنتِ تملكين النواة</b> (محطاتك + تانكي + بياناتك)، وDsquares تضيف العمق الخارجي دون تسليم برنامجك.</li>
        </ul>
      </div>
    </div>
  );
}
