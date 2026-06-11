import { useState } from "react";
import { ScenarioProvider } from "./context/ScenarioContext";
import ScenarioPanel from "./components/ScenarioPanel";
import ExecutiveSummary from "./components/ExecutiveSummary";
import CustomerSimulator from "./components/CustomerSimulator";
import TiersAndValue from "./components/TiersAndValue";
import StationCost from "./components/StationCost";
import StationCalculator from "./components/StationCalculator";
import StationsDB from "./components/StationsDB";
import Redemption from "./components/Redemption";
import WalletRedemption from "./components/WalletRedemption";
import FranchiseCollection from "./components/FranchiseCollection";
import Projections from "./components/Projections";
import ScenarioCompare from "./components/ScenarioCompare";
import ExtendedEconomics from "./components/ExtendedEconomics";
import ROICalculator from "./components/ROICalculator";
import TenantTool from "./components/TenantTool";
import HowItWorks from "./components/HowItWorks";

const TABS = [
  { id: "summary", label: "الملخص التنفيذي", icon: "📊", el: <ExecutiveSummary /> },
  { id: "scenarios", label: "السيناريوهات", icon: "⚙️", el: <ScenarioPanel /> },
  { id: "simulator", label: "محاكي العميل", icon: "🎮", el: <CustomerSimulator /> },
  { id: "redemption", label: "قنوات الاستبدال", icon: "🔁", el: <Redemption /> },
  { id: "wallet", label: "المحفظة وسوق الاستبدال", icon: "👛", el: <WalletRedemption /> },
  { id: "tiers", label: "المستويات وقيمة النقطة", icon: "🎯", el: <TiersAndValue /> },
  { id: "extended", label: "الاقتصاد الموسّع", icon: "🆕", el: <ExtendedEconomics /> },
  { id: "stations", label: "تكلفة المحطات", icon: "🏗️", el: <StationCost /> },
  { id: "stationcalc", label: "حاسبة النقطة لكل محطة", icon: "🧮", el: <StationCalculator /> },
  { id: "mystations", label: "محطاتي", icon: "⛽", el: <StationsDB /> },
  { id: "tenants", label: "أداة المستأجرين", icon: "🤝", el: <TenantTool /> },
  { id: "franchise", label: "التحصيل من الامتياز", icon: "🧾", el: <FranchiseCollection /> },
  { id: "projections", label: "التوقعات السنوية", icon: "📈", el: <Projections /> },
  { id: "roi", label: "حاسبة ROI", icon: "💎", el: <ROICalculator /> },
  { id: "compare", label: "مقارنة السيناريوهات", icon: "🆚", el: <ScenarioCompare /> },
  { id: "how", label: "كيف يعمل + التوسع", icon: "📚", el: <HowItWorks /> },
];

export default function App() {
  const [active, setActive] = useState("summary");
  const current = TABS.find((t) => t.id === active) ?? TABS[0];

  return (
    <ScenarioProvider>
      <div className="min-h-screen flex flex-col">
        <header className="sticky top-0 z-20 backdrop-blur bg-darb-bg/80 border-b border-darb-line">
          <div className="max-w-7xl mx-auto px-4 py-3 flex items-center justify-between gap-3">
            <div className="flex items-center gap-3">
              <div className="w-9 h-9 rounded-xl bg-gradient-to-br from-darb-orange to-darb-accent grid place-items-center font-extrabold text-darb-bg">
                ت
              </div>
              <div>
                <h1 className="text-base font-extrabold leading-none">
                  تانكي · درب
                </h1>
                <p className="text-[11px] text-darb-mut mt-0.5">
                  النموذج المالي التفاعلي · 2026
                </p>
              </div>
            </div>
            <span className="text-[11px] text-darb-mut hidden sm:block">
              كل الأرقام حيّة · عدّل السيناريوهات وشاهد التحديث
            </span>
          </div>
          <nav className="max-w-7xl mx-auto px-2 pb-2 flex gap-1 overflow-x-auto">
            {TABS.map((t) => (
              <button
                key={t.id}
                onClick={() => setActive(t.id)}
                className={`whitespace-nowrap px-3 py-1.5 rounded-lg text-xs font-bold transition ${
                  active === t.id
                    ? "bg-darb-accent/15 text-darb-accent"
                    : "text-darb-mut hover:text-darb-ink hover:bg-darb-panel"
                }`}
              >
                <span className="ml-1">{t.icon}</span>
                {t.label}
              </button>
            ))}
          </nav>
        </header>

        <main className="flex-1 max-w-7xl w-full mx-auto px-4 py-5">{current.el}</main>

        <footer className="border-t border-darb-line py-4 text-center text-[11px] text-darb-mut">
          درب · إدارة الشراكات · النموذج المالي · 2026 — تُحفظ مدخلاتك تلقائياً في المتصفح
        </footer>
      </div>
    </ScenarioProvider>
  );
}
