import { useState } from "react";
import { ScenarioProvider } from "./context/ScenarioContext";
import { DarbLogo } from "./components/DarbLogo";
import ExecutivePresentation from "./components/ExecutivePresentation";
import ScenarioPanel from "./components/ScenarioPanel";
import ExecutiveSummary from "./components/ExecutiveSummary";
import CustomerSimulator from "./components/CustomerSimulator";
import CustomerJourney from "./components/CustomerJourney";
import CustomerFlow from "./components/CustomerFlow";
import Bundles from "./components/Bundles";
import EarnRules from "./components/EarnRules";
import TiersAndValue from "./components/TiersAndValue";
import PointValueModel from "./components/PointValueModel";
import StationCost from "./components/StationCost";
import StationCalculator from "./components/StationCalculator";
import StationsDB from "./components/StationsDB";
import Redemption from "./components/Redemption";
import WalletRedemption from "./components/WalletRedemption";
import OffersMarketplace from "./components/OffersMarketplace";
import OffersAdmin from "./components/OffersAdmin";
import NetworkModel from "./components/NetworkModel";
import FranchiseCollection from "./components/FranchiseCollection";
import Projections from "./components/Projections";
import ScenarioCompare from "./components/ScenarioCompare";
import Benchmark from "./components/Benchmark";
import ExtendedEconomics from "./components/ExtendedEconomics";
import ROICalculator from "./components/ROICalculator";
import TenantTool from "./components/TenantTool";
import MerchantPitch from "./components/MerchantPitch";
import BillToPoints from "./components/BillToPoints";
import HowItWorks from "./components/HowItWorks";

const GROUPS = [
  { key: "overview", label: "نظرة عامة", icon: "📊" },
  { key: "customer", label: "العميل", icon: "🙋" },
  { key: "merchants", label: "التجار", icon: "🤝" },
  { key: "stations", label: "المحطات", icon: "⛽" },
  { key: "financial", label: "المالية", icon: "💰" },
];

const TABS = [
  // نظرة عامة
  { id: "presentation", group: "overview", label: "العرض التنفيذي", icon: "📑", el: <ExecutivePresentation /> },
  { id: "summary", group: "overview", label: "الملخص التنفيذي", icon: "📊", el: <ExecutiveSummary /> },
  { id: "scenarios", group: "overview", label: "السيناريوهات", icon: "⚙️", el: <ScenarioPanel /> },
  { id: "how", group: "overview", label: "كيف يعمل + التوسع", icon: "📚", el: <HowItWorks /> },
  // العميل
  { id: "simulator", group: "customer", label: "محاكي العميل", icon: "🎮", el: <CustomerSimulator /> },
  { id: "journey", group: "customer", label: "رحلة العميل", icon: "🧭", el: <CustomerJourney /> },
  { id: "flow", group: "customer", label: "محاكي الرحلة المتكامل", icon: "🎬", el: <CustomerFlow /> },
  { id: "tiers", group: "customer", label: "المستويات وقيمة النقطة", icon: "🎯", el: <TiersAndValue /> },
  { id: "pointvalue", group: "customer", label: "نموذج قيمة النقطة والاستبدال", icon: "💠", el: <PointValueModel /> },
  { id: "earnrules", group: "customer", label: "قواعد كسب النقاط", icon: "📐", el: <EarnRules /> },
  { id: "offers", group: "customer", label: "سوق العروض", icon: "🎁", el: <OffersMarketplace /> },
  { id: "bundles", group: "customer", label: "باقات درب", icon: "🪪", el: <Bundles /> },
  { id: "redemption", group: "customer", label: "قنوات الاستبدال", icon: "🔁", el: <Redemption /> },
  { id: "wallet", group: "customer", label: "المحفظة والاستبدال", icon: "👛", el: <WalletRedemption /> },
  // التجار
  { id: "tenants", group: "merchants", label: "أداة المستأجرين", icon: "🤝", el: <TenantTool /> },
  { id: "pitch", group: "merchants", label: "حاسبة إقناع التاجر", icon: "🧮", el: <MerchantPitch /> },
  { id: "billpoints", group: "merchants", label: "الفاتورة → النقاط والتحصيل", icon: "🧾", el: <BillToPoints /> },
  { id: "offersadmin", group: "merchants", label: "إدارة العروض", icon: "🛠️", el: <OffersAdmin /> },
  { id: "network", group: "merchants", label: "شبكة التجار (درب/Dsquares)", icon: "🕸️", el: <NetworkModel /> },
  // المحطات
  { id: "stations", group: "stations", label: "تكلفة المحطات", icon: "🏗️", el: <StationCost /> },
  { id: "stationcalc", group: "stations", label: "حاسبة النقطة لكل محطة", icon: "🧮", el: <StationCalculator /> },
  { id: "mystations", group: "stations", label: "محطاتي", icon: "⛽", el: <StationsDB /> },
  { id: "franchise", group: "stations", label: "التحصيل من الامتياز", icon: "🧾", el: <FranchiseCollection /> },
  // المالية
  { id: "roi", group: "financial", label: "حاسبة ROI", icon: "💎", el: <ROICalculator /> },
  { id: "projections", group: "financial", label: "التوقعات السنوية", icon: "📈", el: <Projections /> },
  { id: "extended", group: "financial", label: "الاقتصاد الموسّع", icon: "🆕", el: <ExtendedEconomics /> },
  { id: "compare", group: "financial", label: "مقارنة السيناريوهات", icon: "🆚", el: <ScenarioCompare /> },
  { id: "benchmark", group: "financial", label: "معايير المنافسين", icon: "🌍", el: <Benchmark /> },
];

export default function App() {
  const [group, setGroup] = useState("overview");
  const [active, setActive] = useState("presentation");
  const current = TABS.find((t) => t.id === active) ?? TABS[0];
  const tabsInGroup = TABS.filter((t) => t.group === group);

  const selectGroup = (g: string) => {
    setGroup(g);
    const first = TABS.find((t) => t.group === g);
    if (first) setActive(first.id);
  };

  return (
    <ScenarioProvider>
      <div className="min-h-screen flex flex-col">
        <header className="sticky top-0 z-20 backdrop-blur bg-darb-bg/80 border-b border-darb-line">
          <div className="max-w-7xl mx-auto px-4 py-3 flex items-center justify-between gap-3">
            <div className="flex items-center gap-3">
              <DarbLogo className="h-10 w-auto" />
              <div className="border-r border-darb-line pr-3">
                <h1 className="text-base font-extrabold leading-none text-darb-orange">تانكي</h1>
                <p className="text-[11px] text-darb-mut mt-0.5">النموذج المالي التفاعلي · 2026</p>
              </div>
            </div>
            <span className="text-[11px] text-darb-mut hidden sm:block">
              كل الأرقام حيّة · عدّل السيناريوهات وشاهد التحديث
            </span>
          </div>

          {/* المجموعات */}
          <nav className="max-w-7xl mx-auto px-2 flex gap-1 overflow-x-auto border-b border-darb-line/60">
            {GROUPS.map((g) => (
              <button
                key={g.key}
                onClick={() => selectGroup(g.key)}
                className={`whitespace-nowrap px-3.5 py-2 text-sm font-extrabold transition border-b-2 ${
                  group === g.key
                    ? "border-darb-orange text-darb-orange"
                    : "border-transparent text-darb-mut hover:text-darb-ink"
                }`}
              >
                <span className="ml-1">{g.icon}</span>
                {g.label}
              </button>
            ))}
          </nav>

          {/* تبويبات المجموعة */}
          <nav className="max-w-7xl mx-auto px-2 py-2 flex gap-1 overflow-x-auto">
            {tabsInGroup.map((t) => (
              <button
                key={t.id}
                onClick={() => setActive(t.id)}
                className={`whitespace-nowrap px-3 py-1.5 rounded-lg text-xs font-bold transition ${
                  active === t.id
                    ? "bg-darb-orange/15 text-darb-orange"
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
