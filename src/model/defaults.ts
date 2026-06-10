import type { ScenarioInputs } from "./types";

/** القيم الموصى بها (نسخة v5) — مبنية على أسعار يونيو 2026 وهوامش درب الفعلية */
export const DEFAULT_INPUTS: ScenarioInputs = {
  // 1
  pointValue: 10000,
  earnWhite: 50, // 0.5%
  earnSilver: 75, // 0.75%
  earnOrange: 100, // 1.0%
  costPerRiyalHalala: 1,
  literPrice: 2.33, // بنزين 95 — يونيو 2026
  welcomeGift: 50000,
  // 2
  silverThreshold: 3600,
  orangeThreshold: 14400,
  spendPeriodMonths: 12,
  graceDays: 30,
  // 3
  stationsOperating: 30,
  stationsInvestment: 70,
  stationsFranchise: 73,
  marginOperating: 13, // هللة/لتر (هامش درب الفعلي)
  marginInvestment: 13,
  marginFranchise: 1.5,
  darbShareFranchise: 0.5,
  stationMonthlyLiters: 80000,
  // 4 — الاقتصاد الموسّع
  dieselPrice: 1.79, // الديزل — يناير 2026
  dieselMarginHalala: 4,
  earnDiesel: 25, // 0.25%
  earnPartner: 400, // ~4% ممولة من الشريك
  partnerContributionHalala: 1, // الشريك يدفع 1 هللة/ريال لدرب
};

/** نمو المحطات عبر 5 سنوات (ورقة التوقعات السنوية) — قابل للاشتقاق من السنة الأولى */
export const PROJECTION_YEARS = [2026, 2027, 2028, 2029, 2030];
export const STATION_GROWTH = {
  operating: [30, 35, 42, 50, 58],
  investment: [70, 80, 88, 95, 100],
  franchise: [73, 90, 105, 120, 135],
};

/** عيّنة 10 محطات امتياز (ورقة التحصيل من الامتياز) */
export const FRANCHISE_SAMPLE = [
  { code: "AMT-001", city: "الرياض", liters: 95000 },
  { code: "AMT-002", city: "جدة", liters: 82000 },
  { code: "AMT-003", city: "مكة", liters: 78000 },
  { code: "AMT-004", city: "الدمام", liters: 88000 },
  { code: "AMT-005", city: "الطائف", liters: 65000 },
  { code: "AMT-006", city: "تبوك", liters: 58000 },
  { code: "AMT-007", city: "حائل", liters: 52000 },
  { code: "AMT-008", city: "بريدة", liters: 71000 },
  { code: "AMT-009", city: "أبها", liters: 68000 },
  { code: "AMT-010", city: "جازان", liters: 49000 },
];

/** عملاء نموذجيون (محاكي العميل) */
export const CUSTOMER_SAMPLES = [
  { label: "🚗 سيارة اقتصادية", monthly: 600, note: "كورولا/سنتافي · استخدام يومي" },
  { label: "🚙 سيارة عادية", monthly: 1000, note: "سيارة متوسطة" },
  { label: "🚕 SUV", monthly: 1500, note: "تاهو/باترول صغير" },
  { label: "🚐 عائلة بسيارتين", monthly: 2500, note: "عائلة نشطة" },
  { label: "🏢 عائلة كبيرة", monthly: 4000, note: "3-4 سيارات" },
  { label: "🚚 أسطول صغير", monthly: 8000, note: "شركة · 4-5 سيارات" },
];

/** سيناريوهات ثابتة للمقارنة (ورقة مقارنة السيناريوهات) */
export const COMPARISON_SCENARIOS = [
  { key: "sasco", label: "🔴 ساسكو", pv: 100, w: 0.43, s: 0.43, o: 0.43, cost: 0.43 },
  { key: "conservative", label: "🟢 محافظ", pv: 20000, w: 50, s: 75, o: 100, cost: 0.5 },
  { key: "moderate", label: "🟢 معتدل", pv: 10000, w: 100, s: 150, o: 200, cost: 1 },
  { key: "attractive", label: "🟡 جذاب", pv: 10000, w: 150, s: 200, o: 250, cost: 1.5 },
  { key: "generous", label: "🔴 سخي", pv: 10000, w: 200, s: 250, o: 300, cost: 2 },
];
