import type {
  ScenarioInputs,
  TierRow,
  TierKey,
  CustomerResult,
  StationEconomics,
} from "./types";
import {
  STATION_GROWTH,
  PROJECTION_YEARS,
  FRANCHISE_SAMPLE,
  COMPARISON_SCENARIOS,
} from "./defaults";

// ============================================================
//  محرك الحساب — يطابق معادلات ملف الإكسل خلية بخلية
// ============================================================

export const TIER_META: Record<TierKey, { label: string; emoji: string }> = {
  white: { label: "الأبيض", emoji: "🤍" },
  silver: { label: "الفضي", emoji: "🩶" },
  orange: { label: "البرتقالي", emoji: "🟠" },
};

/** المستويات الثلاثة (هيكل + كاش باك) */
export function tiers(s: ScenarioInputs): TierRow[] {
  const make = (key: TierKey, earnRate: number, minSpend: number): TierRow => ({
    key,
    label: TIER_META[key].label,
    emoji: TIER_META[key].emoji,
    earnRate,
    minSpend,
    cashbackPct: earnRate / s.pointValue,
  });
  return [
    make("white", s.earnWhite, 0),
    make("silver", s.earnSilver, s.silverThreshold),
    make("orange", s.earnOrange, s.orangeThreshold),
  ];
}

/** قيمة النقطة والمعادلات الأساسية (ورقة قيمة النقطة) */
export function pointMath(s: ScenarioInputs) {
  return {
    valuePerPointRiyal: 1 / s.pointValue, // C6
    valuePerPointHalala: (1 / s.pointValue) * 100, // C7
    sascoPointsPerRiyal: 1 / s.literPrice, // C8 = لتر/ريال
    cashbackWhite: s.earnWhite / s.pointValue,
    cashbackSilver: s.earnSilver / s.pointValue,
    cashbackOrange: s.earnOrange / s.pointValue,
  };
}

/** تحديد مستوى العميل من إنفاقه السنوي */
export function tierForSpend(s: ScenarioInputs, annualSpend: number): TierRow {
  const [white, silver, orange] = tiers(s);
  if (annualSpend < s.silverThreshold) return white;
  if (annualSpend < s.orangeThreshold) return silver;
  return orange;
}

/** محاكاة عميل: من الإنفاق الشهري إلى النقاط/المستوى/التقدم */
export function simulateCustomer(
  s: ScenarioInputs,
  monthlySpend: number
): CustomerResult {
  const annualSpend = monthlySpend * 12;
  const tier = tierForSpend(s, annualSpend);
  const pointsPerYear = annualSpend * tier.earnRate;
  const valueRiyal = pointsPerYear / s.pointValue;
  const cashbackPct = tier.earnRate / s.pointValue;

  let nextTierLabel = "✨ أعلى مستوى";
  let remainingSpend = 0;
  let progress = 1;
  if (tier.key === "white") {
    nextTierLabel = "🩶 الفضي";
    remainingSpend = Math.max(0, s.silverThreshold - annualSpend);
    progress = annualSpend / s.silverThreshold;
  } else if (tier.key === "silver") {
    nextTierLabel = "🟠 البرتقالي";
    remainingSpend = Math.max(0, s.orangeThreshold - annualSpend);
    progress = (annualSpend - s.silverThreshold) / (s.orangeThreshold - s.silverThreshold);
  }
  progress = Math.max(0, Math.min(1, progress));

  // تكلفة درب الفعلية على هذا العميل = الإنفاق × الكاش باك
  const darbCostRiyal = valueRiyal;
  const vsSasco = tier.earnRate / (1 / s.literPrice);

  return {
    monthlySpend,
    annualSpend,
    tier,
    pointsPerYear,
    valueRiyal,
    cashbackPct,
    nextTierLabel,
    remainingSpend,
    progress,
    darbCostRiyal,
    vsSasco,
  };
}

/** اقتصاديات كل نوع محطة (ورقة تكلفة المحطات) */
export function stationEconomics(s: ScenarioInputs): StationEconomics[] {
  const loyalty = s.costPerRiyalHalala; // هللة/ريال
  const mk = (
    key: StationEconomics["key"],
    label: string,
    marginPerLiter: number,
    darbShare: number,
    ownerNet: number
  ): StationEconomics => {
    const feasible: StationEconomics["feasible"] =
      ownerNet > 0.5 ? "good" : ownerNet > 0 ? "tight" : "loss";
    return {
      key,
      label,
      marginPerLiter,
      marginPerRiyal: marginPerLiter / s.literPrice,
      loyaltyCostPerRiyal: loyalty,
      darbSharePct: darbShare,
      darbCostPerRiyal: loyalty * darbShare,
      ownerNetPerRiyal: ownerNet,
      feasible,
    };
  };
  const opMargin = s.marginOperating / s.literPrice;
  const invMargin = s.marginInvestment / s.literPrice;
  const frMargin = s.marginFranchise / s.literPrice;
  return [
    // تشغيل: درب تملكها وتتحمل 100% — صافي الهامش يُعرض إجمالياً
    mk("operating", "محطة تشغيل", s.marginOperating, 1, opMargin),
    // استثمار: درب 50% — المالك يتحمل (1 - حصة درب)
    mk("investment", "محطة استثمار", s.marginInvestment, 0.5, invMargin - loyalty * (1 - 0.5)),
    // امتياز: حصة درب = darbShareFranchise
    mk(
      "franchise",
      "محطة امتياز",
      s.marginFranchise,
      s.darbShareFranchise,
      frMargin - loyalty * (1 - s.darbShareFranchise)
    ),
  ];
}

/** التحصيل من محطات الامتياز (ورقة التحصيل) */
export function franchiseCollection(s: ScenarioInputs) {
  const ownerShare = 1 - s.darbShareFranchise;
  const rows = FRANCHISE_SAMPLE.map((r) => {
    const spend = r.liters * s.literPrice; // E = لتر × سعر
    const ownerPay = (spend * s.costPerRiyalHalala * ownerShare) / 100;
    const darbPay = (spend * s.costPerRiyalHalala * s.darbShareFranchise) / 100;
    return { ...r, spend, ownerPay, darbPay };
  });
  const sum = (f: (x: (typeof rows)[number]) => number) =>
    rows.reduce((a, x) => a + f(x), 0);
  const monthly = {
    liters: sum((x) => x.liters),
    spend: sum((x) => x.spend),
    ownerPay: sum((x) => x.ownerPay),
    darbPay: sum((x) => x.darbPay),
  };
  const annual = {
    liters: monthly.liters * 12,
    spend: monthly.spend * 12,
    ownerPay: monthly.ownerPay * 12,
    darbPay: monthly.darbPay * 12,
  };
  return { rows, monthly, annual };
}

/** التوقعات السنوية: التكلفة على درب لكل نوع محطة (مليون ريال) */
export function projections(s: ScenarioInputs) {
  const factor = (count: number, share: number) =>
    (count *
      s.stationMonthlyLiters *
      12 *
      s.literPrice *
      s.costPerRiyalHalala *
      share) /
    100 /
    1_000_000;

  const operating = STATION_GROWTH.operating.map((c) => factor(c, 1));
  const investment = STATION_GROWTH.investment.map((c) => factor(c, 0.5));
  const franchise = STATION_GROWTH.franchise.map((c) => factor(c, s.darbShareFranchise));
  const total = PROJECTION_YEARS.map(
    (_, i) => operating[i] + investment[i] + franchise[i]
  );
  const totalStations = PROJECTION_YEARS.map(
    (_, i) =>
      STATION_GROWTH.operating[i] +
      STATION_GROWTH.investment[i] +
      STATION_GROWTH.franchise[i]
  );
  return {
    years: PROJECTION_YEARS,
    operating,
    investment,
    franchise,
    total,
    totalStations,
    fiveYearTotal: total.reduce((a, b) => a + b, 0),
  };
}

/** الاقتصاد الموسّع: الديزل والشركاء — مربوط فعلياً بالمدخلات */
export function extendedEconomics(s: ScenarioInputs) {
  // الديزل
  const dieselMarginPct = s.dieselMarginHalala / (s.dieselPrice * 100);
  const dieselCashback = s.earnDiesel / s.pointValue;
  const dieselDarbCostHalala = s.earnDiesel / s.pointValue * 100; // هللة/ريال
  const dieselNetMarginHalala = s.dieselMarginHalala / s.dieselPrice - dieselDarbCostHalala;
  const dieselFeasible =
    dieselNetMarginHalala > 0.5 ? "good" : dieselNetMarginHalala > 0 ? "tight" : "loss";

  // الشركاء
  const partnerCashback = s.earnPartner / s.pointValue; // قيمة النقاط للعميل
  // مكافأة العميل ممولة من الشريك → تكلفة درب ≈ 0، وفوقها درب تكسب المساهمة
  const partnerNetToDarbHalala = s.partnerContributionHalala; // هللة/ريال صافية لدرب
  const partnerCostToDarb = 0;

  // مقارنة: ريال إنفاق على البنزين مقابل ريال عند الشريك (تأثيره على درب)
  const fuelDarbMarginHalala = s.marginOperating / s.literPrice; // ربح
  const fuelLoyaltyCostHalala = (s.earnWhite / s.pointValue) * 100; // تكلفة الولاء

  return {
    diesel: {
      price: s.dieselPrice,
      marginHalala: s.dieselMarginHalala,
      marginPct: dieselMarginPct,
      earn: s.earnDiesel,
      cashback: dieselCashback,
      darbCostHalala: dieselDarbCostHalala,
      netMarginHalala: dieselNetMarginHalala,
      feasible: dieselFeasible as "good" | "tight" | "loss",
    },
    partner: {
      earn: s.earnPartner,
      cashback: partnerCashback,
      contributionHalala: s.partnerContributionHalala,
      netToDarbHalala: partnerNetToDarbHalala,
      costToDarb: partnerCostToDarb,
    },
    fuel: {
      darbMarginHalala: fuelDarbMarginHalala,
      loyaltyCostHalala: fuelLoyaltyCostHalala,
      netHalala: fuelDarbMarginHalala - fuelLoyaltyCostHalala,
    },
  };
}

/** مقارنة السيناريوهات: العمود "الحالي" يعكس المدخلات الحيّة */
export function scenarioComparison(s: ScenarioInputs) {
  const current = {
    key: "current",
    label: "⭐ الحالي",
    pv: s.pointValue,
    w: s.earnWhite,
    s: s.earnSilver,
    o: s.earnOrange,
    cost: s.costPerRiyalHalala,
  };
  const list = [
    COMPARISON_SCENARIOS[0],
    COMPARISON_SCENARIOS[1],
    current,
    COMPARISON_SCENARIOS[2],
    COMPARISON_SCENARIOS[3],
    COMPARISON_SCENARIOS[4],
  ];
  // التكلفة السنوية (مليون ريال) — افتراض 230 محطة × 80,000 لتر × 12 × سعر اللتر = ~514.5م ريال
  const annualFuelRevenueM =
    (230 * 80000 * 12 * s.literPrice) / 1_000_000; // ≈ 514.5
  return list.map((sc) => {
    const cashbackWhite = sc.pv > 1000 ? sc.w / sc.pv : sc.w / 100; // ساسكو نسبة مئوية مباشرة
    const annualCostM = cashbackWhite * annualFuelRevenueM * 0.7;
    return { ...sc, cashbackWhite, annualCostM };
  });
}

/** الأثر المالي لكل مستأجر (ورقة محرر المستأجرين) */
export interface TenantImpact {
  annualRevenue: number;
  annualProfit: number;
  pointsCost: number; // تكلفة النقاط على المستأجر/سنة
  pctOfProfit: number; // % من ربح المستأجر
  upliftRevenue: number; // الزيادة المتوقعة في الإيراد
  extraProfit: number; // الربح الإضافي
  netReturn: number; // صافي العائد للمستأجر
  roi: number; // العائد على الاستثمار للمستأجر
}

export function tenantImpact(
  monthlyRevenue: number,
  marginPct: number,
  contributionHalala: number,
  upliftPct = 0.15
): TenantImpact {
  const annualRevenue = monthlyRevenue * 12;
  const annualProfit = annualRevenue * marginPct;
  const pointsCost = (annualRevenue * contributionHalala) / 100;
  const upliftRevenue = annualRevenue * upliftPct;
  const extraProfit = upliftRevenue * marginPct;
  const netReturn = extraProfit - pointsCost;
  return {
    annualRevenue,
    annualProfit,
    pointsCost,
    pctOfProfit: annualProfit ? pointsCost / annualProfit : 0,
    upliftRevenue,
    extraProfit,
    netReturn,
    roi: pointsCost ? netReturn / pointsCost : 0,
  };
}

/** حاسبة النقطة لكل محطة على حدة — بناءً على هامشها الفعلي */
export interface StationPointResult {
  literPrice: number;
  marginPerLiter: number; // هللة/لتر
  marginPerRiyal: number; // هللة/ريال
  pointCostHalala: number; // تكلفة النقطة (هللة/ريال) — القرار
  pctOfMargin: number; // النسبة المستهلكة من الهامش
  cashbackPct: number; // الكاش باك الفعلي
  earnRate: number; // معدل الكسب المكافئ (نقطة/ريال)
  netMarginPerRiyal: number; // صافي هامش المحطة بعد الولاء
  monthlyRevenue: number;
  monthlyCost: number; // تكلفة الولاء الشهرية على المحطة
  annualCost: number;
  feasible: "good" | "tight" | "loss";
  recommendedHalala: number; // الموصى به = نسبة من الهامش
}

export function stationPointCalc(opts: {
  literPrice: number;
  marginHalalaPerLiter: number;
  litersPerMonth: number;
  pointCostHalala: number; // تكلفة النقطة المختارة (هللة/ريال)
  pointValue: number;
  budgetPctOfMargin?: number; // نسبة الهامش الموصى بتخصيصها (افتراضي 15%)
}): StationPointResult {
  const {
    literPrice,
    marginHalalaPerLiter,
    litersPerMonth,
    pointCostHalala,
    pointValue,
    budgetPctOfMargin = 0.15,
  } = opts;
  const marginPerRiyal = literPrice ? marginHalalaPerLiter / literPrice : 0;
  const pctOfMargin = marginPerRiyal ? pointCostHalala / marginPerRiyal : 0;
  const cashbackPct = pointCostHalala / 100;
  const earnRate = cashbackPct * pointValue;
  const netMarginPerRiyal = marginPerRiyal - pointCostHalala;
  const monthlyRevenue = litersPerMonth * literPrice;
  const monthlyCost = monthlyRevenue * cashbackPct;
  const feasible: StationPointResult["feasible"] =
    netMarginPerRiyal <= 0 ? "loss" : pctOfMargin <= 0.25 ? "good" : pctOfMargin <= 0.5 ? "tight" : "loss";
  return {
    literPrice,
    marginPerLiter: marginHalalaPerLiter,
    marginPerRiyal,
    pointCostHalala,
    pctOfMargin,
    cashbackPct,
    earnRate,
    netMarginPerRiyal,
    monthlyRevenue,
    monthlyCost,
    annualCost: monthlyCost * 12,
    feasible,
    recommendedHalala: marginPerRiyal * budgetPctOfMargin,
  };
}

// ---------------- أدوات تنسيق ----------------
const ar = "ar-SA";
export const fmtInt = (n: number) =>
  new Intl.NumberFormat(ar, { maximumFractionDigits: 0 }).format(n);
export const fmtNum = (n: number, d = 2) =>
  new Intl.NumberFormat(ar, { maximumFractionDigits: d }).format(n);
export const fmtSar = (n: number, d = 0) =>
  new Intl.NumberFormat(ar, { maximumFractionDigits: d }).format(n) + " ﷼";
export const fmtPct = (n: number, d = 2) =>
  new Intl.NumberFormat(ar, { maximumFractionDigits: d }).format(n * 100) + "%";
export const fmtMillion = (n: number) => fmtNum(n, 2) + " م";
