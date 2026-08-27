// أنواع النموذج المالي لبرنامج تانكي

export type TierKey = "white" | "silver" | "orange";

/** المدخلات (الخلايا الصفراء) — كل الحسابات تشتق منها */
export interface ScenarioInputs {
  // 1) قيمة النقطة ومعدلات الكسب
  pointValue: number; // كم نقطة = 1 ريال عند الاستبدال (C6)
  earnWhite: number; // نقطة/ريال — الأبيض (C7)
  earnSilver: number; // نقطة/ريال — الفضي (C8)
  earnOrange: number; // نقطة/ريال — البرتقالي (C9)
  costPerRiyalHalala: number; // تكلفة الولاء على درب (هللة/ريال) (C10)
  literPrice: number; // سعر لتر البنزين 95 (ريال) (C11)
  welcomeGift: number; // هدية الترحيب (نقطة) (C12)

  // 2) معايير الترقية (إنفاق سنوي بالريال)
  silverThreshold: number; // حد الفضي (C16)
  orangeThreshold: number; // حد البرتقالي (C17)
  spendPeriodMonths: number; // فترة احتساب الإنفاق (C18)
  graceDays: number; // فترة سماح بعد التنزيل (C19)

  // 3) المحطات والهوامش
  stationsOperating: number; // محطات تشغيل (C23)
  stationsInvestment: number; // محطات استثمار (C24)
  stationsFranchise: number; // محطات امتياز (C25)
  marginOperating: number; // هامش التشغيل هللة/لتر (C26)
  marginInvestment: number; // هامش الاستثمار هللة/لتر (C27)
  marginFranchise: number; // هامش الامتياز هللة/لتر (C28)
  darbShareFranchise: number; // مساهمة درب في الامتياز % (C29)
  stationMonthlyLiters: number; // متوسط مبيعات المحطة لتر/شهر (C31)

  // 4) الاقتصاد الموسّع (مربوط بالحسابات فعلياً)
  dieselPrice: number; // سعر لتر الديزل (ريال)
  dieselMarginHalala: number; // هامش الديزل هللة/لتر
  earnDiesel: number; // نقطة/ريال على الديزل
  earnPartner: number; // نقطة/ريال عند الشركاء
  partnerContributionHalala: number; // مساهمة الشريك لدرب (هللة/ريال مبيعات)
}

export interface TierRow {
  key: TierKey;
  label: string;
  emoji: string;
  earnRate: number; // نقطة/ريال
  minSpend: number; // الحد الأدنى للإنفاق السنوي
  cashbackPct: number; // الكاش باك الفعلي = earnRate/pointValue
}

export interface CustomerResult {
  monthlySpend: number;
  annualSpend: number;
  tier: TierRow;
  pointsPerYear: number;
  valueRiyal: number;
  cashbackPct: number;
  nextTierLabel: string;
  remainingSpend: number;
  progress: number; // 0..1
  darbCostRiyal: number; // تكلفة درب على هذا العميل سنوياً
  vsSasco: number; // مضاعف النقاط مقابل ساسكو
}

export interface StationEconomics {
  key: "operating" | "investment" | "franchise";
  label: string;
  marginPerLiter: number; // هللة/لتر
  marginPerRiyal: number; // هللة/ريال
  loyaltyCostPerRiyal: number; // هللة/ريال
  darbSharePct: number;
  darbCostPerRiyal: number; // هللة/ريال
  ownerNetPerRiyal: number; // صافي هامش المالك هللة/ريال
  feasible: "good" | "tight" | "loss";
}
