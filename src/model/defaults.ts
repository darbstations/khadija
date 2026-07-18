import type { ScenarioInputs } from "./types";

/** القيم الموصى بها (نسخة v5) — مبنية على أسعار يونيو 2026 وهوامش درب الفعلية */
export const DEFAULT_INPUTS: ScenarioInputs = {
  // 1
  pointValue: 200,
  earnWhite: 1, // كل ريال = نقطة · النقطة = نص هللة → 0.5%
  earnSilver: 1,
  earnOrange: 1,
  costPerRiyalHalala: 0.5,
  literPrice: 2.33, // بنزين 95 — يونيو 2026
  welcomeGift: 1000,
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
  earnDiesel: 0.5, // 0.25% (هامش الديزل ضعيف)
  earnPartner: 8, // ~4% ممولة من الشريك
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

/** نماذج التفاوض الأربعة مع المستأجرين (ورقة محرر المستأجرين) */
export const NEGOTIATION_MODELS = [
  {
    key: "tenant100",
    label: "🟢 المستأجر 100%",
    tenantHalala: 2,
    darbHalala: 0,
    fit: "هامش عالي (>60%) · مغاسل/كافيهات محلية",
    argument:
      "العميل يأتيك من تطبيق درب مجاناً · تربح من المبيعات الإضافية · المقابل 2 هللة من كل ريال (أقل من 1% من إيرادك)",
    rating: 5,
  },
  {
    key: "split7030",
    label: "🔵 70/30 (المستأجر)",
    tenantHalala: 1.5,
    darbHalala: 0.5,
    fit: "هامش متوسط (40-60%) · مطاعم وسلاسل صغيرة",
    argument: "أنت تتحمل 70% فقط · درب تشارك 30% · مقابل بيانات العملاء + تسويق مجاني داخل التطبيق",
    rating: 4,
  },
  {
    key: "split5050",
    label: "🟡 50/50",
    tenantHalala: 1,
    darbHalala: 1,
    fit: "مستأجر كبير له ثقل · سلاسل عالمية · شركاء استراتيجيون",
    argument: "شراكة عادلة · كلانا يستفيد · درب توفّر التسويق وقاعدة البيانات وأنت الخدمة",
    rating: 3,
  },
  {
    key: "split3070",
    label: "🟠 30/70 (درب)",
    tenantHalala: 0.5,
    darbHalala: 1.5,
    fit: "هامش منخفض (<25%) · سوبرماركت · حالات خاصة",
    argument: "درب تتحمل 70% · أنت 30% فقط · نقدّر هامشك الضعيف لكنك تساعدنا في تكامل المنظومة",
    rating: 2,
  },
];

/** قاعدة بيانات المستأجرين المحتملين (عيّنة) */
export const TENANT_SAMPLES = [
  { name: "سلسلة كافيه (ستاربكس)", type: "كافيه", stations: 30, monthly: 150000, margin: 0.65, model: "split5050", status: "لم نتواصل", note: "هدف استراتيجي · حضور قوي" },
  { name: "كافيه محلي (دانكن)", type: "كافيه", stations: 25, monthly: 100000, margin: 0.55, model: "split7030", status: "لم نتواصل", note: "مرونة أعلى من العالمية" },
  { name: "سوبرماركت (العثيم)", type: "سوبرماركت", stations: 40, monthly: 500000, margin: 0.2, model: "split3070", status: "لم نتواصل", note: "حجم كبير · هامش منخفض" },
  { name: "مغسلة (Magic Wash)", type: "مغسلة", stations: 50, monthly: 80000, margin: 0.7, model: "tenant100", status: "لم نتواصل", note: "خدمة درب الخاصة" },
  { name: "مطعم سريع (البيك)", type: "مطعم", stations: 35, monthly: 200000, margin: 0.45, model: "split5050", status: "لم نتواصل", note: "حركة عالية · فرص ممتازة" },
];

export const TENANT_STATUSES = ["لم نتواصل", "تواصل أولي", "اجتماع تم", "عرض مُقدّم", "Pilot", "موقّع", "مرفوض"];

/** سكربت محادثة التفاوض (6 مراحل) */
export const NEGOTIATION_SCRIPT = [
  { stage: "الافتتاح", min: "5 د", goal: "بناء الثقة وتحديد السياق", script: "«شكراً لاستضافتنا. درب اليوم 173 محطة وننمو بسرعة. أطلقنا برنامج تانكي وفي يومه الأول وصلنا 50K مستخدم نشط. اليوم نريد نوسّع التجربة لتشمل [اسم المستأجر].»", tip: "لا تتكلم عن المال في البداية · ابدأ بأرقام درب القوية" },
  { stage: "عرض القيمة", min: "10 د", goal: "إظهار الفائدة قبل التكلفة", script: "«قاعدة عملاء درب الـ 50K هي عملاؤك المحتملون. العميل الذي يدخل للوقود يصرف +30% إذا وجد خدمات إضافية. لو رأى عرضكم في تطبيقنا، سيدخل عندكم.»", tip: "استخدم أرقاماً واقعية · اعرض شاشة التطبيق إن أمكن" },
  { stage: "عرض النموذج", min: "10 د", goal: "تقديم الخيار الأنسب", script: "«اقتراحنا بسيط: عملاء درب يحصلون على نقاط عند الشراء منكم. مقابل ذلك تساهمون بـ [2/1.5/1/0.5] هللة من كل ريال مبيعات — أقل من 1% من إيراداتكم.»", tip: "اختر النموذج حسب حجمهم وهامش ربحهم" },
  { stage: "الرد على الاعتراضات", min: "15 د", goal: "إزالة المخاوف", script: "«هذا ضغط على هامشي» → «لكنك تكسب +15% مبيعات على الأقل = صافي إيجابي بعد التكاليف.» · «التكامل صعب» → «API بسيط ودعم فني مجاني · أسبوع واحد.»", tip: "تحضّر للاعتراضات مسبقاً · لا تتفاجأ" },
  { stage: "التفاوض على التفاصيل", min: "15 د", goal: "الوصول لاتفاق", script: "«لنبدأ بـ Pilot 3 شهور في محطة واحدة. لو حقق نتائج نوسّع للجميع. خلال هذي الفترة لا التزامات طويلة. ما رأيكم؟»", tip: "Pilot يخفض المقاومة 70% · ابدأ صغير" },
  { stage: "الإغلاق", min: "5 د", goal: "تحديد الخطوة التالية", script: "«ممتاز. سأرسل اتفاقية الـ Pilot خلال 3 أيام، وأقترح اجتماع متابعة الأسبوع القادم. هل يوم [التاريخ] مناسب؟»", tip: "لا تخرج من الاجتماع بدون موعد محدد للخطوة التالية" },
];

/** الاعتراضات الشائعة وكيف تردّ */
export const OBJECTIONS = [
  { q: "التكلفة عالية ولا نقدر", a: "تدفع فقط من المبيعات التي تأتيك من تطبيق درب. لو لم يأتك عميل، لا تدفع شيئاً.", data: "التكلفة ≈ 1% من المبيعات · العائد +15% مبيعات" },
  { q: "التكامل التقني معقد", a: "نوفّر API بسيط جداً ودعماً فنياً مجانياً في الإعداد. التكامل أسبوع واحد. عندنا 15 شريك مدمج بنجاح.", data: "مدة التكامل 5-7 أيام · الدعم مجاني" },
  { q: "عندنا برامج ولاء خاصة", a: "ممتاز! تانكي لا يحل محلها بل يضيف طبقة. العميل يكسب نقاطك ونقاط تانكي معاً = ضعف الجاذبية.", data: "برامج موازية = نمو 30% بدل 15%" },
  { q: "نريد ضمانات على زيادة المبيعات", a: "نقدّم Pilot 3 شهور بدون التزامات. لو لم تحقق زيادة 10% على الأقل، تنسحب بدون أي رسوم.", data: "Pilot 3 شهور · مخاطرة منخفضة · خروج سهل" },
  { q: "نريد نسبة أقل", a: "النسبة مدروسة حسب هامشك. لكن ما رأيك بهيكلة مختلفة: نسبة منخفضة + رسم اشتراك سنوي ثابت؟", data: "مرونة في النموذج · بدائل متعددة" },
  { q: "لا نثق في حماية بياناتنا", a: "درب ملتزمة بنظام حماية البيانات السعودي (PDPL). لن نشارك أي بيانات حساسة. تستطيع مراجعة العقد قبل التوقيع.", data: "PDPL Compliant · عقد قانوني · مراجعة كاملة" },
];

/** أنواع الوقود وأسعارها (يونيو 2026) */
export const FUEL_TYPES = [
  { key: "p95", label: "بنزين 95", price: 2.33, margin: 13 },
  { key: "p91", label: "بنزين 91", price: 2.18, margin: 13 },
  { key: "diesel", label: "ديزل", price: 1.79, margin: 4 },
];

/** عيّنة محطات لقاعدة «محطاتي» */
export const STATION_SAMPLES = [
  { name: "محطة العليا", city: "الرياض", fuel: "p95", margin: 13, liters: 95000 },
  { name: "محطة التحلية", city: "جدة", fuel: "p95", margin: 11, liters: 82000 },
  { name: "محطة الكورنيش", city: "الدمام", fuel: "p91", margin: 12, liters: 88000 },
  { name: "محطة الطريق السريع", city: "الرياض", fuel: "diesel", margin: 4, liters: 120000 },
  { name: "محطة الامتياز - أبها", city: "أبها", fuel: "p95", margin: 6, liters: 68000 },
];

/** قنوات الاستبدال — لكل قناة تكلفة مختلفة على درب */
export const REDEMPTION_CHANNELS = [
  {
    key: "fuel",
    label: "⛽ خصم بنزين",
    defaultMix: 0.3,
    darbCostFactor: 1.0, // خصم مباشر من هامش الوقود الرفيع
    fundedBy: "درب",
    note: "خصم مباشر من هامش الوقود — الأغلى على درب",
  },
  {
    key: "tenant",
    label: "🛍️ عروض وخصومات المستأجرين",
    defaultMix: 0.35,
    darbCostFactor: 0.0, // ممولة من المستأجر
    fundedBy: "المستأجر",
    note: "يموّلها المستأجر من هامشه العالي — الأرخص والأعلى قيمة",
  },
  {
    key: "external",
    label: "🎁 شركات خارجية (جرير/أمازون/الفرسان)",
    defaultMix: 0.2,
    darbCostFactor: 0.97, // شراء القسائم بخصم جملة ~3%
    fundedBy: "درب (بخصم جملة)",
    note: "درب تشتري القسائم بخصم جملة بسيط · تنوّع وجاذبية",
  },
  {
    key: "recharge",
    label: "📱 كرت شحن",
    defaultMix: 0.15,
    darbCostFactor: 0.95, // عمولة المشغّل ~5%
    fundedBy: "درب (بعمولة المشغّل)",
    note: "عمولة شركة الاتصالات ~5% · مكافأة سائلة ومحبوبة",
  },
];

/** باقات/بطاقات درب — منتجات مدفوعة مسبقاً تُشترى من محفظة درب (مستقلة عن النقاط) */
export const BUNDLES = [
  {
    id: "theeban",
    name: "بطاقة ذيبان",
    emoji: "🦊",
    tag: "اقتصادية",
    price: 500,
    items: [
      { label: "رصيد بنزين", qty: "300 ﷼", worth: 300 },
      { label: "غسلات مجانية", qty: "7 غسلات", worth: 175 },
      { label: "أكواب قهوة", qty: "3 أكواب", worth: 54 },
    ],
  },
  {
    id: "faza3",
    name: "بطاقة فزيع",
    emoji: "🦅",
    tag: "الأكثر طلباً",
    price: 1000,
    items: [
      { label: "رصيد بنزين", qty: "600 ﷼", worth: 600 },
      { label: "اشتراك مغسلة واش واي", qty: "شهر كامل", worth: 250 },
      { label: "أكواب قهوة", qty: "10 أكواب", worth: 180 },
    ],
  },
  {
    id: "shaheen",
    name: "بطاقة الشاهين",
    emoji: "🦅",
    tag: "بريميوم",
    price: 2000,
    items: [
      { label: "رصيد بنزين", qty: "1,300 ﷼", worth: 1300 },
      { label: "اشتراك مغسلة", qty: "3 شهور", worth: 600 },
      { label: "أكواب قهوة", qty: "20 كوب", worth: 360 },
    ],
  },
];

/** الحد الأدنى للاستبدال (نقطة) */
export const REDEEM_MIN_POINTS = 200;

/** مميزات المستويات */
export const TIER_PERKS = [
  { key: "white", emoji: "🤍", label: "الأبيض", perks: ["كاش باك 0.5%", "استبدال فوري", "دعم فني (للجميع)", "عروض خارجية أساسية"] },
  { key: "silver", emoji: "🩶", label: "الفضي", perks: ["كل مزايا الأبيض", "نقاط مضاعفة ×1.5 بالمناسبات", "خصومات شركاء 10%", "عروض خارجية أوسع", "إهداء نقاط للعائلة (محدود)"] },
  { key: "orange", emoji: "🟠", label: "البرتقالي", perks: ["كل مزايا الفضي", "نقاط مضاعفة ×2", "خصومات شركاء 15%", "خصم على باقات درب", "عروض خارجية حصرية (بريميوم)", "أولوية وصول مبكر للعروض المحدودة", "مكافأة ولاء سنوية عند التجديد"] },
  { key: "business", emoji: "🏢", label: "الأعمال/الأساطيل", business: true, perks: ["بالدعوة · للشركات والأساطيل", "👤 مدير حساب مخصّص", "فوترة مجمّعة وتقارير إنفاق", "أسعار وشروط خاصة", "تحكّم بحدود بطاقات الموظفين"] },
];

/** سوق عروض الاستبدال (واجهة العميل · موجّه بالموقع) */
export const OFFER_LOCATIONS = [
  "كل المواقع",
  "محطة العليا · الرياض",
  "محطة التحلية · جدة",
  "محطة الكورنيش · الدمام",
];

export const OFFER_CATEGORIES = [
  { key: "all", label: "الكل", emoji: "✨" },
  { key: "restaurant", label: "مطاعم", emoji: "🍔" },
  { key: "cafe", label: "كافيهات", emoji: "☕" },
  { key: "wash", label: "مغاسل", emoji: "🚿" },
  { key: "service", label: "مراكز خدمة", emoji: "🔧" },
  { key: "grocery", label: "بقالة", emoji: "🛒" },
];

export const SAMPLE_OFFERS = [
  { id: 1, merchant: "كافيه أرابيكا", title: "قهوة مختصة مجاناً", cat: "cafe", emoji: "☕", loc: "محطة العليا · الرياض", points: 1200, value: 16 },
  { id: 2, merchant: "برجر هاوس", title: "خصم 20 ﷼ على وجبة", cat: "restaurant", emoji: "🍔", loc: "محطة العليا · الرياض", points: 1500, value: 20 },
  { id: 3, merchant: "ماجيك واش", title: "غسلة خارجية مجانية", cat: "wash", emoji: "🚿", loc: "محطة العليا · الرياض", points: 1000, value: 25 },
  { id: 4, merchant: "ستاربكس", title: "مشروب متوسط مجاناً", cat: "cafe", emoji: "☕", loc: "محطة العليا · الرياض", points: 1800, value: 18 },
  { id: 5, merchant: "مركز الخدمة السريعة", title: "تغيير زيت بخصم 30 ﷼", cat: "service", emoji: "🔧", loc: "محطة التحلية · جدة", points: 2500, value: 30 },
  { id: 6, merchant: "بقالة درب", title: "قسيمة شراء 10 ﷼", cat: "grocery", emoji: "🛒", loc: "محطة التحلية · جدة", points: 2000, value: 10 },
  { id: 7, merchant: "دانكن", title: "دونت + قهوة", cat: "cafe", emoji: "☕", loc: "محطة التحلية · جدة", points: 1100, value: 14 },
  { id: 8, merchant: "كودو", title: "وجبة عائلية بخصم 40 ﷼", cat: "restaurant", emoji: "🍔", loc: "محطة التحلية · جدة", points: 3000, value: 40 },
  { id: 9, merchant: "شاورما الطازج", title: "ساندويتش مجاني", cat: "restaurant", emoji: "🍔", loc: "محطة الكورنيش · الدمام", points: 800, value: 12 },
  { id: 10, merchant: "سباركل", title: "غسلة كاملة بخصم 50%", cat: "wash", emoji: "🚿", loc: "محطة الكورنيش · الدمام", points: 1300, value: 20 },
  { id: 11, merchant: "إطارات برو", title: "فحص إطارات مجاني", cat: "service", emoji: "🔧", loc: "محطة الكورنيش · الدمام", points: 600, value: 15 },
  { id: 12, merchant: "سوبرماركت العثيم", title: "خصم 25 ﷼", cat: "grocery", emoji: "🛒", loc: "محطة الكورنيش · الدمام", points: 4000, value: 25 },
];

/** معايير المنافسين — برامج ولاء الوقود (بيانات بحثية 2026) */
export const BENCHMARKS = [
  { name: "🐚 Shell (Go+)", region: "عالمي", method: "نقطة/لتر", earn: "≈ 1 نقطة/لتر", redeem: "500 نقطة = £2.50", eff: "~0.5–1%" },
  { name: "⛽ ADNOC Rewards", region: "الإمارات", method: "نقاط/درهم", earn: "حتى 3 نقاط/درهم + 25% محفظة", redeem: "دفع مباشر بالنقاط", eff: "غير معلنة رسمياً" },
  { name: "🟢 Emarat (EmCan)", region: "الإمارات", method: "متدرّج/لتر", earn: "برونزي 2 · فضي 3 · ذهبي 4 /لتر", redeem: "عروض ومكافآت", eff: "متدرّج" },
  { name: "🔴 TotalEnergies (Club)", region: "عالمي", method: "قيمة/لتر", earn: "15 سنت/لتر + 2.5% متجر", redeem: "وقود مجاني/غسيل", eff: "~0.6–0.7%" },
  { name: "🔵 ساسكو", region: "السعودية", method: "نقاط/ريال", earn: "0.43 نقطة/ريال", redeem: "100 نقطة = ريال", eff: "~0.43%" },
];

/** كتالوج قنوات الاستبدال الشامل (مجموعات حسب التمويل) */
export const REDEMPTION_CATALOG = [
  { group: "🟢 من الشركاء (مموّلة — الأرخص)", cost: "≈ 0%", items: ["عروض وخصومات التجار", "عروض خارجية حسب المستوى"] },
  { group: "🔵 منتجات درب", cost: "تكلفتك الفعلية", items: ["خصم بنزين", "غسيل/تغيير زيت/خدمة", "متجر المحطة", "باقات درب", "تحويل لرصيد المحفظة"] },
  { group: "🟠 قسائم خارجية (بخصم جملة)", cost: "95–97%", items: ["جرير · أمازون · نون", "كرت شحن", "اشتراكات رقمية (شاهد/أنغامي)"] },
  { group: "🟣 تحويل / تحالف", cost: "حسب الاتفاق", items: ["أميال طيران / قطاف", "إهداء النقاط للعائلة"] },
  { group: "🔴 تحفيزي / سحوبات", cost: "منخفضة", items: ["سحوبات جوائز كبرى (سيارات/آيفونات)", "تبرّع خيري"] },
];

/** سحوبات الجوائز الكبرى (استبدال بالنقاط كدخول) */
export const RAFFLE_PRIZES = [
  { icon: "🚗", name: "سيارة (تويوتا)", value: 90000, pointsPerEntry: 1000 },
  { icon: "📱", name: "آيفون", value: 5000, pointsPerEntry: 500 },
  { icon: "🏍️", name: "سكوتر", value: 8000, pointsPerEntry: 500 },
  { icon: "⌚", name: "ساعة ذكية", value: 1500, pointsPerEntry: 300 },
];

export const BENCHMARK_SOURCES = [
  { label: "ADNOC Rewards", url: "https://www.adnocdistribution.ae/en/rewards" },
  { label: "Emarat EmCan", url: "https://www.emcan.com/en/faq" },
  { label: "TotalEnergies Club", url: "https://totalenergies.co.za/service-stations/loyalty-programmes/club-totalenergies" },
  { label: "Shell Go+", url: "https://www.canadaloyalty.com/en/programs/shell-go-plus" },
];

/** سيناريوهات ثابتة للمقارنة (ورقة مقارنة السيناريوهات) */
export const COMPARISON_SCENARIOS = [
  { key: "sasco", label: "🔴 ساسكو", pv: 100, w: 0.43, s: 0.43, o: 0.43, cost: 0.43 },
  { key: "conservative", label: "🟢 محافظ", pv: 20000, w: 50, s: 75, o: 100, cost: 0.5 },
  { key: "moderate", label: "🟢 معتدل", pv: 10000, w: 100, s: 150, o: 200, cost: 1 },
  { key: "attractive", label: "🟡 جذاب", pv: 10000, w: 150, s: 200, o: 250, cost: 1.5 },
  { key: "generous", label: "🔴 سخي", pv: 10000, w: 200, s: 250, o: 300, cost: 2 },
];
