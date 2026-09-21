import { useState, useEffect } from "react";

export interface Merchant {
  id: number;
  name: string;
  type: string;
  margin: number; // هامش الربح (كسر)
  earnPct: number; // Earn · النسبة التي يدفعها التاجر عند الشراء (MDR)
  customerPct: number; // حصة العميل (نقاط) من نسبة الـEarn
  burnPct: number; // Burn · رسوم تسويق يدفعها التاجر عند الاستبدال عنده
}

const STORAGE = "tanki.merchants.v1";

export const MERCHANT_SEED: Omit<Merchant, "id">[] = [
  { name: "ستاربكس", type: "كافيه", margin: 0.65, earnPct: 3, customerPct: 1, burnPct: 5 },
  { name: "دانكن", type: "كافيه", margin: 0.55, earnPct: 3, customerPct: 1, burnPct: 5 },
  { name: "العثيم", type: "سوبرماركت", margin: 0.2, earnPct: 1.5, customerPct: 0.5, burnPct: 2 },
  { name: "Magic Wash", type: "مغسلة", margin: 0.7, earnPct: 3, customerPct: 1, burnPct: 6 },
  { name: "البيك", type: "مطعم", margin: 0.45, earnPct: 2, customerPct: 1, burnPct: 4 },
];

export function useMerchants() {
  const [merchants, setMerchants] = useState<Merchant[]>(() => {
    try {
      const s = localStorage.getItem(STORAGE);
      if (s) return JSON.parse(s);
    } catch {
      /* تجاهل */
    }
    return MERCHANT_SEED.map((m, i) => ({ id: i + 1, ...m }));
  });
  useEffect(() => {
    try {
      localStorage.setItem(STORAGE, JSON.stringify(merchants));
    } catch {
      /* تجاهل */
    }
  }, [merchants]);
  return { merchants, setMerchants };
}
