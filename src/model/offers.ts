import { useState, useEffect } from "react";
import { SAMPLE_OFFERS, OFFER_CATEGORIES } from "./defaults";

export type OfferStatus = "approved" | "pending" | "rejected";

export interface Offer {
  id: number;
  merchant: string;
  title: string;
  cat: string;
  emoji: string;
  loc: string;
  points: number;
  value: number;
  status: OfferStatus; // درب هي من تعتمد — التاجر يقترح فقط
  active: boolean;
  manager: string; // مدير حساب درب المسؤول عن مجموعة التجار
}

const STORAGE = "tanki.offers.v2";

/** مدراء حسابات درب — كل واحد يدير مجموعة تجار/منطقة */
export const ACCOUNT_MANAGERS = [
  "سارة · المنطقة الوسطى",
  "خالد · المنطقة الغربية",
  "نورة · المنطقة الشرقية",
];

const managerForLoc = (loc: string) =>
  loc.includes("الرياض") ? ACCOUNT_MANAGERS[0] : loc.includes("جدة") ? ACCOUNT_MANAGERS[1] : ACCOUNT_MANAGERS[2];

export const emojiForCat = (k: string) =>
  OFFER_CATEGORIES.find((c) => c.key === k)?.emoji ?? "🎁";
export const labelForCat = (k: string) =>
  OFFER_CATEGORIES.find((c) => c.key === k)?.label ?? k;

/** مصدر واحد للعروض يشاركه سوق العميل ولوحة إدارة درب (مخزّن في المتصفح) */
export function useOffers() {
  const [offers, setOffers] = useState<Offer[]>(() => {
    try {
      const s = localStorage.getItem(STORAGE);
      if (s) return JSON.parse(s);
    } catch {
      /* تجاهل */
    }
    // البذرة: كل العروض النموذجية معتمدة ونشطة، ومسندة لمدير حسب الموقع
    return SAMPLE_OFFERS.map((o) => ({
      ...o,
      status: "approved" as OfferStatus,
      active: true,
      manager: managerForLoc(o.loc),
    }));
  });
  useEffect(() => {
    try {
      localStorage.setItem(STORAGE, JSON.stringify(offers));
    } catch {
      /* تجاهل */
    }
  }, [offers]);
  return { offers, setOffers };
}

/** ما يراه العميل فقط: معتمد + نشط */
export const visibleOffers = (offers: Offer[]) =>
  offers.filter((o) => o.status === "approved" && o.active);
