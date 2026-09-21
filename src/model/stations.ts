import { useState, useEffect } from "react";
import { FUEL_TYPES, STATION_SAMPLES } from "./defaults";
import { stationPointCalc } from "./engine";

export interface StationRow {
  id: number;
  name: string;
  city: string;
  fuel: string;
  margin: number; // هللة/لتر
  liters: number; // لتر/شهر
}

const STORAGE = "tanki.stations.v1";
export const priceOf = (fuelKey: string) =>
  FUEL_TYPES.find((f) => f.key === fuelKey)?.price ?? 2.33;
export const fuelLabel = (k: string) => FUEL_TYPES.find((f) => f.key === k)?.label ?? k;

/** مصدر واحد لبيانات المحطات يشاركه «محطاتي» وحاسبة ROI (مخزّن في المتصفح) */
export function useStations() {
  const [stations, setStations] = useState<StationRow[]>(() => {
    try {
      const s = localStorage.getItem(STORAGE);
      if (s) return JSON.parse(s);
    } catch {
      /* تجاهل */
    }
    return STATION_SAMPLES.map((s, idx) => ({ id: idx + 1, ...s }));
  });
  useEffect(() => {
    try {
      localStorage.setItem(STORAGE, JSON.stringify(stations));
    } catch {
      /* تجاهل */
    }
  }, [stations]);
  return { stations, setStations };
}

/** تكلفة الولاء على شبكة المحطات عند تكلفة النقطة الموصى بها لكل محطة */
export function stationsLoyaltyCost(
  stations: StationRow[],
  pointValue: number,
  budgetPctOfMargin = 0.15
) {
  let monthlyCost = 0;
  let monthlyRevenue = 0;
  for (const s of stations) {
    const literPrice = priceOf(s.fuel);
    const base = stationPointCalc({
      literPrice,
      marginHalalaPerLiter: s.margin,
      litersPerMonth: s.liters,
      pointCostHalala: 0,
      pointValue,
      budgetPctOfMargin,
    });
    const at = stationPointCalc({
      literPrice,
      marginHalalaPerLiter: s.margin,
      litersPerMonth: s.liters,
      pointCostHalala: +base.recommendedHalala.toFixed(3),
      pointValue,
      budgetPctOfMargin,
    });
    monthlyCost += at.monthlyCost;
    monthlyRevenue += at.monthlyRevenue;
  }
  return {
    monthlyCost,
    annualCost: monthlyCost * 12,
    annualRevenue: monthlyRevenue * 12,
    blendedCashback: monthlyRevenue ? monthlyCost / monthlyRevenue : 0,
  };
}
