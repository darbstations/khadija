import { createContext, useContext, useEffect, useState, type ReactNode } from "react";
import type { ScenarioInputs } from "../model/types";
import { DEFAULT_INPUTS } from "../model/defaults";

interface Ctx {
  inputs: ScenarioInputs;
  set: <K extends keyof ScenarioInputs>(k: K, v: ScenarioInputs[K]) => void;
  reset: () => void;
}

const ScenarioCtx = createContext<Ctx | null>(null);
// عند تغيير القيم الموصى بها نرفع الإصدار لتُحمَّل تلقائياً (تتجاهل المحفوظ القديم)
const STORAGE_KEY = "tanki.inputs.v6";

export function ScenarioProvider({ children }: { children: ReactNode }) {
  const [inputs, setInputs] = useState<ScenarioInputs>(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY);
      if (saved) return { ...DEFAULT_INPUTS, ...JSON.parse(saved) };
    } catch {
      /* تجاهل */
    }
    return DEFAULT_INPUTS;
  });

  useEffect(() => {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(inputs));
    } catch {
      /* تجاهل */
    }
  }, [inputs]);

  const set: Ctx["set"] = (k, v) =>
    setInputs((prev) => ({ ...prev, [k]: Number.isNaN(v as number) ? 0 : v }));
  const reset = () => setInputs(DEFAULT_INPUTS);

  return (
    <ScenarioCtx.Provider value={{ inputs, set, reset }}>
      {children}
    </ScenarioCtx.Provider>
  );
}

export function useScenario() {
  const ctx = useContext(ScenarioCtx);
  if (!ctx) throw new Error("useScenario must be used within ScenarioProvider");
  return ctx;
}
