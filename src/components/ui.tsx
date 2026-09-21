import type { ReactNode } from "react";

export function Card({
  title,
  subtitle,
  children,
  className = "",
}: {
  title?: string;
  subtitle?: string;
  children: ReactNode;
  className?: string;
}) {
  return (
    <section className={`card ${className}`}>
      {title && (
        <header className="mb-3">
          <h3 className="text-base font-extrabold text-darb-ink">{title}</h3>
          {subtitle && <p className="text-xs text-darb-mut mt-0.5">{subtitle}</p>}
        </header>
      )}
      {children}
    </section>
  );
}

export function Stat({
  label,
  value,
  hint,
  tone = "default",
}: {
  label: string;
  value: ReactNode;
  hint?: string;
  tone?: "default" | "good" | "warn" | "bad" | "accent";
}) {
  const toneCls = {
    default: "text-darb-ink",
    good: "text-darb-good",
    warn: "text-darb-warn",
    bad: "text-darb-bad",
    accent: "text-darb-accent",
  }[tone];
  return (
    <div className="stat">
      <div className="text-[11px] text-darb-mut mb-1">{label}</div>
      <div className={`text-xl font-extrabold leading-tight ${toneCls}`}>{value}</div>
      {hint && <div className="text-[10px] text-darb-mut mt-1">{hint}</div>}
    </div>
  );
}

/** حقل إدخال أصفر — كل التغييرات تنتشر فوراً */
export function NumberInput({
  label,
  value,
  onChange,
  step = 1,
  suffix,
  hint,
  status,
}: {
  label: string;
  value: number;
  onChange: (v: number) => void;
  step?: number;
  suffix?: string;
  hint?: string;
  status?: { text: string; tone: "good" | "warn" | "bad" };
}) {
  return (
    <label className="block">
      <div className="flex items-center justify-between gap-2 mb-1">
        <span className="text-xs text-darb-ink font-bold">{label}</span>
        {status && (
          <span
            className={`pill ${
              status.tone === "good"
                ? "bg-darb-good/15 text-darb-good"
                : status.tone === "warn"
                ? "bg-darb-warn/15 text-darb-warn"
                : "bg-darb-bad/15 text-darb-bad"
            }`}
          >
            {status.text}
          </span>
        )}
      </div>
      <div className="relative">
        <input
          type="number"
          step={step}
          value={Number.isFinite(value) ? value : ""}
          onChange={(e) => onChange(parseFloat(e.target.value))}
          className="w-full bg-darb-yellow/10 border border-darb-yellow/40 focus:border-darb-yellow rounded-lg px-3 py-2 text-darb-ink font-bold text-left outline-none transition"
          dir="ltr"
        />
        {suffix && (
          <span className="absolute left-3 top-1/2 -translate-y-1/2 text-[10px] text-darb-mut pointer-events-none">
            {suffix}
          </span>
        )}
      </div>
      {hint && <div className="text-[10px] text-darb-mut mt-1">{hint}</div>}
    </label>
  );
}

export function Badge({
  tone,
  children,
}: {
  tone: "good" | "warn" | "bad" | "accent";
  children: ReactNode;
}) {
  const cls = {
    good: "bg-darb-good/15 text-darb-good",
    warn: "bg-darb-warn/15 text-darb-warn",
    bad: "bg-darb-bad/15 text-darb-bad",
    accent: "bg-darb-accent/15 text-darb-accent",
  }[tone];
  return <span className={`pill ${cls}`}>{children}</span>;
}

export function Bar({ value, tone = "accent" }: { value: number; tone?: string }) {
  const pct = Math.max(0, Math.min(100, value * 100));
  const color =
    tone === "orange" ? "bg-darb-orange" : tone === "good" ? "bg-darb-good" : "bg-darb-accent";
  return (
    <div className="w-full h-2.5 rounded-full bg-darb-line overflow-hidden">
      <div className={`h-full ${color} transition-all`} style={{ width: `${pct}%` }} />
    </div>
  );
}
