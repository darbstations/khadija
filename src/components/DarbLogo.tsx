// شعار درب — wordmark نظيف بألوان الهوية (برتقالي #F18A2B · رمادي #6D6E70 · أبيض)
// ملاحظة: نسخة مؤقتة لحين توفّر ملف الشعار الأصلي (PNG/SVG) لاستبداله.
export function DarbLogo({ className = "" }: { className?: string }) {
  return (
    <svg viewBox="0 0 220 64" className={className} role="img" aria-label="درب · Darb">
      <rect x="2" y="2" width="216" height="60" rx="30" fill="#6D6E70" stroke="#F18A2B" strokeWidth="3" />
      {/* شيفرون برتقالي نظيف (إيحاء الطريق/التقدّم) */}
      <path
        d="M170 19 L185 32 L170 45"
        fill="none"
        stroke="#F18A2B"
        strokeWidth="6"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
      <text x="150" y="31" textAnchor="end" fontFamily="Tajawal, sans-serif" fontSize="27" fontWeight="800" fill="#ffffff">
        درب
      </text>
      <text x="150" y="49" textAnchor="end" fontFamily="Tajawal, sans-serif" fontSize="14" fontWeight="700" fill="#ffffff" letterSpacing="2">
        Darb
      </text>
    </svg>
  );
}
