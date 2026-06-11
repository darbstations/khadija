// شعار درب — إعادة رسم متجهي بألوان الهوية (برتقالي #F18A2B · رمادي #6D6E70)
export function DarbLogo({ className = "" }: { className?: string }) {
  return (
    <svg viewBox="0 0 250 72" className={className} role="img" aria-label="درب · Darb">
      <rect x="3" y="3" width="244" height="66" rx="33" fill="#6D6E70" stroke="#F18A2B" strokeWidth="5" />
      {/* العلامة البرتقالية المميزة (تشبه حرف الدال) */}
      <g fill="#F18A2B">
        <path d="M158 23 h42 a12 12 0 0 1 12 12 v1 a17 17 0 0 1 -17 17 h-12 a4 4 0 0 1 -4 -4 v-6 a4 4 0 0 1 4 -4 h9 a5 5 0 0 0 5 -5 a4 4 0 0 0 -4 -4 h-35 a4 4 0 0 1 -4 -4 v-4 a3 3 0 0 1 4 -3 z" />
        <rect x="158" y="42" width="20" height="9" rx="4.5" />
      </g>
      {/* النص */}
      <text x="128" y="35" textAnchor="end" fontFamily="Tajawal, sans-serif" fontSize="27" fontWeight="800" fill="#ffffff">درب</text>
      <text x="128" y="56" textAnchor="end" fontFamily="Tajawal, sans-serif" fontSize="18" fontWeight="700" fill="#ffffff" letterSpacing="1">Darb</text>
    </svg>
  );
}
