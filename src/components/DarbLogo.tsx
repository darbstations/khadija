// شعار درب — إعادة رسم متجهي دقيقة بألوان الهوية (برتقالي #F18A2B · رمادي #6D6E70 · أبيض)
export function DarbLogo({ className = "" }: { className?: string }) {
  return (
    <svg viewBox="0 0 300 100" className={className} role="img" aria-label="درب · Darb">
      <rect x="4" y="4" width="292" height="92" rx="46" fill="#F18A2B" />
      <rect x="13" y="13" width="274" height="74" rx="37" fill="#6D6E70" />
      {/* العلامة المميزة (حرف الدال البرتقالي) */}
      <path
        d="M 186 31 H 246 a 16 16 0 0 1 16 16 v 5 a 17 17 0 0 1 -17 17 H 210 l 20 -13 h 14 a 4 4 0 0 0 4 -4 v -3 a 4 4 0 0 0 -4 -4 H 186 Z"
        fill="#F18A2B"
      />
      <path d="M 186 60 H 208 l -8 13 H 186 Z" fill="#F18A2B" />
      {/* النص */}
      <text x="110" y="52" textAnchor="middle" fontFamily="Tajawal, sans-serif" fontSize="42" fontWeight="800" fill="#ffffff">
        درب
      </text>
      <text x="110" y="84" textAnchor="middle" fontFamily="Tajawal, sans-serif" fontSize="30" fontWeight="700" fill="#ffffff">
        Darb
      </text>
    </svg>
  );
}
