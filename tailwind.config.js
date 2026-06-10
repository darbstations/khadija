/** @type {import('tailwindcss').Config} */
export default {
  content: ["./index.html", "./src/**/*.{ts,tsx}"],
  theme: {
    extend: {
      fontFamily: {
        sans: ["Tajawal", "Cairo", "Segoe UI", "Tahoma", "Arial", "sans-serif"],
      },
      colors: {
        darb: {
          bg: "#0b1220",
          panel: "#111a2e",
          card: "#16223c",
          line: "#23314f",
          ink: "#e8eefc",
          mut: "#94a3c4",
          // tiers
          white: "#e8eefc",
          silver: "#c0c7d6",
          orange: "#ff8a3d",
          accent: "#3dd6c4",
          yellow: "#ffd54a",
          good: "#34d399",
          warn: "#fbbf24",
          bad: "#f87171",
        },
      },
    },
  },
  plugins: [],
};
