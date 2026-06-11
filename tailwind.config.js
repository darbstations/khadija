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
          // هوية درب · رمادي #6D6E70 + برتقالي #F18A2B + أبيض
          bg: "#1b1c1e",
          panel: "#242528",
          card: "#2b2c2f",
          line: "#3b3d41",
          ink: "#f4f5f6",
          mut: "#9a9da2",
          // tiers
          white: "#f4f5f6",
          silver: "#bcbfc4",
          orange: "#F18A2B",
          accent: "#F18A2B",
          grey: "#6D6E70",
          yellow: "#ffd54a",
          good: "#3fb27f",
          warn: "#e0922e",
          bad: "#e06666",
        },
      },
    },
  },
  plugins: [],
};
