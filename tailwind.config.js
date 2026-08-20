/** @type {import('tailwindcss').Config} */
export default {
  content: ["./index.html", "./src/**/*.{js,jsx}"],
  theme: {
    extend: {
      fontFamily: {
        sans: ["Inter", "Plus Jakarta Sans", "ui-sans-serif", "system-ui", "sans-serif"],
      },
      colors: {
        app: "#F8FAFC",
        card: "#FFFFFF",
        line: "#E2E8F0",
        ink: "#0F172A",
        inksoft: "#334155",
        inkmuted: "#64748B",
        accent: {
  DEFAULT: "#2563EB",
  deep: "#1D4ED8",
  soft: "#EFF6FF",
},
        badge: {
          green: { bg: "#DCFCE7", text: "#15803D" },
          red: { bg: "#FEE2E2", text: "#B91C1C" },
          purple: { bg: "#EDE9FE", text: "#6D28D9" },
          amber: { bg: "#FEF3C7", text: "#B45309" },
          slate: { bg: "#F1F5F9", text: "#475569" },
        },
      },
      boxShadow: {
        xs: "0 1px 2px 0 rgba(15,23,42,0.04)",
        card: "0 1px 3px 0 rgba(15,23,42,0.06), 0 1px 2px -1px rgba(15,23,42,0.06)",
      },
      borderRadius: {
        xl2: "14px",
      },
    },
  },
  plugins: [],
}
