/** @type {import('tailwindcss').Config} */
export default {
  content: ["./index.html", "./src/**/*.{js,jsx}"],
  theme: {
    extend: {
      fontFamily: {
        sans: ["Inter", "Plus Jakarta Sans", "ui-sans-serif", "system-ui", "sans-serif"],
      },
      colors: {
        app: "var(--color-app)",
        card: "var(--color-card)",
        line: "var(--color-line)",
        ink: "var(--color-ink)",
        inksoft: "var(--color-inksoft)",
        inkmuted: "var(--color-inkmuted)",
        accent: {
  DEFAULT: "var(--color-accent)",
  deep: "var(--color-accent-deep)",
  soft: "var(--color-accent-soft)",
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
