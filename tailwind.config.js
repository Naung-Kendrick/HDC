/** @type {import('tailwindcss').Config} */
export default {
  content: [
    "./index.html",
    "./src/**/*.{js,ts,jsx,tsx}",
  ],
  theme: {
    extend: {
      colors: {
        primary: '#1E40AF',
        secondary: '#475569',
        'n-bg': '#FFFFFF',
        'n-sidebar': '#FAFAFA',
        'n-text': '#1A1A1A',
        'n-text-2': '#737373',
        'n-border': '#E5E7EB',
        'n-hover': '#F3F4F6',
        'n-active': '#E5E7EB',
        'n-accent': '#000000',
      },
      fontFamily: {
        sans: ['Inter', 'Noto Sans Myanmar', 'Public Sans', 'Pyidaungsu', 'Myanmar Text', 'sans-serif'],
      },
      borderRadius: {
        'none': '0px',
      },
      keyframes: {
        fadeInUp: {
          '0%': { opacity: '0', transform: 'translateY(8px)' },
          '100%': { opacity: '1', transform: 'translateY(0)' },
        },
        checkDraw: {
          '0%': { opacity: '0', transform: 'scale(0.7)' },
          '60%': { transform: 'scale(1.08)' },
          '100%': { opacity: '1', transform: 'scale(1)' },
        },
      },
      animation: {
        'fade-in-up': 'fadeInUp 0.35s ease-out both',
        'check-draw': 'checkDraw 0.4s ease-out both',
      },
    },
  },
  plugins: [],
}
