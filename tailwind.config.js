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
      }
    },
  },
  plugins: [],
}
