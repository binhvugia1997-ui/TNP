/** @type {import('tailwindcss').Config} */
export default {
  content: ['./index.html', './src/**/*.{ts,tsx}'],
  theme: {
    extend: {
      colors: {
        app: '#eef2f7',
        card: '#ffffff',
        line: '#d8e1ec',
        lineSoft: '#e7edf5',
        header: '#0f2b4c',
        body: '#1f3247',
        muted: '#5b7290',
        brand: {
          50: '#eef6ff',
          100: '#d9eaff',
          200: '#bcd9fb',
          400: '#5aa2ea',
          500: '#1a73d1',
          600: '#1565c0',
          700: '#0f4f99',
        },
        ok: { 50: '#e9f8ee', 500: '#15803d', 600: '#116b33' },
        warn: { 50: '#fef5e5', 500: '#b45309', 600: '#96500a' },
        danger: { 50: '#fdecec', 500: '#b91c1c', 600: '#9c1616' },
        info: { 50: '#eef6ff', 500: '#1d63b8' },
      },
      fontSize: {
        xxs: ['10px', '14px'],
        xs2: ['11px', '15px'],
        sm2: ['12px', '17px'],
        base2: ['13px', '19px'],
      },
      borderRadius: {
        sm2: '3px',
      },
      boxShadow: {
        card: '0 1px 2px rgba(15,43,76,0.06)',
        pop: '0 10px 30px rgba(15,43,76,0.18)',
      },
    },
  },
  plugins: [],
}
