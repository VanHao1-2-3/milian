/** @type {import('tailwindcss').Config} */
export default {
  content: ['./index.html', './src/**/*.{js,jsx}'],
  theme: {
    extend: {
      fontFamily: {
        sans: ['Inter', 'ui-sans-serif', 'system-ui', 'sans-serif'],
        mono: ['ui-monospace', 'SFMono-Regular', 'Menlo', 'Consolas', 'monospace']
      },
      colors: {
        ink: {
          DEFAULT: '#1a1d24',
          soft: '#565f6b',
          faint: '#8b93a0'
        },
        surface: '#ffffff',
        canvas: '#f3f2ee',
        sidebar: '#15171d',
        accent: {
          DEFAULT: '#c8842f',
          dark: '#a1691f',
          soft: '#f4e4c8'
        },
        line: '#e2ded2'
      },
      boxShadow: {
        card: '0 1px 2px 0 rgba(20,20,24,0.04), 0 1px 1px 0 rgba(20,20,24,0.03)',
        popover: '0 8px 24px -4px rgba(20,20,24,0.12)'
      },
      borderRadius: {
        xl2: '14px'
      }
    }
  },
  plugins: []
};
