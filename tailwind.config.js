/** @type {import('tailwindcss').Config} */
module.exports = {
  content: ["./index.html", "./main.js"],
  theme: {
    extend: {
      colors: {
        brand: {
          green: '#0B3D2E',
          cream: '#F8F5F0',
          gold:  '#A57C2B',
          text:  '#1A1A1A',
          espresso: '#5A3A2E'
        }
      },
      fontFamily: {
        sans: ['Inter','system-ui','sans-serif'],
        serif: ['Playfair Display','serif']
      },
      boxShadow: { soft: '0 10px 30px rgba(0,0,0,.08)' },
      borderRadius: { xl2: '1rem' }
    }
  },
  plugins: [],
};