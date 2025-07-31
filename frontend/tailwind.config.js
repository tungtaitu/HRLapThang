/** @type {import('tailwindcss').Config} */
module.exports = {
  content: [
    "./src/**/*.{js,jsx,ts,tsx}",
  ],
  theme: {
    extend: {
      fontFamily: {
        // Ghi đè font sans-serif mặc định bằng font mới của bạn
        'sans': ['TÊN_FONT', 'Times New Roman'],
      },
    },
  },
  plugins: [],
}