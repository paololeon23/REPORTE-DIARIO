/** Configuración — Q Berries · Reporte diario de cosecha */
window.APP = window.APP || {};

APP.CONFIG = {
  _ep: "https://script.google.com/macros/s/AKfycbwM2wRWmQDlGEcZbRzDc2B6p73GgS2uW2tHurRWw2g0hEYiCa_ZRqjqYwk0mJoF5LU35w/exec",
  APP_NAME: "Reporte KG",
  VERSION: "2.1.42",
  TZ: "America/Lima",
  LOCALE: "es-PE",
  YEAR: 2026,
  FUNDOS: ["LICAPA I", "LICAPA II", "LICAPA III", "LICAPA IV"],
  /** kg por jarra. El campo de la fila lo puede cambiar. */
  JARRA_KG: 1.14,
  /** QR de instalación. En iPhone abre el aviso de Safari. */
  PUBLIC_URL: "https://reporte-diario-sup.netlify.app/install.html",
};
