/**
 * ITR Generator – Annual V2.0 — frontend configuration
 * Paste the Annual Apps Script Web App URL below (it is a separate deployment from the quarterly app).
 * You can also set it from "Backend settings" on the sign-in pop-up; that value is saved in the browser.
 */
window.ITR_CONFIG = {
  API_URL: 'PASTE_YOUR_ANNUAL_APPS_SCRIPT_WEB_APP_URL_HERE',
  FIRST_YEAR: 2025,          // earliest taxable year offered (annual returns for 2025 are filed in 2026)
  YEARS_AHEAD: 4,
  PARALLEL_UPLOADS: 3,
  IMAGE_MAX_DIMENSION: 2000,
  IMAGE_QUALITY: 0.85,
  PDF_RASTER_OVER_MB: 6,     // PDFs larger than this are converted to page images in the browser
  PDF_MAX_PAGES: 40,
  PDF_PAGE_WIDTH: 1400
};
