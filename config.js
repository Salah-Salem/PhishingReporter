/* eslint-disable no-undef */
const config = {
  // API_URL: process.env.API_URL || "https://resilience-stag-api.orgate.io/api/v1",
  // COMPANY_NAME: process.env.COMPANY_NAME || "Resilience",
  // SUPPORT_EMAIL_ADDRESS: process.env.SUPPORT_EMAIL_ADDRESS || "Omar.ali@resilience.sa",
  API_URL: "https://resilience-stag-api.orgate.io/api/v1",
  COMPANY_NAME: "Resilience",
  SUPPORT_EMAIL_ADDRESS: "Omar.ali@resilience.sa",
};

// Fallback for development
if (window.location.hostname === "localhost" || window.location.hostname === "127.0.0.1") {
  config.API_URL = config.API_URL || "https://resilience-dev-api.orgate.io/api/v1";
}

export default config;
