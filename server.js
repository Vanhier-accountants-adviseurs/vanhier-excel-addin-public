const https = require("https");
const express = require("express");
const devCerts = require("office-addin-dev-certs");

const PORT = 3000;

async function start() {
  const app = express();

  // Logt elk binnenkomend verzoek, zodat je in de terminal ziet of Excel de bestanden ophaalt
  app.use((req, res, next) => {
    const timestamp = new Date().toLocaleTimeString();
    console.log(`[${timestamp}] ${req.method} ${req.url}`);
    next();
  });

  app.use(express.static(__dirname));

  // Zorgt dat browsers/Excel de lokale bestanden mogen laden (CORS voor sideload-scenario's)
  app.use((req, res, next) => {
    res.setHeader("Access-Control-Allow-Origin", "*");
    next();
  });

  const options = await devCerts.getHttpsServerOptions();

  https.createServer(options, app).listen(PORT, () => {
    console.log(`Server draait op https://localhost:${PORT}`);
    console.log(`Manifest:      https://localhost:${PORT}/manifest.xml`);
    console.log(`Function file: https://localhost:${PORT}/functions.html`);
  });
}

start().catch((err) => {
  console.error("Kon server niet starten:", err);
  console.error('Heb je "npm run cert" al gedraaid?');
  process.exit(1);
});