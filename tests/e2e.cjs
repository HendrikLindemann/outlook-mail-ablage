// Durchlauftest: Vorschau-Server starten, office.js durch den Mock ersetzen, Ablage im Demo-Speicher prüfen.
// Aufruf nach "npm run build": npm run e2e
const { chromium } = require("playwright");
const { spawn } = require("node:child_process");
const { readFileSync } = require("node:fs");
const path = require("node:path");

const PORT = 4179;
const mock = readFileSync(path.join(__dirname, "office-mock.js"), "utf8");

function startServer() {
  const proc = spawn("npx", ["vite", "preview", "--port", String(PORT), "--strictPort"], {
    cwd: path.join(__dirname, ".."),
    shell: true,
    env: { ...process.env },
  });
  proc.stderr.on("data", (d) => process.stderr.write(d));
  // Warten, bis der Port Verbindungen annimmt.
  return new Promise((resolve, reject) => {
    const until = Date.now() + 20000;
    const probe = () => {
      const s = require("node:net").connect({ port: PORT, host: "localhost", autoSelectFamily: true });
      s.on("connect", () => {
        s.destroy();
        resolve(proc);
      });
      s.on("error", () => {
        if (Date.now() > until) {
          stopServer(proc);
          reject(new Error("Server startet nicht"));
        } else setTimeout(probe, 300);
      });
    };
    probe();
  });
}

function stopServer(proc) {
  if (process.platform === "win32") spawn("taskkill", ["/pid", String(proc.pid), "/f", "/t"]);
  else proc.kill();
}

const assert = (cond, msg) => {
  if (!cond) throw new Error(`Fehlgeschlagen: ${msg}`);
  console.log(`  ✓ ${msg}`);
};

(async () => {
  const server = await startServer();
  const browser = await chromium.launch({ channel: "msedge" });
  let failed = false;
  try {
    const page = await browser.newPage({ viewport: { width: 360, height: 800 }, ignoreHTTPSErrors: true });
    const errors = [];
    page.on("pageerror", (e) => errors.push(e.message));
    await page.route("**/office.js", (r) => r.fulfill({ contentType: "text/javascript", body: mock }));
    // Ohne Client-ID läuft das Add-in mit dem Demo-Speicher – unabhängig von der echten config.json.
    await page.route("**/config.json", (r) => r.fulfill({ contentType: "application/json", body: "{}" }));
    // vite.config.ts nutzt HTTPS nur, wenn das Office-Entwicklerzertifikat installiert ist.
    const certs = require("node:fs").existsSync(path.join(require("node:os").homedir(), ".office-addin-dev-certs", "localhost.key"));
    await page.goto(`${certs ? "https" : "http"}://localhost:${PORT}/`);

    console.log("Hauptansicht");
    await page.getByText("Mail-Ablage", { exact: true }).waitFor();
    const name = await page.getByLabel("Dateiname (ohne Endung)").inputValue();
    assert(name === "26-11-19 Dr. Erika Müller - AW Ortstermin 12.12. Gutachten", `Dateiname wie im VBA-Makro (${name})`);
    assert(await page.getByText("Demo-Modus").isVisible(), "Demo-Hinweis ohne Client-ID");
    assert(await page.getByText("Tipp: Ordner, die der OneDrive-Client").isVisible(), "Desktop-Hinweis");
    assert(await page.getByLabel(/Lageplan\.pdf/).isChecked(), "normale Anlage vorausgewählt");
    assert(!(await page.getByLabel(/image001\.png/).isChecked()), "eingebettetes Bild nicht vorausgewählt");

    await page.getByRole("button", { name: "Ablegen" }).click();
    assert(await page.getByText("Bitte zuerst einen Zielordner wählen.").isVisible(), "Meldung ohne Ziel");

    console.log("Ordnerauswahl");
    await page.getByRole("button", { name: "Ordner wählen …" }).click();
    await page.getByRole("button", { name: /Mein OneDrive/ }).click();
    await page.getByRole("button", { name: /Projekte/ }).click();
    await page.getByRole("button", { name: "Neuer Ordner …" }).click();
    await page.getByLabel("Neuer Ordner").fill("Müller Ortstermin");
    await page.getByRole("button", { name: "Anlegen" }).click();
    await page.getByRole("heading", { name: /Projekte \/ Müller Ortstermin/ }).waitFor();
    await page.getByRole("button", { name: "Diesen Ordner wählen" }).click();
    const label = await page.locator("#target-label").textContent();
    assert(label === "Mein OneDrive / Projekte / Müller Ortstermin", `Zielordner übernommen (${label})`);

    console.log("Ablegen");
    await page.getByRole("button", { name: "Ablegen" }).click();
    await page.getByRole("heading", { name: "Abgelegt" }).waitFor();
    const results = await page.locator(".result").allTextContents();
    assert(results.length === 4, `vier Dateien gespeichert (${results.length})`);
    assert(results.some((r) => r.includes("AW Ortstermin 12.12. Gutachten.pdf")), "PDF abgelegt");
    assert(results.some((r) => r.includes(".eml") && r.includes("Gutachten")), "EML abgelegt");
    assert(results.some((r) => r.includes("Lageplan.pdf")), "Anlage abgelegt");
    assert(results.some((r) => r.includes("Weitergeleitet Termin.eml")), "angehängte E-Mail als .eml abgelegt");

    await page.screenshot({ path: "test-results/ergebnis.png", fullPage: true });

    console.log("Zuletzt verwendet");
    await page.reload();
    await page.getByText("Mail-Ablage", { exact: true }).waitFor();
    const label2 = await page.locator("#target-label").textContent();
    assert(label2 === "Mein OneDrive / Projekte / Müller Ortstermin", "letzter Ordner wird vorgeschlagen");

    await page.getByRole("button", { name: "Ordner wählen …" }).click();
    await page.getByRole("button", { name: /Team Gutachten/ }).click();
    await page.getByRole("button", { name: /Dokumente/ }).click();
    await page.getByRole("button", { name: /Allgemein/ }).waitFor();
    assert(true, "SharePoint-Website und Bibliothek navigierbar");

    await page.screenshot({ path: "test-results/picker.png" });
    assert(errors.length === 0, `keine Skriptfehler ${errors.join("; ")}`);
  } catch (e) {
    failed = true;
    console.error(e.message);
  } finally {
    await browser.close();
    stopServer(server);
    process.exit(failed ? 1 : 0);
  }
})();
