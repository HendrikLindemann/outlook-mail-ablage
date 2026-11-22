# Mail-Ablage – Hinweise für Änderungen

Outlook-Web-Add-in (Office.js, Lesemodus), das die geöffnete E-Mail als PDF und EML samt Anlagen in OneDrive
oder SharePoint ablegt. Nachfolger der VBA-Makros `SaveAsPDFfile`/`SaveAttachments2`. Für alle Benutzer des
Mandanten SBHL als Organisations-Add-in installiert. Antworten auf Deutsch; der Benutzer führt Schritte am
eigenen Rechner selbst aus (nummerierte Anleitungen, je Befehl ein Codeblock, PowerShell 7) und committet mit
GitHub Desktop. Hintergrundwissen: Skill `outlook-addin-entwicklung`, besonders `references/lesemodus-ablage.md`.

## Eckdaten

| | |
|---|---|
| Repository | `HendrikLindemann/outlook-mail-ablage` (öffentlich), Branch `main` |
| Veröffentlicht | `https://hendriklindemann.github.io/outlook-mail-ablage/` (GitHub Pages über Actions) |
| Manifest | `manifest.template.xml` → `dist/manifest.xml`, Add-in-ID `5475c721-578c-4dff-965e-1659d7a1b9a2` |
| App-Registrierung | „Mail-Ablage", Client-ID `03b25622-47cb-4234-9c91-2d80b51a5e84`, nur Mandant `a7d6612e-dc90-4650-b8be-19612e7eadcb`; delegiert `Files.ReadWrite.All`, `Sites.Read.All`; Umleitungs-URI `brk-multihub://hendriklindemann.github.io` |
| Konfiguration | `public/config.json` (`clientId`, `authority`). Leere `clientId` ⇒ Demo-Modus ohne echte Ablage |

## Aufbau

- `src/host.ts` – alle Office.js-Zugriffe (Mail, Anlagen, EML, Roaming Settings) und Demo-Host für den Browser
- `src/core/names.ts` – Dateiname `yy-mm-dd Absender - Betreff`, Bereinigung für SharePoint
- `src/core/printhtml.ts` – Druck-HTML (Kopftabelle, Text, `cid:`-Bilder als `data:`-URLs)
- `src/core/save.ts` – Ablauf: PDF, EML, Anlagen; Fehler je Teil, kein Gesamtabbruch
- `src/drive/graph.ts` – Microsoft Graph: Ordner, Upload (einfach/Upload-Sitzung), PDF über `$batch`
- `src/drive/memory.ts` – Graph-Nachbildung für Demo und Tests
- `src/drive/auth.ts` – Anmeldung per Nested App Authentication (MSAL)
- `src/ui/app.ts`, `src/ui/picker.ts` – Hauptansicht und Ordnerauswahl
- `scripts/verteilen.ps1` – Installation als Organisations-Add-in

## Befehle

```powershell
npm install
npm test           # Vitest
npm run build      # tsc + vite build + dist/manifest.xml
npm run e2e        # Playwright mit Edge und Office-Mock (nach build)
npm run icons      # Symbole neu erzeugen
```

Vor jedem Push: `npm test`, `npm run build`, `npm run e2e` müssen grün sein (laufen auch in `deploy.yml`).

## Änderungen ausrollen

**Nur Code, Oberfläche oder `config.json`** (z. B. Dateinamen, PDF-Layout, neue Optionen):
1. Ändern, Tests anpassen, lokal prüfen.
2. Benutzer committet und pusht mit GitHub Desktop.
3. `deploy.yml` veröffentlicht automatisch; Outlook lädt die neue Version beim nächsten Öffnen. Keine Neuinstallation.
4. Prüfen: Actions-Lauf grün (`https://api.github.com/repos/HendrikLindemann/outlook-mail-ablage/actions/runs?per_page=3`), geänderte Datei unter der Pages-Adresse abrufen.

**Manifest** (Name, Beschreibung, Schaltfläche, Symbole-Adressen, Berechtigungen, ExtensionPoints, Mindestversion):
1. In `manifest.template.xml` ändern und **`<Version>` erhöhen** (z. B. `1.0.0.0` → `1.0.1.0`). Die `<Id>` nie ändern.
2. Lokal `npm run build`, danach `npx office-addin-manifest validate -p dist/manifest.xml` (warnt über localhost-Adressen; maßgeblich ist der Lauf von `validate.yml`).
3. Commit, Push, warten bis Deploy grün ist und `…/manifest.xml` die neue Version zeigt.
4. Benutzer verteilt neu (PowerShell 7, im Projektordner; vorher `Import-Module ExchangeOnlineManagement` und `Connect-ExchangeOnline` im selben Tab, sonst erledigt das Skript die Anmeldung):
   ```powershell
   .\scripts\verteilen.ps1 -Alle -WhatIf
   ```
   ```powershell
   .\scripts\verteilen.ps1 -Alle
   ```
5. Sichtbarkeit bei allen kann Stunden bis zwei Tage dauern – nicht vorschnell erneut installieren.

**Graph-Berechtigungen** zusätzlich: in Entra unter App-Registrierungen › Mail-Ablage › API-Berechtigungen ergänzen und Administratorzustimmung erteilen; Bereiche in `src/drive/auth.ts` (`SCOPES`) anpassen.

## Bekannte Grenzen

- Freigegebene Postfächer (z. B. `post@sbhl.de`) brauchen `SupportsSharedFolders` im `VersionOverridesV1_1`-Block des
  Manifests (ab Version 1.0.1.0 enthalten). Beide `VersionOverrides`-Blöcke bei Änderungen gleich halten.

- PDF-Umwandlung durch Graph nur bis 2 MB Quell-HTML; darüber ohne Bilder (Hinweis im Ergebnis).
- `.eml` braucht Mailbox 1.14; ältere Clients speichern nur PDF und Anlagen.
- Per Drag-and-drop angehängte Mails liefert Office.js im Web/neuen Outlook nicht.
- Kein Zugriff auf lokale Laufwerke oder Explorer-Fenster (Web-Add-in); lokal nur über OneDrive-Synchronisierung.

## Stand (21.11.2026)

Live getestet in Outlook im Web: OneDrive-Ablage, PDF, EML, Anlagen. Offen: SharePoint-Ablage live, Anzeige im
klassischen Outlook nach der Verteilung.
