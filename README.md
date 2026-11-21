# Mail-Ablage – Outlook-Add-in

Nachfolger der VBA-Makros `SaveAsPDFfile` und `SaveAttachments2`. Das Add-in legt die geöffnete E-Mail in einem
OneDrive- oder SharePoint-Ordner ab:

- **E-Mail als PDF** – Kopfzeilen (Von, Gesendet, An, Cc, Betreff, Anlagen) und Text samt eingebetteter Bilder,
  umgewandelt von Microsoft Graph (ersetzt den Word-Export).
- **E-Mail im Original (.eml)** – vollständige Kopie mit allen Anlagen, revisionssicher.
- **Anlagen** einzeln im selben Ordner. Angehängte E-Mails werden `.eml`, Termine `.ics`,
  Cloud-Anlagen eine Verknüpfung (`.url`). Eingebettete Signaturbilder sind abgewählt.
- Dateiname wie im Makro: `yy-mm-dd Absender - Betreff`, unzulässige Zeichen entfernt, vor dem Speichern änderbar.
- Gleichnamige Dateien werden nicht überschrieben, OneDrive hängt eine Nummer an.
- Zuletzt verwendete Ordner (8) und die Optionen werden im Postfach gespeichert (Roaming Settings).

## Unterschied zum VBA-Makro

Ein Web-Add-in läuft auch im Desktop-Outlook in einer abgeschotteten Browser-Umgebung. Es kann weder geöffnete
Explorer-Fenster lesen noch direkt auf lokale Laufwerke schreiben. Deshalb speichert es überall – Desktop, neues Outlook,
Outlook im Web – in OneDrive/SharePoint. Ordner, die der OneDrive-Client synchronisiert, erscheinen trotzdem im Explorer.

## Grenzen

- PDF-Umwandlung durch Graph: Quelle höchstens 2 MB. Wird die Grenze mit Bildern überschritten, entstehen die PDFs
  ohne eingebettete Bilder (Hinweis im Ergebnis). Anlagen sind davon nicht betroffen.
- `.eml` braucht Mailbox 1.14 (aktuelles Outlook für Windows, neues Outlook, Web, Mac). Ältere Versionen speichern nur PDF und Anlagen.
- In Outlook im Web und im neuen Outlook liefert Office.js per Drag-and-drop angehängte E-Mails nicht aus; das Ergebnis meldet diese Anlage als fehlgeschlagen.
- Für die Ablage in fremde SharePoint-Websites reichen die eigenen Rechte des Benutzers; das Add-in kann nichts, was der Benutzer nicht darf.
- Ohne Client-ID in `public/config.json` läuft das Add-in im Demo-Modus (nichts wird gespeichert).

## Entwicklung

```powershell
npm install
npm test          # Vitest: Dateinamen, Druck-HTML, Ablauf mit Demo-Speicher
npm run build
npm run e2e       # Playwright mit Edge und Office-Mock
npm run certs     # einmalig: Entwicklerzertifikat für https://localhost:3000
npm run dev       # erzeugt manifest.dev.xml für localhost
```

Aufbau: `src/host.ts` (alle Office.js-Zugriffe, Demo-Host), `src/core/` (Logik ohne Outlook-Bezug),
`src/drive/graph.ts` (Microsoft Graph), `src/drive/memory.ts` (Demo-/Testspeicher), `src/ui/` (Oberfläche).

## Einrichtung im Mandanten

1. **App-Registrierung** in https://entra.microsoft.com › Identität › Anwendungen › App-Registrierungen › Neue Registrierung
   - Name „Mail-Ablage" (der PDF-Zwischenspeicher liegt kurz in OneDrive › Apps › Mail-Ablage)
   - Kontotyp „Nur Konten in diesem Organisationsverzeichnis"
   - Umleitungs-URI, Plattform „Single-Page-Anwendung": `brk-multihub://hendriklindemann.github.io`,
     für Tests zusätzlich `brk-multihub://localhost:3000`
   - API-Berechtigungen › Microsoft Graph › Delegiert: `Files.ReadWrite.All`, `Sites.Read.All`,
     danach „Administratorzustimmung erteilen"
2. Anwendungs-ID in `public/config.json` als `clientId` eintragen, `authority` auf
   `https://login.microsoftonline.com/<Mandanten-ID>` setzen.
3. **Hosting**: Repository `outlook-mail-ablage` (öffentlich) anlegen, unter Settings › Pages „GitHub Actions" wählen.
   Der Workflow `deploy.yml` baut und veröffentlicht nach `https://hendriklindemann.github.io/outlook-mail-ablage/`.
4. **Installation** per Exchange Online PowerShell mit `dist/manifest.xml` von der veröffentlichten Seite:
   `New-App -Mailbox <Postfach> -FileData ([System.IO.File]::ReadAllBytes("<Pfad>\manifest.xml"))`
