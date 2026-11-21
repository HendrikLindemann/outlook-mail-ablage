/** Dateinamen für OneDrive/SharePoint bilden – Gegenstück zum RegEx-Teil des VBA-Makros. */

// Unzulässig in Windows, OneDrive und SharePoint Online, dazu Steuerzeichen.
const INVALID = /["*:<>?/\\|\u0000-\u001f]/g;
// SharePoint lehnt diese Namen ab (unabhängig von der Endung).
const RESERVED = /^(con|prn|aux|nul|com\d|lpt\d|_vti_|desktop\.ini)$/i;

/** Maximale Länge des Namens ohne Endung; lässt Reserve für Pfad und "(1)"-Zusätze. */
export const MAX_BASE_LENGTH = 120;

const two = (n: number) => String(n).padStart(2, "0");

/** Datum wie VBA Format(…, "yy-mm-dd"). */
export function formatDate(d: Date): string {
  return `${two(d.getFullYear() % 100)}-${two(d.getMonth() + 1)}-${two(d.getDate())}`;
}

/** Entfernt unzulässige Zeichen, fasst Leerraum zusammen und kürzt den Namen. */
export function sanitizeFileName(name: string, maxLength = MAX_BASE_LENGTH): string {
  let s = name.replace(INVALID, "").replace(/\s+/g, " ").trim();
  if (s.length > maxLength) s = s.slice(0, maxLength).trim();
  // Führende/abschließende Punkte und Leerzeichen sind in SharePoint unzulässig.
  s = s.replace(/^[.\s]+|[.\s]+$/g, "");
  if (!s || RESERVED.test(s)) s = s ? `_${s}` : "E-Mail";
  return s;
}

/** "yy-mm-dd Absender - Betreff" wie im VBA-Makro. */
export function buildBaseName(date: Date, senderName: string, subject: string): string {
  const sender = senderName.trim() || "Unbekannt";
  const subj = subject.trim() || "(ohne Betreff)";
  return sanitizeFileName(`${formatDate(date)} ${sender} - ${subj}`);
}

export function splitExtension(name: string): [string, string] {
  const i = name.lastIndexOf(".");
  return i > 0 && i < name.length - 1 ? [name.slice(0, i), name.slice(i)] : [name, ""];
}

/** Bereinigt einen Anlagennamen und erzwingt bei Bedarf eine Endung. */
export function attachmentFileName(name: string, forceExt?: string): string {
  let [base, ext] = splitExtension(name.trim());
  if (forceExt && ext.toLowerCase() !== forceExt) {
    base = base + ext;
    ext = forceExt;
  }
  const cleanExt = ext.replace(INVALID, "");
  return sanitizeFileName(base, 200 - cleanExt.length) + cleanExt;
}
