/**
 * Baut aus Kopfdaten und HTML-Text der E-Mail ein eigenständiges HTML-Dokument,
 * das Microsoft Graph in ein PDF umwandelt (Ersatz für den Word-Export im VBA-Makro).
 */

export interface PrintPerson {
  name: string;
  email: string;
}

export interface PrintMail {
  subject: string;
  from: PrintPerson;
  to: PrintPerson[];
  cc: PrintPerson[];
  date: Date;
  attachmentNames: string[];
}

/** Inline-Bild: Content-ID bzw. Dateiname → data:-URL. */
export interface InlineImage {
  contentId?: string;
  name: string;
  dataUrl: string;
}

export const escapeHtml = (s: string) =>
  s.replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]!);

const person = (p: PrintPerson) =>
  p.name && p.email && p.name !== p.email ? `${p.name} <${p.email}>` : p.name || p.email;

export function formatDateTime(d: Date): string {
  return d.toLocaleString("de-DE", {
    weekday: "long",
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  });
}

/** Entfernt aktive Inhalte; Graph rendert ohnehin kein JavaScript, aber sicher ist sicher. */
export function stripActiveContent(html: string): string {
  return html
    .replace(/<script\b[\s\S]*?<\/script\s*>/gi, "")
    .replace(/<(iframe|object|embed)\b[\s\S]*?<\/\1\s*>/gi, "")
    .replace(/<\/?(iframe|object|embed)\b[^>]*>/gi, "")
    .replace(/\son\w+\s*=\s*("[^"]*"|'[^']*'|[^\s>]+)/gi, "");
}

/** Liefert den Inhalt von <body> bzw. das ganze Fragment, plus Stilblöcke aus <head>. */
export function extractBody(html: string): { styles: string; body: string } {
  const styles = (html.match(/<style\b[\s\S]*?<\/style\s*>/gi) || []).join("\n");
  const m = html.match(/<body\b[^>]*>([\s\S]*)<\/body\s*>/i);
  const body = m ? m[1] : html.replace(/<head\b[\s\S]*?<\/head\s*>/i, "").replace(/<\/?html\b[^>]*>/gi, "");
  return { styles, body };
}

/**
 * Ersetzt "cid:"-Verweise durch eingebettete Bilder. Gesucht wird nach Content-ID
 * und – für Outlook-Versionen ohne contentId – nach dem Dateinamen vor dem "@".
 * Nicht auflösbare Verweise werden entfernt, damit kein kaputtes Bild entsteht.
 */
export function replaceInlineImages(html: string, images: InlineImage[]): { html: string; missing: number } {
  let missing = 0;
  const find = (cid: string) => {
    const id = decodeURIComponent(cid).toLowerCase();
    const short = id.split("@")[0];
    return images.find(
      (img) =>
        (img.contentId && img.contentId.toLowerCase().replace(/^<|>$/g, "") === id) ||
        img.name.toLowerCase() === short ||
        img.name.toLowerCase() === id,
    );
  };
  const out = html.replace(/(src\s*=\s*)(["'])cid:([^"']+)\2/gi, (_m, pre: string, q: string, cid: string) => {
    const img = find(cid);
    if (img) return `${pre}${q}${img.dataUrl}${q}`;
    missing++;
    return `${pre}${q}${q}`;
  });
  return { html: out, missing };
}

/** Entfernt alle eingebetteten Bilder (Rückfall bei Überschreitung der 2-MB-Grenze). */
export function dropDataImages(html: string): string {
  return html.replace(/<img\b[^>]*src\s*=\s*["']data:[^"']*["'][^>]*>/gi, "<span>[Bild entfernt]</span>");
}

export function buildPrintHtml(mail: PrintMail, bodyHtml: string): string {
  const { styles, body } = extractBody(stripActiveContent(bodyHtml));
  const rows: [string, string][] = [
    ["Von", person(mail.from)],
    ["Gesendet", formatDateTime(mail.date)],
    ["An", mail.to.map(person).join("; ")],
  ];
  if (mail.cc.length) rows.push(["Cc", mail.cc.map(person).join("; ")]);
  rows.push(["Betreff", mail.subject]);
  if (mail.attachmentNames.length) rows.push(["Anlagen", mail.attachmentNames.join("; ")]);

  const header = rows
    .map(([k, v]) => `<tr><th>${escapeHtml(k)}:</th><td>${escapeHtml(v)}</td></tr>`)
    .join("");

  return `<!DOCTYPE html>
<html lang="de"><head><meta charset="utf-8"><title>${escapeHtml(mail.subject)}</title>
<style>
body{font-family:Calibri,"Segoe UI",Arial,sans-serif;font-size:11pt}
table.mail-header{border-collapse:collapse;margin-bottom:12pt;width:100%}
table.mail-header th{text-align:left;vertical-align:top;padding:1pt 8pt 1pt 0;white-space:nowrap;width:1%}
table.mail-header td{padding:1pt 0}
hr.mail-sep{border:none;border-top:1px solid #888;margin:0 0 12pt}
</style>
${styles}
</head><body>
<table class="mail-header">${header}</table>
<hr class="mail-sep">
${body}
</body></html>`;
}
