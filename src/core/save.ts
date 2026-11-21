import type { DriveApi, FolderRef } from "../drive/types";
import type { AttachmentInfo, Host, MailInfo } from "../host";
import { base64ToBytes, utf8Bytes } from "./base64";
import { attachmentFileName } from "./names";
import { buildPrintHtml, dropDataImages, replaceInlineImages, type InlineImage } from "./printhtml";

export interface SaveOptions {
  baseName: string;
  pdf: boolean;
  eml: boolean;
  /** IDs der zu speichernden Anlagen */
  attachmentIds: string[];
}

export interface SaveResult {
  label: string;
  ok: boolean;
  name?: string;
  webUrl?: string;
  message?: string;
}

const PDF_LIMIT = 2 * 1024 * 1024;

const mimeFromName = (name: string) => {
  const ext = name.split(".").pop()?.toLowerCase();
  return ext === "png" ? "image/png" : ext === "gif" ? "image/gif" : ext === "bmp" ? "image/bmp" : "image/jpeg";
};

/** Anlagen, die standardmäßig gespeichert werden: keine eingebetteten Bilder der Signatur o. Ä. */
export const defaultAttachments = (mail: MailInfo) => mail.attachments.filter((a) => !a.isInline);

async function loadInlineImages(host: Host, attachments: AttachmentInfo[]): Promise<InlineImage[]> {
  const images: InlineImage[] = [];
  for (const a of attachments.filter((x) => x.isInline && x.kind === "file")) {
    try {
      const data = await host.getAttachment(a.id);
      if (data.format !== "base64") continue;
      const type = a.contentType?.startsWith("image/") ? a.contentType : mimeFromName(a.name);
      images.push({ contentId: a.contentId, name: a.name, dataUrl: `data:${type};base64,${data.content}` });
    } catch {
      /* Bild fehlt im PDF – kein Abbruch */
    }
  }
  return images;
}

/** Erzeugt das Druck-HTML; bleibt es mit Bildern über 2 MB, werden die Bilder entfernt. */
export async function buildPdfSource(host: Host, mail: MailInfo): Promise<{ html: string; note?: string }> {
  const body = await host.getBodyHtml();
  const images = await loadInlineImages(host, mail.attachments);
  const { html: withImages } = replaceInlineImages(body, images);
  const printMail = { ...mail, attachmentNames: defaultAttachments(mail).map((a) => a.name) };
  let html = buildPrintHtml(printMail, withImages);
  if (utf8Bytes(html).byteLength > PDF_LIMIT) {
    html = buildPrintHtml(printMail, dropDataImages(withImages));
    if (utf8Bytes(html).byteLength > PDF_LIMIT) {
      throw new Error("Der E-Mail-Text ist auch ohne Bilder größer als 2 MB und kann nicht umgewandelt werden.");
    }
    return { html, note: "Eingebettete Bilder wurden wegen der 2-MB-Grenze weggelassen." };
  }
  return { html };
}

/**
 * Legt PDF, EML und Anlagen im Zielordner ab. Fehler einzelner Teile brechen
 * den Vorgang nicht ab, sondern erscheinen in der Ergebnisliste.
 */
export async function saveMail(
  host: Host,
  drive: DriveApi,
  folder: FolderRef,
  opts: SaveOptions,
  progress: (text: string) => void = () => {},
): Promise<SaveResult[]> {
  const mail = host.getMail();
  const results: SaveResult[] = [];
  const run = async (label: string, work: () => Promise<{ name: string; webUrl?: string; message?: string }>) => {
    progress(`${label} …`);
    try {
      const r = await work();
      results.push({ label, ok: true, ...r });
    } catch (e) {
      results.push({ label, ok: false, message: e instanceof Error ? e.message : String(e) });
    }
  };

  if (opts.pdf) {
    await run("E-Mail als PDF", async () => {
      const { html, note } = await buildPdfSource(host, mail);
      const pdf = await drive.htmlToPdf(html);
      const saved = await drive.upload(folder, `${opts.baseName}.pdf`, pdf, "application/pdf");
      return { ...saved, message: note };
    });
  }

  if (opts.eml) {
    await run("E-Mail als EML", async () => {
      if (!host.emlSupported) throw new Error("Diese Outlook-Version kann die E-Mail nicht als Datei liefern (Mailbox 1.14 nötig).");
      const b64 = await host.getEmlBase64();
      return drive.upload(folder, `${opts.baseName}.eml`, base64ToBytes(b64), "message/rfc822");
    });
  }

  for (const a of mail.attachments.filter((x) => opts.attachmentIds.includes(x.id))) {
    await run(`Anlage ${a.name}`, async () => {
      const data = await host.getAttachment(a.id);
      switch (data.format) {
        case "base64": {
          const bytes = base64ToBytes(data.content);
          return drive.upload(folder, attachmentFileName(a.name), bytes, a.contentType);
        }
        case "eml":
          return drive.upload(folder, attachmentFileName(a.name, ".eml"), utf8Bytes(data.content), "message/rfc822");
        case "iCalendar":
          return drive.upload(folder, attachmentFileName(a.name, ".ics"), utf8Bytes(data.content), "text/calendar");
        case "url": {
          // Cloud-Anlage: Verknüpfung statt Datei, der Inhalt liegt bereits in OneDrive/SharePoint.
          const shortcut = `[InternetShortcut]\r\nURL=${data.content}\r\n`;
          const saved = await drive.upload(folder, attachmentFileName(a.name, ".url"), utf8Bytes(shortcut), "text/plain");
          return { ...saved, message: "Cloud-Anlage – als Verknüpfung gespeichert." };
        }
        default:
          throw new Error(`Unbekanntes Anlagenformat "${data.format}".`);
      }
    });
  }

  progress("");
  return results;
}

export const formatSize = (bytes: number) =>
  bytes < 1024 ? `${bytes} B` : bytes < 1024 * 1024 ? `${Math.round(bytes / 1024)} KB` : `${(bytes / 1024 / 1024).toFixed(1)} MB`;

