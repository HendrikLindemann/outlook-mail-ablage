/** Kapselt alle Zugriffe auf Outlook, damit die Oberfläche auch im Browser (Demo-Modus) läuft. */

export interface Person {
  name: string;
  email: string;
}

export type AttachmentKind = "file" | "item" | "cloud";

export interface AttachmentInfo {
  id: string;
  name: string;
  size: number;
  kind: AttachmentKind;
  isInline: boolean;
  contentId?: string;
  contentType?: string;
}

export interface MailInfo {
  subject: string;
  from: Person;
  to: Person[];
  cc: Person[];
  date: Date;
  attachments: AttachmentInfo[];
}

export interface AttachmentData {
  /** "base64" | "eml" | "iCalendar" | "url" (Werte von Office.MailboxEnums.AttachmentContentFormat) */
  format: string;
  content: string;
}

export type Platform = "desktop" | "web" | "mac" | "other";

export interface Host {
  inOutlook: boolean;
  naaSupported: boolean;
  platform: Platform;
  /** getAsFileAsync (Mailbox 1.14) verfügbar */
  emlSupported: boolean;
  getMail(): MailInfo;
  getBodyHtml(): Promise<string>;
  /** Ganze E-Mail als Base64-EML */
  getEmlBase64(): Promise<string>;
  getAttachment(id: string): Promise<AttachmentData>;
  roamingGet(key: string): unknown;
  roamingSet(key: string, value: unknown): Promise<void>;
  /** Meldet einen Wechsel der markierten E-Mail (angeheftete Aufgabenbereiche). */
  onItemChanged(handler: () => void): void;
}

function promisify<T>(call: (cb: (r: Office.AsyncResult<T>) => void) => void): Promise<T> {
  return new Promise((resolve, reject) =>
    call((r) =>
      r.status === Office.AsyncResultStatus.Succeeded
        ? resolve(r.value)
        : reject(new Error(r.error?.message || "Outlook hat die Anfrage abgelehnt.")),
    ),
  );
}

const toPerson = (p?: Office.EmailAddressDetails | null): Person => ({
  name: p?.displayName || "",
  email: p?.emailAddress || "",
});

function platformOf(): Platform {
  const p = Office.context.diagnostics?.platform as unknown as string | undefined;
  if (p === "PC") return "desktop";
  if (p === "OfficeOnline") return "web";
  if (p === "Mac") return "mac";
  return "other";
}

function outlookHost(): Host {
  const mailbox = Office.context.mailbox;
  const item = () => {
    const i = mailbox.item as Office.MessageRead | undefined;
    if (!i || i.itemType !== Office.MailboxEnums.ItemType.Message) {
      throw new Error("Bitte markieren oder öffnen Sie eine empfangene E-Mail.");
    }
    return i;
  };
  const settings = Office.context.roamingSettings;
  const req = Office.context.requirements;
  return {
    inOutlook: true,
    naaSupported: req.isSetSupported("NestedAppAuth", "1.1"),
    platform: platformOf(),
    emlSupported: req.isSetSupported("Mailbox", "1.14"),
    getMail: () => {
      const i = item();
      return {
        subject: i.subject || "",
        // "from" ist der Absender wie in VBA SenderName (bei "im Auftrag von" der Vertretene).
        from: toPerson(i.from || i.sender),
        to: (i.to || []).map(toPerson),
        cc: (i.cc || []).map(toPerson),
        date: new Date(i.dateTimeCreated),
        attachments: (i.attachments || []).map((a) => ({
          id: a.id,
          name: a.name,
          size: a.size,
          kind: (a.attachmentType as string) as AttachmentKind,
          isInline: !!a.isInline,
          contentId: (a as { contentId?: string }).contentId || undefined,
          contentType: a.contentType,
        })),
      };
    },
    getBodyHtml: () => promisify<string>((cb) => item().body.getAsync(Office.CoercionType.Html, cb)),
    getEmlBase64: () => promisify<string>((cb) => item().getAsFileAsync(cb)),
    getAttachment: async (id) => {
      const r = await promisify<Office.AttachmentContent>((cb) => item().getAttachmentContentAsync(id, cb));
      return { format: r.format as string, content: r.content };
    },
    roamingGet: (key) => settings.get(key),
    roamingSet: (key, value) => {
      if (value === undefined) settings.remove(key);
      else settings.set(key, value);
      return promisify<void>((cb) => settings.saveAsync(cb));
    },
    onItemChanged: (handler) => {
      try {
        mailbox.addHandlerAsync(Office.EventType.ItemChanged, handler);
      } catch {
        /* ältere Clients ohne Anheften */
      }
    },
  };
}

// ---------------------------------------------------------------------------
// Demo-Host außerhalb von Outlook
// ---------------------------------------------------------------------------

const PIXEL =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkaPhfDwAEhgGAWn5BqwAAAABJRU5ErkJggg==";

function demoHost(): Host {
  const KEY = "mailablage.demo.";
  const files: Record<string, AttachmentData> = {
    a1: { format: "base64", content: btoa("%PDF-1.4\n% Demo-Rechnung\n") },
    a2: { format: "base64", content: btoa("Spalte1;Spalte2\n1;2\n") },
    a3: { format: "base64", content: PIXEL },
  };
  return {
    inOutlook: false,
    naaSupported: false,
    platform: "other",
    emlSupported: true,
    getMail: () => ({
      subject: "Re: Angebot Nr. 4711 / Gutachten",
      from: { name: "Müller, Erika", email: "erika.mueller@example.com" },
      to: [{ name: "Max Beispiel", email: "max.beispiel@example.com" }],
      cc: [],
      date: new Date(2026, 10, 20, 9, 30),
      attachments: [
        { id: "a1", name: "Rechnung 2026-11.pdf", size: 24, kind: "file", isInline: false },
        { id: "a2", name: "Aufstellung.csv", size: 20, kind: "file", isInline: false },
        { id: "a3", name: "image001.png", size: 68, kind: "file", isInline: true, contentId: "image001.png@01DB" },
      ],
    }),
    getBodyHtml: async () =>
      '<html><body><p>Guten Tag,</p><p>anbei die <b>Unterlagen</b>.</p><p><img src="cid:image001.png@01DB"></p></body></html>',
    getEmlBase64: async () => btoa("From: erika.mueller@example.com\r\nSubject: Demo\r\n\r\nDemo"),
    getAttachment: async (id) => {
      const f = files[id];
      if (!f) throw new Error("Anlage nicht gefunden.");
      return f;
    },
    roamingGet: (key) => {
      try {
        const raw = localStorage.getItem(KEY + key);
        return raw ? JSON.parse(raw) : undefined;
      } catch {
        return undefined;
      }
    },
    roamingSet: async (key, value) => {
      try {
        if (value === undefined) localStorage.removeItem(KEY + key);
        else localStorage.setItem(KEY + key, JSON.stringify(value));
      } catch {
        /* Speicher gesperrt */
      }
    },
    onItemChanged: () => {},
  };
}

export function initHost(): Promise<Host> {
  if (typeof Office === "undefined") return Promise.resolve(demoHost());
  return new Promise((resolve) => {
    const fallback = setTimeout(() => resolve(demoHost()), 4000);
    Office.onReady((info) => {
      clearTimeout(fallback);
      resolve(info.host === Office.HostType.Outlook ? outlookHost() : demoHost());
    });
  });
}
