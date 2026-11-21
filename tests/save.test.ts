import { describe, expect, it } from "vitest";
import { defaultAttachments, saveMail } from "../src/core/save";
import { MemoryDrive } from "../src/drive/memory";
import type { AttachmentData, Host, MailInfo } from "../src/host";

function fakeHost(overrides: Partial<Host> = {}, attachments: Record<string, AttachmentData> = {}): Host {
  const mail: MailInfo = {
    subject: "Angebot",
    from: { name: "Erika", email: "e@example.com" },
    to: [],
    cc: [],
    date: new Date(2026, 10, 20),
    attachments: [
      { id: "a1", name: "Rechnung.pdf", size: 3, kind: "file", isInline: false },
      { id: "a2", name: "Rechnung.pdf", size: 3, kind: "file", isInline: false },
      { id: "a3", name: "Weiterleitung", size: 3, kind: "item", isInline: false },
      { id: "a4", name: "Plan.docx", size: 3, kind: "cloud", isInline: false },
      { id: "img", name: "image001.png", size: 3, kind: "file", isInline: true, contentId: "image001.png@x" },
    ],
  };
  const data: Record<string, AttachmentData> = {
    a1: { format: "base64", content: "QUJD" },
    a2: { format: "base64", content: "REVG" },
    a3: { format: "eml", content: "Subject: x\r\n\r\ny" },
    a4: { format: "url", content: "https://example.invalid/Plan.docx" },
    img: { format: "base64", content: "iVBO" },
    ...attachments,
  };
  return {
    inOutlook: true,
    naaSupported: true,
    platform: "web",
    emlSupported: true,
    getMail: () => mail,
    getBodyHtml: async () => '<p>Hallo</p><img src="cid:image001.png@x">',
    getEmlBase64: async () => "RU1M",
    getAttachment: async (id) => {
      if (!data[id]) throw new Error("weg");
      return data[id];
    },
    roamingGet: () => undefined,
    roamingSet: async () => {},
    onItemChanged: () => {},
    ...overrides,
  };
}

const folder = { driveId: "me", itemId: "p1", label: "Projekte" };

describe("saveMail", () => {
  it("legt PDF, EML und alle gewählten Anlagen im selben Ordner ab", async () => {
    const drive = new MemoryDrive();
    const host = fakeHost();
    const ids = defaultAttachments(host.getMail()).map((a) => a.id);
    expect(ids).toEqual(["a1", "a2", "a3", "a4"]);

    const results = await saveMail(host, drive, folder, { baseName: "26-11-20 Erika - Angebot", pdf: true, eml: true, attachmentIds: ids });

    expect(results.every((r) => r.ok)).toBe(true);
    const names = drive.filesIn("me", "p1").map((n) => n.name).sort();
    expect(names).toEqual([
      "26-11-20 Erika - Angebot.eml",
      "26-11-20 Erika - Angebot.pdf",
      "Plan.docx.url",
      "Rechnung 1.pdf",
      "Rechnung.pdf",
      "Weiterleitung.eml",
    ]);
    // Inline-Bild ist im PDF eingebettet
    expect(drive.pdfSources[0]).toContain("data:image/png;base64,iVBO");
  });

  it("macht bei Fehlern einzelner Teile weiter", async () => {
    const drive = new MemoryDrive();
    const host = fakeHost({ emlSupported: false });
    const results = await saveMail(host, drive, folder, { baseName: "x", pdf: false, eml: true, attachmentIds: ["a1", "fehlt"] });
    expect(results.map((r) => [r.label, r.ok])).toEqual([
      ["E-Mail als EML", false],
      ["Anlage Rechnung.pdf", true],
    ]);
  });

  it("meldet nicht ladbare Anlagen", async () => {
    const drive = new MemoryDrive();
    const host = fakeHost({ getAttachment: async () => Promise.reject(new Error("Drag-and-drop-Element")) });
    const [r] = await saveMail(host, drive, folder, { baseName: "x", pdf: false, eml: false, attachmentIds: ["a3"] });
    expect(r.ok).toBe(false);
    expect(r.message).toContain("Drag-and-drop");
  });
});
