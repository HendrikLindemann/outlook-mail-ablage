import { describe, expect, it } from "vitest";
import { attachmentFileName, buildBaseName, formatDate, sanitizeFileName } from "../src/core/names";
import { buildPrintHtml, dropDataImages, extractBody, replaceInlineImages, stripActiveContent } from "../src/core/printhtml";
import { base64Size, base64ToBytes } from "../src/core/base64";

describe("Dateinamen", () => {
  it("formatiert das Datum wie VBA yy-mm-dd", () => {
    expect(formatDate(new Date(2026, 0, 5))).toBe("26-01-05");
  });

  it("baut den Namen aus Datum, Absender und Betreff", () => {
    expect(buildBaseName(new Date(2026, 10, 20), "Müller, Erika", "Re: Angebot 4711/2026?")).toBe(
      "26-11-20 Müller, Erika - Re Angebot 47112026",
    );
  });

  it("ersetzt leere Angaben", () => {
    expect(buildBaseName(new Date(2026, 10, 20), " ", "")).toBe("26-11-20 Unbekannt - (ohne Betreff)");
  });

  it("entfernt unzulässige Zeichen, Punkte am Rand und kürzt", () => {
    expect(sanitizeFileName(' a<b>c:"d|e*f\\g/h. ')).toBe("abcdefgh");
    expect(sanitizeFileName("x".repeat(300)).length).toBe(120);
    expect(sanitizeFileName("...")).toBe("E-Mail");
    expect(sanitizeFileName("CON")).toBe("_CON");
  });

  it("behält Endungen von Anlagen und erzwingt sie bei Bedarf", () => {
    expect(attachmentFileName("Rechnung: 11/2026.pdf")).toBe("Rechnung 112026.pdf");
    expect(attachmentFileName("Weitergeleitet: Termin")).toBe("Weitergeleitet Termin");
    expect(attachmentFileName("Weitergeleitet: Termin", ".eml")).toBe("Weitergeleitet Termin.eml");
    expect(attachmentFileName("nachricht.EML", ".eml")).toBe("nachricht.EML");
  });
});

describe("Druck-HTML", () => {
  const mail = {
    subject: "Test <1>",
    from: { name: "Erika", email: "e@example.com" },
    to: [{ name: "Max", email: "m@example.com" }],
    cc: [],
    date: new Date(2026, 10, 20, 9, 30),
    attachmentNames: ["a.pdf"],
  };

  it("enthält Kopfdaten escaped und den Text", () => {
    const html = buildPrintHtml(mail, "<html><head><style>p{}</style></head><body><p>Hallo</p></body></html>");
    expect(html).toContain("Test &lt;1&gt;");
    expect(html).toContain("Erika &lt;e@example.com&gt;");
    expect(html).toContain("<p>Hallo</p>");
    expect(html).toContain("<style>p{}</style>");
    expect(html).toContain("a.pdf");
    expect(html).not.toContain("Cc:");
  });

  it("entfernt Skripte und Ereignisattribute", () => {
    const s = stripActiveContent('<p onclick="x()">a</p><script>alert(1)</script><iframe src="x"></iframe>');
    expect(s).toBe("<p>a</p>");
  });

  it("nimmt Fragmente ohne body unverändert", () => {
    expect(extractBody("<p>x</p>").body).toBe("<p>x</p>");
  });

  it("ersetzt cid-Verweise per Content-ID oder Dateiname", () => {
    const { html, missing } = replaceInlineImages(
      '<img src="cid:image001.png@01DB"><img src=\'cid:abc\'><img src="cid:fehlt">',
      [
        { name: "image001.png", dataUrl: "data:image/png;base64,AAA" },
        { name: "logo.jpg", contentId: "<abc>", dataUrl: "data:image/jpeg;base64,BBB" },
      ],
    );
    expect(html).toBe('<img src="data:image/png;base64,AAA"><img src=\'data:image/jpeg;base64,BBB\'><img src="">');
    expect(missing).toBe(1);
  });

  it("entfernt eingebettete Bilder", () => {
    expect(dropDataImages('<p><img alt="x" src="data:image/png;base64,AAA"></p>')).toBe("<p><span>[Bild entfernt]</span></p>");
  });
});

describe("Base64", () => {
  it("dekodiert und schätzt die Größe", () => {
    expect([...base64ToBytes("QUJD")]).toEqual([65, 66, 67]);
    expect(base64Size("QUI=")).toBe(2);
  });
});
