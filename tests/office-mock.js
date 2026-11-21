// Simuliert Office.context.mailbox im Lesemodus für Durchlauftests (ersetzt office.js).
(() => {
  const ok = (value) => ({ status: "succeeded", value });
  const fail = (message) => ({ status: "failed", error: { message } });
  const state = (window.__mail = { settings: JSON.parse(sessionStorage.getItem("rs") || "{}") });
  const cb = (args) => args[args.length - 1];
  const contents = {
    att1: { format: "base64", content: btoa("%PDF-1.4 Anlage") },
    att2: { format: "eml", content: "Subject: Weitergeleitet\r\n\r\nText" },
    img1: { format: "base64", content: "iVBORw0KGgo=" },
  };
  window.Office = {
    AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
    CoercionType: { Html: "html" },
    HostType: { Outlook: "Outlook" },
    EventType: { ItemChanged: "olkItemSelectedChanged" },
    MailboxEnums: { ItemType: { Message: "message" } },
    onReady: (f) => {
      const r = { host: "Outlook" };
      f && f(r);
      return Promise.resolve(r);
    },
    context: {
      diagnostics: { platform: "PC" },
      requirements: { isSetSupported: (name, v) => name === "Mailbox" },
      roamingSettings: {
        get: (k) => state.settings[k],
        set: (k, v) => { state.settings[k] = v; },
        remove: (k) => { delete state.settings[k]; },
        saveAsync: (f) => { sessionStorage.setItem("rs", JSON.stringify(state.settings)); f(ok()); },
      },
      mailbox: {
        userProfile: { displayName: "Max Beispiel", emailAddress: "max@example.com" },
        addHandlerAsync: () => {},
        item: {
          itemType: "message",
          subject: "AW: Ortstermin 12.12.? <Gutachten>",
          from: { displayName: "Dr. Erika Müller", emailAddress: "erika@example.com" },
          to: [{ displayName: "Max Beispiel", emailAddress: "max@example.com" }],
          cc: [],
          dateTimeCreated: new Date(2026, 10, 19, 14, 5).toISOString(),
          attachments: [
            { id: "att1", name: "Lageplan.pdf", size: 120000, attachmentType: "file", isInline: false, contentType: "application/pdf" },
            { id: "att2", name: "Weitergeleitet: Termin", size: 4000, attachmentType: "item", isInline: false },
            { id: "img1", name: "image001.png", size: 900, attachmentType: "file", isInline: true, contentType: "image/png" },
          ],
          body: {
            getAsync: (...a) => cb(a)(ok('<html><body><p>Sehr geehrter Herr Beispiel,</p><img src="cid:image001.png@01DB0000"></body></html>')),
          },
          getAsFileAsync: (...a) => cb(a)(ok(btoa("Subject: Ortstermin\r\n\r\nText"))),
          getAttachmentContentAsync: (id, ...a) => {
            const c = contents[id];
            cb(a)(c ? ok(c) : fail("nicht gefunden"));
          },
        },
      },
    },
  };
})();
