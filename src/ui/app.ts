import { buildBaseName } from "../core/names";
import { defaultAttachments, formatSize, saveMail, type SaveResult } from "../core/save";
import { AuthRequiredError, type DriveApi, type FolderRef } from "../drive/types";
import type { Host, MailInfo } from "../host";
import { errorText, field, h, toast } from "./dom";
import { createPicker } from "./picker";

const RECENT_KEY = "mailablage.recent";
const OPTIONS_KEY = "mailablage.options";
const MAX_RECENT = 8;

interface StoredOptions {
  pdf: boolean;
  eml: boolean;
}

export function createApp(host: Host, drive: DriveApi, root: HTMLElement, setupHint?: string) {
  const recent = (): FolderRef[] => {
    const v = host.roamingGet(RECENT_KEY);
    return Array.isArray(v) ? (v as FolderRef[]).filter((f) => f && f.driveId && f.itemId) : [];
  };
  const remember = async (f: FolderRef) => {
    const list = [f, ...recent().filter((r) => !(r.driveId === f.driveId && r.itemId === f.itemId))].slice(0, MAX_RECENT);
    await host.roamingSet(RECENT_KEY, list).catch(() => undefined);
  };
  const storedOptions = (): StoredOptions => {
    const v = host.roamingGet(OPTIONS_KEY) as Partial<StoredOptions> | undefined;
    return { pdf: v?.pdf ?? true, eml: v?.eml ?? true };
  };

  let target: FolderRef | undefined = recent()[0];

  /** Führt eine Graph-Aktion aus und bietet bei fehlender Anmeldung eine Schaltfläche an. */
  async function withAuth<T>(work: () => Promise<T>, onRetry: () => void): Promise<T | undefined> {
    try {
      return await work();
    } catch (e) {
      if (e instanceof AuthRequiredError) {
        showSignIn(e.message, onRetry);
        return undefined;
      }
      throw e;
    }
  }

  function showSignIn(message: string, onRetry: () => void) {
    const box = h(
      "div",
      { class: "view" },
      h("header", {}, h("h1", {}, "Anmeldung")),
      h(
        "div",
        { class: "notice" },
        h("p", {}, message),
        h(
          "button",
          {
            class: "primary",
            onclick: async () => {
              try {
                await drive.signIn();
                onRetry();
              } catch (err) {
                toast(errorText(err), "error");
              }
            },
          },
          "Anmelden",
        ),
      ),
      h("button", { class: "link", onclick: () => showMain() }, "Zurück"),
    );
    root.replaceChildren(box);
  }

  function showPicker() {
    createPicker({
      root,
      drive,
      recent: recent(),
      withAuth,
      onCancel: () => showMain(),
      onPick: (f) => {
        target = f;
        showMain();
      },
    }).open();
  }

  let lastMailKey = "";

  function showMain(results?: SaveResult[]) {
    let mail: MailInfo;
    try {
      mail = host.getMail();
    } catch (e) {
      root.replaceChildren(h("div", { class: "view" }, h("div", { class: "notice" }, errorText(e))));
      return;
    }
    lastMailKey = `${mail.subject}|${mail.date.getTime()}`;
    const opts = storedOptions();

    const nameInput = h("input", { type: "text", value: buildBaseName(mail.date, mail.from.name || mail.from.email, mail.subject) });
    const pdfBox = h("input", { type: "checkbox", id: "opt-pdf", checked: opts.pdf });
    const emlBox = h("input", { type: "checkbox", id: "opt-eml", checked: opts.eml, disabled: !host.emlSupported });

    const preselected = new Set(defaultAttachments(mail).map((a) => a.id));
    const attBoxes = mail.attachments.map((a) => {
      const box = h("input", { type: "checkbox", id: `att-${a.id}`, checked: preselected.has(a.id) });
      box.dataset.id = a.id;
      const extra = [a.isInline ? "eingebettet" : "", a.kind === "cloud" ? "Cloud-Anlage" : "", a.kind === "item" ? "E-Mail/Element" : ""]
        .filter(Boolean)
        .join(", ");
      return {
        box,
        row: h(
          "label",
          { class: "check", htmlFor: box.id },
          box,
          h("span", { class: "att-name" }, a.name, h("span", { class: "muted" }, ` · ${formatSize(a.size)}${extra ? ` · ${extra}` : ""}`)),
        ),
      };
    });

    const recentSelect = h("select", { "aria-label": "Zuletzt verwendete Ordner" });
    const recentList = recent();
    recentSelect.append(h("option", { value: "" }, recentList.length ? "Zuletzt verwendet …" : "Noch keine Ordner verwendet"));
    recentList.forEach((f, i) => recentSelect.append(h("option", { value: String(i) }, f.label)));
    recentSelect.disabled = !recentList.length;
    recentSelect.addEventListener("change", () => {
      const f = recentList[Number(recentSelect.value)];
      if (f) {
        target = f;
        showMain();
      }
    });

    const progress = h("p", { class: "muted", role: "status" });
    const saveBtn = h("button", { class: "primary", id: "save" }, "Ablegen");

    saveBtn.addEventListener("click", async () => {
      if (!target) {
        toast("Bitte zuerst einen Zielordner wählen.", "error");
        return;
      }
      const baseName = nameInput.value.trim();
      if (!baseName) {
        toast("Bitte einen Dateinamen angeben.", "error");
        return;
      }
      const attachmentIds = attBoxes.filter((x) => x.box.checked).map((x) => x.box.dataset.id!);
      if (!pdfBox.checked && !emlBox.checked && !attachmentIds.length) {
        toast("Es ist nichts zum Speichern ausgewählt.", "error");
        return;
      }
      await host.roamingSet(OPTIONS_KEY, { pdf: pdfBox.checked, eml: emlBox.checked }).catch(() => undefined);
      saveBtn.disabled = true;
      const folder = target;
      try {
        const res = await withAuth(
          () =>
            saveMail(host, drive, folder, { baseName, pdf: pdfBox.checked, eml: emlBox.checked && host.emlSupported, attachmentIds }, (t) => {
              progress.textContent = t;
            }),
          () => showMain(),
        );
        if (!res) return;
        // Fehlt die Anmeldung erst beim Hochladen, kommt sie als Teilfehler zurück.
        if (res.length && res.every((r) => !r.ok) && /Anmeldung/.test(res[0].message || "")) {
          showSignIn(res[0].message!, () => showMain());
          return;
        }
        if (res.some((r) => r.ok)) await remember(folder);
        showMain(res);
      } catch (e) {
        toast(errorText(e), "error");
      } finally {
        saveBtn.disabled = false;
        progress.textContent = "";
      }
    });

    const targetBlock = h(
      "div",
      { class: "target" },
      h("div", { class: "target-label", id: "target-label" }, target ? target.label : h("span", { class: "muted" }, "Kein Ordner gewählt")),
      h("button", { onclick: () => showPicker(), id: "pick" }, "Ordner wählen …"),
    );

    const resultBlock = results
      ? h(
          "section",
          { class: "results", "aria-label": "Ergebnis" },
          h("h2", {}, results.every((r) => r.ok) ? "Abgelegt" : "Ergebnis"),
          ...results.map((r) =>
            h(
              "div",
              { class: `result ${r.ok ? "ok" : "fail"}` },
              h("span", { class: "mark", "aria-hidden": "true" }, r.ok ? "✓" : "✗"),
              h(
                "div",
                {},
                r.ok && r.webUrl ? h("a", { href: r.webUrl, target: "_blank", rel: "noopener" }, r.name || r.label) : h("span", {}, r.name || r.label),
                r.message ? h("div", { class: "muted" }, r.message) : null,
              ),
            ),
          ),
          target?.webUrl ? h("a", { href: target.webUrl, target: "_blank", rel: "noopener", class: "open-folder" }, "Ordner im Browser öffnen") : null,
        )
      : null;

    const view = h(
      "div",
      { class: "view" },
      h("header", {}, h("h1", {}, "Mail-Ablage")),
      setupHint ? h("div", { class: "notice" }, setupHint) : null,
      h(
        "section",
        { class: "mail" },
        h("div", { class: "mail-subject" }, mail.subject || "(ohne Betreff)"),
        h("div", { class: "muted" }, `${mail.from.name || mail.from.email} · ${mail.date.toLocaleString("de-DE")}`),
      ),
      resultBlock,
      h("h2", {}, "Ziel"),
      targetBlock,
      recentSelect,
      host.platform === "desktop"
        ? h("p", { class: "muted" }, "Tipp: Ordner, die der OneDrive-Client synchronisiert, erscheinen auch im Windows-Explorer.")
        : null,
      h("h2", {}, "Speichern als"),
      field("Dateiname (ohne Endung)", nameInput),
      h("label", { class: "check", htmlFor: pdfBox.id }, pdfBox, "E-Mail als PDF"),
      h(
        "label",
        { class: "check", htmlFor: emlBox.id },
        emlBox,
        host.emlSupported ? "E-Mail im Original (.eml)" : "E-Mail im Original (.eml) – von dieser Outlook-Version nicht unterstützt",
      ),
      h("h2", {}, `Anlagen (${mail.attachments.length})`),
      attBoxes.length ? h("div", { class: "list" }, ...attBoxes.map((x) => x.row)) : h("p", { class: "muted" }, "Diese E-Mail hat keine Anlagen."),
      h("div", { class: "actions" }, saveBtn),
      progress,
    );
    root.replaceChildren(view);
  }

  return {
    start() {
      showMain();
      host.onItemChanged(() => {
        try {
          const m = host.getMail();
          if (`${m.subject}|${m.date.getTime()}` !== lastMailKey) showMain();
        } catch {
          showMain();
        }
      });
    },
  };
}
