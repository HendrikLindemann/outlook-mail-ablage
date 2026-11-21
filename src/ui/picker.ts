import type { DriveApi, Entry, FolderRef } from "../drive/types";
import { errorText, field, h, toast } from "./dom";

interface PickerOptions {
  root: HTMLElement;
  drive: DriveApi;
  recent: FolderRef[];
  withAuth<T>(work: () => Promise<T>, onRetry: () => void): Promise<T | undefined>;
  onPick(folder: FolderRef): void;
  onCancel(): void;
}

/** Eine Ebene der Navigation: Startseite, Website (Bibliotheken) oder Ordner. */
type Level =
  | { type: "start" }
  | { type: "site"; siteId: string; name: string }
  | { type: "folder"; folder: FolderRef; name: string };

export function createPicker(o: PickerOptions) {
  const stack: Level[] = [{ type: "start" }];
  let siteQuery = "";

  const current = () => stack[stack.length - 1];
  const pathLabel = () =>
    stack
      .slice(1)
      .map((l) => (l.type === "start" ? "" : l.name))
      .join(" / ");

  const push = (l: Level) => {
    stack.push(l);
    render();
  };

  function row(label: string, action: () => void | Promise<void>, meta?: string, icon = "📁") {
    const onclick = async () => {
      try {
        await action();
      } catch (e) {
        toast(errorText(e), "error");
      }
    };
    return h(
      "div",
      { class: "row" },
      h("button", { class: "row-main", onclick }, h("span", { "aria-hidden": "true" }, `${icon} `), label, meta ? h("span", { class: "muted" }, ` · ${meta}`) : null),
    );
  }

  const openFolderEntry = (e: Entry, base: FolderRef) =>
    push({
      type: "folder",
      name: e.name,
      folder: { driveId: e.driveId || base.driveId, itemId: e.id, label: `${base.label} / ${e.name}`, webUrl: e.webUrl },
    });

  async function loadList(level: Level): Promise<HTMLElement[]> {
    const d = o.drive;
    if (level.type === "start") {
      const rows: HTMLElement[] = [];
      rows.push(
        row(
          "Mein OneDrive",
          async () => {
            const me = await o.withAuth(() => d.myDrive(), render);
            if (me) push({ type: "folder", folder: me, name: "Mein OneDrive" });
          },
          undefined,
          "☁️",
        ),
      );
      if (o.recent.length) {
        rows.push(h("h2", {}, "Zuletzt verwendet"));
        for (const f of o.recent) rows.push(row(f.label, () => o.onPick(f), undefined, "🕘"));
      }
      rows.push(h("h2", {}, "SharePoint-Websites"));
      const search = h("input", { type: "search", value: siteQuery, placeholder: "Website suchen …" });
      const form = h(
        "form",
        {
          class: "link-row",
          onsubmit: (ev: Event) => {
            ev.preventDefault();
            siteQuery = search.value.trim();
            render();
          },
        },
        field("Website suchen", search),
        h("button", { type: "submit" }, "Suchen"),
      );
      rows.push(form);
      const sites = await o.withAuth(() => (siteQuery ? d.searchSites(siteQuery) : d.followedSites()), render);
      if (sites === undefined) return [];
      if (!sites.length) {
        rows.push(
          h("p", { class: "muted" }, siteQuery ? "Keine Website gefunden." : "Sie folgen noch keiner Website. Suchen Sie oben nach einer Website."),
        );
      }
      for (const s of sites) rows.push(row(s.name, () => push({ type: "site", siteId: s.id, name: s.name }), undefined, "🏢"));
      return rows;
    }
    if (level.type === "site") {
      const drives = await o.withAuth(() => d.siteDrives(level.siteId), render);
      if (drives === undefined) return [];
      if (!drives.length) return [h("p", { class: "muted" }, "Keine Dokumentbibliotheken gefunden.")];
      return drives.map((dr) =>
        row(
          dr.name,
          () =>
            push({
              type: "folder",
              name: dr.name,
              folder: { driveId: dr.id, itemId: "root", label: `${level.name} / ${dr.name}`, webUrl: dr.webUrl },
            }),
          undefined,
          "📚",
        ),
      );
    }
    const folders = await o.withAuth(() => d.childFolders(level.folder.driveId, level.folder.itemId), render);
    if (folders === undefined) return [];
    if (!folders.length) return [h("p", { class: "muted" }, "Keine Unterordner.")];
    return folders.map((f) => row(f.name, () => openFolderEntry(f, level.folder)));
  }

  async function render() {
    const level = current();
    const list = h("div", { class: "list", "aria-busy": "true" }, h("p", { class: "muted" }, "Wird geladen …"));
    const footer = h("div", { class: "actions" });

    if (level.type === "folder") {
      const nameInput = h("input", { type: "text", placeholder: "Name des neuen Ordners" });
      const newRow = h(
        "form",
        {
          class: "link-row",
          hidden: true,
          onsubmit: async (ev: Event) => {
            ev.preventDefault();
            const name = nameInput.value.trim();
            if (!name) return;
            try {
              const created = await o.withAuth(() => o.drive.createFolder(level.folder.driveId, level.folder.itemId, name), render);
              if (created) openFolderEntry(created, level.folder);
            } catch (e) {
              toast(errorText(e), "error");
            }
          },
        },
        field("Neuer Ordner", nameInput),
        h("button", { type: "submit" }, "Anlegen"),
      );
      footer.append(
        h("button", { class: "primary", id: "choose", onclick: () => o.onPick(level.folder) }, "Diesen Ordner wählen"),
        h(
          "button",
          {
            onclick: () => {
              newRow.hidden = false;
              nameInput.focus();
            },
          },
          "Neuer Ordner …",
        ),
      );
      o.root.replaceChildren(view(list, footer, newRow));
    } else {
      o.root.replaceChildren(view(list, footer));
    }

    try {
      const rows = await loadList(level);
      if (current() !== level) return; // inzwischen weiter navigiert
      list.replaceChildren(...rows);
    } catch (e) {
      list.replaceChildren(h("div", { class: "notice" }, errorText(e)));
    } finally {
      list.removeAttribute("aria-busy");
    }
  }

  function view(list: HTMLElement, footer: HTMLElement, extra?: HTMLElement) {
    return h(
      "div",
      { class: "view" },
      h(
        "header",
        {},
        h(
          "button",
          {
            class: "icon",
            "aria-label": "Zurück",
            onclick: () => {
              if (stack.length > 1) {
                stack.pop();
                render();
              } else o.onCancel();
            },
          },
          "‹",
        ),
        h("h1", {}, stack.length > 1 ? pathLabel() : "Ordner wählen"),
        h("button", { class: "link", onclick: () => o.onCancel() }, "Abbrechen"),
      ),
      o.drive.demo ? h("p", { class: "muted" }, "Demo-Daten") : null,
      list,
      footer,
      extra,
    );
  }

  return { open: render };
}
