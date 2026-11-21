import { base64ToBytes, utf8Bytes } from "../core/base64";
import type { DriveApi, Entry, FolderRef, SavedFile } from "./types";

const GRAPH = "https://graph.microsoft.com/v1.0";
/** Bis zu dieser Größe genügt ein einfacher PUT (Graph-Grenze 4 MB). */
const SIMPLE_LIMIT = 3.5 * 1024 * 1024;
/** Blockgröße für Upload-Sitzungen: Vielfaches von 320 KiB. */
const CHUNK = 10 * 320 * 1024;
/** Graph wandelt nur Dateien bis 2 MB um. */
export const PDF_SOURCE_LIMIT = 2 * 1024 * 1024;

type TokenFn = (interactive: boolean) => Promise<string>;

interface GraphItem {
  id: string;
  name: string;
  webUrl?: string;
  folder?: unknown;
  parentReference?: { driveId?: string };
}

const enc = encodeURIComponent;

export class GraphDrive implements DriveApi {
  demo = false;
  constructor(private getToken: TokenFn) {}

  async signIn() {
    await this.getToken(true);
  }

  private async fetch(path: string, init: RequestInit = {}): Promise<Response> {
    const token = await this.getToken(false);
    const url = path.startsWith("https://") ? path : GRAPH + path;
    const res = await fetch(url, {
      ...init,
      headers: { ...(init.headers || {}), Authorization: `Bearer ${token}` },
    });
    if (!res.ok) {
      let detail = "";
      try {
        detail = ((await res.json()) as { error?: { message?: string } }).error?.message || "";
      } catch {
        /* kein JSON */
      }
      const hint =
        res.status === 403
          ? "Keine Berechtigung für diesen Ort."
          : res.status === 404
            ? "Ordner oder Datei nicht gefunden."
            : `Microsoft Graph meldet Fehler ${res.status}.`;
      throw new Error(detail ? `${hint} (${detail})` : hint);
    }
    return res;
  }

  private async json<T>(path: string, init?: RequestInit): Promise<T> {
    return (await (await this.fetch(path, init)).json()) as T;
  }

  /** Folgt @odata.nextLink, damit auch große Ordner vollständig erscheinen. */
  private async list<T>(path: string, max = 1000): Promise<T[]> {
    const out: T[] = [];
    let next: string | undefined = path;
    while (next && out.length < max) {
      const page: { value: T[]; "@odata.nextLink"?: string } = await this.json(next);
      out.push(...page.value);
      next = page["@odata.nextLink"];
    }
    return out;
  }

  async myDrive(): Promise<FolderRef> {
    const d = await this.json<{ id: string; webUrl?: string }>("/me/drive?$select=id,webUrl");
    return { driveId: d.id, itemId: "root", label: "Mein OneDrive", webUrl: d.webUrl };
  }

  async followedSites(): Promise<Entry[]> {
    const sites = await this.list<{ id: string; displayName?: string; name?: string; webUrl?: string }>(
      "/me/followedSites?$select=id,displayName,name,webUrl",
    );
    return sites.map((s) => ({ kind: "site", id: s.id, name: s.displayName || s.name || s.webUrl || s.id, webUrl: s.webUrl }));
  }

  async searchSites(query: string): Promise<Entry[]> {
    const sites = await this.json<{ value: { id: string; displayName?: string; name?: string; webUrl?: string }[] }>(
      `/sites?search=${enc(query)}&$select=id,displayName,name,webUrl&$top=50`,
    );
    return sites.value.map((s) => ({ kind: "site", id: s.id, name: s.displayName || s.name || s.id, webUrl: s.webUrl }));
  }

  async siteDrives(siteId: string): Promise<Entry[]> {
    const drives = await this.list<{ id: string; name: string; webUrl?: string; driveType?: string }>(
      `/sites/${enc(siteId)}/drives?$select=id,name,webUrl,driveType`,
    );
    return drives
      .filter((d) => d.driveType === undefined || d.driveType === "documentLibrary")
      .map((d) => ({ kind: "drive", id: d.id, name: d.name, webUrl: d.webUrl }));
  }

  async childFolders(driveId: string, itemId: string): Promise<Entry[]> {
    const items = await this.list<GraphItem>(
      `/drives/${enc(driveId)}/items/${enc(itemId)}/children?$select=id,name,webUrl,folder&$top=200`,
    );
    return items
      .filter((i) => i.folder)
      .map((i) => ({ kind: "folder" as const, id: i.id, name: i.name, driveId, webUrl: i.webUrl }))
      .sort((a, b) => a.name.localeCompare(b.name, "de"));
  }

  async createFolder(driveId: string, parentId: string, name: string): Promise<Entry> {
    const i = await this.json<GraphItem>(`/drives/${enc(driveId)}/items/${enc(parentId)}/children`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ name, folder: {}, "@microsoft.graph.conflictBehavior": "fail" }),
    });
    return { kind: "folder", id: i.id, name: i.name, driveId, webUrl: i.webUrl };
  }

  async upload(folder: FolderRef, name: string, data: Uint8Array, contentType = "application/octet-stream"): Promise<SavedFile> {
    const target = `/drives/${enc(folder.driveId)}/items/${enc(folder.itemId)}:/${enc(name)}:`;
    let item: GraphItem;
    if (data.byteLength <= SIMPLE_LIMIT) {
      item = await this.json<GraphItem>(`${target}/content?@microsoft.graph.conflictBehavior=rename`, {
        method: "PUT",
        headers: { "Content-Type": contentType },
        body: data as BodyInit,
      });
    } else {
      item = await this.uploadLarge(target, data);
    }
    return { name: item.name, webUrl: item.webUrl };
  }

  /** Upload-Sitzung für große Anlagen; die Upload-Adresse ist vorab signiert (ohne Authorization). */
  private async uploadLarge(target: string, data: Uint8Array): Promise<GraphItem> {
    const session = await this.json<{ uploadUrl: string }>(`${target}/createUploadSession`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ item: { "@microsoft.graph.conflictBehavior": "rename" } }),
    });
    const total = data.byteLength;
    for (let start = 0; start < total; start += CHUNK) {
      const end = Math.min(start + CHUNK, total);
      const res = await fetch(session.uploadUrl, {
        method: "PUT",
        headers: { "Content-Range": `bytes ${start}-${end - 1}/${total}` },
        body: data.slice(start, end) as BodyInit,
      });
      if (!res.ok) {
        await fetch(session.uploadUrl, { method: "DELETE" }).catch(() => undefined);
        throw new Error(`Hochladen abgebrochen (${res.status}).`);
      }
      if (res.status === 200 || res.status === 201) return (await res.json()) as GraphItem;
    }
    throw new Error("Hochladen unvollständig.");
  }

  /**
   * HTML → PDF über "content?format=pdf". Die Quelle liegt kurz im App-Ordner
   * (OneDrive › Apps › <App-Name>) und wird danach gelöscht.
   *
   * Der Aufruf antwortet mit einer 302-Umleitung, der ein Browser wegen CORS nicht
   * folgen darf. Deshalb läuft er über $batch: Dort kommt die Umleitung als
   * Antwortobjekt mit Location-Kopf zurück; die Zieladresse ist vorab signiert.
   */
  async htmlToPdf(html: string): Promise<Uint8Array> {
    const bytes = utf8Bytes(html);
    if (bytes.byteLength > PDF_SOURCE_LIMIT) {
      throw new Error("Die E-Mail ist für die PDF-Umwandlung zu groß (Grenze 2 MB).");
    }
    const tmpName = `mailablage-${Date.now()}-${Math.random().toString(36).slice(2, 8)}.html`;
    const tmp = await this.json<GraphItem>(`/me/drive/special/approot:/${enc(tmpName)}:/content`, {
      method: "PUT",
      headers: { "Content-Type": "text/html; charset=utf-8" },
      body: bytes as BodyInit,
    });
    const driveId = tmp.parentReference?.driveId;
    const itemPath = driveId ? `/drives/${enc(driveId)}/items/${enc(tmp.id)}` : `/me/drive/items/${enc(tmp.id)}`;
    try {
      const batch = await this.json<{
        responses: { id: string; status: number; headers?: Record<string, string>; body?: unknown }[];
      }>("/$batch", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ requests: [{ id: "1", method: "GET", url: `${itemPath}/content?format=pdf` }] }),
      });
      const r = batch.responses[0];
      const location = r?.headers?.Location || r?.headers?.location;
      if (r && r.status >= 300 && r.status < 400 && location) {
        const pdf = await fetch(location);
        if (!pdf.ok) throw new Error(`PDF konnte nicht geladen werden (${pdf.status}).`);
        return new Uint8Array(await pdf.arrayBuffer());
      }
      if (r && r.status === 200 && typeof r.body === "string") {
        // Binäre Antworten liefert $batch Base64-kodiert.
        return base64ToBytes(r.body);
      }
      const msg = (r?.body as { error?: { message?: string } } | undefined)?.error?.message;
      throw new Error(`PDF-Umwandlung fehlgeschlagen (${r?.status ?? "?"}${msg ? `: ${msg}` : ""}).`);
    } finally {
      await this.fetch(itemPath, { method: "DELETE" }).catch(() => undefined);
    }
  }
}
