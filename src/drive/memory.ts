import { splitExtension } from "../core/names";
import type { DriveApi, Entry, FolderRef, SavedFile } from "./types";

interface Node {
  id: string;
  name: string;
  driveId: string;
  parent?: string;
  folder: boolean;
  data?: Uint8Array;
}

/**
 * Speicher im Arbeitsspeicher für Demo und Tests – ohne Anmeldung, ohne Graph.
 * Bildet OneDrive, eine gefolgte Website und das Umbenennen bei Namensgleichheit nach.
 */
export class MemoryDrive implements DriveApi {
  demo = true;
  nodes = new Map<string, Node>();
  pdfSources: string[] = [];
  private seq = 0;

  constructor() {
    this.add("me", "root", "Mein OneDrive", undefined, true);
    const projekte = this.add("me", "p1", "Projekte", "root", true);
    this.add("me", "p2", "Gutachten 2026", projekte.id, true);
    this.add("me", "p3", "Korrespondenz", "root", true);
    this.add("site-docs", "root", "Dokumente", undefined, true);
    this.add("site-docs", "s1", "Allgemein", "root", true);
  }

  private key(driveId: string, id: string) {
    return `${driveId}|${id}`;
  }

  private add(driveId: string, id: string, name: string, parent: string | undefined, folder: boolean, data?: Uint8Array) {
    const n: Node = { id, name, driveId, parent, folder, data };
    this.nodes.set(this.key(driveId, id), n);
    return n;
  }

  /** Alle Dateien eines Ordners (für Tests). */
  filesIn(driveId: string, folderId: string): Node[] {
    return [...this.nodes.values()].filter((n) => n.driveId === driveId && n.parent === folderId && !n.folder);
  }

  async signIn() {}

  async myDrive(): Promise<FolderRef> {
    return { driveId: "me", itemId: "root", label: "Mein OneDrive", webUrl: "https://example.invalid/onedrive" };
  }

  async followedSites(): Promise<Entry[]> {
    return [{ kind: "site", id: "site1", name: "Team Gutachten", webUrl: "https://example.invalid/sites/team" }];
  }

  async searchSites(query: string): Promise<Entry[]> {
    return (await this.followedSites()).filter((s) => s.name.toLowerCase().includes(query.toLowerCase()));
  }

  async siteDrives(): Promise<Entry[]> {
    return [{ kind: "drive", id: "site-docs", name: "Dokumente" }];
  }

  async childFolders(driveId: string, itemId: string): Promise<Entry[]> {
    return [...this.nodes.values()]
      .filter((n) => n.driveId === driveId && n.parent === itemId && n.folder)
      .map((n) => ({ kind: "folder" as const, id: n.id, name: n.name, driveId }))
      .sort((a, b) => a.name.localeCompare(b.name, "de"));
  }

  async createFolder(driveId: string, parentId: string, name: string): Promise<Entry> {
    if ((await this.childFolders(driveId, parentId)).some((e) => e.name.toLowerCase() === name.toLowerCase())) {
      throw new Error("Ein Ordner mit diesem Namen existiert bereits.");
    }
    const n = this.add(driveId, `n${++this.seq}`, name, parentId, true);
    return { kind: "folder", id: n.id, name: n.name, driveId };
  }

  async upload(folder: FolderRef, name: string, data: Uint8Array): Promise<SavedFile> {
    const taken = new Set(
      [...this.nodes.values()]
        .filter((n) => n.driveId === folder.driveId && n.parent === folder.itemId)
        .map((n) => n.name.toLowerCase()),
    );
    let final = name;
    const [base, ext] = splitExtension(name);
    for (let i = 1; taken.has(final.toLowerCase()); i++) final = `${base} ${i}${ext}`;
    this.add(folder.driveId, `f${++this.seq}`, final, folder.itemId, false, data);
    return { name: final, webUrl: `https://example.invalid/${encodeURIComponent(final)}` };
  }

  async htmlToPdf(html: string): Promise<Uint8Array> {
    this.pdfSources.push(html);
    return new TextEncoder().encode("%PDF-1.7\n% Demo\n");
  }
}
