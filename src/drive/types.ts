/** Ziel- und Navigationsobjekte für OneDrive/SharePoint. */

export interface FolderRef {
  driveId: string;
  /** Element-ID des Ordners; "root" für das Stammverzeichnis einer Bibliothek */
  itemId: string;
  /** Anzeigepfad, z. B. "Projekte / Gutachten / 2026" */
  label: string;
  webUrl?: string;
}

export type EntryKind = "site" | "drive" | "folder";

export interface Entry {
  kind: EntryKind;
  /** Site-ID, Drive-ID bzw. Element-ID */
  id: string;
  name: string;
  /** bei Ordnern: zugehörige Bibliothek */
  driveId?: string;
  webUrl?: string;
}

export interface SavedFile {
  name: string;
  webUrl?: string;
}

export class AuthRequiredError extends Error {}

export interface DriveApi {
  demo: boolean;
  /** Stammordner des eigenen OneDrive */
  myDrive(): Promise<FolderRef>;
  followedSites(): Promise<Entry[]>;
  searchSites(query: string): Promise<Entry[]>;
  siteDrives(siteId: string): Promise<Entry[]>;
  childFolders(driveId: string, itemId: string): Promise<Entry[]>;
  createFolder(driveId: string, parentId: string, name: string): Promise<Entry>;
  /** Lädt eine Datei hoch; bei Namensgleichheit vergibt OneDrive einen neuen Namen ("… 1.pdf"). */
  upload(folder: FolderRef, name: string, data: Uint8Array, contentType?: string): Promise<SavedFile>;
  /** Wandelt ein HTML-Dokument serverseitig in PDF um. */
  htmlToPdf(html: string): Promise<Uint8Array>;
  /** Fordert die Anmeldung interaktiv an (nach AuthRequiredError). */
  signIn(): Promise<void>;
}
