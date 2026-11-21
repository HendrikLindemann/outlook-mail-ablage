import { createTokenProvider } from "./drive/auth";
import { GraphDrive } from "./drive/graph";
import { MemoryDrive } from "./drive/memory";
import type { DriveApi } from "./drive/types";
import { initHost } from "./host";
import "./styles.css";
import { createApp } from "./ui/app";

interface Config {
  clientId?: string;
  authority?: string;
}

async function loadConfig(): Promise<Config> {
  try {
    const res = await fetch("./config.json", { cache: "no-store" });
    return res.ok ? ((await res.json()) as Config) : {};
  } catch {
    return {};
  }
}

async function start() {
  const root = document.getElementById("app")!;
  const [host, config] = await Promise.all([initHost(), loadConfig()]);
  let drive: DriveApi;
  let setupHint: string | undefined;
  if (!config.clientId) {
    drive = new MemoryDrive();
    setupHint = "Demo-Modus: In config.json ist keine Client-ID eingetragen. Dateien werden nicht wirklich gespeichert.";
  } else if (host.inOutlook && !host.naaSupported) {
    drive = new MemoryDrive();
    setupHint = "Diese Outlook-Version unterstützt die Anmeldung für Add-ins (NAA) nicht. Bitte Outlook aktualisieren.";
  } else {
    drive = new GraphDrive(createTokenProvider(config.clientId, config.authority || "https://login.microsoftonline.com/organizations"));
  }
  createApp(host, drive, root, setupHint).start();
}

start();
