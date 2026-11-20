import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import samplePlanning from "./sample-planning.json";
import { parseExcelWorkbook } from "./parse-excel";
import { fetchSharePointExcel, resolveShareUrl } from "./sharepoint";
import type { TimetablePayload } from "./types";

const CACHE_TTL_MS = 30_000;
const UPLOAD_PATH = path.join(process.cwd(), "data", "latest.xlsx");

type CacheEntry = {
  expiresAt: number;
  payload: TimetablePayload;
};

let cache: CacheEntry | null = null;
let uploadedPayload: TimetablePayload | null = null;

function samplePayload(error?: string): TimetablePayload {
  const now = new Date().toISOString();
  return {
    generatedAt: samplePlanning.generatedAt,
    refreshedAt: now,
    source: "sample",
    sourceLabel: "Données d'exemple (importez le Excel ISGA pour les vrais horaires)",
    live: false,
    classes: samplePlanning.classes,
    schedules: samplePlanning.schedules as TimetablePayload["schedules"],
    error,
  };
}

function payloadFromBuffer(
  buffer: Buffer,
  source: "sharepoint" | "upload",
  sourceLabel: string,
  live: boolean,
): TimetablePayload {
  const parsed = parseExcelWorkbook(buffer);
  if (parsed.classes.length === 0) {
    throw new Error("Le classeur Excel a été lu, mais aucune filière/cours n'a été détecté.");
  }
  const now = new Date().toISOString();
  return {
    generatedAt: now,
    refreshedAt: now,
    source,
    sourceLabel,
    live,
    classes: parsed.classes,
    schedules: parsed.schedules,
  };
}

async function loadFromDiskUpload(): Promise<TimetablePayload | null> {
  try {
    const buffer = await readFile(UPLOAD_PATH);
    return payloadFromBuffer(
      buffer,
      "upload",
      "Fichier Excel importé manuellement",
      false,
    );
  } catch {
    return null;
  }
}

async function loadLive(): Promise<TimetablePayload> {
  const { buffer, mode } = await fetchSharePointExcel();
  return payloadFromBuffer(
    buffer,
    "sharepoint",
    mode === "graph"
      ? "SharePoint (Microsoft Graph) — synchro live"
      : "SharePoint (lien de partage) — synchro live",
    true,
  );
}

export async function ingestExcelUpload(buffer: Buffer): Promise<TimetablePayload> {
  const payload = payloadFromBuffer(
    buffer,
    "upload",
    "Fichier Excel importé manuellement",
    false,
  );
  await mkdir(path.dirname(UPLOAD_PATH), { recursive: true });
  await writeFile(UPLOAD_PATH, buffer);
  uploadedPayload = payload;
  cache = { payload, expiresAt: Date.now() + CACHE_TTL_MS };
  return payload;
}

export async function getTimetable(force = false): Promise<TimetablePayload> {
  const now = Date.now();
  if (!force && cache && cache.expiresAt > now) {
    return cache.payload;
  }

  // Prefer SharePoint when reachable; otherwise keep a previous upload.
  try {
    const payload = await loadLive();
    cache = { payload, expiresAt: now + CACHE_TTL_MS };
    return payload;
  } catch (err) {
    const message = err instanceof Error ? err.message : "Erreur SharePoint inconnue";

    if (cache?.payload.source === "sharepoint") {
      return {
        ...cache.payload,
        refreshedAt: new Date().toISOString(),
        error: `Synchro temporairement indisponible: ${message}`,
      };
    }

    if (uploadedPayload) {
      const refreshed = {
        ...uploadedPayload,
        refreshedAt: new Date().toISOString(),
      };
      cache = { payload: refreshed, expiresAt: now + CACHE_TTL_MS };
      return refreshed;
    }

    const fromDisk = await loadFromDiskUpload();
    if (fromDisk) {
      uploadedPayload = fromDisk;
      cache = { payload: fromDisk, expiresAt: now + CACHE_TTL_MS };
      return fromDisk;
    }

    const fallback = samplePayload(
      "Pas d’accès SharePoint automatique. Importez le fichier Excel téléchargé depuis SharePoint (aucune clé Azure nécessaire).",
    );
    cache = { payload: fallback, expiresAt: now + CACHE_TTL_MS };
    return fallback;
  }
}

export function getSharePointTarget(): string {
  return resolveShareUrl();
}
