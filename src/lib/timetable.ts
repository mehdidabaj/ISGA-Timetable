import samplePlanning from "./sample-planning.json";
import { parseExcelWorkbook } from "./parse-excel";
import { fetchSharePointExcel, resolveShareUrl } from "./sharepoint";
import type { TimetablePayload } from "./types";

const CACHE_TTL_MS = 30_000;

type CacheEntry = {
  expiresAt: number;
  payload: TimetablePayload;
};

let cache: CacheEntry | null = null;

function samplePayload(error?: string): TimetablePayload {
  const now = new Date().toISOString();
  return {
    generatedAt: samplePlanning.generatedAt,
    refreshedAt: now,
    source: "sample",
    sourceLabel: "Données d'exemple (SharePoint non connecté)",
    live: false,
    classes: samplePlanning.classes,
    schedules: samplePlanning.schedules as TimetablePayload["schedules"],
    error,
  };
}

async function loadLive(): Promise<TimetablePayload> {
  const { buffer, mode } = await fetchSharePointExcel();
  const parsed = parseExcelWorkbook(buffer);
  if (parsed.classes.length === 0) {
    throw new Error("Le classeur Excel a été lu, mais aucune filière/cours n'a été détecté.");
  }

  const now = new Date().toISOString();
  return {
    generatedAt: now,
    refreshedAt: now,
    source: "sharepoint",
    sourceLabel:
      mode === "graph"
        ? "SharePoint (Microsoft Graph) — synchro live"
        : "SharePoint (lien de partage) — synchro live",
    live: true,
    classes: parsed.classes,
    schedules: parsed.schedules,
  };
}

export async function getTimetable(force = false): Promise<TimetablePayload> {
  const now = Date.now();
  if (!force && cache && cache.expiresAt > now) {
    return cache.payload;
  }

  try {
    const payload = await loadLive();
    cache = { payload, expiresAt: now + CACHE_TTL_MS };
    return payload;
  } catch (err) {
    const message = err instanceof Error ? err.message : "Erreur SharePoint inconnue";
    // Keep serving a previous live cache if we have one.
    if (cache?.payload.source === "sharepoint") {
      const stale: TimetablePayload = {
        ...cache.payload,
        refreshedAt: new Date().toISOString(),
        error: `Synchro temporairement indisponible: ${message}`,
      };
      return stale;
    }
    const fallback = samplePayload(message);
    cache = { payload: fallback, expiresAt: now + CACHE_TTL_MS };
    return fallback;
  }
}

export function getSharePointTarget(): string {
  return resolveShareUrl();
}
