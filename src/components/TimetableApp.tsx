"use client";

import { useEffect, useMemo, useRef, useState, useTransition } from "react";
import type { ScheduleSlot, TimetablePayload } from "@/lib/types";

const REFRESH_MS = 30_000;

const DAY_ORDER = [
  "LUNDI",
  "MARDI",
  "MERCREDI",
  "JEUDI",
  "VENDREDI",
  "SAMEDI",
  "DIMANCHE",
];

function formatRelative(iso?: string): string {
  if (!iso) return "—";
  const then = new Date(iso).getTime();
  if (Number.isNaN(then)) return "—";
  const diff = Date.now() - then;
  const mins = Math.max(0, Math.round(diff / 60_000));
  if (mins < 1) return "à l'instant";
  if (mins === 1) return "il y a 1 min";
  if (mins < 60) return `il y a ${mins} min`;
  const hours = Math.round(mins / 60);
  if (hours < 48) return `il y a ${hours} h`;
  return new Date(iso).toLocaleString("fr-FR");
}

function weekNumber(week: string): number {
  const m = week.match(/(\d+)/);
  return m ? Number(m[1]) : 0;
}

function currentWeekKey(slots: ScheduleSlot[]): string | null {
  if (slots.length === 0) return null;
  const weeks = [...new Set(slots.map((s) => s.week))].sort(
    (a, b) => weekNumber(a) - weekNumber(b),
  );
  const now = Date.now();
  let best: { week: string; distance: number } | null = null;
  for (const week of weeks) {
    const sample = slots.find((s) => s.week === week && s.weekDate.includes("/"));
    if (!sample) continue;
    const datePart = sample.weekDate.split("-").slice(1).join("-");
    const parsed = Date.parse(
      datePart.replace(/(\d+)\/(\w+)\/(\d+)/, (_, d, mon, y) => {
        const months: Record<string, string> = {
          jan: "01",
          january: "01",
          fev: "02",
          feb: "02",
          february: "02",
          mar: "03",
          march: "03",
          avr: "04",
          apr: "04",
          april: "04",
          mai: "05",
          may: "05",
          jun: "06",
          june: "06",
          jul: "07",
          july: "07",
          aou: "08",
          aug: "08",
          august: "08",
          sep: "09",
          sept: "09",
          september: "09",
          oct: "10",
          october: "10",
          nov: "11",
          november: "11",
          dec: "12",
          december: "12",
        };
        const key = String(mon).slice(0, 3).toLowerCase();
        const mm = months[key] || months[String(mon).toLowerCase()] || "01";
        const year = String(y).length === 2 ? `20${y}` : y;
        return `${year}-${mm}-${String(d).padStart(2, "0")}T12:00:00Z`;
      }),
    );
    if (Number.isNaN(parsed)) continue;
    const distance = Math.abs(parsed - now);
    if (!best || distance < best.distance) best = { week, distance };
  }
  return best?.week || weeks[0] || null;
}

export function TimetableApp() {
  const [data, setData] = useState<TimetablePayload | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [classe, setClasse] = useState("");
  const [week, setWeek] = useState("current");
  const [uploadMsg, setUploadMsg] = useState<string | null>(null);
  const [pending, startTransition] = useTransition();
  const fileRef = useRef<HTMLInputElement>(null);

  async function load(force = false) {
    try {
      const res = await fetch(`/api/timetable${force ? "?force=1" : ""}`, {
        cache: "no-store",
      });
      if (!res.ok) throw new Error(`API ${res.status}`);
      const json = (await res.json()) as TimetablePayload;
      setData(json);
      setError(json.error || null);
      setClasse((prev) =>
        prev && json.classes.includes(prev) ? prev : json.classes[0] || "",
      );
    } catch (err) {
      setError(err instanceof Error ? err.message : "Erreur de chargement");
    }
  }

  async function onUpload(file: File | null) {
    if (!file) return;
    setUploadMsg(null);
    const body = new FormData();
    body.append("file", file);
    try {
      const res = await fetch("/api/upload", { method: "POST", body });
      const json = (await res.json()) as TimetablePayload & { error?: string };
      if (!res.ok) throw new Error(json.error || `Import ${res.status}`);
      setData(json);
      setError(null);
      setClasse(json.classes[0] || "");
      setWeek("current");
      setUploadMsg(
        `Import OK — ${json.classes.length} filière(s), ${Object.values(json.schedules).reduce((n, s) => n + s.length, 0)} séance(s).`,
      );
    } catch (err) {
      setUploadMsg(err instanceof Error ? err.message : "Import impossible");
    } finally {
      if (fileRef.current) fileRef.current.value = "";
    }
  }

  useEffect(() => {
    void load(true);
    const id = window.setInterval(() => void load(false), REFRESH_MS);
    return () => window.clearInterval(id);
  }, []);

  const slots = useMemo(() => {
    if (!data || !classe) return [] as ScheduleSlot[];
    return data.schedules[classe] || [];
  }, [data, classe]);

  const weeks = useMemo(() => {
    return [...new Set(slots.map((s) => s.week))].sort(
      (a, b) => weekNumber(a) - weekNumber(b),
    );
  }, [slots]);

  const activeWeek = useMemo(() => {
    if (week === "all") return null;
    if (week === "current") return currentWeekKey(slots);
    return week;
  }, [week, slots]);

  const filtered = useMemo(() => {
    const list = activeWeek ? slots.filter((s) => s.week === activeWeek) : slots;
    return [...list].sort((a, b) => {
      return (
        DAY_ORDER.indexOf(a.day) - DAY_ORDER.indexOf(b.day) ||
        a.startTime.localeCompare(b.startTime)
      );
    });
  }, [slots, activeWeek]);

  const byDay = useMemo(() => {
    const map = new Map<string, ScheduleSlot[]>();
    for (const slot of filtered) {
      const arr = map.get(slot.day) || [];
      arr.push(slot);
      map.set(slot.day, arr);
    }
    return DAY_ORDER.filter((d) => map.has(d)).map((d) => ({
      day: d,
      slots: map.get(d) || [],
      weekDate: map.get(d)?.[0]?.weekDate || "",
    }));
  }, [filtered]);

  const statusLabel =
    data?.source === "sharepoint"
      ? "Synchro temps réel"
      : data?.source === "upload"
        ? "Excel importé"
        : "Mode aperçu";

  return (
    <div className="app-shell">
      <header className="hero">
        <p className={`status-pill ${data?.live ? "live" : ""}`}>
          <span className="dot" />
          {statusLabel}
        </p>
        <h1 className="brand">
          ISGA <span>Emplois du temps</span>
        </h1>
        <p className="hero-copy">
          Consultez les plannings des filières. Aucune clé Azure n’est requise : importez
          simplement le fichier Excel téléchargé depuis SharePoint.
        </p>
        <div className="hero-actions">
          <button
            className="btn"
            type="button"
            disabled={pending}
            onClick={() => startTransition(() => void load(true))}
          >
            {pending ? "Actualisation…" : "Actualiser maintenant"}
          </button>
          <button
            className="btn secondary"
            type="button"
            onClick={() => fileRef.current?.click()}
          >
            Importer le Excel
          </button>
          <span className="status-pill">
            Dernière synchro : {formatRelative(data?.refreshedAt)}
          </span>
        </div>
      </header>

      <section className="panel" aria-label="Sélection de filière">
        <h2>Choisir une filière</h2>
        <p>
          {data?.sourceLabel || "Chargement du planning…"}
          {data?.live ? " · rafraîchi toutes les 30 s" : ""}
        </p>

        {!data ? <div className="loading">Chargement des emplois du temps…</div> : null}

        {data?.source === "sample" ? (
          <div className="import-box">
            <h3>Importer le planning (sans clés)</h3>
            <p>
              1. Ouvrez le fichier SharePoint ISGA dans votre navigateur · 2. Téléchargez une
              copie (.xlsx) · 3. Importez-la ici. Pour une synchro auto plus tard, un simple
              lien de partage public suffit (pas de clés Microsoft).
            </p>
            <div className="import-actions">
              <input
                ref={fileRef}
                type="file"
                accept=".xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel"
                onChange={(e) => void onUpload(e.target.files?.[0] || null)}
              />
            </div>
            {uploadMsg ? <p>{uploadMsg}</p> : null}
          </div>
        ) : (
          <>
            <input
              ref={fileRef}
              type="file"
              hidden
              accept=".xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel"
              onChange={(e) => void onUpload(e.target.files?.[0] || null)}
            />
            {uploadMsg ? <div className="empty-state">{uploadMsg}</div> : null}
            {error && data?.source === "sharepoint" ? (
              <div className="error-banner">{error}</div>
            ) : null}
          </>
        )}

        {data ? (
          <>
            <div className="controls">
              <label>
                Filière / classe
                <select
                  value={classe}
                  onChange={(e) => {
                    setClasse(e.target.value);
                    setWeek("current");
                  }}
                >
                  {data.classes.map((c) => (
                    <option key={c} value={c}>
                      {c}
                    </option>
                  ))}
                </select>
              </label>

              <label>
                Semaine
                <select value={week} onChange={(e) => setWeek(e.target.value)}>
                  <option value="current">Semaine actuelle</option>
                  <option value="all">Toutes les semaines</option>
                  {weeks.map((w) => (
                    <option key={w} value={w}>
                      {w.replace(/week/i, "Semaine")}
                    </option>
                  ))}
                </select>
              </label>

              <button
                className="btn"
                type="button"
                onClick={() => startTransition(() => void load(true))}
              >
                Synchroniser
              </button>
            </div>

            <div className="meta-row">
              <span>
                {filtered.length} séance{filtered.length > 1 ? "s" : ""}
              </span>
              <span>
                {activeWeek
                  ? activeWeek.replace(/week/i, "Semaine")
                  : "Vue complète"}
              </span>
            </div>

            {weeks.length > 0 ? (
              <div className="week-nav" role="tablist" aria-label="Navigation semaines">
                <button
                  type="button"
                  className={`chip ${week === "current" ? "active" : ""}`}
                  onClick={() => setWeek("current")}
                >
                  Actuelle
                </button>
                {weeks.map((w) => (
                  <button
                    key={w}
                    type="button"
                    className={`chip ${week === w ? "active" : ""}`}
                    onClick={() => setWeek(w)}
                  >
                    {w.replace(/week/i, "S")}
                  </button>
                ))}
              </div>
            ) : null}

            {byDay.length === 0 ? (
              <div className="empty-state">Aucun cours trouvé pour cette sélection.</div>
            ) : (
              <div className="schedule">
                {byDay.map((group) => (
                  <section key={group.day} className="day-block">
                    <h3>
                      {group.day.charAt(0) + group.day.slice(1).toLowerCase()}
                      <span>{group.weekDate}</span>
                    </h3>
                    {group.slots.map((slot, idx) => (
                      <article
                        key={`${slot.rawText}-${idx}`}
                        className={`slot ${slot.status === "cancelled" ? "cancelled" : ""}`}
                      >
                        <time dateTime={`${slot.startTime}`}>{slot.timeRange}</time>
                        <div>
                          <p className="subject">{slot.subject}</p>
                          <p className="details">
                            {slot.teacher !== "N/A" ? slot.teacher : "Enseignant non précisé"}
                            {" · "}
                            {slot.room !== "N/A" ? slot.room : "Salle non précisée"}
                          </p>
                        </div>
                      </article>
                    ))}
                  </section>
                ))}
              </div>
            )}
          </>
        ) : null}
      </section>

      <footer className="footer">
        ISGA · import Excel ou lien SharePoint public · année 2025-2026
      </footer>
    </div>
  );
}
