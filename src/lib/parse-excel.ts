import * as XLSX from "xlsx";
import type { ScheduleSlot, SlotStatus } from "./types";

const DAY_MAP: Record<string, { fr: string; en: string }> = {
  lundi: { fr: "LUNDI", en: "Monday" },
  mardi: { fr: "MARDI", en: "Tuesday" },
  mercredi: { fr: "MERCREDI", en: "Wednesday" },
  jeudi: { fr: "JEUDI", en: "Thursday" },
  vendredi: { fr: "VENDREDI", en: "Friday" },
  samedi: { fr: "SAMEDI", en: "Saturday" },
  dimanche: { fr: "DIMANCHE", en: "Sunday" },
  monday: { fr: "LUNDI", en: "Monday" },
  tuesday: { fr: "MARDI", en: "Tuesday" },
  wednesday: { fr: "MERCREDI", en: "Wednesday" },
  thursday: { fr: "JEUDI", en: "Thursday" },
  friday: { fr: "VENDREDI", en: "Friday" },
  saturday: { fr: "SAMEDI", en: "Saturday" },
  sunday: { fr: "DIMANCHE", en: "Sunday" },
};

const WEEK_RE = /\bWEEK\s*(\d+)\b/i;
const DAY_DATE_RE =
  /\b(LUNDI|MARDI|MERCREDI|JEUDI|VENDREDI|SAMEDI|DIMANCHE)\s*[-–]?\s*(\d{1,2}\/[A-Za-z]+\/\d{2,4})/i;
const TIME_RE =
  /(\d{1,2})\s*[hH:.]?\s*(\d{2})?\s*(?:à|a|-|–|to)\s*(\d{1,2})\s*[hH:.]?\s*(\d{2})?/i;
const TEACHER_RE = /\bavec\s+((?:M\.|Mme|Mr|Mrs|Miss|Dr\.?)\s*[A-Za-zÀ-ÿ' -]+)/i;
const ROOM_RE =
  /\b(?:salle|salles|room)\s*[:\-]?\s*([A-Za-zÀ-ÿ0-9 /.'’-]+?)(?=\s+(?:avec|de\s+\d)|$)/i;
const CANCEL_RE = /\b(annul[ée]+|cancelled|canceled|report[ée]+)\b/i;

function pad2(n: number): string {
  return String(n).padStart(2, "0");
}

function normalizeTime(h: string, m?: string): string {
  return `${pad2(Number(h))}:${pad2(Number(m || "0"))}`;
}

function cellText(value: unknown): string {
  if (value == null) return "";
  if (typeof value === "string") return value.replace(/\s+/g, " ").trim();
  if (typeof value === "number") return String(value);
  if (value instanceof Date) return value.toISOString();
  return String(value).replace(/\s+/g, " ").trim();
}

function detectDay(text: string): { fr: string; en: string } | null {
  const key = text.toLowerCase();
  for (const [token, mapped] of Object.entries(DAY_MAP)) {
    if (key.includes(token)) return mapped;
  }
  return null;
}

function parseSlotText(
  raw: string,
  ctx: { week: string; weekDate: string; day: string; dayEn: string; roomHint?: string },
): ScheduleSlot | null {
  const text = raw.trim();
  if (!text || text.length < 3) return null;
  if (/^(week|semaine|horaire|jour|salle|room|time|heure)/i.test(text) && text.length < 20) {
    return null;
  }

  const timeMatch = text.match(TIME_RE);
  const startTime = timeMatch
    ? normalizeTime(timeMatch[1], timeMatch[2])
    : "00:00";
  const endTime = timeMatch
    ? normalizeTime(timeMatch[3], timeMatch[4])
    : "00:00";

  const teacherMatch = text.match(TEACHER_RE);
  const roomMatch = text.match(ROOM_RE);
  const status: SlotStatus = CANCEL_RE.test(text) ? "cancelled" : "scheduled";

  let subject = text;
  if (teacherMatch) subject = subject.replace(teacherMatch[0], "").trim();
  subject = subject.replace(/\s{2,}/g, " ").replace(/[–-]\s*$/, "").trim();

  return {
    week: ctx.week,
    weekDate: ctx.weekDate || `${ctx.day}`,
    day: ctx.day,
    dayEn: ctx.dayEn,
    startTime,
    endTime,
    timeRange: `${startTime} - ${endTime}`,
    subject: subject || text,
    teacher: teacherMatch?.[1]?.trim() || "N/A",
    room: roomMatch?.[1]?.trim() || ctx.roomHint || "N/A",
    status,
    rawText: text,
  };
}

function sheetToMatrix(sheet: XLSX.WorkSheet): string[][] {
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1");
  const rows: string[][] = [];
  for (let r = range.s.r; r <= range.e.r; r += 1) {
    const row: string[] = [];
    for (let c = range.s.c; c <= range.e.c; c += 1) {
      const addr = XLSX.utils.encode_cell({ r, c });
      row.push(cellText(sheet[addr]?.v));
    }
    rows.push(row);
  }
  return rows;
}

function parseSheet(name: string, sheet: XLSX.WorkSheet): ScheduleSlot[] {
  const rows = sheetToMatrix(sheet);
  const slots: ScheduleSlot[] = [];

  let currentWeek = "WEEK 1";
  let currentWeekDate = "";
  let dayColumns: Array<{ index: number; day: string; dayEn: string }> = [];

  for (let r = 0; r < rows.length; r += 1) {
    const row = rows[r];
    const joined = row.join(" ").trim();
    if (!joined) continue;

    const weekMatch = joined.match(WEEK_RE);
    if (weekMatch) {
      currentWeek = `WEEK ${weekMatch[1]}`;
    }

    const dayDateMatch = joined.match(DAY_DATE_RE);
    if (dayDateMatch) {
      currentWeekDate = `${dayDateMatch[1].toUpperCase()}-${dayDateMatch[2]}`;
    }

    const maybeHeader = row
      .map((cell, index) => {
        const day = detectDay(cell);
        return day ? { index, day: day.fr, dayEn: day.en } : null;
      })
      .filter(Boolean) as Array<{ index: number; day: string; dayEn: string }>;

    if (maybeHeader.length >= 2) {
      dayColumns = maybeHeader;
      continue;
    }

    if (dayColumns.length === 0) {
      // Fallback: treat each non-empty cell as a slot on an unknown day inferred from text.
      for (const cell of row) {
        if (!cell) continue;
        const day = detectDay(cell) || { fr: "SAMEDI", en: "Saturday" };
        const parsed = parseSlotText(cell, {
          week: currentWeek,
          weekDate: currentWeekDate || `${day.fr}`,
          day: day.fr,
          dayEn: day.en,
        });
        if (parsed) slots.push(parsed);
      }
      continue;
    }

    for (const col of dayColumns) {
      const cell = row[col.index];
      if (!cell) continue;
      // Nearby room hint: look one column to the right if it looks like a room.
      const right = row[col.index + 1] || "";
      const roomHint = /salle|étage|etage|^[0-9]?[A-D]$/i.test(right) ? right : undefined;
      const parsed = parseSlotText(cell, {
        week: currentWeek,
        weekDate: currentWeekDate || `${col.day}`,
        day: col.day,
        dayEn: col.dayEn,
        roomHint,
      });
      if (parsed) slots.push(parsed);
    }
  }

  // De-duplicate identical raw slots.
  const seen = new Set<string>();
  return slots.filter((slot) => {
    const key = `${slot.week}|${slot.day}|${slot.startTime}|${slot.rawText}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

export function parseExcelWorkbook(buffer: ArrayBuffer | Buffer): {
  classes: string[];
  schedules: Record<string, ScheduleSlot[]>;
} {
  const workbook = XLSX.read(buffer, { type: "buffer", cellDates: true });
  const schedules: Record<string, ScheduleSlot[]> = {};

  for (const sheetName of workbook.SheetNames) {
    const name = sheetName.trim();
    if (!name || /^sheet\d+$/i.test(name) && workbook.SheetNames.length > 1) {
      // Keep generic Sheet1 only if it's the sole sheet.
    }
    if (/^(légende|legende|info|readme|index)$/i.test(name)) continue;

    const sheet = workbook.Sheets[sheetName];
    const slots = parseSheet(name, sheet);
    if (slots.length === 0) continue;
    schedules[name] = slots;
  }

  const classes = Object.keys(schedules).sort((a, b) => a.localeCompare(b, "fr"));
  return { classes, schedules };
}
