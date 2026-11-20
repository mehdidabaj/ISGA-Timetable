export type SlotStatus = "scheduled" | "cancelled" | "empty";

export type ScheduleSlot = {
  week: string;
  weekDate: string;
  day: string;
  dayEn: string;
  startTime: string;
  endTime: string;
  timeRange: string;
  subject: string;
  teacher: string;
  room: string;
  status: SlotStatus;
  rawText: string;
};

export type TimetableSource = "sharepoint" | "upload" | "sample";

export type TimetablePayload = {
  generatedAt: string;
  refreshedAt: string;
  source: TimetableSource;
  sourceLabel: string;
  live: boolean;
  classes: string[];
  schedules: Record<string, ScheduleSlot[]>;
  error?: string;
};
