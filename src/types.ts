export interface Config {
  start: string; // e.g., '06:45'
  workHours: number; // e.g., 7
  autoCheckout: string; // e.g., '13:30' (from 13:30 onwards is auto-checkout)
  checkoutSunWed?: string; // e.g. '12:30' (Approved checkout Sun-Wed)
  checkoutThu?: string; // e.g. '11:45' (Approved checkout Thu)
  weekend: number[]; // e.g., [5, 6] (Friday and Saturday)
  holidays: string; // holiday list as string
  warn1: number;
  warn2: number;
  ramadanStart?: string; // e.g., '09:00'
  ramadanWorkHours?: number; // e.g., 5
  ramadanAutoCheckout?: string; // e.g., '08:45'
  ramadanDates?: string[]; // e.g. ['2026-03-10', ...]
  enableAutoAlert?: boolean; // Automatic Alert feature toggle
  calendarMode?: 'both' | 'hijri' | 'gregorian'; // 'both' | 'hijri' | 'gregorian'
  suspendedDates?: string[]; // e.g. ['2026-10-01', ...] for school-wide suspended/remote learning days
  autoDetectSuspendedDays?: boolean; // automatically detect days with 0 school-wide punches as suspended
}

export interface TimeObj {
  h: number;
  m: number;
  total: number;
  text: string;
}

export interface ParsedRow {
  name: string;
  job: string;
  civil: string;
  date: Date;
  dateKey: string;
  month: string;
  status: string;
  tin: TimeObj | null;
  tout: TimeObj | null;
  expected: number | null;
  checkout: string;
  earlyMins: number;
  type: 'present' | 'absence' | 'excuse' | 'other' | 'missing';
  late: number;
}

export interface EmployeeSummary {
  name: string;
  job: string;
  civil?: string;
  month: string;
  work: number;
  present: number;
  absence: number;
  excuse: number;
  lateCount: number;
  lateMins: number;
  completeCount: number;
  earlyCount: number;
  earlyMins: number;
  autoCount: number;
  missingCheckoutCount: number;
  wrongTimeExitsCount?: number;
  emergencyNoExitCount?: number;
  emergencyWithExitCount?: number;
  attendanceRate?: number;
  rate?: number; // discipline score (0-100)
  rating?: string; // discipline classification label
  disciplineCategory?: 'ideal' | 'notes' | 'follow' | 'admin' | 'none';
  disciplineReason?: string;
}

export interface MessageLogEntry {
  time: string;
  name: string;
  date: string;
  type: string;
  result: string;
}

export interface ImportStats {
  totalRows: number;
  imported: number;
  present: number;
  absence: number;
  excuse: number;
  other: number;
  missing: number;
  noName: number;
  noDate: number;
  ignored: number;
}
