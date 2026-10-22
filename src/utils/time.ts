import { Config, TimeObj } from '../types';

export function norm(s: any): string {
  return String(s ?? '').trim();
}

export function normKey(s: any): string {
  return norm(s)
    .replace(/[إأآا]/g, 'ا')
    .replace(/ى/g, 'ي')
    .replace(/ة/g, 'ه')
    .replace(/[^\u0600-\u06FFa-zA-Z0-9]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

export function getVal(o: any, names: string[]): string {
  if (!o) return '';
  // 1. Direct exact normalized key match
  for (const n of names) {
    const k = normKey(n);
    if (o[k] !== undefined && o[k] !== null && o[k] !== '') {
      return String(o[k]).trim();
    }
  }
  // 2. Fuzzy / Substring match against any key in object
  const objKeys = Object.keys(o);
  for (const n of names) {
    const target = normKey(n);
    if (!target) continue;
    for (const key of objKeys) {
      const normK = normKey(key);
      if (normK === target || normK.includes(target) || (target.length >= 4 && target.includes(normK))) {
        if (o[key] !== undefined && o[key] !== null && o[key] !== '') {
          return String(o[key]).trim();
        }
      }
    }
  }
  return '';
}

export function pad(n: number): string {
  return String(n).padStart(2, '0');
}

export function parseTime(v: any): TimeObj | null {
  if (v == null || v === '') return null;
  if (typeof v === 'number') {
    let m = Math.round(v * 24 * 60);
    return {
      h: Math.floor(m / 60) % 24,
      m: m % 60,
      total: m,
      text: `${pad(Math.floor(m / 60) % 24)}:${pad(m % 60)}`,
    };
  }
  let s = norm(v).replace(/[٠-٩]/g, d => '٠١٢٣٤٥٦٧٨٩'.indexOf(d).toString());
  if (s === '-' || s === '—' || s === 'لا يوجد' || s === 'بلا بصمة' || s === 'غير متوفر') {
    return { text: '-', total: 0 } as any; // special block or empty
  }
  let m = s.match(/(\d{1,2})[:٫.](\d{1,2})/);
  if (!m) return null;
  let h = +m[1];
  let mi = +m[2];
  return {
    h,
    m: mi,
    total: h * 60 + mi,
    text: `${pad(h)}:${pad(mi)}`,
  };
}

export function parseDate(v: any): Date | null {
  if (v == null || v === '') return null;
  if (v instanceof Date && !isNaN(v.getTime())) {
    return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  }
  
  // 1. Numeric or numeric-string Excel serial numbers (e.g. 45678 or 46250)
  // In Excel, 1899-12-30 is base (serial 0). Every integer is 1 day.
  // We use UTC methods to prevent any timezone offset from shifting the day backwards/forwards.
  if (typeof v === 'number' && !isNaN(v) && v > 20000 && v < 90000) {
    const totalDays = Math.round(v);
    const d = new Date(Date.UTC(1899, 11, 30 + totalDays));
    return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
  }

  let s = norm(v).replace(/[٠-٩]/g, d => '٠١٢٣٤٥٦٧٨٩'.indexOf(d).toString());

  if (/^\d+(\.\d+)?$/.test(s)) {
    const num = parseFloat(s);
    if (num > 20000 && num < 90000) {
      const totalDays = Math.round(num);
      const d = new Date(Date.UTC(1899, 11, 30 + totalDays));
      return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
    }
  }

  // 2. Check 4-digit year at beginning: YYYY-MM-DD or YYYY/MM/DD
  let mYearFirst = s.match(/(\d{4})[-\/\.٫](\d{1,2})[-\/\.٫](\d{1,2})/);
  if (mYearFirst) {
    const y = +mYearFirst[1];
    const mo = +mYearFirst[2] - 1;
    const day = +mYearFirst[3];
    const d = new Date(y, mo, day);
    if (!isNaN(d.getTime())) return d;
  }

  // 3. Check 4-digit year at end: DD-MM-YYYY or DD/MM/YYYY
  let mYearLast = s.match(/(\d{1,2})[-\/\.٫](\d{1,2})[-\/\.٫](\d{4})/);
  if (mYearLast) {
    const day = +mYearLast[1];
    const mo = +mYearLast[2] - 1;
    const y = +mYearLast[3];
    const d = new Date(y, mo, day);
    if (!isNaN(d.getTime())) return d;
  }

  // 4. Check 2-digit year patterns: e.g. 16-08-26, 26-08-16, 16/08/26, 26/08/16
  let mThreeNums = s.match(/(\d{1,2})[-\/\.٫](\d{1,2})[-\/\.٫](\d{1,2})/);
  if (mThreeNums) {
    const n1 = +mThreeNums[1];
    const n2 = +mThreeNums[2];
    const n3 = +mThreeNums[3];

    let y = 2026;
    let mo = 7; // 0-indexed August
    let day = 1;

    if (n3 >= 20 && n3 <= 50) {
      // Format: DD-MM-YY (e.g. 16-08-26 -> Day 16, Month 8, Year 2026)
      y = 2000 + n3;
      mo = n2 - 1;
      day = n1;
    } else if (n1 >= 20 && n1 <= 50) {
      // Format: YY-MM-DD (e.g. 26-08-16 -> Year 2026, Month 8, Day 16)
      y = 2000 + n1;
      mo = n2 - 1;
      day = n3;
    } else if (n1 > 12) {
      // n1 must be day
      y = n3 < 100 ? 2000 + n3 : n3;
      mo = n2 - 1;
      day = n1;
    } else if (n3 > 12) {
      // n3 must be day
      y = n1 < 100 ? 2000 + n1 : n1;
      mo = n2 - 1;
      day = n3;
    } else {
      // Default standard DD-MM-YY
      y = 2000 + n3;
      mo = n2 - 1;
      day = n1;
    }

    const d = new Date(y, mo, day);
    if (!isNaN(d.getTime())) return d;
  }

  // 5. Try standard Date constructor
  const parsed = new Date(s);
  if (!isNaN(parsed.getTime())) {
    return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
  }

  return null;
}

export function dateKey(d: any): string {
  if (!d) return '';
  if (typeof d === 'string') {
    const match = d.match(/^\d{4}-\d{2}-\d{2}/);
    if (match) return match[0];
    const parsed = parseDate(d);
    if (parsed) return `${parsed.getFullYear()}-${pad(parsed.getMonth() + 1)}-${pad(parsed.getDate())}`;
    return d.trim();
  }
  if (d instanceof Date && !isNaN(d.getTime())) {
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  }
  return '';
}

export function arDate(d: any): string {
  if (!d) return '';
  let dt: Date | null = d instanceof Date ? d : parseDate(d);
  if (!dt || isNaN(dt.getTime())) return String(d);
  return `${pad(dt.getDate())}-${pad(dt.getMonth() + 1)}-${dt.getFullYear()}`;
}

export function toHijri(d: any): string {
  if (!d) return '';
  let dt: Date | null = d instanceof Date ? d : parseDate(d);
  if (!dt || isNaN(dt.getTime())) return '';
  try {
    const parts = new Intl.DateTimeFormat('ar-SA-u-ca-islamic-umalqura', {
      day: '2-digit',
      month: '2-digit',
      year: 'numeric',
      calendar: 'islamic-umalqura',
      numberingSystem: 'latn',
    }).formatToParts(dt);

    const day = parts.find(p => p.type === 'day')?.value || '';
    const month = parts.find(p => p.type === 'month')?.value || '';
    const year = parts.find(p => p.type === 'year')?.value || '';
    if (day && month && year) {
      return `${day.padStart(2, '0')}-${month.padStart(2, '0')}-${year}هـ`;
    }
  } catch (e) {}

  try {
    const formatted = new Intl.DateTimeFormat('ar-SA-u-ca-islamic', {
      day: 'numeric',
      month: 'numeric',
      year: 'numeric',
    }).format(dt);
    const converted = formatted.replace(/[٠-٩]/g, d => '٠١٢٣٤٥٦٧٨٩'.indexOf(d).toString());
    return `${converted}هـ`;
  } catch (e) {
    return '';
  }
}

export function toHijriLong(d: any): string {
  if (!d) return '';
  let dt: Date | null = d instanceof Date ? d : parseDate(d);
  if (!dt || isNaN(dt.getTime())) return '';
  try {
    return new Intl.DateTimeFormat('ar-SA-u-ca-islamic-umalqura', {
      day: 'numeric',
      month: 'long',
      year: 'numeric',
    }).format(dt) + ' هـ';
  } catch (e) {
    return toHijri(dt);
  }
}

export function formatDateWithCalendar(d: any, mode: 'both' | 'hijri' | 'gregorian' = 'both'): string {
  if (!d) return '';
  const greg = arDate(d);
  const hij = toHijri(d);
  if (!hij) return greg ? `${greg}م` : '';

  if (mode === 'hijri') return hij;
  if (mode === 'gregorian') return `${greg}م`;
  return `${greg}م (${hij})`;
}

export function monthKey(d: any): string {
  if (!d) return '';
  if (typeof d === 'string') {
    const match = d.match(/^\d{4}-\d{2}/);
    if (match) return match[0];
    const parsed = parseDate(d);
    if (parsed) return `${parsed.getFullYear()}-${pad(parsed.getMonth() + 1)}`;
    return d.trim();
  }
  if (d instanceof Date && !isNaN(d.getTime())) {
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}`;
  }
  return '';
}

export function minsAfterStart(t: TimeObj | null, startStr: string): number {
  const st = parseTime(startStr);
  if (!t || t.total == null || !st) return 0;
  return Math.max(0, t.total - st.total);
}

export function getApprovedCheckoutTime(
  date?: Date | string | null,
  checkoutSunWed: string = '12:30',
  checkoutThu: string = '11:45'
): TimeObj | null {
  if (!date) return parseTime(checkoutSunWed) || { h: 12, m: 30, total: 750, text: '12:30' };
  const dt = date instanceof Date ? date : parseDate(date);
  if (!dt || isNaN(dt.getTime())) {
    return parseTime(checkoutSunWed) || { h: 12, m: 30, total: 750, text: '12:30' };
  }
  // 4 is Thursday (الخميس)
  if (dt.getDay() === 4) {
    return parseTime(checkoutThu) || { h: 11, m: 45, total: 705, text: '11:45' };
  }
  // Sunday (0), Monday (1), Tuesday (2), Wednesday (3) -> 12:30
  return parseTime(checkoutSunWed) || { h: 12, m: 30, total: 750, text: '12:30' };
}

export function expectedCheckout(
  tin: TimeObj | null,
  workHours: number = 7,
  date?: Date | string | null,
  checkoutSunWed: string = '12:30',
  checkoutThu: string = '11:45'
): number | null {
  if (date) {
    const approved = getApprovedCheckoutTime(date, checkoutSunWed, checkoutThu);
    if (approved && approved.total != null) return approved.total;
  }
  if (!tin || tin.total == null) {
    return parseTime(checkoutSunWed)?.total ?? 750;
  }
  return tin.total + (Number(workHours) || 7) * 60;
}

export function isAbnormalMorningCheckout(
  tin: TimeObj | null,
  tout: TimeObj | null
): boolean {
  if (!tout || tout.total == null || tout.text === '-') return false;
  // If checkout is recorded in the early morning before 10:00 AM (e.g. 06:46, 07:15, 08:30)
  // or within 45 minutes of check-in in the morning before 10:30 AM:
  // This cannot be a genuine dismissal punch for school. It signifies that the teacher
  // either punched twice in the morning or the biometric reader carried over yesterday's
  // missing departure onto today's morning punch (خلل تسجيل انصراف في اليوم التالي / لم يسجل انصراف في موعده).
  if (tout.total < 10 * 60) {
    return true;
  }
  if (tin && tin.total != null && tout.total < 10.5 * 60 && (tout.total - tin.total) <= 45) {
    return true;
  }
  return false;
}

export function isEarlyCheckout(checkoutStateStr: string): boolean {
  return checkoutStateStr === 'خروج قبل نهاية الدوام' || checkoutStateStr === 'انصراف مبكر';
}

export function isEmergencyLeave(s: any): boolean {
  if (!s) return false;
  const k = normKey(s);
  return (
    k.includes('استئذان') ||
    k.includes('استئيذان') ||
    (k.includes('اذن') && (k.includes('طارئ') || k.includes('خروج') || k.includes('عمل')))
  );
}

export function checkoutState(
  status: string,
  tin: TimeObj | null,
  tout: TimeObj | null,
  autoCheckoutStr: string = '13:30',
  workHours: number = 7,
  date?: Date | string | null,
  checkoutSunWed: string = '12:30',
  checkoutThu: string = '11:45'
): string {
  const st = norm(status);
  const hasIn = !!(tin && tin.total != null && tin.text !== '-');
  const hasOut = !!(tout && tout.total != null && tout.text !== '-');
  const isEmergency = isEmergencyLeave(st);

  // If no check-in and no check-out, the employee did not attend
  if (!hasIn && !hasOut) {
    if (isEmergency) return 'استئذان بدون بصمة انصراف';
    return '';
  }

  // If did not check out:
  // If they have an emergency permission (استئذان طارئ) but no checkout punch,
  // this is a specific violation per school regulations!
  if (!hasOut) {
    if (isEmergency) return 'استئذان بدون بصمة انصراف';
    if (hasIn) return 'لا توجد بصمة انصراف';
    return '';
  }

  // If checked out with an Emergency Permission:
  if (isEmergency) {
    return 'استئذان طارئ موثق (بصمة)';
  }

  // 1. Abnormal morning / next-day checkout punch (e.g. in 06:40 and out 06:46 AM)
  // Teacher did not record checkout at 12:30, and punched the next morning at 06:46
  if (
    st.includes('انصراف في غير') ||
    st.includes('بصمة صباحية') ||
    st.includes('غير وقت الفعلي') ||
    isAbnormalMorningCheckout(tin, tout)
  ) {
    return 'انصراف في غير وقت الفعلي';
  }

  // 2. Auto-checkout: from 13:30 and above (من الساعة 13:30 وما زاد), or status text mentions auto checkout
  const auto = parseTime(autoCheckoutStr) || { total: 13 * 60 + 30 };
  if (st.includes('انصراف تلقائي') || tout.total >= auto.total) {
    return 'انصراف تلقائي';
  }

  // 3. Early checkout check based on approved checkout time:
  // Sun-Wed: 12:30 (750 mins), Thu: 11:45 (705 mins)
  // Per user instruction: always designated as 'خروج قبل نهاية الدوام' (never 'معفى')
  const exp = expectedCheckout(tin, workHours, date, checkoutSunWed, checkoutThu);
  if (exp != null && tout.total < exp) {
    return 'خروج قبل نهاية الدوام';
  }

  return 'مكتمل';
}

export function earlyCheckoutMins(
  tin: TimeObj | null,
  tout: TimeObj | null,
  status: string,
  autoCheckoutStr: string = '13:30',
  workHours: number = 7,
  date?: Date | string | null,
  checkoutSunWed: string = '12:30',
  checkoutThu: string = '11:45'
): number {
  if (!tin || tin.total == null || !tout || tout.total == null || tout.text === '-') return 0;
  const state = checkoutState(status, tin, tout, autoCheckoutStr, workHours, date, checkoutSunWed, checkoutThu);
  if (
    state === 'انصراف تلقائي' ||
    state === 'انصراف في غير وقت الفعلي' ||
    state === 'لا توجد بصمة انصراف' ||
    state === 'استئذان بدون بصمة انصراف' ||
    state === 'استئذان طارئ موثق (بصمة)' ||
    state === 'معفى'
  ) {
    return 0;
  }
  const exp = expectedCheckout(tin, workHours, date, checkoutSunWed, checkoutThu);
  return exp == null ? 0 : Math.max(0, exp - tout.total);
}

export function isWeekend(d: Date | string | null | undefined, weekend: number[] = [5, 6]): boolean {
  if (!d) return false;
  let dt: Date | null = d instanceof Date ? d : parseDate(d);
  if (!dt || isNaN(dt.getTime())) return false;
  const days = Array.isArray(weekend) && weekend.length > 0 ? weekend : [5, 6];
  return days.includes(dt.getDay());
}

export function parseHolidayDate(s: string): Date | null {
  s = norm(s).replace(/[٠-٩]/g, d => '٠١٢٣٤٥٦٧٨٩'.indexOf(d).toString());
  let m = s.match(/(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  m = s.match(/(\d{1,2})[-\/](\d{1,2})[-\/](\d{2,4})/);
  if (m) {
    let y = +m[3];
    if (y < 100) y += 2000;
    return new Date(y, +m[2] - 1, +m[1]);
  }
  return null;
}

export function addHolidayRange(set: Set<string>, a: Date | null, b: Date | null) {
  if (!a) return;
  if (!b) b = a;
  const start = a <= b ? a : b;
  const end = a <= b ? b : a;
  for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
    set.add(dateKey(d));
  }
}

export function holidaySet(holidaysStr: string): Set<string> {
  const set = new Set<string>();
  const lines = String(holidaysStr || '')
    .split(/[\n,،]+/)
    .map(x => x.trim())
    .filter(Boolean);
  for (const line of lines) {
    const parts = line.match(/\d{4}[-\/]\d{1,2}[-\/]\d{1,2}|\d{1,2}[-\/]\d{1,2}[-\/]\d{2,4}/g) || [];
    if (parts.length >= 2) {
      addHolidayRange(set, parseHolidayDate(parts[0]), parseHolidayDate(parts[1]));
    } else if (parts.length === 1) {
      addHolidayRange(set, parseHolidayDate(parts[0]), null);
    }
  }
  return set;
}

export function isHoliday(d: Date | string | null | undefined, holidaysStr: string): boolean {
  if (!d) return false;
  return holidaySet(holidaysStr).has(dateKey(d));
}

export function nonWorkReason(d: Date | string | null | undefined, weekend: number[] = [5, 6], holidaysStr: string = ''): string {
  if (isWeekend(d, weekend)) return 'عطلة أسبوعية';
  if (isHoliday(d, holidaysStr)) return 'إجازة رسمية';
  return '';
}

export function dayName(d: Date | string | null | undefined): string {
  if (!d) return '';
  let dt: Date | null = d instanceof Date ? d : parseDate(d);
  if (!dt || isNaN(dt.getTime())) return '';
  return ['الأحد', 'الاثنين', 'الثلاثاء', 'الأربعاء', 'الخميس', 'الجمعة', 'السبت'][dt.getDay()];
}
