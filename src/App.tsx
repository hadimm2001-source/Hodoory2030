import React, { useState, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx';
import {
  LayoutDashboard,
  Users,
  Clock,
  CalendarX,
  FileText,
  Calendar,
  Settings as SettingsIcon,
  Upload,
  CheckCircle2,
  AlertCircle,
  Share2,
  Printer,
  FileDown,
  Undo2,
  Copy,
  Trash2,
  Check,
  X,
  Plus,
  ArrowLeft,
  MessageSquare,
  AlertTriangle,
  Edit,
  FileSpreadsheet
} from 'lucide-react';

import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Legend,
  ResponsiveContainer,
  Cell
} from 'recharts';

import {
  Config,
  TimeObj,
  ParsedRow,
  EmployeeSummary,
  MessageLogEntry,
  ImportStats
} from './types';

import {
  norm,
  normKey,
  getVal,
  pad,
  parseTime,
  parseDate,
  dateKey,
  arDate,
  toHijri,
  toHijriLong,
  formatDateWithCalendar,
  monthKey,
  minsAfterStart,
  expectedCheckout,
  checkoutState,
  earlyCheckoutMins,
  isEmergencyLeave,
  isWeekend,
  isHoliday,
  holidaySet,
  nonWorkReason,
  dayName
} from './utils/time';

import {
  disciplineScore,
  disciplineReason,
  disciplineClassify
} from './utils/discipline';

// Default Holidays String
const DEFAULT_HOLIDAYS = `2025-09-23
2025-10-12
2025-11-21 إلى 2025-11-29
2025-12-11 إلى 2025-12-14
2026-01-09 إلى 2026-01-17
2026-02-22
2026-03-06 إلى 2026-03-28
2026-05-22 إلى 2026-06-01
2026-06-25 إلى 2026-08-15`;

// Keys for local storage
const LOCAL_STORAGE_KEYS = {
  CONFIG: 'jashah_config_v1',
  ROWS: 'jashah_rows_v2', // Versioned to support newer formats
  FARES_SUBMITTED: 'jashah_fares_submitted_v1',
  MESSAGE_LOG: 'jashah_message_log_v1',
  IMPORT_STATS: 'jashah_import_stats_v1',
};

interface CustomTooltipProps {
  active?: boolean;
  payload?: any[];
}

const CustomTooltip = ({ active, payload }: CustomTooltipProps) => {
  if (active && payload && payload.length) {
    const data = payload[0].payload;
    return (
      <div className="bg-slate-900 text-white p-3.5 rounded-2xl border border-slate-800 shadow-xl text-xs font-bold space-y-1" dir="rtl">
        <p className="text-slate-300">{data.name}</p>
        <p className="text-sm font-black flex items-center gap-1">
          <span>العدد:</span>
          <span style={{ color: data.fill }}>{data.count} موظف</span>
        </p>
        <p className="text-[10px] text-slate-400 font-normal">{data.label}</p>
      </div>
    );
  }
  return null;
};

export default function App() {
  // 1. Core State
  const [config, setConfig] = useState<Config>(() => {
    const saved = localStorage.getItem(LOCAL_STORAGE_KEYS.CONFIG);
    if (saved) {
      try { 
        const parsed = JSON.parse(saved);
        let holidaysStr = typeof parsed.holidays === 'string' ? parsed.holidays : DEFAULT_HOLIDAYS;
        // Fix any stale summer holiday string that was blocking mid-August dates
        if (holidaysStr.includes('2026-08-22')) {
          holidaysStr = holidaysStr.replace(/2026-08-22/g, '2026-08-15');
        }
        return {
          start: '06:45',
          workHours: 7,
          weekend: [5, 6], // 5=Friday, 6=Saturday in JS Date.getDay()
          warn1: 30,
          warn2: 60,
          calendarMode: 'both',
          ...parsed,
          holidays: holidaysStr,
          autoCheckout: (!parsed.autoCheckout || parsed.autoCheckout === '16:00') ? '13:30' : parsed.autoCheckout,
          checkoutSunWed: parsed.checkoutSunWed || '12:30',
          checkoutThu: parsed.checkoutThu || '11:45',
          enableAutoAlert: parsed.enableAutoAlert !== undefined ? parsed.enableAutoAlert : true,
          suspendedDates: Array.isArray(parsed.suspendedDates) ? parsed.suspendedDates : [],
          autoDetectSuspendedDays: parsed.autoDetectSuspendedDays !== undefined ? parsed.autoDetectSuspendedDays : true,
        };
      } catch (e) {}
    }
    return {
      start: '06:45',
      workHours: 7,
      autoCheckout: '13:30',
      checkoutSunWed: '12:30',
      checkoutThu: '11:45',
      weekend: [5, 6], // 5=Friday, 6=Saturday in JS Date.getDay()
      holidays: DEFAULT_HOLIDAYS,
      warn1: 30,
      warn2: 60,
      enableAutoAlert: true,
      calendarMode: 'both', // 'both' | 'hijri' | 'gregorian'
      suspendedDates: [],
      autoDetectSuspendedDays: true,
    };
  });

  const [rawRows, setRawRows] = useState<any[]>(() => {
    const saved = localStorage.getItem(LOCAL_STORAGE_KEYS.ROWS);
    return saved ? JSON.parse(saved) : [];
  });

  const [faresSubmitted, setFaresSubmitted] = useState<Record<string, boolean>>(() => {
    const saved = localStorage.getItem(LOCAL_STORAGE_KEYS.FARES_SUBMITTED);
    return saved ? JSON.parse(saved) : {};
  });

  const [messageLogs, setMessageLogs] = useState<MessageLogEntry[]>(() => {
    const saved = localStorage.getItem(LOCAL_STORAGE_KEYS.MESSAGE_LOG);
    return saved ? JSON.parse(saved) : [];
  });

  const [importStats, setImportStats] = useState<ImportStats | null>(() => {
    const saved = localStorage.getItem(LOCAL_STORAGE_KEYS.IMPORT_STATS);
    return saved ? JSON.parse(saved) : null;
  });

  // UI Navigation / Filter state
  const [currentTab, setCurrentTab] = useState<string>('dashboard');
  const [selectedMonth, setSelectedMonth] = useState<string>('all');
  const [selectedPeriodType, setSelectedPeriodType] = useState<'month' | 'range'>('month');
  const [selectedStartDate, setSelectedStartDate] = useState<string>('');
  const [selectedEndDate, setSelectedEndDate] = useState<string>('');
  const [selectedEmployee, setSelectedEmployee] = useState<string>('all');
  const [selectedStatus, setSelectedStatus] = useState<string>('all');
  const [searchTerm, setSearchTerm] = useState<string>('');
  const [excuseSubFilter, setExcuseSubFilter] = useState<'all' | 'leaves' | 'emergency_all' | 'emergency_no_exit' | 'emergency_with_exit'>('all');

  // Monthly Report interactive selections
  const [monthlyEmployee, setMonthlyEmployee] = useState<string>('all');
  const [monthlyMonth, setMonthlyMonth] = useState<string>('all');
  const [monthlyPeriodType, setMonthlyPeriodType] = useState<'month' | 'range'>('month');
  const [monthlyStartDate, setMonthlyStartDate] = useState<string>('');
  const [monthlyEndDate, setMonthlyEndDate] = useState<string>('');
  const [detailCardFilter, setDetailCardFilter] = useState<string | null>(null);
  const [comprehensiveSearch, setComprehensiveSearch] = useState<string>('');

  useEffect(() => {
    setDetailCardFilter(null);
  }, [monthlyEmployee, monthlyMonth, monthlyPeriodType, monthlyStartDate, monthlyEndDate]);

  // Synchronize global filters to monthly report filters so selections carry over seamlessly
  useEffect(() => {
    if (selectedEmployee && selectedEmployee !== 'all') {
      setMonthlyEmployee(selectedEmployee);
    }
  }, [selectedEmployee]);

  useEffect(() => {
    if (selectedPeriodType === 'range') {
      setMonthlyPeriodType('range');
      if (selectedStartDate) setMonthlyStartDate(selectedStartDate);
      if (selectedEndDate) setMonthlyEndDate(selectedEndDate);
    } else {
      setMonthlyPeriodType('month');
      if (selectedMonth && selectedMonth !== 'all') {
        setMonthlyMonth(selectedMonth);
      }
    }
  }, [selectedPeriodType, selectedStartDate, selectedEndDate, selectedMonth]);

  // Custom modals/popups
  const [modalOpen, setModalOpen] = useState(false);
  const [modalTitle, setModalTitle] = useState('');
  const [modalText, setModalText] = useState('');
  const [modalActiveRecordKey, setModalActiveRecordKey] = useState<string | null>(null);
  const [showClearConfirm, setShowClearConfirm] = useState(false);
  const [sentAlerts, setSentAlerts] = useState<Record<string, boolean>>({});
  const [formatStatus, setFormatStatus] = useState<'idle' | 'processing' | 'success' | 'error'>('idle');
  const [formatFileName, setFormatFileName] = useState<string>('');
  const [formatRowsCount, setFormatRowsCount] = useState<number>(0);
  const [pendingImport, setPendingImport] = useState<{
    rows: ParsedRow[];
    stats: ImportStats;
    fileDates: string[];
    duplicateDates: string[];
  } | null>(null);

  // Manual record edit/override state
  const [editingRecord, setEditingRecord] = useState<ParsedRow | null>(null);
  const [editCheckIn, setEditCheckIn] = useState<string>('');
  const [editCheckOut, setEditCheckOut] = useState<string>('');
  const [editType, setEditType] = useState<'present' | 'absence' | 'excuse'>('present');
  const [editStatusText, setEditStatusText] = useState<string>('');

  useEffect(() => {
    document.title = 'تقارير برنامج حضوري';
  }, []);

  useEffect(() => {
    if (editingRecord) {
      setEditCheckIn(editingRecord.tin?.text && editingRecord.tin.text !== '-' ? editingRecord.tin.text : '');
      setEditCheckOut(editingRecord.tout?.text && editingRecord.tout.text !== '-' ? editingRecord.tout.text : '');
      setEditType(editingRecord.type === 'missing' || editingRecord.type === 'other' ? 'present' : editingRecord.type);
      setEditStatusText(editingRecord.status || '');
    }
  }, [editingRecord]);

  const [newSuspendedDateInput, setNewSuspendedDateInput] = useState<string>('');

  const autoDetectedSuspendedDays = React.useMemo(() => {
    if (config.autoDetectSuspendedDays === false) return [];
    const holidays = holidaySet(config.holidays);
    const dateBiometricCounts = new Map<string, number>();

    rawRows.forEach((r: any) => {
      const dk = r.dateKey || dateKey(r.date);
      if (!dk) return;
      const tinText = r.tin?.text || (typeof r.tin === 'string' ? r.tin : '');
      const toutText = r.tout?.text || (typeof r.tout === 'string' ? r.tout : '');
      const hasIn = tinText && tinText !== '-';
      const hasOut = toutText && toutText !== '-';
      const st = norm(r.status);
      const isForce = ['استئذان', 'استئيذان', 'مهمة', 'انتداب', 'تكليف', 'مهمة عمل'].some(w => st.includes(w));
      if (hasIn || hasOut || isForce) {
        dateBiometricCounts.set(dk, (dateBiometricCounts.get(dk) || 0) + 1);
      }
    });

    const allDks = Array.from(new Set(rawRows.map((r: any) => r.dateKey || dateKey(r.date)).filter(Boolean))) as string[];
    const detected: string[] = [];
    allDks.forEach(dk => {
      const d = parseDate(dk);
      if (!d) return;
      if (isWeekend(d, config.weekend) || holidays.has(dk) || (config.suspendedDates || []).includes(dk)) {
        return;
      }
      if ((dateBiometricCounts.get(dk) || 0) === 0) {
        detected.push(dk);
      }
    });
    return detected.sort();
  }, [rawRows, config]);

  const toggleSuspendedDate = (dateKeyStr: string) => {
    const current = config.suspendedDates || [];
    const exists = current.includes(dateKeyStr);
    const updated = exists ? current.filter(d => d !== dateKeyStr) : [...current, dateKeyStr];
    setConfig(prev => ({ ...prev, suspendedDates: updated }));
  };

  const handleAddSuspendedDate = () => {
    if (!newSuspendedDateInput) return;
    const dk = dateKey(newSuspendedDateInput);
    if (!dk) return;
    if (!(config.suspendedDates || []).includes(dk)) {
      setConfig(prev => ({ ...prev, suspendedDates: [...(prev.suspendedDates || []), dk] }));
    }
    setNewSuspendedDateInput('');
  };

  const handleSaveEditedRecord = () => {
    if (!editingRecord) return;

    const { name, dateKey, date } = editingRecord;
    
    const tinObj = editCheckIn ? parseTime(editCheckIn) : null;
    const toutObj = editCheckOut ? parseTime(editCheckOut) : null;

    const updatedRows = [...rawRows];
    const existingIdx = updatedRows.findIndex(r => r.name === name && r.dateKey === dateKey);

    const targetType = editType;
    let targetStatus = editStatusText.trim();
    if (!targetStatus) {
      if (targetType === 'present') targetStatus = 'حاضر';
      else if (targetType === 'absence') targetStatus = 'بدون سجل بصمة';
      else if (targetType === 'excuse') targetStatus = 'غياب بعذر';
    }

    if (existingIdx > -1) {
      updatedRows[existingIdx] = {
        ...updatedRows[existingIdx],
        tin: tinObj,
        tout: toutObj,
        type: targetType,
        status: targetStatus,
      };
    } else {
      const empInfo = employees.find(e => e.name === name) || { job: 'موظف', civil: '' };
      const newRow = {
        name,
        job: empInfo.job,
        civil: empInfo.civil,
        date: date instanceof Date ? date.toISOString() : new Date(date).toISOString(),
        dateKey,
        month: monthKey(date instanceof Date ? date : new Date(date)),
        tin: tinObj,
        tout: toutObj,
        type: targetType,
        status: targetStatus,
      };
      updatedRows.push(newRow);
    }

    setRawRows(updatedRows);
    setEditingRecord(null);
  };

  // Computed collections
  const [employees, setEmployees] = useState<{ name: string; job: string; civil: string }[]>([]);
  const [dailyRecords, setDailyRecords] = useState<ParsedRow[]>([]);
  const [monthlySummaries, setMonthlySummaries] = useState<EmployeeSummary[]>([]);

  const fileInputRef = useRef<HTMLInputElement>(null);

  // 2. Persistent Sync Effects
  useEffect(() => {
    localStorage.setItem(LOCAL_STORAGE_KEYS.CONFIG, JSON.stringify(config));
  }, [config]);

  useEffect(() => {
    localStorage.setItem(LOCAL_STORAGE_KEYS.ROWS, JSON.stringify(rawRows));
  }, [rawRows]);

  useEffect(() => {
    localStorage.setItem(LOCAL_STORAGE_KEYS.FARES_SUBMITTED, JSON.stringify(faresSubmitted));
  }, [faresSubmitted]);

  useEffect(() => {
    localStorage.setItem(LOCAL_STORAGE_KEYS.MESSAGE_LOG, JSON.stringify(messageLogs));
  }, [messageLogs]);

  useEffect(() => {
    if (importStats) {
      localStorage.setItem(LOCAL_STORAGE_KEYS.IMPORT_STATS, JSON.stringify(importStats));
    } else {
      localStorage.removeItem(LOCAL_STORAGE_KEYS.IMPORT_STATS);
    }
  }, [importStats]);

  // 3. Process & Re-compute Core Data whenever raw rows or config changes
  useEffect(() => {
    if (rawRows.length === 0) {
      setEmployees([]);
      setDailyRecords([]);
      setMonthlySummaries([]);
      return;
    }

    // A. Re-instantiate Date objects in the raw records
    const parsedRows: ParsedRow[] = rawRows
      .map(r => {
        const d = r.date instanceof Date ? r.date : parseDate(r.date);
        const validD = d && !isNaN(d.getTime()) ? d : new Date();
        const dk = dateKey(validD);
        const mk = monthKey(validD);
        return {
          ...r,
          date: validD,
          dateKey: dk,
          month: mk,
          tin: r.tin ? { ...r.tin } : (typeof r.tin === 'string' ? parseTime(r.tin) : null),
          tout: r.tout ? { ...r.tout } : (typeof r.tout === 'string' ? parseTime(r.tout) : null),
          expected: r.expected,
        };
      })
      .filter(r => r.name);

    if (parsedRows.length === 0) {
      setEmployees([]);
      setDailyRecords([]);
      setMonthlySummaries([]);
      return;
    }

    // B. Group unique employees
    const empMap = new Map<string, { name: string; job: string; civil: string }>();
    parsedRows.forEach(r => {
      if (r.name && !empMap.has(r.name)) {
        empMap.set(r.name, { name: r.name, job: r.job || 'موظف', civil: r.civil || '' });
      }
    });
    const uniqueEmployees = Array.from(empMap.values()).sort((a, b) =>
      a.name.localeCompare(b.name, 'ar')
    );
    setEmployees(uniqueEmployees);

    // C. Reconstruct continuous daily records
    const validTimes = parsedRows.map(r => r.date.getTime()).filter(t => !isNaN(t));
    if (validTimes.length === 0) {
      setDailyRecords([]);
      setMonthlySummaries([]);
      return;
    }

    const minTime = Math.min(...validTimes);
    const maxTime = Math.max(...validTimes);
    const minDate = new Date(minTime);
    const maxDate = new Date(maxTime);

    // Build map of original imported records indexed by 'name|YYYY-MM-DD'
    const importedByEmpAndDate = new Map<string, ParsedRow>();
    const importedDatesSet = new Set<string>();
    const dateBiometricCounts = new Map<string, number>();

    parsedRows.forEach(r => {
      if (r.name && r.dateKey) {
        importedByEmpAndDate.set(`${r.name}|${r.dateKey}`, r);
        importedDatesSet.add(r.dateKey);

        const hasIn = !!(r.tin && r.tin.total != null && r.tin.text !== '-');
        const hasOut = !!(r.tout && r.tout.total != null && r.tout.text !== '-');
        const isEmergency = isEmergencyLeave(r.status);
        const forcePresentWords = ['مهمة', 'انتداب', 'تكليف', 'مهمة عمل'];
        const normalizedStatus = norm(r.status);
        const isForcePresent = forcePresentWords.some(w => normalizedStatus.includes(w));
        if (hasIn || hasOut || isForcePresent || isEmergency) {
          dateBiometricCounts.set(r.dateKey, (dateBiometricCounts.get(r.dateKey) || 0) + 1);
        }
      }
    });

    const dailyList: ParsedRow[] = [];
    const holidays = holidaySet(config.holidays);

    // If date range is reasonable (<= 365 days), iterate day by day
    const dayDifference = Math.round((maxDate.getTime() - minDate.getTime()) / (1000 * 3600 * 24));
    
    if (dayDifference >= 0 && dayDifference <= 365) {
      for (let d = new Date(minDate); d <= maxDate; d.setDate(d.getDate() + 1)) {
        const dk = dateKey(d);
        const isWeekendDay = isWeekend(d, config.weekend);
        const isHolidayDay = holidays.has(dk) || isHoliday(d, config.holidays);
        const isExplicitSuspended = (config.suspendedDates || []).includes(dk);

        // 1. Critical Safeguard: If this date was NOT present in the imported file at all,
        // NEVER manufacture mass absence records for the entire school staff!
        if (!importedDatesSet.has(dk)) {
          continue;
        }

        // 2. School-wide Zero-Biometrics Detection:
        // If a date is in the file, but 0 employees in the whole school have any biometric punch or assignment,
        // this indicates a suspended attendance day (rain/weather suspension, remote learning via مدرستي, or unrecorded day).
        const punchCount = dateBiometricCounts.get(dk) || 0;
        const isAutoSuspended = (config.autoDetectSuspendedDays !== false) && punchCount === 0;

        const isOffDay = isWeekendDay || isHolidayDay || isExplicitSuspended || isAutoSuspended;
        
        // NEVER include weekends, holidays, suspended days, or non-imported dates in attendance/absence calculations
        if (isOffDay) {
          continue;
        }

        const isRamadanDay = config.ramadanDates?.includes(dk) || false;
        const currentStart = isRamadanDay ? (config.ramadanStart || '09:00') : config.start;
        const currentWorkHours = isRamadanDay ? (config.ramadanWorkHours ?? 5) : config.workHours;
        const currentAutoCheckout = isRamadanDay ? (config.ramadanAutoCheckout || '08:45') : config.autoCheckout;

        uniqueEmployees.forEach(emp => {
          const key = `${emp.name}|${dk}`;
          const imported = importedByEmpAndDate.get(key);

          let baseRecord: ParsedRow;
          if (imported) {
            const hasIn = !!(imported.tin && imported.tin.total != null && imported.tin.text !== '-');
            const hasOut = !!(imported.tout && imported.tout.total != null && imported.tout.text !== '-');
            const hasBiometrics = hasIn || hasOut;

            let status = imported.status || '';
            const normalizedStatus = norm(status);

            const isFaresExcuse = faresSubmitted[key];
            const isEmergency = isEmergencyLeave(status);
            const forcePresentWords = ['مهمة', 'انتداب', 'تكليف', 'مهمة عمل'];
            const excuseWords = ['إجازة', 'اجازة', 'طبي', 'عذر', 'مرضي', 'مرافقة', 'وضع', 'أمومة', 'امومة', 'وفاة', 'استثنائية', 'اضطرارية', 'غياب بعذر'];

            const isForcePresent = forcePresentWords.some(w => normalizedStatus.includes(w));
            const isExcuse = isFaresExcuse || excuseWords.some(w => normalizedStatus.includes(w));

            let type: ParsedRow['type'] = 'absence';
            if (isEmergency) {
              type = hasBiometrics ? 'present' : 'excuse';
              status = status || 'استئذان طارئ';
            } else if (isForcePresent) {
              type = 'present';
              status = status || 'مهمة عمل';
            } else if (isExcuse) {
              type = 'excuse';
              status = isFaresExcuse ? 'غياب بعذر (فارس)' : (status || 'غياب بعذر');
            } else if (hasBiometrics) {
              type = 'present';
              status = (status && !['غير مكتملة', 'غير مكتمله', 'انصراف تلقائي', 'بدون سجل بصمة', 'غائب', 'غياب'].includes(status)) ? status : 'حاضر';
            } else {
              // No punch in, no punch out, no official excuse/duty -> ABSENT!
              type = 'absence';
              status = (status && (status.includes('غياب') || status.includes('غائب'))) ? status : 'غياب بدون عذر';
            }

            let lateMins = 0;
            let earlyMins = 0;
            let checkState = '';
            let expCheck = null;

            const approvedSunWed = config.checkoutSunWed || '12:30';
            const approvedThu = config.checkoutThu || '11:45';

            if (type === 'present' || isEmergency) {
              lateMins = minsAfterStart(imported.tin, currentStart);
              expCheck = expectedCheckout(imported.tin, currentWorkHours, d, approvedSunWed, approvedThu);
              checkState = checkoutState(status, imported.tin, imported.tout, currentAutoCheckout, currentWorkHours, d, approvedSunWed, approvedThu);
              earlyMins = earlyCheckoutMins(imported.tin, imported.tout, status, currentAutoCheckout, currentWorkHours, d, approvedSunWed, approvedThu);

              if (isForcePresent && !hasBiometrics) {
                // Official assignment without biometric punches
                lateMins = 0;
                earlyMins = 0;
                checkState = '';
              }
            }

            baseRecord = {
              ...imported,
              status,
              type,
              tin: hasIn ? imported.tin : null,
              tout: hasOut ? imported.tout : null,
              late: lateMins,
              expected: expCheck,
              checkout: checkState,
              earlyMins,
            };
          } else {
            // No punch card for this day -> treated as Absence by default
            const isFaresExcuse = faresSubmitted[key];
            baseRecord = {
              name: emp.name,
              job: emp.job,
              civil: emp.civil,
              date: new Date(d),
              dateKey: dk,
              month: monthKey(d),
              status: isFaresExcuse ? 'غياب بعذر (فارس)' : 'غياب بدون عذر',
              tin: null,
              tout: null,
              expected: null,
              checkout: '',
              earlyMins: 0,
              type: isFaresExcuse ? 'excuse' : 'absence',
              late: 0,
            };
          }

          dailyList.push(baseRecord);
        });
      }
    } else {
      // Fallback if dates are disjoint or span across years: use imported records directly
      parsedRows.forEach(imported => {
        const dk = imported.dateKey;
        const isWeekendDay = isWeekend(imported.date, config.weekend);
        const isHolidayDay = holidays.has(dk) || isHoliday(imported.date, config.holidays);
        const isExplicitSuspended = (config.suspendedDates || []).includes(dk);
        const punchCount = dateBiometricCounts.get(dk) || 0;
        const isAutoSuspended = (config.autoDetectSuspendedDays !== false) && punchCount === 0;

        if (isWeekendDay || isHolidayDay || isExplicitSuspended || isAutoSuspended) {
          return; // Skip weekend/holiday/suspended completely
        }
        dailyList.push(imported);
      });
    }

    setDailyRecords(dailyList);

    // D. Compute monthly aggregate scores
    const monthlyMap = new Map<string, EmployeeSummary>();
    dailyList.forEach(r => {
      const isOff = isWeekend(r.date, config.weekend) || holidays.has(r.dateKey);
      if (isOff && r.type !== 'present') {
        return; // Never count non-working weekend/holiday rows as work or absence
      }

      const groupKey = `${r.name}|${r.month}`;
      if (!monthlyMap.has(groupKey)) {
        monthlyMap.set(groupKey, {
          name: r.name,
          job: r.job,
          month: r.month,
          work: 0,
          present: 0,
          absence: 0,
          excuse: 0,
          lateCount: 0,
          lateMins: 0,
          completeCount: 0,
          earlyCount: 0,
          earlyMins: 0,
          autoCount: 0,
          missingCheckoutCount: 0,
          wrongTimeExitsCount: 0,
          emergencyNoExitCount: 0,
          emergencyWithExitCount: 0,
        });
      }

      const summary = monthlyMap.get(groupKey)!;

      if (r.type === 'present') {
        summary.work++;
        summary.present++;
        if (!isOff) {
          if (r.late > 0) {
            summary.lateCount++;
            summary.lateMins += r.late;
          }
          if (r.checkout === 'مكتمل') {
            summary.completeCount++;
          }
          if (r.checkout === 'خروج قبل نهاية الدوام' || r.checkout === 'انصراف مبكر') {
            summary.earlyCount++;
            summary.earlyMins += r.earlyMins || 0;
          }
          if (r.checkout === 'انصراف تلقائي') {
            summary.autoCount++;
          }
          if (r.checkout === 'لا توجد بصمة انصراف') {
            summary.missingCheckoutCount++;
          }
          if (r.checkout === 'انصراف في غير وقت الفعلي') {
            summary.wrongTimeExitsCount = (summary.wrongTimeExitsCount || 0) + 1;
          }
          if (r.checkout === 'استئذان بدون بصمة انصراف') {
            summary.emergencyNoExitCount = (summary.emergencyNoExitCount || 0) + 1;
          }
          if (r.checkout === 'استئذان طارئ موثق (بصمة)') {
            summary.emergencyWithExitCount = (summary.emergencyWithExitCount || 0) + 1;
          }
        } else {
          summary.completeCount++;
        }
      } else if (r.type === 'excuse') {
        summary.work++;
        summary.excuse++;
        if (r.checkout === 'استئذان بدون بصمة انصراف') {
          summary.emergencyNoExitCount = (summary.emergencyNoExitCount || 0) + 1;
        }
        if (r.checkout === 'استئذان طارئ موثق (بصمة)') {
          summary.emergencyWithExitCount = (summary.emergencyWithExitCount || 0) + 1;
        }
      } else if (r.type === 'absence') {
        summary.work++;
        summary.absence++;
      }
    });

    const summaries = Array.from(monthlyMap.values()).map(m => {
      const attendanceRate = m.work ? Math.round(((m.present + m.excuse) / m.work) * 100) : 0;
      const classification = disciplineClassify(m);
      return {
        ...m,
        attendanceRate,
        rate: classification.score,
        rating: classification.disciplineLabel,
        disciplineCategory: classification.disciplineCategory,
        disciplineReason: classification.reason,
      };
    });
    setMonthlySummaries(summaries);
  }, [rawRows, faresSubmitted, config]);

  // 4. File Reading / Excel Import Handler
  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onerror = () => alert('تعذر قراءة الملف من الجهاز.');
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        
        let parsedList: ParsedRow[] = [];
        let ignoredNoName = 0;
        let ignoredNoDate = 0;
        let totalRawRowsCount = 0;

        let stats: ImportStats = {
          totalRows: 0,
          imported: 0,
          present: 0,
          absence: 0,
          excuse: 0,
          other: 0,
          missing: 0,
          noName: 0,
          noDate: 0,
          ignored: 0
        };

        const presentWords = ['مكتملة', 'مكتمله', 'غير مكتملة', 'غير مكتمله', 'انصراف تلقائي', 'حاضر', 'حضور'];
        const excuseWords = ['استئذان', 'استئيذان', 'إجازة', 'اجازة', 'انتداب', 'مهمة', 'طبي', 'عذر', 'مرضي', 'مهمة عمل'];

        // Collect all sheets that contain relevant attendance/employee data
        const sheetsToProcess: { sheetName: string; sheetJson: any[][]; headerIndex: number }[] = [];

        for (const sheetName of workbook.SheetNames) {
          const ws = workbook.Sheets[sheetName];
          const sheetJson: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
          if (!sheetJson || sheetJson.length === 0) continue;

          // Locate header row in this sheet
          const hIdx = sheetJson.findIndex(r =>
            Array.isArray(r) && r.some((c: any) => {
              const normalized = normKey(c);
              return (
                normalized.includes('السجل') ||
                normalized.includes('الاسم') ||
                normalized.includes('الموظف') ||
                normalized.includes('المعلم') ||
                normalized.includes('حضور') ||
                normalized.includes('تاريخ') ||
                normalized.includes('هويه') ||
                normalized.includes('الهوية') ||
                normalized.includes('بصمه') ||
                normalized.includes('تحضير')
              );
            })
          );

          if (hIdx >= 0) {
            sheetsToProcess.push({ sheetName, sheetJson, headerIndex: hIdx });
          }
        }

        // Fallback to first sheet if no specific header was detected
        if (sheetsToProcess.length === 0 && workbook.SheetNames.length > 0) {
          const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
          const sheetJson: any[][] = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: '' });
          if (sheetJson && sheetJson.length > 0) {
            sheetsToProcess.push({ sheetName: workbook.SheetNames[0], sheetJson, headerIndex: 0 });
          }
        }

        if (sheetsToProcess.length === 0) {
          alert('ملف Excel فارغ أو غير متوافق.');
          return;
        }

        // Process every identified sheet (e.g. September sheet + October sheet)
        for (const { sheetName, sheetJson, headerIndex } of sheetsToProcess) {
          // Try extracting any global date embedded in header/title rows or sheet name
          let globalDateFromTitle: Date | null = parseDate(sheetName);
          for (let i = 0; i < headerIndex; i++) {
            const rowText = (sheetJson[i] || []).join(' ');
            const extractedD = parseDate(rowText);
            if (extractedD) {
              globalDateFromTitle = extractedD;
              break;
            }
          }

          const headerRow = (sheetJson[headerIndex] || []).map((h: any) => normKey(h) || '');
          const dataRows = sheetJson.slice(headerIndex + 1);
          totalRawRowsCount += dataRows.length;

          const objects = dataRows.map((r: any[]) => {
            const obj: any = {};
            headerRow.forEach((h: string, idx: number) => {
              if (h) obj[h] = r[idx];
            });
            return obj;
          });

          let lastValidDate: Date | null = globalDateFromTitle;

          objects.forEach(o => {
            const name = norm(getVal(o, [
              'الإسم', 'الاسم', 'اسم الموظف', 'الموظف', 'اسم المعلم', 'المعلم', 
              'اسم المستخدم', 'اسم المنسوب', 'المنسوب', 'المستخدم', 'الاسم الكامل', 
              'اسم الموظف/المعلم', 'اسم المعلم/الموظف', 'Name', 'Employee Name'
            ]));

            if (!name || name === '-' || name.includes('الإجمالي') || name.includes('المجموع')) {
              ignoredNoName++;
              return;
            }

            const job = norm(getVal(o, [
              'المسمى الوظيفي', 'المسمي الوظيفي', 'الوظيفة', 'الوظيفه', 'المسمى', 
              'الصفة', 'الرتبة', 'المهنة', 'Job', 'Title', 'Position'
            ])) || 'معلم / موظف';

            const civil = norm(getVal(o, [
              'السجل المدني', 'رقم الهوية', 'الهويه', 'الهوية', 'السجل', 
              'رقم السجل المدني', 'الهوية الوطنية', 'الرقم الوظيفي', 'رقم الموظف', 'Civil ID', 'National ID', 'ID'
            ]));

            let rawDateVal = getVal(o, [
              'التاريخ', 'اليوم والتاريخ', 'تاريخ التحضير', 'تاريخ الحضور', 
              'تاريخ الحركة', 'تاريخ اليوم', 'تاريخ التسجيل', 'اليوم', 'التاريخ والوقت', 'Date',
              'تاريخ البصمة', 'تاريخ الحضور والانصراف', 'تاريخ الحضور والإنصراف', 'التاريخ الميلادي', 'التاريخ الهجري', 'تاريخ المعاملة', 'تاريخ السجل'
            ]);

            let d = parseDate(rawDateVal);

            // If row has no explicit date, inherit last valid date (for grouped/merged rows) or global title date
            if (!d) {
              d = lastValidDate;
            } else {
              lastValidDate = d;
            }

            if (!d) {
              d = new Date();
            }

            const status = norm(getVal(o, [
              'حالة التحضير', 'حاله التحضير', 'الحالة', 'الحاله', 'حالة الحضور', 
              'حالة البصمة', 'نوع التحضير', 'البيان', 'الوضع', 'حالة السجل', 'Status'
            ]));

            const tin = parseTime(getVal(o, [
              'توقيت الحضور', 'وقت الحضور', 'الحضور', 'الدخول', 'توقيت الدخول', 
              'وقت الدخول', 'بصمة الدخول', 'أول بصمة', 'بصمة الحضور', 'حضور', 'دخول', 
              'وقت الحضور الفعلي', 'Check In', 'Check-in', 'Time In', 'In'
            ]));

            const tout = parseTime(getVal(o, [
              'توقيت الإنصراف', 'توقيت الانصراف', 'وقت الانصراف', 'وقت الإنصراف', 
              'الانصراف', 'الإنصراف', 'الخروج', 'وقت الخروج', 'توقيت الخروج', 
              'بصمة الخروج', 'آخر بصمة', 'بصمة الانصراف', 'انصراف', 'خروج', 
              'وقت الانصراف الفعلي', 'Check Out', 'Check-out', 'Time Out', 'Out'
            ]));

            // Determine record type
            const normalizedStatus = norm(status);
            const hasIn = !!(tin && tin.total != null && tin.text !== '-');
            const hasOut = !!(tout && tout.total != null && tout.text !== '-');
            const hasBiometrics = hasIn || hasOut;

            const isEmergency = isEmergencyLeave(status);
            const forcePresentWords = ['مهمة', 'انتداب', 'تكليف', 'مهمة عمل'];
            const forceExcuseWords = ['إجازة', 'اجازة', 'طبي', 'عذر', 'مرضي', 'مرافقة', 'وضع', 'أمومة', 'امومة', 'وفاة', 'استثنائية', 'اضطرارية', 'غياب بعذر'];
            const weekendWords = ['عطلة', 'عطله', 'نهاية الاسبوع', 'نهاية الأسبوع', 'راحة أسبوعية', 'راحة اسبوعية', 'إجازة أسبوعية', 'اجازة اسبوعية'];

            const isForcePresent = forcePresentWords.some(w => normalizedStatus.includes(w));
            const isExcuse = forceExcuseWords.some(w => normalizedStatus.includes(w));
            const isWeekendDay = isWeekend(d, config.weekend);
            const isOfficialHoliday = isHoliday(d, config.holidays);

            let type: ParsedRow['type'] = 'absence';
            let finalStatus = status;

            if (weekendWords.some(w => normalizedStatus.includes(w)) || isWeekendDay || isOfficialHoliday) {
              type = 'other';
              finalStatus = 'عطلة أسبوعية / رسمية';
            } else if (isEmergency) {
              type = hasBiometrics ? 'present' : 'excuse';
              finalStatus = status || 'استئذان طارئ';
            } else if (isForcePresent) {
              type = 'present';
              finalStatus = status || 'مهمة عمل';
            } else if (isExcuse) {
              type = 'excuse';
              finalStatus = status || 'غياب بعذر';
            } else if (hasBiometrics) {
              type = 'present';
              finalStatus = (status && !['غير مكتملة', 'غير مكتمله', 'انصراف تلقائي', 'بدون سجل بصمة', 'غائب', 'غياب'].includes(status)) ? status : 'حاضر';
            } else {
              type = 'absence';
              finalStatus = (status && (status.includes('غياب') || status.includes('غائب'))) ? status : 'غياب بدون عذر';
            }

            let checkState = '';
            let earlyMins = 0;
            let lateMins = 0;

            const approvedSunWed = config.checkoutSunWed || '12:30';
            const approvedThu = config.checkoutThu || '11:45';

            if (type === 'present' || isEmergency) {
              checkState = checkoutState(finalStatus, tin, tout, config.autoCheckout, config.workHours, d, approvedSunWed, approvedThu);
              earlyMins = earlyCheckoutMins(tin, tout, finalStatus, config.autoCheckout, config.workHours, d, approvedSunWed, approvedThu);
              lateMins = minsAfterStart(tin, config.start);

              if (isForcePresent && !hasBiometrics) {
                checkState = '';
                earlyMins = 0;
                lateMins = 0;
              }
            }

            parsedList.push({
              name,
              job,
              civil,
              date: d,
              dateKey: dateKey(d),
              month: monthKey(d),
              status: finalStatus,
              tin: hasIn ? tin : null,
              tout: hasOut ? tout : null,
              expected: expectedCheckout(tin, config.workHours, d, approvedSunWed, approvedThu),
              checkout: checkState,
              earlyMins,
              type,
              late: lateMins,
            });

            stats.imported++;
            if (type === 'present') stats.present++;
            else if (type === 'absence') stats.absence++;
            else if (type === 'excuse') stats.excuse++;
            else if (type === 'other') stats.other++;
            else stats.missing++;
          });
        }

        stats.totalRows = totalRawRowsCount;
        stats.noName = ignoredNoName;
        stats.noDate = ignoredNoDate;
        stats.ignored = ignoredNoName + ignoredNoDate;

        if (parsedList.length === 0) {
          alert('تعذر استخراج بيانات الموظفين من ملف Excel. يرجى التأكد من أن الملف يحتوي على أعمدة الأسماء والحضور.');
          return;
        }

        // 1. Identify all unique date keys from the imported parsedList
        const fileDatesSet = new Set<string>();
        parsedList.forEach(r => {
          if (r.dateKey) fileDatesSet.add(r.dateKey);
        });
        const fileDates = Array.from(fileDatesSet).sort();

        // 2. Identify which of these dates already exist in the database (rawRows)
        const existingDatesSet = new Set<string>();
        rawRows.forEach((r: any) => {
          if (r.dateKey) existingDatesSet.add(r.dateKey);
        });
        const duplicateDates = fileDates.filter(d => existingDatesSet.has(d));

        // 3. Set the pending import state
        setPendingImport({
          rows: parsedList,
          stats,
          fileDates,
          duplicateDates
        });

        // Reset file input
        if (fileInputRef.current) {
          fileInputRef.current.value = '';
        }
      } catch (err: any) {
        console.error(err);
        alert('حدث خطأ أثناء تحليل ملف Excel: ' + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  // 4b. Format / Clean Fingerprint Excel File (Removes F, H, J, L, M, N, O, P, Q, R, S, T, U)
  const handleFormatExcelFile = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setFormatFileName(file.name);
    setFormatStatus('processing');

    const reader = new FileReader();
    reader.onerror = () => {
      alert('تعذر قراءة الملف من الجهاز.');
      setFormatStatus('error');
    };
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array', raw: true });
        
        let totalProcessedRows = 0;

        workbook.SheetNames.forEach(sheetName => {
          const worksheet = workbook.Sheets[sheetName];
          const rawJson: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
          
          if (!rawJson || rawJson.length === 0) return;
          
          totalProcessedRows += rawJson.length;

          // Indices to delete (0-based):
          // F(5), H(7), J(9), L(11), M(12), N(13), O(14), P(15), Q(16), R(17), S(18), T(19), U(20)
          const indicesToDelete = [5, 7, 9, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20];
          
          const processedRows = rawJson.map(row => {
            return row.filter((_, colIdx) => !indicesToDelete.includes(colIdx));
          });
          
          const newWorksheet = XLSX.utils.aoa_to_sheet(processedRows);
          workbook.Sheets[sheetName] = newWorksheet;
        });

        const outData = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([outData], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `ملف_البصمة_المنسق_${file.name.replace(/\.[^/.]+$/, "")}.xlsx`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);

        setFormatRowsCount(totalProcessedRows);
        setFormatStatus('success');

        // Reset the input value so user can upload again
        e.target.value = '';
      } catch (err: any) {
        console.error(err);
        alert('حدث خطأ أثناء معالجة وتنسيق الملف: ' + err.message);
        setFormatStatus('error');
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const handleClearData = () => {
    setRawRows([]);
    setImportStats(null);
    setFaresSubmitted({});
    setMessageLogs([]);
    setSelectedEmployee('all');
    setSelectedMonth('all');
    setSelectedPeriodType('month');
    setSelectedStartDate('');
    setSelectedEndDate('');
    setSelectedStatus('all');
    setSearchTerm('');
    setMonthlyEmployee('all');
    setMonthlyMonth('all');
    setShowClearConfirm(false);
  };

  const handleSavePendingImport = () => {
    if (!pendingImport) return;

    // Filter out any existing rows in rawRows that share a dateKey with the incoming rows, then append all new rows
    const merged = rawRows.filter((r: any) => !pendingImport.fileDates.includes(r.dateKey));
    merged.push(...pendingImport.rows);

    setRawRows(merged);
    setImportStats(pendingImport.stats);
    setPendingImport(null);
    setSelectedEmployee('all');
    setSelectedMonth('all');
    setSelectedPeriodType('month');
    setSelectedStartDate('');
    setSelectedEndDate('');
    setSelectedStatus('all');
    setSearchTerm('');
    setMonthlyEmployee('all');
    setMonthlyMonth('all');
    setCurrentTab('dashboard');
  };

  // Ramadan dynamic toggles & bulk actions
  const toggleRamadanDate = (dateStr: string) => {
    const current = config.ramadanDates || [];
    let updated: string[];
    if (current.includes(dateStr)) {
      updated = current.filter(d => d !== dateStr);
    } else {
      updated = [...current, dateStr];
    }
    setConfig({ ...config, ramadanDates: updated });
  };

  const setBulkRamadanDatesByMonth = (monthName: string) => {
    const targetDates = dailyRecords
      .filter(r => r.month === monthName)
      .map(r => r.dateKey);
    const uniqueTargetDates = Array.from(new Set(targetDates));

    const current = config.ramadanDates || [];
    const updated = Array.from(new Set([...current, ...uniqueTargetDates]));
    setConfig({ ...config, ramadanDates: updated });
  };

  const clearAllRamadanDates = () => {
    setConfig({ ...config, ramadanDates: [] });
  };

  const setAllAsRamadanDates = () => {
    const systemDates = Array.from(new Set(dailyRecords.map(r => r.dateKey)));
    setConfig({ ...config, ramadanDates: systemDates });
  };

  const setRamadanDateRange = (startDateStr: string, endDateStr: string) => {
    const start = new Date(startDateStr);
    const end = new Date(endDateStr);
    if (isNaN(start.getTime()) || isNaN(end.getTime())) return;

    const datesInRange: string[] = [];
    const current = new Date(start);
    while (current <= end) {
      const y = current.getFullYear();
      const m = String(current.getMonth() + 1).padStart(2, '0');
      const d = String(current.getDate()).padStart(2, '0');
      datesInRange.push(`${y}-${m}-${d}`);
      current.setDate(current.getDate() + 1);
    }

    const currentRamadan = config.ramadanDates || [];
    const updated = Array.from(new Set([...currentRamadan, ...datesInRange]));
    setConfig({ ...config, ramadanDates: updated });
  };

  const formatArabicDate = (dateKeyStr: string) => {
    try {
      const parts = dateKeyStr.split('-');
      if (parts.length === 3) {
        const d = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
        return d.toLocaleDateString('ar-EG', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });
      }
    } catch (e) {}
    return dateKeyStr;
  };

  // 5. Navigation Filters and Selection Methods
  const getMonthsList = () => {
    return Array.from(new Set(dailyRecords.map(r => r.month))).sort();
  };

  const getEmployeesNames = () => {
    return Array.from(new Set(dailyRecords.map(r => r.name))).sort((a: string, b: string) =>
      a.localeCompare(b, 'ar')
    );
  };

  const filterDailyRecords = (records: ParsedRow[]) => {
    return records.filter(r => {
      // Exclude weekends and public holidays
      if (isWeekend(r.date, config.weekend) || isHoliday(r.date, config.holidays)) {
        return false;
      }

      let matchMonth = true;
      if (selectedPeriodType === 'range') {
        const dk = dateKey(r.date);
        const matchStart = !selectedStartDate || dk >= selectedStartDate;
        const matchEnd = !selectedEndDate || dk <= selectedEndDate;
        matchMonth = matchStart && matchEnd;
      } else {
        matchMonth = selectedMonth === 'all' || r.month === selectedMonth;
      }
      const matchEmployee = selectedEmployee === 'all' || r.name === selectedEmployee;
      const matchSearch = searchTerm === '' ||
        normKey(r.name).includes(normKey(searchTerm)) ||
        normKey(r.job).includes(normKey(searchTerm)) ||
        normKey(r.civil || '').includes(normKey(searchTerm));

      let matchStatus = true;
      if (selectedStatus === 'early') {
        matchStatus = r.checkout === 'خروج قبل نهاية الدوام' || r.checkout === 'انصراف مبكر' || r.earlyMins > 0;
      } else if (selectedStatus === 'late') {
        matchStatus = r.late > 0;
      } else if (selectedStatus === 'absence') {
        matchStatus = r.type === 'absence';
      } else if (selectedStatus === 'auto') {
        matchStatus = r.checkout === 'انصراف تلقائي';
      } else if (selectedStatus === 'missingCheckout') {
        matchStatus = r.checkout === 'لا توجد بصمة انصراف' || (r.type === 'present' && (!r.tout || r.tout.total == null));
      } else if (selectedStatus === 'emergency_no_exit') {
        matchStatus = r.checkout === 'استئذان بدون بصمة انصراف' || (isEmergencyLeave(r.status) && (!r.tout || r.tout.total == null || r.tout.text === '-'));
      } else if (selectedStatus === 'emergency_with_exit') {
        matchStatus = r.checkout === 'استئذان طارئ موثق (بصمة)' || (isEmergencyLeave(r.status) && !!(r.tout && r.tout.total != null && r.tout.text !== '-'));
      } else if (selectedStatus === 'wrong_time') {
        matchStatus = r.checkout === 'انصراف في غير وقت الفعلي';
      }

      return matchMonth && matchEmployee && matchSearch && matchStatus;
    });
  };

  const filterMonthlySummaries = (summaries: EmployeeSummary[]) => {
    if (selectedPeriodType === 'range') {
      // Dynamically compute range-based summaries directly from dailyRecords matching range & filters
      const matchingDaily = dailyRecords.filter(r => {
        if (isWeekend(r.date, config.weekend) || isHoliday(r.date, config.holidays)) {
          return false;
        }
        const dk = dateKey(r.date);
        const matchStart = !selectedStartDate || dk >= selectedStartDate;
        const matchEnd = !selectedEndDate || dk <= selectedEndDate;
        const matchEmployee = selectedEmployee === 'all' || r.name === selectedEmployee;
        const matchSearch = searchTerm === '' ||
          normKey(r.name).includes(normKey(searchTerm)) ||
          normKey(r.job).includes(normKey(searchTerm)) ||
          normKey(r.civil || '').includes(normKey(searchTerm));
        
        return matchStart && matchEnd && matchEmployee && matchSearch;
      });

      const holidays = holidaySet(config.holidays);
      const empMap = new Map<string, EmployeeSummary>();
      matchingDaily.forEach(r => {
        const isOff = isWeekend(r.date, config.weekend) || holidays.has(r.dateKey);
        if (isOff && r.type !== 'present') {
          return;
        }

        if (!empMap.has(r.name)) {
          empMap.set(r.name, {
            name: r.name,
            job: r.job,
            civil: r.civil,
            month: `من ${selectedStartDate || 'البداية'} إلى ${selectedEndDate || 'النهاية'}`,
            work: 0,
            present: 0,
            absence: 0,
            excuse: 0,
            lateCount: 0,
            lateMins: 0,
            completeCount: 0,
            earlyCount: 0,
            earlyMins: 0,
            autoCount: 0,
            missingCheckoutCount: 0,
            wrongTimeExitsCount: 0,
            emergencyNoExitCount: 0,
            emergencyWithExitCount: 0,
          });
        }
        const s = empMap.get(r.name)!;
        if (r.type === 'present') {
          s.work++;
          s.present++;
          if (!isOff) {
            if (r.late > 0) {
              s.lateCount++;
              s.lateMins += r.late;
            }
            if (r.checkout === 'مكتمل') {
              s.completeCount++;
            }
            if (r.checkout === 'خروج قبل نهاية الدوام' || r.checkout === 'انصراف مبكر') {
              s.earlyCount++;
              s.earlyMins += r.earlyMins || 0;
            }
            if (r.checkout === 'انصراف تلقائي') {
              s.autoCount++;
            }
            if (r.checkout === 'لا توجد بصمة انصراف') {
              s.missingCheckoutCount++;
            }
            if (r.checkout === 'انصراف في غير وقت الفعلي') {
              s.wrongTimeExitsCount = (s.wrongTimeExitsCount || 0) + 1;
            }
            if (r.checkout === 'استئذان بدون بصمة انصراف') {
              s.emergencyNoExitCount = (s.emergencyNoExitCount || 0) + 1;
            }
            if (r.checkout === 'استئذان طارئ موثق (بصمة)') {
              s.emergencyWithExitCount = (s.emergencyWithExitCount || 0) + 1;
            }
          } else {
            s.completeCount++;
          }
        } else if (r.type === 'excuse') {
          s.work++;
          s.excuse++;
          if (r.checkout === 'استئذان بدون بصمة انصراف') {
            s.emergencyNoExitCount = (s.emergencyNoExitCount || 0) + 1;
          }
          if (r.checkout === 'استئذان طارئ موثق (بصمة)') {
            s.emergencyWithExitCount = (s.emergencyWithExitCount || 0) + 1;
          }
        } else if (r.type === 'absence') {
          s.work++;
          s.absence++;
        }
      });

      const rangeSummaries = Array.from(empMap.values()).map(m => {
        const attendanceRate = m.work ? Math.round(((m.present + m.excuse) / m.work) * 100) : 0;
        const classification = disciplineClassify(m);
        return {
          ...m,
          attendanceRate,
          rate: classification.score,
          rating: classification.disciplineLabel,
          disciplineCategory: classification.disciplineCategory,
          disciplineReason: classification.reason,
        };
      });

      return rangeSummaries.filter(s => {
        let matchStatus = true;
        if (selectedStatus === 'early') {
          matchStatus = s.earlyCount > 0;
        } else if (selectedStatus === 'late') {
          matchStatus = s.lateCount > 0;
        } else if (selectedStatus === 'absence') {
          matchStatus = s.absence > 0;
        } else if (selectedStatus === 'auto') {
          matchStatus = s.autoCount > 0;
        } else if (selectedStatus === 'missingCheckout') {
          matchStatus = s.missingCheckoutCount > 0;
        } else if (selectedStatus === 'emergency_no_exit') {
          matchStatus = (s.emergencyNoExitCount || 0) > 0;
        } else if (selectedStatus === 'emergency_with_exit') {
          matchStatus = (s.emergencyWithExitCount || 0) > 0;
        }
        return matchStatus;
      });
    }

    return summaries.filter(s => {
      const matchMonth = selectedMonth === 'all' || s.month === selectedMonth;
      const matchEmployee = selectedEmployee === 'all' || s.name === selectedEmployee;
      const matchSearch = searchTerm === '' ||
        normKey(s.name).includes(normKey(searchTerm)) ||
        normKey(s.job).includes(normKey(searchTerm));

      let matchStatus = true;
      if (selectedStatus === 'early') {
        matchStatus = s.earlyCount > 0;
      } else if (selectedStatus === 'late') {
        matchStatus = s.lateCount > 0;
      } else if (selectedStatus === 'absence') {
        matchStatus = s.absence > 0;
      } else if (selectedStatus === 'auto') {
        matchStatus = s.autoCount > 0;
      } else if (selectedStatus === 'missingCheckout') {
        matchStatus = s.missingCheckoutCount > 0;
      } else if (selectedStatus === 'emergency_no_exit') {
        matchStatus = (s.emergencyNoExitCount || 0) > 0;
      } else if (selectedStatus === 'emergency_with_exit') {
        matchStatus = (s.emergencyWithExitCount || 0) > 0;
      }

      return matchMonth && matchEmployee && matchSearch && matchStatus;
    });
  };

  // 6. Action Toggles and Utilities
  const toggleFaresStatus = (recordKey: string) => {
    setFaresSubmitted(prev => {
      const updated = { ...prev, [recordKey]: !prev[recordKey] };
      return updated;
    });
  };

  const getRecordByKey = (key: string): ParsedRow | undefined => {
    return dailyRecords.find(r => `${r.name}|${r.dateKey}` === key);
  };

  const getMessageText = (r: ParsedRow): string => {
    if (!r) return '';
    const dStr = arDate(r.date);
    const dName = dayName(r.date);
    const submitted = faresSubmitted[`${r.name}|${r.dateKey}`];

    // Priority 1: Emergency Leave without exit punch (Teacher took emergency leave but didn't punch departure)
    if (
      r.checkout === 'استئذان بدون بصمة انصراف' ||
      (isEmergencyLeave(r.status) && (!r.tout || r.tout.total == null || r.tout.text === '-'))
    ) {
      return `السلام عليكم ورحمة الله وبركاته\n\nالأستاذ الموقر: ${r.name}\n\nنود إحاطتكم بأنه تم رصد تسجيل (استئذان طارئ) ليوم ${dName} الموافق ${dStr}، ولكن لوحظ عدم تسجيل بصمة الانصراف في جهاز البصمة عند مغادرتكم المدرسة، وهو تصرف غير مكتمل وإجراء ملزم لتوثيق وقت الخروج الفعلي نظاماً.\n\nنأمل التكرم بالحرص على تسجيل بصمة الخروج عند الاستئذان مستقبلاً ومراجعة إدارة المدرسة للتأكيد.\n\nشاكرين لكم حسن تعاونكم الدائم.`;
    }

    if (submitted) {
      return `السلام عليكم ورحمة الله وبركاته\n\nالأستاذ الموقر: ${r.name}\n\nتم تقييد غيابكم ليوم ${dName} الموافق ${dStr} كغياب بعذر في نظام متابعة حضور وانضباط منسوبي المدرسة بعد تقديم العذر المعتمد.\n\nشاكرين لكم التزامكم الدائم.`;
    }

    if (r.type === 'absence') {
      return `السلام عليكم ورحمة الله وبركاته\n\nالأستاذ الموقر: ${r.name}\n\nيرجى التكرم بتقديم مبرر الغياب المعتمد ليوم ${dName} الموافق ${dStr} في نظام فارس وتقديمه للموجه الطلابي لاعتماده.\n\nشاكرين لكم تعاونكم وحرصكم المستمر.`;
    }

    if (r.late > 0) {
      return `السلام عليكم ورحمة الله وبركاته\n\nالأستاذ الموقر: ${r.name}\n\nنود إحاطتكم بأنه تم رصد تأخر في الحضور ليوم ${dName} الموافق ${dStr} بمقدار ${r.late} دقيقة عن الموعد المعتمد (${config.start}).\n\nنأمل منكم الالتزام بمواعيد الحضور لتعزيز انتظام العمل المدرسي.`;
    }

    if (r.checkout === 'خروج قبل نهاية الدوام' || r.checkout === 'انصراف مبكر') {
      return `السلام عليكم ورحمة الله وبركاته\n\nالأستاذ الموقر: ${r.name}\n\nنود إحاطتكم بأنه تم رصد خروج قبل نهاية الدوام ليوم ${dName} الموافق ${dStr} بمقدار ${r.earlyMins} دقيقة قبل اكتمال الساعات المقررة للدوام.\n\nشاكرين لكم تعاونكم المعهود.`;
    }

    if (r.checkout === 'انصراف في غير وقت الفعلي') {
      return `السلام عليكم ورحمة الله وبركاته\n\nالأستاذ الموقر: ${r.name}\n\nنود إحاطتكم بأنه تم رصد بصمة انصراف في غير وقتها الفعلي (بصمة صباحية الساعة ${r.tout?.text || ''}) ليوم ${dName} الموافق ${dStr} بدلاً من موعد الانصراف المعتمد، مما يشير إلى عدم تسجيل بصمة الانصراف في الموعد المحدد.\n\nشاكرين لكم تعاونكم ومراجعة إدارة المدرسة للتأكيد.`;
    }

    if (r.checkout === 'انصراف تلقائي') {
      return `السلام عليكم ورحمة الله وبركاته\n\nالأستاذ الموقر: ${r.name}\n\nنود تنبيهكم إلى رصد حالة انصراف تلقائي ليوم ${dName} الموافق ${dStr} (انقضاء وقت الدوام دون تسجيل بصمة خروج يدوية).\n\nنرجو التأكد من تسجيل البصمة بشكل منتظم.`;
    }

    if (r.checkout === 'لا توجد بصمة انصراف') {
      return `السلام عليكم ورحمة الله وبركاته\n\nالأستاذ الموقر: ${r.name}\n\nنحيطكم علماً بعدم وجود بصمة انصراف مسجلة ليوم ${dName} الموافق ${dStr} في جهاز البصمة.\n\nيرجى مراجعة إدارة المدرسة لتأكيد سبب عدم البصم.`;
    }

    return `السلام عليكم ورحمة الله وبركاته\n\nالأستاذ الموقر: ${r.name}\n\nبخصوص سجل حضوركم ليوم ${dName} الموافق ${dStr}، تم تسجيل الحالة كـ (${r.status || 'عطلة'}).\n\nمع خالص تحيات الإدارة المدرسية.`;
  };

  const handleOpenMessageModal = (recordKey: string) => {
    const record = getRecordByKey(recordKey);
    if (!record) return;

    const text = getMessageText(record);
    setModalActiveRecordKey(recordKey);
    setModalTitle(`إرسال إشعار إلى الأستاذ / ${record.name}`);
    setModalText(text);
    setModalOpen(true);
  };

  const handleCopyModalText = () => {
    navigator.clipboard.writeText(modalText);
    alert('تم نسخ نص الرسالة بنجاح.');

    if (modalActiveRecordKey) {
      const record = getRecordByKey(modalActiveRecordKey);
      logMessage(record, 'نسخ نص الإشعار');
    }
  };

  const handleShareWhatsApp = () => {
    const encoded = encodeURIComponent(modalText);
    window.open(`https://wa.me/?text=${encoded}`, '_blank');

    if (modalActiveRecordKey) {
      const record = getRecordByKey(modalActiveRecordKey);
      logMessage(record, 'مشاركة عبر واتساب');
    }
  };

  const logMessage = (record: ParsedRow | undefined, resultAction: string) => {
    if (!record) return;
    const nowStr = new Date().toLocaleString('ar-SA');
    const newEntry: MessageLogEntry = {
      time: nowStr,
      name: record.name,
      date: arDate(record.date),
      type: record.type === 'absence' ? 'إشعار غياب' : (record.late > 0 ? 'تنبيه تأخر' : 'تنبيه انصراف'),
      result: resultAction,
    };
    setMessageLogs(prev => [newEntry, ...prev].slice(0, 200));
  };

  // 7. General statistics for current filtered list
  const getAggregatedActiveSummaries = () => {
    const activeSummaries = filterMonthlySummaries(monthlySummaries);
    const empMap = new Map<string, EmployeeSummary>();
    
    activeSummaries.forEach(s => {
      if (!empMap.has(s.name)) {
        empMap.set(s.name, {
          name: s.name,
          job: s.job,
          month: selectedMonth === 'all' ? 'جميع الأشهر' : s.month,
          work: 0,
          present: 0,
          absence: 0,
          excuse: 0,
          lateCount: 0,
          lateMins: 0,
          completeCount: 0,
          earlyCount: 0,
          earlyMins: 0,
          autoCount: 0,
          missingCheckoutCount: 0,
          wrongTimeExitsCount: 0,
          emergencyNoExitCount: 0,
          emergencyWithExitCount: 0,
          attendanceRate: 0,
          rate: 100,
          rating: '',
          disciplineCategory: 'ideal',
          disciplineReason: '',
        });
      }
      const agg = empMap.get(s.name)!;
      agg.work += s.work || 0;
      agg.present += s.present || 0;
      agg.absence += s.absence || 0;
      agg.excuse += s.excuse || 0;
      agg.lateCount += s.lateCount || 0;
      agg.lateMins += s.lateMins || 0;
      agg.completeCount += s.completeCount || 0;
      agg.earlyCount += s.earlyCount || 0;
      agg.earlyMins += s.earlyMins || 0;
      agg.autoCount += s.autoCount || 0;
      agg.missingCheckoutCount += s.missingCheckoutCount || 0;
      agg.wrongTimeExitsCount = (agg.wrongTimeExitsCount || 0) + (s.wrongTimeExitsCount || 0);
      agg.emergencyNoExitCount = (agg.emergencyNoExitCount || 0) + (s.emergencyNoExitCount || 0);
      agg.emergencyWithExitCount = (agg.emergencyWithExitCount || 0) + (s.emergencyWithExitCount || 0);
      if (s.job && !agg.job) {
        agg.job = s.job;
      }
    });

    return Array.from(empMap.values()).map(agg => {
      const attendanceRate = agg.work ? Math.round(((agg.present + agg.excuse) / agg.work) * 100) : 0;
      const classification = disciplineClassify(agg);
      return {
        ...agg,
        attendanceRate,
        rate: classification.score,
        rating: classification.disciplineLabel,
        disciplineCategory: classification.disciplineCategory,
        disciplineReason: classification.reason,
      };
    });
  };

  const getDisciplineStats = () => {
    const aggregatedSummaries = getAggregatedActiveSummaries();
    const totalCount = aggregatedSummaries.length;

    const excellent = aggregatedSummaries.filter(m => 
      (m.lateCount || 0) <= 2 && 
      (m.earlyCount || 0) <= 1 && 
      (m.autoCount || 0) === 0 && 
      (m.missingCheckoutCount || 0) === 0 && 
      (m.absence || 0) === 0 &&
      (m.emergencyNoExitCount || 0) === 0
    );

    const needsSupport = aggregatedSummaries.filter(m => 
      (m.lateCount || 0) >= 3 || 
      (m.earlyCount || 0) >= 5 || 
      ((m.autoCount || 0) + (m.missingCheckoutCount || 0) + (m.emergencyNoExitCount || 0)) >= 3
    );

    const ideal = aggregatedSummaries.filter(m => m.disciplineCategory === 'ideal');
    const notes = aggregatedSummaries.filter(m => m.disciplineCategory === 'notes');
    const follow = aggregatedSummaries.filter(m => m.disciplineCategory === 'follow');
    const admin = aggregatedSummaries.filter(m => m.disciplineCategory === 'admin');

    const totalAbsence = aggregatedSummaries.reduce((sum, m) => sum + (m.absence || 0), 0);
    const totalLateMins = aggregatedSummaries.reduce((sum, m) => sum + (m.lateMins || 0), 0);
    const totalLateCount = aggregatedSummaries.reduce((sum, m) => sum + (m.lateCount || 0), 0);
    const totalEarlyCount = aggregatedSummaries.reduce((sum, m) => sum + (m.earlyCount || 0), 0);
    const totalEarlyMins = aggregatedSummaries.reduce((sum, m) => sum + (m.earlyMins || 0), 0);
    const totalAutoCount = aggregatedSummaries.reduce((sum, m) => sum + (m.autoCount || 0), 0);
    const totalMissingCheckout = aggregatedSummaries.reduce((sum, m) => sum + (m.missingCheckoutCount || 0), 0);
    const totalEmergencyNoExit = aggregatedSummaries.reduce((sum, m) => sum + (m.emergencyNoExitCount || 0), 0);
    const totalEmergencyWithExit = aggregatedSummaries.reduce((sum, m) => sum + (m.emergencyWithExitCount || 0), 0);

    return {
      totalCount,
      excellent,
      needsSupport,
      ideal,
      notes,
      follow,
      admin,
      totalAbsence,
      totalLateMins,
      totalLateCount,
      totalEarlyCount,
      totalEarlyMins,
      totalAutoCount,
      totalMissingCheckout,
      totalEmergencyNoExit,
      totalEmergencyWithExit,
    };
  };

  const stats = getDisciplineStats();

  // Dynamic calculations for Leaves & Excuses (Annual vs Emergency)
  const dashboardLeaves = filterDailyRecords(dailyRecords).filter(r => r.type === 'excuse');
  const dashboardAnnualCount = dashboardLeaves.filter(r => {
    const s = r.status || '';
    return s.includes('سنوية') || s.includes('سنويه') || s.includes('اعتياد');
  }).length;
  const dashboardEmergencyCount = dashboardLeaves.filter(r => {
    const s = r.status || '';
    return s.includes('طارئ') || s.includes('اضطرار');
  }).length;

  const isExcludedDay = (d: any) => {
    const dk = dateKey(d);
    return isWeekend(d, config.weekend) || 
      isHoliday(d, config.holidays) || 
      (config.suspendedDates || []).includes(dk) ||
      autoDetectedSuspendedDays.includes(dk);
  };

  // Calculations for Absence Tab Statistics
  const allSchoolAbsences = dailyRecords.filter(r => {
    if (isExcludedDay(r.date)) {
      return false;
    }
    const matchType = r.type === 'absence';
    let matchPeriod = true;
    if (selectedPeriodType === 'range') {
      const dk = dateKey(r.date);
      const matchStart = !selectedStartDate || dk >= selectedStartDate;
      const matchEnd = !selectedEndDate || dk <= selectedEndDate;
      matchPeriod = matchStart && matchEnd;
    } else {
      matchPeriod = selectedMonth === 'all' || r.month === selectedMonth;
    }
    return matchType && matchPeriod;
  });
  const uniqueSchoolAbsentNames = Array.from(new Set(allSchoolAbsences.map(r => r.name)));
  
  const selectedEmployeeAbsences = dailyRecords.filter(r => {
    if (isExcludedDay(r.date)) {
      return false;
    }
    const matchType = r.type === 'absence';
    const matchEmployee = r.name === selectedEmployee;
    let matchPeriod = true;
    if (selectedPeriodType === 'range') {
      const dk = dateKey(r.date);
      const matchStart = !selectedStartDate || dk >= selectedStartDate;
      const matchEnd = !selectedEndDate || dk <= selectedEndDate;
      matchPeriod = matchStart && matchEnd;
    } else {
      matchPeriod = selectedMonth === 'all' || r.month === selectedMonth;
    }
    return matchType && matchEmployee && matchPeriod;
  });
  const currentEmpSummaryInAbsence = getAggregatedActiveSummaries().find(s => s.name === selectedEmployee);

  // 8. Individual aggregate generator
  const getIndividualReport = (
    name: string, 
    month: string, 
    customStartDate?: string, 
    customEndDate?: string
  ) => {
    let effStart = (customStartDate || '').trim();
    let effEnd = (customEndDate || '').trim();

    // Auto-detect range if encoded in month string (e.g. from filterMonthlySummaries: "من 2026-08-01 إلى 2026-08-31")
    if (!effStart && !effEnd && typeof month === 'string' && month.includes(' إلى ') && (month.startsWith('من ') || month.includes('الفترة من'))) {
      const cleaned = month.replace('الفترة من', '').replace('من', '').trim();
      const parts = cleaned.split(' إلى ');
      if (parts.length === 2) {
        const p1 = parts[0].trim();
        const p2 = parts[1].trim();
        if (p1 && p1 !== 'البداية') effStart = p1;
        if (p2 && p2 !== 'النهاية') effEnd = p2;
      }
    }

    const startKey = effStart ? dateKey(effStart) : '';
    const endKey = effEnd ? dateKey(effEnd) : '';
    const [minKey, maxKey] = (startKey && endKey && startKey > endKey) ? [endKey, startKey] : [startKey, endKey];

    const records = dailyRecords.filter(r => {
      // Exclude weekends, public holidays, and suspended dates
      if (isExcludedDay(r.date)) {
        return false;
      }
      const matchName = name === 'all' || r.name === name;
      let matchPeriod = true;
      if (minKey || maxKey) {
        const dk = r.dateKey || dateKey(r.date);
        const matchStart = !minKey || dk >= minKey;
        const matchEnd = !maxKey || dk <= maxKey;
        matchPeriod = matchStart && matchEnd;
      } else {
        matchPeriod = month === 'all' || r.month === month;
      }
      return matchName && matchPeriod;
    });

    const formatRangeDate = (dk: string) => {
      if (!dk) return '';
      const parts = dk.split('-');
      if (parts.length === 3) {
        return `${parts[2]}-${parts[1]}-${parts[0]}`;
      }
      return dk;
    };

    const summary: EmployeeSummary = {
      name: name === 'all' ? 'جميع الموظفين' : name,
      job: '',
      month: (minKey || maxKey)
        ? `الفترة من ${formatRangeDate(minKey) || 'البداية'} إلى ${formatRangeDate(maxKey) || 'النهاية'}`
        : (month === 'all' ? 'كل الأشهر' : month),
      work: records.length,
      present: 0,
      absence: 0,
      excuse: 0,
      lateCount: 0,
      lateMins: 0,
      completeCount: 0,
      earlyCount: 0,
      earlyMins: 0,
      autoCount: 0,
      missingCheckoutCount: 0,
      wrongTimeExitsCount: 0,
      emergencyNoExitCount: 0,
      emergencyWithExitCount: 0,
    };

    if (name !== 'all') {
      const emp = employees.find(e => e.name === name);
      if (emp) summary.job = emp.job;
    }

    records.forEach(r => {
      if (r.type === 'present') {
        summary.present++;
        if (r.late > 0) {
          summary.lateCount++;
          summary.lateMins += r.late;
        }
        if (r.checkout === 'مكتمل') {
          summary.completeCount++;
        }
        if (r.checkout === 'خروج قبل نهاية الدوام' || r.checkout === 'انصراف مبكر') {
          summary.earlyCount++;
          summary.earlyMins += r.earlyMins || 0;
        }
        if (r.checkout === 'انصراف تلقائي') {
          summary.autoCount++;
        }
        if (r.checkout === 'لا توجد بصمة انصراف') {
          summary.missingCheckoutCount++;
        }
        if (r.checkout === 'انصراف في غير وقت الفعلي') {
          summary.wrongTimeExitsCount = (summary.wrongTimeExitsCount || 0) + 1;
        }
        if (r.checkout === 'استئذان بدون بصمة انصراف') {
          summary.emergencyNoExitCount = (summary.emergencyNoExitCount || 0) + 1;
        }
        if (r.checkout === 'استئذان طارئ موثق (بصمة)') {
          summary.emergencyWithExitCount = (summary.emergencyWithExitCount || 0) + 1;
        }
      } else if (r.type === 'excuse') {
        summary.excuse++;
        if (r.checkout === 'استئذان بدون بصمة انصراف') {
          summary.emergencyNoExitCount = (summary.emergencyNoExitCount || 0) + 1;
        }
        if (r.checkout === 'استئذان طارئ موثق (بصمة)') {
          summary.emergencyWithExitCount = (summary.emergencyWithExitCount || 0) + 1;
        }
      } else if (r.type === 'absence') {
        summary.absence++;
      }
    });

    const attRate = summary.work ? Math.round(((summary.present + summary.excuse) / summary.work) * 100) : 0;
    const classification = disciplineClassify(summary);

    summary.attendanceRate = attRate;
    summary.rate = classification.score;
    summary.rating = classification.disciplineLabel;
    summary.disciplineCategory = classification.disciplineCategory;
    summary.disciplineReason = classification.reason;

    const sortedRecords = [...records].sort((a, b) => new Date(a.date).getTime() - new Date(b.date).getTime());

    return {
      summary,
      details: sortedRecords,
      presents: sortedRecords.filter(r => r.type === 'present'),
      absences: sortedRecords.filter(r => r.type === 'absence'),
      lates: sortedRecords.filter(r => r.late > 0),
      earlies: sortedRecords.filter(r => r.checkout === 'خروج قبل نهاية الدوام' || r.checkout === 'انصراف مبكر'),
      autos: sortedRecords.filter(r => r.checkout === 'انصراف تلقائي'),
      wrongTimeExits: sortedRecords.filter(r => r.checkout === 'انصراف في غير وقت الفعلي'),
      excusesList: sortedRecords.filter(r => r.type === 'excuse'),
    };
  };

  const renderDateCell = (date: Date) => {
    const mode = config.calendarMode || 'both';
    if (mode === 'hijri') {
      return (
        <div className="text-center">
          <span className="font-bold text-slate-800 text-xs">{toHijri(date)}</span>
        </div>
      );
    }
    if (mode === 'gregorian') {
      return (
        <div className="text-center">
          <span className="font-bold text-slate-800 text-xs">{arDate(date)}م</span>
        </div>
      );
    }
    return (
      <div className="text-center leading-tight">
        <span className="font-bold text-slate-800 block text-xs">{arDate(date)}م</span>
        <span className="text-[10px] text-teal-700 font-bold block mt-0.5">{toHijri(date)}</span>
      </div>
    );
  };

  const formatPrintDate = (d: any) => {
    const mode = config.calendarMode || 'both';
    if (mode === 'hijri') return toHijri(d);
    if (mode === 'gregorian') return `${arDate(d)}م`;
    return `${arDate(d)}م<br><span style="font-size:10px; color:#0f766e; font-weight:bold;">${toHijri(d)}</span>`;
  };

  const formatWordDate = (d: any) => {
    const mode = config.calendarMode || 'both';
    if (mode === 'hijri') return toHijri(d);
    if (mode === 'gregorian') return `${arDate(d)}م`;
    return `${arDate(d)}م (${toHijri(d)})`;
  };

  const getSmartExecutiveSummaryText = () => {
    const activeSummaries = filterMonthlySummaries(monthlySummaries);
    const total = activeSummaries.length;
    if (total === 0) return 'لا توجد بيانات كافية لإجراء التحليل والملخص التنفيذي.';

    const excellentCount = activeSummaries.filter(m => 
      (m.lateCount || 0) <= 2 && 
      (m.earlyCount || 0) <= 1 && 
      (m.autoCount || 0) === 0 && 
      (m.missingCheckoutCount || 0) === 0 && 
      (m.absence || 0) === 0
    ).length;

    const needsSupportCount = activeSummaries.filter(m => 
      (m.lateCount || 0) >= 3 || 
      (m.earlyCount || 0) >= 5 || 
      ((m.autoCount || 0) + (m.missingCheckoutCount || 0)) >= 3
    ).length;

    return `بناءً على سجلات الحضور والانتظام النشطة للهيئة التعليمية والإدارية خلال هذه الفترة، تميز عدد ${excellentCount} من المنسوبين بالانضباط والانتظام العالي (فئة المنضبطون المتميزون) مع تسجيل تأخر لا يتجاوز يومين، وانصراف مبكر لا يتجاوز مرة واحدة، ودون تسجيل أي انصراف تلقائي أو غياب بدون عذر. بينما يبلغ عدد الزملاء الذين تظهر مؤشراتهم حاجة للمتابعة والدعم لتكرار التأخر، أو الانصراف المبكر، أو الانصراف التلقائي ${needsSupportCount} من الزملاء؛ وذلك بهدف تقديم الدعم والمساندة لهم لرفع مستوى الانتظام بالمدرسة.`;
  };

  const handlePrintReport = () => {
    if (monthlyEmployee === 'all') {
      const allReports = employees.map(e => {
        return monthlyPeriodType === 'range'
          ? getIndividualReport(e.name, 'all', monthlyStartDate, monthlyEndDate)
          : getIndividualReport(e.name, monthlyMonth);
      });

      const totalEmployees = allReports.length;
      const totalWorkDays = allReports.reduce((acc, r) => acc + r.summary.work, 0);
      const totalPresent = allReports.reduce((acc, r) => acc + r.summary.present, 0);
      const totalAbsence = allReports.reduce((acc, r) => acc + r.summary.absence, 0);
      const totalExcuse = allReports.reduce((acc, r) => acc + r.summary.excuse, 0);
      const totalLateCount = allReports.reduce((acc, r) => acc + r.summary.lateCount, 0);
      const totalLateMins = allReports.reduce((acc, r) => acc + r.summary.lateMins, 0);
      const totalEarlyCount = allReports.reduce((acc, r) => acc + r.summary.earlyCount, 0);
      const totalAutoCount = allReports.reduce((acc, r) => acc + r.summary.autoCount, 0);

      const avgAttendanceRate = totalEmployees > 0 
        ? Math.round(allReports.reduce((acc, r) => acc + (r.summary.attendanceRate || 0), 0) / totalEmployees) 
        : 0;

      const avgDisciplineRate = totalEmployees > 0 
        ? Math.round(allReports.reduce((acc, r) => acc + (r.summary.rate || 0), 0) / totalEmployees) 
        : 0;

      const periodLabel = monthlyPeriodType === 'range'
        ? `الفترة من ${monthlyStartDate || 'البداية'} إلى ${monthlyEndDate || 'النهاية'}`
        : (monthlyMonth === 'all' ? 'جميع الأشهر المسجلة' : monthlyMonth);

      const printWindow = window.open('', '_blank', 'width=1100,height=850');
      if (!printWindow) {
        alert('الرجاء السماح بالنوافذ المنبثقة للطباعة.');
        return;
      }

      const printHtml = `
        <!doctype html>
        <html lang="ar" dir="rtl">
        <head>
          <meta charset="utf-8">
          <title>التقرير الشامل لحضور وانضباط منسوبي المدرسة</title>
          <style>
            @import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@400;500;700;800&display=swap');
            body {
              font-family: 'Tajawal', sans-serif;
              color: #1e293b;
              margin: 30px;
              direction: rtl;
              text-align: right;
              line-height: 1.5;
            }
            .report-header {
              display: flex;
              justify-content: space-between;
              align-items: center;
              border-bottom: 3px solid #0f766e;
              padding-bottom: 15px;
              margin-bottom: 25px;
            }
            .school-info {
              font-size: 13px;
              font-weight: bold;
            }
            .report-title {
              text-align: center;
              font-size: 20px;
              color: #0f766e;
              font-weight: 800;
            }
            .stats-grid {
              display: grid;
              grid-template-columns: repeat(4, 1fr);
              gap: 10px;
              margin-bottom: 20px;
            }
            .stat-card {
              background-color: #f8fafc;
              border: 1px solid #cbd5e1;
              border-radius: 8px;
              padding: 10px;
              text-align: center;
            }
            .stat-label {
              font-size: 10px;
              color: #64748b;
              font-weight: bold;
            }
            .stat-val {
              font-size: 18px;
              font-weight: 800;
              color: #0f766e;
              margin-top: 2px;
            }
            table {
              width: 100%;
              border-collapse: collapse;
              margin-top: 10px;
              margin-bottom: 25px;
              font-size: 11px;
            }
            th, td {
              border: 1px solid #cbd5e1;
              padding: 8px;
              text-align: right;
            }
            th {
              background-color: #0f766e;
              color: white;
              font-weight: bold;
            }
            tr:nth-child(even) {
              background-color: #f8fafc;
            }
            .badge {
              display: inline-block;
              padding: 2px 6px;
              border-radius: 999px;
              font-size: 9px;
              font-weight: bold;
            }
            .badge-green { background-color: #dcfce7; color: #166534; }
            .badge-amber { background-color: #fef3c7; color: #92400e; }
            .badge-red { background-color: #fee2e2; color: #991b1b; }
            .badge-blue { background-color: #dbeafe; color: #1e40af; }
            .signature-box {
              margin-top: 40px;
              display: flex;
              justify-content: space-between;
              font-size: 12px;
            }
            .sig-col {
              text-align: center;
              width: 220px;
            }
            @media print {
              body { margin: 15px; }
              .no-print { display: none; }
            }
          </style>
        </head>
        <body>
          <div class="report-header">
            <div class="school-info">
              وزارة التعليم<br>
              إدارة التعليم بالأحساء<br>
              مدرسة الجشة المتوسطة
            </div>
            <div class="report-title">التقرير الشامل لحضور وانضباط منسوبي المدرسة</div>
            <div class="school-info" style="text-align: left;">
              التاريخ: ${new Date().toLocaleDateString('ar-SA')}<br>
              الفترة: ${periodLabel}
            </div>
          </div>

          <div class="stats-grid">
            <div class="stat-card">
              <div class="stat-label">إجمالي المنسوبين</div>
              <div class="stat-val">${totalEmployees} موظف</div>
            </div>
            <div class="stat-card">
              <div class="stat-label">متوسط نسبة الحضور والانتظام</div>
              <div class="stat-val" style="color:#16a34a">${avgAttendanceRate}%</div>
            </div>
            <div class="stat-card">
              <div class="stat-label">متوسط درجة الانضباط المدرسي</div>
              <div class="stat-val" style="color:#0f766e">${avgDisciplineRate}%</div>
            </div>
            <div class="stat-card">
              <div class="stat-label">إجمالي الغيابات والتأخيرات</div>
              <div class="stat-val" style="color:#dc2626">${totalAbsence} غياب | ${totalLateCount} تأخر</div>
            </div>
          </div>

          <h3>جدول رصد انضباط المعلمين والموظفين</h3>
          <table>
            <thead>
              <tr>
                <th>المعلم / الموظف</th>
                <th>المسمى الوظيفي</th>
                <th style="text-align:center">أيام العمل</th>
                <th style="text-align:center">حاضر</th>
                <th style="text-align:center">غياب (عذر)</th>
                <th style="text-align:center">غياب (بدون)</th>
                <th style="text-align:center">مرات التأخر</th>
                <th style="text-align:center">دقائق التأخر</th>
                <th style="text-align:center">انصراف مبكر / تلقائي</th>
                <th style="text-align:center">الدرجة</th>
                <th style="text-align:center">التقييم</th>
              </tr>
            </thead>
            <tbody>
              ${allReports.map(r => `
                <tr>
                  <td><b>${r.summary.name}</b></td>
                  <td>${r.summary.job || 'غير محدد'}</td>
                  <td style="text-align:center">${r.summary.work}</td>
                  <td style="text-align:center" style="color:#166534">${r.summary.present}</td>
                  <td style="text-align:center" style="color:#1e40af">${r.summary.excuse}</td>
                  <td style="text-align:center" style="color:#991b1b">${r.summary.absence}</td>
                  <td style="text-align:center">${r.summary.lateCount}</td>
                  <td style="text-align:center">${r.summary.lateMins}</td>
                  <td style="text-align:center">${r.summary.earlyCount + r.summary.autoCount}</td>
                  <td style="text-align:center"><b>${r.summary.rate}%</b></td>
                  <td style="text-align:center">
                    <span class="badge ${
                      (r.summary.rate || 0) >= 95 ? 'badge-green' :
                      (r.summary.rate || 0) >= 85 ? 'badge-blue' :
                      (r.summary.rate || 0) >= 75 ? 'badge-amber' : 'badge-red'
                    }">${r.summary.rating || 'غير محدد'}</span>
                  </td>
                </tr>
              `).join('')}
            </tbody>
          </table>

          <div class="signature-box">
            <div class="sig-col">
              <b>المعلم / الموظف: ............................................</b><br><br><br>
              التوقيع: ............................................
            </div>
            <div class="sig-col">
              <b>مدير المدرسة: ............................................</b><br><br><br>
              التوقيع: ............................................
            </div>
          </div>

          <script>
            window.onload = function() {
              window.focus();
              window.print();
            }
          <\/script>
        </body>
        </html>
      `;

      printWindow.document.open();
      printWindow.document.write(printHtml);
      printWindow.document.close();
      return;
    }

    const rep = monthlyPeriodType === 'range'
      ? getIndividualReport(monthlyEmployee, 'all', monthlyStartDate, monthlyEndDate)
      : getIndividualReport(monthlyEmployee, monthlyMonth);
    const m = rep.summary;

    const printWindow = window.open('', '_blank', 'width=1000,height=800');
    if (!printWindow) {
      alert('الرجاء السماح بالنوافذ المنبثقة للطباعة.');
      return;
    }

    const printHtml = `
      <!doctype html>
      <html lang="ar" dir="rtl">
      <head>
        <meta charset="utf-8">
        <title>تقرير حضور وانضباط - ${m.name}</title>
        <style>
          @import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@400;500;700;800&display=swap');
          body {
            font-family: 'Tajawal', sans-serif;
            color: #1e293b;
            margin: 40px;
            direction: rti;
            text-align: right;
            line-height: 1.6;
          }
          .report-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-bottom: 3px solid #0f766e;
            padding-bottom: 20px;
            margin-bottom: 30px;
          }
          .school-info {
            font-size: 14px;
            font-weight: bold;
          }
          .report-title {
            text-align: center;
            font-size: 24px;
            color: #0f766e;
            font-weight: 800;
          }
          .profile-box {
            background-color: #f8fafc;
            border: 1px solid #e2e8f0;
            border-radius: 12px;
            padding: 20px;
            margin-bottom: 24px;
          }
          .profile-title {
            font-size: 20px;
            font-weight: 800;
            color: #0f766e;
            margin: 0 0 8px 0;
          }
          .profile-meta {
            font-size: 14px;
            color: #64748b;
          }
          .stats-grid {
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 12px;
            margin-bottom: 24px;
          }
          .stat-card {
            background-color: #fff;
            border: 1px solid #cbd5e1;
            border-radius: 10px;
            padding: 12px;
            text-align: center;
          }
          .stat-label {
            font-size: 11px;
            color: #64748b;
            font-weight: bold;
          }
          .stat-val {
            font-size: 22px;
            font-weight: 800;
            color: #0f766e;
            margin-top: 4px;
          }
          table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 16px;
            margin-bottom: 30px;
            font-size: 12px;
          }
          th, td {
            border: 1px solid #cbd5e1;
            padding: 10px;
            text-align: right;
          }
          th {
            background-color: #0f766e;
            color: white;
            font-weight: bold;
          }
          tr:nth-child(even) {
            background-color: #f8fafc;
          }
          .badge {
            display: inline-block;
            padding: 3px 8px;
            border-radius: 999px;
            font-size: 10px;
            font-weight: bold;
          }
          .badge-green { background-color: #dcfce7; color: #166534; }
          .badge-amber { background-color: #fef3c7; color: #92400e; }
          .badge-red { background-color: #fee2e2; color: #991b1b; }
          .badge-blue { background-color: #dbeafe; color: #1e40af; }
          .badge-gray { background-color: #f1f5f9; color: #334155; }
          .signature-box {
            margin-top: 50px;
            display: flex;
            justify-content: space-between;
          }
          .sig-col {
            text-align: center;
            width: 250px;
          }
          @media print {
            body { margin: 20px; }
            .no-print { display: none; }
          }
        </style>
      </head>
      <body>
        <div class="report-header">
          <div class="school-info">
            وزارة التعليم<br>
            إدارة التعليم بالأحساء<br>
            مدرسة الجشة المتوسطة
          </div>
          <div class="report-title">تقرير الانضباط والحضور اليومي</div>
          <div class="school-info" style="text-align: left;">
            التاريخ: ${new Date().toLocaleDateString('ar-SA')}<br>
          </div>
        </div>

        <div class="profile-box">
          <h2 class="profile-title">${m.name}</h2>
          <div class="profile-meta">المسمى الوظيفي: ${m.job || 'غير محدد'} | السجل المدني: ${m.civil || 'غير محدد'}</div>
          <div class="profile-meta" style="margin-top:4px;">الفترة المشمولة بالتقرير: ${m.month === 'all' ? 'جميع الأشهر' : m.month}</div>
        </div>

        <div class="stats-grid">
          <div class="stat-card">
            <div class="stat-label">أيام الدوام المشمولة</div>
            <div class="stat-val">${m.work}</div>
          </div>
          <div class="stat-card">
            <div class="stat-label">أيام الحضور الفعلي</div>
            <div class="stat-val">${m.present}</div>
          </div>
          <div class="stat-card">
            <div class="stat-label">أيام الغياب غير المبرر</div>
            <div class="stat-val" style="color:#dc2626">${m.absence}</div>
          </div>
          <div class="stat-card">
            <div class="stat-label">أيام الغياب المبرر/الأعذار</div>
            <div class="stat-val" style="color:#2563eb">${m.excuse}</div>
          </div>
          <div class="stat-card">
            <div class="stat-label">مرات التأخر الصباحي</div>
            <div class="stat-val" style="color:#d97706">${m.lateCount}</div>
          </div>
          <div class="stat-card">
            <div class="stat-label">إجمالي دقائق التأخر</div>
            <div class="stat-val" style="color:#d97706">${m.lateMins}</div>
          </div>
          <div class="stat-card">
            <div class="stat-label">انصراف مبكر / تلقائي</div>
            <div class="stat-val" style="color:#ea580c">${m.earlyCount + m.autoCount}</div>
          </div>
          <div class="stat-card">
            <div class="stat-label">درجة الانضباط النهائية</div>
            <div class="stat-val" style="color:#16a34a">${m.rate}%</div>
          </div>
        </div>

        <div style="background-color: #fff7ed; border: 1px solid #fed7aa; border-radius: 8px; padding: 12px; margin-bottom: 20px; font-size:13px; color:#9a3412">
          <strong>سبب تقييم الانضباط:</strong> ${m.disciplineReason}
        </div>

        <h3>جدول الحضور والتأخر التفصيلي</h3>
        <table>
          <thead>
            <tr>
              <th>التاريخ</th>
              <th>اليوم</th>
              <th>حالة الدوام</th>
              <th>وقت الدخول</th>
              <th>وقت الانصراف</th>
              <th>حالة الانصراف</th>
              <th>التأخر (دقيقة)</th>
              <th>الخروج المبكر</th>
            </tr>
          </thead>
          <tbody>
            ${rep.details.map(r => `
              <tr>
                <td>${formatPrintDate(r.date)}</td>
                <td>${dayName(r.date)}</td>
                <td>${r.type === 'present' ? 'حاضر' : (r.type === 'excuse' ? 'غياب بعذر' : 'غياب بدون عذر')}</td>
                <td>${r.tin?.text || '-'}</td>
                <td>${r.tout?.text || '-'}</td>
                <td>${r.checkout || '-'}</td>
                <td>${r.late || 0}</td>
                <td>${r.earlyMins || 0}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>

        <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-top: 45px; font-weight:bold; font-size:13px;">
          <div style="text-align: right; min-width: 260px;">
            <div>المعلم / الموظف: ${m.name}</div>
            <div style="margin-top: 18px;">التوقيع: ............................................</div>
          </div>
          <div style="text-align: left; min-width: 260px;">
            <div>مدير المدرسة: ............................................</div>
            <div style="margin-top: 18px;">التوقيع: ............................................</div>
          </div>
        </div>

        <script>
          window.onload = function() {
            window.focus();
            window.print();
          }
        <\/script>
      </body>
      </html>
    `;

    printWindow.document.open();
    printWindow.document.write(printHtml);
    printWindow.document.close();
  };

  const handleExportWord = () => {
    if (monthlyEmployee === 'all') {
      const allReports = employees.map(e => {
        return monthlyPeriodType === 'range'
          ? getIndividualReport(e.name, 'all', monthlyStartDate, monthlyEndDate)
          : getIndividualReport(e.name, monthlyMonth);
      });

      const totalEmployees = allReports.length;
      const totalWorkDays = allReports.reduce((acc, r) => acc + r.summary.work, 0);
      const totalPresent = allReports.reduce((acc, r) => acc + r.summary.present, 0);
      const totalAbsence = allReports.reduce((acc, r) => acc + r.summary.absence, 0);
      const totalExcuse = allReports.reduce((acc, r) => acc + r.summary.excuse, 0);
      const totalLateCount = allReports.reduce((acc, r) => acc + r.summary.lateCount, 0);
      const totalLateMins = allReports.reduce((acc, r) => acc + r.summary.lateMins, 0);
      const totalEarlyCount = allReports.reduce((acc, r) => acc + r.summary.earlyCount, 0);
      const totalAutoCount = allReports.reduce((acc, r) => acc + r.summary.autoCount, 0);

      const avgAttendanceRate = totalEmployees > 0 
        ? Math.round(allReports.reduce((acc, r) => acc + (r.summary.attendanceRate || 0), 0) / totalEmployees) 
        : 0;

      const avgDisciplineRate = totalEmployees > 0 
        ? Math.round(allReports.reduce((acc, r) => acc + (r.summary.rate || 0), 0) / totalEmployees) 
        : 0;

      const periodLabel = monthlyPeriodType === 'range'
        ? `الفترة من ${monthlyStartDate || 'البداية'} إلى ${monthlyEndDate || 'النهاية'}`
        : (monthlyMonth === 'all' ? 'جميع الأشهر المسجلة' : monthlyMonth);

      const wordHtml = `
        <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
        <head>
          <title>التقرير الشامل لحضور وانضباط منسوبي المدرسة</title>
          <style>
            body { font-family: 'Arial', sans-serif; direction: rtl; text-align: right; }
            .header { border-bottom: 2px solid #0f766e; padding-bottom: 10px; margin-bottom: 20px; }
            .table { width: 100%; border-collapse: collapse; margin-top: 15px; }
            .table th, .table td { border: 1px solid #999; padding: 8px; text-align: right; }
            .table th { background-color: #0f766e; color: #ffffff; }
          </style>
        </head>
        <body>
          <div class="header">
            <h2>مدرسة الجشة المتوسطة - إدارة التعليم بالأحساء</h2>
            <h3>التقرير الشامل لحضور وانضباط منسوبي المدرسة</h3>
            <p>الفترة المشمولة بالتقرير: <b>${periodLabel}</b></p>
            <p>تاريخ التصدير: ${new Date().toLocaleDateString('ar-SA')}م (${toHijri(new Date())})</p>
          </div>

          <h3>المؤشرات العامة للجميع:</h3>
          <ul>
            <li>إجمالي عدد المعلمين والموظفين: ${totalEmployees}</li>
            <li>متوسط نسبة الحضور والانتظام: ${avgAttendanceRate}%</li>
            <li>متوسط درجة الانضباط المدرسي: ${avgDisciplineRate}%</li>
            <li>إجمالي الغيابات بدون عذر: ${totalAbsence} يوم</li>
            <li>إجمالي حالات التأخر الصباحي: ${totalLateCount} (بمجموع ${totalLateMins} دقيقة)</li>
            <li>إجمالي حالات الخروج المبكر / الانصراف التلقائي: ${totalEarlyCount + totalAutoCount}</li>
          </ul>

          <h3>جدول رصد انضباط المعلمين والموظفين الشامل:</h3>
          <table class="table">
            <thead>
              <tr>
                <th>المعلم / الموظف</th>
                <th>المسمى الوظيفي</th>
                <th>أيام العمل</th>
                <th>حضور</th>
                <th>غياب بعذر</th>
                <th>غياب بدون عذر</th>
                <th>مرات التأخر</th>
                <th>دقائق التأخر</th>
                <th>خروج مبكر / تلقائي</th>
                <th>درجة الانضباط</th>
                <th>التصنيف والمؤشر</th>
              </tr>
            </thead>
            <tbody>
              ${allReports.map(r => `
                <tr>
                  <td><b>${r.summary.name}</b></td>
                  <td>${r.summary.job || 'غير محدد'}</td>
                  <td>${r.summary.work}</td>
                  <td>${r.summary.present}</td>
                  <td>${r.summary.excuse}</td>
                  <td>${r.summary.absence}</td>
                  <td>${r.summary.lateCount}</td>
                  <td>${r.summary.lateMins}</td>
                  <td>${r.summary.earlyCount + r.summary.autoCount}</td>
                  <td><b>${r.summary.rate}%</b></td>
                  <td>${r.summary.rating || 'غير محدد'}</td>
                </tr>
              `).join('')}
            </tbody>
          </table>

          <div style="margin-top: 40px; font-size: 13px; font-weight: bold;">
            <table style="width: 100%; border: none;">
              <tr>
                <td style="border: none; text-align: right; width: 50%;">
                  المعلم / الموظف: ............................................<br><br>
                  التوقيع: ............................................
                </td>
                <td style="border: none; text-align: left; width: 50%;">
                  مدير المدرسة: ............................................<br><br>
                  التوقيع: ............................................
                </td>
              </tr>
            </table>
          </div>
        </body>
        </html>
      `;

      const blob = new Blob(['\ufeff' + wordHtml], { type: 'application/msword' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `التقرير_الشامل_للانضباط_والحضور_${periodLabel.replace(/\s+/g, '_')}.doc`;
      a.click();
      URL.revokeObjectURL(url);
      return;
    }

    const rep = monthlyPeriodType === 'range'
      ? getIndividualReport(monthlyEmployee, 'all', monthlyStartDate, monthlyEndDate)
      : getIndividualReport(monthlyEmployee, monthlyMonth);
    const m = rep.summary;

    const wordHtml = `
      <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
      <head>
        <title>تقرير حضور وانضباط</title>
        <style>
          body { font-family: 'Arial', sans-serif; direction: rtl; text-align: right; }
          .header { border-bottom: 2px solid #0f766e; padding-bottom: 10px; margin-bottom: 20px; }
          .table { width: 100%; border-collapse: collapse; margin-top: 15px; }
          .table th, .table td { border: 1px solid #999; padding: 8px; text-align: right; }
          .table th { background-color: #0f766e; color: #ffffff; }
        </style>
      </head>
      <body>
        <div class="header">
          <h2>مدرسة الجشة المتوسطة - إدارة التعليم بالأحساء</h2>
          <h3>تقرير الانضباط والحضور الشامل</h3>
          <p>الموظف: <b>${m.name}</b></p>
          <p>المسمى الوظيفي: ${m.job || '-'} | السجل المدني: ${m.civil || '-'}</p>
          <p>الفترة: ${m.month} | تاريخ التصدير: ${new Date().toLocaleDateString('ar-SA')}م (${toHijri(new Date())})</p>
        </div>

        <h3>الملخص الرقمي:</h3>
        <ul>
          <li>أيام الدوام المقررة: ${m.work}</li>
          <li>أيام الحضور الفعلي: ${m.present}</li>
          <li>أيام الغياب غير المبرر: ${m.absence}</li>
          <li>أيام الغياب بعذر: ${m.excuse}</li>
          <li>مرات التأخر الصباحي: ${m.lateCount} (بمجموع ${m.lateMins} دقيقة)</li>
          <li>حالات الانصراف غير المكتمل أو المبكر: ${m.earlyCount + m.autoCount}</li>
          <li>درجة الانضباط العام: ${m.rate}%</li>
          <li>المؤشر العام والتصنيف: <b>${m.rating}</b></li>
        </ul>

        <p><b>مبررات التقييم:</b> ${m.disciplineReason}</p>

        <h3>سجل الدوام اليومي التفصيلي:</h3>
        <table class="table">
          <thead>
            <tr>
              <th>التاريخ</th>
              <th>اليوم</th>
              <th>حالة الحضور</th>
              <th>وقت الدخول</th>
              <th>وقت الانصراف</th>
              <th>حالة الانصراف</th>
              <th>التأخر بالدقائق</th>
              <th>الخروج المبكر</th>
            </tr>
          </thead>
          <tbody>
            ${rep.details.map(r => `
              <tr>
                <td>${formatWordDate(r.date)}</td>
                <td>${dayName(r.date)}</td>
                <td>${r.type === 'present' ? 'حاضر' : (r.type === 'excuse' ? 'غياب بعذر' : 'غياب بدون عذر')}</td>
                <td>${r.tin?.text || '-'}</td>
                <td>${r.tout?.text || '-'}</td>
                <td>${r.checkout || '-'}</td>
                <td>${r.late || 0}</td>
                <td>${r.earlyMins || 0}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>

        <div style="margin-top: 40px; font-size: 13px; font-weight: bold;">
          <table style="width: 100%; border: none;">
            <tr>
              <td style="border: none; text-align: right; width: 50%;">
                المعلم / الموظف: ${m.name}<br><br>
                التوقيع: ............................................
              </td>
              <td style="border: none; text-align: left; width: 50%;">
                مدير المدرسة: ............................................<br><br>
                التوقيع: ............................................
              </td>
            </tr>
          </table>
        </div>
      </body>
      </html>
    `;

    const blob = new Blob(['\ufeff' + wordHtml], { type: 'application/msword' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `تقرير_انضباط_${m.name.replace(/\s+/g, '_')}.doc`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const handlePrintAttendanceList = () => {
    const list = filterDailyRecords(dailyRecords)
      .filter(r => r.type === 'present')
      .sort((a, b) => new Date(a.date).getTime() - new Date(b.date).getTime());
    const printWindow = window.open('', '_blank', 'width=1000,height=800');
    if (!printWindow) {
      alert('الرجاء السماح بالنوافذ المنبثقة للطباعة.');
      return;
    }

    const printHtml = `
      <!doctype html>
      <html lang="ar" dir="rtl">
      <head>
        <meta charset="utf-8">
        <title>سجل حضور منسوبي مدرسة الجشة المتوسطة</title>
        <style>
          @import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@400;500;700;800&display=swap');
          body {
            font-family: 'Tajawal', sans-serif;
            color: #1e293b;
            margin: 40px;
            direction: rtl;
            text-align: right;
            line-height: 1.6;
          }
          .report-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-bottom: 3px solid #0f766e;
            padding-bottom: 20px;
            margin-bottom: 30px;
          }
          .school-info {
            font-size: 14px;
            font-weight: bold;
          }
          .report-title {
            text-align: center;
            font-size: 24px;
            color: #0f766e;
            font-weight: 800;
          }
          table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 30px;
          }
          th, td {
            border: 1px solid #cbd5e1;
            padding: 10px;
            text-align: right;
            font-size: 12px;
          }
          th {
            background-color: #f1f5f9;
            color: #0f766e;
            font-weight: bold;
          }
        </style>
      </head>
      <body>
        <div class="report-header">
          <div class="school-info">
            وزارة التعليم<br>
            إدارة التعليم بالأحساء<br>
            مدرسة الجشة المتوسطة
          </div>
          <div class="report-title">سجل الحضور والمنسوبين</div>
          <div class="school-info" style="text-align: left;">
            التاريخ: ${new Date().toLocaleDateString('ar-SA')}م (${toHijri(new Date())})<br>
          </div>
        </div>

        <div style="background-color: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 12px; margin-bottom: 20px; font-size:13px; font-weight: bold; color: #0f766e;">
          ${selectedPeriodType === 'range'
            ? `الفترة المستهدفة: من ${selectedStartDate || 'البداية'} إلى ${selectedEndDate || 'النهاية'}`
            : (selectedMonth === 'all' 
              ? 'الفترة المستهدفة: جميع الأشهر المسجلة' 
              : `الشهر المستهدف: ${selectedMonth}`)}
        </div>

        <table>
          <thead>
            <tr>
              <th>اسم المعلم / الموظف</th>
              <th>المسمى الوظيفي</th>
              <th style="text-align:center;">التاريخ</th>
              <th style="text-align:center;">اليوم</th>
              <th style="text-align:center;">وقت الدخول</th>
              <th style="text-align:center;">وقت الانصراف</th>
              <th style="text-align:center;">حالة الانصراف</th>
            </tr>
          </thead>
          <tbody>
            ${list.map(r => `
              <tr>
                <td><b>${r.name}</b></td>
                <td>${r.job || '-'}</td>
                <td style="text-align:center;">${formatPrintDate(r.date)}</td>
                <td style="text-align:center;">${dayName(r.date)}</td>
                <td style="color:#047857; font-weight: bold; text-align:center;">${r.tin?.text || '-'}</td>
                <td style="color:#475569; font-weight: bold; text-align:center;">${r.tout?.text || '-'}</td>
                <td style="text-align:center;">${r.checkout || '-'}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>

        <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-top: 45px; font-weight:bold; font-size:13px;">
          <div style="text-align: right; min-width: 260px;">
            <div>المعلم / الموظف: ${selectedEmployee !== 'all' ? selectedEmployee : (Array.from(new Set(list.map(r => r.name))).length === 1 ? Array.from(new Set(list.map(r => r.name)))[0] : '...........................................')}</div>
            <div style="margin-top: 18px;">التوقيع: ............................................</div>
          </div>
          <div style="text-align: left; min-width: 260px;">
            <div>مدير المدرسة: ............................................</div>
            <div style="margin-top: 18px;">التوقيع: ............................................</div>
          </div>
        </div>

        <script>
          window.onload = function() {
            window.focus();
            window.print();
          }
        <\/script>
      </body>
      </html>
    `;

    printWindow.document.open();
    printWindow.document.write(printHtml);
    printWindow.document.close();
  };

  const handleExportAttendanceListWord = () => {
    const list = filterDailyRecords(dailyRecords)
      .filter(r => r.type === 'present')
      .sort((a, b) => new Date(a.date).getTime() - new Date(b.date).getTime());
    const targetTeacher = selectedEmployee !== 'all' 
      ? selectedEmployee 
      : (Array.from(new Set(list.map(r => r.name))).length === 1 
          ? Array.from(new Set(list.map(r => r.name)))[0] 
          : '');
    const wordHtml = `
      <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
      <head>
        <title>سجل الحضور والمنسوبين</title>
        <style>
          body { font-family: 'Arial', sans-serif; direction: rtl; text-align: right; }
          .header { border-bottom: 2px solid #0f766e; padding-bottom: 10px; margin-bottom: 20px; }
          .table { width: 100%; border-collapse: collapse; margin-top: 15px; }
          .table th, .table td { border: 1px solid #999; padding: 8px; text-align: right; }
          .table th { background-color: #0f766e; color: #ffffff; }
        </style>
      </head>
      <body>
        <div class="header">
          <h2>مدرسة الجشة المتوسطة - إدارة التعليم بالأحساء</h2>
          <h3>سجل الحضور والمنسوبين</h3>
          <p>
            ${selectedPeriodType === 'range'
              ? `الفترة المستهدفة: من ${selectedStartDate || 'البداية'} إلى ${selectedEndDate || 'النهاية'}`
              : (selectedMonth === 'all' 
                ? 'الفترة المستهدفة: جميع الأشهر المسجلة' 
                : `الشهر المستهدف: ${selectedMonth}`)}
            | تاريخ التصدير: ${new Date().toLocaleDateString('ar-SA')}م (${toHijri(new Date())})
          </p>
        </div>

        <table class="table">
          <thead>
            <tr>
              <th>اسم المعلم / الموظف</th>
              <th>المسمى الوظيفي</th>
              <th style="text-align:center;">التاريخ</th>
              <th style="text-align:center;">اليوم</th>
              <th style="text-align:center;">وقت الدخول</th>
              <th style="text-align:center;">وقت الانصراف</th>
              <th style="text-align:center;">حالة الانصراف</th>
            </tr>
          </thead>
          <tbody>
            ${list.map(r => `
              <tr>
                <td><b>${r.name}</b></td>
                <td>${r.job || '-'}</td>
                <td style="text-align:center;">${formatWordDate(r.date)}</td>
                <td style="text-align:center;">${dayName(r.date)}</td>
                <td style="text-align:center;">${r.tin?.text || '-'}</td>
                <td style="text-align:center;">${r.tout?.text || '-'}</td>
                <td style="text-align:center;">${r.checkout || '-'}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>

        <div style="margin-top: 40px; font-size: 13px; font-weight: bold;">
          <table style="width: 100%; border: none;">
            <tr>
              <td style="border: none; text-align: right; width: 50%;">
                المعلم / الموظف: ${targetTeacher || '...........................................'}<br><br>
                التوقيع: ............................................
              </td>
              <td style="border: none; text-align: left; width: 50%;">
                مدير المدرسة: ............................................<br><br>
                التوقيع: ............................................
              </td>
            </tr>
          </table>
        </div>
      </body>
      </html>
    `;

    const blob = new Blob(['\ufeff' + wordHtml], { type: 'application/msword' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `سجل_الحضور_والمنسوبين_${(selectedMonth === 'all' ? 'كامل' : selectedMonth).replace(/\s+/g, '_')}.doc`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const handlePrintLateList = () => {
    const list = filterDailyRecords(dailyRecords)
      .filter(r => r.late > 0)
      .sort((a, b) => new Date(a.date).getTime() - new Date(b.date).getTime());
    const totalLateMins = list.reduce((acc, r) => acc + r.late, 0);
    const printWindow = window.open('', '_blank', 'width=1000,height=800');
    if (!printWindow) {
      alert('الرجاء السماح بالنوافذ المنبثقة للطباعة.');
      return;
    }

    const printHtml = `
      <!doctype html>
      <html lang="ar" dir="rtl">
      <head>
        <meta charset="utf-8">
        <title>كشف التأخر الصباحي - مدرسة الجشة المتوسطة</title>
        <style>
          @import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@400;500;700;800&display=swap');
          body {
            font-family: 'Tajawal', sans-serif;
            color: #1e293b;
            margin: 40px;
            direction: rtl;
            text-align: right;
            line-height: 1.6;
          }
          .report-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-bottom: 3px solid #d97706;
            padding-bottom: 20px;
            margin-bottom: 25px;
          }
          .school-info {
            font-size: 14px;
            font-weight: bold;
          }
          .report-title {
            text-align: center;
            font-size: 24px;
            color: #b45309;
            font-weight: 800;
          }
          .summary-cards {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 15px;
            margin-bottom: 25px;
          }
          .card {
            background-color: #fffbeb;
            border: 1px solid #fde68a;
            border-radius: 8px;
            padding: 12px 18px;
            text-align: center;
          }
          .card-num {
            font-size: 22px;
            font-weight: 800;
            color: #b45309;
          }
          .card-lbl {
            font-size: 12px;
            color: #78350f;
            font-weight: bold;
          }
          table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 30px;
          }
          th, td {
            border: 1px solid #cbd5e1;
            padding: 10px;
            text-align: right;
            font-size: 12px;
          }
          th {
            background-color: #fef3c7;
            color: #92400e;
            font-weight: bold;
          }
        </style>
      </head>
      <body>
        <div class="report-header">
          <div class="school-info">
            وزارة التعليم<br>
            إدارة التعليم بالأحساء<br>
            مدرسة الجشة المتوسطة
          </div>
          <div class="report-title">كشف رصد التأخر الصباحي</div>
          <div class="school-info" style="text-align: left;">
            التاريخ: ${new Date().toLocaleDateString('ar-SA')}م (${toHijri(new Date())})<br>
            بداية الدوام: ${config.start}
          </div>
        </div>

        <div style="background-color: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 12px; margin-bottom: 20px; font-size:13px; font-weight: bold; color: #b45309;">
          ${selectedMonth === 'all' ? 'الفترة المستهدفة: كامل الفصل الدراسي / كل الأشهر' : `الشهر المستهدف: ${selectedMonth}`}
          ${selectedEmployee !== 'all' ? ` | الموظف: ${selectedEmployee}` : ''}
        </div>

        <div class="summary-cards">
          <div class="card">
            <div class="card-num">${list.length}</div>
            <div class="card-lbl">إجمالي حالات التأخر المسجلة</div>
          </div>
          <div class="card">
            <div class="card-num">${totalLateMins} دقيقة</div>
            <div class="card-lbl">مجموع دقائق التأخر</div>
          </div>
        </div>

        <table>
          <thead>
            <tr>
              <th style="width: 5%; text-align:center;">#</th>
              <th>اسم المعلم / الموظف</th>
              <th>المسمى الوظيفي</th>
              <th style="text-align:center;">التاريخ</th>
              <th style="text-align:center;">اليوم</th>
              <th style="text-align:center;">وقت الحضور الفعلي</th>
              <th style="text-align:center;">مدة التأخر</th>
            </tr>
          </thead>
          <tbody>
            ${list.length === 0 ? `
              <tr>
                <td colspan="7" style="text-align:center; padding: 20px; color: #047857; font-weight:bold;">
                  لا توجد أي حالات تأخر مسجلة خلال الفترة المحددة ✨
                </td>
              </tr>
            ` : list.map((r, idx) => `
              <tr>
                <td style="text-align:center;">${idx + 1}</td>
                <td><b>${r.name}</b></td>
                <td>${r.job || '-'}</td>
                <td style="text-align:center;">${formatPrintDate(r.date)}</td>
                <td style="text-align:center;">${dayName(r.date)}</td>
                <td style="text-align:center; color:#b91c1c; font-weight: bold;">${r.tin?.text || '-'}</td>
                <td style="text-align:center; font-weight: bold; color:#b45309;">${r.late} دقيقة</td>
              </tr>
            `).join('')}
          </tbody>
        </table>

        <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-top: 45px; font-weight:bold; font-size:13px;">
          <div style="text-align: right; min-width: 260px;">
            <div>المعلم / الموظف: ${selectedEmployee !== 'all' ? selectedEmployee : (Array.from(new Set(list.map(r => r.name))).length === 1 ? Array.from(new Set(list.map(r => r.name)))[0] : '...........................................')}</div>
            <div style="margin-top: 18px;">التوقيع: ............................................</div>
          </div>
          <div style="text-align: left; min-width: 260px;">
            <div>مدير المدرسة: ............................................</div>
            <div style="margin-top: 18px;">التوقيع: ............................................</div>
          </div>
        </div>

        <script>
          window.onload = function() {
            window.focus();
            window.print();
          }
        <\/script>
      </body>
      </html>
    `;

    printWindow.document.open();
    printWindow.document.write(printHtml);
    printWindow.document.close();
  };

  const handleExportLateListWord = () => {
    const list = filterDailyRecords(dailyRecords)
      .filter(r => r.late > 0)
      .sort((a, b) => new Date(a.date).getTime() - new Date(b.date).getTime());
    const totalLateMins = list.reduce((acc, r) => acc + r.late, 0);
    const targetTeacher = selectedEmployee !== 'all' 
      ? selectedEmployee 
      : (Array.from(new Set(list.map(r => r.name))).length === 1 
          ? Array.from(new Set(list.map(r => r.name)))[0] 
          : '');
    const wordHtml = `
      <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
      <head>
        <title>كشف التأخر الصباحي</title>
        <style>
          body { font-family: 'Arial', sans-serif; direction: rtl; text-align: right; }
          .header { border-bottom: 2px solid #d97706; padding-bottom: 10px; margin-bottom: 20px; }
          .table { width: 100%; border-collapse: collapse; margin-top: 15px; }
          .table th, .table td { border: 1px solid #999; padding: 8px; text-align: right; }
          .table th { background-color: #d97706; color: #ffffff; }
        </style>
      </head>
      <body>
        <div class="header">
          <h2>مدرسة الجشة المتوسطة - إدارة التعليم بالأحساء</h2>
          <h3>كشف التأخر الصباحي (إجمالي الحالات: ${list.length} | مجموع الدقائق: ${totalLateMins} دقيقة)</h3>
          <p>
            ${selectedMonth === 'all' ? 'الفترة: كامل الفصل الدراسي' : `الشهر: ${selectedMonth}`}
            | تاريخ التصدير: ${new Date().toLocaleDateString('ar-SA')}م (${toHijri(new Date())})
          </p>
        </div>

        <table class="table">
          <thead>
            <tr>
              <th>#</th>
              <th>اسم المعلم / الموظف</th>
              <th>المسمى الوظيفي</th>
              <th style="text-align:center;">التاريخ</th>
              <th style="text-align:center;">اليوم</th>
              <th style="text-align:center;">وقت الحضور</th>
              <th style="text-align:center;">مدة التأخر</th>
            </tr>
          </thead>
          <tbody>
            ${list.map((r, idx) => `
              <tr>
                <td>${idx + 1}</td>
                <td><b>${r.name}</b></td>
                <td>${r.job || '-'}</td>
                <td style="text-align:center;">${formatWordDate(r.date)}</td>
                <td style="text-align:center;">${dayName(r.date)}</td>
                <td style="text-align:center;">${r.tin?.text || '-'}</td>
                <td style="text-align:center;">${r.late} دقيقة</td>
              </tr>
            `).join('')}
          </tbody>
        </table>

        <div style="margin-top: 40px; font-size: 13px; font-weight: bold;">
          <table style="width: 100%; border: none;">
            <tr>
              <td style="border: none; text-align: right; width: 50%;">
                المعلم / الموظف: ${targetTeacher || '...........................................'}<br><br>
                التوقيع: ............................................
              </td>
              <td style="border: none; text-align: left; width: 50%;">
                مدير المدرسة: ............................................<br><br>
                التوقيع: ............................................
              </td>
            </tr>
          </table>
        </div>
      </body>
      </html>
    `;

    const blob = new Blob(['\ufeff' + wordHtml], { type: 'application/msword' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `كشف_التأخر_الصباحي_${(selectedMonth === 'all' ? 'كامل' : selectedMonth).replace(/\s+/g, '_')}.doc`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const handlePrintAbsenceList = () => {
    const list = filterDailyRecords(dailyRecords)
      .filter(r => r.type === 'absence')
      .sort((a, b) => new Date(a.date).getTime() - new Date(b.date).getTime());
    const absentTeachers = Array.from(new Set(list.map(r => r.name)));
    const printWindow = window.open('', '_blank', 'width=1000,height=800');
    if (!printWindow) {
      alert('الرجاء السماح بالنوافذ المنبثقة للطباعة.');
      return;
    }

    const printHtml = `
      <!doctype html>
      <html lang="ar" dir="rtl">
      <head>
        <meta charset="utf-8">
        <title>كشف الغياب بدون عذر - مدرسة الجشة المتوسطة</title>
        <style>
          @import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@400;500;700;800&display=swap');
          body {
            font-family: 'Tajawal', sans-serif;
            color: #1e293b;
            margin: 40px;
            direction: rtl;
            text-align: right;
            line-height: 1.6;
          }
          .report-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-bottom: 3px solid #e11d48;
            padding-bottom: 20px;
            margin-bottom: 25px;
          }
          .school-info {
            font-size: 14px;
            font-weight: bold;
          }
          .report-title {
            text-align: center;
            font-size: 24px;
            color: #be123c;
            font-weight: 800;
          }
          .summary-cards {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 15px;
            margin-bottom: 25px;
          }
          .card {
            background-color: #fff1f2;
            border: 1px solid #fecdd3;
            border-radius: 8px;
            padding: 12px 18px;
            text-align: center;
          }
          .card-num {
            font-size: 22px;
            font-weight: 800;
            color: #e11d48;
          }
          .card-lbl {
            font-size: 12px;
            color: #9f1239;
            font-weight: bold;
          }
          table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 30px;
          }
          th, td {
            border: 1px solid #cbd5e1;
            padding: 10px;
            text-align: right;
            font-size: 12px;
          }
          th {
            background-color: #ffe4e6;
            color: #9f1239;
            font-weight: bold;
          }
        </style>
      </head>
      <body>
        <div class="report-header">
          <div class="school-info">
            وزارة التعليم<br>
            إدارة التعليم بالأحساء<br>
            مدرسة الجشة المتوسطة
          </div>
          <div class="report-title">كشف رصد الغياب بدون عذر</div>
          <div class="school-info" style="text-align: left;">
            التاريخ: ${new Date().toLocaleDateString('ar-SA')}م (${toHijri(new Date())})<br>
          </div>
        </div>

        <div style="background-color: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 12px; margin-bottom: 20px; font-size:13px; font-weight: bold; color: #be123c;">
          ${selectedMonth === 'all' ? 'الفترة المستهدفة: كامل الفصل الدراسي / كل الأشهر' : `الشهر المستهدف: ${selectedMonth}`}
          ${selectedEmployee !== 'all' ? ` | الموظف: ${selectedEmployee}` : ''}
        </div>

        <div class="summary-cards">
          <div class="card">
            <div class="card-num">${list.length}</div>
            <div class="card-lbl">إجمالي أيام الغياب غير المبرر</div>
          </div>
          <div class="card">
            <div class="card-num">${absentTeachers.length}</div>
            <div class="card-lbl">عدد المعلمين/الموظفين المتغيبين</div>
          </div>
        </div>

        <table>
          <thead>
            <tr>
              <th style="width: 5%; text-align:center;">#</th>
              <th>اسم المعلم / الموظف</th>
              <th>المسمى الوظيفي</th>
              <th>السجل المدني</th>
              <th style="text-align:center;">تاريخ الغياب</th>
              <th style="text-align:center;">اليوم</th>
              <th style="text-align:center;">الحالة</th>
            </tr>
          </thead>
          <tbody>
            ${list.length === 0 ? `
              <tr>
                <td colspan="7" style="text-align:center; padding: 20px; color: #047857; font-weight:bold;">
                  لا توجد أي حالات غياب بدون عذر مسجلة خلال الفترة ✨
                </td>
              </tr>
            ` : list.map((r, idx) => `
              <tr>
                <td style="text-align:center;">${idx + 1}</td>
                <td><b>${r.name}</b></td>
                <td>${r.job || '-'}</td>
                <td>${r.civil || '-'}</td>
                <td style="text-align:center;">${formatPrintDate(r.date)}</td>
                <td style="text-align:center;">${dayName(r.date)}</td>
                <td style="text-align:center; font-weight: bold; color:#e11d48;">غياب بدون عذر</td>
              </tr>
            `).join('')}
          </tbody>
        </table>

        <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-top: 45px; font-weight:bold; font-size:13px;">
          <div style="text-align: right; min-width: 260px;">
            <div>المعلم / الموظف: ${selectedEmployee !== 'all' ? selectedEmployee : (absentTeachers.length === 1 ? absentTeachers[0] : '...........................................')}</div>
            <div style="margin-top: 18px;">التوقيع: ............................................</div>
          </div>
          <div style="text-align: left; min-width: 260px;">
            <div>مدير المدرسة: ............................................</div>
            <div style="margin-top: 18px;">التوقيع: ............................................</div>
          </div>
        </div>

        <script>
          window.onload = function() {
            window.focus();
            window.print();
          }
        <\/script>
      </body>
      </html>
    `;

    printWindow.document.open();
    printWindow.document.write(printHtml);
    printWindow.document.close();
  };

  const handleExportAbsenceListWord = () => {
    const list = filterDailyRecords(dailyRecords)
      .filter(r => r.type === 'absence')
      .sort((a, b) => new Date(a.date).getTime() - new Date(b.date).getTime());
    const absentTeachers = Array.from(new Set(list.map(r => r.name)));
    const targetTeacher = selectedEmployee !== 'all' 
      ? selectedEmployee 
      : (absentTeachers.length === 1 ? absentTeachers[0] : '');
    const wordHtml = `
      <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
      <head>
        <title>كشف الغياب بدون عذر</title>
        <style>
          body { font-family: 'Arial', sans-serif; direction: rtl; text-align: right; }
          .header { border-bottom: 2px solid #e11d48; padding-bottom: 10px; margin-bottom: 20px; }
          .table { width: 100%; border-collapse: collapse; margin-top: 15px; }
          .table th, .table td { border: 1px solid #999; padding: 8px; text-align: right; }
          .table th { background-color: #e11d48; color: #ffffff; }
        </style>
      </head>
      <body>
        <div class="header">
          <h2>مدرسة الجشة المتوسطة - إدارة التعليم بالأحساء</h2>
          <h3>كشف الغياب بدون عذر (إجمالي أيام الغياب: ${list.length})</h3>
          <p>
            ${selectedMonth === 'all' ? 'الفترة: كامل الفصل الدراسي' : `الشهر: ${selectedMonth}`}
            | تاريخ التصدير: ${new Date().toLocaleDateString('ar-SA')}م (${toHijri(new Date())})
          </p>
        </div>

        <table class="table">
          <thead>
            <tr>
              <th>#</th>
              <th>اسم المعلم / الموظف</th>
              <th>المسمى الوظيفي</th>
              <th>السجل المدني</th>
              <th style="text-align:center;">تاريخ الغياب</th>
              <th style="text-align:center;">اليوم</th>
              <th style="text-align:center;">الحالة</th>
            </tr>
          </thead>
          <tbody>
            ${list.map((r, idx) => `
              <tr>
                <td>${idx + 1}</td>
                <td><b>${r.name}</b></td>
                <td>${r.job || '-'}</td>
                <td>${r.civil || '-'}</td>
                <td style="text-align:center;">${formatWordDate(r.date)}</td>
                <td style="text-align:center;">${dayName(r.date)}</td>
                <td style="text-align:center;">غياب بدون عذر</td>
              </tr>
            `).join('')}
          </tbody>
        </table>

        <div style="margin-top: 40px; font-size: 13px; font-weight: bold;">
          <table style="width: 100%; border: none;">
            <tr>
              <td style="border: none; text-align: right; width: 50%;">
                المعلم / الموظف: ${targetTeacher || '...........................................'}<br><br>
                التوقيع: ............................................
              </td>
              <td style="border: none; text-align: left; width: 50%;">
                مدير المدرسة: ............................................<br><br>
                التوقيع: ............................................
              </td>
            </tr>
          </table>
        </div>
      </body>
      </html>
    `;

    const blob = new Blob(['\ufeff' + wordHtml], { type: 'application/msword' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `كشف_الغياب_بدون_عذر_${(selectedMonth === 'all' ? 'كامل' : selectedMonth).replace(/\s+/g, '_')}.doc`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const handlePrintExcusesList = () => {
    const list = filterDailyRecords(dailyRecords)
      .filter(r => r.type === 'excuse')
      .sort((a, b) => new Date(a.date).getTime() - new Date(b.date).getTime());
    const printWindow = window.open('', '_blank', 'width=1000,height=800');
    if (!printWindow) {
      alert('الرجاء السماح بالنوافذ المنبثقة للطباعة.');
      return;
    }

    const printHtml = `
      <!doctype html>
      <html lang="ar" dir="rtl">
      <head>
        <meta charset="utf-8">
        <title>كشف الغياب بعذر والأعذار المعتمدة - مدرسة الجشة المتوسطة</title>
        <style>
          @import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@400;500;700;800&display=swap');
          body {
            font-family: 'Tajawal', sans-serif;
            color: #1e293b;
            margin: 40px;
            direction: rtl;
            text-align: right;
            line-height: 1.6;
          }
          .report-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-bottom: 3px solid #0284c7;
            padding-bottom: 20px;
            margin-bottom: 25px;
          }
          .school-info {
            font-size: 14px;
            font-weight: bold;
          }
          .report-title {
            text-align: center;
            font-size: 24px;
            color: #0369a1;
            font-weight: 800;
          }
          .summary-cards {
            display: grid;
            grid-template-columns: 1fr;
            gap: 15px;
            margin-bottom: 25px;
          }
          .card {
            background-color: #f0f9ff;
            border: 1px solid #bae6fd;
            border-radius: 8px;
            padding: 12px 18px;
            text-align: center;
          }
          .card-num {
            font-size: 22px;
            font-weight: 800;
            color: #0284c7;
          }
          .card-lbl {
            font-size: 12px;
            color: #0369a1;
            font-weight: bold;
          }
          table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 30px;
          }
          th, td {
            border: 1px solid #cbd5e1;
            padding: 10px;
            text-align: right;
            font-size: 12px;
          }
          th {
            background-color: #e0f2fe;
            color: #0369a1;
            font-weight: bold;
          }
        </style>
      </head>
      <body>
        <div class="report-header">
          <div class="school-info">
            وزارة التعليم<br>
            إدارة التعليم بالأحساء<br>
            مدرسة الجشة المتوسطة
          </div>
          <div class="report-title">كشف الغياب بعذر والأعذار المعتمدة</div>
          <div class="school-info" style="text-align: left;">
            التاريخ: ${new Date().toLocaleDateString('ar-SA')}م (${toHijri(new Date())})<br>
          </div>
        </div>

        <div style="background-color: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 12px; margin-bottom: 20px; font-size:13px; font-weight: bold; color: #0284c7;">
          ${selectedMonth === 'all' ? 'الفترة المستهدفة: كامل الفصل الدراسي / كل الأشهر' : `الشهر المستهدف: ${selectedMonth}`}
          ${selectedEmployee !== 'all' ? ` | الموظف: ${selectedEmployee}` : ''}
        </div>

        <div class="summary-cards">
          <div class="card">
            <div class="card-num">${list.length}</div>
            <div class="card-lbl">إجمالي أيام الغياب المبرر والإجازات المعتمدة</div>
          </div>
        </div>

        <table>
          <thead>
            <tr>
              <th style="width: 5%; text-align:center;">#</th>
              <th>اسم المعلم / الموظف</th>
              <th>المسمى الوظيفي</th>
              <th style="text-align:center;">تاريخ العذر</th>
              <th style="text-align:center;">اليوم</th>
              <th style="text-align:center;">نوع الإجازة / العذر</th>
              <th style="text-align:center;">حالة الاعتماد</th>
            </tr>
          </thead>
          <tbody>
            ${list.length === 0 ? `
              <tr>
                <td colspan="7" style="text-align:center; padding: 20px; color: #64748b; font-weight:bold;">
                  لا توجد أي أعذار أو إجازات مسجلة خلال الفترة
                </td>
              </tr>
            ` : list.map((r, idx) => `
              <tr>
                <td style="text-align:center;">${idx + 1}</td>
                <td><b>${r.name}</b></td>
                <td>${r.job || '-'}</td>
                <td style="text-align:center;">${formatPrintDate(r.date)}</td>
                <td style="text-align:center;">${dayName(r.date)}</td>
                <td style="text-align:center; font-weight: bold; color:#0284c7;">${r.status || 'غياب مبرر'}</td>
                <td style="text-align:center; color:#059669; font-weight:bold;">معتمد رسمياً ✓</td>
              </tr>
            `).join('')}
          </tbody>
        </table>

        <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-top: 45px; font-weight:bold; font-size:13px;">
          <div style="text-align: right; min-width: 260px;">
            <div>المعلم / الموظف: ${selectedEmployee !== 'all' ? selectedEmployee : (Array.from(new Set(list.map(r => r.name))).length === 1 ? Array.from(new Set(list.map(r => r.name)))[0] : '...........................................')}</div>
            <div style="margin-top: 18px;">التوقيع: ............................................</div>
          </div>
          <div style="text-align: left; min-width: 260px;">
            <div>مدير المدرسة: ............................................</div>
            <div style="margin-top: 18px;">التوقيع: ............................................</div>
          </div>
        </div>

        <script>
          window.onload = function() {
            window.focus();
            window.print();
          }
        <\/script>
      </body>
      </html>
    `;

    printWindow.document.open();
    printWindow.document.write(printHtml);
    printWindow.document.close();
  };

  const handleExportExcusesListWord = () => {
    const list = filterDailyRecords(dailyRecords)
      .filter(r => r.type === 'excuse')
      .sort((a, b) => new Date(a.date).getTime() - new Date(b.date).getTime());
    const targetTeacher = selectedEmployee !== 'all' 
      ? selectedEmployee 
      : (Array.from(new Set(list.map(r => r.name))).length === 1 
          ? Array.from(new Set(list.map(r => r.name)))[0] 
          : '');
    const wordHtml = `
      <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
      <head>
        <title>كشف الغياب بعذر والأعذار</title>
        <style>
          body { font-family: 'Arial', sans-serif; direction: rtl; text-align: right; }
          .header { border-bottom: 2px solid #0284c7; padding-bottom: 10px; margin-bottom: 20px; }
          .table { width: 100%; border-collapse: collapse; margin-top: 15px; }
          .table th, .table td { border: 1px solid #999; padding: 8px; text-align: right; }
          .table th { background-color: #0284c7; color: #ffffff; }
        </style>
      </head>
      <body>
        <div class="header">
          <h2>مدرسة الجشة المتوسطة - إدارة التعليم بالأحساء</h2>
          <h3>كشف الغياب بعذر والأعذار المعتمدة (إجمالي الأيام: ${list.length})</h3>
          <p>
            ${selectedMonth === 'all' ? 'الفترة: كامل الفصل الدراسي' : `الشهر: ${selectedMonth}`}
            | تاريخ التصدير: ${new Date().toLocaleDateString('ar-SA')}م (${toHijri(new Date())})
          </p>
        </div>

        <table class="table">
          <thead>
            <tr>
              <th>#</th>
              <th>اسم المعلم / الموظف</th>
              <th>المسمى الوظيفي</th>
              <th style="text-align:center;">تاريخ العذر</th>
              <th style="text-align:center;">اليوم</th>
              <th style="text-align:center;">نوع الإجازة / العذر</th>
              <th style="text-align:center;">حالة الاعتماد</th>
            </tr>
          </thead>
          <tbody>
            ${list.map((r, idx) => `
              <tr>
                <td>${idx + 1}</td>
                <td><b>${r.name}</b></td>
                <td>${r.job || '-'}</td>
                <td style="text-align:center;">${formatWordDate(r.date)}</td>
                <td style="text-align:center;">${dayName(r.date)}</td>
                <td style="text-align:center;">${r.status || 'غياب مبرر'}</td>
                <td style="text-align:center;">معتمد رسمياً ✓</td>
              </tr>
            `).join('')}
          </tbody>
        </table>

        <div style="margin-top: 40px; font-size: 13px; font-weight: bold;">
          <table style="width: 100%; border: none;">
            <tr>
              <td style="border: none; text-align: right; width: 50%;">
                المعلم / الموظف: ${targetTeacher || '...........................................'}<br><br>
                التوقيع: ............................................
              </td>
              <td style="border: none; text-align: left; width: 50%;">
                مدير المدرسة: ............................................<br><br>
                التوقيع: ............................................
              </td>
            </tr>
          </table>
        </div>
      </body>
      </html>
    `;

    const blob = new Blob(['\ufeff' + wordHtml], { type: 'application/msword' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `كشف_الأعذار_والإجازات_${(selectedMonth === 'all' ? 'كامل' : selectedMonth).replace(/\s+/g, '_')}.doc`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const handlePrintEmployeeReport = (employeeName: string) => {
    const rep = selectedPeriodType === 'range'
      ? getIndividualReport(employeeName, 'all', selectedStartDate, selectedEndDate)
      : getIndividualReport(employeeName, selectedMonth);
    const m = rep.summary;

    const printWindow = window.open('', '_blank', 'width=1000,height=800');
    if (!printWindow) {
      alert('الرجاء السماح بالنوافذ المنبثقة للطباعة.');
      return;
    }

    const printHtml = `
      <!doctype html>
      <html lang="ar" dir="rtl">
      <head>
        <meta charset="utf-8">
        <title>تقرير حضور وانضباط - ${m.name}</title>
        <style>
          @import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@400;500;700;800&display=swap');
          body {
            font-family: 'Tajawal', sans-serif;
            color: #1e293b;
            margin: 40px;
            direction: rtl;
            text-align: right;
            line-height: 1.6;
          }
          .report-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-bottom: 3px solid #0f766e;
            padding-bottom: 20px;
            margin-bottom: 30px;
          }
          .school-info {
            font-size: 14px;
            font-weight: bold;
          }
          .report-title {
            text-align: center;
            font-size: 24px;
            color: #0f766e;
            font-weight: 800;
          }
          .profile-box {
            background-color: #f8fafc;
            border: 1px solid #e2e8f0;
            border-radius: 12px;
            padding: 20px;
            margin-bottom: 24px;
          }
          .profile-title {
            margin: 0 0 8px 0;
            font-size: 18px;
            color: #0f766e;
          }
          .profile-meta {
            font-size: 13px;
            color: #64748b;
          }
          .stats-grid {
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 16px;
            margin-bottom: 30px;
          }
          .stat-card {
            background-color: #f8fafc;
            border: 1px solid #e2e8f0;
            border-radius: 12px;
            padding: 16px;
            text-align: center;
          }
          .stat-label {
            font-size: 11px;
            color: #64748b;
            font-weight: bold;
            margin-bottom: 6px;
          }
          .stat-val {
            font-size: 20px;
            font-weight: 800;
            color: #0f766e;
          }
          table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 30px;
          }
          th, td {
            border: 1px solid #cbd5e1;
            padding: 10px;
            text-align: right;
            font-size: 12px;
          }
          th {
            background-color: #f1f5f9;
            color: #0f766e;
            font-weight: bold;
          }
          @media print {
            body { margin: 20px; }
          }
        </style>
      </head>
      <body>
        <div class="report-header">
          <div class="school-info">
            وزارة التعليم<br>
            إدارة التعليم بالأحساء<br>
            مدرسة الجشة المتوسطة
          </div>
          <div class="report-title">تقرير الانضباط والحضور اليومي</div>
          <div class="school-info" style="text-align: left;">
            التاريخ: ${new Date().toLocaleDateString('ar-SA')}م (${toHijri(new Date())})<br>
          </div>
        </div>

        <div class="profile-box">
          <h2 class="profile-title">${m.name}</h2>
          <div class="profile-meta">المسمى الوظيفي: ${m.job || 'غير محدد'} | السجل المدني: ${m.civil || 'غير محدد'}</div>
          <div class="profile-meta" style="margin-top:4px;">الفترة المشمولة بالتقرير: ${m.month === 'all' ? 'جميع الأشهر' : m.month}</div>
        </div>

        <div class="stats-grid">
          <div class="stat-card">
            <div class="stat-label">أيام الدوام المشمولة</div>
            <div class="stat-val">${m.work}</div>
          </div>
          <div class="stat-card">
            <div class="stat-label">أيام الحضور الفعلي</div>
            <div class="stat-val">${m.present}</div>
          </div>
          <div class="stat-card">
            <div class="stat-label">أيام الغياب غير المبرر</div>
            <div class="stat-val" style="color:#dc2626">${m.absence}</div>
          </div>
          <div class="stat-card">
            <div class="stat-label">أيام الغياب المبرر/الأعذار</div>
            <div class="stat-val" style="color:#2563eb">${m.excuse}</div>
          </div>
          <div class="stat-card">
            <div class="stat-label">مرات التأخر الصباحي</div>
            <div class="stat-val" style="color:#d97706">${m.lateCount}</div>
          </div>
          <div class="stat-card">
            <div class="stat-label">إجمالي دقائق التأخر</div>
            <div class="stat-val" style="color:#d97706">${m.lateMins}</div>
          </div>
          <div class="stat-card">
            <div class="stat-label">انصراف مبكر / تلقائي</div>
            <div class="stat-val" style="color:#ea580c">${m.earlyCount + m.autoCount}</div>
          </div>
          <div class="stat-card">
            <div class="stat-label">درجة الانضباط النهائية</div>
            <div class="stat-val" style="color:#16a34a">${m.rate}%</div>
          </div>
        </div>

        <div style="background-color: #fff7ed; border: 1px solid #fed7aa; border-radius: 8px; padding: 12px; margin-bottom: 20px; font-size:13px; color:#9a3412">
          <strong>سبب تقييم الانضباط:</strong> ${m.disciplineReason}
        </div>

        <h3>جدول الحضور والتأخر التفصيلي</h3>
        <table>
          <thead>
            <tr>
              <th>التاريخ</th>
              <th>اليوم</th>
              <th>حالة الدوام</th>
              <th>وقت الدخول</th>
              <th>وقت الانصراف</th>
              <th>حالة الانصراف</th>
              <th>التأخر (دقيقة)</th>
              <th>الخروج المبكر</th>
            </tr>
          </thead>
          <tbody>
            ${rep.details.map(r => `
              <tr>
                <td>${formatPrintDate(r.date)}</td>
                <td>${dayName(r.date)}</td>
                <td>${r.type === 'present' ? 'حاضر' : (r.type === 'excuse' ? 'غياب بعذر' : 'غياب بدون عذر')}</td>
                <td>${r.tin?.text || '-'}</td>
                <td>${r.tout?.text || '-'}</td>
                <td>${r.checkout || '-'}</td>
                <td>${r.late || 0}</td>
                <td>${r.earlyMins || 0}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>

        <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-top: 45px; font-weight:bold; font-size:13px;">
          <div style="text-align: right; min-width: 260px;">
            <div>المعلم / الموظف: ${m.name}</div>
            <div style="margin-top: 18px;">التوقيع: ............................................</div>
          </div>
          <div style="text-align: left; min-width: 260px;">
            <div>مدير المدرسة: ............................................</div>
            <div style="margin-top: 18px;">التوقيع: ............................................</div>
          </div>
        </div>

        <script>
          window.onload = function() {
            window.focus();
            window.print();
          }
        <\/script>
      </body>
      </html>
    `;

    printWindow.document.open();
    printWindow.document.write(printHtml);
    printWindow.document.close();
  };

  const handleExportEmployeeWord = (employeeName: string) => {
    const rep = selectedPeriodType === 'range'
      ? getIndividualReport(employeeName, 'all', selectedStartDate, selectedEndDate)
      : getIndividualReport(employeeName, selectedMonth);
    const m = rep.summary;

    const wordHtml = `
      <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
      <head>
        <title>تقرير حضور وانضباط</title>
        <style>
          body { font-family: 'Arial', sans-serif; direction: rtl; text-align: right; }
          .header { border-bottom: 2px solid #0f766e; padding-bottom: 10px; margin-bottom: 20px; }
          .table { width: 100%; border-collapse: collapse; margin-top: 15px; }
          .table th, .table td { border: 1px solid #999; padding: 8px; text-align: right; }
          .table th { background-color: #0f766e; color: #ffffff; }
        </style>
      </head>
      <body>
        <div class="header">
          <h2>مدرسة الجشة المتوسطة - إدارة التعليم بالأحساء</h2>
          <h3>تقرير الانضباط والحضور الشامل</h3>
          <p>الموظف: <b>${m.name}</b></p>
          <p>المسمى الوظيفي: ${m.job || '-'} | السجل المدني: ${m.civil || '-'}</p>
          <p>الفترة: ${m.month} | تاريخ التصدير: ${new Date().toLocaleDateString('ar-SA')}م (${toHijri(new Date())})</p>
        </div>

        <h3>الملخص الرقمي:</h3>
        <ul>
          <li>أيام الدوام المقررة: ${m.work}</li>
          <li>أيام الحضور الفعلي: ${m.present}</li>
          <li>أيام الغياب غير المبرر: ${m.absence}</li>
          <li>أيام الغياب بعذر: ${m.excuse}</li>
          <li>مرات التأخر الصباحي: ${m.lateCount} (بمجموع ${m.lateMins} دقيقة)</li>
          <li>حالات الانصراف غير المكتمل أو المبكر: ${m.earlyCount + m.autoCount}</li>
          <li>درجة الانضباط العام: ${m.rate}%</li>
          <li>المؤشر العام والتصنيف: <b>${m.rating}</b></li>
        </ul>

        <p><b>مبررات التقييم:</b> ${m.disciplineReason}</p>

        <h3>سجل الدوام اليومي التفصيلي:</h3>
        <table class="table">
          <thead>
            <tr>
              <th>التاريخ</th>
              <th>اليوم</th>
              <th>حالة الحضور</th>
              <th>وقت الدخول</th>
              <th>وقت الانصراف</th>
              <th>حالة الانصراف</th>
              <th>التأخر بالدقائق</th>
              <th>الخروج المبكر</th>
            </tr>
          </thead>
          <tbody>
            ${rep.details.map(r => `
              <tr>
                <td>${formatWordDate(r.date)}</td>
                <td>${dayName(r.date)}</td>
                <td>${r.type === 'present' ? 'حاضر' : (r.type === 'excuse' ? 'غياب بعذر' : 'غياب بدون عذر')}</td>
                <td>${r.tin?.text || '-'}</td>
                <td>${r.tout?.text || '-'}</td>
                <td>${r.checkout || '-'}</td>
                <td>${r.late || 0}</td>
                <td>${r.earlyMins || 0}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>

        <div style="margin-top: 40px; font-size: 13px; font-weight: bold;">
          <table style="width: 100%; border: none;">
            <tr>
              <td style="border: none; text-align: right; width: 50%;">
                المعلم / الموظف: ${m.name}<br><br>
                التوقيع: ............................................
              </td>
              <td style="border: none; text-align: left; width: 50%;">
                مدير المدرسة: ............................................<br><br>
                التوقيع: ............................................
              </td>
            </tr>
          </table>
        </div>
      </body>
      </html>
    `;

    const blob = new Blob(['\\ufeff' + wordHtml], { type: 'application/msword' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `تقرير_انضباط_${m.name.replace(/\\s+/g, '_')}.doc`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const getWhatsAppMessageForEmployee = (summary: EmployeeSummary) => {
    const isExcellent = 
      (summary.lateCount || 0) <= 2 && 
      (summary.earlyCount || 0) <= 1 && 
      (summary.autoCount || 0) === 0 && 
      (summary.missingCheckoutCount || 0) === 0 && 
      (summary.absence || 0) === 0;

    const isNeedsSupport = 
      (summary.lateCount || 0) >= 3 || 
      (summary.earlyCount || 0) >= 5 || 
      ((summary.autoCount || 0) + (summary.missingCheckoutCount || 0)) >= 3;

    if (isExcellent) {
      return `السلام عليكم ورحمة الله وبركاته\n\nالأستاذ الموقر: ${summary.name}\n\nيسر إدارة مدرسة الجشة المتوسطة الإشادة بمستوى انضباطكم الاستثنائي وحصولكم على تصنيف (المنضبطون المتميزون) بمعدل استثنائي قدره (${summary.rate ?? 100}%) خلال هذه الفترة.\n\nنشكر لكم هذا الحرص الدائم والعمل الدؤوب، متمنين لكم دوام التوفيق والتميز.`;
    } else if (isNeedsSupport) {
      let reasons = [];
      if ((summary.lateCount || 0) >= 3) reasons.push(`تكرار التأخر الصباحي: ${summary.lateCount} مرات`);
      if ((summary.earlyCount || 0) >= 5) reasons.push(`تكرار الانصراف المبكر: ${summary.earlyCount} مرات`);
      const autoAndMissing = (summary.autoCount || 0) + (summary.missingCheckoutCount || 0);
      if (autoAndMissing >= 3) reasons.push(`تكرار الانصراف التلقائي: ${autoAndMissing} مرات`);

      return `السلام عليكم ورحمة الله وبركاته\n\nالأستاذ الموقر: ${summary.name}\n\nنحيطكم علماً بوجود مؤشرات أداء تتطلب المتابعة والدعم والمراجعة معكم للعمل على تحسينها وتلافيها مستقبلاً:\n${reasons.map(r => `• ${r}`).join('\n')}\n\nنأمل منكم مراجعة إدارة المدرسة وتقديم الأعذار اللازمة ومناقشة الدعم والمساندة لرفع معدلات الانتظام. شاكرين تعاونكم.`;
    } else {
      return `السلام عليكم ورحمة الله وبركاته\n\nالأستاذ الموقر: ${summary.name}\n\nنود إفادتكم بتقرير انضباطكم ومؤشر حضوركم المسجل خلال هذه الفترة بمتوسط (${summary.rate ?? 100}%). \n\nنشكر لكم جهودكم وتفهمكم المستمر.`;
    }
  };

  // 9. Quick WhatsApp for Top 10 / Custom lists
  const handleQuickWhatsApp = (summary: EmployeeSummary) => {
    const msg = getWhatsAppMessageForEmployee(summary);
    setModalActiveRecordKey(null);
    setModalTitle(`رسالة واتساب - ${summary.name}`);
    setModalText(msg);
    setModalOpen(true);
  };

  return (
    <div className="grid grid-cols-1 md:grid-cols-[280px_1fr] min-height-screen bg-[#f4f7fb] text-[#14213d] min-h-screen">
      {/* SIDEBAR */}
      <aside className="bg-radial-[circle_at_20%_0%] from-teal-800/20 via-teal-900 to-[#052d3c] text-white p-6 flex flex-col justify-between md:sticky md:top-0 md:h-screen overflow-y-auto">
        <div>
          <div className="text-center mb-6">
            <h1 className="font-extrabold text-2xl tracking-wide text-shadow-md">تقارير برنامج حضوري</h1>
            <div className="h-0.5 w-24 bg-gradient-to-r from-transparent via-cyan-400 to-transparent mx-auto mt-3"></div>
            <p className="text-xs text-teal-300 font-medium mt-1">مدرسة الجشة المتوسطة</p>
          </div>

          <nav className="space-y-2 mt-4">
            <button
              onClick={() => { setCurrentTab('dashboard'); setSelectedStatus('all'); }}
              className={`w-full flex items-center justify-between p-3.5 rounded-xl text-right font-bold transition-all border border-white/5 ${
                currentTab === 'dashboard'
                  ? 'bg-white text-teal-950 shadow-md'
                  : 'bg-white/5 text-white hover:bg-white/10'
              }`}
            >
              <span className="flex items-center gap-3">
                <LayoutDashboard className="w-5 h-5 text-teal-400" />
                <span>لوحة التحكم</span>
              </span>
            </button>

            <button
              onClick={() => { setCurrentTab('attendance'); setSelectedStatus('all'); }}
              className={`w-full flex items-center justify-between p-3.5 rounded-xl text-right font-bold transition-all border border-white/5 ${
                currentTab === 'attendance'
                  ? 'bg-white text-teal-950 shadow-md'
                  : 'bg-white/5 text-white hover:bg-white/10'
              }`}
            >
              <span className="flex items-center gap-3">
                <Users className="w-5 h-5 text-teal-400" />
                <span>الحضور والمنسوبين</span>
              </span>
              {dailyRecords.length > 0 && (
                <span className="bg-teal-500/20 text-teal-300 text-xs px-2 py-1 rounded-full">
                  {employees.length}
                </span>
              )}
            </button>

            <button
              onClick={() => { setCurrentTab('late'); setSelectedStatus('all'); }}
              className={`w-full flex items-center justify-between p-3.5 rounded-xl text-right font-bold transition-all border border-white/5 ${
                currentTab === 'late'
                  ? 'bg-white text-teal-950 shadow-md'
                  : 'bg-white/5 text-white hover:bg-white/10'
              }`}
            >
              <span className="flex items-center gap-3">
                <Clock className="w-5 h-5 text-amber-400" />
                <span>التأخر الصباحي</span>
              </span>
              {dailyRecords.length > 0 && stats.totalLateCount > 0 && (
                <span className="bg-amber-500/30 text-amber-300 text-xs px-2 py-1 rounded-full">
                  {stats.totalLateCount}
                </span>
              )}
            </button>

            <button
              onClick={() => { setCurrentTab('absence'); setSelectedStatus('all'); }}
              className={`w-full flex items-center justify-between p-3.5 rounded-xl text-right font-bold transition-all border border-white/5 ${
                currentTab === 'absence'
                  ? 'bg-white text-teal-950 shadow-md'
                  : 'bg-white/5 text-white hover:bg-white/10'
              }`}
            >
              <span className="flex items-center gap-3">
                <CalendarX className="w-5 h-5 text-rose-400" />
                <span>الغياب بدون عذر</span>
              </span>
              {dailyRecords.length > 0 && stats.totalAbsence > 0 && (
                <span className="bg-rose-500/30 text-rose-300 text-xs px-2 py-1 rounded-full">
                  {stats.totalAbsence}
                </span>
              )}
            </button>

            <button
              onClick={() => { setCurrentTab('excuses'); setSelectedStatus('all'); }}
              className={`w-full flex items-center justify-between p-3.5 rounded-xl text-right font-bold transition-all border border-white/5 ${
                currentTab === 'excuses'
                  ? 'bg-white text-teal-950 shadow-md'
                  : 'bg-white/5 text-white hover:bg-white/10'
              }`}
            >
              <span className="flex items-center gap-3">
                <FileText className="w-5 h-5 text-cyan-400" />
                <span>الغياب بعذر والأعذار</span>
              </span>
            </button>

            <button
              onClick={() => {
                if (selectedEmployee && selectedEmployee !== 'all') {
                  setMonthlyEmployee(selectedEmployee);
                }
                if (selectedPeriodType === 'range') {
                  setMonthlyPeriodType('range');
                  if (selectedStartDate) setMonthlyStartDate(selectedStartDate);
                  if (selectedEndDate) setMonthlyEndDate(selectedEndDate);
                } else {
                  setMonthlyPeriodType('month');
                  setMonthlyMonth(selectedMonth || 'all');
                }
                setCurrentTab('monthly');
              }}
              className={`w-full flex items-center justify-between p-3.5 rounded-xl text-right font-bold transition-all border border-white/5 ${
                currentTab === 'monthly'
                  ? 'bg-white text-teal-950 shadow-md'
                  : 'bg-white/5 text-white hover:bg-white/10'
              }`}
            >
              <span className="flex items-center gap-3">
                <Calendar className="w-5 h-5 text-purple-400" />
                <span>التقرير الشهري الشامل</span>
              </span>
            </button>

            <button
              onClick={() => setCurrentTab('settings')}
              className={`w-full flex items-center justify-between p-3.5 rounded-xl text-right font-bold transition-all border border-white/5 ${
                currentTab === 'settings'
                  ? 'bg-white text-teal-950 shadow-md'
                  : 'bg-white/5 text-white hover:bg-white/10'
              }`}
            >
              <span className="flex items-center gap-3">
                <SettingsIcon className="w-5 h-5 text-slate-400" />
                <span>إعدادات النظام</span>
              </span>
            </button>
          </nav>
        </div>

        <div className="mt-8 border-t border-teal-800/40 pt-4 text-center">
          <div className="bg-white/5 p-3.5 rounded-xl border border-teal-500/20 shadow-inner">
            <p className="text-[10px] text-teal-300 font-semibold mb-1">إعداد الموجه الطلابي:</p>
            <p className="text-xs font-bold tracking-tight text-white">عبدالهادي بن محمد المحسن</p>
          </div>
        </div>
      </aside>

      {/* MAIN CONTENT AREA */}
      <main className="p-6 md:p-8 flex flex-col justify-between overflow-x-hidden">
        <div>
          {/* TOP ACTIONS & TITLE */}
          <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4 mb-6 pb-4 border-b border-slate-200">
            <div>
              <h2 className="text-2xl font-extrabold text-teal-900">
                {currentTab === 'dashboard' && '📊 لوحة التحكم والمؤشرات'}
                {currentTab === 'attendance' && '👤 سجل الحضور والمنسوبين'}
                {currentTab === 'late' && '🕒 رصد التأخر الصباحي'}
                {currentTab === 'absence' && '📅 رصد الغياب بدون عذر'}
                {currentTab === 'excuses' && '📄 سجل الغياب بعذر والمسوغات'}
                {currentTab === 'monthly' && '📈 التقارير الدورية والملخصات'}
                {currentTab === 'settings' && '⚙️ إعدادات الحضور والتقييم'}
              </h2>
              <p className="text-sm text-slate-500 mt-1">
                نظام إدارة ومتابعة الانضباط المدرسي اليومي والشهري بمدرسة الجشة المتوسطة
              </p>
            </div>

            <div className="flex flex-wrap items-center gap-2">
              <input
                ref={fileInputRef}
                type="file"
                accept=".xlsx,.xls,.csv"
                onChange={handleFileUpload}
                className="hidden"
                id="excel-file-upload-input"
              />
              <button
                onClick={() => fileInputRef.current?.click()}
                className="bg-teal-700 text-white font-bold px-4 py-2.5 rounded-xl hover:bg-teal-800 transition-all flex items-center gap-2 shadow-sm text-sm"
              >
                <Upload className="w-4 h-4" />
                <span>استيراد ملف البصمة Excel</span>
              </button>

              {rawRows.length > 0 && (
                <button
                  onClick={() => setShowClearConfirm(true)}
                  className="bg-rose-100 text-rose-700 border border-rose-200 font-bold px-3 py-2.5 rounded-xl hover:bg-rose-200 transition-all flex items-center gap-2 text-sm"
                >
                  <Trash2 className="w-4 h-4" />
                  <span>تفريغ البيانات</span>
                </button>
              )}
            </div>
          </div>



          {/* FALLBACK IF NO DATA IMPORTED */}
          {rawRows.length === 0 && currentTab !== 'settings' ? (
            <div className="bg-white border border-slate-200 rounded-2xl p-12 text-center max-w-xl mx-auto my-12 shadow-md">
              <div className="bg-teal-50 w-16 h-16 rounded-full flex items-center justify-center mx-auto mb-4 text-teal-600">
                <Upload className="w-8 h-8" />
              </div>
              <h3 className="text-xl font-extrabold text-teal-950 mb-2">مرحباً بك في نظام تقارير الحضور والانضباط</h3>
              <p className="text-slate-500 text-sm leading-relaxed mb-6">
                البداية تبدأ برفع ملف إكسل المصدر المستخرج مباشرة من جهاز البصمة المدرسي لترجمة البيانات آلياً إلى لوحات تفاعلية، إحصائيات دقيقة، وتقارير رسمية جاهزة للطباعة والتصدير.
              </p>
              <button
                onClick={() => fileInputRef.current?.click()}
                className="bg-teal-700 text-white font-bold px-5 py-3 rounded-xl hover:bg-teal-800 transition-all shadow-md text-sm inline-flex items-center gap-2"
              >
                <Upload className="w-5 h-5" />
                <span>اختر ملف البصمة من جهازك</span>
              </button>
            </div>
          ) : (
            <>
              {/* TABS VIEW RENDERING */}

              {/* A. DASHBOARD TAB */}
              {currentTab === 'dashboard' && (
                <div className="space-y-6">
                  {/* AUTO-DETECTED SUSPENDED DAYS BANNER */}
                  {autoDetectedSuspendedDays.length > 0 && (
                    <div className="bg-amber-50/80 border border-amber-200 p-4 rounded-2xl flex items-center justify-between gap-4 text-xs shadow-xs" dir="rtl">
                      <div className="flex items-center gap-3">
                        <div className="bg-amber-100 text-amber-800 p-2 rounded-xl text-lg flex-shrink-0">
                          🌤️
                        </div>
                        <div>
                          <h4 className="font-extrabold text-amber-950 text-sm">
                            تم استبعاد {autoDetectedSuspendedDays.length} يوم دراسي تلقائياً (تعليق دراسة / غير دراسي لعدم تسجيل أي بصمات لجميع المعلمين):
                          </h4>
                          <div className="flex flex-wrap items-center gap-2 mt-1">
                            {autoDetectedSuspendedDays.map(dk => (
                              <span key={dk} className="bg-white border border-amber-300 font-bold px-2.5 py-0.5 rounded-lg text-amber-900 text-xs">
                                {dayName(dk)} {dk}م ({toHijri(dk)})
                              </span>
                            ))}
                          </div>
                          <p className="text-[11px] text-amber-800 mt-1 leading-relaxed">
                            تمت حماية تقييمات وانضباط المعلمين واستبعاد هذه الأيام تلقائياً حتى لا تُحسب كغياب بدون عذر جماعي.
                          </p>
                        </div>
                      </div>
                      <button
                        onClick={() => setCurrentTab('settings')}
                        className="bg-amber-200 hover:bg-amber-300 text-amber-950 font-extrabold px-3.5 py-2 rounded-xl text-xs transition-all flex-shrink-0 cursor-pointer shadow-2xs"
                      >
                        إدارة أيام التعليق
                      </button>
                    </div>
                  )}

                  {/* SEARCH & FILTER FOR DASHBOARD STATS */}
                  <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-xs flex flex-wrap items-center justify-between gap-4">
                    <div className="flex flex-wrap items-center gap-3">
                      <div>
                        <label className="block text-[10px] font-bold text-slate-400 mb-1">تصفية الفترة حسب</label>
                        <select
                          value={selectedPeriodType}
                          onChange={(e) => setSelectedPeriodType(e.target.value as 'month' | 'range')}
                          className="bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-sm font-bold text-slate-700 outline-none focus:border-teal-500 transition-all"
                        >
                          <option value="month">📅 الشهر</option>
                          <option value="range">🗓️ نطاق تاريخ</option>
                        </select>
                      </div>

                      {selectedPeriodType === 'month' ? (
                        <div>
                          <label className="block text-[10px] font-bold text-slate-400 mb-1">
                            {selectedMonth === 'all' ? 'الفترة المستهدفة' : 'الشهر المستهدف'}
                          </label>
                          <select
                            value={selectedMonth}
                            onChange={(e) => setSelectedMonth(e.target.value)}
                            className="bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-sm font-bold text-slate-700 outline-none focus:border-teal-500 transition-all"
                          >
                            <option value="all">كل الأشهر المسجلة</option>
                            {getMonthsList().map(m => (
                              <option key={m} value={m}>{m}</option>
                            ))}
                          </select>
                        </div>
                      ) : (
                        <>
                          <div>
                            <label className="block text-[10px] font-bold text-slate-400 mb-1">من تاريخ</label>
                            <input
                              type="date"
                              value={selectedStartDate}
                              onChange={(e) => setSelectedStartDate(e.target.value)}
                              className="bg-slate-50 border border-slate-200 rounded-xl px-3 py-1.5 text-sm font-bold text-slate-700 outline-none focus:border-teal-500 transition-all h-[38px]"
                            />
                          </div>
                          <div>
                            <label className="block text-[10px] font-bold text-slate-400 mb-1">إلى تاريخ</label>
                            <input
                              type="date"
                              value={selectedEndDate}
                              onChange={(e) => setSelectedEndDate(e.target.value)}
                              className="bg-slate-50 border border-slate-200 rounded-xl px-3 py-1.5 text-sm font-bold text-slate-700 outline-none focus:border-teal-500 transition-all h-[38px]"
                            />
                          </div>
                        </>
                      )}

                      <div>
                        <label className="block text-[10px] font-bold text-slate-400 mb-1">الموظف / المعلم</label>
                        <select
                          value={selectedEmployee}
                          onChange={(e) => setSelectedEmployee(e.target.value)}
                          className="bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-sm font-bold text-slate-700 outline-none focus:border-teal-500 transition-all max-w-[240px]"
                        >
                          <option value="all">جميع المعلمين والموظفين</option>
                          {getEmployeesNames().map(name => (
                            <option key={name} value={name}>{name}</option>
                          ))}
                        </select>
                      </div>
                    </div>

                    <div className="w-full md:w-auto md:min-w-[260px]">
                      <label className="block text-[10px] font-bold text-slate-400 mb-1">بحث سريع بالاسم أو السجل</label>
                      <input
                        type="text"
                        placeholder="اكتب هنا للبحث..."
                        value={searchTerm}
                        onChange={(e) => setSearchTerm(e.target.value)}
                        className="bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-sm w-full outline-none focus:border-teal-500 transition-all font-medium"
                      />
                    </div>
                  </div>

                  {/* METRIC CARD STRIPS */}
                  <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-5 xl:grid-cols-9 gap-4">
                    <div className="bg-white p-4 rounded-2xl border border-slate-200 shadow-sm flex flex-col justify-between">
                      <span className="text-xs text-slate-400 font-bold">المنسوبين</span>
                      <span className="text-3xl font-extrabold text-blue-700 mt-2">
                        {selectedEmployee !== 'all' ? 1 : employees.length}
                      </span>
                    </div>

                    <div className="bg-white p-4 rounded-2xl border border-slate-200 shadow-sm flex flex-col justify-between">
                      <span className="text-xs text-slate-400 font-bold">الحضور المسجل</span>
                      <span className="text-3xl font-extrabold text-emerald-600 mt-2">
                        {filterDailyRecords(dailyRecords).filter(r => r.type === 'present').length}
                      </span>
                    </div>

                    {/* LEAVES & EXCUSES CARD */}
                    <div className="bg-white p-3.5 rounded-2xl border border-slate-200 shadow-sm flex flex-col justify-between">
                      <div className="flex items-center justify-between">
                        <span className="text-xs text-slate-400 font-bold">بطاقة الإجازات</span>
                        <span className="text-[10px] bg-teal-50 text-teal-700 px-1.5 py-0.5 rounded-md font-bold">فارس</span>
                      </div>
                      <div className="grid grid-cols-2 gap-2 mt-2 pt-1.5 border-t border-slate-50">
                        <div className="border-l border-slate-100 pl-1">
                          <span className="block text-[9px] font-bold text-slate-400">سنوية</span>
                          <span className="text-sm font-extrabold text-emerald-600 block mt-0.5">{dashboardAnnualCount} <span className="text-[9px] font-normal text-slate-400">يوم</span></span>
                        </div>
                        <div className="pr-1">
                          <span className="block text-[9px] font-bold text-slate-400">طارئة</span>
                          <span className="text-sm font-extrabold text-blue-600 block mt-0.5">{dashboardEmergencyCount} <span className="text-[9px] font-normal text-slate-400">يوم</span></span>
                        </div>
                      </div>
                    </div>

                    <div className="bg-white p-4 rounded-2xl border border-slate-200 shadow-sm flex flex-col justify-between">
                      <span className="text-xs text-slate-400 font-bold">حالات التأخر</span>
                      <span className="text-3xl font-extrabold text-amber-600 mt-2">{stats.totalLateCount}</span>
                    </div>

                    <div className="bg-white p-4 rounded-2xl border border-slate-200 shadow-sm flex flex-col justify-between">
                      <span className="text-xs text-slate-400 font-bold">غياب بدون عذر</span>
                      <span className="text-3xl font-extrabold text-rose-600 mt-2">{stats.totalAbsence}</span>
                    </div>

                    <div className="bg-white p-4 rounded-2xl border border-slate-200 shadow-sm flex flex-col justify-between">
                      <span className="text-xs text-slate-400 font-bold">دقائق التأخر</span>
                      <span className="text-3xl font-extrabold text-slate-700 mt-2">{stats.totalLateMins}</span>
                    </div>

                    <div className="bg-white p-4 rounded-2xl border border-slate-200 shadow-sm flex flex-col justify-between">
                      <span className="text-xs text-slate-400 font-bold">دقائق الخروج المبكر</span>
                      <span className="text-3xl font-extrabold text-rose-500 mt-2">{stats.totalEarlyMins}</span>
                    </div>

                    <div className="bg-white p-4 rounded-2xl border border-slate-200 shadow-sm flex flex-col justify-between">
                      <span className="text-xs text-slate-400 font-bold">مفقود الانصراف</span>
                      <span className="text-3xl font-extrabold text-amber-500 mt-2">{stats.totalMissingCheckout}</span>
                    </div>

                    {/* EMERGENCY LEAVE WITHOUT CHECKOUT CARD */}
                    <div
                      onClick={() => {
                        setSelectedStatus('emergency_no_exit');
                        setCurrentTab('attendance');
                      }}
                      className="bg-white p-3.5 rounded-2xl border border-amber-300 shadow-sm flex flex-col justify-between cursor-pointer hover:border-amber-500 hover:shadow-md transition-all group relative overflow-hidden"
                      title="اضغط للتصفية وعرض المعلمين المستأذنين بدون تسجيل بصمة انصراف"
                    >
                      <div className="flex items-center justify-between">
                        <span className="text-xs text-amber-900 font-bold">استئذان بلا بصمة</span>
                        <span className="text-[9px] bg-rose-100 text-rose-800 px-1.5 py-0.5 rounded-md font-extrabold">مخالف</span>
                      </div>
                      <span className="text-3xl font-extrabold text-amber-700 mt-2">{stats.totalEmergencyNoExit}</span>
                      <div className="mt-2 text-[9px] text-amber-800 border-t border-amber-100 pt-1 flex items-center justify-between font-bold group-hover:text-amber-950">
                        <span>عرض الحالات</span>
                        <span>←</span>
                      </div>
                    </div>
                  </div>

                  {/* SMART EXECUTIVE ANALYSIS SECTION */}
                  <div className="bg-gradient-to-r from-teal-50 to-cyan-50 border border-teal-200 rounded-3xl p-5 shadow-xs">
                    <div className="flex items-start gap-3.5">
                      <div className="bg-teal-600 text-white p-2 rounded-xl mt-0.5">
                        <AlertCircle className="w-5 h-5" />
                      </div>
                      <div>
                        <h3 className="font-extrabold text-teal-950 text-base">الملخص التنفيذي الذكي للأداء والانضباط</h3>
                        <p className="text-slate-600 text-sm leading-relaxed mt-1.5">
                          {getSmartExecutiveSummaryText()}
                        </p>
                      </div>
                    </div>
                  </div>

                  {/* RECHARTS DISCIPLINE DISTRIBUTION BAR CHART */}
                  <div className="bg-white border border-slate-200 rounded-3xl p-6 shadow-sm">
                    <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 mb-6 pb-4 border-b border-slate-100">
                      <div>
                        <h3 className="text-lg font-extrabold text-teal-950 flex items-center gap-2">
                          <span>📊 التوزيع البياني لمؤشرات الانضباط والانتظام</span>
                        </h3>
                        <p className="text-xs text-slate-400 mt-1">
                          رسم بياني تفاعلي يوضح توزيع المعلمين والموظفين على فئات الانضباط الأربعة بالمدرسة
                        </p>
                      </div>
                      <div className="bg-teal-50 px-3.5 py-1.5 rounded-full text-xs font-bold text-teal-700 flex items-center gap-1.5 self-start md:self-auto">
                        <span>إجمالي الموظفين الخاضعين للمؤشر:</span>
                        <strong className="font-black text-sm text-teal-900">{stats.totalCount}</strong>
                      </div>
                    </div>

                    <div className="w-full h-[320px]" dir="ltr">
                      <ResponsiveContainer width="100%" height="100%">
                        <BarChart
                          data={[
                            { name: 'منضبط متميز', count: stats.ideal.length, fill: '#10b981', label: 'نسبة التزام استثنائية (>= 90%)' },
                            { name: 'منضبط مع ملاحظات', count: stats.notes.length, fill: '#3b82f6', label: 'يحتاج تحسين طفيف (80-89%)' },
                            { name: 'يحتاج متابعة', count: stats.follow.length, fill: '#f59e0b', label: 'تكرار تأخر أو تقصير (70-79%)' },
                            { name: 'يحتاج إجراء إداري', count: stats.admin.length, fill: '#ef4444', label: 'تجاوز للحدود المسموحة (< 70%)' },
                          ]}
                          margin={{ top: 20, right: 30, left: 10, bottom: 5 }}
                        >
                          <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" vertical={false} />
                          <XAxis 
                            dataKey="name" 
                            stroke="#64748b" 
                            fontSize={12} 
                            tickLine={false} 
                            axisLine={false}
                            dy={10}
                          />
                          <YAxis 
                            stroke="#64748b" 
                            fontSize={12} 
                            tickLine={false} 
                            axisLine={false}
                            allowDecimals={false}
                            dx={-10}
                          />
                          <Tooltip content={<CustomTooltip />} cursor={{ fill: '#f8fafc' }} />
                          <Bar 
                            dataKey="count" 
                            radius={[8, 8, 0, 0]} 
                            maxBarSize={60}
                          >
                            {[
                              { name: 'منضبط متميز', fill: '#10b981' },
                              { name: 'منضبط مع ملاحظات', fill: '#3b82f6' },
                              { name: 'يحتاج متابعة', fill: '#f59e0b' },
                              { name: 'يحتاج إجراء إداري', fill: '#ef4444' },
                            ].map((entry, index) => (
                              <Cell key={`cell-${index}`} fill={entry.fill} />
                            ))}
                          </Bar>
                        </BarChart>
                      </ResponsiveContainer>
                    </div>

                    <div className="grid grid-cols-2 lg:grid-cols-4 gap-4 mt-6 pt-4 border-t border-slate-100" dir="rtl">
                      <div className="flex items-center gap-2.5">
                        <span className="w-3.5 h-3.5 rounded-full bg-emerald-500 block shrink-0" />
                        <div>
                          <span className="block text-[10px] font-bold text-slate-400">منضبط متميز</span>
                          <strong className="text-xs text-slate-700 font-extrabold">{stats.ideal.length} موظف ({stats.totalCount ? Math.round((stats.ideal.length / stats.totalCount) * 100) : 0}%)</strong>
                        </div>
                      </div>
                      <div className="flex items-center gap-2.5">
                        <span className="w-3.5 h-3.5 rounded-full bg-blue-500 block shrink-0" />
                        <div>
                          <span className="block text-[10px] font-bold text-slate-400">منضبط مع ملاحظات</span>
                          <strong className="text-xs text-slate-700 font-extrabold">{stats.notes.length} موظف ({stats.totalCount ? Math.round((stats.notes.length / stats.totalCount) * 100) : 0}%)</strong>
                        </div>
                      </div>
                      <div className="flex items-center gap-2.5">
                        <span className="w-3.5 h-3.5 rounded-full bg-amber-500 block shrink-0" />
                        <div>
                          <span className="block text-[10px] font-bold text-slate-400">يحتاج متابعة</span>
                          <strong className="text-xs text-slate-700 font-extrabold">{stats.follow.length} موظف ({stats.totalCount ? Math.round((stats.follow.length / stats.totalCount) * 100) : 0}%)</strong>
                        </div>
                      </div>
                      <div className="flex items-center gap-2.5">
                        <span className="w-3.5 h-3.5 rounded-full bg-red-500 block shrink-0" />
                        <div>
                          <span className="block text-[10px] font-bold text-slate-400">يحتاج إجراء إداري</span>
                          <strong className="text-xs text-slate-700 font-extrabold">{stats.admin.length} موظف ({stats.totalCount ? Math.round((stats.admin.length / stats.totalCount) * 100) : 0}%)</strong>
                        </div>
                      </div>
                    </div>
                  </div>

                  {/* AUTO ALERT PREPARED CONTROL PANEL */}
                  {config.enableAutoAlert && (
                    <div className="bg-white rounded-3xl border border-teal-200 shadow-md overflow-hidden" dir="rtl">
                      <div className="bg-gradient-to-r from-teal-700 to-teal-900 text-white p-5 flex flex-wrap items-center justify-between gap-4">
                        <div className="flex items-center gap-3">
                          <div className="bg-white/15 p-2.5 rounded-2xl text-teal-200">
                            <MessageSquare className="w-5 h-5 animate-pulse" />
                          </div>
                          <div>
                            <h3 className="font-extrabold text-base">📢 التنبيهات التلقائية المجهزة (تأخر أو انصراف متكرر)</h3>
                            <p className="text-xs text-teal-100/80 mt-0.5">رسائل واتساب منسقة ومعدة للإرسال الفوري لزملاء فئة المتابعة والدعم</p>
                          </div>
                        </div>
                        <div className="bg-teal-950/40 border border-teal-500/20 px-3.5 py-1.5 rounded-full text-xs font-bold text-teal-200 flex items-center gap-2">
                          <span>قيد المتابعة: {stats.needsSupport.filter(m => !sentAlerts[m.name]).length}</span>
                          <span className="opacity-40">|</span>
                          <span className="text-emerald-300">تم إرسالها: {stats.needsSupport.filter(m => sentAlerts[m.name]).length}</span>
                        </div>
                      </div>

                      <div className="p-5">
                        {stats.needsSupport.length === 0 ? (
                          <div className="text-center py-8 text-slate-400">
                            <CheckCircle2 className="w-12 h-12 text-emerald-500 mx-auto mb-2" />
                            <p className="text-sm font-bold text-slate-500">رائع! جميع المنسوبين لديهم مستويات التزام مثالية وممتازة حالياً.</p>
                            <p className="text-xs mt-1">لا توجد حالات تنبيه تلقائي تتطلب المتابعة والدعم.</p>
                          </div>
                        ) : (
                          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                            {stats.needsSupport.map(m => {
                              const msgText = getWhatsAppMessageForEmployee(m);
                              const isSent = !!sentAlerts[m.name];
                              return (
                                <div 
                                  key={m.name} 
                                  className={`p-4 rounded-2xl border transition-all flex flex-col justify-between ${
                                    isSent 
                                      ? 'bg-emerald-50/50 border-emerald-100 opacity-75' 
                                      : 'bg-slate-50/80 border-slate-200 hover:border-teal-300 shadow-xs'
                                  }`}
                                >
                                  <div>
                                    <div className="flex items-start justify-between gap-2 mb-2 pb-2 border-b border-slate-200/60">
                                      <div>
                                        <h4 className="font-extrabold text-slate-800 text-sm truncate max-w-[160px]">{m.name}</h4>
                                        <p className="text-[10px] text-slate-400">{m.job || 'منسوب المدرسة'}</p>
                                      </div>
                                      <div className="flex flex-col items-end gap-1">
                                        <span className={`text-[10px] font-bold px-2 py-0.5 rounded-full ${
                                          isSent 
                                            ? 'bg-emerald-100 text-emerald-800 border border-emerald-200' 
                                            : 'bg-rose-100 text-rose-800 border border-rose-200'
                                        }`}>
                                          {isSent ? '✓ تم الإرسال' : '⚠️ معلق'}
                                        </span>
                                      </div>
                                    </div>

                                    {/* Stats badging */}
                                    <div className="flex flex-wrap gap-1 mb-3">
                                      {(m.lateCount || 0) >= 3 && (
                                        <span className="bg-amber-50 text-amber-700 text-[10px] px-2 py-0.5 rounded-md border border-amber-200 font-bold">
                                          ⏱️ تأخر: {m.lateCount}
                                        </span>
                                      )}
                                      {(m.earlyCount || 0) >= 5 && (
                                        <span className="bg-orange-50 text-orange-700 text-[10px] px-2 py-0.5 rounded-md border border-orange-200 font-bold">
                                          🚪 خروج مبكر: {m.earlyCount}
                                        </span>
                                      )}
                                      {((m.autoCount || 0) + (m.missingCheckoutCount || 0)) >= 3 && (
                                        <span className="bg-rose-50 text-rose-700 text-[10px] px-2 py-0.5 rounded-md border border-rose-200 font-bold">
                                          🔄 تلقائي: {(m.autoCount || 0) + (m.missingCheckoutCount || 0)}
                                        </span>
                                      )}
                                    </div>

                                    {/* Message preview block */}
                                    <div className="bg-white border border-slate-200/80 rounded-xl p-2.5 text-xs text-slate-600 font-sans leading-relaxed max-h-[100px] overflow-y-auto mb-3 text-right" dir="rtl">
                                      {msgText.split('\n').map((line, idx) => (
                                        <p key={idx} className="min-h-[1.2em]">{line}</p>
                                      ))}
                                    </div>
                                  </div>

                                  <div className="flex items-center gap-2 pt-1 border-t border-slate-200/50 mt-2">
                                    <button
                                      onClick={() => {
                                        const encoded = encodeURIComponent(msgText);
                                        window.open(`https://wa.me/?text=${encoded}`, '_blank');
                                        setSentAlerts(prev => ({ ...prev, [m.name]: true }));
                                        
                                        // log to messages
                                        const matchedRecord = dailyRecords.find(r => r.name === m.name);
                                        if (matchedRecord) {
                                          logMessage(matchedRecord, 'تنبيه تلقائي مباشر');
                                        }
                                      }}
                                      className={`flex-1 flex items-center justify-center gap-1.5 py-2 px-3 rounded-xl text-xs font-bold transition-all ${
                                        isSent
                                          ? 'bg-emerald-600 hover:bg-emerald-700 text-white shadow-sm'
                                          : 'bg-teal-600 hover:bg-teal-700 text-white shadow-md hover:shadow-lg'
                                      }`}
                                    >
                                      <MessageSquare className="w-3.5 h-3.5" />
                                      <span>إرسال واتساب 💬</span>
                                    </button>

                                    <button
                                      onClick={() => {
                                        navigator.clipboard.writeText(msgText);
                                        alert(`تم نسخ رسالة ${m.name} بنجاح.`);
                                      }}
                                      className="p-2 bg-slate-100 hover:bg-slate-200 text-slate-600 hover:text-slate-800 rounded-xl border border-slate-200 transition-all"
                                      title="نسخ نص الرسالة"
                                    >
                                      <Copy className="w-3.5 h-3.5" />
                                    </button>

                                    <button
                                      onClick={() => {
                                        setSentAlerts(prev => ({ ...prev, [m.name]: !isSent }));
                                      }}
                                      className={`p-2 rounded-xl border transition-all ${
                                        isSent 
                                          ? 'bg-amber-50 border-amber-200 text-amber-600 hover:bg-amber-100' 
                                          : 'bg-emerald-50 border-emerald-200 text-emerald-600 hover:bg-emerald-100'
                                      }`}
                                      title={isSent ? "تعليم كمعلق / غير مرسل" : "تعليم كمرسل"}
                                    >
                                      <Check className="w-3.5 h-3.5" />
                                    </button>
                                  </div>
                                </div>
                              );
                            })}
                          </div>
                        )}
                      </div>
                    </div>
                  )}

                  {/* BENTO GRID INDEX CLASSIFICATIONS */}
                  <div>
                    <div className="flex flex-wrap items-center justify-between gap-3 mb-3">
                      <h3 className="text-lg font-bold text-slate-700">حالة تصنيف الهيئة التعليمية والإدارية</h3>
                      <button
                        type="button"
                        onClick={() => {
                          const updated = !config.enableAutoAlert;
                          const newConfig = { ...config, enableAutoAlert: updated };
                          setConfig(newConfig);
                          localStorage.setItem(LOCAL_STORAGE_KEYS.CONFIG, JSON.stringify(newConfig));
                        }}
                        className={`inline-flex items-center gap-2 px-3.5 py-1.5 rounded-full text-xs font-black transition-all shadow-xs border cursor-pointer ${
                          config.enableAutoAlert
                            ? 'bg-teal-700 hover:bg-teal-800 text-white border-teal-800'
                            : 'bg-amber-50 hover:bg-amber-100 text-amber-900 border-amber-300'
                        }`}
                      >
                        <MessageSquare className="w-3.5 h-3.5" />
                        <span>
                          {config.enableAutoAlert
                            ? '📢 لوحة التنبيهات التلقائية المجهزة (مفعلة ✓)'
                            : '📢 إظهار التنبيهات التلقائية المجهزة (اضغط للتفعيل)'}
                        </span>
                      </button>
                    </div>
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                      {/* Excellent (Green) */}
                      <div className="bg-white border-t-4 border-emerald-600 rounded-2xl p-6 shadow-sm flex flex-col justify-between min-h-[340px]">
                        <div>
                          <div className="flex justify-between items-start">
                            <div>
                              <h4 className="font-bold text-emerald-800 text-sm">🏆 المنضبطون المتميزون</h4>
                              <p className="text-[10px] text-slate-400 mt-1">
                                تأخر ≤ يومين، انصراف مبكر ≤ يوم واحد، لا يوجد انصراف تلقائي ولا غياب بدون عذر
                              </p>
                            </div>
                            <span className="text-2xl font-black text-emerald-600">{stats.excellent.length}</span>
                          </div>
                          <div className="grid grid-cols-2 gap-2 mt-4 text-[11px] bg-emerald-50/50 p-2.5 rounded-xl border border-emerald-100">
                            <div>
                              <span className="block text-slate-400 font-medium">النسبة من الإجمالي</span>
                              <strong className="text-emerald-700 font-extrabold">
                                {stats.totalCount ? Math.round((stats.excellent.length / stats.totalCount) * 100) : 0}%
                              </strong>
                            </div>
                            <div>
                              <span className="block text-slate-400 font-medium">مجموع أيام الحضور</span>
                              <strong className="text-emerald-700 font-extrabold">
                                {stats.excellent.reduce((sum, m) => sum + (m.present || 0), 0)} يوم
                              </strong>
                            </div>
                          </div>
                        </div>
                        <div className="mt-4 flex-1 overflow-y-auto max-h-[160px] space-y-2 pr-1">
                          {stats.excellent.length === 0 ? (
                            <p className="text-xs text-slate-400 text-center my-4">لا يوجد منسوبون في هذه الفئة حالياً</p>
                          ) : (
                            stats.excellent.map(m => (
                              <div key={m.name} className="flex items-center justify-between text-xs bg-slate-50 p-2 rounded-lg border border-slate-200">
                                <span className="font-bold text-slate-700 truncate max-w-[200px]" title={m.name}>{m.name}</span>
                                <div className="flex items-center gap-1.5 flex-shrink-0">
                                  <button
                                    onClick={() => handleQuickWhatsApp(m)}
                                    className="p-1 hover:bg-slate-200 rounded text-teal-600"
                                    title="إرسال رسالة شكر"
                                  >
                                    <MessageSquare className="w-3.5 h-3.5" />
                                  </button>
                                </div>
                              </div>
                            ))
                          )}
                        </div>
                      </div>

                      {/* Needs Support (Red) */}
                      <div className="bg-white border-t-4 border-rose-500 rounded-2xl p-6 shadow-sm flex flex-col justify-between min-h-[340px]">
                        <div>
                          <div className="flex justify-between items-start">
                            <div>
                              <h4 className="font-bold text-rose-800 text-sm">⚠️ يحتاج متابعة ودعم</h4>
                              <p className="text-[10px] text-slate-400 mt-1">
                                تأخر ≥ 3 أيام، انصراف مبكر ≥ 5 أيام، انصراف تلقائي ≥ 3 أيام
                              </p>
                            </div>
                            <span className="text-2xl font-black text-rose-600">{stats.needsSupport.length}</span>
                          </div>
                          <div className="grid grid-cols-2 gap-2 mt-4 text-[11px] bg-rose-50/50 p-2.5 rounded-xl border border-rose-100">
                            <div>
                              <span className="block text-slate-400 font-medium">النسبة من الإجمالي</span>
                              <strong className="text-rose-700 font-extrabold">
                                {stats.totalCount ? Math.round((stats.needsSupport.length / stats.totalCount) * 100) : 0}%
                              </strong>
                            </div>
                            <div>
                              <span className="block text-slate-400 font-medium">إجمالي أيام التأخر</span>
                              <strong className="text-rose-700 font-extrabold">
                                {stats.needsSupport.reduce((sum, m) => sum + (m.lateCount || 0), 0)} يوم
                              </strong>
                            </div>
                          </div>
                        </div>
                        <div className="mt-4 flex-1 overflow-y-auto max-h-[160px] space-y-2 pr-1">
                          {stats.needsSupport.length === 0 ? (
                            <p className="text-xs text-slate-400 text-center my-4">لا يوجد منسوبون في هذه الفئة حالياً</p>
                          ) : (
                            stats.needsSupport.map(m => (
                              <div key={m.name} className="flex flex-col gap-1 text-xs bg-slate-50 p-2.5 rounded-lg border border-slate-200">
                                <div className="flex items-center justify-between">
                                  <span className="font-bold text-slate-700 truncate max-w-[180px]" title={m.name}>{m.name}</span>
                                  <button
                                    onClick={() => handleQuickWhatsApp(m)}
                                    className="p-1 hover:bg-slate-200 rounded text-rose-600 animate-pulse"
                                    title="إرسال رسالة تنبيه ودعم"
                                  >
                                    <MessageSquare className="w-3.5 h-3.5" />
                                  </button>
                                </div>
                                <div className="flex flex-wrap gap-1 mt-0.5">
                                  {(m.lateCount || 0) >= 3 && (
                                    <span className="bg-amber-50 text-amber-700 text-[9px] px-1.5 py-0.5 rounded border border-amber-200 font-semibold">
                                      ⏱️ تأخر: {m.lateCount}
                                    </span>
                                  )}
                                  {(m.earlyCount || 0) >= 5 && (
                                    <span className="bg-orange-50 text-orange-700 text-[9px] px-1.5 py-0.5 rounded border border-orange-200 font-semibold">
                                      🚪 خروج مبكر: {m.earlyCount}
                                    </span>
                                  )}
                                  {((m.autoCount || 0) + (m.missingCheckoutCount || 0)) >= 3 && (
                                    <span className="bg-rose-50 text-rose-700 text-[9px] px-1.5 py-0.5 rounded border border-rose-200 font-semibold">
                                      🔄 تلقائي: {(m.autoCount || 0) + (m.missingCheckoutCount || 0)}
                                    </span>
                                  )}
                                </div>
                              </div>
                            ))
                          )}
                        </div>
                      </div>
                    </div>
                  </div>

                  {/* TOP 10 RANKING CHART */}
                  <div className="bg-white border border-slate-200 rounded-2xl p-6 shadow-sm">
                    <h3 className="text-lg font-extrabold text-teal-950 mb-4 flex items-center gap-2">
                      <span>🏆 قائمة الـ 10 الأكثر انضباطاً بالمدرسة</span>
                      <span className="text-xs text-slate-400 font-medium">(مرتبة تنازلياً حسب المؤشر والأقل تأخراً)</span>
                    </h3>
                    <div className="overflow-x-auto">
                      <table className="w-full text-right text-sm border-collapse">
                        <thead>
                          <tr className="bg-slate-50 text-slate-600 border-b border-slate-200">
                            <th className="p-3 font-bold">الترتيب</th>
                            <th className="p-3 font-bold">المعلم / الموظف</th>
                            <th className="p-3 font-bold">المسمى الوظيفي</th>
                            <th className="p-3 font-bold text-center">الغياب</th>
                            <th className="p-3 font-bold text-center">دقائق التأخر</th>
                            <th className="p-3 font-bold text-center">الخروج المبكر</th>
                            <th className="p-3 font-bold text-center">الإجراء المباشر</th>
                          </tr>
                        </thead>
                        <tbody>
                          {(() => {
                            const disciplinedList = [...getAggregatedActiveSummaries()]
                              .filter(m => (m.rate || 0) >= 80 && m.disciplineCategory !== 'follow' && m.disciplineCategory !== 'admin')
                              .sort((a, b) => (b.rate || 0) - (a.rate || 0) || a.lateMins - b.lateMins || a.earlyCount - b.earlyCount)
                              .slice(0, 10);

                            if (disciplinedList.length === 0) {
                              return (
                                <tr>
                                  <td colSpan={7} className="p-8 text-center text-slate-400 font-bold">
                                    لا يوجد معلمون أو موظفون مؤهلون لقائمة الأكثر انضباطاً حالياً (درجة الانضباط لجميع المنسوبين أقل من 80٪)
                                  </td>
                                </tr>
                              );
                            }

                            return disciplinedList.map((m, index) => (
                              <tr key={m.name} className="border-b border-slate-100 hover:bg-slate-50/50 transition-all">
                                <td className="p-3 font-bold text-slate-500">
                                  <span className={`inline-flex items-center justify-center w-6 h-6 rounded-full font-black text-xs ${
                                    index === 0 ? 'bg-amber-100 text-amber-700' :
                                    index === 1 ? 'bg-slate-100 text-slate-700' :
                                    index === 2 ? 'bg-orange-100 text-orange-700' :
                                    'text-slate-500'
                                  }`}>
                                    {index + 1}
                                  </span>
                                </td>
                                <td className="p-3 font-bold text-slate-800">{m.name}</td>
                                <td className="p-3 text-slate-500 text-xs">{m.job || '-'}</td>
                                <td className="p-3 text-center text-slate-600 font-bold">{m.absence} ي</td>
                                <td className="p-3 text-center text-slate-600 font-bold">{m.lateMins} د</td>
                                <td className="p-3 text-center text-slate-600 font-bold">{m.earlyMins} د</td>
                                <td className="p-3 text-center">
                                  <button
                                    onClick={() => handleQuickWhatsApp(m)}
                                    className="bg-teal-50 hover:bg-teal-100 text-teal-800 font-bold text-xs px-2.5 py-1.5 rounded-lg transition-all inline-flex items-center gap-1"
                                  >
                                    <Share2 className="w-3 h-3" />
                                    <span>مراسلة شكر</span>
                                  </button>
                                </td>
                              </tr>
                            ));
                          })()}
                        </tbody>
                      </table>
                    </div>
                  </div>

                  {/* MASTER TABULAR DATA */}
                  <div className="bg-white border border-slate-200 rounded-2xl p-6 shadow-sm">
                    <h3 className="text-lg font-extrabold text-teal-950 mb-4">سجل المتابعة والانتظام الإحصائي لجميع الموظفين</h3>
                    <div className="overflow-x-auto">
                      <table className="w-full text-right text-sm border-collapse">
                        <thead>
                          <tr className="bg-slate-50 text-slate-600 border-b border-slate-200">
                            <th className="p-3 font-bold">اسم الموظف</th>
                            <th className="p-3 font-bold">الوظيفة</th>
                            <th className="p-3 font-bold text-center">الفترة</th>
                            <th className="p-3 font-bold text-center">أيام العمل</th>
                            <th className="p-3 font-bold text-center">الحضور</th>
                            <th className="p-3 font-bold text-center">غياب غير مبرر</th>
                            <th className="p-3 font-bold text-center">تأخر صباحي</th>
                            <th className="p-3 font-bold text-center">مجموع دقائق التأخر</th>
                            <th className="p-3 font-bold text-center">مراجعة</th>
                          </tr>
                        </thead>
                        <tbody>
                          {filterMonthlySummaries(monthlySummaries).map(m => (
                            <tr
                              key={`${m.name}-${m.month}`}
                              onClick={() => {
                                setMonthlyEmployee(m.name);
                                setSelectedEmployee(m.name);
                                if (selectedPeriodType === 'range') {
                                  setMonthlyPeriodType('range');
                                  setMonthlyStartDate(selectedStartDate);
                                  setMonthlyEndDate(selectedEndDate);
                                  setMonthlyMonth('all');
                                } else {
                                  setMonthlyPeriodType('month');
                                  setMonthlyMonth(m.month);
                                }
                                setCurrentTab('monthly');
                              }}
                              className="border-b border-slate-100 hover:bg-slate-50 transition-all cursor-pointer"
                            >
                              <td className="p-3 font-bold text-teal-950">{m.name}</td>
                              <td className="p-3 text-slate-500 text-xs">{m.job || '-'}</td>
                              <td className="p-3 text-center text-slate-500 text-xs font-semibold">{m.month}</td>
                              <td className="p-3 text-center text-slate-700 font-bold">{m.work}</td>
                              <td className="p-3 text-center text-emerald-600 font-bold">{m.present}</td>
                              <td className="p-3 text-center text-rose-600 font-bold">{m.absence}</td>
                              <td className="p-3 text-center text-amber-600 font-bold">{m.lateCount}</td>
                              <td className="p-3 text-center text-slate-600 font-bold">{m.lateMins}</td>
                              <td className="p-3 text-center text-teal-600 text-xs font-bold hover:underline">
                                عرض التقرير ←
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                </div>
              )}

              {/* B. ATTENDANCE TAB */}
              {currentTab === 'attendance' && (
                <div className="space-y-6">
                  {/* FILTERS PANEL */}
                  <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-xs flex flex-wrap items-center justify-between gap-4">
                    <div className="flex flex-wrap items-center gap-3">
                      <div>
                        <label className="block text-[10px] font-bold text-slate-400 mb-1">تصفية الفترة حسب</label>
                        <select
                          value={selectedPeriodType}
                          onChange={(e) => setSelectedPeriodType(e.target.value as 'month' | 'range')}
                          className="bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-sm font-bold text-slate-700 outline-none focus:border-teal-500 transition-all"
                        >
                          <option value="month">📅 الشهر</option>
                          <option value="range">🗓️ نطاق تاريخ</option>
                        </select>
                      </div>

                      {selectedPeriodType === 'month' ? (
                        <div>
                          <label className="block text-[10px] font-bold text-slate-400 mb-1">
                            {selectedMonth === 'all' ? 'الفترة المستهدفة' : 'الشهر المستهدف'}
                          </label>
                          <select
                            value={selectedMonth}
                            onChange={(e) => setSelectedMonth(e.target.value)}
                            className="bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-sm font-bold text-slate-700 outline-none focus:border-teal-500 transition-all"
                          >
                            <option value="all">كل الأشهر المسجلة</option>
                            {getMonthsList().map(m => (
                              <option key={m} value={m}>{m}</option>
                            ))}
                          </select>
                        </div>
                      ) : (
                        <>
                          <div>
                            <label className="block text-[10px] font-bold text-slate-400 mb-1">من تاريخ</label>
                            <input
                              type="date"
                              value={selectedStartDate}
                              onChange={(e) => setSelectedStartDate(e.target.value)}
                              className="bg-slate-50 border border-slate-200 rounded-xl px-3 py-1.5 text-sm font-bold text-slate-700 outline-none focus:border-teal-500 transition-all h-[38px]"
                            />
                          </div>
                          <div>
                            <label className="block text-[10px] font-bold text-slate-400 mb-1">إلى تاريخ</label>
                            <input
                              type="date"
                              value={selectedEndDate}
                              onChange={(e) => setSelectedEndDate(e.target.value)}
                              className="bg-slate-50 border border-slate-200 rounded-xl px-3 py-1.5 text-sm font-bold text-slate-700 outline-none focus:border-teal-500 transition-all h-[38px]"
                            />
                          </div>
                        </>
                      )}

                      <div>
                        <label className="block text-[10px] font-bold text-slate-400 mb-1">الموظف / المعلم</label>
                        <select
                          value={selectedEmployee}
                          onChange={(e) => setSelectedEmployee(e.target.value)}
                          className="bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-sm font-bold text-slate-700 outline-none focus:border-teal-500 transition-all max-w-[240px]"
                        >
                          <option value="all">جميع المعلمين والموظفين</option>
                          {getEmployeesNames().map(name => (
                            <option key={name} value={name}>{name}</option>
                          ))}
                        </select>
                      </div>

                      <div>
                        <label className="block text-[10px] font-bold text-slate-400 mb-1">تصفية حسب حالة الانصراف</label>
                        <select
                          value={selectedStatus}
                          onChange={(e) => setSelectedStatus(e.target.value)}
                          className="bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-sm font-bold text-slate-700 outline-none focus:border-teal-500 transition-all"
                        >
                          <option value="all">كل حالات الانصراف اليومية</option>
                          <option value="early">خروج قبل نهاية الدوام (انصراف مبكر)</option>
                          <option value="auto">انصراف تلقائي من النظام</option>
                          <option value="missingCheckout">غياب بصمة الانصراف اليومية</option>
                          <option value="emergency_no_exit">⚠️ استئذان طارئ بدون بصمة انصراف (مخالف)</option>
                          <option value="emergency_with_exit">✓ استئذان طارئ موثق ببصمة انصراف</option>
                          <option value="wrong_time">انصراف في غير وقت الفعلي (بصمة صباحية/اليوم التالي)</option>
                        </select>
                      </div>
                    </div>

                    <div className="flex flex-wrap items-end gap-3 w-full lg:w-auto">
                      <div className="w-full md:w-80">
                        <label className="block text-[10px] font-bold text-slate-400 mb-1">البحث السريع بالاسم</label>
                        <input
                          type="text"
                          placeholder="ابحث عن اسم موظف..."
                          value={searchTerm}
                          onChange={(e) => setSearchTerm(e.target.value)}
                          className="bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-sm w-full outline-none font-medium"
                        />
                      </div>
                      <div className="flex items-center gap-2">
                        <button
                          onClick={handlePrintAttendanceList}
                          className="bg-teal-600 hover:bg-teal-700 text-white font-bold px-4 py-2 rounded-xl text-xs flex items-center gap-1.5 transition-all shadow-sm h-[38px]"
                          title="طباعة السجل الحالي بصيغة PDF"
                        >
                          <Printer className="w-4 h-4" />
                          <span>طباعة</span>
                        </button>
                      </div>
                    </div>
                  </div>

                  <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
                    <div className="overflow-x-auto">
                      <table className="w-full text-right border-collapse">
                        <thead>
                          <tr className="bg-slate-50 text-slate-700 text-xs border-b border-slate-200">
                            <th className="p-3 font-bold text-right">المعلم / الموظف</th>
                            <th className="p-3 font-bold text-right">الوظيفة</th>
                            <th className="p-3 font-bold text-center">التاريخ</th>
                            <th className="p-3 font-bold text-center">اليوم</th>
                            <th className="p-3 font-bold text-center">وقت الدخول</th>
                            <th className="p-3 font-bold text-center">وقت الانصراف</th>
                            <th className="p-3 font-bold text-center">حالة الانصراف</th>
                            <th className="p-3 font-bold text-center">الإجراء</th>
                          </tr>
                        </thead>
                        <tbody>
                          {filterDailyRecords(dailyRecords)
                            .filter(r => r.type === 'present' || (selectedStatus === 'emergency_no_exit' && r.checkout === 'استئذان بدون بصمة انصراف') || (selectedStatus === 'emergency_with_exit' && r.checkout === 'استئذان طارئ موثق (بصمة)'))
                            .sort((a, b) => new Date(a.date).getTime() - new Date(b.date).getTime())
                            .map((r, i) => (
                              <tr key={i} className="border-b border-slate-100 hover:bg-slate-50/50 transition-all">
                                <td className="p-3 font-bold text-slate-800">{r.name}</td>
                                <td className="p-3 text-slate-400 text-xs">{r.job || '-'}</td>
                                <td className="p-3 text-center">{renderDateCell(r.date)}</td>
                                <td className="p-3 text-center text-slate-500 text-xs font-semibold">{dayName(r.date)}</td>
                                <td className="p-3 text-center font-mono font-bold text-emerald-700">{r.tin?.text || '-'}</td>
                                <td className="p-3 text-center font-mono font-bold text-slate-700">{r.tout?.text || '-'}</td>
                                <td className="p-3 text-center">
                                  <span className={`text-[10px] font-bold px-2.5 py-1 rounded-full ${
                                    r.checkout === 'مكتمل' ? 'bg-emerald-100 text-emerald-800' :
                                    r.checkout === 'انصراف تلقائي' ? 'bg-amber-100 text-amber-800' :
                                    (r.checkout === 'خروج قبل نهاية الدوام' || r.checkout === 'انصراف مبكر') ? 'bg-rose-100 text-rose-800' :
                                    r.checkout === 'انصراف في غير وقت الفعلي' ? 'bg-purple-100 text-purple-800 border border-purple-300 font-extrabold shadow-xs' :
                                    r.checkout === 'استئذان بدون بصمة انصراف' ? 'bg-amber-100 text-amber-950 border border-amber-300 font-black shadow-2xs' :
                                    r.checkout === 'استئذان طارئ موثق (بصمة)' ? 'bg-sky-100 text-sky-800 border border-sky-300 font-bold' :
                                    (r.checkout && !['لا توجد بصمة انصراف', 'غير متاح'].includes(r.checkout)) ? 'bg-blue-100 text-blue-800' :
                                    'bg-rose-50 text-rose-600 border border-rose-200'
                                  }`}>
                                    {r.checkout === 'انصراف في غير وقت الفعلي' ? 'انصراف في غير وقته الفعلي (بصمة صباحية)' :
                                     r.checkout === 'استئذان بدون بصمة انصراف' ? '⚠️ استئذان طارئ بدون بصمة (مخالف)' :
                                     r.checkout === 'استئذان طارئ موثق (بصمة)' ? '✓ استئذان طارئ موثق (ببصمة)' :
                                     (r.checkout || 'غير متاح')}
                                    {(r.checkout === 'خروج قبل نهاية الدوام' || r.checkout === 'انصراف مبكر') && r.earlyMins > 0 && ` (${r.earlyMins} دقيقة)`}
                                  </span>
                                </td>
                                <td className="p-3 text-center">
                                  <div className="flex items-center justify-center gap-1.5">
                                    <button
                                      onClick={() => handlePrintEmployeeReport(r.name)}
                                      className="text-teal-700 bg-teal-50 hover:bg-teal-100 p-2 rounded-lg text-xs transition-all flex items-center gap-1 font-bold"
                                      title="طباعة تقرير المعلم PDF"
                                    >
                                      <Printer className="w-3.5 h-3.5" />
                                      <span className="hidden sm:inline">PDF</span>
                                    </button>
                                    <button
                                      onClick={() => handleExportEmployeeWord(r.name)}
                                      className="text-indigo-700 bg-indigo-50 hover:bg-indigo-100 p-2 rounded-lg text-xs transition-all flex items-center gap-1 font-bold"
                                      title="تصدير تقرير المعلم Word"
                                    >
                                      <FileText className="w-3.5 h-3.5" />
                                      <span className="hidden sm:inline">Word</span>
                                    </button>
                                    <button
                                      onClick={() => handleSaveEditedRecord ? setEditingRecord(r) : null}
                                      className="text-amber-700 bg-amber-50 hover:bg-amber-100 p-2 rounded-lg text-xs transition-all flex items-center gap-1 font-bold"
                                      title="تعديل وقت البصمة وحالة الحضور"
                                    >
                                      <Edit className="w-3.5 h-3.5" />
                                      <span className="hidden sm:inline">تعديل</span>
                                    </button>
                                    {r.checkout !== 'مكتمل' && (
                                      <button
                                        onClick={() => handleOpenMessageModal(`${r.name}|${r.dateKey}`)}
                                        className={`font-bold p-2 rounded-lg text-xs transition-all flex items-center gap-1 ${
                                          r.checkout === 'استئذان بدون بصمة انصراف'
                                            ? 'text-amber-900 bg-amber-100 hover:bg-amber-200 border border-amber-300'
                                            : 'text-rose-700 bg-rose-50 hover:bg-rose-100'
                                        }`}
                                        title={r.checkout === 'استئذان بدون بصمة انصراف' ? 'إرسال تنبيه بالالتزام ببصمة الانصراف عند الاستئذان' : 'إرسال تنبيه خروج'}
                                      >
                                        <MessageSquare className="w-3.5 h-3.5" />
                                        <span>تنبيه</span>
                                      </button>
                                    )}
                                  </div>
                                </td>
                              </tr>
                            ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                </div>
              )}

              {/* C. LATE TAB */}
              {currentTab === 'late' && (
                <div className="space-y-6">
                  {/* METRIC CHIPS FOR LATE TAB */}
                  <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                    <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-xs">
                      <span className="text-xs text-slate-400 font-bold">إجمالي مرات التأخر</span>
                      <p className="text-2xl font-black text-amber-600 mt-1">{stats.totalLateCount} مرة</p>
                    </div>
                    <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-xs">
                      <span className="text-xs text-slate-400 font-bold">إجمالي دقائق التأخر</span>
                      <p className="text-2xl font-black text-slate-800 mt-1">{stats.totalLateMins} دقيقة</p>
                    </div>
                    <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-xs text-slate-600 font-medium text-xs leading-normal flex items-center">
                      تعتمد عمليات حساب التأخر على وقت بداية الدوام المعتمد في الإعدادات ({config.start}). أي دقيقة بعد هذا الوقت تخصم 0.25 درجة من التقييم الشهري للموظف.
                    </div>
                  </div>

                  {/* FILTERS PANEL */}
                  <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-xs flex flex-wrap items-center justify-between gap-4">
                    <div className="flex flex-wrap items-center gap-3">
                      <div>
                        <label className="block text-[10px] font-bold text-slate-400 mb-1">
                          {selectedMonth === 'all' ? 'الفترة المستهدفة' : 'الشهر المستهدف'}
                        </label>
                        <select
                          value={selectedMonth}
                          onChange={(e) => setSelectedMonth(e.target.value)}
                          className="bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-sm font-bold text-slate-700 outline-none focus:border-teal-500 transition-all"
                        >
                          <option value="all">كل الأشهر المسجلة</option>
                          {getMonthsList().map(m => (
                            <option key={m} value={m}>{m}</option>
                          ))}
                        </select>
                      </div>

                      <div>
                        <label className="block text-[10px] font-bold text-slate-400 mb-1">الموظف / المعلم</label>
                        <select
                          value={selectedEmployee}
                          onChange={(e) => setSelectedEmployee(e.target.value)}
                          className="bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-sm font-bold text-slate-700 outline-none focus:border-teal-500 transition-all max-w-[240px]"
                        >
                          <option value="all">جميع المعلمين والموظفين</option>
                          {getEmployeesNames().map(name => (
                            <option key={name} value={name}>{name}</option>
                          ))}
                        </select>
                      </div>
                    </div>

                    <div className="flex flex-wrap items-center gap-3">
                      <div className="w-full sm:w-64">
                        <label className="block text-[10px] font-bold text-slate-400 mb-1">البحث السريع بالاسم</label>
                        <input
                          type="text"
                          placeholder="ابحث عن اسم موظف..."
                          value={searchTerm}
                          onChange={(e) => setSearchTerm(e.target.value)}
                          className="bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-sm w-full outline-none focus:border-amber-500 transition-all font-medium"
                        />
                      </div>
                      <div className="flex items-end gap-2">
                        <button
                          onClick={handlePrintLateList}
                          className="bg-amber-600 hover:bg-amber-700 text-white font-bold px-4 py-2 rounded-xl text-xs flex items-center gap-1.5 transition-all shadow-xs h-[38px]"
                          title="طباعة كشف التأخر الصباحي بصيغة PDF"
                        >
                          <Printer className="w-4 h-4" />
                          <span>طباعة PDF</span>
                        </button>
                        <button
                          onClick={handleExportLateListWord}
                          className="bg-indigo-600 hover:bg-indigo-700 text-white font-bold px-3 py-2 rounded-xl text-xs flex items-center gap-1.5 transition-all shadow-xs h-[38px]"
                          title="تصدير كشف التأخر الصباحي Word"
                        >
                          <FileDown className="w-4 h-4" />
                          <span>Word</span>
                        </button>
                      </div>
                    </div>
                  </div>

                  {/* LATE TABLE */}
                  <div className="bg-white border border-slate-200 rounded-2xl overflow-hidden shadow-sm">
                    <div className="overflow-x-auto">
                      <table className="w-full text-right text-sm border-collapse">
                        <thead>
                          <tr className="bg-slate-50 text-slate-600 border-b border-slate-200">
                            <th className="p-3 font-bold">اسم المعلم / الموظف</th>
                            <th className="p-3 font-bold">المسمى الوظيفي</th>
                            <th className="p-3 font-bold text-center">التاريخ</th>
                            <th className="p-3 font-bold text-center">اليوم</th>
                            <th className="p-3 font-bold text-center">وقت الحضور</th>
                            <th className="p-3 font-bold text-center">مدة التأخر</th>
                            <th className="p-3 font-bold text-center">الإجراء المتاح</th>
                          </tr>
                        </thead>
                        <tbody>
                          {filterDailyRecords(dailyRecords)
                            .filter(r => r.late > 0)
                            .sort((a, b) => new Date(a.date).getTime() - new Date(b.date).getTime())
                            .map((r, i) => (
                              <tr key={i} className="border-b border-slate-100 hover:bg-slate-50/50 transition-all">
                                <td className="p-3 font-bold text-slate-800">{r.name}</td>
                                <td className="p-3 text-slate-400 text-xs">{r.job || '-'}</td>
                                <td className="p-3 text-center">{renderDateCell(r.date)}</td>
                                <td className="p-3 text-center text-slate-500 text-xs font-semibold">{dayName(r.date)}</td>
                                <td className="p-3 text-center font-mono font-bold text-rose-600">{r.tin?.text || '-'}</td>
                                <td className="p-3 text-center">
                                  <span className="bg-amber-100 text-amber-800 font-bold text-xs px-2 py-1 rounded-full">
                                    {r.late} دقيقة
                                  </span>
                                </td>
                                <td className="p-3 text-center">
                                  <div className="flex items-center justify-center gap-1.5">
                                    <button
                                      onClick={() => handleOpenMessageModal(`${r.name}|${r.dateKey}`)}
                                      className="text-teal-700 bg-teal-50 hover:bg-teal-100 font-bold px-2.5 py-1.5 rounded-lg text-xs transition-all inline-flex items-center gap-1"
                                    >
                                      <MessageSquare className="w-3.5 h-3.5" />
                                      <span>تنبيه تأخير</span>
                                    </button>
                                    <button
                                      onClick={() => setEditingRecord(r)}
                                      className="text-amber-700 bg-amber-50 hover:bg-amber-100 font-bold px-2.5 py-1.5 rounded-lg text-xs transition-all inline-flex items-center gap-1"
                                      title="تعديل وقت البصمة وحالة الحضور"
                                    >
                                      <Edit className="w-3.5 h-3.5" />
                                      <span>تعديل</span>
                                    </button>
                                  </div>
                                </td>
                              </tr>
                            ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                </div>
              )}

              {/* D. ABSENCE TAB */}
              {currentTab === 'absence' && (
                <div className="space-y-6">
                  {/* ABSENCE INFO NOTE */}
                  <div className="bg-amber-50 border border-amber-200 text-amber-900 rounded-2xl p-4 text-xs leading-relaxed">
                    <strong>💡 رصد ومتابعة الأعذار عبر نظام فارس:</strong> يدرج غياب الموظف بصفة مبدئية كـ <strong>غياب بدون عذر</strong> ويخصم 10 درجات كاملة من تقييمه. في حال قيام الموظف بتقديم عذره الطبي أو الرسمي بنجاح في نظام فارس وموافقة الإدارة المدرسية، يرجى الضغط على زر التفعيل (✓) في عمود <strong>حالة العذر</strong> لتحويل الغياب إلى غياب بعذر واستعادة الدرجات المخصومة تلقائياً.
                  </div>

                  {/* ABSENCE STATISTICS CARDS */}
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4" dir="rtl">
                    {/* CARD 1: ALL TEACHERS ABSENCES */}
                    <div className="bg-gradient-to-br from-rose-50 to-white p-5 rounded-3xl border border-rose-100 shadow-sm flex items-start justify-between gap-4">
                      <div className="space-y-2">
                        <span className="text-[10px] font-bold text-rose-800 bg-rose-100/60 px-2.5 py-1 rounded-full">
                          📊 إحصاءات المدرسة الإجمالية
                        </span>
                        <h4 className="text-sm font-bold text-slate-700 mt-2">إجمالي الغياب بدون عذر (لكافة المعلمين)</h4>
                        <div className="flex items-baseline gap-2">
                          <span className="text-4xl font-black text-rose-600 tracking-tight">
                            {allSchoolAbsences.length}
                          </span>
                          <span className="text-xs font-bold text-slate-400">أيام غياب معلقة</span>
                        </div>
                        <p className="text-xs text-slate-500 font-medium">
                          عدد المعلمين الغائبين على الأقل مرة واحدة: <strong className="font-bold text-slate-700">{uniqueSchoolAbsentNames.length}</strong> معلم
                        </p>
                      </div>
                      <div className="bg-rose-500/10 p-3.5 rounded-2xl text-rose-600">
                        <CalendarX className="w-8 h-8" />
                      </div>
                    </div>

                    {/* CARD 2: SELECTED TEACHER ABSENCE */}
                    {selectedEmployee === 'all' ? (
                      <div className="bg-slate-50 p-5 rounded-3xl border border-slate-200/60 shadow-xs flex items-center justify-between gap-4">
                        <div className="space-y-1">
                          <h4 className="text-sm font-extrabold text-slate-700 flex items-center gap-1.5">
                            <span className="text-lg">👤</span>
                            <span>إحصاءات الغياب الفردية</span>
                          </h4>
                          <p className="text-xs text-slate-400 leading-relaxed max-w-[280px]">
                            يرجى اختيار اسم معلم محدد من قائمة التصفية أدناه لعرض سجل وتفاصيل غيابه الفردي ومعدل انضباطه.
                          </p>
                        </div>
                        <div className="bg-slate-200/55 p-3.5 rounded-2xl text-slate-400">
                          <Users className="w-8 h-8" />
                        </div>
                      </div>
                    ) : (
                      <div className="bg-gradient-to-br from-teal-50 to-white p-5 rounded-3xl border border-teal-100 shadow-sm flex items-start justify-between gap-4">
                        <div className="space-y-2">
                          <span className="text-[10px] font-bold text-teal-800 bg-teal-100/60 px-2.5 py-1 rounded-full truncate max-w-[180px] inline-block">
                            👤 المعلم: {selectedEmployee}
                          </span>
                          <h4 className="text-xs font-bold text-slate-400 mt-2">
                            {currentEmpSummaryInAbsence?.job || 'منسوب المدرسة'}
                          </h4>
                          <div className="flex items-baseline gap-2">
                            <span className="text-4xl font-black text-teal-600 tracking-tight">
                              {selectedEmployeeAbsences.length}
                            </span>
                            <span className="text-xs font-bold text-slate-400">أيام غياب</span>
                          </div>
                          <div className="text-[11px] text-slate-500 font-medium flex flex-wrap gap-x-3 gap-y-1">
                            <span>معدل الانضباط: <strong className="font-extrabold text-teal-700">{currentEmpSummaryInAbsence?.rate ?? 100}%</strong></span>
                            <span className="text-slate-300">|</span>
                            <span>التصنيف: <strong className="font-extrabold text-teal-700">{currentEmpSummaryInAbsence?.rating || 'مكتمل'}</strong></span>
                          </div>
                        </div>
                        <div className="bg-teal-500/10 p-3.5 rounded-2xl text-teal-600">
                          <Users className="w-8 h-8" />
                        </div>
                      </div>
                    )}
                  </div>

                  {/* FILTERS PANEL */}
                  <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-xs flex flex-wrap items-center justify-between gap-4">
                    <div className="flex flex-wrap items-center gap-3">
                      <div>
                        <label className="block text-[10px] font-bold text-slate-400 mb-1">
                          {selectedMonth === 'all' ? 'الفترة المستهدفة' : 'الشهر المستهدف'}
                        </label>
                        <select
                          value={selectedMonth}
                          onChange={(e) => setSelectedMonth(e.target.value)}
                          className="bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-sm font-bold text-slate-700 outline-none focus:border-teal-500 transition-all"
                        >
                          <option value="all">كل الأشهر المسجلة</option>
                          {getMonthsList().map(m => (
                            <option key={m} value={m}>{m}</option>
                          ))}
                        </select>
                      </div>

                      <div>
                        <label className="block text-[10px] font-bold text-slate-400 mb-1">الموظف / المعلم</label>
                        <select
                          value={selectedEmployee}
                          onChange={(e) => setSelectedEmployee(e.target.value)}
                          className="bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-sm font-bold text-slate-700 outline-none focus:border-teal-500 transition-all max-w-[240px]"
                        >
                          <option value="all">جميع المعلمين والموظفين</option>
                          {getEmployeesNames().map(name => (
                            <option key={name} value={name}>{name}</option>
                          ))}
                        </select>
                      </div>
                    </div>

                    <div className="flex flex-wrap items-center gap-3">
                      <div className="w-full sm:w-64">
                        <label className="block text-[10px] font-bold text-slate-400 mb-1">البحث السريع بالاسم</label>
                        <input
                          type="text"
                          placeholder="ابحث عن اسم موظف..."
                          value={searchTerm}
                          onChange={(e) => setSearchTerm(e.target.value)}
                          className="bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-sm w-full outline-none focus:border-rose-500 transition-all font-medium"
                        />
                      </div>
                      <div className="flex items-end gap-2">
                        <button
                          onClick={handlePrintAbsenceList}
                          className="bg-rose-600 hover:bg-rose-700 text-white font-bold px-4 py-2 rounded-xl text-xs flex items-center gap-1.5 transition-all shadow-xs h-[38px]"
                          title="طباعة كشف الغياب بدون عذر بصيغة PDF"
                        >
                          <Printer className="w-4 h-4" />
                          <span>طباعة PDF</span>
                        </button>
                        <button
                          onClick={handleExportAbsenceListWord}
                          className="bg-indigo-600 hover:bg-indigo-700 text-white font-bold px-3 py-2 rounded-xl text-xs flex items-center gap-1.5 transition-all shadow-xs h-[38px]"
                          title="تصدير كشف الغياب بدون عذر Word"
                        >
                          <FileDown className="w-4 h-4" />
                          <span>Word</span>
                        </button>
                      </div>
                    </div>
                  </div>

                  {/* AUTO-DETECTED SUSPENDED DAYS BANNER IN ABSENCE TAB */}
                  {autoDetectedSuspendedDays.length > 0 && (
                    <div className="bg-amber-50/90 border border-amber-200 p-4 rounded-2xl flex items-center justify-between gap-4 text-xs shadow-xs" dir="rtl">
                      <div className="flex items-center gap-3">
                        <div className="bg-amber-100 text-amber-800 p-2.5 rounded-xl text-xl flex-shrink-0">
                          🌤️
                        </div>
                        <div>
                          <h4 className="font-extrabold text-amber-950 text-sm">
                            تم استبعاد {autoDetectedSuspendedDays.length} يوم دراسي تلقائياً من الغياب (تعليق دراسة لعدم وجود أي بصمات لجميع المعلمين):
                          </h4>
                          <div className="flex flex-wrap items-center gap-2 mt-1">
                            {autoDetectedSuspendedDays.map(dk => (
                              <span key={dk} className="bg-white border border-amber-300 font-bold px-2.5 py-0.5 rounded-lg text-amber-900 text-xs">
                                {dayName(dk)} {dk}م ({toHijri(dk)})
                              </span>
                            ))}
                          </div>
                          <p className="text-[11px] text-amber-800 mt-1">
                            النظام ذكي: لا يقوم بتسجيل غياب بدون عذر جماعي على معلمي المدرسة في الأيام التي لم يسجل فيها أحد أي بصمة.
                          </p>
                        </div>
                      </div>
                      <button
                        onClick={() => setCurrentTab('settings')}
                        className="bg-amber-200 hover:bg-amber-300 text-amber-950 font-extrabold px-3.5 py-2 rounded-xl text-xs transition-all flex-shrink-0 cursor-pointer"
                      >
                        إدارة أيام التعليق
                      </button>
                    </div>
                  )}

                  {/* ABSENCE LIST */}
                  <div className="bg-white border border-slate-200 rounded-2xl overflow-hidden shadow-sm">
                    <div className="overflow-x-auto">
                      <table className="w-full text-right text-sm border-collapse">
                        <thead>
                          <tr className="bg-slate-50 text-slate-600 border-b border-slate-200">
                            <th className="p-3 font-bold">اسم المعلم / الموظف</th>
                            <th className="p-3 font-bold">المسمى الوظيفي</th>
                            <th className="p-3 font-bold text-center">تاريخ الغياب</th>
                            <th className="p-3 font-bold text-center">اليوم</th>
                            <th className="p-3 font-bold text-center">نوع الغياب</th>
                            <th className="p-3 font-bold text-center">طلب العذر وتنبيه واتساب</th>
                            <th className="p-3 font-bold text-center">موجّه بنظام فارس؟</th>
                            <th className="p-3 font-bold text-center">استثناء اليوم (تعليق)</th>
                            <th className="p-3 font-bold text-center">تعديل الوقت</th>
                          </tr>
                        </thead>
                        <tbody>
                          {filterDailyRecords(dailyRecords)
                            .filter(r => r.type === 'absence')
                            .sort((a, b) => new Date(a.date).getTime() - new Date(b.date).getTime())
                            .map((r, i) => {
                              const rKey = `${r.name}|${r.dateKey}`;
                              return (
                                <tr key={i} className="border-b border-slate-100 hover:bg-rose-50/20 transition-all">
                                  <td className="p-3 font-bold text-rose-950">{r.name}</td>
                                  <td className="p-3 text-slate-400 text-xs">{r.job || '-'}</td>
                                  <td className="p-3 text-center">{renderDateCell(r.date)}</td>
                                  <td className="p-3 text-center text-slate-500 text-xs font-semibold">{dayName(r.date)}</td>
                                  <td className="p-3 text-center">
                                    <span className="bg-rose-100 text-rose-800 font-bold text-[10px] px-2 py-0.5 rounded-full">
                                      غياب بدون عذر
                                    </span>
                                  </td>
                                  <td className="p-3 text-center">
                                    <button
                                      onClick={() => handleOpenMessageModal(rKey)}
                                      className="bg-teal-50 hover:bg-teal-100 text-teal-800 font-bold px-3 py-1.5 rounded-xl text-xs transition-all inline-flex items-center gap-1.5 shadow-xs"
                                    >
                                      <MessageSquare className="w-3.5 h-3.5" />
                                      <span>إشعار طلب مبرر</span>
                                    </button>
                                  </td>
                                  <td className="p-3 text-center">
                                    <button
                                      onClick={() => toggleFaresStatus(rKey)}
                                      className="bg-slate-100 hover:bg-emerald-100 text-slate-400 hover:text-emerald-700 w-9 h-9 rounded-xl flex items-center justify-center border border-slate-200 transition-all mx-auto"
                                      title="اضغط هنا لتأكيد تقديم العذر في نظام فارس"
                                    >
                                      <Check className="w-5 h-5" />
                                    </button>
                                  </td>
                                  <td className="p-3 text-center">
                                    <button
                                      onClick={() => {
                                        if (confirm(`هل ترغب في استثناء يوم (${dayName(r.date)} ${r.dateKey}) كتعليق دراسة / تعليم عن بعد لجميع معلمي المدرسة وإلغاء الغياب؟`)) {
                                          toggleSuspendedDate(r.dateKey);
                                        }
                                      }}
                                      className="bg-amber-50 hover:bg-amber-100 text-amber-800 border border-amber-200 font-bold px-2.5 py-1.5 rounded-xl text-xs transition-all inline-flex items-center gap-1 shadow-2xs"
                                      title="استثناء هذا اليوم بالكامل لجميع المعلمين وإلغاء الغياب"
                                    >
                                      <AlertTriangle className="w-3.5 h-3.5 text-amber-600" />
                                      <span>تعليق دراسة</span>
                                    </button>
                                  </td>
                                  <td className="p-3 text-center">
                                    <button
                                      onClick={() => setEditingRecord(r)}
                                      className="text-amber-700 bg-amber-50 hover:bg-amber-100 font-bold px-2.5 py-1.5 rounded-lg text-xs transition-all inline-flex items-center gap-1"
                                      title="تعديل وقت البصمة وحالة الحضور"
                                    >
                                      <Edit className="w-3.5 h-3.5" />
                                      <span>تعديل</span>
                                    </button>
                                  </td>
                                </tr>
                              );
                            })}
                        </tbody>
                      </table>
                    </div>
                  </div>
                </div>
              )}

              {/* E. EXCUSES & LEAVES TAB */}
              {currentTab === 'excuses' && (
                <div className="space-y-6">
                  {/* EXCUSES NOTE & COMPLIANCE WARNING */}
                  <div className="bg-blue-50 border border-blue-200 text-blue-900 rounded-2xl p-4 text-xs leading-relaxed">
                    يعرض هذا الجدول جميع حالات الغياب والإجازات التي تم إثباتها وتقديم مسوغاتها الرسمية أو الطبية بنجاح، بالإضافة إلى حالات الاستئذان الطارئ ومتابعة مدى التزام المعلمين بتسجيل بصمة الانصراف عند الخروج.
                  </div>

                  {/* EMERGENCY PERMISSION WITHOUT EXIT ALERT BANNER */}
                  {stats.totalEmergencyNoExit > 0 && (
                    <div className="bg-gradient-to-r from-amber-50 to-orange-50 border border-amber-300 p-4 rounded-2xl flex flex-col md:flex-row md:items-center justify-between gap-4 text-xs shadow-xs" dir="rtl">
                      <div className="flex items-start gap-3">
                        <div className="bg-amber-200 text-amber-900 p-2.5 rounded-xl text-xl flex-shrink-0">
                          ⚠️
                        </div>
                        <div>
                          <h4 className="font-extrabold text-amber-950 text-sm">
                            تنبيه انضباطي: تم رصد {stats.totalEmergencyNoExit} حالة استئذان طارئ بدون تسجيل بصمة انصراف عند المغادرة!
                          </h4>
                          <p className="text-[11px] text-amber-800 mt-1 leading-relaxed">
                            نظام العمل يلزم كل موظف أو معلم مستأذن طارئ بتسجيل بصمة الانصراف في جهاز البصمة لحظة الخروج الفعلي لتوثيق وقت المغادرة نظاماً. الخروج بدون بصمة انصراف يعد تصرفاً غير مكتمل ومخالفاً للتعليمات.
                          </p>
                        </div>
                      </div>
                      <button
                        onClick={() => setExcuseSubFilter('emergency_no_exit')}
                        className="bg-amber-300 hover:bg-amber-400 text-amber-950 font-black px-4 py-2.5 rounded-xl text-xs transition-all flex-shrink-0 cursor-pointer shadow-2xs self-start md:self-auto"
                      >
                        تصفية المخالفين ({stats.totalEmergencyNoExit})
                      </button>
                    </div>
                  )}

                  {/* FILTERS PANEL */}
                  <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-xs flex flex-wrap items-center justify-between gap-4">
                    <div className="flex flex-wrap items-center gap-3">
                      <div>
                        <label className="block text-[10px] font-bold text-slate-400 mb-1">
                          {selectedMonth === 'all' ? 'الفترة المستهدفة' : 'الشهر المستهدف'}
                        </label>
                        <select
                          value={selectedMonth}
                          onChange={(e) => setSelectedMonth(e.target.value)}
                          className="bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-sm font-bold text-slate-700 outline-none focus:border-teal-500 transition-all"
                        >
                          <option value="all">كل الأشهر المسجلة</option>
                          {getMonthsList().map(m => (
                            <option key={m} value={m}>{m}</option>
                          ))}
                        </select>
                      </div>

                      <div>
                        <label className="block text-[10px] font-bold text-slate-400 mb-1">الموظف / المعلم</label>
                        <select
                          value={selectedEmployee}
                          onChange={(e) => setSelectedEmployee(e.target.value)}
                          className="bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-sm font-bold text-slate-700 outline-none focus:border-teal-500 transition-all max-w-[240px]"
                        >
                          <option value="all">جميع المعلمين والموظفين</option>
                          {getEmployeesNames().map(name => (
                            <option key={name} value={name}>{name}</option>
                          ))}
                        </select>
                      </div>

                      <div>
                        <label className="block text-[10px] font-bold text-slate-400 mb-1">نوع المسوغ / الاستئذان</label>
                        <select
                          value={excuseSubFilter}
                          onChange={(e) => setExcuseSubFilter(e.target.value as any)}
                          className="bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-sm font-bold text-slate-700 outline-none focus:border-teal-500 transition-all"
                        >
                          <option value="all">كل الإجازات والاستئذانات</option>
                          <option value="leaves">إجازات رسمية وأعذار غياب فقط</option>
                          <option value="emergency_all">جميع الاستئذانات الطارئة</option>
                          <option value="emergency_no_exit">⚠️ استئذان طارئ بدون بصمة انصراف (مخالف)</option>
                          <option value="emergency_with_exit">✓ استئذان طارئ موثق ببصمة خروج</option>
                        </select>
                      </div>
                    </div>

                    <div className="flex flex-wrap items-center gap-3">
                      <div className="w-full sm:w-64">
                        <label className="block text-[10px] font-bold text-slate-400 mb-1">البحث السريع بالاسم</label>
                        <input
                          type="text"
                          placeholder="ابحث عن اسم موظف..."
                          value={searchTerm}
                          onChange={(e) => setSearchTerm(e.target.value)}
                          className="bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-sm w-full outline-none focus:border-teal-500 transition-all font-medium"
                        />
                      </div>
                      <div className="flex items-end gap-2">
                        <button
                          onClick={handlePrintExcusesList}
                          className="bg-sky-600 hover:bg-sky-700 text-white font-bold px-4 py-2 rounded-xl text-xs flex items-center gap-1.5 transition-all shadow-xs h-[38px]"
                          title="طباعة كشف الغياب بعذر والأعذار بصيغة PDF"
                        >
                          <Printer className="w-4 h-4" />
                          <span>طباعة PDF</span>
                        </button>
                        <button
                          onClick={handleExportExcusesListWord}
                          className="bg-indigo-600 hover:bg-indigo-700 text-white font-bold px-3 py-2 rounded-xl text-xs flex items-center gap-1.5 transition-all shadow-xs h-[38px]"
                          title="تصدير كشف الغياب بعذر والأعذار Word"
                        >
                          <FileDown className="w-4 h-4" />
                          <span>Word</span>
                        </button>
                      </div>
                    </div>
                  </div>

                  {/* EXCUSES TABLE */}
                  <div className="bg-white border border-slate-200 rounded-2xl overflow-hidden shadow-sm">
                    <div className="overflow-x-auto">
                      <table className="w-full text-right text-sm border-collapse">
                        <thead>
                          <tr className="bg-slate-50 text-slate-600 border-b border-slate-200">
                            <th className="p-3 font-bold">اسم المعلم / الموظف</th>
                            <th className="p-3 font-bold">المسمى الوظيفي</th>
                            <th className="p-3 font-bold text-center">التاريخ</th>
                            <th className="p-3 font-bold text-center">اليوم</th>
                            <th className="p-3 font-bold text-center">نوع الإجازة / الاستئذان</th>
                            <th className="p-3 font-bold text-center">بصمة الانصراف والتوثيق</th>
                            <th className="p-3 font-bold text-center">المراسلة والتنبيه</th>
                            <th className="p-3 font-bold text-center">نظام فارس</th>
                            <th className="p-3 font-bold text-center">تعديل الوقت</th>
                          </tr>
                        </thead>
                        <tbody>
                          {filterDailyRecords(dailyRecords)
                            .filter(r => {
                              const isEm = isEmergencyLeave(r.status) || r.checkout === 'استئذان بدون بصمة انصراف' || r.checkout === 'استئذان طارئ موثق (بصمة)';
                              const isExc = r.type === 'excuse';
                              if (!isEm && !isExc) return false;
                              if (excuseSubFilter === 'leaves') return isExc && !isEm;
                              if (excuseSubFilter === 'emergency_all') return isEm;
                              if (excuseSubFilter === 'emergency_no_exit') return isEm && (!r.tout || r.tout.total == null || r.tout.text === '-');
                              if (excuseSubFilter === 'emergency_with_exit') return isEm && !!(r.tout && r.tout.total != null && r.tout.text !== '-');
                              return true;
                            })
                            .sort((a, b) => new Date(a.date).getTime() - new Date(b.date).getTime())
                            .map((r, i) => {
                              const rKey = `${r.name}|${r.dateKey}`;
                              const isEm = isEmergencyLeave(r.status) || r.checkout === 'استئذان بدون بصمة انصراف' || r.checkout === 'استئذان طارئ موثق (بصمة)';
                              const hasNoExit = !r.tout || r.tout.total == null || r.tout.text === '-';

                              return (
                                <tr key={i} className={`border-b border-slate-100 transition-all ${
                                  isEm && hasNoExit ? 'bg-amber-50/40 hover:bg-amber-50/70' : 'hover:bg-slate-50'
                                }`}>
                                  <td className="p-3 font-bold text-slate-800">{r.name}</td>
                                  <td className="p-3 text-slate-400 text-xs">{r.job || '-'}</td>
                                  <td className="p-3 text-center">{renderDateCell(r.date)}</td>
                                  <td className="p-3 text-center text-slate-500 text-xs font-semibold">{dayName(r.date)}</td>
                                  <td className="p-3 text-center">
                                    <span className={`font-bold text-[10px] px-2.5 py-0.5 rounded-full ${
                                      isEm
                                        ? 'bg-amber-100 text-amber-950 border border-amber-300'
                                        : 'bg-blue-100 text-blue-800'
                                    }`}>
                                      {r.status || (isEm ? 'استئذان طارئ' : 'غياب مبرر')}
                                    </span>
                                  </td>
                                  <td className="p-3 text-center">
                                    {isEm ? (
                                      hasNoExit ? (
                                        <span className="bg-rose-100 text-rose-800 border border-rose-300 font-black text-[10px] px-2.5 py-1 rounded-full shadow-2xs">
                                          ⚠️ بدون بصمة خروج (مخالف)
                                        </span>
                                      ) : (
                                        <span className="bg-emerald-100 text-emerald-800 border border-emerald-300 font-bold text-[10px] px-2.5 py-1 rounded-full">
                                          ✓ موثق ببصمة ({r.tout?.text})
                                        </span>
                                      )
                                    ) : (
                                      <span className="text-slate-400 text-xs">معفى / إجازة رسمية</span>
                                    )}
                                  </td>
                                  <td className="p-3 text-center">
                                    <button
                                      onClick={() => handleOpenMessageModal(rKey)}
                                      className={`font-bold px-3 py-1.5 rounded-xl text-xs transition-all inline-flex items-center gap-1.5 shadow-xs ${
                                        isEm && hasNoExit
                                          ? 'bg-amber-100 hover:bg-amber-200 text-amber-950 border border-amber-300'
                                          : 'bg-slate-50 hover:bg-slate-100 text-slate-700'
                                      }`}
                                    >
                                      <MessageSquare className={`w-3.5 h-3.5 ${isEm && hasNoExit ? 'text-amber-700' : 'text-blue-500'}`} />
                                      <span>{isEm && hasNoExit ? 'تنبيه بصمة الخروج' : 'إشعار اعتماد'}</span>
                                    </button>
                                  </td>
                                  <td className="p-3 text-center">
                                    {faresSubmitted[rKey] ? (
                                      <button
                                        onClick={() => toggleFaresStatus(rKey)}
                                        className="bg-rose-50 hover:bg-rose-100 text-rose-700 font-bold text-xs px-2.5 py-1.5 rounded-xl transition-all border border-rose-200 inline-flex items-center gap-1"
                                      >
                                        <Undo2 className="w-3.5 h-3.5" />
                                        <span>إلغاء الاعتماد</span>
                                      </button>
                                    ) : (
                                      <span className="text-xs text-slate-400 italic">عذر معتمد بالبصمة</span>
                                    )}
                                  </td>
                                  <td className="p-3 text-center">
                                    <button
                                      onClick={() => setEditingRecord(r)}
                                      className="text-amber-700 bg-amber-50 hover:bg-amber-100 font-bold px-2.5 py-1.5 rounded-lg text-xs transition-all inline-flex items-center gap-1"
                                      title="تعديل وقت البصمة وحالة الحضور"
                                    >
                                      <Edit className="w-3.5 h-3.5" />
                                      <span>تعديل</span>
                                    </button>
                                  </td>
                                </tr>
                              );
                            })}
                        </tbody>
                      </table>
                    </div>
                  </div>
                </div>
              )}

              {/* F. MONTHLY COMPREHENSIVE REPORT TAB */}
              {currentTab === 'monthly' && (
                <div className="space-y-6">
                  {/* COMPREHENSIVE SELECTION CONTROLS */}
                  <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm">
                    <div className="flex flex-wrap items-center justify-between gap-2 mb-3">
                      <h3 className="text-base font-extrabold text-teal-950">توليد تقرير الموظف المتكامل والدوري</h3>
                      {dailyRecords.length > 0 && (() => {
                        const validTimes = dailyRecords.map(r => r.date.getTime()).filter(t => !isNaN(t));
                        if (validTimes.length === 0) return null;
                        const minD = new Date(Math.min(...validTimes));
                        const maxD = new Date(Math.max(...validTimes));
                        const arMonthsNames = ['يناير', 'فبراير', 'مارس', 'أبريل', 'مايو', 'يونيو', 'يوليو', 'أغسطس', 'سبتمبر', 'أكتوبر', 'نوفمبر', 'ديسمبر'];
                        return (
                          <div className="text-xs bg-teal-50 border border-teal-200 text-teal-900 px-3 py-1 rounded-xl font-bold flex items-center gap-1.5 shadow-2xs">
                            <span className="w-2 h-2 rounded-full bg-teal-600 animate-pulse"></span>
                            <span>التواريخ المتاحة في قاعدة البيانات:</span>
                            <span className="text-teal-950 font-extrabold underline decoration-dotted">
                              من {minD.getDate()} {arMonthsNames[minD.getMonth()]} {minD.getFullYear()}م إلى {maxD.getDate()} {arMonthsNames[maxD.getMonth()]} {maxD.getFullYear()}م
                            </span>
                          </div>
                        );
                      })()}
                    </div>
                    <div className="flex flex-wrap items-end gap-4">
                      <div className="w-full md:w-72">
                        <label className="block text-xs font-bold text-slate-500 mb-1.5">اختر الموظف المستهدف</label>
                        <select
                          value={monthlyEmployee}
                          onChange={(e) => {
                            const val = e.target.value;
                            setMonthlyEmployee(val);
                            setSelectedEmployee(val);
                          }}
                          className="bg-slate-50 border border-slate-200 rounded-xl px-3 py-2.5 text-sm font-bold text-slate-700 w-full outline-none focus:border-teal-500"
                        >
                          <option value="all">📊 تقرير شامل لجميع المعلمين والموظفين</option>
                          {employees.map(e => (
                            <option key={e.name} value={e.name}>{e.name}</option>
                          ))}
                        </select>
                      </div>

                      <div className="w-full md:w-44">
                        <label className="block text-xs font-bold text-slate-500 mb-1.5">تصفية الفترة حسب</label>
                        <select
                          value={monthlyPeriodType}
                          onChange={(e) => {
                            const val = e.target.value as 'month' | 'range';
                            setMonthlyPeriodType(val);
                            setSelectedPeriodType(val);
                          }}
                          className="bg-slate-50 border border-slate-200 rounded-xl px-3 py-2.5 text-sm font-bold text-slate-700 w-full outline-none focus:border-teal-500"
                        >
                          <option value="month">📅 الشهر</option>
                          <option value="range">🗓️ نطاق تاريخ مخصص</option>
                        </select>
                      </div>

                      {monthlyPeriodType === 'month' ? (
                        <div className="w-full md:w-44">
                          <label className="block text-xs font-bold text-slate-500 mb-1.5">اختر الشهر</label>
                          <select
                            value={monthlyMonth}
                            onChange={(e) => {
                              const val = e.target.value;
                              setMonthlyMonth(val);
                              setSelectedMonth(val);
                            }}
                            className="bg-slate-50 border border-slate-200 rounded-xl px-3 py-2.5 text-sm font-bold text-slate-700 w-full outline-none focus:border-teal-500"
                          >
                            <option value="all">كل الأشهر المسجلة</option>
                            {getMonthsList().map(m => (
                              <option key={m} value={m}>{m}</option>
                            ))}
                          </select>
                        </div>
                      ) : (
                        <>
                          <div className="w-full md:w-40">
                            <label className="block text-xs font-bold text-slate-500 mb-1.5">من تاريخ (البدء)</label>
                            <input
                              type="date"
                              value={monthlyStartDate}
                              onChange={(e) => {
                                const val = e.target.value;
                                setMonthlyStartDate(val);
                                setSelectedStartDate(val);
                              }}
                              className="bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-sm font-bold text-slate-700 w-full outline-none focus:border-teal-500 h-[42px]"
                            />
                          </div>
                          <div className="w-full md:w-40">
                            <label className="block text-xs font-bold text-slate-500 mb-1.5">إلى تاريخ (النهاية)</label>
                            <input
                              type="date"
                              value={monthlyEndDate}
                              onChange={(e) => {
                                const val = e.target.value;
                                setMonthlyEndDate(val);
                                setSelectedEndDate(val);
                              }}
                              className="bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-sm font-bold text-slate-700 w-full outline-none focus:border-teal-500 h-[42px]"
                            />
                          </div>
                        </>
                      )}

                      <div className="flex items-end gap-2 mt-4 md:mt-0 w-full md:w-auto pt-4 md:pt-0">
                        <button
                          onClick={handlePrintReport}
                          disabled={employees.length === 0}
                          className="bg-teal-700 hover:bg-teal-800 disabled:bg-slate-200 text-white font-bold px-4 py-2.5 rounded-xl text-sm shadow-sm transition-all flex items-center gap-2 cursor-pointer"
                        >
                          <Printer className="w-4 h-4" />
                          <span>طباعة التقرير / PDF</span>
                        </button>

                        <button
                          onClick={handleExportWord}
                          disabled={employees.length === 0}
                          className="bg-amber-600 hover:bg-amber-700 disabled:bg-slate-200 text-white font-bold px-4 py-2.5 rounded-xl text-sm shadow-sm transition-all flex items-center gap-2 cursor-pointer"
                        >
                          <FileDown className="w-4 h-4" />
                          <span>تصدير Word</span>
                        </button>
                      </div>
                    </div>
                  </div>

                  {monthlyEmployee === 'all' ? (
                    (() => {
                      const allReports = employees.map(e => {
                        return monthlyPeriodType === 'range'
                          ? getIndividualReport(e.name, 'all', monthlyStartDate, monthlyEndDate)
                          : getIndividualReport(e.name, monthlyMonth);
                      });

                      const totalEmployees = allReports.length;
                      const totalWorkDays = allReports.reduce((acc, r) => acc + r.summary.work, 0);
                      const totalPresent = allReports.reduce((acc, r) => acc + r.summary.present, 0);
                      const totalAbsence = allReports.reduce((acc, r) => acc + r.summary.absence, 0);
                      const totalExcuse = allReports.reduce((acc, r) => acc + r.summary.excuse, 0);
                      const totalLateCount = allReports.reduce((acc, r) => acc + r.summary.lateCount, 0);
                      const totalLateMins = allReports.reduce((acc, r) => acc + r.summary.lateMins, 0);
                      const totalEarlyCount = allReports.reduce((acc, r) => acc + r.summary.earlyCount, 0);
                      const totalAutoCount = allReports.reduce((acc, r) => acc + r.summary.autoCount, 0);

                      const avgAttendanceRate = totalEmployees > 0 
                        ? Math.round(allReports.reduce((acc, r) => acc + (r.summary.attendanceRate || 0), 0) / totalEmployees) 
                        : 0;

                      const avgDisciplineRate = totalEmployees > 0 
                        ? Math.round(allReports.reduce((acc, r) => acc + (r.summary.rate || 0), 0) / totalEmployees) 
                        : 0;

                      const filteredReports = allReports.filter(r => {
                        if (!comprehensiveSearch) return true;
                        const term = comprehensiveSearch.trim().toLowerCase();
                        return r.summary.name.toLowerCase().includes(term) || (r.summary.job || '').toLowerCase().includes(term);
                      });

                      return (
                        <div className="space-y-6" dir="rtl">
                          {/* Aggregates Dashboard */}
                          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
                            <div className="bg-gradient-to-br from-teal-50 to-teal-100/30 p-5 rounded-2xl border border-teal-100 shadow-xs">
                              <span className="text-xs text-slate-400 font-extrabold block mb-1">إجمالي منسوبي المدرسة</span>
                              <div className="flex items-baseline gap-1.5">
                                <span className="text-2xl font-black text-teal-950">{totalEmployees}</span>
                                <span className="text-xs text-slate-500 font-bold">موظف ومعلم</span>
                              </div>
                            </div>

                            <div className="bg-gradient-to-br from-emerald-50 to-emerald-100/30 p-5 rounded-2xl border border-emerald-100 shadow-xs">
                              <span className="text-xs text-slate-400 font-extrabold block mb-1">متوسط نسبة الحضور والانتظام</span>
                              <div className="flex items-baseline gap-1.5">
                                <span className="text-2xl font-black text-emerald-700">{avgAttendanceRate}%</span>
                                <span className="text-xs text-slate-500 font-bold">للفترة المحددة</span>
                              </div>
                            </div>

                            <div className="bg-gradient-to-br from-blue-50 to-blue-100/30 p-5 rounded-2xl border border-blue-100 shadow-xs">
                              <span className="text-xs text-slate-400 font-extrabold block mb-1">متوسط درجة الانضباط المدرسي</span>
                              <div className="flex items-baseline gap-1.5">
                                <span className="text-2xl font-black text-blue-700">{avgDisciplineRate}%</span>
                                <span className="text-xs text-slate-500 font-bold">التقييم العام</span>
                              </div>
                            </div>

                            <div className="bg-gradient-to-br from-rose-50 to-rose-100/30 p-5 rounded-2xl border border-rose-100 shadow-xs">
                              <span className="text-xs text-slate-400 font-extrabold block mb-1">إجمالي الغياب والتأخر الصباحي</span>
                              <div className="flex items-baseline gap-1.5">
                                <span className="text-xl font-black text-rose-700">{totalAbsence} غياب</span>
                                <span className="text-xs text-slate-400 font-bold">|</span>
                                <span className="text-xl font-black text-amber-600">{totalLateCount} تأخر</span>
                              </div>
                            </div>
                          </div>

                          {/* Unified Table view */}
                          <div className="bg-white border border-slate-200 rounded-2xl shadow-sm overflow-hidden">
                            <div className="p-4 border-b border-slate-200 bg-slate-50/50 flex flex-wrap items-center justify-between gap-4">
                              <div>
                                <h4 className="text-sm font-extrabold text-teal-950">بيانات الحضور التفصيلية والدرجات لعموم المنسوبين</h4>
                                <p className="text-[11px] text-slate-400 mt-0.5">انقر على اسم أي موظف أو على زر "التفاصيل" لاستعراض تقريره الفردي وسجل بصماته اليومي.</p>
                              </div>
                              
                              <div className="w-full md:w-72">
                                <input
                                  type="text"
                                  placeholder="🔍 ابحث بالاسم أو المسمى الوظيفي..."
                                  value={comprehensiveSearch}
                                  onChange={(e) => setComprehensiveSearch(e.target.value)}
                                  className="bg-white border border-slate-200 rounded-xl px-3 py-2 text-xs w-full outline-none font-bold text-slate-700 focus:border-teal-500 transition-all shadow-xs"
                                />
                              </div>
                            </div>

                            <div className="overflow-x-auto">
                              <table className="w-full text-right text-xs border-collapse">
                                <thead>
                                  <tr className="bg-slate-50/30 text-slate-500 border-b border-slate-200">
                                    <th className="p-3.5 font-bold">المعلم / الموظف</th>
                                    <th className="p-3.5 font-bold">المسمى الوظيفي</th>
                                    <th className="p-3.5 font-bold text-center">أيام العمل</th>
                                    <th className="p-3.5 font-bold text-center text-emerald-700">الحضور</th>
                                    <th className="p-3.5 font-bold text-center text-blue-700">غياب بعذر</th>
                                    <th className="p-3.5 font-bold text-center text-rose-700">غياب بدون عذر</th>
                                    <th className="p-3.5 font-bold text-center">حالات التأخر</th>
                                    <th className="p-3.5 font-bold text-center">دقائق التأخر</th>
                                    <th className="p-3.5 font-bold text-center">انصراف مبكر/تلقائي</th>
                                    <th className="p-3.5 font-bold text-center">درجة الانضباط</th>
                                    <th className="p-3.5 font-bold text-center">المؤشر والتقييم</th>
                                    <th className="p-3.5 font-bold text-center">الإجراءات</th>
                                  </tr>
                                </thead>
                                <tbody>
                                  {filteredReports.length === 0 ? (
                                    <tr>
                                      <td colSpan={12} className="p-8 text-center text-slate-400 font-bold">
                                        لا توجد نتائج مطابقة لعملية البحث.
                                      </td>
                                    </tr>
                                  ) : (
                                    filteredReports.map((r) => (
                                      <tr key={r.summary.name} className="border-b border-slate-100 hover:bg-slate-50/50 transition-all">
                                        <td className="p-3 font-bold text-slate-800">
                                          <button
                                            onClick={() => setMonthlyEmployee(r.summary.name)}
                                            className="hover:text-teal-600 transition-all font-bold text-right outline-none cursor-pointer"
                                          >
                                            {r.summary.name}
                                          </button>
                                        </td>
                                        <td className="p-3 text-slate-500 font-semibold">{r.summary.job || 'غير محدد'}</td>
                                        <td className="p-3 text-center font-bold">{r.summary.work}</td>
                                        <td className="p-3 text-center font-bold text-emerald-600">{r.summary.present}</td>
                                        <td className="p-3 text-center font-bold text-blue-600">{r.summary.excuse}</td>
                                        <td className="p-3 text-center font-bold text-rose-600">{r.summary.absence}</td>
                                        <td className="p-3 text-center font-semibold text-amber-600">{r.summary.lateCount}</td>
                                        <td className="p-3 text-center font-semibold text-amber-700">{r.summary.lateMins} د</td>
                                        <td className="p-3 text-center font-semibold text-rose-500">
                                          {r.summary.earlyCount + r.summary.autoCount}
                                        </td>
                                        <td className="p-3 text-center font-black text-slate-900">{r.summary.rate}%</td>
                                        <td className="p-3 text-center">
                                          <span className={`px-2.5 py-1 rounded-full text-[10px] font-bold ${
                                            (r.summary.rate || 0) >= 95 ? 'bg-emerald-100 text-emerald-800' :
                                            (r.summary.rate || 0) >= 85 ? 'bg-blue-100 text-blue-800' :
                                            (r.summary.rate || 0) >= 75 ? 'bg-amber-100 text-amber-800' : 'bg-rose-100 text-rose-800'
                                          }`}>
                                            {r.summary.rating || 'غير محدد'}
                                          </span>
                                        </td>
                                        <td className="p-3 text-center">
                                          <button
                                            onClick={() => setMonthlyEmployee(r.summary.name)}
                                            className="bg-slate-100 hover:bg-teal-50 hover:text-teal-700 text-slate-700 px-3 py-1.5 rounded-lg text-[11px] font-bold transition-all inline-flex items-center gap-1 cursor-pointer"
                                          >
                                            <span>تفاصيل الملف</span>
                                            <span>←</span>
                                          </button>
                                        </td>
                                      </tr>
                                    ))
                                  )}
                                </tbody>
                              </table>
                            </div>
                          </div>
                        </div>
                      );
                    })()
                  ) : (
                    (() => {
                      const rep = monthlyPeriodType === 'range'
                        ? getIndividualReport(monthlyEmployee, 'all', monthlyStartDate, monthlyEndDate)
                        : getIndividualReport(monthlyEmployee, monthlyMonth);
                      const m = rep.summary;

                      const getDatesForCard = (filterName: string) => {
                        let matched: any[] = [];
                        if (filterName === 'absence') matched = rep.details.filter(r => r.type === 'absence');
                        else if (filterName === 'excuse') matched = rep.details.filter(r => r.type === 'excuse');
                        else if (filterName === 'late') matched = rep.details.filter(r => r.late > 0);
                        else if (filterName === 'late_mins') matched = rep.details.filter(r => r.late > 0);
                        else if (filterName === 'early') matched = rep.details.filter(r => r.checkout === 'خروج قبل نهاية الدوام' || r.checkout === 'انصراف مبكر' || r.earlyMins > 0);
                        else if (filterName === 'auto') matched = rep.details.filter(r => r.checkout === 'انصراف تلقائي');
                        else if (filterName === 'missing') matched = rep.details.filter(r => r.checkout === 'لا توجد بصمة انصراف');
                        else if (filterName === 'wrong_time') matched = rep.details.filter(r => r.checkout === 'انصراف في غير وقت الفعلي');
                        else if (filterName === 'emergency_no_exit') matched = rep.details.filter(r => r.checkout === 'استئذان بدون بصمة انصراف');
                        else if (filterName === 'emergency_with_exit') matched = rep.details.filter(r => r.checkout === 'استئذان طارئ موثق (بصمة)');
                        
                        if (matched.length === 0) return '';
                        const arMonthNamesList = ['يناير', 'فبراير', 'مارس', 'أبريل', 'مايو', 'يونيو', 'يوليو', 'أغسطس', 'سبتمبر', 'أكتوبر', 'نوفمبر', 'ديسمبر'];
                        return matched.map(r => {
                          const d = r.date;
                          const mName = arMonthNamesList[d.getMonth()] || `${d.getMonth() + 1}`;
                          return `${d.getDate()} ${mName}`;
                        }).join(' ، ');
                      };

                      const wrongTimeExitsCount = rep.details.filter(r => r.checkout === 'انصراف في غير وقت الفعلي').length;

                      const handleCardClick = (filterType: string | null) => {
                        setDetailCardFilter(prev => {
                          const next = prev === filterType ? null : filterType;
                          setTimeout(() => {
                            document.getElementById('report-details-table-container')?.scrollIntoView({ behavior: 'smooth' });
                          }, 100);
                          return next;
                        });
                      };

                      const cardStyle = (filterType: string | null) => {
                        const isActive = detailCardFilter === filterType;
                        const base = "bg-white p-4 rounded-xl border transition-all text-right select-none flex flex-col justify-between min-h-[90px] cursor-pointer hover:border-teal-500 hover:shadow-md hover:-translate-y-0.5 active:scale-[0.98]";
                        const active = isActive 
                          ? "border-teal-600 bg-teal-50/20 ring-2 ring-teal-600/20 shadow-sm font-extrabold" 
                          : "border-slate-200 shadow-xs";
                        return `${base} ${active}`;
                      };

                      return (
                        <div className="space-y-6">
                          {/* Profile Header Card */}
                          <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-xs">
                            <h2 className="text-2xl font-black text-teal-950">{m.name}</h2>
                            <p className="text-xs text-slate-400 font-bold mt-1">
                              المسمى الوظيفي: {m.job || 'غير محدد'} | السجل المدني: {m.civil || 'غير محدد'} | الفترة: {m.month}
                            </p>
                          </div>

                          {/* Interactive Stat Grid Guidance */}
                          <div className="text-xs text-slate-500 font-medium bg-slate-50 border border-slate-200 p-2.5 rounded-xl flex items-center gap-2">
                            <span className="flex h-2 w-2 relative">
                              <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-teal-400 opacity-75"></span>
                              <span className="relative inline-flex rounded-full h-2 w-2 bg-teal-500"></span>
                            </span>
                            <span>البطاقات أدناه تفاعلية: انقر على أي بطاقة لتصفية السجل والذهاب مباشرة لليوم المحدد.</span>
                          </div>

                          {/* Stat Grid */}
                          <div className="grid grid-cols-2 md:grid-cols-4 lg:grid-cols-5 gap-4">
                            
                            {/* 1. أيام العمل */}
                            <div
                              onClick={() => handleCardClick(null)}
                              className={cardStyle(null)}
                            >
                              <div>
                                <span className="text-[10px] text-slate-400 font-bold block">أيام العمل الفعلية</span>
                                <p className="text-xl font-extrabold mt-1 text-slate-800">{m.work}</p>
                              </div>
                              <div className="mt-2 text-[9px] text-slate-400 border-t border-slate-100 pt-1 leading-normal">
                                عرض كل السجلات
                              </div>
                            </div>

                            {/* 2. الحضور */}
                            <div
                              onClick={() => handleCardClick('present')}
                              className={cardStyle('present')}
                            >
                              <div>
                                <span className="text-[10px] text-slate-400 font-bold block">الحضور الفعلي</span>
                                <p className="text-xl font-extrabold mt-1 text-emerald-600">{m.present}</p>
                              </div>
                              <div className="mt-2 text-[9px] text-slate-400 border-t border-slate-100 pt-1 leading-normal font-medium">
                                تصفية أيام الحضور
                              </div>
                            </div>

                            {/* 3. غياب بدون عذر */}
                            <div
                              onClick={() => handleCardClick('absence')}
                              className={cardStyle('absence')}
                            >
                              <div>
                                <span className="text-[10px] text-slate-400 font-bold block">غياب بدون عذر</span>
                                <p className="text-xl font-extrabold mt-1 text-rose-600">{m.absence}</p>
                              </div>
                              {getDatesForCard('absence') ? (
                                <div className="mt-2 text-[9px] text-rose-500 border-t border-rose-100 pt-1 leading-normal">
                                  <span className="font-bold">التواريخ:</span>{' '}
                                  <span className="font-mono font-bold">{getDatesForCard('absence')}</span>
                                </div>
                              ) : (
                                <div className="mt-2 text-[9px] text-slate-400 border-t border-slate-100 pt-1 leading-normal">
                                  لا يوجد غياب
                                </div>
                              )}
                            </div>

                            {/* 4. غياب بعذر/إجازة */}
                            <div
                              onClick={() => handleCardClick('excuse')}
                              className={cardStyle('excuse')}
                            >
                              <div>
                                <span className="text-[10px] text-slate-400 font-bold block">غياب بعذر/إجازة</span>
                                <p className="text-xl font-extrabold mt-1 text-blue-600">{m.excuse}</p>
                              </div>
                              {getDatesForCard('excuse') ? (
                                <div className="mt-2 text-[9px] text-blue-500 border-t border-blue-100 pt-1 leading-normal">
                                  <span className="font-bold">التواريخ:</span>{' '}
                                  <span className="font-mono font-bold">{getDatesForCard('excuse')}</span>
                                </div>
                              ) : (
                                <div className="mt-2 text-[9px] text-slate-400 border-t border-slate-100 pt-1 leading-normal">
                                  لا توجد إجازات
                                </div>
                              )}
                            </div>

                            {/* 5. مرات التأخر الصباحي */}
                            <div
                              onClick={() => handleCardClick('late')}
                              className={cardStyle('late')}
                            >
                              <div>
                                <span className="text-[10px] text-slate-400 font-bold block">مرات التأخر الصباحي</span>
                                <p className="text-xl font-extrabold mt-1 text-amber-600">{m.lateCount}</p>
                              </div>
                              {getDatesForCard('late') ? (
                                <div className="mt-2 text-[9px] text-amber-600 border-t border-amber-100 pt-1 leading-normal">
                                  <span className="font-bold">التواريخ:</span>{' '}
                                  <span className="font-mono font-bold">{getDatesForCard('late')}</span>
                                </div>
                              ) : (
                                <div className="mt-2 text-[9px] text-slate-400 border-t border-slate-100 pt-1 leading-normal">
                                  لا يوجد تأخر
                                </div>
                              )}
                            </div>

                            {/* 6. إجمالي دقائق التأخر */}
                            <div
                              onClick={() => handleCardClick('late_mins')}
                              className={cardStyle('late_mins')}
                            >
                              <div>
                                <span className="text-[10px] text-slate-400 font-bold block">إجمالي دقائق التأخر</span>
                                <p className="text-xl font-extrabold mt-1 text-amber-700">{m.lateMins}</p>
                              </div>
                              {getDatesForCard('late_mins') ? (
                                <div className="mt-2 text-[9px] text-amber-700 border-t border-amber-100 pt-1 leading-normal">
                                  <span className="font-bold">التواريخ:</span>{' '}
                                  <span className="font-mono font-bold">{getDatesForCard('late_mins')}</span>
                                </div>
                              ) : (
                                <div className="mt-2 text-[9px] text-slate-400 border-t border-slate-100 pt-1 leading-normal">
                                  0 دقيقة
                                </div>
                              )}
                            </div>

                            {/* 7. خروج مبكر / ناقص */}
                            <div
                              onClick={() => handleCardClick('early')}
                              className={cardStyle('early')}
                            >
                              <div>
                                <span className="text-[10px] text-slate-400 font-bold block">خروج قبل نهاية الدوام</span>
                                <p className="text-xl font-extrabold mt-1 text-rose-500">{m.earlyCount}</p>
                              </div>
                              {getDatesForCard('early') ? (
                                <div className="mt-2 text-[9px] text-rose-500 border-t border-rose-100 pt-1 leading-normal">
                                  <span className="font-bold">التواريخ:</span>{' '}
                                  <span className="font-mono font-bold">{getDatesForCard('early')}</span>
                                </div>
                              ) : (
                                <div className="mt-2 text-[9px] text-slate-400 border-t border-slate-100 pt-1 leading-normal">
                                  لا يوجد خروج مبكر
                                </div>
                              )}
                            </div>

                            {/* 8. انصراف تلقائي */}
                            <div
                              onClick={() => handleCardClick('auto')}
                              className={cardStyle('auto')}
                            >
                              <div>
                                <span className="text-[10px] text-slate-400 font-bold block">انصراف تلقائي</span>
                                <p className="text-xl font-extrabold mt-1 text-amber-500">{m.autoCount}</p>
                              </div>
                              {getDatesForCard('auto') ? (
                                <div className="mt-2 text-[9px] text-amber-600 border-t border-amber-100 pt-1 leading-normal">
                                  <span className="font-bold">التواريخ:</span>{' '}
                                  <span className="font-mono font-bold">{getDatesForCard('auto')}</span>
                                </div>
                              ) : (
                                <div className="mt-2 text-[9px] text-slate-400 border-t border-slate-100 pt-1 leading-normal">
                                  لا يوجد انصراف تلقائي
                                </div>
                              )}
                            </div>

                            {/* 9. لا توجد بصمة انصراف */}
                            <div
                              onClick={() => handleCardClick('missing')}
                              className={cardStyle('missing')}
                            >
                              <div>
                                <span className="text-[10px] text-slate-400 font-bold block">لا توجد بصمة انصراف</span>
                                <p className="text-xl font-extrabold mt-1 text-rose-500">{m.missingCheckoutCount}</p>
                              </div>
                              {getDatesForCard('missing') ? (
                                <div className="mt-2 text-[9px] text-rose-500 border-t border-rose-100 pt-1 leading-normal">
                                  <span className="font-bold">التواريخ:</span>{' '}
                                  <span className="font-mono font-bold">{getDatesForCard('missing')}</span>
                                </div>
                              ) : (
                                <div className="mt-2 text-[9px] text-slate-400 border-t border-slate-100 pt-1 leading-normal">
                                  لا يوجد بصمات ناقصة
                                </div>
                              )}
                            </div>

                            {/* 10. انصراف في غير وقت الفعلي */}
                            <div
                              onClick={() => handleCardClick('wrong_time')}
                              className={cardStyle('wrong_time')}
                            >
                              <div>
                                <span className="text-[10px] text-slate-400 font-bold block">انصراف في غير وقته الفعلي</span>
                                <p className="text-xl font-extrabold mt-1 text-purple-600">{wrongTimeExitsCount}</p>
                              </div>
                              {getDatesForCard('wrong_time') ? (
                                <div className="mt-2 text-[9px] text-purple-600 border-t border-purple-100 pt-1 leading-normal">
                                  <span className="font-bold">التواريخ:</span>{' '}
                                  <span className="font-mono font-bold">{getDatesForCard('wrong_time')}</span>
                                </div>
                              ) : (
                                <div className="mt-2 text-[9px] text-slate-400 border-t border-slate-100 pt-1 leading-normal">
                                  لا توجد بصمات خاطئة
                                </div>
                              )}
                            </div>

                            {/* 11. استئذان بلا بصمة انصراف */}
                            <div
                              onClick={() => handleCardClick('emergency_no_exit')}
                              className={cardStyle('emergency_no_exit')}
                            >
                              <div>
                                <span className="text-[10px] text-amber-600 font-bold block">استئذان بلا بصمة انصراف</span>
                                <p className="text-xl font-extrabold mt-1 text-amber-600">{m.emergencyNoExitCount || 0}</p>
                              </div>
                              {getDatesForCard('emergency_no_exit') ? (
                                <div className="mt-2 text-[9px] text-amber-600 border-t border-amber-100 pt-1 leading-normal">
                                  <span className="font-bold">التواريخ:</span>{' '}
                                  <span className="font-mono font-bold">{getDatesForCard('emergency_no_exit')}</span>
                                </div>
                              ) : (
                                <div className="mt-2 text-[9px] text-slate-400 border-t border-slate-100 pt-1 leading-normal">
                                  لا توجد مخالفات
                                </div>
                              )}
                            </div>

                          </div>

                          {/* Deduction Reasons */}
                          <div className="bg-orange-50 border border-orange-200 p-4 rounded-2xl text-orange-950 text-xs">
                            <strong>ملاحظات وتنبيهات سجل الحضور والانتظام:</strong> {m.disciplineReason}
                          </div>

                          {/* Table details */}
                          <div id="report-details-table-container" className="bg-white border border-slate-200 rounded-2xl overflow-hidden shadow-sm">
                            <div className="p-4 border-b border-slate-200 bg-slate-50 font-bold text-slate-700 text-sm flex flex-wrap items-center justify-between gap-2">
                              <span>السجل اليومي التفصيلي الشامل للحضور والانصراف</span>
                              {detailCardFilter && (
                                <button
                                  onClick={() => setDetailCardFilter(null)}
                                  className="text-xs bg-rose-50 border border-rose-200 text-rose-600 hover:bg-rose-100/50 px-2.5 py-1 rounded-lg font-bold flex items-center gap-1 cursor-pointer transition-all"
                                >
                                  × إلغاء التصفية وعرض الكل
                                </button>
                              )}
                            </div>
                            <div className="overflow-x-auto">
                              <table className="w-full text-right text-sm border-collapse">
                                <thead>
                                  <tr className="bg-slate-50/50 text-slate-600 border-b border-slate-200 text-xs">
                                    <th className="p-3 font-bold">التاريخ</th>
                                    <th className="p-3 font-bold">اليوم</th>
                                    <th className="p-3 font-bold">حالة اليوم</th>
                                    <th className="p-3 font-bold text-center">وقت الحضور</th>
                                    <th className="p-3 font-bold text-center">وقت الانصراف</th>
                                    <th className="p-3 font-bold text-center">حالة الانصراف</th>
                                    <th className="p-3 font-bold text-center">دقائق التأخر</th>
                                    <th className="p-3 font-bold text-center">الخروج المبكر</th>
                                  </tr>
                                </thead>
                                <tbody>
                                  {(() => {
                                    const filteredDetails = rep.details.filter(r => {
                                      if (!detailCardFilter) return true;
                                      if (detailCardFilter === 'present') return r.type === 'present';
                                      if (detailCardFilter === 'absence') return r.type === 'absence';
                                      if (detailCardFilter === 'excuse') return r.type === 'excuse';
                                      if (detailCardFilter === 'late' || detailCardFilter === 'late_mins') return r.late > 0;
                                      if (detailCardFilter === 'early') return r.earlyMins > 0 || r.checkout === 'خروج قبل نهاية الدوام' || r.checkout === 'انصراف مبكر';
                                      if (detailCardFilter === 'auto') return r.checkout === 'انصراف تلقائي';
                                      if (detailCardFilter === 'missing') return r.checkout === 'لا توجد بصمة انصراف';
                                      if (detailCardFilter === 'wrong_time') return r.checkout === 'انصراف في غير وقت الفعلي';
                                      if (detailCardFilter === 'emergency_no_exit') return r.checkout === 'استئذان بدون بصمة انصراف';
                                      if (detailCardFilter === 'emergency_with_exit') return r.checkout === 'استئذان طارئ موثق (بصمة)';
                                      return true;
                                    });

                                    if (filteredDetails.length === 0) {
                                      return (
                                        <tr>
                                          <td colSpan={8} className="p-8 text-center text-slate-400 font-bold">
                                            <div className="flex flex-col items-center justify-center gap-2">
                                              <span>
                                                {detailCardFilter
                                                  ? 'لا توجد سجلات مطابقة لتصفية البطاقة المحددة في هذه الفترة.'
                                                  : (monthlyPeriodType === 'range'
                                                      ? 'لا توجد سجلات مسجلة لهذا الموظف في نطاق التواريخ المحددة.'
                                                      : 'لا توجد سجلات مسجلة لهذا الموظف في هذا الشهر.')}
                                              </span>
                                              {detailCardFilter && (
                                                <button
                                                  onClick={() => setDetailCardFilter(null)}
                                                  className="text-xs bg-teal-50 text-teal-700 hover:bg-teal-100 px-3 py-1.5 rounded-lg font-bold border border-teal-200 transition-all cursor-pointer inline-flex items-center gap-1"
                                                >
                                                  <span>عرض جميع سجلات الفترة (إلغاء تصفية البطاقة)</span>
                                                </button>
                                              )}
                                            </div>
                                          </td>
                                        </tr>
                                      );
                                    }

                                    return filteredDetails.map((r, i) => (
                                      <tr key={i} className={`border-b border-slate-100 hover:bg-slate-50/50 transition-all text-xs ${
                                        detailCardFilter ? 'bg-teal-50/10' : ''
                                      }`}>
                                        <td className="p-3 font-bold text-slate-700">{renderDateCell(r.date)}</td>
                                        <td className="p-3 text-slate-500 font-semibold">{dayName(r.date)}</td>
                                        <td className="p-3">
                                          <span className={`px-2 py-0.5 rounded-full font-bold text-[10px] ${
                                            r.type === 'present' ? 'bg-emerald-100 text-emerald-800' :
                                            r.type === 'excuse' ? 'bg-blue-100 text-blue-800' :
                                            'bg-rose-100 text-rose-800'
                                          }`}>
                                            {r.type === 'present' ? 'حاضر' : r.status}
                                          </span>
                                        </td>
                                        <td className="p-3 text-center font-mono font-bold">{r.tin?.text || '-'}</td>
                                        <td className="p-3 text-center font-mono font-bold">{r.tout?.text || '-'}</td>
                                        <td className="p-3 text-center">
                                          {(r.type === 'present' || r.checkout?.includes('استئذان')) && (
                                            <span className={`text-[10px] font-bold px-1.5 py-0.5 rounded ${
                                              r.checkout === 'مكتمل' ? 'bg-emerald-100 text-emerald-800' :
                                              r.checkout === 'انصراف تلقائي' ? 'bg-amber-100 text-amber-800' :
                                              (r.checkout === 'خروج قبل نهاية الدوام' || r.checkout === 'انصراف مبكر') ? 'bg-rose-100 text-rose-800' :
                                              r.checkout === 'انصراف في غير وقت الفعلي' ? 'bg-purple-100 text-purple-800 border border-purple-300 font-extrabold shadow-xs' :
                                              r.checkout === 'استئذان بدون بصمة انصراف' ? 'bg-amber-100 text-amber-900 border border-amber-300 font-extrabold shadow-xs' :
                                              r.checkout === 'استئذان طارئ موثق (بصمة)' ? 'bg-blue-100 text-blue-800 border border-blue-200 font-bold' :
                                              (r.checkout && !['لا توجد بصمة انصراف', 'غير متاح'].includes(r.checkout)) ? 'bg-blue-100 text-blue-800' :
                                              'bg-rose-100 text-rose-800'
                                            }`}>
                                              {r.checkout === 'انصراف في غير وقت الفعلي' ? 'انصراف في غير وقت الفعلي (بصمة صباحية)' : r.checkout}
                                            </span>
                                          )}
                                        </td>
                                        <td className="p-3 text-center font-bold text-amber-600">{r.late || 0} د</td>
                                        <td className="p-3 text-center font-bold text-rose-600">{r.earlyMins || 0} د</td>
                                      </tr>
                                    ));
                                  })()}
                                </tbody>
                              </table>
                            </div>
                          </div>
                        </div>
                      );
                    })()
                  )}
                </div>
              )}

              {/* G. SYSTEM CONFIGURATION / SETTINGS TAB */}
              {currentTab === 'settings' && (
                <div className="space-y-6" dir="rtl">
                  {/* Master Timing Dashboard Header */}
                  <div className="bg-gradient-to-r from-teal-800 to-teal-950 text-white p-6 rounded-3xl border border-teal-700/30 shadow-lg relative overflow-hidden">
                    <div className="absolute top-0 left-0 translate-y-[-20%] translate-x-[-20%] opacity-10 text-[180px] pointer-events-none select-none">
                      🌙
                    </div>
                    <div className="relative z-10">
                      <span className="bg-amber-400/20 text-amber-300 text-[11px] font-black tracking-wider uppercase px-3 py-1 rounded-full border border-amber-400/30">
                        🌙 ميزة التوقيت المزدوج الذكي
                      </span>
                      <h3 className="text-xl font-extrabold text-white mt-2">تشغيل وإدارة كلا التوقيتين معاً بالتزامن</h3>
                      <p className="text-xs text-teal-100/80 leading-relaxed max-w-3xl mt-1.5">
                        الآن يمكنك تفعيل الدوام الاعتيادي ودوام شهر رمضان المبارك في نفس الوقت! سيقوم النظام تلقائياً بتطبيق توقيت شهر رمضان (09:00 ص، 5 ساعات عمل، وانصراف تلقائي 08:45) على التواريخ التي تحددها بالأسفل، مع الحفاظ على الدوام المدرسي الاعتيادي لبقية أيام السنة دون أي تداخل في التقارير.
                      </p>
                    </div>
                  </div>

                  {/* Dual Timings Columns */}
                  <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                    {/* Column 1: Normal Timings */}
                    <div className="bg-white p-6 rounded-3xl border border-slate-200 shadow-sm space-y-4">
                      <div className="flex items-center gap-2 pb-3 border-b border-slate-100">
                        <div className="bg-slate-100 p-2 rounded-xl text-slate-700">
                          🏫
                        </div>
                        <div>
                          <h4 className="font-extrabold text-slate-900 text-sm">1. إعدادات الدوام الرسمي الاعتيادي</h4>
                          <p className="text-[10px] text-slate-400">يطبق تلقائياً على الأيام العادية طوال العام</p>
                        </div>
                      </div>

                      <div className="space-y-3.5">
                        <div>
                          <label className="block text-xs font-bold text-slate-500 mb-1">بداية الدوام المعتمد للتأخير</label>
                          <input
                            type="time"
                            value={config.start}
                            onChange={(e) => setConfig({ ...config, start: e.target.value })}
                            className="bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-sm font-bold text-slate-700 outline-none w-full focus:border-teal-500"
                          />
                        </div>
                        <div className="grid grid-cols-2 gap-2">
                          <div>
                            <label className="block text-[11px] font-bold text-slate-600 mb-1">
                              خروج معتمد (الأحد - الأربعاء)
                            </label>
                            <input
                              type="time"
                              value={config.checkoutSunWed || '12:30'}
                              onChange={(e) => setConfig({ ...config, checkoutSunWed: e.target.value })}
                              className="bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-sm font-bold text-slate-700 outline-none w-full focus:border-teal-500 font-mono"
                            />
                          </div>
                          <div>
                            <label className="block text-[11px] font-bold text-slate-600 mb-1">
                              خروج معتمد (الخميس)
                            </label>
                            <input
                              type="time"
                              value={config.checkoutThu || '11:45'}
                              onChange={(e) => setConfig({ ...config, checkoutThu: e.target.value })}
                              className="bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-sm font-bold text-slate-700 outline-none w-full focus:border-teal-500 font-mono"
                            />
                          </div>
                        </div>
                        <div>
                          <label className="block text-xs font-bold text-slate-500 mb-1">
                            وقت الدخول في الانصراف التلقائي (من 13:30 وما زاد)
                          </label>
                          <input
                            type="time"
                            value={config.autoCheckout}
                            onChange={(e) => setConfig({ ...config, autoCheckout: e.target.value })}
                            className="bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-sm font-bold text-slate-700 outline-none w-full focus:border-teal-500 font-mono"
                          />
                          <p className="text-[10px] text-slate-400 mt-1">أي بصمة انصراف من هذا الوقت فما بعد تُصنف كـ انصراف تلقائي</p>
                        </div>
                        <div>
                          <label className="block text-xs font-bold text-slate-500 mb-1">عدد ساعات العمل اليومية (استرشادي)</label>
                          <input
                            type="number"
                            step="0.25"
                            value={config.workHours}
                            onChange={(e) => setConfig({ ...config, workHours: parseFloat(e.target.value) || 7 })}
                            className="bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-sm font-bold text-slate-700 outline-none w-full focus:border-teal-500"
                          />
                        </div>
                      </div>
                    </div>

                    {/* Column 2: Ramadan Timings */}
                    <div className="bg-emerald-50/40 p-6 rounded-3xl border border-emerald-100 shadow-sm space-y-4">
                      <div className="flex items-center gap-2 pb-3 border-b border-emerald-100">
                        <div className="bg-emerald-50 p-2 rounded-xl text-emerald-700">
                          🌙
                        </div>
                        <div>
                          <h4 className="font-extrabold text-emerald-950 text-sm">2. إعدادات دوام شهر رمضان المبارك</h4>
                          <p className="text-[10px] text-emerald-600/80">يطبق فقط على الأيام التي يتم تحديدها كـ دوام رمضاني</p>
                        </div>
                      </div>

                      <div className="space-y-3.5">
                        <div>
                          <label className="block text-xs font-bold text-emerald-800 mb-1">بداية الدوام المعتمد للتأخير</label>
                          <input
                            type="time"
                            value={config.ramadanStart || '09:00'}
                            onChange={(e) => setConfig({ ...config, ramadanStart: e.target.value })}
                            className="bg-white border border-emerald-200 rounded-xl px-3 py-2 text-sm font-bold text-emerald-950 outline-none w-full focus:border-emerald-500"
                          />
                        </div>
                        <div>
                          <label className="block text-xs font-bold text-emerald-800 mb-1">عدد ساعات العمل اليومية المطلوبة</label>
                          <input
                            type="number"
                            step="0.25"
                            value={config.ramadanWorkHours ?? 5}
                            onChange={(e) => setConfig({ ...config, ramadanWorkHours: parseFloat(e.target.value) || 5 })}
                            className="bg-white border border-emerald-200 rounded-xl px-3 py-2 text-sm font-bold text-emerald-950 outline-none w-full focus:border-emerald-500"
                          />
                        </div>
                        <div>
                          <label className="block text-xs font-bold text-emerald-800 mb-1">وقت الدخول في الانصراف التلقائي</label>
                          <input
                            type="time"
                            value={config.ramadanAutoCheckout || '08:45'}
                            onChange={(e) => setConfig({ ...config, ramadanAutoCheckout: e.target.value })}
                            className="bg-white border border-emerald-200 rounded-xl px-3 py-2 text-sm font-bold text-emerald-950 outline-none w-full focus:border-emerald-500"
                          />
                        </div>
                      </div>
                    </div>
                  </div>

                  {/* Section 3: Interactive Dates Grid & Selection */}
                  <div className="bg-white p-6 rounded-3xl border border-slate-200 shadow-sm space-y-4">
                    <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 pb-3 border-b border-slate-100">
                      <div className="flex items-center gap-2">
                        <div className="bg-amber-50 p-2 rounded-xl text-amber-600">
                          📅
                        </div>
                        <div>
                          <h4 className="font-extrabold text-slate-900 text-sm">3. تحديد أيام وتواريخ الدوام الرمضاني</h4>
                          <p className="text-[11px] text-slate-400">انقر على اليوم للتبديل الفوري بين الدوام الاعتيادي والرمضاني</p>
                        </div>
                      </div>
                      <div className="bg-amber-50 text-amber-900 font-extrabold text-xs px-3.5 py-1.5 rounded-full border border-amber-100">
                        الأيام الرمضانية النشطة: {(config.ramadanDates || []).length} يوم
                      </div>
                    </div>

                    {/* Bulk controls */}
                    {dailyRecords.length > 0 ? (
                      <div className="space-y-4">
                        {/* Ramadan 2026 Preset Alert/Control */}
                        <div className="bg-gradient-to-r from-emerald-50/90 to-teal-50/90 border border-emerald-100 p-4 rounded-2xl space-y-3">
                          <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
                            <div className="space-y-1">
                              <h5 className="text-xs font-black text-emerald-950 flex items-center gap-1.5">
                                <span>🌙</span>
                                <span>تطبيق فوري لفترة شهر رمضان المبارك لعام 2026م (1447هـ)</span>
                              </h5>
                              <p className="text-[11px] text-emerald-800 leading-relaxed">
                                يبدأ شهر رمضان المبارك لعام 2026م من <strong className="font-extrabold text-emerald-950">18 فبراير 2026</strong> إلى <strong className="font-extrabold text-emerald-950">19 مارس 2026</strong>.
                                <br />
                                انقر على الزر بالجانب لتفعيل توقيت شهر رمضان على هذه الفترة تلقائياً دفعة واحدة!
                              </p>
                            </div>
                            <button
                              onClick={() => setRamadanDateRange('2026-02-18', '2026-03-19')}
                              className="bg-emerald-600 hover:bg-emerald-700 active:scale-95 text-white text-xs font-black px-4.5 py-2.5 rounded-xl shadow-sm transition-all flex items-center gap-1.5 shrink-0 self-start md:self-center"
                            >
                              <span>🌙 تفعيل هذه الفترة الآن (18-2 إلى 19-3)</span>
                            </button>
                          </div>

                          <div className="border-t border-emerald-100/60 pt-3 flex flex-wrap items-center gap-3">
                            <span className="text-[10px] font-bold text-teal-800">أو حدد أي نطاق تاريخ مخصص آخر:</span>
                            <div className="flex items-center gap-1.5">
                              <input
                                type="date"
                                id="custom-ram-start"
                                defaultValue="2026-02-18"
                                className="bg-white border border-slate-200 rounded-lg px-2 py-1 text-xs font-bold text-slate-700 outline-none focus:border-teal-500"
                              />
                              <span className="text-slate-400 text-[10px]">إلى</span>
                              <input
                                type="date"
                                id="custom-ram-end"
                                defaultValue="2026-03-19"
                                className="bg-white border border-slate-200 rounded-lg px-2 py-1 text-xs font-bold text-slate-700 outline-none focus:border-teal-500"
                              />
                            </div>
                            <button
                              onClick={() => {
                                const startInput = document.getElementById('custom-ram-start') as HTMLInputElement;
                                const endInput = document.getElementById('custom-ram-end') as HTMLInputElement;
                                if (startInput && endInput) {
                                  setRamadanDateRange(startInput.value, endInput.value);
                                }
                              }}
                              className="bg-teal-700 hover:bg-teal-800 text-white text-[10px] font-bold px-3 py-1.5 rounded-lg transition-all"
                            >
                              تطبيق الفترة المحددة
                            </button>
                          </div>
                        </div>

                        <div className="flex flex-wrap gap-2">
                          <button
                            onClick={setAllAsRamadanDates}
                            className="bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold px-3.5 py-2 rounded-xl transition-all"
                          >
                            🌙 تعيين كل الأيام المتاحة كـ دوام رمضان
                          </button>
                          <button
                            onClick={clearAllRamadanDates}
                            className="bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold px-3.5 py-2 rounded-xl transition-all border border-slate-200"
                          >
                            🏫 إعادة تعيين الكل كـ دوام اعتيادي
                          </button>
                        </div>

                        {/* Month-based bulk actions */}
                        <div className="bg-slate-50 p-3 rounded-2xl border border-slate-100 space-y-2">
                          <span className="block text-[10px] font-bold text-slate-400">تفعيل الدوام الرمضاني دفعة واحدة لشهور كاملة:</span>
                          <div className="flex flex-wrap gap-2">
                            {(getMonthsList() as any[]).map((m: string) => (
                              <button
                                key={m}
                                onClick={() => setBulkRamadanDatesByMonth(m)}
                                className="bg-white hover:bg-amber-50 hover:text-amber-800 text-slate-700 text-xs font-bold px-3 py-1.5 rounded-lg transition-all border border-slate-200 flex items-center gap-1 shadow-sm"
                              >
                                <span>🌙 تفعيل شهر {m} كاملاً</span>
                              </button>
                            ))}
                          </div>
                        </div>

                        {/* Interactive Clickable Date Grid */}
                        <div className="pt-2">
                          <span className="block text-[10px] font-bold text-slate-400 mb-2">اضغط لتفعيل/إلغاء التوقيت الرمضاني لتواريخ محددة:</span>
                          <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-5 gap-2.5 max-h-[280px] overflow-y-auto p-1 border border-slate-100 rounded-2xl bg-slate-50/50">
                            {(Array.from(new Set(dailyRecords.map(r => r.dateKey))).sort() as any[]).map((dk: string) => {
                              const isRam = config.ramadanDates?.includes(dk);
                              return (
                                <button
                                  key={dk}
                                  onClick={() => toggleRamadanDate(dk)}
                                  className={`p-2.5 rounded-xl border text-right transition-all flex flex-col justify-between h-16 relative overflow-hidden group ${
                                    isRam
                                      ? 'bg-emerald-600 border-emerald-500 text-white shadow-sm'
                                      : 'bg-white hover:bg-slate-50 border-slate-200 text-slate-700'
                                  }`}
                                >
                                  <span className={`text-[10px] font-bold block ${isRam ? 'text-emerald-100' : 'text-slate-400'}`}>
                                    {isRam ? '🌙 دوام رمضان' : '🏫 دوام اعتيادي'}
                                  </span>
                                  <span className="text-xs font-bold block truncate mt-1">
                                    {formatArabicDate(dk)}
                                  </span>
                                  <div className="absolute left-1 bottom-1 opacity-10 group-hover:opacity-20 text-lg transition-opacity">
                                    {isRam ? '🌙' : '🏫'}
                                  </div>
                                </button>
                              );
                            })}
                          </div>
                        </div>
                      </div>
                    ) : (
                      <div className="bg-slate-50 p-6 rounded-2xl text-center border border-dashed border-slate-200">
                        <p className="text-xs text-slate-400 font-medium">يرجى رفع أو استيراد ملف البصمة أولاً لتظهر لك التواريخ هنا وتستطيع فرز وتخصيص أيام شهر رمضان المبارك.</p>
                      </div>
                    )}
                  </div>

                  {/* Weekend days settings */}
                  <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm space-y-3">
                    <div className="flex items-center justify-between gap-3 pb-2 border-b border-slate-100">
                      <div>
                        <h4 className="font-extrabold text-teal-950 text-sm">عطلة نهاية الأسبوع الرسمية (مستثناة تلقائياً من الغياب والدراسة)</h4>
                        <p className="text-xs text-slate-400 mt-0.5">
                          الأيام المحددة أدناه تعتبر أيام إجازة أسبوعية ولا يتم احتسابها كأيام غياب أو دوام لمنسوبي المدرسة
                        </p>
                      </div>
                      <span className="bg-teal-50 text-teal-800 text-[11px] font-extrabold px-3 py-1 rounded-full border border-teal-200">
                        الجمعة والسبت (معتمدة)
                      </span>
                    </div>

                    <div className="flex flex-wrap gap-2.5 pt-1">
                      {[
                        { dayIndex: 0, label: 'الأحد' },
                        { dayIndex: 1, label: 'الاثنين' },
                        { dayIndex: 2, label: 'الثلاثاء' },
                        { dayIndex: 3, label: 'الأربعاء' },
                        { dayIndex: 4, label: 'الخميس' },
                        { dayIndex: 5, label: 'الجمعة' },
                        { dayIndex: 6, label: 'السبت' },
                      ].map((item) => {
                        const isSelected = (config.weekend || [5, 6]).includes(item.dayIndex);
                        return (
                          <button
                            key={item.dayIndex}
                            type="button"
                            onClick={() => {
                              const currentWeekend = config.weekend || [5, 6];
                              const newWeekend = isSelected
                                ? currentWeekend.filter((d: number) => d !== item.dayIndex)
                                : [...currentWeekend, item.dayIndex];
                              setConfig({ ...config, weekend: newWeekend });
                            }}
                            className={`px-4 py-2 rounded-xl text-xs font-black transition-all border ${
                              isSelected
                                ? 'bg-teal-700 text-white border-teal-800 shadow-xs'
                                : 'bg-slate-50 hover:bg-slate-100 text-slate-600 border-slate-200'
                            }`}
                          >
                            {item.label} {isSelected ? '✓ (عطلة)' : ''}
                          </button>
                        );
                      })}
                    </div>
                  </div>

                  {/* Public Holidays Text block */}
                  <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm">
                    <h4 className="font-extrabold text-teal-950 text-sm mb-1.5">قائمة العطلات الرسمية المستثناة من المتابعة</h4>
                    <p className="text-xs text-slate-400 mb-3 leading-relaxed">
                      اكتب تاريخاً واحداً أو مدى زمنياً في كل سطر. سيقوم النظام باستبعاد هذه التواريخ تلقائياً من الإحصائيات لعدم احتساب غيابات على منسوبي المدرسة.
                    </p>
                    <textarea
                      rows={6}
                      value={config.holidays}
                      onChange={(e) => setConfig({ ...config, holidays: e.target.value })}
                      className="bg-slate-50 border border-slate-200 rounded-xl p-4 text-sm w-full outline-none font-mono focus:border-teal-500"
                      placeholder="YYYY-MM-DD&#10;YYYY-MM-DD إلى YYYY-MM-DD"
                    />
                  </div>

                  {/* Suspended School Days & Remote Learning Management Card */}
                  <div className="bg-white p-6 rounded-3xl border border-amber-200 shadow-sm space-y-5" dir="rtl">
                    <div className="flex items-start justify-between gap-4 pb-3 border-b border-amber-100">
                      <div className="flex items-center gap-3">
                        <div className="bg-amber-100 p-2.5 rounded-2xl text-amber-700">
                          <AlertTriangle className="w-6 h-6" />
                        </div>
                        <div>
                          <h4 className="font-extrabold text-slate-900 text-base">🌤️ أيام تعليق الدراسة الحضورية والتعليم عن بُعد (منصة مدرستي)</h4>
                          <p className="text-xs text-slate-400 mt-0.5">
                            استبعاد أيام تعليق الدوام الحضوري أو الظروف الجوية والطارئة من احتساب الغياب لحماية تقييمات وانضباط المعلمين
                          </p>
                        </div>
                      </div>
                      <span className="bg-amber-100 text-amber-800 text-[11px] font-extrabold px-3 py-1 rounded-full border border-amber-200 whitespace-nowrap">
                        حماية تقييم المعلمين ⭐
                      </span>
                    </div>

                    {/* Auto-detection Toggle */}
                    <div className="bg-amber-50/60 p-4 rounded-2xl border border-amber-200/60 flex items-center justify-between gap-4">
                      <div>
                        <h5 className="text-xs font-bold text-amber-950">الكشف التلقائي الذكي لأيام تعليق الدراسة</h5>
                        <p className="text-[11px] text-amber-800 mt-0.5 leading-relaxed">
                          عند تفعيل هذا الخيار، إذا كان هناك يوم دراسي (مثل يوم الخميس) لم تسجل فيه أي بصمات لجميع معلمي المدرسة (نسبة الحضور 0%)، سيتعرف عليه النظام تلقائياً كيوم تعليق دراسة ويستبعده فوراً من الغياب بدون عذر بدلاً من معاقبة جميع المعلمين.
                        </p>
                      </div>
                      <button
                        type="button"
                        dir="ltr"
                        onClick={() => setConfig({ ...config, autoDetectSuspendedDays: config.autoDetectSuspendedDays === false ? true : false })}
                        className={`relative inline-flex h-6 w-11 flex-shrink-0 cursor-pointer items-center rounded-full border-2 border-transparent transition-colors duration-200 ease-in-out focus:outline-none ${
                          config.autoDetectSuspendedDays !== false ? 'bg-amber-600' : 'bg-slate-200'
                        }`}
                      >
                        <span
                          aria-hidden="true"
                          className={`pointer-events-none inline-block h-5 w-5 transform rounded-full bg-white shadow-sm ring-0 transition duration-200 ease-in-out ${
                            config.autoDetectSuspendedDays !== false ? 'translate-x-5' : 'translate-x-0'
                          }`}
                        />
                      </button>
                    </div>

                    {/* Add Specific Suspended Date */}
                    <div>
                      <label className="block text-xs font-bold text-slate-600 mb-1.5">إضافة تاريخ تعليق دراسة محدد يدوياً</label>
                      <div className="flex flex-wrap items-center gap-2">
                        <input
                          type="date"
                          value={newSuspendedDateInput}
                          onChange={(e) => setNewSuspendedDateInput(e.target.value)}
                          className="bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-sm font-bold text-slate-700 outline-none focus:border-amber-500 h-[42px]"
                        />
                        <button
                          type="button"
                          onClick={handleAddSuspendedDate}
                          disabled={!newSuspendedDateInput}
                          className="bg-amber-700 hover:bg-amber-800 disabled:opacity-50 text-white font-bold px-4 py-2 rounded-xl text-xs flex items-center gap-1.5 transition-all shadow-xs h-[42px]"
                        >
                          <Plus className="w-4 h-4" />
                          <span>إضافة للتاريخ المستثنى</span>
                        </button>
                      </div>
                    </div>

                    {/* List of active suspended dates */}
                    <div className="space-y-2">
                      <span className="block text-xs font-bold text-slate-500">الأيام المستثناة حالياً من الغياب والدوام:</span>
                      
                      {/* Explicitly added dates */}
                      <div className="flex flex-wrap gap-2">
                        {(config.suspendedDates || []).length === 0 && autoDetectedSuspendedDays.length === 0 && (
                          <span className="text-xs text-slate-400 italic">لا توجد أيام تعليق مسجلة حالياً. يتم استبعاد الأيام تلقائياً عند عدم وجود بصمات بالكامل.</span>
                        )}

                        {(config.suspendedDates || []).map(dk => (
                          <span
                            key={dk}
                            className="bg-amber-50 border border-amber-300 text-amber-900 text-xs font-bold px-3 py-1.5 rounded-xl flex items-center gap-2 shadow-2xs"
                          >
                            <span>📅 {dayName(dk)} {dk}م ({toHijri(dk)})</span>
                            <span className="bg-amber-200 text-amber-900 text-[10px] px-1.5 py-0.2 rounded font-extrabold">محدد يدوياً</span>
                            <button
                              type="button"
                              onClick={() => toggleSuspendedDate(dk)}
                              className="text-amber-700 hover:text-rose-700 p-0.5 rounded transition-all"
                              title="حذف هذا الاستثناء"
                            >
                              <X className="w-3.5 h-3.5" />
                            </button>
                          </span>
                        ))}

                        {/* Auto-detected days */}
                        {autoDetectedSuspendedDays.map(dk => (
                          <span
                            key={dk}
                            className="bg-emerald-50 border border-emerald-300 text-emerald-900 text-xs font-bold px-3 py-1.5 rounded-xl flex items-center gap-2 shadow-2xs"
                          >
                            <span>🌤️ {dayName(dk)} {dk}م ({toHijri(dk)})</span>
                            <span className="bg-emerald-200 text-emerald-900 text-[10px] px-1.5 py-0.2 rounded font-extrabold">تم الكشف الذكي (0% بصمات)</span>
                            <button
                              type="button"
                              onClick={() => toggleSuspendedDate(dk)}
                              className="text-emerald-700 hover:text-rose-700 p-0.5 rounded transition-all"
                              title="تثبيت كيوم مستثنى يدوي"
                            >
                              <Check className="w-3.5 h-3.5" />
                            </button>
                          </span>
                        ))}
                      </div>
                    </div>
                  </div>

                  {/* Automatic Alert Settings */}
                  <div className="bg-white p-6 rounded-3xl border border-slate-200 shadow-sm space-y-4">
                    <div className="flex items-start justify-between gap-4">
                      <div className="flex gap-3">
                        <div className="bg-teal-50 p-2.5 rounded-2xl text-teal-600 mt-0.5">
                          <MessageSquare className="w-5 h-5 animate-pulse" />
                        </div>
                        <div>
                          <h4 className="font-extrabold text-slate-900 text-sm">تفعيل 'التنبيه التلقائي' الذكي لمنسوبي المدرسة</h4>
                          <p className="text-xs text-slate-400 mt-1 leading-relaxed">
                            عند تفعيل هذا الخيار، سيقوم النظام تلقائياً بتوليد وتجهيز رسائل واتساب جاهزة للإرسال في لوحة التحكم لجميع المنسوبين الذين لديهم تأخر متكرر أو انصراف مبكر/تلقائي، بدلاً من الدخول يدوياً لكل سجل. يمكنك الإرسال المتوالي والسريع بضغطة زر واحدة.
                          </p>
                        </div>
                      </div>
                      <button
                        type="button"
                        dir="ltr"
                        onClick={() => setConfig({ ...config, enableAutoAlert: !config.enableAutoAlert })}
                        className={`relative inline-flex h-6 w-11 flex-shrink-0 cursor-pointer items-center rounded-full border-2 border-transparent transition-colors duration-200 ease-in-out focus:outline-none ${
                          config.enableAutoAlert ? 'bg-teal-600' : 'bg-slate-200'
                        }`}
                      >
                        <span
                          aria-hidden="true"
                          className={`pointer-events-none inline-block h-5 w-5 transform rounded-full bg-white shadow-sm ring-0 transition duration-200 ease-in-out ${
                            config.enableAutoAlert ? 'translate-x-5' : 'translate-x-0'
                          }`}
                        />
                      </button>
                    </div>
                  </div>

                  {/* Warning limits */}
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                    <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm">
                      <h4 className="font-extrabold text-teal-950 text-sm mb-1.5">حد الإنذار الشفهي الأول (بالدقائق شهرياً)</h4>
                      <input
                        type="number"
                        value={config.warn1}
                        onChange={(e) => setConfig({ ...config, warn1: parseInt(e.target.value) || 30 })}
                        className="bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-sm font-bold text-slate-700 outline-none w-full"
                      />
                    </div>

                    <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm">
                      <h4 className="font-extrabold text-teal-950 text-sm mb-1.5">حد الإنذار الكتابي الثاني (بالدقائق شهرياً)</h4>
                      <input
                        type="number"
                        value={config.warn2}
                        onChange={(e) => setConfig({ ...config, warn2: parseInt(e.target.value) || 60 })}
                        className="bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-sm font-bold text-slate-700 outline-none w-full"
                      />
                    </div>
                  </div>

                  {/* Calendar Mode / Hijri & Gregorian Settings */}
                  <div className="bg-white p-6 rounded-3xl border border-teal-200 shadow-sm space-y-4" dir="rtl">
                    <div className="flex items-center gap-3 pb-3 border-b border-slate-100">
                      <div className="bg-teal-50 p-2.5 rounded-2xl text-teal-600">
                        <Calendar className="w-6 h-6" />
                      </div>
                      <div>
                        <h4 className="font-extrabold text-slate-900 text-base">نظام عرض التواريخ (الهجري والميلادي) في التقارير والجداول</h4>
                        <p className="text-xs text-slate-400 mt-0.5">اختر طريقة عرض التواريخ المناسبة لبيئة العمل المدرسية والتقارير الرسمية</p>
                      </div>
                    </div>

                    <div className="grid grid-cols-1 md:grid-cols-3 gap-3.5">
                      <button
                        type="button"
                        onClick={() => setConfig({ ...config, calendarMode: 'both' })}
                        className={`p-4 rounded-2xl border-2 text-right transition-all flex flex-col justify-between ${
                          (config.calendarMode || 'both') === 'both'
                            ? 'border-teal-600 bg-teal-50/60 shadow-sm'
                            : 'border-slate-200 hover:border-slate-300 bg-slate-50/50'
                        }`}
                      >
                        <div className="flex items-center justify-between mb-2">
                          <span className="text-xs font-black text-teal-950">التاريخ المزدوج (الميلادي + الهجري)</span>
                          <span className="text-[10px] bg-teal-100 text-teal-800 font-extrabold px-2 py-0.5 rounded-full">موصى به للمدارس ⭐</span>
                        </div>
                        <p className="text-[11px] text-slate-500 mb-3 leading-relaxed">
                          يعرض التاريخ الميلادي مع التاريخ الهجري جنباً إلى جنب في كافة الجداول، والتقارير، وملفات الطباعة وWord.
                        </p>
                        <div className="bg-white p-2 rounded-xl border border-teal-100 text-center">
                          <span className="text-xs font-bold text-slate-800 block">22/02/2026م</span>
                          <span className="text-[11px] font-bold text-teal-700 block">04 رمضان 1447هـ</span>
                        </div>
                      </button>

                      <button
                        type="button"
                        onClick={() => setConfig({ ...config, calendarMode: 'hijri' })}
                        className={`p-4 rounded-2xl border-2 text-right transition-all flex flex-col justify-between ${
                          config.calendarMode === 'hijri'
                            ? 'border-teal-600 bg-teal-50/60 shadow-sm'
                            : 'border-slate-200 hover:border-slate-300 bg-slate-50/50'
                        }`}
                      >
                        <div className="flex items-center justify-between mb-2">
                          <span className="text-xs font-black text-teal-950">التاريخ الهجري فقط</span>
                          <span className="text-[10px] bg-amber-100 text-amber-800 font-bold px-2 py-0.5 rounded-full">التقويم الهجري</span>
                        </div>
                        <p className="text-[11px] text-slate-500 mb-3 leading-relaxed">
                          يعرض التاريخ الهجري فقط (اليوم والشهر الهجري والسنة) في كامل شاشات البرنامج ومخرجات الطباعة.
                        </p>
                        <div className="bg-white p-2 rounded-xl border border-slate-200 text-center">
                          <span className="text-xs font-bold text-teal-700 block">04 رمضان 1447هـ</span>
                        </div>
                      </button>

                      <button
                        type="button"
                        onClick={() => setConfig({ ...config, calendarMode: 'gregorian' })}
                        className={`p-4 rounded-2xl border-2 text-right transition-all flex flex-col justify-between ${
                          config.calendarMode === 'gregorian'
                            ? 'border-teal-600 bg-teal-50/60 shadow-sm'
                            : 'border-slate-200 hover:border-slate-300 bg-slate-50/50'
                        }`}
                      >
                        <div className="flex items-center justify-between mb-2">
                          <span className="text-xs font-black text-teal-950">التاريخ الميلادي فقط</span>
                          <span className="text-[10px] bg-slate-100 text-slate-700 font-bold px-2 py-0.5 rounded-full">التقويم الميلادي</span>
                        </div>
                        <p className="text-[11px] text-slate-500 mb-3 leading-relaxed">
                          يعرض التاريخ الميلادي القياسي فقط في الجداول ومستندات الطباعة والتصدير.
                        </p>
                        <div className="bg-white p-2 rounded-xl border border-slate-200 text-center">
                          <span className="text-xs font-bold text-slate-800 block">22/02/2026م</span>
                        </div>
                      </button>
                    </div>
                  </div>

                  {/* Excel Formatter Tool Card */}
                  <div className="bg-white p-6 rounded-3xl border border-teal-200 shadow-md space-y-4" dir="rtl">
                    <div className="flex items-center gap-3 pb-3 border-b border-slate-100">
                      <div className="bg-teal-50 p-2.5 rounded-2xl text-teal-600">
                        <FileSpreadsheet className="w-6 h-6" />
                      </div>
                      <div>
                        <h4 className="font-extrabold text-slate-900 text-base">🛠️ أداة تنسيق وترتيب ملف بصمة حضور المنسوبين</h4>
                        <p className="text-xs text-slate-400 mt-0.5">قم برفع ملف البصمة الخام، وسيتولى النظام حذف الأعمدة غير الضرورية فوراً لتنظيمه</p>
                      </div>
                    </div>

                    <div className="bg-slate-50 border border-dashed border-slate-200 rounded-2xl p-5 text-center relative hover:bg-teal-50/20 transition-all group">
                      <input
                        type="file"
                        accept=".xlsx,.xls,.csv"
                        onChange={handleFormatExcelFile}
                        className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
                        id="formatter-file-input"
                      />
                      <div className="flex flex-col items-center justify-center space-y-2">
                        <div className="bg-white p-3 rounded-full border border-slate-200 text-teal-600 shadow-xs group-hover:scale-110 transition-transform">
                          <Upload className="w-6 h-6" />
                        </div>
                        <span className="text-xs font-bold text-slate-700">اضغط هنا أو قم بسحب وإفلات ملف البصمة لتنسيقه فوراً</span>
                        <span className="text-[10px] text-slate-400">يدعم صيغ Excel (.xlsx, .xls, .csv)</span>
                      </div>
                    </div>

                    {/* Columns to be deleted list info */}
                    <div className="bg-amber-50/50 border border-amber-200/60 rounded-2xl p-4">
                      <h5 className="text-xs font-bold text-amber-800 mb-2 flex items-center gap-1.5">
                        <span>⚠️</span>
                        <span>الأعمدة التي سيتم حذفها تلقائياً من الملف للتنظيف:</span>
                      </h5>
                      <div className="grid grid-cols-3 sm:grid-cols-4 md:grid-cols-6 gap-2 text-center text-[11px] font-bold text-slate-600">
                        {['F', 'H', 'J', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U'].map((col) => (
                          <span key={col} className="bg-white border border-slate-200/80 px-2 py-1 rounded-lg shadow-2xs">
                            عمود {col}
                          </span>
                        ))}
                      </div>
                    </div>

                    {/* Formatter status feedback */}
                    {formatStatus === 'processing' && (
                      <div className="bg-blue-50 border border-blue-100 p-4 rounded-2xl flex items-center gap-3 text-blue-700 text-xs font-bold">
                        <div className="animate-spin rounded-full h-4 w-4 border-2 border-blue-600 border-t-transparent" />
                        <span>جاري معالجة وتنسيق الملف الآن، يرجى الانتظار...</span>
                      </div>
                    )}

                    {formatStatus === 'success' && (
                      <div className="bg-emerald-50 border border-emerald-100 p-4 rounded-2xl flex items-start gap-3">
                        <div className="bg-emerald-100 p-1.5 rounded-full text-emerald-600 mt-0.5">
                          <Check className="w-4 h-4" />
                        </div>
                        <div>
                          <h5 className="text-xs font-black text-emerald-950">تم تنسيق وتنزيل الملف بنجاح! 🎉</h5>
                          <p className="text-[11px] text-emerald-800 mt-0.5">اسم الملف المعالج: <strong className="font-extrabold">{formatFileName}</strong> (إجمالي الصفوف المعالجة: {formatRowsCount} صفاً)</p>
                        </div>
                      </div>
                    )}

                    {formatStatus === 'error' && (
                      <div className="bg-rose-50 border border-rose-100 p-4 rounded-2xl flex items-start gap-3">
                        <div className="bg-rose-100 p-1.5 rounded-full text-rose-600 mt-0.5">
                          <X className="w-4 h-4" />
                        </div>
                        <div>
                          <h5 className="text-xs font-black text-rose-950">تعذر تنسيق الملف!</h5>
                          <p className="text-[11px] text-rose-800 mt-0.5">تأكد من اختيار ملف إكسل أو بصمة صحيح وصالح للملف المطلوب.</p>
                        </div>
                      </div>
                    )}
                  </div>
                </div>
              )}
            </>
          )}
        </div>

        {/* MODAL WINDOW FOR MESSAGES & ACTIONS */}
        {modalOpen && (
          <div className="fixed inset-0 bg-slate-950/60 flex items-center justify-center p-4 z-50 animate-fade-in">
            <div className="bg-white rounded-3xl max-w-lg w-full overflow-hidden shadow-2xl border border-slate-200">
              <div className="bg-slate-50 p-5 border-b border-slate-200 flex justify-between items-center">
                <h3 className="font-extrabold text-teal-950 text-sm">{modalTitle}</h3>
                <button
                  onClick={() => setModalOpen(false)}
                  className="p-1.5 hover:bg-slate-200 rounded-xl text-slate-400 hover:text-slate-600 transition-all"
                >
                  <X className="w-5 h-5" />
                </button>
              </div>

              <div className="p-5 space-y-4">
                <label className="block text-xs font-bold text-slate-400 mb-1">تعديل أو مراجعة نص الرسالة قبل الإرسال</label>
                <textarea
                  rows={8}
                  value={modalText}
                  onChange={(e) => setModalText(e.target.value)}
                  className="bg-slate-50 border border-slate-200 rounded-2xl p-4 text-sm w-full outline-none focus:border-teal-500 font-sans leading-relaxed text-slate-800"
                />

                <div className="flex flex-wrap items-center gap-2.5 pt-2">
                  <button
                    onClick={handleShareWhatsApp}
                    className="bg-teal-700 hover:bg-teal-800 text-white font-bold px-4 py-2.5 rounded-xl text-sm shadow-sm transition-all flex items-center gap-2"
                  >
                    <MessageSquare className="w-4 h-4" />
                    <span>إرسال عبر واتساب</span>
                  </button>

                  <button
                    onClick={handleCopyModalText}
                    className="bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold px-4 py-2.5 rounded-xl text-sm transition-all flex items-center gap-2 border border-slate-200"
                  >
                    <Copy className="w-4 h-4" />
                    <span>نسخ النص</span>
                  </button>

                  <button
                    onClick={() => setModalOpen(false)}
                    className="mr-auto text-xs font-bold text-slate-400 hover:text-slate-600 p-2"
                  >
                    إغلاق النافذة
                  </button>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* MODAL WINDOW FOR PENDING IMPORT CONFIRMATION */}
        {pendingImport && (
          <div className="fixed inset-0 bg-slate-950/60 flex items-center justify-center p-4 z-50 animate-fade-in" dir="rtl">
            <div className="bg-white rounded-3xl max-w-lg w-full overflow-hidden shadow-2xl border border-slate-200">
              <div className="p-6">
                <div className="flex items-center gap-3 mb-4 border-b border-slate-100 pb-3">
                  <div className="bg-teal-50 w-10 h-10 rounded-full flex items-center justify-center text-teal-600">
                    <Upload className="w-5 h-5" />
                  </div>
                  <div>
                    <h3 className="text-lg font-extrabold text-teal-950">مراجعة وحفظ ملف البصمة المستورد</h3>
                    <p className="text-xs text-slate-400">الرجاء تأكيد حفظ البيانات المتراكمة قبل الإضافة</p>
                  </div>
                </div>

                {/* File summary stats */}
                <div className="grid grid-cols-2 gap-3 mb-4">
                  <div className="bg-slate-50 p-3 rounded-2xl border border-slate-100">
                    <span className="block text-[10px] font-bold text-slate-400">إجمالي السجلات المقروءة</span>
                    <span className="text-lg font-bold text-slate-700">{pendingImport.stats.imported} سجل</span>
                  </div>
                  <div className="bg-slate-50 p-3 rounded-2xl border border-slate-100">
                    <span className="block text-[10px] font-bold text-slate-400">التواريخ المكتشفة بالملف</span>
                    <span className="text-sm font-bold text-slate-700 block truncate">
                      {pendingImport.fileDates.join(' ، ')}
                    </span>
                  </div>
                </div>

                <div className="bg-teal-50/50 p-3.5 rounded-2xl border border-teal-100 mb-4 text-xs space-y-1.5 text-teal-950">
                  <div className="font-bold flex items-center gap-1.5 mb-1 text-teal-900">
                    <CheckCircle2 className="w-4 h-4 text-teal-600" />
                    <span>تفاصيل تحليل الحضور والغياب:</span>
                  </div>
                  <div className="grid grid-cols-3 gap-2 text-center text-[11px] font-semibold mt-1">
                    <div className="bg-emerald-50 text-emerald-800 p-1.5 rounded-xl border border-emerald-100">
                      حضور: {pendingImport.stats.present}
                    </div>
                    <div className="bg-rose-50 text-rose-800 p-1.5 rounded-xl border border-rose-100">
                      غياب بدون عذر: {pendingImport.stats.absence}
                    </div>
                    <div className="bg-blue-50 text-blue-800 p-1.5 rounded-xl border border-blue-100">
                      غياب بعذر: {pendingImport.stats.excuse}
                    </div>
                  </div>
                </div>

                {/* Duplicates check and alert */}
                {pendingImport.duplicateDates.length > 0 ? (
                  <div className="bg-amber-50 p-4 rounded-2xl border border-amber-200 mb-6">
                    <div className="flex gap-2.5">
                      <AlertTriangle className="w-5 h-5 text-amber-600 shrink-0 mt-0.5" />
                      <div>
                        <h4 className="text-xs font-bold text-amber-900 mb-1">
                          ⚠️ تنبيه: بيانات هذا اليوم مسجلة مسبقاً!
                        </h4>
                        <p className="text-[11px] text-amber-700 leading-relaxed">
                          تم العثور بالفعل على سجلات سابقة للتواريخ: <strong className="text-amber-900 font-extrabold">{pendingImport.duplicateDates.join(' ، ')}</strong>.
                        </p>
                        <p className="text-[11px] text-amber-700 leading-relaxed mt-1">
                          الضغط على <strong>حفظ</strong> سيقوم <span className="underline decoration-dotted font-bold text-amber-900">بتحديث وتعديل</span> سجلات تلك الأيام واستبدالها بالكامل ببيانات الملف الجديد، بينما يتم الاحتفاظ بجميع الأيام والتواريخ القديمة الأخرى دون حذف.
                        </p>
                      </div>
                    </div>
                  </div>
                ) : (
                  <div className="bg-emerald-50/75 p-4 rounded-2xl border border-emerald-100 mb-6">
                    <div className="flex gap-2.5">
                      <Check className="w-5 h-5 text-emerald-600 shrink-0 mt-0.5" />
                      <div>
                        <h4 className="text-xs font-bold text-emerald-900 mb-1">
                          ✨ تجميع وحفظ تراكمي آمن
                        </h4>
                        <p className="text-[11px] text-emerald-700 leading-relaxed">
                          التواريخ الواردة بالملف جديدة كلياً ولا تتقاطع مع أي بيانات مسجلة مسبقاً.
                        </p>
                        <p className="text-[11px] text-emerald-700 leading-relaxed mt-1">
                          سيقوم البرنامج بإضافة وحفظ هذه البيانات الجديدة في قاعدة البيانات التراكمية لتنضم إلى سجل الموظفين بنجاح.
                        </p>
                      </div>
                    </div>
                  </div>
                )}

                {/* Action controls */}
                <div className="flex items-center gap-3">
                  <button
                    onClick={handleSavePendingImport}
                    className="flex-1 bg-teal-700 hover:bg-teal-800 text-white font-bold px-5 py-3 rounded-xl text-sm shadow-md transition-all flex items-center justify-center gap-2"
                  >
                    <Check className="w-4 h-4" />
                    <span>حفظ واعتماد البيانات</span>
                  </button>
                  <button
                    onClick={() => setPendingImport(null)}
                    className="bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold px-5 py-3 rounded-xl text-sm transition-all border border-slate-200"
                  >
                    إلغاء واستبعاد
                  </button>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* MODAL WINDOW FOR CLEAR DATA CONFIRMATION */}
        {showClearConfirm && (
          <div className="fixed inset-0 bg-slate-950/60 flex items-center justify-center p-4 z-50 animate-fade-in">
            <div className="bg-white rounded-3xl max-w-md w-full overflow-hidden shadow-2xl border border-slate-200">
              <div className="p-6 text-center">
                <div className="bg-rose-50 w-16 h-16 rounded-full flex items-center justify-center mx-auto mb-4 text-rose-600">
                  <Trash2 className="w-8 h-8" />
                </div>
                <h3 className="text-lg font-extrabold text-teal-950 mb-2">تفريغ كافة البيانات المستوردة</h3>
                <p className="text-slate-500 text-sm leading-relaxed mb-6">
                  هل أنت متأكد من رغبتك في حذف كافة السجلات والبيانات المستوردة؟ لا يمكن التراجع عن هذا الإجراء وسيتم تصفير لوحة المؤشرات.
                </p>
                <div className="flex items-center justify-center gap-3">
                  <button
                    onClick={handleClearData}
                    className="bg-rose-600 hover:bg-rose-700 text-white font-bold px-5 py-2.5 rounded-xl text-sm shadow-sm transition-all"
                  >
                    نعم، احذف البيانات
                  </button>
                  <button
                    onClick={() => setShowClearConfirm(false)}
                    className="bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold px-5 py-2.5 rounded-xl text-sm transition-all border border-slate-200"
                  >
                    تراجع وإلغاء
                  </button>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* MODAL WINDOW FOR EDITING/OVERRIDING RECORD */}
        {editingRecord && (
          <div className="fixed inset-0 bg-slate-950/60 flex items-center justify-center p-4 z-50 animate-fade-in" dir="rtl">
            <div className="bg-white rounded-3xl max-w-md w-full overflow-hidden shadow-2xl border border-slate-200">
              <div className="bg-slate-50 p-5 border-b border-slate-200 flex justify-between items-center">
                <div className="flex items-center gap-2">
                  <span className="text-xl">✏️</span>
                  <h3 className="font-extrabold text-teal-950 text-sm">تعديل وقت وحالة البصمة اليومية</h3>
                </div>
                <button
                  onClick={() => setEditingRecord(null)}
                  className="p-1.5 hover:bg-slate-200 rounded-xl text-slate-400 hover:text-slate-600 transition-all"
                >
                  <X className="w-5 h-5" />
                </button>
              </div>

              <div className="p-5 space-y-4">
                {/* Employee / Day summary */}
                <div className="bg-teal-50/50 p-3 rounded-2xl border border-teal-100/50">
                  <div className="text-xs font-bold text-slate-400 mb-1">اسم المعلم / الموظف</div>
                  <div className="text-sm font-black text-teal-950">{editingRecord.name}</div>
                  <div className="text-[11px] text-teal-800 font-bold mt-1">
                    🗓️ {dayName(editingRecord.date)} - {arDate(editingRecord.date)}م ({toHijri(editingRecord.date)})
                  </div>
                </div>

                {/* Status Selection */}
                <div>
                  <label className="block text-xs font-bold text-slate-500 mb-2">حالة الحضور اليومية</label>
                  <div className="grid grid-cols-3 gap-2">
                    <button
                      type="button"
                      onClick={() => setEditType('present')}
                      className={`py-2 px-3 rounded-xl text-xs font-bold transition-all border text-center ${
                        editType === 'present'
                          ? 'bg-emerald-600 border-emerald-600 text-white shadow-sm'
                          : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                      }`}
                    >
                      حاضر / بصمة
                    </button>
                    <button
                      type="button"
                      onClick={() => setEditType('absence')}
                      className={`py-2 px-3 rounded-xl text-xs font-bold transition-all border text-center ${
                        editType === 'absence'
                          ? 'bg-rose-600 border-rose-600 text-white shadow-sm'
                          : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                      }`}
                    >
                      غائب بدون عذر
                    </button>
                    <button
                      type="button"
                      onClick={() => setEditType('excuse')}
                      className={`py-2 px-3 rounded-xl text-xs font-bold transition-all border text-center ${
                        editType === 'excuse'
                          ? 'bg-blue-600 border-blue-600 text-white shadow-sm'
                          : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                      }`}
                    >
                      غائب بعذر
                    </button>
                  </div>
                </div>

                {/* Punch Times Inputs - only relevant if Present */}
                {editType === 'present' && (
                  <div className="grid grid-cols-2 gap-3 animate-fade-in">
                    <div>
                      <label htmlFor="manual-check-in" className="block text-xs font-bold text-slate-500 mb-1">وقت الحضور الفعلي</label>
                      <input
                        type="time"
                        id="manual-check-in"
                        value={editCheckIn}
                        onChange={(e) => setEditCheckIn(e.target.value)}
                        className="bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-sm w-full font-bold text-slate-700 outline-none focus:border-teal-500 font-sans"
                      />
                    </div>
                    <div>
                      <label htmlFor="manual-check-out" className="block text-xs font-bold text-slate-500 mb-1">وقت الانصراف الفعلي</label>
                      <input
                        type="time"
                        id="manual-check-out"
                        value={editCheckOut}
                        onChange={(e) => setEditCheckOut(e.target.value)}
                        className="bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-sm w-full font-bold text-slate-700 outline-none focus:border-teal-500 font-sans"
                      />
                    </div>
                  </div>
                )}

                {/* Custom Status Label / Notes */}
                <div>
                  <label htmlFor="manual-status-text" className="block text-xs font-bold text-slate-500 mb-1">
                    نص الحالة أو المبرر المخصص (اختياري)
                  </label>
                  <input
                    type="text"
                    id="manual-status-text"
                    placeholder={
                      editType === 'present'
                        ? 'مثال: حضور يدوي - عطل في جهاز البصمة'
                        : editType === 'excuse'
                        ? 'مثال: إجازة مرضية، عذر طبي معتمد'
                        : 'مثال: غياب بدون مبرر رسمي'
                    }
                    value={editStatusText}
                    onChange={(e) => setEditStatusText(e.target.value)}
                    className="bg-slate-50 border border-slate-200 rounded-xl px-3 py-2.5 text-xs w-full outline-none focus:border-teal-500 text-slate-700 font-medium font-sans"
                  />
                  {editType === 'excuse' && (
                    <div className="mt-2.5 space-y-1.5 animate-fade-in">
                      <span className="block text-[10px] font-bold text-slate-400">تحديد نوع الإجازة بسرعة:</span>
                      <div className="flex flex-wrap gap-1.5">
                        {[
                          { label: 'إجازة سنوية 🌴', value: 'إجازة سنوية' },
                          { label: 'إجازة طارئة 🚨', value: 'إجازة طارئة' },
                          { label: 'إجازة مرضية 🏥', value: 'إجازة مرضية' },
                          { label: 'مهمة عمل 💼', value: 'مهمة عمل' },
                        ].map((item) => (
                          <button
                            key={item.value}
                            type="button"
                            onClick={() => setEditStatusText(item.value)}
                            className={`px-2.5 py-1 rounded-lg text-[10px] font-bold transition-all border ${
                              editStatusText === item.value
                                ? 'bg-teal-50 border-teal-500 text-teal-700 shadow-sm'
                                : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                            }`}
                          >
                            {item.label}
                          </button>
                        ))}
                      </div>
                    </div>
                  )}
                  <p className="text-[10px] text-slate-400 mt-1 leading-relaxed">
                    يفيد هذا الخيار في تجاوز حالات تعليق جهاز البصمة أو الأخطاء التقنية التي تحول دون قيام الموظف بالبصم في الموعد المطلوب.
                  </p>
                </div>

                {/* Actions */}
                <div className="flex gap-2.5 pt-3 border-t border-slate-100">
                  <button
                    onClick={handleSaveEditedRecord}
                    className="flex-1 bg-teal-700 hover:bg-teal-800 text-white font-bold py-2.5 rounded-xl text-xs shadow-sm transition-all text-center"
                  >
                    حفظ التعديلات
                  </button>
                  <button
                    onClick={() => setEditingRecord(null)}
                    className="bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold px-4 py-2.5 rounded-xl text-xs transition-all border border-slate-200 text-center"
                  >
                    إلغاء
                  </button>
                </div>
              </div>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}
