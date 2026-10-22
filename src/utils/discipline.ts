import { EmployeeSummary } from '../types';

export function disciplineScore(m: EmployeeSummary): number {
  const latePenalty = (+m.lateMins || 0) * 0.25; // 0.25 point deducted per minute late
  const earlyPenalty = (+m.earlyMins || 0) * 0.25; // 0.25 point deducted per minute of early checkout
  const autoPenalty = (+m.autoCount || 0) * 2.0; // 2 points deducted per auto checkout
  const missingPenalty = (+m.missingCheckoutCount || 0) * 3.0; // 3 points deducted per missing checkout
  const wrongTimePenalty = (+m.wrongTimeExitsCount || 0) * 3.0; // 3 points deducted per wrong-time / morning checkout (خلل في تسجيل الانصراف)
  const emergencyNoExitPenalty = (+m.emergencyNoExitCount || 0) * 3.0; // 3 points deducted per emergency permission without departure punch
  const absencePenalty = (+m.absence || 0) * 10.0; // 10 points deducted per unexcused absence day

  const raw = 100 - latePenalty - earlyPenalty - autoPenalty - missingPenalty - wrongTimePenalty - emergencyNoExitPenalty - absencePenalty;
  return Math.max(0, Math.min(100, Math.round(raw)));
}

export function disciplineReason(m: EmployeeSummary): string {
  const parts: string[] = [];
  if (m.absence > 0) {
    parts.push(`غياب بدون عذر: ${m.absence} يوم`);
  }
  if (m.lateMins > 0) {
    parts.push(`حضور متأخر: ${m.lateCount} مرة / ${m.lateMins} دقيقة`);
  }
  if (m.earlyMins > 0 || m.earlyCount > 0) {
    parts.push(`خروج قبل نهاية الدوام: ${m.earlyCount} مرة / ${m.earlyMins} دقيقة`);
  }
  if (m.autoCount > 0) {
    parts.push(`انصراف تلقائي: ${m.autoCount} مرة`);
  }
  if (m.missingCheckoutCount > 0) {
    parts.push(`لا يوجد بصمة انصراف: ${m.missingCheckoutCount} مرة`);
  }
  if ((m.wrongTimeExitsCount || 0) > 0) {
    parts.push(`انصراف في غير وقت الفعلي (بصمة صباحية/اليوم التالي): ${m.wrongTimeExitsCount} مرة`);
  }
  if ((m.emergencyNoExitCount || 0) > 0) {
    parts.push(`استئذان طارئ بدون بصمة انصراف: ${m.emergencyNoExitCount} مرة`);
  }
  return parts.length ? parts.join('، ') : 'لا توجد مخالفات محتسبة في التقييم';
}

export function disciplineClassify(m: EmployeeSummary): {
  score: number;
  reason: string;
  disciplineCategory: 'ideal' | 'notes' | 'follow' | 'admin' | 'none';
  disciplineLabel: string;
} {
  const score = disciplineScore(m);
  const reason = disciplineReason(m);

  if (m.work <= 0) {
    return {
      score,
      reason,
      disciplineCategory: 'none',
      disciplineLabel: 'لا توجد بيانات',
    };
  }

  const attendanceRate = m.work ? Math.round(((m.present + m.excuse) / m.work) * 100) : 0;

  // Strict Fair Outstanding (Ideal) Criteria:
  // - Unexcused absences == 0
  // - Lateness minutes and count == 0
  // - Early checkout minutes and count == 0
  // - Auto checkout count == 0
  // - Missing checkout count == 0
  // - Wrong-time checkout count == 0
  // - Emergency permission without checkout count == 0
  // - Attendance rate >= 90%
  let disciplineCategory: 'ideal' | 'notes' | 'follow' | 'admin' = 'notes';
  let disciplineLabel = 'منضبط مع ملاحظات';

  if (
    score >= 90 &&
    (m.absence || 0) === 0 &&
    (m.autoCount || 0) === 0 &&
    (m.missingCheckoutCount || 0) === 0 &&
    (m.wrongTimeExitsCount || 0) === 0 &&
    (m.emergencyNoExitCount || 0) === 0
  ) {
    disciplineCategory = 'ideal';
    disciplineLabel = 'منضبط متميز';
  } else if (score >= 80) {
    disciplineCategory = 'notes';
    disciplineLabel = 'منضبط مع ملاحظات';
  } else if (score >= 70) {
    disciplineCategory = 'follow';
    disciplineLabel = 'يحتاج متابعة';
  } else {
    disciplineCategory = 'admin';
    disciplineLabel = 'يحتاج إجراء إداري';
  }

  return {
    score,
    reason,
    disciplineCategory,
    disciplineLabel,
  };
}
