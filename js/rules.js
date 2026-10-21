/**
 * Timesheet rules — pure functions, no storage. Same idea as the monthly / individual timesheet:
 *  - Lateness = whole minutes after the shift start (08:06:59 → 6). Grace, then tiers; each tier keeps its own
 *    counter per calendar month; from the tier's starting occurrence every time is deducted.
 *  - Missing check-in / check-out: own monthly counter, own deduction.
 *  - No punch on a working day → absence (for review); deducted only when the absence rule is switched on.
 *  - Early departure is recorded, never deducted; time after the shift end is shown, never counted as overtime.
 *  - Day hours = standard hours − deducted hours − deducted days × standard hours.
 *  - Net hours = sum of the period + «cut» hours + «missing» hours (both in the employee's favour);
 *    difference = net − standard hours of the period.
 */

export const CATEGORIES = {
  WORK: { label: 'دوام', color: null },
  WEEKEND: { label: 'راحة', color: 'C6E7B5' },
  L_DAYS: { label: 'إجازة', color: 'C9C9C9', hours: 'std' },
  B_TRIP: { label: 'مهمة عمل', color: '9CC3E6', hours: 'std' },
  EID: { label: 'عيد', color: 'CDBFE8', hours: 'std' },
  P_TIME: { label: 'إذن', color: 'FFF2A8', hours: 'std' },
  Z_TIME: { label: 'غياب', color: 'F4A6A6', hours: 0 },
  X_TIME: { label: 'استثناء', color: 'F8C58C', hours: 'std' },
  LWP: { label: 'بدون راتب', color: 'F6C4DC', hours: 0 },
  M_TIME: { label: 'مرضية', color: 'B99AD6', hours: 'std' },
  ABSENCE_REVIEW: { label: 'غياب للمراجعة', color: 'F4A6A6' },
}
/** Codes HR can put on a day (the legend of the sheet). */
export const DAY_CODES = ['L_DAYS', 'B_TRIP', 'EID', 'P_TIME', 'Z_TIME', 'X_TIME', 'LWP', 'M_TIME']
export const LEGEND = ['L_DAYS', 'B_TRIP', 'WEEKEND', 'EID', 'P_TIME', 'Z_TIME', 'X_TIME', 'LWP', 'M_TIME']
export const SOURCES = { lateness: 'تأخير', missing_punch: 'بصمة ناقصة', absence: 'غياب' }
export const WEEKDAYS = ['الأحد', 'الاثنين', 'الثلاثاء', 'الأربعاء', 'الخميس', 'الجمعة', 'السبت']
export const MONTHS = ['يناير', 'فبراير', 'مارس', 'أبريل', 'مايو', 'يونيو', 'يوليو', 'أغسطس', 'سبتمبر', 'أكتوبر', 'نوفمبر', 'ديسمبر']

export const DEFAULT_SETTINGS = {
  periodStartDay: 21,
  standardDayHours: 8,
  standardPeriodDays: 30,
  restDays: [5], // Friday
  shiftStart: '08:00',
  shiftEnd: '',
  /** Two punches closer than this are one punch (double press on the device). */
  duplicateMinutes: 2,
  lateness: {
    graceMinutes: 5,
    tiers: [
      { fromMinute: 6, toMinute: 20, unit: 'hours', value: 1, startingOccurrence: 3 },
      { fromMinute: 21, toMinute: 35, unit: 'hours', value: 2, startingOccurrence: 2 },
      { fromMinute: 36, toMinute: 95, unit: 'hours', value: 4, startingOccurrence: 1 },
      { fromMinute: 96, toMinute: null, unit: 'day', value: 1, startingOccurrence: 1 },
    ],
  },
  missingPunch: { active: true, unit: 'hours', value: 4, startingOccurrence: 1 },
  absence: { active: false, unit: 'day', value: 1 },
}

export const round2 = n => Math.round((Number(n) || 0) * 100) / 100
export const pad = n => String(n).padStart(2, '0')
export const toMin = hhmm => {
  const m = /^(\d{1,2}):(\d{2})/.exec(hhmm ?? '')
  return m ? Number(m[1]) * 60 + Number(m[2]) : null
}
export const hhmm = min => `${pad(Math.floor(min / 60))}:${pad(min % 60)}`

export function addDays(date, n) {
  const d = new Date(`${date}T12:00:00Z`)
  d.setUTCDate(d.getUTCDate() + n)
  return d.toISOString().slice(0, 10)
}
export function datesBetween(from, to) {
  const out = []
  for (let d = from; d <= to; d = addDays(d, 1)) out.push(d)
  return out
}
export const weekday = date => new Date(`${date}T12:00:00Z`).getUTCDay()

/** Period of a month label YYYY-MM: start day of the previous month → the day before it in this month. */
export function periodFor(label, settings) {
  const [y, m] = label.split('-').map(Number)
  const start = settings.periodStartDay
  if (start === 1) {
    const last = new Date(Date.UTC(y, m, 0)).getUTCDate()
    return { from: `${label}-01`, to: `${label}-${pad(last)}` }
  }
  return {
    from: new Date(Date.UTC(y, m - 2, start)).toISOString().slice(0, 10),
    to: new Date(Date.UTC(y, m - 1, start - 1)).toISOString().slice(0, 10),
  }
}

export function periodTitle(from, to) {
  const [, m] = to.split('-').map(Number)
  return `${MONTHS[m - 1]} ${to.slice(0, 4)}`
}

function tierIndex(lateness, minutesLate) {
  if (!(minutesLate > lateness.graceMinutes)) return -1
  return lateness.tiers.findIndex(t => minutesLate >= t.fromMinute && (t.toMinute === null || minutesLate <= t.toMinute))
}

/**
 * In / out of one day from the device punches (and HR's fill / correction, which win over the device).
 * One punch only: it is the check-in, unless the shift end is set and the punch is closer to it.
 */
function resolveInOut(punches, edits, settings, shiftStart) {
  const times = punches.map(p => toMin(p.slice(11, 16)))
  let inMin = null
  let outMin = null
  if (times.length >= 2) {
    inMin = times[0]
    outMin = times[times.length - 1]
  } else if (times.length === 1) {
    const end = toMin(settings.shiftEnd)
    const start = toMin(shiftStart)
    if (end !== null && start !== null && Math.abs(times[0] - end) < Math.abs(times[0] - start)) outMin = times[0]
    else inMin = times[0]
  }
  const deviceIn = inMin
  const deviceOut = outMin
  if (edits.in) inMin = toMin(edits.in.value)
  if (edits.out) outMin = toMin(edits.out.value)
  return { inMin, outMin, deviceIn, deviceOut }
}

/** Keep the last HR entry per field (each entry stays in the log). */
export function latestEdits(edits, deviceId, date) {
  const out = {}
  for (const e of edits) if (e.deviceId === deviceId && e.date === date) out[e.field] = e
  return out
}

/**
 * Every day of [from, to] for one employee, with the monthly counters run over whole calendar months
 * (a period from the 21st to the 20th spans two months).
 * @param {{ employee, punchesByDate: Map<string,string[]>, edits, marks, waivers, settings, from, to }} p
 */
export function computeDays(p) {
  const { employee, punchesByDate, edits, marks, waivers, settings, from, to } = p
  const std = settings.standardDayHours
  const shiftStart = employee.shiftStart || settings.shiftStart
  const calcFrom = `${from.slice(0, 7)}-01`
  const lastOfTo = new Date(Date.UTC(Number(to.slice(0, 4)), Number(to.slice(5, 7)), 0)).getUTCDate()
  const calcTo = `${to.slice(0, 7)}-${pad(lastOfTo)}`
  const counters = new Map()
  const days = []
  for (const date of datesBetween(calcFrom, calcTo)) {
    const punches = punchesByDate.get(date) ?? []
    const dayEdits = latestEdits(edits, employee.deviceId, date)
    const mark = marks.find(m => m.deviceId === employee.deviceId && m.date === date) ?? null
    const rest = settings.restDays.includes(weekday(date))
    const io = resolveInOut(punches, dayEdits, settings, shiftStart)
    const hasTime = io.inMin !== null || io.outMin !== null
    const d = {
      date,
      punches: punches.map(x => x.slice(11, 16)),
      inMin: io.inMin,
      outMin: io.outMin,
      deviceIn: io.deviceIn,
      deviceOut: io.deviceOut,
      edits: dayEdits,
      mark,
      status: 'work',
      category: 'WORK',
      lateMinutes: null,
      lateOccurrence: null,
      earlyMinutes: null,
      extraMinutes: null,
      missingIn: false,
      missingOut: false,
      items: [],
      hours: null,
    }
    if (mark) {
      d.status = 'mark'
      d.category = mark.code
    } else if (rest) {
      d.status = hasTime ? 'rest_worked' : 'rest'
      d.category = 'WEEKEND'
    } else if (!hasTime) {
      d.status = 'absence'
      d.category = 'ABSENCE_REVIEW'
      if (settings.absence.active) d.items.push({ source: 'absence', unit: settings.absence.unit, proposed: settings.absence.value })
    } else {
      const start = toMin(shiftStart)
      if (io.inMin !== null && start !== null) {
        d.lateMinutes = Math.max(0, io.inMin - start)
        const ti = tierIndex(settings.lateness, d.lateMinutes)
        if (ti >= 0) {
          const tier = settings.lateness.tiers[ti]
          const key = `${date.slice(0, 7)}|${ti}`
          const occ = (counters.get(key) ?? 0) + 1
          counters.set(key, occ)
          d.lateTier = ti
          d.lateOccurrence = occ
          if (occ >= tier.startingOccurrence) d.items.push({ source: 'lateness', unit: tier.unit, proposed: tier.value, occurrence: occ })
        }
      }
      d.missingIn = io.inMin === null
      d.missingOut = io.outMin === null
      if (d.missingIn || d.missingOut) {
        const key = `${date.slice(0, 7)}|missing`
        const occ = (counters.get(key) ?? 0) + 1
        counters.set(key, occ)
        const mp = settings.missingPunch
        if (mp.active && occ >= mp.startingOccurrence) d.items.push({ source: 'missing_punch', unit: mp.unit, proposed: mp.value, occurrence: occ })
      }
      const end = toMin(settings.shiftEnd)
      if (end !== null && io.outMin !== null) {
        if (io.outMin < end) d.earlyMinutes = end - io.outMin
        else if (io.outMin > end) d.extraMinutes = io.outMin - end
      }
    }
    for (const it of d.items) {
      const w = waivers.find(x => x.deviceId === employee.deviceId && x.date === date && x.source === it.source)
      it.final = w ? w.final : it.proposed
      it.waiver = w ?? null
    }
    d.proposedHours = round2(d.items.filter(i => i.unit === 'hours').reduce((s, i) => s + i.proposed, 0))
    d.proposedDays = round2(d.items.filter(i => i.unit === 'day').reduce((s, i) => s + i.proposed, 0))
    d.finalHours = round2(d.items.filter(i => i.unit === 'hours').reduce((s, i) => s + i.final, 0))
    d.finalDays = round2(d.items.filter(i => i.unit === 'day').reduce((s, i) => s + i.final, 0))
    if (d.status === 'mark') {
      const h = CATEGORIES[mark.code]?.hours
      d.hours = h === 'std' ? std : 0
    } else if (d.status === 'rest' || d.status === 'rest_worked') d.hours = std
    else d.hours = round2(std - d.finalHours - d.finalDays * std)
    d.adjusted = d.items.some(i => i.final !== i.proposed)
    days.push(d)
  }
  return days.filter(d => d.date >= from && d.date <= to)
}

/** Totals of one employee's period. */
export function totalsOf(days, carry, settings) {
  const std = settings.standardDayHours
  const c = { cut: Number(carry?.cut) || 0, missing: Number(carry?.missing) || 0, overtime: Number(carry?.overtime) || 0 }
  const sum = round2(days.reduce((s, d) => s + (d.hours ?? 0), 0))
  const totalH = round2(sum + c.cut + c.missing)
  const standardHours = std * settings.standardPeriodDays
  const ph = round2(days.reduce((s, d) => s + d.proposedHours, 0))
  const pd = round2(days.reduce((s, d) => s + d.proposedDays, 0))
  const fh = round2(days.reduce((s, d) => s + d.finalHours, 0))
  const fd = round2(days.reduce((s, d) => s + d.finalDays, 0))
  return {
    standardDays: settings.standardPeriodDays,
    standardHours,
    cut: c.cut,
    missing: c.missing,
    overtime: c.overtime,
    totalH,
    actH: round2(totalH - standardHours),
    proposedHours: ph,
    proposedDays: pd,
    finalHours: fh,
    finalDays: fd,
    calcDeduction: round2(ph + pd * std),
    finalDeduction: round2(fh + fd * std),
    lateDays: days.filter(d => d.lateOccurrence).length,
    missingDays: days.filter(d => d.missingIn || d.missingOut).length,
    absenceDays: days.filter(d => d.status === 'absence').length,
    editedDays: days.filter(d => d.edits.in || d.edits.out).length,
    extraMinutes: days.reduce((s, d) => s + (d.extraMinutes ?? 0), 0),
  }
}

/** Short note for a day (what the sheet shows in the note column). */
export function dayNote(d) {
  const parts = []
  if (d.mark) parts.push(CATEGORIES[d.mark.code]?.label + (d.mark.note ? ` — ${d.mark.note}` : ''))
  else if (d.status === 'rest' || d.status === 'rest_worked') parts.push('راحة')
  if (d.status === 'absence') parts.push('غياب للمراجعة')
  if (d.missingIn) parts.push('بصمة دخول ناقصة')
  if (d.missingOut) parts.push('بصمة خروج ناقصة')
  if (d.edits.in?.kind === 'fill' || d.edits.out?.kind === 'fill') parts.push('إعادة تعبئة')
  if (d.edits.in?.kind === 'correct' || d.edits.out?.kind === 'correct') parts.push('تصحيح')
  if (d.lateOccurrence) parts.push(`تأخير — المرة ${d.lateOccurrence}`)
  if (d.earlyMinutes) parts.push(`خروج مبكر ${d.earlyMinutes} د`)
  for (const i of d.items) if (i.final !== i.proposed) parts.push(`${SOURCES[i.source]}: ${fmt(i.proposed)} → ${fmt(i.final)}`)
  return parts.join(' · ')
}

export const fmt = n => (n === null || n === undefined || n === '' ? '' : Number.isInteger(Number(n)) ? String(Number(n)) : Number(n).toFixed(2).replace(/\.?0+$/, ''))

/** ZKTeco attlog (USB): «id<TAB>YYYY-MM-DD HH:MM:SS<TAB>…» — the later fields are ignored. */
export function parseAttlog(text) {
  const out = []
  let bad = 0
  for (const line of text.split(/\r?\n/)) {
    if (!line.trim()) continue
    const parts = line.split('\t')
    const id = (parts[0] ?? '').trim()
    const at = (parts[1] ?? '').trim()
    if (!/^\d+$/.test(id) || !/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}(:\d{2})?$/.test(at)) {
      bad++
      continue
    }
    out.push({ deviceId: String(Number(id)), at: at.length === 16 ? `${at}:00` : at })
  }
  return { rows: out, bad }
}

/** Punches of one employee by date, sorted, with double presses folded into one. */
export function punchesByDate(list, duplicateMinutes) {
  const m = new Map()
  for (const at of [...list].sort()) {
    const date = at.slice(0, 10)
    if (!m.has(date)) m.set(date, [])
    const day = m.get(date)
    const prev = day[day.length - 1]
    if (prev && toMin(at.slice(11, 16)) - toMin(prev.slice(11, 16)) < duplicateMinutes) continue
    day.push(at)
  }
  return m
}
