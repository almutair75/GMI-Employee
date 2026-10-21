/**
 * Exports — Excel (ExcelJS, ready to print: header, logo, colours, borders, frozen panes, A4/A3 landscape)
 * and PDF (a print page at A4 or A3 → «Save as PDF»). HR corrections always get their own page.
 */

import { CATEGORIES, LEGEND, SOURCES, WEEKDAYS, dayNote, fmt, hhmm, toMin } from './rules.js?v=4'

const COMPANY = 'جفت مارت العالمية'
const PURPLE = '2E1F5E'
const BURGUNDY = '9B2C5E'
const HEAD_FILL = 'ECE8F4'
const FILL_KIND = { fill: 'DCEBFA', correct: 'FDE3C8' }
const KIND_LABEL = { fill: 'إعادة تعبئة', correct: 'تصحيح' }
const SHORT_DAYS = ['أحد', 'إثن', 'ثلا', 'أرب', 'خمي', 'جمع', 'سبت']
const t = min => (min === null || min === undefined ? '' : hhmm(min))
const esc = v => String(v ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c])
const wd = date => new Date(`${date}T12:00:00Z`).getUTCDay()
const num = v => (v === null || v === undefined || v === '' ? null : Number(v))

let logoCache = null
async function logoBuffer() {
  if (!logoCache) logoCache = await (await fetch('assets/logo.png')).arrayBuffer()
  return logoCache
}

function fileName(base, meta, ext) {
  const [, m, d] = meta.at.slice(0, 10).split('-')
  const months = ['يناير', 'فبراير', 'مارس', 'أبريل', 'مايو', 'يونيو', 'يوليو', 'أغسطس', 'سبتمبر', 'أكتوبر', 'نوفمبر', 'ديسمبر']
  return `${base} - ${meta.title} - ${Number(d)} ${months[Number(m) - 1]} ${meta.at.slice(0, 4)}.${ext}`
}

function download(blob, name) {
  const a = document.createElement('a')
  a.href = URL.createObjectURL(blob)
  a.download = name
  document.body.appendChild(a)
  a.click()
  a.remove()
  setTimeout(() => URL.revokeObjectURL(a.href), 2000)
}

/** All HR entries for the exported employees and period, oldest first per employee. */
function correctionsOf(rows) {
  const out = []
  for (const r of rows) {
    for (const d of r.days) {
      for (const f of ['in', 'out']) {
        // every entry of the log for that day / field, not only the latest
        for (const e of r.allEdits?.filter(x => x.date === d.date && x.field === f) ?? []) out.push({ r, e })
      }
    }
  }
  return out
}

function waiversOf(rows) {
  return rows.flatMap(r => r.days.flatMap(d => d.items.filter(i => i.final !== i.proposed).map(i => ({ r, d, i }))))
}

function marksOf(rows) {
  return rows.flatMap(r => r.days.filter(d => d.mark).map(d => ({ r, d })))
}

// ————————————————— Excel —————————————————

function border(cell, color = 'A8A2B5') {
  const s = { style: 'thin', color: { argb: `FF${color}` } }
  cell.border = { top: s, left: s, bottom: s, right: s }
}
function fill(cell, hex) {
  if (hex) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${hex}` } }
}
function headCell(cell) {
  cell.font = { bold: true, color: { argb: `FF${PURPLE}` }, size: 10, name: 'Arial' }
  cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }
  fill(cell, HEAD_FILL)
  border(cell)
}
function bodyCell(cell, opts = {}) {
  cell.font = { size: 10, name: 'Arial', bold: Boolean(opts.bold), underline: Boolean(opts.underline), color: opts.color ? { argb: `FF${opts.color}` } : undefined }
  cell.alignment = { horizontal: opts.align ?? 'center', vertical: 'middle', wrapText: Boolean(opts.wrap) }
  if (opts.fill) fill(cell, opts.fill)
  border(cell)
}

async function sheetHeader(wb, ws, meta, title, lastCol, sub) {
  const logo = wb.addImage({ buffer: await logoBuffer(), extension: 'png' })
  ws.getRow(1).height = 30
  ws.getRow(2).height = 20
  ws.getRow(3).height = 18
  ws.mergeCells(1, 1, 3, Math.min(3, lastCol))
  ws.addImage(logo, { tl: { col: 0.15, row: 0.2 }, ext: { width: 190, height: 52 } })
  const titleFrom = Math.min(4, lastCol)
  ws.mergeCells(1, titleFrom, 1, lastCol)
  const c1 = ws.getCell(1, titleFrom)
  c1.value = `${COMPANY} — ${title}`
  c1.font = { bold: true, size: 15, color: { argb: `FF${BURGUNDY}` }, name: 'Arial' }
  c1.alignment = { horizontal: 'center', vertical: 'middle' }
  ws.mergeCells(2, titleFrom, 2, lastCol)
  const c2 = ws.getCell(2, titleFrom)
  c2.value = sub
  c2.font = { bold: true, size: 11, color: { argb: `FF${PURPLE}` }, name: 'Arial' }
  c2.alignment = { horizontal: 'center', vertical: 'middle' }
  ws.mergeCells(3, titleFrom, 3, lastCol)
  const c3 = ws.getCell(3, titleFrom)
  c3.value = `صدر بواسطة: ${meta.by} — ${meta.at}${meta.source ? ` — مصدر البصمة: ${meta.source}` : ''}`
  c3.font = { size: 9, color: { argb: 'FF6B637E' }, name: 'Arial' }
  c3.alignment = { horizontal: 'center', vertical: 'middle' }
}

function pageSetup(ws, size, orientation, headerRows) {
  ws.pageSetup = {
    paperSize: size === 'A3' ? 8 : 9,
    orientation,
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
    horizontalCentered: true,
    margins: { left: 0.25, right: 0.25, top: 0.4, bottom: 0.45, header: 0.2, footer: 0.2 },
    printTitlesRow: headerRows,
  }
  ws.headerFooter = { oddFooter: '&Cصفحة &P من &N' }
}

function legendRow(ws, row, fromCol) {
  let c = fromCol
  for (const k of LEGEND) {
    const cell = ws.getCell(row, c)
    cell.value = CATEGORIES[k].label
    bodyCell(cell, { fill: CATEGORIES[k].color })
    c++
  }
  const a = ws.getCell(row, c)
  a.value = 'غياب للمراجعة'
  bodyCell(a, { fill: CATEGORIES.ABSENCE_REVIEW.color })
  const f = ws.getCell(row + 1, fromCol)
  f.value = 'الرقم بخط عريض وتسطير = خصم معدّل · خلية زرقاء = إعادة تعبئة · خلية برتقالية = تصحيح'
  f.font = { size: 9, italic: true, name: 'Arial', color: { argb: 'FF5C5470' } }
}

async function monthlySheet(wb, meta, rows) {
  const ws = wb.addWorksheet('الكشف الشامل', { views: [{ rightToLeft: true, state: 'frozen', xSplit: 3, ySplit: 5 }] })
  const dates = rows[0]?.days.map(d => d.date) ?? []
  const heads = ['م', 'رقم البصمة', 'الاسم', 'الرقم الوظيفي', 'مقطوع (س)', 'ناقص (س)', ...dates.map(d => `${Number(d.slice(8))}\n${SHORT_DAYS[wd(d)]}`), 'إضافي معتمد', 'الخصم المحسوب (س)', 'الخصم النهائي (س)', 'الفرق', 'الإجمالي']
  const last = heads.length
  await sheetHeader(wb, ws, meta, `كشف الدوام الشهري${meta.branchName ? ` — ${meta.branchName}` : ''}`, last, `${meta.title} — من ${meta.period.from} إلى ${meta.period.to} — ${meta.settings.standardPeriodDays} يوم · ${meta.settings.standardPeriodDays * meta.settings.standardDayHours} ساعة`)
  const hr = ws.getRow(5)
  heads.forEach((h, i) => {
    const c = hr.getCell(i + 1)
    c.value = h
    headCell(c)
    if (i >= 6 && i < 6 + dates.length && meta.settings.restDays.includes(wd(dates[i - 6]))) fill(c, 'DDEFD3')
  })
  hr.height = 34
  rows.forEach((r, n) => {
    const row = ws.getRow(6 + n)
    const vals = [n + 1, num(r.employee.deviceId), r.employee.name, r.employee.jobNo || '', num(r.totals.cut) || null, num(r.totals.missing) || null]
    vals.forEach((v, i) => {
      const c = row.getCell(i + 1)
      c.value = v
      bodyCell(c, { align: i === 2 ? 'right' : 'center' })
    })
    r.days.forEach((d, i) => {
      const c = row.getCell(7 + i)
      c.value = d.hours
      const color = d.status === 'absence' ? CATEGORIES.ABSENCE_REVIEW.color : CATEGORIES[d.category]?.color
      const edited = d.edits.in || d.edits.out
      bodyCell(c, { fill: color ?? (edited ? FILL_KIND[(d.edits.in ?? d.edits.out).kind] : null), bold: d.adjusted, underline: d.adjusted })
      const note = dayNote(d)
      if (note) c.note = note
    })
    const tail = [r.totals.overtime || null, r.totals.calcDeduction, r.totals.finalDeduction, r.totals.actH, r.totals.totalH]
    tail.forEach((v, i) => {
      const c = row.getCell(7 + r.days.length + i)
      c.value = v
      bodyCell(c, { bold: i >= 3 || (i === 2 && r.totals.finalDeduction !== r.totals.calcDeduction), underline: i === 2 && r.totals.finalDeduction !== r.totals.calcDeduction, fill: i >= 3 ? 'F8F6FD' : null })
    })
    row.height = 18
  })
  ws.getColumn(1).width = 4
  ws.getColumn(2).width = 8
  ws.getColumn(3).width = 26
  ws.getColumn(4).width = 10
  ws.getColumn(5).width = 8
  ws.getColumn(6).width = 8
  for (let i = 0; i < dates.length; i++) ws.getColumn(7 + i).width = 4.6
  for (let i = 0; i < 5; i++) ws.getColumn(7 + dates.length + i).width = 10
  legendRow(ws, 6 + rows.length + 1, 3)
  ws.autoFilter = { from: { row: 5, column: 1 }, to: { row: 5 + rows.length, column: last } }
  pageSetup(ws, meta.size, 'landscape', '5:5')
}

async function individualSheet(wb, meta, r, name) {
  const ws = wb.addWorksheet(name, { views: [{ rightToLeft: true, state: 'frozen', ySplit: 8 }] })
  const heads = ['التاريخ', 'اليوم', 'البصمات', 'دخول', 'خروج', 'تأخير (د)', 'بعد نهاية الدوام (د)', 'الخصم', 'ملاحظة', 'الساعات']
  const last = heads.length
  await sheetHeader(wb, ws, meta, 'كشف الدوام الفردي', last, `${r.employee.name}${meta.branchName ? ` — ${meta.branchName}` : ''} — من ${meta.period.from} إلى ${meta.period.to}`)
  const facts = [
    ['الاسم', r.employee.name, 'رقم البصمة', num(r.employee.deviceId), 'الرقم الوظيفي', r.employee.jobNo || '—'],
    ['رقم الهوية', r.employee.idNumber || '—', 'تاريخ المباشرة', r.employee.hireDate || '—', 'بداية الدوام', r.employee.shiftStart || meta.settings.shiftStart],
  ]
  facts.forEach((f, i) => {
    const row = ws.getRow(5 + i)
    // label/value pairs spread over the 10 columns: 1-2, 3-4|5, 6-7|8, 9|10
    const spots = [[1, 1], [2, 3], [4, 4], [5, 6], [7, 8], [9, 10]]
    f.forEach((v, j) => {
      const [a, b] = spots[j]
      if (b > a) ws.mergeCells(5 + i, a, 5 + i, b)
      const c = row.getCell(a)
      c.value = v
      bodyCell(c, j % 2 === 0 ? { fill: HEAD_FILL, bold: true, color: PURPLE } : { align: 'right' })
    })
  })
  const hr = ws.getRow(8)
  heads.forEach((h, i) => {
    const c = hr.getCell(i + 1)
    c.value = h
    headCell(c)
  })
  hr.height = 30
  r.days.forEach((d, n) => {
    const row = ws.getRow(9 + n)
    const color = d.status === 'absence' ? CATEGORIES.ABSENCE_REVIEW.color : CATEGORIES[d.category]?.color
    const ded = d.items.map(i => `${SOURCES[i.source]} ${fmt(i.final)}${i.unit === 'day' ? ' يوم' : ' س'}${i.final !== i.proposed ? ` (المحسوب ${fmt(i.proposed)})` : ''}`).join('\n')
    const vals = [d.date, WEEKDAYS[wd(d.date)], d.punches.join('  '), t(d.inMin), t(d.outMin), d.lateMinutes || null, d.extraMinutes || null, ded, dayNote(d), d.hours]
    vals.forEach((v, i) => {
      const c = row.getCell(i + 1)
      c.value = v === '' ? null : v
      const kindFill = i === 3 && d.edits.in ? FILL_KIND[d.edits.in.kind] : i === 4 && d.edits.out ? FILL_KIND[d.edits.out.kind] : null
      bodyCell(c, { fill: i === 9 ? color : kindFill ?? (d.status === 'rest' ? 'F4F9F1' : null), wrap: i === 7 || i === 8 || i === 2, align: i === 8 ? 'right' : 'center', bold: i === 9 || (i === 3 && d.edits.in) || (i === 4 && d.edits.out) })
    })
  })
  ;[11, 9, 20, 8, 8, 9, 12, 16, 40, 9].forEach((w, i) => (ws.getColumn(i + 1).width = w))
  const tt = r.totals
  const sumStart = 9 + r.days.length + 1
  const summary = [
    ['أيام الفترة', tt.standardDays],
    ['الساعات الرسمية', tt.standardHours],
    ['مقطوع — إعادة خصم سابق بالخطأ (س)', tt.cut],
    ['ناقص — ساعات عمل لم تُحتسب (س)', tt.missing],
    ['الإضافي المعتمد (س)', tt.overtime],
    ['الخصم المحسوب (س)', tt.calcDeduction],
    ['الخصم النهائي (س)', tt.finalDeduction],
    ['الفرق (+/−)', tt.actH],
    ['صافي الساعات', tt.totalH],
  ]
  const st = ws.getCell(sumStart, 1)
  ws.mergeCells(sumStart, 1, sumStart, 4)
  st.value = 'الملخص'
  headCell(st)
  summary.forEach(([k, v], i) => {
    ws.mergeCells(sumStart + 1 + i, 1, sumStart + 1 + i, 3)
    const a = ws.getCell(sumStart + 1 + i, 1)
    a.value = k
    bodyCell(a, { align: 'right', fill: 'F8F6FD' })
    const b = ws.getCell(sumStart + 1 + i, 4)
    b.value = v
    bodyCell(b, { bold: true })
  })
  legendRow(ws, sumStart + 1, 6)
  pageSetup(ws, meta.size, 'portrait', '8:8')
}

async function correctionsSheet(wb, meta, rows) {
  const ws = wb.addWorksheet('التصحيحات', { views: [{ rightToLeft: true, state: 'frozen', ySplit: 5 }] })
  const heads = ['رقم البصمة', 'الاسم', 'اليوم', 'التاريخ', 'الخانة', 'النوع', 'وقت الجهاز', 'الوقت المعتمد', 'ملاحظة الموارد البشرية', 'بواسطة', 'وقت التسجيل']
  const last = heads.length
  await sheetHeader(wb, ws, meta, 'سجل التصحيحات وإعادة التعبئة', last, `${meta.title} — من ${meta.period.from} إلى ${meta.period.to}`)
  const hr = ws.getRow(5)
  heads.forEach((h, i) => {
    const c = hr.getCell(i + 1)
    c.value = h
    headCell(c)
  })
  const list = correctionsOf(rows)
  list.forEach(({ r, e }, n) => {
    const row = ws.getRow(6 + n)
    const vals = [num(r.employee.deviceId), r.employee.name, WEEKDAYS[wd(e.date)], e.date, e.field === 'in' ? 'دخول' : 'خروج', KIND_LABEL[e.kind], e.original ?? '—', e.value, e.note, e.by, e.at]
    vals.forEach((v, i) => {
      const c = row.getCell(i + 1)
      c.value = v
      bodyCell(c, { fill: i === 5 ? FILL_KIND[e.kind] : null, bold: i === 5 || i === 7, wrap: i === 8, align: i === 1 || i === 8 ? 'right' : 'center' })
    })
  })
  if (!list.length) {
    ws.mergeCells(6, 1, 6, last)
    const c = ws.getCell(6, 1)
    c.value = 'لا توجد تصحيحات أو إعادة تعبئة في هذه الفترة'
    bodyCell(c)
  }
  // Day codes and changed deductions, so the reviewer sees every manual decision
  let at = 6 + Math.max(1, list.length) + 2
  const waivers = waiversOf(rows)
  const marks = marksOf(rows)
  if (waivers.length || marks.length) {
    ws.mergeCells(at, 1, at, last)
    const h = ws.getCell(at, 1)
    h.value = 'قرارات أخرى من الموارد البشرية: تصنيف الأيام وتعديل الخصم'
    headCell(h)
    at++
    const h2 = ['رقم البصمة', 'الاسم', 'اليوم', 'التاريخ', 'القرار', 'التفاصيل', '', '', 'السبب / الملاحظة', 'بواسطة', 'وقت التسجيل']
    h2.forEach((x, i) => {
      const c = ws.getCell(at, i + 1)
      c.value = x
      headCell(c)
    })
    at++
    for (const { r, d } of marks) {
      ;[num(r.employee.deviceId), r.employee.name, WEEKDAYS[wd(d.date)], d.date, 'تصنيف اليوم', CATEGORIES[d.mark.code]?.label, '', '', d.mark.note, d.mark.by, d.mark.at].forEach((v, i) => {
        const c = ws.getCell(at, i + 1)
        c.value = v
        bodyCell(c, { fill: i === 5 ? CATEGORIES[d.mark.code]?.color : null, wrap: i === 8, align: i === 1 || i === 8 ? 'right' : 'center' })
      })
      at++
    }
    for (const { r, d, i: it } of waivers) {
      ;[num(r.employee.deviceId), r.employee.name, WEEKDAYS[wd(d.date)], d.date, 'تعديل خصم', `${SOURCES[it.source]}: ${fmt(it.proposed)} ← ${fmt(it.final)}`, '', '', it.waiver?.reason, it.waiver?.by, it.waiver?.at].forEach((v, i) => {
        const c = ws.getCell(at, i + 1)
        c.value = v
        bodyCell(c, { wrap: i === 8, align: i === 1 || i === 8 ? 'right' : 'center' })
      })
      at++
    }
  }
  ;[9, 24, 9, 11, 7, 12, 10, 11, 42, 14, 16].forEach((w, i) => (ws.getColumn(i + 1).width = w))
  pageSetup(ws, meta.size, 'landscape', '5:5')
}

export async function exportExcel(payload) {
  if (!window.ExcelJS) throw new Error('مكتبة Excel لم تُحمّل — تأكد من الاتصال بالإنترنت')
  const meta = { ...payload, size: 'A4' }
  const rows = withEdits(payload.rows)
  const wb = new window.ExcelJS.Workbook()
  wb.creator = 'GMI Employee'
  wb.created = new Date()
  let base = 'كشف الدوام'
  if (payload.kind === 'monthly') await monthlySheet(wb, meta, rows)
  if (payload.kind === 'individual') {
    await individualSheet(wb, meta, rows[0], 'كشف الموظف')
    base = `كشف ${rows[0].employee.name}`
  }
  if (payload.kind === 'corrections') base = 'سجل التصحيحات'
  await correctionsSheet(wb, meta, rows)
  const buf = await wb.xlsx.writeBuffer()
  download(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), fileName(base, meta, 'xlsx'))
}

/** Rows plus the full HR log of their days (the view keeps only the latest entry per field). */
function withEdits(rows) {
  let all = []
  try {
    all = JSON.parse(localStorage.getItem('gmi-employee:v1') ?? '{}').edits ?? []
  } catch {}
  return rows.map(r => ({ ...r, allEdits: all.filter(e => e.deviceId === r.employee.deviceId && e.date >= r.days[0]?.date && e.date <= r.days[r.days.length - 1]?.date).sort((a, b) => (a.at < b.at ? -1 : 1)) }))
}

// ————————————————— PDF (print page) —————————————————

function pdfHead(meta, title, sub) {
  return `<header class="hd"><img src="${new URL('assets/logo.png', location.href)}" alt=""><div class="ttl"><h1>${esc(COMPANY)} — ${esc(title)}</h1><p>${esc(sub)}</p></div><div class="meta">صدر بواسطة<br><b>${esc(meta.by)}</b><br>${esc(meta.at)}${meta.source ? `<br>مصدر البصمة: ${esc(meta.source)}` : ''}</div></header>`
}

function legendHtml() {
  return `<div class="legend">${LEGEND.map(k => `<span><i style="background:#${CATEGORIES[k].color}"></i>${CATEGORIES[k].label}</span>`).join('')}<span><i style="background:#${CATEGORIES.ABSENCE_REVIEW.color}"></i>غياب للمراجعة</span><span><u><b>8</b></u> خصم معدّل</span><span><i style="background:#${FILL_KIND.fill}"></i>إعادة تعبئة</span><span><i style="background:#${FILL_KIND.correct}"></i>تصحيح</span></div>`
}

function monthlyHtml(meta, rows) {
  const dates = rows[0]?.days.map(d => d.date) ?? []
  return `${pdfHead(meta, `كشف الدوام الشهري${meta.branchName ? ` — ${meta.branchName}` : ''}`, `${meta.title} — من ${meta.period.from} إلى ${meta.period.to} — ${meta.settings.standardPeriodDays} يوم · ${meta.settings.standardPeriodDays * meta.settings.standardDayHours} ساعة`)}
  <table class="ts"><colgroup><col class="c-s"><col class="c-id"><col class="c-nm"><col class="c-n"><col class="c-n">${dates.map(() => '<col>').join('')}<col class="c-n"><col class="c-w"><col class="c-w"><col class="c-n"><col class="c-w"></colgroup>
  <thead><tr><th>م</th><th>رقم البصمة</th><th>الاسم</th><th>مقطوع</th><th>ناقص</th>${dates.map(d => `<th class="${meta.settings.restDays.includes(wd(d)) ? 'rest' : ''}">${Number(d.slice(8))}<small>${SHORT_DAYS[wd(d)]}</small></th>`).join('')}<th>إضافي معتمد</th><th>الخصم المحسوب (س)</th><th>الخصم النهائي (س)</th><th>الفرق</th><th>الإجمالي</th></tr></thead>
  <tbody>${rows.map((r, n) => `<tr><td>${n + 1}</td><td>${esc(r.employee.deviceId)}</td><td class="nm">${esc(r.employee.name)}</td><td>${fmt(r.totals.cut || '')}</td><td>${fmt(r.totals.missing || '')}</td>
    ${r.days.map(d => {
      const color = d.status === 'absence' ? CATEGORIES.ABSENCE_REVIEW.color : CATEGORIES[d.category]?.color
      const edited = d.edits.in || d.edits.out
      const bg = color ?? (edited ? FILL_KIND[(d.edits.in ?? d.edits.out).kind] : null)
      return `<td${bg ? ` style="background:#${bg}"` : ''}${d.adjusted ? ' class="adj"' : ''}>${fmt(d.hours)}</td>`
    }).join('')}
    <td>${fmt(r.totals.overtime || '')}</td><td>${fmt(r.totals.calcDeduction)}</td><td class="${r.totals.finalDeduction !== r.totals.calcDeduction ? 'adj' : ''}">${fmt(r.totals.finalDeduction)}</td><td class="b">${fmt(r.totals.actH)}</td><td class="b">${fmt(r.totals.totalH)}</td></tr>`).join('')}</tbody></table>
  ${legendHtml()}
  <div class="sign"><div>مدير إدارة الموارد البشرية</div><div>مراجعة الإدارة المالية</div><div>اعتماد المدير العام</div></div>`
}

function individualHtml(meta, r) {
  const tt = r.totals
  return `${pdfHead(meta, 'كشف الدوام الفردي', `${meta.branchName ? `${meta.branchName} — ` : ''}من ${meta.period.from} إلى ${meta.period.to}`)}
  <table class="facts"><tr><th>الاسم</th><td>${esc(r.employee.name)}</td><th>رقم البصمة</th><td>${esc(r.employee.deviceId)}</td><th>الرقم الوظيفي</th><td>${esc(r.employee.jobNo || '—')}</td></tr>
  <tr><th>رقم الهوية</th><td>${esc(r.employee.idNumber || '—')}</td><th>تاريخ المباشرة</th><td>${esc(r.employee.hireDate || '—')}</td><th>بداية الدوام</th><td>${esc(r.employee.shiftStart || meta.settings.shiftStart)}</td></tr></table>
  <table class="dy"><thead><tr><th>التاريخ</th><th>البصمات</th><th>دخول</th><th>خروج</th><th>تأخير (د)</th><th>بعد النهاية (د)</th><th>الخصم</th><th>ملاحظة</th><th>الساعات</th></tr></thead>
  <tbody>${r.days.map(d => {
    const color = d.status === 'absence' ? CATEGORIES.ABSENCE_REVIEW.color : CATEGORIES[d.category]?.color
    const ded = d.items.map(i => `${SOURCES[i.source]} ${fmt(i.final)}${i.unit === 'day' ? ' يوم' : ' س'}`).join('، ')
    const io = f => {
      const e = d.edits[f]
      return `<td${e ? ` style="background:#${FILL_KIND[e.kind]}" class="b"` : ''}>${t(f === 'in' ? d.inMin : d.outMin)}</td>`
    }
    return `<tr class="${d.status === 'rest' ? 'rest' : ''}"><td class="dt">${WEEKDAYS[wd(d.date)]} ${d.date}</td><td class="fp">${d.punches.join(' ')}</td>${io('in')}${io('out')}<td>${d.lateMinutes || ''}</td><td>${d.extraMinutes || ''}</td><td>${esc(ded)}</td><td class="nt">${esc(dayNote(d))}</td><td class="b"${color ? ` style="background:#${color}"` : ''}>${fmt(d.hours)}</td></tr>`
  }).join('')}</tbody></table>
  <div class="bottom"><table class="sum"><tbody>
    <tr><td>أيام الفترة</td><td>${tt.standardDays}</td></tr><tr><td>الساعات الرسمية</td><td>${tt.standardHours}</td></tr>
    <tr><td>مقطوع — إعادة خصم سابق بالخطأ</td><td>${fmt(tt.cut)}</td></tr><tr><td>ناقص — ساعات عمل لم تُحتسب</td><td>${fmt(tt.missing)}</td></tr>
    <tr><td>الإضافي المعتمد</td><td>${fmt(tt.overtime)}</td></tr><tr><td>الخصم — المحسوب / النهائي (س)</td><td>${fmt(tt.calcDeduction)} / ${fmt(tt.finalDeduction)}</td></tr>
    <tr><td>الفرق (+/−)</td><td class="b">${fmt(tt.actH)}</td></tr><tr><td>صافي الساعات</td><td class="b">${fmt(tt.totalH)}</td></tr></tbody></table>
  ${legendHtml()}</div>`
}

function correctionsHtml(meta, rows) {
  const list = correctionsOf(rows)
  const waivers = waiversOf(rows)
  const marks = marksOf(rows)
  return `${pdfHead(meta, 'سجل التصحيحات وإعادة التعبئة', `${meta.title} — من ${meta.period.from} إلى ${meta.period.to}`)}
  ${list.length ? `<table class="cr"><thead><tr><th>رقم البصمة</th><th>الاسم</th><th>اليوم</th><th>الخانة</th><th>النوع</th><th>وقت الجهاز</th><th>الوقت المعتمد</th><th>ملاحظة الموارد البشرية</th><th>بواسطة</th><th>وقت التسجيل</th></tr></thead>
  <tbody>${list.map(({ r, e }) => `<tr><td>${esc(r.employee.deviceId)}</td><td class="nm">${esc(r.employee.name)}</td><td>${WEEKDAYS[wd(e.date)]} ${e.date}</td><td>${e.field === 'in' ? 'دخول' : 'خروج'}</td><td class="b" style="background:#${FILL_KIND[e.kind]}">${KIND_LABEL[e.kind]}</td><td>${esc(e.original ?? '—')}</td><td class="b">${esc(e.value)}</td><td class="nt">${esc(e.note)}</td><td>${esc(e.by)}</td><td>${esc(e.at)}</td></tr>`).join('')}</tbody></table>`
    : '<p class="none">لا توجد تصحيحات أو إعادة تعبئة في هذه الفترة.</p>'}
  ${waivers.length || marks.length ? `<h2>قرارات أخرى من الموارد البشرية</h2><table class="cr"><thead><tr><th>رقم البصمة</th><th>الاسم</th><th>اليوم</th><th>القرار</th><th>التفاصيل</th><th>السبب / الملاحظة</th><th>بواسطة</th><th>وقت التسجيل</th></tr></thead><tbody>
    ${marks.map(({ r, d }) => `<tr><td>${esc(r.employee.deviceId)}</td><td class="nm">${esc(r.employee.name)}</td><td>${WEEKDAYS[wd(d.date)]} ${d.date}</td><td>تصنيف اليوم</td><td style="background:#${CATEGORIES[d.mark.code]?.color}">${esc(CATEGORIES[d.mark.code]?.label)}</td><td class="nt">${esc(d.mark.note)}</td><td>${esc(d.mark.by)}</td><td>${esc(d.mark.at)}</td></tr>`).join('')}
    ${waivers.map(({ r, d, i }) => `<tr><td>${esc(r.employee.deviceId)}</td><td class="nm">${esc(r.employee.name)}</td><td>${WEEKDAYS[wd(d.date)]} ${d.date}</td><td>تعديل خصم</td><td>${SOURCES[i.source]}: ${fmt(i.proposed)} ← ${fmt(i.final)}</td><td class="nt">${esc(i.waiver?.reason)}</td><td>${esc(i.waiver?.by)}</td><td>${esc(i.waiver?.at)}</td></tr>`).join('')}
  </tbody></table>` : ''}`
}

function pdfCss(size, orientation, kind) {
  const big = size === 'A3'
  const fs = kind === 'monthly' ? (big ? 8.6 : 6.2) : big ? 11 : 8.4
  return `
@page { size: ${size} ${orientation}; margin: 9mm 8mm 11mm; @bottom-center { content: "صفحة " counter(page) " من " counter(pages); font: 8pt 'Cairo', sans-serif; color: #6B637E; } }
* { box-sizing: border-box; }
html, body { margin: 0; }
body { font-family: 'Cairo', sans-serif; color: #181228; font-size: ${fs}pt; -webkit-print-color-adjust: exact; print-color-adjust: exact; direction: rtl; }
.page { page-break-after: always; }
.page:last-child { page-break-after: auto; }
.hd { display: grid; grid-template-columns: auto 1fr auto; align-items: center; gap: 6mm; padding-bottom: 2.5mm; margin-bottom: 3mm; border-bottom: 1.2pt solid #${BURGUNDY}; }
.hd img { height: ${big ? 17 : 13}mm; }
.hd .ttl { text-align: center; }
.hd h1 { margin: 0; font-size: ${big ? 17 : 13.5}pt; color: #${PURPLE}; }
.hd p { margin: 1mm 0 0; font-size: ${big ? 11 : 9}pt; color: #${BURGUNDY}; font-weight: 700; }
.hd .meta { font-size: ${big ? 9 : 7.5}pt; color: #5C5470; text-align: left; line-height: 1.5; }
table { border-collapse: collapse; width: 100%; }
thead { display: table-header-group; }
tr { break-inside: avoid; page-break-inside: avoid; }
th, td { border: .45pt solid #A8A2B5; padding: .5mm .6mm; text-align: center; vertical-align: middle; }
th { background: #${HEAD_FILL}; color: #${PURPLE}; font-weight: 700; line-height: 1.2; }
th.rest { background: #DDEFD3; }
th small { display: block; font-weight: 400; font-size: .82em; }
.ts { table-layout: fixed; }
.ts col.c-s { width: ${big ? 7 : 5}mm; } .ts col.c-id { width: ${big ? 12 : 9}mm; } .ts col.c-nm { width: ${big ? 48 : 33}mm; }
.ts col.c-n { width: ${big ? 11 : 8}mm; } .ts col.c-w { width: ${big ? 14 : 10}mm; }
.ts td { white-space: nowrap; overflow: hidden; }
td.nm { text-align: right; padding-inline: 1.2mm; text-overflow: ellipsis; }
td.adj { font-weight: 700; text-decoration: underline; }
td.b { font-weight: 700; }
.legend { display: flex; flex-wrap: wrap; gap: 1.2mm 4mm; margin-top: 2.5mm; font-size: .92em; align-items: center; }
.legend i { display: inline-block; width: 4.5mm; height: 3mm; border: .4pt solid #777; vertical-align: middle; margin-inline-end: 1mm; }
.sign { display: grid; grid-template-columns: repeat(3, 1fr); gap: 12mm; margin-top: 9mm; text-align: center; font-size: 1.1em; break-inside: avoid; }
.sign div { border-top: .7pt solid #555; padding-top: 1.5mm; min-height: 14mm; }
.facts th { width: 13%; } .facts td { text-align: right; padding-inline: 2mm; }
.facts { margin-bottom: 3mm; }
.dy td.fp { direction: ltr; font-size: .9em; } .dy td.nt, .cr td.nt { text-align: right; } .dy td.dt { white-space: nowrap; }
.dy tr.rest td { background: #F4F9F1; }
.bottom { display: grid; grid-template-columns: 1fr 1fr; gap: 6mm; margin-top: 3mm; align-items: start; break-inside: avoid; }
.sum td:first-child { text-align: right; background: #F8F6FD; } .sum td:last-child { width: 26%; font-weight: 700; }
h2 { font-size: 1.25em; color: #${PURPLE}; margin: 5mm 0 2mm; }
.none { text-align: center; color: #5C5470; padding: 8mm; font-size: 1.2em; }`
}

/** Print page at A4 or A3 (the browser's «Save as PDF»). Corrections always start on their own page. */
export async function exportPdf(payload, size) {
  const rows = withEdits(payload.rows)
  const meta = { ...payload, size }
  let pages = []
  let orientation = 'landscape'
  if (payload.kind === 'monthly') pages = [monthlyHtml(meta, rows), correctionsHtml(meta, rows)]
  else if (payload.kind === 'individual') {
    orientation = 'portrait'
    pages = [individualHtml(meta, rows[0]), correctionsHtml(meta, rows)]
  } else pages = [correctionsHtml(meta, rows)]
  const title = payload.kind === 'individual' ? `كشف ${rows[0].employee.name}` : payload.kind === 'monthly' ? 'كشف الدوام' : 'سجل التصحيحات'
  const html = `<!doctype html><html lang="ar" dir="rtl"><head><meta charset="utf-8"><title>${esc(fileName(title, meta, 'pdf').replace(/\.pdf$/, ''))}</title>
  <link rel="preconnect" href="https://fonts.googleapis.com"><link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;700&display=swap" rel="stylesheet">
  <style>${pdfCss(size, orientation, payload.kind)}</style></head><body>${pages.map(p => `<section class="page">${p}</section>`).join('')}</body></html>`
  const frame = document.createElement('iframe')
  frame.style.cssText = 'position:fixed;inset-inline-end:0;bottom:0;width:0;height:0;border:0;visibility:hidden'
  document.body.appendChild(frame)
  const doc = frame.contentDocument
  doc.open()
  doc.write(html)
  doc.close()
  await new Promise(res => (frame.contentWindow.document.readyState === 'complete' ? res() : frame.contentWindow.addEventListener('load', res, { once: true })))
  try {
    await frame.contentWindow.document.fonts.ready
  } catch {}
  await Promise.all([...doc.images].map(img => (img.complete ? null : new Promise(r => (img.onload = img.onerror = r)))))
  // The page title becomes the suggested PDF file name in most browsers
  const prevTitle = document.title
  document.title = doc.title
  frame.contentWindow.focus()
  frame.contentWindow.print()
  setTimeout(() => {
    document.title = prevTitle
    frame.remove()
  }, 1500)
}

// ————————————————— Employees file —————————————————

const COLS = [
  ['deviceId', ['رقم البصمة', 'رقم الجهاز', 'device', 'no', 'id']],
  ['name', ['الاسم', 'الاسم الكامل', 'name']],
  ['jobNo', ['الرقم الوظيفي', 'job', 'jobno']],
  ['idNumber', ['رقم الهوية', 'الهوية', 'id number', 'iqama']],
  ['hireDate', ['تاريخ المباشرة', 'المباشرة', 'hire', 'start date']],
  ['branch', ['الفرع', 'branch']],
  ['shiftStart', ['بداية الدوام', 'shift', 'shift start']],
]

export async function employeesTemplate() {
  const wb = new window.ExcelJS.Workbook()
  const ws = wb.addWorksheet('الموظفون', { views: [{ rightToLeft: true, state: 'frozen', ySplit: 1 }] })
  const heads = ['رقم البصمة', 'الاسم', 'الرقم الوظيفي', 'رقم الهوية', 'تاريخ المباشرة', 'الفرع', 'بداية الدوام']
  heads.forEach((h, i) => {
    const c = ws.getCell(1, i + 1)
    c.value = h
    headCell(c)
  })
  ;[12, 34, 14, 16, 16, 20, 14].forEach((w, i) => (ws.getColumn(i + 1).width = w))
  ws.getCell(2, 1).value = 1
  ws.getCell(2, 2).value = 'اكتب الاسم الكامل هنا'
  ws.getCell(2, 5).value = '2024-01-01'
  ws.getCell(2, 6).value = 'الفرع الرئيسي'
  const buf = await wb.xlsx.writeBuffer()
  download(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), 'نموذج الموظفين.xlsx')
}

function cellText(v) {
  if (v === null || v === undefined) return ''
  if (v instanceof Date) {
    // Excel time-only cells come as 1899-12-30; dates as midnight
    if (v.getUTCFullYear() < 1901) return `${String(v.getUTCHours()).padStart(2, '0')}:${String(v.getUTCMinutes()).padStart(2, '0')}`
    return v.toISOString().slice(0, 10)
  }
  if (typeof v === 'object') return String(v.text ?? v.result ?? v.richText?.map(x => x.text).join('') ?? '')
  return String(v).trim()
}

export async function readEmployeesFile(file) {
  let table = []
  if (/\.csv$/i.test(file.name)) {
    const text = (await file.text()).replace(/^﻿/, '')
    table = text.split(/\r?\n/).filter(l => l.trim()).map(l => l.split(/[,;\t]/).map(x => x.trim().replace(/^"|"$/g, '')))
  } else {
    const wb = new window.ExcelJS.Workbook()
    await wb.xlsx.load(await file.arrayBuffer())
    const ws = wb.worksheets[0]
    ws.eachRow({ includeEmpty: false }, row => {
      const r = []
      for (let i = 1; i <= Math.max(7, row.cellCount); i++) r.push(cellText(row.getCell(i).value))
      table.push(r)
    })
  }
  if (!table.length) throw new Error('الملف فاضي')
  const head = table[0].map(h => h.toLowerCase())
  const idx = {}
  for (const [key, names] of COLS) idx[key] = head.findIndex(h => names.some(n => h === n.toLowerCase() || h.includes(n.toLowerCase())))
  if (idx.deviceId < 0 || idx.name < 0) throw new Error('لازم يكون في الملف عمود «رقم البصمة» وعمود «الاسم»')
  const out = []
  for (const r of table.slice(1)) {
    const id = String(r[idx.deviceId] ?? '').replace(/\D/g, '')
    const name = String(r[idx.name] ?? '').trim()
    if (!id || !name) continue
    const rec = { deviceId: String(Number(id)), name }
    for (const k of ['jobNo', 'idNumber', 'hireDate', 'branch', 'shiftStart']) rec[k] = idx[k] >= 0 ? String(r[idx[k]] ?? '').trim() : ''
    if (rec.shiftStart && toMin(rec.shiftStart) === null) rec.shiftStart = ''
    out.push(rec)
  }
  if (!out.length) throw new Error('ما لقيت صفوف فيها رقم واسم')
  return out
}
