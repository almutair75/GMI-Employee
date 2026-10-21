import {
  CATEGORIES, DAY_CODES, LEGEND, SOURCES, WEEKDAYS, MONTHS, DEFAULT_SETTINGS,
  computeDays, totalsOf, dayNote, fmt, hhmm, toMin, parseAttlog, punchesByDate, periodFor, periodTitle, round2, pad,
} from './rules.js'
import { USERS, db, save, login, logout, currentUser, nowStamp, replaceAll, resetAll } from './store.js'
import { exportExcel, exportPdf, employeesTemplate, readEmployeesFile } from './exports.js'

const $ = (sel, root = document) => root.querySelector(sel)
const esc = v => String(v ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c])
const app = $('#app')

const ui = {
  tab: 'sheet',
  label: null, // YYYY-MM
  custom: null, // { from, to }
  employee: null, // deviceId for the individual sheet
  search: '',
  importPreview: null,
}

// —— helpers ——

function toast(msg, kind = 'ok') {
  const t = document.createElement('div')
  t.className = `toast toast--${kind}`
  t.textContent = msg
  document.body.appendChild(t)
  setTimeout(() => t.classList.add('show'), 10)
  setTimeout(() => {
    t.classList.remove('show')
    setTimeout(() => t.remove(), 300)
  }, 3200)
}

function persist() {
  if (!save()) toast('تعذّر الحفظ في المتصفح — المساحة ممتلئة؟ صدّر نسخة احتياطية.', 'err')
}

function period() {
  const s = db().settings
  if (ui.custom) return ui.custom
  return periodFor(ui.label, s)
}
const carryKey = p => `${p.from}|${p.to}`

function defaultLabel() {
  const today = new Date()
  const s = db().settings
  // The period that contains today: after the start day it belongs to next month's label
  let y = today.getFullYear()
  let m = today.getMonth() + 1
  if (s.periodStartDay > 1 && today.getDate() >= s.periodStartDay) {
    m++
    if (m > 12) {
      m = 1
      y++
    }
  }
  return `${y}-${pad(m)}`
}

function labelsAround() {
  const out = []
  const [y, m] = defaultLabel().split('-').map(Number)
  for (let i = 0; i < 24; i++) {
    const d = new Date(Date.UTC(y, m - 1 - i, 1))
    out.push(`${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}`)
  }
  return out
}

function employeesSorted() {
  return [...db().employees].sort((a, b) => Number(a.deviceId) - Number(b.deviceId))
}

/** Days + totals of one employee for the current period. */
export function viewOf(emp, p = period()) {
  const s = db().settings
  const days = computeDays({
    employee: emp,
    punchesByDate: punchesByDate(db().punches[emp.deviceId] ?? [], s.duplicateMinutes),
    edits: db().edits,
    marks: db().marks,
    waivers: db().waivers,
    settings: s,
    from: p.from,
    to: p.to,
  })
  const carry = db().carry[carryKey(p)]?.[emp.deviceId] ?? {}
  return { employee: emp, days, carry, totals: totalsOf(days, carry, s) }
}

const dayLabel = date => `${WEEKDAYS[new Date(`${date}T12:00:00Z`).getUTCDay()]} ${date.slice(8)}/${date.slice(5, 7)}`
const t = min => (min === null || min === undefined ? '' : hhmm(min))

// —— shell ——

function render() {
  const user = currentUser()
  if (!user) return renderLogin()
  if (!ui.label) ui.label = defaultLabel()
  const tabs = [
    ['sheet', 'الكشف الشهري'],
    ['individual', 'كشف الموظف'],
    ['corrections', 'سجل التصحيحات'],
    ['employees', 'الموظفون'],
    ['import', 'استيراد البصمة'],
    ['settings', 'الإعدادات'],
  ]
  app.innerHTML = `
  <header class="top">
    <div class="brand"><img src="assets/logo.png" alt="Gift Mart"><div><b>GMI Employee</b><small>كشف الدوام — تجربة</small></div></div>
    <div class="who"><span>${esc(user.name)}</span><button class="btn btn--ghost btn--sm" id="logout">خروج</button></div>
  </header>
  <nav class="tabs" role="tablist">${tabs.map(([k, l]) => `<button role="tab" class="tab${ui.tab === k ? ' is-active' : ''}" data-tab="${k}">${l}</button>`).join('')}</nav>
  <main id="view"></main>`
  $('#logout').onclick = () => {
    logout()
    render()
  }
  app.querySelectorAll('[data-tab]').forEach(b => (b.onclick = () => {
    ui.tab = b.dataset.tab
    ui.search = ''
    render()
  }))
  const view = $('#view')
  try {
    ;({ sheet: renderSheet, individual: renderIndividual, corrections: renderCorrections, employees: renderEmployees, import: renderImport, settings: renderSettings })[ui.tab](view, user)
  } catch (err) {
    console.error(err)
    view.innerHTML = `<div class="card empty"><h2>تعذّر عرض هذه الصفحة</h2><p>${esc(err?.message ?? err)}</p></div>`
  }
}

function renderLogin() {
  app.innerHTML = `
  <div class="login">
    <form class="card login__card" id="login-form">
      <img src="assets/logo.png" alt="Gift Mart" class="login__logo">
      <h1>GMI Employee</h1>
      <p class="muted">كشف الدوام الشهري وكشف الموظف</p>
      <label class="field"><span>المستخدم</span>
        <select id="lg-user">${USERS.map(u => `<option value="${u.id}">${esc(u.name)}</option>`).join('')}</select></label>
      <label class="field"><span>كلمة المرور</span><input id="lg-pass" type="password" inputmode="numeric" autocomplete="current-password" required></label>
      <p class="err" id="lg-err" hidden>كلمة المرور غير صحيحة</p>
      <button class="btn btn--primary" type="submit">دخول</button>
    </form>
  </div>`
  $('#login-form').onsubmit = async e => {
    e.preventDefault()
    const u = await login($('#lg-user').value, $('#lg-pass').value)
    if (!u) {
      $('#lg-err').hidden = false
      return
    }
    render()
  }
}

function periodBar(extra = '') {
  const p = period()
  return `
  <div class="bar">
    <label class="field field--inline"><span>الفترة</span>
      <select id="pb-label">${labelsAround().map(l => `<option value="${l}"${!ui.custom && l === ui.label ? ' selected' : ''}>${MONTHS[Number(l.slice(5)) - 1]} ${l.slice(0, 4)}</option>`).join('')}<option value="custom"${ui.custom ? ' selected' : ''}>فترة مخصصة…</option></select></label>
    ${ui.custom ? `<label class="field field--inline"><span>من</span><input type="date" id="pb-from" value="${p.from}"></label><label class="field field--inline"><span>إلى</span><input type="date" id="pb-to" value="${p.to}"></label>` : `<span class="chip">من ${p.from} إلى ${p.to}</span>`}
    ${extra}
  </div>`
}

function wirePeriodBar(onChange) {
  $('#pb-label').onchange = e => {
    if (e.target.value === 'custom') ui.custom = { ...period() }
    else {
      ui.custom = null
      ui.label = e.target.value
    }
    onChange()
  }
  const f = $('#pb-from')
  const to = $('#pb-to')
  if (f && to) {
    const upd = () => {
      if (f.value && to.value && f.value <= to.value) {
        ui.custom = { from: f.value, to: to.value }
        onChange()
      }
    }
    f.onchange = upd
    to.onchange = upd
  }
}

function legendHtml() {
  return `<div class="legend">${LEGEND.map(k => `<span><i style="background:#${CATEGORIES[k].color}"></i>${CATEGORIES[k].label}</span>`).join('')}<span><i class="lg-abs"></i>غياب للمراجعة</span><span><b>*</b> خصم معدّل</span><span><i class="lg-edit"></i>وقت مُدخل من الموارد</span></div>`
}

function exportButtons(kind) {
  return `<div class="export">
    <button class="btn btn--outline btn--sm" data-x="excel">تصدير Excel</button>
    <button class="btn btn--outline btn--sm" data-x="a4">PDF ‏A4</button>
    <button class="btn btn--outline btn--sm" data-x="a3">PDF ‏A3</button>
  </div>`
}

function wireExports(root, getPayload) {
  root.querySelectorAll('[data-x]').forEach(b => (b.onclick = async () => {
    b.disabled = true
    try {
      const payload = getPayload()
      if (b.dataset.x === 'excel') await exportExcel(payload)
      else await exportPdf(payload, b.dataset.x === 'a3' ? 'A3' : 'A4')
    } catch (err) {
      console.error(err)
      toast('تعذّر التصدير: ' + (err?.message ?? err), 'err')
    } finally {
      b.disabled = false
    }
  }))
}

function exportMeta() {
  const p = period()
  return { period: p, title: periodTitle(p.from, p.to), settings: db().settings, by: currentUser()?.name ?? '', at: nowStamp() }
}

// —— monthly sheet ——

function renderSheet(view) {
  const p = period()
  const s = db().settings
  const emps = employeesSorted().filter(e => !ui.search || `${e.deviceId} ${e.name} ${e.jobNo ?? ''}`.includes(ui.search.trim()))
  if (!db().employees.length) {
    view.innerHTML = `${periodBar()}<div class="card empty"><h2>لا يوجد موظفون بعد</h2><p>ابدأ من «الموظفون» لرفع الأسماء مع أرقامها في جهاز البصمة، ثم «استيراد البصمة».</p></div>`
    wirePeriodBar(render)
    return
  }
  const rows = emps.map(e => viewOf(e))
  const dates = rows[0]?.days.map(d => d.date) ?? []
  view.innerHTML = `
  ${periodBar(`<input class="search" id="sh-search" placeholder="بحث بالاسم أو الرقم" value="${esc(ui.search)}">${exportButtons()}`)}
  <section class="card">
    <div class="card__head"><h2>كشف الدوام — ${esc(periodTitle(p.from, p.to))}</h2><span class="muted">${s.standardPeriodDays} يوم · ${s.standardPeriodDays * s.standardDayHours} ساعة · اضغط على اسم الموظف لفتح كشفه</span></div>
    <div class="table-wrap">
      <table class="ts">
        <thead><tr><th class="sticky">م</th><th class="sticky2">رقم البصمة</th><th class="sticky3 nm">الاسم</th><th>مقطوع</th><th>ناقص</th>
          ${dates.map(d => `<th class="d${s.restDays.includes(new Date(`${d}T12:00:00Z`).getUTCDay()) ? ' rest' : ''}" title="${d}">${Number(d.slice(8))}</th>`).join('')}
          <th>إضافي معتمد</th><th>الخصم المحسوب (س)</th><th>الخصم النهائي (س)</th><th>الفرق</th><th>الإجمالي</th></tr></thead>
        <tbody>${rows.map((r, i) => `<tr>
          <td class="sticky">${i + 1}</td><td class="sticky2">${esc(r.employee.deviceId)}</td>
          <td class="sticky3 nm"><button class="link" data-emp="${esc(r.employee.deviceId)}">${esc(r.employee.name)}</button></td>
          <td>${fmt(r.totals.cut)}</td><td>${fmt(r.totals.missing)}</td>
          ${r.days.map(d => {
            const c = d.status === 'absence' ? CATEGORIES.ABSENCE_REVIEW.color : CATEGORIES[d.category]?.color
            const edited = d.edits.in || d.edits.out
            return `<td class="d${edited ? ' ed' : ''}"${c ? ` style="background:#${c}"` : ''} title="${esc(dayLabel(d.date))} ${esc(dayNote(d))}">${fmt(d.hours)}${d.adjusted ? '*' : ''}</td>`
          }).join('')}
          <td>${fmt(r.totals.overtime)}</td><td>${fmt(r.totals.calcDeduction)}</td><td class="${r.totals.finalDeduction !== r.totals.calcDeduction ? 'chg' : ''}">${fmt(r.totals.finalDeduction)}</td>
          <td class="b">${fmt(r.totals.actH)}</td><td class="b">${fmt(r.totals.totalH)}</td></tr>`).join('')}
        </tbody>
      </table>
    </div>
    ${legendHtml()}
  </section>`
  wirePeriodBar(render)
  const se = $('#sh-search')
  se.oninput = () => {
    ui.search = se.value
    const pos = se.selectionStart
    render()
    const n = $('#sh-search')
    n.focus()
    n.setSelectionRange(pos, pos)
  }
  view.querySelectorAll('[data-emp]').forEach(b => (b.onclick = () => {
    ui.employee = b.dataset.emp
    ui.tab = 'individual'
    render()
  }))
  wireExports(view, () => ({ ...exportMeta(), kind: 'monthly', rows: employeesSorted().map(e => viewOf(e)) }))
}

// —— individual sheet ——

function renderIndividual(view, user) {
  const emps = employeesSorted()
  if (!emps.length) {
    view.innerHTML = `<div class="card empty"><h2>لا يوجد موظفون بعد</h2><p>ارفع قائمة الموظفين من «الموظفون».</p></div>`
    return
  }
  if (!ui.employee || !emps.some(e => e.deviceId === ui.employee)) ui.employee = emps[0].deviceId
  const emp = emps.find(e => e.deviceId === ui.employee)
  const v = viewOf(emp)
  const p = period()
  const s = db().settings
  const tt = v.totals
  view.innerHTML = `
  ${periodBar(`<label class="field field--inline"><span>الموظف</span><select id="iv-emp">${emps.map(e => `<option value="${esc(e.deviceId)}"${e.deviceId === emp.deviceId ? ' selected' : ''}>${esc(e.deviceId)} — ${esc(e.name)}</option>`).join('')}</select></label>${exportButtons()}`)}
  <section class="card">
    <div class="card__head"><h2>كشف الدوام الفردي</h2><span class="muted">من ${p.from} إلى ${p.to} · اضغط على خانة الدخول أو الخروج للإضافة أو التصحيح، وعلى التاريخ لتصنيف اليوم أو تعديل الخصم</span></div>
    <div class="facts">
      <div><span>الاسم</span><b>${esc(emp.name)}</b></div>
      <div><span>رقم البصمة</span><b>${esc(emp.deviceId)}</b></div>
      <div><span>الرقم الوظيفي</span><b>${esc(emp.jobNo || '—')}</b></div>
      <div><span>رقم الهوية</span><b>${esc(emp.idNumber || '—')}</b></div>
      <div><span>تاريخ المباشرة</span><b>${esc(emp.hireDate || '—')}</b></div>
      <div><span>بداية الدوام</span><b>${esc(emp.shiftStart || s.shiftStart)}</b></div>
    </div>
    <div class="table-wrap">
      <table class="days">
        <thead><tr><th>التاريخ</th><th>البصمات</th><th>دخول</th><th>خروج</th><th>تأخير (د)</th><th>بعد نهاية الدوام (د)</th><th>الخصم (س/يوم)</th><th class="nt">ملاحظة</th><th>الساعات</th></tr></thead>
        <tbody>${v.days.map(d => {
          const c = d.status === 'absence' ? CATEGORIES.ABSENCE_REVIEW.color : CATEGORIES[d.category]?.color
          const ded = d.items.map(i => `${SOURCES[i.source]} ${fmt(i.final)}${i.unit === 'day' ? ' يوم' : ' س'}${i.final !== i.proposed ? '*' : ''}`).join('<br>')
          const cell = f => {
            const val = f === 'in' ? d.inMin : d.outMin
            const e = d.edits[f]
            const cls = e ? (e.kind === 'fill' ? 'tm tm--fill' : 'tm tm--fix') : val === null ? 'tm tm--empty' : 'tm'
            return `<button class="${cls}" data-edit="${f}" data-date="${d.date}" title="${e ? `${e.kind === 'fill' ? 'إعادة تعبئة' : 'تصحيح'}: ${esc(e.note)} — ${esc(e.by)}` : val === null ? 'إضافة وقت' : 'تصحيح الوقت'}">${val === null ? '+' : t(val)}</button>`
          }
          return `<tr class="${d.status === 'rest' ? 'is-rest' : ''}">
            <td><button class="link" data-day="${d.date}">${esc(dayLabel(d.date))}</button></td>
            <td class="fp">${d.punches.join(' ')}</td>
            <td>${d.status === 'rest' && !d.edits.in ? '' : cell('in')}</td><td>${d.status === 'rest' && !d.edits.out ? '' : cell('out')}</td>
            <td>${d.lateMinutes || ''}</td><td>${d.extraMinutes || ''}</td><td class="dd">${ded}</td>
            <td class="nt">${esc(dayNote(d))}</td>
            <td class="h"${c ? ` style="background:#${c}"` : ''}>${fmt(d.hours)}</td></tr>`
        }).join('')}</tbody>
      </table>
    </div>
    <div class="bottom">
      <div class="summary">
        <h3>الملخص</h3>
        <dl>
          <dt>أيام الفترة</dt><dd>${tt.standardDays}</dd>
          <dt>الساعات الرسمية</dt><dd>${tt.standardHours}</dd>
          <dt>الخصم — المحسوب / النهائي (س)</dt><dd>${fmt(tt.calcDeduction)} / ${fmt(tt.finalDeduction)}</dd>
          <dt>أيام التأخير</dt><dd>${tt.lateDays}</dd>
          <dt>أيام البصمة الناقصة</dt><dd>${tt.missingDays}</dd>
          <dt>أيام الغياب للمراجعة</dt><dd>${tt.absenceDays}</dd>
          <dt>الفرق (+/−)</dt><dd class="b">${fmt(tt.actH)}</dd>
          <dt>صافي الساعات</dt><dd class="b">${fmt(tt.totalH)}</dd>
        </dl>
      </div>
      <form class="carry" id="carry-form">
        <h3>ساعات مرحّلة وإضافي (لصالح الموظف)</h3>
        <label class="field"><span>مقطوع — إعادة خصم سابق بالخطأ (س)</span><input type="number" step="0.25" min="0" id="cr-cut" value="${esc(v.carry.cut ?? '')}"></label>
        <label class="field"><span>ناقص — ساعات عمل لم تُحتسب (س)</span><input type="number" step="0.25" min="0" id="cr-missing" value="${esc(v.carry.missing ?? '')}"></label>
        <label class="field"><span>إضافي معتمد هذه الفترة (س)</span><input type="number" step="0.25" min="0" id="cr-ot" value="${esc(v.carry.overtime ?? '')}"></label>
        <button class="btn btn--primary btn--sm" type="submit">حفظ</button>
      </form>
    </div>
    ${legendHtml()}
  </section>`
  wirePeriodBar(render)
  $('#iv-emp').onchange = e => {
    ui.employee = e.target.value
    render()
  }
  $('#carry-form').onsubmit = e => {
    e.preventDefault()
    const k = carryKey(p)
    db().carry[k] = db().carry[k] ?? {}
    db().carry[k][emp.deviceId] = { cut: Number($('#cr-cut').value) || 0, missing: Number($('#cr-missing').value) || 0, overtime: Number($('#cr-ot').value) || 0 }
    persist()
    toast('حُفظت الساعات المرحّلة')
    render()
  }
  view.querySelectorAll('[data-edit]').forEach(b => (b.onclick = () => openTimeDialog(emp, v.days.find(d => d.date === b.dataset.date), b.dataset.edit, user)))
  view.querySelectorAll('[data-day]').forEach(b => (b.onclick = () => openDayDialog(emp, v.days.find(d => d.date === b.dataset.day), user)))
  wireExports(view, () => ({ ...exportMeta(), kind: 'individual', rows: [viewOf(emp)] }))
}

function dialog(html, onMount) {
  const wrap = document.createElement('div')
  wrap.className = 'modal'
  wrap.innerHTML = `<div class="modal__box card" role="dialog" aria-modal="true">${html}</div>`
  document.body.appendChild(wrap)
  const close = () => wrap.remove()
  wrap.addEventListener('click', e => {
    if (e.target === wrap) close()
  })
  wrap.querySelectorAll('[data-close]').forEach(b => (b.onclick = close))
  onMount(wrap, close)
  wrap.querySelector('input, select, textarea')?.focus()
}

/** HR adds a missing time (fill) or changes a device time (correction); the device punches stay as they were. */
function openTimeDialog(emp, d, field, user) {
  const device = field === 'in' ? d.deviceIn : d.deviceOut
  const current = d.edits[field]
  const kind = device === null ? 'fill' : 'correct'
  const label = field === 'in' ? 'الدخول' : 'الخروج'
  dialog(`
    <h2>${kind === 'fill' ? 'إعادة تعبئة' : 'تصحيح'} وقت ${label}</h2>
    <p class="muted">${esc(emp.name)} — ${esc(dayLabel(d.date))} ${d.date}</p>
    <p class="hint ${kind === 'fill' ? 'hint--fill' : 'hint--fix'}">${kind === 'fill' ? 'لا توجد بصمة لهذا الوقت في الجهاز — ستُضاف يدوياً وتظهر في صفحة التصحيحات كـ«إعادة تعبئة».' : `وقت الجهاز: <b>${t(device)}</b> — سيبقى مسجلاً، والوقت الجديد يظهر في صفحة التصحيحات كـ«تصحيح».`}</p>
    ${current ? `<p class="muted">آخر إدخال: ${t(toMin(current.value))} — ${esc(current.note)} (${esc(current.by)})</p>` : ''}
    <form id="tm-form">
      <label class="field"><span>الوقت</span><input type="time" id="tm-val" required value="${current ? current.value : device !== null ? t(device) : ''}"></label>
      <label class="field"><span>ملاحظة الموارد البشرية (إلزامية)</span><textarea id="tm-note" rows="3" required minlength="3" placeholder="لماذا أُضيف أو صُحّح الوقت؟"></textarea></label>
      <div class="actions">
        ${current ? '<button type="button" class="btn btn--danger btn--sm" id="tm-del">حذف آخر إدخال</button>' : ''}
        <span class="grow"></span>
        <button type="button" class="btn btn--ghost" data-close>إلغاء</button>
        <button type="submit" class="btn btn--primary">حفظ</button>
      </div>
    </form>`, (wrap, close) => {
    wrap.querySelector('#tm-form').onsubmit = e => {
      e.preventDefault()
      const value = wrap.querySelector('#tm-val').value
      const note = wrap.querySelector('#tm-note').value.trim()
      if (!value || note.length < 3) return
      if (kind === 'correct' && toMin(value) === device && !current) {
        toast('الوقت مطابق لوقت الجهاز — لا يوجد تصحيح', 'err')
        return
      }
      db().edits.push({ id: crypto.randomUUID(), deviceId: emp.deviceId, date: d.date, field, kind, original: device === null ? null : t(device), value, note, by: user.name, at: nowStamp() })
      persist()
      close()
      toast(kind === 'fill' ? 'أُضيف الوقت' : 'صُحّح الوقت')
      render()
    }
    const del = wrap.querySelector('#tm-del')
    if (del) del.onclick = () => {
      if (!confirm('حذف آخر إدخال لهذا الوقت؟')) return
      db().edits = db().edits.filter(x => x.id !== current.id)
      persist()
      close()
      render()
    }
  })
}

/** Day code (leave, trip, permission…) and deduction waivers — always with a reason and who decided. */
function openDayDialog(emp, d, user) {
  dialog(`
    <h2>${esc(dayLabel(d.date))} ${d.date}</h2>
    <p class="muted">${esc(emp.name)} · ${esc(dayNote(d)) || 'يوم عادي'}</p>
    <form id="day-form">
      <label class="field"><span>تصنيف اليوم</span><select id="dy-code"><option value="">بدون — حسب البصمة</option>${DAY_CODES.map(c => `<option value="${c}"${d.mark?.code === c ? ' selected' : ''}>${CATEGORIES[c].label}</option>`).join('')}</select></label>
      <label class="field"><span>ملاحظة التصنيف</span><input id="dy-note" value="${esc(d.mark?.note ?? '')}"></label>
      ${d.items.length ? `<fieldset><legend>تعديل الخصم (الإعفاء أو التخفيض)</legend>${d.items.map((i, n) => `
        <div class="wv"><b>${SOURCES[i.source]}</b> — المحسوب ${fmt(i.proposed)} ${i.unit === 'day' ? 'يوم' : 'س'}
          <label class="field field--inline"><span>النهائي</span><input type="number" step="0.25" min="0" max="${i.proposed}" data-wv="${n}" value="${fmt(i.final)}"></label>
          <label class="field"><span>السبب</span><input data-wr="${n}" value="${esc(i.waiver?.reason ?? '')}"></label></div>`).join('')}</fieldset>` : ''}
      <div class="actions"><span class="grow"></span><button type="button" class="btn btn--ghost" data-close>إلغاء</button><button type="submit" class="btn btn--primary">حفظ</button></div>
    </form>`, (wrap, close) => {
    wrap.querySelector('#day-form').onsubmit = e => {
      e.preventDefault()
      const code = wrap.querySelector('#dy-code').value
      const note = wrap.querySelector('#dy-note').value.trim()
      db().marks = db().marks.filter(m => !(m.deviceId === emp.deviceId && m.date === d.date))
      if (code) db().marks.push({ deviceId: emp.deviceId, date: d.date, code, note, by: user.name, at: nowStamp() })
      for (const [n, i] of d.items.entries()) {
        const val = Number(wrap.querySelector(`[data-wv="${n}"]`).value)
        const reason = wrap.querySelector(`[data-wr="${n}"]`).value.trim()
        db().waivers = db().waivers.filter(w => !(w.deviceId === emp.deviceId && w.date === d.date && w.source === i.source))
        if (Number.isFinite(val) && val !== i.proposed) {
          if (reason.length < 3) {
            toast('اكتب سبب تعديل الخصم', 'err')
            return
          }
          db().waivers.push({ deviceId: emp.deviceId, date: d.date, source: i.source, proposed: i.proposed, final: Math.min(Math.max(0, val), i.proposed), reason, by: user.name, at: nowStamp() })
        }
      }
      persist()
      close()
      render()
    }
  })
}

// —— corrections log ——

function renderCorrections(view) {
  const p = period()
  const names = new Map(db().employees.map(e => [e.deviceId, e]))
  const list = db().edits.filter(e => e.date >= p.from && e.date <= p.to).sort((a, b) => (a.deviceId === b.deviceId ? (a.date + a.at < b.date + b.at ? -1 : 1) : Number(a.deviceId) - Number(b.deviceId)))
  view.innerHTML = `
  ${periodBar(exportButtons())}
  <section class="card">
    <div class="card__head"><h2>سجل التصحيحات وإعادة التعبئة</h2><span class="muted">كل وقت أضافته أو صحّحته الموارد البشرية — يطلع مع التصدير في صفحة مستقلة</span></div>
    ${list.length ? `<div class="table-wrap"><table class="days">
      <thead><tr><th>رقم البصمة</th><th>الاسم</th><th>اليوم</th><th>الخانة</th><th>النوع</th><th>وقت الجهاز</th><th>الوقت المعتمد</th><th class="nt">ملاحظة الموارد البشرية</th><th>بواسطة</th><th>وقت التسجيل</th></tr></thead>
      <tbody>${list.map(e => `<tr><td>${esc(e.deviceId)}</td><td>${esc(names.get(e.deviceId)?.name ?? '')}</td><td>${esc(dayLabel(e.date))} ${e.date}</td><td>${e.field === 'in' ? 'دخول' : 'خروج'}</td>
        <td><span class="badge ${e.kind === 'fill' ? 'badge--fill' : 'badge--fix'}">${e.kind === 'fill' ? 'إعادة تعبئة' : 'تصحيح'}</span></td><td>${esc(e.original ?? '—')}</td><td><b>${esc(e.value)}</b></td><td class="nt">${esc(e.note)}</td><td>${esc(e.by)}</td><td>${esc(e.at)}</td></tr>`).join('')}</tbody></table></div>`
      : '<p class="muted pad">لا توجد تصحيحات في هذه الفترة.</p>'}
  </section>`
  wirePeriodBar(render)
  wireExports(view, () => ({ ...exportMeta(), kind: 'corrections', rows: employeesSorted().map(e => viewOf(e)) }))
}

// —— employees ——

function renderEmployees(view, user) {
  const emps = employeesSorted()
  const known = new Set(emps.map(e => e.deviceId))
  const unknown = Object.entries(db().punches).filter(([id]) => !known.has(id)).map(([id, l]) => ({ id, n: l.length, last: [...l].sort().pop() }))
  unknown.sort((a, b) => Number(a.id) - Number(b.id))
  view.innerHTML = `
  <section class="card">
    <div class="card__head"><h2>الموظفون (${emps.length})</h2>
      <div class="row">
        <button class="btn btn--outline btn--sm" id="em-tpl">تنزيل نموذج Excel</button>
        <label class="btn btn--primary btn--sm file">رفع ملف الموظفين<input type="file" id="em-file" accept=".xlsx,.csv" hidden></label>
      </div>
    </div>
    <p class="muted pad">الأعمدة: <b>رقم البصمة</b> (رقم الموظف في الجهاز) · <b>الاسم</b> (الاسم الكامل) · الرقم الوظيفي · رقم الهوية · تاريخ المباشرة · بداية الدوام (إذا كانت تختلف عن العامة). الرفع يحدّث الموجود ويضيف الجديد.</p>
    <form class="add-row" id="em-add">
      <input id="ea-id" placeholder="رقم البصمة" inputmode="numeric" required>
      <input id="ea-name" placeholder="الاسم الكامل" required>
      <input id="ea-job" placeholder="الرقم الوظيفي">
      <input id="ea-idn" placeholder="رقم الهوية" inputmode="numeric">
      <input id="ea-hire" type="date" title="تاريخ المباشرة">
      <input id="ea-shift" type="time" title="بداية الدوام (اختياري)">
      <button class="btn btn--primary btn--sm" type="submit">إضافة / تحديث</button>
    </form>
    ${emps.length ? `<div class="table-wrap"><table class="days">
      <thead><tr><th>رقم البصمة</th><th>الاسم</th><th>الرقم الوظيفي</th><th>رقم الهوية</th><th>تاريخ المباشرة</th><th>بداية الدوام</th><th>عدد البصمات</th><th></th></tr></thead>
      <tbody>${emps.map(e => `<tr><td>${esc(e.deviceId)}</td><td class="nt">${esc(e.name)}</td><td>${esc(e.jobNo ?? '')}</td><td>${esc(e.idNumber ?? '')}</td><td>${esc(e.hireDate ?? '')}</td><td>${esc(e.shiftStart || '')}</td><td>${(db().punches[e.deviceId] ?? []).length}</td>
        <td><button class="btn btn--ghost btn--sm" data-edit-emp="${esc(e.deviceId)}">تعديل</button>${user.role === 'admin' ? ` <button class="btn btn--ghost btn--sm danger" data-del-emp="${esc(e.deviceId)}">حذف</button>` : ''}</td></tr>`).join('')}</tbody></table></div>` : ''}
  </section>
  ${unknown.length ? `<section class="card">
    <div class="card__head"><h2>أرقام في جهاز البصمة بدون اسم (${unknown.length})</h2><span class="muted">اضغط «ربط» واكتب الاسم الكامل</span></div>
    <div class="table-wrap"><table class="days"><thead><tr><th>رقم البصمة</th><th>عدد البصمات</th><th>آخر بصمة</th><th></th></tr></thead>
    <tbody>${unknown.map(u => `<tr><td>${esc(u.id)}</td><td>${u.n}</td><td>${esc(u.last)}</td><td><button class="btn btn--outline btn--sm" data-link="${esc(u.id)}">ربط باسم</button></td></tr>`).join('')}</tbody></table></div>
  </section>` : ''}`
  $('#em-tpl').onclick = () => employeesTemplate().catch(err => toast(String(err), 'err'))
  $('#em-file').onchange = async e => {
    const f = e.target.files[0]
    if (!f) return
    try {
      const list = await readEmployeesFile(f)
      let added = 0
      let updated = 0
      for (const x of list) {
        const i = db().employees.findIndex(y => y.deviceId === x.deviceId)
        if (i >= 0) {
          db().employees[i] = { ...db().employees[i], ...Object.fromEntries(Object.entries(x).filter(([, v]) => v)) }
          updated++
        } else {
          db().employees.push(x)
          added++
        }
      }
      persist()
      toast(`أُضيف ${added} وحُدّث ${updated}`)
      render()
    } catch (err) {
      toast('تعذّرت قراءة الملف: ' + (err?.message ?? err), 'err')
    }
  }
  $('#em-add').onsubmit = e => {
    e.preventDefault()
    const deviceId = String(Number($('#ea-id').value.trim()))
    if (!/^\d+$/.test(deviceId)) return toast('رقم البصمة أرقام فقط', 'err')
    const rec = { deviceId, name: $('#ea-name').value.trim(), jobNo: $('#ea-job').value.trim(), idNumber: $('#ea-idn').value.trim(), hireDate: $('#ea-hire').value, shiftStart: $('#ea-shift').value }
    const i = db().employees.findIndex(y => y.deviceId === deviceId)
    if (i >= 0) db().employees[i] = rec
    else db().employees.push(rec)
    persist()
    render()
  }
  view.querySelectorAll('[data-edit-emp]').forEach(b => (b.onclick = () => {
    const e = db().employees.find(x => x.deviceId === b.dataset.editEmp)
    $('#ea-id').value = e.deviceId
    $('#ea-name').value = e.name
    $('#ea-job').value = e.jobNo ?? ''
    $('#ea-idn').value = e.idNumber ?? ''
    $('#ea-hire').value = e.hireDate ?? ''
    $('#ea-shift').value = e.shiftStart ?? ''
    $('#ea-name').focus()
  }))
  view.querySelectorAll('[data-del-emp]').forEach(b => (b.onclick = () => {
    if (!confirm('حذف الموظف من القائمة؟ بصماته تبقى محفوظة وتظهر في «أرقام بدون اسم».')) return
    db().employees = db().employees.filter(x => x.deviceId !== b.dataset.delEmp)
    persist()
    render()
  }))
  view.querySelectorAll('[data-link]').forEach(b => (b.onclick = () => {
    $('#ea-id').value = b.dataset.link
    $('#ea-name').focus()
    $('#em-add').scrollIntoView({ behavior: 'smooth', block: 'center' })
  }))
}

// —— import ——

function renderImport(view, user) {
  const pv = ui.importPreview
  view.innerHTML = `
  <section class="card">
    <div class="card__head"><h2>استيراد ملف البصمة (من الفلاشة)</h2></div>
    <ol class="steps">
      <li>من جهاز البصمة: <b>USB ← تنزيل سجل الحضور</b>، ينزل ملف اسمه مثل <code>…_attlog.dat</code>.</li>
      <li>اختر الملف، وحدد الفترة اللي تبي تنزلها (من — إلى). اتركها فاضية لكل الملف.</li>
      <li>راجع الملخص ثم اضغط «استيراد». البصمات المكررة لا تتكرر، والبصمات الأصلية لا تُعدّل أبداً.</li>
    </ol>
    <div class="bar">
      <label class="btn btn--primary btn--sm file">اختيار الملف<input type="file" id="im-file" accept=".dat,.txt,.csv" hidden></label>
      <label class="field field--inline"><span>من</span><input type="date" id="im-from" value="${pv?.from ?? ''}"></label>
      <label class="field field--inline"><span>إلى</span><input type="date" id="im-to" value="${pv?.to ?? ''}"></label>
    </div>
    ${pv ? `<div class="preview">
      <div class="facts">
        <div><span>الملف</span><b>${esc(pv.file)}</b></div>
        <div><span>السطور المقروءة</span><b>${pv.rows.length}</b></div>
        <div><span>سطور غير مفهومة</span><b>${pv.bad}</b></div>
        <div><span>أول بصمة في الملف</span><b>${esc(pv.min)}</b></div>
        <div><span>آخر بصمة في الملف</span><b>${esc(pv.max)}</b></div>
        <div><span>داخل الفترة المختارة</span><b id="im-in">${pv.inRange}</b></div>
        <div><span>أرقام موظفين</span><b>${pv.ids}</b></div>
        <div><span>أرقام بدون اسم</span><b>${pv.unknown}</b></div>
      </div>
      <div class="actions"><span class="grow"></span><button class="btn btn--ghost" id="im-cancel">إلغاء</button><button class="btn btn--primary" id="im-go">استيراد ${pv.inRange} بصمة</button></div>
    </div>` : ''}
  </section>
  <section class="card">
    <div class="card__head"><h2>سجل الاستيراد</h2></div>
    ${db().imports.length ? `<div class="table-wrap"><table class="days"><thead><tr><th>الوقت</th><th>بواسطة</th><th>الملف</th><th>الفترة</th><th>المقروء</th><th>الجديد</th></tr></thead>
      <tbody>${[...db().imports].reverse().map(i => `<tr><td>${esc(i.at)}</td><td>${esc(i.by)}</td><td>${esc(i.file)}</td><td>${esc(i.from || 'كل الملف')}${i.to ? ' ← ' + esc(i.to) : ''}</td><td>${i.read}</td><td>${i.added}</td></tr>`).join('')}</tbody></table></div>` : '<p class="muted pad">لم يُستورد شيء بعد.</p>'}
  </section>`
  const recount = () => {
    if (!ui.importPreview) return
    const from = $('#im-from').value
    const to = $('#im-to').value
    ui.importPreview.from = from
    ui.importPreview.to = to
    ui.importPreview.inRange = ui.importPreview.rows.filter(r => (!from || r.at.slice(0, 10) >= from) && (!to || r.at.slice(0, 10) <= to)).length
    render()
  }
  $('#im-from').onchange = recount
  $('#im-to').onchange = recount
  $('#im-file').onchange = async e => {
    const f = e.target.files[0]
    if (!f) return
    const { rows, bad } = parseAttlog(await f.text())
    if (!rows.length) return toast('الملف لا يحتوي بصمات بالشكل المتوقع', 'err')
    const ats = rows.map(r => r.at).sort()
    const ids = new Set(rows.map(r => r.deviceId))
    const known = new Set(db().employees.map(x => x.deviceId))
    ui.importPreview = { file: f.name, rows, bad, min: ats[0], max: ats[ats.length - 1], ids: ids.size, unknown: [...ids].filter(i => !known.has(i)).length, from: $('#im-from').value, to: $('#im-to').value, inRange: rows.length }
    recount()
  }
  if (pv) {
    $('#im-cancel').onclick = () => {
      ui.importPreview = null
      render()
    }
    $('#im-go').onclick = () => {
      const { from, to } = pv
      let added = 0
      let read = 0
      const sets = new Map()
      for (const r of pv.rows) {
        const d = r.at.slice(0, 10)
        if ((from && d < from) || (to && d > to)) continue
        read++
        if (!sets.has(r.deviceId)) sets.set(r.deviceId, new Set(db().punches[r.deviceId] ?? []))
        const s = sets.get(r.deviceId)
        if (!s.has(r.at)) {
          s.add(r.at)
          added++
        }
      }
      for (const [id, s] of sets) db().punches[id] = [...s].sort()
      db().imports.push({ at: nowStamp(), by: user.name, file: pv.file, from, to, read, added })
      ui.importPreview = null
      persist()
      toast(`استُوردت ${added} بصمة جديدة`)
      render()
    }
  }
}

// —— settings ——

function renderSettings(view, user) {
  const s = db().settings
  const admin = user.role === 'admin'
  const dis = admin ? '' : ' disabled'
  view.innerHTML = `
  <section class="card">
    <div class="card__head"><h2>قواعد الكشف</h2>${admin ? '' : '<span class="muted">التعديل لحساب admin فقط</span>'}</div>
    <form id="st-form" class="grid">
      <fieldset><legend>الفترة والدوام</legend>
        <label class="field"><span>يوم بداية الفترة</span><input type="number" min="1" max="28" id="st-start" value="${s.periodStartDay}"${dis}></label>
        <label class="field"><span>ساعات اليوم</span><input type="number" min="1" max="24" step="0.5" id="st-std" value="${s.standardDayHours}"${dis}></label>
        <label class="field"><span>أيام الفترة الرسمية</span><input type="number" min="1" max="31" id="st-days" value="${s.standardPeriodDays}"${dis}></label>
        <label class="field"><span>بداية الدوام</span><input type="time" id="st-shift" value="${s.shiftStart}"${dis}></label>
        <label class="field"><span>نهاية الدوام (اختياري)</span><input type="time" id="st-end" value="${s.shiftEnd}"${dis}></label>
        <label class="field"><span>دمج البصمات المتكررة خلال (دقيقة)</span><input type="number" min="0" max="30" id="st-dup" value="${s.duplicateMinutes}"${dis}></label>
        <div class="field"><span>أيام الراحة</span><div class="checks">${WEEKDAYS.map((w, i) => `<label><input type="checkbox" data-rest="${i}"${s.restDays.includes(i) ? ' checked' : ''}${dis}> ${w}</label>`).join('')}</div></div>
      </fieldset>
      <fieldset class="fs-late"><legend>التأخير</legend>
        <label class="field"><span>دقائق السماح</span><input type="number" min="0" max="120" id="st-grace" value="${s.lateness.graceMinutes}"${dis}></label>
        <div class="tiers-wrap"><table class="days tiers"><thead><tr><th>من (د)</th><th>إلى (د)</th><th>الخصم</th><th>الوحدة</th><th>يبدأ من المرة</th></tr></thead>
        <tbody>${s.lateness.tiers.map((tr, i) => `<tr>
          <td><input type="number" data-tf="${i}" value="${tr.fromMinute}"${dis}></td>
          <td><input type="number" data-tt="${i}" value="${tr.toMinute ?? ''}" placeholder="مفتوح"${dis}></td>
          <td><input type="number" step="0.25" data-tv="${i}" value="${tr.value}"${dis}></td>
          <td><select data-tu="${i}"${dis}><option value="hours"${tr.unit === 'hours' ? ' selected' : ''}>ساعة</option><option value="day"${tr.unit === 'day' ? ' selected' : ''}>يوم</option></select></td>
          <td><input type="number" min="1" max="31" data-to="${i}" value="${tr.startingOccurrence}"${dis}></td></tr>`).join('')}</tbody></table></div>
        <p class="muted">كل فئة لها عدّاد مستقل يبدأ من جديد أول كل شهر، ومن المرة المحددة تُخصم كل مرة.</p>
      </fieldset>
      <fieldset><legend>البصمة الناقصة والغياب</legend>
        <label class="check"><input type="checkbox" id="st-mp-on"${s.missingPunch.active ? ' checked' : ''}${dis}> خصم البصمة الناقصة</label>
        <label class="field"><span>الخصم (ساعة)</span><input type="number" step="0.25" id="st-mp-v" value="${s.missingPunch.value}"${dis}></label>
        <label class="field"><span>يبدأ من المرة</span><input type="number" min="1" id="st-mp-o" value="${s.missingPunch.startingOccurrence}"${dis}></label>
        <label class="check"><input type="checkbox" id="st-ab-on"${s.absence.active ? ' checked' : ''}${dis}> خصم الغياب بدون بصمة (يوم كامل)</label>
        <p class="muted">إذا كان الخصم مطفأ يظهر اليوم باللون الأحمر «غياب للمراجعة» بدون خصم، حتى تصنّفه الموارد (إجازة، مرضية، غياب…).</p>
      </fieldset>
      ${admin ? '<div class="actions"><button type="button" class="btn btn--ghost" id="st-default">استرجاع الافتراضي</button><span class="grow"></span><button class="btn btn--primary" type="submit">حفظ القواعد</button></div>' : ''}
    </form>
  </section>
  <section class="card">
    <div class="card__head"><h2>النسخ الاحتياطي</h2></div>
    <p class="muted pad">البيانات محفوظة في هذا المتصفح فقط. لنقلها لجهاز ثاني أو لحفظها: صدّر نسخة واسترجعها هناك.</p>
    <div class="bar">
      <button class="btn btn--outline btn--sm" id="bk-out">تصدير نسخة احتياطية</button>
      <label class="btn btn--outline btn--sm file">استرجاع نسخة<input type="file" id="bk-in" accept=".json" hidden></label>
      ${admin ? '<button class="btn btn--danger btn--sm" id="bk-reset">مسح كل البيانات</button>' : ''}
    </div>
  </section>`
  if (admin) {
    $('#st-form').onsubmit = e => {
      e.preventDefault()
      const n = id => Number($(id).value)
      const tiers = s.lateness.tiers.map((_, i) => ({
        fromMinute: Number($(`[data-tf="${i}"]`).value),
        toMinute: $(`[data-tt="${i}"]`).value === '' ? null : Number($(`[data-tt="${i}"]`).value),
        value: Number($(`[data-tv="${i}"]`).value),
        unit: $(`[data-tu="${i}"]`).value,
        startingOccurrence: Number($(`[data-to="${i}"]`).value),
      }))
      const grace = n('#st-grace')
      for (let i = 0; i < tiers.length; i++) {
        const prev = i === 0 ? grace : tiers[i - 1].toMinute
        if (prev === null || tiers[i].fromMinute !== prev + 1 || (i < tiers.length - 1 && tiers[i].toMinute === null) || (i === tiers.length - 1 && tiers[i].toMinute !== null)) {
          return toast(`الفئات لازم تكون متتابعة: الأولى تبدأ بعد السماح، وكل فئة بعد اللي قبلها، والأخيرة مفتوحة (الفئة ${i + 1})`, 'err')
        }
      }
      db().settings = {
        ...s,
        periodStartDay: n('#st-start'),
        standardDayHours: n('#st-std'),
        standardPeriodDays: n('#st-days'),
        shiftStart: $('#st-shift').value || '08:00',
        shiftEnd: $('#st-end').value,
        duplicateMinutes: n('#st-dup'),
        restDays: [...view.querySelectorAll('[data-rest]')].filter(x => x.checked).map(x => Number(x.dataset.rest)),
        lateness: { graceMinutes: grace, tiers },
        missingPunch: { ...s.missingPunch, active: $('#st-mp-on').checked, value: n('#st-mp-v'), startingOccurrence: n('#st-mp-o') },
        absence: { ...s.absence, active: $('#st-ab-on').checked },
      }
      persist()
      toast('حُفظت القواعد')
      render()
    }
    $('#st-default').onclick = () => {
      if (!confirm('استرجاع القواعد الافتراضية؟')) return
      db().settings = structuredClone(DEFAULT_SETTINGS)
      persist()
      render()
    }
    $('#bk-reset').onclick = () => {
      if (!confirm('مسح كل البيانات من هذا المتصفح؟ صدّر نسخة احتياطية أولاً.')) return
      if (prompt('اكتب: مسح') !== 'مسح') return
      resetAll()
      render()
    }
  }
  $('#bk-out').onclick = () => {
    const blob = new Blob([JSON.stringify(db())], { type: 'application/json' })
    const a = document.createElement('a')
    a.href = URL.createObjectURL(blob)
    a.download = `GMI-Employee-نسخة-${nowStamp().slice(0, 10)}.json`
    a.click()
    URL.revokeObjectURL(a.href)
  }
  $('#bk-in').onchange = async e => {
    const f = e.target.files[0]
    if (!f) return
    try {
      const data = JSON.parse(await f.text())
      if (!data || !Array.isArray(data.employees) || typeof data.punches !== 'object') throw new Error('ليس ملف نسخة صحيح')
      if (!confirm('استرجاع النسخة يستبدل كل البيانات الحالية في هذا المتصفح. متابعة؟')) return
      replaceAll(data)
      toast('استُرجعت النسخة')
      render()
    } catch (err) {
      toast('تعذّر الاسترجاع: ' + (err?.message ?? err), 'err')
    }
  }
}

try {
  render()
} catch (err) {
  console.error(err)
  app.innerHTML = `<div class="card empty"><h2>حدث خطأ</h2><p>${esc(err?.message ?? err)}</p></div>`
}
