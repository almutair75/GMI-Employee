/**
 * Storage — everything stays in this browser (localStorage). Nothing is sent to any server.
 * The backup file moves the data to another device.
 */

import { DEFAULT_SETTINGS } from './rules.js'

const KEY = 'gmi-employee:v1'

/** Passwords are kept as hashes only (SHA-256 of «GMI-Employee|password»). */
export const USERS = [
  { id: 'tahani', name: 'تهاني العنزي', role: 'hr', hash: '5ec34ef0863c3b7f1755d59a9e3aedf1b41b6ae64d70e6f6514971c3317f968e' },
  { id: 'fatih', name: 'محمد الفاتح', role: 'hr', hash: '30312156baf07064468aba7cea5155187228b6011c4b5211e73a34e3fc9dc3e8' },
  { id: 'noor', name: 'نور بوزيد', role: 'hr', hash: 'c1bf3355f3bff45971a0d0588ff2bfa886005501a3051d5cadb635f0eaa841bc' },
  { id: 'admin', name: 'admin', role: 'admin', hash: '001da4c7b1380bf6bb7a0bee6cae822b4dfffac595c7fc522eed8035365ba23f' },
]

function empty() {
  return {
    settings: structuredClone(DEFAULT_SETTINGS),
    employees: [], // { deviceId, name, jobNo, idNumber, hireDate, shiftStart }
    punches: {}, // deviceId → ['YYYY-MM-DD HH:MM:SS', …] (as read from the device, never edited)
    imports: [], // { at, by, file, from, to, read, added }
    edits: [], // { id, deviceId, date, field: 'in'|'out', kind: 'fill'|'correct', original, value, note, by, at }
    marks: [], // { deviceId, date, code, note, by, at }
    waivers: [], // { deviceId, date, source, proposed, final, reason, by, at }
    carry: {}, // 'from|to' → deviceId → { cut, missing, overtime }
  }
}

let state = load()

function load() {
  try {
    const raw = localStorage.getItem(KEY)
    if (!raw) return empty()
    const s = JSON.parse(raw)
    const base = empty()
    return { ...base, ...s, settings: { ...base.settings, ...(s.settings ?? {}) } }
  } catch {
    return empty()
  }
}

export function save() {
  try {
    localStorage.setItem(KEY, JSON.stringify(state))
    return true
  } catch {
    return false
  }
}

export const db = () => state

export function replaceAll(next) {
  const base = empty()
  state = { ...base, ...next, settings: { ...base.settings, ...(next.settings ?? {}) } }
  save()
}

export function resetAll() {
  state = empty()
  save()
}

export async function sha256(text) {
  const buf = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(text))
  return [...new Uint8Array(buf)].map(b => b.toString(16).padStart(2, '0')).join('')
}

export async function login(userId, password) {
  const u = USERS.find(x => x.id === userId)
  if (!u) return null
  if ((await sha256(`GMI-Employee|${password}`)) !== u.hash) return null
  try {
    sessionStorage.setItem('gmi-employee:user', u.id)
  } catch {}
  return u
}

export function currentUser() {
  try {
    return USERS.find(u => u.id === sessionStorage.getItem('gmi-employee:user')) ?? null
  } catch {
    return null
  }
}

export function logout() {
  try {
    sessionStorage.removeItem('gmi-employee:user')
  } catch {}
}

export const nowStamp = () => {
  const d = new Date()
  const p = n => String(n).padStart(2, '0')
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`
}
