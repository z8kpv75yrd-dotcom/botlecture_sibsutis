/**
 * Журнал посещаемости — серверная часть на Google Apps Script.
 * Заменяет bot.py: база = листы этой Google-таблицы, сервер = бесплатный Google.
 * Настройки — в «Свойствах скрипта» (Project Settings → Script Properties), см. README.md.
 */

const SCHEMA = {
  grps:     ['id', 'name', 'url'],
  students: ['id', 'group_id', 'name', 'tg_id'],
  lessons:  ['id', 'group_id', 'title', 'teacher', 'room', 'ts_start', 'ts_end', 'uid'],
  marks:    ['lesson_id', 'student_id', 'status', 'ts'],
};
const TEXT_COLS = ['name', 'url', 'title', 'teacher', 'room', 'uid', 'status'];

const PROPS = PropertiesService.getScriptProperties();
const cfg = (k, d) => PROPS.getProperty(k) || d;
const TOKEN = cfg('BOT_TOKEN', '');
const OWNERS = cfg('OWNER_IDS', '').replace(/\s/g, '').split(',').filter(Boolean).map(Number);
const TZ = cfg('TZ_NAME', 'Asia/Novosibirsk');
const TPL = cfg('SCHEDULE_URL_TEMPLATE', '');
const BEFORE = Number(cfg('MARK_BEFORE_MIN', '15'));
const AFTER = Number(cfg('MARK_AFTER_MIN', '30'));

class Err extends Error {}

// ---------------- «БАЗА» (листы таблицы) ----------------
function sheet(n) {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(n);
  if (!s) {
    s = ss.insertSheet(n);
    s.appendRow(SCHEMA[n]);
    s.setFrozenRows(1);
    SCHEMA[n].forEach((k, i) => { if (TEXT_COLS.indexOf(k) >= 0) s.getRange(1, i + 1, s.getMaxRows(), 1).setNumberFormat('@'); });
  }
  return s;
}
function all(n) {
  const s = sheet(n), last = s.getLastRow();
  if (last < 2) return [];
  return s.getRange(2, 1, last - 1, SCHEMA[n].length).getValues().map(r => {
    const o = {}; SCHEMA[n].forEach((k, i) => o[k] = r[i]); return o;
  });
}
function saveAll(n, rows) {
  const s = sheet(n), cols = SCHEMA[n].length, last = s.getLastRow();
  if (last > 1) s.getRange(2, 1, last - 1, cols).clearContent();
  if (rows.length) s.getRange(2, 1, rows.length, cols)
    .setValues(rows.map(o => SCHEMA[n].map(k => (o[k] === undefined || o[k] === null) ? '' : o[k])));
}
const nextId = rows => rows.reduce((m, r) => Math.max(m, Number(r.id) || 0), 0) + 1;
const clean = s => String(s == null ? '' : s).split(/\s+/).filter(Boolean).join(' ');
const byName = (a, b) => String(a.name).localeCompare(String(b.name), 'ru');

// ---------------- ВРЕМЯ ----------------
const fmt = ts => Utilities.formatDate(new Date(ts * 1000), TZ, 'HH:mm');
const today = () => Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
const toTs = (date, hm) => Utilities.parseDate(date + ' ' + hm, TZ, 'yyyy-MM-dd HH:mm').getTime() / 1000;
const toTsFull = (date, hms) => Utilities.parseDate(date + ' ' + hms, TZ, 'yyyy-MM-dd HH:mm:ss').getTime() / 1000;
function addDay(ds, n) {
  const p = ds.split('-').map(Number);
  return new Date(Date.UTC(p[0], p[1] - 1, p[2] + n)).toISOString().slice(0, 10);
}
const dayBounds = ds => [toTs(ds, '00:00'), toTs(addDay(ds, 1), '00:00')];

// ---------------- АВТОРИЗАЦИЯ MINI APP ----------------
function verify(init) {
  if (!init || !TOKEN) return null;
  const d = {};
  init.split('&').forEach(p => {
    const i = p.indexOf('='); if (i < 0) return;
    d[decodeURIComponent(p.slice(0, i))] = decodeURIComponent(p.slice(i + 1).replace(/\+/g, ' '));
  });
  const h = d.hash; delete d.hash;
  if (!h || !d.user) return null;
  const s = Object.keys(d).sort().map(k => k + '=' + d[k]).join('\n');
  const key = Utilities.computeHmacSha256Signature(TOKEN, 'WebAppData');
  const sig = Utilities.computeHmacSha256Signature(s, key);
  const hex = sig.map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('');
  if (hex !== h) return null;
  if (Date.now() / 1000 - Number(d.auth_date || 0) > 86400) return null;
  return JSON.parse(d.user);
}
const isOwner = u => OWNERS.indexOf(u.id) >= 0;
const student = u => all('students').filter(s => Number(s.tg_id) === u.id)[0] || null;

// ---------------- API ----------------
function aMe(u) {
  const s = student(u);
  const r = { owner: isOwner(u), today: today(), tpl: !!TPL, before: BEFORE, after: AFTER, student: null,
    groups: all('grps').sort(byName).map(g => ({ id: g.id, name: g.name })) };
  if (s) {
    const g = all('grps').filter(x => Number(x.id) === Number(s.group_id))[0];
    r.student = { id: s.id, name: s.name, group: g ? g.name : '' };
  }
  return r;
}

function aFree(u, b) {
  return { students: all('students').filter(s => Number(s.group_id) === Number(b.group_id) && !s.tg_id)
    .sort(byName).map(s => ({ id: s.id, name: s.name })) };
}

function aRegister(u, b) {
  if (student(u)) throw new Err('Вы уже зарегистрированы');
  const gid = Number(b.group_id), rows = all('students');
  if (b.student_id) {
    const s = rows.filter(x => Number(x.id) === Number(b.student_id) && Number(x.group_id) === gid)[0];
    if (!s || s.tg_id) throw new Err('Это имя уже занято');
    s.tg_id = u.id;
  } else {
    const name = clean(b.name);
    if (name.length < 3) throw new Err('Введите фамилию и имя');
    if (rows.some(x => Number(x.group_id) === gid && x.name === name))
      throw new Err('Такое имя уже есть в списке — выберите его из списка');
    rows.push({ id: nextId(rows), group_id: gid, name: name, tg_id: u.id });
  }
  saveAll('students', rows);
  return {};
}

function aLessons(u, b) {
  const s = student(u), asTeacher = isOwner(u) && b.group_id;
  let gid;
  if (asTeacher) gid = Number(b.group_id); else if (s) gid = Number(s.group_id); else throw new Err('Сначала зарегистрируйтесь');
  const bounds = dayBounds(b.date), now = Date.now() / 1000;
  const total = all('students').filter(x => Number(x.group_id) === gid).length;
  const marks = all('marks');
  const lessons = all('lessons').filter(l => Number(l.group_id) === gid && l.ts_start >= bounds[0] && l.ts_start < bounds[1])
    .sort((x, y) => x.ts_start - y.ts_start).map(l => {
      const d = { id: l.id, title: l.title, teacher: l.teacher, room: l.room, start: fmt(l.ts_start), end: fmt(l.ts_end) };
      const ms = marks.filter(m => Number(m.lesson_id) === Number(l.id));
      if (asTeacher) {
        d.present = ms.filter(m => m.status === 'present').length;
        d.absent = ms.filter(m => m.status === 'absent').length;
        d.total = total;
      } else {
        const m = ms.filter(x => Number(x.student_id) === Number(s.id))[0];
        d.mark = m ? m.status : null;
        d.can_mark = l.ts_start - BEFORE * 60 <= now && now <= l.ts_end + AFTER * 60;
      }
      return d;
    });
  return { lessons: lessons };
}

function setMark(lessonId, studentId, status) {
  const marks = all('marks');
  const m = marks.filter(x => Number(x.lesson_id) === lessonId && Number(x.student_id) === studentId)[0];
  const ts = Math.floor(Date.now() / 1000);
  if (m) { m.status = status; m.ts = ts; } else marks.push({ lesson_id: lessonId, student_id: studentId, status: status, ts: ts });
  saveAll('marks', marks);
}

function aMark(u, b) {
  const s = student(u);
  if (!s) throw new Err('Сначала зарегистрируйтесь');
  const l = all('lessons').filter(x => Number(x.id) === Number(b.lesson_id) && Number(x.group_id) === Number(s.group_id))[0];
  if (!l) throw new Err('Пара не найдена');
  if (b.status !== 'present' && b.status !== 'absent') throw new Err('Неверный статус');
  const now = Date.now() / 1000;
  if (!(l.ts_start - BEFORE * 60 <= now && now <= l.ts_end + AFTER * 60)) throw new Err('Сейчас отметка недоступна');
  setMark(Number(l.id), Number(s.id), b.status);
  return {};
}

// ---- преподаватель ----
function aGroups() {
  const studs = all('students');
  return { groups: all('grps').sort(byName).map(g => ({
    id: g.id, name: g.name, url: g.url,
    students: studs.filter(s => Number(s.group_id) === Number(g.id)).sort(byName)
      .map(s => ({ id: s.id, name: s.name, linked: s.tg_id ? 1 : 0 })) })) };
}

function aGroupAdd(u, b) {
  const name = clean(b.name);
  if (!name) throw new Err('Введите название группы');
  const rows = all('grps');
  if (rows.some(g => g.name === name)) throw new Err('Такая группа уже есть');
  rows.push({ id: nextId(rows), name: name, url: clean(b.url) });
  saveAll('grps', rows);
  return {};
}

function dropLessons(ids) {
  saveAll('marks', all('marks').filter(m => ids.indexOf(Number(m.lesson_id)) < 0));
  saveAll('lessons', all('lessons').filter(l => ids.indexOf(Number(l.id)) < 0));
}
function dropStudents(ids) {
  saveAll('marks', all('marks').filter(m => ids.indexOf(Number(m.student_id)) < 0));
  saveAll('students', all('students').filter(s => ids.indexOf(Number(s.id)) < 0));
}

function aGroupDel(u, b) {
  const gid = Number(b.group_id);
  dropLessons(all('lessons').filter(l => Number(l.group_id) === gid).map(l => Number(l.id)));
  dropStudents(all('students').filter(s => Number(s.group_id) === gid).map(s => Number(s.id)));
  saveAll('grps', all('grps').filter(g => Number(g.id) !== gid));
  return {};
}

function aStudentsAdd(u, b) {
  const gid = Number(b.group_id), rows = all('students');
  let n = 0, id = nextId(rows);
  String(b.names || '').split(/\r?\n/).forEach(line => {
    const name = clean(line.replace(/^[0-9.)\-\s]+/, ''));
    if (name.length > 2 && !rows.some(r => Number(r.group_id) === gid && r.name === name)) {
      rows.push({ id: id++, group_id: gid, name: name, tg_id: '' }); n++;
    }
  });
  saveAll('students', rows);
  return { added: n };
}

function aStudentDel(u, b) { dropStudents([Number(b.student_id)]); return {}; }

function aStudentUnlink(u, b) {
  const rows = all('students');
  rows.forEach(s => { if (Number(s.id) === Number(b.student_id)) s.tg_id = ''; });
  saveAll('students', rows);
  return {};
}

function aLessonAdd(u, b) {
  const title = clean(b.title);
  if (!title) throw new Err('Введите название пары');
  const ts = toTs(b.date, b.start), te = toTs(b.date, b.end);
  if (te <= ts) throw new Err('Конец пары раньше начала');
  const rows = all('lessons');
  rows.push({ id: nextId(rows), group_id: Number(b.group_id), title: title, teacher: clean(b.teacher),
    room: clean(b.room), ts_start: ts, ts_end: te, uid: '' });
  saveAll('lessons', rows);
  return {};
}

function aLessonDel(u, b) { dropLessons([Number(b.lesson_id)]); return {}; }

function aAttendance(u, b) {
  const l = all('lessons').filter(x => Number(x.id) === Number(b.lesson_id))[0];
  if (!l) throw new Err('Пара не найдена');
  const marks = all('marks').filter(m => Number(m.lesson_id) === Number(l.id));
  return {
    lesson: { id: l.id, title: l.title, start: fmt(l.ts_start), end: fmt(l.ts_end) },
    students: all('students').filter(s => Number(s.group_id) === Number(l.group_id)).sort(byName).map(s => {
      const m = marks.filter(x => Number(x.student_id) === Number(s.id))[0];
      return { id: s.id, name: s.name, linked: !!s.tg_id, status: m ? m.status : null, time: m ? fmt(m.ts) : '' };
    }),
  };
}

function aOverride(u, b) {
  const lid = Number(b.lesson_id), sid = Number(b.student_id);
  if (b.status === 'none') saveAll('marks', all('marks').filter(m => !(Number(m.lesson_id) === lid && Number(m.student_id) === sid)));
  else if (b.status === 'present' || b.status === 'absent') setMark(lid, sid, b.status);
  else throw new Err('Неверный статус');
  return {};
}

// ---------------- ИМПОРТ РАСПИСАНИЯ (ICS) ----------------
function parseDT(v, p) {
  const m = String(v).match(/^(\d{4})(\d{2})(\d{2})(?:T(\d{2})(\d{2})(\d{2})?(Z)?)?$/);
  if (!m || !m[4]) return null; // события на весь день пропускаем
  const str = m[1] + '-' + m[2] + '-' + m[3] + ' ' + m[4] + ':' + m[5] + ':' + (m[6] || '00');
  let tz = TZ;
  if (m[7]) tz = 'UTC'; else if (p && p.TZID) tz = p.TZID.replace(/"/g, '');
  let t;
  try { t = Utilities.parseDate(str, tz, 'yyyy-MM-dd HH:mm:ss'); } catch (e) { t = Utilities.parseDate(str, TZ, 'yyyy-MM-dd HH:mm:ss'); }
  return { ts: t.getTime() / 1000, date: Utilities.formatDate(t, TZ, 'yyyy-MM-dd'), time: Utilities.formatDate(t, TZ, 'HH:mm:ss') };
}

function parseICS(text) {
  const lines = text.replace(/\r?\n[ \t]/g, '').split(/\r?\n/);
  const events = []; let cur = null, depth = 0;
  lines.forEach(line => {
    if (line === 'BEGIN:VEVENT') { cur = { ex: [] }; depth = 0; return; }
    if (!cur) return;
    if (line === 'END:VEVENT') { events.push(cur); cur = null; return; }
    if (line.indexOf('BEGIN:') === 0) { depth++; return; }
    if (line.indexOf('END:') === 0) { depth--; return; }
    if (depth > 0) return; // вложенные блоки (VALARM)
    const i = line.indexOf(':'); if (i < 0) return;
    const parts = line.slice(0, i).split(';'), name = parts[0].toUpperCase(), value = line.slice(i + 1), params = {};
    parts.slice(1).forEach(x => { const kv = x.split('='); params[kv[0].toUpperCase()] = kv[1]; });
    if (name === 'EXDATE') value.split(',').forEach(v => { const d = parseDT(v, params); if (d) cur.ex.push(d.ts); });
    else cur[name] = { value: value.replace(/\\n/gi, ' ').replace(/\\([,;\\])/g, '$1'), params: params };
  });
  return events;
}

// Повторяющиеся события: поддерживаются DAILY и WEEKLY (INTERVAL, COUNT, UNTIL, BYDAY, EXDATE)
function occurrences(ev, s, from, to) {
  if (!ev.RRULE) return (s.ts >= from && s.ts <= to) ? [s.ts] : [];
  const rule = {};
  ev.RRULE.value.split(';').forEach(x => { const kv = x.split('='); rule[kv[0]] = kv[1]; });
  if (rule.FREQ !== 'DAILY' && rule.FREQ !== 'WEEKLY') return (s.ts >= from && s.ts <= to) ? [s.ts] : [];
  const interval = Number(rule.INTERVAL) || 1, count = Number(rule.COUNT) || Infinity;
  const until = rule.UNTIL ? parseDT(rule.UNTIL.length === 8 ? rule.UNTIL + 'T235959' : rule.UNTIL, {}) : null;
  const names = ['SU', 'MO', 'TU', 'WE', 'TH', 'FR', 'SA'];
  const dow0 = new Date(s.date + 'T00:00:00Z').getUTCDay();
  const byday = rule.BYDAY ? rule.BYDAY.split(',').map(x => names.indexOf(x.slice(-2))) : [dow0];
  const dFrom = Utilities.formatDate(new Date(from * 1000), TZ, 'yyyy-MM-dd');
  const dTo = Utilities.formatDate(new Date(to * 1000), TZ, 'yyyy-MM-dd');
  const out = []; let n = 0;
  for (let i = 0; i < 5000; i++) {
    const date = addDay(s.date, i);
    if (date > dTo || (until && date > until.date)) break;
    let ok;
    if (rule.FREQ === 'DAILY') ok = i % interval === 0;
    else ok = byday.indexOf(new Date(date + 'T00:00:00Z').getUTCDay()) >= 0 && Math.floor((i + (dow0 + 6) % 7) / 7) % interval === 0;
    if (!ok) continue;
    if (++n > count) break;
    if (date < dFrom) continue;
    const ts = toTsFull(date, s.time);
    if (ev.ex.indexOf(ts) < 0) out.push(ts);
  }
  return out;
}

function aImport(u, b) {
  const g = all('grps').filter(x => Number(x.id) === Number(b.group_id))[0];
  if (!g) throw new Err('Группа не найдена');
  const url = clean(b.url) || g.url || (TPL ? TPL.replace('{group}', encodeURIComponent(g.name)) : '');
  if (!url) throw new Err('Укажите ссылку на расписание (ICS)');
  if (clean(b.url)) { const gs = all('grps'); gs.forEach(x => { if (Number(x.id) === Number(g.id)) x.url = url; }); saveAll('grps', gs); }
  let resp;
  try { resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0' } }); }
  catch (e) { throw new Err('Не удалось открыть ссылку: ' + e.message); }
  if (resp.getResponseCode() !== 200) throw new Err('Сайт ответил кодом ' + resp.getResponseCode());
  const text = resp.getContentText('UTF-8');
  if (text.indexOf('BEGIN:VCALENDAR') < 0) throw new Err('Ссылка не ведёт на календарь в формате ICS');

  const from = Date.now() / 1000 - 7 * 86400, to = from + 150 * 86400;
  const rows = all('lessons'), idx = {};
  rows.forEach((l, i) => { if (Number(l.group_id) === Number(g.id) && l.uid) idx[l.uid] = i; });
  let id = nextId(rows), n = 0;
  parseICS(text).forEach(ev => {
    if (!ev.DTSTART) return;
    const s = parseDT(ev.DTSTART.value, ev.DTSTART.params);
    if (!s) return;
    const e = ev.DTEND ? parseDT(ev.DTEND.value, ev.DTEND.params) : null;
    const dur = e && e.ts > s.ts ? e.ts - s.ts : 5400;
    const uid0 = ev.UID ? ev.UID.value : 'x';
    occurrences(ev, s, from, to).forEach(ts => {
      const rec = { group_id: Number(g.id), title: clean(ev.SUMMARY ? ev.SUMMARY.value : 'Пара'),
        teacher: clean(ev.DESCRIPTION ? ev.DESCRIPTION.value : '').slice(0, 120),
        room: clean(ev.LOCATION ? ev.LOCATION.value : ''), ts_start: ts, ts_end: ts + dur, uid: uid0 + '|' + ts };
      if (idx[rec.uid] !== undefined) Object.assign(rows[idx[rec.uid]], rec);
      else { rec.id = id++; idx[rec.uid] = rows.push(rec) - 1; }
      n++;
    });
  });
  saveAll('lessons', rows);
  return { imported: n };
}

const ROUTES = { // имя: [функция, только_владелец, пишет_в_базу]
  me: [aMe, false, false], free: [aFree, false, false], register: [aRegister, false, true],
  lessons: [aLessons, false, false], mark: [aMark, false, true],
  groups: [aGroups, true, false], group_add: [aGroupAdd, true, true], group_del: [aGroupDel, true, true],
  students_add: [aStudentsAdd, true, true], student_del: [aStudentDel, true, true],
  student_unlink: [aStudentUnlink, true, true], lesson_add: [aLessonAdd, true, true],
  lesson_del: [aLessonDel, true, true], attendance: [aAttendance, true, false],
  override: [aOverride, true, true], import: [aImport, true, true],
};

const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);

function doGet() { return ContentService.createTextOutput('ok'); }

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    const u = verify(req.init);
    if (!u) return out({ error: 'Откройте приложение через Telegram' });
    const route = ROUTES[req.a];
    if (!route) return out({ error: 'Не найдено' });
    if (route[1] && !isOwner(u)) return out({ error: 'Только для преподавателя' });
    const lock = LockService.getScriptLock();
    if (route[2]) lock.waitLock(25000);
    try { return out(route[0](u, req.b || {})); }
    finally { if (route[2]) lock.releaseLock(); }
  } catch (err) {
    if (err instanceof Err) return out({ error: err.message });
    console.error(err);
    return out({ error: 'Ошибка сервера' });
  }
}

// ---------------- ОТЧЁТЫ (замена /export) ----------------
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Журнал').addItem('Обновить отчёты по группам', 'buildReport').addToUi();
}

function buildReport() {
  const ss = SpreadsheetApp.getActive(), marks = all('marks'), studs = all('students');
  const sym = { present: '+', absent: 'н' };
  all('grps').forEach(g => {
    const name = ('Отчёт ' + g.name).replace(/[\[\]*?:\/\\]/g, '_').slice(0, 99);
    let sh = ss.getSheetByName(name); if (sh) sh.clear(); else sh = ss.insertSheet(name);
    const lessons = all('lessons').filter(l => Number(l.group_id) === Number(g.id)).sort((a, b) => a.ts_start - b.ts_start);
    const gs = studs.filter(s => Number(s.group_id) === Number(g.id)).sort(byName);
    const head = ['Студент'].concat(lessons.map(l => Utilities.formatDate(new Date(l.ts_start * 1000), TZ, 'dd.MM HH:mm') + ' ' + l.title));
    const body = gs.map(s => [s.name].concat(lessons.map(l => {
      const m = marks.filter(x => Number(x.lesson_id) === Number(l.id) && Number(x.student_id) === Number(s.id))[0];
      return m ? sym[m.status] : '';
    })));
    sh.getRange(1, 1, body.length + 1, head.length).setValues([head].concat(body));
    sh.setFrozenRows(1); sh.setFrozenColumns(1);
  });
}
