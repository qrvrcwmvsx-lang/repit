/**
 * Бэкенд сайта: окна берутся из Apple Календаря, записи — в Google Таблицу,
 * а на каждую запись вам приходит приглашение в календарь на почту iCloud.
 *
 *   Apple Календарь «Окна»  ──(публичная ссылка)──►  этот скрипт  ◄──  сайт
 *                                                       │
 *                     лист «Записи» ◄───────────────────┤
 *       приглашение на iCloud (через Google Календарь) ◄─┤
 *                            Telegram (по желанию) ◄────┘
 *
 * Как подключить — в README.md.
 */

// ====== Настройки ======
const CALENDAR_URL = '';       // публичная ссылка на календарь «Окна» из Apple Календаря (webcal://…)
const NOTIFY_EMAIL = '';       // ваша почта iCloud — туда придёт приглашение на каждый урок
const TELEGRAM_BOT_TOKEN = ''; // по желанию: токен от @BotFather
const TELEGRAM_CHAT_ID = '';   // по желанию: ваш chat id (узнать у @userinfobot)

const SLOT_MINUTES = 60;       // длина урока: событие 15:00–19:00 превратится в 4 окна
const MIN_HOURS_BEFORE = 3;    // окна, до которых меньше 3 часов, на сайте не показываются
const DAYS_AHEAD = 30;         // насколько вперёд показывать расписание

// ====== Дальше лучше не менять ======
const LESSONS_CALENDAR = 'Уроки с сайта'; // календарь Google, из которого уходят приглашения
const BOOKINGS_SHEET = 'Записи';
const COL = { CREATED: 1, DATE: 2, TIME: 3, NAME: 4, TG: 5, COMMENT: 6, STATUS: 7 };
const NEW = 'новая';
const CANCELLED = 'отменена';
const WEEKDAYS = ['SU', 'MO', 'TU', 'WE', 'TH', 'FR', 'SA'];
const MIN = 60 * 1000;
const DAY = 24 * 60 * MIN;

// ---------- Сайт → скрипт ----------

function doGet() {
  try {
    const slots = freeSlots_().map((s) => ({ id: s.id, date: s.date, time: s.time, duration: s.duration }));
    return json_({ ok: true, slots });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); // две заявки на одно окно в одну секунду не пройдут
    const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');

    if (data.website) return json_({ ok: true }); // спам-бот заполнил скрытое поле

    const name = clean_(data.name, 80);
    const telegram = clean_(data.telegram, 40).replace(/^@/, '');
    const comment = clean_(data.comment, 1000);
    const slotId = clean_(data.slotId, 40);

    if (name.length < 2) return json_({ ok: false, error: 'name' });
    if (!/^[A-Za-z0-9_]{4,32}$/.test(telegram)) return json_({ ok: false, error: 'telegram' });

    const slot = freeSlots_().find((s) => s.id === slotId);
    if (!slot) return json_({ ok: false, error: 'taken' });

    // апостроф — чтобы таблица не превращала дату и время в свои форматы
    getSheet_(BOOKINGS_SHEET).appendRow([new Date(), "'" + slot.date, "'" + slot.time, cell_(name), '@' + telegram, cell_(comment), NEW]);
    SpreadsheetApp.flush();
    lock.releaseLock();

    // запись уже сохранена: ошибки уведомлений не должны ломать ответ ученику
    try { sendInvite_(slot, name, telegram, comment); } catch (err) { console.error('Приглашение:', err); }
    try { notifyTelegram_(slot, name, telegram, comment); } catch (err) { console.error('Telegram:', err); }
    return json_({ ok: true });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (_) { /* уже отпущен */ }
  }
}

// ---------- Свободные окна ----------

/** Окна из календаря, нарезанные на уроки, минус уже занятые. */
function freeSlots_() {
  const tz = tz_();
  const now = Date.now();
  const from = now + MIN_HOURS_BEFORE * 60 * MIN;
  const to = now + DAYS_AHEAD * DAY;
  const booked = bookedIds_();
  const step = SLOT_MINUTES * MIN;
  const seen = new Set();
  const out = [];

  windows_(from, to).forEach((w) => {
    for (let t = w.start; t + step <= w.end; t += step) {
      if (t < from || t > to) continue;
      const date = Utilities.formatDate(new Date(t), tz, 'yyyy-MM-dd');
      const time = Utilities.formatDate(new Date(t), tz, 'HH:mm');
      const id = `${date} ${time}`;
      if (seen.has(id) || booked.has(id)) continue;
      seen.add(id);
      out.push({ id, date, time, duration: SLOT_MINUTES, start: t });
    }
  });
  return out.sort((a, b) => a.start - b.start);
}

function bookedIds_() {
  const sheet = getSheet_(BOOKINGS_SHEET);
  const last = sheet.getLastRow();
  if (last < 2) return new Set();
  const rows = sheet.getRange(2, COL.DATE, last - 1, COL.STATUS - COL.DATE + 1).getDisplayValues();
  const ids = new Set();
  rows.forEach((r) => {
    const status = String(r[COL.STATUS - COL.DATE] || '').trim().toLowerCase();
    if (status !== CANCELLED && r[0] && r[1]) ids.add(`${r[0].trim()} ${r[1].trim()}`);
  });
  return ids;
}

/** Все события календаря «Окна» в промежутке [from, to): { start, end } в миллисекундах. */
function windows_(from, to) {
  if (!CALENDAR_URL) throw new Error('Не указана ссылка CALENDAR_URL');
  const events = parseIcs_(fetchIcs_());
  return expandEvents_(events, from, to);
}

function fetchIcs_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('ics');
  if (cached) return cached;
  const url = CALENDAR_URL.trim().replace(/^webcals?:\/\//i, 'https://');
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
  if (res.getResponseCode() !== 200) throw new Error(`Календарь не открылся (код ${res.getResponseCode()}). Проверьте, что он публичный.`);
  const text = res.getContentText('UTF-8');
  if (text.indexOf('BEGIN:VCALENDAR') === -1) throw new Error('По ссылке не календарь');
  try { cache.put('ics', text, 60); } catch (_) { /* большой календарь просто не кэшируем */ }
  return text;
}

// ---------- Разбор календаря (формат iCalendar / .ics) ----------

function parseIcs_(text) {
  const lines = text.replace(/\r?\n[ \t]/g, '').split(/\r?\n/);
  const events = [];
  let ev = null;
  let depth = 0;

  lines.forEach((line) => {
    if (line === 'BEGIN:VEVENT') { ev = { exdates: [] }; depth = 0; return; }
    if (!ev) return;
    if (/^BEGIN:/.test(line)) { depth++; return; }        // например, VALARM внутри события
    if (/^END:/.test(line)) {
      if (depth === 0 && line === 'END:VEVENT') { events.push(ev); ev = null; } else depth--;
      return;
    }
    if (depth > 0) return;

    const m = line.match(/^([A-Za-z-]+)((?:;(?:[^:;"]|"[^"]*")*)*):(.*)$/);
    if (!m) return;
    const prop = m[1].toUpperCase();
    const params = {};
    m[2].split(';').slice(1).forEach((p) => {
      const i = p.indexOf('=');
      if (i > 0) params[p.slice(0, i).toUpperCase()] = p.slice(i + 1).replace(/^"|"$/g, '');
    });
    const value = m[3];

    if (prop === 'DTSTART') {
      const d = parseIcsDate_(value, params);
      if (d) { ev.start = d.ms; ev.allDay = d.allDay; ev.zone = d.zone; }
    } else if (prop === 'DTEND') {
      const d = parseIcsDate_(value, params);
      if (d) ev.end = d.ms;
    } else if (prop === 'DURATION') {
      ev.durationMs = parseDuration_(value);
    } else if (prop === 'RRULE') {
      ev.rrule = {};
      value.split(';').forEach((p) => { const [k, v] = p.split('='); if (k && v) ev.rrule[k.toUpperCase()] = v; });
    } else if (prop === 'EXDATE') {
      value.split(',').forEach((v) => { const d = parseIcsDate_(v, params); if (d) ev.exdates.push(d.ms); });
    } else if (prop === 'RECURRENCE-ID') {
      const d = parseIcsDate_(value, params);
      if (d) ev.recurrenceId = d.ms;
    } else if (prop === 'STATUS') {
      ev.status = value.toUpperCase();
    } else if (prop === 'UID') {
      ev.uid = value;
    }
  });
  return events;
}

/** 20261014T180000Z | 20261014T180000 (+TZID) | 20261014 → { ms, allDay, zone } */
function parseIcsDate_(value, params, fallbackZone) {
  const m = String(value).trim().match(/^(\d{4})(\d{2})(\d{2})(?:T(\d{2})(\d{2})(\d{2})(Z)?)?$/);
  if (!m) return null;
  const zone = m[7] ? 'UTC' : zoneOrDefault_(params && params.TZID, fallbackZone);
  if (!m[4]) return { ms: localMs_(`${m[1]}-${m[2]}-${m[3]} 00:00:00`, zone), allDay: true, zone };
  if (m[7]) return { ms: Date.UTC(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +m[6]), allDay: false, zone: tz_() };
  return { ms: localMs_(`${m[1]}-${m[2]}-${m[3]} ${m[4]}:${m[5]}:${m[6]}`, zone), allDay: false, zone };
}

/** Apple пишет часовые пояса как Europe/Moscow. Непонятные названия заменяем поясом таблицы. */
function zoneOrDefault_(tzid, fallback) {
  const z = String(tzid || '').trim();
  return /^[A-Za-z]+(?:\/[A-Za-z0-9_+\-]+)+$/.test(z) || z === 'UTC' ? z : (fallback || tz_());
}

function localMs_(s, zone) {
  return Utilities.parseDate(s, zone, 'yyyy-MM-dd HH:mm:ss').getTime();
}

function parseDuration_(v) {
  const m = String(v).match(/^P(?:(\d+)W)?(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?)?$/);
  if (!m) return 0;
  return ((+m[1] || 0) * 7 * DAY) + ((+m[2] || 0) * DAY) + ((+m[3] || 0) * 60 * MIN) + ((+m[4] || 0) * MIN) + ((+m[5] || 0) * 1000);
}

/** Разворачивает повторяющиеся события и убирает отменённые/перенесённые повторы. */
function expandEvents_(events, from, to) {
  const moved = {}; // uid → начала повторов, которые изменены отдельным событием
  events.forEach((e) => {
    if (e.recurrenceId != null && e.uid) (moved[e.uid] = moved[e.uid] || []).push(e.recurrenceId);
  });

  const out = [];
  events.forEach((e) => {
    if (e.start == null || e.allDay || e.status === 'CANCELLED') return;
    const len = e.end != null ? e.end - e.start : (e.durationMs || SLOT_MINUTES * MIN);
    if (len <= 0) return;

    const isSeries = e.rrule && e.recurrenceId == null;
    const starts = isSeries ? occurrences_(e, to) : [e.start];
    const skip = new Set(e.exdates.concat(isSeries ? (moved[e.uid] || []) : []));
    starts.forEach((s) => {
      if (!skip.has(s) && s + len > from && s < to) out.push({ start: s, end: s + len });
    });
  });
  return out;
}

/** Повторы DAILY / WEEKLY (с BYDAY) / MONTHLY (то же число) с INTERVAL, COUNT, UNTIL. */
function occurrences_(e, to) {
  const r = e.rrule;
  const zone = e.zone || tz_();
  const startIso = Utilities.formatDate(new Date(e.start), zone, 'yyyy-MM-dd');
  const clock = Utilities.formatDate(new Date(e.start), zone, 'HH:mm:ss');
  const interval = Math.max(1, Number(r.INTERVAL) || 1);
  const count = r.COUNT ? Number(r.COUNT) : Infinity;
  let until = Infinity;
  if (r.UNTIL) {
    const u = parseIcsDate_(r.UNTIL, {}, zone);
    if (u) until = u.allDay ? u.ms + DAY - 1 : u.ms;
  }
  const byday = r.BYDAY ? r.BYDAY.split(',').map((d) => WEEKDAYS.indexOf(d.slice(-2).toUpperCase())).filter((d) => d >= 0) : null;

  const d0 = dayNum_(startIso);
  const startDow = weekday_(d0);
  const out = [];
  let n = 0;
  for (let d = d0; d - d0 < 40000; d++) {
    const iso = isoFromDay_(d);
    const dow = weekday_(d);
    let match;
    if (r.FREQ === 'DAILY') {
      match = (d - d0) % interval === 0 && (!byday || byday.indexOf(dow) !== -1);
    } else if (r.FREQ === 'WEEKLY') {
      const weeks = Math.round((mondayOf_(d) - mondayOf_(d0)) / 7);
      match = weeks % interval === 0 && (byday || [startDow]).indexOf(dow) !== -1;
    } else if (r.FREQ === 'MONTHLY') {
      const months = (+iso.slice(0, 4) - +startIso.slice(0, 4)) * 12 + (+iso.slice(5, 7) - +startIso.slice(5, 7));
      match = iso.slice(8) === startIso.slice(8) && months % interval === 0;
    } else {
      return [e.start]; // редкие правила (YEARLY и т.п.) — берём только первое событие
    }
    if (!match) continue;
    const ms = localMs_(`${iso} ${clock}`, zone);
    if (ms < e.start) continue;
    if (ms > until || n >= count || ms >= to) break;
    n++;
    out.push(ms);
  }
  return out;
}

function dayNum_(iso) { return Date.UTC(+iso.slice(0, 4), +iso.slice(5, 7) - 1, +iso.slice(8, 10)) / DAY; }
function isoFromDay_(n) { return new Date(n * DAY).toISOString().slice(0, 10); }
function weekday_(n) { return new Date(n * DAY).getUTCDay(); }
function mondayOf_(n) { return n - ((weekday_(n) + 6) % 7); }

// ---------- Уведомления ----------

/** Создаёт урок в Google Календаре с вами в гостях — Google сам отправит приглашение на iCloud. */
function sendInvite_(slot, name, telegram, comment) {
  if (!NOTIFY_EMAIL) return;
  const start = new Date(slot.start);
  const end = new Date(slot.start + slot.duration * MIN);
  const description = [
    `Ученик: ${name}`,
    `Telegram: https://t.me/${telegram}`,
    comment ? `Комментарий: ${comment}` : '',
    '',
    'Запись с сайта. Отменить: в таблице «Записи» поставьте статус «отменена» — окно снова станет свободным.',
  ].filter((l, i) => l !== '' || i === 3).join('\n');

  lessonsCalendar_().createEvent(`Урок: ${name} (@${telegram})`, start, end, {
    description,
    guests: NOTIFY_EMAIL,
    sendInvites: true,
  });
}

function lessonsCalendar_() {
  return CalendarApp.getCalendarsByName(LESSONS_CALENDAR)[0]
    || CalendarApp.createCalendar(LESSONS_CALENDAR, { timeZone: tz_() });
}

function notifyTelegram_(slot, name, telegram, comment) {
  if (!TELEGRAM_BOT_TOKEN || !TELEGRAM_CHAT_ID) return;
  const d = new Date(dayNum_(slot.date) * DAY);
  const days = ['вс', 'пн', 'вт', 'ср', 'чт', 'пт', 'сб'];
  const pretty = `${days[d.getUTCDay()]}, ${slot.date.slice(8)}.${slot.date.slice(5, 7)} в ${slot.time}`;
  const text = [
    '🫒 <b>Новая запись на урок</b>',
    '',
    `🗓 ${pretty} (${slot.duration} мин)`,
    `👤 ${esc_(name)}`,
    `✈️ <a href="https://t.me/${telegram}">@${telegram}</a>`,
    comment ? `💬 ${esc_(comment)}` : '',
  ].filter((l, i) => l !== '' || i === 1).join('\n');

  UrlFetchApp.fetch(`https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage`, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ chat_id: TELEGRAM_CHAT_ID, text, parse_mode: 'HTML', disable_web_page_preview: true }),
    muteHttpExceptions: true,
  });
}

// ---------- Меню в таблице ----------

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🫒 Расписание')
    .addItem('Показать свободные окна', 'showFreeSlots')
    .addItem('Отправить тестовое приглашение', 'testInvite')
    .addItem('Проверить уведомление в Telegram', 'testTelegram')
    .addSeparator()
    .addItem('Первичная настройка', 'setup')
    .addToUi();
}

/** Проверка связи с Apple Календарём: что сейчас видит сайт. */
function showFreeSlots() {
  const ui = SpreadsheetApp.getUi();
  try {
    const slots = freeSlots_();
    if (!slots.length) return ui.alert('Календарь открылся, но свободных окон на ближайшие дни нет.\n\nДобавьте события в календарь «Окна». Изменения доходят до сайта через несколько минут.');
    const list = slots.slice(0, 15).map((s) => `${s.date.slice(8)}.${s.date.slice(5, 7)}  ${s.time}`).join('\n');
    ui.alert(`Свободных окон: ${slots.length}\n\nБлижайшие:\n${list}`);
  } catch (err) {
    ui.alert(`Не получилось прочитать календарь:\n${err.message || err}`);
  }
}

function testInvite() {
  const ui = SpreadsheetApp.getUi();
  if (!NOTIFY_EMAIL) return ui.alert('Сначала впишите NOTIFY_EMAIL (вашу почту iCloud) в начале скрипта.');
  const start = Math.ceil((Date.now() + DAY) / (60 * MIN)) * 60 * MIN;
  const tz = tz_();
  sendInvite_({
    date: Utilities.formatDate(new Date(start), tz, 'yyyy-MM-dd'),
    time: Utilities.formatDate(new Date(start), tz, 'HH:mm'),
    duration: SLOT_MINUTES,
    start,
  }, 'Тестовый ученик', 'durov', 'Это проверка — приглашение можно отклонить 🌿');
  ui.alert(`Приглашение отправлено на ${NOTIFY_EMAIL}.\nЧерез минуту проверьте Календарь на iPhone — раздел «Входящие».`);
}

function testTelegram() {
  if (!TELEGRAM_BOT_TOKEN || !TELEGRAM_CHAT_ID) {
    return SpreadsheetApp.getUi().alert('Сначала впишите TELEGRAM_BOT_TOKEN и TELEGRAM_CHAT_ID в начале скрипта.');
  }
  notifyTelegram_({ date: Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd'), time: '18:00', duration: SLOT_MINUTES }, 'Тестовый ученик', 'durov', 'Это проверка — всё работает 🌿');
  toast_('Отправил тестовое сообщение — проверьте Telegram');
}

/** Создаёт лист «Записи» с оформлением и календарь для приглашений. Запускается один раз. */
function setup() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(BOOKINGS_SHEET);
  if (!sheet) {
    const first = ss.getSheets()[0];
    sheet = first.getLastRow() === 0 ? first.setName(BOOKINGS_SHEET) : ss.insertSheet(BOOKINGS_SHEET);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Когда пришла', 'Дата урока', 'Время', 'Имя', 'Telegram', 'Комментарий', 'Статус']);
  }
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, 7).setBackground('#56612E').setFontColor('#FAF6EE').setFontWeight('bold');
  sheet.getRange('A2:A').setNumberFormat('dd.mm.yyyy hh:mm');
  sheet.getRange('B2:C').setNumberFormat('@');
  sheet.getRange('G2:G').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList([NEW, 'подтверждена', CANCELLED], true).build()
  );
  sheet.getRange('G1').setNote('новая / подтверждена — время занято\nотменена — окно снова появится на сайте');
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=$G2="${CANCELLED}"`).setFontColor('#A39B8D').setStrikethrough(true)
      .setRanges([sheet.getRange('A2:G')]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=$G2="${NEW}"`).setBackground('#F7EAE9')
      .setRanges([sheet.getRange('A2:G')]).build(),
  ]);
  sheet.setColumnWidth(1, 140);
  sheet.setColumnWidth(6, 320);

  lessonsCalendar_(); // заодно Google спросит разрешение на календарь
  onOpen();
  toast_('Готово! Таблица настроена 🌿');
}

// ---------- Вспомогательное ----------

function tz_() {
  return SpreadsheetApp.getActive().getSpreadsheetTimeZone();
}

function getSheet_(name) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sheet) throw new Error(`Нет листа «${name}». Запустите setup().`);
  return sheet;
}

function clean_(v, max) {
  return String(v == null ? '' : v).replace(/[\u0000-\u0009\u000B-\u001F\u007F]/g, ' ').trim().slice(0, max);
}

/** Не даём тексту из формы превратиться в формулу таблицы (=…, +…, -…, @…). */
function cell_(s) {
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function esc_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function toast_(msg) {
  SpreadsheetApp.getActive().toast(msg, '🫒 Расписание', 5);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
