/**
 * Бэкенд сайта-визитки: Google Таблица + Apps Script.
 *
 * Лист «Окна»   — вы добавляете сюда свободное время.
 * Лист «Записи» — сюда автоматически падают заявки с сайта.
 *
 * Как подключить — в README.md, раздел «Подключаем таблицу».
 */

// ====== Настройки ======
const TELEGRAM_BOT_TOKEN = ''; // токен от @BotFather, например '123456:ABC-...'
const TELEGRAM_CHAT_ID = '';   // ваш chat id (узнать у @userinfobot)

const MIN_HOURS_BEFORE = 3;   // окна, до которых меньше 3 часов, на сайте не показываются
const DAYS_AHEAD = 30;        // насколько вперёд показывать расписание
const DEFAULT_DURATION = 60;  // длительность, если в таблице не указана

// ====== Структура таблицы (лучше не менять) ======
const SLOTS_SHEET = 'Окна';
const BOOKINGS_SHEET = 'Записи';
const COL = { DATE: 1, TIME: 2, DURATION: 3, STATUS: 4, WHO: 5 };
const FREE = 'свободно';
const BUSY = 'занято';

// ---------- Сайт → таблица ----------

function doGet() {
  try {
    const now = new Date();
    const from = new Date(now.getTime() + MIN_HOURS_BEFORE * 3600 * 1000);
    const to = new Date(now.getTime() + DAYS_AHEAD * 24 * 3600 * 1000);
    const slots = readSlots_()
      .filter((s) => s.free && s.start >= from && s.start <= to)
      .map((s) => ({ id: s.id, date: s.date, time: s.time, duration: s.duration }));
    return json_({ ok: true, slots });
  } catch (err) {
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

    const minStart = new Date(Date.now() + MIN_HOURS_BEFORE * 3600 * 1000);
    const slot = readSlots_().find((s) => s.id === slotId);
    if (!slot || !slot.free || slot.start < minStart) return json_({ ok: false, error: 'taken' });

    const sheet = getSheet_(SLOTS_SHEET);
    sheet.getRange(slot.row, COL.STATUS).setValue(BUSY);
    sheet.getRange(slot.row, COL.WHO).setValue(cell_(`${name} · @${telegram}`));

    getSheet_(BOOKINGS_SHEET).appendRow([new Date(), slot.date, slot.time, cell_(name), '@' + telegram, cell_(comment)]);
    SpreadsheetApp.flush();

    notifyTelegram_(slot, name, telegram, comment);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// ---------- Чтение окон ----------

function readSlots_() {
  const sheet = getSheet_(SLOTS_SHEET);
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const range = sheet.getRange(2, 1, last - 1, COL.WHO);
  const values = range.getValues();
  const shown = range.getDisplayValues(); // время берём «как видно», чтобы не ловить сдвиги часовых поясов
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const out = [];

  values.forEach((row, i) => {
    const date = normalizeDate_(row[COL.DATE - 1], tz);
    const time = normalizeTime_(shown[i][COL.TIME - 1]);
    if (!date || !time) return; // пустые и кривые строки пропускаем

    const status = String(row[COL.STATUS - 1] || '').trim().toLowerCase();
    out.push({
      row: i + 2,
      id: `${date} ${time}`,
      date,
      time,
      duration: Number(row[COL.DURATION - 1]) || DEFAULT_DURATION,
      free: status === '' || status === FREE,
      start: Utilities.parseDate(`${date} ${time}`, tz, 'yyyy-MM-dd HH:mm'),
    });
  });
  return out;
}

function normalizeDate_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  const m = String(v || '').trim().match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})/); // дата, введённая текстом: 14.10.2026
  return m ? `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}` : '';
}

function normalizeTime_(v) {
  const m = String(v || '').trim().match(/^(\d{1,2})[:.](\d{2})/);
  if (!m) return '';
  const h = Number(m[1]);
  const min = Number(m[2]);
  if (h > 23 || min > 59) return '';
  return `${String(h).padStart(2, '0')}:${m[2]}`;
}

// ---------- Уведомление в Telegram ----------

function notifyTelegram_(slot, name, telegram, comment) {
  if (!TELEGRAM_BOT_TOKEN || !TELEGRAM_CHAT_ID) return;
  const d = Utilities.parseDate(slot.date, 'UTC', 'yyyy-MM-dd');
  const days = ['вс', 'пн', 'вт', 'ср', 'чт', 'пт', 'сб'];
  const pretty = `${days[d.getUTCDay()]}, ${Utilities.formatDate(d, 'UTC', 'dd.MM')} в ${slot.time}`;
  const text = [
    '🫒 <b>Новая запись на урок</b>',
    '',
    `🗓 ${pretty} (${slot.duration} мин)`,
    `👤 ${esc_(name)}`,
    `✈️ <a href="https://t.me/${telegram}">@${telegram}</a>`,
    comment ? `💬 ${esc_(comment)}` : '',
  ].filter((l, i, a) => l !== '' || i === 1).join('\n');

  try {
    UrlFetchApp.fetch(`https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage`, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ chat_id: TELEGRAM_CHAT_ID, text, parse_mode: 'HTML', disable_web_page_preview: true }),
      muteHttpExceptions: true,
    });
  } catch (err) {
    console.error('Telegram:', err); // запись уже сохранена, поэтому ошибку не пробрасываем
  }
}

// ---------- Меню в таблице ----------

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🫒 Расписание')
    .addItem('Повторить последнюю неделю', 'repeatLastWeek')
    .addItem('Убрать прошедшие окна', 'removePastSlots')
    .addSeparator()
    .addItem('Проверить уведомление в Telegram', 'testTelegram')
    .addItem('Первичная настройка таблицы', 'setup')
    .addToUi();
}

/** Копирует окна последних 7 дней расписания на неделю вперёд (статусы — пустые). */
function repeatLastWeek() {
  const slots = readSlots_();
  if (!slots.length) return toast_('В листе «Окна» пока нечего копировать');
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const lastDate = slots.map((s) => s.date).sort().pop();
  const weekStart = shiftDate_(lastDate, -6);
  const existing = new Set(slots.map((s) => s.id));

  const rows = slots
    .filter((s) => s.date >= weekStart && s.date <= lastDate)
    .map((s) => ({ ...s, date: shiftDate_(s.date, 7) }))
    .filter((s) => !existing.has(`${s.date} ${s.time}`))
    .sort((a, b) => (a.date + a.time).localeCompare(b.date + b.time))
    .map((s) => [Utilities.parseDate(s.date, tz, 'yyyy-MM-dd'), s.time, s.duration, FREE, '']);

  if (!rows.length) return toast_('Следующая неделя уже заполнена');
  const sheet = getSheet_(SLOTS_SHEET);
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  toast_(`Добавлено окон: ${rows.length}`);
}

/** Удаляет строки с окнами, которые уже прошли. Записи в листе «Записи» остаются. */
function removePastSlots() {
  const sheet = getSheet_(SLOTS_SHEET);
  const now = new Date();
  const past = readSlots_().filter((s) => s.start < now).map((s) => s.row).sort((a, b) => b - a);
  past.forEach((r) => sheet.deleteRow(r));
  toast_(past.length ? `Удалено прошедших окон: ${past.length}` : 'Прошедших окон нет');
}

function testTelegram() {
  if (!TELEGRAM_BOT_TOKEN || !TELEGRAM_CHAT_ID) {
    return SpreadsheetApp.getUi().alert('Сначала впишите TELEGRAM_BOT_TOKEN и TELEGRAM_CHAT_ID в начале скрипта.');
  }
  notifyTelegram_({ date: Utilities.formatDate(new Date(), 'UTC', 'yyyy-MM-dd'), time: '18:00', duration: 60 }, 'Тестовый ученик', 'durov', 'Это проверка — всё работает 🌿');
  toast_('Отправил тестовое сообщение — проверьте Telegram');
}

/** Создаёт листы с заголовками, оформлением, подсказками и примером окон. Запускается один раз. */
function setup() {
  const ss = SpreadsheetApp.getActive();

  let slots = ss.getSheetByName(SLOTS_SHEET);
  if (!slots) {
    slots = ss.getSheets()[0].getLastRow() === 0 ? ss.getSheets()[0].setName(SLOTS_SHEET) : ss.insertSheet(SLOTS_SHEET);
  }
  if (slots.getLastRow() === 0) {
    slots.appendRow(['Дата', 'Время', 'Длительность, мин', 'Статус', 'Кто записан']);
    const tz = ss.getSpreadsheetTimeZone();
    const base = new Date();
    const sample = [];
    for (let i = 1; i <= 5; i++) {
      const day = Utilities.parseDate(Utilities.formatDate(new Date(base.getTime() + i * 864e5), tz, 'yyyy-MM-dd'), tz, 'yyyy-MM-dd');
      ['16:00', '18:00'].forEach((t) => sample.push([day, t, 60, FREE, '']));
    }
    slots.getRange(2, 1, sample.length, 5).setValues(sample);
  }
  slots.getRange('B:B').setNumberFormat('@'); // время храним текстом «18:00»
  slots.getRange('A2:A').setNumberFormat('dd.mm.yyyy ddd');
  slots.getRange('D2:D').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList([FREE, BUSY], true).setAllowInvalid(false).build()
  );
  styleHeader_(slots, 5);
  slots.setColumnWidths(1, 1, 150);
  slots.setColumnWidths(2, 3, 120);
  slots.setColumnWidth(5, 220);
  slots.getRange('A1').setNote('Дата урока. Можно вводить как 14.10.2026');
  slots.getRange('B1').setNote('Время начала в формате 18:00');
  slots.getRange('D1').setNote('свободно — окно видно на сайте\nзанято — скрыто (ставится само при записи)\nЧтобы освободить окно — выберите «свободно»');

  // занятые окна подсвечиваются винным целой строкой
  slots.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=$D2="${BUSY}"`).setBackground('#F7EAE9').setFontColor('#7A2434')
      .setRanges([slots.getRange('A2:E')]).build(),
  ]);

  let bookings = ss.getSheetByName(BOOKINGS_SHEET);
  if (!bookings) bookings = ss.insertSheet(BOOKINGS_SHEET);
  if (bookings.getLastRow() === 0) {
    bookings.appendRow(['Когда пришла', 'Дата урока', 'Время', 'Имя', 'Telegram', 'Комментарий']);
  }
  bookings.getRange('A2:A').setNumberFormat('dd.mm.yyyy hh:mm');
  styleHeader_(bookings, 6);
  bookings.setColumnWidth(1, 150);
  bookings.setColumnWidth(6, 320);

  onOpen();
  toast_('Готово! Таблица настроена 🌿');
}

// ---------- Вспомогательное ----------

function getSheet_(name) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sheet) throw new Error(`Нет листа «${name}». Запустите setup().`);
  return sheet;
}

function styleHeader_(sheet, cols) {
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, cols).setBackground('#56612E').setFontColor('#FAF6EE').setFontWeight('bold');
}

function shiftDate_(iso, days) {
  const d = Utilities.parseDate(iso, 'UTC', 'yyyy-MM-dd');
  return Utilities.formatDate(new Date(d.getTime() + days * 864e5), 'UTC', 'yyyy-MM-dd');
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
