/**
 * Создаёт новую таблицу «Платежи AI-профессионал» из графика на листе «Лист1».
 *
 * Одна строка новой таблицы = один платёж.
 * Платёж с зелёной заливкой в графике переносится со статусом «Оплачено»,
 * остальные получают статус «Ожидается».
 *
 * Запуск: Расширения → Apps Script → вставить этот код → выбрать функцию
 * createPaymentsTable → «Выполнить» → разрешить доступ.
 */

const SOURCE_SHEET = 'Лист1';
const NEW_TITLE = 'Платежи AI-профессионал';
const GREEN = '#00ff00';
const RED = '#ff0000';
const MOP_NAMES = {
  'Венера Х': 'Венера Хасуева',
  'Крис Менеджер Effect Medici': 'Крис',
  'A. Джанаралиева': 'А. Джанаралиева', // в исходнике первая буква латинская
};
const DEFAULT_PAYMENTS_ID = '13zidU5kRMDcOeUWiH8PzSYQ_EA0gEOQCv415Xiq8EEg';
const STATUSES = ['Ожидается', 'Оплачено', 'Перенос', 'Отказ'];

function createPaymentsTable() {
  const src = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SOURCE_SHEET);
  const lastRow = src.getLastRow();
  const range = src.getRange(1, 1, lastRow, 38);
  const values = range.getValues();
  const display = range.getDisplayValues();
  const bg = range.getBackgrounds();

  const rows = [];
  for (let r = 1; r < values.length; r++) {
    const name = String(values[r][1]).trim();
    if (!name) continue;
    const redRow = bg[r][1].toLowerCase() === RED;
    let tg = String(values[r][2]).trim();
    if (tg === 'нету') tg = '';
    const mopRaw = String(values[r][3]).trim();
    const mop = MOP_NAMES[mopRaw] || mopRaw;
    const link = String(values[r][0]).trim();

    const pays = [];
    for (let k = 0; k < 15; k++) {
      const dc = 8 + 2 * k, sc = 9 + 2 * k;
      const dv = values[r][dc], sv = values[r][sc];
      if (dv === '' && sv === '') continue;
      const green = bg[r][dc].toLowerCase() === GREEN;
      const parsed = parseDate_(dv, display[r][dc]);
      const notes = parsed.note ? [parsed.note] : [];
      let amount = sv;
      if (typeof sv === 'string' && sv !== '') { notes.push('Сумма в исходнике: «' + sv + '»'); amount = ''; }
      else if (sv === '' || sv === 0) { notes.push('Нет суммы'); amount = ''; }
      pays.push({ n: k + 1, date: parsed.date, amount: amount, status: green ? 'Оплачено' : 'Ожидается', notes: notes });
    }

    const seen = {};
    pays.forEach(p => { if (p.date) seen[p.date.getTime()] = (seen[p.date.getTime()] || 0) + 1; });
    pays.forEach(p => {
      if (p.date && seen[p.date.getTime()] > 1) p.notes.push('Дата повторяется');
      if (redRow) p.notes.push('Строка была красной целиком');
      rows.push([name, tg, mop, p.n, p.date || '', p.amount, p.status, '', p.notes.join('; '), link]);
    });
  }

  const ss = SpreadsheetApp.create(NEW_TITLE);
  const sh = ss.getSheets()[0];
  sh.setName('Платежи');
  buildPaymentsSheet_(sh, rows);
  const mops = Array.from(new Set(rows.map(x => x[2]))).sort();
  buildSummarySheet_(ss, mops);
  buildManagersSheet_(ss, mops);
  buildGuideSheet_(ss);

  PropertiesService.getScriptProperties().setProperty('PAYMENTS_SPREADSHEET_ID', ss.getId());
  const msg = 'Готово: ' + rows.length + ' платежей. Новая таблица: ' + ss.getUrl();
  Logger.log(msg);
  SpreadsheetApp.getActiveSpreadsheet().toast(msg, 'Платежи', 30);
}

/** Год в графике не указан: сентябрь–декабрь считаем 2026, январь–август 2027. */
function parseDate_(v, shown) {
  let d, m, note = '';
  if (v instanceof Date) {
    d = v.getDate(); m = v.getMonth() + 1;
  } else {
    const s = String(shown || v).trim();
    if (!s) return { date: null, note: 'Нет даты' };
    const mm = s.match(/(\d{1,2})(-\d{1,2})?\.(\d{1,2})/);
    if (!mm) return { date: null, note: 'В исходнике дата: «' + s + '»' };
    d = +mm[1]; m = +mm[3];
    if (mm[2]) note = 'В исходнике дата: «' + s + '»';
  }
  const y = m >= 9 ? 2026 : 2027;
  const date = new Date(y, m - 1, d);
  if (date.getMonth() !== m - 1) return { date: null, note: 'Несуществующая дата: ' + d + '.' + m };
  return { date: date, note: note };
}

function buildPaymentsSheet_(sh, rows) {
  const header = ['Клиентка', 'Telegram', 'МОП', '№ платежа', 'Дата платежа', 'Сумма, ₽', 'Статус',
    'Дата оплаты', 'Что проверить', 'Сделка в GetCourse', 'Дней до оплаты', 'Сигнал'];
  const n = rows.length + 1;
  sh.getRange(1, 1, 1, header.length).setValues([header])
    .setFontWeight('bold').setFontColor('#ffffff').setBackground('#37474f').setWrap(true);
  sh.getRange(2, 1, rows.length, 10).setValues(rows);

  writePaymentFormulas_(sh, rows.length);
  sh.getRange(2, 5, rows.length, 1).setNumberFormat('dd.mm.yyyy');
  sh.getRange(2, 8, rows.length, 1).setNumberFormat('dd.mm.yyyy');
  sh.getRange(2, 6, rows.length, 1).setNumberFormat('#,##0');
  sh.getRange(2, 11, rows.length, 1).setNumberFormat('0');

  sh.getRange(2, 7, rows.length, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).setAllowInvalid(false).build());

  setColorRules_(sh, rows.length);

  sh.setFrozenRows(1);
  sh.setFrozenColumns(1);
  sh.getRange(1, 1, n, 12).createFilter();
  [200, 150, 140, 80, 105, 90, 105, 105, 260, 220, 80, 140].forEach((w, i) => sh.setColumnWidth(i + 1, w));
}

function buildSummarySheet_(ss, mops) {
  const s = ss.insertSheet('Сводка по МОП');
  const P = "'Платежи'!";
  const open = `${P}$G:$G,"<>Оплачено",${P}$G:$G,"<>Отказ"`;
  const data = [['МОП', 'Просрочено, платежей', 'Просрочено, ₽', 'Ближайшие 7 дней, платежей', 'Ближайшие 7 дней, ₽', 'Оплачено, ₽']];
  mops.forEach((m, idx) => {
    const i = idx + 2;
    data.push([m,
      `=COUNTIFS(${P}$C:$C,$A${i},${open},${P}$K:$K,"<0")`,
      `=SUMIFS(${P}$F:$F,${P}$C:$C,$A${i},${open},${P}$K:$K,"<0")`,
      `=COUNTIFS(${P}$C:$C,$A${i},${open},${P}$K:$K,">=0",${P}$K:$K,"<=7")`,
      `=SUMIFS(${P}$F:$F,${P}$C:$C,$A${i},${open},${P}$K:$K,">=0",${P}$K:$K,"<=7")`,
      `=SUMIFS(${P}$F:$F,${P}$C:$C,$A${i},${P}$G:$G,"Оплачено")`]);
  });
  const t = mops.length + 1;
  data.push(['Итого', `=SUM(B2:B${t})`, `=SUM(C2:C${t})`, `=SUM(D2:D${t})`, `=SUM(E2:E${t})`, `=SUM(F2:F${t})`]);
  const sep = formulaSeparator_(ss);
  s.getRange(1, 1, data.length, 6).setValues(data.map((row, i) => i === 0 ? row : row.map(v => fx_(v, sep))));
  s.getRange(1, 1, 1, 6).setFontWeight('bold').setFontColor('#ffffff').setBackground('#37474f').setWrap(true);
  s.getRange(data.length, 1, 1, 6).setFontWeight('bold');
  s.getRange(2, 2, data.length - 1, 5).setNumberFormat('#,##0');
  s.setColumnWidth(1, 170);
  for (let c = 2; c <= 6; c++) s.setColumnWidth(c, 130);
}

function buildManagersSheet_(ss, mops) {
  const m = ss.insertSheet('Менеджеры');
  const data = [['МОП', 'Telegram chat_id', 'Получает общую сводку (да/нет)']].concat(mops.map(x => [x, '', 'нет']));
  m.getRange(1, 1, data.length, 3).setValues(data);
  m.getRange(1, 1, 1, 3).setFontWeight('bold').setFontColor('#ffffff').setBackground('#37474f');
  m.setColumnWidths(1, 3, 200);
}

function buildGuideSheet_(ss) {
  const g = ss.insertSheet('Как вести');
  const lines = [
    'Как вести таблицу',
    '',
    'Одна строка = один платёж. После оплаты МОП ставит в столбце «Статус» значение «Оплачено» и дату оплаты.',
    'Статусы: Ожидается, Оплачено, Перенос (новую дату вписать в «Дата платежа»), Отказ (напоминания прекращаются).',
    'Цвет строки ставится сам: зелёный — оплачено, жёлтый — оплата в ближайшие 2 дня, красный — просрочено, серый — отказ.',
    'Столбцы «Дней до оплаты» и «Сигнал» считаются формулами, их не трогать.',
    'Столбец «Что проверить» заполнен там, где в старом графике была неточность: диапазон дат, «?», нет суммы, повтор даты, красная строка.',
    'Новый платёж: добавить строку внизу и протянуть формулы из столбцов K и L.',
    'Перенесено из «График платежей AI-профессионал», Лист1. Оплаченным считался платёж с зелёной заливкой.',
  ];
  g.getRange(1, 1, lines.length, 1).setValues(lines.map(x => [x]));
  g.getRange(1, 1).setFontWeight('bold').setFontSize(13);
  g.setColumnWidth(1, 900);
}

/* ---------- Формулы с учётом языка таблицы ---------- */

/** В русской таблице аргументы формул разделяются «;», в английской «,». Проверяем на пробной ячейке. */
function formulaSeparator_(ss) {
  const tmp = ss.insertSheet('_tmp');
  const c = tmp.getRange('A1');
  c.setFormula('=IF(1,1,0)');
  SpreadsheetApp.flush();
  const ok = c.getDisplayValue() === '1';
  ss.deleteSheet(tmp);
  return ok ? ',' : ';';
}

/** Формулы ниже записаны с запятыми и без запятых внутри текста, поэтому замена безопасна. */
function fx_(f, sep) {
  return typeof f === 'string' && f.charAt(0) === '=' ? f.replace(/,/g, sep) : f;
}

function writePaymentFormulas_(sh, count) {
  const sep = formulaSeparator_(sh.getParent());
  const f = [];
  for (let i = 2; i <= count + 1; i++) {
    f.push([
      fx_(`=IF(E${i}="","",E${i}-TODAY())`, sep),
      fx_(`=IF(OR(G${i}="Оплачено",G${i}="Отказ"),"",IF(E${i}="","Нет даты",` +
        `IF(K${i}<0,"Просрочено "&-K${i}&" дн.",IF(K${i}=0,"Оплата сегодня",IF(K${i}<=2,"Через "&K${i}&" дн.","")))))`, sep),
    ]);
  }
  sh.getRange(2, 11, count, 2).setFormulas(f);
}

function setColorRules_(sh, count) {
  const sep = formulaSeparator_(sh.getParent());
  const all = sh.getRange(2, 1, count, 12);
  const rule = (formula, color, font) => {
    const b = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(fx_(formula, sep))
      .setBackground(color).setRanges([all]);
    if (font) b.setFontColor(font);
    return b.build();
  };
  sh.setConditionalFormatRules([
    rule('=$G2="Оплачено"', '#d9ead3'),
    rule('=$G2="Отказ"', '#eeeeee', '#888888'),
    rule('=AND($G2<>"Оплачено",$G2<>"Отказ",$E2<>"",$E2<TODAY())', '#f4cccc'),
    rule('=AND($G2<>"Оплачено",$G2<>"Отказ",$E2<>"",$E2>=TODAY(),$E2-TODAY()<=2)', '#fff2cc'),
  ]);
}

function paymentsSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('PAYMENTS_SPREADSHEET_ID') || DEFAULT_PAYMENTS_ID;
  return SpreadsheetApp.openById(id);
}

/** Один раз: чинит формулы и цвета в уже созданной таблице «Платежи AI-профессионал». */
function fixFormulas() {
  const ss = paymentsSpreadsheet_();
  const sh = ss.getSheetByName('Платежи');
  const count = sh.getLastRow() - 1;
  writePaymentFormulas_(sh, count);
  setColorRules_(sh, count);
  const mops = Array.from(new Set(sh.getRange(2, 3, count, 1).getValues().map(r => r[0]).filter(String))).sort();
  const old = ss.getSheetByName('Сводка по МОП');
  const pos = old.getIndex();
  ss.deleteSheet(old);
  buildSummarySheet_(ss, mops);
  ss.setActiveSheet(ss.getSheetByName('Сводка по МОП'));
  ss.moveActiveSheet(pos);
  SpreadsheetApp.flush();
  Logger.log('Готово. Проверка: K2 = ' + sh.getRange('K2').getDisplayValue() + ', L2 = «' + sh.getRange('L2').getDisplayValue() + '»');
}

/* ---------- Напоминания в Telegram ---------- */
/*
 * Утром (9:00): каждой менеджер свой список, владелице «План на день».
 * Вечером (20:00): владелице «Факт за день»: что собрали и что из плана не оплачено.
 * Владелица = строка на листе «Менеджеры» с «да» в третьем столбце.
 *
 * Когда МОП ставит статус «Оплачено», в «Дата оплаты» сама подставляется сегодняшняя дата.
 * По этой дате вечерний отчёт понимает, что деньги пришли именно сегодня.
 *
 * Настройка:
 * 1. Шестерёнка «Настройки проекта» → «Свойства скрипта» → свойство TELEGRAM_TOKEN.
 * 2. Нажать «Старт» в боте и написать ему любое слово.
 * 3. Запустить showBotChats: в журнале будут имена и chat_id. Вписать их на лист «Менеджеры».
 * 4. Проверить: sendReminders (утро) и sendEveningReport (вечер).
 * 5. Запустить setupDailyTrigger один раз: включает утро, вечер и автодату оплаты.
 */

const REMIND_DAYS_BEFORE = 2;   // за сколько дней предупреждать менеджера
const MORNING_HOUR = 9;
const EVENING_HOUR = 20;

function tg_(method, payload) {
  const token = PropertiesService.getScriptProperties().getProperty('TELEGRAM_TOKEN');
  if (!token) throw new Error('Не задан TELEGRAM_TOKEN в свойствах скрипта.');
  const res = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/' + method, {
    method: 'post', contentType: 'application/json', payload: JSON.stringify(payload || {}), muteHttpExceptions: true,
  });
  const body = JSON.parse(res.getContentText());
  if (!body.ok) throw new Error('Telegram: ' + body.description);
  return body.result;
}

/** Показывает, кто писал боту, с их chat_id. */
function showBotChats() {
  const updates = tg_('getUpdates', {});
  const seen = {};
  updates.forEach(u => {
    const m = u.message || u.edited_message;
    if (!m) return;
    const c = m.chat;
    seen[c.id] = [c.first_name, c.last_name].filter(Boolean).join(' ') + (c.username ? ' (@' + c.username + ')' : '');
  });
  const lines = Object.keys(seen).map(id => seen[id] + ': ' + id);
  Logger.log(lines.length ? lines.join('\n') : 'Боту пока никто не писал. Попросите менеджеров нажать «Старт» и написать любое слово.');
}

function sendText_(chatId, text) {
  // Telegram принимает до 4096 символов, режем по строкам
  const parts = [];
  let cur = '';
  text.split('\n').forEach(line => {
    if ((cur + '\n' + line).length > 3900) { parts.push(cur); cur = line; } else { cur = cur ? cur + '\n' + line : line; }
  });
  if (cur) parts.push(cur);
  parts.forEach(p => tg_('sendMessage', { chat_id: chatId, text: p, disable_web_page_preview: true }));
}

function money_(n) {
  return n === '' || n === null ? 'сумма не указана' : Utilities.formatString('%s ₽', Math.round(n).toLocaleString('ru-RU'));
}

function tz_() { return Session.getScriptTimeZone(); }
function dayKey_(d) { return Utilities.formatDate(d, tz_(), 'yyyy-MM-dd'); }
function sum_(arr) { return arr.reduce((a, p) => a + (Number(p.amount) || 0), 0); }
function total_(arr) { return arr.length + ' на ' + money_(sum_(arr)); }

/** Все платежи с днями до оплаты (отрицательные = просрочка). */
function readAll_() {
  const ss = paymentsSpreadsheet_();
  const v = ss.getSheetByName('Платежи').getDataRange().getValues();
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const all = [];
  for (let i = 1; i < v.length; i++) {
    const [client, tgName, mop, num, date, amount, status, paidDate] = v[i];
    if (!client) continue;
    const p = { client, tgName, mop, num, amount, status, paidDate: paidDate instanceof Date ? paidDate : null, days: null };
    if (date instanceof Date) {
      p.date = new Date(date); p.date.setHours(0, 0, 0, 0);
      p.days = Math.round((p.date - today) / 86400000);
    }
    all.push(p);
  }
  const mgr = ss.getSheetByName('Менеджеры').getDataRange().getValues().slice(1)
    .filter(r => r[0] && r[1])
    .map(r => ({ name: String(r[0]).trim(), chat: String(r[1]).trim(), owner: String(r[2]).trim().toLowerCase() === 'да' }));
  return { ss, all, mgr };
}

function isOpen_(p) { return p.status !== 'Оплачено' && p.status !== 'Отказ'; }
function groupBy_(arr) {
  const g = {};
  arr.forEach(p => { (g[p.mop] = g[p.mop] || []).push(p); });
  return g;
}

function line_(p) {
  const who = p.client + (p.tgName ? ' ' + p.tgName : '');
  const when = p.date ? Utilities.formatDate(p.date, Session.getScriptTimeZone(), 'dd.MM') : 'без даты';
  return '• ' + who + ', платёж №' + p.num + ', ' + money_(p.amount) + ' (' + when + ')';
}

/** Утро: менеджерам их списки, владелице план на день. */
function sendReminders() {
  const { ss, all, mgr } = readAll_();
  const open = all.filter(isOpen_);
  const byMop = groupBy_(open);

  mgr.forEach(m => {
    const mine = byMop[m.name];
    if (!mine) return;
    const overdue = mine.filter(p => p.days !== null && p.days < 0).sort((a, b) => a.days - b.days);
    const today = mine.filter(p => p.days === 0);
    const soon = mine.filter(p => p.days > 0 && p.days <= REMIND_DAYS_BEFORE).sort((a, b) => a.days - b.days);
    if (!overdue.length && !today.length && !soon.length) return;

    const t = ['Доброе утро! Платежи на сегодня:'];
    if (today.length) t.push('', 'Оплата сегодня:', ...today.map(line_));
    if (soon.length) t.push('', 'Скоро, предупредить заранее:', ...soon.map(line_));
    if (overdue.length) t.push('', 'Просрочено (' + total_(overdue) + '):', ...overdue.map(p => line_(p) + ', ' + (-p.days) + ' дн.'));
    t.push('', 'После оплаты поставьте «Оплачено» в таблице: ' + ss.getUrl());
    sendText_(m.chat, t.join('\n'));
  });

  const owners = mgr.filter(m => m.owner);
  if (!owners.length) return;
  const todayAll = open.filter(p => p.days === 0);
  const overdueAll = open.filter(p => p.days !== null && p.days < 0);
  const s = ['План на ' + Utilities.formatDate(new Date(), tz_(), 'dd.MM.yyyy'), '',
    'Сегодня ждём: ' + total_(todayAll),
    'Всего в просрочке: ' + total_(overdueAll)];
  Object.keys(byMop).sort().forEach(name => {
    const td = byMop[name].filter(p => p.days === 0);
    const od = byMop[name].filter(p => p.days !== null && p.days < 0);
    const wk = byMop[name].filter(p => p.days !== null && p.days > 0 && p.days <= 7);
    s.push('', name + ':', 'сегодня ' + total_(td));
    td.forEach(p => s.push(line_(p)));
    s.push('в ближайшие 7 дней ' + total_(wk), 'просрочено ' + total_(od));
    const long = od.filter(p => p.days <= -7).sort((a, b) => a.days - b.days);
    if (long.length) s.push('дольше недели:', ...long.map(p => line_(p) + ', ' + (-p.days) + ' дн.'));
  });
  owners.forEach(m => sendText_(m.chat, s.join('\n')));
}

/** Вечер: владелице факт за день. */
function sendEveningReport() {
  const { all, mgr } = readAll_();
  const owners = mgr.filter(m => m.owner);
  if (!owners.length) return;
  const key = dayKey_(new Date());

  const collected = all.filter(p => p.status === 'Оплачено' && p.paidDate && dayKey_(p.paidDate) === key);
  const plan = all.filter(p => p.days === 0 && p.status !== 'Отказ');
  const planPaid = plan.filter(p => p.status === 'Оплачено');
  const planMissed = plan.filter(isOpen_);
  const overdueAll = all.filter(p => isOpen_(p) && p.days !== null && p.days <= 0);

  const s = ['Факт за ' + Utilities.formatDate(new Date(), tz_(), 'dd.MM.yyyy'), '',
    'Собрано сегодня: ' + total_(collected),
    'План на сегодня: ' + total_(plan) + ', оплачено ' + total_(planPaid),
    'Не оплачено к вечеру, включая старые долги: ' + total_(overdueAll)];

  const mops = Array.from(new Set(collected.concat(planMissed).map(p => p.mop))).sort();
  mops.forEach(name => {
    const c = collected.filter(p => p.mop === name);
    const miss = planMissed.filter(p => p.mop === name);
    s.push('', name + ':');
    if (c.length) s.push('собрано ' + total_(c), ...c.map(line_));
    if (miss.length) s.push('из плана не оплатили:', ...miss.map(line_));
  });
  if (!collected.length && !plan.length) s.push('', 'Сегодня оплат не было и не ожидалось.');
  owners.forEach(m => sendText_(m.chat, s.join('\n')));
}

/** Когда в «Статус» ставят «Оплачено», подставляет сегодняшнюю дату оплаты, если её нет. */
function onPaymentsEdit(e) {
  const r = e.range, sh = r.getSheet();
  if (sh.getName() !== 'Платежи') return;
  const firstCol = r.getColumn(), lastCol = firstCol + r.getNumColumns() - 1;
  if (firstCol > 7 || lastCol < 7) return;
  const rows = sh.getRange(r.getRow(), 7, r.getNumRows(), 2).getValues();
  let changed = false;
  rows.forEach(row => {
    if (row[0] === 'Оплачено' && !row[1]) { row[1] = new Date(); changed = true; }
  });
  if (changed) sh.getRange(r.getRow(), 8, r.getNumRows(), 1).setValues(rows.map(x => [x[1]]));
}

/** Включает утреннюю и вечернюю рассылку и автодату оплаты. Запускать один раз. */
function setupDailyTrigger() {
  const mine = ['sendReminders', 'sendEveningReport', 'onPaymentsEdit'];
  ScriptApp.getProjectTriggers().filter(t => mine.indexOf(t.getHandlerFunction()) >= 0).forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('sendReminders').timeBased().everyDays(1).atHour(MORNING_HOUR).create();
  ScriptApp.newTrigger('sendEveningReport').timeBased().everyDays(1).atHour(EVENING_HOUR).create();
  ScriptApp.newTrigger('onPaymentsEdit').forSpreadsheet(paymentsSpreadsheet_().getId()).onEdit().create();
  Logger.log('Готово: план около ' + MORNING_HOUR + ':00, факт около ' + EVENING_HOUR + ':00 (' + tz_() + ').');
}
