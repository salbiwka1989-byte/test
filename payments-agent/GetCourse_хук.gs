/**
 * Приём оплат из GetCourse по вебхуку («отбивка по хуку»).
 *
 * GetCourse при каждой оплате по заказу вызывает ссылку этого скрипта и передаёт
 * номер заказа, почту и сколько всего оплачено по заказу. Скрипт:
 *  1. считает, сколько пришло сейчас (оплачено всего минус то, что уже записано по этому заказу);
 *  2. записывает поступление на лист «Поступления GK»;
 *  3. ставит «Оплачено» ближайшему неоплаченному платежу клиентки на листе «Платежи»;
 *  4. пишет об оплате её менеджеру в Telegram.
 * Заказы, которых нет в графике (столбец «Заказ GK (номер)» на листе «Платежи»), пропускаются молча.
 * Повторный вызов с той же суммой ничего не задвоит: прирост будет 0.
 *
 * Это отдельный проект Apps Script, он не трогает остальные скрипты.
 * Настройка описана в setupHook() ниже.
 */

const HOOK_PAYMENTS_ID = '1ExjbIn-C7hIirY2_UYIulFEQ-AoWrqVodWZlcDSSb2I'; // таблица «Платежи AI-профессионал»
const HOOK_TZ = 'Europe/Moscow';

/** Шаг 1 настройки: создаёт секретный ключ и показывает, какие параметры передавать из GetCourse. */
function setupHook() {
  const props = PropertiesService.getScriptProperties();
  let secret = props.getProperty('GK_SECRET');
  if (!secret) {
    secret = Utilities.getUuid().replace(/-/g, '');
    props.setProperty('GK_SECRET', secret);
  }
  const url = ScriptApp.getService().getUrl();
  Logger.log([
    'Секретный ключ: ' + secret,
    url ? 'Ссылка веб-приложения: ' + url : 'Ссылки пока нет: сначала «Начать развертывание» → «Новое развертывание» → «Веб-приложение».',
    '',
    'Параметры для блока «Вызвать URL» в процессе GetCourse:',
    'key=' + secret,
    'order_id=<ID заказа>',
    'order_number=<номер заказа>',
    'email=<email пользователя>',
    'name=<имя пользователя>',
    'paid=<сколько всего оплачено по заказу, числом>',
    'cost=<стоимость заказа, числом>',
    '',
    'Проверка: в браузере откройте ссылку веб-приложения с ?key=' + secret + '&test=1',
  ].join('\n'));
}

function doGet(e) { return handleHook_(e); }
function doPost(e) { return handleHook_(e); }

function handleHook_(e) {
  const p = Object.assign({}, (e && e.parameter) || {});
  if (e && e.postData && e.postData.type === 'application/json') {
    try { Object.assign(p, JSON.parse(e.postData.contents)); } catch (err) { /* не JSON, берём параметры */ }
  }
  const secret = PropertiesService.getScriptProperties().getProperty('GK_SECRET');
  if (!secret || p.key !== secret) return text_('forbidden');
  if (p.test) return text_('ok: скрипт на связи');

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return text_(processPayment_(p));
  } catch (err) {
    notifyOwners_('Ошибка при приёме оплаты из GetCourse: ' + err.message + '\nДанные: ' + JSON.stringify(p));
    return text_('error: ' + err.message);
  } finally {
    lock.releaseLock();
  }
}

function text_(s) { return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.TEXT); }

function num_(v) {
  if (typeof v === 'number') return v;
  const n = parseFloat(String(v || '').replace(/\s| |руб\.?|₽/gi, '').replace(',', '.'));
  return isNaN(n) ? 0 : n;
}
function norm_(s) { return String(s || '').trim().toLowerCase().replace(/ё/g, 'е').replace(/\s+/g, ' '); }

/** Индексы столбцов по заголовкам, чтобы не зависеть от порядка. */
function cols_(header, names) {
  const idx = {};
  names.forEach(n => { idx[n] = header.findIndex(h => norm_(h) === norm_(n)); });
  return idx;
}

function processPayment_(p) {
  const ss = SpreadsheetApp.openById(HOOK_PAYMENTS_ID);
  const inSh = ss.getSheetByName('Поступления GK');
  const paySh = ss.getSheetByName('Платежи');
  const inVals = inSh.getDataRange().getValues();
  const IC = cols_(inVals[0], ['Дата', 'ID заказа', 'Номер заказа', 'Email', 'Клиентка', 'Сумма платежа',
    'Всего оплачено по заказу', 'Стоимость заказа', 'Источник', 'Время записи', 'ID платежа GK']);

  const orderId = String(p.order_id || '').trim();
  const orderNum = String(p.order_number || orderId).trim();
  const email = norm_(p.email);
  const paidTotal = num_(p.paid);
  if (!orderId && !orderNum) return 'skip: нет номера заказа';

  // Работаем только с заказами из графика рассрочек: номер заказа должен стоять в «Платежах»
  const payVals = paySh.getDataRange().getValues();
  const PC = cols_(payVals[0], ['Клиентка', 'МОП', '№ платежа', 'Дата платежа', 'Сумма, ₽', 'Статус', 'Дата оплаты',
    'Заказ GK (номер)', 'Зачтено из GK, ₽']);
  if (PC['Заказ GK (номер)'] < 0) return 'skip: в «Платежах» нет столбца «Заказ GK (номер)»';
  const planRows = [];
  for (let i = 1; i < payVals.length; i++) {
    const v = String(payVals[i][PC['Заказ GK (номер)']]).trim();
    if (v && (v === orderNum || v === orderId)) planRows.push({ i, r: payVals[i] });
  }
  if (!planRows.length) return 'skip: заказ не из графика рассрочек';
  const client = String(planRows[0].r[PC['Клиентка']]).trim();

  // Сколько уже учтено по этому заказу
  const sameOrder = r => {
    const id = String(r[IC['ID заказа']]).trim(), nm = String(r[IC['Номер заказа']]).trim();
    return (orderId && (id === orderId || nm === orderId)) || (orderNum && (id === orderNum || nm === orderNum));
  };
  const prev = inVals.slice(1).filter(sameOrder);
  const sumPrev = prev.reduce((a, r) => a + num_(r[IC['Сумма платежа']]), 0);
  const lastTotal = IC['Всего оплачено по заказу'] < 0 ? 0 :
    prev.reduce((a, r) => Math.max(a, num_(r[IC['Всего оплачено по заказу']])), 0);
  const already = Math.max(sumPrev, lastTotal);
  const amount = Math.round((paidTotal - already) * 100) / 100;
  if (amount <= 0) return 'skip: новых денег нет (оплачено ' + paidTotal + ', уже учтено ' + already + ')';

  // Записываем поступление
  const now = new Date();
  const row = new Array(inVals[0].length).fill('');
  const put = (name, v) => { if (IC[name] >= 0) row[IC[name]] = v; };
  put('Дата', Utilities.formatDate(now, HOOK_TZ, 'dd.MM.yyyy'));
  put('ID заказа', orderId || orderNum);
  put('Номер заказа', orderNum);
  put('Email', p.email || '');
  put('Клиентка', client);
  put('Сумма платежа', amount);
  put('Всего оплачено по заказу', paidTotal);
  put('Стоимость заказа', p.cost ? num_(p.cost) : '');
  put('Источник', 'GetCourse хук');
  put('Время записи', Utilities.formatDate(now, HOOK_TZ, 'dd.MM.yyyy H:mm:ss'));
  inSh.appendRow(row);

  // Отмечаем платежи: по порядку дат, пока хватает суммы
  const open = planRows.filter(o => {
    const st = String(o.r[PC['Статус']]);
    return st !== 'Оплачено' && st !== 'Отказ';
  }).map(o => Object.assign(o, { date: o.r[PC['Дата платежа']] instanceof Date ? o.r[PC['Дата платежа']].getTime() : Infinity }));
  open.sort((a, b) => a.date - b.date);
  let rest = amount;
  const marked = [];
  for (const o of open) {
    const need = num_(o.r[PC['Сумма, ₽']]);
    if (need && rest < need - 1) break; // на следующий платёж денег не хватает
    const rowNo = o.i + 1;
    paySh.getRange(rowNo, PC['Статус'] + 1).setValue('Оплачено');
    paySh.getRange(rowNo, PC['Дата оплаты'] + 1).setValue(now);
    if (PC['Заказ GK (номер)'] >= 0 && !String(o.r[PC['Заказ GK (номер)']]).trim()) paySh.getRange(rowNo, PC['Заказ GK (номер)'] + 1).setValue(orderNum);
    if (PC['Зачтено из GK, ₽'] >= 0) paySh.getRange(rowNo, PC['Зачтено из GK, ₽'] + 1).setValue(need || rest);
    marked.push(o);
    rest -= need || rest;
    if (rest <= 1) break;
  }

  const mop = String(planRows[0].r[PC['МОП']]).trim();
  let msg = client + ' оплатила ' + money2_(amount) + ' (заказ ' + orderNum + ').';
  if (marked.length) {
    msg += '\nОтмечено в таблице: ' + marked.map(o => 'платёж №' + o.r[PC['№ платежа']]).join(', ') + '.';
    if (rest > 1) msg += '\nОсталось сверх графика: ' + money2_(rest) + ', проверьте вручную.';
  } else if (open.length) {
    msg += '\nЭто меньше ближайшего платежа (' + money2_(num_(open[0].r[PC['Сумма, ₽']])) + '), статус не менялся: проверьте вручную.';
  } else {
    msg += '\nНеоплаченных платежей по графику у неё нет, проверьте вручную.';
  }
  notifyMop_(mop, msg);
  return 'ok: ' + msg;
}

function money2_(n) { return Math.round(n).toLocaleString('ru-RU') + ' ₽'; }

function managers_() {
  const v = SpreadsheetApp.openById(HOOK_PAYMENTS_ID).getSheetByName('Менеджеры').getDataRange().getValues().slice(1);
  return v.filter(r => r[0] && r[1]).map(r => ({ name: String(r[0]).trim(), chat: String(r[1]).trim(), owner: norm_(r[2]) === 'да' }));
}

function notifyMop_(mop, text) {
  const m = managers_();
  const targets = m.filter(x => x.name === mop);
  (targets.length ? targets : m.filter(x => x.owner)).forEach(x => tgSend_(x.chat, text));
}

function notifyOwners_(text) { managers_().filter(x => x.owner).forEach(x => tgSend_(x.chat, text)); }

function tgSend_(chat, text) {
  const token = PropertiesService.getScriptProperties().getProperty('TELEGRAM_TOKEN');
  if (!token) return;
  UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    payload: JSON.stringify({ chat_id: chat, text: text, disable_web_page_preview: true }),
  });
}

/** Проверка без GetCourse: имитирует оплату. Поменяйте данные на реальный заказ из «Поступлений GK». */
function testHook() {
  const secret = PropertiesService.getScriptProperties().getProperty('GK_SECRET');
  const res = handleHook_({ parameter: { key: secret, order_id: '22640', order_number: '22640', email: '', name: 'Тест', paid: '0' } });
  Logger.log(res.getContent());
}

/**
 * Один раз: убирает следы первого запуска 19.11.2026, когда GetCourse прислал сразу 93 заказа.
 * 1. Удаляет из «Поступлений GK» строки «GetCourse хук» по заказам, которых нет в графике.
 * 2. Возвращает статус «Ожидается» платежам, которые хук отметил по чужому заказу:
 *    Алена Гасанбекова №2–4 и Фатима (@Fatima_Bokova) №1.
 */
function cleanupFirstRun() {
  const ss = SpreadsheetApp.openById(HOOK_PAYMENTS_ID);
  const paySh = ss.getSheetByName('Платежи');
  const pv = paySh.getDataRange().getValues();
  const PC = cols_(pv[0], ['Клиентка', 'Telegram', '№ платежа', 'Статус', 'Дата оплаты', 'Заказ GK (номер)', 'Зачтено из GK, ₽']);
  const plan = new Set(pv.slice(1).map(r => String(r[PC['Заказ GK (номер)']]).trim()).filter(String));

  const inSh = ss.getSheetByName('Поступления GK');
  const iv = inSh.getDataRange().getValues();
  const IC = cols_(iv[0], ['ID заказа', 'Номер заказа', 'Источник']);
  let removed = 0;
  for (let i = iv.length - 1; i >= 1; i--) {
    if (String(iv[i][IC['Источник']]) !== 'GetCourse хук') continue;
    const id = String(iv[i][IC['ID заказа']]).trim(), nm = String(iv[i][IC['Номер заказа']]).trim();
    if (!plan.has(id) && !plan.has(nm)) { inSh.deleteRow(i + 1); removed++; }
  }

  const wrong = [
    r => norm_(r[PC['Клиентка']]) === 'алена гасанбекова' && [2, 3, 4].indexOf(Number(r[PC['№ платежа']])) >= 0,
    r => String(r[PC['Telegram']]).trim() === '@Fatima_Bokova' && Number(r[PC['№ платежа']]) === 1,
  ];
  const reverted = [];
  for (let i = 1; i < pv.length; i++) {
    const r = pv[i];
    const d = r[PC['Дата оплаты']];
    const today = d instanceof Date && Utilities.formatDate(d, HOOK_TZ, 'dd.MM.yyyy') === '19.11.2026';
    if (!today || r[PC['Статус']] !== 'Оплачено' || !wrong.some(f => f(r))) continue;
    paySh.getRange(i + 1, PC['Статус'] + 1).setValue('Ожидается');
    paySh.getRange(i + 1, PC['Дата оплаты'] + 1).clearContent();
    paySh.getRange(i + 1, PC['Зачтено из GK, ₽'] + 1).clearContent();
    reverted.push(r[PC['Клиентка']] + ' №' + r[PC['№ платежа']]);
  }
  Logger.log('Удалено строк из «Поступлений GK»: ' + removed + '. Возвращено в «Ожидается»: ' + (reverted.join(', ') || 'ничего'));
}
