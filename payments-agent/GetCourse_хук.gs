/**
 * Приём оплат из GetCourse по вебхуку («отбивка по хуку»).
 *
 * GetCourse при каждой оплате по заказу вызывает ссылку этого скрипта и передаёт
 * номер заказа, почту и сколько всего оплачено по заказу. Скрипт:
 *  1. считает, сколько пришло сейчас (оплачено всего минус то, что уже записано по этому заказу);
 *  2. записывает поступление на лист «Поступления GK»;
 *  3. ставит «Оплачено» ближайшему неоплаченному платежу клиентки на листе «Платежи»;
 *  4. пишет об оплате её менеджеру в Telegram, а если клиентку найти не удалось, владелицам.
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

  // Сколько уже записано по этому заказу
  const sameOrder = r => {
    const id = String(r[IC['ID заказа']]).trim(), nm = String(r[IC['Номер заказа']]).trim();
    return (orderId && (id === orderId || nm === orderId)) || (orderNum && (id === orderNum || nm === orderNum));
  };
  const prev = inVals.slice(1).filter(sameOrder);
  const already = prev.reduce((a, r) => a + num_(r[IC['Сумма платежа']]), 0);
  const amount = Math.round((paidTotal - already) * 100) / 100;
  if (amount <= 0) return 'skip: новых денег нет (оплачено ' + paidTotal + ', уже записано ' + already + ')';

  // Кто клиентка
  const payVals = paySh.getDataRange().getValues();
  const PC = cols_(payVals[0], ['Клиентка', 'МОП', '№ платежа', 'Дата платежа', 'Сумма, ₽', 'Статус', 'Дата оплаты',
    'Заказ GK (номер)', 'Зачтено из GK, ₽']);
  let client = '';
  if (PC['Заказ GK (номер)'] >= 0) {
    const hit = payVals.slice(1).find(r => {
      const v = String(r[PC['Заказ GK (номер)']]).trim();
      return v && (v === orderNum || v === orderId);
    });
    if (hit) client = String(hit[PC['Клиентка']]).trim();
  }
  if (!client) {
    const named = prev.find(r => String(r[IC['Клиентка']]).trim());
    if (named) client = String(named[IC['Клиентка']]).trim();
  }
  if (!client && email && IC['Email'] >= 0) {
    const byMail = inVals.slice(1).find(r => norm_(r[IC['Email']]) === email && String(r[IC['Клиентка']]).trim());
    if (byMail) client = String(byMail[IC['Клиентка']]).trim();
  }

  // Записываем поступление
  const now = new Date();
  const row = new Array(inVals[0].length).fill('');
  const put = (name, v) => { if (IC[name] >= 0) row[IC[name]] = v; };
  put('Дата', Utilities.formatDate(now, HOOK_TZ, 'dd.MM.yyyy'));
  put('ID заказа', orderId || orderNum);
  put('Номер заказа', orderNum);
  put('Email', p.email || '');
  put('Клиентка', client || p.name || '');
  put('Сумма платежа', amount);
  put('Всего оплачено по заказу', paidTotal);
  put('Стоимость заказа', p.cost ? num_(p.cost) : '');
  put('Источник', 'GetCourse хук');
  put('Время записи', Utilities.formatDate(now, HOOK_TZ, 'dd.MM.yyyy H:mm:ss'));
  inSh.appendRow(row);

  if (!client) {
    notifyOwners_('Пришла оплата из GetCourse, но клиентку в таблице найти не удалось.\n' +
      'Заказ ' + orderNum + ', ' + (p.name || 'без имени') + ', ' + (p.email || 'без почты') + ', ' + money2_(amount) +
      '.\nОплата записана на лист «Поступления GK», статус в «Платежах» нужно поставить вручную.');
    return 'ok: записано, клиентка не найдена';
  }

  // Отмечаем платежи: по порядку дат, пока хватает суммы
  const open = [];
  for (let i = 1; i < payVals.length; i++) {
    const r = payVals[i];
    if (norm_(r[PC['Клиентка']]) !== norm_(client)) continue;
    const st = String(r[PC['Статус']]);
    if (st === 'Оплачено' || st === 'Отказ') continue;
    open.push({ i, r, date: r[PC['Дата платежа']] instanceof Date ? r[PC['Дата платежа']].getTime() : Infinity });
  }
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

  const mop = open.length ? String(open[0].r[PC['МОП']]).trim() : '';
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
