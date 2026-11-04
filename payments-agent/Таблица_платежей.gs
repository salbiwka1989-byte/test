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

  const f = [];
  for (let i = 2; i <= n; i++) {
    f.push([
      `=IF(E${i}="","",E${i}-TODAY())`,
      `=IF(OR(G${i}="Оплачено",G${i}="Отказ"),"",IF(E${i}="","Нет даты",` +
      `IF(K${i}<0,"Просрочено "&-K${i}&" дн.",IF(K${i}=0,"Оплата сегодня",IF(K${i}<=2,"Через "&K${i}&" дн.","")))))`,
    ]);
  }
  sh.getRange(2, 11, rows.length, 2).setFormulas(f);
  sh.getRange(2, 5, rows.length, 1).setNumberFormat('dd.mm.yyyy');
  sh.getRange(2, 8, rows.length, 1).setNumberFormat('dd.mm.yyyy');
  sh.getRange(2, 6, rows.length, 1).setNumberFormat('#,##0');
  sh.getRange(2, 11, rows.length, 1).setNumberFormat('0');

  sh.getRange(2, 7, rows.length, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).setAllowInvalid(false).build());

  const all = sh.getRange(2, 1, rows.length, 12);
  const rule = (formula, color, font) => {
    const b = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(formula).setBackground(color).setRanges([all]);
    if (font) b.setFontColor(font);
    return b.build();
  };
  sh.setConditionalFormatRules([
    rule('=$G2="Оплачено"', '#d9ead3'),
    rule('=$G2="Отказ"', '#eeeeee', '#888888'),
    rule('=AND($G2<>"Оплачено",$G2<>"Отказ",$E2<>"",$E2<TODAY())', '#f4cccc'),
    rule('=AND($G2<>"Оплачено",$G2<>"Отказ",$E2<>"",$E2>=TODAY(),$E2-TODAY()<=2)', '#fff2cc'),
  ]);

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
  s.getRange(1, 1, data.length, 6).setValues(data);
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
