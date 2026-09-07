// Приём анкет «Анкета рабочих задач» в Google-таблицу.
//
// Куда вставлять: таблица → Extensions → Apps Script → заменить весь Code.gs этим файлом.
// Затем Deploy → Manage deployments → карандаш → Version: New version → Deploy.
// Без обновления версии по ссылке продолжит работать прежний код.
//
// Пишется ОДИН лист «Анкеты»: одна строка — одна анкета целиком, включая все
// задачи человека списком в последней ячейке. Отдельного листа задач нет
// намеренно: заказчик читает таблицу глазами, и разложенные по двум вкладкам
// ответы выглядели как потерянные.

const COLUMNS = [
  ['submittedAt', 'Заполнено'],
  ['name',        'ФИО'],
  ['position',    'Должность'],
  ['org',         'Организация'],
  ['dept',        'Подразделение'],
  ['exp',         'Стаж в должности'],
  ['email',       'Почта'],
  ['contact',     'Телефон / Telegram'],
  ['duties',      'Зона ответственности'],
  ['flow',        'Кто ставит задачи / кому результат'],
  ['apps',        'Программы (отмеченные)'],
  ['apps_other',  'Программы (прочие)'],
  ['apps_pain',   'Ручной перенос данных'],
  ['heavy',       'Наиболее трудоёмкие задачи'],
  ['manual',      'Выполняется вручную'],
  ['delegate',    'Передал бы помощнику'],
  ['comment',     'Дополнительно'],
];

// Три колонки в конце строки: сколько задач, суммарная нагрузка и сами задачи.
const TAIL = ['Задач указано', 'Часов в месяц (оценка)', 'Задачи (списком)'];

// Оценка нагрузки: частота × длительность, обе по нижней границе диапазона.
// Коэффициенты живут только здесь — на странице анкеты их нет.
const PER_MONTH = {
  day_many: 42,    // ~2 раза в день × 21 рабочий день
  day: 21,
  week_few: 11,    // ~2,5 раза в неделю
  week: 4.3,
  month_few: 2.5,
  month: 1,
  rare: 0.3,
};
const HOURS = {
  m15: 0.2,        // до 15 минут — считаем как 12
  m30: 0.4,
  m60: 0.75,
  h3: 2,
  h8: 5,
  d1: 9,
};

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000); // чтобы одновременные анкеты не затирали друг друга

  try {
    const d = JSON.parse(e.postData.contents);
    const tasks = d.tasks || [];
    const sh = sheet(COLUMNS.map(function (c) { return c[1]; }).concat(TAIL));

    sh.appendRow(
      COLUMNS.map(function (c) { return d[c[0]] || ''; })
             .concat([tasks.length, total(tasks), summary(tasks)]));

    return json({ ok: true, tasks: tasks.length });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Нагрузка одной задачи в часах за месяц.
function hours(t) {
  const load = (PER_MONTH[t.freq] || 0) * (HOURS[t.dur] || 0);
  return load ? Math.round(load * 10) / 10 : 0;
}

// Сумма по всем задачам — число, а не текст: по этой колонке сортируют.
function total(tasks) {
  let sum = 0;
  tasks.forEach(function (t) { sum += hours(t); });
  return sum ? Math.round(sum * 10) / 10 : '';
}

// Все задачи одной ячейкой, по строке на задачу.
function summary(tasks) {
  return tasks.map(function (t, i) {
    const h = hours(t);
    return (i + 1) + '. ' + (t.what || 'без названия') +
      ' | периодичность: ' + (t.freqLabel || 'не указана') +
      ' | длительность: ' + (t.durLabel || 'не указана') +
      (t.apps ? ' | программы: ' + t.apps : '') +
      (t.out ? ' | результат: ' + t.out : '') +
      (h ? ' | около ' + h + ' ч в месяц' : '') +
      (t.pain ? ' | трудоёмкая часть: ' + t.pain : '');
  }).join('\n');
}

// Лист «Анкеты»: создаётся при первой анкете вместе с шапкой.
function sheet(header) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Анкеты');
  if (!sh) sh = ss.insertSheet('Анкеты');
  if (sh.getLastRow() === 0) {
    sh.appendRow(header);
    sh.getRange(1, 1, 1, header.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

// Самопроверка. Выбрать функцию test в списке наверху и нажать Run:
// в лист ляжет строка «Проверка связи». Заодно видно, не поехала ли
// кириллица при вставке кода — в шапке должны быть читаемые слова.
function test() {
  const demo = {
    submittedAt: new Date().toLocaleString('ru-RU'),
    name: 'Проверка связи', position: 'Тест', org: 'Тест', dept: 'Тест',
    exp: '1–3 года', email: 'test@test.ru', contact: '@test',
    duties: 'Тестовая строка, её можно удалить.',
    flow: '', apps: 'Excel, 1С', apps_other: '', apps_pain: '',
    heavy: 'Тестовая строка, её можно удалить.', manual: '', delegate: '', comment: '',
    tasks: [{
      what: 'Тестовая задача', freq: 'day', freqLabel: 'Каждый день',
      dur: 'm60', durLabel: '30–60 минут', apps: '1С', out: 'Документ', pain: '',
    }],
  };
  Logger.log(doPost({ postData: { contents: JSON.stringify(demo) } }).getContent());
}

function doGet() {
  return json({ ok: true, service: 'employee-tasks' });
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
