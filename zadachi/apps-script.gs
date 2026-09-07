// Приём анкет «Задачи и программы сотрудника» в Google-таблицу.
//
// Куда вставлять: таблица → Расширения → Apps Script → заменить весь Код.gs этим файлом.
// Затем «Начать развёртывание → Новое развёртывание», тип «Веб-приложение»,
// запуск от имени «Я», доступ «Все», и полученный URL вписать в index.html
// в строку const ENDPOINT = "...".
//
// Пишется ДВА листа:
//   «Сотрудники» — одна строка на анкету (кто, обязанности, программы, боли);
//   «Задачи»     — одна строка на задачу, с оценкой часов в месяц.
// Задачи отдельным листом, потому что их у человека много: так их можно
// отсортировать по часам и сравнить между людьми, не разбирая ячейку с текстом.

const PEOPLE = [
  ['submittedAt', 'Заполнено'],
  ['name',        'ФИО'],
  ['position',    'Должность'],
  ['org',         'Организация'],
  ['dept',        'Подразделение'],
  ['exp',         'Стаж в должности'],
  ['email',       'Почта'],
  ['contact',     'Телефон / Telegram'],
  ['duties',      'Обязанности'],
  ['flow',        'Кто ставит задачи / кому результат'],
  ['apps',        'Программы (отмеченные)'],
  ['apps_other',  'Программы (прочие)'],
  ['apps_pain',   'Ручной перенос данных'],
  ['heavy',       'Отнимает больше всего времени'],
  ['manual',      'Делается руками, хотя повторяется'],
  ['delegate',    'Отдал бы помощнику'],
  ['comment',     'Дополнительно'],
];

const TASKS = [
  'Заполнено', 'ФИО', 'Должность', 'Организация', 'Подразделение',
  '№', 'Задача', 'Как часто', 'Сколько занимает раз',
  'Программы', 'Результат', 'Что муторно', '≈ часов в месяц',
];

// Оценка нагрузки. Коэффициенты живут ТОЛЬКО здесь — на странице анкеты их нет,
// иначе две оценки одной задачи разъехались бы. Это прикидка, а не замер:
// частота × длительность, обе взяты по нижней границе диапазона.
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
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // --- лист «Сотрудники» ---
    const people = sheet(ss, 'Сотрудники', PEOPLE.map(function (c) { return c[1]; }));
    people.appendRow(PEOPLE.map(function (c) { return d[c[0]] || ''; }));

    // --- лист «Задачи» ---
    const tasks = d.tasks || [];
    if (tasks.length) {
      const sh = sheet(ss, 'Задачи', TASKS);
      const rows = tasks.map(function (t, i) {
        const load = (PER_MONTH[t.freq] || 0) * (HOURS[t.dur] || 0);
        return [
          d.submittedAt || '', d.name || '', d.position || '', d.org || '', d.dept || '',
          i + 1, t.what || '', t.freqLabel || '', t.durLabel || '',
          t.apps || '', t.out || '', t.pain || '',
          load ? Math.round(load * 10) / 10 : '',
        ];
      });
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, TASKS.length).setValues(rows);
    }

    return json({ ok: true, tasks: tasks.length });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Лист по имени: создаётся при первой анкете вместе с шапкой.
function sheet(ss, name, header) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(header);
    sh.getRange(1, 1, 1, header.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function doGet() {
  return json({ ok: true, service: 'employee-tasks' });
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
