// Приём анкет «Задачи и программы сотрудника» в Google-таблицу.
//
// Куда вставлять: таблица → Расширения → Apps Script → заменить весь Код.gs этим файлом.
// Затем «Начать развёртывание → Новое развёртывание», тип «Веб-приложение»,
// запуск от имени «Я», доступ «Все», и полученный URL вписать в index.html
// в строку const ENDPOINT = "...".
//
// Пишется ДВА листа:
//   «Сотрудники» — одна строка на анкету, ЦЕЛИКОМ: кто, обязанности, программы,
//                  затраты времени и все задачи списком в последней ячейке;
//   «Задачи»     — одна строка на задачу, с оценкой часов в месяц.
// Задачи продублированы намеренно: строкой сотрудника анкету читают глазами,
// а листом «Задачи» её сортируют по нагрузке. Источник у обоих один и тот же
// массив tasks, поэтому разойтись они не могут.

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

// Две колонки в конце строки сотрудника: сколько задач и все они текстом.
const PEOPLE_TAIL = ['Задач указано', 'Задачи (списком)'];

const TASKS = [
  'Заполнено', 'ФИО', 'Должность', 'Организация', 'Подразделение',
  '№', 'Задача', 'Как часто', 'Сколько занимает раз',
  'Программы', 'Результат', 'Наиболее трудоёмкая часть', '≈ часов в месяц',
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

    const tasks = d.tasks || [];

    // --- лист «Сотрудники» ---
    const people = sheet(ss, 'Сотрудники',
      PEOPLE.map(function (c) { return c[1]; }).concat(PEOPLE_TAIL));
    people.appendRow(
      PEOPLE.map(function (c) { return d[c[0]] || ''; })
            .concat([tasks.length, summary(tasks)]));

    // --- лист «Задачи» ---
    if (tasks.length) {
      const sh = sheet(ss, 'Задачи', TASKS);
      const rows = tasks.map(function (t, i) {
        const load = hours(t);
        return [
          d.submittedAt || '', d.name || '', d.position || '', d.org || '', d.dept || '',
          i + 1, t.what || '', t.freqLabel || '', t.durLabel || '',
          t.apps || '', t.out || '', t.pain || '',
          load,
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

// Оценка нагрузки одной задачи. Считается в одном месте: и для листа «Задачи»,
// и для сводки в строке сотрудника — иначе одна задача получила бы два числа.
function hours(t) {
  const load = (PER_MONTH[t.freq] || 0) * (HOURS[t.dur] || 0);
  return load ? Math.round(load * 10) / 10 : '';
}

// Все задачи человека одной ячейкой — чтобы анкету можно было прочитать
// целиком, не переходя на второй лист.
function summary(tasks) {
  return tasks.map(function (t, i) {
    const h = hours(t);
    return (i + 1) + '. ' + (t.what || 'без названия') +
      ' — ' + (t.freqLabel || 'периодичность не указана') +
      ', ' + (t.durLabel || 'длительность не указана') +
      (t.apps ? ', программы: ' + t.apps : '') +
      (t.out ? ', результат: ' + t.out : '') +
      (h ? ', ≈' + String(h).replace('.', ',') + ' ч/мес' : '') +
      (t.pain ? '. Трудоёмкая часть: ' + t.pain : '');
  }).join('\n');
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

// Самопроверка. Запускается прямо в редакторе: выбрать функцию test в списке
// наверху и нажать «Выполнить». В таблицу ляжет строка «Проверка связи» и одна
// тестовая задача. Если строки появились — код и права в порядке, остаётся
// только развернуть веб-приложение. Строки потом удалите руками.
function test() {
  const demo = {
    submittedAt: new Date().toLocaleString('ru-RU'),
    name: 'Проверка связи', position: 'Тест', org: 'Тест', dept: 'Тест',
    exp: '1–3 года', email: 'test@test.ru', contact: '@test',
    duties: 'Тестовая строка, её можно удалить.',
    flow: '', apps: 'Excel / Google Таблицы, 1С', apps_other: '', apps_pain: '',
    heavy: 'Тестовая строка, её можно удалить.', manual: '', delegate: '', comment: '',
    tasks: [{
      what: 'Тестовая задача', freq: 'day', freqLabel: 'Каждый день',
      dur: 'm60', durLabel: '30–60 минут', apps: '1С', out: 'Документ', pain: '',
    }],
  };
  const res = doPost({ postData: { contents: JSON.stringify(demo) } });
  Logger.log(res.getContent());
}

function doGet() {
  return json({ ok: true, service: 'employee-tasks' });
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
