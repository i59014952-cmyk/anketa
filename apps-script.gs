// Приём анкет в Google-таблицу.
// Вставьте этот код в Apps Script вашей таблицы (Расширения → Apps Script),
// затем разверните как веб-приложение с доступом «Все» и вставьте выданный URL
// в index.html в строку const ENDPOINT = "...".

const COLUMNS = [
  ['submittedAt', 'Заполнено'],
  ['date',        'Дата лекции'],
  ['topic',       'Тема лекции'],
  ['lecturer',    'Лектор'],
  ['org',         'Организация / подразделение'],
  ['useful',      'Польза для работы (1-5)'],
  ['clarity',     'Понятность изложения (1-5)'],
  ['pace',        'Темп подачи'],
  ['comment',     'Что было полезно, чего не хватило'],
  ['questions',   'Вопросы, которые не успели задать'],
  ['next',        'Темы для следующих лекций'],
  ['name',        'ФИО'],
  ['email',       'E-mail'],
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000); // чтобы одновременные анкеты не затирали друг друга

  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];

    if (sheet.getLastRow() === 0) {
      const header = COLUMNS.map(function (c) { return c[1]; });
      sheet.appendRow(header);
      sheet.getRange(1, 1, 1, header.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    sheet.appendRow(COLUMNS.map(function (c) { return data[c[0]] || ''; }));

    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return json({ ok: true, service: 'lecture-feedback' });
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
