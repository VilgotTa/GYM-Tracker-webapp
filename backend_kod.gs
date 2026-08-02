/**
 * Träningslogg – JSON-API för fristående frontend (GitHub Pages).
 *
 * SETUP:
 * 1. Skapa ett NYTT Google Sheet för träning, kopiera dess ID hit nedan.
 *    (Kan också vara en ny flik i ett befintligt Sheet – ändra bara SHEET_ID
 *     till det Sheetets ID.)
 * 2. Tillägg -> Apps Script, klistra in denna kod.
 * 3. Byt TOKEN till en egen hemlig sträng (samma som i index.html).
 * 4. Deploy -> New deployment -> Web app
 *      Execute as: Me
 *      Who has access: Anyone      <-- viktigt!
 * 5. Kopiera /exec-URL:en till API_URL i index.html.
 *
 * Sheet-kolumner som skapas automatiskt:
 * Datum | Tid | Övning | Set | Reps | Vikt | RIR | Kommentar | SessionID
 */

var SHEET_ID = 'DITT_TRANINGS_SHEET_ID_HÄR';
var SHEET_NAME = 'Traning';
var TOKEN = 'BYT_TILL_EN_EGEN_HEMLIG_KOD';

var DEFAULT_EXERCISES = [
  'Knäböj', 'Bänkpress', 'Marklyft', 'Militärpress', 'Chins',
  'Skivstångsrodd', 'Dips', 'Benpress', 'Bicepscurl', 'Triceps pushdown'
];

function doGet(e) {
  if (e.parameter.token !== TOKEN) return jsonOut_({ error: 'unauthorized' });
  var action = e.parameter.action;
  if (action === 'getExercises') return jsonOut_(getExercises());
  if (action === 'getLastSessions') return jsonOut_(getLastSessions());
  return jsonOut_({ error: 'unknown action' });
}

function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); }
  catch (err) { return jsonOut_({ error: 'bad request' }); }
  if (body.token !== TOKEN) return jsonOut_({ error: 'unauthorized' });

  var action = body.action;
  if (action === 'saveWorkouts') return jsonOut_(saveWorkouts(body.entries));
  if (action === 'addExercise') return jsonOut_(addExercise(body.name));
  if (action === 'removeExercise') return jsonOut_(removeExercise(body.name));
  return jsonOut_({ error: 'unknown action' });
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Datum', 'Tid', 'Övning', 'Set', 'Reps', 'Vikt', 'RIR', 'Kommentar', 'SessionID']);
  }
  return sheet;
}

// Sparar en kö av övningar (varje övning = flera set = flera rader)
function saveWorkouts(entries) {
  var sheet = getSheet_();
  var rows = [];

  for (var i = 0; i < entries.length; i++) {
    var en = entries[i];
    var d = new Date(en.sessionId);
    var datum = Utilities.formatDate(d, 'Europe/Stockholm', 'yyyy-MM-dd');
    var tid = Utilities.formatDate(d, 'Europe/Stockholm', 'HH:mm');

    for (var j = 0; j < en.sets.length; j++) {
      var s = en.sets[j];
      rows.push([
        datum,
        tid,
        en.exercise,
        j + 1,
        s.reps,
        s.weight,
        (s.rir === '' || s.rir === null || s.rir === undefined) ? '' : s.rir,
        j === 0 ? (en.comment || '') : '',
        en.sessionId,
      ]);
    }
  }

  if (rows.length > 0) {
    var startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
  }
  return { ok: true, rows: rows.length };
}

// Bygger { övning: { date, sets:[{weight,reps,rir}], comment } } för senaste passet per övning
function getLastSessions() {
  var sheet = getSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};

  var data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
  var best = {}; // övning -> sessionId
  var out = {};

  for (var i = 0; i < data.length; i++) {
    var r = data[i];
    var exercise = r[2];
    var sessionId = r[8];
    if (!exercise || !sessionId) continue;

    if (!best[exercise] || sessionId > best[exercise]) {
      best[exercise] = sessionId;
      out[exercise] = { date: formatCell_(r[0]), sets: [], comment: '' };
    }
    if (sessionId === best[exercise]) {
      out[exercise].sets.push({ weight: r[5], reps: r[4], rir: r[6] });
      if (r[7]) out[exercise].comment = r[7];
    }
  }
  return out;
}

function formatCell_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Europe/Stockholm', 'yyyy-MM-dd');
  return String(v);
}

// ---- Övningslista ----
function getExercises() {
  var props = PropertiesService.getUserProperties();
  var str = props.getProperty('EXERCISES');
  if (!str) {
    props.setProperty('EXERCISES', JSON.stringify(DEFAULT_EXERCISES));
    return DEFAULT_EXERCISES;
  }
  return JSON.parse(str);
}

function addExercise(name) {
  var props = PropertiesService.getUserProperties();
  var list = JSON.parse(props.getProperty('EXERCISES') || '[]');
  if (list.indexOf(name) === -1) {
    list.push(name);
    list.sort();
    props.setProperty('EXERCISES', JSON.stringify(list));
  }
  return list;
}

function removeExercise(name) {
  var props = PropertiesService.getUserProperties();
  var list = JSON.parse(props.getProperty('EXERCISES') || '[]');
  var idx = list.indexOf(name);
  if (idx !== -1) {
    list.splice(idx, 1);
    props.setProperty('EXERCISES', JSON.stringify(list));
  }
  return list;
}
