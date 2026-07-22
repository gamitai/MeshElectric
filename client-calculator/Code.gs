/**
 * מחשבון רווחיות ללקוח – Google Apps Script Web App
 * -------------------------------------------------------------
 * הרעיון: כל הנוסחאות נשארות בגיליון (סודיות לגמרי).
 * הקוד כותב את הקלט של הלקוח לתאי B2:B9, נותן לגיליון לחשב,
 * קורא את התוצאות, ואז משחזר את ערכי ברירת המחדל.
 * הלקוח מקבל רק מספרים – אף פעם לא את הנוסחאות.
 */

// ---- הגדרות ----------------------------------------------------
// שם הטאב בגיליון. אם השם שונה – אין צורך לשנות: יש נפילה חכמה לטאב הראשון.
var SHEET_NAME = 'מחשבון';

// טווחי התאים (לפי המבנה בגיליון). שנה כאן בלבד אם המבנה זז.
var INPUT_RANGE     = 'B2:B9';   // ערכי הקלט שהלקוח ממלא
var INPUT_LABELS    = 'A2:A9';   // שמות שדות הקלט
var INPUT_UNITS     = 'C2:C9';   // יחידות (שח / kw / שנים ...)
var INPUT_HINT_CELL  = 'A10';     // הערת עזר למטר-לקילוואט
var BLOCK_EXTRA      = 'A34:C35'; // שדות (שטח פנלים, גודל מתקן אגירה)
var BLOCK_INVESTOR   = 'A25:C30'; // פיננסי למשקיע (ללא שורת ה-IRR שמוצגת בכרטיס נפרד)
var BLOCK_PROJECT    = 'A16:C22'; // פיננסי לפרויקט
var PRIVATE_IRR_CELL = 'B31';     // Private Leveraged IRR (CoC) – "IRR עם הלוואה", מוצג רק לשותף
var SPITZER_RANGE    = 'E2:Q52';  // כותרות (שורה 2) + נתוני השפיצר (טווח רחב; מוצגות רק שנים תקינות לפי B5)
var PARTNER_CELL     = 'B8';      // אחוז שותף – אם > 0 הלקוח נכנס כשותף

// עמודות השפיצר (אינדקס יחסי לתוך E..Q: E=0,F=1,...,Q=12)
var SPITZER_PROJECT_COLS = [0, 1, 2, 3, 4, 5, 6, 7];      // E–L (פרויקטלי)
var SPITZER_PRIVATE_COLS = [0, 1, 8, 9, 10, 11, 12];      // E,F,M–Q (פרטי)

// ---- נקודת כניסה ל-Web App ------------------------------------
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('מחשבון רווחיות')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  return sh || ss.getSheets()[0]; // נפילה חכמה אם שם הטאב שונה
}

// ---- בניית הטופס: מחזיר את שדות הקלט ----------------------------
function getInputs() {
  var sh = getSheet_();
  var labels  = sh.getRange(INPUT_LABELS).getValues();
  var values  = sh.getRange(INPUT_RANGE).getValues();
  var units   = sh.getRange(INPUT_UNITS).getValues();
  var formats = sh.getRange(INPUT_RANGE).getNumberFormats();

  var inputs = [];
  for (var i = 0; i < labels.length; i++) {
    var isPct = String(formats[i][0]).indexOf('%') > -1;
    inputs.push({
      row: 2 + i,
      label: labels[i][0],
      value: isPct ? Math.round(values[i][0] * 10000) / 100 : values[i][0],
      unit: isPct ? '%' : units[i][0],
      isPercent: isPct
    });
  }
  return { inputs: inputs, hint9: sh.getRange(INPUT_HINT_CELL).getValue() };
}

// ---- החישוב הראשי ---------------------------------------------
function calculate(inputMap) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000); // מונע התנגשות בין לקוחות במקביל

  var sh = getSheet_();
  var range = sh.getRange(INPUT_RANGE);
  var original = range.getValues();          // שמירת ברירת המחדל
  var formats = range.getNumberFormats();

  try {
    // כתיבת הקלט של הלקוח (אחוזים מומרים חזרה לשבר)
    var newVals = original.map(function (o, i) {
      var v = inputMap[String(2 + i)];
      if (v === '' || v === null || v === undefined || isNaN(Number(v))) {
        return [o[0]];
      }
      v = Number(v);
      if (String(formats[i][0]).indexOf('%') > -1) v = v / 100;
      return [v];
    });
    range.setValues(newVals);
    SpreadsheetApp.flush(); // מכריח חישוב מחדש של הנוסחאות

    var partnerPct = sh.getRange(PARTNER_CELL).getValue();
    var isPartner = !!(partnerPct && Number(partnerPct) > 0);

    // שדות + פיננסי + שפיצר פרויקטלי מוצגים לכולם.
    // ה-IRR (CoC) והשפיצר הפרטי – רק לשותף.
    var spitzer = readSpitzer_(sh);
    if (!isPartner) spitzer.private = null;  // לקוח רגיל רואה רק פרויקטלי

    return {
      isPartner: isPartner,
      fields:    readBlock_(sh, BLOCK_EXTRA),
      investor:  readBlock_(sh, BLOCK_INVESTOR),
      project:   readBlock_(sh, BLOCK_PROJECT),
      irr:       isPartner ? sh.getRange(PRIVATE_IRR_CELL).getDisplayValue() : null,
      spitzer:   spitzer
    };
  } finally {
    range.setValues(original);  // שחזור ברירת המחדל תמיד
    SpreadsheetApp.flush();
    lock.releaseLock();
  }
}

// ---- קריאת בלוק תוויות/ערכים (כמו שמוצג בגיליון) ---------------
function readBlock_(sh, a1) {
  var disp = sh.getRange(a1).getDisplayValues();
  var out = [];
  for (var i = 0; i < disp.length; i++) {
    var r = disp[i];
    if (r[0] !== '' || r[1] !== '') {
      out.push({ label: r[0], value: r[1], unit: r[2] });
    }
  }
  return out;
}

// ---- קריאת טבלאות השפיצר (פרויקטלי + פרטי) --------------------
function readSpitzer_(sh) {
  var disp = sh.getRange(SPITZER_RANGE).getDisplayValues();
  var vals = sh.getRange(SPITZER_RANGE).getValues();
  var headers = disp[0];

  var rows = [];
  for (var i = 1; i < disp.length; i++) {
    if (typeof vals[i][0] === 'number') rows.push(disp[i]); // שורות שנה תקינות בלבד
  }

  function pick(cols) {
    return {
      headers: cols.map(function (c) { return headers[c]; }),
      rows: rows.map(function (r) { return cols.map(function (c) { return r[c]; }); })
    };
  }

  return {
    project: pick(SPITZER_PROJECT_COLS),
    private: pick(SPITZER_PRIVATE_COLS)
  };
}
