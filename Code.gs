/**
 * מתכנן ביקורים — צד השרת (Google Apps Script)
 * הקובץ הזה רץ אצל גוגל. הוא קורא וכותב לגיליון, ממקם כתובות על המפה
 * ומסדר את סדר הנסיעה. אין צורך לשנות בו כלום.
 */

var SHEETS = {
  biz: 'עסקים',
  visits: 'ביקורים',
  routes: 'קווים',
  settings: 'הגדרות'
};

var BIZ_COLS = ['id', 'שם העסק', 'סוג', 'סטטוס', 'אזור', 'איש קשר', 'טלפון', 'כתובת', 'הערה', 'מה לוקח', 'lat', 'lng', 'מיקום משוער', 'נוצר',
  'אימייל', 'שעות פתיחה', 'אתר / אינסטגרם', 'שם חברה / ח.פ'];
var VISIT_COLS = ['id', 'תאריך', 'businessId', 'שם העסק', 'תוצאה', 'כפתורים', 'מה נאמר', 'פעולת המשך', 'תאריך המשך', 'התעניינות', 'המשך פתוח', 'נוצר'];
var ROUTE_COLS = ['תאריך', 'עצירות', 'נקודת התחלה'];
var RESULT_HE = { prep: 'הכנה לפגישה', note: 'הערה', done: 'ביקרתי', owner_absent: 'הבעלים לא היה', closed_door: 'היה סגור', phone: 'שיחת טלפון', cancelled: 'בוטל' };
function resultCode_(he) { for (var k in RESULT_HE) if (RESULT_HE[k] === he) return k; return he; }

/* ---------- פתיחת האתר ---------- */

/**
 * המסכים נטענים מ-GitHub, כך ש-Claude יכול לעדכן אותם בלי שתצטרך להדביק קוד.
 * אם GitHub לא זמין, נטען העותק המקומי (הקובץ index כאן בעורך).
 */
var SCREENS_URL = 'https://raw.githubusercontent.com/maestracciariel/visit-planner/main/index.html';

function loadScreens_() {
  try {
    var res = UrlFetchApp.fetch(SCREENS_URL + '?t=' + Date.now(), { muteHttpExceptions: true });
    var html = res.getContentText('UTF-8');
    if (res.getResponseCode() === 200 && html.indexOf('google.script.run') > 0) return HtmlService.createHtmlOutput(html);
  } catch (e) {}
  return HtmlService.createHtmlOutputFromFile('index');
}

function doGet() {
  return loadScreens_()
    .setTitle('מתכנן ביקורים')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('מתכנן ביקורים')
    .addItem('1. הכנת הגיליונות', 'setup')
    .addItem('2. מיקום כתובות על המפה', 'geocodeAll')
    .addItem('3. בדיקה ותיקון של כל המיקומים', 'recheckLocations')
    .addSeparator()
    .addItem('4. הפעלת עדכון אוטומטי (פעם אחת)', 'setupAutoUpdate')
    .addItem('5. עדכון עכשיו', 'updateNow')
    .addToUi();
}

/* ---------- הקמה ---------- */

function setup() {
  var ss = SpreadsheetApp.getActive();
  ensureSheet_(ss, SHEETS.biz, BIZ_COLS);
  ensureSheet_(ss, SHEETS.visits, VISIT_COLS);
  ensureSheet_(ss, SHEETS.routes, ROUTE_COLS);
  var st = ensureSheet_(ss, SHEETS.settings, ['הגדרה', 'ערך']);
  if (st.getLastRow() < 2) {
    st.appendRow(['כתובת יציאה', '']);
    st.appendRow(['ימים לצהוב', 30]);
    st.appendRow(['ימים לכתום', 60]);
  }
  try { SpreadsheetApp.getUi().alert('הגיליונות מוכנים.'); } catch (e) {}
}

function ensureSheet_(ss, name, cols) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(cols);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, cols.length).setFontWeight('bold');
    sh.setRightToLeft(true);
  }
  return sh;
}

/* ---------- קריאה וכתיבה ---------- */

function readTable_(name) {
  var sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  var values = sh.getDataRange().getValues();
  var head = values.shift();
  return values.filter(function (r) { return r[0] !== ''; }).map(function (r) {
    var o = {};
    head.forEach(function (h, i) {
      var v = r[i];
      if (v instanceof Date) v = Utilities.formatDate(v, 'Asia/Jerusalem', 'yyyy-MM-dd');
      o[h] = v;
    });
    return o;
  });
}

function upsert_(name, cols, keyCol, obj) {
  var sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (sh.getLastColumn() > 0) {
    var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].filter(String);
    if (head.length >= cols.length) cols = head;
  }
  var row = cols.map(function (c) { return obj[c] === undefined || obj[c] === null ? '' : obj[c]; });
  var keyIdx = cols.indexOf(keyCol);
  var last = sh.getLastRow();
  if (last >= 2) {
    var keys = sh.getRange(2, keyIdx + 1, last - 1, 1).getDisplayValues();
    for (var i = 0; i < keys.length; i++) {
      if (String(keys[i][0]) === String(obj[keyCol])) {
        sh.getRange(i + 2, 1, 1, cols.length).setValues([row]);
        return;
      }
    }
  }
  sh.appendRow(row);
}

function newId_(prefix) {
  return prefix + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
}

/** מוסיף לגיליון עמודות חדשות שנוספו לקוד, כדי שלא תצטרך לעשות את זה ידנית */
function ensureColumns_(name, cols) {
  var sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) return;
  var head = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
  cols.forEach(function (c) {
    if (head.indexOf(c) < 0) { sh.getRange(1, head.length + 1).setValue(c).setFontWeight('bold'); head.push(c); }
  });
}

/** כל הנתונים שהאתר צריך, בקריאה אחת */
function getData() {
  ensureColumns_(SHEETS.biz, BIZ_COLS);
  var settings = {};
  readTable_(SHEETS.settings).forEach(function (r) { settings[r['הגדרה']] = r['ערך']; });
  return {
    businesses: readTable_(SHEETS.biz).map(function (r) {
      return {
        id: r.id, name: r['שם העסק'], kind: r['סוג'] === 'עסק חדש' ? 'lead' : 'client',
        status: r['סטטוס'] || 'פעיל', region: r['אזור'], contact: r['איש קשר'], phone: String(r['טלפון'] || ''),
        address: r['כתובת'], note: r['הערה'], takes: r['מה לוקח'],
        email: r['אימייל'] || '', hours: r['שעות פתיחה'] || '', web: r['אתר / אינסטגרם'] || '', company: r['שם חברה / ח.פ'] || '',
        lat: r.lat === '' ? null : Number(r.lat), lng: r.lng === '' ? null : Number(r.lng),
        approx: r['מיקום משוער'] === true || r['מיקום משוער'] === 'כן'
      };
    }),
    visits: readTable_(SHEETS.visits).map(function (r) {
      return {
        id: r.id, date: r['תאריך'], businessId: r.businessId, result: resultCode_(r['תוצאה']),
        chips: r['כפתורים'] ? String(r['כפתורים']).split(' | ') : [],
        text: r['מה נאמר'], followup: r['פעולת המשך'], followupDate: r['תאריך המשך'],
        interest: r['התעניינות'], followupOpen: r['המשך פתוח'] === 'כן'
      };
    }),
    routes: readTable_(SHEETS.routes).map(function (r) {
      return { date: r['תאריך'], stops: r['עצירות'] ? String(r['עצירות']).split(',') : [], start: r['נקודת התחלה'] };
    }),
    settings: settings
  };
}

/** שמירת סיכום ביקור. אפשר גם לעדכן סטטוס של העסק באותה פעולה */
function saveVisit(v) {
  ensureColumns_(SHEETS.visits, VISIT_COLS);
  var biz = readTable_(SHEETS.biz).filter(function (b) { return b.id === v.businessId; })[0];
  var row = {
    id: v.id || newId_('v'), 'תאריך': v.date, businessId: v.businessId, 'שם העסק': biz ? biz['שם העסק'] : '',
    'תוצאה': RESULT_HE[v.result] || v.result, 'כפתורים': (v.chips || []).join(' | '), 'מה נאמר': v.text || '',
    'פעולת המשך': v.followup || '', 'תאריך המשך': v.followupDate || '', 'התעניינות': v.interest || '',
    'המשך פתוח': v.followup ? 'כן' : 'לא', 'נוצר': new Date()
  };
  upsert_(SHEETS.visits, VISIT_COLS, 'id', row);
  if (v.newStatus && biz) {
    biz['סטטוס'] = v.newStatus;
    if (v.newStatus === 'פעיל' && biz['סוג'] === 'עסק חדש') biz['סוג'] = 'לקוח קיים';
    upsert_(SHEETS.biz, BIZ_COLS, 'id', biz);
  }
  return getData();
}

/** סימון פעולת המשך כבוצעה */
function closeFollowup(visitId) {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.visits);
  var ids = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getDisplayValues();
  for (var i = 0; i < ids.length; i++) {
    if (ids[i][0] === visitId) sh.getRange(i + 2, VISIT_COLS.indexOf('המשך פתוח') + 1).setValue('לא');
  }
  return getData();
}

/** הוספה או עדכון של עסק. כתובת חדשה ממוקמת על המפה מיד */
function saveBusiness(b) {
  var existing = b.id ? readTable_(SHEETS.biz).filter(function (x) { return x.id === b.id; })[0] : null;
  var row = existing || { id: newId_('b'), 'נוצר': new Date() };
  row['שם העסק'] = b.name;
  row['סוג'] = b.kind === 'lead' ? 'עסק חדש' : 'לקוח קיים';
  row['סטטוס'] = b.status || 'פעיל';
  row['אזור'] = b.region || '';
  row['איש קשר'] = b.contact || '';
  row['טלפון'] = b.phone || '';
  row['הערה'] = b.note || '';
  row['מה לוקח'] = b.takes || '';
  row['אימייל'] = b.email || '';
  row['שעות פתיחה'] = b.hours || '';
  row['אתר / אינסטגרם'] = b.web || '';
  row['שם חברה / ח.פ'] = b.company || '';
  var addressChanged = !existing || existing['כתובת'] !== b.address;
  row['כתובת'] = b.address || '';
  if (b.lat && b.lng) { row.lat = b.lat; row.lng = b.lng; row['מיקום משוער'] = 'לא'; }
  else if (addressChanged && b.address) {
    var g = locate_(b.address);
    if (g) { row.lat = g.lat; row.lng = g.lng; row['מיקום משוער'] = g.exact ? 'לא' : 'כן'; }
  }
  upsert_(SHEETS.biz, BIZ_COLS, 'id', row);
  return getData();
}

function saveRoute(date, stops, start) {
  upsert_(SHEETS.routes, ROUTE_COLS, 'תאריך', { 'תאריך': date, 'עצירות': stops.join(','), 'נקודת התחלה': start || '' });
  return getData();
}

/* ---------- מפות של גוגל ---------- */

function geocode_(address) {
  try {
    var res = Maps.newGeocoder().setRegion('il').setLanguage('iw').geocode(address);
    if (res.status === 'OK' && res.results.length) {
      var r0 = res.results[0], loc = r0.geometry.location, places = [];
      var SKIP = ['route', 'street_number', 'country', 'administrative_area_level_1', 'postal_code'];
      (r0.address_components || []).forEach(function (c) {
        if (!c.types.some(function (t) { return SKIP.indexOf(t) >= 0; })) places.push(c.long_name, c.short_name);
      });
      return { lat: loc.lat, lng: loc.lng, places: places };
    }
  } catch (e) {}
  return null;
}

function norm_(s) { return String(s || '').replace(/["'״׳`]/g, '').replace(/[()\-,.\/]/g, ' ').replace(/\s+/g, ' ').trim(); }
/** השוואה סלחנית לכתיב: בלי ו' ו-י' (שהם/שוהם, קיסריה/קיסרייה) */
function loose_(s) { return ' ' + norm_(s).replace(/[וי]/g, '') + ' '; }

/**
 * גוגל מצא מקום אמין רק אם שם היישוב שהוא מצא כתוב גם בכתובת שלנו, ובתוך ישראל.
 * לא סומכים על שם הרחוב: "הדקל 28" קיים בהרבה ערים.
 */
function trusted_(address, g) {
  if (!g || g.lat < 29.4 || g.lat > 33.4 || g.lng < 34.2 || g.lng > 35.9) return false;
  var a = loose_(address);
  return (g.places || []).some(function (name) {
    return name.split(/\s*-\s*/).concat([name]).some(function (part) {
      var p = loose_(part);
      return p.trim().length > 1 && a.indexOf(p) >= 0;
    });
  });
}

/**
 * מיקום לכתובת: קודם הכתובת המלאה. אם גוגל מצא יישוב אחר,
 * מנסים רק את שם היישוב מסוף הכתובת (מיקום משוער, אבל באזור הנכון).
 */
function locate_(address) {
  var g = geocode_(address);
  if (trusted_(address, g)) return { lat: g.lat, lng: g.lng, exact: true };
  var words = norm_(address).split(' ');
  for (var n = Math.min(3, words.length); n >= 1; n--) {
    var q = words.slice(-n).join(' ');
    var t = geocode_(q + ', ישראל');
    if (trusted_(q, t)) return { lat: t.lat, lng: t.lng, exact: false };
  }
  return null;
}

/** ממקם מחדש את כל הכתובות שסומנו "משוער" או שאין להן מיקום */
function geocodeAll() {
  return relocate_(false);
}

/** בודק מחדש את כל המיקומים, ומתקן כל עסק שגוגל שם ביישוב אחר מזה שבכתובת */
function recheckLocations() {
  return relocate_(true);
}

function relocate_(all) {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.biz);
  var values = sh.getDataRange().getValues();
  var c = function (n) { return BIZ_COLS.indexOf(n); };
  var exact = 0, approx = 0, moved = [], failed = [];
  for (var i = 1; i < values.length; i++) {
    var r = values[i];
    var wasApprox = r[c('מיקום משוער')] === 'כן' || r[c('lat')] === '';
    var addr = String(r[c('כתובת')] || '').trim();
    if ((!all && !wasApprox) || !addr || addr === 'כל הארץ') continue;
    var g = locate_(addr);
    if (!g) {
      failed.push(r[c('שם העסק')]);
      if (r[c('lat')] !== '') sh.getRange(i + 1, c('מיקום משוער') + 1).setValue('כן');
      continue;
    }
    var oldLat = Number(r[c('lat')]), oldLng = Number(r[c('lng')]);
    var far = r[c('lat')] !== '' && (Math.abs(oldLat - g.lat) > 0.05 || Math.abs(oldLng - g.lng) > 0.05);
    if (far) moved.push(r[c('שם העסק')]);
    sh.getRange(i + 1, c('lat') + 1, 1, 3).setValues([[g.lat, g.lng, g.exact ? 'לא' : 'כן']]);
    if (g.exact) exact++; else approx++;
    Utilities.sleep(150);
  }
  var msg = 'מיקום מדויק: ' + exact + '\nמיקום ברמת יישוב: ' + approx;
  if (moved.length) msg += '\n\nתוקנו (היו במקום שגוי):\n' + moved.join('\n');
  if (failed.length) msg += '\n\nלא נמצא מיקום — כדאי להוסיף עיר לכתובת:\n' + failed.join('\n');
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
  return msg;
}

/**
 * סידור העצירות בסדר הנסיעה הקצר ביותר, לפי גוגל.
 * מתחיל ומסתיים בכתובת היציאה (אם הוגדרה), אחרת מתחיל בעצירה הראשונה.
 */
function optimizeRoute(stopIds, startAddress) {
  var all = readTable_(SHEETS.biz);
  var byId = {};
  all.forEach(function (b) { byId[b.id] = b; });
  var stops = stopIds.map(function (id) { return byId[id]; }).filter(Boolean);
  if (stops.length < 3 && !startAddress) return { order: stopIds };
  var point = function (b) { return b.lat !== '' ? b.lat + ',' + b.lng : b['כתובת']; };
  var origin, destination, middle;
  if (startAddress) { origin = startAddress; destination = startAddress; middle = stops; }
  else { origin = point(stops[0]); destination = point(stops[stops.length - 1]); middle = stops.slice(1, -1); }
  var finder = Maps.newDirectionFinder().setOrigin(origin).setDestination(destination)
    .setMode(Maps.DirectionFinder.Mode.DRIVING).setOptimizeWaypoints(true).setLanguage('iw');
  middle.forEach(function (b) { finder.addWaypoint(point(b)); });
  var res = finder.getDirections();
  if (!res.routes || !res.routes.length) return { order: stopIds };
  var order = res.routes[0].waypoint_order.map(function (i) { return middle[i].id; });
  var legs = res.routes[0].legs;
  var km = 0, min = 0;
  legs.forEach(function (l) { km += l.distance.value / 1000; min += l.duration.value / 60; });
  if (!startAddress) order = [stops[0].id].concat(order, [stops[stops.length - 1].id]);
  return { order: order, km: Math.round(km), minutes: Math.round(min) };
}

/* ---------- עדכון אוטומטי ----------
 * Claude מעלה גרסאות חדשות ל-GitHub. פעם בשעה הסקריפט בודק אם יש גרסה חדשה,
 * מחליף לבד את הקוד שלו, יוצר גרסה ומפרסם אותה. לא צריך להדביק יותר כלום.
 * דורש פעם אחת: להפעיל את Google Apps Script API בכתובת script.google.com/home/usersettings
 */
var CODE_VERSION = 4;
var REPO_RAW = 'https://raw.githubusercontent.com/maestracciariel/visit-planner/main/';

function setupAutoUpdate() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'autoUpdate_') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('autoUpdate_').timeBased().everyHours(1).create();
  var msg = updateNow_(true, true);
  try { SpreadsheetApp.getUi().alert('העדכון האוטומטי פועל.\nמעכשיו האתר מתעדכן לבד פעם בשעה.\n\n' + msg); } catch (e) {}
  return msg;
}

function updateNow() {
  var msg = updateNow_(true);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
  return msg;
}

function autoUpdate_() { updateNow_(false); }

function fetchText_(name) {
  var res = UrlFetchApp.fetch(REPO_RAW + name + '?t=' + Date.now(), { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error('לא הצלחתי להוריד את ' + name + ' מ-GitHub');
  return res.getContentText('UTF-8');
}

function scriptApi_(method, path, body) {
  var res = UrlFetchApp.fetch('https://script.googleapis.com/v1/projects/' + ScriptApp.getScriptId() + path, {
    method: method, contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    payload: body ? JSON.stringify(body) : undefined
  });
  var code = res.getResponseCode(), text = res.getContentText();
  if (code >= 300) {
    if (text.indexOf('has not been used') >= 0 || text.indexOf('disabled') >= 0 || code === 403) {
      throw new Error('צריך להפעיל פעם אחת את Google Apps Script API: script.google.com/home/usersettings');
    }
    throw new Error('שגיאה מגוגל (' + code + '): ' + text.slice(0, 200));
  }
  return text ? JSON.parse(text) : {};
}

/** מחזיר הודעה קצרה בעברית על מה שקרה */
function updateNow_(manual, force) {
  var props = PropertiesService.getScriptProperties();
  try {
    var code = fetchText_('Code.gs');
    var m = code.match(/var CODE_VERSION = (\d+);/);
    var remote = m ? Number(m[1]) : 0;
    if (!remote || (!force && remote <= CODE_VERSION)) {
      props.setProperty('lastCheck', new Date().toISOString());
      return 'אין עדכון חדש. הגרסה שלך: ' + CODE_VERSION + '.';
    }
    var html = fetchText_('index.html');
    var manifest = fetchText_('appsscript.json');
    scriptApi_('put', '/content', { files: [
      { name: 'appsscript', type: 'JSON', source: manifest },
      { name: 'Code', type: 'SERVER_JS', source: code },
      { name: 'index', type: 'HTML', source: html }
    ] });
    var ver = scriptApi_('post', '/versions', { description: 'עדכון אוטומטי לגרסה ' + remote });
    var deps = scriptApi_('get', '/deployments').deployments || [];
    var moved = 0;
    deps.forEach(function (d) {
      var cfg = d.deploymentConfig || {};
      if (!cfg.versionNumber) return; // הפריסה של מצב העריכה (HEAD), לא נוגעים
      scriptApi_('put', '/deployments/' + d.deploymentId, { deploymentConfig: {
        scriptId: ScriptApp.getScriptId(), versionNumber: ver.versionNumber,
        manifestFileName: 'appsscript', description: 'גרסה ' + remote + ' (אוטומטי)'
      } });
      moved++;
    });
    props.setProperty('lastUpdate', new Date().toISOString() + ' → ' + remote);
    return 'עודכן לגרסה ' + remote + '. פורסמו ' + moved + ' פריסות.';
  } catch (e) {
    props.setProperty('lastError', new Date().toISOString() + ' ' + e.message);
    if (manual) return 'העדכון לא הצליח: ' + e.message;
    return '';
  }
}

