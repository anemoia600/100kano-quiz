/**
 * 100그녀 이름 대기 — 랭킹 서버 (Google Apps Script)
 *
 * 쓰는 법
 * 1. 새 Google 스프레드시트를 만든다.
 * 2. 확장 프로그램 > Apps Script 를 열고, 기본 코드를 지운 뒤 이 파일 내용을 통째로 붙여 넣는다.
 * 3. 배포 > 새 배포 > 유형: 웹 앱, 실행 계정: 나, 액세스 권한: 모든 사용자 로 배포한다.
 * 4. 나온 웹 앱 URL(.../exec)을 사이트의 config.js 에 넣는다.
 *
 * 기록은 스프레드시트의 "scores" 시트에 한 줄씩 쌓인다. 이상한 기록은 그 줄을 지우면 랭킹에서도 사라진다.
 */
var SHEET = 'scores';
var TOTAL = 37;          // 그녀 수. 사이트의 명단이 늘면 같이 올린다.
var MAX_ROWS = 20000;    // 이보다 많이 쌓이면 새 기록을 받지 않는다.

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.appendRow(['at', 'name', 'mode', 'count', 'ms', 'done']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// 닉네임: 앞뒤 공백 제거, 12자 제한, 수식으로 해석될 수 있는 첫 글자 제거
function cleanName_(v) {
  var s = String(v == null ? '' : v).replace(/[\u0000-\u001f]/g, '').trim();
  s = s.replace(/^[=+\-@'\s]+/, '');
  return s.slice(0, 12);
}

function doGet() {
  var sh = sheet_();
  var last = sh.getLastRow();
  var best = {};
  if (last > 1) {
    var values = sh.getRange(2, 1, last - 1, 6).getValues();
    for (var i = 0; i < values.length; i++) {
      var v = values[i];
      var name = cleanName_(v[1]);
      var mode = String(v[2]);
      var count = Number(v[3]);
      var ms = Number(v[4]);
      if (!name || (mode !== 'order' && mode !== 'free')) continue;
      if (!(count >= 0 && count <= TOTAL) || !(ms >= 0)) continue;
      var key = mode + '\n' + name;
      var b = best[key];
      if (!b || count > b.count || (count === b.count && ms < b.ms)) {
        best[key] = { name: name, mode: mode, count: count, ms: ms, done: count === TOTAL,
                      at: v[0] instanceof Date ? v[0].toISOString() : String(v[0]) };
      }
    }
  }
  var rows = [];
  for (var k in best) rows.push(best[k]);
  rows.sort(function (a, b) { return a.count !== b.count ? b.count - a.count : a.ms - b.ms; });
  var out = [], perMode = { order: 0, free: 0 };
  for (var j = 0; j < rows.length; j++) {
    if (perMode[rows[j].mode] < 100) { out.push(rows[j]); perMode[rows[j].mode]++; }
  }
  return json_({ ok: true, rows: out });
}

function doPost(e) {
  var d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_json' }); }
  var name = cleanName_(d && d.name);
  var mode = d && d.mode;
  var count = Number(d && d.count);
  var ms = Math.round(Number(d && d.ms));
  if (!name) return json_({ ok: false, error: 'name' });
  if (mode !== 'order' && mode !== 'free') return json_({ ok: false, error: 'mode' });
  if (!(count >= 0 && count <= TOTAL) || Math.floor(count) !== count) return json_({ ok: false, error: 'count' });
  // 풀네임 하나를 0.4초보다 빨리 칠 수는 없다. 말이 안 되는 기록은 받지 않는다.
  if (!(ms >= count * 400) || ms > 6 * 60 * 60 * 1000) return json_({ ok: false, error: 'time' });

  var lock = LockService.getScriptLock();
  try { lock.waitLock(8000); } catch (err) { return json_({ ok: false, error: 'busy' }); }
  try {
    var sh = sheet_();
    if (sh.getLastRow() > MAX_ROWS) return json_({ ok: false, error: 'full' });
    sh.appendRow([new Date(), name, mode, count, ms, count === TOTAL]);
  } finally {
    lock.releaseLock();
  }
  return json_({ ok: true });
}
