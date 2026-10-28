/**
 * パスタマン ブレイカー みんなのランキング (Google Apps Script)
 *
 * 1人1行で自己ベストを保存する。プレイヤーはブラウザごとのランダムなキーで区別。
 * 設置手順は game/README.md を参照。
 */
const SHEET_NAME = 'ranking';
const TOP_N = 50;
const MAX_SCORE = 50000000;

// スプレッドシートから開いたスクリプトならそのシートを使う。
// script.google.com で単体作成したスクリプトなら、初回に「パスタマン ランキング」を自動で作って覚える。
function book_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHEET_ID');
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (e) { /* 消されていたら作り直す */ }
  }
  const ss = SpreadsheetApp.create('パスタマン ランキング');
  props.setProperty('SHEET_ID', ss.getId());
  return ss;
}

function sheet_() {
  const ss = book_();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['key', 'name', 'score', 'stage', 'combo', 'updated']);
  }
  return sh;
}

function top_(key) {
  const rows = sheet_().getDataRange().getValues().slice(1).filter(function (r) { return r[0]; });
  rows.sort(function (a, b) { return Number(b[2]) - Number(a[2]); });
  return rows.slice(0, TOP_N).map(function (r) {
    return { name: String(r[1]), score: Number(r[2]), stage: Number(r[3]), combo: Number(r[4]), me: !!key && r[0] === key };
  });
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function cleanName_(v) {
  let name = String(v || '').replace(/[\u0000-\u001f\u007f<>]/g, '').trim().slice(0, 12);
  if (!name) name = '名無しのパスタ';
  // スプレッドシートの数式として解釈させない
  if (/^[=+\-@]/.test(name)) name = "'" + name;
  return name;
}

function doGet(e) {
  const key = String((e && e.parameter && e.parameter.key) || '');
  return out_({ ok: true, rows: top_(key) });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const key = String(d.key || '');
    if (!/^[A-Za-z0-9-]{8,64}$/.test(key)) return out_({ ok: false, error: 'bad key' });
    const score = Math.floor(Number(d.score));
    if (!(score > 0 && score <= MAX_SCORE)) return out_({ ok: false, error: 'bad score' });
    const stage = Math.max(1, Math.min(999, Math.floor(Number(d.stage)) || 1));
    const combo = Math.max(0, Math.min(99999, Math.floor(Number(d.combo)) || 0));
    const name = cleanName_(d.name);

    const sh = sheet_();
    const vals = sh.getDataRange().getValues();
    let row = -1;
    for (let i = 1; i < vals.length; i++) if (vals[i][0] === key) { row = i + 1; break; }

    let improved = false, best = score;
    if (row < 0) {
      sh.appendRow([key, name, score, stage, combo, new Date()]);
      improved = true;
    } else {
      best = Number(vals[row - 1][2]) || 0;
      if (score > best) {
        sh.getRange(row, 2, 1, 5).setValues([[name, score, stage, combo, new Date()]]);
        improved = true;
        best = score;
      } else {
        sh.getRange(row, 2).setValue(name);
      }
    }
    return out_({ ok: true, improved: improved, best: best, rows: top_(key) });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
