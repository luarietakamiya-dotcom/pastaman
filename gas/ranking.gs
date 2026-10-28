/**
 * パスタマン ブレイカー みんなのランキング (Google Apps Script)
 *
 * 登録ごとに log シートへ1行追記し、読み出し時に期間ごと（今日・週間・月間・年間・総合）の
 * 自己ベストランキングを組み立てる。期間の区切りは日本時間（週は月曜はじまり）。
 * プレイヤーはブラウザごとのランダムなキーで区別。設置手順は game/README.md を参照。
 */
const LOG_SHEET = 'log';
const OLD_SHEET = 'ranking'; // 旧バージョンの自己ベスト表（あれば初回に log へ取り込む）
const TOP_N = 50;
const MAX_SCORE = 50000000;
const JST = 9 * 3600 * 1000;
const DAY = 24 * 3600 * 1000;
const PERIODS = ['day', 'week', 'month', 'year', 'all'];

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

function log_() {
  const ss = book_();
  let sh = ss.getSheetByName(LOG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(LOG_SHEET);
    sh.appendRow(['time', 'key', 'name', 'score', 'stage', 'combo']);
    const old = ss.getSheetByName(OLD_SHEET);
    if (old) {
      const rows = old.getDataRange().getValues().slice(1).filter(function (r) { return r[0]; });
      const add = rows.map(function (r) { return [r[5] instanceof Date ? r[5] : new Date(), r[0], r[1], r[2], r[3], r[4]]; });
      if (add.length) sh.getRange(2, 1, add.length, 6).setValues(add);
    }
  }
  return sh;
}

// 各期間の開始時刻（ミリ秒）
function starts_(now) {
  const j = new Date(now + JST);
  const y = j.getUTCFullYear(), m = j.getUTCMonth(), d = j.getUTCDate();
  const day = Date.UTC(y, m, d) - JST;
  const dow = (j.getUTCDay() + 6) % 7; // 月曜=0
  return { day: day, week: day - dow * DAY, month: Date.UTC(y, m, 1) - JST, year: Date.UTC(y, 0, 1) - JST, all: 0 };
}

function boards_(key) {
  const rows = log_().getDataRange().getValues().slice(1);
  const st = starts_(Date.now());
  const latestName = {};
  const best = {}; PERIODS.forEach(function (p) { best[p] = {}; });
  rows.forEach(function (r) {
    const t = r[0] instanceof Date ? r[0].getTime() : Number(new Date(r[0])) || 0;
    const k = String(r[1]); const score = Number(r[3]) || 0;
    if (!k || score <= 0) return;
    if (!latestName[k] || t >= latestName[k].t) latestName[k] = { t: t, name: String(r[2]) };
    PERIODS.forEach(function (p) {
      if (t < st[p]) return;
      const b = best[p][k];
      if (!b || score > b.score) best[p][k] = { score: score, stage: Number(r[4]) || 1, combo: Number(r[5]) || 0 };
    });
  });
  const out = {}, mine = {};
  PERIODS.forEach(function (p) {
    const list = Object.keys(best[p]).map(function (k) {
      const b = best[p][k];
      return { k: k, name: latestName[k] ? latestName[k].name : '名無しのパスタ', score: b.score, stage: b.stage, combo: b.combo };
    });
    list.sort(function (a, b) { return b.score - a.score; });
    let myRank = 0, myBest = 0;
    list.forEach(function (e, i) { if (key && e.k === key) { myRank = i + 1; myBest = e.score; } });
    mine[p] = { rank: myRank, best: myBest, total: list.length };
    out[p] = list.slice(0, TOP_N).map(function (e) {
      return { name: e.name, score: e.score, stage: e.stage, combo: e.combo, me: !!key && e.k === key };
    });
  });
  return { boards: out, mine: mine };
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function cleanName_(v) {
  let name = String(v || '').replace(/[\u0000-\u001f\u007f<>]/g, '').trim().slice(0, 12);
  if (!name) name = '名無しのパスタ';
  if (/^[=+\-@]/.test(name)) name = "'" + name; // 数式として解釈させない
  return name;
}

function doGet(e) {
  const key = String((e && e.parameter && e.parameter.key) || '');
  const b = boards_(key);
  return out_({ ok: true, boards: b.boards, mine: b.mine });
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
    const before = boards_(key).mine.all.best;
    log_().appendRow([new Date(), key, name, score, stage, combo]);
    const b = boards_(key);
    return out_({ ok: true, improved: score > before, best: Math.max(before, score), boards: b.boards, mine: b.mine });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
