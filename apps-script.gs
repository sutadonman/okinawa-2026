/**
 * 沖縄2026 スポット希望リスト — wishlist.html の提出フォームの受け口
 *
 * Google スプレッドシート「沖縄2026 スポット希望リスト」に紐づけて、
 * ウェブアプリとしてデプロイして使う。手順は README.md を参照。
 *
 *   doGet  … シートの中身を JSONP で返す（ダッシュボードの集計用）
 *   doPost … action に応じて 追加 / 優先度などの修正 / 取り下げ を行う
 *
 * 修正と取り下げは、行のIDと提出者名の両方が一致したときだけ通す。
 * ログインを求めない代わりの簡易な持ち主チェックなので、
 * 「他人のふりをすれば消せる」点は許容している（身内3人での運用のため）。
 */

// ── 設定 ─────────────────────────────────────────────
// シートのURL /d/ と /edit の間にある文字列
var SHEET_ID = '12VYd7jl6_IPttbCkA974p-PhVPWRuAK9A6laz8hJABI';

// wishlist.html 側の TOKEN と同じ文字列にすること。
// いたずら投稿を止めるだけの合言葉で、秘密の情報は入れない。
var TOKEN = 'okinawa2026-3nin-2f9a41c7';

// シートの列の並び。7列目のIDは、修正・取り下げのために後から足したもの
var HEADERS = ['提出者', 'GoogleマップURL', '優先度', 'スポット名', 'ひとこと', '採否', 'ID'];
var COL_WHO = 1, COL_URL = 2, COL_PRI = 3, COL_SPOT = 4, COL_MEMO = 5, COL_STATUS = 6, COL_ID = 7;

// 提出者はこの3人だけ。wishlist.html の選択肢と揃えること
var MEMBERS = ['Otsu', 'Sugi', 'Runto'];

// 一度に受け付ける最大件数（取りこぼしより暴走を止めることを優先）
var MAX_ITEMS = 20;
// ─────────────────────────────────────────────────────


function sheet_() {
  return SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
}

/**
 * callback が来ていれば JSONP、無ければ素の JSON で返す。
 * ブラウザから別オリジンで読むため、GET は JSONP を使う。
 */
function out_(obj, callback) {
  var body = JSON.stringify(obj);
  if (callback && /^[A-Za-z_$][A-Za-z0-9_$]*$/.test(callback)) {
    return ContentService
      .createTextOutput(callback + '(' + body + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(body)
    .setMimeType(ContentService.MimeType.JSON);
}

function trim_(v, max) {
  return String(v == null ? '' : v).replace(/[\r\n\t]+/g, ' ').trim().slice(0, max);
}

function normPri_(v) {
  var s = String(v || '').trim().toLowerCase();
  return (s === 'high' || s === 'middle' || s === 'low') ? s : 'middle';
}

function newId_() {
  return 'r' + Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
}

/** ID列の見出しと、まだIDが無い既存行を埋める。何も無ければ書き込まない */
function ensureIds_(sh) {
  var last = sh.getLastRow();
  if (last < 1) return;

  if (trim_(sh.getRange(1, COL_ID).getValue(), 20) !== 'ID') {
    sh.getRange(1, COL_ID).setValue('ID');
  }
  if (last < 2) return;

  var ids = sh.getRange(2, COL_ID, last - 1, 1).getValues();
  var changed = false;
  for (var i = 0; i < ids.length; i++) {
    if (!trim_(ids[i][0], 40)) { ids[i][0] = newId_(); changed = true; }
  }
  if (changed) sh.getRange(2, COL_ID, ids.length, 1).setValues(ids);
}

/** IDから行番号を引く。見つからなければ -1 */
function rowOfId_(sh, id) {
  var last = sh.getLastRow();
  if (last < 2) return -1;
  var ids = sh.getRange(2, COL_ID, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (trim_(ids[i][0], 40) === id) return i + 2;
  }
  return -1;
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}


/** ダッシュボードがシートを読むための入口 */
function doGet(e) {
  var cb = (e && e.parameter) ? e.parameter.callback : null;
  try {
    var values = withLock_(function () {
      var sh = sheet_();
      ensureIds_(sh);
      SpreadsheetApp.flush();
      return sh.getDataRange().getDisplayValues();
    });
    return out_({ ok: true, rows: values, fetchedAt: new Date().toISOString() }, cb);
  } catch (err) {
    return out_({ ok: false, error: String(err) }, cb);
  }
}


/** 追加・修正・取り下げの入口 */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return out_({ ok: false, error: 'no body' });
    }
    var body = JSON.parse(e.postData.contents);
    if (body.token !== TOKEN) return out_({ ok: false, error: 'token' });

    var action = body.action || 'add';
    if (action === 'add') return addItems_(body);
    if (action === 'update') return updateItem_(body);
    if (action === 'remove') return removeItem_(body);
    return out_({ ok: false, error: 'unknown action' });

  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}


function addItems_(body) {
  var items = Array.isArray(body.items) ? body.items : [];
  if (!items.length) return out_({ ok: false, error: 'empty' });
  if (items.length > MAX_ITEMS) return out_({ ok: false, error: 'too many' });

  var rows = [], ids = [];
  for (var i = 0; i < items.length; i++) {
    var it = items[i] || {};
    var who = trim_(it.who, 40);
    var url = trim_(it.url, 500);
    var spot = trim_(it.spot, 80);
    // 提出者は決め打ちの3人のみ。URLかスポット名のどちらかがあれば受け付ける
    if (MEMBERS.indexOf(who) < 0 || (!url && !spot)) continue;
    if (url && !/^https?:\/\//i.test(url)) continue;
    var id = newId_();
    ids.push(id);
    rows.push([who, url, normPri_(it.pri), spot, trim_(it.memo, 200), '未定', id]);
  }
  if (!rows.length) return out_({ ok: false, error: 'invalid' });

  // 3人が同時に送ったときに同じ行へ重ね書きしないよう直列化する
  return withLock_(function () {
    var sh = sheet_();
    ensureIds_(sh);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    }
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, HEADERS.length).setValues(rows);
    SpreadsheetApp.flush();
    return out_({ ok: true, added: rows.length, ids: ids });
  });
}


function updateItem_(body) {
  var id = trim_(body.id, 40);
  var who = trim_(body.who, 40);
  if (!id || MEMBERS.indexOf(who) < 0) return out_({ ok: false, error: 'invalid' });

  return withLock_(function () {
    var sh = sheet_();
    ensureIds_(sh);
    var row = rowOfId_(sh, id);
    if (row < 0) return out_({ ok: false, error: 'not found' });
    // 自分が出した行以外は触らせない
    if (trim_(sh.getRange(row, COL_WHO).getValue(), 40) !== who) {
      return out_({ ok: false, error: 'not yours' });
    }

    var changed = 0;
    if (body.pri != null) { sh.getRange(row, COL_PRI).setValue(normPri_(body.pri)); changed++; }
    if (body.spot != null) { sh.getRange(row, COL_SPOT).setValue(trim_(body.spot, 80)); changed++; }
    if (body.memo != null) { sh.getRange(row, COL_MEMO).setValue(trim_(body.memo, 200)); changed++; }
    if (!changed) return out_({ ok: false, error: 'nothing to change' });

    SpreadsheetApp.flush();
    return out_({ ok: true, updated: 1 });
  });
}


function removeItem_(body) {
  var id = trim_(body.id, 40);
  var who = trim_(body.who, 40);
  if (!id || MEMBERS.indexOf(who) < 0) return out_({ ok: false, error: 'invalid' });

  return withLock_(function () {
    var sh = sheet_();
    ensureIds_(sh);
    var row = rowOfId_(sh, id);
    if (row < 0) return out_({ ok: false, error: 'not found' });
    if (trim_(sh.getRange(row, COL_WHO).getValue(), 40) !== who) {
      return out_({ ok: false, error: 'not yours' });
    }
    sh.deleteRow(row);
    SpreadsheetApp.flush();
    return out_({ ok: true, removed: 1 });
  });
}
