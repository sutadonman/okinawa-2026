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

// 採否の取りうる値。不採用はシートに残したまま、地図と集計から外す
var STATUSES = ['採用', '未定', '不採用'];

// 一度に受け付ける最大件数（取りこぼしより暴走を止めることを優先）
var MAX_ITEMS = 20;
// ─────────────────────────────────────────────────────


function sheet_() {
  return SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
}

/**
 * 承認を取り直すための関数。エディタから手で1回実行する。
 * 短縮URLの展開には外部URLへのアクセス権が要るが、ウェブアプリを
 * 再デプロイしただけでは承認画面が出ないことがあるため、ここで明示的に呼ぶ。
 * 例外を握りつぶさないので、権限が無ければそのまま実行ログに出る。
 */
function authorize() {
  var res = UrlFetchApp.fetch('https://maps.app.goo.gl/JwsBq7HeYfy36uPe8', {
    followRedirects: false,
    muteHttpExceptions: true,
    headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' }
  });
  var h = res.getAllHeaders();
  Logger.log('code=%s location=%s', res.getResponseCode(), h['Location'] || h['location'] || '(なし)');
  Logger.log('シート=%s', SpreadsheetApp.openById(SHEET_ID).getName());
  return 'OK';
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

// 展開に失敗した理由を残す。doGet?debug=... で外から読めるようにして、
// 「黙って解決されない」状態を追えるようにする
var EXPAND_NOTE = '';

/**
 * 短縮URL（maps.app.goo.gl）を実URLに展開する。
 * ブラウザからは転送先を辿れないので、サーバー側であるここで解決してしまう。
 * 展開できなければ null。
 */
function expandUrl_(url) {
  if (!/^https?:\/\/(maps\.app\.goo\.gl|goo\.gl\/maps)/i.test(url)) return null;
  var cur = url;
  for (var i = 0; i < 6; i++) {
    var res;
    try {
      res = UrlFetchApp.fetch(cur, {
        followRedirects: false,
        muteHttpExceptions: true,
        // UA を伏せると転送ではなく案内ページを返されることがあるため、ブラウザを名乗る
        headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' }
      });
    } catch (e) {
      EXPAND_NOTE = 'fetch例外: ' + e;
      return null;
    }

    var code = res.getResponseCode();

    if (code >= 300 && code < 400) {
      var h = res.getAllHeaders();
      var loc = h['Location'] || h['location'];
      if (loc instanceof Array) loc = loc[0];
      if (!loc) { EXPAND_NOTE = code + ' だが Location なし'; return null; }
      cur = String(loc);
      if (/\/maps\/place\/|!3d-?\d/.test(cur)) return cur;
      continue;
    }

    // 200 が返る場合、本文（HTML）の中に本当の遷移先が入っている
    var body = '';
    try { body = res.getContentText(); } catch (e) { body = ''; }
    var m = body.match(/https:\/\/www\.google\.com\/maps\/[^"'\\\s<>]+/);
    if (m) return m[0].replace(/&amp;/g, '&');

    EXPAND_NOTE = 'code=' + code + ' 本文冒頭=' + body.slice(0, 150);
    return null;
  }
  EXPAND_NOTE = '転送が多すぎます';
  return null;
}

/**
 * 展開後のURLから名称と座標を取り出す。
 * `@lat,lng` は地図の表示中心で実際の地点とズレることがあるため、
 * 施設の座標である `!3d/!4d` を優先する。
 */
function parsePlace_(url) {
  var out = { name: '', lat: null, lng: null };
  var m = url.match(/!3d(-?\d+\.\d+)!4d(-?\d+\.\d+)/) ||
          url.match(/[?&]q=(-?\d+\.\d+),\s*(-?\d+\.\d+)/) ||
          url.match(/@(-?\d+\.\d+),(-?\d+\.\d+)/);
  if (m) { out.lat = parseFloat(m[1]); out.lng = parseFloat(m[2]); }
  var n = url.match(/\/maps\/place\/([^\/@?]+)/);
  if (n) {
    try { out.name = decodeURIComponent(n[1]).replace(/\+/g, ' ').trim(); } catch (e) {}
  }
  return out;
}

/**
 * 短縮URLを、座標つきの素直な形に置き換える。
 * 返り値は { url, spot }。解決できなければ null。
 */
function resolveShort_(url, spot) {
  var expanded = expandUrl_(url);
  if (!expanded) return null;
  var p = parsePlace_(expanded);
  if (p.lat == null) return null;
  return {
    url: 'https://www.google.com/maps/search/?api=1&query=' + p.lat + ',' + p.lng,
    spot: spot || p.name
  };
}

/**
 * シートに残っている短縮URLを解決して書き戻す。
 * 1回のリクエストで触る件数を絞って、doGet が重くならないようにする。
 */
function backfillShortUrls_(sh, max) {
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var rng = sh.getRange(2, COL_URL, last - 1, 3);   // URL / 優先度 / スポット名
  var vals = rng.getValues();
  var done = 0;
  for (var i = 0; i < vals.length && done < max; i++) {
    var url = trim_(vals[i][0], 500);
    if (!/^https?:\/\/(maps\.app\.goo\.gl|goo\.gl\/maps)/i.test(url)) continue;
    var r = resolveShort_(url, trim_(vals[i][2], 80));
    if (!r) continue;
    vals[i][0] = r.url;
    vals[i][2] = r.spot;
    done++;
  }
  if (done) rng.setValues(vals);
  return done;
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}


/** ダッシュボードがシートを読むための入口 */
function doGet(e) {
  var cb = (e && e.parameter) ? e.parameter.callback : null;

  // 診断用：?debug=<短縮URL> で展開の結果と失敗理由をそのまま返す
  if (e && e.parameter && e.parameter.debug) {
    var target = e.parameter.debug;
    EXPAND_NOTE = '';
    var expanded = expandUrl_(target);
    return out_({
      ok: true, target: target, expanded: expanded,
      parsed: expanded ? parsePlace_(expanded) : null,
      note: EXPAND_NOTE
    }, cb);
  }

  try {
    var values = withLock_(function () {
      var sh = sheet_();
      ensureIds_(sh);
      // 短縮URLのまま残っている行をここで解決する。1回の読み込みにつき数件ずつ
      backfillShortUrls_(sh, 5);
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
    if (action === 'status') return setStatus_(body);
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
    // 短縮URLはこの場で座標つきに直す。名前が空なら地点名も貰う
    var res = url ? resolveShort_(url, spot) : null;
    if (res) { url = res.url; spot = res.spot; }
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


/**
 * 採否を切り替える。
 * 優先度と違って「みんなで決めるもの」なので、行の持ち主でなくても変更できる。
 * 同じスポットが複数人から出ていると行も複数あるため、IDをまとめて受け取る。
 */
function setStatus_(body) {
  var who = trim_(body.who, 40);
  var status = trim_(body.status, 10);
  var ids = Array.isArray(body.ids) ? body.ids : (body.id ? [body.id] : []);
  if (MEMBERS.indexOf(who) < 0) return out_({ ok: false, error: 'invalid' });
  if (STATUSES.indexOf(status) < 0) return out_({ ok: false, error: 'bad status' });
  if (!ids.length || ids.length > MAX_ITEMS) return out_({ ok: false, error: 'invalid' });

  return withLock_(function () {
    var sh = sheet_();
    ensureIds_(sh);
    var last = sh.getLastRow();
    if (last < 2) return out_({ ok: false, error: 'not found' });

    var idCol = sh.getRange(2, COL_ID, last - 1, 1).getValues();
    var stCol = sh.getRange(2, COL_STATUS, last - 1, 1).getValues();
    var changed = 0;
    for (var i = 0; i < idCol.length; i++) {
      if (ids.indexOf(trim_(idCol[i][0], 40)) >= 0) { stCol[i][0] = status; changed++; }
    }
    if (!changed) return out_({ ok: false, error: 'not found' });

    sh.getRange(2, COL_STATUS, stCol.length, 1).setValues(stCol);
    SpreadsheetApp.flush();
    return out_({ ok: true, changed: changed });
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
