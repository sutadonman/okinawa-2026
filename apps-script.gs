/**
 * 沖縄2026 スポット希望リスト — wishlist.html の提出フォームの受け口
 *
 * Google スプレッドシート「沖縄2026 スポット希望リスト」に紐づけて、
 * ウェブアプリとしてデプロイして使う。手順は README.md を参照。
 *
 *   doPost … フォームから届いた提出をシートの末尾に追記する
 *   doGet  … シートの中身を JSONP で返す（ダッシュボードの集計用）
 */

// ── 設定 ─────────────────────────────────────────────
// シートのURL /d/ と /edit の間にある文字列
var SHEET_ID = '12VYd7jl6_IPttbCkA974p-PhVPWRuAK9A6laz8hJABI';

// wishlist.html 側の TOKEN と同じ文字列にすること。
// いたずら投稿を止めるだけの合言葉で、秘密の情報は入れない。
var TOKEN = 'okinawa2026-3nin-2f9a41c7';

// シートの列の並び。テンプレートCSVと同じ順序
var HEADERS = ['提出者', 'GoogleマップURL', '優先度', 'スポット名', 'ひとこと', '採否'];

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


/** ダッシュボードがシートを読むための入口 */
function doGet(e) {
  var cb = (e && e.parameter) ? e.parameter.callback : null;
  try {
    var values = sheet_().getDataRange().getDisplayValues();
    return out_({ ok: true, rows: values, fetchedAt: new Date().toISOString() }, cb);
  } catch (err) {
    return out_({ ok: false, error: String(err) }, cb);
  }
}


/** 提出フォームからの書き込み */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return out_({ ok: false, error: 'no body' });
    }

    var body = JSON.parse(e.postData.contents);
    if (body.token !== TOKEN) {
      return out_({ ok: false, error: 'token' });
    }

    var items = Array.isArray(body.items) ? body.items : [];
    if (!items.length) {
      return out_({ ok: false, error: 'empty' });
    }
    if (items.length > MAX_ITEMS) {
      return out_({ ok: false, error: 'too many' });
    }

    var rows = [];
    for (var i = 0; i < items.length; i++) {
      var it = items[i] || {};
      var who = trim_(it.who, 40);
      var url = trim_(it.url, 500);
      var spot = trim_(it.spot, 80);
      // 提出者は決め打ちの3人のみ。URLかスポット名のどちらかがあれば受け付ける
      if (MEMBERS.indexOf(who) < 0 || (!url && !spot)) continue;
      if (url && !/^https?:\/\//i.test(url)) continue;
      rows.push([who, url, normPri_(it.pri), spot, trim_(it.memo, 200), '未定']);
    }
    if (!rows.length) {
      return out_({ ok: false, error: 'invalid' });
    }

    // 3人が同時に送ったときに同じ行へ重ね書きしないよう直列化する
    var lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      var sh = sheet_();
      if (sh.getLastRow() === 0) {
        sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      }
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, HEADERS.length).setValues(rows);
      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }

    return out_({ ok: true, added: rows.length });

  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}
