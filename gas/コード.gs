/**
 * 大泉衛生 営業進捗ツール ― サーバー側(Google Apps Script)
 *
 * 役割: Webツール(スマホ)からの登録・更新を受け取り、スプレッドシートに読み書きする。
 *
 * ■ 使う前にやること
 *   1. 下の SHEET_ID を、Webツール用のスプレッドシートのIDに書き換える
 *      （URLの /d/ と /edit の間の文字列）
 *   2. 「デプロイ」→「新しいデプロイ」→種類「ウェブアプリ」
 *      次のユーザーとして実行: 自分
 *      アクセスできるユーザー: 全員
 *   3. 発行されたURLを Webツール側に設定する
 *
 * 列の順番はスプレッドシートの見出し行から自動で読み取るので、
 * 列を増やしたり並べ替えたりしても動きます（見出しの文字は変えないでください）。
 */

var SHEET_ID = 'ここにスプレッドシートのIDを貼る';
var SHEET_NAME = '';            // 空ならブックの最初のシートを使う
var TARGETS = { '木村': 22000, '原田': 22000, '藤川': 16000 };   // 個人の月間目標(円)

// 見出し行に必ずある列名。この行を見出し行と判断する目印にする。
var KEY_HEADER = '見込み日';
// ツールが読み書きする列（スプレッドシートの見出しと同じ文字）
var COL = {
  no: 'No.', date: '見込み日', person: '担当', industry: '業種', area: 'エリア',
  address: '住所詳細', company: '管理会社', done: '完成日', status: 'ステータス',
  quote: '見積金額(円)', contract: '契約金額(円)', next: '次回アクション',
  reason: 'クローズ理由', memo: '備考・メモ', id: 'ID', updated: '更新日時'
};

// スプレッドシートが空のときに使う既定の選択肢
var FALLBACK = {
  status: ['🔵 アプローチ中', '🟡 見積・交渉中', '🟢 契約済', '⚫ クローズ', '⬜ 時期待ち'],
  next: ['📝 見積提出', '📄 見積フォロー', '📞 TEL', '🔄 再訪問', '🏁 完成前再訪', '✅ 完了', '—'],
  reason: ['契約済み業者あり', '指定業者あり', '他社価格が安い', '予算なし', '連絡不通', 'オーナー不在', 'その他'],
  person: ['木村', '原田', '藤川'],
  industry: ['居酒屋', '焼き肉', 'ラーメン', '和食', 'イタリアン', 'カレー', 'バー', '立ち飲み', '喫茶店',
             '美容室', 'サービス業', '病院', '介護施設', 'ホテル', '民泊', 'マンション', 'アパート',
             'テナントビル', 'オフィスビル', '事務所', 'スーパー', 'その他'],
  area: ['北区', '都島区', '福島区', '此花区', '中央区', '西区', '港区', '大正区', '天王寺区', '浪速区',
         '西淀川区', '淀川区', '東淀川区', '東成区', '生野区', '旭区', '城東区', '鶴見区', '阿倍野区',
         '住之江区', '住吉区', '東住吉区', '西成区', '平野区',
         '堺市堺区', '堺市中区', '堺市東区', '堺市西区', '堺市南区', '堺市北区', '堺市美原区', 'その他']
};

/* ================= 共通 ================= */

function sheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  return SHEET_NAME ? ss.getSheetByName(SHEET_NAME) : ss.getSheets()[0];
}

/** 見出し行の位置と、列名→列番号の対応を返す */
function layout_(sh) {
  var scan = sh.getRange(1, 1, Math.min(10, sh.getLastRow()), sh.getLastColumn()).getValues();
  for (var r = 0; r < scan.length; r++) {
    if (scan[r].indexOf(KEY_HEADER) >= 0) {
      var map = {};
      scan[r].forEach(function (h, i) { if (h) map[String(h).trim()] = i + 1; });
      return { headerRow: r + 1, map: map, firstData: r + 2 };
    }
  }
  throw new Error('見出し行が見つかりません（「' + KEY_HEADER + '」の列が必要です）');
}

/** ID・更新日時の列がなければ右端に足す（既存のデータには触らない） */
function ensureColumns_(sh, lay) {
  [COL.id, COL.updated].forEach(function (name) {
    if (!lay.map[name]) {
      var col = sh.getLastColumn() + 1;
      sh.getRange(lay.headerRow, col).setValue(name);
      lay.map[name] = col;
    }
  });
  return lay;
}

function readAll_() {
  var sh = sheet_();
  var lay = ensureColumns_(sh, layout_(sh));
  var last = sh.getLastRow();
  if (last < lay.firstData) return { sh: sh, lay: lay, rows: [] };
  var values = sh.getRange(lay.firstData, 1, last - lay.firstData + 1, sh.getLastColumn()).getValues();
  var rows = [];
  values.forEach(function (v, i) {
    var get = function (name) {
      var c = lay.map[name];
      return c ? v[c - 1] : '';
    };
    // 完全な空行は飛ばす
    if (!String(get(COL.date)) && !String(get(COL.address)) && !String(get(COL.status))) return;
    rows.push({
      row: lay.firstData + i,
      id: String(get(COL.id) || ''),
      no: String(get(COL.no) || ''),
      date: fmtDate_(get(COL.date)),
      person: String(get(COL.person) || ''),
      industry: String(get(COL.industry) || ''),
      area: String(get(COL.area) || ''),
      address: String(get(COL.address) || ''),
      company: String(get(COL.company) || ''),
      done: fmtDate_(get(COL.done)),
      status: String(get(COL.status) || ''),
      quote: num_(get(COL.quote)),
      contract: num_(get(COL.contract)),
      next: String(get(COL.next) || ''),
      reason: String(get(COL.reason) || ''),
      memo: String(get(COL.memo) || ''),
      updated: fmtDate_(get(COL.updated))
    });
  });
  return { sh: sh, lay: lay, rows: rows };
}

function fmtDate_(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy/MM/dd');
  }
  return String(v);
}

function num_(v) {
  if (v === '' || v === null || v === undefined) return 0;
  var n = Number(String(v).replace(/[^0-9.\-]/g, ''));
  return isNaN(n) ? 0 : n;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** その列に実際に入っている値を多い順に返す（選択肢を実データに合わせる） */
function optionsOf_(rows, key, fallback) {
  var count = {};
  rows.forEach(function (r) {
    var v = String(r[key] || '').trim();
    if (v) count[v] = (count[v] || 0) + 1;
  });
  var list = Object.keys(count).sort(function (a, b) { return count[b] - count[a]; });
  fallback.forEach(function (v) { if (list.indexOf(v) < 0) list.push(v); });
  return list;
}

/* ================= 読み取り ================= */

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'bootstrap';
  var data = readAll_();
  if (action === 'list') {
    return json_({ ok: true, rows: data.rows });
  }
  // bootstrap: 選択肢・目標・全データをまとめて返す（ツールはこれ1回で動く）
  return json_({
    ok: true,
    targets: TARGETS,
    masters: {
      person: optionsOf_(data.rows, 'person', FALLBACK.person),
      industry: optionsOf_(data.rows, 'industry', FALLBACK.industry),
      area: optionsOf_(data.rows, 'area', FALLBACK.area),
      status: optionsOf_(data.rows, 'status', FALLBACK.status),
      next: optionsOf_(data.rows, 'next', FALLBACK.next),
      reason: optionsOf_(data.rows, 'reason', FALLBACK.reason)
    },
    rows: data.rows
  });
}

/* ================= 書き込み ================= */

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var p = JSON.parse(e.postData.contents);
    if (p.action === 'create') return json_(create_(p.item));
    if (p.action === 'update') return json_(update_(p.item));
    return json_({ ok: false, error: '不明な操作です' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function setCells_(sh, lay, row, item) {
  var put = function (name, value) {
    var c = lay.map[name];
    if (c) sh.getRange(row, c).setValue(value);
  };
  if (item.date !== undefined) put(COL.date, item.date ? new Date(item.date) : '');
  if (item.person !== undefined) put(COL.person, item.person);
  if (item.industry !== undefined) put(COL.industry, item.industry);
  if (item.area !== undefined) put(COL.area, item.area);
  if (item.address !== undefined) put(COL.address, item.address);
  if (item.company !== undefined) put(COL.company, item.company);
  if (item.done !== undefined) put(COL.done, item.done ? new Date(item.done) : '');
  if (item.status !== undefined) put(COL.status, item.status);
  if (item.quote !== undefined) put(COL.quote, item.quote === '' ? '' : Number(item.quote));
  if (item.contract !== undefined) put(COL.contract, item.contract === '' ? '' : Number(item.contract));
  if (item.next !== undefined) put(COL.next, item.next);
  if (item.reason !== undefined) put(COL.reason, item.reason);
  if (item.memo !== undefined) put(COL.memo, item.memo);
  put(COL.updated, new Date());
}

function create_(item) {
  var sh = sheet_();
  var lay = ensureColumns_(sh, layout_(sh));
  var row = Math.max(sh.getLastRow() + 1, lay.firstData);
  var id = Utilities.getUuid();
  sh.getRange(row, lay.map[COL.id]).setValue(id);
  // No. は既存の最大値+1
  if (lay.map[COL.no]) {
    var maxNo = 0;
    if (sh.getLastRow() >= lay.firstData) {
      sh.getRange(lay.firstData, lay.map[COL.no], sh.getLastRow() - lay.firstData + 1, 1)
        .getValues().forEach(function (v) { maxNo = Math.max(maxNo, num_(v[0])); });
    }
    sh.getRange(row, lay.map[COL.no]).setValue(maxNo + 1);
  }
  setCells_(sh, lay, row, item);
  return { ok: true, id: id, row: row };
}

function update_(item) {
  var data = readAll_();
  var hit = null;
  data.rows.forEach(function (r) {
    if (item.id && r.id === item.id) hit = r;
    else if (!item.id && item.row && r.row === item.row) hit = r;
  });
  if (!hit) return { ok: false, error: '対象の行が見つかりませんでした' };
  if (!hit.id) {
    var newId = Utilities.getUuid();
    data.sh.getRange(hit.row, data.lay.map[COL.id]).setValue(newId);
    hit.id = newId;
  }
  setCells_(data.sh, data.lay, hit.row, item);
  return { ok: true, id: hit.id, row: hit.row };
}

/* ================= 動作確認用 ================= */

function テスト_読み取り() {
  var d = readAll_();
  Logger.log('件数: ' + d.rows.length);
  Logger.log(d.rows.slice(0, 3));
}
