/**
 * 大泉衛生 営業進捗ツール ― サーバー側(Google Apps Script)
 *
 * 役割: Webツール(スマホ)からの登録・更新・訪問記録を受け取り、スプレッドシートに読み書きする。
 *
 * ■ 使う前にやること
 *   1. SHEET_ID は設定済みです（（新）大泉衛生営業進捗管理）。書き換え不要。
 *   2. 「デプロイ」→「新しいデプロイ」→種類「ウェブアプリ」
 *      次のユーザーとして実行: 自分
 *      アクセスできるユーザー: 全員
 *   3. 発行されたURLを Webツール側に設定する
 *
 * ■ シートの構成
 *   1枚目 … 案件のリスト（既存の「営業進捗リスト」）
 *   活動ログ … 訪問1回につき1行。なければ自動で作ります。
 *
 * 列の順番は見出し行から自動で読み取るので、列を増やしたり並べ替えても動きます
 * （見出しの文字は変えないでください）。
 */

var SHEET_ID = '10vKZyOb7-RBx6jTgQMev-ojPoEqJpQkm41qRKJFB24Q';   // （新）大泉衛生営業進捗管理
var SHEET_NAME = '';            // 空ならブックの最初のシートを使う
var LOG_SHEET_NAME = '活動ログ';
var TARGETS = { '木村': 22000, '原田': 22000, '藤川': 16000 };   // 個人の月間目標(円)

// 見出し行に必ずある列名。この行を見出し行と判断する目印にする。
var KEY_HEADER = '見込み日';
var COL = {
  no: 'No.', date: '見込み日', person: '担当', industry: '業種', area: 'エリア',
  address: '住所詳細', company: '管理会社', done: '完成日', status: 'ステータス',
  quote: '見積金額(円)', contract: '契約金額(円)', next: '次回アクション',
  reason: 'クローズ理由', memo: '備考・メモ',
  quoteDate: '見積提出日', contractDate: '契約日', closeDate: 'クローズ日',
  lastAp: '最終AP日', apCount: 'AP回数',
  id: 'ID', updated: '更新日時'
};
var LOG_COL = ['日時', '案件ID', '担当', 'エリア', '住所詳細', '方法', '結果', 'メモ', '登録日時'];

// スプレッドシートが空のときに使う既定の選択肢
var FALLBACK = {
  status: ['🔵 アプローチ中', '🟡 見積・交渉中', '🟠 契約予定', '🟢 契約済', '⚫ クローズ', '⬜ 時期待ち'],
  method: ['飛び込み訪問', '再訪問', 'TEL', 'DM・郵送', 'メール', '紹介・つながり'],
  result: ['担当者と面談できた', '名刺・パンフを置いた', '不在だった', '門前払い',
           '見積依頼をもらった', '見積を提出した', '後日連絡の約束', '断られた', '契約になった'],
  next: ['📝 見積提出', '📄 見積フォロー', '📞 TEL', '🔄 再訪問', '🏁 完成前再訪', '✅ 完了', '—'],
  reason: ['契約済み業者あり', '指定業者あり', '他社価格が安い', '予算なし', '連絡不通', 'オーナー不在', 'その他'],
  person: ['木村', '原田', '藤川'],
  industry: ['居酒屋', '焼き肉', 'ラーメン', '和食', 'イタリアン', 'カレー', 'バー', '立ち飲み', '喫茶店',
             '美容室', 'サービス業', '病院', '介護施設', 'ホテル', '民泊', 'マンション', 'アパート',
             'テナントビル', 'オフィスビル', '事務所', '店舗', 'スーパー', 'その他'],
  area: ['北区', '都島区', '福島区', '此花区', '中央区', '西区', '港区', '大正区', '天王寺区', '浪速区',
         '西淀川区', '淀川区', '東淀川区', '東成区', '生野区', '旭区', '城東区', '鶴見区', '阿倍野区',
         '住之江区', '住吉区', '東住吉区', '西成区', '平野区',
         '堺市堺区', '堺市中区', '堺市東区', '堺市西区', '堺市南区', '堺市北区', '堺市美原区', 'その他']
};

// 結果からステータスを決める。進んだときだけ書き換え、後戻りはさせない。
var RESULT_TO_STATUS = {
  '契約になった': '🟢 契約済',
  '断られた': '⚫ クローズ',
  '見積を提出した': '🟡 見積・交渉中',
  '見積依頼をもらった': '🟡 見積・交渉中'
};
var STATUS_RANK = { 'アプローチ中': 1, '見積': 2, '交渉': 2, '契約予定': 3, '契約済': 4 };

/* ================= 共通 ================= */

function book_() { return SpreadsheetApp.openById(SHEET_ID); }

function sheet_() {
  var ss = book_();
  return SHEET_NAME ? ss.getSheetByName(SHEET_NAME) : ss.getSheets()[0];
}

/** 活動ログのシート。なければ作る。 */
function logSheet_() {
  var ss = book_();
  var sh = ss.getSheetByName(LOG_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(LOG_SHEET_NAME);
    sh.getRange(1, 1, 1, LOG_COL.length).setValues([LOG_COL]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
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

/** ツールが使う列がなければ右端に足す（既存のデータには触らない） */
function ensureColumns_(sh, lay) {
  [COL.quoteDate, COL.contractDate, COL.closeDate, COL.lastAp, COL.apCount, COL.id, COL.updated]
    .forEach(function (name) {
      if (!lay.map[name]) {
        var col = sh.getLastColumn() + 1;
        sh.getRange(lay.headerRow, col).setValue(name);
        lay.map[name] = col;
      }
    });
  return lay;
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

function today_() { return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd'); }

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
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
      quoteDate: fmtDate_(get(COL.quoteDate)),
      contractDate: fmtDate_(get(COL.contractDate)),
      closeDate: fmtDate_(get(COL.closeDate)),
      lastAp: fmtDate_(get(COL.lastAp)),
      apCount: num_(get(COL.apCount)),
      updated: fmtDate_(get(COL.updated))
    });
  });
  return { sh: sh, lay: lay, rows: rows };
}

function readLogs_() {
  var sh = logSheet_();
  var last = sh.getLastRow();
  if (last < 2) return [];
  var v = sh.getRange(2, 1, last - 1, LOG_COL.length).getValues();
  return v.filter(function (r) { return r[0] || r[1]; }).map(function (r) {
    return {
      date: fmtDate_(r[0]), caseId: String(r[1] || ''), person: String(r[2] || ''),
      area: String(r[3] || ''), address: String(r[4] || ''),
      method: String(r[5] || ''), result: String(r[6] || ''), memo: String(r[7] || '')
    };
  });
}

/** 実データに出てきた値を多い順に返す（業種・エリア用） */
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

/** 決まった順番で返す（ステータス・方法・結果など、並び順に意味があるもの用） */
function orderedOptionsOf_(rows, key, fixed) {
  var seen = {};
  rows.forEach(function (r) {
    var v = String(r[key] || '').trim();
    if (v) seen[v] = true;
  });
  var list = fixed.slice();
  Object.keys(seen).forEach(function (v) { if (list.indexOf(v) < 0) list.push(v); });
  return list;
}

/* ================= 読み取り ================= */

function doGet(e) {
  var data = readAll_();
  var logs = readLogs_();
  return json_({
    ok: true,
    targets: TARGETS,
    masters: {
      person: optionsOf_(data.rows, 'person', FALLBACK.person),
      industry: optionsOf_(data.rows, 'industry', FALLBACK.industry),
      area: optionsOf_(data.rows, 'area', FALLBACK.area),
      status: orderedOptionsOf_(data.rows, 'status', FALLBACK.status),
      next: orderedOptionsOf_(data.rows, 'next', FALLBACK.next),
      reason: orderedOptionsOf_(data.rows, 'reason', FALLBACK.reason),
      method: orderedOptionsOf_(logs, 'method', FALLBACK.method),
      result: orderedOptionsOf_(logs, 'result', FALLBACK.result)
    },
    rows: data.rows,
    logs: logs
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
    if (p.action === 'log') return json_(addLog_(p.item));
    return json_({ ok: false, error: '不明な操作です' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function rankOf_(status) {
  var s = String(status || ''), best = 0;
  Object.keys(STATUS_RANK).forEach(function (k) {
    if (s.indexOf(k) >= 0) best = Math.max(best, STATUS_RANK[k]);
  });
  return best;
}

/** 結果からステータスを決める。進んだときだけ更新し、後戻りはさせない。 */
function statusFromResult_(result, current) {
  var mapped = RESULT_TO_STATUS[result];
  if (!mapped) return current || FALLBACK.status[0];
  if (mapped.indexOf('クローズ') >= 0 || mapped.indexOf('契約済') >= 0) return mapped;
  return rankOf_(mapped) > rankOf_(current) ? mapped : (current || mapped);
}

/**
 * ステータスに応じて日付を自動で入れる（すでに日付が入っていれば触らない）。
 * 成績は契約日、活動は見積提出日で集計するため、入力の手間を増やさずに日付を残す。
 */
function stampDates_(item, current) {
  current = current || {};
  var st = String(item.status || current.status || '');
  var t = today_();
  var has = function (key) {
    return (item[key] !== undefined && item[key] !== '') || (current[key] && String(current[key]) !== '');
  };
  if (st.indexOf('契約済') >= 0 && !has('contractDate')) item.contractDate = t;
  if ((st.indexOf('見積') >= 0 || st.indexOf('交渉') >= 0 || st.indexOf('契約') >= 0) && !has('quoteDate')) {
    item.quoteDate = t;
  }
  if (st.indexOf('クローズ') >= 0 && !has('closeDate')) item.closeDate = t;
  return item;
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
  if (item.quoteDate !== undefined) put(COL.quoteDate, item.quoteDate ? new Date(item.quoteDate) : '');
  if (item.contractDate !== undefined) put(COL.contractDate, item.contractDate ? new Date(item.contractDate) : '');
  if (item.closeDate !== undefined) put(COL.closeDate, item.closeDate ? new Date(item.closeDate) : '');
  if (item.lastAp !== undefined) put(COL.lastAp, item.lastAp ? new Date(item.lastAp) : '');
  if (item.apCount !== undefined) put(COL.apCount, item.apCount);
  put(COL.updated, new Date());
}

/** 活動ログに1行足す */
function writeLog_(caseId, item) {
  if (!item.method && !item.result) return;
  var sh = logSheet_();
  sh.appendRow([
    new Date(item.logDate || today_()), caseId, item.person || '', item.area || '',
    item.address || '', item.method || '', item.result || '', item.logMemo || '', new Date()
  ]);
}

/**
 * 実データが入っている最終行を返す。
 * getLastRow() は空白文字だけの行も「内容あり」と数えるため、
 * 全セルが空白（スペースのみを含む）の行はデータなしとみなす。
 */
function lastDataRow_(sh, lay) {
  var last = sh.getLastRow();
  if (last < lay.firstData) return lay.firstData - 1;
  var values = sh.getRange(lay.firstData, 1, last - lay.firstData + 1, sh.getLastColumn()).getValues();
  for (var i = values.length - 1; i >= 0; i--) {
    var filled = values[i].some(function (v) { return String(v == null ? '' : v).trim() !== ''; });
    if (filled) return lay.firstData + i;
  }
  return lay.firstData - 1;
}

/** 追加する行を決める。データのすぐ下に足す。 */
function appendRow_(sh, lay) {
  return Math.max(lastDataRow_(sh, lay) + 1, lay.firstData);
}

function create_(item) {
  var sh = sheet_();
  var lay = ensureColumns_(sh, layout_(sh));
  var row = appendRow_(sh, lay);
  var id = Utilities.getUuid();
  sh.getRange(row, lay.map[COL.id]).setValue(id);

  if (lay.map[COL.no]) {
    var maxNo = 0;
    var upto = Math.max(row - 1, lay.firstData);
    if (sh.getLastRow() >= lay.firstData) {
      sh.getRange(lay.firstData, lay.map[COL.no], upto - lay.firstData + 1, 1)
        .getValues().forEach(function (v) { maxNo = Math.max(maxNo, num_(v[0])); });
    }
    sh.getRange(row, lay.map[COL.no]).setValue(maxNo + 1);
  }

  if (!item.status && item.result) item.status = statusFromResult_(item.result, '');
  item.lastAp = item.logDate || item.date || today_();
  item.apCount = 1;
  stampDates_(item, null);
  setCells_(sh, lay, row, item);
  writeLog_(id, item);
  return { ok: true, id: id, row: row };
}

function findRow_(data, item) {
  var hit = null;
  data.rows.forEach(function (r) {
    if (item.id && r.id === item.id) hit = r;
    else if (!item.id && item.row && r.row === item.row) hit = r;
  });
  return hit;
}

function update_(item) {
  var data = readAll_();
  var hit = findRow_(data, item);
  if (!hit) return { ok: false, error: '対象の行が見つかりませんでした' };
  if (!hit.id) {
    var newId = Utilities.getUuid();
    data.sh.getRange(hit.row, data.lay.map[COL.id]).setValue(newId);
    hit.id = newId;
  }
  stampDates_(item, hit);
  setCells_(data.sh, data.lay, hit.row, item);
  return { ok: true, id: hit.id, row: hit.row };
}

/** 訪問を1件記録する。案件側のステータス・最終AP日・AP回数も更新する。 */
function addLog_(item) {
  var data = readAll_();
  var hit = findRow_(data, item);
  if (!hit) return { ok: false, error: '対象の案件が見つかりませんでした' };
  if (!hit.id) {
    var newId = Utilities.getUuid();
    data.sh.getRange(hit.row, data.lay.map[COL.id]).setValue(newId);
    hit.id = newId;
  }
  var patch = {
    status: item.status || statusFromResult_(item.result, hit.status),
    lastAp: item.logDate || today_(),
    apCount: (hit.apCount || 0) + 1
  };
  if (item.next) patch.next = item.next;
  if (item.logMemo) patch.memo = (hit.memo ? hit.memo + ' / ' : '') + item.logMemo;
  stampDates_(patch, hit);
  setCells_(data.sh, data.lay, hit.row, patch);
  writeLog_(hit.id, {
    logDate: item.logDate, person: item.person || hit.person, area: hit.area,
    address: hit.address, method: item.method, result: item.result, logMemo: item.logMemo
  });
  return { ok: true, id: hit.id, row: hit.row, status: patch.status };
}

/* ================= お手入れ用 ================= */

/**
 * データより下にある空っぽの行をまとめて削除する。
 * 見た目は空でもスペースなどが残っていると、新しい行がずっと下に追加されてしまうため。
 * データが入っている行は絶対に消しません。実行前に必ず件数を確認してください。
 */
function 整理_下の空行を削除() {
  var sh = sheet_();
  var lay = ensureColumns_(sh, layout_(sh));
  var lastData = lastDataRow_(sh, lay);
  var maxRows = sh.getMaxRows();
  Logger.log('データの最終行: ' + lastData + ' 行目');
  Logger.log('シートの行数: ' + maxRows + ' 行');
  if (maxRows <= lastData + 1) {
    Logger.log('削除する空行はありません。');
    return;
  }
  var from = lastData + 1;
  var count = maxRows - lastData - 1;   // 入力用に1行だけ残す
  sh.deleteRows(from, count);
  Logger.log(from + ' 行目から ' + count + ' 行を削除しました。');
  Logger.log('次に登録される行: ' + (lastData + 1) + ' 行目');
}

/* ================= 動作確認用 ================= */

function テスト_読み取り() {
  var d = readAll_();
  Logger.log('案件: ' + d.rows.length + ' 件');
  Logger.log('活動ログ: ' + readLogs_().length + ' 件');
  Logger.log(d.rows.slice(0, 3));
}
