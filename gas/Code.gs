/**
 * 入退室管理（iPad版）Google側のプログラム
 *
 * 使い方：
 *   1. スプレッドシートの「拡張機能 → Apps Script」にこのファイルの中身を貼り付ける
 *   2. 関数「setup」を実行する（シート・写真フォルダ・合言葉・毎日の自動処理を作る）
 *   3. ウェブアプリとしてデプロイする（実行ユーザー：自分／アクセス：全員）
 *   4. 関数「makeIpadLink」を実行し、表示されたリンクをiPadのSafariで開く
 */

const TZ = 'Asia/Tokyo';
const SHEET_EMP = '従業員';
const SHEET_LOG = '入退室ログ';
const SHEET_DAILY = '日別回数';
const SHEET_CHECK = '要確認';
const FOLDER_NAME = '入退室写真';
const KEEP_DAYS = 90;   // 写真の保存日数（これより古い写真は自動でゴミ箱へ）
const PAGE_URL = 'https://yamauchi917.github.io/factory-entry/';

const EMP_HEADER = ['社員ID', '氏名', 'ふりがな', '状態', '最終更新'];
const LOG_HEADER = ['日時', '社員ID', '氏名', '区分', '撮影方法', '写真', '備考', '受付ID', '受信日時'];

// ---------------------------------------------------------------- 初期設定

function setup() {
  const ss = SpreadsheetApp.getActive();
  const emp = getOrCreateSheet_(ss, SHEET_EMP, EMP_HEADER);
  const log = getOrCreateSheet_(ss, SHEET_LOG, LOG_HEADER);
  emp.getRange('E:E').setNumberFormat('yyyy/MM/dd HH:mm:ss');
  log.getRange('A:A').setNumberFormat('yyyy/MM/dd HH:mm:ss');
  log.getRange('I:I').setNumberFormat('yyyy/MM/dd HH:mm:ss');
  setupSummarySheets_(ss);

  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('TOKEN')) {
    props.setProperty('TOKEN', Utilities.getUuid().replace(/-/g, ''));
  }
  if (!props.getProperty('FOLDER_ID')) {
    props.setProperty('FOLDER_ID', DriveApp.createFolder(FOLDER_NAME).getId());
  }

  // 毎日の自動処理（同じものが二重にできないよう、一度消してから作る）
  ScriptApp.getProjectTriggers()
    .filter(t => ['nightlyReset', 'deleteOldPhotos'].includes(t.getHandlerFunction()))
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('nightlyReset').timeBased().everyDays(1).atHour(0).inTimezone(TZ).create();
  ScriptApp.newTrigger('deleteOldPhotos').timeBased().everyDays(1).atHour(2).inTimezone(TZ).create();

  Logger.log('準備ができました。');
  Logger.log('写真フォルダ：' + DriveApp.getFolderById(props.getProperty('FOLDER_ID')).getUrl());
  Logger.log('「従業員」シートに社員ID・氏名・ふりがなを入力してください（状態・最終更新は自動で入ります）。');
}

/** デプロイ後に実行：iPadの設定用リンクを表示する */
function makeIpadLink() {
  // 会社のアカウントだと …/a/macros/ドメイン/… の形になり、ログインしていないiPadからつながらないので直す
  const url = (ScriptApp.getService().getUrl() || '').replace(/\/a\/macros\/[^/]+\//, '/macros/');
  const token = PropertiesService.getScriptProperties().getProperty('TOKEN');
  if (!url || !token) {
    Logger.log('先に setup を実行し、ウェブアプリとしてデプロイしてください。');
    return;
  }
  Logger.log('ウェブアプリのURL：' + url);
  Logger.log('合言葉（トークン）：' + token);
  Logger.log('iPadの設定用リンク（このリンクは他の人に教えないでください）：');
  Logger.log(PAGE_URL + '#gas=' + encodeURIComponent(url) + '&token=' + token);
}

/**
 * 集計用のシート（ログから自動で計算されるので、手で入力しない）
 *  - 日別回数：日付・氏名ごとの入室回数と退室回数（深夜の自動退室は数えない）
 *  - 要確認　：押し忘れの可能性がある記録の一覧
 */
function setupSummarySheets_(ss) {
  const log = `'${SHEET_LOG}'`;
  const daily = ss.getSheetByName(SHEET_DAILY) || ss.insertSheet(SHEET_DAILY);
  daily.getRange('A1').setFormula(
    `=QUERY(${log}!A:E,"select toDate(A), C, count(B) where D <> '' and E <> '-' ` +
    `group by toDate(A), C pivot D order by toDate(A) desc ` +
    `label toDate(A) '日付', C '氏名', count(B) '' format toDate(A) 'yyyy/MM/dd'",1)`);
  daily.setFrozenRows(1);

  const check = ss.getSheetByName(SHEET_CHECK) || ss.insertSheet(SHEET_CHECK);
  check.getRange('A1').setFormula(
    `=QUERY(${log}!A:G,"select A, C, D, G where G <> '' order by A desc",1)`);
  check.getRange('A:A').setNumberFormat('yyyy/MM/dd HH:mm:ss');
  check.setFrozenRows(1);
}

function getOrCreateSheet_(ss, name, header) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(header);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, header.length).setFontWeight('bold');
  }
  return sh;
}

// ---------------------------------------------------------------- iPadからの受付

function doGet() {
  return ContentService.createTextOutput('入退室管理：動作中');
}

function doPost(e) {
  let req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, code: 'bad_request' });
  }
  const token = PropertiesService.getScriptProperties().getProperty('TOKEN');
  if (!token || req.token !== token) return json_({ ok: false, code: 'auth' });

  try {
    if (req.action === 'list') return json_({ ok: true, employees: listEmployees_() });
    if (req.action === 'punch') return json_(punch_(req));
    return json_({ ok: false, code: 'bad_request' });
  } catch (err) {
    return json_({ ok: false, code: 'server_error', error: String(err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** 従業員一覧（ふりがな順）と、いまの状態 */
function listEmployees_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_EMP);
  const rows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues() : [];
  return rows
    .filter(r => String(r[0]).trim() && String(r[1]).trim())
    .map(r => ({
      id: String(r[0]).trim(),
      name: String(r[1]).trim(),
      kana: String(r[2]).trim(),
      status: r[3] === '入室' ? 'in' : 'out',
    }))
    .sort((a, b) => (a.kana || a.name).localeCompare(b.kana || b.name, 'ja'));
}

/** 入室・退室を記録する */
function punch_(req) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActive();
    const emp = ss.getSheetByName(SHEET_EMP);
    const log = ss.getSheetByName(SHEET_LOG);

    // 通信が途切れて同じ記録が2回届いた場合は、2回目を無視する
    if (req.qid && isDuplicate_(log, req.qid)) return { ok: true, duplicate: true };

    const ids = emp.getLastRow() > 1 ? emp.getRange(2, 1, emp.getLastRow() - 1, 1).getValues() : [];
    const idx = ids.findIndex(r => String(r[0]).trim() === String(req.id));
    if (idx < 0) return { ok: false, code: 'unknown_employee' };
    const row = idx + 2;
    const name = String(emp.getRange(row, 2).getValue()).trim();
    const current = emp.getRange(row, 4).getValue();

    const type = req.type === 'in' ? '入室' : '退室';
    let time = new Date(req.time);
    if (isNaN(time.getTime())) time = new Date();

    const notes = [];
    if (req.note) notes.push(req.note);
    if (!req.note && type === '入室' && current === '入室') notes.push('退室記録なし');

    let photoUrl = '';
    if (req.photo) photoUrl = savePhoto_(req.photo, time, name, type);

    log.appendRow([time, req.id, name, type, req.method || '', photoUrl, notes.join('・'), req.qid || '', new Date()]);
    emp.getRange(row, 4, 1, 2).setValues([[type, time]]);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function isDuplicate_(log, qid) {
  const last = log.getLastRow();
  if (last < 2) return false;
  const n = Math.min(500, last - 1);
  return log.getRange(last - n + 1, 8, n, 1).getValues().some(r => r[0] === qid);
}

/** 写真をドライブの「入退室写真 / 年-月」フォルダに保存して、URLを返す */
function savePhoto_(dataUrl, time, name, type) {
  const root = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('FOLDER_ID'));
  const month = Utilities.formatDate(time, TZ, 'yyyy-MM');
  const it = root.getFoldersByName(month);
  const folder = it.hasNext() ? it.next() : root.createFolder(month);

  const base64 = String(dataUrl).replace(/^data:image\/\w+;base64,/, '');
  const fileName = Utilities.formatDate(time, TZ, 'yyyyMMdd_HHmmss') + '_' + name + '_' + type + '.jpg';
  const blob = Utilities.newBlob(Utilities.base64Decode(base64), 'image/jpeg', fileName);
  return folder.createFile(blob).getUrl();
}

// ---------------------------------------------------------------- 毎日の自動処理

/** 深夜0時：退室の打刻を忘れた人を「退室」に戻す */
function nightlyReset() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActive();
    const emp = ss.getSheetByName(SHEET_EMP);
    const log = ss.getSheetByName(SHEET_LOG);
    if (emp.getLastRow() < 2) return;
    const now = new Date();
    const rows = emp.getRange(2, 1, emp.getLastRow() - 1, 5).getValues();
    rows.forEach((r, i) => {
      if (r[3] !== '入室') return;
      log.appendRow([now, r[0], r[1], '退室', '-', '', '自動退室（退室の打刻忘れ）', '', now]);
      emp.getRange(i + 2, 4, 1, 2).setValues([['退室', now]]);
    });
  } finally {
    lock.releaseLock();
  }
}

/** 深夜2時：保存日数を過ぎた写真をゴミ箱へ移す */
function deleteOldPhotos() {
  const root = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('FOLDER_ID'));
  const limit = new Date(Date.now() - KEEP_DAYS * 24 * 60 * 60 * 1000);
  const folders = root.getFolders();
  while (folders.hasNext()) {
    const files = folders.next().getFiles();
    while (files.hasNext()) {
      const f = files.next();
      if (f.getDateCreated() < limit) f.setTrashed(true);
    }
  }
}
