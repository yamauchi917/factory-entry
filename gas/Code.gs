/**
 * 入退室管理（iPad版）Google側のプログラム
 *
 * 使い方：
 *   1. スプレッドシートの「拡張機能 → Apps Script」にこのファイルの中身を貼り付ける
 *   2. 関数「setup」を実行する（シート・写真フォルダ・合言葉・毎日の自動処理を作る）
 *   3. ウェブアプリとしてデプロイする（実行ユーザー：自分／アクセス：全員）
 *   4. 関数「makeIpadLink」を実行し、表示されたリンクをiPadのSafariで開く
 *
 * 健康観察記録（別のスプレッドシート HEALTH_SS_ID）にも書き込む（列は1行目の見出しの名前で探す）。
 *   - GASを実行するアカウントに、そのスプレッドシートの編集権限が必要
 *   - 書きかえたら setup をもう一度実行する（別のスプレッドシートを開く許可を求められるので「許可」する）→ 新バージョンでデプロイ
 */

const TZ = 'Asia/Tokyo';
const SHEET_EMP = '従業員';
const SHEET_LOG = '入退室ログ';
const SHEET_DAILY = '日別回数';
const SHEET_CHECK = '要確認';
const FOLDER_NAME = '入退室写真';
const KEEP_DAYS = 90;   // 写真の保存日数（これより古い写真は自動でゴミ箱へ）
const PAGE_URL = 'https://yamauchi917.github.io/factory-entry/';
// ウェブアプリのURL（「デプロイ」→「デプロイを管理」に出る、最後が /exec のもの）
const WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbwSgMaylQL7XtzeZLxoRSs16hQK9pqkI_6P0VulHGJWR_xkLjwoX6c1_dvcAM3W_uio/exec';

// 健康観察記録（保健所の監査で見せるため、入退室管理とは別のスプレッドシート）
const HEALTH_SS_ID = '1boZs4RvttupJeJrm412GAWFnNyvlBd0crlgdI0CdY4E';
const SHEET_HEALTH = '健康観察記録';
const SHEET_HEALTH_DAY = '日別確認';   // 工場長・栄養士が1日1回確認する（紙の票の確認印のかわり）
const SHEET_HEALTH_VIEW = '日別表示';  // 紙の票と同じ並び（人が横・項目が縦）で1日分を表示
const FEVER_TEMP = 37.5;               // この体温以上は「上司に報告」（iPad側の CFG.feverTemp と合わせる）
const HEALTH_MARKS = 18;               // ①〜⑬（本人）＋同居人①〜⑤
const HEALTH_RULES = [
  '★点検結果は○×で記入する。項目で異常があった場合は、速やかに上司に報告し指示を仰ぐこと。',
  '★化膿性疾患のある場合は調理作業に従事することを禁止する。',
  '◆点検結果で×の際は措置を記録する。',
];

const EMP_HEADER = ['社員ID', '氏名', 'ふりがな', '状態', '最終更新', '健康チェック日'];
const LOG_HEADER = ['日時', '社員ID', '氏名', '区分', '撮影方法', '写真', '備考', '受付ID', '受信日時'];

// ---------------------------------------------------------------- 初期設定

function setup() {
  const ss = SpreadsheetApp.getActive();
  const emp = getOrCreateSheet_(ss, SHEET_EMP, EMP_HEADER);
  const log = getOrCreateSheet_(ss, SHEET_LOG, LOG_HEADER);
  if (!emp.getRange('F1').getValue()) emp.getRange('F1').setValue(EMP_HEADER[5]).setFontWeight('bold');
  emp.getRange('E:F').setNumberFormat('yyyy/MM/dd HH:mm:ss');
  log.getRange('A:A').setNumberFormat('yyyy/MM/dd HH:mm:ss');
  log.getRange('I:I').setNumberFormat('yyyy/MM/dd HH:mm:ss');
  setupSummarySheets_(ss);
  setupHealth_();

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
  // エディタから実行すると、テスト用のURL（…/dev）や会社アカウント用のURL（…/a/ドメイン/…）が返ってくることがあり、
  // ログインしていないiPadからはつながらない。そのため「デプロイを管理」に出る …/exec のURLを WEB_APP_URL に書いておく
  const url = WEB_APP_URL;
  const token = PropertiesService.getScriptProperties().getProperty('TOKEN');
  if (!/^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/exec$/.test(url) || !token) {
    Logger.log('先に setup を実行し、「デプロイを管理」に出る …/exec のURLを、このファイルの上のほうの WEB_APP_URL に書いてください。');
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

/**
 * 健康観察記録のスプレッドシートを準備する
 *  - 健康観察記録：1行目の見出し（入場時刻・社員ID・氏名・体温・①〜⑬・同居人①〜⑤）で列を探すので、列の順番は自由。
 *                  社員IDの列がなければ氏名の右に、「措置」「備考」がなければ右端に足す
 *  - 日別確認　　：毎日0時に前日の行が増える。工場長・栄養士が確認したら名前を入れる
 *  - 日別表示　　：B2の日付を書きかえると、その日の記録が紙の票と同じ並びで表示される（印刷用）
 */
function setupHealth_() {
  const hs = SpreadsheetApp.openById(HEALTH_SS_ID);
  const sh = hs.getSheetByName(SHEET_HEALTH);
  if (!sh) throw new Error(`健康観察のスプレッドシートに「${SHEET_HEALTH}」タブが見つかりません。`);

  let c = healthCols_(sh);
  const missing = [[c.time, '入場時刻'], [c.name, '氏名']].filter(m => !m[0]).map(m => m[1]);
  if (c.marks.length !== HEALTH_MARKS) missing.push(`①〜⑬・同居人①〜⑤（${HEALTH_MARKS}個のところ ${c.marks.length}個）`);
  if (missing.length) {
    throw new Error(`「${SHEET_HEALTH}」の1行目に見つからない見出し：${missing.join('、')}。` +
      `／いまの1行目：${c.head.map((h, i) => `${colLetter_(i + 1)}「${h}」`).join(' ')}`);
  }
  // 体温の列がなければ氏名の右に、社員IDの列がなければ氏名の左に足す（入場時刻・社員ID・氏名・体温・①… の並び）
  if (!c.temp) {
    sh.insertColumnAfter(c.name);
    sh.getRange(1, c.name + 1).setValue('体温（℃）').setFontWeight('bold');
  }
  if (!c.id) {
    sh.insertColumnBefore(c.name);
    sh.getRange(1, c.name).setValue('社員ID').setFontWeight('bold');
  }
  ['措置', '備考'].forEach(h => {
    if (!healthCols_(sh)[h === '措置' ? 'measure' : 'note']) {
      sh.getRange(1, sh.getLastColumn() + 1).setValue(h).setFontWeight('bold');
    }
  });
  c = healthCols_(sh);
  const L = colLetter_;
  sh.setFrozenRows(1);
  sh.getRange(`${L(c.time)}:${L(c.time)}`).setNumberFormat('yyyy/MM/dd HH:mm');
  sh.getRange(`${L(c.temp)}:${L(c.temp)}`).setNumberFormat('0.0');
  sh.getRange(`${L(c.id)}:${L(c.id)}`).setNumberFormat('@');   // 社員IDの先頭の0が消えないよう文字として扱う

  // ×や発熱のある行を赤くする（同じルールが二重にならないよう、前に付けたものを外してから付ける）
  const t = `$${L(c.temp)}2`;
  const countX = runs_(c.marks).map(([a, b]) => `COUNTIF($${L(a)}2:$${L(b)}2,"×")`).join('+');
  const formula = `=OR(${countX}>0,AND(ISNUMBER(${t}),${t}>=${FEVER_TEMP}))`;
  const rules = sh.getConditionalFormatRules().filter(r => {
    const b = r.getBooleanCondition();
    return !(b && /COUNTIF\(.*"×"\)/.test(String(b.getCriteriaValues()[0])));
  });
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(formula).setBackground('#fde2e2')
    .setRanges([sh.getRange(`A2:${L(c.width)}`)]).build());
  sh.setConditionalFormatRules(rules);

  // 日別確認
  const day = getOrCreateSheet_(hs, SHEET_HEALTH_DAY, ['日付', '記録人数', '異常のあった人数', '工場長', '栄養士', '備考']);
  day.getRange('A:A').setNumberFormat('yyyy/MM/dd（ddd）');

  // 日別表示（紙の票と同じ並び）
  const view = hs.getSheetByName(SHEET_HEALTH_VIEW) || hs.insertSheet(SHEET_HEALTH_VIEW);
  view.clear();
  const src = `'${SHEET_HEALTH}'!`;
  const rng = n => `${src}${L(n)}2:${L(n)}`;
  const hit = `INT(${rng(c.time)})=$B$2`;
  const pick = n => `=IFERROR(TRANSPOSE(FILTER(${rng(n)},${hit})),"")`;
  const rows = [
    ['従事者個人別健康観察記録票'],
    ['日付（書きかえる）', '=TODAY()'],
    ['氏名', `=IFERROR(TRANSPOSE(FILTER(${rng(c.name)},${hit})),"この日の記録はありません")`],
    [c.head[c.temp - 1], pick(c.temp)],
    [c.head[c.time - 1], `=IFERROR(TRANSPOSE(ARRAYFORMULA(TEXT(FILTER(${rng(c.time)},${hit}),"H:mm"))),"")`],
  ];
  c.marks.forEach((n, i) => {
    if (i === 13) rows.push(['同居人について']);
    rows.push([c.head[n - 1], pick(n)]);
  });
  rows.push(['措置', pick(c.measure)]);
  rows.push([]);
  HEALTH_RULES.forEach(t => rows.push([t]));
  rows.forEach((r, i) => r.forEach((v, j) => {
    const cell = view.getRange(i + 1, j + 1);
    if (String(v).startsWith('=')) cell.setFormula(v); else cell.setValue(v);
  }));
  view.getRange('A1').setFontWeight('bold').setFontSize(14);
  view.getRange('B2').setNumberFormat('yyyy/MM/dd（ddd）').setBackground('#fff7d6');
  view.getRange('A3:A').setFontWeight('bold');
  view.setColumnWidth(1, 420);
  view.setFrozenColumns(1);
  view.setFrozenRows(3);
}

/** 健康観察記録の1行目の見出しから、各項目が何列目か（1始まり、ないときは0）を調べる */
function healthCols_(sh) {
  const width = Math.max(sh.getLastColumn(), 1);
  const head = sh.getRange(1, 1, 1, width).getValues()[0].map(v => String(v).trim());
  const find = test => head.findIndex(test) + 1;
  return {
    head, width,
    time: find(h => /入場|入室|時刻|日時/.test(h)),
    id: find(h => /社員\s*(ID|ＩＤ|番号)/i.test(h)),
    name: find(h => /氏名|名前/.test(h)),
    temp: find(h => h.includes('体温')),
    marks: head.map((h, i) => /^(同居人)?\s*[①-⑬]/.test(h) ? i + 1 : 0).filter(n => n),   // 本人①〜⑬ → 同居人①〜⑤ の順
    measure: find(h => h === '措置'),
    note: find(h => h === '備考'),
  };
}

/** 列番号 → 列の文字（1→A、27→AA） */
function colLetter_(n) {
  let s = '';
  for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + (n - 1) % 26) + s;
  return s;
}

/** [4,5,6,9,10] → [[4,6],[9,10]]（となり合う列をまとめる） */
function runs_(cols) {
  const out = [];
  cols.forEach(n => {
    const last = out[out.length - 1];
    if (last && last[1] === n - 1) last[1] = n; else out.push([n, n]);
  });
  return out;
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
  const rows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues() : [];
  return rows
    .filter(r => String(r[0]).trim() && String(r[1]).trim())
    .map(r => ({
      id: String(r[0]).trim(),
      name: String(r[1]).trim(),
      kana: String(r[2]).trim(),
      status: r[3] === '入室' ? 'in' : 'out',
      healthDate: r[5] instanceof Date ? Utilities.formatDate(r[5], TZ, 'yyyy-MM-dd') : '',   // 健康チェックをした日
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

    // その日最初の打刻では健康チェックの結果も届く → 健康観察記録へ書く
    const health = parseHealth_(req.health);
    if (health) {
      writeHealth_(time, req.id, name, health, type);
      emp.getRange(row, 6).setValue(time);
      if (health.abnormal) notes.push('健康チェックで異常あり');
    }

    let photoUrl = '';
    if (req.photo) photoUrl = savePhoto_(req.photo, time, name, type);

    log.appendRow([time, req.id, name, type, req.method || '', photoUrl, notes.join('・'), req.qid || '', new Date()]);
    emp.getRange(row, 4, 1, 2).setValues([[type, time]]);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/** iPadから届いた健康チェック {temp, marks:['○'|'×' ×18]} を確かめる（おかしければ null） */
function parseHealth_(h) {
  if (!h || !Array.isArray(h.marks) || h.marks.length !== HEALTH_MARKS) return null;
  if (!h.marks.every(m => m === '○' || m === '×')) return null;
  const temp = Number(h.temp);
  if (!(temp >= 34 && temp <= 42)) return null;
  return { temp, marks: h.marks, abnormal: h.marks.includes('×') || temp >= FEVER_TEMP };
}

function writeHealth_(time, id, name, health, type) {
  const sh = SpreadsheetApp.openById(HEALTH_SS_ID).getSheetByName(SHEET_HEALTH);
  const c = healthCols_(sh);
  if (c.marks.length !== HEALTH_MARKS) throw new Error(`「${SHEET_HEALTH}」の見出しが変わっています（setupを実行して確認してください）`);
  const row = Array(c.width).fill('');
  const put = (n, v) => { if (n) row[n - 1] = v; };
  put(c.time, time);
  put(c.id, String(id));
  put(c.name, name);
  put(c.temp, health.temp);
  c.marks.forEach((n, i) => put(n, health.marks[i]));
  put(c.note, type === '退室' ? '退室時に記録（入室の押し忘れ）' : '');
  sh.appendRow(row);
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
  try { addHealthDay_(new Date(Date.now() - 12 * 60 * 60 * 1000)); }
  catch (err) { console.error('日別確認の追加に失敗', err); }
}

/** 日別確認に、指定した日の行（記録人数・異常のあった人数）を足す。記録のない日は足さない */
function addHealthDay_(date) {
  const hs = SpreadsheetApp.openById(HEALTH_SS_ID);
  const sh = hs.getSheetByName(SHEET_HEALTH);
  const day = hs.getSheetByName(SHEET_HEALTH_DAY);
  if (!sh || !day || sh.getLastRow() < 2) return;
  const key = Utilities.formatDate(date, TZ, 'yyyy-MM-dd');
  const done = day.getLastRow() > 1 ? day.getRange(2, 1, day.getLastRow() - 1, 1).getValues() : [];
  if (done.some(r => r[0] instanceof Date && Utilities.formatDate(r[0], TZ, 'yyyy-MM-dd') === key)) return;

  const c = healthCols_(sh);
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, c.width).getValues()
    .filter(r => r[c.time - 1] instanceof Date && Utilities.formatDate(r[c.time - 1], TZ, 'yyyy-MM-dd') === key);
  if (!rows.length) return;
  const ng = rows.filter(r => c.marks.some(n => r[n - 1] === '×') || Number(r[c.temp - 1]) >= FEVER_TEMP).length;
  const [y, m, d] = key.split('-').map(Number);
  day.appendRow([new Date(y, m - 1, d), rows.length, ng, '', '', '']);
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
