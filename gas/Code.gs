/**
 * きづきの診断 — 回答受け取り用 Google Apps Script
 *
 * 使い方（詳しくは README.md）:
 *   1. 新しい Google スプレッドシートを作り、［拡張機能］→［Apps Script］を開く
 *   2. このファイルの中身をすべて貼り付けて保存
 *   3. ［デプロイ］→［新しいデプロイ］→ 種類「ウェブアプリ」
 *        実行ユーザー: 自分 / アクセスできるユーザー: 全員
 *   4. 表示された URL（…/exec）を config.js の sheetUrl に貼る
 *
 * 「回答」シートに1人1行で追記され、「集計」シートに回答者別の平均が自動で出ます。
 * Excel で使うときは［ファイル］→［ダウンロード］→［Microsoft Excel (.xlsx)］。
 */

const ANSWER_SHEET = '回答';
const SUMMARY_SHEET = '集計';
const ID_HEADER = '回答ID';
const ROLE_HEADER = '回答者';
const MAX_COLS = 80;
const MAX_CELL = 2000;

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const body = JSON.parse(e.postData.contents);
    const headers = body.headers, row = body.row, id = String(body.id || '');
    if (!Array.isArray(headers) || !Array.isArray(row) || headers.length !== row.length ||
        headers.length === 0 || headers.length > MAX_COLS || !/^[a-z0-9]{6,32}$/.test(id)) {
      return reply({ ok: false, error: 'bad_request' });
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(ANSWER_SHEET) || ss.insertSheet(ANSWER_SHEET, 0);

    // 見出し行：初回に作成し、新しい列があれば右に足す
    let cols = sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String) : [];
    if (cols.length === 0) {
      cols = headers.map(String);
      sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
      sh.setFrozenRows(1);
    } else {
      headers.forEach(h => {
        h = String(h);
        if (cols.indexOf(h) < 0 && cols.length < MAX_COLS) { cols.push(h); sh.getRange(1, cols.length).setValue(h).setFontWeight('bold'); }
      });
    }

    const out = cols.map(h => { const i = headers.indexOf(h); return i < 0 ? '' : clean(row[i]); });
    out[cols.indexOf(ID_HEADER)] = id;

    // 同じ回答IDの行があれば上書き（気づいたことの追記・予想一致数の更新）、なければ追記
    const idCol = cols.indexOf(ID_HEADER) + 1, last = sh.getLastRow();
    let target = 0;
    if (idCol > 0 && last > 1) {
      const ids = sh.getRange(2, idCol, last - 1, 1).getValues();
      for (let i = ids.length - 1; i >= 0; i--) if (String(ids[i][0]) === id) { target = i + 2; break; }
    }
    if (target) sh.getRange(target, 1, 1, out.length).setValues([out]);
    else sh.getRange(last + 1, 1, 1, out.length).setValues([out]);

    ensureSummary(ss, cols);
    return reply({ ok: true });
  } catch (err) {
    return reply({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return reply({ ok: true, service: 'kizuki' });
}

/** 数値はそのまま、文字列は長さを制限し、数式として解釈されないようにする */
function clean(v) {
  if (typeof v === 'number' && isFinite(v)) return v;
  if (v === null || v === undefined) return '';
  let s = String(v).slice(0, MAX_CELL);
  if (/^[=+\-@\t\r]/.test(s)) s = "'" + s;
  return s;
}

/** 「集計」シート：回答者ごとの件数と平均。見出し名で列を探すので列の並びが変わっても動く */
function ensureSummary(ss, cols) {
  if (ss.getSheetByName(SUMMARY_SHEET)) return;
  const sh = ss.insertSheet(SUMMARY_SHEET);
  const metrics = cols.filter(h => h === '柔軟度' || h === '自分を知る' || h === 'お手本・選択肢' ||
    h === '挑戦の機会' || h === '流されやすさ' || h === 'サポーター適性' || /^Q\d /.test(h) || h === '予想の一致数(子ども本人と)');
  const head = ['回答者', '件数'].concat(metrics);
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
  sh.setFrozenRows(1);
  const A = "'" + ANSWER_SHEET + "'";
  const colOf = ref => `INDEX(${A}!$A:$CZ,0,MATCH(${ref},${A}!$1:$1,0))`;
  const roles = ['親', '子ども', '大学生', '全体'];
  const rows = roles.map((r, ri) => {
    const rr = ri + 2;
    const count = r === '全体'
      ? `=MAX(0,COUNTA(${colOf('"' + ID_HEADER + '"')})-1)`
      : `=COUNTIF(${colOf('"' + ROLE_HEADER + '"')},$A${rr})`;
    const avgs = metrics.map((_, mi) => {
      const c = columnLetter(mi + 3) + '$1';
      return r === '全体'
        ? `=IFERROR(ROUND(AVERAGE(${colOf(c)}),1),"")`
        : `=IFERROR(ROUND(AVERAGEIFS(${colOf(c)},${colOf('"' + ROLE_HEADER + '"')},$A${rr}),1),"")`;
    });
    return [r, count].concat(avgs);
  });
  sh.getRange(2, 1, rows.length, head.length).setFormulas(rows.map(r => r.map(String)));
  sh.getRange(2, 1, rows.length, 1).setValues(roles.map(r => [r]));
  sh.getRange(7, 1).setValue('Q1〜Q8 は 1=ちがう〜4=そう の平均（素点）。スコアは 0〜100。「流されやすさ」は高いほど周囲に合わせる傾向。');
}

function columnLetter(n) {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
