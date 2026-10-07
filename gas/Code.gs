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
  const roles = ['子育て経験のある方', '子育て経験のない方（高校生以下）', '子育て経験のない方（大学生以上）', '全体'];
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

/* ======================================================================
 * 分析シート（グラフ付き）
 *   スプレッドシートを開くとメニュー「きづきの診断 → 分析シートを作る／作り直す」が出ます。
 *   すべて数式でできているので、回答が増えると表もグラフも自動で更新されます。
 *   B2 のプルダウンでグループ（?g= で指定したイベント名など）を絞り込めます。
 * ====================================================================== */

const ANALYSIS_SHEET = '分析';
const ROLES = ['子育て経験のある方', '子育て経験のない方（高校生以下）', '子育て経験のない方（大学生以上）'];
const ROLE_SHORT = ['子育て経験あり', '高校生以下', '大学生以上'];
const q = s => '"' + s + '"';
const ROLE_COLORS = ['#1F72B8', '#C9781E', '#5552B0'];
const TYPE_NAMES = ['自己分析派', 'ロールモデル派', '機会活用派', '空気読み派', 'これから探す派'];
const TYPE_COLORS = ['#1F72B8', '#C9781E', '#2E9A6A', '#5552B0', '#C2455A'];
const AXIS_ROWS = ['自分を知る', 'お手本・選択肢', '挑戦の機会', '流されやすさ', '柔軟度'];
// 回答シートの見出し（診断ページと同じ文言）と、グラフ用の短い名前
const QUESTIONS = [
  ['Q1 自分が好きなこと・夢中になっていることを聞かれたら、すぐに答える', 'Q1 好きなことをすぐ答える'],
  ['Q2 「こんな人になりたい」と思うお手本が身近にいる', 'Q2 お手本が身近にいる'],
  ['Q3 「周りがそうしているから」という理由で決めることが多い', 'Q3 周りに合わせて決める'],
  ['Q4 やってみたいことを、実際に試している', 'Q4 実際に試している'],
  ['Q5 自分が何をしたいのか、よく分からないことが多い', 'Q5 したいことが分からない'],
  ['Q6 身近な人の生き方は、どれも似たようなものばかりだ', 'Q6 周りの生き方が似ている'],
  ['Q7 周りと違っても、自分がやりたいことを選ぶ', 'Q7 違ってもやりたいことを選ぶ'],
  ['Q8 やってみたいことがあっても、時間や周りの目を理由にあきらめることが多い', 'Q8 あきらめることが多い']
];
const SCENE_COLS = ['場面1（学校に行きたくないとき）', '場面2（何かを選ぶとき）', '場面3（身近な人が「学校に行きたくない」と言ったら）'];
const SCENE_OPTS = ['A 自分で考えて決める', 'B 身近な人を参考にする', 'C まず試してみる', 'D 周りに合わせる'];
const FLEX_BINS = [[0, 19], [20, 39], [40, 59], [60, 79], [80, 100]];

function onOpen() {
  SpreadsheetApp.getUi().createMenu('きづきの診断')
    .addItem('分析シートを作る／作り直す', 'setupAnalysis')
    .addToUi();
}

function setupAnalysis() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(ANSWER_SHEET)) ss.insertSheet(ANSWER_SHEET, 0);
  const old = ss.getSheetByName(ANALYSIS_SHEET);
  if (old) ss.deleteSheet(old);
  const sh = ss.insertSheet(ANALYSIS_SHEET, 1);

  // 回答シートの列を見出し名で探す（列の並びが変わっても動く）
  const A = "'" + ANSWER_SHEET + "'";
  const col = h => `INDEX(${A}!$A$2:$CZ,0,MATCH(${h},${A}!$1:$1,0))`;
  const ROLE = col('"' + ROLE_HEADER + '"'), GROUP = col('"グループ"');
  // グループ絞り込み：B2 が「すべて」なら条件なし
  const ifs = (fn, valueRange, conds) => {
    const base = conds.map(c => c.join(',')).join(',');
    const head = fn === 'COUNTIFS' ? '' : valueRange + ',';
    return `IF($B$2="すべて",${fn}(${head}${base}),${fn}(${head}${base},${GROUP},$B$2))`;
  };
  const avg = (valueRange, conds) => `=IFERROR(ROUND(${ifs('AVERAGEIFS', valueRange, conds)},1),"")`;
  const cnt = conds => `=IFERROR(${ifs('COUNTIFS', '', conds)},0)`;

  sh.setHiddenGridlines(true);
  sh.setColumnWidth(1, 230);
  sh.setColumnWidths(2, 5, 92);
  sh.getRange('A1').setValue('きづきの診断 分析').setFontSize(16).setFontWeight('bold');
  sh.getRange('A2').setValue('グループ').setFontWeight('bold');
  sh.getRange('B2:C2').merge().setValue('すべて').setBackground('#EEF3F8');
  sh.getRange('A3').setValue('B2 のプルダウンでイベント・グループを絞り込めます。表とグラフは回答が増えると自動で更新されます。')
    .setFontColor('#5A6778').setFontSize(9);

  // 補助列（非表示）：グループ一覧
  sh.getRange('Z1').setValue('すべて');
  sh.getRange('Z2').setFormula(`=IFERROR(SORT(UNIQUE(FILTER(${GROUP},${GROUP}<>""))),"")`);
  sh.getRange('B2').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(sh.getRange('Z1:Z300'), true).setAllowInvalid(false).build());

  const blocks = [];
  const title = (r, text, note) => {
    sh.getRange(r, 1).setValue(text).setFontSize(12).setFontWeight('bold');
    if (note) sh.getRange(r, 2).setValue(note).setFontColor('#5A6778').setFontSize(9);
  };
  const header = (r, cells) => sh.getRange(r, 1, 1, cells.length).setValues([cells])
    .setFontWeight('bold').setBackground('#EEF3F8').setHorizontalAlignment('center');

  // 1. 回答者別の平均スコア
  let r = 5;
  title(r, '1. 回答者別の平均スコア', '0〜100。流されやすさは高いほど周囲に合わせる傾向');
  header(r + 1, ['項目'].concat(ROLE_SHORT));
  AXIS_ROWS.forEach((name, i) => {
    sh.getRange(r + 2 + i, 1).setValue(name);
    ROLES.forEach((_, j) => sh.getRange(r + 2 + i, 2 + j).setFormula(
      avg(col('$A' + (r + 2 + i)), [[ROLE, q(ROLES[j])]])));
  });
  sh.getRange(r + 7, 1).setValue('サポーター適性（大学生以上）');
  sh.getRange(r + 7, 4).setFormula(avg(col('"サポーター適性"'), [[ROLE, q(ROLES[2])]]));
  sh.getRange(r + 8, 1).setValue('回答数').setFontColor('#5A6778');
  ROLES.forEach((_, j) => sh.getRange(r + 8, 2 + j).setFormula(cnt([[ROLE, q(ROLES[j])]])).setFontColor('#5A6778'));
  blocks.push({ row: r, range: sh.getRange(r + 1, 1, 6, 4), type: Charts.ChartType.COLUMN, colors: ROLE_COLORS,
    title: '回答者別の平均スコア（0〜100）', axis: { v: [0, 100] }, height: 320 });

  // 2. 質問ごとの平均
  r = 21;
  title(r, '2. 質問ごとの平均', '1=ちがう 〜 4=そう（素点）');
  header(r + 1, ['質問'].concat(ROLE_SHORT));
  QUESTIONS.forEach(([full, short], i) => {
    sh.getRange(r + 2 + i, 1).setValue(short);
    sh.getRange(r + 2 + i, 28).setValue(full); // AB列（非表示）に元の見出し
    ROLES.forEach((_, j) => sh.getRange(r + 2 + i, 2 + j).setFormula(
      avg(col('$AB' + (r + 2 + i)), [[ROLE, q(ROLES[j])]])));
  });
  scale(sh.getRange(r + 2, 2, 8, 3));
  blocks.push({ row: r, range: sh.getRange(r + 1, 1, 9, 4), type: Charts.ChartType.BAR, colors: ROLE_COLORS,
    title: '質問ごとの平均（1=ちがう〜4=そう）', axis: { h: [1, 4] }, height: 420 });

  // 3. 子ども本人と大人の予想
  r = 41;
  title(r, '3. 高校生以下の答えと大人の予想', '高校生以下の方の平均と、子育て経験あり・大学生以上の方が予想した答えの平均');
  header(r + 1, ['質問', '高校生以下（本人）', '子育て経験ありの予想', '大学生以上の予想']);
  QUESTIONS.forEach(([full, short], i) => {
    const rr = r + 2 + i;
    sh.getRange(rr, 1).setValue(short);
    sh.getRange(rr, 28).setValue(full);
    sh.getRange(rr, 29).setValue('子どもの予想Q' + (i + 1)); // AC列（非表示）
    sh.getRange(rr, 2).setFormula(avg(col('$AB' + rr), [[ROLE, q(ROLES[1])]]));
    sh.getRange(rr, 3).setFormula(avg(col('$AC' + rr), [[ROLE, q(ROLES[0])]]));
    sh.getRange(rr, 4).setFormula(avg(col('$AC' + rr), [[ROLE, q(ROLES[2])]]));
  });
  scale(sh.getRange(r + 2, 2, 8, 3));
  sh.getRange(r + 10, 1).setValue('予想の一致数の平均（8問中）');
  sh.getRange(r + 10, 3).setFormula(avg(col('"予想の一致数(子ども本人と)"'), [[ROLE, q(ROLES[0])]]));
  sh.getRange(r + 10, 4).setFormula(avg(col('"予想の一致数(子ども本人と)"'), [[ROLE, q(ROLES[2])]]));
  sh.getRange(r + 11, 1).setValue('一致数は、同じ端末で高校生以下の方も答えた回のみ').setFontColor('#5A6778').setFontSize(9);
  blocks.push({ row: r, range: sh.getRange(r + 1, 1, 9, 4), type: Charts.ChartType.BAR, colors: ROLE_COLORS.slice(1, 2).concat([ROLE_COLORS[0], ROLE_COLORS[2]]),
    title: '高校生以下の方の答え と 大人の予想（1〜4）', axis: { h: [1, 4] }, height: 420 });

  // 4. タイプの分布
  r = 61;
  title(r, '4. タイプの分布（人数）', '高校生以下の方も大人向けのタイプ名で集計');
  header(r + 1, ['回答者'].concat(TYPE_NAMES));
  ROLES.forEach((role, i) => {
    sh.getRange(r + 2 + i, 1).setValue(ROLE_SHORT[i]);
    TYPE_NAMES.forEach((_, j) => sh.getRange(r + 2 + i, 2 + j).setFormula(
      cnt([[ROLE, q(role)], [col('"タイプ"'), columnLetter(2 + j) + '$' + (r + 1)]])));
  });
  blocks.push({ row: r, range: sh.getRange(r + 1, 1, 4, 6), type: Charts.ChartType.COLUMN, colors: TYPE_COLORS,
    title: 'タイプの割合（回答者別）', stacked: 'percent', height: 320 });

  // 5. 柔軟度の分布
  r = 77;
  title(r, '5. 柔軟度の分布（人数）', '柔軟度＝自分を知る・お手本・挑戦の機会・流されにくさの平均');
  header(r + 1, ['柔軟度'].concat(ROLE_SHORT));
  FLEX_BINS.forEach(([lo, hi], i) => {
    const rr = r + 2 + i;
    sh.getRange(rr, 1).setValue(lo + '〜' + hi);
    ROLES.forEach((_, j) => sh.getRange(rr, 2 + j).setFormula(cnt([
      [ROLE, q(ROLES[j])],
      [col('"柔軟度"'), '">=' + lo + '"'], [col('"柔軟度"'), '"<=' + hi + '"']])));
  });
  blocks.push({ row: r, range: sh.getRange(r + 1, 1, 6, 4), type: Charts.ChartType.COLUMN, colors: ROLE_COLORS,
    title: '柔軟度の分布（人数）', height: 320 });

  // 6. ふりかえりの選択（場面1〜3の合計）
  r = 93;
  title(r, '6. ふりかえりの選択（場面1〜3の合計）', '各場面でA〜Dのどれを選んだかの合計。聞き方は立場ごとに少し違う');
  header(r + 1, ['回答者'].concat(SCENE_OPTS));
  ROLES.forEach((role, i) => {
    sh.getRange(r + 2 + i, 1).setValue(ROLE_SHORT[i]);
    SCENE_OPTS.forEach((opt, j) => {
      const parts = SCENE_COLS.map(c => ifs('COUNTIFS', '', [[ROLE, q(role)], [col(q(c)), q(opt.charAt(0) + '*')]]));
      sh.getRange(r + 2 + i, 2 + j).setFormula('=IFERROR(' + parts.join('+') + ',0)');
    });
  });
  blocks.push({ row: r, range: sh.getRange(r + 1, 1, 4, 5), type: Charts.ChartType.COLUMN, colors: TYPE_COLORS.slice(0, 4),
    title: 'ふりかえりで選んだ行動の割合（回答者別）', stacked: 'percent', height: 320 });

  sh.getRange('B5:F110').setHorizontalAlignment('center').setNumberFormat('0.0');
  sh.getRange(95, 2, 3, 4).setNumberFormat('0');
  // 人数の表は整数表示
  sh.getRange(13, 2, 1, 3).setNumberFormat('0');
  sh.getRange(63, 2, 3, 5).setNumberFormat('0');
  sh.getRange(79, 2, 5, 3).setNumberFormat('0');

  // グラフ（表の右側、H列から）
  blocks.forEach(b => {
    let cb = sh.newChart().setChartType(b.type).addRange(b.range).setNumHeaders(1)
      .setPosition(b.row, 8, 0, 0)
      .setOption('title', b.title)
      .setOption('titleTextStyle', { fontSize: 13, bold: true, color: '#1D2938' })
      .setOption('colors', b.colors)
      .setOption('legend', { position: 'top', textStyle: { color: '#1D2938' } })
      .setOption('width', 640).setOption('height', b.height)
      .setOption('backgroundColor', '#FFFFFF');
    if (b.axis && b.axis.v) cb = cb.setOption('vAxis', { viewWindow: { min: b.axis.v[0], max: b.axis.v[1] }, gridlines: { color: '#E3EAF2' } });
    if (b.axis && b.axis.h) cb = cb.setOption('hAxis', { viewWindow: { min: b.axis.h[0], max: b.axis.h[1] }, gridlines: { color: '#E3EAF2' } });
    if (b.stacked) cb = cb.setOption('isStacked', b.stacked);
    sh.insertChart(cb.build());
  });

  blocks.forEach(b => sh.getRange(b.row, 2).setHorizontalAlignment('left'));
  sh.hideColumns(26, 4); // Z〜AC（補助列。Z はグループ一覧）
  sh.setFrozenRows(3);
  ss.setActiveSheet(sh);
}

function scale(range) {
  const sh = range.getSheet();
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpointWithValue('#FFFFFF', SpreadsheetApp.InterpolationType.NUMBER, '1')
    .setGradientMaxpointWithValue('#A9C9E8', SpreadsheetApp.InterpolationType.NUMBER, '4')
    .setRanges([range]).build();
  sh.setConditionalFormatRules(sh.getConditionalFormatRules().concat([rule]));
}
