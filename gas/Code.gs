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

    // 同じ回答IDの行があれば上書き（アンケート・気づいたことの追記・予想一致数の更新）、なければ追記
    const idCol = cols.indexOf(ID_HEADER) + 1, last = sh.getLastRow();
    let target = 0;
    if (idCol > 0 && last > 1) {
      const ids = sh.getRange(2, idCol, last - 1, 1).getValues();
      for (let i = ids.length - 1; i >= 0; i--) if (String(ids[i][0]) === id) { target = i + 2; break; }
    }
    // ページが送らない列（チームが入力する「自由記述の分類」など）は、上書きのときも今の値を残す
    const existing = target ? sh.getRange(target, 1, 1, cols.length).getValues()[0] : null;
    const out = cols.map((h, c) => { const i = headers.indexOf(h); return i < 0 ? (existing ? existing[c] : '') : clean(row[i]); });
    out[cols.indexOf(ID_HEADER)] = id;
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
    h === '挑戦の機会' || h === '流されやすさ' || h === 'サポーター適性' || /^Q\d /.test(h) || h === '予想の一致数(子ども本人と)' ||
    /^指標：/.test(h));
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
    .addItem('シートをまとめて作る（集計・回答・分析）', 'setupAll')
    .addItem('分析シートだけ作り直す', 'setupAnalysis')
    .addToUi();
}

/** 診断ページが送る列（index.html の HEADERS と同じ並び） */
const HEADERS = [
  "回答日時",
  "グループ",
  "セッションID",
  "回答ID",
  "回答者",
  "学年（高校生以下）",
  "タイプ",
  "サブタイプ",
  "柔軟度",
  "自分を知る",
  "お手本・選択肢",
  "挑戦の機会",
  "流されやすさ",
  "回答の一貫性",
  "場面1（学校に行きたくないとき）",
  "場面2（何かを選ぶとき）",
  "場面3（身近な人が「学校に行きたくない」と言ったら）",
  "Q1 自分が好きなこと・夢中になっていることを聞かれたら、すぐに答える",
  "Q2 「こんな人になりたい」と思うお手本が身近にいる",
  "Q3 「周りがそうしているから」という理由で決めることが多い",
  "Q4 やってみたいことを、実際に試している",
  "Q5 自分が何をしたいのか、よく分からないことが多い",
  "Q6 身近な人の生き方は、どれも似たようなものばかりだ",
  "Q7 周りと違っても、自分がやりたいことを選ぶ",
  "Q8 やってみたいことがあっても、時間や周りの目を理由にあきらめることが多い",
  "子どもの予想Q1",
  "子どもの予想Q2",
  "子どもの予想Q3",
  "子どもの予想Q4",
  "子どもの予想Q5",
  "子どもの予想Q6",
  "子どもの予想Q7",
  "子どもの予想Q8",
  "予想の一致数(子ども本人と)",
  "S1 年下の子の話を、口をはさまずに最後まで聞く",
  "S2 自分の経験は「正解」ではなく、ひとつの例として話す",
  "S3 学校に行かないなど、自分と違う選択をした子がいたら、その選択を尊重する",
  "S4 相手のためを思うと、つい「こうした方がいい」とアドバイスしたくなる",
  "サポーター適性",
  "アンケート回答",
  "B1 診断の結果は、自分に当てはまっていると思う",
  "B2 結果の説明は分かりやすかった",
  "C1 診断を受ける前から、自分がどうやって選んできたかを考えたことがあった",
  "C2 診断を受けて、自分の選びかたの傾向がはっきりした",
  "C3 これまで気づいていなかった自分の一面に気づいた",
  "C4 「周りに合わせて選んだ場面」や「お手本にした人」など、具体的な出来事を思い出した",
  "D1 診断の前、身近な子どもの好きなことや気持ちを分かっているつもりだった",
  "D2 予想と高校生以下の方の答えがずれた質問があり、意外に感じた（高校生以下の方が回答していない場合は、ずれそうだと感じた）",
  "D3 子どもの気持ちを、もっと聞いてみたいと思った",
  "D4 自分の考えや経験を、子どもに当てはめていた場面があったと気づいた",
  "E1 この結果について、子どもや身近な人と話してみたい",
  "E2 1週間以内に、子どもに好きなことや最近の気持ちを聞いてみようと思う",
  "E3 年下の子と関わる機会があれば、参加してみたい",
  "K1 けっかは、じぶんに当てはまっていると思う",
  "K2 じぶんのことで、あたらしくわかったことがあった",
  "K3 このけっかについて、おうちの人や大人と話してみたい",
  "K4 おうちの人に、じぶんの「すき」をもっと知ってほしいと思った",
  "指標：結果への納得感",
  "指標：自分への気づき",
  "指標：子どもへの理解",
  "指標：気づきの伸び(C2-C1)",
  "指標：思い込みへの気づき(1=あり)",
  "指標：行動意思",
  "気づいたこと",
  "F2 その気づきをきっかけに、やってみたいことがあれば書いてください",
  "F3 分かりにくかった質問や、改善してほしい点があれば書いてください"
];
/** チームが自由記述を読んで入力する列（ページからは送らない） */
const TEAM_HEADER = '自由記述の分類（0〜2・チーム入力）';

/**
 * 「集計」「回答」「分析」の3シートをこの順に作る。
 * 回答シートにすでにデータがあれば残し、見出しが足りない列だけ右に足す。集計・分析は作り直す。
 */
function setupAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ans = ss.getSheetByName(ANSWER_SHEET) || ss.insertSheet(ANSWER_SHEET);
  let cols = ans.getLastColumn() ? ans.getRange(1, 1, 1, ans.getLastColumn()).getValues()[0].map(String) : [];
  if (cols.length === 0) {
    cols = HEADERS.slice();
    ans.getRange(1, 1, 1, cols.length).setValues([cols]);
  } else {
    HEADERS.forEach(h => { if (cols.indexOf(h) < 0) { cols.push(h); ans.getRange(1, cols.length).setValue(h); } });
  }
  if (cols.indexOf(TEAM_HEADER) < 0) { cols.push(TEAM_HEADER); ans.getRange(1, cols.length).setValue(TEAM_HEADER); }
  ans.getRange(1, 1, 1, cols.length).setFontWeight('bold').setBackground('#EEF3F8');
  ans.getRange(1, cols.indexOf(TEAM_HEADER) + 1).setBackground('#FFF2B3');
  ans.getRange(2, cols.indexOf(TEAM_HEADER) + 1, Math.max(ans.getMaxRows() - 1, 1), 1).setBackground('#FFFBE6')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['0', '1', '2'], true).setAllowInvalid(true).build());
  ans.setFrozenRows(1);

  const oldSum = ss.getSheetByName(SUMMARY_SHEET);
  if (oldSum) ss.deleteSheet(oldSum);
  ensureSummary(ss, cols);
  setupAnalysis();

  // 並び順：集計 → 回答 → 分析
  [SUMMARY_SHEET, ANSWER_SHEET, ANALYSIS_SHEET].forEach((name, i) => {
    ss.setActiveSheet(ss.getSheetByName(name));
    ss.moveActiveSheet(i + 1);
  });
  // 何も入っていない初期シート（「シート1」など）は片づける
  ss.getSheets().forEach(sh => {
    const n = sh.getName();
    if ([SUMMARY_SHEET, ANSWER_SHEET, ANALYSIS_SHEET].indexOf(n) < 0 &&
        sh.getLastRow() === 0 && sh.getLastColumn() === 0 && sh.getCharts().length === 0) ss.deleteSheet(sh);
  });
  ss.setActiveSheet(ss.getSheetByName(SUMMARY_SHEET));
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

  // 7. 高校生以下の学年別
  r = 109;
  const GRADES = ['小学生', '中学生', '高校生'], GRADE = col('"学年（高校生以下）"');
  title(r, '7. 高校生以下の学年別の平均スコア', '0〜100。最初の質問で選んだ学年ごと');
  header(r + 1, ['項目'].concat(GRADES));
  AXIS_ROWS.forEach((name, i) => {
    sh.getRange(r + 2 + i, 1).setValue(name);
    GRADES.forEach((g, j) => sh.getRange(r + 2 + i, 2 + j).setFormula(
      avg(col('$A' + (r + 2 + i)), [[ROLE, q(ROLES[1])], [GRADE, q(g)]])));
  });
  sh.getRange(r + 7, 1).setValue('回答数').setFontColor('#5A6778');
  GRADES.forEach((g, j) => sh.getRange(r + 7, 2 + j).setFormula(cnt([[ROLE, q(ROLES[1])], [GRADE, q(g)]])).setFontColor('#5A6778'));
  blocks.push({ row: r, range: sh.getRange(r + 1, 1, 6, 4), type: Charts.ChartType.COLUMN, colors: ['#E0A04A', '#C9781E', '#8A4F12'],
    title: '高校生以下の学年別 平均スコア（0〜100）', axis: { v: [0, 100] }, height: 320 });

  // 8. アンケート：自分の現状への気づき・理解
  r = 125;
  const SURV = col('"アンケート回答"'), done = [[SURV, '"済"']];
  const byRole = j => j === 0 ? done : done.concat([[ROLE, q(ROLES[j - 1])]]);
  const share = (X, crit, conds) =>
    `=IFERROR(${ifs('COUNTIFS', '', conds.concat([[X, crit]]))}/${ifs('COUNTIFS', '', conds.concat([[X, '">=0"']]))},"")`;
  title(r, '8. アンケート：自分の現状への気づき・理解', '診断のあとのアンケートに答えた人のみ。スコアは0〜100（67≒「ややそう」）');
  header(r + 1, ['指標', '全体'].concat(ROLE_SHORT, ['目標', '判定']));
  const S8 = [
    ['結果への納得感', '指標：結果への納得感', 'avg', ''],
    ['自分への気づき', '指標：自分への気づき', 'avg', 66.6],
    ['子どもへの理解（大人のみ）', '指標：子どもへの理解', 'avg', 66.6],
    ['行動意思', '指標：行動意思', 'avg', 66.6],
    ['アンケート回答数', '', 'cnt', ''],
    ['自分への気づき 67以上の人の割合', '指標：自分への気づき', 'ge67', ''],
    ['気づきの伸び（C2−C1、−3〜+3）', '指標：気づきの伸び(C2-C1)', 'avg', ''],
    ['思い込みに気づいた人の割合（大人のみ）', '指標：思い込みへの気づき(1=あり)', 'eq1', 0.3],
    ['具体的な気づきがあった人の割合', TEAM_HEADER, 'team', 0.5]
  ];
  S8.forEach(([label, colName, kind, target], i) => {
    const rr = r + 2 + i;
    sh.getRange(rr, 1).setValue(label);
    for (let j = 0; j < 4; j++) {
      const X = colName ? col(q(colName)) : '', c = sh.getRange(rr, 2 + j);
      if (kind === 'avg') c.setFormula(avg(X, byRole(j)));
      if (kind === 'cnt') c.setFormula(cnt(byRole(j)));
      if (kind === 'ge67') c.setFormula(share(X, '">=66.6"', byRole(j)));
      if (kind === 'team') c.setFormula(share(X, '2', byRole(j)));
      if (kind === 'eq1') c.setFormula(share(X, '1', byRole(j)));
    }
    if (target !== '') {
      sh.getRange(rr, 6).setValue(target).setBackground('#FFF2B3');
      sh.getRange(rr, 7).setFormula(`=IF(ISNUMBER(B${rr}),IF(B${rr}>=F${rr},"達成","未達"),"−")`);
    }
    const fmt = kind === 'cnt' ? '0' : (kind === 'ge67' || kind === 'team' || kind === 'eq1') ? '0%' : i === 6 ? '0.00' : '0.0';
    sh.getRange(rr, 2, 1, 5).setNumberFormat(fmt);
  });
  sh.getRange(r + 2, 2, S8.length, 6).setHorizontalAlignment('center');
  sh.getRange(r + 11, 1).setValue('目標の数字は仮置きです。黄色のセルを書き換えると判定が変わります。「具体的な気づき」は回答シートの黄色の列（' + TEAM_HEADER + '）にチームが入力した分だけ数えます。')
    .setFontColor('#5A6778').setFontSize(9);
  blocks.push({ row: r, range: sh.getRange(r + 1, 1, 5, 5), type: Charts.ChartType.COLUMN, colors: ['#8A94A3'].concat(ROLE_COLORS),
    title: 'アンケート：気づき・理解のスコア（0〜100）', axis: { v: [0, 100] }, height: 320 });

  // 9. 予想の一致数と子どもへの理解
  r = 141;
  title(r, '9. 予想の一致数と子どもへの理解（大人のみ）', 'ずれが多かった人ほど理解のスコアが高ければ、すれ違いを見ることが気づきにつながっている');
  header(r + 1, ['予想の一致数', '人数', '子どもへの理解の平均']);
  const MATCH_COL = col('"予想の一致数(子ども本人と)"'), KIDU = col('"指標：子どもへの理解"');
  [['0〜4（ずれが多い）', '"<=4"'], ['5〜8（ずれが少ない）', '">=5"']].forEach(([label, crit], i) => {
    const rr = r + 2 + i, conds = done.concat([[MATCH_COL, crit]]);
    sh.getRange(rr, 1).setValue(label);
    sh.getRange(rr, 2).setFormula(cnt(conds)).setNumberFormat('0');
    sh.getRange(rr, 3).setFormula(avg(KIDU, conds)).setNumberFormat('0.0');
  });
  sh.getRange(r + 2, 2, 2, 2).setHorizontalAlignment('center');
  sh.getRange(r + 4, 1).setValue('同じ端末で高校生以下の方も答えた回だけが対象です。').setFontColor('#5A6778').setFontSize(9);

  // 自由記述の分類基準
  r = 147;
  title(r, '自由記述の分類基準（回答シートの黄色の列に 0・1・2 を入力）', '迷ったら低い方の数字にします。2人で別々に分類し、ずれたものを話し合うと判断がぶれにくくなります');
  header(r + 1, ['分類', '基準', '', '', '', '例']);
  [[0, '気づきが書かれていない。感想や評価だけ、または空欄。', '「楽しかった」「だいたい予想どおりだった」'],
   [1, '自分や子どもについての気づきはあるが、一般的・抽象的。', '「自分は流されやすいと思った」「子どものことを分かっていなかった」'],
   [2, '具体的な出来事・場面・人と結びついた気づき、または具体的な行動の予定がある。', '「中学の部活を友だちに合わせて選んだことを思い出した」「今週、最近夢中なことを聞いてみる」']
  ].forEach(([n, rule, ex], i) => {
    const rr = r + 2 + i;
    sh.getRange(rr, 1).setValue(n).setHorizontalAlignment('center');
    sh.getRange(rr, 2, 1, 4).merge().setValue(rule).setWrap(true);
    sh.getRange(rr, 6, 1, 2).merge().setValue(ex).setWrap(true);
  });

  sh.getRange('B5:F124').setHorizontalAlignment('center').setNumberFormat('0.0');
  sh.getRange(116, 2, 1, 3).setNumberFormat('0');
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

  blocks.concat([{ row: 141 }, { row: 147 }]).forEach(b => sh.getRange(b.row, 2).setHorizontalAlignment('left'));
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
