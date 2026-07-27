/**
 * JQuants_AccountingRisk 純ロジック検証（GAS不要 / 依存なし）
 *
 *   npm test        （= node tests/verify.js）
 *
 * src/Code.js と共有モジュールを GAS API のモック上へ読み込み、実際に動かす。
 * Sakata_Screener / Abitus-Automation / gas-shared と同じ骨格。
 */
const fs   = require('fs');
const path = require('path');

/* ── GAS モック ───────────────────────────────────────────────────────────── */

const logs = [];
let fetchPlan = [], fetchCalls = 0, slept = [];

const sandbox = {
  Logger: { log: m => logs.push(String(m)) },
  Utilities: { sleep: ms => slept.push(ms), formatDate: d => d.toISOString().slice(0, 10) },
  UrlFetchApp: {
    fetch: () => {
      const p = fetchPlan[fetchCalls++];
      if (p == null) throw new Error('想定より多く fetch が呼ばれた');
      if (p instanceof Error) throw p;
      return { getResponseCode: () => p.code, getContentText: () => p.body || '{}' };
    },
  },
  PropertiesService: {
    getScriptProperties: () => ({ getProperty: k => (k === 'JQUANTS_API_KEY' ? 'test-key' : null),
      setProperty: () => {}, deleteProperty: () => {} }),
  },
  SpreadsheetApp: { getActive: () => ({ toast: () => {} }) },
  ScriptApp: { getProjectTriggers: () => [], deleteTrigger: () => {} },
  DriveApp: {}, MailApp: {}, Session: {},
};

const src = f => fs.readFileSync(path.join(__dirname, '..', 'src', f), 'utf8');
const EXPORTS = ['JQ', 'F', 'num_', 'monthsSince_', 'rowFromStatement_', 'HEADER_STATEMENTS_',
  'riskLevel_', 'riskReasons_', 'riskComment_', 'fmt_', 'jqGet_', 'fetchWithRetry_'];

// 共通モジュール（symlink）も読み込む。本体が fetchWithRetry_ / to4_ を呼ぶため。
const M = new Function(...Object.keys(sandbox), `
${src('FetchRetry.js')}
${src('StockCode.js')}
${src('Code.js')}
return { ${EXPORTS.join(', ')} };
`)(...Object.values(sandbox));

/* ── アサーション ─────────────────────────────────────────────────────────── */

let pass = 0, fail = 0;
const eq = (a, b, label) => {
  const x = JSON.stringify(a), y = JSON.stringify(b);
  if (x === y) { pass++; console.log('  ✅ ' + label); }
  else { fail++; console.log('  ❌ ' + label + '\n     期待: ' + y + '\n     実際: ' + x); }
};
const throws = (fn, re, label) => {
  try { fn(); fail++; console.log('  ❌ ' + label + '（例外が出なかった）'); }
  catch (e) {
    if (re.test(e.message)) { pass++; console.log('  ✅ ' + label); }
    else { fail++; console.log('  ❌ ' + label + '\n     期待パターン: ' + re + '\n     実際: ' + e.message); }
  }
};

/* ── 1. 数値の解釈 ───────────────────────────────────────────────────────── */

console.log('\n【1】num_ — 財務値の数値化');
{
  eq(M.num_('1,234,567'), 1234567, 'カンマ区切りを数値にする');
  eq(M.num_('-500'), -500, '負の値');
  eq(M.num_(0), 0, '0 は 0（null にしない。売上0と未開示は別物）');
  eq(M.num_(''), null, '空文字は null');
  eq(M.num_(null), null, 'null は null');
  eq(M.num_(undefined), null, 'undefined は null');
  eq(M.num_('－'), null, '数値でない記号は null');
  eq(M.num_('1.5e3'), 1500, '指数表記も解釈する');
  eq(M.num_(Infinity), null, '無限大は採用しない');
}

/* ── 2. 開示の鮮度 ───────────────────────────────────────────────────────── */

console.log('\n【2】monthsSince_ — 開示日の古さ');
{
  eq(M.monthsSince_(''), Infinity, '空は Infinity（＝古い扱い）');
  eq(M.monthsSince_(null), Infinity, '未取得も Infinity');
  eq(M.monthsSince_('not-a-date'), Infinity, '解釈できない文字列も Infinity');
  const now = new Date();
  const iso = d => d.toISOString().slice(0, 10);
  eq(M.monthsSince_(iso(now)) < 0.1, true, '今日ならほぼ0か月');
  const old = new Date(now.getTime() - 365 * 24 * 3600 * 1000);
  const m = M.monthsSince_(iso(old));
  eq(m > 11.5 && m < 12.5, true, '1年前なら約12か月');
  eq(Math.abs(M.monthsSince_('2026/01/01') - M.monthsSince_('2026-01-01')) < 1e-9, true,
    'スラッシュ区切りもハイフン区切りと同じに扱う');
  // STALE_MONTHS=15 の境界。ここを跨ぐと「参考度低」になり出力から外れる
  const m16 = M.monthsSince_(iso(new Date(now.getTime() - 16 * 30.44 * 24 * 3600 * 1000)));
  eq(m16 > 15, true, '16か月前は 15か月の閾値を超える（＝出力対象から外れる）');
}

/* ── 3. 収集した決算の行化 ───────────────────────────────────────────────── */

console.log('\n【3】rowFromStatement_ — 財務データシートの1行');
{
  const H = M.HEADER_STATEMENTS_();
  const row = M.rowFromStatement_({
    [M.F.code]: '13010', [M.F.disclosed]: '2026-05-14', [M.F.periodType]: 'FY', [M.F.docType]: 'FYFinancialStatements',
    [M.F.netSales]: '1,000', [M.F.opProfit]: '100', [M.F.ordProfit]: '110', [M.F.profit]: '80',
    [M.F.totalAssets]: '5,000', [M.F.equity]: '2,000', [M.F.cfo]: '90', [M.F.eps]: '12.5',
  });
  eq(row.length, H.length, 'ヘッダーと列数が一致する（ずれると全列が1つずつ横にずれる）');
  eq(row[0], '1301', '5桁コードを4桁へ寄せる（旧データとの名寄せ）');
  eq(row.slice(1, 4), ['2026-05-14', 'FY', 'FYFinancialStatements'], '開示日・期種別・文書種別');
  eq(row.slice(4), [1000, 100, 110, 80, 5000, 2000, 90, 12.5], '財務値は数値化して入れる');
  eq(H.indexOf('営業CF'), 10, '営業CFの位置（computeRiskScores が列名で引くため）');

  const empty = M.rowFromStatement_({});
  eq(empty.length, H.length, '空の決算でも列数は保たれる');
  eq(empty.slice(4).every(v => v === null), true, '値が無い項目は null（0と区別する）');
}

/* ── 4. リスク区分と解説 ─────────────────────────────────────────────────── */

console.log('\n【4】riskLevel_ / riskReasons_ / riskComment_');
{
  const rec = o => Object.assign({ hasData: true, stale: false, risk: 50, accruals: null, flagCF: 0,
    opMarginChg: null, equityRatio: null, specialDep: null, disclosed: '2026-05-14' }, o);
  eq(M.riskLevel_(rec({ hasData: false })), 'データなし', '決算が無ければデータなし');
  eq(M.riskLevel_(rec({ risk: null })), 'データなし', 'スコアが出せなければデータなし');
  eq(M.riskLevel_(rec({ stale: true, risk: 90 })), '参考度低（データ古）', '開示が古ければスコアより鮮度を優先して示す');
  eq(M.riskLevel_(rec({ risk: 65 })), '【高リスク】', '65は高リスク（境界を含む）');
  eq(M.riskLevel_(rec({ risk: 64.9 })), '【中リスク】', '65未満は中リスク');
  eq(M.riskLevel_(rec({ risk: 52 })), '【中リスク】', '52は中リスク（境界を含む）');
  eq(M.riskLevel_(rec({ risk: 51.9 })), '【低リスク】', '52未満は低リスク');
}
{
  const rec = o => Object.assign({ hasData: true, stale: false, risk: 50, accruals: null, flagCF: 0,
    opMarginChg: null, equityRatio: null, specialDep: null, disclosed: '2026-05-14' }, o);
  eq(M.riskReasons_(rec({})), [], '該当が無ければ理由は空');
  eq(M.riskReasons_(rec({ accruals: 0.10 })).length, 1, 'アクルーアル0.10で1件（境界を含む）');
  eq(/営業CFの裏付けが弱い/.test(M.riskReasons_(rec({ accruals: 0.10 }))[0]), true, '強い方の文言になる');
  eq(/やや高め/.test(M.riskReasons_(rec({ accruals: 0.05 }))[0]), true, '0.05〜0.10 は「やや高め」');
  eq(M.riskReasons_(rec({ accruals: 0.049 })), [], '0.05未満は理由にしない');
  eq(M.riskReasons_(rec({ accruals: 0.2 })).length, 1, '強弱の文言は排他（二重計上しない）');
  eq(M.riskReasons_(rec({ flagCF: 1 })).length, 1, '黒字だが営業CFがマイナス');
  eq(M.riskReasons_(rec({ opMarginChg: -0.05 })).length, 1, '営業利益率の悪化（境界を含む）');
  eq(M.riskReasons_(rec({ opMarginChg: -0.049 })), [], '悪化が閾値未満なら挙げない');
  eq(M.riskReasons_(rec({ equityRatio: 0.199 })).length, 1, '自己資本比率20%未満');
  eq(M.riskReasons_(rec({ equityRatio: 0.20 })), [], '20%ちょうどは挙げない');
  eq(M.riskReasons_(rec({ specialDep: -0.031 })).length, 1, '特別損益の影響は絶対値で見る（負でも挙げる）');
  eq(M.riskReasons_(rec({ specialDep: 0.03 })), [], '0.03ちょうどは挙げない');
  eq(M.riskReasons_(rec({ accruals: 0.2, flagCF: 1, equityRatio: 0.1 })).length, 3, '複数該当は全部挙げる');
}
{
  const rec = o => Object.assign({ hasData: true, stale: false, risk: 50, accruals: null, flagCF: 0,
    opMarginChg: null, equityRatio: null, specialDep: null, disclosed: '2026-05-14' }, o);
  eq(M.riskComment_(rec({ hasData: false })), '決算(FY)データ未取得', 'データ無しの解説');
  eq(M.riskComment_(rec({})), '・目立った会計リスクの兆候は少ない', '該当が無いときの解説');
  eq(M.riskComment_(rec({ flagCF: 1 })).indexOf('・'), 0, '理由は箇条書きにする');
  eq(M.riskComment_(rec({ stale: true, flagCF: 1 })).split('\n')[0], '※最新開示が古く参考度は低い（2026-05-14）',
    '開示が古い場合は先頭に断り書きを入れる');
  eq(M.riskComment_(rec({ accruals: 0.2, flagCF: 1 })).split('\n').length, 2, '理由が複数なら改行で並べる');
}

/* ── 5. 表示の丸め ───────────────────────────────────────────────────────── */

console.log('\n【5】fmt_');
{
  eq(M.fmt_(null, 3), '', 'null は空欄（0と区別する）');
  eq(M.fmt_(0, 3), 0, '0 はそのまま0');
  eq(M.fmt_(0.123456, 3), 0.123, '小数3桁へ丸める');
  eq(M.fmt_(-0.123456, 3), -0.123, '負の値も丸める');
  eq(M.fmt_(1.5, 0), 2, '桁0なら整数へ');
}

/* ── 6. API取得（ページング・エラー・再試行） ────────────────────────────── */

console.log('\n【6】jqGet_ — ページングと失敗の扱い');
const runFetch = plan => { fetchPlan = plan; fetchCalls = 0; slept = []; logs.length = 0; };
{
  runFetch([{ code: 200, body: JSON.stringify({ data: [{ Code: '1301' }] }) }]);
  eq(M.jqGet_('/fins/summary', { code: '1301' }), [{ Code: '1301' }], '1ページなら1回の取得で返す');
  eq(fetchCalls, 1, '余計な取得をしない');
}
{
  runFetch([
    { code: 200, body: JSON.stringify({ data: [{ Code: '1' }], pagination_key: 'k1' }) },
    { code: 200, body: JSON.stringify({ data: [{ Code: '2' }], pagination_key: 'k2' }) },
    { code: 200, body: JSON.stringify({ data: [{ Code: '3' }] }) },
  ]);
  eq(M.jqGet_('/fins/summary').map(x => x.Code), ['1', '2', '3'], 'pagination_key を追って全ページを連結する');
  eq(fetchCalls, 3, 'ページ数だけ取得する');
}
{
  runFetch([{ code: 200, body: JSON.stringify({}) }]);
  eq(M.jqGet_('/equities/master'), [], 'data が無い応答は空配列（例外にしない）');
}
{
  runFetch([{ code: 401, body: '{"message":"unauthorized"}' }]);
  throws(() => M.jqGet_('/fins/summary'), /失敗\(401\)/, '401は再試行せず即エラー');
  eq(fetchCalls, 1, '恒久エラーで無駄に叩かない');
}
{
  // collectStatements は「失敗(401)/失敗(429)/失敗(50x)」の文言でキューへ戻すか判断している
  runFetch([{ code: 429, body: 'rate limit' }, { code: 429, body: 'rate limit' }, { code: 429, body: 'rate limit' }]);
  throws(() => M.jqGet_('/fins/summary'), /失敗\(429\)/, '再試行しきった429はメッセージに残る（キューへ戻す判定に使う）');
  eq([fetchCalls, slept], [3, [1500, 3000]], '429は設定どおり2回まで指数バックオフで再試行する');
}
{
  runFetch([{ code: 503, body: 'busy' }, { code: 200, body: JSON.stringify({ data: [{ Code: 'X' }] }) }]);
  eq(M.jqGet_('/fins/summary').length, 1, '5xxから回復すれば通常どおり返る');
  eq(slept, [1500], '1回だけ待つ');
}
{
  eq([M.JQ.FETCH_RETRY, M.JQ.FETCH_BACKOFF_MS], [2, 1500], '再試行の設定値');
  eq(typeof M.fetchWithRetry_, 'function', '共通モジュール FetchRetry.js が読み込まれている');
}

console.log('\n' + '─'.repeat(62));
console.log(fail === 0 ? `全 ${pass} 項目 合格` : `${pass} 合格 / ${fail} 失敗`);
process.exit(fail === 0 ? 0 : 1);
