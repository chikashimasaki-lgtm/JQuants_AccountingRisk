// ============================================================================
//  東証プライム 会計リスク・スクリーナー（J-Quants API）
//  ---------------------------------------------------------------------------
//  無料Freeプラン（約12週遅延）で取得できるサマリー財務から「利益の質」を評価し、
//  会計リスク（利益操作・急悪化）の兆候をランク付けする。投資助言ではない。
//
//  使い方（初回）:
//   1) スクリプトプロパティに JQUANTS_API_KEY を設定（J-Quants V2はAPIキー方式）
//      ダッシュボードで発行したAPIキーを x-api-key ヘッダーで送る。トークン交換は不要。
//   2) メニュー「会計リスク」→ セットアップ
//   3) ① プライム銘柄を取得 → ② 財務データを収集 → ③ リスクスコアを計算
// ============================================================================

const JQ = {
  BASE: 'https://api.jquants.com/v2',   // J-Quants V2
  SHEETS: {
    UNIVERSE:   '銘柄マスタ',        // プライム銘柄一覧
    STATEMENTS: '財務データ',        // 収集した決算（生データ）
    RANKING:    'リスクランキング',  // スコア計算結果
    USAGE:      '使い方',            // 説明シート
    WORKFLOW:   'ワークフロー',      // 処理フロー図
  },
  MARKET_NAME_PRIME: 'プライム',     // /equities/master の MktNm
  MARKET_CODE_PRIME: '0111',         // /equities/master の Mkt（0111=プライム）
  TIME_BUDGET_MS:   4.5 * 60 * 1000, // 1回の実行で使う時間上限（GAS 6分制限対策）
};

// J-Quants V2 /fins/summary の項目名（短縮化されている）
const F = {
  code:       'Code',
  disclosed:  'DiscDate',
  periodType: 'CurPerType',   // 'FY' | '1Q' | '2Q' | '3Q'
  docType:    'DocType',
  netSales:   'Sales',
  opProfit:   'OP',
  ordProfit:  'OdP',
  profit:     'NP',
  totalAssets:'TA',
  equity:     'Eq',
  cfo:        'CFO',
  eps:        'EPS',
};

// ============================================================================
//  メニュー
// ============================================================================

function onOpen() {
  SpreadsheetApp.getUi().createMenu('会計リスク')
    .addItem('セットアップ（シート作成）', 'setup')
    .addSeparator()
    .addItem('① プライム銘柄を取得',        'fetchPrimeUniverse')
    .addItem('② 財務データを収集/続行',      'collectStatements')
    .addItem('③ リスクスコアを計算',        'computeRiskScores')
    .addSeparator()
    .addItem('JSON出力（Pages用）',          'exportJson')
    .addItem('収集の進捗リセット',           'resetCollectQueue')
    .addSeparator()
    .addItem('使い方シートを作成/更新',      'createUsageSheet')
    .addItem('ワークフロー図を作成/更新',    'refreshWorkflowDiagram')
    .addToUi();
}

function setup() {
  const ss = SpreadsheetApp.getActive();
  [JQ.SHEETS.UNIVERSE, JQ.SHEETS.STATEMENTS, JQ.SHEETS.RANKING]
    .forEach(name => { if (!ss.getSheetByName(name)) ss.insertSheet(name); });
  createUsageSheet();
  writeWorkflowDiagram_();
  ss.toast('シート一式（使い方・ワークフロー含む）を準備しました', '会計リスク', 5);
}

// ============================================================================
//  認証（V2: APIキー方式。x-api-key ヘッダー。トークン交換は廃止）
// ============================================================================

function getApiKey_() {
  const key = PropertiesService.getScriptProperties().getProperty('JQUANTS_API_KEY');
  if (!key) throw new Error('JQUANTS_API_KEY をスクリプトプロパティに設定してください（J-Quants V2はAPIキー方式）');
  return key;
}

// GET（pagination_key 自動追従）。戻り値はデータ配列（V2は "data" キー）。
function jqGet_(path, params) {
  const apiKey = getApiKey_();
  const base = JQ.BASE + path;
  const q = params
    ? Object.keys(params).filter(k => params[k] != null && params[k] !== '')
        .map(k => k + '=' + encodeURIComponent(params[k])).join('&')
    : '';
  let url = q ? base + '?' + q : base;

  const out = [];
  let pagination = null;
  do {
    const u = pagination ? url + (url.includes('?') ? '&' : '?') + 'pagination_key=' + encodeURIComponent(pagination) : url;
    const res  = UrlFetchApp.fetch(u, { headers: { 'x-api-key': apiKey }, muteHttpExceptions: true });
    const code = res.getResponseCode();
    if (code !== 200) throw new Error('GET ' + path + ' 失敗(' + code + '): ' + res.getContentText().slice(0, 300));
    const json = JSON.parse(res.getContentText());
    if (Array.isArray(json.data)) out.push.apply(out, json.data);
    pagination = json.pagination_key || null;
  } while (pagination);
  return out;
}

// ============================================================================
//  ① プライム銘柄マスタ
// ============================================================================

function fetchPrimeUniverse() {
  const info  = jqGet_('/equities/master');
  const prime = info.filter(x => x.Mkt === JQ.MARKET_CODE_PRIME || x.MktNm === JQ.MARKET_NAME_PRIME);

  const rows = prime.map(x => [
    to4_(x.Code),
    x.CoName || '',
    x.S17Nm || x.S33Nm || '',
    x.MktNm || '',
  ]);

  const sh = SpreadsheetApp.getActive().getSheetByName(JQ.SHEETS.UNIVERSE);
  sh.clearContents();
  sh.getRange(1, 1, 1, 4).setValues([['コード', '企業名', '業種', '市場']]);
  if (rows.length) sh.getRange(2, 1, rows.length, 4).setValues(rows);
  sh.autoResizeColumns(1, 4);  // 列幅をデータ内容に合わせて自動調整
  Logger.log('プライム銘柄: ' + rows.length + '件 / 全上場 ' + info.length + '件');
  SpreadsheetApp.getActive().toast('プライム ' + rows.length + '件を取得', '会計リスク', 5);
}

// ============================================================================
//  ② 財務データ収集（銘柄コードごとに全履歴を取得・時間内で分割実行）
// ============================================================================

function collectStatements() {
  const ss  = SpreadsheetApp.getActive();
  const uni = ss.getSheetByName(JQ.SHEETS.UNIVERSE);
  const st  = ss.getSheetByName(JQ.SHEETS.STATEMENTS);
  if (!uni || uni.getLastRow() < 2) throw new Error('先に「① プライム銘柄を取得」を実行してください');

  const props = PropertiesService.getScriptProperties();
  let queue = JSON.parse(props.getProperty('JQ_COLLECT_QUEUE') || 'null');
  if (!queue) {
    queue = uni.getRange(2, 1, uni.getLastRow() - 1, 1).getValues().flat().map(String).filter(Boolean);
    // ヘッダを（無ければ）用意
    if (st.getLastRow() === 0) st.appendRow(HEADER_STATEMENTS_());
  }

  // 既存キー（重複防止）
  const seen = new Set();
  if (st.getLastRow() > 1) {
    st.getRange(2, 1, st.getLastRow() - 1, 3).getValues()
      .forEach(r => seen.add(r[0] + '|' + r[1] + '|' + r[2]));  // code|disclosed|periodType
  }

  const start = Date.now();
  const buffer = [];
  let processed = 0;

  while (queue.length > 0) {
    if (Date.now() - start > JQ.TIME_BUDGET_MS) break;
    const code = queue.shift();
    try {
      const list = jqGet_('/fins/summary', { code: code });
      list.forEach(s => {
        const key = s[F.code] + '|' + s[F.disclosed] + '|' + s[F.periodType];
        if (seen.has(key)) return;
        seen.add(key);
        buffer.push(rowFromStatement_(s));
      });
      processed++;
    } catch (e) {
      Logger.log('収集エラー(' + code + '): ' + e.message);
      // 認証切れ等の一時失敗はキューに戻して次回再試行
      if (/失敗\(401\)|失敗\(429\)|失敗\(50/.test(e.message)) { queue.unshift(code); break; }
    }
    Utilities.sleep(150);
  }

  if (buffer.length) st.getRange(st.getLastRow() + 1, 1, buffer.length, buffer[0].length).setValues(buffer);

  // 続きがあれば保存して自動再開トリガーを張る
  clearResumeTriggers_();
  if (queue.length > 0) {
    props.setProperty('JQ_COLLECT_QUEUE', JSON.stringify(queue));
    ScriptApp.newTrigger('collectStatements').timeBased().after(90 * 1000).create();
    Logger.log('一時停止: ' + processed + '銘柄処理 / 残り ' + queue.length + '銘柄。90秒後に自動再開。');
    SpreadsheetApp.getActive().toast('残り ' + queue.length + '銘柄。自動再開します', '会計リスク', 5);
  } else {
    props.deleteProperty('JQ_COLLECT_QUEUE');
    if (st.getLastColumn() > 0) st.autoResizeColumns(1, st.getLastColumn());  // 列幅を自動調整
    Logger.log('収集完了');
    SpreadsheetApp.getActive().toast('財務データ収集が完了しました', '会計リスク', 5);
  }
}

function resetCollectQueue() {
  PropertiesService.getScriptProperties().deleteProperty('JQ_COLLECT_QUEUE');
  clearResumeTriggers_();
  SpreadsheetApp.getActive().toast('収集の進捗をリセットしました', '会計リスク', 5);
}

function clearResumeTriggers_() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'collectStatements')
    .forEach(t => ScriptApp.deleteTrigger(t));
}

function HEADER_STATEMENTS_() {
  return ['コード', '開示日', '期種別', '文書種別', '売上高', '営業利益', '経常利益',
          '当期純利益', '総資産', '純資産', '営業CF', 'EPS'];
}

function rowFromStatement_(s) {
  return [
    to4_(s[F.code] || ''), s[F.disclosed] || '', s[F.periodType] || '', s[F.docType] || '',
    num_(s[F.netSales]), num_(s[F.opProfit]), num_(s[F.ordProfit]), num_(s[F.profit]),
    num_(s[F.totalAssets]), num_(s[F.equity]), num_(s[F.cfo]), num_(s[F.eps]),
  ];
}

// 文字列/空文字を数値化（空・非数は null）
function num_(v) {
  if (v === '' || v == null) return null;
  const n = Number(String(v).replace(/,/g, ''));
  return isFinite(n) ? n : null;
}

// J-Quantsの5桁コード（4桁ティッカー+末尾0）を、見慣れた4桁に正規化する。
// 例: "72030" → "7203" / "130A0" → "130A"。5桁で末尾0のときのみ落とす。
function to4_(code) {
  const c = String(code == null ? '' : code);
  return (c.length === 5 && c.slice(-1) === '0') ? c.slice(0, 4) : c;
}

// ============================================================================
//  ③ リスクスコア計算（通期FYを対象。最新FYと前期FYを比較）
// ============================================================================

function computeRiskScores() {
  const ss = SpreadsheetApp.getActive();
  const st = ss.getSheetByName(JQ.SHEETS.STATEMENTS);
  if (!st || st.getLastRow() < 2) throw new Error('先に「② 財務データを収集」を実行してください');

  const uni  = ss.getSheetByName(JQ.SHEETS.UNIVERSE);
  const meta = new Map();
  if (uni && uni.getLastRow() > 1) {
    uni.getRange(2, 1, uni.getLastRow() - 1, 3).getValues()
      .forEach(r => meta.set(String(r[0]), { name: r[1], sector: r[2] }));
  }

  // コードごとに FY 決算を開示日昇順で並べる
  const H = HEADER_STATEMENTS_();
  const idx = Object.fromEntries(H.map((h, i) => [h, i]));
  const byCode = new Map();
  st.getRange(2, 1, st.getLastRow() - 1, H.length).getValues().forEach(r => {
    if (r[idx['期種別']] !== 'FY') return;   // 通期のみ
    const code = String(r[idx['コード']]);
    (byCode.get(code) || byCode.set(code, []).get(code)).push(r);
  });

  const recs = [];
  byCode.forEach((rows, code) => {
    rows.sort((a, b) => String(a[idx['開示日']]).localeCompare(String(b[idx['開示日']])));
    const cur  = rows[rows.length - 1];
    const prev = rows.length >= 2 ? rows[rows.length - 2] : null;

    const g = (r, k) => r ? r[idx[k]] : null;
    const profit = g(cur, '当期純利益'), cfo = g(cur, '営業CF'), assets = g(cur, '総資産');
    const sales  = g(cur, '売上高'),     op  = g(cur, '営業利益'), ord = g(cur, '経常利益'), eq = g(cur, '純資産');

    // 主指標: アクルーアル比率 (純利益 − 営業CF) / 総資産
    const accruals = (profit != null && cfo != null && assets) ? (profit - cfo) / assets : null;

    // 補助フラグ
    const flagCF   = (profit != null && cfo != null && profit > 0 && cfo < 0) ? 1 : 0; // 黒字なのに営業CFマイナス
    const opMargin = (op != null && sales) ? op / sales : null;
    const opMarginPrev = (g(prev, '営業利益') != null && g(prev, '売上高')) ? g(prev, '営業利益') / g(prev, '売上高') : null;
    const opMarginChg  = (opMargin != null && opMarginPrev != null) ? opMargin - opMarginPrev : null;
    const equityRatio  = (eq != null && assets) ? eq / assets : null;
    const specialDep   = (ord != null && profit != null && sales) ? (ord - profit) / sales : null; // 特別損益依存

    recs.push({ code, name: (meta.get(code) || {}).name || '', sector: (meta.get(code) || {}).sector || '',
      disclosed: cur[idx['開示日']], accruals, flagCF, opMarginChg, equityRatio, specialDep });
  });

  // アクルーアルを母集団で偏差値化（高いほど利益の質が低い＝リスク）
  const vals = recs.map(r => r.accruals).filter(v => v != null);
  const mean = vals.reduce((a, b) => a + b, 0) / (vals.length || 1);
  const sd   = Math.sqrt(vals.reduce((a, b) => a + (b - mean) ** 2, 0) / (vals.length || 1)) || 1;

  recs.forEach(r => {
    const dev = r.accruals != null ? 50 + 10 * ((r.accruals - mean) / sd) : null; // 会計リスク偏差値
    // フラグで加点（説明可能な範囲で軽く）
    let bonus = 0;
    if (r.flagCF) bonus += 8;
    if (r.opMarginChg != null && r.opMarginChg < -0.05) bonus += 4;   // 営業利益率5pt超悪化
    if (r.equityRatio != null && r.equityRatio < 0.2)   bonus += 3;   // 自己資本比率20%未満
    r.risk = dev != null ? Math.round((dev + bonus) * 10) / 10 : (r.flagCF ? 60 + bonus : null);
  });

  recs.sort((a, b) => (b.risk ?? -1) - (a.risk ?? -1));

  const out = recs.map(r => [
    r.code, r.name, r.sector, r.risk,
    fmt_(r.accruals, 3), r.flagCF ? '⚠' : '',
    fmt_(r.opMarginChg, 3), fmt_(r.equityRatio, 3), fmt_(r.specialDep, 3), r.disclosed,
  ]);

  const rank = ss.getSheetByName(JQ.SHEETS.RANKING);
  rank.clearContents();
  rank.getRange(1, 1, 1, 10).setValues([[
    'コード', '企業名', '業種', '会計リスク',
    'アクルーアル', '黒字CF-', '営業益率変化', '自己資本比率', '特別損益依存', '最新開示日']]);
  if (out.length) rank.getRange(2, 1, out.length, 10).setValues(out);
  rank.setFrozenRows(1);
  rank.autoResizeColumns(1, 10);  // 列幅をデータ内容に合わせて自動調整
  Logger.log('リスク計算完了: ' + out.length + '社');
  SpreadsheetApp.getActive().toast(out.length + '社をランク付けしました', '会計リスク', 5);
}

function fmt_(v, d) { return v == null ? '' : Math.round(v * 10 ** d) / 10 ** d; }

// ============================================================================
//  JSON出力（GitHub Pages 表示用）
// ============================================================================

function exportJson() {
  const ss   = SpreadsheetApp.getActive();
  const rank = ss.getSheetByName(JQ.SHEETS.RANKING);
  if (!rank || rank.getLastRow() < 2) throw new Error('先に「③ リスクスコアを計算」を実行してください');

  const header = rank.getRange(1, 1, 1, rank.getLastColumn()).getValues()[0];
  const keys   = ['code', 'name', 'sector', 'risk', 'accruals', 'cfFlag', 'opMarginChg', 'equityRatio', 'specialDep', 'disclosed'];
  const data   = rank.getRange(2, 1, rank.getLastRow() - 1, header.length).getValues()
    .map(r => Object.fromEntries(keys.map((k, i) => [k, r[i]])));

  const json = JSON.stringify({ updated: new Date().toISOString(), market: 'プライム', items: data });
  const file = DriveApp.createFile('accounting_risk_prime.json', json, 'application/json');
  Logger.log('JSON出力: ' + file.getUrl());
  SpreadsheetApp.getActive().toast('JSONをDriveに出力しました', '会計リスク', 5);
  return file.getUrl();
}

// ============================================================================
//  使い方シート
// ============================================================================

function createUsageSheet() {
  const ss = SpreadsheetApp.getActive();
  const old = ss.getSheetByName(JQ.SHEETS.USAGE);
  if (old) ss.deleteSheet(old);
  const sh = ss.insertSheet(JQ.SHEETS.USAGE, 0);
  sh.setHiddenGridlines(true);
  sh.setColumnWidth(1, 760);

  // [テキスト, 種別]  種別: title / h(見出し) / p(本文) / code / note
  const rows = [
    ['会計リスク・スクリーナー　使い方', 'title'],
    ['', 'p'],
    ['■ これは何？', 'h'],
    ['J-Quants API（無料・約12週遅延）の決算データから「利益の質」を評価し、会計リスク（利益操作・急悪化）の兆候がある東証プライム銘柄をランク付けするツール。投資助言ではありません。', 'p'],
    ['', 'p'],
    ['■ 事前準備（初回のみ）', 'h'],
    ['1. J-Quants（https://jpx-jquants.com/）に無料登録', 'p'],
    ['2. ダッシュボードで APIキー を発行（V2はAPIキー方式・有効期限なし）', 'p'],
    ['3. スクリプトエディタ → ⚙ プロジェクトの設定 → スクリプトプロパティ に登録:', 'p'],
    ['JQUANTS_API_KEY = 発行したAPIキー', 'code'],
    ['', 'p'],
    ['■ 使い方（上部メニュー「会計リスク」）', 'h'],
    ['① プライム銘柄を取得 … /equities/master からプライム銘柄を「銘柄マスタ」へ', 'p'],
    ['② 財務データを収集/続行 … 各銘柄の決算を /fins/summary から「財務データ」へ（時間分割で自動再開）', 'p'],
    ['③ リスクスコアを計算 … 通期(FY)決算から会計リスク偏差値を算出し「リスクランキング」へ', 'p'],
    ['JSON出力 … 公開ページ用 accounting_risk_prime.json を Drive に出力', 'p'],
    ['', 'p'],
    ['■ シートの説明', 'h'],
    ['銘柄マスタ … 対象のプライム銘柄一覧（コード・企業名・業種・市場）', 'p'],
    ['財務データ … 収集した決算の生データ（重複防止つき）', 'p'],
    ['リスクランキング … スコア計算結果。会計リスクが高い順', 'p'],
    ['', 'p'],
    ['■ 指標の見方', 'h'],
    ['アクルーアル比率 =(当期純利益−営業CF)/総資産。高いほど利益が現金の裏付けを欠く＝利益操作の代表的赤信号（主指標）', 'p'],
    ['黒字CF- … 黒字なのに営業CFがマイナス（⚠）', 'p'],
    ['営業益率変化 … 前期比の営業利益率の増減', 'p'],
    ['自己資本比率 … 純資産 / 総資産', 'p'],
    ['特別損益依存 =(経常利益−純利益)/売上。特別項目での利益調整の度合い', 'p'],
    ['', 'p'],
    ['■ 注意', 'h'],
    ['・無料プランは約12週遅延（決算分析用途のため実用上問題なし）', 'note'],
    ['・四半期は営業CF未開示が多いため通期(FY)を主対象にしています', 'note'],
    ['・本ツールはスクリーニング／監査教育目的。投資判断を保証しません', 'note'],
  ];

  sh.getRange(1, 1, rows.length, 1).setValues(rows.map(r => [r[0]]));
  rows.forEach((r, i) => {
    const cell = sh.getRange(i + 1, 1);
    if (r[1] === 'title') {
      cell.setFontSize(16).setFontWeight('bold').setFontColor('#ffffff').setBackground('#1a1e3a');
      sh.setRowHeight(i + 1, 40);
    } else if (r[1] === 'h') {
      cell.setFontSize(12).setFontWeight('bold').setFontColor('#1a3c6e').setBackground('#e7effb');
      sh.setRowHeight(i + 1, 26);
    } else if (r[1] === 'code') {
      cell.setFontFamily('Consolas').setBackground('#f2f2f2').setFontColor('#b3261e');
    } else if (r[1] === 'note') {
      cell.setFontColor('#666666').setWrap(true);
    } else {
      cell.setWrap(true);
    }
  });
  sh.getRange(1, 1, rows.length, 1).setVerticalAlignment('middle');
  ss.setActiveSheet(sh);
  return sh;
}

// ============================================================================
//  ワークフロー図シート
// ============================================================================

function refreshWorkflowDiagram() { writeWorkflowDiagram_(); }

function writeWorkflowDiagram_() {
  const ss = SpreadsheetApp.getActive();
  const old = ss.getSheetByName(JQ.SHEETS.WORKFLOW);
  if (old) ss.deleteSheet(old);
  const sh = ss.insertSheet(JQ.SHEETS.WORKFLOW, 1);
  sh.setHiddenGridlines(true);
  sh.setColumnWidth(1, 430);
  sh.setColumnWidth(2, 20);
  sh.setColumnWidth(3, 220);

  let r = 1;
  const box = (text, ref, bg, fg) => {
    sh.getRange(r, 1).setValue(text)
      .setBackground(bg).setFontColor(fg || '#000000').setWrap(true)
      .setVerticalAlignment('middle').setHorizontalAlignment('left')
      .setBorder(true, true, true, true, false, false, '#888888', SpreadsheetApp.BorderStyle.SOLID);
    if (ref) sh.getRange(r, 3).setValue(ref).setFontColor('#555555').setFontSize(9)
      .setVerticalAlignment('middle').setWrap(true);
    sh.setRowHeight(r, 46);
    r++;
  };
  const arrow = () => {
    sh.getRange(r, 1).setValue('▼').setHorizontalAlignment('center')
      .setFontColor('#888888').setFontWeight('bold');
    sh.setRowHeight(r, 22);
    r++;
  };

  sh.getRange(r, 1, 1, 3).merge().setValue('会計リスク・スクリーナー　ワークフロー')
    .setFontSize(14).setFontWeight('bold').setFontColor('#ffffff').setBackground('#1a1e3a')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(r, 44); r++;
  sh.setRowHeight(r, 10); r++;

  box('事前準備：J-Quantsに登録 → APIキー発行 → スクリプトプロパティ JQUANTS_API_KEY を設定',
      'スクリプトプロパティ', '#fff3cd', '#5c4a00');
  arrow();
  box('① プライム銘柄を取得\n/equities/master（Mkt=0111）から抽出', '→ 銘柄マスタ シート', '#dce9f8', '#1a3c6e');
  arrow();
  box('② 財務データを収集/続行\n各銘柄の /fins/summary を取得（時間分割・90秒後自動再開・重複防止）', '→ 財務データ シート', '#dce9f8', '#1a3c6e');
  arrow();
  box('③ リスクスコアを計算\n通期(FY)決算からアクルーアル比率を主軸に会計リスク偏差値を算出', '→ リスクランキング シート', '#d9ead3', '#1a4a1a');
  arrow();
  box('JSON出力\naccounting_risk_prime.json を生成', '→ Google Drive', '#d9ead3', '#1a4a1a');
  arrow();
  box('GitHub Pages で公開\nindex.html が JSON を読み込み、ランキングを表示', '→ 公開ページ', '#e8e8e8', '#333333');

  sh.getRange(1, 1, r - 1, 3).setVerticalAlignment('middle');
  ss.setActiveSheet(sh);
  return sh;
}
