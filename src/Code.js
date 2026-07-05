// ============================================================================
//  東証プライム 会計リスク・スクリーナー（J-Quants API）
//  ---------------------------------------------------------------------------
//  無料Freeプラン（約12週遅延）で取得できるサマリー財務から「利益の質」を評価し、
//  会計リスク（利益操作・急悪化）の兆候をランク付けする。投資助言ではない。
//
//  使い方（初回）:
//   1) スクリプトプロパティに JQUANTS_MAIL / JQUANTS_PASSWORD を設定
//      （または JQUANTS_REFRESH_TOKEN を直接設定してもよい）
//   2) メニュー「会計リスク」→ セットアップ
//   3) ① プライム銘柄を取得 → ② 財務データを収集 → ③ リスクスコアを計算
//
//  ※「要確認」コメントの箇所は J-Quants の仕様変更で名称が変わり得る点。
// ============================================================================

const JQ = {
  BASE: 'https://api.jquants.com/v1',
  SHEETS: {
    UNIVERSE:   '銘柄マスタ',        // プライム銘柄一覧
    STATEMENTS: '財務データ',        // 収集した決算（生データ）
    RANKING:    'リスクランキング',  // スコア計算結果
  },
  MARKET_NAME_PRIME: 'プライム',     // /listed/info の MarketCodeName（要確認）
  MARKET_CODE_PRIME: '0111',         // /listed/info の MarketCode（要確認）
  ID_TOKEN_TTL_SEC: 23 * 3600,       // idTokenは24h有効 → 安全に23hだけキャッシュ
  TIME_BUDGET_MS:   4.5 * 60 * 1000, // 1回の実行で使う時間上限（GAS 6分制限対策）
};

// J-Quants /fins/statements で参照する項目名（要確認: 公式ドキュメント準拠）
const F = {
  code:      'LocalCode',
  disclosed: 'DisclosedDate',
  periodType:'TypeOfCurrentPeriod',   // 'FY' | '1Q' | '2Q' | '3Q'
  docType:   'TypeOfDocument',
  netSales:  'NetSales',
  opProfit:  'OperatingProfit',
  ordProfit: 'OrdinaryProfit',
  profit:    'Profit',
  totalAssets:'TotalAssets',
  equity:    'Equity',
  cfo:       'CashFlowsFromOperatingActivities',
  eps:       'EarningsPerShare',
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
    .addToUi();
}

function setup() {
  const ss = SpreadsheetApp.getActive();
  Object.values(JQ.SHEETS).forEach(name => { if (!ss.getSheetByName(name)) ss.insertSheet(name); });
  SpreadsheetApp.getActive().toast('シートを準備しました', '会計リスク', 5);
}

// ============================================================================
//  認証（auth_user → refreshToken → auth_refresh → idToken）
// ============================================================================

function getIdToken_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('JQ_ID_TOKEN');
  if (cached) return cached;

  const props   = PropertiesService.getScriptProperties();
  let   refresh = props.getProperty('JQUANTS_REFRESH_TOKEN');

  if (!refresh) {
    const mail = props.getProperty('JQUANTS_MAIL');
    const pass = props.getProperty('JQUANTS_PASSWORD');
    if (!mail || !pass) throw new Error('JQUANTS_MAIL / JQUANTS_PASSWORD（または JQUANTS_REFRESH_TOKEN）を設定してください');
    const res = UrlFetchApp.fetch(JQ.BASE + '/token/auth_user', {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify({ mailaddress: mail, password: pass }),
      muteHttpExceptions: true,
    });
    if (res.getResponseCode() !== 200) throw new Error('auth_user失敗: ' + res.getContentText().slice(0, 300));
    refresh = JSON.parse(res.getContentText()).refreshToken;
  }

  const r2 = UrlFetchApp.fetch(
    JQ.BASE + '/token/auth_refresh?refreshtoken=' + encodeURIComponent(refresh),
    { method: 'post', muteHttpExceptions: true }
  );
  if (r2.getResponseCode() !== 200) throw new Error('auth_refresh失敗: ' + r2.getContentText().slice(0, 300));
  const idToken = JSON.parse(r2.getContentText()).idToken;

  cache.put('JQ_ID_TOKEN', idToken, JQ.ID_TOKEN_TTL_SEC);
  return idToken;
}

// GET（pagination_key 自動追従）。戻り値はレスポンスJSONの配列。
function jqGet_(path, params) {
  const idToken = getIdToken_();
  const base = JQ.BASE + path;
  const q = params
    ? Object.keys(params).filter(k => params[k] != null && params[k] !== '')
        .map(k => k + '=' + encodeURIComponent(params[k])).join('&')
    : '';
  let url = q ? base + '?' + q : base;

  const pages = [];
  let pagination = null;
  do {
    const u = pagination ? url + (url.includes('?') ? '&' : '?') + 'pagination_key=' + encodeURIComponent(pagination) : url;
    const res  = UrlFetchApp.fetch(u, { headers: { Authorization: 'Bearer ' + idToken }, muteHttpExceptions: true });
    const code = res.getResponseCode();
    if (code === 401) { CacheService.getScriptCache().remove('JQ_ID_TOKEN'); }
    if (code !== 200) throw new Error('GET ' + path + ' 失敗(' + code + '): ' + res.getContentText().slice(0, 300));
    const json = JSON.parse(res.getContentText());
    pages.push(json);
    pagination = json.pagination_key || null;
  } while (pagination);
  return pages;
}

// ============================================================================
//  ① プライム銘柄マスタ
// ============================================================================

function fetchPrimeUniverse() {
  const pages = jqGet_('/listed/info');
  const info  = pages.reduce((a, p) => a.concat(p.info || []), []);
  const prime = info.filter(x => x.MarketCode === JQ.MARKET_CODE_PRIME || x.MarketCodeName === JQ.MARKET_NAME_PRIME);

  const rows = prime.map(x => [
    String(x.Code),
    x.CompanyName || '',
    x.Sector17CodeName || x.Sector33CodeName || '',
    x.MarketCodeName || '',
  ]);

  const sh = SpreadsheetApp.getActive().getSheetByName(JQ.SHEETS.UNIVERSE);
  sh.clearContents();
  sh.getRange(1, 1, 1, 4).setValues([['コード', '企業名', '業種', '市場']]);
  if (rows.length) sh.getRange(2, 1, rows.length, 4).setValues(rows);
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
      const pages = jqGet_('/fins/statements', { code: code });
      const list  = pages.reduce((a, p) => a.concat(p.statements || []), []);
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
    String(s[F.code] || ''), s[F.disclosed] || '', s[F.periodType] || '', s[F.docType] || '',
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
