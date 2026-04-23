/*************************************************
 * 決算チェックツール 安全統合版
 * v2026-04-15.05
 *
 * 方針
 * - 起動 / フォルダ取込 / old移動 は既存版を踏襲
 * - ルール部分のみ 新_統合ルール / 新_検索定義 に差し替え
 * - 日本語列名ベースで運用しやすくする
 *
 * 前提シート
 * - 決算書
 * - 新_統合ルール
 * - 新_検索定義
 * - 新_科目グループ
 * - 新_科目正規化
 * - 新_除外科目
 * - 新_AIチェック対象（任意）
 *
 * スクリプトプロパティ（任意）
 * - GEMINI_API_KEY
 * - GEMINI_MODEL
 * - AI_CHECK_ENABLED = true / false
 * - MOVE_EXCEL_TO_OLD = true / false
 * - MOVE_CONVERTED_TO_OLD = true / false
 *
 * 注意
 * - Excel取込のため Advanced Drive API を ON にしてください
 * - ルールシートの加減算記号は半角 + - を使用してください
 *************************************************/



const CONFIG = {
  TEMPLATE_SHEET_ID: '',
  RESULT_FILE_PREFIX: '決算チェック_',

  SHEET_DECISION: '決算書',

  SHEET_RULES: '新_統合ルール',
  SHEET_LOOKUP: '新_検索定義',
  SHEET_GROUP_MASTER: '新_科目グループ',
  SHEET_NORMALIZE_MASTER: '新_科目正規化',
  SHEET_EXCLUDE_MASTER: '新_除外科目',
  SHEET_AI_TARGETS: '新_AIチェック対象',

  SHEET_RESULT: 'check_result',
  SHEET_LOG: 'check_log',
  SHEET_ACCOUNT_MATCH_LOG: 'account_match_log',
  SHEET_AI_CHECK_LOG: 'ai_check_log',
  SHEET_REPORT: 'report_A4',

  AI_CHECK_START_ROW: 5,
};

/* =========================
 * メニュー
 * ========================= */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('照合ツール')
    .addItem('照合実行（このブック）', 'runChecks')
    .addItem('結果シート初期化', 'resetResultSheets')
    .addItem('テスト実行（folderId手入力）', 'testRunFromPrompt')
    .addToUi();
}

/* =========================
 * Webアプリ入口
 * ========================= */

function doGet(e) {
  e = e || {};
  const params = e.parameter || {};
  const action = params.action || '';
  const folderId = params.folderId || '';

  if (!folderId) {
    return ContentService.createTextOutput('folderId missing');
  }

  if (action === 'run') {
    try {
      const result = runFromFolder_(folderId);
      return HtmlService
        .createHtmlOutput(
          `<p>処理が完了しました。</p><p><a href="${result.url}" target="_blank">結果スプレッドシートを開く</a></p>`
        )
        .setTitle('決算チェック');
    } catch (err) {
      return HtmlService
        .createHtmlOutput(
          `<p>エラーが発生しました。</p><pre>${escapeHtml_(String(err && err.stack ? err.stack : err))}</pre>`
        )
        .setTitle('決算チェック');
    }
  }

  return ContentService.createTextOutput('no action');
}

function testRunFromPrompt() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('folderIdを入力してください');
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const folderId = res.getResponseText().trim();
  if (!folderId) {
    ui.alert('folderIdが空です');
    return;
  }
  const result = runFromFolder_(folderId);
  ui.alert(`完了しました\n${result.url}`);
}

/* =========================
 * フォルダ単位実行
 * ========================= */

function getFlag_(key, defaultValue) {
  const v = PropertiesService.getScriptProperties().getProperty(key);
  if (v === null || v === '') return !!defaultValue;
  return String(v).toLowerCase() === 'true';
}

function getMoveExcelToOldFlag_() {
  return getFlag_('MOVE_EXCEL_TO_OLD', true);
}

function getMoveConvertedToOldFlag_() {
  return getFlag_('MOVE_CONVERTED_TO_OLD', true);
}

function getAiCheckEnabledFlag_() {
  return getFlag_('AI_CHECK_ENABLED', true);
}

function getOrCreateOldFolder_(parentFolder) {
  const folders = parentFolder.getFoldersByName('old');
  if (folders.hasNext()) return folders.next();
  return parentFolder.createFolder('old');
}

function moveFileToOldIfNeeded_(file, oldFolder, enabled) {
  if (!enabled || !file || !oldFolder) return;

  try {
    const parents = file.getParents();
    while (parents.hasNext()) {
      const p = parents.next();
      if (p.getId() === oldFolder.getId()) return;
    }

    oldFolder.addFile(file);

    const parents2 = file.getParents();
    while (parents2.hasNext()) {
      const p = parents2.next();
      if (p.getId() !== oldFolder.getId()) {
        p.removeFile(file);
      }
    }
  } catch (e) {
    Logger.log(`old移動失敗: ${file.getName()} / ${e.message}`);
  }
}

function runFromFolder_(folderId) {
  const folder = DriveApp.getFolderById(folderId);
  const oldFolder = getOrCreateOldFolder_(folder);

  const templateSs = getTemplateSpreadsheet_();
  const resultSs = createResultSpreadsheetFromTemplate_(templateSs, folder);

  importFilesFromFolder_(folder, oldFolder, resultSs);
  runChecksCore_(resultSs);

  return {
    spreadsheetId: resultSs.getId(),
    url: resultSs.getUrl(),
  };
}

function getTemplateSpreadsheet_() {
  if (CONFIG.TEMPLATE_SHEET_ID) {
    return SpreadsheetApp.openById(CONFIG.TEMPLATE_SHEET_ID);
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('テンプレートスプレッドシートを取得できません。CONFIG.TEMPLATE_SHEET_ID を設定してください。');
  }
  return ss;
}

function createResultSpreadsheetFromTemplate_(templateSs, folder) {
  const templateFile = DriveApp.getFileById(templateSs.getId());
  const name = `${CONFIG.RESULT_FILE_PREFIX}${getNowStr_()}`;
  const copiedFile = templateFile.makeCopy(name, folder);
  return SpreadsheetApp.openById(copiedFile.getId());
}

function importFilesFromFolder_(folder, oldFolder, resultSs) {
  const files = folder.getFiles();
  const templateId = getTemplateSpreadsheet_().getId();
  let fileCount = 0;
  let importedGoogleCount = 0;
  let importedExcelCount = 0;
  let skippedCount = 0;

  appendLogRow_(resultSs, `取込開始: folder=${folder.getName()} (${folder.getId()})`);

  while (files.hasNext()) {
    const file = files.next();
    fileCount++;

    const fileId = file.getId();
    const fileName = file.getName();
    const mimeType = file.getMimeType();

    appendLogRow_(resultSs, `検出: name=${fileName} / mimeType=${mimeType}`);

    if (fileId === resultSs.getId()) {
      appendLogRow_(resultSs, `除外: 結果ブック ${fileName}`);
      skippedCount++;
      continue;
    }

    if (fileId === templateId) {
      appendLogRow_(resultSs, `除外: テンプレート ${fileName}`);
      skippedCount++;
      continue;
    }

    if (fileName === 'old') {
      appendLogRow_(resultSs, `除外: old`);
      skippedCount++;
      continue;
    }

    if (mimeType === MimeType.GOOGLE_SHEETS) {
      importGoogleSpreadsheetFile_(file, resultSs);
      importedGoogleCount++;
      appendLogRow_(resultSs, `Googleスプレッドシート取込: ${fileName}`);
      continue;
    }

    if (isExcelFile_(file)) {
      importExcelFile_(file, folder, oldFolder, resultSs);
      importedExcelCount++;
      appendLogRow_(resultSs, `Excel取込完了: ${fileName}`);
      continue;
    }

    if (mimeType === MimeType.PDF || /\.pdf$/i.test(fileName)) {
      appendLogRow_(resultSs, `PDFは未取込（OCR未実装）: ${fileName}`);
      skippedCount++;
      continue;
    }

    appendLogRow_(resultSs, `未対応形式のためスキップ: ${fileName}`);
    skippedCount++;
  }

  appendLogRow_(
    resultSs,
    `取込完了: 総件数=${fileCount}, Googleスプレッドシート=${importedGoogleCount}, Excel=${importedExcelCount}, スキップ=${skippedCount}`
  );
}

function importGoogleSpreadsheetFile_(file, resultSs) {
  const sourceSs = SpreadsheetApp.openById(file.getId());
  copySourceSheets_(sourceSs, resultSs, file.getName());
}

function importExcelFile_(file, folder, oldFolder, resultSs) {
  const moveExcelToOld = getMoveExcelToOldFlag_();
  const moveConvertedToOld = getMoveConvertedToOldFlag_();
  const fileName = file.getName();

  appendLogRow_(resultSs, `Excel変換開始: ${fileName}`);

  const tempSs = convertExcelToSpreadsheet_(file, folder);

  if (!tempSs) {
    throw new Error(`Excel変換に失敗しました: ${fileName}`);
  }

  try {
    appendLogRow_(resultSs, `Excel変換成功: ${fileName} -> ${tempSs.getName()} (${tempSs.getId()})`);
    copySourceSheets_(tempSs, resultSs, fileName);
    appendLogRow_(resultSs, `Excelシート統合完了: ${fileName}`);
  } finally {
    try {
      moveFileToOldIfNeeded_(file, oldFolder, moveExcelToOld);
      appendLogRow_(resultSs, `元Excel old移動: ${fileName} / enabled=${moveExcelToOld}`);
    } catch (e1) {
      appendLogRow_(resultSs, `元Excel old移動失敗: ${fileName} / ${e1.message}`);
    }

    try {
      moveFileToOldIfNeeded_(DriveApp.getFileById(tempSs.getId()), oldFolder, moveConvertedToOld);
      appendLogRow_(resultSs, `変換後スプレッドシート old移動: ${tempSs.getName()} / enabled=${moveConvertedToOld}`);
    } catch (e2) {
      appendLogRow_(resultSs, `変換後スプレッドシート old移動失敗: ${fileName} / ${e2.message}`);
    }
  }
}

function convertExcelToSpreadsheet_(file, folder) {
  const blob = file.getBlob();
  const fileName = file.getName();

  const resource = {
    title: `[tmp]${fileName}`,
    mimeType: MimeType.GOOGLE_SHEETS,
    parents: [{ id: folder.getId() }],
  };

  try {
    const converted = Drive.Files.insert(resource, blob);
    if (!converted || !converted.id) {
      throw new Error('Drive.Files.insert の戻り値に id がありません');
    }
    return SpreadsheetApp.openById(converted.id);
  } catch (e) {
    throw new Error(`Excel変換エラー: ${fileName} / ${e.message}`);
  }
}

function copySourceSheets_(sourceSs, targetSs, sourceFileName) {
  const sheets = sourceSs.getSheets();
  const multiple = sheets.length > 1;

  for (let i = 0; i < sheets.length; i++) {
    const sh = sheets[i];
    const targetName = buildImportedSheetName_(sourceFileName, sh.getName(), multiple, i);

    const existing = targetSs.getSheetByName(targetName);
    if (existing) targetSs.deleteSheet(existing);

    const copied = sh.copyTo(targetSs);
    copied.setName(targetName);
  }
}

function buildImportedSheetName_(fileName, originalSheetName, multiple, index) {
  const base = stripExtension_(fileName);

  if (base.includes('決算書')) {
    return CONFIG.SHEET_DECISION;
  }

  let name = multiple ? `${base}_${originalSheetName}` : base;
  name = name.replace(/[\\\/\?\*\[\]:]/g, '_');
  return name.substring(0, 95);
}

function stripExtension_(name) {
  return String(name || '').replace(/\.[^/.]+$/, '');
}

function isExcelMimeType_(mimeType) {
  const s = String(mimeType || '');
  return [
    MimeType.MICROSOFT_EXCEL,
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'application/vnd.ms-excel',
    'application/vnd.ms-excel.sheet.macroenabled.12',
    'application/vnd.ms-excel.sheet.binary.macroenabled.12',
    'application/octet-stream'
  ].includes(s);
}

function isExcelFile_(file) {
  const fileName = String(file.getName() || '');
  const mimeType = String(file.getMimeType() || '');

  return (
    isExcelMimeType_(mimeType) ||
    /\.xlsx$/i.test(fileName) ||
    /\.xls$/i.test(fileName) ||
    /\.xlsm$/i.test(fileName) ||
    /\.xlsb$/i.test(fileName)
  );
}

/* =========================
 * 手動実行（このブック）
 * ========================= */

function runChecks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  runChecksCore_(ss);
}

function runChecksCore_(ss) {
  const startedAt = new Date();

  ensureSheetInSpreadsheet_(ss, CONFIG.SHEET_RESULT);
  ensureSheetInSpreadsheet_(ss, CONFIG.SHEET_LOG);
  resetResultSheetsInSpreadsheet_(ss);

  try {
    appendLogRow_(ss, '照合開始');

    const ctx = buildContext_(ss);

    const ruleResults = runUnifiedRules_(ss, ctx);
    const aiCheckEnabled = getAiCheckEnabledFlag_();
    const typoResults = aiCheckEnabled ? runGlobalInputAnomalyChecks_(ss) : [];
    if (!aiCheckEnabled) {
      appendLogRow_(ss, 'AI入力値チェックをスキップしました（AI_CHECK_ENABLED=false）');
    }

    writeAccountMatchLog_(ss, ctx);
    const bsUnusedResults = buildUnusedBSAccountResults_(ctx);

    const allResults = [...ruleResults, ...typoResults, ...bsUnusedResults];
    writeResults_(ss, allResults);
    buildA4Report_(ss);

    appendLogRow_(ss, `照合完了 件数=${(allResults || []).filter(r => r != null).length}`);
  } catch (err) {
    appendLogRow_(ss, `ERROR: ${err.stack || err}`);
    throw err;
  } finally {
    const sec = Math.round((new Date() - startedAt) / 1000);
    appendLogRow_(ss, `処理時間 ${sec} 秒`);
  }
}

function resetResultSheets() {
  resetResultSheetsInSpreadsheet_(SpreadsheetApp.getActiveSpreadsheet());
}

function resetResultSheetsInSpreadsheet_(ss) {
  const result = ensureSheetInSpreadsheet_(ss, CONFIG.SHEET_RESULT);
  result.clearContents().clearFormats();
  result.getRange(1, 1, 1, 18).setValues([[
    '判定',
    '区分',
    'ルールID',
    '対象シート',
    '対象項目',
    '対象セル',
    'ジャンプURL',
    '決算書値',
    '比較値',
    '差額',
    '条件',
    'メッセージ',
    '詳細',
    'AI判定',
    'AI理由',
    'AI修正候補1',
    'AI修正候補2',
    '日時',
  ]]);

  const log = ensureSheetInSpreadsheet_(ss, CONFIG.SHEET_LOG);
  log.clearContents().clearFormats();
  log.getRange(1, 1, 1, 2).setValues([['日時', 'ログ']]);

  const accountLog = ensureSheetInSpreadsheet_(ss, CONFIG.SHEET_ACCOUNT_MATCH_LOG);
  accountLog.clearContents().clearFormats();
  accountLog.getRange(1, 1, 1, 6).setValues([[
    'No',
    '勘定科目',
    '決算書値',
    '照合利用有無',
    '対象レンジ内',
    '備考'
  ]]);

  const aiLog = ensureSheetInSpreadsheet_(ss, CONFIG.SHEET_AI_CHECK_LOG);
  aiLog.clearContents().clearFormats();
  aiLog.getRange(1, 1, 1, 7).setValues([[
    '日時',
    'シート名',
    'セル',
    '入力値',
    '取得元',
    'AI判定',
    'AI理由'
  ]]);

  const report = ss.getSheetByName(CONFIG.SHEET_REPORT);
  if (report) ss.deleteSheet(report);
}

/* =========================
 * Context
 * ========================= */

function buildContext_(ss) {
  const decisionSheet = getRequiredSheet_(ss, CONFIG.SHEET_DECISION);
  const parsed = parseDecisionSheet_(decisionSheet);

  return {
    ss,
    decisionSheet,
    decisionValues: decisionSheet.getDataRange().getDisplayValues(),
    decisionMap: parsed.map,
    bsAccounts: parsed.bsAccounts,
    usedDecisionAccounts: new Set(),

    rules: loadUnifiedRules_(ss),
    lookupMaster: loadLookupMaster_(ss),
    groups: loadGroupMasterJa_(ss),
    normalizeRules: loadNormalizeMasterJa_(ss),
    excludeMaster: loadExcludeMasterJa_(ss),

    trackDecisionUsage: true,
  };
}

/* =========================
 * 統合ルール読込
 * ========================= */

function loadUnifiedRules_(ss) {
  const sheet = getRequiredSheet_(ss, CONFIG.SHEET_RULES);
  const rows = loadRowsAsObjectsJapanese_(sheet);
  
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  if (headerRow.indexOf('ルールID') === -1) {
  appendLogRow_(ss, '警告: 新_統合ルール のヘッダーに「ルールID」がありません。ヘッダー崩れの可能性があります。');
  }

  const mapped = rows
    .map((r, idx) => {
      const fallbackRuleId =
        String(r['ルールID'] || '').trim() ||
        String(r['照合先の値'] || '').trim(); // ヘッダー崩れ時の救済

      return {
        rule_id: fallbackRuleId,
        enabled: toBooleanJa_(r['有効']),
        check_type: normalizeCheckTypeJa_(r['チェック種別']),
        category: String(r['区分'] || '').trim(),

        src_sheet: String(r['照合元シート名（部分一致）'] || r['照合元シート名'] || '').trim(),
        src_key: String(r['照合元項目・見出し'] || r['照合元項目名'] || '').trim(),
        src_account_col: String(r['照合元科目列'] || '').trim(),
        src_amount_col: String(r['照合元金額列'] || '').trim(),
        src_agg: String(r['照合元集計方法'] || '').trim(),
        src_value_expr: String(r['照合元値指定'] || '').trim(),

        dst_sheet: String(r['照合先シート名（部分一致）'] || r['照合先シート名'] || '').trim(),
        dst_key: String(r['照合先項目・見出し'] || r['照合先項目名'] || '').trim(),
        dst_account_col: String(r['照合先科目列'] || '').trim(),
        dst_amount_col: String(r['照合先金額列'] || '').trim(),
        dst_agg: String(r['照合先集計方法'] || '').trim(),
        dst_value_expr: String(r['照合先値指定'] || '').trim(),

        condition_expr: String(r['条件式'] || '').trim(),
        match_mode: String(r['一致方法'] || '').trim(),
        lookup_id: String(r['検索定義ID'] || '').trim(),

        severity: String(r['重要度'] || '要確認').trim(),
        message: String(r['メッセージ'] || '').trim(),
        note: String(r['移行メモ'] || '').trim(),
        __row_index: idx + 2,
      };
    })
    .filter(rule => rule.rule_id);

  const deduped = [];
  const seen = new Set();
  let headerLikeSkipped = 0;
  let duplicatedSkipped = 0;

  mapped.forEach(rule => {
    if (isHeaderLikeUnifiedRule_(rule)) {
      headerLikeSkipped++;
      return;
    }

    const key = buildUnifiedRuleDedupKey_(rule);
    if (seen.has(key)) {
      duplicatedSkipped++;
      return;
    }
    seen.add(key);
    deduped.push(rule);
  });

  if (headerLikeSkipped > 0) {
    appendLogRow_(ss, `統合ルール読込: ヘッダー行混入を ${headerLikeSkipped} 行スキップしました`);
  }
  if (duplicatedSkipped > 0) {
    appendLogRow_(ss, `統合ルール読込: 重複ルールを ${duplicatedSkipped} 行スキップしました`);
  }

  return deduped;
}

function isHeaderLikeUnifiedRule_(rule) {
  const id = String(rule.rule_id || '').trim();
  const checkType = String(rule.check_type || '').trim();
  const srcSheet = String(rule.src_sheet || '').trim();
  const srcKey = String(rule.src_key || '').trim();
  const srcAgg = String(rule.src_agg || '').trim();
  const dstSheet = String(rule.dst_sheet || '').trim();
  const dstKey = String(rule.dst_key || '').trim();
  const dstAgg = String(rule.dst_agg || '').trim();

  return (
    id === 'ルールID' ||
    checkType === 'チェック種別' ||
    srcSheet.includes('照合元シート名') ||
    srcKey.includes('照合元項目') ||
    srcAgg.includes('照合元集計方法') ||
    dstSheet.includes('照合先シート名') ||
    dstKey.includes('照合先項目') ||
    dstAgg.includes('照合先集計方法')
  );
}

function buildUnifiedRuleDedupKey_(rule) {
  return [
    rule.rule_id,
    rule.enabled ? '1' : '0',
    rule.check_type,
    rule.src_sheet,
    rule.src_key,
    rule.src_account_col,
    rule.src_amount_col,
    rule.src_agg,
    rule.src_value_expr,
    rule.dst_sheet,
    rule.dst_key,
    rule.dst_account_col,
    rule.dst_amount_col,
    rule.dst_agg,
    rule.dst_value_expr,
    rule.condition_expr,
    rule.match_mode,
    rule.lookup_id,
    rule.severity,
    rule.message
  ].join('|');
}

function loadLookupMaster_(ss) {
  const sheet = ss.getSheetByName(CONFIG.SHEET_LOOKUP);
  if (!sheet) return {};

  const rows = loadRowsAsObjectsJapanese_(sheet);
  const map = {};

  rows.forEach(r => {
    const id = String(r['検索定義ID'] || r['検索ID'] || '').trim();
    if (!id) return;

    map[id] = {
      enabled: toBooleanJa_(r['有効']),
      status: String(r['設定状況'] || '').trim(),
      sheet_pattern: String(r['対象シート名（部分一致）'] || r['対象シート名'] || '').trim(),
      match_col: String(r['検索列'] || '').trim(),
      value_col_expr: String(r['取得列・式'] || r['取得列'] || '').trim(),
      key_type: String(r['検索値の取得方法'] || r['検索値種別'] || '').trim(),
      key_value: String(r['検索値（固定）'] || r['検索値'] || '').trim(),
      key_sheet: String(r['検索値取得元シート名'] || r['検索値取得元シート'] || '').trim(),
      key_cell: String(r['検索値取得セル'] || '').trim(),
      match_mode: String(r['一致方法'] || r['検索方法'] || '完全一致').trim(),
      old_rule_id: String(r['旧ルールID'] || '').trim(),
      note: String(r['備考'] || '').trim(),
    };
  });

  return map;
}

function loadGroupMasterJa_(ss) {
  const sheet = ss.getSheetByName(CONFIG.SHEET_GROUP_MASTER);
  if (!sheet) return {};

  const rows = loadRowsAsObjectsJapanese_(sheet);
  const map = {};

  rows.forEach(r => {
    const groupName = String(r['グループ名'] || '').trim();
    const account = String(r['勘定科目'] || r['科目名'] || '').trim();
    if (!groupName || !account) return;

    if (!map[groupName]) map[groupName] = [];
    map[groupName].push(account);
  });

  return map;
}

function loadNormalizeMasterJa_(ss) {
  const sheet = ss.getSheetByName(CONFIG.SHEET_NORMALIZE_MASTER);
  if (!sheet) return [];

  const rows = loadRowsAsObjectsJapanese_(sheet);
  return rows
    .map(r => ({
      from: normalizeText_(r['変換前'] || r['元キーワード'] || ''),
      to: normalizeText_(r['変換後'] || r['正規化後科目'] || ''),
    }))
    .filter(r => r.from && r.to);
}

function loadExcludeMasterJa_(ss) {
  const sheet = ss.getSheetByName(CONFIG.SHEET_EXCLUDE_MASTER);
  if (!sheet) return [];

  const rows = loadRowsAsObjectsJapanese_(sheet);
  return rows
    .filter(r => String(r['除外科目'] || '').trim())
    .map(r => ({
      rule_id: String(r['対象ルールID'] || '').trim(),
      account: normalizeText_(r['除外科目']),
    }));
}

/* =========================
 * 統合ルール実行
 * ========================= */

function runUnifiedRules_(ss, ctx) {
  const results = [];
  const ruleValueRows = [];

  for (const rule of ctx.rules) {
    if (!rule.enabled) continue;

    try {
      switch (rule.check_type) {
        case 'VALUE_MATCH': {
          const src = resolveRuleSideValue_(ss, rule, 'src', ctx);
          const dst = resolveRuleSideValue_(ss, rule, 'dst', ctx);
          ruleValueRows.push({
            rowIndex: rule.__row_index,
            srcValue: src.value,
            dstValue: dst.value,
          });

          const sheetName = src.sheetName || dst.sheetName || '';
          const itemName = rule.src_key || rule.dst_key || rule.rule_id;
          const targetCell = src.a1 || dst.a1 || '';
          const jumpUrl = targetCell && sheetName ? buildRangeUrl_(ss, sheetName, targetCell) : '';

          results.push(compareNumbersResult_(
            rule,
            sheetName,
            itemName,
            dst.value,
            src.value,
            rule.category || '',
            targetCell,
            jumpUrl
          ));
          break;
        }

        case 'NOT_BLANK': {
          const src = resolveRuleSideValue_(ss, rule, 'src', ctx);
          ruleValueRows.push({
            rowIndex: rule.__row_index,
            srcValue: src.value,
            dstValue: '',
          });
          const bad = isBlank_(src.value);

          results.push(makeResult_({
            status: bad ? (rule.severity || '要確認') : 'OK',
            category: rule.category || '',
            ruleId: rule.rule_id,
            sheetName: src.sheetName || '',
            itemName: rule.src_key || rule.rule_id,
            targetCell: src.a1 || '',
            jumpUrl: src.a1 && src.sheetName ? buildRangeUrl_(ss, src.sheetName, src.a1) : '',
            decisionValue: '',
            compareValue: src.value || '',
            diff: '',
            condition: 'NOT_BLANK',
            message: bad ? (rule.message || '空欄です') : 'OK',
            detail: '',
          }));
          break;
        }

        case 'CONDITIONAL_NOT_BLANK': {
          const src = resolveRuleSideValue_(ss, rule, 'src', ctx);
          ruleValueRows.push({
            rowIndex: rule.__row_index,
            srcValue: src.value,
            dstValue: '',
          });
          const cond = evaluateSimpleCondition_(rule.condition_expr, ctx);
          const bad = cond && isBlank_(src.value);

          results.push(makeResult_({
            status: bad ? (rule.severity || '要確認') : 'OK',
            category: rule.category || '',
            ruleId: rule.rule_id,
            sheetName: src.sheetName || '',
            itemName: rule.src_key || rule.rule_id,
            targetCell: src.a1 || '',
            jumpUrl: src.a1 && src.sheetName ? buildRangeUrl_(ss, src.sheetName, src.a1) : '',
            decisionValue: '',
            compareValue: src.value || '',
            diff: '',
            condition: 'CONDITIONAL_NOT_BLANK',
            message: bad ? (rule.message || '条件を満たすのに空欄です') : 'OK',
            detail: '',
          }));
          break;
        }

        case 'TEXT_MATCH': {
          const src = resolveRuleSideValue_(ss, rule, 'src', ctx);
          const dst = resolveRuleSideValue_(ss, rule, 'dst', ctx);
          ruleValueRows.push({
            rowIndex: rule.__row_index,
            srcValue: src.value,
            dstValue: dst.value,
          });

          if (isBlank_(src.value) && isBlank_(dst.value)) {
            results.push(null);
            break;
          }

          results.push(compareTextsResult_(
            rule,
            src.sheetName || dst.sheetName || '',
            rule.src_key || rule.rule_id,
            String(dst.value || ''),
            String(src.value || ''),
            rule.category || '',
            src.a1 || '',
            src.a1 && src.sheetName ? buildRangeUrl_(ss, src.sheetName, src.a1) : ''
          ));
          break;
        }

        case 'HEADER_EXISTS': {
          const sheet = findSheetFlexible_(ss, rule.src_sheet);
          const values = sheet ? sheet.getDataRange().getDisplayValues() : [];
          const found = sheet ? findCellByText_(values, rule.src_key) : null;
          const a1 = found ? toA1_(found.row + 1, found.col + 1) : '';
          ruleValueRows.push({
            rowIndex: rule.__row_index,
            srcValue: found ? rule.src_key : '',
            dstValue: '',
          });

          results.push(makeResult_({
            status: found ? 'OK' : (rule.severity || '要確認'),
            category: rule.category || '',
            ruleId: rule.rule_id,
            sheetName: sheet ? sheet.getName() : '',
            itemName: rule.src_key || '',
            targetCell: a1,
            jumpUrl: a1 && sheet ? buildRangeUrl_(ss, sheet.getName(), a1) : '',
            decisionValue: '',
            compareValue: found ? rule.src_key : '',
            diff: '',
            condition: 'HEADER_EXISTS',
            message: found ? 'OK' : (rule.message || '見出しが見つかりません'),
            detail: '',
          }));
          break;
        }

        case 'FIXED_MATCH': {
          const src = resolveRuleSideValue_(ss, rule, 'src', ctx);
          const expected = rule.dst_value_expr;
          ruleValueRows.push({
            rowIndex: rule.__row_index,
            srcValue: src.value,
            dstValue: expected,
          });

          if (isBlank_(src.value) && isBlank_(expected)) {
            results.push(null);
            break;
          }

          results.push(compareTextsResult_(
            rule,
            src.sheetName || '',
            rule.src_key || rule.rule_id,
            String(expected || ''),
            String(src.value || ''),
            rule.category || '',
            src.a1 || '',
            src.a1 && src.sheetName ? buildRangeUrl_(ss, src.sheetName, src.a1) : ''
          ));
          break;
        }

        case 'EXISTS': {
          const src = resolveRuleSideValue_(ss, rule, 'src', ctx);
          ruleValueRows.push({
            rowIndex: rule.__row_index,
            srcValue: src.value,
            dstValue: '',
          });
          const exists = !isBlank_(src.value);

          results.push(makeResult_({
            status: exists ? 'OK' : (rule.severity || '要確認'),
            category: rule.category || '',
            ruleId: rule.rule_id,
            sheetName: src.sheetName || '',
            itemName: rule.src_key || rule.rule_id,
            targetCell: src.a1 || '',
            jumpUrl: src.a1 && src.sheetName ? buildRangeUrl_(ss, src.sheetName, src.a1) : '',
            decisionValue: '',
            compareValue: src.value || '',
            diff: '',
            condition: 'EXISTS',
            message: exists ? 'OK' : (rule.message || '対象が存在しません'),
            detail: '',
          }));
          break;
        }

        default:
          ruleValueRows.push({
            rowIndex: rule.__row_index,
            srcValue: '',
            dstValue: '',
          });
          results.push(makeResult_({
            status: '要確認',
            category: rule.category || '',
            ruleId: rule.rule_id,
            sheetName: '',
            itemName: '',
            targetCell: '',
            jumpUrl: '',
            decisionValue: '',
            compareValue: '',
            diff: '',
            condition: '未対応チェック種別',
            message: `未対応のチェック種別です: ${rule.check_type}`,
            detail: '',
          }));
      }
    } catch (e) {
      results.push(makeResult_({
        status: '要確認',
        category: rule.category || '',
        ruleId: rule.rule_id,
        sheetName: rule.src_sheet || '',
        itemName: rule.src_key || rule.rule_id,
        targetCell: '',
        jumpUrl: '',
        decisionValue: '',
        compareValue: '',
        diff: '',
        condition: '実行エラー',
        message: e.message,
        detail: e.stack || '',
      }));
      ruleValueRows.push({
        rowIndex: rule.__row_index,
        srcValue: '',
        dstValue: '',
      });
    }
  }

  writeRuleSideValuesToSheet_(ss, ruleValueRows);
  return results;
}

function writeRuleSideValuesToSheet_(ss, rows) {
  if (!rows || rows.length === 0) return;

  const sheet = ss.getSheetByName(CONFIG.SHEET_RULES);
  if (!sheet) return;

  const header = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getDisplayValues()[0];
  const srcCol = header.indexOf('照合元の値') + 1;
  const dstCol = header.indexOf('照合先の値') + 1;
  if (srcCol <= 0 && dstCol <= 0) return;

  rows.forEach(row => {
    if (!row || !row.rowIndex || row.rowIndex < 2) return;
    if (srcCol > 0) sheet.getRange(row.rowIndex, srcCol).setValue(row.srcValue == null ? '' : row.srcValue);
    if (dstCol > 0) sheet.getRange(row.rowIndex, dstCol).setValue(row.dstValue == null ? '' : row.dstValue);
  });
}

/* =========================
 * 照合元/照合先 値取得
 * ========================= */

function resolveRuleSideValue_(ss, rule, side, ctx) {
  const sheetName = side === 'src' ? rule.src_sheet : rule.dst_sheet;
  const key = side === 'src' ? rule.src_key : rule.dst_key;
  const accountCol = side === 'src' ? rule.src_account_col : rule.dst_account_col;
  const amountCol = side === 'src' ? rule.src_amount_col : rule.dst_amount_col;
  let agg = String(side === 'src' ? rule.src_agg : rule.dst_agg).trim();
  const valueExpr = String(side === 'src' ? rule.src_value_expr : rule.dst_value_expr).trim();
  const matchMode = String(rule.match_mode || '').trim();

  if (agg === '項目値') agg = 'FIRST';
  if (agg === '項目値あり') agg = 'FIRST_PRESENCE';
  if (agg === '行があれば必須') agg = 'REQUIRED_IF_ANY_ROW';
  if (agg === '見出し下合計') agg = 'SUM_BELOW_HEADER';
  if (agg === '科目別合計') agg = 'SUM_BY_ACCOUNT';
  if (agg === '科目正規化後の科目別合計') agg = 'SUM_BY_ACCOUNT_NORMALIZED';
  if (agg === '見出し列合計') agg = 'SUM_HEADER_COLUMN';
  if (agg === '決算書式') agg = 'DECISION_EXPR';
  if (agg === '検索定義') agg = 'LOOKUP';
  if (agg === 'セル計算式') agg = 'CELL_EXPR';

  if (agg === 'DECISION_EXPR') {
    let decisionValue = resolveDecisionExprSimple_(ctx, valueExpr || key, matchMode);

    const isKilo = shouldApplyKiloScale_(rule, side, key);
    if (decisionValue != null && isKilo) {
      decisionValue = Math.floor(decisionValue / 1000);
    }

    return {
      value: decisionValue,
      sheetName: CONFIG.SHEET_DECISION,
      a1: '',
    };
  }

  if (agg === 'LOOKUP') {
    let lookupId = '';

    if (/^LOOKUP:/i.test(valueExpr)) {
      lookupId = valueExpr.replace(/^LOOKUP:/i, '').trim();
    } else if (rule.lookup_id) {
      lookupId = String(rule.lookup_id).trim();
    }

    let lookupValue = lookupId ? resolveLookupById_(ss, lookupId, ctx) : null;
    const lookupDef = lookupId ? ctx.lookupMaster[lookupId] : null;

    const isKilo = shouldApplyKiloScale_(rule, side, key);
    const alreadyKiloScaledByLookup = isLookupValueAlreadyKiloScaled_(lookupDef);

    if (lookupValue != null && isKilo && !alreadyKiloScaledByLookup) {
      const n = toNumber_(lookupValue);
      if (n != null) {
        lookupValue = Math.floor(n / 1000);
      }
    }

    return {
      value: lookupValue,
      sheetName: '',
      a1: '',
    };
  }

  if (agg === 'CELL_EXPR') {
    const sheet = findSheetFlexible_(ss, sheetName);
    if (!sheet) {
      return { value: null, sheetName: '', a1: '' };
    }
    const values = sheet.getDataRange().getDisplayValues();
    return {
      value: evaluateCellExpression_(valueExpr, values),
      sheetName: sheet.getName(),
      a1: extractFirstA1FromExpr_(valueExpr),
    };
  }

  if (/^LOOKUP:/i.test(valueExpr)) {
    const lookupId = valueExpr.replace(/^LOOKUP:/i, '').trim();
    return {
      value: resolveLookupById_(ss, lookupId, ctx),
      sheetName: '',
      a1: '',
    };
  }

  if (valueExpr && /[+\-]/.test(valueExpr) && !sheetName) {
    return {
      value: resolveDecisionExprSimple_(ctx, valueExpr, matchMode),
      sheetName: CONFIG.SHEET_DECISION,
      a1: '',
    };
  }

  if (valueExpr && !sheetName) {
    const directValue = resolveDecisionExprSimple_(ctx, valueExpr, matchMode);
    if (directValue != null) {
      return {
        value: directValue,
        sheetName: CONFIG.SHEET_DECISION,
        a1: '',
      };
    }
  }

  if (valueExpr && isNumericLike_(valueExpr)) {
    return {
      value: toNumber_(valueExpr),
      sheetName: '',
      a1: '',
    };
  }

  if (!sheetName) {
    return { value: valueExpr || '', sheetName: '', a1: '' };
  }

  const sheet = findSheetFlexible_(ss, sheetName);
  if (!sheet) {
    return { value: null, sheetName: '', a1: '' };
  }

  const values = sheet.getDataRange().getDisplayValues();

  if (agg === 'FIRST' && key && /[+\-]/.test(key)) {
    const exprValue = resolveSheetFirstExpressionValue_(values, key, amountCol);
    return {
      value: exprValue.value,
      sheetName: sheet.getName(),
      a1: exprValue.a1,
    };
  }

  const found = key ? findCellByText_(values, key) : null;

  if (agg === 'FIRST_PRESENCE') {
    if (!found) return { value: null, sheetName: sheet.getName(), a1: '' };
    const val = findRightTextValue_(values, found.row, found.col);
    return {
      value: val,
      sheetName: sheet.getName(),
      a1: toA1_(found.row + 1, found.col + 1),
    };
  }

  if (agg === 'REQUIRED_IF_ANY_ROW') {
    if (!found) return { value: null, sheetName: sheet.getName(), a1: '' };
    const requiredColIndex = found.col;
    const triggerColIndex = accountCol ? colToIndex_(accountCol) : 0;
    const check = checkRequiredColumnIfAnyRow_(values, found.row, triggerColIndex, requiredColIndex);

    return {
      value: check.isValid ? 'OK' : '',
      sheetName: sheet.getName(),
      a1: check.a1,
    };
  }

  if (agg === 'FIRST') {
    if (!found) return { value: null, sheetName: sheet.getName(), a1: '' };

    const colIndex = amountCol ? colToIndex_(amountCol) : found.col;
    if (colIndex < 0) {
      return {
        value: values[found.row][found.col],
        sheetName: sheet.getName(),
        a1: toA1_(found.row + 1, found.col + 1),
      };
    }

    return {
      value: values[found.row][colIndex],
      sheetName: sheet.getName(),
      a1: toA1_(found.row + 1, colIndex + 1),
    };
  }

  if (agg === 'SUM_BELOW_HEADER') {
    if (!found || !amountCol) {
      return { value: null, sheetName: sheet.getName(), a1: '' };
    }
    const colIndex = colToIndex_(amountCol);
    return {
      value: sumColumnBelowHeader_(values, found.row, colIndex, ''),
      sheetName: sheet.getName(),
      a1: toA1_(found.row + 1, colIndex + 1),
    };
  }

  if (agg === 'SUM_HEADER_COLUMN') {
    if (!found) {
      return { value: null, sheetName: sheet.getName(), a1: '' };
    }
    const sumCol = found.col;
    return {
      value: sumColumnBelowHeader_(values, found.row, sumCol, ''),
      sheetName: sheet.getName(),
      a1: toA1_(found.row + 1, sumCol + 1),
    };
  }

  if (agg === 'SUM_BY_ACCOUNT' || agg === 'SUM_BY_ACCOUNT_NORMALIZED') {
    if (!found || !accountCol || !amountCol) {
      return { value: null, sheetName: sheet.getName(), a1: '' };
    }

    const accountIndex = colToIndex_(accountCol);
    const amountIndex = colToIndex_(amountCol);
    const excludeSet = buildExcludeSetForRule_(rule.rule_id, ctx.excludeMaster);

    const sums = {};
    let firstRow = -1;

    for (let r = found.row + 1; r < values.length; r++) {
      const rawAccount = normalizeText_(values[r][accountIndex]);
      if (!rawAccount) continue;

      let account = rawAccount;
      if (agg === 'SUM_BY_ACCOUNT_NORMALIZED') {
        account = normalizeAccountByRule_(rawAccount, ctx.normalizeRules) || rawAccount;
      }

      if (excludeSet.has(account)) continue;

      const n = toNumber_(values[r][amountIndex]);
      if (n == null) continue;

      if (firstRow < 0) firstRow = r;
      sums[account] = (sums[account] || 0) + n;
    }

    const expr = valueExpr || key;
    const resultValue = resolveAccountExpressionFromMap_(expr, sums);

    markAccountsUsedFromExpr_(ctx, expr);

    return {
      value: resultValue,
      sheetName: sheet.getName(),
      a1: firstRow >= 0 ? toA1_(firstRow + 1, amountIndex + 1) : '',
    };
  }

  if (valueExpr !== '') {
    return { value: valueExpr, sheetName: '', a1: '' };
  }

  return { value: null, sheetName: sheet.getName(), a1: '' };
}

function resolveSheetFirstExpressionValue_(values, expr, amountCol) {
  const tokens = String(expr || '').match(/[+\-]?[^+\-]+/g);
  if (!tokens) return { value: null, a1: '' };

  const forcedCol = amountCol ? colToIndex_(amountCol) : -1;
  let total = 0;
  let foundAny = false;
  let firstA1 = '';

  for (const token of tokens) {
    const raw = String(token || '').trim();
    if (!raw) continue;

    const sign = raw.startsWith('-') ? -1 : 1;
    const label = raw.replace(/^[+\-]/, '').trim();
    if (!label) continue;

    const found = findCellByText_(values, label);
    if (!found) continue;

    const colIndex = forcedCol >= 0 ? forcedCol : found.col;
    if (colIndex < 0) continue;

    const n = toNumber_(values[found.row][colIndex]);
    if (n == null) continue;

    total += sign * n;
    foundAny = true;
    if (!firstA1) firstA1 = toA1_(found.row + 1, colIndex + 1);
  }

  return {
    value: foundAny ? total : null,
    a1: firstA1,
  };
}

function checkRequiredColumnIfAnyRow_(values, headerRow, triggerColIndex, requiredColIndex) {
  if (!values || values.length === 0) {
    return { isValid: true, a1: '' };
  }

  const safeTriggerCol = triggerColIndex >= 0 ? triggerColIndex : 0;
  const safeRequiredCol = requiredColIndex >= 0 ? requiredColIndex : 0;
  let hasTargetRows = false;

  for (let r = headerRow + 1; r < values.length; r++) {
    const row = values[r] || [];
    const trigger = String(row[safeTriggerCol] || '').trim();
    if (!trigger) continue;

    hasTargetRows = true;

    const required = String(row[safeRequiredCol] || '').trim();
    if (!required) {
      return {
        isValid: false,
        a1: toA1_(r + 1, safeRequiredCol + 1),
      };
    }
  }

  if (!hasTargetRows) {
    return { isValid: true, a1: '' };
  }

  return { isValid: true, a1: '' };
}

function shouldApplyKiloScale_(rule, side, sideKeyRaw) {
  const sideKey = normalizeText_(sideKeyRaw || '');
  const pairKey = normalizeText_(side === 'src' ? rule.dst_key : rule.src_key);
  return hasKiloLabel_(sideKey) || hasKiloLabel_(pairKey);
}

function hasKiloLabel_(text) {
  const s = normalizeText_(text || '');
  return s.includes('_千円') || s.includes('千円');
}

function isLookupValueAlreadyKiloScaled_(lookupDef) {
  if (!lookupDef) return false;
  const expr = String(lookupDef.value_col_expr || '').toUpperCase();
  if (!expr) return false;

  return (
    /\/\s*1000(?:\D|$)/.test(expr) ||
    /ROUNDDOWN\s*\(.+\/\s*1000\s*,\s*0\s*\)/.test(expr) ||
    /ROUND\s*\(.+\/\s*1000\s*,\s*0\s*\)/.test(expr)
  );
}

/* =========================
 * LOOKUP
 * ========================= */

function resolveLookupById_(ss, lookupId, ctx) {
  const def = ctx.lookupMaster[lookupId];
  if (!def || !def.enabled) return null;

  const sheet = findSheetFlexible_(ss, def.sheet_pattern);
  if (!sheet) return null;

  const values = sheet.getDataRange().getDisplayValues();
  const matchCol = colToIndex_(def.match_col);
  if (matchCol < 0) return null;

  let searchValue = '';
  let useAutoRepresentative = false;

  if (def.key_type === '他シートセル') {
    const srcSheet = findSheetFlexible_(ss, def.key_sheet);
    if (!srcSheet) return null;
    if (!def.key_cell) return null;
    searchValue = srcSheet.getRange(def.key_cell).getDisplayValue();
  } else if (def.key_type === '固定値') {
    searchValue = def.key_value;
  } else if (def.key_type === '要設定' || def.key_type === '') {
    useAutoRepresentative = true;
  } else {
    searchValue = def.key_value;
  }

  // 通常検索
  if (!useAutoRepresentative) {
    if (isBlank_(searchValue)) return null;

    for (let r = 0; r < values.length; r++) {
      const cell = String(values[r][matchCol] || '').trim();
      if (!cell) continue;

      if (isTextMatchJa_(cell, searchValue, def.match_mode)) {
        return getMultiColValue_(values[r], def.value_col_expr);
      }
    }
    return null;
  }

  // 要設定時の自動代表者行検出
  const repRow = findRepresentativeRowForLookup_(values);
  if (repRow < 0) return null;

  return getMultiColValue_(values[repRow], def.value_col_expr);
}

function findRepresentativeRowForLookup_(values) {
  if (!values || values.length === 0) return -1;

  // ヘッダー行っぽいところを飛ばして 5行目以降を優先
  const startRow = Math.min(4, values.length - 1);

  let foundByHonin = -1;
  let foundByGaito = -1;
  let foundByDaihyo = -1;

  for (let r = startRow; r < values.length; r++) {
    const row = values[r] || [];
    const joined = row.map(v => normalizeText_(v)).join('|');
    if (!joined) continue;

    // 明細行らしさがない行は飛ばす
    const hasAnyText = row.some(v => String(v || '').trim() !== '');
    if (!hasAnyText) continue;

    if (joined.includes('本人') && foundByHonin < 0) {
      foundByHonin = r;
    }
    if (joined.includes('該当') && foundByGaito < 0) {
      foundByGaito = r;
    }
    if (joined.includes('代表') && foundByDaihyo < 0) {
      foundByDaihyo = r;
    }
  }

  if (foundByHonin >= 0) return foundByHonin;
  if (foundByGaito >= 0) return foundByGaito;
  if (foundByDaihyo >= 0) return foundByDaihyo;

  return -1;
}

function isTextMatchJa_(cell, target, mode) {
  const a = normalizeText_(cell);
  const b = normalizeText_(target);
  const m = String(mode || '完全一致').trim();

  if (!a || !b) return false;
  if (m === '完全一致') return a === b;
  if (m === '部分一致') return a.includes(b);
  if (m === '前方一致') return a.startsWith(b);
  if (m === '後方一致') return a.endsWith(b);
  return a === b;
}

function getMultiColValue_(row, colExpr) {
  if (!colExpr) return null;
  let expr = String(colExpr).trim();
  if (!expr) return null;

  let roundDownDigits = null;
  const roundDownMatch = expr.match(/^ROUNDDOWN\((.+),\s*(-?\d+)\)$/i);
  if (roundDownMatch) {
    expr = roundDownMatch[1];
    roundDownDigits = Number(roundDownMatch[2]);
  }

  const value = evaluateRowMathExpr_(row, expr);
  if (value == null) return null;

  if (roundDownDigits == null) return value;

  const factor = Math.pow(10, -roundDownDigits);
  return Math.floor(value / factor) * factor;
}

function evaluateRowMathExpr_(row, expr) {
  const replaced = String(expr)
    .toUpperCase()
    .replace(/[A-Z]+/g, col => {
      const idx = colToIndex_(col);
      const v = idx >= 0 ? toNumber_(row[idx]) : 0;
      return String(v == null ? 0 : v);
    });

  if (!/^[0-9+\-*/().\s]+$/.test(replaced)) return null;

  try {
    return Number(Function(`"use strict"; return (${replaced});`)());
  } catch (e) {
    return null;
  }
}

/* =========================
 * 決算書式
 * ========================= */

function resolveDecisionExprSimple_(ctx, expr, matchMode) {
  const text = String(expr || '').trim();
  if (!text) return null;

  const tokens = text.match(/[+-]?[^+-]+/g);
  if (!tokens) return null;

  let total = 0;
  let foundAny = false;

  for (const token of tokens) {
    const sign = token.startsWith('-') ? -1 : 1;
    const name = token.replace(/^[+-]/, '').trim();
    if (!name) continue;

    let value = null;

    if (ctx.groups[name]) {
      value = 0;
      let foundGroup = false;

      ctx.groups[name].forEach(acc => {
        const n = findDecisionValueByMode_(ctx.decisionMap, acc, matchMode);
        if (n != null) {
          value += n;
          foundGroup = true;
          markDecisionAccountUsed_(ctx, acc);
        }
      });

      if (!foundGroup) value = null;
    } else {
      value = findDecisionValueByMode_(ctx.decisionMap, name, matchMode);
      if (value != null) markDecisionAccountUsed_(ctx, name);
    }

    if (value != null) {
      total += sign * value;
      foundAny = true;
    }
  }

  return foundAny ? total : null;
}

function findDecisionValueByMode_(decisionMap, name, matchMode) {
  const target = normalizeText_(name);
  const mode = normalizeMatchMode_(matchMode);

  if (mode === 'exact') {
    return decisionMap[target] != null ? decisionMap[target] : null;
  }

  let sum = 0;
  let found = false;

  Object.keys(decisionMap).forEach(k => {
    let hit = false;
    if (mode === 'contains') hit = k.includes(target);
    else if (mode === 'prefix') hit = k.startsWith(target);
    else if (mode === 'suffix') hit = k.endsWith(target);
    else hit = (k === target);

    if (hit && decisionMap[k] != null) {
      sum += decisionMap[k];
      found = true;
    }
  });

  return found ? sum : null;
}

function resolveAccountExpressionFromMap_(expr, valueMap) {
  const text = String(expr || '').trim();
  if (!text) return null;

  const tokens = text.match(/[+-]?[^+-]+/g);
  if (!tokens) return null;

  let total = 0;
  let foundAny = false;

  for (const token of tokens) {
    const sign = token.startsWith('-') ? -1 : 1;
    const name = normalizeText_(token.replace(/^[+-]/, ''));
    if (!name) continue;

    if (valueMap[name] != null) {
      total += sign * valueMap[name];
      foundAny = true;
    }
  }

  return foundAny ? total : null;
}

function markAccountsUsedFromExpr_(ctx, expr) {
  const text = String(expr || '').trim();
  if (!text) return;
  const tokens = text.match(/[+-]?[^+-]+/g);
  if (!tokens) return;

  tokens.forEach(token => {
    const name = normalizeText_(token.replace(/^[+-]/, ''));
    if (!name) return;
    if (ctx.groups[name]) {
      ctx.groups[name].forEach(acc => markDecisionAccountUsed_(ctx, acc));
    } else {
      markDecisionAccountUsed_(ctx, name);
    }
  });
}

/* =========================
 * 決算書解析
 * ========================= */

function parseDecisionSheet_(sheet) {
  const values = sheet.getDataRange().getDisplayValues();
  const map = {};
  const bsAccounts = [];
  let inBsRange = false;
  let bsRangeFinished = false;

  for (let r = 0; r < values.length; r++) {
    const account = normalizeText_(values[r][0]);
    if (!account) continue;

    let amount = null;
    for (let c = values[r].length - 1; c >= 1; c--) {
      const n = toNumber_(values[r][c]);
      if (n != null) {
        amount = n;
        break;
      }
    }

    if (amount != null) {
      map[account] = amount;

      if (!bsRangeFinished) {
        inBsRange = true;
      }

      if (inBsRange && !isSummaryAccountForBsUnused_(account)) {
        bsAccounts.push(account);
      }

      if (account === normalizeText_('固定負債合計')) {
        bsRangeFinished = true;
        inBsRange = false;
      }
    }
  }

  return { map, bsAccounts };
}

function isSummaryAccountForBsUnused_(accountName) {
  const s = normalizeText_(accountName);
  if (!s) return false;
  return /合計$/.test(s);
}

/* =========================
 * AIチェック（ラフ版）
 * ========================= */

function runGlobalInputAnomalyChecks_(ss) {
  const results = [];
  const apiKey = getGeminiApiKey_();
  if (!apiKey) return results;

  const ignoreSheets = new Set([
    CONFIG.SHEET_RULES,
    CONFIG.SHEET_LOOKUP,
    CONFIG.SHEET_GROUP_MASTER,
    CONFIG.SHEET_NORMALIZE_MASTER,
    CONFIG.SHEET_EXCLUDE_MASTER,
    CONFIG.SHEET_AI_TARGETS,
    CONFIG.SHEET_RESULT,
    CONFIG.SHEET_LOG,
    CONFIG.SHEET_REPORT,
    CONFIG.SHEET_ACCOUNT_MATCH_LOG,
    CONFIG.SHEET_AI_CHECK_LOG
  ]);

  const aiTargets = loadAiCheckTargets_(ss);
  const hasTargetConfig = Object.keys(aiTargets).length > 0;
  const aiCache = {};
  const aiStats = {
    checked: 0,
    apiErrors: 0,
    errorSamples: []
  };
  const aiAuditRows = [];

  ss.getSheets().forEach(sheet => {
    const sheetName = sheet.getName();
    if (ignoreSheets.has(sheetName)) return;

    const targets = resolveAiTargetsForSheet_(aiTargets, sheetName);
    if (hasTargetConfig) {
      if (!targets.length) return;
    } else {
      if (!normalizeText_(sheetName).includes('の内訳')) return;
    }

    const values = sheet.getDataRange().getDisplayValues();
    const ranges = hasTargetConfig
      ? targets.filter(t => t.enabled)
      : [{ startCol: 2, endCol: 21 }];

    if (!ranges.length) return;

    const visited = new Set();

    for (const range of ranges) {
      for (let r = CONFIG.AI_CHECK_START_ROW - 1; r < values.length; r++) {
        for (let c = range.startCol - 1; c <= range.endCol - 1; c++) {
          const key = `${r}:${c}`;
          if (visited.has(key)) continue;
          visited.add(key);

          const text = String(values[r][c] || '').trim();
          if (!text) continue;
          if (/^[0-9,\.\-△()\/年月日円千百]+$/.test(text)) continue;
          if (text.length <= 1) continue;
          if (isAccountingTerm_(text)) continue;
          aiStats.checked++;

          let ai = aiCache[text];
          const fromCache = !!ai;
          if (!ai) {
            ai = callGeminiBreakdownCheckSafe_(text, sheetName, aiStats);
            aiCache[text] = ai;
            Utilities.sleep(120);
          }

          const a1 = toA1_(r + 1, c + 1);
          aiAuditRows.push([
            new Date(),
            sheetName,
            a1,
            text,
            fromCache ? 'CACHE' : 'API',
            ai && ai.is_suspicious ? '不自然' : '問題なし',
            ai && ai.reason ? ai.reason : ''
          ]);

          if (!ai || !ai.is_suspicious) continue;
          if (!shouldKeepAiFinding_(text, ai)) continue;

          const jumpUrl = buildRangeUrl_(ss, sheetName, a1);

          results.push(makeResult_({
            status: '要確認',
            category: '入力値チェック',
            ruleId: 'AI001',
            sheetName: sheetName,
            itemName: `${a1}: ${text}`,
            targetCell: a1,
            jumpUrl: jumpUrl,
            decisionValue: '',
            compareValue: text,
            diff: '',
            condition: 'BREAKDOWN_TEXT_GEMINI_CHECK',
            message: ai.reason || '不自然な入力の可能性があります',
            detail: '',
            aiJudge: ai.is_suspicious ? '不自然' : '問題なし',
            aiReason: ai.reason || '',
            aiSuggestion1: ai.suggestions && ai.suggestions[0] ? ai.suggestions[0] : '',
            aiSuggestion2: ai.suggestions && ai.suggestions[1] ? ai.suggestions[1] : '',
          }));
        }
      }
    }
  });

  appendLogRow_(
    ss,
    `AI入力値チェック: チェック件数=${aiStats.checked}, 指摘件数=${results.length}, APIエラー=${aiStats.apiErrors}`
  );
  if (aiStats.errorSamples.length) {
    aiStats.errorSamples.forEach((msg, i) => {
      appendLogRow_(ss, `AI APIエラー詳細(${i + 1}): ${msg}`);
    });
  }
  writeAiCheckLog_(ss, aiAuditRows);

  return results;
}

function writeAiCheckLog_(ss, rows) {
  const sheet = ensureSheetInSpreadsheet_(ss, CONFIG.SHEET_AI_CHECK_LOG);
  sheet.clearContents().clearFormats();
  sheet.getRange(1, 1, 1, 7).setValues([[
    '日時',
    'シート名',
    'セル',
    '入力値',
    '取得元',
    'AI判定',
    'AI理由'
  ]]);

  if (!rows || rows.length === 0) return;
  sheet.getRange(2, 1, rows.length, 7).setValues(rows);
}

function isAccountingTerm_(text) {
  const t = normalizeText_(text);
  const keywords = [
    '預り金','仮受金','仮払金','前払費用','前払金','前渡金',
    '未払金','未払費用','売掛金','買掛金','貸付金','借入金',
    '減価償却費','仕入高','役員給与','役員報酬','役員賞与',
    '役員退職金','地代家賃','保険積立金','預託金','立替金',
    '雑給','給与','雑収入','雑損失','敷金','保証金',
    '源泉所得税','住民税','普通預金','当座預金','定期預金',
    '普通','当座','定期','社宅','事務所','家賃','常勤','非常勤',
    '本人','該当','その他','登録番号'
  ];
  return keywords.some(k => t === normalizeText_(k));
}

function loadAiCheckTargets_(ss) {
  const sheet = ss.getSheetByName(CONFIG.SHEET_AI_TARGETS);
  if (!sheet) return {};

  const rows = loadRowsAsObjectsJapanese_(sheet);
  const map = {};

  rows.forEach(r => {
    const sheetName = String(r['対象シート名'] || '').trim();
    if (!sheetName) return;

    if (!map[sheetName]) map[sheetName] = [];
    map[sheetName].push({
      enabled: toBooleanJa_(r['有効']),
      startCol: parseColRefTo1Based_(r['開始列']) || 2,
      endCol: parseColRefTo1Based_(r['終了列']) || 21,
    });
  });

  return map;
}

function resolveAiTargetsForSheet_(aiTargets, sheetName) {
  const out = [];
  const normSheet = normalizeText_(sheetName);

  Object.keys(aiTargets).forEach(k => {
    if (normalizeText_(k) && normSheet.includes(normalizeText_(k))) {
      out.push(...aiTargets[k]);
    }
  });

  return out;
}

function parseColRefTo1Based_(v) {
  const s = String(v || '').trim().toUpperCase();
  if (!s) return null;
  if (/^\d+$/.test(s)) return Number(s);
  const idx = colToIndex_(s);
  return idx >= 0 ? idx + 1 : null;
}

function shouldKeepAiFinding_(text, ai) {
  const t = String(text || '').trim();
  const reason = String((ai && ai.reason) || '');
  const norm = normalizeText_(t);

  if (!t) return false;
  if (isAccountingTerm_(t)) return false;
  if (norm.length <= 2) return false;
  if (/^T\d{13}$/i.test(norm)) return false;
  if (/^(令和|平成|昭和)\d+年\d+月\d+日$/u.test(norm)) return false;
  if (/^[0-9０-９\-－ー\/／.．]+$/.test(norm)) return false;

  const blockedReasonSignals = [
    /情報不足/u,
    /正式名称不足/u,
    /法人格/u,
    /不完全/u,
    /内容が足りない/u,
    /説明不足/u
  ];

  const textLooksSuspicious =
    /�/.test(t) ||
    /[Ａ-Ｚａ-ｚ]{4,}/.test(t) ||
    /(.)\1{2,}/.test(norm);

  if (textLooksSuspicious) return true;
  if (!reason) return false;
  if (blockedReasonSignals.some(re => re.test(reason))) return false;
  return reason.length >= 4;
}

function getGeminiApiKey_() {
  return PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || '';
}

function getGeminiModel_() {
  const raw = PropertiesService.getScriptProperties().getProperty('GEMINI_MODEL') || 'gemini-2.5-flash';
  return String(raw).replace(/^models\//, '');
}

function callGeminiBreakdownCheck_(text, sheetName) {
  const apiKey = getGeminiApiKey_();
  if (!apiKey) {
    return {
      is_suspicious: false,
      reason: '',
      suggestions: []
    };
  }

  const model = getGeminiModel_();
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;

  const prompt = `
あなたは税務申告書の内訳書チェック支援AIです。
対象は「○○の内訳」シートなどの入力値です。

次の文字列について、
「誤字・脱字・文字化け・存在しない地名の可能性・存在しないブランド名の可能性・明らかな不自然表記」
だけを中心に判定してください。

次の理由では原則として指摘しないでください。
- 情報不足
- 正式名称不足
- 法人格（株式会社など）がない
- 銀行名だけ、支店名だけ、住所だけ等の不完全さ
- 勘定科目として妥当かどうか
- 内訳として内容が足りないという理由だけ
- 一般的な略称や通称
- 単なる説明不足

特に残したいもの：
- 明らかな誤字脱字
- 会社名や氏名が途中で切れている
- 文字化け
- 不自然な空白や重複
- 実在しない地名の可能性
- 実在しないブランド名・社名の可能性

JSONだけ返してください。

シート名: ${sheetName}
入力文字列: ${text}

返却形式:
{
  "is_suspicious": true,
  "reason": "理由",
  "suggestions": ["候補1", "候補2"]
}
`.trim();

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      temperature: 0.1,
      responseMimeType: 'application/json'
    }
  };

  const maxAttempts = 4;
  const baseSleepMs = 800;

  let lastCode = 0;
  let lastBody = '';

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const code = res.getResponseCode();
    const body = res.getContentText();
    lastCode = code;
    lastBody = body;

    if (code >= 200 && code < 300) {
      const json = JSON.parse(body);
      const textOut = json.candidates?.[0]?.content?.parts?.[0]?.text;
      if (!textOut) throw new Error('Gemini応答が空です');

      const parsed = parseGeminiJsonResponse_(textOut);

      return {
        is_suspicious: !!parsed.is_suspicious,
        reason: parsed.reason || '',
        suggestions: Array.isArray(parsed.suggestions) ? parsed.suggestions.slice(0, 2) : []
      };
    }

    if (!shouldRetryGeminiError_(code, body) || attempt === maxAttempts) {
      break;
    }

    const sleepMs = baseSleepMs * Math.pow(2, attempt - 1) + Math.floor(Math.random() * 250);
    Utilities.sleep(sleepMs);
  }

  throw new Error(`Gemini API error ${lastCode}: ${summarizeGeminiError_(lastBody)}`);
}

function parseGeminiJsonResponse_(textOut) {
  const s = String(textOut || '').trim();
  if (!s) throw new Error('Gemini応答JSONが空です');

  try {
    return JSON.parse(s);
  } catch (e) {
    const fenced = s.match(/```(?:json)?\s*([\s\S]*?)\s*```/i);
    if (fenced && fenced[1]) {
      return JSON.parse(fenced[1]);
    }
    const firstObj = s.match(/\{[\s\S]*\}/);
    if (firstObj && firstObj[0]) {
      return JSON.parse(firstObj[0]);
    }
    throw e;
  }
}

function callGeminiBreakdownCheckSafe_(text, sheetName, aiStats) {
  try {
    return callGeminiBreakdownCheck_(text, sheetName);
  } catch (e) {
    if (aiStats) {
      aiStats.apiErrors = (aiStats.apiErrors || 0) + 1;
      if ((aiStats.errorSamples || []).length < 3) {
        aiStats.errorSamples.push(String(e && e.message ? e.message : e));
      }
    }
    return {
      is_suspicious: false,
      reason: '',
      suggestions: []
    };
  }
}

function shouldRetryGeminiError_(code, body) {
  if (code === 429 || code === 503 || code === 504) return true;
  const msg = normalizeText_(String(body || ''));
  return (
    msg.includes(normalizeText_('帯域幅の上限')) ||
    msg.includes(normalizeText_('resource_exhausted')) ||
    msg.includes(normalizeText_('rate limit')) ||
    msg.includes(normalizeText_('quota'))
  );
}

function summarizeGeminiError_(body) {
  const text = String(body || '');
  const noApiKey = text.replace(/key=[^&"\s]+/gi, 'key=***');

  try {
    const json = JSON.parse(noApiKey);
    const msg = json?.error?.message || json?.message || noApiKey;
    return String(msg).slice(0, 400);
  } catch (e) {
    return noApiKey.slice(0, 400);
  }
}

/* =========================
 * 未照合BS科目
 * ========================= */

function buildUnusedBSAccountResults_(ctx) {
  const results = [];
  const excludedAccounts = new Set([
    normalizeText_('普通預金'),
    normalizeText_('未払消費税等'),
    normalizeText_('未払法人税等')
  ]);

  ctx.bsAccounts.forEach(account => {
    const a = normalizeText_(account);
    if (!a) return;
    if (excludedAccounts.has(a)) return;

    const isUnused = !ctx.usedDecisionAccounts.has(a);
    const decisionValue = toNumber_(ctx.decisionMap[a]);
    const hasAmount = decisionValue != null && decisionValue >= 1;

    if (isUnused && hasAmount) {
      results.push(makeResult_({
        status: '要確認',
        category: '未照合BS科目',
        ruleId: 'BS001',
        sheetName: CONFIG.SHEET_DECISION,
        itemName: a,
        targetCell: '',
        jumpUrl: '',
        decisionValue: decisionValue,
        compareValue: '',
        diff: '',
        condition: 'UNUSED_BS_ACCOUNT',
        message: '決算書のBS科目ですが、今回どの照合にも使用されていません',
        detail: '',
      }));
    }
  });

  return results;
}

/* =========================
 * 結果出力
 * ========================= */

function writeResults_(ss, results) {
  const sheet = getRequiredSheet_(ss, CONFIG.SHEET_RESULT);
  const filtered = (results || []).filter(r => r != null);
  if (!filtered.length) return;

  const rows = filtered.map(r => ([
    r.status || '',
    r.category || '',
    r.ruleId || '',
    r.sheetName || '',
    r.itemName || '',
    r.targetCell || '',
    r.jumpUrl || '',
    r.decisionValue === undefined ? '' : r.decisionValue,
    r.compareValue === undefined ? '' : r.compareValue,
    r.diff === undefined ? '' : r.diff,
    r.condition || '',
    r.message || '',
    r.detail || '',
    r.aiJudge || '',
    r.aiReason || '',
    r.aiSuggestion1 || '',
    r.aiSuggestion2 || '',
    Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'),
  ]));

  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

function writeAccountMatchLog_(ss, ctx) {
  const sheet = ensureSheetInSpreadsheet_(ss, CONFIG.SHEET_ACCOUNT_MATCH_LOG);
  sheet.clearContents().clearFormats();
  sheet.getRange(1, 1, 1, 6).setValues([[
    'No',
    '勘定科目',
    '決算書値',
    '照合利用有無',
    '対象レンジ内',
    '備考'
  ]]);

  const rows = (ctx.bsAccounts || []).map((account, idx) => {
    const key = normalizeText_(account);
    const used = key ? ctx.usedDecisionAccounts.has(key) : false;
    return [
      idx + 1,
      key || '',
      key ? (ctx.decisionMap[key] != null ? ctx.decisionMap[key] : '') : '',
      used ? '照合済み' : '未照合',
      '対象',
      used ? '' : '今回のルールで未使用'
    ];
  });

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 6).setValues(rows);
  }
}

function buildA4Report_(ss) {
  const src = getRequiredSheet_(ss, CONFIG.SHEET_RESULT);
  const values = src.getDataRange().getValues();
  if (values.length < 2) return;

  const header = values[0];
  const rows = values.slice(1);

  let report = ss.getSheetByName(CONFIG.SHEET_REPORT);
  if (report) ss.deleteSheet(report);
  report = ss.insertSheet(CONFIG.SHEET_REPORT);

  const colIndex = {};
  header.forEach((h, i) => colIndex[h] = i);

  const picked = rows.filter(r => {
    const status = String(r[colIndex['判定']] || '');
    if (status === 'OK') return false;

    const decisionRaw = r[colIndex['決算書値']];
    const compareRaw = r[colIndex['比較値']];

    if (isBlank_(decisionRaw) && isBlank_(compareRaw)) return false;

    const decisionValue = toNumber_(decisionRaw);
    const compareValue = toNumber_(compareRaw);
    const bothZero = decisionValue === 0 && compareValue === 0;

    return !bothZero;
  });

  const out = [[
    '判定',
    '区分',
    '対象シート',
    '対象項目',
    '対象セル',
    'メッセージ',
    '比較値',
    '決算書値',
    '差額',
    'AI理由',
    'ジャンプ'
  ]];

  picked.forEach(r => {
    out.push([
      r[colIndex['判定']],
      r[colIndex['区分']],
      r[colIndex['対象シート']],
      r[colIndex['対象項目']],
      r[colIndex['対象セル']],
      r[colIndex['メッセージ']],
      r[colIndex['比較値']],
      r[colIndex['決算書値']],
      r[colIndex['差額']],
      r[colIndex['AI理由']],
      r[colIndex['ジャンプURL']] ? '開く' : ''
    ]);
  });

  report.getRange(1, 1, out.length, out[0].length).setValues(out);

  for (let i = 2; i <= out.length; i++) {
    const srcRow = picked[i - 2];
    const jumpUrl = srcRow[colIndex['ジャンプURL']];
    if (jumpUrl) {
      const rich = SpreadsheetApp.newRichTextValue()
        .setText('開く')
        .setLinkUrl(jumpUrl)
        .build();
      report.getRange(i, 11).setRichTextValue(rich);
    }
  }

  report.getRange(1, 1, 1, 11).setFontWeight('bold').setBackground('#d9ead3');
  report.setFrozenRows(1);
  report.setColumnWidths(1, 11, 120);
  report.setColumnWidth(4, 240);
  report.setColumnWidth(6, 320);
  report.setColumnWidth(10, 240);
  report.setColumnWidth(11, 80);

  if (report.getFilter()) report.getFilter().remove();
  report.getDataRange().createFilter();
  report.setHiddenGridlines(true);
}

/* =========================
 * 比較結果
 * ========================= */
function compareNumbersResult_(rule, sheetName, itemName, expected, actual, category, targetCell, jumpUrl) {
  const n1 = toNumber_(expected);
  const n2 = toNumber_(actual);

  const isBlank1 = isBlank_(expected);
  const isBlank2 = isBlank_(actual);

  const isZero1 = n1 === 0;
  const isZero2 = n2 === 0;

  if (
    (isBlank1 && isBlank2) ||
    (isZero1 && isBlank2) ||
    (isBlank1 && isZero2) ||
    (isZero1 && isZero2)
  ) {
    return null;
  }

  if (n1 == null || n2 == null) {
    return makeResult_({
      status: rule.severity || '要確認',
      category: category || '',
      ruleId: rule.rule_id,
      sheetName: sheetName || '',
      itemName: itemName || '',
      targetCell: targetCell || '',
      jumpUrl: jumpUrl || '',
      decisionValue: expected == null ? '' : expected,
      compareValue: actual == null ? '' : actual,
      diff: '',
      condition: rule.check_type || 'VALUE_MATCH',
      message: rule.message || '比較値または参照値を取得できませんでした',
      detail: '',
    });
  }

  const diff = n2 - n1;
  const ok = Math.abs(diff) < 0.000001;

  return makeResult_({
    status: ok ? 'OK' : (rule.severity || '要確認'),
    category: category || '',
    ruleId: rule.rule_id,
    sheetName: sheetName || '',
    itemName: itemName || '',
    targetCell: targetCell || '',
    jumpUrl: jumpUrl || '',
    decisionValue: n1,
    compareValue: n2,
    diff: diff,
    condition: rule.check_type || 'VALUE_MATCH',
    message: ok ? 'OK' : (rule.message || '値が一致しません'),
    detail: '',
  });
}

function compareTextsResult_(rule, sheetName, itemName, expected, actual, category, targetCell, jumpUrl) {
  const ok = String(expected || '').trim() === String(actual || '').trim();

  return makeResult_({
    status: ok ? 'OK' : (rule.severity || '要確認'),
    category: category || '',
    ruleId: rule.rule_id,
    sheetName: sheetName || '',
    itemName: itemName || '',
    targetCell: targetCell || '',
    jumpUrl: jumpUrl || '',
    decisionValue: expected || '',
    compareValue: actual || '',
    diff: '',
    condition: rule.check_type || 'TEXT_MATCH',
    message: ok ? 'OK' : (rule.message || '文字列が一致しません'),
    detail: '',
  });
}

function makeResult_(obj) {
  return {
    status: obj.status || '',
    category: obj.category || '',
    ruleId: obj.ruleId || '',
    sheetName: obj.sheetName || '',
    itemName: obj.itemName || '',
    targetCell: obj.targetCell || '',
    jumpUrl: obj.jumpUrl || '',
    decisionValue: obj.decisionValue,
    compareValue: obj.compareValue,
    diff: obj.diff,
    condition: obj.condition || '',
    message: obj.message || '',
    detail: obj.detail || '',
    aiJudge: obj.aiJudge || '',
    aiReason: obj.aiReason || '',
    aiSuggestion1: obj.aiSuggestion1 || '',
    aiSuggestion2: obj.aiSuggestion2 || '',
  };
}

/* =========================
 * 汎用ユーティリティ
 * ========================= */

function loadRowsAsObjectsJapanese_(sheet) {
  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) return [];

  const headers = values[0].map(v => String(v || '').trim());
  const rows = [];

  for (let r = 1; r < values.length; r++) {
    const row = {};
    let hasValue = false;

    for (let c = 0; c < headers.length; c++) {
      const key = headers[c];
      if (!key) continue;
      row[key] = values[r][c];
      if (String(values[r][c] || '').trim() !== '') hasValue = true;
    }

    if (hasValue) rows.push(row);
  }
  return rows;
}

function normalizeCheckTypeJa_(v) {
  const s = String(v || '').trim();
  const map = {
    '数値照合': 'VALUE_MATCH',
    '空欄チェック': 'NOT_BLANK',
    '条件付空欄チェック': 'CONDITIONAL_NOT_BLANK',
    '文字列照合': 'TEXT_MATCH',
    '見出し確認': 'HEADER_EXISTS',
    '固定値一致': 'FIXED_MATCH',
    '存在確認': 'EXISTS',
    '検索一致': 'VALUE_MATCH',
  };
  return map[s] || s;
}

function normalizeMatchMode_(v) {
  const s = String(v || '').trim();
  if (s === '完全一致') return 'exact';
  if (s === '部分一致') return 'contains';
  if (s === '前方一致') return 'prefix';
  if (s === '後方一致') return 'suffix';
  return 'exact';
}

function toBooleanJa_(v) {
  const s = String(v == null ? '' : v).trim().toUpperCase();
  return ['TRUE', 'T', '1', 'YES', 'Y', '有効', 'ON'].includes(s);
}

function normalizeText_(v) {
  return String(v == null ? '' : v)
    .replace(/\s+/g, '')
    .replace(/[　]/g, '')
    .trim();
}

function toNumber_(v) {
  if (v == null || v === '') return null;
  let s = String(v).trim();
  if (!s) return null;

  let negative = false;
  if (/^\(.*\)$/.test(s)) {
    negative = true;
    s = s.replace(/^\(|\)$/g, '');
  }

  s = s
    .replace(/,/g, '')
    .replace(/△/g, '-')
    .replace(/▲/g, '-')
    .replace(/円|千円|百万円/g, '');

  if (!/^-?\d+(\.\d+)?$/.test(s)) return null;
  let n = Number(s);
  if (isNaN(n)) return null;
  if (negative) n = -Math.abs(n);
  return n;
}

function isNumericLike_(v) {
  return toNumber_(v) != null;
}

function isBlank_(v) {
  return v == null || String(v).trim() === '';
}

function ensureSheetInSpreadsheet_(ss, sheetName) {
  let sh = ss.getSheetByName(sheetName);
  if (!sh) sh = ss.insertSheet(sheetName);
  return sh;
}

function getRequiredSheet_(ss, sheetName) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error(`必要シートが見つかりません: ${sheetName}`);
  return sh;
}

function appendLogRow_(ss, message) {
  const sheet = ensureSheetInSpreadsheet_(ss, CONFIG.SHEET_LOG);
  sheet.appendRow([
    Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'),
    message
  ]);
}

function findSheetFlexible_(ss, pattern) {
  const p = String(pattern || '').trim();
  if (!p) return null;

  const sheets = ss.getSheets();

  for (const sh of sheets) {
    if (sh.getName() === p) return sh;
  }
  for (const sh of sheets) {
    if (sh.getName().includes(p)) return sh;
  }
  return null;
}

function findCellByText_(values, targetText) {
  const target = normalizeText_(targetText);
  if (!target) return null;

  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      if (normalizeText_(values[r][c]) === target) {
        return { row: r, col: c };
      }
    }
  }

  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      if (normalizeText_(values[r][c]).includes(target)) {
        return { row: r, col: c };
      }
    }
  }

  return null;
}

function sumColumnBelowHeader_(values, headerRow, amountCol, rowCond) {
  let sum = 0;

  for (let r = headerRow + 1; r < values.length; r++) {
    const row = values[r];
    if (!row) continue;

    if (rowCond && !isRowMatchCondition_(row, rowCond, values, r)) continue;

    const n = toNumber_(row[amountCol]);
    if (n == null) continue;
    sum += n;
  }

  return sum;
}

function isRowMatchCondition_(row, rowCond, values, rowIndex) {
  const s = String(rowCond || '').trim();
  if (!s) return true;

  const m1 = s.match(/^([A-Z]+)列\s*<>\s*空欄$/i);
  if (m1) {
    const col = colToIndex_(m1[1]);
    return normalizeText_(row[col]) !== '';
  }

  const m2 = s.match(/^([A-Z]+)列\s*=\s*(.+)$/i);
  if (m2) {
    const col = colToIndex_(m2[1]);
    return normalizeText_(row[col]) === normalizeText_(m2[2]);
  }

  return true;
}

function findRightTextValue_(values, row, startCol) {
  for (let c = startCol + 1; c < values[row].length; c++) {
    const v = String(values[row][c] || '').trim();
    if (v !== '') return v;
  }
  return '';
}

function evaluateSimpleCondition_(expr, ctx) {
  const text = String(expr || '').trim();
  if (!text) return false;

  const m = text.match(/^(.+)\s*>\s*0$/);
  if (m) {
    const value = resolveDecisionExprSimple_(ctx, m[1], 'exact');
    return toNumber_(value) != null && toNumber_(value) > 0;
  }

  const m2 = text.match(/^(.+)\s*>=\s*0$/);
  if (m2) {
    const value = resolveDecisionExprSimple_(ctx, m2[1], 'exact');
    return toNumber_(value) != null && toNumber_(value) >= 0;
  }

  return false;
}

function evaluateCellExpression_(expr, values) {
  const replaced = String(expr || '')
    .toUpperCase()
    .replace(/[A-Z]+\d+/g, ref => {
      const m = ref.match(/^([A-Z]+)(\d+)$/);
      if (!m) return '0';
      const col = colToIndex_(m[1]);
      const row = Number(m[2]) - 1;
      if (row < 0 || col < 0) return '0';
      const v = values[row] && values[row][col] != null ? values[row][col] : '';
      const n = toNumber_(v);
      return String(n == null ? 0 : n);
    });

  if (!/^[0-9+\-*/().\s]+$/.test(replaced)) return null;

  try {
    return Number(Function(`"use strict"; return (${replaced});`)());
  } catch (e) {
    return null;
  }
}

function extractFirstA1FromExpr_(expr) {
  const m = String(expr || '').match(/[A-Z]+\d+/);
  return m ? m[0] : '';
}

function buildExcludeSetForRule_(ruleId, excludeMaster) {
  const set = new Set();
  (excludeMaster || []).forEach(r => {
    if (!r.account) return;
    if (!r.rule_id || r.rule_id === ruleId) {
      set.add(r.account);
    }
  });
  return set;
}

function normalizeAccountByRule_(rawLabel, normalizeRules) {
  const raw = normalizeText_(rawLabel);
  if (!raw) return raw;

  const rules = (normalizeRules || [])
    .slice()
    .sort((a, b) => (b.from || '').length - (a.from || '').length);

  for (let i = 0; i < rules.length; i++) {
    const from = rules[i].from;
    const to = rules[i].to;
    if (!from || !to) continue;

    if (raw === from || raw.includes(from)) {
      return to;
    }
  }

  return raw;
}

function colToIndex_(colRef) {
  const s = String(colRef || '').trim().toUpperCase();
  if (!s) return -1;
  if (!/^[A-Z]+$/.test(s)) return -1;

  let n = 0;
  for (let i = 0; i < s.length; i++) {
    n = n * 26 + (s.charCodeAt(i) - 64);
  }
  return n - 1;
}

function toA1_(row, col) {
  return `${indexToCol_(col - 1)}${row}`;
}

function indexToCol_(index) {
  let n = Number(index) + 1;
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - m) / 26);
  }
  return s;
}

function buildRangeUrl_(ss, sheetName, a1) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || !a1) return '';
  return `${ss.getUrl()}#gid=${sheet.getSheetId()}&range=${encodeURIComponent(a1)}`;
}

function markDecisionAccountUsed_(ctx, account) {
  const key = normalizeText_(account);
  if (key) ctx.usedDecisionAccounts.add(key);
}

function getNowStr_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
}

function escapeHtml_(s) {
  return String(s || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}
