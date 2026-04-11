/*************************************************
 * 決算チェックツール 完全版
 * v2026-04-11.03
 *
 * できること
 * - DriveフォルダIDを受け取り、同じフォルダに結果スプレッドシートを作成
 * - テンプレート（このGASが紐づくスプレッドシート）を丸ごとコピー
 * - フォルダ内の Googleスプレッドシート / Excel を取り込み
 * - 変換に使った一時スプレッドシートと元Excelを old フォルダへ移動（設定で切替）
 * - rules_main / rules_kokyo に基づいて照合
 * - check_result / check_log / report_A4 を出力
 * - 入力値の全般異常チェック + Gemini補助判定
 * - 概況書の MATCH_DECISION_EXPR, account_match_mode, value_pick_rule 対応
 * - MATCH_BREAKDOWN_LOOKUP 対応
 * - lookup_value_col の K+L のような複数列合算対応
 * - 決算書BS主要科目の未照合チェック
 *
 * 注意
 * - PDFのOCR読取はこの版では未実装
 * - 決算書は「決算書」というシート名で結果ブックに入る前提
 * - Advanced Drive API を ON にしてください
 * - スクリプトプロパティ（任意）:
 *     GEMINI_API_KEY = Gemini API Key
 *     AI_CHECK_ENABLED = true / false
 *     MOVE_EXCEL_TO_OLD = true / false
 *     MOVE_CONVERTED_TO_OLD = true / false
 *************************************************/

const CONFIG = {
  TEMPLATE_SHEET_ID: '',
  RESULT_FILE_PREFIX: '決算チェック_',

  SHEET_DECISION: '決算書',
  SHEET_RULES_MAIN: 'rules_main',
  SHEET_RULES_KOKYO: 'rules_kokyo',
  SHEET_GROUP_MASTER: 'account_group_master',
  SHEET_NORMALIZE_MASTER: 'account_normalize_master',
  SHEET_EXCLUDE_MASTER: 'account_exclude_master',
  SHEET_AI_TARGETS: 'ai_check_targets',

  SHEET_KOKYO_FRONT: '概況書表面',
  SHEET_KOKYO_BACK: '概況書裏面',

  SHEET_RESULT: 'check_result',
  SHEET_LOG: 'check_log',
  SHEET_REPORT: 'report_A4',
  AI_CHECK_START_ROW: 5,
};

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
    const result = runFromFolder_(folderId);
    return HtmlService
      .createHtmlOutput(
        `<p>処理が完了しました。</p><p><a href="${result.url}" target="_blank">結果スプレッドシートを開く</a></p>`
      )
      .setTitle('決算チェック');
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

  while (files.hasNext()) {
    const file = files.next();
    const fileId = file.getId();
    const fileName = file.getName();
    const mimeType = file.getMimeType();

    if (fileId === resultSs.getId()) continue;
    if (fileId === getTemplateSpreadsheet_().getId()) continue;
    if (fileName === 'old') continue;

    if (mimeType === MimeType.GOOGLE_SHEETS) {
      importGoogleSpreadsheetFile_(file, resultSs);
      continue;
    }

    if (isExcelMimeType_(mimeType) || /\.xlsx?$/i.test(fileName)) {
      importExcelFile_(file, folder, oldFolder, resultSs);
      continue;
    }

    if (mimeType === MimeType.PDF || /\.pdf$/i.test(fileName)) {
      appendLogRow_(resultSs, `PDFは未取込（OCR未実装）: ${fileName}`);
      continue;
    }
  }
}

function importGoogleSpreadsheetFile_(file, resultSs) {
  const sourceSs = SpreadsheetApp.openById(file.getId());
  copySourceSheets_(sourceSs, resultSs, file.getName());
  appendLogRow_(resultSs, `Googleスプレッドシート取込: ${file.getName()}`);
}

function importExcelFile_(file, folder, oldFolder, resultSs) {
  const moveExcelToOld = getMoveExcelToOldFlag_();
  const moveConvertedToOld = getMoveConvertedToOldFlag_();

  const tempSs = convertExcelToSpreadsheet_(file, folder);

  try {
    copySourceSheets_(tempSs, resultSs, file.getName());
    appendLogRow_(resultSs, `Excel取込: ${file.getName()}`);
  } finally {
    moveFileToOldIfNeeded_(file, oldFolder, moveExcelToOld);
    moveFileToOldIfNeeded_(DriveApp.getFileById(tempSs.getId()), oldFolder, moveConvertedToOld);
  }
}

function convertExcelToSpreadsheet_(file, folder) {
  const blob = file.getBlob();

  const resource = {
    title: `[tmp]${file.getName()}`,
    mimeType: MimeType.GOOGLE_SHEETS,
    parents: [{ id: folder.getId() }],
  };

  const converted = Drive.Files.insert(resource, blob);
  return SpreadsheetApp.openById(converted.id);
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

  if (base.includes('法人事業概況説明書') && base.includes('表面')) {
    return CONFIG.SHEET_KOKYO_FRONT;
  }
  if (base.includes('法人事業概況説明書') && base.includes('裏面')) {
    return CONFIG.SHEET_KOKYO_BACK;
  }
  if (base.includes('決算書')) {
    return CONFIG.SHEET_DECISION;
  }

  let name = multiple ? `${base}_${originalSheetName}` : base;
  name = name.replace(/[\\\/\?\*\[\]:]/g, '_');
  return name.substring(0, 95);
}

function stripExtension_(name) {
  return String(name || '').replace(/\.[^/.]+$/,'');
}

function isExcelMimeType_(mimeType) {
  return [
    MimeType.MICROSOFT_EXCEL,
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'application/vnd.ms-excel',
  ].includes(mimeType);
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

    const mainResults = runMainRules_(ss, ctx);
    const kokyoResults = runKokyoRules_(ss, ctx);
    const aiCheckEnabled = getAiCheckEnabledFlag_();
    const typoResults = aiCheckEnabled ? runGlobalInputAnomalyChecks_(ss) : [];
    if (!aiCheckEnabled) {
      appendLogRow_(ss, 'AI入力値チェックをスキップしました（AI_CHECK_ENABLED=false）');
    }
    const bsUnusedResults = buildUnusedBSAccountResults_(ctx);

    const allResults = [...mainResults, ...kokyoResults, ...typoResults, ...bsUnusedResults];
    writeResults_(ss, allResults);
    buildA4Report_(ss);

    appendLogRow_(ss, `照合完了 件数=${allResults.length}`);
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

  const report = ss.getSheetByName(CONFIG.SHEET_REPORT);
  if (report) ss.deleteSheet(report);
}

function buildContext_(ss) {
  const decisionSheet = getRequiredSheet_(ss, CONFIG.SHEET_DECISION);
  const parsed = parseDecisionSheet_(decisionSheet);

  return {
    ss,
    decisionSheet,
    decisionValues: decisionSheet.getDataRange().getValues(),
    decisionMap: parsed.map,
    bsAccounts: parsed.bsAccounts,
    usedDecisionAccounts: new Set(),
    groups: loadGroupMaster_(ss),
    normalizeRules: loadNormalizeMaster_(ss),
    excludeMaster: loadExcludeMaster_(ss),
    rulesMain: loadRulesMain_(ss),
    rulesKokyo: loadRulesKokyo_(ss),
  };
}

function buildUnusedBSAccountResults_(ctx) {
  const results = [];

  const targetSet = new Set([
    '現金及び預金',
    '売掛金',
    '買掛金',
    '未払金',
    '未払費用',
    '長期借入金',
    '短期借入金',
    '役員借入金',
    '貸付金',
    '短期貸付金',
    '長期貸付金',
    '土地',
    '建物',
    '機械装置',
    '車両',
    '車両運搬具',
    '商品',
    '製品',
    '半製品',
    '仕掛品',
    '原材料',
    '貯蔵品',
    '有価証券',
    '出資金',
    '預託金',
    '仮払金',
    '仮受金',
    '受取手形',
    '支払手形'
  ]);

  ctx.bsAccounts.forEach(account => {
    const a = normalizeText_(account);
    if (!a) return;
    if (!targetSet.has(a)) return;

    if (!ctx.usedDecisionAccounts.has(a)) {
      results.push(makeResult_({
        status: '要確認',
        category: '未照合BS科目',
        ruleId: 'BS001',
        sheetName: CONFIG.SHEET_DECISION,
        itemName: a,
        targetCell: '',
        jumpUrl: '',
        decisionValue: ctx.decisionMap[a] || '',
        compareValue: '',
        diff: '',
        condition: 'UNUSED_BS_ACCOUNT',
        message: '決算書のBS科目ですが、今回どの内訳書照合にも使用されていません',
        detail: '',
      }));
    }
  });

  return results;
}

/* =========================
 * rules_main
 * ========================= */

function runMainRules_(ss, ctx) {
  const results = [];

  for (const rule of ctx.rulesMain) {
    if (!toBoolean_(rule.enabled)) continue;
    const checkType = normalizeCheckType_(rule.check_type || rule.checktype || rule['check type']);

    const targetSheet = findTargetSheetByPattern_(ss, rule.file_pattern);
    if (!targetSheet) {
      results.push(makeResult_({
        status: '要確認',
        category: '内訳書',
        ruleId: rule.rule_id,
        sheetName: '',
        itemName: rule.document_type,
        targetCell: '',
        jumpUrl: '',
        decisionValue: '',
        compareValue: '',
        diff: '',
        condition: 'シート未検出',
        message: `対象シートが見つかりません: ${rule.file_pattern}`,
        detail: '',
      }));
      continue;
    }

    const sheetData = getSheetValues_(targetSheet);
    const rows = sheetData.values;
    const display = sheetData.displayValues;

    try {
      switch (checkType) {
        case 'SUM_MATCH':
          results.push(...checkSumMatch_(rule, targetSheet, rows, ctx));
          break;
        case 'SUM_BY_ACCOUNT':
          results.push(...checkSumByAccount_(rule, targetSheet, rows, ctx));
          break;
        case 'GROUP_SUM':
          results.push(...checkGroupSum_(rule, targetSheet, rows, ctx));
          break;
        case 'FIXED_VALUE':
          results.push(...checkFixedValue_(rule, targetSheet, display, ctx));
          break;
        case 'NOT_BLANK_WHEN_PRESENT':
          results.push(...checkNotBlankWhenPresent_(rule, targetSheet, display, ctx));
          break;
        case 'NOT_BLANK_WHEN_ROW_EXISTS':
          results.push(...checkNotBlankWhenRowExists_(rule, targetSheet, display, ctx));
          break;
        case 'NOT_BLANK_WHEN_AMOUNT_EXISTS':
          results.push(...checkNotBlankWhenAmountExists_(rule, targetSheet, rows, ctx));
          break;
        case 'HEADER_CHECK':
          results.push(...checkHeader_(rule, targetSheet, display, ctx));
          break;
        case 'TITLE_CHECK':
          results.push(...checkTitle_(rule, targetSheet, display, ctx));
          break;
        case 'NORMALIZE_SUM':
          results.push(...checkNormalizeSum_(rule, targetSheet, rows, ctx));
          break;
        default:
          results.push(makeResult_({
            status: '要確認',
            category: '内訳書',
            ruleId: rule.rule_id,
            sheetName: targetSheet.getName(),
            itemName: rule.document_type,
            targetCell: '',
            jumpUrl: '',
            decisionValue: '',
            compareValue: '',
            diff: '',
            condition: checkType || '',
            message: `未対応のcheck_typeです: ${checkType || '(空欄)'}`,
            detail: '',
          }));
      }
    } catch (err) {
      results.push(makeResult_({
        status: '要確認',
        category: '内訳書',
        ruleId: rule.rule_id,
        sheetName: targetSheet.getName(),
        itemName: rule.document_type,
        targetCell: '',
        jumpUrl: '',
        decisionValue: '',
        compareValue: '',
        diff: '',
        condition: '実行エラー',
        message: `ルール実行中にエラー: ${err.message}`,
        detail: err.stack || '',
      }));
    }
  }

  return results;
}

function checkSumMatch_(rule, sheet, rows, ctx) {
  let amountCol = colToIndex_(rule.amount_col);
  let headerRow = findHeaderRowByText_(rows, rule.header_name, amountCol);

  if (headerRow < 0) {
    const found = findCellByText_(sheet.getDataRange().getDisplayValues(), rule.header_name);
    if (found) {
      headerRow = found.row;
      amountCol = found.col;
    }
  }

  if (headerRow < 0) return [ngHeaderNotFound_(rule, sheet)];

  const sum = sumColumnBelowHeader_(rows, headerRow, amountCol);
  const target = resolveDecisionTargetValue_(rule, ctx);
  const a1 = toA1_(headerRow + 1, amountCol + 1);

  return [compareNumbersResult_(rule, sheet.getName(), rule.header_name || rule.target_account || rule.target_account_group, target, sum, '内訳書', a1, buildRangeUrl_(ctx.ss, sheet.getName(), a1))];
}

function checkSumByAccount_(rule, sheet, rows, ctx) {
  const accountCol = colToIndex_(rule.account_col);
  let amountCol = colToIndex_(rule.amount_col);

  let headerRow = findHeaderRowByText_(rows, rule.header_name, amountCol);

  if (normalizeText_(rule.file_pattern).includes('売掛金の内訳')) {
    const accountHeaderRow = findHeaderRowByText_(rows, '種類', accountCol);
    if (accountHeaderRow >= 0) headerRow = accountHeaderRow;
  }

  if (headerRow < 0) {
    const found = findCellByText_(sheet.getDataRange().getDisplayValues(), rule.header_name);
    if (found) {
      headerRow = found.row;
      amountCol = found.col;
    }
  }

  if (headerRow < 0) return [ngHeaderNotFound_(rule, sheet)];

  const exclude = mergeExcludeAccounts_(rule.exclude_accounts, ctx.excludeMaster);
  const sums = {};
  const rowMap = {};

  for (let r = headerRow + 1; r < rows.length; r++) {
    const account = normalizeText_(rows[r][accountCol]);
    if (!account) continue;
    if (isSectionLike_(account)) continue;
    if (exclude.has(account)) continue;

    const amount = toNumber_(rows[r][amountCol]);
    if (amount == null) continue;

    sums[account] = (sums[account] || 0) + amount;
    if (!rowMap[account]) rowMap[account] = r;
  }

  const results = [];
  Object.keys(sums).sort().forEach(account => {
    markDecisionAccountUsed_(ctx, account);

    const a1 = toA1_(rowMap[account] + 1, amountCol + 1);
    const jumpUrl = buildRangeUrl_(ctx.ss, sheet.getName(), a1);
    const decisionValue = ctx.decisionMap[account];

    if (decisionValue == null) {
      results.push(makeResult_({
        status: 'SKIP',
        category: '内訳書',
        ruleId: rule.rule_id,
        sheetName: sheet.getName(),
        itemName: account,
        targetCell: a1,
        jumpUrl: jumpUrl,
        decisionValue: '',
        compareValue: sums[account],
        diff: '',
        condition: '決算書科目なし',
        message: '決算書に対象科目が無いためスキップ',
        detail: '',
      }));
      return;
    }

    results.push(compareNumbersResult_(rule, sheet.getName(), account, decisionValue, sums[account], '内訳書', a1, jumpUrl));
  });

  if (results.length === 0) {
    results.push(makeResult_({
      status: '要確認',
      category: '内訳書',
      ruleId: rule.rule_id,
      sheetName: sheet.getName(),
      itemName: rule.document_type,
      targetCell: '',
      jumpUrl: '',
      decisionValue: '',
      compareValue: '',
      diff: '',
      condition: '比較対象なし',
      message: '決算書と一致判定できる科目が見つかりませんでした',
      detail: '',
    }));
  }

  return results;
}

function checkGroupSum_(rule, sheet, rows, ctx) {
  const displayValues = sheet.getDataRange().getDisplayValues();

  let amountCol = colToIndex_(rule.amount_col);
  let headerRow = findHeaderRowByText_(displayValues, rule.header_name, amountCol);

  if (headerRow < 0) {
    const found = findCellByText_(displayValues, rule.header_name);
    if (found) {
      headerRow = found.row;
      amountCol = found.col;
    }
  }

  if (headerRow < 0) return [ngHeaderNotFound_(rule, sheet)];

  const sum = sumColumnBelowHeader_(rows, headerRow, amountCol);
  const target = resolveDecisionTargetValue_(rule, ctx);
  const a1 = toA1_(headerRow + 1, amountCol + 1);

  return [compareNumbersResult_(rule, sheet.getName(), rule.target_account_group || rule.header_name, target, sum, '内訳書', a1, buildRangeUrl_(ctx.ss, sheet.getName(), a1))];
}

function checkFixedValue_(rule, sheet, displayValues, ctx) {
  const found = findCellByText_(displayValues, rule.header_name);
  if (!found) {
    return [makeResult_({
      status: '要確認',
      category: '内訳書',
      ruleId: rule.rule_id,
      sheetName: sheet.getName(),
      itemName: rule.header_name,
      targetCell: '',
      jumpUrl: '',
      decisionValue: '',
      compareValue: '',
      diff: '',
      condition: '固定項目未検出',
      message: `${rule.header_name} が見つかりません`,
      detail: '',
    })];
  }

  const value = findRightNumericValue_(displayValues, found.row, found.col);
  const target = resolveDecisionTargetValue_(rule, ctx);
  const a1 = toA1_(found.row + 1, found.col + 1);

  return [compareNumbersResult_(rule, sheet.getName(), rule.header_name, target, value, '内訳書', a1, buildRangeUrl_(ctx.ss, sheet.getName(), a1))];
}

function checkNotBlankWhenPresent_(rule, sheet, displayValues, ctx) {
  const found = findCellByText_(displayValues, rule.header_name);
  if (!found) {
    return [makeResult_({
      status: '要確認',
      category: '内訳書',
      ruleId: rule.rule_id,
      sheetName: sheet.getName(),
      itemName: rule.header_name,
      targetCell: '',
      jumpUrl: '',
      decisionValue: '',
      compareValue: '',
      diff: '',
      condition: '固定項目未検出',
      message: `${rule.header_name} が見つかりません`,
      detail: '',
    })];
  }

  const value = findRightTextValue_(displayValues, found.row, found.col);
  const a1 = toA1_(found.row + 1, found.col + 1);

  return [makeResult_({
    status: value ? 'OK' : (rule.severity || '要確認'),
    category: '内訳書',
    ruleId: rule.rule_id,
    sheetName: sheet.getName(),
    itemName: rule.header_name,
    targetCell: a1,
    jumpUrl: buildRangeUrl_(ctx.ss, sheet.getName(), a1),
    decisionValue: '',
    compareValue: value || '',
    diff: '',
    condition: 'NOT_BLANK',
    message: value ? 'OK' : rule.message,
    detail: '',
  })];
}

function checkNotBlankWhenAmountExists_(rule, sheet, rows, ctx) {
  const amountCol = colToIndex_(rule.amount_col);
  const targetBlankCol = colToIndex_('M');

  let foundBlank = false;
  let foundRow = -1;

  for (let r = 0; r < rows.length; r++) {
    const amount = toNumber_(rows[r][amountCol]);
    if (amount != null && amount !== 0) {
      const v = normalizeText_(rows[r][targetBlankCol]);
      if (!v) {
        foundBlank = true;
        foundRow = r;
        break;
      }
    }
  }

  const a1 = foundRow >= 0 ? toA1_(foundRow + 1, targetBlankCol + 1) : '';
  return [makeResult_({
    status: foundBlank ? (rule.severity || '要確認') : 'OK',
    category: '内訳書',
    ruleId: rule.rule_id,
    sheetName: sheet.getName(),
    itemName: rule.document_type,
    targetCell: a1,
    jumpUrl: a1 ? buildRangeUrl_(ctx.ss, sheet.getName(), a1) : '',
    decisionValue: '',
    compareValue: '',
    diff: '',
    condition: rule.condition || '',
    message: foundBlank ? rule.message : 'OK',
    detail: '',
  })];
}

function checkNotBlankWhenRowExists_(rule, sheet, displayValues, ctx) {
  const found = findCellByText_(displayValues, rule.header_name);
  if (!found) {
    return [makeResult_({
      status: '要確認',
      category: '内訳書',
      ruleId: rule.rule_id,
      sheetName: sheet.getName(),
      itemName: rule.header_name,
      targetCell: '',
      jumpUrl: '',
      decisionValue: '',
      compareValue: '',
      diff: '',
      condition: '固定項目未検出',
      message: `${rule.header_name} が見つかりません`,
      detail: '',
    })];
  }

  const targetCol = found.col;
  const startRow = found.row + 1;
  const triggerCols = parseColumnList_(rule.account_col || 'B');

  let foundBlank = false;
  let foundRow = -1;
  for (let r = startRow; r < displayValues.length; r++) {
    const hasTrigger = triggerCols.some(col => normalizeText_(displayValues[r][col]) !== '');
    if (!hasTrigger) continue;

    const targetText = normalizeText_(displayValues[r][targetCol]);
    if (!targetText) {
      foundBlank = true;
      foundRow = r;
      break;
    }
  }

  const a1 = foundBlank ? toA1_(foundRow + 1, targetCol + 1) : '';
  return [makeResult_({
    status: foundBlank ? (rule.severity || '要確認') : 'OK',
    category: '内訳書',
    ruleId: rule.rule_id,
    sheetName: sheet.getName(),
    itemName: rule.header_name,
    targetCell: a1,
    jumpUrl: a1 ? buildRangeUrl_(ctx.ss, sheet.getName(), a1) : '',
    decisionValue: '',
    compareValue: '',
    diff: '',
    condition: 'NOT_BLANK_WHEN_ROW_EXISTS',
    message: foundBlank ? rule.message : 'OK',
    detail: '',
  })];
}

function parseColumnList_(expr) {
  const raw = String(expr || '').trim();
  if (!raw) return [1]; // B
  const parts = raw.split(/[+,]/).map(s => s.trim()).filter(Boolean);
  const cols = parts
    .map(p => colToIndex_(p))
    .filter(i => i >= 0);
  return cols.length ? cols : [1];
}

function checkHeader_(rule, sheet, displayValues, ctx) {
  const found = findCellByText_(displayValues, rule.header_name);
  const a1 = found ? toA1_(found.row + 1, found.col + 1) : '';

  return [makeResult_({
    status: found ? 'OK' : (rule.severity || 'NG'),
    category: '内訳書',
    ruleId: rule.rule_id,
    sheetName: sheet.getName(),
    itemName: rule.header_name,
    targetCell: a1,
    jumpUrl: a1 ? buildRangeUrl_(ctx.ss, sheet.getName(), a1) : '',
    decisionValue: '',
    compareValue: found ? rule.header_name : '',
    diff: '',
    condition: 'HEADER_CHECK',
    message: found ? 'OK' : rule.message,
    detail: '',
  })];
}

function checkTitle_(rule, sheet, displayValues, ctx) {
  const found = findCellByText_(displayValues, rule.header_name);
  const a1 = found ? toA1_(found.row + 1, found.col + 1) : '';

  return [makeResult_({
    status: found ? 'OK' : (rule.severity || 'NG'),
    category: '内訳書',
    ruleId: rule.rule_id,
    sheetName: sheet.getName(),
    itemName: rule.header_name,
    targetCell: a1,
    jumpUrl: a1 ? buildRangeUrl_(ctx.ss, sheet.getName(), a1) : '',
    decisionValue: '',
    compareValue: found ? rule.header_name : '',
    diff: '',
    condition: 'TITLE_CHECK',
    message: found ? 'OK' : rule.message,
    detail: '',
  })];
}

function checkNormalizeSum_(rule, sheet, rows, ctx) {
  const accountCol = colToIndex_(rule.account_col);
  let amountCol = colToIndex_(rule.amount_col);
  let headerRow = findHeaderRowByText_(rows, rule.header_name, amountCol);

  if (headerRow < 0) {
    const found = findCellByText_(sheet.getDataRange().getDisplayValues(), rule.header_name);
    if (found) {
      headerRow = found.row;
      amountCol = found.col;
    }
  }

  if (headerRow < 0) return [ngHeaderNotFound_(rule, sheet)];

  const sums = {};
  const rowMap = {};

  for (let r = headerRow + 1; r < rows.length; r++) {
    const rawLabel = normalizeText_(rows[r][accountCol]);
    if (!rawLabel) continue;

    const normalizedAccount = normalizeAccountByRule_(rawLabel, ctx.normalizeRules) || rawLabel;
    const amount = toNumber_(rows[r][amountCol]);
    if (amount == null) continue;

    sums[normalizedAccount] = (sums[normalizedAccount] || 0) + amount;
    if (!rowMap[normalizedAccount]) rowMap[normalizedAccount] = r;
  }

  const results = [];
  Object.keys(sums).sort().forEach(account => {
    markDecisionAccountUsed_(ctx, account);

    const decisionValue = ctx.decisionMap[account];
    const a1 = toA1_(rowMap[account] + 1, amountCol + 1);
    const jumpUrl = buildRangeUrl_(ctx.ss, sheet.getName(), a1);

    if (decisionValue == null) {
      results.push(makeResult_({
        status: 'SKIP',
        category: '内訳書',
        ruleId: rule.rule_id,
        sheetName: sheet.getName(),
        itemName: account,
        targetCell: a1,
        jumpUrl: jumpUrl,
        decisionValue: '',
        compareValue: sums[account],
        diff: '',
        condition: '決算書科目なし',
        message: '決算書に対象科目が無いためスキップ',
        detail: '',
      }));
      return;
    }

    results.push(compareNumbersResult_(rule, sheet.getName(), account, decisionValue, sums[account], '内訳書', a1, jumpUrl));
  });

  if (results.length === 0) {
    results.push(makeResult_({
      status: '要確認',
      category: '内訳書',
      ruleId: rule.rule_id,
      sheetName: sheet.getName(),
      itemName: rule.document_type,
      targetCell: '',
      jumpUrl: '',
      decisionValue: '',
      compareValue: '',
      diff: '',
      condition: '比較対象なし',
      message: '正規化後に比較対象科目が見つかりませんでした',
      detail: '',
    }));
  }

  return results;
}

/* =========================
 * rules_kokyo
 * ========================= */

function runKokyoRules_(ss, ctx) {
  const results = [];
  const frontSheet = ss.getSheetByName(CONFIG.SHEET_KOKYO_FRONT);
  const backSheet = ss.getSheetByName(CONFIG.SHEET_KOKYO_BACK);

  const frontMap = frontSheet ? parseKeyValueSheet_(frontSheet) : {};
  const backValues = backSheet ? getSheetValues_(backSheet).values : [];

  for (const rule of ctx.rulesKokyo) {
    if (!toBoolean_(rule.enabled)) continue;
    const checkType = normalizeCheckType_(rule.check_type || rule.checktype || rule['check type']);

    try {
      switch (checkType) {
        case 'NOT_BLANK': {
          const value = frontMap[rule.item_name];
          results.push(makeResult_({
            status: isBlank_(value) ? (rule.severity || '要確認') : 'OK',
            category: '概況書',
            ruleId: rule.rule_id,
            sheetName: CONFIG.SHEET_KOKYO_FRONT,
            itemName: rule.item_name,
            targetCell: '',
            jumpUrl: '',
            decisionValue: '',
            compareValue: value || '',
            diff: '',
            condition: 'NOT_BLANK',
            message: isBlank_(value) ? rule.message : 'OK',
            detail: '',
          }));
          break;
        }

        case 'CONDITIONAL_NOT_BLANK': {
          const condOk = evaluateConditionAgainstDecision_(rule.condition, ctx);
          const value = frontMap[rule.item_name];
          const bad = condOk && isBlank_(value);

          results.push(makeResult_({
            status: bad ? (rule.severity || '要確認') : 'OK',
            category: '概況書',
            ruleId: rule.rule_id,
            sheetName: CONFIG.SHEET_KOKYO_FRONT,
            itemName: rule.item_name,
            targetCell: '',
            jumpUrl: '',
            decisionValue: '',
            compareValue: value || '',
            diff: '',
            condition: rule.condition || '',
            message: bad ? rule.message : 'OK',
            detail: '',
          }));
          break;
        }

        case 'MATCH_DECISION': {
          const frontValue = toNumber_(frontMap[rule.item_name]);
          let decisionValue = resolveDecisionExprWithPickRule_(
            rule.source_detail,
            ctx.decisionValues,
            rule.value_pick_rule || 'C>B',
            ctx,
            rule.account_match_mode || 'exact'
          );

          if (decisionValue != null) {
            decisionValue = toKiloYenFloor_(decisionValue);
          }

          results.push(compareNumbersResult_(rule, CONFIG.SHEET_KOKYO_FRONT, rule.item_name, decisionValue, frontValue, '概況書', '', ''));
          break;
        }

        case 'MATCH_DECISION_EXPR': {
          const frontValue = toNumber_(frontMap[rule.item_name]);
          let decisionValue = resolveDecisionExprWithPickRule_(
            rule.source_detail,
            ctx.decisionValues,
            rule.value_pick_rule || 'C>B',
            ctx,
            rule.account_match_mode || 'exact'
          );

          if (decisionValue != null) {
            decisionValue = toKiloYenFloor_(decisionValue);
          }

          results.push(compareNumbersResult_(rule, CONFIG.SHEET_KOKYO_FRONT, rule.item_name, decisionValue, frontValue, '概況書', '', ''));
          break;
        }

        case 'MATCH_BREAKDOWN': {
          results.push(makeResult_({
            status: rule.severity || '要確認',
            category: '概況書',
            ruleId: rule.rule_id,
            sheetName: CONFIG.SHEET_KOKYO_FRONT,
            itemName: rule.item_name,
            targetCell: '',
            jumpUrl: '',
            decisionValue: '',
            compareValue: '',
            diff: '',
            condition: 'MATCH_BREAKDOWN',
            message: rule.message,
            detail: rule.source_detail || '',
          }));
          break;
        }

        case 'MATCH_BREAKDOWN_LOOKUP': {
          const frontValue = toNumber_(frontMap[rule.item_name]);
          const lookupValue = resolveBreakdownLookupValue_(ss, rule);
          results.push(compareNumbersResult_(rule, CONFIG.SHEET_KOKYO_FRONT, rule.item_name, lookupValue, frontValue, '概況書', '', ''));
          break;
        }

        case 'CALC_MATCH': {
          if (!backSheet) {
            results.push(makeResult_({
              status: '要確認',
              category: '概況書',
              ruleId: rule.rule_id,
              sheetName: '',
              itemName: rule.item_name,
              targetCell: '',
              jumpUrl: '',
              decisionValue: '',
              compareValue: '',
              diff: '',
              condition: '概況書裏面未検出',
              message: '概況書裏面シートが見つかりません',
              detail: '',
            }));
            break;
          }

          const calcValue = evaluateCellExpression_(rule.source_detail, backValues);
          let targetValue = null;

          if (rule.rule_id === 'K100') {
            targetValue = toNumber_(frontMap['売上_収入_高_千円']);
          } else if (rule.rule_id === 'K101') {
            targetValue = toNumber_(frontMap['売上_収入_原価_千円']);
          } else if (rule.rule_id === 'K102') {
            targetValue =
              (toNumber_(frontMap['販管費のうち_役員報酬_千円']) || 0) +
              (toNumber_(frontMap['販管費のうち_従業員給料_千円']) || 0);
          }

          results.push(compareNumbersResult_(rule, CONFIG.SHEET_KOKYO_BACK, rule.item_name, targetValue, calcValue, '概況書', '', ''));
          break;
        }

        default:
          const debugMsg = checkType
            ? `未対応のcheck_typeです: ${checkType}`
            : `check_type が空です: row=${rule.__row_index || '?'} header_row=${rule.__header_row || '?'} rule_id=${rule.rule_id || ''} item=${rule.item_name || ''} raw=${rule.__raw_first || ''}（rules_kokyoの列ずれ・1セルTSV貼付の可能性）`;
          results.push(makeResult_({
            status: '要確認',
            category: '概況書',
            ruleId: rule.rule_id,
            sheetName: '',
            itemName: rule.item_name,
            targetCell: '',
            jumpUrl: '',
            decisionValue: '',
            compareValue: '',
            diff: '',
            condition: checkType || '',
            message: debugMsg,
            detail: '',
          }));
      }
    } catch (err) {
      results.push(makeResult_({
        status: '要確認',
        category: '概況書',
        ruleId: rule.rule_id,
        sheetName: '',
        itemName: rule.item_name,
        targetCell: '',
        jumpUrl: '',
        decisionValue: '',
        compareValue: '',
        diff: '',
        condition: '実行エラー',
        message: `概況書ルール実行中にエラー: ${err.message}`,
        detail: err.stack || '',
      }));
    }
  }

  return results;
}

function resolveBreakdownLookupValue_(ss, rule) {
  const nameSourceSheet = findTargetSheetByPattern_(ss, rule.lookup_name_source);
  if (!nameSourceSheet) return null;

  const nameCell = String(rule.lookup_name_cell || '').trim();
  let representativeName = '';
  if (nameCell) {
    representativeName = normalizeText_(nameSourceSheet.getRange(nameCell).getDisplayValue());
  }

  const targetSheet = findTargetSheetByPattern_(ss, rule.lookup_sheet_pattern || rule.source_detail);
  if (!targetSheet) return null;

  const values = targetSheet.getDataRange().getDisplayValues();
  const matchCol = colToIndex_(rule.lookup_match_col);
  const matchMode = rule.account_match_mode || 'contains';

  if (matchCol < 0) return null;

  for (let r = 0; r < values.length; r++) {
    const cellText = values[r][matchCol];
    if (!cellText) continue;

    if (!representativeName || isMatchByMode_(cellText, representativeName, matchMode)) {
      return getMultiColValue_(values[r], rule.lookup_value_col);
    }
  }

  return null;
}

function isMatchByMode_(cellText, targetText, mode) {
  const cell = normalizeText_(cellText);
  const target = normalizeText_(targetText);
  const m = String(mode || 'exact').toLowerCase();

  if (!cell || !target) return false;

  if (m === 'exact') return cell === target;
  if (m === 'prefix') return cell.startsWith(target);
  if (m === 'suffix') return cell.endsWith(target);
  if (m === 'contains') return cell.includes(target);

  return cell === target;
}

function getMultiColValue_(row, colExpr) {
  if (!colExpr) return 0;

  const cols = String(colExpr).split('+');
  let total = 0;

  cols.forEach(c => {
    const idx = colToIndex_(c.trim());
    if (idx >= 0) {
      total += toNumber_(row[idx] || 0);
    }
  });

  return total;
}

function resolveDecisionExprWithPickRule_(expr, values, valuePickRule, ctx, matchMode) {
  const text = normalizeText_(expr);
  if (!text) return null;

  const tokens = text.match(/[+-]?[^+-]+/g);
  if (!tokens) return null;

  let total = 0;
  let foundAny = false;

  for (const token of tokens) {
    const sign = token.startsWith('-') ? -1 : 1;
    const raw = token.replace(/^[+-]/, '');
    const part = normalizeText_(raw);

    let value = null;

    if (ctx.groups[part]) {
      value = 0;
      let foundGroup = false;

      for (const acc of ctx.groups[part]) {
        const matchedRows = findDecisionRowsByAccount_(values, acc, 'exact');
        for (const row of matchedRows) {
          const v = pickValueByRule_(row, valuePickRule);
          if (v != null) {
            value += v;
            foundGroup = true;
            markDecisionAccountUsed_(ctx, acc);
          }
        }
      }

      if (!foundGroup) value = null;
    } else {
      const matchedRows = findDecisionRowsByAccount_(values, part, matchMode || 'exact');

      if (matchedRows.length > 0) {
        value = 0;
        let foundRowValue = false;

        for (const row of matchedRows) {
          const v = pickValueByRule_(row, valuePickRule);
          if (v != null) {
            value += v;
            foundRowValue = true;
          }
        }

        if (!foundRowValue) {
          value = null;
        } else {
          markDecisionAccountUsed_(ctx, part);
        }
      }
    }

    if (value != null) {
      total += sign * value;
      foundAny = true;
    }
  }

  return foundAny ? total : null;
}

function findDecisionRowsByAccount_(values, accountName, matchMode) {
  const target = normalizeText_(accountName);
  const rows = [];
  const mode = String(matchMode || 'exact').toLowerCase();

  for (let r = 0; r < values.length; r++) {
    const cell = normalizeText_(values[r][0]);
    if (!cell) continue;

    let hit = false;

    if (mode === 'exact') {
      hit = (cell === target);
    } else if (mode === 'prefix') {
      hit = cell.startsWith(target);
    } else if (mode === 'suffix') {
      hit = cell.endsWith(target);
    } else if (mode === 'contains') {
      hit = cell.includes(target);
    } else {
      hit = (cell === target);
    }

    if (hit) rows.push(values[r]);
  }

  return rows;
}

function pickValueByRule_(row, ruleText) {
  const rule = String(ruleText || 'C>B').trim().toUpperCase();

  const colMap = {
    A: 0, B: 1, C: 2, D: 3, E: 4, F: 5, G: 6, H: 7, I: 8, J: 9, K: 10, L: 11, M: 12, N: 13, O: 14
  };

  if (rule === 'B_ONLY') return toNumber_(row[colMap.B]);
  if (rule === 'C_ONLY') return toNumber_(row[colMap.C]);

  if (rule === 'RIGHTMOST') {
    for (let i = row.length - 1; i >= 0; i--) {
      const n = toNumber_(row[i]);
      if (n != null) return n;
    }
    return null;
  }

  const order = rule.split('>').map(s => s.trim()).filter(Boolean);
  for (const col of order) {
    if (colMap[col] == null) continue;
    const n = toNumber_(row[colMap[col]]);
    if (n != null) return n;
  }

  for (let i = row.length - 1; i >= 0; i--) {
    const n = toNumber_(row[i]);
    if (n != null) return n;
  }

  return null;
}

/* =========================
 * 入力値全般チェック + Gemini
 * ========================= */

function runGlobalInputAnomalyChecks_(ss) {
  const results = [];
  const ignoreSheets = new Set([
    CONFIG.SHEET_RULES_MAIN,
    CONFIG.SHEET_RULES_KOKYO,
    CONFIG.SHEET_GROUP_MASTER,
    CONFIG.SHEET_NORMALIZE_MASTER,
    CONFIG.SHEET_EXCLUDE_MASTER,
    CONFIG.SHEET_AI_TARGETS,
    CONFIG.SHEET_RESULT,
    CONFIG.SHEET_LOG,
    CONFIG.SHEET_REPORT
  ]);

  const aiCache = {};
  const aiTargets = loadAiCheckTargets_(ss);
  const hasTargetConfig = Object.keys(aiTargets).length > 0;

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
    if (hasTargetConfig && ranges.length === 0) return;
    const visited = new Set();

    for (const range of ranges) {
      for (let r = CONFIG.AI_CHECK_START_ROW - 1; r < values.length; r++) {
        for (let c = range.startCol - 1; c <= range.endCol - 1; c++) {
          const cellKey = `${r}:${c}`;
          if (visited.has(cellKey)) continue;
          visited.add(cellKey);

          const raw = values[r][c];
          const text = String(raw || '').trim();
          if (!text) continue;
          if (/^[0-9,\.\-△()\/年月日円千百]+$/.test(text)) continue;
          if (text.length <= 1) continue;

          let ai = aiCache[text];
          if (!ai) {
            ai = callGeminiBreakdownCheckSafe_(text, sheetName);
            aiCache[text] = ai;
            Utilities.sleep(120);
          }

          if (!ai || !ai.is_suspicious) continue;
          if (!shouldKeepAiFinding_(text, ai)) continue;

          const a1 = toA1_(r + 1, c + 1);
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

  return results;
}

function resolveAiTargetsForSheet_(aiTargets, sheetName) {
  const out = [];
  const seenKeys = new Set();
  const normSheet = normalizeText_(sheetName);

  if (aiTargets[sheetName]) {
    out.push(...aiTargets[sheetName]);
    seenKeys.add(sheetName);
  }

  Object.keys(aiTargets).forEach(k => {
    if (seenKeys.has(k)) return;
    const normKey = normalizeText_(k);
    if (!normKey) return;
    if (normSheet.includes(normKey)) {
      out.push(...aiTargets[k]);
    }
  });

  return out;
}

function loadAiCheckTargets_(ss) {
  const sheet = ss.getSheetByName(CONFIG.SHEET_AI_TARGETS);
  if (!sheet) return {};

  const rows = loadRowsAsObjects_(sheet);
  const map = {};

  rows.forEach(r => {
    const sheetName = String(r.sheet_name || r.sheetname || r['sheet name'] || '').trim();
    if (!sheetName) return;

    const enabled = String(r.enabled == null ? 'TRUE' : r.enabled).trim().toUpperCase();
    const startCol = parseColRefTo1Based_(r.start_col || r.startcol || r['start col']) || 2;
    const endCol = parseColRefTo1Based_(r.end_col || r.endcol || r['end col']) || 21;

    if (!map[sheetName]) map[sheetName] = [];
    map[sheetName].push({
      enabled: enabled !== 'FALSE' && enabled !== '0' && enabled !== 'NO',
      startCol: Math.max(1, Math.min(startCol, endCol)),
      endCol: Math.max(startCol, endCol)
    });
  });

  return map;
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

  // 今回は「明らかな誤入力」を中心に拾うため、業務上許容される語は除外
  if (/^(登録番号|売掛金|買掛金|仮払金|仮受金|本人|該当|該当なし)$/u.test(norm)) return false;
  if (/^T\d{13}$/i.test(norm)) return false; // インボイス登録番号
  if (/^(令和|平成|昭和)\d+年\d+月\d+日$/u.test(norm)) return false;
  if (/^\d+[-‐－~～]\d+月/.test(norm)) return false; // 例: 1-2月利用分

  const hasClearTypoSignal =
    /(誤字|脱字|文字化け|重複|欠落|誤入力|入力ミス|途切れ|存在しません)/u.test(reason) ||
    /(.)\1/u.test(norm); // 例: 北沢沢

  // 全角英字だけの一般語（例: Ｌａｂｏｒａｔｏｒｙ）も残す
  const fullWidthAlphaOnly = /^[Ａ-Ｚａ-ｚ]+$/u.test(norm);

  return hasClearTypoSignal || fullWidthAlphaOnly;
}

function callGemini_(text) {
  const apiKey = getGeminiApiKey_();
  const model = getGeminiModel_();

  if (!apiKey) {
    return "Geminiキー未設定";
  }

  const url =
    "https://generativelanguage.googleapis.com/v1beta/models/" +
    model +
    ":generateContent?key=" +
    apiKey;

  const payload = {
    contents: [
      {
        parts: [
          {
            text:
              "以下の日本語の誤字・不自然表現・文字化けをチェックし、修正案を簡潔に提示してください。\n\n" +
              text
          }
        ]
      }
    ],
    generationConfig: {
      temperature: 0.2
    }
  };

  const res = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const json = JSON.parse(res.getContentText());

  if (json.error) {
    return "Geminiエラー: " + json.error.message;
  }

  return json.candidates?.[0]?.content?.parts?.[0]?.text || "応答なし";
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
      is_suspicious: true,
      reason: 'GEMINI_API_KEY が未設定です',
      suggestions: []
    };
  }

  const model = getGeminiModel_();
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;

  const prompt = `
あなたは税務申告書の内訳書チェック支援AIです。
対象は「○○の内訳」シートの入力値です。
次の文字列が、税務書類の入力値として不自然かどうか判定してください。

特に次の観点を重視してください。
- 誤字脱字
- 文字化け
- 同じ語の重複（例: 所沢沢）
- 半角カナと全角カナの混在
- 全角英数字と半角英数字の不自然混在
- 日本語と英字の不自然な連結（例: F菱UFJ銀行）
- 銀行名・会社名・地名・氏名として不自然
- 明らかな入力ミスの可能性

注意:
- 珍しい固有名詞の可能性もあるため、断定しすぎず要確認ベースで判断してください。
- 明らかに自然で問題ない場合は is_suspicious を false にしてください。
- 修正候補は最大2件まで。
- JSONだけ返してください。

シート名: ${sheetName}
入力文字列: ${text}

返却形式:
{
  "is_suspicious": true,
  "reason": "不自然だと考える理由",
  "suggestions": ["候補1", "候補2"]
}
`.trim();

  const payload = {
    contents: [
      {
        parts: [{ text: prompt }]
      }
    ],
    generationConfig: {
      temperature: 0.1,
      responseMimeType: 'application/json'
    }
  };

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  const body = res.getContentText();

  if (code < 200 || code >= 300) {
    throw new Error(`Gemini API error ${code}: ${body}`);
  }

  const json = JSON.parse(body);
  const textOut = json.candidates?.[0]?.content?.parts?.[0]?.text;
  if (!textOut) throw new Error('Gemini応答が空です');

  const parsed = JSON.parse(textOut);

  return {
    is_suspicious: !!parsed.is_suspicious,
    reason: parsed.reason || '',
    suggestions: Array.isArray(parsed.suggestions) ? parsed.suggestions.slice(0, 2) : []
  };
}

function callGeminiBreakdownCheckSafe_(text, sheetName) {
  try {
    return callGeminiBreakdownCheck_(text, sheetName);
  } catch (e) {
    return {
      is_suspicious: true,
      reason: `Gemini判定エラー: ${e && e.message ? e.message : String(e)}`,
      suggestions: []
    };
  }
}

/* =========================
 * report_A4
 * ========================= */

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
    return status !== 'OK';
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
 * 決算書解析
 * ========================= */

function parseDecisionSheet_(sheet) {
  const { values } = getSheetValues_(sheet);
  const map = {};
  const bsAccounts = [];

  let section = '';
  let inBS = false;

  for (let r = 0; r < values.length; r++) {
    const colA = normalizeText_(values[r][0]);
    const colB = values[r][1];
    const colC = values[r][2];
    const colD = values[r][3];

    const key = colA;
    const amount = pickDecisionAmount_(colB, colC, colD);

    if (!key) continue;

    if (key.includes('資産の部')) {
      section = '資産の部';
      inBS = true;
    } else if (key.includes('負債の部')) {
      section = '負債の部';
      inBS = true;
    } else if (key.includes('純資産の部')) {
      section = '純資産の部';
      inBS = false;
    }

    if (amount != null) {
      map[key] = amount;
    }

    if (section && amount != null && key && !key.includes('の部')) {
      map[`${section}:${key}`] = amount;
    }

    if (inBS && amount != null && isRealBSAccount_(key)) {
      bsAccounts.push(key);
    }
  }

  return {
    map,
    bsAccounts: [...new Set(bsAccounts)],
  };
}

function pickDecisionAmount_(colB, colC, colD) {
  const c = toNumber_(colC);
  if (c != null) return c;
  const b = toNumber_(colB);
  if (b != null) return b;
  return toNumber_(colD);
}

function isRealBSAccount_(key) {
  const k = normalizeText_(key);
  if (!k) return false;

  const excludes = new Set([
    '資産の部合計',
    '負債の部合計',
    '純資産の部合計',
    '流動資産',
    '固定資産',
    '繰延資産',
    '流動負債',
    '固定負債',
    '株主資本',
    '資本金',
    '利益剰余金',
    'その他利益剰余金',
    '繰越利益剰余金',
    '評価・換算差額等',
    '新株予約権'
  ]);

  if (excludes.has(k)) return false;
  if (k.startsWith('【') && k.endsWith('】')) return false;
  if (k.includes('合計')) return false;
  if (k.includes('うち')) return false;

  return true;
}

function loadRulesMain_(ss) {
  return loadRowsAsObjects_(getRequiredSheet_(ss, CONFIG.SHEET_RULES_MAIN));
}

function loadRulesKokyo_(ss) {
  const sheet = getRequiredSheet_(ss, CONFIG.SHEET_RULES_KOKYO);
  const { displayValues } = getSheetValues_(sheet);
  if (displayValues.length < 2) return [];

  const headerRow = detectHeaderRow_(displayValues);
  const headerKeys = new Set(displayValues[headerRow].map(v => normalizeText_(v).toLowerCase()));
  const hasHeader =
    (headerKeys.has('rule_id') || headerKeys.has('ruleid')) &&
    (headerKeys.has('check_type') || headerKeys.has('checktype'));
  if (hasHeader) return loadRowsAsObjects_(sheet);

  return parseRulesKokyoByPosition_(displayValues);
}

function parseRulesKokyoByPosition_(displayValues) {
  const out = [];
  for (let r = 0; r < displayValues.length; r++) {
    let row = displayValues[r];
    if (!row || row.join('').trim() === '') continue;

    // 1セルにTSV形式で入っているケースを救済
    if (
      row.length > 0 &&
      String(row[0] || '').includes('\t') &&
      row.slice(1).every(v => String(v || '').trim() === '')
    ) {
      row = String(row[0]).split('\t');
    }

    const idMatch = String(row[0] || '').match(/K\d{3}/);
    if (!idMatch) continue;

    out.push({
      rule_id: idMatch[0],
      enabled: row[1],
      item_name: row[2],
      check_type: row[3],
      document_type: row[4],
      source_detail: row[5],
      value_pick_rule: row[6],
      account_match_mode: row[7],
      lookup_sheet_pattern: row[8],
      lookup_name_source: row[9],
      lookup_name_cell: row[10],
      lookup_match_col: row[11],
      lookup_value_col: row[12],
      condition: row[13],
      severity: row[14],
      message: row[15],
      __row_index: r + 1,
      __raw_first: String(row[0] || '')
    });
  }
  return out;
}

function loadGroupMaster_(ss) {
  const rows = loadRowsAsObjects_(getRequiredSheet_(ss, CONFIG.SHEET_GROUP_MASTER));
  const map = {};
  rows.forEach(r => {
    const g = normalizeText_(r.group_name);
    const a = normalizeText_(r.account_name);
    if (!g || !a) return;
    if (!map[g]) map[g] = [];
    map[g].push(a);
  });
  return map;
}

function loadNormalizeMaster_(ss) {
  return loadRowsAsObjects_(getRequiredSheet_(ss, CONFIG.SHEET_NORMALIZE_MASTER))
    .map(r => ({
      keyword: normalizeText_(r.keyword),
      normalized: normalizeText_(r.normalized),
    }))
    .filter(r => r.keyword && r.normalized);
}

function loadExcludeMaster_(ss) {
  const rows = loadRowsAsObjects_(getRequiredSheet_(ss, CONFIG.SHEET_EXCLUDE_MASTER));
  const set = new Set();
  rows.forEach(r => {
    const name = normalizeText_(r.account_name);
    if (name) set.add(name);
  });
  return set;
}

/* =========================
 * 決算書値解決
 * ========================= */

function resolveDecisionTargetValue_(rule, ctx) {
  if (normalizeText_(rule.target_account)) {
    markDecisionExpressionUsed_(ctx, rule.target_account);
    return resolveDecisionExpression_(rule.target_account, ctx);
  }
  if (normalizeText_(rule.target_account_group)) {
    const groupName = normalizeText_(rule.target_account_group);
    const accounts = ctx.groups[groupName] || [];
    const filtered = accounts.filter(a => ctx.decisionMap[a] != null);
    filtered.forEach(a => markDecisionAccountUsed_(ctx, a));
    if (filtered.length === 0) return null;
    return filtered.reduce((sum, a) => sum + (ctx.decisionMap[a] || 0), 0);
  }
  return null;
}

function resolveDecisionExpression_(expr, ctx) {
  const text = normalizeText_(expr);
  if (!text) return null;

  const parts = text.split('+').map(s => normalizeText_(s)).filter(Boolean);
  if (parts.length === 0) return null;

  let foundAny = false;
  const total = parts.reduce((sum, part) => {
    if (ctx.groups[part]) {
      const groupAccounts = ctx.groups[part].filter(acc => ctx.decisionMap[acc] != null);
      if (groupAccounts.length === 0) return sum;
      foundAny = true;
      groupAccounts.forEach(acc => markDecisionAccountUsed_(ctx, acc));
      return sum + groupAccounts.reduce((s, acc) => s + (ctx.decisionMap[acc] || 0), 0);
    }

    if (ctx.decisionMap[part] != null) {
      foundAny = true;
      markDecisionAccountUsed_(ctx, part);
      return sum + (ctx.decisionMap[part] || 0);
    }

    return sum;
  }, 0);

  return foundAny ? total : null;
}

function markDecisionAccountUsed_(ctx, accountName) {
  const a = normalizeText_(accountName);
  if (!a) return;
  ctx.usedDecisionAccounts.add(a);
}

function markDecisionExpressionUsed_(ctx, expr) {
  const text = normalizeText_(expr);
  if (!text) return;

  const parts = text.split('+').map(s => normalizeText_(s)).filter(Boolean);
  parts.forEach(part => {
    if (ctx.groups[part]) {
      ctx.groups[part].forEach(acc => markDecisionAccountUsed_(ctx, acc));
    } else {
      markDecisionAccountUsed_(ctx, part);
    }
  });
}

function toKiloYenFloor_(value) {
  if (value === null || value === '' || isNaN(value)) return null;
  return Math.floor(Number(value) / 1000);
}

/* =========================
 * 汎用
 * ========================= */

function loadRowsAsObjects_(sheet) {
  const { displayValues } = getSheetValues_(sheet);
  if (displayValues.length < 2) return [];

  const headerRow = detectHeaderRow_(displayValues);
  const headers = displayValues[headerRow].map(h => normalizeText_(h));
  const rows = [];

  for (let r = headerRow + 1; r < displayValues.length; r++) {
    const row = displayValues[r];
    if (row.join('').trim() === '') continue;

    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    obj.__row_index = r + 1;
    obj.__header_row = headerRow + 1;
    obj.__raw_first = String(row[0] || '');
    rows.push(obj);
  }
  return rows;
}

function detectHeaderRow_(displayValues) {
  const limit = Math.min(displayValues.length, 20);
  for (let r = 0; r < limit; r++) {
    const keys = new Set(displayValues[r].map(v => normalizeText_(v).toLowerCase()));
    const hasRuleId = keys.has('rule_id') || keys.has('ruleid');
    const hasCheckType =
      keys.has('check_type') ||
      keys.has('checktype');
    const hasItemName = keys.has('item_name') || keys.has('itemname');

    if (hasRuleId && (hasCheckType || hasItemName)) {
      return r;
    }
  }
  return 0;
}

function getSheetValues_(sheet) {
  const range = sheet.getDataRange();
  return {
    values: range.getValues(),
    displayValues: range.getDisplayValues(),
  };
}

function getRequiredSheet_(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error(`シートがありません: ${name}`);
  return sh;
}

function ensureSheetInSpreadsheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function findTargetSheetByPattern_(ss, pattern) {
  const p = normalizeText_(pattern);
  if (!p) return null;

  const sheets = ss.getSheets();
  for (const sh of sheets) {
    if (normalizeText_(sh.getName()).includes(p)) return sh;
  }
  return null;
}

function parseKeyValueSheet_(sheet) {
  const { displayValues } = getSheetValues_(sheet);
  const map = {};

  for (let r = 0; r < displayValues.length; r++) {
    const key = normalizeText_(displayValues[r][0]);
    const value = displayValues[r][1];
    if (!key) continue;
    map[key] = value;
  }
  return map;
}

function evaluateConditionAgainstDecision_(condition, ctx) {
  const text = normalizeText_(condition);
  if (!text) return false;

  const m = text.match(/^(.+?)(>=|<=|>|<|=)(.+)$/);
  if (!m) return false;

  const expr = normalizeText_(m[1]);
  const op = m[2];
  const rhs = Number(m[3]);
  const value = resolveDecisionExpression_(expr, ctx) || 0;

  switch (op) {
    case '>': return value > rhs;
    case '<': return value < rhs;
    case '>=': return value >= rhs;
    case '<=': return value <= rhs;
    case '=': return value === rhs;
    default: return false;
  }
}

function evaluateCellExpression_(expr, values) {
  const text = normalizeText_(expr);
  if (!text) return null;

  const tokens = text.match(/[+-]?[^+-]+/g);
  if (!tokens) return null;

  let total = 0;
  let foundAny = false;

  for (const token of tokens) {
    const sign = token.startsWith('-') ? -1 : 1;
    const ref = token.replace(/^[+-]/, '');
    const v = getCellByA1_(values, ref);
    if (v != null) {
      total += sign * v;
      foundAny = true;
    }
  }

  return foundAny ? total : null;
}

function getCellByA1_(values, a1) {
  const m = String(a1).match(/^([A-Z]+)(\d+)$/i);
  if (!m) return 0;
  const col = colToIndex_(m[1]);
  const row = Number(m[2]) - 1;
  if (row < 0 || row >= values.length) return 0;
  if (col < 0 || col >= values[row].length) return 0;
  return toNumber_(values[row][col]) || 0;
}

function colToIndex_(col) {
  const s = String(col || '').trim().toUpperCase();
  if (!s) return -1;
  let n = 0;
  for (let i = 0; i < s.length; i++) {
    n = n * 26 + (s.charCodeAt(i) - 64);
  }
  return n - 1;
}

function findHeaderRowByText_(rows, text, colIndex) {
  const target = normalizeText_(text);
  if (!target || colIndex < 0) return -1;

  for (let r = 0; r < rows.length; r++) {
    const cell = rows[r][colIndex];
    if (normalizeText_(cell) === target) return r;
  }
  return -1;
}

function sumColumnBelowHeader_(rows, headerRow, colIndex) {
  let sum = 0;
  for (let r = headerRow + 1; r < rows.length; r++) {
    const n = toNumber_(rows[r][colIndex]);
    if (n != null) sum += n;
  }
  return sum;
}

function findCellByText_(displayValues, text) {
  const target = normalizeText_(text);
  for (let r = 0; r < displayValues.length; r++) {
    for (let c = 0; c < displayValues[r].length; c++) {
      if (normalizeText_(displayValues[r][c]) === target) {
        return { row: r, col: c };
      }
    }
  }
  return null;
}

function findRightNumericValue_(displayValues, row, col) {
  for (let c = col + 1; c < Math.min(displayValues[row].length, col + 8); c++) {
    const n = toNumber_(displayValues[row][c]);
    if (n != null) return n;
  }
  return null;
}

function findRightTextValue_(displayValues, row, col) {
  for (let c = col + 1; c < Math.min(displayValues[row].length, col + 8); c++) {
    const v = normalizeText_(displayValues[row][c]);
    if (v) return v;
  }
  return '';
}

function normalizeAccountByRule_(text, normalizeRules) {
  const t = normalizeText_(text);
  for (const rule of normalizeRules) {
    if (t.includes(rule.keyword)) return rule.normalized;
  }
  return '';
}

function mergeExcludeAccounts_(ruleExclude, masterSet) {
  const set = new Set(masterSet ? [...masterSet] : []);
  normalizeText_(ruleExclude).split('|').map(s => s.trim()).filter(Boolean).forEach(v => set.add(v));
  return set;
}

function buildRangeUrl_(ss, sheetName, a1) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh || !a1) return '';
  return `${ss.getUrl()}#gid=${sh.getSheetId()}&range=${encodeURIComponent(a1)}`;
}

function compareNumbersResult_(rule, sheetName, itemName, decisionValue, compareValue, category, targetCell, jumpUrl) {
  const d = toNumber_(decisionValue);
  const c = toNumber_(compareValue);

  if (d == null) {
    return makeResult_({
      status: 'SKIP',
      category: category || '内訳書',
      ruleId: rule.rule_id,
      sheetName,
      itemName,
      targetCell: targetCell || '',
      jumpUrl: jumpUrl || '',
      decisionValue: '',
      compareValue: c ?? '',
      diff: '',
      condition: '決算書科目なし',
      message: '決算書に対象科目が無いためスキップ',
      detail: '',
    });
  }

  const diff = (c == null) ? '' : c - d;
  const ok = (c != null && Math.round(d) === Math.round(c));

  return makeResult_({
    status: ok ? 'OK' : (rule.severity || 'NG'),
    category: category || '内訳書',
    ruleId: rule.rule_id,
    sheetName,
    itemName,
    targetCell: targetCell || '',
    jumpUrl: jumpUrl || '',
    decisionValue: d,
    compareValue: c,
    diff,
    condition: rule.check_type || '',
    message: ok ? 'OK' : rule.message,
    detail: '',
  });
}

function ngHeaderNotFound_(rule, sheet) {
  return makeResult_({
    status: rule.severity || 'NG',
    category: '内訳書',
    ruleId: rule.rule_id,
    sheetName: sheet.getName(),
    itemName: rule.header_name,
    targetCell: '',
    jumpUrl: '',
    decisionValue: '',
    compareValue: '',
    diff: '',
    condition: 'ヘッダー未検出',
    message: `${rule.header_name} が見つかりません`,
    detail: '',
  });
}

function toA1_(row, col) {
  let s = '';
  while (col > 0) {
    const m = (col - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    col = Math.floor((col - 1) / 26);
  }
  return `${s}${row}`;
}

function makeResult_(obj) {
  return [
    obj.status || '',
    obj.category || '',
    obj.ruleId || '',
    obj.sheetName || '',
    obj.itemName || '',
    obj.targetCell || '',
    obj.jumpUrl || '',
    obj.decisionValue ?? '',
    obj.compareValue ?? '',
    obj.diff ?? '',
    obj.condition || '',
    obj.message || '',
    obj.detail || '',
    obj.aiJudge || '',
    obj.aiReason || '',
    obj.aiSuggestion1 || '',
    obj.aiSuggestion2 || '',
    new Date(),
  ];
}

function writeResults_(ss, rows) {
  const sh = getRequiredSheet_(ss, CONFIG.SHEET_RESULT);
  if (!rows.length) return;
  sh.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

function appendLogRow_(ss, message) {
  const sh = ensureSheetInSpreadsheet_(ss, CONFIG.SHEET_LOG);
  sh.appendRow([new Date(), message]);
}

function getNowStr_() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
}

function normalizeText_(v) {
  return String(v || '')
    .replace(/\s/g, '')
    .replace(/[　]/g, '')
    .trim();
}

function normalizeCheckType_(v) {
  const t = String(v || '').trim().toUpperCase();
  if (!t) return '';

  const aliasMap = {
    'NOTBLANK': 'NOT_BLANK',
    'CONDITIONALNOTBLANK': 'CONDITIONAL_NOT_BLANK',
    'MATCHDECISION': 'MATCH_DECISION',
    'MATCHDECISIONEXPR': 'MATCH_DECISION_EXPR',
    'MATCHBREAKDOWN': 'MATCH_BREAKDOWN',
    'MATCHBREAKDOWNLOOKUP': 'MATCH_BREAKDOWN_LOOKUP',
    'CALCMATCH': 'CALC_MATCH',
    'NOTBLANKWHENROWEXISTS': 'NOT_BLANK_WHEN_ROW_EXISTS',
  };
  return aliasMap[t.replace(/[_\-\s]/g, '')] || t;
}

function isBlank_(v) {
  return normalizeText_(v) === '';
}

function toBoolean_(v) {
  const s = String(v == null ? '' : v).toUpperCase();
  return s === 'TRUE' || s === '1' || s === 'YES';
}

function toNumber_(v) {
  if (v == null || v === '') return null;
  if (typeof v === 'number') return v;

  let s = String(v).trim();
  if (!s) return null;

  s = s
    .replace(/,/g, '')
    .replace(/△/g, '-')
    .replace(/[()]/g, '')
    .replace(/円/g, '');

  if (!s || s === '-') return null;

  const n = Number(s);
  return isNaN(n) ? null : n;
}

function isSectionLike_(text) {
  const t = normalizeText_(text);
  return (
    t.startsWith('【') ||
    t.includes('合計') ||
    t.includes('計') ||
    t.includes('の部')
  );
}
