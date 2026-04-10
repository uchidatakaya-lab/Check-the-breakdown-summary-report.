/*********************************
 * 決算書PDF × Excel 突合ツール（GAS完結版）
 * - ブックマークレットで folderId を受け取る
 * - フォルダ直下の決算書PDFとExcelを取得
 * - Excelは一時変換して5行目以降が空欄なら old フォルダへ移動
 * - 採用Excelは A列=勘定科目 / F列=金額 で集計
 * - 決算書PDFは一時Googleドキュメント化して貸借対照表から科目残高抽出
 * - 決算書BS残高とExcel集計残高を突合
 *********************************/

const APP_VERSION = '2026-04-10.02';

const CONFIG = {
  RESULT_FILE_PREFIX: '決算書Excelチェック結果',
  TZ: Session.getScriptTimeZone() || 'Asia/Tokyo',

  DELETE_TEMP_CONVERTED_SHEETS: true,
  DELETE_TEMP_CONVERTED_DOCS: true,

  SHEET_NAMES: {
    summary: '実行サマリ',
    files: 'ファイル一覧',
    excelCheck: 'Excel判定結果',
    adopted: '採用対象',
    excelBalance: 'Excel残高集計',
    bs: '決算書BS残高',
    compare: '決算書突合',
    log: 'ログ',
  },

  PDF_NAME_HINTS: ['決算', '決算書', '決算報告書', '貸借対照表', '損益計算書'],
};

const CHECK_CONFIG = {
  EXCEL_ACCOUNT_COL: 1, // A列
  EXCEL_AMOUNT_COL: 6,  // F列
  DATA_START_ROW: 5,
};

const EXCEL_MIME_TYPES = [
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  'application/vnd.ms-excel',
  'application/vnd.ms-excel.sheet.macroEnabled.12',
];

/** =========================
 * 公開入口
 * ========================= */

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';
  const folderId = (e && e.parameter && e.parameter.folderId) || '';

  try {
    if (action === 'run') {
      if (!folderId) throw new Error('folderId が指定されていません。');

      const result = runCheckSetupByFolderId_(folderId);
      return HtmlService.createHtmlOutput(renderHtmlResult_(result))
        .setTitle('決算書Excelチェック起動結果');
    }

    return HtmlService.createHtmlOutput([
      '<h3>決算書Excelチェック起動GAS</h3>',
      '<p>ブックマークレットから起動してください。</p>',
      '<p>パラメータ: ?action=run&folderId=...</p>'
    ].join(''));
  } catch (err) {
    return HtmlService.createHtmlOutput(
      '<h3>エラー</h3><pre>' + escapeHtml_(String(err && err.stack || err)) + '</pre>'
    );
  }
}

/** =========================
 * メイン処理
 * ========================= */

function runCheckSetupByFolderId_(folderId) {
  const startedAt = new Date();
  const logs = [];
  log_(logs, 'START', '処理開始 folderId=' + folderId);

  const folder = DriveApp.getFolderById(folderId);
  log_(logs, 'INFO', 'フォルダ名: ' + folder.getName());

  const scanned = scanFolderFiles_(folder, logs);
  const selectedPdf = selectMainFinancialPdf_(scanned.pdfFiles, logs);
  const excelAnalysis = analyzeExcelFiles_(folder, scanned.excelFiles, logs);

  const adoptedExcelFiles = excelAnalysis.filter(x => x.isTarget);
  const skippedExcelFiles = excelAnalysis.filter(x => !x.isTarget);

  log_(logs, 'INFO', '決算書PDF採用: ' + (selectedPdf ? selectedPdf.fileName : 'なし'));
  log_(logs, 'INFO', 'Excel採用件数: ' + adoptedExcelFiles.length);
  log_(logs, 'INFO', 'Excel除外件数: ' + skippedExcelFiles.length);

  let excelBalance = { accountMap: {}, detailRows: [] };
  let bsMap = {};
  let comparisonRows = [];

  if (adoptedExcelFiles.length) {
    excelBalance = buildExcelAccountBalanceMap_(folder, adoptedExcelFiles, logs);
    log_(logs, 'INFO', 'Excel残高集計科目数: ' + Object.keys(excelBalance.accountMap).length);
  }

  if (selectedPdf) {
    bsMap = extractBsMapFromPdf_(folder, selectedPdf, logs);
    log_(logs, 'INFO', '決算書BS科目数: ' + Object.keys(bsMap).length);
  }

  comparisonRows = compareBsVsExcel_(bsMap, excelBalance.accountMap);
  log_(logs, 'INFO', '突合件数: ' + comparisonRows.length);

  const resultSs = createResultSpreadsheet_({
    folder,
    folderId,
    scanned,
    selectedPdf,
    excelAnalysis,
    excelBalance,
    bsMap,
    comparisonRows,
    startedAt,
    logs,
  });

  log_(logs, 'DONE', '結果スプレッドシート作成: ' + resultSs.getUrl());

  return {
    ok: true,
    folderId,
    folderName: folder.getName(),
    resultSpreadsheetId: resultSs.getId(),
    resultSpreadsheetUrl: resultSs.getUrl(),
    pdfCount: scanned.pdfFiles.length,
    excelCount: scanned.excelFiles.length,
    adoptedExcelCount: adoptedExcelFiles.length,
    skippedExcelCount: skippedExcelFiles.length,
    selectedPdfName: selectedPdf ? selectedPdf.fileName : '',
  };
}

/** =========================
 * フォルダ走査
 * ========================= */

function scanFolderFiles_(folder, logs) {
  const allFiles = [];
  const pdfFiles = [];
  const excelFiles = [];
  const otherFiles = [];

  const it = folder.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    const info = {
      fileId: f.getId(),
      fileName: f.getName(),
      mimeType: f.getMimeType(),
      url: 'https://drive.google.com/file/d/' + f.getId() + '/view',
      size: safeGetSize_(f),
      updatedAt: safeGetUpdatedAt_(f),
    };

    allFiles.push(info);

    if (isPdfFile_(info)) {
      pdfFiles.push(info);
    } else if (isExcelFile_(info)) {
      excelFiles.push(info);
    } else {
      otherFiles.push(info);
    }
  }

  log_(logs, 'INFO', '全ファイル数: ' + allFiles.length);
  log_(logs, 'INFO', 'PDF件数: ' + pdfFiles.length);
  log_(logs, 'INFO', 'Excel件数: ' + excelFiles.length);
  log_(logs, 'INFO', 'その他件数: ' + otherFiles.length);

  return {
    allFiles,
    pdfFiles,
    excelFiles,
    otherFiles,
  };
}

function isPdfFile_(fileInfo) {
  return fileInfo.mimeType === MimeType.PDF || /\.pdf$/i.test(fileInfo.fileName);
}

function isExcelFile_(fileInfo) {
  if (EXCEL_MIME_TYPES.includes(fileInfo.mimeType)) return true;
  return /\.(xlsx|xlsm|xls)$/i.test(fileInfo.fileName);
}

function safeGetSize_(file) {
  try {
    return file.getSize();
  } catch (e) {
    return '';
  }
}

function safeGetUpdatedAt_(file) {
  try {
    return file.getLastUpdated();
  } catch (e) {
    return '';
  }
}

/** =========================
 * old フォルダ
 * ========================= */

function getOrCreateOldFolder_(parentFolder) {
  const name = 'old';
  const it = parentFolder.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return parentFolder.createFolder(name);
}

/** =========================
 * 決算書PDF選定
 * ========================= */

function selectMainFinancialPdf_(pdfFiles, logs) {
  if (!pdfFiles || !pdfFiles.length) return null;

  const scored = pdfFiles.map(f => {
    const name = String(f.fileName || '');
    let score = 0;

    CONFIG.PDF_NAME_HINTS.forEach(hint => {
      if (name.indexOf(hint) >= 0) score += 10;
    });

    if (/決算書/i.test(name)) score += 20;
    if (/決算報告書/i.test(name)) score += 20;

    return Object.assign({}, f, { score });
  });

  scored.sort((a, b) => {
    if (b.score !== a.score) return b.score - a.score;
    return String(a.fileName).localeCompare(String(b.fileName), 'ja');
  });

  const selected = scored[0] || null;
  if (selected) {
    log_(logs, 'INFO', '採用PDF: ' + selected.fileName + ' / score=' + selected.score);
  }

  return selected;
}

/** =========================
 * Excel変換・判定
 * ========================= */

function convertExcelToGoogleSheet_(folder, fileInfo) {
  const originalFile = DriveApp.getFileById(fileInfo.fileId);
  const blob = originalFile.getBlob();

  const resource = {
    name: '[TEMP]_' + fileInfo.fileName + '_' + Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyyMMdd_HHmmss'),
    mimeType: MimeType.GOOGLE_SHEETS,
    parents: [folder.getId()],
  };

  const converted = Drive.Files.create(resource, blob, {
    supportsAllDrives: true,
  });

  return {
    spreadsheetId: converted.id,
    spreadsheetUrl: 'https://docs.google.com/spreadsheets/d/' + converted.id + '/edit',
  };
}

function inspectConvertedSpreadsheet_(spreadsheetId) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheets = ss.getSheets();

  let hasDataAfterRow5 = false;
  const nonEmptySheetNames = [];
  let checkedSheetCount = 0;

  sheets.forEach(sh => {
    checkedSheetCount++;

    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();

    if (lastRow < CHECK_CONFIG.DATA_START_ROW || lastCol < 1) return;

    const startRow = CHECK_CONFIG.DATA_START_ROW;
    const numRows = lastRow - startRow + 1;
    if (numRows <= 0) return;

    const values = sh.getRange(startRow, 1, numRows, lastCol).getDisplayValues();
    const hasAny = values.some(row => row.some(v => String(v).trim() !== ''));

    if (hasAny) {
      hasDataAfterRow5 = true;
      nonEmptySheetNames.push(sh.getName());
    }
  });

  return {
    hasDataAfterRow5,
    nonEmptySheetNames,
    checkedSheetCount,
  };
}

function analyzeExcelFiles_(folder, excelFiles, logs) {
  const results = [];
  const oldFolder = getOrCreateOldFolder_(folder);

  for (let i = 0; i < excelFiles.length; i++) {
    const fileInfo = excelFiles[i];
    log_(logs, 'INFO', 'Excel判定開始: ' + fileInfo.fileName);

    let tempSpreadsheetId = '';
    let tempSpreadsheetUrl = '';
    let isTarget = false;
    let hasDataAfterRow5 = false;
    let nonEmptySheetNames = [];
    let checkedSheetCount = 0;
    let error = '';
    let statusJa = '';

    try {
      const converted = convertExcelToGoogleSheet_(folder, fileInfo);
      tempSpreadsheetId = converted.spreadsheetId;
      tempSpreadsheetUrl = converted.spreadsheetUrl;

      const check = inspectConvertedSpreadsheet_(tempSpreadsheetId);
      hasDataAfterRow5 = check.hasDataAfterRow5;
      nonEmptySheetNames = check.nonEmptySheetNames;
      checkedSheetCount = check.checkedSheetCount;

      isTarget = hasDataAfterRow5;

      if (isTarget) {
        statusJa = 'OK（使用）';
      } else {
        statusJa = '空欄（oldへ移動）';
        try {
          const file = DriveApp.getFileById(fileInfo.fileId);
          oldFolder.addFile(file);
          folder.removeFile(file);
          log_(logs, 'INFO', 'oldフォルダへ移動: ' + fileInfo.fileName);
        } catch (moveErr) {
          log_(logs, 'WARN', 'old移動失敗: ' + fileInfo.fileName + ' / ' + moveErr);
        }
      }

      log_(logs, 'INFO',
        'Excel判定完了: ' + fileInfo.fileName +
        ' / status=' + statusJa +
        ' / nonEmptySheets=' + nonEmptySheetNames.join(', ')
      );

    } catch (e) {
      error = String(e && e.message || e);
      statusJa = '判定失敗';
      log_(logs, 'ERROR', 'Excel判定失敗: ' + fileInfo.fileName + ' / ' + error);
    } finally {
      if (tempSpreadsheetId && CONFIG.DELETE_TEMP_CONVERTED_SHEETS) {
        try {
          DriveApp.getFileById(tempSpreadsheetId).setTrashed(true);
        } catch (e) {
          log_(logs, 'WARN', '一時変換Sheet削除失敗: ' + tempSpreadsheetId);
        }
      }
    }

    results.push({
      fileId: fileInfo.fileId,
      fileName: fileInfo.fileName,
      mimeType: fileInfo.mimeType,
      url: fileInfo.url,
      tempSpreadsheetId,
      tempSpreadsheetUrl,
      checkedSheetCount,
      hasDataAfterRow5,
      nonEmptySheetNames,
      isTarget,
      statusJa,
      error,
    });
  }

  return results;
}

/** =========================
 * Excel集計
 * ========================= */

function buildExcelAccountBalanceMap_(folder, adoptedExcelFiles, logs) {
  const accountMap = {};
  const detailRows = [];

  for (const fileInfo of adoptedExcelFiles) {
    let tempSpreadsheetId = '';
    try {
      const converted = convertExcelToGoogleSheet_(folder, fileInfo);
      tempSpreadsheetId = converted.spreadsheetId;

      const ss = SpreadsheetApp.openById(tempSpreadsheetId);
      const sheets = ss.getSheets();

      sheets.forEach(sh => {
        const lastRow = sh.getLastRow();
        const lastCol = sh.getLastColumn();
        if (lastRow < CHECK_CONFIG.DATA_START_ROW || lastCol < CHECK_CONFIG.EXCEL_AMOUNT_COL) return;

        const numRows = lastRow - CHECK_CONFIG.DATA_START_ROW + 1;
        const values = sh.getRange(
          CHECK_CONFIG.DATA_START_ROW,
          1,
          numRows,
          Math.max(lastCol, CHECK_CONFIG.EXCEL_AMOUNT_COL)
        ).getDisplayValues();

        values.forEach((row, idx) => {
          const rawAccount = String(row[CHECK_CONFIG.EXCEL_ACCOUNT_COL - 1] || '').trim();
          const rawAmount = String(row[CHECK_CONFIG.EXCEL_AMOUNT_COL - 1] || '').trim();

          if (!rawAccount && !rawAmount) return;

          const account = rawAccount || inferAccountFromFileName_(fileInfo.fileName);
          const amount = parseAmount_(rawAmount);

          if (!account || amount === null) return;

          if (!accountMap[account]) accountMap[account] = 0;
          accountMap[account] += amount;

          detailRows.push({
            fileName: fileInfo.fileName,
            sheetName: sh.getName(),
            rowNo: CHECK_CONFIG.DATA_START_ROW + idx,
            account,
            amount,
          });
        });
      });

      log_(logs, 'INFO', 'Excel集計完了: ' + fileInfo.fileName);

    } catch (e) {
      log_(logs, 'ERROR', 'Excel集計失敗: ' + fileInfo.fileName + ' / ' + e);
    } finally {
      if (tempSpreadsheetId && CONFIG.DELETE_TEMP_CONVERTED_SHEETS) {
        try {
          DriveApp.getFileById(tempSpreadsheetId).setTrashed(true);
        } catch (e) {}
      }
    }
  }

  return {
    accountMap,
    detailRows,
  };
}

function parseAmount_(value) {
  const s = String(value || '').replace(/,/g, '').replace(/△/g, '-').trim();
  if (!s) return null;
  const n = Number(s);
  return isNaN(n) ? null : n;
}

function inferAccountFromFileName_(fileName) {
  const name = String(fileName || '');

  if (name.indexOf('預貯金') >= 0) return '現金及び預金';
  if (name.indexOf('売掛金') >= 0) return '売掛金';
  if (name.indexOf('受取手形') >= 0) return '受取手形';
  if (name.indexOf('棚卸資産') >= 0) return '仕掛品';
  if (name.indexOf('仮払金') >= 0) return '仮払金';
  if (name.indexOf('貸付金') >= 0) return '貸付金';
  if (name.indexOf('買掛金') >= 0) return '買掛金';
  if (name.indexOf('未払金') >= 0) return '未払金';
  if (name.indexOf('未払費用') >= 0) return '未払費用';
  if (name.indexOf('借入金') >= 0) return '長期借入金';
  if (name.indexOf('役員給与') >= 0) return '役員報酬';
  if (name.indexOf('地代家賃') >= 0) return '地代家賃';
  if (name.indexOf('仮受金') >= 0) return '仮受金';
  if (name.indexOf('固定資産') >= 0) return '固定資産';
  if (name.indexOf('雑益') >= 0) return '雑益';
  if (name.indexOf('雑損失') >= 0) return '雑損失';

  return '';
}

/** =========================
 * 決算書PDF解析
 * ========================= */

function extractBsMapFromPdf_(folder, pdfInfo, logs) {
  let tempDocId = '';
  try {
    const file = DriveApp.getFileById(pdfInfo.fileId);
    const blob = file.getBlob();

    const resource = {
      name: '[TEMP_PDFDOC]_' + pdfInfo.fileName + '_' + Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyyMMdd_HHmmss'),
      mimeType: MimeType.GOOGLE_DOCS,
      parents: [folder.getId()],
    };

    const converted = Drive.Files.create(resource, blob, {
      supportsAllDrives: true,
    });

    tempDocId = converted.id;

    const text = DocumentApp.openById(tempDocId).getBody().getText();
    const bsMap = parseBsAmountsFromText_(text);

    log_(logs, 'INFO', '決算書PDF解析完了: ' + pdfInfo.fileName + ' / 科目数=' + Object.keys(bsMap).length);
    return bsMap;

  } finally {
    if (tempDocId && CONFIG.DELETE_TEMP_CONVERTED_DOCS) {
      try {
        DriveApp.getFileById(tempDocId).setTrashed(true);
      } catch (e) {}
    }
  }
}

function parseBsAmountsFromText_(text) {
  const map = {};
  const lines = String(text || '').split(/\r?\n/).map(s => s.trim()).filter(Boolean);

  let inBs = false;
  for (const line of lines) {
    if (line.indexOf('貸借対照表') >= 0) {
      inBs = true;
      continue;
    }
    if (inBs && line.indexOf('損益計算書') >= 0) {
      break;
    }
    if (!inBs) continue;

    const m = line.match(/^(.+?)\s+([△\-]?\s*[\d,]+)$/);
    if (!m) continue;

    const account = normalizeAccountName_(m[1]);
    const amount = parseAmount_(m[2]);

    if (!account || amount === null) continue;
    if (isIgnoredBsAccount_(account)) continue;

    map[account] = amount;
  }

  return map;
}

function isIgnoredBsAccount_(account) {
  const ng = [
    '資産の部',
    '負債の部',
    '純資産の部',
    '流動資産',
    '固定資産',
    '流動負債',
    '固定負債',
    '株主資本',
    '利益剰余金',
    'その他利益剰余金',
    '繰越利益剰余金',
    '資産の部合計',
    '負債の部合計',
    '純資産の部合計',
    '負債・純資産の部合計',
  ];
  return ng.indexOf(account) >= 0;
}

function normalizeAccountName_(name) {
  return String(name || '')
    .replace(/[【】\[\]]/g, '')
    .replace(/\s+/g, '')
    .replace('現金預金', '現金及び預金')
    .replace('棚卸資産', '仕掛品')
    .trim();
}

/** =========================
 * 突合
 * ========================= */

function compareBsVsExcel_(bsMap, excelMap) {
  const rows = [];
  const keys = Object.keys(bsMap || {}).sort();

  keys.forEach(account => {
    const bsAmount = bsMap[account];
    const excelAmount = findExcelAmountByAccount_(excelMap, account);

    let result = '';
    let comment = '';

    if (excelAmount === null) {
      result = 'NG';
      comment = 'Excel側に対応する勘定科目が見つかりません';
    } else if (bsAmount === excelAmount) {
      result = 'OK';
      comment = '一致';
    } else {
      result = 'NG';
      comment = '残高不一致';
    }

    rows.push({
      account,
      bsAmount,
      excelAmount,
      diff: excelAmount === null ? '' : bsAmount - excelAmount,
      result,
      comment,
    });
  });

  return rows;
}

function findExcelAmountByAccount_(excelMap, account) {
  if (Object.prototype.hasOwnProperty.call(excelMap, account)) return excelMap[account];

  const aliasMap = {
    '現金及び預金': ['預貯金'],
    '長期借入金': ['借入金'],
    '役員借入金': ['役員借入金'],
    '仕掛品': ['仕掛品', '棚卸資産'],
    '前払費用': ['前払費用'],
    '敷金': ['敷金'],
    '保険積立金': ['保険積立金'],
    '買掛金': ['買掛金'],
    '未払金': ['未払金'],
    '未払費用': ['未払費用'],
    '預り金': ['預り金'],
    '売掛金': ['売掛金'],
    '受取手形': ['受取手形'],
    '仮払金': ['仮払金'],
    '貸付金': ['貸付金'],
  };

  const aliases = aliasMap[account] || [];
  for (const a of aliases) {
    if (Object.prototype.hasOwnProperty.call(excelMap, a)) return excelMap[a];
  }

  return null;
}

/** =========================
 * 結果出力
 * ========================= */

function createResultSpreadsheet_(params) {
  const folder = params.folder;
  const folderId = params.folderId;
  const scanned = params.scanned;
  const selectedPdf = params.selectedPdf;
  const excelAnalysis = params.excelAnalysis || [];
  const excelBalance = params.excelBalance || { accountMap: {}, detailRows: [] };
  const bsMap = params.bsMap || {};
  const comparisonRows = params.comparisonRows || [];
  const startedAt = params.startedAt;
  const logs = params.logs || [];

  const timestamp = Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyyMMdd_HHmmss');
  const name = CONFIG.RESULT_FILE_PREFIX + '_' + timestamp;
  const ss = SpreadsheetApp.create(name);

  const ssFile = DriveApp.getFileById(ss.getId());
  folder.addFile(ssFile);
  try {
    DriveApp.getRootFolder().removeFile(ssFile);
  } catch (e) {}

  const firstSheet = ss.getSheets()[0];
  firstSheet.setName(CONFIG.SHEET_NAMES.summary);

  const filesSheet = ss.insertSheet(CONFIG.SHEET_NAMES.files);
  const excelCheckSheet = ss.insertSheet(CONFIG.SHEET_NAMES.excelCheck);
  const adoptedSheet = ss.insertSheet(CONFIG.SHEET_NAMES.adopted);
  const excelBalanceSheet = ss.insertSheet(CONFIG.SHEET_NAMES.excelBalance);
  const bsSheet = ss.insertSheet(CONFIG.SHEET_NAMES.bs);
  const compareSheet = ss.insertSheet(CONFIG.SHEET_NAMES.compare);
  const logSheet = ss.insertSheet(CONFIG.SHEET_NAMES.log);

  writeSummarySheet_(firstSheet, {
    folder,
    folderId,
    scanned,
    selectedPdf,
    excelAnalysis,
    startedAt,
    comparisonRows,
  });

  writeFilesSheet_(filesSheet, scanned.allFiles || []);
  writeExcelCheckSheet_(excelCheckSheet, excelAnalysis);
  writeAdoptedSheet_(adoptedSheet, selectedPdf, excelAnalysis);
  writeExcelBalanceSheet_(excelBalanceSheet, excelBalance.accountMap);
  writeBsSheet_(bsSheet, bsMap);
  writeCompareSheet_(compareSheet, comparisonRows);
  writeLogSheet_(logSheet, logs);

  return ss;
}

function writeSummarySheet_(sh, data) {
  sh.clearContents();

  const adoptedExcel = (data.excelAnalysis || []).filter(x => x.isTarget);
  const skippedExcel = (data.excelAnalysis || []).filter(x => !x.isTarget);
  const ngCount = (data.comparisonRows || []).filter(x => x.result === 'NG').length;
  const okCount = (data.comparisonRows || []).filter(x => x.result === 'OK').length;

  const rows = [
    ['項目', '値'],
    ['GAS版', APP_VERSION],
    ['実行開始', data.startedAt],
    ['実行時刻', new Date()],
    ['フォルダ名', data.folder.getName()],
    ['フォルダID', data.folderId],
    ['全ファイル数', (data.scanned.allFiles || []).length],
    ['PDF件数', (data.scanned.pdfFiles || []).length],
    ['Excel件数', (data.scanned.excelFiles || []).length],
    ['採用PDF', data.selectedPdf ? data.selectedPdf.fileName : ''],
    ['採用Excel件数', adoptedExcel.length],
    ['除外Excel件数', skippedExcel.length],
    ['突合OK件数', okCount],
    ['突合NG件数', ngCount],
  ];

  sh.getRange(1, 1, rows.length, 2).setValues(rows);
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, 2);
  sh.getRange(1, 1, sh.getLastRow(), 2).setWrap(true);
}

function writeFilesSheet_(sh, files) {
  sh.clearContents();

  const headers = ['No', 'fileId', 'fileName', 'mimeType', 'size', 'updatedAt', 'url'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);

  const rows = (files || []).map((f, i) => [
    i + 1,
    f.fileId || '',
    f.fileName || '',
    f.mimeType || '',
    f.size || '',
    f.updatedAt || '',
    f.url || '',
  ]);

  if (rows.length) {
    sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, headers.length);
  sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), headers.length).setWrap(true);
}

function writeExcelCheckSheet_(sh, excelAnalysis) {
  sh.clearContents();

  const headers = [
    'No',
    'fileId',
    'fileName',
    '判定結果',
    '確認シート数',
    '5行目以降にデータあり',
    'データありシート名',
    'エラー',
    '元ファイルURL',
    '一時変換Sheet URL'
  ];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);

  const rows = (excelAnalysis || []).map((x, i) => [
    i + 1,
    x.fileId || '',
    x.fileName || '',
    x.statusJa || '',
    x.checkedSheetCount || 0,
    x.hasDataAfterRow5 ? 'あり' : 'なし',
    (x.nonEmptySheetNames || []).join('\n'),
    x.error || '',
    x.url || '',
    x.tempSpreadsheetUrl || '',
  ]);

  if (rows.length) {
    sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  sh.setFrozenRows(1);
  sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), headers.length).setWrap(true);

  if (sh.getLastRow() >= 2) {
    const range = sh.getRange(2, 4, sh.getLastRow() - 1, 1);
    const values = range.getValues();
    const bg = values.map(([v]) => {
      if (v === 'OK（使用）') return ['#d9ead3'];
      if (v === '空欄（oldへ移動）') return ['#fce5cd'];
      if (v === '判定失敗') return ['#f4cccc'];
      return ['#ffffff'];
    });
    range.setBackgrounds(bg);
  }
}

function writeAdoptedSheet_(sh, selectedPdf, excelAnalysis) {
  sh.clearContents();

  const headers = ['区分', 'fileName', 'fileId', 'url', '備考'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);

  const rows = [];

  rows.push([
    '決算書PDF',
    selectedPdf ? selectedPdf.fileName : '',
    selectedPdf ? selectedPdf.fileId : '',
    selectedPdf ? selectedPdf.url : '',
    selectedPdf ? '採用' : '未検出'
  ]);

  (excelAnalysis || []).filter(x => x.isTarget).forEach(x => {
    rows.push([
      'Excel',
      x.fileName || '',
      x.fileId || '',
      x.url || '',
      '採用'
    ]);
  });

  if (rows.length) {
    sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, headers.length);
  sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), headers.length).setWrap(true);
}

function writeExcelBalanceSheet_(sh, accountMap) {
  sh.clearContents();
  const headers = ['勘定科目', 'Excel集計残高'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);

  const keys = Object.keys(accountMap || {}).sort();
  const rows = keys.map(k => [k, accountMap[k]]);

  if (rows.length) {
    sh.getRange(2, 1, rows.length, 2).setValues(rows);
  }

  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, 2);
}

function writeBsSheet_(sh, bsMap) {
  sh.clearContents();
  const headers = ['勘定科目', '決算書残高'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);

  const keys = Object.keys(bsMap || {}).sort();
  const rows = keys.map(k => [k, bsMap[k]]);

  if (rows.length) {
    sh.getRange(2, 1, rows.length, 2).setValues(rows);
  }

  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, 2);
}

function writeCompareSheet_(sh, rowsData) {
  sh.clearContents();
  const headers = ['勘定科目', '決算書残高', 'Excel集計残高', '差額', '判定', 'コメント'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);

  const rows = (rowsData || []).map(r => [
    r.account || '',
    r.bsAmount || '',
    r.excelAmount === null ? '' : r.excelAmount,
    r.diff === '' ? '' : r.diff,
    r.result || '',
    r.comment || '',
  ]);

  if (rows.length) {
    sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, headers.length);
  sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), headers.length).setWrap(true);

  if (sh.getLastRow() >= 2) {
    const range = sh.getRange(2, 5, sh.getLastRow() - 1, 1);
    const values = range.getValues();
    const bg = values.map(([v]) => [v === 'OK' ? '#d9ead3' : '#f4cccc']);
    range.setBackgrounds(bg);
  }
}

function writeLogSheet_(sh, logs) {
  sh.clearContents();

  const headers = ['時刻', '区分', '内容'];
  sh.getRange(1, 1, 1, 3).setValues([headers]);

  const rows = (logs || []).map(l => [l.at, l.level, l.message]);
  if (rows.length) {
    sh.getRange(2, 1, rows.length, 3).setValues(rows);
  }

  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, 3);
  sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 3).setWrap(true);
}

/** =========================
 * 補助
 * ========================= */

function log_(logs, level, message) {
  logs.push({
    at: new Date(),
    level,
    message,
  });
}

function escapeHtml_(s) {
  return String(s || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function renderHtmlResult_(result) {
  return [
    '<h3>決算書Excelチェック起動完了</h3>',
    '<p>フォルダ名: ' + escapeHtml_(result.folderName || '') + '</p>',
    '<p>PDF件数: ' + escapeHtml_(String(result.pdfCount || 0)) + '</p>',
    '<p>Excel件数: ' + escapeHtml_(String(result.excelCount || 0)) + '</p>',
    '<p>採用Excel件数: ' + escapeHtml_(String(result.adoptedExcelCount || 0)) + '</p>',
    '<p>除外Excel件数: ' + escapeHtml_(String(result.skippedExcelCount || 0)) + '</p>',
    '<p>採用PDF: ' + escapeHtml_(result.selectedPdfName || '') + '</p>',
    '<p><a href="' + result.resultSpreadsheetUrl + '" target="_blank">結果スプレッドシートを開く</a></p>',
  ].join('');
}
