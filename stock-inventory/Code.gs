// 股票庫存試算表（可編輯版）— Google Apps Script
// 前往 script.google.com → 新增空白專案 → 貼入全部內容 → 執行 createStockSpreadsheet()
// 建立完成後，再執行一次 setupDailyTrigger() 以啟用每日 13:30 自動更新。

// ─── 全域常數 ───────────────────────────────────────────────────────────────────
var STOCK_START      = 20;   // 持股資料起始列
var STOCK_END        = 50;   // 持股資料結束列（最多 31 檔）
var EXCH_ROW         = 53;   // 匯率區塊起始列
var CHART_DATA_COL   = 14;   // 圖表輔助資料欄（N 欄，隱藏）
var CHART_ANCHOR_ROW = 1;
var CHART_ANCHOR_COL = 8;    // 欄 H

// ─── 主函式 ────────────────────────────────────────────────────────────────────
function createStockSpreadsheet() {
  var ss = SpreadsheetApp.create('股票庫存試算表');
  var sh = ss.getActiveSheet();
  sh.setName('股票庫存');

  _setupMarketIndices(sh);
  _setupPortfolioSummary(sh);
  _setupInstructionRow(sh);
  _setupStockHeaders(sh);
  _setupStockFormulas(sh);
  _setupSampleStocks(sh);
  _setupExchangeRates(sh);
  _setupChartDataHelper(sh);   // 圖表輔助欄（N欄）
  _applyFormatting(sh);
  _createDonutChart(sh);

  // 命名儲存格（USD_TWD 供市值公式使用）
  ss.setNamedRange('USD_TWD', sh.getRange(EXCH_ROW + 1, 2));

  Logger.log('試算表網址: ' + ss.getUrl());
  SpreadsheetApp.getUi().alert(
    '試算表已建立！\n\n' + ss.getUrl() +
    '\n\n請再執行一次 setupDailyTrigger() 以啟用每日 13:30 自動更新。'
  );
}

// ─── 每日 13:30 自動更新（執行一次即永久啟用）──────────────────────────────────
function setupDailyTrigger() {
  // 先刪除所有舊 trigger，避免重複
  ScriptApp.getProjectTriggers().forEach(function(t) {
    ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger('refreshPortfolio')
    .timeBased()
    .atHour(13)
    .nearMinute(30)
    .everyDays(1)
    .inTimezone('Asia/Taipei')
    .create();

  Logger.log('每日 13:30 (台北時間) 自動更新已啟用');
  SpreadsheetApp.getUi().alert('每日 13:30 (台北時間) 自動更新已啟用！');
}

// 自動更新觸發的函式：強制重新整理 GOOGLEFINANCE 公式
function refreshPortfolio() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('股票庫存');
  if (!sh) return;

  // 寫入時間戳記到隱藏欄強制試算表重新計算
  var ts = sh.getRange(1, 13); // M1（隱藏）
  ts.setValue(new Date());
  SpreadsheetApp.flush();

  // 重新寫入 B 欄（股價）和 J 欄（今日漲跌%）公式以強制更新
  for (var r = STOCK_START; r <= STOCK_END; r++) {
    var a = sh.getRange(r, 1).getValue();
    if (!a) continue;
    sh.getRange(r, 2).setFormula(
      '=IFERROR(IF(A' + r + '="","",GOOGLEFINANCE(A' + r + ',"price")),"")'
    );
    sh.getRange(r, 10).setFormula(
      '=IFERROR(IF(A' + r + '="","",GOOGLEFINANCE(A' + r + ',"changepct")/100),"")'
    );
  }
  SpreadsheetApp.flush();

  Logger.log('Portfolio refreshed at ' + new Date());
}

// ─── 1. 大盤指數（列 1–5）───────────────────────────────────────────────────────
function _setupMarketIndices(sh) {
  sh.getRange('A1').setValue('大盤指數');
  sh.getRange('C1').setValue('價格');
  sh.getRange('D1').setValue('今日漲跌%');

  var indices = [
    ['台灣加權指數',   'TPE:IX0001'],
    ['美國道瓊指數',   'INDEXDJX:.DJI'],
    ['那斯達克',       'INDEXNASDAQ:IXIC'],
    ['費城半導體指數', 'INDEXNASDAQ:SOX'],
  ];
  for (var i = 0; i < indices.length; i++) {
    var r = i + 2;
    sh.getRange(r, 1).setValue(indices[i][0]);
    sh.getRange(r, 2).setValue(indices[i][1]);
    sh.getRange(r, 3).setFormula('=IFERROR(GOOGLEFINANCE("' + indices[i][1] + '","price"),"")');
    sh.getRange(r, 4).setFormula('=IFERROR(GOOGLEFINANCE("' + indices[i][1] + '","changepct")/100,"")');
  }
}

// ─── 2. 持股分類彙總（列 7–11）─────────────────────────────────────────────────
function _setupPortfolioSummary(sh) {
  sh.getRange(7, 2, 1, 5).setValues([['成本(台幣)', '損益', '報酬%', '市值', '今日損益']]);

  var categories = [['美股', 8], ['台股', 9], ['英股', 10]];
  categories.forEach(function(cat) {
    var name = cat[0], row = cat[1];
    var s = STOCK_START, e = STOCK_END;
    sh.getRange(row, 1).setValue(name);
    sh.getRange(row, 2).setFormula('=IFERROR(SUMIF($L$' + s + ':$L$' + e + ',"' + name + '",$D$' + s + ':$D$' + e + '),0)');
    sh.getRange(row, 3).setFormula('=IFERROR(SUMIF($L$' + s + ':$L$' + e + ',"' + name + '",$G$' + s + ':$G$' + e + '),0)');
    sh.getRange(row, 4).setFormula('=IFERROR(C' + row + '/B' + row + ',"")');
    sh.getRange(row, 5).setFormula('=IFERROR(SUMIF($L$' + s + ':$L$' + e + ',"' + name + '",$I$' + s + ':$I$' + e + '),0)');
    sh.getRange(row, 6).setFormula('=IFERROR(SUMIF($L$' + s + ':$L$' + e + ',"' + name + '",$K$' + s + ':$K$' + e + '),0)');
  });

  sh.getRange(11, 1).setValue('Total');
  sh.getRange(11, 2).setFormula('=SUM(B8:B10)');
  sh.getRange(11, 3).setFormula('=SUM(C8:C10)');
  sh.getRange(11, 4).setFormula('=IFERROR(C11/B11,"")');
  sh.getRange(11, 5).setFormula('=SUM(E8:E10)');
  sh.getRange(11, 6).setFormula('=SUM(F8:F10)');
}

// ─── 3. 提示列（列 18）────────────────────────────────────────────────────────
function _setupInstructionRow(sh) {
  sh.getRange('A18').setValue('填寫藍色方塊位置（股票代號、股數、成本台幣、成本美金）');
}

// ─── 4. 持股表格標題（列 19）───────────────────────────────────────────────────
function _setupStockHeaders(sh) {
  var h = ['股票', '股價', '股數', '成本(台幣)', '成本(美金)',
           '均價', '損益', '報酬%', '市值', '今日漲跌%', '今日損益'];
  sh.getRange(19, 1, 1, h.length).setValues([h]);
}

// ─── 5. 持股公式（列 20–50）────────────────────────────────────────────────────
function _setupStockFormulas(sh) {
  for (var r = STOCK_START; r <= STOCK_END; r++) {
    var a = 'A'+r, b = 'B'+r, c = 'C'+r,
        d = 'D'+r, e = 'E'+r, g = 'G'+r,
        i = 'I'+r, j = 'J'+r;

    // B: 股價
    sh.getRange(r, 2).setFormula(
      '=IFERROR(IF('+a+'="","",GOOGLEFINANCE('+a+',"price")),"")'
    );
    // F: 均價（外股用美金成本/股數；台股用台幣成本/股數）
    sh.getRange(r, 6).setFormula(
      '=IFERROR(IF('+a+'="","",IF(OR('+e+'="",'+e+'=0),'+d+'/'+c+','+e+'/'+c+')),"")'
    );
    // I: 市值（台股直接×股數；外股×股數×匯率）
    sh.getRange(r, 9).setFormula(
      '=IFERROR(IF('+a+'="","",IF(OR('+e+'="",'+e+'=0),'+b+'*'+c+','+b+'*'+c+'*USD_TWD)),"")'
    );
    // G: 損益 = 市值 - 成本台幣
    sh.getRange(r, 7).setFormula('=IFERROR(IF('+a+'="","",'+i+'-'+d+'),"")');
    // H: 報酬%
    sh.getRange(r, 8).setFormula('=IFERROR(IF('+a+'="","",'+g+'/'+d+'),"")');
    // J: 今日漲跌%
    sh.getRange(r, 10).setFormula(
      '=IFERROR(IF('+a+'="","",GOOGLEFINANCE('+a+',"changepct")/100),"")'
    );
    // K: 今日損益
    sh.getRange(r, 11).setFormula('=IFERROR(IF('+a+'="","",'+i+'*'+j+'),"")');
    // L（隱藏）: 分類
    sh.getRange(r, 12).setFormula(
      '=IF('+a+'="","",IF(ISNUMBER(FIND("LON:",'+a+')),"英股",' +
      'IF(OR(ISNUMBER(FIND("TPE:",'+a+')),ISNUMBER(FIND(".TW",'+a+'))),"台股","美股")))'
    );
  }
}

// ─── 6. 範例持股（藍色欄位）────────────────────────────────────────────────────
function _setupSampleStocks(sh) {
  var samples = [
    ['VOO',       1,    10000,  316.69],
    ['VT',        1,    4000,   126.68],
    ['NVDA',      1,    6000,   190.01],
    ['NVDA',      1,    6315.3, 200.00],
    ['TPE:0050',  1000, 90000,  ''],
    ['00679B.TW', 1000, 26000,  0],
    ['LON:VWRA',  1,    5000,   158.35],
    ['LON:VUAA',  1,    3000,   95.01],
  ];
  samples.forEach(function(s, idx) {
    var r = STOCK_START + idx;
    sh.getRange(r, 1).setValue(s[0]);
    sh.getRange(r, 3).setValue(s[1]);
    sh.getRange(r, 4).setValue(s[2]);
    if (s[3] !== '') sh.getRange(r, 5).setValue(s[3]);
  });
  // 冷門股提示
  sh.getRange(STOCK_START + 5, 13)
    .setValue('提示：GOOGLEFINANCE 無法抓取此股票時，請手動在 B 欄輸入股價，或改用 IMPORTXML。')
    .setFontColor('#cc0000').setFontStyle('italic');
}

// ─── 7. 匯率區塊 ───────────────────────────────────────────────────────────────
function _setupExchangeRates(sh) {
  sh.getRange(EXCH_ROW,     1).setValue('美金兌換匯率');
  sh.getRange(EXCH_ROW,     2).setValue('USD');
  sh.getRange(EXCH_ROW + 1, 1).setValue('TWD');
  sh.getRange(EXCH_ROW + 1, 2).setFormula('=IFERROR(GOOGLEFINANCE("CURRENCY:USDTWD"),"")');
  sh.getRange(EXCH_ROW + 2, 1).setValue('JPY');
  sh.getRange(EXCH_ROW + 2, 2).setFormula('=IFERROR(GOOGLEFINANCE("CURRENCY:USDJPY"),"")');
  sh.getRange(EXCH_ROW + 3, 1).setValue('CNY');
  sh.getRange(EXCH_ROW + 3, 2).setFormula('=IFERROR(GOOGLEFINANCE("CURRENCY:USDCNY"),"")');
}

// ─── 8. 圖表輔助欄（N 欄）— 讓甜甜圈圖讀取相鄰的標籤+數值 ──────────────────────
// 直接用兩個不相鄰 addRange 抓 A 和 I 欄，Google Charts 有時會把第一欄當 X 軸而非標籤；
// 把資料複製到相鄰的 N:O 欄可確保圖表正確繪製。
function _setupChartDataHelper(sh) {
  sh.getRange(STOCK_START - 1, CHART_DATA_COL    ).setValue('股票');   // N19
  sh.getRange(STOCK_START - 1, CHART_DATA_COL + 1).setValue('市值');   // O19

  for (var r = STOCK_START; r <= STOCK_END; r++) {
    sh.getRange(r, CHART_DATA_COL    ).setFormula('=IF(A'+r+'="","",A'+r+')');   // N欄
    sh.getRange(r, CHART_DATA_COL + 1).setFormula('=IF(I'+r+'="",0,I'+r+')');   // O欄
  }
}

// ─── 9. 格式設定 ───────────────────────────────────────────────────────────────
function _applyFormatting(sh) {
  sh.setColumnWidth(1,  110);  // A
  sh.setColumnWidth(2,  105);  // B
  sh.setColumnWidth(3,  70);   // C
  sh.setColumnWidth(4,  105);  // D
  sh.setColumnWidth(5,  105);  // E
  sh.setColumnWidth(6,  90);   // F
  sh.setColumnWidth(7,  100);  // G
  sh.setColumnWidth(8,  80);   // H
  sh.setColumnWidth(9,  110);  // I
  sh.setColumnWidth(10, 105);  // J
  sh.setColumnWidth(11, 105);  // K
  sh.setColumnWidth(12, 2);    // L 隱藏分類
  sh.setColumnWidth(13, 2);    // M 時間戳（隱藏）
  sh.setColumnWidth(CHART_DATA_COL,     2);  // N 圖表輔助（隱藏）
  sh.setColumnWidth(CHART_DATA_COL + 1, 2);  // O 圖表輔助（隱藏）

  // 大盤指數
  sh.getRange('A1').setBackground('#1a1a2e').setFontColor('#ffffff').setFontWeight('bold');
  sh.getRange('C1:D1').setBackground('#434343').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange('A2:A5').setFontWeight('bold');
  sh.getRange('C2:C5').setNumberFormat('#,##0.00');
  sh.getRange('D2:D5').setNumberFormat('0.00%');
  _cfPnl(sh, 'D2:D5');

  // 彙總區
  sh.getRange('B7:G7').setBackground('#434343').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange('A8:G8').setBackground('#FF9900').setFontWeight('bold');
  sh.getRange('A9:G9').setBackground('#6FA8DC').setFontWeight('bold');
  sh.getRange('A10:G10').setBackground('#93C47D').setFontWeight('bold');
  sh.getRange('A11:G11').setBackground('#FFD966').setFontWeight('bold');
  sh.getRange('B8:B11').setNumberFormat('#,##0');
  sh.getRange('C8:C11').setNumberFormat('#,##0');
  sh.getRange('D8:D11').setNumberFormat('0.00%');
  sh.getRange('E8:E11').setNumberFormat('#,##0');
  sh.getRange('F8:F11').setNumberFormat('#,##0');
  _cfNeg(sh, 'F8:F11');

  // 提示
  sh.getRange('A18').setFontColor('#cc0000').setFontStyle('italic').setFontWeight('bold');

  // 持股標題
  sh.getRange(19, 1, 1, 11)
    .setBackground('#1a1a2e').setFontColor('#ffffff')
    .setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(19, 32);

  // 持股資料列
  var editBg = '#cfe2f3'; // 藍色可編輯欄
  for (var r = STOCK_START; r <= STOCK_END; r++) {
    sh.getRange(r, 1, 1, 11).setBackground(r % 2 === 0 ? '#f8f9fa' : '#ffffff');
    sh.getRange(r, 1).setBackground(editBg);
    sh.getRange(r, 3).setBackground(editBg);
    sh.getRange(r, 4).setBackground(editBg);
    sh.getRange(r, 5).setBackground(editBg);
  }

  var s = STOCK_START, len = STOCK_END - STOCK_START + 1;
  sh.getRange(s, 2, len).setNumberFormat('#,##0.00');   // 股價
  sh.getRange(s, 3, len).setNumberFormat('#,##0');      // 股數
  sh.getRange(s, 4, len, 2).setNumberFormat('#,##0.00');// 成本
  sh.getRange(s, 6, len).setNumberFormat('#,##0.00');   // 均價
  sh.getRange(s, 7, len).setNumberFormat('#,##0.00');   // 損益
  sh.getRange(s, 8, len).setNumberFormat('0.00%');      // 報酬%
  sh.getRange(s, 9, len).setNumberFormat('#,##0.00');   // 市值
  sh.getRange(s, 10, len).setNumberFormat('0.00%');     // 今日漲跌%
  sh.getRange(s, 11, len).setNumberFormat('#,##0.00');  // 今日損益

  _cfPnl(sh, 'G'+s+':G'+STOCK_END+',H'+s+':H'+STOCK_END);
  _cfPnl(sh, 'J'+s+':J'+STOCK_END+',K'+s+':K'+STOCK_END);

  // 匯率
  sh.getRange(EXCH_ROW, 1, 1, 2).setBackground('#ffe599').setFontWeight('bold');
  sh.getRange(EXCH_ROW + 1, 2, 3).setNumberFormat('#,##0.0000');

  sh.setFrozenRows(19);
}

// ─── 10. 甜甜圈圖 ──────────────────────────────────────────────────────────────
// 使用相鄰的 N:O 欄（圖表輔助欄），確保標籤和數值正確對應
function _createDonutChart(sh) {
  var dataRange = sh.getRange(STOCK_START - 1, CHART_DATA_COL,
                               STOCK_END - STOCK_START + 2, 2); // N19:O50

  var chart = sh.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(dataRange)
    .setOption('title', '股票總資產')
    .setOption('titleTextStyle', { fontSize: 14, bold: true, color: '#1a1a2e' })
    .setOption('pieHole', 0.4)
    .setOption('is3D', false)
    .setOption('width', 520)
    .setOption('height', 380)
    .setOption('legend', { position: 'labeled' })
    .setOption('pieSliceText', 'percentage')
    .setPosition(CHART_ANCHOR_ROW, CHART_ANCHOR_COL, 10, 10)
    .build();

  sh.insertChart(chart);
}

// ─── 條件格式工具 ───────────────────────────────────────────────────────────────
function _cfPnl(sh, notation) {
  var ranges = notation.split(',').map(function(n) { return sh.getRange(n.trim()); });
  var rules = sh.getConditionalFormatRules();
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0).setFontColor('#137333').setBackground('#d9ead3')
      .setRanges(ranges).build()
  );
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0).setFontColor('#cc0000').setBackground('#fce8e6')
      .setRanges(ranges).build()
  );
  sh.setConditionalFormatRules(rules);
}

function _cfNeg(sh, notation) {
  var ranges = notation.split(',').map(function(n) { return sh.getRange(n.trim()); });
  var rules = sh.getConditionalFormatRules();
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0).setFontColor('#cc0000').setBackground('#fce8e6')
      .setRanges(ranges).build()
  );
  sh.setConditionalFormatRules(rules);
}
