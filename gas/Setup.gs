/**
 * BSCA CUP タイムライン自動生成ツール - 初期セットアップ
 */

/**
 * スプレッドシートの初期セットアップ
 */
function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 入力シートを作成
  createInputSheet(ss);
  
  // 出力シートを作成
  createOutputSheets(ss);
  
  // メニューを追加
  onOpen();
  
  SpreadsheetApp.getUi().alert('セットアップが完了しました！\n「入力」シートに必要事項を入力し、メニューから「タイムライン生成」を実行してください。');
}

/**
 * 入力シートを作成
 */
function createInputSheet(ss) {
  let inputSheet = ss.getSheetByName(CONFIG.SHEETS.INPUT);
  
  if (!inputSheet) {
    inputSheet = ss.insertSheet(CONFIG.SHEETS.INPUT, 0);
  } else {
    inputSheet.clear();
  }
  
  // タイトル
  inputSheet.getRange('A1').setValue('BSCA CUP タイムライン生成ツール');
  inputSheet.getRange('A1').setFontSize(16).setFontWeight('bold');
  inputSheet.getRange('A1:D1').merge();
  
  // 大会共通セクション
  inputSheet.getRange('A2').setValue('【大会共通設定】');
  inputSheet.getRange('A2').setFontWeight('bold').setBackground('#e8f5e9');
  inputSheet.getRange('A2:D2').merge();
  
  const commonLabels = [
    ['B3', '利用開始時刻'],
    ['B4', '利用終了時刻'],
    ['B5', '日数'],
    ['B6', 'コート数'],
    ['B7', 'エキシビジョン'],
    ['B8', '表彰式'],
    ['B9', '写真撮影']
  ];
  
  for (const [cell, label] of commonLabels) {
    inputSheet.getRange(cell).setValue(label);
  }
  
  // デフォルト値
  inputSheet.getRange(INPUT_CELLS.DAYS).setValue(1);
  inputSheet.getRange(INPUT_CELLS.COURTS).setValue(1);
  inputSheet.getRange(INPUT_CELLS.HAS_EXHIBITION).setValue('OFF');
  inputSheet.getRange(INPUT_CELLS.HAS_CEREMONY).setValue('ON');
  inputSheet.getRange(INPUT_CELLS.HAS_PHOTO).setValue('ON');
  
  // U12セクション
  inputSheet.getRange('A11').setValue('【U12 設定】');
  inputSheet.getRange('A11').setFontWeight('bold').setBackground('#e3f2fd');
  inputSheet.getRange('A11:D11').merge();
  
  const u12Labels = [
    ['B12', '開催'],
    ['B13', 'チーム数'],
    ['B14', '試合方式']
  ];
  
  for (const [cell, label] of u12Labels) {
    inputSheet.getRange(cell).setValue(label);
  }
  
  inputSheet.getRange(INPUT_CELLS.U12_ENABLED).setValue('OFF');
  
  // U15セクション
  inputSheet.getRange('A16').setValue('【U15 設定】');
  inputSheet.getRange('A16').setFontWeight('bold').setBackground('#fff3e0');
  inputSheet.getRange('A16:D16').merge();
  
  const u15Labels = [
    ['B17', '開催'],
    ['B18', 'チーム数'],
    ['B19', '試合方式']
  ];
  
  for (const [cell, label] of u15Labels) {
    inputSheet.getRange(cell).setValue(label);
  }
  
  inputSheet.getRange(INPUT_CELLS.U15_ENABLED).setValue('OFF');
  
  // 説明セクション
  inputSheet.getRange('A21').setValue('【入力ガイド】');
  inputSheet.getRange('A21').setFontWeight('bold');
  inputSheet.getRange('A21:D21').merge();
  
  const guide = [
    '・時刻は「9:00」のように入力してください',
    '・日数: 1 または 2',
    '・コート数: 1 または 2',
    '・開催: ON / OFF',
    '・試合方式: リーグ戦 / トーナメント',
    '',
    '入力後、メニュー「BSCA CUP」→「タイムライン生成」を実行'
  ];
  
  for (let i = 0; i < guide.length; i++) {
    inputSheet.getRange(22 + i, 1).setValue(guide[i]);
    inputSheet.getRange(22 + i, 1, 1, 4).merge();
  }
  
  // ドロップダウンの設定
  setupDropdowns(inputSheet);
  
  // 書式設定
  formatInputSheet(inputSheet);
}

/**
 * ドロップダウンを設定
 */
function setupDropdowns(sheet) {
  // 日数
  const daysRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['1', '2'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(INPUT_CELLS.DAYS).setDataValidation(daysRule);
  
  // コート数
  const courtsRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['1', '2'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(INPUT_CELLS.COURTS).setDataValidation(courtsRule);
  
  // ON/OFF選択
  const onOffRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['ON', 'OFF'], true)
    .setAllowInvalid(false)
    .build();
  
  sheet.getRange(INPUT_CELLS.HAS_EXHIBITION).setDataValidation(onOffRule);
  sheet.getRange(INPUT_CELLS.HAS_CEREMONY).setDataValidation(onOffRule);
  sheet.getRange(INPUT_CELLS.HAS_PHOTO).setDataValidation(onOffRule);
  sheet.getRange(INPUT_CELLS.U12_ENABLED).setDataValidation(onOffRule);
  sheet.getRange(INPUT_CELLS.U15_ENABLED).setDataValidation(onOffRule);
  
  // 試合方式
  const matchTypeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList([CONFIG.MATCH_TYPES.LEAGUE, CONFIG.MATCH_TYPES.TOURNAMENT], true)
    .setAllowInvalid(false)
    .build();
  
  sheet.getRange(INPUT_CELLS.U12_MATCH_TYPE).setDataValidation(matchTypeRule);
  sheet.getRange(INPUT_CELLS.U15_MATCH_TYPE).setDataValidation(matchTypeRule);
  
  // チーム数（2〜16）
  const teamsRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13', '14', '15', '16'], true)
    .setAllowInvalid(false)
    .build();
  
  sheet.getRange(INPUT_CELLS.U12_TEAMS).setDataValidation(teamsRule);
  sheet.getRange(INPUT_CELLS.U15_TEAMS).setDataValidation(teamsRule);
}

/**
 * 入力シートの書式設定
 */
function formatInputSheet(sheet) {
  // 列幅
  sheet.setColumnWidth(1, 30);   // A列（余白）
  sheet.setColumnWidth(2, 150);  // B列（ラベル）
  sheet.setColumnWidth(3, 150);  // C列（入力）
  sheet.setColumnWidth(4, 200);  // D列（補足）
  
  // 入力セルの背景色
  const inputCells = [
    INPUT_CELLS.START_TIME, INPUT_CELLS.END_TIME, INPUT_CELLS.DAYS, INPUT_CELLS.COURTS,
    INPUT_CELLS.HAS_EXHIBITION, INPUT_CELLS.HAS_CEREMONY, INPUT_CELLS.HAS_PHOTO,
    INPUT_CELLS.U12_ENABLED, INPUT_CELLS.U12_TEAMS, INPUT_CELLS.U12_MATCH_TYPE,
    INPUT_CELLS.U15_ENABLED, INPUT_CELLS.U15_TEAMS, INPUT_CELLS.U15_MATCH_TYPE
  ];
  
  for (const cell of inputCells) {
    sheet.getRange(cell).setBackground('#fffde7'); // 薄い黄色
    sheet.getRange(cell).setBorder(true, true, true, true, false, false, '#bdbdbd', SpreadsheetApp.BorderStyle.SOLID);
  }
  
  // ラベルの右寄せ
  sheet.getRange('B3:B19').setHorizontalAlignment('right');
  
  // 入力セルの中央揃え
  for (const cell of inputCells) {
    sheet.getRange(cell).setHorizontalAlignment('center');
  }
}

/**
 * 出力シートを作成
 */
function createOutputSheets(ss) {
  // タイムラインシート
  let timelineSheet = ss.getSheetByName(CONFIG.SHEETS.TIMELINE);
  if (!timelineSheet) {
    timelineSheet = ss.insertSheet(CONFIG.SHEETS.TIMELINE);
  }
  timelineSheet.clear();
  timelineSheet.getRange('A1').setValue('タイムライン生成後に表示されます');
  
  // 試合スケジュールシート
  let scheduleSheet = ss.getSheetByName(CONFIG.SHEETS.SCHEDULE);
  if (!scheduleSheet) {
    scheduleSheet = ss.insertSheet(CONFIG.SHEETS.SCHEDULE);
  }
  scheduleSheet.clear();
  scheduleSheet.getRange('A1').setValue('タイムライン生成後に表示されます');
}

/**
 * 生成ボタンを追加（図形として）
 */
function addGenerateButton() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName(CONFIG.SHEETS.INPUT);
  
  if (!inputSheet) {
    SpreadsheetApp.getUi().alert('入力シートが見つかりません。先にセットアップを実行してください。');
    return;
  }
  
  // ボタンはメニューから実行する形式にしているため、
  // 代わりに説明テキストを追加
  inputSheet.getRange('E3').setValue('←入力後、メニュー「BSCA CUP」→「タイムライン生成」を実行');
  inputSheet.getRange('E3').setFontColor('#666666');
}

/**
 * サンプルデータを入力（テスト用）
 */
function inputSampleData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName(CONFIG.SHEETS.INPUT);
  
  if (!inputSheet) {
    SpreadsheetApp.getUi().alert('入力シートが見つかりません。先にセットアップを実行してください。');
    return;
  }
  
  // サンプルデータを入力
  inputSheet.getRange(INPUT_CELLS.START_TIME).setValue('8:00');
  inputSheet.getRange(INPUT_CELLS.END_TIME).setValue('18:00');
  inputSheet.getRange(INPUT_CELLS.DAYS).setValue(1);
  inputSheet.getRange(INPUT_CELLS.COURTS).setValue(2);
  inputSheet.getRange(INPUT_CELLS.HAS_EXHIBITION).setValue('OFF');
  inputSheet.getRange(INPUT_CELLS.HAS_CEREMONY).setValue('ON');
  inputSheet.getRange(INPUT_CELLS.HAS_PHOTO).setValue('ON');
  
  inputSheet.getRange(INPUT_CELLS.U12_ENABLED).setValue('ON');
  inputSheet.getRange(INPUT_CELLS.U12_TEAMS).setValue(4);
  inputSheet.getRange(INPUT_CELLS.U12_MATCH_TYPE).setValue(CONFIG.MATCH_TYPES.LEAGUE);
  
  inputSheet.getRange(INPUT_CELLS.U15_ENABLED).setValue('ON');
  inputSheet.getRange(INPUT_CELLS.U15_TEAMS).setValue(4);
  inputSheet.getRange(INPUT_CELLS.U15_MATCH_TYPE).setValue(CONFIG.MATCH_TYPES.LEAGUE);
  
  SpreadsheetApp.getUi().alert('サンプルデータを入力しました。\nメニュー「BSCA CUP」→「タイムライン生成」を実行してください。');
}
