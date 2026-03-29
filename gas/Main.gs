/**
 * BSCA CUP タイムライン自動生成ツール - メイン処理
 */

/**
 * メニューを追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('BSCA CUP')
    .addItem('タイムライン生成', 'generateTimelineMain')
    .addItem('入力をクリア', 'clearInput')
    .addItem('出力をクリア', 'clearOutput')
    .addToUi();
}

/**
 * メイン処理：タイムライン生成
 */
function generateTimelineMain() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  try {
    // 1. 入力値を読み込み
    const input = readInputValues(ss);
    
    // 2. 入力値のバリデーション
    const validationResult = validateInput(input);
    if (!validationResult.isValid) {
      showError(validationResult.errors.join('\n'));
      return;
    }
    
    // 3. 最適なレギュレーションを選択
    const regulations = selectOptimalRegulations(input);
    if (!regulations) {
      const error = generateTimeError(input, null);
      showError(error.message);
      return;
    }
    
    // 4. 試合スケジュールを生成
    const schedule = generateSchedule(input, regulations);
    
    // 5. スケジュールの検証
    const scheduleValidation = validateSchedule(schedule, input);
    if (!scheduleValidation.isValid) {
      showError(scheduleValidation.errors.join('\n'));
      return;
    }
    
    // 警告があれば表示
    if (scheduleValidation.warnings.length > 0) {
      ui.alert('警告', scheduleValidation.warnings.join('\n'), ui.ButtonSet.OK);
    }
    
    // 6. タイムラインを生成
    const timeline = generateTimeline(input, schedule);
    
    // 7. 出力
    outputResults(ss, input, schedule, timeline, regulations);
    
    // 8. 完了メッセージ
    const summary = generateSummary(input, schedule, timeline, regulations);
    showSuccess(summary);
    
  } catch (e) {
    showError('エラーが発生しました: ' + e.message);
    Logger.log(e);
  }
}

/**
 * 入力値を読み込み
 */
function readInputValues(ss) {
  const inputSheet = ss.getSheetByName(CONFIG.SHEETS.INPUT);
  
  if (!inputSheet) {
    throw new Error('入力シートが見つかりません');
  }
  
  const input = createInputStructure();
  
  // 大会共通
  input.common.startTime = parseTime(inputSheet.getRange(INPUT_CELLS.START_TIME).getValue());
  input.common.endTime = parseTime(inputSheet.getRange(INPUT_CELLS.END_TIME).getValue());
  input.common.days = inputSheet.getRange(INPUT_CELLS.DAYS).getValue() || 1;
  input.common.courts = inputSheet.getRange(INPUT_CELLS.COURTS).getValue() || 1;
  input.common.hasExhibition = inputSheet.getRange(INPUT_CELLS.HAS_EXHIBITION).getValue() === true || 
                                inputSheet.getRange(INPUT_CELLS.HAS_EXHIBITION).getValue() === 'ON';
  input.common.hasCeremony = inputSheet.getRange(INPUT_CELLS.HAS_CEREMONY).getValue() !== false && 
                              inputSheet.getRange(INPUT_CELLS.HAS_CEREMONY).getValue() !== 'OFF';
  input.common.hasPhoto = inputSheet.getRange(INPUT_CELLS.HAS_PHOTO).getValue() !== false && 
                           inputSheet.getRange(INPUT_CELLS.HAS_PHOTO).getValue() !== 'OFF';
  
  // U12
  const u12Enabled = inputSheet.getRange(INPUT_CELLS.U12_ENABLED).getValue();
  input.u12.enabled = u12Enabled === true || u12Enabled === 'ON';
  input.u12.teams = inputSheet.getRange(INPUT_CELLS.U12_TEAMS).getValue() || 0;
  input.u12.matchType = inputSheet.getRange(INPUT_CELLS.U12_MATCH_TYPE).getValue() || CONFIG.MATCH_TYPES.LEAGUE;
  
  // U15
  const u15Enabled = inputSheet.getRange(INPUT_CELLS.U15_ENABLED).getValue();
  input.u15.enabled = u15Enabled === true || u15Enabled === 'ON';
  input.u15.teams = inputSheet.getRange(INPUT_CELLS.U15_TEAMS).getValue() || 0;
  input.u15.matchType = inputSheet.getRange(INPUT_CELLS.U15_MATCH_TYPE).getValue() || CONFIG.MATCH_TYPES.LEAGUE;
  
  return input;
}

/**
 * 時刻文字列をDate型に変換
 */
function parseTime(value) {
  if (value instanceof Date) {
    return value;
  }
  
  if (typeof value === 'string') {
    const parts = value.split(':');
    if (parts.length >= 2) {
      const date = new Date();
      date.setHours(parseInt(parts[0], 10));
      date.setMinutes(parseInt(parts[1], 10));
      date.setSeconds(0);
      date.setMilliseconds(0);
      return date;
    }
  }
  
  return null;
}

/**
 * 入力値のバリデーション
 */
function validateInput(input) {
  const errors = [];
  
  // 時刻チェック
  if (!input.common.startTime) {
    errors.push('利用開始時刻を入力してください');
  }
  if (!input.common.endTime) {
    errors.push('利用終了時刻を入力してください');
  }
  if (input.common.startTime && input.common.endTime && 
      input.common.startTime >= input.common.endTime) {
    errors.push('利用終了時刻は開始時刻より後にしてください');
  }
  
  // コート数チェック
  if (input.common.courts < 1 || input.common.courts > 2) {
    errors.push('コート数は1または2を指定してください');
  }
  
  // カテゴリチェック
  if (!input.u12.enabled && !input.u15.enabled) {
    errors.push('U12またはU15のいずれかを有効にしてください');
  }
  
  // U12チェック
  if (input.u12.enabled) {
    if (input.u12.teams < 2) {
      errors.push('U12のチーム数は2以上を指定してください');
    }
    if (!input.u12.matchType) {
      errors.push('U12の試合方式を選択してください');
    }
  }
  
  // U15チェック
  if (input.u15.enabled) {
    if (input.u15.teams < 2) {
      errors.push('U15のチーム数は2以上を指定してください');
    }
    if (!input.u15.matchType) {
      errors.push('U15の試合方式を選択してください');
    }
  }
  
  return {
    isValid: errors.length === 0,
    errors: errors
  };
}

/**
 * 結果を出力
 */
function outputResults(ss, input, schedule, timeline, regulations) {
  // タイムラインシート
  let timelineSheet = ss.getSheetByName(CONFIG.SHEETS.TIMELINE);
  if (!timelineSheet) {
    timelineSheet = ss.insertSheet(CONFIG.SHEETS.TIMELINE);
  }
  outputTimelineToSheet(timeline, timelineSheet);
  
  // 試合スケジュールシート
  let scheduleSheet = ss.getSheetByName(CONFIG.SHEETS.SCHEDULE);
  if (!scheduleSheet) {
    scheduleSheet = ss.insertSheet(CONFIG.SHEETS.SCHEDULE);
  }
  outputScheduleToSheet(schedule, scheduleSheet);
  
  // トーナメント表（トーナメント方式の場合）
  if (input.u12.enabled && input.u12.matchType === CONFIG.MATCH_TYPES.TOURNAMENT) {
    let tournamentSheet = ss.getSheetByName('U12_' + CONFIG.SHEETS.TOURNAMENT);
    if (!tournamentSheet) {
      tournamentSheet = ss.insertSheet('U12_' + CONFIG.SHEETS.TOURNAMENT);
    }
    const u12Matches = generateMatchList(input.u12.teams, input.u12.matchType, 'U12', 'U12-');
    outputTournamentToSheet(u12Matches, tournamentSheet, 'U12');
  }
  
  if (input.u15.enabled && input.u15.matchType === CONFIG.MATCH_TYPES.TOURNAMENT) {
    let tournamentSheet = ss.getSheetByName('U15_' + CONFIG.SHEETS.TOURNAMENT);
    if (!tournamentSheet) {
      tournamentSheet = ss.insertSheet('U15_' + CONFIG.SHEETS.TOURNAMENT);
    }
    const u15Matches = generateMatchList(input.u15.teams, input.u15.matchType, 'U15', 'U15-');
    outputTournamentToSheet(u15Matches, tournamentSheet, 'U15');
  }
}

/**
 * エラーメッセージを表示
 */
function showError(message) {
  const ui = SpreadsheetApp.getUi();
  ui.alert('エラー', message, ui.ButtonSet.OK);
}

/**
 * 成功メッセージを表示
 */
function showSuccess(summary) {
  const ui = SpreadsheetApp.getUi();
  const message = `タイムライン生成が完了しました！

【サマリー】
・総試合数: ${summary.totalMatches}試合
  - U12: ${summary.u12Matches}試合 (${summary.u12Regulation})
  - U15: ${summary.u15Matches}試合 (${summary.u15Regulation})
・コート数: ${summary.courts}面
・開始: ${summary.startTime}
・終了予定: ${summary.endTime}
・終了時刻までの余裕: ${summary.marginMinutes}分`;

  ui.alert('完了', message, ui.ButtonSet.OK);
}

/**
 * 入力をクリア
 */
function clearInput() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName(CONFIG.SHEETS.INPUT);
  
  if (inputSheet) {
    // 入力セルのみクリア（ラベルは残す）
    const cellsToClear = [
      INPUT_CELLS.START_TIME, INPUT_CELLS.END_TIME, INPUT_CELLS.DAYS, INPUT_CELLS.COURTS,
      INPUT_CELLS.HAS_EXHIBITION, INPUT_CELLS.HAS_CEREMONY, INPUT_CELLS.HAS_PHOTO,
      INPUT_CELLS.U12_ENABLED, INPUT_CELLS.U12_TEAMS, INPUT_CELLS.U12_MATCH_TYPE,
      INPUT_CELLS.U15_ENABLED, INPUT_CELLS.U15_TEAMS, INPUT_CELLS.U15_MATCH_TYPE
    ];
    
    for (const cell of cellsToClear) {
      inputSheet.getRange(cell).clearContent();
    }
  }
  
  SpreadsheetApp.getUi().alert('入力をクリアしました');
}

/**
 * 出力をクリア
 */
function clearOutput() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const sheetsToClear = [
    CONFIG.SHEETS.TIMELINE,
    CONFIG.SHEETS.SCHEDULE,
    'U12_' + CONFIG.SHEETS.TOURNAMENT,
    'U15_' + CONFIG.SHEETS.TOURNAMENT
  ];
  
  for (const sheetName of sheetsToClear) {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      sheet.clear();
    }
  }
  
  SpreadsheetApp.getUi().alert('出力をクリアしました');
}
