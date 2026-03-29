/**
 * BSCA CUP タイムライン自動生成ツール - タイムライン生成
 */

/**
 * タイムラインを生成
 * @param {Object} input - 入力情報
 * @param {Array} schedule - 試合スケジュール
 * @returns {Array} - タイムラインイベントリスト
 */
function generateTimeline(input, schedule) {
  const timeline = [];
  let currentTime = new Date(input.common.startTime);
  
  // 1. 設営
  timeline.push(createTimelineEvent('設営', currentTime, CONFIG.TIMELINE_EVENTS.SETUP));
  currentTime = addMinutes(currentTime, CONFIG.TIMELINE_EVENTS.SETUP);
  
  // 2. 開場・受付
  timeline.push(createTimelineEvent('開場・受付', currentTime, CONFIG.TIMELINE_EVENTS.RECEPTION));
  currentTime = addMinutes(currentTime, CONFIG.TIMELINE_EVENTS.RECEPTION);
  
  // 3. 試合（スロットごとにグループ化）
  const matchesBySlot = groupMatchesBySlot(schedule);
  let matchNumber = 1;
  
  for (const [slotIndex, matches] of Object.entries(matchesBySlot)) {
    const slotMatches = matches.sort((a, b) => a.court.localeCompare(b.court));
    const slotStartTime = slotMatches[0].startTime;
    const maxSlotMinutes = Math.max(...slotMatches.map(m => m.slotMinutes));
    
    // 各コートの試合を1つのタイムラインエントリにまとめる
    const matchDescriptions = slotMatches.map(m => 
      `[${m.court}コート] ${m.category} ${m.team1} vs ${m.team2}`
    ).join('\n');
    
    timeline.push({
      eventType: 'match',
      name: `第${matchNumber}試合`,
      startTime: new Date(slotStartTime),
      endTime: addMinutes(slotStartTime, maxSlotMinutes),
      duration: maxSlotMinutes,
      description: matchDescriptions,
      matches: slotMatches
    });
    
    matchNumber++;
  }
  
  // 最後の試合終了時刻を取得
  if (schedule.length > 0) {
    const lastMatch = schedule[schedule.length - 1];
    currentTime = new Date(lastMatch.endTime);
  }
  
  // 4. エキシビジョン（任意）
  if (input.common.hasExhibition) {
    timeline.push(createTimelineEvent('エキシビジョン', currentTime, CONFIG.TIMELINE_EVENTS.EXHIBITION));
    currentTime = addMinutes(currentTime, CONFIG.TIMELINE_EVENTS.EXHIBITION);
  }
  
  // 5. 表彰式（任意）
  if (input.common.hasCeremony) {
    timeline.push(createTimelineEvent('表彰式', currentTime, CONFIG.TIMELINE_EVENTS.CEREMONY));
    currentTime = addMinutes(currentTime, CONFIG.TIMELINE_EVENTS.CEREMONY);
  }
  
  // 6. 写真撮影（任意）
  if (input.common.hasPhoto) {
    timeline.push(createTimelineEvent('1〜3位 写真撮影', currentTime, CONFIG.TIMELINE_EVENTS.PHOTO));
    currentTime = addMinutes(currentTime, CONFIG.TIMELINE_EVENTS.PHOTO);
  }
  
  // 7. 完全撤退
  timeline.push(createTimelineEvent('完全撤退', currentTime, CONFIG.TIMELINE_EVENTS.CLEANUP));
  
  return timeline;
}

/**
 * タイムラインイベントを作成
 */
function createTimelineEvent(name, startTime, duration) {
  return {
    eventType: 'general',
    name: name,
    startTime: new Date(startTime),
    endTime: addMinutes(startTime, duration),
    duration: duration,
    description: ''
  };
}

/**
 * 分を追加した新しいDateを返す
 */
function addMinutes(date, minutes) {
  const result = new Date(date);
  result.setMinutes(result.getMinutes() + minutes);
  return result;
}

/**
 * 試合をスロットごとにグループ化
 */
function groupMatchesBySlot(schedule) {
  const grouped = {};
  for (const match of schedule) {
    const slotIndex = match.slotIndex;
    if (!grouped[slotIndex]) {
      grouped[slotIndex] = [];
    }
    grouped[slotIndex].push(match);
  }
  return grouped;
}

/**
 * タイムラインをシートに出力
 */
function outputTimelineToSheet(timeline, sheet) {
  // ヘッダー行
  const headers = ['時刻', 'イベント', '所要時間', '備考'];
  
  // データ行を作成
  const data = [headers];
  
  for (const event of timeline) {
    const timeStr = formatTimeRange(event.startTime, event.endTime);
    const durationStr = `${event.duration}分`;
    
    data.push([
      timeStr,
      event.name,
      durationStr,
      event.description || ''
    ]);
  }
  
  // シートをクリアして書き込み
  sheet.clear();
  sheet.getRange(1, 1, data.length, headers.length).setValues(data);
  
  // 書式設定
  formatTimelineSheet(sheet, data.length);
}

/**
 * 試合スケジュールをシートに出力
 */
function outputScheduleToSheet(schedule, sheet) {
  // ヘッダー行
  const headers = ['試合No.', 'コート', '開始', '終了', 'カテゴリ', 'チーム1', 'チーム2', 'グループ/ラウンド'];
  
  // データ行を作成
  const data = [headers];
  
  for (const match of schedule) {
    data.push([
      match.matchNumber,
      match.court,
      formatTime(match.startTime),
      formatTime(match.endTime),
      match.category,
      match.team1,
      match.team2,
      match.groupName || match.roundName || ''
    ]);
  }
  
  // シートをクリアして書き込み
  sheet.clear();
  sheet.getRange(1, 1, data.length, headers.length).setValues(data);
  
  // 書式設定
  formatScheduleSheet(sheet, data.length);
}

/**
 * 時刻をフォーマット
 */
function formatTime(date) {
  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');
  return `${hours}:${minutes}`;
}

/**
 * 時間範囲をフォーマット
 */
function formatTimeRange(startTime, endTime) {
  return `${formatTime(startTime)} - ${formatTime(endTime)}`;
}

/**
 * タイムラインシートの書式設定
 */
function formatTimelineSheet(sheet, rowCount) {
  // ヘッダー行の書式
  const headerRange = sheet.getRange(1, 1, 1, 4);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('#ffffff');
  
  // 列幅の設定
  sheet.setColumnWidth(1, 150); // 時刻
  sheet.setColumnWidth(2, 120); // イベント
  sheet.setColumnWidth(3, 80);  // 所要時間
  sheet.setColumnWidth(4, 300); // 備考
  
  // 罫線
  if (rowCount > 1) {
    const dataRange = sheet.getRange(1, 1, rowCount, 4);
    dataRange.setBorder(true, true, true, true, true, true);
  }
  
  // 中央揃え
  sheet.getRange(1, 1, rowCount, 3).setHorizontalAlignment('center');
  
  // 備考列は左揃え・折り返し
  sheet.getRange(1, 4, rowCount, 1).setWrap(true);
}

/**
 * 試合スケジュールシートの書式設定
 */
function formatScheduleSheet(sheet, rowCount) {
  // ヘッダー行の書式
  const headerRange = sheet.getRange(1, 1, 1, 8);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#34a853');
  headerRange.setFontColor('#ffffff');
  
  // 列幅の設定
  sheet.setColumnWidth(1, 70);  // 試合No.
  sheet.setColumnWidth(2, 60);  // コート
  sheet.setColumnWidth(3, 60);  // 開始
  sheet.setColumnWidth(4, 60);  // 終了
  sheet.setColumnWidth(5, 80);  // カテゴリ
  sheet.setColumnWidth(6, 100); // チーム1
  sheet.setColumnWidth(7, 100); // チーム2
  sheet.setColumnWidth(8, 150); // グループ/ラウンド
  
  // 罫線
  if (rowCount > 1) {
    const dataRange = sheet.getRange(1, 1, rowCount, 8);
    dataRange.setBorder(true, true, true, true, true, true);
  }
  
  // 中央揃え
  sheet.getRange(1, 1, rowCount, 8).setHorizontalAlignment('center');
  
  // カテゴリ別に背景色を設定
  for (let i = 2; i <= rowCount; i++) {
    const category = sheet.getRange(i, 5).getValue();
    const rowRange = sheet.getRange(i, 1, 1, 8);
    if (category === 'U12') {
      rowRange.setBackground('#e8f5e9'); // 薄い緑
    } else if (category === 'U15') {
      rowRange.setBackground('#e3f2fd'); // 薄い青
    }
  }
}

/**
 * トーナメント表を生成してシートに出力
 */
function outputTournamentToSheet(matches, sheet, category) {
  const tournamentMatches = matches.filter(m => m.type === 'tournament' && m.category === category);
  
  if (tournamentMatches.length === 0) {
    return;
  }
  
  sheet.clear();
  
  // タイトル
  sheet.getRange(1, 1).setValue(`${category} トーナメント表`);
  sheet.getRange(1, 1).setFontWeight('bold').setFontSize(14);
  
  // トーナメントの構造を描画
  let row = 3;
  const roundGroups = {};
  
  for (const match of tournamentMatches) {
    if (!roundGroups[match.round]) {
      roundGroups[match.round] = [];
    }
    roundGroups[match.round].push(match);
  }
  
  let col = 1;
  for (const [round, roundMatches] of Object.entries(roundGroups).sort((a, b) => a[0] - b[0])) {
    // ラウンド名
    sheet.getRange(2, col).setValue(roundMatches[0].roundName);
    sheet.getRange(2, col).setFontWeight('bold');
    
    row = 3;
    for (const match of roundMatches) {
      sheet.getRange(row, col).setValue(match.team1);
      sheet.getRange(row + 1, col).setValue('vs');
      sheet.getRange(row + 2, col).setValue(match.team2);
      row += 4;
    }
    
    col += 2;
  }
  
  // 3位決定戦
  const thirdPlace = tournamentMatches.find(m => m.isThirdPlace);
  if (thirdPlace) {
    sheet.getRange(row + 1, 1).setValue('【3位決定戦】');
    sheet.getRange(row + 2, 1).setValue(`${thirdPlace.team1} vs ${thirdPlace.team2}`);
  }
}

/**
 * サマリー情報を生成
 */
function generateSummary(input, schedule, timeline, regulations) {
  const u12MatchCount = schedule.filter(m => m.category === 'U12').length;
  const u15MatchCount = schedule.filter(m => m.category === 'U15').length;
  
  const lastEvent = timeline[timeline.length - 1];
  const endTime = lastEvent.endTime;
  
  const margin = input.common.endTime - endTime;
  const marginMinutes = Math.floor(margin / (1000 * 60));
  
  return {
    totalMatches: schedule.length,
    u12Matches: u12MatchCount,
    u15Matches: u15MatchCount,
    u12Regulation: regulations.u12Regulation ? regulations.u12Regulation.label : 'なし',
    u15Regulation: regulations.u15Regulation ? regulations.u15Regulation.label : 'なし',
    startTime: formatTime(input.common.startTime),
    endTime: formatTime(endTime),
    marginMinutes: marginMinutes,
    courts: input.common.courts
  };
}
