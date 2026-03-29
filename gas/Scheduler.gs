/**
 * BSCA CUP タイムライン自動生成ツール - スケジュール最適化ロジック
 */

/**
 * 最適なレギュレーションを選択
 * @param {Object} input - 入力情報
 * @returns {Object} - 選択されたレギュレーションと評価結果
 */
function selectOptimalRegulations(input) {
  const availableTime = calculateAvailableTime(input);
  const u12Matches = input.u12.enabled ? calculateMatchCount(input.u12.teams, input.u12.matchType) : null;
  const u15Matches = input.u15.enabled ? calculateMatchCount(input.u15.teams, input.u15.matchType) : null;
  
  const u12Regs = input.u12.enabled ? getRegulationsByCategory('U12') : [];
  const u15Regs = input.u15.enabled ? getRegulationsByCategory('U15') : [];
  
  let bestResult = null;
  let bestMargin = -Infinity;
  
  // 全組み合わせを評価
  const u12Options = u12Regs.length > 0 ? u12Regs : [null];
  const u15Options = u15Regs.length > 0 ? u15Regs : [null];
  
  for (const u12Reg of u12Options) {
    for (const u15Reg of u15Options) {
      const evaluation = evaluateRegulationCombination(
        input, availableTime, u12Matches, u15Matches, u12Reg, u15Reg
      );
      
      if (evaluation.isValid && evaluation.margin > bestMargin) {
        bestMargin = evaluation.margin;
        bestResult = {
          u12Regulation: u12Reg,
          u15Regulation: u15Reg,
          ...evaluation
        };
      }
    }
  }
  
  return bestResult;
}

/**
 * 利用可能時間を計算（分）
 */
function calculateAvailableTime(input) {
  const start = input.common.startTime;
  const end = input.common.endTime;
  
  // 時刻から分に変換して差分を計算
  const startMinutes = start.getHours() * 60 + start.getMinutes();
  const endMinutes = end.getHours() * 60 + end.getMinutes();
  
  let totalMinutes = endMinutes - startMinutes;
  
  // 固定イベントの時間を差し引く
  totalMinutes -= CONFIG.TIMELINE_EVENTS.SETUP;
  totalMinutes -= CONFIG.TIMELINE_EVENTS.RECEPTION;
  totalMinutes -= CONFIG.TIMELINE_EVENTS.CLEANUP;
  
  if (input.common.hasExhibition) {
    totalMinutes -= CONFIG.TIMELINE_EVENTS.EXHIBITION;
  }
  if (input.common.hasCeremony) {
    totalMinutes -= CONFIG.TIMELINE_EVENTS.CEREMONY;
  }
  if (input.common.hasPhoto) {
    totalMinutes -= CONFIG.TIMELINE_EVENTS.PHOTO;
  }
  
  return totalMinutes;
}

/**
 * レギュレーション組み合わせを評価
 */
function evaluateRegulationCombination(input, availableTime, u12Matches, u15Matches, u12Reg, u15Reg) {
  let totalSlots = 0;
  let totalTime = 0;
  
  if (u12Matches && u12Reg) {
    const u12Slots = Math.ceil(u12Matches.total / input.common.courts);
    totalSlots += u12Slots;
    totalTime += u12Slots * u12Reg.slot_minutes;
  }
  
  if (u15Matches && u15Reg) {
    const u15Slots = Math.ceil(u15Matches.total / input.common.courts);
    totalSlots += u15Slots;
    totalTime += u15Slots * u15Reg.slot_minutes;
  }
  
  const margin = availableTime - totalTime;
  const isValid = margin >= 0;
  
  return {
    isValid: isValid,
    totalSlots: totalSlots,
    totalTime: totalTime,
    availableTime: availableTime,
    margin: margin
  };
}

/**
 * 試合スケジュールを生成
 * @param {Object} input - 入力情報
 * @param {Object} regulations - 選択されたレギュレーション
 * @returns {Array} - スケジュールされた試合リスト
 */
function generateSchedule(input, regulations) {
  const u12Matches = input.u12.enabled ? 
    generateMatchList(input.u12.teams, input.u12.matchType, 'U12', 'U12-') : [];
  const u15Matches = input.u15.enabled ? 
    generateMatchList(input.u15.teams, input.u15.matchType, 'U15', 'U15-') : [];
  
  // 連戦NGを考慮した配置
  const schedule = arrangeMatchesWithConstraints(
    u12Matches, u15Matches, input, regulations
  );
  
  return schedule;
}

/**
 * 連戦制約を考慮した試合配置
 */
function arrangeMatchesWithConstraints(u12Matches, u15Matches, input, regulations) {
  const courts = input.common.courts;
  const schedule = [];
  
  // 試合をスロットに配置
  const slots = [];
  let slotIndex = 0;
  
  // U12とU15を交互に配置（U12優先）
  const u12Queue = [...u12Matches];
  const u15Queue = [...u15Matches];
  
  // 各チームの最後に試合したスロット番号を記録
  const lastMatchSlot = {};
  
  while (u12Queue.length > 0 || u15Queue.length > 0) {
    const slot = {
      index: slotIndex,
      matches: []
    };
    
    // このスロットに配置する試合を選択（コート数分）
    for (let court = 0; court < courts; court++) {
      let selectedMatch = null;
      let selectedQueue = null;
      
      // U12優先で配置
      selectedMatch = findValidMatch(u12Queue, lastMatchSlot, slotIndex);
      if (selectedMatch) {
        selectedQueue = u12Queue;
      } else {
        selectedMatch = findValidMatch(u15Queue, lastMatchSlot, slotIndex);
        if (selectedMatch) {
          selectedQueue = u15Queue;
        }
      }
      
      // 見つからなければU15からも試す
      if (!selectedMatch && u15Queue.length > 0) {
        selectedMatch = findValidMatch(u15Queue, lastMatchSlot, slotIndex);
        if (selectedMatch) {
          selectedQueue = u15Queue;
        }
      }
      
      // それでも見つからなければ、U12から強制的に選択
      if (!selectedMatch && u12Queue.length > 0) {
        selectedMatch = u12Queue[0];
        selectedQueue = u12Queue;
      }
      
      if (!selectedMatch && u15Queue.length > 0) {
        selectedMatch = u15Queue[0];
        selectedQueue = u15Queue;
      }
      
      if (selectedMatch) {
        // キューから削除
        const idx = selectedQueue.indexOf(selectedMatch);
        if (idx > -1) {
          selectedQueue.splice(idx, 1);
        }
        
        // 最後の試合スロットを更新
        lastMatchSlot[selectedMatch.team1] = slotIndex;
        lastMatchSlot[selectedMatch.team2] = slotIndex;
        
        slot.matches.push({
          ...selectedMatch,
          court: COURT_NAMES[court],
          slotIndex: slotIndex
        });
      }
    }
    
    if (slot.matches.length > 0) {
      slots.push(slot);
      slotIndex++;
    } else {
      // 配置できる試合がない場合は終了
      break;
    }
  }
  
  // 時刻を割り当て
  return assignTimes(slots, input, regulations);
}

/**
 * 連戦制約を満たす試合を探す
 */
function findValidMatch(queue, lastMatchSlot, currentSlot) {
  for (const match of queue) {
    const team1LastSlot = lastMatchSlot[match.team1];
    const team2LastSlot = lastMatchSlot[match.team2];
    
    // 最低1枠空ける（連戦NG）
    const team1OK = team1LastSlot === undefined || currentSlot - team1LastSlot > 1;
    const team2OK = team2LastSlot === undefined || currentSlot - team2LastSlot > 1;
    
    if (team1OK && team2OK) {
      return match;
    }
  }
  return null;
}

/**
 * 時刻を割り当て
 */
function assignTimes(slots, input, regulations) {
  const schedule = [];
  
  // 試合開始時刻を計算
  const startTime = new Date(input.common.startTime);
  startTime.setMinutes(startTime.getMinutes() + CONFIG.TIMELINE_EVENTS.SETUP + CONFIG.TIMELINE_EVENTS.RECEPTION);
  
  let currentTime = new Date(startTime);
  
  for (const slot of slots) {
    for (const match of slot.matches) {
      const slotMinutes = match.category === 'U12' ? 
        regulations.u12Regulation.slot_minutes : 
        regulations.u15Regulation.slot_minutes;
      
      const endTime = new Date(currentTime);
      endTime.setMinutes(endTime.getMinutes() + slotMinutes);
      
      schedule.push({
        ...match,
        startTime: new Date(currentTime),
        endTime: endTime,
        slotMinutes: slotMinutes
      });
    }
    
    // 次のスロットの開始時刻を計算（最大のslot_minutesを使用）
    const maxSlotMinutes = Math.max(...slot.matches.map(m => 
      m.category === 'U12' ? 
        regulations.u12Regulation.slot_minutes : 
        regulations.u15Regulation.slot_minutes
    ));
    currentTime.setMinutes(currentTime.getMinutes() + maxSlotMinutes);
  }
  
  return schedule;
}

/**
 * スケジュールの検証
 */
function validateSchedule(schedule, input) {
  const endTime = input.common.endTime;
  const errors = [];
  const warnings = [];
  
  if (schedule.length === 0) {
    errors.push('配置可能な試合がありません');
    return { isValid: false, errors, warnings };
  }
  
  // 最終試合の終了時刻をチェック
  const lastMatch = schedule[schedule.length - 1];
  const lastMatchEnd = lastMatch.endTime;
  
  // 終了後のイベント時間を追加
  const finalEnd = new Date(lastMatchEnd);
  if (input.common.hasExhibition) {
    finalEnd.setMinutes(finalEnd.getMinutes() + CONFIG.TIMELINE_EVENTS.EXHIBITION);
  }
  if (input.common.hasCeremony) {
    finalEnd.setMinutes(finalEnd.getMinutes() + CONFIG.TIMELINE_EVENTS.CEREMONY);
  }
  if (input.common.hasPhoto) {
    finalEnd.setMinutes(finalEnd.getMinutes() + CONFIG.TIMELINE_EVENTS.PHOTO);
  }
  finalEnd.setMinutes(finalEnd.getMinutes() + CONFIG.TIMELINE_EVENTS.CLEANUP);
  
  if (finalEnd > endTime) {
    const overMinutes = Math.ceil((finalEnd - endTime) / (1000 * 60));
    errors.push(`利用終了時刻を${overMinutes}分超過します`);
  }
  
  // 連戦チェック
  const teamSchedule = {};
  for (const match of schedule) {
    for (const team of [match.team1, match.team2]) {
      if (!teamSchedule[team]) {
        teamSchedule[team] = [];
      }
      teamSchedule[team].push(match.slotIndex);
    }
  }
  
  for (const [team, slots] of Object.entries(teamSchedule)) {
    slots.sort((a, b) => a - b);
    for (let i = 1; i < slots.length; i++) {
      if (slots[i] - slots[i-1] <= 1) {
        warnings.push(`${team}が連戦になっています（スロット${slots[i-1]+1}→${slots[i]+1}）`);
      }
    }
  }
  
  return {
    isValid: errors.length === 0,
    errors: errors,
    warnings: warnings,
    finalEndTime: finalEnd
  };
}

/**
 * 時間不足時のエラー情報を生成
 */
function generateTimeError(input, regulations) {
  const availableTime = calculateAvailableTime(input);
  const u12Matches = input.u12.enabled ? calculateMatchCount(input.u12.teams, input.u12.matchType) : null;
  const u15Matches = input.u15.enabled ? calculateMatchCount(input.u15.teams, input.u15.matchType) : null;
  
  // 最小のレギュレーションで計算
  const u12MinReg = input.u12.enabled ? getRegulationsByCategory('U12')[0] : null;
  const u15MinReg = input.u15.enabled ? getRegulationsByCategory('U15')[0] : null;
  
  let requiredTime = 0;
  if (u12Matches && u12MinReg) {
    const slots = Math.ceil(u12Matches.total / input.common.courts);
    requiredTime += slots * u12MinReg.slot_minutes;
  }
  if (u15Matches && u15MinReg) {
    const slots = Math.ceil(u15Matches.total / input.common.courts);
    requiredTime += slots * u15MinReg.slot_minutes;
  }
  
  const shortage = requiredTime - availableTime;
  const shortageSlots = Math.ceil(shortage / 60);
  
  return {
    availableTime: availableTime,
    requiredTime: requiredTime,
    shortageMinutes: shortage,
    shortageSlots: shortageSlots,
    message: `時間が不足しています。不足: ${shortage}分（約${shortageSlots}枠）`
  };
}
