/**
 * ============================================================
 * BSCA CUP タイムライン自動生成ツール - 全コード統合版
 * ============================================================
 *
 * 使い方：
 * 1. Apps Scriptで既存のファイルをすべて削除
 * 2. このファイルの内容をコピー＆ペースト
 * 3. 保存してsetupSpreadsheet() を実行
 *
 * ============================================================
 */

// ============================================================
// Config.gs - 設定
// ============================================================

const CONFIG = {
  SHEETS: {
    INPUT: "入力",
    TIMELINE: "タイムライン",
    SCHEDULE: "試合スケジュール",
    TOURNAMENT: "トーナメント表",
  },

  // デフォルト値（入力がない場合に使用）
  DEFAULT_TIMES: {
    SETUP: 30,
    RECEPTION: 15,
    EXHIBITION: 30,
    CEREMONY: 15,
    PHOTO: 10,
    CLEANUP: 30,
  },

  CATEGORIES: {
    U12: "U12",
    U15: "U15",
  },

  MATCH_TYPES: {
    LEAGUE: "リーグ戦",
    TOURNAMENT: "トーナメント",
  },

  GROUP_THRESHOLDS: {
    TWO_GROUPS: 6,
    THREE_GROUPS: 9,
  },
};

// スロット時間の選択肢（45分〜90分、5分刻み）
const SLOT_OPTIONS = [45, 50, 55, 60, 65, 70, 75, 80, 85, 90];

// REGULATIONSを動的に生成
const REGULATIONS = {};
for (const minutes of SLOT_OPTIONS) {
  REGULATIONS[`SLOT${minutes}`] = {
    slot_minutes: minutes,
    label: `${minutes}分`,
  };
}

function getRegulationsByCategory(category) {
  // カテゴリに関係なく全てのスロットオプションを返す
  const result = [];
  for (const [id, reg] of Object.entries(REGULATIONS)) {
    result.push({ id: id, ...reg, category: category });
  }
  return result.sort((a, b) => a.slot_minutes - b.slot_minutes);
}

function getSlotMinutes(regulationId) {
  const reg = REGULATIONS[regulationId];
  return reg ? reg.slot_minutes : null;
}

function createInputStructure() {
  return {
    common: {
      startTime: null,
      endTime: null,
      courts: 1,
      // イベント時間設定
      setupMinutes: 30,
      receptionMinutes: 15,
      firstInterval15: false, // 1〜2試合間を15分にするか
      intervalMinutes: 10, // その他試合間インターバル（デフォルト10分）
      hasExhibition: false,
      exhibitionMinutes: 30,
      hasCeremony: true,
      ceremonyMinutes: 15,
      hasPhoto: true,
      photoMinutes: 10,
      cleanupMinutes: 30,
    },
    u12: {
      enabled: false,
      teams: 0,
      matchType: null,
      regulation: "自動", // 自動 / 60分 / 65分 / 70分
    },
    u15: {
      enabled: false,
      teams: 0,
      matchType: null,
      regulation: "自動", // 自動 / 75分 / 80分 / 85分
    },
  };
}

const INPUT_CELLS = {
  // 大会共通
  START_TIME: "C3",
  END_TIME: "C4",
  COURTS: "C5",
  // イベント時間（0分ならスキップ）
  SETUP_MINUTES: "C8",
  RECEPTION_MINUTES: "C9",
  FIRST_INTERVAL_15: "C10", // 1〜2試合間15分（ON/OFF）
  INTERVAL_MINUTES: "C11", // その他試合間インターバル（分）
  EXHIBITION_MINUTES: "C12",
  CEREMONY_MINUTES: "C13",
  PHOTO_MINUTES: "C14",
  CLEANUP_MINUTES: "C15",
  // U12
  U12_ENABLED: "C18",
  U12_TEAMS: "C19",
  U12_MATCH_TYPE: "C20",
  U12_REGULATION: "C21",
  // U15
  U15_ENABLED: "C24",
  U15_TEAMS: "C25",
  U15_MATCH_TYPE: "C26",
  U15_REGULATION: "C27",
};

// 固定値
const FIRST_INTERVAL_MINUTES = 15; // 1〜2試合間インターバル（ONの場合）

const TEAM_NAMES = [
  "A",
  "B",
  "C",
  "D",
  "E",
  "F",
  "G",
  "H",
  "I",
  "J",
  "K",
  "L",
  "M",
  "N",
  "O",
  "P",
];
const COURT_NAMES = ["A", "B"];

// ============================================================
// MatchCalculator.gs - 試合数計算
// ============================================================

function combination(n, r) {
  if (r > n || r < 0) return 0;
  if (r === 0 || r === n) return 1;

  let result = 1;
  for (let i = 0; i < r; i++) {
    result = (result * (n - i)) / (i + 1);
  }
  return Math.round(result);
}

function calculateMatchCount(teams, matchType) {
  if (teams < 2) {
    return {
      total: 0,
      groups: 0,
      groupMatches: 0,
      finalMatches: 0,
      structure: null,
    };
  }

  if (matchType === CONFIG.MATCH_TYPES.TOURNAMENT) {
    return calculateTournamentMatches(teams);
  } else {
    return calculateLeagueMatches(teams);
  }
}

function calculateTournamentMatches(teams) {
  // 通常のトーナメント試合数: n - 1
  // 3位決定戦: 4チーム以上で +1
  const baseMatches = teams - 1;
  const thirdPlaceMatch = teams >= 4 ? 1 : 0;

  return {
    total: baseMatches + thirdPlaceMatch,
    groups: 0,
    groupMatches: 0,
    finalMatches: baseMatches + thirdPlaceMatch,
    structure: "tournament",
  };
}

function calculateLeagueMatches(teams) {
  if (teams <= 5) {
    const matches = combination(teams, 2);
    return {
      total: matches,
      groups: 1,
      groupMatches: matches,
      finalMatches: 0,
      structure: "single_league",
      teamsPerGroup: [teams],
    };
  } else if (teams <= 8) {
    return calculateTwoGroupLeague(teams);
  } else {
    return calculateThreeGroupLeague(teams);
  }
}

/**
 * 2グループ制リーグ戦（グループ総当たりのみ、順位決定戦なし）
 */
function calculateTwoGroupLeague(teams) {
  const groupA = Math.ceil(teams / 2);
  const groupB = Math.floor(teams / 2);

  const groupAMatches = combination(groupA, 2);
  const groupBMatches = combination(groupB, 2);
  const totalGroupMatches = groupAMatches + groupBMatches;

  return {
    total: totalGroupMatches,
    groups: 2,
    groupMatches: totalGroupMatches,
    finalMatches: 0,
    structure: "two_group_league",
    teamsPerGroup: [groupA, groupB],
    groupMatchBreakdown: [groupAMatches, groupBMatches],
  };
}

/**
 * 3グループ制リーグ戦（グループ総当たりのみ、順位決定戦なし）
 */
function calculateThreeGroupLeague(teams) {
  const base = Math.floor(teams / 3);
  const remainder = teams % 3;

  const groupSizes = [base, base, base];
  for (let i = 0; i < remainder; i++) {
    groupSizes[i]++;
  }

  const groupMatches = groupSizes.map((size) => combination(size, 2));
  const totalGroupMatches = groupMatches.reduce((a, b) => a + b, 0);

  return {
    total: totalGroupMatches,
    groups: 3,
    groupMatches: totalGroupMatches,
    finalMatches: 0,
    structure: "three_group_league",
    teamsPerGroup: groupSizes,
    groupMatchBreakdown: groupMatches,
  };
}

function generateMatchList(teams, matchType, category, prefix) {
  const info = calculateMatchCount(teams, matchType);

  if (matchType === CONFIG.MATCH_TYPES.TOURNAMENT) {
    return generateTournamentMatchList(teams, category, prefix);
  } else {
    return generateLeagueMatchList(teams, category, prefix, info);
  }
}

function generateTournamentMatchList(teams, category, prefix) {
  const matches = [];
  let matchNum = 1;

  // チーム数に応じたトーナメント構造を生成
  if (teams === 2) {
    // 2チーム: 決勝のみ
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 1,
      roundName: "決勝",
      team1: `${prefix}A`,
      team2: `${prefix}B`,
      type: "tournament",
      isFinal: true,
    });
  } else if (teams === 3) {
    // 3チーム: 準決勝1試合 + 決勝
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 1,
      roundName: "準決勝",
      team1: `${prefix}B`,
      team2: `${prefix}C`,
      type: "tournament",
      isFinal: false,
    });
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 2,
      roundName: "決勝",
      team1: `${prefix}A`,
      team2: `${prefix}準決勝勝者`,
      type: "tournament",
      isFinal: true,
    });
  } else if (teams === 4) {
    // 4チーム: 準決勝2試合 + 3位決定戦 + 決勝
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 1,
      roundName: "準決勝",
      team1: `${prefix}A`,
      team2: `${prefix}D`,
      type: "tournament",
      isFinal: false,
    });
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 1,
      roundName: "準決勝",
      team1: `${prefix}B`,
      team2: `${prefix}C`,
      type: "tournament",
      isFinal: false,
    });
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 2,
      roundName: "3位決定戦",
      team1: `${prefix}準決勝敗者1`,
      team2: `${prefix}準決勝敗者2`,
      type: "tournament",
      isFinal: false,
      isThirdPlace: true,
    });
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 2,
      roundName: "決勝",
      team1: `${prefix}準決勝勝者1`,
      team2: `${prefix}準決勝勝者2`,
      type: "tournament",
      isFinal: true,
    });
  } else if (teams === 5) {
    // 5チーム: シードA,B,C + 1回戦D vs E
    // 1回戦
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 1,
      roundName: "1回戦",
      team1: `${prefix}D`,
      team2: `${prefix}E`,
      type: "tournament",
      isFinal: false,
    });
    // 準決勝: A vs 1回戦勝者, B vs C
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 2,
      roundName: "準決勝",
      team1: `${prefix}A`,
      team2: `${prefix}1回戦勝者`,
      type: "tournament",
      isFinal: false,
    });
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 2,
      roundName: "準決勝",
      team1: `${prefix}B`,
      team2: `${prefix}C`,
      type: "tournament",
      isFinal: false,
    });
    // 3位決定戦 + 決勝
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 3,
      roundName: "3位決定戦",
      team1: `${prefix}準決勝敗者1`,
      team2: `${prefix}準決勝敗者2`,
      type: "tournament",
      isFinal: false,
      isThirdPlace: true,
    });
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 3,
      roundName: "決勝",
      team1: `${prefix}準決勝勝者1`,
      team2: `${prefix}準決勝勝者2`,
      type: "tournament",
      isFinal: true,
    });
  } else if (teams === 6) {
    // 6チーム: シードA,B + 1回戦C vs F, D vs E
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 1,
      roundName: "1回戦",
      team1: `${prefix}C`,
      team2: `${prefix}F`,
      type: "tournament",
      isFinal: false,
    });
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 1,
      roundName: "1回戦",
      team1: `${prefix}D`,
      team2: `${prefix}E`,
      type: "tournament",
      isFinal: false,
    });
    // 準決勝: A vs 1回戦勝者1, B vs 1回戦勝者2
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 2,
      roundName: "準決勝",
      team1: `${prefix}A`,
      team2: `${prefix}1回戦勝者1`,
      type: "tournament",
      isFinal: false,
    });
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 2,
      roundName: "準決勝",
      team1: `${prefix}B`,
      team2: `${prefix}1回戦勝者2`,
      type: "tournament",
      isFinal: false,
    });
    // 3位決定戦 + 決勝
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 3,
      roundName: "3位決定戦",
      team1: `${prefix}準決勝敗者1`,
      team2: `${prefix}準決勝敗者2`,
      type: "tournament",
      isFinal: false,
      isThirdPlace: true,
    });
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 3,
      roundName: "決勝",
      team1: `${prefix}準決勝勝者1`,
      team2: `${prefix}準決勝勝者2`,
      type: "tournament",
      isFinal: true,
    });
  } else if (teams === 7) {
    // 7チーム: シードA + 1回戦B vs G, C vs F, D vs E
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 1,
      roundName: "1回戦",
      team1: `${prefix}B`,
      team2: `${prefix}G`,
      type: "tournament",
      isFinal: false,
    });
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 1,
      roundName: "1回戦",
      team1: `${prefix}C`,
      team2: `${prefix}F`,
      type: "tournament",
      isFinal: false,
    });
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 1,
      roundName: "1回戦",
      team1: `${prefix}D`,
      team2: `${prefix}E`,
      type: "tournament",
      isFinal: false,
    });
    // 準決勝: A vs 1回戦勝者1, 1回戦勝者2 vs 1回戦勝者3
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 2,
      roundName: "準決勝",
      team1: `${prefix}A`,
      team2: `${prefix}1回戦勝者1`,
      type: "tournament",
      isFinal: false,
    });
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 2,
      roundName: "準決勝",
      team1: `${prefix}1回戦勝者2`,
      team2: `${prefix}1回戦勝者3`,
      type: "tournament",
      isFinal: false,
    });
    // 3位決定戦 + 決勝
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 3,
      roundName: "3位決定戦",
      team1: `${prefix}準決勝敗者1`,
      team2: `${prefix}準決勝敗者2`,
      type: "tournament",
      isFinal: false,
      isThirdPlace: true,
    });
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 3,
      roundName: "決勝",
      team1: `${prefix}準決勝勝者1`,
      team2: `${prefix}準決勝勝者2`,
      type: "tournament",
      isFinal: true,
    });
  } else if (teams === 8) {
    // 8チーム: シードなし、1回戦4試合
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 1,
      roundName: "1回戦",
      team1: `${prefix}A`,
      team2: `${prefix}H`,
      type: "tournament",
      isFinal: false,
    });
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 1,
      roundName: "1回戦",
      team1: `${prefix}B`,
      team2: `${prefix}G`,
      type: "tournament",
      isFinal: false,
    });
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 1,
      roundName: "1回戦",
      team1: `${prefix}C`,
      team2: `${prefix}F`,
      type: "tournament",
      isFinal: false,
    });
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 1,
      roundName: "1回戦",
      team1: `${prefix}D`,
      team2: `${prefix}E`,
      type: "tournament",
      isFinal: false,
    });
    // 準決勝
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 2,
      roundName: "準決勝",
      team1: `${prefix}1回戦勝者1`,
      team2: `${prefix}1回戦勝者2`,
      type: "tournament",
      isFinal: false,
    });
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 2,
      roundName: "準決勝",
      team1: `${prefix}1回戦勝者3`,
      team2: `${prefix}1回戦勝者4`,
      type: "tournament",
      isFinal: false,
    });
    // 3位決定戦 + 決勝
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 3,
      roundName: "3位決定戦",
      team1: `${prefix}準決勝敗者1`,
      team2: `${prefix}準決勝敗者2`,
      type: "tournament",
      isFinal: false,
      isThirdPlace: true,
    });
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 3,
      roundName: "決勝",
      team1: `${prefix}準決勝勝者1`,
      team2: `${prefix}準決勝勝者2`,
      type: "tournament",
      isFinal: true,
    });
  } else if (teams === 9) {
    // 9チーム: シード7 + 1回戦1試合
    // 1回戦: H vs I
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 1,
      roundName: "1回戦",
      team1: `${prefix}H`,
      team2: `${prefix}I`,
      type: "tournament",
      isFinal: false,
    });
    // 準々決勝: A vs 1回戦勝者, B vs G, C vs F, D vs E
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 2,
      roundName: "準々決勝",
      team1: `${prefix}A`,
      team2: `${prefix}1回戦勝者`,
      type: "tournament",
      isFinal: false,
    });
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 2,
      roundName: "準々決勝",
      team1: `${prefix}B`,
      team2: `${prefix}G`,
      type: "tournament",
      isFinal: false,
    });
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 2,
      roundName: "準々決勝",
      team1: `${prefix}C`,
      team2: `${prefix}F`,
      type: "tournament",
      isFinal: false,
    });
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 2,
      roundName: "準々決勝",
      team1: `${prefix}D`,
      team2: `${prefix}E`,
      type: "tournament",
      isFinal: false,
    });
    // 準決勝
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 3,
      roundName: "準決勝",
      team1: `${prefix}準々決勝勝者1`,
      team2: `${prefix}準々決勝勝者2`,
      type: "tournament",
      isFinal: false,
    });
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 3,
      roundName: "準決勝",
      team1: `${prefix}準々決勝勝者3`,
      team2: `${prefix}準々決勝勝者4`,
      type: "tournament",
      isFinal: false,
    });
    // 3位決定戦 + 決勝
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 4,
      roundName: "3位決定戦",
      team1: `${prefix}準決勝敗者1`,
      team2: `${prefix}準決勝敗者2`,
      type: "tournament",
      isFinal: false,
      isThirdPlace: true,
    });
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 4,
      roundName: "決勝",
      team1: `${prefix}準決勝勝者1`,
      team2: `${prefix}準決勝勝者2`,
      type: "tournament",
      isFinal: true,
    });
  } else if (teams === 10) {
    // 10チーム: シード6 + 1回戦2試合
    // 1回戦: G vs J, H vs I
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 1,
      roundName: "1回戦",
      team1: `${prefix}G`,
      team2: `${prefix}J`,
      type: "tournament",
      isFinal: false,
    });
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 1,
      roundName: "1回戦",
      team1: `${prefix}H`,
      team2: `${prefix}I`,
      type: "tournament",
      isFinal: false,
    });
    // 準々決勝
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 2,
      roundName: "準々決勝",
      team1: `${prefix}A`,
      team2: `${prefix}1回戦勝者1`,
      type: "tournament",
      isFinal: false,
    });
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 2,
      roundName: "準々決勝",
      team1: `${prefix}B`,
      team2: `${prefix}1回戦勝者2`,
      type: "tournament",
      isFinal: false,
    });
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 2,
      roundName: "準々決勝",
      team1: `${prefix}C`,
      team2: `${prefix}F`,
      type: "tournament",
      isFinal: false,
    });
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 2,
      roundName: "準々決勝",
      team1: `${prefix}D`,
      team2: `${prefix}E`,
      type: "tournament",
      isFinal: false,
    });
    // 準決勝
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 3,
      roundName: "準決勝",
      team1: `${prefix}準々決勝勝者1`,
      team2: `${prefix}準々決勝勝者2`,
      type: "tournament",
      isFinal: false,
    });
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 3,
      roundName: "準決勝",
      team1: `${prefix}準々決勝勝者3`,
      team2: `${prefix}準々決勝勝者4`,
      type: "tournament",
      isFinal: false,
    });
    // 3位決定戦 + 決勝
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 4,
      roundName: "3位決定戦",
      team1: `${prefix}準決勝敗者1`,
      team2: `${prefix}準決勝敗者2`,
      type: "tournament",
      isFinal: false,
      isThirdPlace: true,
    });
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 4,
      roundName: "決勝",
      team1: `${prefix}準決勝勝者1`,
      team2: `${prefix}準決勝勝者2`,
      type: "tournament",
      isFinal: true,
    });
  } else if (teams === 11) {
    // 11チーム: シード5 + 1回戦3試合
    // 1回戦: F vs K, G vs J, H vs I
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 1,
      roundName: "1回戦",
      team1: `${prefix}F`,
      team2: `${prefix}K`,
      type: "tournament",
      isFinal: false,
    });
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 1,
      roundName: "1回戦",
      team1: `${prefix}G`,
      team2: `${prefix}J`,
      type: "tournament",
      isFinal: false,
    });
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 1,
      roundName: "1回戦",
      team1: `${prefix}H`,
      team2: `${prefix}I`,
      type: "tournament",
      isFinal: false,
    });
    // 準々決勝
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 2,
      roundName: "準々決勝",
      team1: `${prefix}A`,
      team2: `${prefix}1回戦勝者1`,
      type: "tournament",
      isFinal: false,
    });
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 2,
      roundName: "準々決勝",
      team1: `${prefix}B`,
      team2: `${prefix}1回戦勝者2`,
      type: "tournament",
      isFinal: false,
    });
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 2,
      roundName: "準々決勝",
      team1: `${prefix}C`,
      team2: `${prefix}1回戦勝者3`,
      type: "tournament",
      isFinal: false,
    });
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 2,
      roundName: "準々決勝",
      team1: `${prefix}D`,
      team2: `${prefix}E`,
      type: "tournament",
      isFinal: false,
    });
    // 準決勝
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 3,
      roundName: "準決勝",
      team1: `${prefix}準々決勝勝者1`,
      team2: `${prefix}準々決勝勝者2`,
      type: "tournament",
      isFinal: false,
    });
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 3,
      roundName: "準決勝",
      team1: `${prefix}準々決勝勝者3`,
      team2: `${prefix}準々決勝勝者4`,
      type: "tournament",
      isFinal: false,
    });
    // 3位決定戦 + 決勝
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 4,
      roundName: "3位決定戦",
      team1: `${prefix}準決勝敗者1`,
      team2: `${prefix}準決勝敗者2`,
      type: "tournament",
      isFinal: false,
      isThirdPlace: true,
    });
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 4,
      roundName: "決勝",
      team1: `${prefix}準決勝勝者1`,
      team2: `${prefix}準決勝勝者2`,
      type: "tournament",
      isFinal: true,
    });
  } else if (teams === 12) {
    // 12チーム: シード4 + 1回戦4試合
    // 1回戦: E vs L, F vs K, G vs J, H vs I
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 1,
      roundName: "1回戦",
      team1: `${prefix}E`,
      team2: `${prefix}L`,
      type: "tournament",
      isFinal: false,
    });
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 1,
      roundName: "1回戦",
      team1: `${prefix}F`,
      team2: `${prefix}K`,
      type: "tournament",
      isFinal: false,
    });
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 1,
      roundName: "1回戦",
      team1: `${prefix}G`,
      team2: `${prefix}J`,
      type: "tournament",
      isFinal: false,
    });
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 1,
      roundName: "1回戦",
      team1: `${prefix}H`,
      team2: `${prefix}I`,
      type: "tournament",
      isFinal: false,
    });
    // 準々決勝
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 2,
      roundName: "準々決勝",
      team1: `${prefix}A`,
      team2: `${prefix}1回戦勝者1`,
      type: "tournament",
      isFinal: false,
    });
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 2,
      roundName: "準々決勝",
      team1: `${prefix}B`,
      team2: `${prefix}1回戦勝者2`,
      type: "tournament",
      isFinal: false,
    });
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 2,
      roundName: "準々決勝",
      team1: `${prefix}C`,
      team2: `${prefix}1回戦勝者3`,
      type: "tournament",
      isFinal: false,
    });
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 2,
      roundName: "準々決勝",
      team1: `${prefix}D`,
      team2: `${prefix}1回戦勝者4`,
      type: "tournament",
      isFinal: false,
    });
    // 準決勝
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 3,
      roundName: "準決勝",
      team1: `${prefix}準々決勝勝者1`,
      team2: `${prefix}準々決勝勝者2`,
      type: "tournament",
      isFinal: false,
    });
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 3,
      roundName: "準決勝",
      team1: `${prefix}準々決勝勝者3`,
      team2: `${prefix}準々決勝勝者4`,
      type: "tournament",
      isFinal: false,
    });
    // 3位決定戦 + 決勝
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 4,
      roundName: "3位決定戦",
      team1: `${prefix}準決勝敗者1`,
      team2: `${prefix}準決勝敗者2`,
      type: "tournament",
      isFinal: false,
      isThirdPlace: true,
    });
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: 4,
      roundName: "決勝",
      team1: `${prefix}準決勝勝者1`,
      team2: `${prefix}準決勝勝者2`,
      type: "tournament",
      isFinal: true,
    });
  } else {
    // 13チーム以上: 汎用ロジック（簡易版）
    const rounds = Math.ceil(Math.log2(teams));
    const bracketSize = Math.pow(2, rounds);
    const byeCount = bracketSize - teams;

    // 1回戦
    const firstRoundMatches = bracketSize / 2 - byeCount;
    for (let i = 0; i < firstRoundMatches; i++) {
      const team1Index = teams - firstRoundMatches * 2 + i * 2;
      const team2Index = team1Index + 1;
      matches.push({
        matchNumber: matchNum++,
        category: category,
        round: 1,
        roundName: "1回戦",
        team1: `${prefix}${TEAM_NAMES[team1Index]}`,
        team2: `${prefix}${TEAM_NAMES[team2Index]}`,
        type: "tournament",
        isFinal: false,
      });
    }

    // 2回戦以降
    for (let round = 2; round <= rounds; round++) {
      const matchesInRound = Math.pow(2, rounds - round);
      const roundName = getTournamentRoundName(round, rounds);

      for (let i = 0; i < matchesInRound; i++) {
        matches.push({
          matchNumber: matchNum++,
          category: category,
          round: round,
          roundName: roundName,
          team1: `${prefix}${roundName}進出${i * 2 + 1}`,
          team2: `${prefix}${roundName}進出${i * 2 + 2}`,
          type: "tournament",
          isFinal: round === rounds,
        });
      }
    }

    // 3位決定戦
    if (teams >= 4) {
      matches.push({
        matchNumber: matchNum++,
        category: category,
        round: rounds,
        roundName: "3位決定戦",
        team1: `${prefix}準決勝敗者1`,
        team2: `${prefix}準決勝敗者2`,
        type: "tournament",
        isFinal: false,
        isThirdPlace: true,
      });
    }
  }

  return matches;
}

function getTournamentRoundName(round, totalRounds) {
  const remaining = totalRounds - round;
  if (remaining === 0) return "決勝";
  if (remaining === 1) return "準決勝";
  if (remaining === 2) return "準々決勝";
  return `${round}回戦`;
}

function getRoundName(round, totalRounds) {
  return getTournamentRoundName(round, totalRounds);
}

/**
 * リーグ戦試合リスト生成（順位決定戦なし）
 */
function generateLeagueMatchList(teams, category, prefix, info) {
  const matches = [];
  let matchNum = 1;

  if (info.structure === "single_league") {
    const teamList = [];
    for (let i = 0; i < teams; i++) {
      teamList.push(`${prefix}${TEAM_NAMES[i]}`);
    }

    // ラウンドロビン方式で試合順序を最適化
    const roundRobinPairs = generateRoundRobinOrder(teams);
    for (const pair of roundRobinPairs) {
      matches.push({
        matchNumber: matchNum++,
        category: category,
        group: 1,
        groupName: "",
        team1: teamList[pair[0]],
        team2: teamList[pair[1]],
        type: "league",
        phase: "group",
      });
    }
  } else if (info.structure === "two_group_league") {
    // グループA
    matches.push(
      ...generateGroupMatchList(
        info.teamsPerGroup[0],
        category,
        prefix,
        "A",
        matchNum,
      ),
    );
    matchNum = matches.length + 1;
    // グループB
    matches.push(
      ...generateGroupMatchList(
        info.teamsPerGroup[1],
        category,
        prefix,
        "B",
        matchNum,
      ),
    );
    // 順位決定戦は生成しない
  } else if (info.structure === "three_group_league") {
    const groupLabels = ["A", "B", "C"];
    for (let g = 0; g < 3; g++) {
      matches.push(
        ...generateGroupMatchList(
          info.teamsPerGroup[g],
          category,
          prefix,
          groupLabels[g],
          matchNum,
        ),
      );
      matchNum = matches.length + 1;
    }
    // 順位決定戦は生成しない
  }

  return matches;
}

function generateGroupMatchList(
  teamCount,
  category,
  prefix,
  groupLabel,
  startNum,
) {
  const matches = [];
  const teamList = [];

  for (let i = 0; i < teamCount; i++) {
    teamList.push(`${prefix}${groupLabel}${i + 1}`);
  }

  // ラウンドロビン方式で試合順序を最適化（連戦を最小化）
  const roundRobinPairs = generateRoundRobinOrder(teamCount);

  let matchNum = startNum;
  for (const pair of roundRobinPairs) {
    matches.push({
      matchNumber: matchNum++,
      category: category,
      group: groupLabel,
      groupName: `グループ${groupLabel}`,
      team1: teamList[pair[0]],
      team2: teamList[pair[1]],
      type: "league",
      phase: "group",
    });
  }

  return matches;
}

/**
 * ラウンドロビン方式で試合ペアを生成（連戦を最小化する順序）
 * ベルジュ・テーブル方式（Circle Method）を使用
 */
function generateRoundRobinOrder(n) {
  const pairs = [];

  if (n === 3) {
    // 3チーム: 最適順序は難しいが、できるだけ分散
    pairs.push([0, 1], [0, 2], [1, 2]);
  } else if (n === 4) {
    // 4チーム: ラウンドロビン最適順序
    // ラウンド1: 0vs3, 1vs2
    // ラウンド2: 0vs2, 1vs3
    // ラウンド3: 0vs1, 2vs3
    pairs.push([0, 3], [1, 2], [0, 2], [1, 3], [0, 1], [2, 3]);
  } else if (n === 5) {
    // 5チーム: 各チーム4試合
    pairs.push(
      [0, 4],
      [1, 3],
      [0, 2],
      [3, 4],
      [1, 2],
      [0, 3],
      [2, 4],
      [0, 1],
      [2, 3],
      [1, 4],
    );
  } else if (n >= 6) {
    // ベルジュ・テーブル方式で生成（6チーム以上）
    // 奇数の場合はダミーチームを追加して偶数にする
    const isOdd = n % 2 === 1;
    const teamCount = isOdd ? n + 1 : n;
    const rounds = generateBergerTable(teamCount);

    for (const round of rounds) {
      for (const pair of round) {
        // ダミーチーム（インデックス n）との対戦は除外
        if (!isOdd || (pair[0] < n && pair[1] < n)) {
          pairs.push(pair);
        }
      }
    }
  } else {
    // フォールバック: 標準的な総当たり
    for (let i = 0; i < n; i++) {
      for (let j = i + 1; j < n; j++) {
        pairs.push([i, j]);
      }
    }
  }

  return pairs;
}

/**
 * ベルジュ・テーブル（Circle Method）で対戦表を生成
 * n チーム（偶数）で n-1 ラウンド、各ラウンド n/2 試合
 * 各チームは各ラウンドで1試合のみ（連戦なし）
 */
function generateBergerTable(n) {
  const rounds = [];
  const teams = [];

  // チーム0を固定、残りを回転させる
  for (let i = 1; i < n; i++) {
    teams.push(i);
  }

  for (let round = 0; round < n - 1; round++) {
    const roundPairs = [];

    // チーム0 vs 現在の先頭チーム
    roundPairs.push([0, teams[0]]);

    // 残りのペアリング（対角線で組む）
    // teams配列の長さは n-1 なので、teams[i] と teams[n-1-i] をペアにする
    for (let i = 1; i < n / 2; i++) {
      const team1 = teams[i];
      const team2 = teams[n - 1 - i];
      roundPairs.push([Math.min(team1, team2), Math.max(team1, team2)]);
    }

    rounds.push(roundPairs);

    // 配列を回転（最後の要素を先頭に移動）
    const last = teams.pop();
    teams.unshift(last);
  }

  return rounds;
}

// ============================================================
// Scheduler.gs - スケジュール最適化
// ============================================================

function selectOptimalRegulations(input) {
  const availableTime = calculateAvailableTime(input);
  const u12Matches = input.u12.enabled
    ? calculateMatchCount(input.u12.teams, input.u12.matchType)
    : null;
  const u15Matches = input.u15.enabled
    ? calculateMatchCount(input.u15.teams, input.u15.matchType)
    : null;

  // ユーザー選択のレギュレーションを取得
  const u12SelectedReg = getSelectedRegulation(input.u12.regulation, "U12");
  const u15SelectedReg = getSelectedRegulation(input.u15.regulation, "U15");

  // ユーザーが選択した場合はそれを使用
  if (
    u12SelectedReg ||
    u15SelectedReg ||
    (input.u12.regulation !== "自動" && input.u12.enabled) ||
    (input.u15.regulation !== "自動" && input.u15.enabled)
  ) {
    const u12Reg = input.u12.enabled
      ? u12SelectedReg || getRegulationsByCategory("U12")[0]
      : null;
    const u15Reg = input.u15.enabled
      ? u15SelectedReg || getRegulationsByCategory("U15")[0]
      : null;

    const evaluation = evaluateRegulationCombination(
      input,
      availableTime,
      u12Matches,
      u15Matches,
      u12Reg,
      u15Reg,
    );

    return {
      u12Regulation: u12Reg,
      u15Regulation: u15Reg,
      ...evaluation,
    };
  }

  // 自動選択の場合は最適なものを探す
  // 時間内に収まる中で、最大のスロット時間（時間を最大限活用）を選択
  const u12Regs = input.u12.enabled ? getRegulationsByCategory("U12") : [];
  const u15Regs = input.u15.enabled ? getRegulationsByCategory("U15") : [];

  let bestResult = null;
  let bestSlotTotal = -1; // スロット時間の合計が最大のものを選ぶ

  // スロット時間が大きい順にソート（降順）
  const u12Options =
    u12Regs.length > 0
      ? [...u12Regs].sort((a, b) => b.slot_minutes - a.slot_minutes)
      : [null];
  const u15Options =
    u15Regs.length > 0
      ? [...u15Regs].sort((a, b) => b.slot_minutes - a.slot_minutes)
      : [null];

  for (const u12Reg of u12Options) {
    for (const u15Reg of u15Options) {
      const evaluation = evaluateRegulationCombination(
        input,
        availableTime,
        u12Matches,
        u15Matches,
        u12Reg,
        u15Reg,
      );

      // 時間内に収まる（margin >= 0）かつスロット時間の合計が最大
      if (evaluation.isValid && evaluation.margin >= 0) {
        const slotTotal =
          (u12Reg ? u12Reg.slot_minutes : 0) +
          (u15Reg ? u15Reg.slot_minutes : 0);

        if (slotTotal > bestSlotTotal) {
          bestSlotTotal = slotTotal;
          bestResult = {
            u12Regulation: u12Reg,
            u15Regulation: u15Reg,
            ...evaluation,
          };
        }
      }
    }
  }

  // 時間内に収まるものがない場合、最小余裕時間（負でも）のものを選択
  if (!bestResult) {
    let bestMargin = -Infinity;
    for (const u12Reg of u12Options) {
      for (const u15Reg of u15Options) {
        const evaluation = evaluateRegulationCombination(
          input,
          availableTime,
          u12Matches,
          u15Matches,
          u12Reg,
          u15Reg,
        );

        if (evaluation.margin > bestMargin) {
          bestMargin = evaluation.margin;
          bestResult = {
            u12Regulation: u12Reg,
            u15Regulation: u15Reg,
            ...evaluation,
          };
        }
      }
    }
  }

  return bestResult;
}

/**
 * ユーザー選択からレギュレーションを取得
 */
function getSelectedRegulation(selection, category) {
  if (!selection || selection === "自動") return null;

  // "60分" -> 60 のように数値を抽出
  const match = selection.match(/(\d+)/);
  if (!match) return null;

  const minutes = parseInt(match[1]);

  // 該当するレギュレーションを探す
  const regId = `SLOT${minutes}`;
  if (REGULATIONS[regId]) {
    return {
      id: regId,
      ...REGULATIONS[regId],
      category: category,
    };
  }

  return null;
}

function calculateAvailableTime(input) {
  const start = input.common.startTime;
  const end = input.common.endTime;

  const startMinutes = start.getHours() * 60 + start.getMinutes();
  const endMinutes = end.getHours() * 60 + end.getMinutes();

  let totalMinutes = endMinutes - startMinutes;

  // 入力された時間を使用
  totalMinutes -= input.common.setupMinutes;
  totalMinutes -= input.common.receptionMinutes;
  totalMinutes -= input.common.cleanupMinutes;

  if (input.common.hasExhibition) {
    totalMinutes -= input.common.exhibitionMinutes;
  }
  if (input.common.hasCeremony) {
    totalMinutes -= input.common.ceremonyMinutes;
  }
  if (input.common.hasPhoto) {
    totalMinutes -= input.common.photoMinutes;
  }

  return totalMinutes;
}

function evaluateRegulationCombination(
  input,
  availableTime,
  u12Matches,
  u15Matches,
  u12Reg,
  u15Reg,
) {
  const courts = input.common.courts;
  const u12Total = u12Matches ? u12Matches.total : 0;
  const u15Total = u15Matches ? u15Matches.total : 0;

  // スロット時間
  const u12SlotMin = u12Reg ? u12Reg.slot_minutes : 0;
  const u15SlotMin = u15Reg ? u15Reg.slot_minutes : 0;

  // トーナメント戦かどうかを確認
  const u12IsTournament =
    input.u12.enabled && input.u12.matchType === CONFIG.MATCH_TYPES.TOURNAMENT;
  const u15IsTournament =
    input.u15.enabled && input.u15.matchType === CONFIG.MATCH_TYPES.TOURNAMENT;

  let totalSlots;
  let totalTime;

  if (u12IsTournament || u15IsTournament) {
    // トーナメント戦の場合
    if (courts >= 2 && u12Total > 0 && u15Total > 0) {
      // 2コートで両カテゴリ: Aコート=U12、Bコート=U15
      totalSlots = Math.max(u12Total, u15Total);
      const maxSlotMinutes = Math.max(u12SlotMin, u15SlotMin);
      totalTime = totalSlots * maxSlotMinutes;
    } else if (courts >= 2) {
      // 2コートで1カテゴリのみ: 同じラウンドの試合を2コートで同時に
      const singleTotal = u12Total + u15Total;
      totalSlots = Math.ceil(singleTotal / courts);
      const maxSlotMinutes = Math.max(u12SlotMin, u15SlotMin);
      totalTime = totalSlots * maxSlotMinutes;
    } else {
      // 1コート
      totalSlots = u12Total + u15Total;
      const maxSlotMinutes = Math.max(u12SlotMin, u15SlotMin);
      totalTime = totalSlots * maxSlotMinutes;
    }
  } else {
    // リーグ戦の場合：カテゴリ交互配置を考慮
    if (courts >= 2 && u12Total > 0 && u15Total > 0) {
      // 2コートで両カテゴリがある場合、カテゴリごとにスロット数を計算
      const u12Slots = Math.ceil(u12Total / courts);
      const u15Slots = Math.ceil(u15Total / courts);
      totalSlots = u12Slots + u15Slots;
      // 各カテゴリの時間を個別に計算
      totalTime = u12Slots * u12SlotMin + u15Slots * u15SlotMin;
    } else {
      // 1コートまたは片方のカテゴリのみ
      totalSlots = Math.ceil((u12Total + u15Total) / courts);
      const maxSlotMinutes = Math.max(u12SlotMin, u15SlotMin);
      totalTime = totalSlots * maxSlotMinutes;
    }
  }

  // 試合間インターバルを追加
  if (totalSlots > 1) {
    // 1〜2試合間は15分（ONの場合）、それ以外は入力値（デフォルト10分）
    const firstInterval = input.common.firstInterval15
      ? FIRST_INTERVAL_MINUTES
      : input.common.intervalMinutes;
    totalTime += firstInterval; // 1〜2試合間
    if (totalSlots > 2) {
      totalTime += (totalSlots - 2) * input.common.intervalMinutes; // 残りは入力値
    }
  }

  const margin = availableTime - totalTime;
  const isValid = margin >= 0;

  return {
    isValid: isValid,
    totalSlots: totalSlots,
    totalTime: totalTime,
    availableTime: availableTime,
    margin: margin,
  };
}

function generateSchedule(input, regulations) {
  const u12Matches = input.u12.enabled
    ? generateMatchList(input.u12.teams, input.u12.matchType, "U12", "U12-")
    : [];
  const u15Matches = input.u15.enabled
    ? generateMatchList(input.u15.teams, input.u15.matchType, "U15", "U15-")
    : [];

  const schedule = arrangeMatchesWithConstraints(
    u12Matches,
    u15Matches,
    input,
    regulations,
  );

  return schedule;
}

function arrangeMatchesWithConstraints(
  u12Matches,
  u15Matches,
  input,
  regulations,
) {
  const courts = input.common.courts;
  const slots = [];
  let slotIndex = 0;

  // トーナメント戦はラウンド順にソート、リーグ戦はグループ別にシャッフル
  const u12Queue = sortMatchesForScheduling([...u12Matches]);
  const u15Queue = sortMatchesForScheduling([...u15Matches]);
  const lastMatchSlot = {};
  const usedInSlot = new Set();

  const hasBothCategories = u12Queue.length > 0 && u15Queue.length > 0;

  // トーナメント戦が含まれるかチェック
  const hasTournament =
    (input.u12.enabled &&
      input.u12.matchType === CONFIG.MATCH_TYPES.TOURNAMENT) ||
    (input.u15.enabled &&
      input.u15.matchType === CONFIG.MATCH_TYPES.TOURNAMENT);

  // 3チームグループがあるかチェック
  const hasSmallGroup = checkHasSmallGroup(u12Matches, u15Matches);

  // トーナメント戦で2コートの場合
  if (hasTournament && courts >= 2) {
    if (hasBothCategories) {
      // 両カテゴリ: Aコート=U12、Bコート=U15
      return arrangeMatchesByCourtCategory(
        u12Queue,
        u15Queue,
        input,
        regulations,
      );
    } else {
      // 1カテゴリのみ: 同じラウンドの試合を2コートで同時に行う
      const singleQueue = u12Queue.length > 0 ? u12Queue : u15Queue;
      return arrangeTournamentSingleCategory(singleQueue, input, regulations);
    }
  }

  // リーグ戦で両カテゴリ + 2コート + 小グループあり → コート分離配置で連戦回避
  if (hasBothCategories && courts >= 2 && hasSmallGroup) {
    // Aコート=U12、Bコート=U15として分離配置
    // 各カテゴリは既にshuffleByGroupでグループ交互配置されているので連戦なし
    return arrangeMatchesByCourtCategory(
      u12Queue,
      u15Queue,
      input,
      regulations,
    );
  }

  // リーグ戦のみの場合: カテゴリ交互配置（効率重視）
  let currentCategory = "U12";

  while (u12Queue.length > 0 || u15Queue.length > 0) {
    const slot = {
      index: slotIndex,
      matches: [],
    };
    usedInSlot.clear();

    // このスロットで使用するキューを決定
    let primaryQueue, secondaryQueue;
    if (hasBothCategories && courts >= 2) {
      // 同じスロット内は同じカテゴリで統一
      if (currentCategory === "U12" && u12Queue.length > 0) {
        primaryQueue = u12Queue;
        secondaryQueue = u15Queue;
      } else if (currentCategory === "U15" && u15Queue.length > 0) {
        primaryQueue = u15Queue;
        secondaryQueue = u12Queue;
      } else if (u12Queue.length > 0) {
        primaryQueue = u12Queue;
        secondaryQueue = u15Queue;
      } else {
        primaryQueue = u15Queue;
        secondaryQueue = u12Queue;
      }
    } else {
      // 1コートまたは片方のカテゴリのみ
      primaryQueue = u12Queue.length > 0 ? u12Queue : u15Queue;
      secondaryQueue = u12Queue.length > 0 ? u15Queue : u12Queue;
    }

    // 3チームグループがある場合、最初は1コートで開始して連戦回避
    // 連戦なしで2コート配置可能になったら2コートに切り替え
    const effectiveCourts =
      hasSmallGroup &&
      shouldUseSingleCourt(primaryQueue, lastMatchSlot, slotIndex)
        ? 1
        : courts;

    for (let court = 0; court < effectiveCourts; court++) {
      let selectedMatch = null;
      let selectedQueue = null;

      // まずプライマリキューから連戦なしの試合を探す
      selectedMatch = findValidMatchStrict(
        primaryQueue,
        lastMatchSlot,
        slotIndex,
        usedInSlot,
      );
      if (selectedMatch) {
        selectedQueue = primaryQueue;
      }

      // プライマリで見つからない場合、セカンダリから探す（片方が終わった場合用）
      if (!selectedMatch && primaryQueue.length === 0) {
        selectedMatch = findValidMatchStrict(
          secondaryQueue,
          lastMatchSlot,
          slotIndex,
          usedInSlot,
        );
        if (selectedMatch) selectedQueue = secondaryQueue;
      }

      // 連戦なしの試合が見つからない場合の処理
      if (!selectedMatch) {
        if (court === 0) {
          // 1コート目で連戦なしがない場合は、連戦覚悟で1試合だけ配置
          // （スキップせず運営を続ける）
          selectedMatch = findBestMatchStrict(
            primaryQueue,
            lastMatchSlot,
            slotIndex,
            usedInSlot,
          );
          if (selectedMatch) {
            selectedQueue = primaryQueue;
          } else if (primaryQueue.length === 0) {
            selectedMatch = findBestMatchStrict(
              secondaryQueue,
              lastMatchSlot,
              slotIndex,
              usedInSlot,
            );
            if (selectedMatch) selectedQueue = secondaryQueue;
          }
          // 連戦覚悟で配置した場合、試合配置後に1コートで終了するフラグ
          if (selectedMatch) {
            // 試合を配置した後breakでこのスロットを1コートで終了
            const idx = selectedQueue.indexOf(selectedMatch);
            if (idx > -1) {
              selectedQueue.splice(idx, 1);
            }
            lastMatchSlot[selectedMatch.team1] = slotIndex;
            lastMatchSlot[selectedMatch.team2] = slotIndex;
            usedInSlot.add(selectedMatch.team1);
            usedInSlot.add(selectedMatch.team2);
            slot.matches.push({
              ...selectedMatch,
              court: COURT_NAMES[court],
              slotIndex: slotIndex,
            });
            break; // 1コートで終了
          }
        }
        // 2コート目以降、または1コート目でも配置できない場合はスロット終了
        break;
      }

      // 試合を配置
      if (selectedMatch) {
        const idx = selectedQueue.indexOf(selectedMatch);
        if (idx > -1) {
          selectedQueue.splice(idx, 1);
        }

        lastMatchSlot[selectedMatch.team1] = slotIndex;
        lastMatchSlot[selectedMatch.team2] = slotIndex;
        usedInSlot.add(selectedMatch.team1);
        usedInSlot.add(selectedMatch.team2);

        slot.matches.push({
          ...selectedMatch,
          court: COURT_NAMES[court],
          slotIndex: slotIndex,
        });
      }
    }

    if (slot.matches.length > 0) {
      slots.push(slot);
      slotIndex++;
      // 次のスロットはカテゴリを切り替え（両方残っている場合のみ）
      if (u12Queue.length > 0 && u15Queue.length > 0) {
        currentCategory = currentCategory === "U12" ? "U15" : "U12";
      }
    } else {
      break;
    }
  }

  return assignTimes(slots, input, regulations);
}

/**
 * トーナメント戦用: Aコート=U12、Bコート=U15として配置
 */
function arrangeMatchesByCourtCategory(u12Queue, u15Queue, input, regulations) {
  const slots = [];
  let slotIndex = 0;

  // 各カテゴリのキューをラウンド順にソート（既にソート済みのはず）
  let u12Index = 0;
  let u15Index = 0;

  while (u12Index < u12Queue.length || u15Index < u15Queue.length) {
    const slot = {
      index: slotIndex,
      matches: [],
    };

    // Aコート: U12の試合
    if (u12Index < u12Queue.length) {
      const u12Match = u12Queue[u12Index];
      slot.matches.push({
        ...u12Match,
        court: COURT_NAMES[0], // Aコート
        slotIndex: slotIndex,
      });
      u12Index++;
    }

    // Bコート: U15の試合
    if (u15Index < u15Queue.length) {
      const u15Match = u15Queue[u15Index];
      slot.matches.push({
        ...u15Match,
        court: COURT_NAMES[1], // Bコート
        slotIndex: slotIndex,
      });
      u15Index++;
    }

    if (slot.matches.length > 0) {
      slots.push(slot);
      slotIndex++;
    } else {
      break;
    }
  }

  return assignTimes(slots, input, regulations);
}

/**
 * 1カテゴリのトーナメント戦で2コートを使う場合
 * 同じラウンドの試合を2コートで同時に行う
 */
function arrangeTournamentSingleCategory(queue, input, regulations) {
  const courts = input.common.courts;
  const slots = [];
  let slotIndex = 0;

  // ラウンドごとにグループ化
  const roundGroups = {};
  for (const match of queue) {
    const roundKey = match.isThirdPlace ? "thirdPlace" : match.round;
    if (!roundGroups[roundKey]) {
      roundGroups[roundKey] = [];
    }
    roundGroups[roundKey].push(match);
  }

  // ラウンド順にソート（3位決定戦は決勝の前）
  const sortedRounds = Object.keys(roundGroups).sort((a, b) => {
    if (a === "thirdPlace") return 1; // 3位決定戦は最後の方
    if (b === "thirdPlace") return -1;
    return parseInt(a) - parseInt(b);
  });

  // 決勝と3位決定戦を同時にできるように調整
  const finalIndex = sortedRounds.indexOf(
    sortedRounds.find(
      (r) => r !== "thirdPlace" && roundGroups[r].some((m) => m.isFinal),
    ),
  );
  const thirdPlaceIndex = sortedRounds.indexOf("thirdPlace");
  if (finalIndex !== -1 && thirdPlaceIndex !== -1) {
    // 3位決定戦を決勝と同じ位置に移動
    sortedRounds.splice(thirdPlaceIndex, 1);
    sortedRounds.splice(finalIndex, 0, "thirdPlace");
  }

  for (const roundKey of sortedRounds) {
    const roundMatches = roundGroups[roundKey];
    let matchIndex = 0;

    while (matchIndex < roundMatches.length) {
      const slot = {
        index: slotIndex,
        matches: [],
      };

      // 最大courtsコート分の試合を配置
      for (
        let court = 0;
        court < courts && matchIndex < roundMatches.length;
        court++
      ) {
        const match = roundMatches[matchIndex];
        slot.matches.push({
          ...match,
          court: COURT_NAMES[court],
          slotIndex: slotIndex,
        });
        matchIndex++;
      }

      if (slot.matches.length > 0) {
        slots.push(slot);
        slotIndex++;
      }
    }
  }

  return assignTimes(slots, input, regulations);
}

/**
 * 試合をスケジューリング用にソート
 * - トーナメント戦: ラウンド順（準決勝→3位決定戦→決勝）
 * - リーグ戦: グループ交互配置
 */
function sortMatchesForScheduling(matches) {
  if (matches.length === 0) return matches;

  // トーナメント戦かどうかをチェック
  const hasTournament = matches.some((m) => m.type === "tournament");

  if (hasTournament) {
    // トーナメント戦: ラウンド順にソート
    // round 1（準決勝など）→ 3位決定戦 → 決勝 の順
    return [...matches].sort((a, b) => {
      // 3位決定戦は決勝の前（roundは同じだがisFinalがfalse）
      if (a.isThirdPlace && !b.isThirdPlace && b.isFinal) return -1;
      if (b.isThirdPlace && !a.isThirdPlace && a.isFinal) return 1;
      // ラウンド順
      if (a.round !== b.round) return a.round - b.round;
      return a.matchNumber - b.matchNumber;
    });
  } else {
    // リーグ戦: グループ交互配置
    return shuffleByGroup(matches);
  }
}

/**
 * グループ別に試合を並べ替える
 * グループ間で1試合ずつ交互に配置することで連戦を最小化
 * 例: A1, B1, A2, B2, A3, B3...
 * これにより1コートでも連戦回避が可能になる
 */
function shuffleByGroup(matches) {
  const byGroup = {};
  for (const match of matches) {
    const group = match.groupName || match.roundName || "default";
    if (!byGroup[group]) byGroup[group] = [];
    byGroup[group].push(match);
  }

  const groups = Object.keys(byGroup).sort();
  const result = [];

  // 各グループから1試合ずつ交互に取り出す
  const groupIndices = {};
  for (const group of groups) {
    groupIndices[group] = 0;
  }

  let hasMore = true;
  while (hasMore) {
    hasMore = false;
    for (const group of groups) {
      const groupMatches = byGroup[group];
      const idx = groupIndices[group];

      if (idx < groupMatches.length) {
        result.push(groupMatches[idx]);
        groupIndices[group] = idx + 1;
        hasMore = true;
      }
    }
  }

  return result;
}

/**
 * 3チームグループがあるかチェック
 * 3チームグループは連戦回避が難しいため、特別な配置が必要
 */
function checkHasSmallGroup(u12Matches, u15Matches) {
  const allMatches = [...u12Matches, ...u15Matches];
  const groupSizes = {};

  for (const match of allMatches) {
    if (match.type === "league") {
      // グループ名がない場合は"default"として扱う（単一リーグ）
      const groupName = match.groupName || "default";
      const key = `${match.category}-${groupName}`;
      if (!groupSizes[key]) {
        groupSizes[key] = new Set();
      }
      groupSizes[key].add(match.team1);
      groupSizes[key].add(match.team2);
    }
  }

  // いずれかのグループが3チーム以下ならtrue
  for (const key in groupSizes) {
    if (groupSizes[key].size <= 3) {
      return true;
    }
  }
  return false;
}

/**
 * 1コートで配置すべきかチェック
 * 2コートで配置すると次のスロットで全試合が連戦になる場合はtrue
 */
function shouldUseSingleCourt(queue, lastMatchSlot, currentSlot) {
  if (queue.length < 2) return false;

  // 連戦なしで配置可能な試合を数える
  let validCount = 0;
  const tempUsedInSlot = new Set();

  for (const match of queue) {
    if (tempUsedInSlot.has(match.team1) || tempUsedInSlot.has(match.team2)) {
      continue;
    }

    const team1LastSlot = lastMatchSlot[match.team1];
    const team2LastSlot = lastMatchSlot[match.team2];
    const team1OK =
      team1LastSlot === undefined || currentSlot - team1LastSlot > 1;
    const team2OK =
      team2LastSlot === undefined || currentSlot - team2LastSlot > 1;

    if (team1OK && team2OK) {
      validCount++;
      tempUsedInSlot.add(match.team1);
      tempUsedInSlot.add(match.team2);
    }

    if (validCount >= 2) break;
  }

  // 連戦なしで2試合以上配置可能なら2コートOK
  if (validCount >= 2) {
    // さらに、2試合配置した後に残り試合が全て連戦になるかチェック
    // 2試合配置後のシミュレーション
    const simLastMatchSlot = { ...lastMatchSlot };
    const simUsedInSlot = new Set();
    let simCount = 0;

    for (const match of queue) {
      if (simUsedInSlot.has(match.team1) || simUsedInSlot.has(match.team2)) {
        continue;
      }

      const team1LastSlot = simLastMatchSlot[match.team1];
      const team2LastSlot = simLastMatchSlot[match.team2];
      const team1OK =
        team1LastSlot === undefined || currentSlot - team1LastSlot > 1;
      const team2OK =
        team2LastSlot === undefined || currentSlot - team2LastSlot > 1;

      if (team1OK && team2OK) {
        simLastMatchSlot[match.team1] = currentSlot;
        simLastMatchSlot[match.team2] = currentSlot;
        simUsedInSlot.add(match.team1);
        simUsedInSlot.add(match.team2);
        simCount++;
        if (simCount >= 2) break;
      }
    }

    if (simCount >= 2) {
      // 2試合配置後、次のスロットで連戦なしの試合があるかチェック
      const nextSlot = currentSlot + 1;
      // 配置済みの試合を除外（両方のチームがsimUsedInSlotにある試合）
      const remainingQueue = queue.filter(
        (m) => !(simUsedInSlot.has(m.team1) && simUsedInSlot.has(m.team2)),
      );

      for (const match of remainingQueue) {
        const team1LastSlot = simLastMatchSlot[match.team1];
        const team2LastSlot = simLastMatchSlot[match.team2];
        const team1OK =
          team1LastSlot === undefined || nextSlot - team1LastSlot > 1;
        const team2OK =
          team2LastSlot === undefined || nextSlot - team2LastSlot > 1;

        if (team1OK && team2OK) {
          // 次のスロットで連戦なしの試合がある → 2コートOK
          return false;
        }
      }

      // 次のスロットで全試合が連戦 → 1コートにして回避
      return true;
    }
  }

  return false;
}

/**
 * 連戦にならない試合を探す（同スロット内の重複も避ける）
 */
function findValidMatchStrict(queue, lastMatchSlot, currentSlot, usedInSlot) {
  // トーナメント戦の場合、前のラウンドが終わるまで次のラウンドは選択しない
  const minRound = getMinRoundInQueue(queue);

  for (const match of queue) {
    // トーナメント戦で、最小ラウンドより後の試合はスキップ
    if (match.type === "tournament" && match.round > minRound) {
      continue;
    }

    // 同じスロット内で既に使用されているチームは除外
    if (usedInSlot.has(match.team1) || usedInSlot.has(match.team2)) {
      continue;
    }

    const team1LastSlot = lastMatchSlot[match.team1];
    const team2LastSlot = lastMatchSlot[match.team2];

    const team1OK =
      team1LastSlot === undefined || currentSlot - team1LastSlot > 1;
    const team2OK =
      team2LastSlot === undefined || currentSlot - team2LastSlot > 1;

    if (team1OK && team2OK) {
      return match;
    }
  }
  return null;
}

/**
 * キュー内の最小ラウンド番号を取得（トーナメント用）
 */
function getMinRoundInQueue(queue) {
  let minRound = Infinity;
  for (const match of queue) {
    if (match.type === "tournament" && match.round < minRound) {
      minRound = match.round;
    }
  }
  return minRound === Infinity ? 1 : minRound;
}

/**
 * 連戦にならない試合を探す（旧版、互換性用）
 */
function findValidMatch(queue, lastMatchSlot, currentSlot) {
  return findValidMatchStrict(queue, lastMatchSlot, currentSlot, new Set());
}

/**
 * 最も休息時間が長いチームの試合を選択（連戦を最小化）
 */
/**
 * 最も休息時間が長いチームの試合を選択（同スロット内の重複も避ける）
 */
function findBestMatchStrict(queue, lastMatchSlot, currentSlot, usedInSlot) {
  if (queue.length === 0) return null;

  // トーナメント戦の場合、前のラウンドが終わるまで次のラウンドは選択しない
  const minRound = getMinRoundInQueue(queue);

  let bestMatch = null;
  let bestScore = -Infinity;

  for (const match of queue) {
    // トーナメント戦で、最小ラウンドより後の試合はスキップ
    if (match.type === "tournament" && match.round > minRound) {
      continue;
    }

    // 同じスロット内で既に使用されているチームは除外
    if (usedInSlot.has(match.team1) || usedInSlot.has(match.team2)) {
      continue;
    }

    const team1LastSlot = lastMatchSlot[match.team1];
    const team2LastSlot = lastMatchSlot[match.team2];

    // 休息スロット数（未プレイなら大きな値）
    const team1Rest =
      team1LastSlot === undefined ? 100 : currentSlot - team1LastSlot;
    const team2Rest =
      team2LastSlot === undefined ? 100 : currentSlot - team2LastSlot;

    // 両チームの最小休息時間をスコアとする
    const score = Math.min(team1Rest, team2Rest);

    if (score > bestScore) {
      bestScore = score;
      bestMatch = match;
    }
  }

  return bestMatch;
}

/**
 * 旧版（互換性用）
 */
function findBestMatch(queue, lastMatchSlot, currentSlot) {
  return findBestMatchStrict(queue, lastMatchSlot, currentSlot, new Set());
}

function assignTimes(slots, input, regulations) {
  const schedule = [];

  const startTime = new Date(input.common.startTime);
  startTime.setMinutes(
    startTime.getMinutes() +
      input.common.setupMinutes +
      input.common.receptionMinutes,
  );

  let currentTime = new Date(startTime);

  for (let i = 0; i < slots.length; i++) {
    const slot = slots[i];

    for (const match of slot.matches) {
      const slotMinutes =
        match.category === "U12"
          ? regulations.u12Regulation.slot_minutes
          : regulations.u15Regulation.slot_minutes;

      const endTime = new Date(currentTime);
      endTime.setMinutes(endTime.getMinutes() + slotMinutes);

      schedule.push({
        ...match,
        startTime: new Date(currentTime),
        endTime: endTime,
        slotMinutes: slotMinutes,
      });
    }

    const maxSlotMinutes = Math.max(
      ...slot.matches.map((m) =>
        m.category === "U12"
          ? regulations.u12Regulation.slot_minutes
          : regulations.u15Regulation.slot_minutes,
      ),
    );

    // 次のスロットへ（試合時間 + インターバル）
    currentTime.setMinutes(currentTime.getMinutes() + maxSlotMinutes);

    // 最後のスロット以外はインターバルを追加
    if (i < slots.length - 1) {
      // 1〜2試合間は15分（ONの場合）、それ以外は入力値（デフォルト10分）
      const interval =
        i === 0 && input.common.firstInterval15
          ? FIRST_INTERVAL_MINUTES
          : input.common.intervalMinutes;
      currentTime.setMinutes(currentTime.getMinutes() + interval);
    }
  }

  return schedule;
}

function validateSchedule(schedule, input) {
  const endTime = input.common.endTime;
  const errors = [];
  const warnings = [];
  const consecutiveWarnings = []; // 連戦警告を別途管理

  if (schedule.length === 0) {
    errors.push("配置可能な試合がありません");
    return { isValid: false, errors, warnings, consecutiveWarnings };
  }

  const lastMatch = schedule[schedule.length - 1];
  const lastMatchEnd = lastMatch.endTime;

  const finalEnd = new Date(lastMatchEnd);
  if (input.common.hasExhibition) {
    finalEnd.setMinutes(finalEnd.getMinutes() + input.common.exhibitionMinutes);
  }
  if (input.common.hasCeremony) {
    finalEnd.setMinutes(finalEnd.getMinutes() + input.common.ceremonyMinutes);
  }
  if (input.common.hasPhoto) {
    finalEnd.setMinutes(finalEnd.getMinutes() + input.common.photoMinutes);
  }
  finalEnd.setMinutes(finalEnd.getMinutes() + input.common.cleanupMinutes);

  let timeOverflow = null;
  if (finalEnd > endTime) {
    const overMinutes = Math.ceil((finalEnd - endTime) / (1000 * 60));
    timeOverflow = {
      overMinutes: overMinutes,
      finalEnd: finalEnd,
      endTime: endTime,
    };
    warnings.push(`利用終了時刻を${overMinutes}分超過します。`);
  }

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
      if (slots[i] - slots[i - 1] <= 1) {
        const msg = `${team}が連戦になっています（スロット${slots[i - 1] + 1}→${slots[i] + 1}）`;
        warnings.push(msg);
        consecutiveWarnings.push(msg);
      }
    }
  }

  return {
    isValid: errors.length === 0,
    errors: errors,
    warnings: warnings,
    consecutiveWarnings: consecutiveWarnings,
    finalEndTime: finalEnd,
    timeOverflow: timeOverflow,
  };
}

/**
 * 2日間に分割してスケジュールを生成
 * 試合を半分ずつに分けて、各日のタイムラインを生成
 * スロットの境界で分割（同じスロットの試合は同じ日に配置）
 */
function generateTwoDaySchedule(input, schedule, regulations) {
  // スロットごとにグループ化
  const slotGroups = {};
  for (const match of schedule) {
    if (!slotGroups[match.slotIndex]) {
      slotGroups[match.slotIndex] = [];
    }
    slotGroups[match.slotIndex].push(match);
  }

  const slotIndices = Object.keys(slotGroups)
    .map(Number)
    .sort((a, b) => a - b);
  const totalSlots = slotIndices.length;
  const halfSlotPoint = Math.ceil(totalSlots / 2);

  // スロット単位で2日に分割
  const day1SlotIndices = slotIndices.slice(0, halfSlotPoint);
  const day2SlotIndices = slotIndices.slice(halfSlotPoint);

  const day1Matches = [];
  for (const slotIdx of day1SlotIndices) {
    day1Matches.push(...slotGroups[slotIdx]);
  }

  const day2Matches = [];
  for (const slotIdx of day2SlotIndices) {
    day2Matches.push(...slotGroups[slotIdx]);
  }

  // 1日目のスケジュールを再構築
  const day1Schedule = rebuildScheduleForDay(
    day1Matches,
    input,
    regulations,
    1,
  );

  // 2日目のスケジュールを再構築
  const day2Schedule = rebuildScheduleForDay(
    day2Matches,
    input,
    regulations,
    2,
  );

  // 1日目用の入力設定（表彰式・写真撮影なし）
  const day1Input = JSON.parse(JSON.stringify(input));
  day1Input.common.hasCeremony = false;
  day1Input.common.hasPhoto = false;
  day1Input.common.hasExhibition = false;

  // 各日のタイムラインを生成
  const day1Timeline = generateTimeline(day1Input, day1Schedule);
  const day2Timeline = generateTimeline(input, day2Schedule); // 2日目は通常通り

  return {
    day1: {
      schedule: day1Schedule,
      timeline: day1Timeline,
      matchCount: day1Matches.length,
    },
    day2: {
      schedule: day2Schedule,
      timeline: day2Timeline,
      matchCount: day2Matches.length,
    },
    totalMatches: schedule.length,
  };
}

/**
 * 指定された試合リストから1日分のスケジュールを再構築
 * 時間を0からリセットし、新しいスロットインデックスを割り当て
 */
function rebuildScheduleForDay(matches, input, regulations, dayNumber) {
  if (matches.length === 0) return [];

  const rebuiltSchedule = [];
  const startTime = new Date(input.common.startTime);
  startTime.setMinutes(
    startTime.getMinutes() +
      input.common.setupMinutes +
      input.common.receptionMinutes,
  );

  let currentTime = new Date(startTime);
  let newSlotIndex = 0;
  let prevOriginalSlotIndex = matches[0].slotIndex;

  for (let i = 0; i < matches.length; i++) {
    const match = matches[i];
    const slotMinutes =
      match.category === "U12"
        ? regulations.u12Regulation.slot_minutes
        : regulations.u15Regulation.slot_minutes;

    // 元のスロットが変わったら、新しいスロットに移行
    if (match.slotIndex !== prevOriginalSlotIndex) {
      // 前のスロットの試合時間分を進める
      const prevSlotMatches = matches.filter(
        (m, idx) => idx < i && m.slotIndex === prevOriginalSlotIndex,
      );
      if (prevSlotMatches.length > 0) {
        const maxPrevSlotMinutes = Math.max(
          ...prevSlotMatches.map((m) =>
            m.category === "U12"
              ? regulations.u12Regulation.slot_minutes
              : regulations.u15Regulation.slot_minutes,
          ),
        );
        currentTime.setMinutes(currentTime.getMinutes() + maxPrevSlotMinutes);
      }

      // インターバルを追加
      const interval =
        newSlotIndex === 0 && input.common.firstInterval15
          ? 15
          : input.common.intervalMinutes;
      currentTime.setMinutes(currentTime.getMinutes() + interval);

      newSlotIndex++;
      prevOriginalSlotIndex = match.slotIndex;
    }

    const matchStartTime = new Date(currentTime);
    const matchEndTime = new Date(currentTime);
    matchEndTime.setMinutes(matchEndTime.getMinutes() + slotMinutes);

    rebuiltSchedule.push({
      ...match,
      startTime: matchStartTime,
      endTime: matchEndTime,
      slotIndex: newSlotIndex,
      slotMinutes: slotMinutes,
      day: dayNumber,
    });
  }

  return rebuiltSchedule;
}

/**
 * 2日間のタイムラインを出力
 */
function outputTwoDayResults(ss, input, twoDayResult, regulations) {
  // タイムラインシートをクリアして2日間分を出力
  let timelineSheet = ss.getSheetByName(CONFIG.SHEETS.TIMELINE);
  if (!timelineSheet) {
    timelineSheet = ss.insertSheet(CONFIG.SHEETS.TIMELINE);
  }
  timelineSheet.clear();

  // ヘッダー
  timelineSheet
    .getRange("A1:D1")
    .setValues([["時刻", "イベント", "所要時間", "備考"]]);
  timelineSheet
    .getRange("A1:D1")
    .setBackground("#4a86e8")
    .setFontColor("white")
    .setFontWeight("bold");

  let currentRow = 2;

  // 1日目
  timelineSheet.getRange(`A${currentRow}`).setValue("【1日目】");
  timelineSheet
    .getRange(`A${currentRow}:D${currentRow}`)
    .setBackground("#e8f0fe")
    .setFontWeight("bold");
  currentRow++;

  currentRow = outputDayTimeline(
    timelineSheet,
    twoDayResult.day1.timeline,
    currentRow,
  );

  // 空行
  currentRow++;

  // 2日目
  timelineSheet.getRange(`A${currentRow}`).setValue("【2日目】");
  timelineSheet
    .getRange(`A${currentRow}:D${currentRow}`)
    .setBackground("#e8f0fe")
    .setFontWeight("bold");
  currentRow++;

  currentRow = outputDayTimeline(
    timelineSheet,
    twoDayResult.day2.timeline,
    currentRow,
  );

  // 列幅調整
  timelineSheet.setColumnWidth(1, 120);
  timelineSheet.setColumnWidth(2, 150);
  timelineSheet.setColumnWidth(3, 80);
  timelineSheet.setColumnWidth(4, 300);

  // 試合スケジュールシートも2日間分を出力
  outputTwoDayScheduleSheet(ss, input, twoDayResult, regulations);
}

/**
 * 1日分のタイムラインをシートに出力
 */
function outputDayTimeline(sheet, timeline, startRow) {
  let currentRow = startRow;

  for (const event of timeline) {
    const timeStr = formatTimeRange(event.startTime, event.endTime);
    const duration = `${event.duration}分`;

    let remarks = "";
    if (event.eventType === "match" && event.matches) {
      remarks = event.matches
        .map((m) => `[${m.court}] ${m.team1} vs ${m.team2}`)
        .join("\n");
    }

    sheet.getRange(currentRow, 1).setValue(timeStr);
    sheet.getRange(currentRow, 2).setValue(event.name);
    sheet.getRange(currentRow, 3).setValue(duration);
    sheet.getRange(currentRow, 4).setValue(remarks);

    if (event.eventType === "match") {
      sheet.getRange(`A${currentRow}:D${currentRow}`).setBackground("#fff2cc");
    }

    currentRow++;
  }

  return currentRow;
}

/**
 * 2日間の試合スケジュールシートを出力
 */
function outputTwoDayScheduleSheet(ss, input, twoDayResult, regulations) {
  let scheduleSheet = ss.getSheetByName(CONFIG.SHEETS.SCHEDULE);
  if (!scheduleSheet) {
    scheduleSheet = ss.insertSheet(CONFIG.SHEETS.SCHEDULE);
  }
  scheduleSheet.clear();

  // ヘッダー
  scheduleSheet
    .getRange("A1:G1")
    .setValues([["日", "試合#", "カテゴリ", "対戦", "コート", "開始", "終了"]]);
  scheduleSheet
    .getRange("A1:G1")
    .setBackground("#4a86e8")
    .setFontColor("white")
    .setFontWeight("bold");

  let row = 2;
  let matchNum = 1;

  // 1日目
  for (const match of twoDayResult.day1.schedule) {
    scheduleSheet.getRange(row, 1).setValue("1日目");
    scheduleSheet.getRange(row, 2).setValue(matchNum++);
    scheduleSheet.getRange(row, 3).setValue(match.category);
    scheduleSheet.getRange(row, 4).setValue(`${match.team1} vs ${match.team2}`);
    scheduleSheet.getRange(row, 5).setValue(match.court);
    scheduleSheet.getRange(row, 6).setValue(formatTime(match.startTime));
    scheduleSheet.getRange(row, 7).setValue(formatTime(match.endTime));
    row++;
  }

  // 2日目
  for (const match of twoDayResult.day2.schedule) {
    scheduleSheet.getRange(row, 1).setValue("2日目");
    scheduleSheet.getRange(row, 2).setValue(matchNum++);
    scheduleSheet.getRange(row, 3).setValue(match.category);
    scheduleSheet.getRange(row, 4).setValue(`${match.team1} vs ${match.team2}`);
    scheduleSheet.getRange(row, 5).setValue(match.court);
    scheduleSheet.getRange(row, 6).setValue(formatTime(match.startTime));
    scheduleSheet.getRange(row, 7).setValue(formatTime(match.endTime));
    row++;
  }

  // 列幅調整
  scheduleSheet.setColumnWidth(1, 60);
  scheduleSheet.setColumnWidth(2, 50);
  scheduleSheet.setColumnWidth(3, 60);
  scheduleSheet.setColumnWidth(4, 180);
  scheduleSheet.setColumnWidth(5, 70);
  scheduleSheet.setColumnWidth(6, 80);
  scheduleSheet.setColumnWidth(7, 80);
}

/**
 * 2日間開催のサマリーを生成
 */
function generateTwoDaySummary(input, twoDayResult, regulations) {
  const day1LastEvent =
    twoDayResult.day1.timeline[twoDayResult.day1.timeline.length - 1];
  const day2LastEvent =
    twoDayResult.day2.timeline[twoDayResult.day2.timeline.length - 1];

  const u12Reg = regulations.u12Regulation
    ? regulations.u12Regulation.label
    : "なし";
  const u15Reg = regulations.u15Regulation
    ? regulations.u15Regulation.label
    : "なし";

  return `【2日間開催】タイムライン生成完了！

■ 1日目（表彰式・写真撮影なし）
・試合数: ${twoDayResult.day1.matchCount}試合
・終了予定: ${formatTime(day1LastEvent.endTime)}

■ 2日目（表彰式・写真撮影あり）
・試合数: ${twoDayResult.day2.matchCount}試合
・終了予定: ${formatTime(day2LastEvent.endTime)}

■ 合計
・総試合数: ${twoDayResult.totalMatches}試合
・コート数: ${input.common.courts}面
・U12: ${u12Reg}
・U15: ${u15Reg}

タイムラインシートと試合スケジュールシートに
2日間分のスケジュールを出力しました。`;
}

function generateTimeError(input, regulations) {
  const availableTime = calculateAvailableTime(input);
  const u12Matches = input.u12.enabled
    ? calculateMatchCount(input.u12.teams, input.u12.matchType)
    : null;
  const u15Matches = input.u15.enabled
    ? calculateMatchCount(input.u15.teams, input.u15.matchType)
    : null;

  const u12MinReg = input.u12.enabled
    ? getRegulationsByCategory("U12")[0]
    : null;
  const u15MinReg = input.u15.enabled
    ? getRegulationsByCategory("U15")[0]
    : null;

  const u12Total = u12Matches ? u12Matches.total : 0;
  const u15Total = u15Matches ? u15Matches.total : 0;

  // スロット時間
  const u12SlotMin = u12MinReg ? u12MinReg.slot_minutes : 0;
  const u15SlotMin = u15MinReg ? u15MinReg.slot_minutes : 0;
  const maxSlotMinutes = Math.max(u12SlotMin, u15SlotMin);

  // トーナメント戦かどうかを確認
  const u12IsTournament =
    input.u12.enabled && input.u12.matchType === CONFIG.MATCH_TYPES.TOURNAMENT;
  const u15IsTournament =
    input.u15.enabled && input.u15.matchType === CONFIG.MATCH_TYPES.TOURNAMENT;

  let totalSlots;
  let requiredTime;

  if (u12IsTournament || u15IsTournament) {
    // トーナメント戦の場合
    if (input.common.courts >= 2 && u12Total > 0 && u15Total > 0) {
      // 両カテゴリ: 多い方の試合数
      totalSlots = Math.max(u12Total, u15Total);
    } else if (input.common.courts >= 2) {
      // 1カテゴリのみでも2コート使用
      totalSlots = Math.ceil((u12Total + u15Total) / input.common.courts);
    } else {
      totalSlots = u12Total + u15Total;
    }
    requiredTime = totalSlots * maxSlotMinutes;
  } else {
    // リーグ戦の場合
    if (input.common.courts >= 2 && u12Total > 0 && u15Total > 0) {
      const u12Slots = Math.ceil(u12Total / input.common.courts);
      const u15Slots = Math.ceil(u15Total / input.common.courts);
      totalSlots = u12Slots + u15Slots;
      requiredTime = u12Slots * u12SlotMin + u15Slots * u15SlotMin;
    } else {
      totalSlots = Math.ceil((u12Total + u15Total) / input.common.courts);
      requiredTime = totalSlots * maxSlotMinutes;
    }
  }

  // インターバル時間を追加
  if (totalSlots > 1) {
    const firstInterval = input.common.firstInterval15
      ? 15
      : input.common.intervalMinutes;
    requiredTime += firstInterval;
    if (totalSlots > 2) {
      requiredTime += (totalSlots - 2) * input.common.intervalMinutes;
    }
  }

  const shortage = requiredTime - availableTime;
  const shortageSlots = Math.ceil(shortage / maxSlotMinutes);

  return {
    availableTime: availableTime,
    requiredTime: requiredTime,
    shortageMinutes: shortage,
    shortageSlots: shortageSlots,
    message: `時間が不足しています。不足: ${shortage}分（約${shortageSlots}枠）`,
  };
}

// ============================================================
// TimelineGenerator.gs - タイムライン生成
// ============================================================

function generateTimeline(input, schedule) {
  const timeline = [];
  let currentTime = new Date(input.common.startTime);

  // 設営・開場（0分ならスキップ）
  if (input.common.setupMinutes > 0) {
    timeline.push(
      createTimelineEvent("設営・開場", currentTime, input.common.setupMinutes),
    );
    currentTime = addMinutes(currentTime, input.common.setupMinutes);
  }

  // 受付（0分ならスキップ）
  if (input.common.receptionMinutes > 0) {
    timeline.push(
      createTimelineEvent("受付", currentTime, input.common.receptionMinutes),
    );
    currentTime = addMinutes(currentTime, input.common.receptionMinutes);
  }

  const matchesBySlot = groupMatchesBySlot(schedule);
  let matchNumber = 1;

  for (const [slotIndex, matches] of Object.entries(matchesBySlot)) {
    const slotMatches = matches.sort((a, b) => a.court.localeCompare(b.court));
    const slotStartTime = slotMatches[0].startTime;
    const maxSlotMinutes = Math.max(...slotMatches.map((m) => m.slotMinutes));

    const matchDescriptions = slotMatches
      .map((m) => `[${m.court}コート] ${m.team1} vs ${m.team2}`)
      .join("\n");

    timeline.push({
      eventType: "match",
      name: `第${matchNumber}試合`,
      startTime: new Date(slotStartTime),
      endTime: addMinutes(slotStartTime, maxSlotMinutes),
      duration: maxSlotMinutes,
      description: matchDescriptions,
      matches: slotMatches,
    });

    matchNumber++;
  }

  if (schedule.length > 0) {
    const lastMatch = schedule[schedule.length - 1];
    currentTime = new Date(lastMatch.endTime);
  }

  // エキシビジョン（入力された時間を使用）
  if (input.common.hasExhibition) {
    timeline.push(
      createTimelineEvent(
        "エキシビジョン",
        currentTime,
        input.common.exhibitionMinutes,
      ),
    );
    currentTime = addMinutes(currentTime, input.common.exhibitionMinutes);
  }

  // 表彰式
  if (input.common.hasCeremony) {
    timeline.push(
      createTimelineEvent("表彰式", currentTime, input.common.ceremonyMinutes),
    );
    currentTime = addMinutes(currentTime, input.common.ceremonyMinutes);
  }

  // 写真撮影
  if (input.common.hasPhoto) {
    timeline.push(
      createTimelineEvent(
        "1〜3位 写真撮影",
        currentTime,
        input.common.photoMinutes,
      ),
    );
    currentTime = addMinutes(currentTime, input.common.photoMinutes);
  }

  // 完全撤退（0分ならスキップ）
  if (input.common.cleanupMinutes > 0) {
    timeline.push(
      createTimelineEvent("完全撤退", currentTime, input.common.cleanupMinutes),
    );
  }

  return timeline;
}

function createTimelineEvent(name, startTime, duration) {
  return {
    eventType: "general",
    name: name,
    startTime: new Date(startTime),
    endTime: addMinutes(startTime, duration),
    duration: duration,
    description: "",
  };
}

function addMinutes(date, minutes) {
  const result = new Date(date);
  result.setMinutes(result.getMinutes() + minutes);
  return result;
}

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

function outputTimelineToSheet(timeline, sheet) {
  const headers = ["時刻", "イベント", "所要時間", "備考"];
  const data = [headers];

  for (const event of timeline) {
    const timeStr = formatTimeRange(event.startTime, event.endTime);
    const durationStr = `${event.duration}分`;

    data.push([timeStr, event.name, durationStr, event.description || ""]);
  }

  sheet.clear();
  sheet.getRange(1, 1, data.length, headers.length).setValues(data);
  formatTimelineSheet(sheet, data.length);
}

function outputScheduleToSheet(schedule, sheet) {
  const headers = [
    "No.",
    "コート",
    "開始",
    "終了",
    "カテゴリ",
    "チーム1",
    "チーム2",
    "グループ/ラウンド",
  ];
  const data = [headers];

  // 時系列順に番号を振り直す
  let matchNo = 1;
  for (const match of schedule) {
    data.push([
      matchNo++,
      match.court,
      formatTime(match.startTime),
      formatTime(match.endTime),
      match.category,
      match.team1,
      match.team2,
      match.groupName || match.roundName || "",
    ]);
  }

  sheet.clear();
  sheet.getRange(1, 1, data.length, headers.length).setValues(data);
  formatScheduleSheet(sheet, data.length);
}

function formatTime(date) {
  const hours = String(date.getHours()).padStart(2, "0");
  const minutes = String(date.getMinutes()).padStart(2, "0");
  return `${hours}:${minutes}`;
}

function formatTimeRange(startTime, endTime) {
  return `${formatTime(startTime)} - ${formatTime(endTime)}`;
}

/**
 * チーム名からカテゴリプレフィックス（U12-, U15-）を除去
 */
function stripTeamPrefix(teamName) {
  if (!teamName) return teamName;
  return teamName.replace(/^U1[25]-/, "");
}

function formatTimelineSheet(sheet, rowCount) {
  const headerRange = sheet.getRange(1, 1, 1, 4);
  headerRange.setFontWeight("bold");
  headerRange.setBackground("#4285f4");
  headerRange.setFontColor("#ffffff");

  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 80);
  sheet.setColumnWidth(4, 300);

  if (rowCount > 1) {
    const dataRange = sheet.getRange(1, 1, rowCount, 4);
    dataRange.setBorder(true, true, true, true, true, true);
  }

  sheet.getRange(1, 1, rowCount, 3).setHorizontalAlignment("center");
  sheet.getRange(1, 4, rowCount, 1).setWrap(true);

  // 試合行に色を付ける（備考欄の内容からU12/U15を判定）
  for (let i = 2; i <= rowCount; i++) {
    const eventName = sheet.getRange(i, 2).getValue();
    const description = sheet.getRange(i, 4).getValue();
    const rowRange = sheet.getRange(i, 1, 1, 4);

    if (eventName.includes("試合")) {
      const hasU12 = description.includes("U12");
      const hasU15 = description.includes("U15");

      if (hasU12 && hasU15) {
        // 両方ある場合はグラデーション風（左半分青、右半分オレンジは難しいので薄紫）
        rowRange.setBackground("#e1bee7"); // 薄紫（両方の混合イメージ）
      } else if (hasU12) {
        rowRange.setBackground("#bbdefb"); // 明るい青
      } else if (hasU15) {
        rowRange.setBackground("#ffe0b2"); // 明るいオレンジ
      }
    }
  }
}

function formatScheduleSheet(sheet, rowCount) {
  const headerRange = sheet.getRange(1, 1, 1, 8);
  headerRange.setFontWeight("bold");
  headerRange.setBackground("#34a853");
  headerRange.setFontColor("#ffffff");

  sheet.setColumnWidth(1, 70);
  sheet.setColumnWidth(2, 60);
  sheet.setColumnWidth(3, 60);
  sheet.setColumnWidth(4, 60);
  sheet.setColumnWidth(5, 80);
  sheet.setColumnWidth(6, 100);
  sheet.setColumnWidth(7, 100);
  sheet.setColumnWidth(8, 150);

  if (rowCount > 1) {
    const dataRange = sheet.getRange(1, 1, rowCount, 8);
    dataRange.setBorder(true, true, true, true, true, true);
  }

  sheet.getRange(1, 1, rowCount, 8).setHorizontalAlignment("center");

  for (let i = 2; i <= rowCount; i++) {
    const category = sheet.getRange(i, 5).getValue();
    const rowRange = sheet.getRange(i, 1, 1, 8);
    if (category === "U12") {
      rowRange.setBackground("#bbdefb"); // 明るい青
    } else if (category === "U15") {
      rowRange.setBackground("#ffe0b2"); // 明るいオレンジ
    }
  }
}

function outputTournamentToSheet(matches, sheet, category) {
  const tournamentMatches = matches.filter(
    (m) => m.type === "tournament" && m.category === category,
  );

  if (tournamentMatches.length === 0) {
    return;
  }

  sheet.clear();

  sheet.getRange(1, 1).setValue(`${category} トーナメント表`);
  sheet.getRange(1, 1).setFontWeight("bold").setFontSize(14);

  // 3位決定戦を除外してラウンドグループを作成
  const roundGroups = {};
  const thirdPlace = tournamentMatches.find((m) => m.isThirdPlace);

  for (const match of tournamentMatches) {
    if (match.isThirdPlace) continue; // 3位決定戦は別枠
    if (!roundGroups[match.round]) {
      roundGroups[match.round] = [];
    }
    roundGroups[match.round].push(match);
  }

  let col = 1;
  let maxRow = 3;

  for (const [round, roundMatches] of Object.entries(roundGroups).sort(
    (a, b) => a[0] - b[0],
  )) {
    sheet.getRange(2, col).setValue(roundMatches[0].roundName);
    sheet.getRange(2, col).setFontWeight("bold");

    let row = 3;
    for (const match of roundMatches) {
      sheet.getRange(row, col).setValue(match.team1);
      sheet.getRange(row + 1, col).setValue("vs");
      sheet.getRange(row + 2, col).setValue(match.team2);
      row += 4;
    }
    maxRow = Math.max(maxRow, row);

    col += 2;
  }

  // 3位決定戦を表の下に別枠で表示
  if (thirdPlace) {
    const thirdPlaceRow = maxRow + 2;
    sheet.getRange(thirdPlaceRow, 1).setValue("【3位決定戦】");
    sheet.getRange(thirdPlaceRow, 1).setFontWeight("bold");
    sheet.getRange(thirdPlaceRow + 1, 1).setValue(thirdPlace.team1);
    sheet.getRange(thirdPlaceRow + 2, 1).setValue("vs");
    sheet.getRange(thirdPlaceRow + 3, 1).setValue(thirdPlace.team2);
  }
}

function generateSummary(input, schedule, timeline, regulations) {
  const u12MatchCount = schedule.filter((m) => m.category === "U12").length;
  const u15MatchCount = schedule.filter((m) => m.category === "U15").length;

  const lastEvent = timeline[timeline.length - 1];
  const endTime = lastEvent.endTime;

  const margin = input.common.endTime - endTime;
  const marginMinutes = Math.floor(margin / (1000 * 60));

  return {
    totalMatches: schedule.length,
    u12Matches: u12MatchCount,
    u15Matches: u15MatchCount,
    u12Regulation: regulations.u12Regulation
      ? regulations.u12Regulation.label
      : "なし",
    u15Regulation: regulations.u15Regulation
      ? regulations.u15Regulation.label
      : "なし",
    startTime: formatTime(input.common.startTime),
    endTime: formatTime(endTime),
    marginMinutes: marginMinutes,
    courts: input.common.courts,
  };
}

// ============================================================
// Main.gs - メイン処理
// ============================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("BSCA CUP")
    .addItem("タイムライン生成", "generateTimelineMain")
    .addItem("入力シートに色付け", "formatInputSheet")
    .addItem("入力をクリア", "clearInput")
    .addItem("出力をクリア", "clearOutput")
    .addToUi();
}

/**
 * 入力シートに色付け＆プルダウン設定
 */
function formatInputSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.INPUT);

  if (!sheet) {
    SpreadsheetApp.getUi().alert("入力シートが見つかりません");
    return;
  }

  // 全てのプルダウンをクリア
  sheet.getRange("C1:D31").clearDataValidations();
  sheet.getRange("D1:D31").clearContent();

  // 全体の背景をリセット
  sheet.getRange("A1:D31").setBackground(null);

  // === プルダウン設定 ===
  // コート数（1, 2）
  const courtsRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["1", "2"], true)
    .build();
  sheet.getRange("C5").setDataValidation(courtsRule);

  // ON/OFF共通
  const onOffRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["ON", "OFF"], true)
    .build();

  // 1〜2試合間15分（ON/OFF）
  sheet.getRange("C10").setDataValidation(onOffRule);

  // U12開催（ON/OFF）
  sheet.getRange("C18").setDataValidation(onOffRule);

  // U12チーム数（2-16）
  const teamsRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(
      [
        "2",
        "3",
        "4",
        "5",
        "6",
        "7",
        "8",
        "9",
        "10",
        "11",
        "12",
        "13",
        "14",
        "15",
        "16",
      ],
      true,
    )
    .build();
  sheet.getRange("C19").setDataValidation(teamsRule);

  // U12試合方式
  const matchRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["リーグ戦", "トーナメント"], true)
    .build();
  sheet.getRange("C20").setDataValidation(matchRule);

  // レギュレーション（共通：45分〜90分、5分刻み）
  const slotOptions = ["自動"];
  for (let m = 45; m <= 90; m += 5) {
    slotOptions.push(`${m}分`);
  }
  const regRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(slotOptions, true)
    .build();

  // U12レギュレーション
  sheet.getRange("C21").setDataValidation(regRule);

  // U15開催（ON/OFF）
  sheet.getRange("C24").setDataValidation(onOffRule);

  // U15チーム数（2-16）
  sheet.getRange("C25").setDataValidation(teamsRule);

  // U15試合方式
  sheet.getRange("C26").setDataValidation(matchRule);

  // U15レギュレーション
  sheet.getRange("C27").setDataValidation(regRule);

  // === 色付け ===
  // タイトル
  sheet.getRange("A1").setFontSize(14).setFontWeight("bold");

  // セクションヘッダー（行全体に色付け）
  sheet.getRange("A2:D2").setBackground("#c8e6c9").setFontWeight("bold"); // 大会共通（緑）
  sheet.getRange("A7:D7").setBackground("#fff9c4").setFontWeight("bold"); // イベント時間（黄）
  sheet.getRange("A17:D17").setBackground("#bbdefb").setFontWeight("bold"); // U12（青）
  sheet.getRange("A23:D23").setBackground("#ffccbc").setFontWeight("bold"); // U15（オレンジ）

  // 入力セル（C列のみ）
  sheet.getRange("C3:C5").setBackground("#e8f5e9"); // 大会共通（薄緑）
  sheet.getRange("C8:C15").setBackground("#fffde7"); // イベント時間（薄黄）
  sheet.getRange("C18:C21").setBackground("#e3f2fd"); // U12（薄青）
  sheet.getRange("C24:C27").setBackground("#fbe9e7"); // U15（薄オレンジ）

  // 中央揃え
  sheet.getRange("C3:C27").setHorizontalAlignment("center");
  sheet.getRange("B3:B27").setHorizontalAlignment("right");

  // 列幅調整
  sheet.setColumnWidth(1, 30);
  sheet.setColumnWidth(2, 140);
  sheet.setColumnWidth(3, 100);

  SpreadsheetApp.getUi().alert("色付け＆プルダウン設定が完了しました");
}

function generateTimelineMain() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  try {
    const input = readInputValues(ss);

    const validationResult = validateInput(input);
    if (!validationResult.isValid) {
      showError(validationResult.errors.join("\n"));
      return;
    }

    const regulations = selectOptimalRegulations(input);
    if (!regulations) {
      const error = generateTimeError(input, null);
      showError(error.message);
      return;
    }

    let schedule = generateSchedule(input, regulations);

    let scheduleValidation = validateSchedule(schedule, input);
    if (!scheduleValidation.isValid) {
      showError(scheduleValidation.errors.join("\n"));
      return;
    }

    // 連戦警告がある場合、連戦を回避できる最大コート数を探す
    if (
      scheduleValidation.consecutiveWarnings.length > 0 &&
      input.common.courts > 1
    ) {
      // 元のコート数から1コートまで順に試して、連戦回避できる最大コート数を探す
      let bestCourts = 0;
      let bestSchedule = null;
      let bestValidation = null;

      for (
        let testCourts = input.common.courts - 1;
        testCourts >= 1;
        testCourts--
      ) {
        const testInput = JSON.parse(JSON.stringify(input));
        testInput.common.courts = testCourts;
        const testSchedule = generateSchedule(testInput, regulations);
        const testValidation = validateSchedule(testSchedule, testInput);

        if (testValidation.consecutiveWarnings.length === 0) {
          bestCourts = testCourts;
          bestSchedule = testSchedule;
          bestValidation = testValidation;
          break;
        }
      }

      if (bestCourts > 0) {
        const response = ui.alert(
          "連戦回避の提案",
          scheduleValidation.warnings.join("\n") +
            "\n\n【提案】コート数を" +
            bestCourts +
            "にすれば連戦を回避できます。\n変更しますか？",
          ui.ButtonSet.YES_NO,
        );

        if (response === ui.Button.YES) {
          input.common.courts = bestCourts;
          schedule = bestSchedule;
          scheduleValidation = bestValidation;

          // 連戦回避後に時間超過をチェック
          if (scheduleValidation.timeOverflow) {
            const timeResponse = ui.alert(
              "時間超過 - 2日間開催の提案",
              `連戦は回避できましたが、利用終了時刻を${scheduleValidation.timeOverflow.overMinutes}分超過します。\n\n` +
                "【提案】2日間に分けて開催しますか？\n" +
                "試合を半分ずつに分けてタイムラインを生成します。",
              ui.ButtonSet.YES_NO,
            );

            if (timeResponse === ui.Button.YES) {
              const twoDayResult = generateTwoDaySchedule(
                input,
                schedule,
                regulations,
              );
              outputTwoDayResults(ss, input, twoDayResult, regulations);
              const summary = generateTwoDaySummary(
                input,
                twoDayResult,
                regulations,
              );
              showSuccess(summary);
              return;
            }
          }
        }
      } else {
        ui.alert(
          "警告",
          scheduleValidation.warnings.join("\n") +
            "\n\n※このチーム構成では連戦回避が困難です",
          ui.ButtonSet.OK,
        );
      }
    } else if (scheduleValidation.timeOverflow) {
      // 時間超過の場合、2日間分割を提案
      const response = ui.alert(
        "時間超過 - 2日間開催の提案",
        scheduleValidation.warnings.join("\n") +
          "\n\n【提案】2日間に分けて開催しますか？\n" +
          "試合を半分ずつに分けてタイムラインを生成します。",
        ui.ButtonSet.YES_NO,
      );

      if (response === ui.Button.YES) {
        // 2日間分割でスケジュール再生成
        const twoDayResult = generateTwoDaySchedule(
          input,
          schedule,
          regulations,
        );
        outputTwoDayResults(ss, input, twoDayResult, regulations);
        const summary = generateTwoDaySummary(input, twoDayResult, regulations);
        showSuccess(summary);
        return;
      }
    } else if (scheduleValidation.warnings.length > 0) {
      // その他の警告
      ui.alert("警告", scheduleValidation.warnings.join("\n"), ui.ButtonSet.OK);
    }

    const timeline = generateTimeline(input, schedule);

    outputResults(ss, input, schedule, timeline, regulations);

    const summary = generateSummary(input, schedule, timeline, regulations);
    showSuccess(summary);
  } catch (e) {
    showError("エラーが発生しました: " + e.message);
    Logger.log(e);
  }
}

function readInputValues(ss) {
  const inputSheet = ss.getSheetByName(CONFIG.SHEETS.INPUT);

  if (!inputSheet) {
    throw new Error("入力シートが見つかりません");
  }

  const input = createInputStructure();

  // 大会共通
  input.common.startTime = parseTime(
    inputSheet.getRange(INPUT_CELLS.START_TIME).getValue(),
  );
  input.common.endTime = parseTime(
    inputSheet.getRange(INPUT_CELLS.END_TIME).getValue(),
  );
  input.common.courts = inputSheet.getRange(INPUT_CELLS.COURTS).getValue() || 1;

  // イベント時間（0ならスキップ）
  input.common.setupMinutes =
    parseInt(inputSheet.getRange(INPUT_CELLS.SETUP_MINUTES).getValue()) || 0;
  input.common.receptionMinutes =
    parseInt(inputSheet.getRange(INPUT_CELLS.RECEPTION_MINUTES).getValue()) ||
    0;

  // 1〜2試合間15分設定
  const firstInterval15 = inputSheet
    .getRange(INPUT_CELLS.FIRST_INTERVAL_15)
    .getValue();
  input.common.firstInterval15 =
    firstInterval15 === true || firstInterval15 === "ON";

  // その他試合間インターバル（デフォルト10分）
  input.common.intervalMinutes =
    parseInt(inputSheet.getRange(INPUT_CELLS.INTERVAL_MINUTES).getValue()) ||
    10;

  input.common.exhibitionMinutes =
    parseInt(inputSheet.getRange(INPUT_CELLS.EXHIBITION_MINUTES).getValue()) ||
    0;
  input.common.ceremonyMinutes =
    parseInt(inputSheet.getRange(INPUT_CELLS.CEREMONY_MINUTES).getValue()) || 0;
  input.common.photoMinutes =
    parseInt(inputSheet.getRange(INPUT_CELLS.PHOTO_MINUTES).getValue()) || 0;
  input.common.cleanupMinutes =
    parseInt(inputSheet.getRange(INPUT_CELLS.CLEANUP_MINUTES).getValue()) || 0;

  // 0より大きければ有効
  input.common.hasExhibition = input.common.exhibitionMinutes > 0;
  input.common.hasCeremony = input.common.ceremonyMinutes > 0;
  input.common.hasPhoto = input.common.photoMinutes > 0;

  // U12
  const u12Enabled = inputSheet.getRange(INPUT_CELLS.U12_ENABLED).getValue();
  input.u12.enabled = u12Enabled === true || u12Enabled === "ON";
  input.u12.teams = inputSheet.getRange(INPUT_CELLS.U12_TEAMS).getValue() || 0;
  input.u12.matchType =
    inputSheet.getRange(INPUT_CELLS.U12_MATCH_TYPE).getValue() ||
    CONFIG.MATCH_TYPES.LEAGUE;
  input.u12.regulation =
    inputSheet.getRange(INPUT_CELLS.U12_REGULATION).getValue() || "自動";

  // U15
  const u15Enabled = inputSheet.getRange(INPUT_CELLS.U15_ENABLED).getValue();
  input.u15.enabled = u15Enabled === true || u15Enabled === "ON";
  input.u15.teams = inputSheet.getRange(INPUT_CELLS.U15_TEAMS).getValue() || 0;
  input.u15.matchType =
    inputSheet.getRange(INPUT_CELLS.U15_MATCH_TYPE).getValue() ||
    CONFIG.MATCH_TYPES.LEAGUE;
  input.u15.regulation =
    inputSheet.getRange(INPUT_CELLS.U15_REGULATION).getValue() || "自動";

  return input;
}

function parseTime(value) {
  if (value instanceof Date) {
    return value;
  }

  if (typeof value === "string") {
    const parts = value.split(":");
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

function validateInput(input) {
  const errors = [];

  if (!input.common.startTime) {
    errors.push("利用開始時刻を入力してください");
  }
  if (!input.common.endTime) {
    errors.push("利用終了時刻を入力してください");
  }
  if (
    input.common.startTime &&
    input.common.endTime &&
    input.common.startTime >= input.common.endTime
  ) {
    errors.push("利用終了時刻は開始時刻より後にしてください");
  }

  if (input.common.courts < 1 || input.common.courts > 2) {
    errors.push("コート数は1または2を指定してください");
  }

  if (!input.u12.enabled && !input.u15.enabled) {
    errors.push("U12またはU15のいずれかを有効にしてください");
  }

  if (input.u12.enabled) {
    if (input.u12.teams < 2) {
      errors.push("U12のチーム数は2以上を指定してください");
    }
    if (!input.u12.matchType) {
      errors.push("U12の試合方式を選択してください");
    }
  }

  if (input.u15.enabled) {
    if (input.u15.teams < 2) {
      errors.push("U15のチーム数は2以上を指定してください");
    }
    if (!input.u15.matchType) {
      errors.push("U15の試合方式を選択してください");
    }
  }

  return {
    isValid: errors.length === 0,
    errors: errors,
  };
}

function outputResults(ss, input, schedule, timeline, regulations) {
  let timelineSheet = ss.getSheetByName(CONFIG.SHEETS.TIMELINE);
  if (!timelineSheet) {
    timelineSheet = ss.insertSheet(CONFIG.SHEETS.TIMELINE);
  }
  outputTimelineToSheet(timeline, timelineSheet);

  let scheduleSheet = ss.getSheetByName(CONFIG.SHEETS.SCHEDULE);
  if (!scheduleSheet) {
    scheduleSheet = ss.insertSheet(CONFIG.SHEETS.SCHEDULE);
  }
  outputScheduleToSheet(schedule, scheduleSheet);

  if (
    input.u12.enabled &&
    input.u12.matchType === CONFIG.MATCH_TYPES.TOURNAMENT
  ) {
    let tournamentSheet = ss.getSheetByName("U12_" + CONFIG.SHEETS.TOURNAMENT);
    if (!tournamentSheet) {
      tournamentSheet = ss.insertSheet("U12_" + CONFIG.SHEETS.TOURNAMENT);
    }
    const u12Matches = generateMatchList(
      input.u12.teams,
      input.u12.matchType,
      "U12",
      "U12-",
    );
    outputTournamentToSheet(u12Matches, tournamentSheet, "U12");
  }

  if (
    input.u15.enabled &&
    input.u15.matchType === CONFIG.MATCH_TYPES.TOURNAMENT
  ) {
    let tournamentSheet = ss.getSheetByName("U15_" + CONFIG.SHEETS.TOURNAMENT);
    if (!tournamentSheet) {
      tournamentSheet = ss.insertSheet("U15_" + CONFIG.SHEETS.TOURNAMENT);
    }
    const u15Matches = generateMatchList(
      input.u15.teams,
      input.u15.matchType,
      "U15",
      "U15-",
    );
    outputTournamentToSheet(u15Matches, tournamentSheet, "U15");
  }
}

function showError(message) {
  const ui = SpreadsheetApp.getUi();
  ui.alert("エラー", message, ui.ButtonSet.OK);
}

function showSuccess(summary) {
  const ui = SpreadsheetApp.getUi();

  // 文字列の場合はそのまま表示
  if (typeof summary === "string") {
    ui.alert("完了", summary, ui.ButtonSet.OK);
    return;
  }

  // オブジェクトの場合はフォーマットして表示
  const message = `タイムライン生成が完了しました！

【サマリー】
・総試合数: ${summary.totalMatches}試合
  - U12: ${summary.u12Matches}試合 (${summary.u12Regulation})
  - U15: ${summary.u15Matches}試合 (${summary.u15Regulation})
・コート数: ${summary.courts}面
・開始: ${summary.startTime}
・終了予定: ${summary.endTime}
・終了時刻までの余裕: ${summary.marginMinutes}分`;

  ui.alert("完了", message, ui.ButtonSet.OK);
}

function clearInput() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName(CONFIG.SHEETS.INPUT);

  if (inputSheet) {
    const cellsToClear = [
      INPUT_CELLS.START_TIME,
      INPUT_CELLS.END_TIME,
      INPUT_CELLS.DAYS,
      INPUT_CELLS.COURTS,
      INPUT_CELLS.HAS_EXHIBITION,
      INPUT_CELLS.HAS_CEREMONY,
      INPUT_CELLS.HAS_PHOTO,
      INPUT_CELLS.U12_ENABLED,
      INPUT_CELLS.U12_TEAMS,
      INPUT_CELLS.U12_MATCH_TYPE,
      INPUT_CELLS.U15_ENABLED,
      INPUT_CELLS.U15_TEAMS,
      INPUT_CELLS.U15_MATCH_TYPE,
    ];

    for (const cell of cellsToClear) {
      inputSheet.getRange(cell).clearContent();
    }
  }

  SpreadsheetApp.getUi().alert("入力をクリアしました");
}

function clearOutput() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheetsToClear = [
    CONFIG.SHEETS.TIMELINE,
    CONFIG.SHEETS.SCHEDULE,
    "U12_" + CONFIG.SHEETS.TOURNAMENT,
    "U15_" + CONFIG.SHEETS.TOURNAMENT,
  ];

  for (const sheetName of sheetsToClear) {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      sheet.clear();
    }
  }

  SpreadsheetApp.getUi().alert("出力をクリアしました");
}

// ============================================================
// Setup.gs - 初期セットアップ（最適化版）
// ============================================================

function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  createInputSheetFast(ss);
  createOutputSheetsFast(ss);

  // メニュー作成（UIコンテキストがある場合のみ）
  try {
    onOpen();
    SpreadsheetApp.getUi().alert(
      "セットアップが完了しました！\n「入力」シートに必要事項を入力し、メニューから「タイムライン生成」を実行してください。",
    );
  } catch (e) {
    // スクリプトエディタから実行した場合はログに出力
    Logger.log(
      "セットアップが完了しました。スプレッドシートを開き直すとメニューが表示されます。",
    );
  }
}

function createInputSheetFast(ss) {
  let inputSheet = ss.getSheetByName(CONFIG.SHEETS.INPUT);

  if (!inputSheet) {
    inputSheet = ss.insertSheet(CONFIG.SHEETS.INPUT, 0);
  } else {
    inputSheet.clear();
  }

  // 一括でデータを設定（新レイアウト）
  const data = [
    ["BSCA CUP タイムライン生成ツール", "", "", ""],
    ["【大会共通設定】", "", "", ""],
    ["", "利用開始時刻", "8:00", ""],
    ["", "利用終了時刻", "18:00", ""],
    ["", "コート数", "2", ""],
    ["", "", "", ""],
    ["【イベント時間設定（分）】", "", "※0でスキップ", ""],
    ["", "設営・開場", "30", ""],
    ["", "受付", "15", ""],
    ["", "1〜2試合間15分", "OFF", "※ONで15分"],
    ["", "その他試合間", "10", "※デフォルト10分"],
    ["", "エキシビジョン", "0", ""],
    ["", "表彰式", "15", ""],
    ["", "写真撮影", "10", ""],
    ["", "完全撤退", "30", ""],
    ["", "", "", ""],
    ["【U12 設定】", "", "", ""],
    ["", "開催", "ON", ""],
    ["", "チーム数", "5", ""],
    ["", "試合方式", "リーグ戦", ""],
    ["", "1試合枠", "自動", ""],
    ["", "", "", ""],
    ["【U15 設定】", "", "", ""],
    ["", "開催", "OFF", ""],
    ["", "チーム数", "", ""],
    ["", "試合方式", "", ""],
    ["", "1試合枠", "自動", ""],
  ];

  inputSheet.getRange(1, 1, data.length, 4).setValues(data);
  inputSheet.setColumnWidths(1, 4, 120);
  inputSheet.setColumnWidth(1, 30);
}

function createOutputSheetsFast(ss) {
  const sheetNames = [CONFIG.SHEETS.TIMELINE, CONFIG.SHEETS.SCHEDULE];

  for (const name of sheetNames) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
    }
    sheet.clear();
    sheet.getRange("A1").setValue("タイムライン生成後に表示されます");
  }
}

function inputSampleData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName(CONFIG.SHEETS.INPUT);

  if (!inputSheet) {
    SpreadsheetApp.getUi().alert("入力シートが見つかりません");
    return;
  }

  // 大会共通
  inputSheet.getRange("C3").setValue("8:00");
  inputSheet.getRange("C4").setValue("18:00");
  inputSheet.getRange("C5").setValue("2");
  // イベント時間（0でスキップ）
  inputSheet.getRange("C8").setValue("30"); // 設営・開場
  inputSheet.getRange("C9").setValue("15"); // 受付
  inputSheet.getRange("C10").setValue("OFF"); // 1〜2試合間15分
  inputSheet.getRange("C11").setValue("10"); // その他試合間インターバル
  inputSheet.getRange("C12").setValue("0"); // エキシビジョン（0=なし）
  inputSheet.getRange("C13").setValue("15"); // 表彰式
  inputSheet.getRange("C14").setValue("10"); // 写真撮影
  inputSheet.getRange("C15").setValue("30"); // 完全撤退
  // U12
  inputSheet.getRange("C18").setValue("ON");
  inputSheet.getRange("C19").setValue("5");
  inputSheet.getRange("C20").setValue("リーグ戦");
  inputSheet.getRange("C21").setValue("自動");
  // U15
  inputSheet.getRange("C24").setValue("OFF");
  inputSheet.getRange("C25").setValue("");
  inputSheet.getRange("C26").setValue("");
  inputSheet.getRange("C27").setValue("自動");

  SpreadsheetApp.getUi().alert("サンプルデータを入力しました");
}
