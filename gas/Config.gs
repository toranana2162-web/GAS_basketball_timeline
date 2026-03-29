/**
 * BSCA CUP タイムライン自動生成ツール - 設定ファイル
 */

const CONFIG = {
  // シート名
  SHEETS: {
    INPUT: '入力',
    TIMELINE: 'タイムライン',
    SCHEDULE: '試合スケジュール',
    TOURNAMENT: 'トーナメント表'
  },
  
  // タイムラインイベント所要時間（分）
  TIMELINE_EVENTS: {
    SETUP: 30,           // 設営
    RECEPTION: 15,       // 開場・受付
    EXHIBITION: 30,      // エキシビジョン
    CEREMONY: 15,        // 表彰式
    PHOTO: 10,           // 写真撮影
    CLEANUP: 30          // 完全撤退
  },
  
  // カテゴリ
  CATEGORIES: {
    U12: 'U12',
    U15: 'U15'
  },
  
  // 試合方式
  MATCH_TYPES: {
    LEAGUE: 'リーグ戦',
    TOURNAMENT: 'トーナメント'
  },
  
  // グループ分けの閾値
  GROUP_THRESHOLDS: {
    TWO_GROUPS: 6,   // 6チーム以上で2グループ
    THREE_GROUPS: 9  // 9チーム以上で3グループ
  }
};

/**
 * レギュレーション定義
 * slot_minutes: 1試合枠の時間（分）
 */
const REGULATIONS = {
  U12_SLOT60: { category: 'U12', slot_minutes: 60, label: 'U12 - 60分枠' },
  U12_SLOT65: { category: 'U12', slot_minutes: 65, label: 'U12 - 65分枠' },
  U12_SLOT70: { category: 'U12', slot_minutes: 70, label: 'U12 - 70分枠' },
  U15_SLOT75: { category: 'U15', slot_minutes: 75, label: 'U15 - 75分枠' },
  U15_SLOT80: { category: 'U15', slot_minutes: 80, label: 'U15 - 80分枠' },
  U15_SLOT85: { category: 'U15', slot_minutes: 85, label: 'U15 - 85分枠' }
};

/**
 * カテゴリ別のレギュレーション候補を取得
 */
function getRegulationsByCategory(category) {
  const result = [];
  for (const [id, reg] of Object.entries(REGULATIONS)) {
    if (reg.category === category) {
      result.push({ id: id, ...reg });
    }
  }
  return result.sort((a, b) => a.slot_minutes - b.slot_minutes);
}

/**
 * レギュレーションIDからslot_minutesを取得
 */
function getSlotMinutes(regulationId) {
  const reg = REGULATIONS[regulationId];
  return reg ? reg.slot_minutes : null;
}

/**
 * 入力値の構造
 */
function createInputStructure() {
  return {
    // 大会共通
    common: {
      startTime: null,      // Date: 利用開始時刻
      endTime: null,        // Date: 利用終了時刻
      days: 1,              // Number: 日数 (1 or 2)
      courts: 1,            // Number: コート数 (1 or 2)
      hasExhibition: false, // Boolean: エキシビジョン有無
      hasCeremony: true,    // Boolean: 表彰式有無
      hasPhoto: true        // Boolean: 写真撮影有無
    },
    // U12設定
    u12: {
      enabled: false,       // Boolean: 開催するか
      teams: 0,             // Number: チーム数
      matchType: null,      // String: 試合方式
      regulation: null      // String: レギュレーションID (自動選択)
    },
    // U15設定
    u15: {
      enabled: false,       // Boolean: 開催するか
      teams: 0,             // Number: チーム数
      matchType: null,      // String: 試合方式
      regulation: null      // String: レギュレーションID (自動選択)
    }
  };
}

/**
 * 入力シートのセル位置定義
 */
const INPUT_CELLS = {
  // 大会共通
  START_TIME: 'C3',
  END_TIME: 'C4',
  DAYS: 'C5',
  COURTS: 'C6',
  HAS_EXHIBITION: 'C7',
  HAS_CEREMONY: 'C8',
  HAS_PHOTO: 'C9',
  
  // U12
  U12_ENABLED: 'C12',
  U12_TEAMS: 'C13',
  U12_MATCH_TYPE: 'C14',
  
  // U15
  U15_ENABLED: 'C17',
  U15_TEAMS: 'C18',
  U15_MATCH_TYPE: 'C19'
};

/**
 * チーム名のプリセット（仮名）
 */
const TEAM_NAMES = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P'];

/**
 * コート名
 */
const COURT_NAMES = ['A', 'B'];
