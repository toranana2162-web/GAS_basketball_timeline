/**
 * BSCA CUP タイムライン自動生成ツール - 試合数計算ロジック
 */

/**
 * 組み合わせ数 nCr を計算
 */
function combination(n, r) {
  if (r > n || r < 0) return 0;
  if (r === 0 || r === n) return 1;

  let result = 1;
  for (let i = 0; i < r; i++) {
    result = (result * (n - i)) / (i + 1);
  }
  return Math.round(result);
}

/**
 * 試合方式に応じた試合数を計算
 * @param {number} teams - チーム数
 * @param {string} matchType - 試合方式
 * @returns {Object} - 試合数情報
 */
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

/**
 * トーナメント方式の試合数を計算
 */
function calculateTournamentMatches(teams) {
  return {
    total: teams - 1,
    groups: 0,
    groupMatches: 0,
    finalMatches: teams - 1,
    structure: "tournament",
  };
}

/**
 * リーグ戦方式の試合数を計算
 */
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
 * 2グループ制リーグ戦の計算（グループ総当たりのみ）
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
 * 3グループ制リーグ戦の計算（グループ総当たりのみ）
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

/**
 * 全試合リストを生成
 * @param {number} teams - チーム数
 * @param {string} matchType - 試合方式
 * @param {string} category - カテゴリ (U12/U15)
 * @param {string} prefix - チーム名プレフィックス
 * @returns {Array} - 試合リスト
 */
function generateMatchList(teams, matchType, category, prefix) {
  const info = calculateMatchCount(teams, matchType);

  if (matchType === CONFIG.MATCH_TYPES.TOURNAMENT) {
    return generateTournamentMatches(teams, category, prefix);
  } else {
    return generateLeagueMatches(teams, category, prefix, info);
  }
}

/**
 * トーナメント試合リスト生成
 */
function generateTournamentMatches(teams, category, prefix) {
  const matches = [];
  const rounds = Math.ceil(Math.log2(teams));

  let matchNum = 1;
  let currentRoundTeams = teams;

  for (let round = 1; round <= rounds; round++) {
    const matchesInRound = Math.ceil(currentRoundTeams / 2);
    const isFinal = round === rounds;
    const isThirdPlace = false;

    for (let i = 0; i < matchesInRound; i++) {
      const roundName = getRoundName(round, rounds);
      matches.push({
        matchNumber: matchNum++,
        category: category,
        round: round,
        roundName: roundName,
        team1:
          round === 1 ? `${prefix}${TEAM_NAMES[i * 2]}` : `${roundName}勝者`,
        team2:
          round === 1
            ? `${prefix}${TEAM_NAMES[i * 2 + 1] || "BYE"}`
            : `${roundName}勝者`,
        type: "tournament",
        isFinal: isFinal,
      });
    }

    currentRoundTeams = Math.ceil(currentRoundTeams / 2);
  }

  // 3位決定戦を追加（4チーム以上の場合）
  if (teams >= 4) {
    matches.push({
      matchNumber: matchNum++,
      category: category,
      round: rounds,
      roundName: "3位決定戦",
      team1: "準決勝敗者1",
      team2: "準決勝敗者2",
      type: "tournament",
      isFinal: false,
      isThirdPlace: true,
    });
  }

  return matches;
}

/**
 * ラウンド名を取得
 */
function getRoundName(round, totalRounds) {
  const remaining = totalRounds - round;
  if (remaining === 0) return "決勝";
  if (remaining === 1) return "準決勝";
  if (remaining === 2) return "準々決勝";
  return `${round}回戦`;
}

/**
 * リーグ戦試合リスト生成
 */
function generateLeagueMatches(teams, category, prefix, info) {
  const matches = [];
  let matchNum = 1;

  if (info.structure === "single_league") {
    const teamList = [];
    for (let i = 0; i < teams; i++) {
      teamList.push(`${prefix}${TEAM_NAMES[i]}`);
    }

    // 総当たり戦の組み合わせを生成
    for (let i = 0; i < teams; i++) {
      for (let j = i + 1; j < teams; j++) {
        matches.push({
          matchNumber: matchNum++,
          category: category,
          group: 1,
          groupName: "",
          team1: teamList[i],
          team2: teamList[j],
          type: "league",
          phase: "group",
        });
      }
    }
  } else if (info.structure === "two_group_league") {
    matches.push(
      ...generateGroupMatches(
        info.teamsPerGroup[0],
        category,
        prefix,
        "A",
        matchNum,
      ),
    );
    matchNum = matches.length + 1;

    matches.push(
      ...generateGroupMatchesWithOffset(
        info.teamsPerGroup[1],
        category,
        prefix,
        "B",
        matchNum,
        info.teamsPerGroup[0],
      ),
    );
  } else if (info.structure === "three_group_league") {
    let offset = 0;
    const groupLabels = ["A", "B", "C"];

    for (let g = 0; g < 3; g++) {
      matches.push(
        ...generateGroupMatchesWithOffset(
          info.teamsPerGroup[g],
          category,
          prefix,
          groupLabels[g],
          matchNum,
          offset,
        ),
      );
      matchNum = matches.length + 1;
      offset += info.teamsPerGroup[g];
    }
  }

  return matches;
}

/**
 * グループ内総当たり戦を生成
 */
function generateGroupMatches(
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

  let matchNum = startNum;
  for (let i = 0; i < teamCount; i++) {
    for (let j = i + 1; j < teamCount; j++) {
      matches.push({
        matchNumber: matchNum++,
        category: category,
        group: groupLabel,
        groupName: `グループ${groupLabel}`,
        team1: teamList[i],
        team2: teamList[j],
        type: "league",
        phase: "group",
      });
    }
  }

  return matches;
}

/**
 * オフセット付きグループ内総当たり戦を生成
 */
function generateGroupMatchesWithOffset(
  teamCount,
  category,
  prefix,
  groupLabel,
  startNum,
  offset,
) {
  const matches = [];
  const teamList = [];

  for (let i = 0; i < teamCount; i++) {
    teamList.push(`${prefix}${groupLabel}${i + 1}`);
  }

  let matchNum = startNum;
  for (let i = 0; i < teamCount; i++) {
    for (let j = i + 1; j < teamCount; j++) {
      matches.push({
        matchNumber: matchNum++,
        category: category,
        group: groupLabel,
        groupName: `グループ${groupLabel}`,
        team1: teamList[i],
        team2: teamList[j],
        type: "league",
        phase: "group",
      });
    }
  }

  return matches;
}

/**
 * 必要な総時間を計算（分）
 */
function calculateTotalTimeRequired(matchInfo, slotMinutes, courts) {
  const totalSlots = Math.ceil(matchInfo.total / courts);
  return totalSlots * slotMinutes;
}

/**
 * カテゴリごとの試合情報をまとめて計算
 */
function calculateCategoryMatchInfo(categoryConfig, category) {
  if (!categoryConfig.enabled || categoryConfig.teams < 2) {
    return null;
  }

  const matchInfo = calculateMatchCount(
    categoryConfig.teams,
    categoryConfig.matchType,
  );

  return {
    category: category,
    teams: categoryConfig.teams,
    matchType: categoryConfig.matchType,
    ...matchInfo,
  };
}
