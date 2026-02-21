// Constants for ELO calculation
const K = 32;
const BASE_RATING = 1000;

// Type Aliases for Excel Data
type ExcelCellValue = string | number | boolean;
type ExcelData = ExcelCellValue[][];

// Helper function interface for ELO stats
interface PlayerStats {
  rating: number;
  nationality: string;
  team: string;
  role: string;
  n_game_played: number;
  wins: number;
  standardRating?: number; // ELO before penalization, used for calculating pointsLost
}

// --- Main Orchestration Function ---
/**
 * Main entry point for the Excel Script.
 * Orchestrates the calculation of ELO ratings, updates the leaderboard,
 * and refreshes all winrate matrices, including rank change calculation.
 */
function main(workbook: ExcelScript.Workbook) {
  // 1. Get necessary sheet references
  const matches = workbook.getWorksheet("Matches Log");
  const ratings = workbook.getWorksheet("Leaderboard");
  const playersDataSheet = workbook.getWorksheet("Challengers cards");
  const definition = workbook.getWorksheet("Definition");

  if (!matches || !ratings || !playersDataSheet || !definition) {
    console.log("One or more required sheets are missing.");
    return;
  }

  // 2. Run ELO calculation and get results, including previous ranks
  const { allRatings, ratingsMap, maxGamesPlayed, columnIndices, previousRanksMap } = setupAndCalculateElo(
    matches, ratings, playersDataSheet, K, BASE_RATING
  );

  // 3. Update the main leaderboard table with rank change
  updateLeaderboardTable(ratings, allRatings, maxGamesPlayed, columnIndices, previousRanksMap);

  // 4. Update the various podium views on the leaderboard
  updatePodiumViews(ratings, allRatings, ratingsMap);

  // 5. Update the country, team, and role winrate matrices
  updateAllWinratesMatrices(workbook);
}

// --- ELO Calculation Helper ---
/**
 * Runs ELO calculations on the match data up to a specified date limit.
 * This is used to calculate both the final ELO and the ELO snapshot for previous ranks.
 * @param limitDateKey YYYYMMDD string to limit matches to.
 */
function runEloCalculations(
  matchesSheet: ExcelScript.Worksheet,
  matchesData: ExcelData, // Updated type
  playersData: ExcelData, // Updated type
  limitDateKey: string | null,
  K: number,
  baseRating: number,
): { ratings: { [player: string]: PlayerStats }, maxGames: number } {

  // Helper function to convert ExcelCellValue (usually a number) to a sortable YYYYMMDD string
  const dateToKey = (d: ExcelCellValue): string => {
    if (typeof d === 'number' && d > 0) {
      // Convert Excel serial date to JS Date object (25569 is the difference in days)
      const date = new Date(Math.round((d - 25569) * 86400000));

      // Format to YYYYMMDD string for comparison (using UTC to prevent timezone shift issues)
      const year = date.getUTCFullYear().toString();
      const month = (date.getUTCMonth() + 1).toString().padStart(2, '0');
      const day = date.getUTCDate().toString().padStart(2, '0');

      return year + month + day;
    }
    return '00000000';
  };

  // 1. Initialize ratings map from players data
  const currentRatings: { [name: string]: PlayerStats } = {};
  for (let i = 1; i < playersData.length; i++) {
    const player = playersData[i][0] as string;
    if (!player) continue;
    currentRatings[player] = {
      rating: baseRating,
      nationality: playersData[i][1] as string,
      team: playersData[i][2] as string,
      role: playersData[i][3] as string,
      n_game_played: 0,
      wins: 0
    };
  }

  // 2. Process matches
  for (let i = 1; i < matchesData.length; i++) {
    const row = matchesData[i];
    const matchDateValue = row[0] as ExcelCellValue; // Read the raw cell value (expected number)

    // Skip match if it's past the limit date
    if (limitDateKey && dateToKey(matchDateValue) > limitDateKey) {
      continue;
    }

    const playerA = row[1] as string;
    const playerB = row[2] as string;
    const winner = row[5] as string;
    const ptsA = Number(row[3]);
    const ptsB = Number(row[4]);
    const margin = Math.abs(ptsA - ptsB);

    if (!playerA || !playerB || isNaN(ptsA) || isNaN(ptsB)) continue;

    // Ensure players exist (initialize if they played but weren't in cards)
    if (!currentRatings[playerA]) currentRatings[playerA] = { rating: baseRating, nationality: "Unknown", team: "Unknown", role: "Unknown", n_game_played: 0, wins: 0 };
    if (!currentRatings[playerB]) currentRatings[playerB] = { rating: baseRating, nationality: "Unknown", team: "Unknown", role: "Unknown", n_game_played: 0, wins: 0 };

    // Track wins
    if (winner && currentRatings[winner]) {
      currentRatings[winner].wins += 1;
    }

    const RA = currentRatings[playerA].rating;
    const RB = currentRatings[playerB].rating;

    const EA = 1 / (1 + Math.pow(10, (RB - RA) / 400));
    const SA = winner === playerA ? 1 : 0;
    const fMargin = 1 + 0.5 * (Math.log(1 + margin) / Math.log(2));
    const deltaA = K * fMargin * (SA - EA);

    currentRatings[playerA].rating += deltaA;
    currentRatings[playerB].rating -= deltaA;

    currentRatings[playerA].n_game_played += 1;
    currentRatings[playerB].n_game_played += 1;

    // Calculate Elo gained/lost per player
    const eloGainA = currentRatings[playerA].rating - RA;
    const eloLossB = currentRatings[playerB].rating - RB;

    // Write Elo gain/loss to columns I (col index 8) and J (col index 9) of Matches Log sheet (assuming zero-based)
    matchesSheet.getCell(i, 8).setValue(Math.abs(Math.round(eloGainA * 100) / 100)); // column I
  }

  let maxGamesPlayed = Math.max(...Object.values(currentRatings).map(p => p.n_game_played));
  if (maxGamesPlayed === 0) maxGamesPlayed = 1;

  return { ratings: currentRatings, maxGames: maxGamesPlayed };
}

// --- ELO Calculation and Data Preparation ---
/**
 * Loads data, calculates ELO ratings for all players based on match history,
 * determines previous ranks for comparison, applies penalization, and sorts the results.
 */
function setupAndCalculateElo(
  matchesSheet: ExcelScript.Worksheet,
  ratingsSheet: ExcelScript.Worksheet,
  playersDataSheet: ExcelScript.Worksheet,
  K: number,
  baseRating: number
): {
  allRatings: [string, PlayerStats][];
  ratingsMap: { [player: string]: PlayerStats };
  maxGamesPlayed: number;
  columnIndices: { [key: string]: number };
  previousRanksMap: { [player: string]: number };
} {
  const matchesData: ExcelData = matchesSheet.getUsedRange().getValues();
  const playersData: ExcelData = playersDataSheet.getUsedRange().getValues();
  const ratingsData: ExcelData = ratingsSheet.getUsedRange().getValues();

  // Map column headers to their indices for dynamic access
  const headerRow = ratingsData[0] as string[];
  const columnIndices = {
    elo: headerRow.indexOf("ELO"),
    players: headerRow.indexOf("Challenger"),
    ranking: headerRow.indexOf("Ranking"),
    gamesPlayed: headerRow.indexOf("Game played"),
    nationality: headerRow.indexOf("Nationality"),
    team: headerRow.indexOf("Team"),
    role: headerRow.indexOf("Role"),
    penalization: headerRow.indexOf("Participation Penalization"),
    winrate: headerRow.indexOf("Winrate"),
  };

  // 1. Find the two most recent distinct date keys (YYYYMMDD)
  const dateToKey = (d: ExcelCellValue): string => {
    if (typeof d === 'number' && d > 0) {
      // Convert Excel serial date to JS Date object (25569 is the difference in days)
      const date = new Date(Math.round((d - 25569) * 86400000));

      // Format to YYYYMMDD string for comparison (using UTC to prevent timezone shift issues)
      const year = date.getUTCFullYear().toString();
      const month = (date.getUTCMonth() + 1).toString().padStart(2, '0');
      const day = date.getUTCDate().toString().padStart(2, '0');

      return year + month + day;
    }
    return '00000000';
  };

  const matchDates: ExcelCellValue[] = matchesData.slice(1).map(row => row[0]).filter(date => date);
  // Convert all valid date values to YYYYMMDD keys
  const uniqueDateKeys = Array.from(new Set(matchDates)).map(d => dateToKey(d)).sort((a, b) => b.localeCompare(a));

  const latestDateKey = uniqueDateKeys[0] || null;
  const secondLatestDateKey = uniqueDateKeys[1] || null;

  // 2. Calculate Final ELO and Stats (using all matches, latest date key)
  const { ratings: finalRatingsMap, maxGames: finalMaxGamesPlayed } = runEloCalculations(matchesSheet, 
    matchesData, playersData, latestDateKey, K, baseRating
  );

  // 3. Apply Penalization to Final ELO and store standard ELO
  for (const player in finalRatingsMap) {
    const stats = finalRatingsMap[player];
    const factor = 0.5 + 0.5 * Math.pow(stats.n_game_played / finalMaxGamesPlayed, 0.1);
    stats.standardRating = stats.rating; // ELO before penalization
    stats.rating *= factor;
  }

  // 4. Calculate Previous Ranks (up to second latest date)
  let previousRanksMap: { [player: string]: number } = {};
  const allRatingsSorted = Object.entries(finalRatingsMap).sort((a, b) => b[1].rating - a[1].rating);

  if (secondLatestDateKey) {
    const { ratings: prevRatings, maxGames: prevMaxGames } = runEloCalculations(matchesSheet, 
      matchesData, playersData, secondLatestDateKey, K, baseRating
    );

    // Apply Penalization to Previous ELO
    const prevMaxGamesFinal = prevMaxGames === 0 ? 1 : prevMaxGames;
    for (const player in prevRatings) {
      const stats = prevRatings[player];
      const factor = 0.5 + 0.5 * Math.pow(stats.n_game_played / prevMaxGamesFinal, 0.1);
      stats.rating *= factor;
    }

    // Calculate Ranks from the previous snapshot
    const prevAllRatings = Object.entries(prevRatings).sort((a, b) => b[1].rating - a[1].rating);
    prevAllRatings.forEach(([player], i) => previousRanksMap[player] = i + 1);
  } else {
    // If no previous date, assume all players are new/unchanged (rank change = 0)
    // Set previous rank equal to current rank for all players
    allRatingsSorted.forEach(([player], i) => previousRanksMap[player] = i + 1);
  }

  return {
    allRatings: allRatingsSorted,
    ratingsMap: finalRatingsMap,
    maxGamesPlayed: finalMaxGamesPlayed,
    columnIndices,
    previousRanksMap
  };
}

// --- Leaderboard Table Update ---
/**
 * Clears and writes the calculated ELO and stats to the main Leaderboard table area,
 * including the calculation and display of rank change.
 */
function updateLeaderboardTable(
  ratings: ExcelScript.Worksheet,
  allRatings: [string, PlayerStats][],
  maxGamesPlayed: number,
  columnIndices: { [key: string]: number },
  previousRanksMap: { [player: string]: number }
): void {
  const { ranking, players, nationality, elo, gamesPlayed, team, role, penalization, winrate } = columnIndices;

  // 1. Clear existing data in relevant columns (rows 1 onwards)
  const columnsToClear = [ranking, players, nationality, elo, gamesPlayed, team, role, penalization, winrate].filter(i => i >= 0);
  // Determine the range to clear (from row 1, down to the number of players)
  const clearRowCount = allRatings.length + 50; // Clear a generous number of rows

  // Clear the main leaderboard columns
  for (let row = 1; row < clearRowCount; row++) {
    columnsToClear.forEach(col => {
      // Use clear(ExcelScript.ClearApplyTo.contents) to preserve formatting
      if (col >= 0) ratings.getCell(row, col).clear(ExcelScript.ClearApplyTo.contents);
    });
  }

  // 2. Write new data to leaderboard
  for (let i = 0; i < allRatings.length; i++) {
    const [player, stats] = allRatings[i];
    const currentRank = i + 1;

    // Retrieve previous rank, default to current rank if player had no data in the previous snapshot
    const previousRank = previousRanksMap[player];

    // Calculate rank change: Previous Rank - Current Rank (Positive = climbed, Negative = dropped)
    const rankChange = (previousRank !== undefined ? previousRank : currentRank) - currentRank;

    // Format Rank String: Current Rank (Change)
    let rankString: string;
    let stringColor: string;
    if (rankChange > 0) {
      rankString = `${currentRank} (↑${rankChange})`;
      stringColor = "#00B050";
    } else if (rankChange < 0) {
      rankString = `${currentRank} (↓${Math.abs(rankChange)})`; // Negative sign is included
      stringColor = "#F00000";
    } else {
      rankString = `${currentRank} (~)`;
      stringColor = "#000000";
    }

    // Calculate points lost due to penalization
    const standardElo = stats.standardRating ?? stats.rating; // Should always be defined after setup, but fallback to current
    const pointsLost = standardElo - stats.rating;

    // Calculate Win/Loss Differential
    const winLossDiff = stats.wins - (stats.n_game_played - stats.wins);

    // Write values
    ratings.getCell(i + 1, ranking).setValue(rankString); // Updated column A (Ranking)
    ratings.getCell(i + 1, ranking).getFormat().getFont().setColor(stringColor);
    ratings.getCell(i + 1, players).setValue(player);
    ratings.getCell(i + 1, nationality).setValue(stats.nationality);
    ratings.getCell(i + 1, team).setValue(stats.team);
    ratings.getCell(i + 1, role).setValue(stats.role);
    ratings.getCell(i + 1, gamesPlayed).setValue(stats.n_game_played);
    ratings.getCell(i + 1, elo).setValue(Math.round(stats.rating));
    // Penalization: Display as a negative number for points lost
    ratings.getCell(i + 1, penalization).setValue(-Math.round(pointsLost * 100) / 100);

    // Write Win/Loss Differential with conditional formatting
    const winLossCell = ratings.getCell(i + 1, winrate);
    winLossCell.setValue(winLossDiff);
    winLossCell.setNumberFormat("+0;-0;0");

    if (winLossDiff > 0) {
      winLossCell.getFormat().getFont().setColor("#00B050"); // Green
    } else if (winLossDiff === 0) {
      winLossCell.getFormat().getFont().setColor("#FFA000"); // Amber/Orange
    } else {
      winLossCell.getFormat().getFont().setColor("#F00000"); // Red
    }
  }
}

// --- Podium Views Update ---
/**
 * Calculates and updates the Player, Country, Team, and Role podiums.
 * (Logic is the same as before, using final ratingsMap)
 */
function updatePodiumViews(
  ratings: ExcelScript.Worksheet,
  allRatings: [string, PlayerStats][],
  ratingsMap: { [player: string]: PlayerStats }
): void {
  // --- 1. Update Player Podium (hardcoded cells: Column K=10, L=11, M=12) ---
  ratings.getCell(7, 11).setValue("🥇 1st");
  ratings.getCell(8, 10).setValue("🥈 2nd");
  ratings.getCell(8, 12).setValue("🥉 3rd");

  // Top 3 players
  ratings.getCell(5, 11).setValue(allRatings[0]?.[0] || "");
  ratings.getCell(6, 10).setValue(allRatings[1]?.[0] || "");
  ratings.getCell(6, 12).setValue(allRatings[2]?.[0] || "");
  ratings.getCell(6, 11).setValue(Math.round(allRatings[0]?.[1].rating ?? 0));
  ratings.getCell(7, 10).setValue(Math.round(allRatings[1]?.[1].rating ?? 0));
  ratings.getCell(7, 12).setValue(Math.round(allRatings[2]?.[1].rating ?? 0));

  // Bottom player ("💩" cell)
  ratings.getCell(5, 15).setValue(allRatings[allRatings.length - 1]?.[0] || "");
  ratings.getCell(6, 15).setValue(Math.round(allRatings[allRatings.length - 1]?.[1].rating ?? 0));
  ratings.getCell(7, 15).setValue("💩");

  // --- 2. Calculate Category ELO Averages ---
  const countryEloSum: { [country: string]: number } = {};
  const countryCount: { [country: string]: number } = {};
  const teamEloSum: { [team: string]: number } = {};
  const teamCount: { [team: string]: number } = {};
  const roleEloSum: { [role: string]: number } = {};
  const roleCount: { [role: string]: number } = {};

  for (const [player, stats] of allRatings) {
    const { nationality, team, role, rating: elo } = stats;

    // Helper function to update sums/counts
    const updateCategory = (category: string, sumMap: { [key: string]: number }, countMap: { [key: string]: number }) => {
      sumMap[category] = (sumMap[category] || 0) + elo;
      countMap[category] = (countMap[category] || 0) + 1;
    };

    updateCategory(nationality, countryEloSum, countryCount);
    updateCategory(team, teamEloSum, teamCount);
    updateCategory(role, roleEloSum, roleCount);
  }

  const avg = (sum: number, count: number) => count > 0 ? sum / count : 0;

  const topCountries = Object.entries(countryEloSum)
    .map(([name, sum]) => ({ name, avg: avg(sum, countryCount[name]) }))
    .sort((a, b) => b.avg - a.avg);

  const topTeams = Object.entries(teamEloSum)
    .map(([name, sum]) => ({ name, avg: avg(sum, teamCount[name]) }))
    .sort((a, b) => b.avg - a.avg);

  const topRoles = Object.entries(roleEloSum)
    .map(([name, sum]) => ({ name, avg: avg(sum, roleCount[name]) }))
    .sort((a, b) => b.avg - a.avg);

  // --- 3. Write Country Podium (hardcoded cells: Rows 11-14) ---
  if (topCountries.length > 0) {
    ratings.getCell(11, 11).setValue(topCountries[0].name);
    ratings.getCell(12, 11).setValue(Math.round(topCountries[0].avg));
    ratings.getCell(13, 11).setValue("🥇 1st");
  }
  if (topCountries.length > 1) {
    ratings.getCell(12, 10).setValue(topCountries[1].name);
    ratings.getCell(13, 10).setValue(Math.round(topCountries[1].avg));
    ratings.getCell(14, 10).setValue("🥈 2nd");
  }
  if (topCountries.length > 2) {
    ratings.getCell(12, 12).setValue(topCountries[2].name);
    ratings.getCell(13, 12).setValue(Math.round(topCountries[2].avg));
    ratings.getCell(14, 12).setValue("🥉 3rd");
  }

  // --- 4. Write Team Podium (hardcoded cells: Rows 17-20) ---
  if (topTeams.length > 0) {
    ratings.getCell(17, 11).setValue(topTeams[0].name);
    ratings.getCell(18, 11).setValue(Math.round(topTeams[0].avg));
    ratings.getCell(19, 11).setValue("🥇 1st");
  }
  if (topTeams.length > 1) {
    ratings.getCell(18, 10).setValue(topTeams[1].name);
    ratings.getCell(19, 10).setValue(Math.round(topTeams[1].avg));
    ratings.getCell(20, 10).setValue("🥈 2nd");
  }
  if (topTeams.length > 2) {
    ratings.getCell(18, 12).setValue(topTeams[2].name);
    ratings.getCell(19, 12).setValue(Math.round(topTeams[2].avg));
    ratings.getCell(20, 12).setValue("🥉 3rd");
  }

  // --- 5. Write Role Podium (hardcoded cells: Rows 23-26) ---
  if (topRoles.length > 0) {
    ratings.getCell(23, 11).setValue(topRoles[0].name);
    ratings.getCell(24, 11).setValue(Math.round(topRoles[0].avg));
    ratings.getCell(25, 11).setValue("🥇 1st");
  }
  if (topRoles.length > 1) {
    ratings.getCell(24, 10).setValue(topRoles[1].name);
    ratings.getCell(25, 10).setValue(Math.round(topRoles[1].avg));
    ratings.getCell(26, 10).setValue("🥈 2nd");
  }
  if (topRoles.length > 2) {
    ratings.getCell(24, 12).setValue(topRoles[2].name);
    ratings.getCell(25, 12).setValue(Math.round(topRoles[2].avg));
    ratings.getCell(26, 12).setValue("🥉 3rd");
  }
}

// --- Winrates Matrix Wrapper ---
/**
 * Wrapper function to call all three matrix update functions.
 */
function updateAllWinratesMatrices(workbook: ExcelScript.Workbook) {
  updateCountryWinratesMatrix(workbook);
  updateTeamWinratesMatrix(workbook);
  updateRoleWinratesMatrix(workbook);
}


// --- Winrates Matrix Functions (Original Logic Maintained) ---

/**
 * Updates the Country vs Country Win/Loss matrix on the "Winrates" sheet.
 */
function updateCountryWinratesMatrix(workbook: ExcelScript.Workbook) {
  const matchesSheet = workbook.getWorksheet("Matches Log");
  const challengersSheet = workbook.getWorksheet("Challengers cards");
  const winratesSheet = workbook.getWorksheet("Winrates");

  if (!matchesSheet || !challengersSheet || !winratesSheet) return;

  const matchesData: ExcelData = matchesSheet.getUsedRange().getValues();
  const challengersData: ExcelData = challengersSheet.getUsedRange().getValues();

  // Player to country map (Challengers cards, country in column B = index 1)
  const playerToCountry: { [player: string]: string } = {};
  for (let i = 1; i < challengersData.length; i++) {
    const player = challengersData[i][0] as string;
    if (player) playerToCountry[player] = challengersData[i][1] as string;
  }

  // Get column countries from B2 to N2 (columns 1 to 13, row index 1)
  const columnCountries: string[] = [];
  for (let c = 1; c <= 13; c++) {
    const country = winratesSheet.getCell(1, c).getText().trim();
    columnCountries.push(country);
  }

  // Get row countries from A3 to A28 (rows 2 to 27)
  const rowCountries: string[] = [];
  for (let r = 2; r <= 27; r += 2) {
    // Each country spans two rows; use the top one for country name
    const country = winratesSheet.getCell(r, 0).getText().trim();
    rowCountries.push(country);
  }

  // Initialize matrices for wins and losses
  const winsMatrix: number[][] = [];
  const lossesMatrix: number[][] = [];
  for (let i = 0; i < rowCountries.length; i++) {
    winsMatrix[i] = new Array(columnCountries.length).fill(0);
    lossesMatrix[i] = new Array(columnCountries.length).fill(0);
  }

  // Count wins and losses per country pair
  for (let i = 1; i < matchesData.length; i++) {
    const playerA = matchesData[i][1] as string;
    const playerB = matchesData[i][2] as string;
    const winner = matchesData[i][5] as string;

    if (!playerA || !playerB || !winner) continue;
    const countryA = playerToCountry[playerA] || "Unknown";
    const countryB = playerToCountry[playerB] || "Unknown";

    const rowIndex = rowCountries.indexOf(countryA);
    const colIndex = columnCountries.indexOf(countryB);

    if (rowIndex >= 0 && colIndex >= 0) {
      if (winner === playerA) {
        // Player A (Country A) won over Player B (Country B)
        winsMatrix[rowIndex][colIndex]++;
        // Country B lost to Country A
        lossesMatrix[colIndex][rowIndex]++;
      } else if (winner === playerB) {
        // Player B (Country B) won over Player A (Country A)
        winsMatrix[colIndex][rowIndex]++;
        // Country A lost to Country B
        lossesMatrix[rowIndex][colIndex]++;
      }
    }
  }

  // Write wins and losses to winrates matrix:
  // Top row (r) gets Wins, bottom row (r+1) gets Losses per country pair
  for (let r = 0; r < rowCountries.length; r++) {
    for (let c = 0; c < columnCountries.length; c++) {
      const topRow = 2 + r * 2;
      const col = c + 1; // Column index is 1-based
      const isDiagonal = rowCountries[r] === columnCountries[c] && r === c;

      // Cells for wins (top) and losses (bottom)
      const winCell = winratesSheet.getCell(topRow, col);
      const lossCell = winratesSheet.getCell(topRow + 1, col);

      // Clear previous fills and contents for safety
      winCell.getFormat().getFill().clear();
      lossCell.getFormat().getFill().clear();

      // Clear contents only if not diagonal (diagonal contains total games)
      if (!isDiagonal) {
        winCell.clear(ExcelScript.ClearApplyTo.contents);
        lossCell.clear(ExcelScript.ClearApplyTo.contents);
      }


      if (isDiagonal) {
        // Diagonal cell: write total games played for that country
        const gamesPlayed = winsMatrix[r][c] + lossesMatrix[r][c];
        if (gamesPlayed > 0) {
          winCell.setValue(gamesPlayed / 2);
          winCell.getFormat().getFill().setColor("#d9e1f2"); // very light blue
        }
      } else {
        // Off-diagonal: wins/losses
        if (winsMatrix[r][c] > 0) {
          winCell.setValue(winsMatrix[r][c]);
          winCell.getFormat().getFill().setColor("#d9f2d9"); // very light green
        } else {
          winCell.clear(ExcelScript.ClearApplyTo.contents);
        }

        if (lossesMatrix[r][c] > 0) {
          lossCell.setValue(lossesMatrix[r][c]);
          lossCell.getFormat().getFill().setColor("#f4cccc"); // very light red
        } else {
          lossCell.clear(ExcelScript.ClearApplyTo.contents);
        }
      }
    }
  }
}

/**
 * Updates the Team vs Team Win/Loss matrix on the "Winrates" sheet.
 */
function updateTeamWinratesMatrix(workbook: ExcelScript.Workbook) {
  const matchesSheet = workbook.getWorksheet("Matches Log");
  const challengersSheet = workbook.getWorksheet("Challengers cards");
  const winratesSheet = workbook.getWorksheet("Winrates");

  if (!matchesSheet || !challengersSheet || !winratesSheet) return;

  const matchesData: ExcelData = matchesSheet.getUsedRange().getValues();
  const challengersData: ExcelData = challengersSheet.getUsedRange().getValues();

  // Player to team map (Challengers cards, team in column C = index 2)
  const playerToTeam: { [player: string]: string } = {};
  for (let i = 1; i < challengersData.length; i++) {
    const player = challengersData[i][0] as string;
    if (player) playerToTeam[player] = challengersData[i][2] as string;
  }

  // Get column teams from B32 to H32 (columns 1 to 7, row 31 zero-based)
  const columnTeams: string[] = [];
  for (let c = 1; c <= 7; c++) {
    const team = winratesSheet.getCell(31, c).getText().trim();
    columnTeams.push(team);
  }

  // Get row teams from A33 to A46 stepping by 2 rows (rows 32 to 45 zero-based)
  const rowTeams: string[] = [];
  for (let r = 32; r <= 45; r += 2) {
    const team = winratesSheet.getCell(r, 0).getText().trim();
    rowTeams.push(team);
  }

  const winsMatrix: number[][] = [];
  const lossesMatrix: number[][] = [];
  for (let i = 0; i < rowTeams.length; i++) {
    winsMatrix[i] = new Array(columnTeams.length).fill(0);
    lossesMatrix[i] = new Array(columnTeams.length).fill(0);
  }

  for (let i = 1; i < matchesData.length; i++) {
    const playerA = matchesData[i][1] as string;
    const playerB = matchesData[i][2] as string;
    const winner = matchesData[i][5] as string;
    if (!playerA || !playerB || !winner) continue;

    const teamA = playerToTeam[playerA] || "Unknown";
    const teamB = playerToTeam[playerB] || "Unknown";

    const rowIndex = rowTeams.indexOf(teamA);
    const colIndex = columnTeams.indexOf(teamB);

    if (rowIndex >= 0 && colIndex >= 0) {
      if (winner === playerA) {
        winsMatrix[rowIndex][colIndex]++;
        lossesMatrix[colIndex][rowIndex]++;
      } else if (winner === playerB) {
        winsMatrix[colIndex][rowIndex]++;
        lossesMatrix[rowIndex][colIndex]++;
      }
    }
  }

  for (let r = 0; r < rowTeams.length; r++) {
    for (let c = 0; c < columnTeams.length; c++) {
      const topRow = 32 + r * 2;
      const col = c + 1;
      const isDiagonal = rowTeams[r] === columnTeams[c] && r === c;

      const winCell = winratesSheet.getCell(topRow, col);
      const lossCell = winratesSheet.getCell(topRow + 1, col);

      // Clear previous fills and contents for safety
      winCell.getFormat().getFill().clear();
      lossCell.getFormat().getFill().clear();

      // Clear contents only if not diagonal (diagonal contains total games)
      if (!isDiagonal) {
        winCell.clear(ExcelScript.ClearApplyTo.contents);
        lossCell.clear(ExcelScript.ClearApplyTo.contents);
      }

      if (isDiagonal) {
        const gamesPlayed = winsMatrix[r][c] + lossesMatrix[r][c];
        if (gamesPlayed > 0) {
          winCell.setValue(gamesPlayed / 2);
          winCell.getFormat().getFill().setColor("#d9e1f2"); // very light blue
        }
      } else {
        if (winsMatrix[r][c] > 0) {
          winCell.setValue(winsMatrix[r][c]);
          winCell.getFormat().getFill().setColor("#d9f2d9"); // very light green
        } else {
          winCell.clear(ExcelScript.ClearApplyTo.contents);
        }

        if (lossesMatrix[r][c] > 0) {
          lossCell.setValue(lossesMatrix[r][c]);
          lossCell.getFormat().getFill().setColor("#f4cccc"); // very light red
        } else {
          lossCell.clear(ExcelScript.ClearApplyTo.contents);
        }
      }
    }
  }
}

/**
 * Updates the Role vs Role Win/Loss matrix on the "Winrates" sheet.
 */
function updateRoleWinratesMatrix(workbook: ExcelScript.Workbook) {
  const matchesSheet = workbook.getWorksheet("Matches Log");
  const challengersSheet = workbook.getWorksheet("Challengers cards");
  const winratesSheet = workbook.getWorksheet("Winrates");

  if (!matchesSheet || !challengersSheet || !winratesSheet) return;

  const matchesData: ExcelData = matchesSheet.getUsedRange().getValues();
  const challengersData: ExcelData = challengersSheet.getUsedRange().getValues();

  // Player to role map (Challengers cards, role in column D = index 3)
  const playerToRole: { [player: string]: string } = {};
  for (let i = 1; i < challengersData.length; i++) {
    const player = challengersData[i][0] as string;
    if (player) playerToRole[player] = challengersData[i][3] as string;
  }

  // Get column roles from B50 to I50 (columns 1 to 8, row 49 zero-based)
  const columnRoles: string[] = [];
  for (let c = 1; c <= 8; c++) {
    const role = winratesSheet.getCell(49, c).getText().trim();
    columnRoles.push(role);
  }

  // Get row roles from A51 to A66 stepping by 2 rows (rows 50 to 65 zero-based)
  const rowRoles: string[] = [];
  for (let r = 50; r <= 65; r += 2) {
    const role = winratesSheet.getCell(r, 0).getText().trim();
    rowRoles.push(role);
  }

  const winsMatrix: number[][] = [];
  const lossesMatrix: number[][] = [];
  for (let i = 0; i < rowRoles.length; i++) {
    winsMatrix[i] = new Array(columnRoles.length).fill(0);
    lossesMatrix[i] = new Array(columnRoles.length).fill(0);
  }

  for (let i = 1; i < matchesData.length; i++) {
    const playerA = matchesData[i][1] as string;
    const playerB = matchesData[i][2] as string;
    const winner = matchesData[i][5] as string;
    if (!playerA || !playerB || !winner) continue;

    const roleA = playerToRole[playerA] || "Unknown";
    const roleB = playerToRole[playerB] || "Unknown";

    const rowIndex = rowRoles.indexOf(roleA);
    const colIndex = columnRoles.indexOf(roleB);

    if (rowIndex >= 0 && colIndex >= 0) {
      if (winner === playerA) {
        winsMatrix[rowIndex][colIndex]++;
        lossesMatrix[colIndex][rowIndex]++;
      } else if (winner === playerB) {
        winsMatrix[colIndex][rowIndex]++;
        lossesMatrix[rowIndex][colIndex]++;
      }
    }
  }

  for (let r = 0; r < rowRoles.length; r++) {
    for (let c = 0; c < columnRoles.length; c++) {
      const topRow = 50 + r * 2;
      const col = c + 1;
      const isDiagonal = rowRoles[r] === columnRoles[c] && r === c;

      const winCell = winratesSheet.getCell(topRow, col);
      const lossCell = winratesSheet.getCell(topRow + 1, col);

      // Clear previous fills and contents for safety
      winCell.getFormat().getFill().clear();
      lossCell.getFormat().getFill().clear();

      // Clear contents only if not diagonal (diagonal contains total games)
      if (!isDiagonal) {
        winCell.clear(ExcelScript.ClearApplyTo.contents);
        lossCell.clear(ExcelScript.ClearApplyTo.contents);
      }

      if (isDiagonal) {
        const gamesPlayed = winsMatrix[r][c] + lossesMatrix[r][c];
        if (gamesPlayed > 0) {
          winCell.setValue(gamesPlayed / 2);
          winCell.getFormat().getFill().setColor("#d9e1f2"); // very light blue
        }
      } else {
        if (winsMatrix[r][c] > 0) {
          winCell.setValue(winsMatrix[r][c]);
          winCell.getFormat().getFill().setColor("#d9f2d9"); // very light green
        } else {
          winCell.clear(ExcelScript.ClearApplyTo.contents);
        }

        if (lossesMatrix[r][c] > 0) {
          lossCell.setValue(lossesMatrix[r][c]);
          lossCell.getFormat().getFill().setColor("#f4cccc"); // very light red
        } else {
          lossCell.clear(ExcelScript.ClearApplyTo.contents);
        }
      }
    }
  }
}
