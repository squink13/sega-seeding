let startTime = new Date();
let clientId = "YOUR_CLIENT_ID";
let clientSecret = "YOUR_CLIENT_SECRET";

const SS = SpreadsheetApp.getActiveSpreadsheet();
const PROPERTIES = PropertiesService.getScriptProperties();

const DuelRatingByRank = {
  1000: 6.75,
  2000: 6.7,
  3000: 6.65,
  4000: 6.6,
  5000: 6.55,
  6000: 6.453,
  7000: 6.413,
  8000: 6.375,
  9000: 6.339,
  10000: 6.303,
  11000: 6.269,
  12000: 6.236,
  13000: 6.205,
  14000: 6.175,
  15000: 6.146,
  16000: 6.118,
  17000: 6.092,
  18000: 6.067,
  19000: 6.043,
  20000: 6.02,
  21000: 5.999,
  22000: 5.979,
  23000: 5.96,
  24000: 5.943,
  25000: 5.927,
  26000: 5.912,
  27000: 5.898,
  28000: 5.886,
  29000: 5.875,
};

function main() {
  Logger.log("Starting script");
  parsePlayers("_import", "_export");
  Logger.log("Finished script");
}

function parsePlayers(importSheetName, exportSheetName) {
  let importedPlayers = readData(importSheetName);
  let players = [];
  let badges, player, osuData;
  let start = parseInt(PROPERTIES.getProperty("start")) || 0;

  for (let i = start; i < importedPlayers.length; i++) {
    badges = 0;
    player = importedPlayers[i];
    osuData = getOsuData(player.userId);

    player.username = osuData.username;
    player.country = osuData.country_code;
    player.rank = osuData.statistics.global_rank;

    Logger.log(`Parsed player ${player.username}`);

    player.badges = filteredBadges(osuData);

    player.bwsRank = Math.pow(
      player.rank,
      Math.pow(0.9937, Math.pow(player.badges, 1.7))
    );

    if (player.provisional || player.outdated) {
      Logger.log(
        `Player ${player.username} is provisional or outdated ${player.provisional} ${player.outdated}`
      );
      let rank = Math.floor(parseInt(player.rank) / 1000) * 1000;
      if (rank in DuelRatingByRank) {
        player.duelRating = DuelRatingByRank[rank];
      } else {
        player.duelRating = 0.001;
      }
    }

    players.push(player);
    // Set the property for the next starting point
    PROPERTIES.setProperty("start", i + 1);

    // You can also add a termination condition based on script execution time
    // to prevent reaching the maximum execution time limit.
    // For example, terminate if the script has been running for over 5 minutes:
    if (new Date().getTime() - startTime.getTime() > 300000) {
      Logger.log("Execution time limit approaching, terminating script...");
      return;
    }
  }

  // Clear the property after all rows are processed
  PROPERTIES.deleteProperty("start");
  printPlayersData(players, exportSheetName);
}

function printPlayersData(players, sheetName) {
  let sheet = SS.getSheetByName(sheetName);
  if (!sheet) {
    sheet = SS.insertSheet(sheetName);
  }

  let headers = [
    "userId",
    "username",
    "country",
    "rank",
    "badges",
    "bwsRank",
    "duelRating",
    "provisional",
    "outdated",
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  players.sort((a, b) => b.duelRating - a.duelRating);
  let data = players.map((player) => [
    player.userId,
    player.username,
    player.country,
    player.rank,
    player.badges,
    player.bwsRank,
    player.duelRating,
    player.provisional,
    player.outdated,
  ]);
  sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
}

function getOsuData(userId) {
  let url = `https://osu.ppy.sh/api/v2/users/${userId}/osu`;
  let headers = {
    Authorization: `Bearer ${getAccessToken()}`,
  };
  let options = {
    contentType: "application/json",
    key: "id",
    headers: headers,
  };
  let response = UrlFetchApp.fetch(url, options);
  response = JSON.parse(response);
  return response;
}

function getAccessToken() {
  let seconds = new Date().getTime() / 1000;

  if (seconds >= PROPERTIES.getProperty("expires_in")) {
    let data = {
      client_id: clientId,
      client_secret: clientSecret,
      grant_type: "client_credentials",
      scope: "public",
    };

    let options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(data),
    };

    let response = UrlFetchApp.fetch("https://osu.ppy.sh/oauth/token", options);

    PROPERTIES.setProperty(
      "expires_in",
      Math.round(new Date().getTime() / 1000 + JSON.parse(response).expires_in)
    );
    PROPERTIES.setProperty("access_token", JSON.parse(response).access_token);

    Logger.log("Got access token");
  }

  return PROPERTIES.getProperty("access_token");
}

function filteredBadges(player) {
  let filterRange = SS.getRangeByName("_filtered_badges!A1")
    .getDataRegion(SpreadsheetApp.Dimension.ROWS)
    .getValues()
    .flat(2)
    .filter((i) => i);

  let expression = filterRange.slice(1).join("|");
  const ignoredBadges = new RegExp(expression, "i");

  let badges = 0;
  for (let badge in player.badges) {
    let currentBadge = player.badges[badge];
    let awardedAt = new Date(currentBadge.awarded_at);
    if (
      awardedAt.getFullYear() >= 2021 &&
      !ignoredBadges.test(currentBadge.description.toLowerCase())
    ) {
      badges++;
    }
  }

  return badges;
}

function readData(sheetName) {
  Logger.log("Reading data");
  let sheet = SS.getSheetByName(sheetName);
  let data = sheet.getDataRange().getValues();
  let players = [];
  let userId, duelRating, provisional, outdated;
  for (var i = 1; i < data.length; i++) {
    userId = parseInt(data[i][0]);
    duelRating = parseFloat(data[i][1]);
    provisional = data[i][2];
    outdated = data[i][3];

    players.push({ userId, duelRating, provisional, outdated });
  }

  return players;
}
