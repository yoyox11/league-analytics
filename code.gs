const API_KEY = "YOUR_RIOT_API_KEY_HERE";
const REGION_ROUTING = "europe"; 
const REGION_PLATFORM = "euw1";  

function refreshLoLStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const riotId = sheet.getRange("A2").getValue();
  const tagline = sheet.getRange("B2").getValue();

  // 1. RECUP AUTO - IMAGE
  const versionUrl = "https://ddragon.leagueoflegends.com/api/versions.json";
  const latestVersion = JSON.parse(UrlFetchApp.fetch(versionUrl).getContentText())[0];

  // 2. PUUID
  const puuidUrl = `https://${REGION_ROUTING}.api.riotgames.com/riot/account/v1/accounts/by-riot-id/${riotId}/${tagline}?api_key=${API_KEY}`;
  const puuid = JSON.parse(UrlFetchApp.fetch(puuidUrl).getContentText()).puuid;

  // 3. RANK (in E2)
  const currentRank = getPlayerRank(puuid, API_KEY);
  sheet.getRange("E2").setValue(currentRank); 

  // 4. GAMES (Buffer : 30)
  const matchIdsUrl = `https://${REGION_ROUTING}.api.riotgames.com/lol/match/v5/matches/by-puuid/${puuid}/ids?type=ranked&start=0&count=30&api_key=${API_KEY}`;
  const matchIds = JSON.parse(UrlFetchApp.fetch(matchIdsUrl).getContentText());

  const rows = [];
  let avg = { kda: 0, visionMin: 0, kp: 0, dpm: 0, deathPct: 0, wins: 0 };
  let validMatchCount = 0;

  for (let i = 0; i < matchIds.length; i++) {
    if (validMatchCount >= 15) break;
    const matchUrl = `https://${REGION_ROUTING}.api.riotgames.com/lol/match/v5/matches/${matchIds[i]}?api_key=${API_KEY}`;
    const response = UrlFetchApp.fetch(matchUrl, {"muteHttpExceptions": true});
    if (response.getResponseCode() !== 200) continue;
    const matchData = JSON.parse(response.getContentText());
    const p = matchData.info.participants.find(part => part.puuid === puuid);

    if (matchData.info.gameDuration < 600 || p.teamPosition !== "UTILITY") continue; 

    const durationMin = matchData.info.gameDuration / 60;
    const kda = (p.kills + p.assists) / Math.max(1, p.deaths);
    const visionMin = p.visionScore / durationMin;
    const kp = p.challenges ? (p.challenges.killParticipation * 100) : 0;
    const dpm = p.totalDamageDealtToChampions / durationMin;
    const deathPct = p.deaths / Math.max(1, p.kills + p.assists + p.deaths);
    
    avg.kda += kda; avg.visionMin += visionMin; avg.kp += kp; avg.dpm += dpm; avg.deathPct += deathPct;
    if (p.win) avg.wins++;
    validMatchCount++;

    rows.push([
      new Date(matchData.info.gameStartTimestamp),
      p.championName,
      `=IMAGE("https://ddragon.leagueoflegends.com/cdn/${latestVersion}/img/champion/${p.championName}.png")`,
      p.win ? "WIN" : "LOSE",
      kda.toFixed(2).replace('.', ','),
      visionMin.toFixed(1).replace('.', ','),
      kp.toFixed(0),
      Math.round(dpm),
      (deathPct * 100).toFixed(1).replace('.', ',')
    ]);
  }

  // 5. DISPLAY
  sheet.getRange("A6:I100").clearContent();
  if (rows.length > 0) {
    sheet.getRange(6, 1, rows.length, 9).setValues(rows);
    const summary = [[avg.kda/validMatchCount, avg.visionMin/validMatchCount, avg.kp/validMatchCount, avg.dpm/validMatchCount, avg.deathPct/validMatchCount]];
    sheet.getRange("L2:P2").setValues(summary); 
  }
  
  // 6. ARCHIVE
  archiveMatches(latestVersion);
}

function getPlayerRank(puuid, key) {
  try {
    const url = `https://${REGION_PLATFORM}.api.riotgames.com/lol/league/v4/entries/by-puuid/${puuid}?api_key=${key}`;
    const res = UrlFetchApp.fetch(url, {"muteHttpExceptions": true});
    if (res.getResponseCode() !== 200) return "ERROR " + res.getResponseCode();
    const lData = JSON.parse(res.getContentText());
    let rank = "UNRANKED";
    lData.forEach(e => {
      if (e.queueType === "RANKED_SOLO_5x5") 
        rank = `${e.tier} ${e.rank} (${e.leaguePoints} LP)`;
    });
    return rank;
  } catch (e) { return "ERROR"; }
}

function archiveMatches(version) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName("Support");
  const archiveSheet = ss.getSheetByName("Archive");
  if (!archiveSheet) return;

  const data = mainSheet.getRange(6, 1, 15, 9).getValues(); 
  const colA = archiveSheet.getRange("A:A").getValues();
  let lastRow = 1;
  for (let i = 1; i < colA.length; i++) {
    if (colA[i][0] instanceof Date) {
      lastRow = i + 1;
  }
}
  
  let existingKeys = new Set();
  if (lastRow > 1) {
    const archiveData = archiveSheet.getRange(2, 1, lastRow - 1, 2).getValues();
    archiveData.forEach(row => {
      let dKey = (row[0] instanceof Date) ? row[0].toISOString() : row[0].toString();
      existingKeys.add(dKey + row[1]);
    });
  }

  let newCount = 0;
let writeRow = lastRow + 1; // commence après la dernière vraie ligne

for (let i = 0; i < data.length; i++) {
  let row = data[i];
  let dateObj = row[0];
  let champ = row[1];

  if (!(dateObj instanceof Date) || !champ) continue;
  
  let uniqueKey = dateObj.toISOString() + champ;
  if (!existingKeys.has(uniqueKey)) {
    let cleanRow = [...row];
    cleanRow[2] = `=IMAGE("https://ddragon.leagueoflegends.com/cdn/${version}/img/champion/${champ}.png")`;
    archiveSheet.getRange(writeRow, 1, 1, cleanRow.length).setValues([cleanRow]);
    writeRow++;
    newCount++;
  }
}
  if (newCount > 0) ss.toast("✅ " + newCount + " Archive");
}
