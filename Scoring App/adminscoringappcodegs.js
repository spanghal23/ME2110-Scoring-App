// --- 1. UI SETUP ---
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ ME 2110 Admin')
    .addItem('1. Generate Round of 32 (From R1/R2 Scores)', 'generateRoundOf32')
    .addSeparator()
    .addItem('2. Advance Winners -> Sweet 16', 'advanceToSweet16')
    .addItem('3. Advance Winners -> Elite 8', 'advanceToElite8')
    .addItem('4. Advance Winners -> Final 4', 'advanceToFinal4')
    .addSeparator()
    .addItem('⚠️ RESET ALL COMPETITION DATA', 'clearAllData')
    .addToUi();
}

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ME 2110 Scoring App')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1'); 
}

function shuffleArray(array) {
  for (var i = array.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var temp = array[i]; array[i] = array[j]; array[j] = temp;
  }
  return array;
}

function getColMap(sheet) {
  var headers = sheet.getDataRange().getValues()[0];
  var map = { round: 1, track: 2, heat: 3, teamId: 5, score: 17, isDq: 18, 
              cIn: 10, cOut: 11, sIn: 12, sOut: 13, mTier: 9, kTop: 14, kDisp: 15, lumas: 16 };
  var findCol = function(names, fallback) {
     for(var j=0; j<names.length; j++) {
         for (var i=0; i<headers.length; i++) { if (String(headers[i]).toLowerCase().trim() === names[j].toLowerCase().trim()) return i; }
         for (var i=0; i<headers.length; i++) { if (String(headers[i]).toLowerCase().trim().indexOf(names[j].toLowerCase().trim()) > -1) return i; }
     }
     return fallback;
  };
  map.round = findCol(["Round"], map.round); map.track = findCol(["Track Number", "Track"], map.track); map.heat = findCol(["Heat"], map.heat);
  map.teamId = findCol(["Team ID", "ID"], map.teamId); map.score = findCol(["Total Score", "Score"], map.score); map.isDq = findCol(["DQ?", "DQ"], map.isDq);
  map.cIn = findCol(["Coins Inner"], map.cIn); map.cOut = findCol(["Coins Outer", "Coin Outer"], map.cOut); map.sIn = findCol(["Starbits Inner"], map.sIn); map.sOut = findCol(["Starbits Outer"], map.sOut);
  map.mTier = findCol(["Mario Tier"], map.mTier); map.kTop = findCol(["Koopas Toppled", "Toppled"], map.kTop); map.kDisp = findCol(["Koopas Displaced", "Displaced"], map.kDisp); map.lumas = findCol(["Unique Lumas", "Lumas"], map.lumas);
  return map;
}

function getSheetNameFromRound(roundStr) {
  if (roundStr.indexOf("Round 1") > -1 || roundStr.indexOf("Round 2") > -1) return "Matches";
  if (roundStr.indexOf("32") > -1) return "Round of 32";
  if (roundStr.indexOf("16") > -1) return "Sweet 16";
  if (roundStr.indexOf("8") > -1) return "Elite 8";
  if (roundStr.indexOf("4") > -1) return "Final 4";
  return "Matches"; 
}

function getRoundNumber(roundStr) {
  if (roundStr.indexOf("32") > -1) return 32;
  if (roundStr.indexOf("16") > -1) return 16;
  if (roundStr.indexOf("8") > -1) return 8;
  if (roundStr.indexOf("4") > -1) return 4;
  return 1; 
}

function isValidTeam(id) {
  if (!id) return false;
  var clean = String(id).trim().toUpperCase();
  return clean !== "" && clean !== "NONE" && clean !== "TBD" && clean !== "EMPTY" && clean !== "UNDEFINED";
}

function getPlayoffTiebreakerWeights(scoreRow, map) {
  var coinPoints = (Number(scoreRow[map.cIn]) * 7) + (Number(scoreRow[map.cOut]) * 3);
  var starbitPoints = (Number(scoreRow[map.sIn]) * 9) + (Number(scoreRow[map.sOut]) * 4);
  var marioPoints = 0;
  if(scoreRow[map.mTier] === "Bottom") marioPoints = 11;
  else if(scoreRow[map.mTier] === "Middle") marioPoints = 22;
  else if(scoreRow[map.mTier] === "Top") marioPoints = 37;
  var koopaPoints = (Number(scoreRow[map.kTop]) * 6) + (Number(scoreRow[map.kDisp]) * 2);
  var lumaPoints = (Number(scoreRow[map.lumas]) * 10);
  return { tier1: coinPoints + starbitPoints, tier2: lumaPoints, tier3: marioPoints, tier4: koopaPoints };
}

// ==========================================
// THE SINGLE SOURCE OF TRUTH LEADERBOARD
// ==========================================
function updateLiveLeaderboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var scoresSheet = ss.getSheetByName("Scores");
  var teamSheet = ss.getSheetByName("Teams");
  var lbSheet = ss.getSheetByName("Live_Leaderboard");
  
  if (!scoresSheet || !teamSheet || !lbSheet) return;

  var scoreData = scoresSheet.getDataRange().getValues();
  var teamData = teamSheet.getDataRange().getValues();
  var map = getColMap(scoresSheet);
  
  var teamDict = {};
  var sprint3Scores = {};
  
  for(var i=1; i<teamData.length; i++) {
    if(teamData[i][0]) {
      var cId = String(teamData[i][0]).trim().toUpperCase();
      teamDict[cId] = String(teamData[i][1]).trim();
      sprint3Scores[cId] = parseInt(teamData[i][2]) || 0;
    }
  }
  
  var bestRuns = {};
  for(var j=1; j<scoreData.length; j++) {
    var roundStr = String(scoreData[j][map.round]).toUpperCase(); 
    if(roundStr.indexOf("ROUND 1") > -1 || roundStr.indexOf("ROUND 2") > -1 || roundStr.indexOf("FLEX") > -1) {
      var tId = String(scoreData[j][map.teamId]).trim().toUpperCase(); 
      if(isValidTeam(tId)) {
        var rawTotalPts = parseInt(scoreData[j][map.score]); 
        if(isNaN(rawTotalPts)) rawTotalPts = 0; 
        
        var isDQ = scoreData[j][map.isDq] === true || String(scoreData[j][map.isDq]).toUpperCase() === "TRUE"; 
        var displayPts = isDQ ? 0 : rawTotalPts; 
          
        if(!bestRuns[tId] || displayPts > bestRuns[tId].max) {
           bestRuns[tId] = { max: displayPts, sprint3: sprint3Scores[tId] || 0 };
        } 
      }
    }
  }
  
  var leaderboard = [];
  for(var teamId in teamDict) {
     if(isValidTeam(teamId)) {
         if(bestRuns[teamId]) {
            leaderboard.push({ id: teamId, name: teamDict[teamId], data: bestRuns[teamId] });
         } else {
            leaderboard.push({ id: teamId, name: teamDict[teamId], data: { max: -1, sprint3: sprint3Scores[teamId] || 0 } });
         }
     }
  }
  
  leaderboard.sort(function(a, b) {
    if(b.data.max !== a.data.max) return b.data.max - a.data.max;
    return b.data.sprint3 - a.data.sprint3; 
  });
  
  if (lbSheet.getLastRow() > 1) {
      lbSheet.getRange(2, 1, lbSheet.getLastRow() - 1, 4).clearContent();
  }
  
  var outputData = [];
  for (var rank = 0; rank < leaderboard.length; rank++) {
      var t = leaderboard[rank];
      outputData.push([rank + 1, t.id, t.name, t.data.max]);
  }
  
  if (outputData.length > 0) {
      lbSheet.getRange(2, 1, outputData.length, 4).setValues(outputData);
  }
}

// ==========================================
// CORE SAVING & PLAYOFF ENGINE
// ==========================================
function saveScoresToSheet(payload) {
  // --- PLACEHOLDER RUN INTERCEPTOR ---
  try {
    var pRound = String(payload.round).trim().toUpperCase();
    var pHeat = parseInt(String(payload.heat).replace(/\D/g, ''));
    var pTrack = parseInt(String(payload.track).replace(/\D/g, ''));

    for (var i = 0; i < payload.teams.length; i++) {
      var tColor = String(payload.teams[i].color).toUpperCase();
      if (pRound === "ROUND 1" && pHeat === 5 && pTrack === 2 && tColor === "WHITE") {
        payload.teams.splice(i, 1); i--; 
      }
      if (pRound === "ROUND 2" && pHeat === 1 && pTrack === 1 && tColor === "YELLOW") {
        payload.teams.splice(i, 1); i--;
      }
    }
  } catch(e) {}
  // --- END INTERCEPTOR ---

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName("Scores");
  var headers = logSheet.getDataRange().getValues()[0];
  var timestamp = new Date(); 

  var findCol = function(names) {
    for(var j=0; j<names.length; j++) {
      for (var i=0; i<headers.length; i++) { if (String(headers[i]).toLowerCase().trim() === names[j].toLowerCase().trim()) return i; }
      for (var i=0; i<headers.length; i++) { if (String(headers[i]).toLowerCase().trim().indexOf(names[j].toLowerCase().trim()) > -1) return i; }
    }
    return -1;
  };

  var colMap = {
    time: findCol(["Timestamp", "Time"]), round: findCol(["Round"]), track: findCol(["Track"]), heat: findCol(["Heat"]),
    color: findCol(["Color", "Zone"]), section: findCol(["Section"]), teamId: findCol(["Team ID", "ID"]), name: findCol(["Team Name", "Name"]),
    activate: findCol(["Activate", "Activation"]), leave: findCol(["Leave", "Left"]), mario: findCol(["Mario"]),
    cIn: findCol(["Coins Inner"]), cOut: findCol(["Coins Outer", "Coin Outer"]), sIn: findCol(["Starbits Inner"]), sOut: findCol(["Starbits Outer"]),
    kTop: findCol(["Koopas Toppled", "Toppled"]), kDisp: findCol(["Koopas Displaced", "Displaced"]), lumas: findCol(["Unique Lumas", "Lumas"]),
    score: findCol(["Total Score", "Score"]), dq: findCol(["DQ?", "DQ"]), reason: findCol(["DQ Reason", "Reason"])
  };
  
  var rowsToWrite = []; // HIGH-SPEED BULK SAVE CACHE

  for (var i = 0; i < payload.teams.length; i++) {
    var team = payload.teams[i];
    var newRow = new Array(headers.length).fill("");
    
    if(colMap.time > -1) newRow[colMap.time] = timestamp;
    if(colMap.round > -1) newRow[colMap.round] = String(payload.round);
    if(colMap.track > -1) newRow[colMap.track] = String(payload.track);
    if(colMap.heat > -1) newRow[colMap.heat] = String(payload.heat);
    if(colMap.color > -1) newRow[colMap.color] = String(team.color);
    if(colMap.section > -1) newRow[colMap.section] = String(team.section);
    if(colMap.teamId > -1) newRow[colMap.teamId] = String(team.teamId);
    if(colMap.name > -1) newRow[colMap.name] = String(team.name);
    if(colMap.activate > -1) newRow[colMap.activate] = String(team.activate);
    if(colMap.leave > -1) newRow[colMap.leave] = String(team.leaveStar);
    if(colMap.mario > -1) newRow[colMap.mario] = String(team.marioTier);
    if(colMap.cIn > -1) newRow[colMap.cIn] = Number(team.coinsInner);
    if(colMap.cOut > -1) newRow[colMap.cOut] = Number(team.coinsOuter);
    if(colMap.sIn > -1) newRow[colMap.sIn] = Number(team.starbitsInner);
    if(colMap.sOut > -1) newRow[colMap.sOut] = Number(team.starbitsOuter);
    if(colMap.kTop > -1) newRow[colMap.kTop] = Number(team.koopasToppled);
    if(colMap.kDisp > -1) newRow[colMap.kDisp] = Number(team.koopasDisplaced);
    if(colMap.lumas > -1) newRow[colMap.lumas] = Number(team.lumasUnique);
    if(colMap.score > -1) newRow[colMap.score] = Number(team.score);
    if(colMap.dq > -1) newRow[colMap.dq] = String(team.isDQ);
    if(colMap.reason > -1) newRow[colMap.reason] = String(team.dqReason);

    rowsToWrite.push(newRow); 
  }

  // BULK WRITE INJECTION
  if (rowsToWrite.length > 0) {
    var startRow = Math.max(logSheet.getLastRow() + 1, 2); 
    logSheet.getRange(startRow, 1, rowsToWrite.length, headers.length).setValues(rowsToWrite);
  }

  var targetRoundNum = getRoundNumber(payload.round);
  if (targetRoundNum === 1 || String(payload.round).toUpperCase().includes("FLEX")) {
      updateLiveLeaderboard();
  }

  var targetSheetName = getSheetNameFromRound(payload.round);
  
  if (targetRoundNum >= 4) { 
    var bracketSheet = ss.getSheetByName(targetSheetName);
    if(bracketSheet) {
       var data = bracketSheet.getDataRange().getValues();
       var numTrack = parseInt(String(payload.track).replace(/\D/g, ''));
       var numHeat = parseInt(String(payload.heat).replace(/\D/g, ''));

       for (var r = 1; r < data.length; r++) {
         if (data[r][0] == targetRoundNum && data[r][1] == numHeat && data[r][2] == numTrack) {
           var evalTeams = [];
           for (var t = 0; t < payload.teams.length; t++) {
             var isDQStr = String(payload.teams[t].isDQ).toUpperCase() === "TRUE" || payload.teams[t].isDQ === true;
             var total = parseInt(payload.teams[t].score) || 0;
             var finalScoreDisplay = isDQStr ? -1 : total; 
             
             var tempRow = new Array(21).fill(0);
             tempRow[colMap.cIn] = payload.teams[t].coinsInner;
             tempRow[colMap.cOut] = payload.teams[t].coinsOuter;
             tempRow[colMap.sIn] = payload.teams[t].starbitsInner;
             tempRow[colMap.sOut] = payload.teams[t].starbitsOuter;
             tempRow[colMap.mTier] = payload.teams[t].marioTier;
             tempRow[colMap.kTop] = payload.teams[t].koopasToppled;
             tempRow[colMap.kDisp] = payload.teams[t].koopasDisplaced;
             tempRow[colMap.lumas] = payload.teams[t].lumasUnique;
             
             var weights = isDQStr ? {tier1:-1, tier2:-1, tier3:-1, tier4:-1} : getPlayoffTiebreakerWeights(tempRow, colMap);
             var rawReason = String(payload.teams[t].dqReason).trim();
             var displayStr = isDQStr ? ((rawReason && rawReason !== "undefined" && rawReason !== "") ? "DQ - " + rawReason : "DQ") : total;

             evalTeams.push({ color: payload.teams[t].color, val: finalScoreDisplay, weights: weights, display: displayStr });
           }

           evalTeams.sort(function(a, b) {
              if(b.val !== a.val) return b.val - a.val;
              if(b.weights.tier1 !== a.weights.tier1) return b.weights.tier1 - a.weights.tier1;
              if(b.weights.tier2 !== a.weights.tier2) return b.weights.tier2 - a.weights.tier2;
              if(b.weights.tier3 !== a.weights.tier3) return b.weights.tier3 - a.weights.tier3;
              return b.weights.tier4 - a.weights.tier4;
           });

           var yScore = "", bkScore = "", blScore = "", wScore = "";
           evalTeams.forEach(function(et, index) {
              var isAdvancing = false;
              if (targetRoundNum === 4) isAdvancing = (index === 0);
              else isAdvancing = (index < 2); 
              
              var appendTag = "";
              if (isAdvancing && et.val !== -1) appendTag = (targetRoundNum === 4) ? " 🏆" : " ✅"; 
              var finalStr = et.display + appendTag;
              
              if (et.color == "YELLOW") yScore = finalStr;
              if (et.color == "BLACK") bkScore = finalStr;
              if (et.color == "BLUE") blScore = finalStr;
              if (et.color == "WHITE") wScore = finalStr;
           });
           bracketSheet.getRange(r + 1, 8, 1, 5).setValues([[yScore, bkScore, blScore, wScore, "Complete"]]);
           break; 
         }
       }
    }
  }
  return true;
}

function getCompletedTracks(roundStr, heatStr) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var scoresSheet = ss.getSheetByName("Scores");
  if (!scoresSheet) return [];
  var map = getColMap(scoresSheet);
  var data = scoresSheet.getDataRange().getValues();
  var completed = {};
  var targetRound = String(roundStr).trim().toUpperCase();
  var targetHeatNum = parseInt(String(heatStr).replace(/\D/g, ''));

  for (var i = 1; i < data.length; i++) {
    var rowRound = String(data[i][map.round]).trim().toUpperCase();
    var rowRawHeat = String(data[i][map.heat]).toUpperCase().replace(/HEAT/g, '').trim();
    var rowRawTrack = String(data[i][map.track]).toUpperCase().replace(/TRACK/g, '').replace(/\(COMPLETED\)/g, '').trim();
    var rowHeatNum = parseInt(rowRawHeat); var rowTrackNum = parseInt(rowRawTrack);
    var roundMatch = false;
    if(targetRound.includes("FLEX") && rowRound === targetRound) roundMatch = true;
    else if(!targetRound.includes("FLEX") && !rowRound.includes("FLEX")) {
       var r1 = parseInt(targetRound.replace(/\D/g, ''));
       var r2 = parseInt(rowRound.replace(/\D/g, ''));
       if(r1 === r2) roundMatch = true;
    }
    if (roundMatch && rowHeatNum === targetHeatNum && !isNaN(rowTrackNum)) completed[String(rowTrackNum)] = true; 
  }
  return Object.keys(completed); 
}

function getMatchData(round, heat, track) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var targetSheetName = getSheetNameFromRound(round);
  var matchSheet = ss.getSheetByName(targetSheetName);
  if(!matchSheet) return null; 
  
  var matchData = matchSheet.getDataRange().getValues();
  var teamSheet = ss.getSheetByName("Teams");
  var teamData = teamSheet.getDataRange().getValues();

  var teamDict = {};
  for(var i=1; i<teamData.length; i++) {
    if(teamData[i][0]) teamDict[String(teamData[i][0]).trim().toUpperCase()] = teamData[i][1];
  }

  var targetRoundNum = parseInt(String(round).replace(/\D/g, ''));
  var numTrack = parseInt(String(track).replace(/\D/g, ''));
  var numHeat = parseInt(String(heat).replace(/\D/g, ''));

  for(var j=1; j<matchData.length; j++) {
    var rawSheetRound = String(matchData[j][0]).trim().toUpperCase();
    var sheetRoundNum = parseInt(rawSheetRound.replace(/\D/g, ''));
    var matchRoundOk = false;
    if (targetSheetName === "Matches") { if(sheetRoundNum === targetRoundNum && !rawSheetRound.includes("FLEX")) matchRoundOk = true; } 
    else { if(sheetRoundNum === targetRoundNum) matchRoundOk = true; }

    var sheetHeatNum = parseInt(String(matchData[j][1]).toUpperCase().replace(/HEAT/g, '').trim());
    var sheetTrackNum = parseInt(String(matchData[j][2]).toUpperCase().replace(/TRACK/g, '').trim());

    if(matchRoundOk && sheetHeatNum === numHeat && sheetTrackNum === numTrack) {
       var yId = String(matchData[j][3] || "").trim().toUpperCase(); var bkId = String(matchData[j][4] || "").trim().toUpperCase();
       var blId = String(matchData[j][5] || "").trim().toUpperCase(); var wId = String(matchData[j][6] || "").trim().toUpperCase();
       return {
         yellow: {id: matchData[j][3] || "None", name: teamDict[yId] || "Empty"}, black:  {id: matchData[j][4] || "None", name: teamDict[bkId] || "Empty"},
         blue:   {id: matchData[j][5] || "None", name: teamDict[blId] || "Empty"}, white:  {id: matchData[j][6] || "None", name: teamDict[wId] || "Empty"}
       };
    }
  }
  return null;
}

function generateRoundOf32() {
  updateLiveLeaderboard(); 
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lbSheet = ss.getSheetByName("Live_Leaderboard");
  var r32Sheet = ss.getSheetByName("Round of 32");
  
  var lbData = lbSheet.getDataRange().getValues();
  var top32 = [];
  for(var m=1; m<=32; m++) {
      if (lbData[m] && lbData[m][1]) top32.push(String(lbData[m][1]));
      else top32.push("TBD");
  }
  if(r32Sheet.getLastRow() > 1) r32Sheet.getRange(2, 1, r32Sheet.getLastRow()-1, 12).clearContent();

  var r32Matches = [
    [top32[0], top32[15], top32[16], top32[31]], [top32[1], top32[14], top32[17], top32[30]], 
    [top32[2], top32[13], top32[18], top32[29]], [top32[3], top32[12], top32[19], top32[28]], 
    [top32[4], top32[11], top32[20], top32[27]], [top32[5], top32[10], top32[21], top32[26]], 
    [top32[6], top32[9],  top32[22], top32[25]], [top32[7], top32[8],  top32[23], top32[24]]
  ];
  
  var h32 = 1, t32 = 1;
  for(var k=0; k<r32Matches.length; k++) {
    var group = shuffleArray(r32Matches[k]); 
    r32Sheet.appendRow([32, h32, t32, group[0], group[1], group[2], group[3], "", "", "", "", "Pending"]);
    t32++; if(t32 > 4) { t32 = 1; h32++; }
  }
  SpreadsheetApp.getUi().alert("Round of 32 generated successfully directly from Live Leaderboard!");
}

function advanceToSweet16() { advanceRound("Round of 32", "Sweet 16", 16, 4); }
function advanceToElite8()  { advanceRound("Sweet 16", "Elite 8", 8, 2); }
function advanceToFinal4()  { advanceRound("Elite 8", "Final 4", 4, 1); }

function advanceRound(sourceSheetName, targetSheetName, targetRoundNum, numTargetMatches) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName(sourceSheetName);
  var targetSheet = ss.getSheetByName(targetSheetName);

  var sourceData = sourceSheet.getDataRange().getValues();
  var advancingPool = [];

  for(var j=1; j<sourceData.length; j++) {
    var teamIds = [sourceData[j][3], sourceData[j][4], sourceData[j][5], sourceData[j][6]];
    var teamResults = [String(sourceData[j][7]), String(sourceData[j][8]), String(sourceData[j][9]), String(sourceData[j][10])]; 

    for (var idx = 0; idx < 4; idx++) {
       if (teamResults[idx].includes("✅") || teamResults[idx].includes("🏆")) {
           advancingPool.push(teamIds[idx]);
       }
    }
  }

  var seededTeams = advancingPool;
  var expectedTeams = numTargetMatches * 4;
  while(seededTeams.length < expectedTeams) seededTeams.push("TBD");

  var snakeMatches = [];
  if (targetRoundNum === 16) {
    snakeMatches = [
      [seededTeams[0], seededTeams[7], seededTeams[8], seededTeams[15]], [seededTeams[1], seededTeams[6], seededTeams[9], seededTeams[14]],
      [seededTeams[2], seededTeams[5], seededTeams[10], seededTeams[13]], [seededTeams[3], seededTeams[4], seededTeams[11], seededTeams[12]]
    ];
  } else if (targetRoundNum === 8) {
    snakeMatches = [
      [seededTeams[0], seededTeams[3], seededTeams[4], seededTeams[7]], [seededTeams[1], seededTeams[2], seededTeams[5], seededTeams[6]]
    ];
  } else if (targetRoundNum === 4) { snakeMatches = [ [seededTeams[0], seededTeams[1], seededTeams[2], seededTeams[3]] ]; }

  if(targetSheet.getLastRow() > 1) targetSheet.getRange(2, 1, targetSheet.getLastRow()-1, 12).clearContent();
  var h = 1, t = 1;
  for(var m=0; m<numTargetMatches; m++) {
    var group = shuffleArray(snakeMatches[m]); 
    targetSheet.appendRow([targetRoundNum, h, t, group[0], group[1], group[2], group[3], "", "", "", "", "Pending"]);
    t++; if(t > 4) { t = 1; h++; }
  }
  SpreadsheetApp.getUi().alert("Winners Advanced successfully!");
}

function clearAllData() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('CRITICAL WARNING', 'This will wipe all submitted scores. Are you sure?', ui.ButtonSet.YES_NO);

  if (response == ui.Button.YES) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var scoresSheet = ss.getSheetByName("Scores");
    if (scoresSheet.getLastRow() > 1) scoresSheet.getRange(2, 1, scoresSheet.getLastRow() - 1, 21).clearContent(); 

    var elimTabs = ["Round of 32", "Sweet 16", "Elite 8", "Final 4"];
    elimTabs.forEach(function(tabName) {
      var sheet = ss.getSheetByName(tabName);
      if (sheet && sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, 12).clearContent();
    });

    var matchesSheet = ss.getSheetByName("Matches");
    if (matchesSheet.getLastRow() > 1) matchesSheet.getRange(2, 8, matchesSheet.getLastRow() - 1, 1).setValue("Pending");
    
    var lbSheet = ss.getSheetByName("Live_Leaderboard");
    if (lbSheet && lbSheet.getLastRow() > 1) lbSheet.getRange(2, 1, lbSheet.getLastRow() - 1, 4).clearContent();
    
    updateLiveLeaderboard(); 

    ui.alert("System Reset Complete.");
  }
}