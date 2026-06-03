var TARGET_SHEET_ID = 'YOUR_SHEET_ID_HERE'; 

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ME 2110 Live Broadcast')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1'); 
}

function getBroadcastLeaderboard() {
  var payload = { leaderboard: [] };
  try {
    var ss = SpreadsheetApp.openById(TARGET_SHEET_ID);
    var lbSheet = ss.getSheetByName("Live_Leaderboard");
    
    if (lbSheet && lbSheet.getLastRow() > 1) {
      var lbData = lbSheet.getDataRange().getValues();
      for (var i = 1; i < lbData.length; i++) {
         if (lbData[i][1]) {
             payload.leaderboard.push({ rank: lbData[i][0], id: String(lbData[i][1]), name: String(lbData[i][2]), data: { max: lbData[i][3] } });
         }
      }
    }
  } catch (error) { payload.error = String(error.message); }
  return JSON.stringify(payload);
}