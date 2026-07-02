function doGet(e) {
  var params = (e && e.parameter) ? e.parameter : {};
  var callback = params.callback ? String(params.callback) : 'callback';
  var team = params.team ? String(params.team).trim() : '';
  var result = claimLink(team);
  var open = String.fromCharCode(40);
  var close = String.fromCharCode(41);
  var body = callback + open + JSON.stringify(result) + close;
  var output = ContentService.createTextOutput(body);
  output.setMimeType(ContentService.MimeType.JAVASCRIPT);
  return output;
}

function claimLink(team) {
  if (!team) {
    return { status: 'error', message: 'Nom d equipe manquant' };
  }
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    return { status: 'busy' };
  }
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    var last = sheet.getLastRow();
    if (last < 2) {
      return { status: 'soldout' };
    }
    var values = sheet.getRange(2, 1, last - 1, 2).getValues();
    var key = team.toLowerCase();
    for (var i = 0; i < values.length; i++) {
      var link = String(values[i][0]).trim();
      var name = String(values[i][1]).trim();
      if (link && name && name.toLowerCase() === key) {
        return { status: 'ok', link: link, reused: true };
      }
    }
    for (var j = 0; j < values.length; j++) {
      var free = String(values[j][0]).trim();
      var taken = String(values[j][1]).trim();
      if (free && !taken) {
        var row = j + 2;
        sheet.getRange(row, 2).setValue(team);
        sheet.getRange(row, 3).setValue(Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy HH:mm:ss'));
        return { status: 'ok', link: free, reused: false };
      }
    }
    return { status: 'soldout' };
  } finally {
    lock.releaseLock();
  }
}
