function myFunction() {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("AGT法人管理")
  for(var row = 2 ; row < 75 ; row ++){
    var id= ss.getRange(row,11).getValue()
    if(!id){continue}
    var name = SpreadsheetApp.openById(id).getName()
    ss.getRange(row,12).setValue(name)
  }
}


function debugSearchedEmails() {
  var q = '"' + CONFIG_CLIENTSYNC.SHARE_SUBJECT_KEYWORD + '"';
  var threads = GmailApp.search(q, 0, CONFIG_CLIENTSYNC.GMAIL_SEARCH_MAX);
  var rows = [['subject', 'from', 'fileId']];
  threads.forEach(function(t){
    t.getMessages().forEach(function(m){
      var idMatch = (m.getBody()||'').match(/spreadsheets\/d\/([a-zA-Z0-9_\-]+)/);
      rows.push([m.getSubject(), m.getFrom(), idMatch ? idMatch[1] : '']);
    });
  });
  Logger.log(rows.map(function(r){ return r.join(' | '); }).join('\n'));
}