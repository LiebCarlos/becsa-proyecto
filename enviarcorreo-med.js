const foldersId = {
  "DICCEA": {
    "email": "email1@example.com",
    "carpeta": "14rDnhsKdFwwSc4oVTiK0zCVrVfNho0x8"
  },
  "NH": {
    "email": "email2@example.com",
    "carpeta": "14rDnhsKdFwwSc4oVTiK0zCVrVfNho0x8"
  },
  "EMED": { 
    "email": "email3@example.com",
    "carpeta": "1fjO3bq2EtS-iD4h_qs0N0WBvyQ6jlE70"
  },
  "BECSA": {
    "email": "email4@example.com",
    "carpeta": "1Z1Fl4H__gVPwsjl7Ue4zHmvlH_-iVAU2"
  },
  "CM": {
    "email": "email5@example.com",
    "carpeta": "1Z1Fl4H__gVPwsjl7Ue4zHmvlH_-iVAU2"
  }  
}

function copyandpaste() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var ss = sheet.getActiveSheet()
  //var ss = SpreadsheetApp.setActiveSheet(sheet.getSheets()[12])
  var headers = ss.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
  var headerRange = ss.getSheetValues(1, 1, -1, -1);
  var folderId = "1fjO3bq2EtS-iD4h_qs0N0WBvyQ6jlE70";
  var folder = DriveApp.getFolderById(folderId);
  var prestS = sheet.getSheetName().split("-");
  const subject = ss.getSheetName()
  var lastR = ss.getLastRow()
  var data = headers;

var wPrest = subject.split("-");
var pName = wPrest[1].trim()
  var spreadsheet = SpreadsheetApp.create(subject);
  var tempSheet = spreadsheet.getActiveSheet();

var spreadsheetId = spreadsheet.getId()

  tempSheet.getRange(1, 1, data.length, data[0].length).setValues(data);

  headerRange.forEach(function(x){

    if(x[15]!= "Paciente vigente\nUGL correcta")
      if(x[15]!= "Resultado"){
      var tmprango = headerRange.indexOf(x)+1;
      tempSheet.getRange("P"+tmprango).activate()
      tempSheet.getActiveRange().setBackground("red")
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
    }
  })

  tempSheet.hideColumns(4);
  tempSheet.setColumnWidth(3, 146);
  tempSheet.autoResizeColumns(3, 1);
  tempSheet.autoResizeColumns(7, 1);
  tempSheet.autoResizeColumns(8, 1);
  tempSheet.setColumnWidth(10, 98);
  tempSheet.setColumnWidth(16, 125);
  tempSheet.autoResizeColumns(10, 1);
  tempSheet.setRowHeights(1, tempSheet.getLastRow(), 35);
  tempSheet.setFrozenRows(1);
  tempSheet.sort(8, true);
  tempSheet.getRange('A1:P'+lastR).activate();
  tempSheet.getActiveRangeList().setFontWeight('bold')
  .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  tempSheet.getActiveRangeList().setHorizontalAlignment('center')
  .setVerticalAlignment('middle');
  tempSheet.getRange('E2:E').activate();
  tempSheet.getActiveRangeList().setNumberFormat('[h]:mm:ss');

moveFileToFolder(spreadsheetId, foldersId[pName].carpeta)



  return sendMail(spreadsheetId, subject, foldersId[pName].email)
}

function sendMail(id, subj, recipiente){
  var recipient = recipiente
  var subject = subj;
  var attachmentName = subj;
  var fileUrl = "https://drive.google.com/open?id=" + id;
  var body = `Buenos días,\n\nLes enviamos en adjunto las transmisiones realizadas el/los día/s mencionado/s. \n\n\nSaludos.\n\n ${fileUrl}`;

 if(typeof recipiente === "string"){
MailApp.sendEmail({
    to: recipiente,
    subject: subject,
    body: body
      })
} else {

  recipiente.forEach(function(x){
     MailApp.sendEmail({
    to: x,
    subject: subject,
    body: body
      })
    })

}

  // Delete the temporary spreadsheet and Excel file
//  SpreadsheetApp.getActiveSpreadsheet().deleteFile(spreadsheet);
//  Drive.Files.remove(fileId);
}


function moveFileToFolder(fileId, folderId) {

    var file = DriveApp.getFileById(fileId);
  var folder = DriveApp.getFolderById(folderId);
  
  file.moveTo(folder)
}
