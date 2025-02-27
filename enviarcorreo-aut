// @ts-nocheck
const yellowPages = {
{
  "O D SAS": {
    "email": "email1@example.com",
    "number": "17wWvvAAsS95ubOrXJKUC-aLY1L41SI_U"
  },
  "LCA SAS": {
    "email": "email2@example.com",
    "number": "1YBCHPfXWxH5HR5_7Jm98SordD5HeRakk"
  },
  "A SRL": {
    "email": ["email3@example.com", "email4@example.com"],
    "number": "1FJJl_PIWXERR8pMB002mmEyI2-6OuaDO"
  },
  "CS SRL": {
    "email": "email5@example.com",
    "number": "1gaf4DNtbkt1e_oQUsPG2XKmzQ5VnKrBx"
  },
  "CD SRL": {
    "email": "email6@example.com",
    "number": "1cZfapIzlsvFk46MrzLExEEX9A0RKamJd"
  },
  "CUD SRL": {
    "email": "email7@example.com",
    "number": "1kLnlUMCn8DRbBNOdkaq9nLw5uTNDUcku"
  },
  "DA SRL": {
    "email": "email8@example.com",
    "number": "1ISK0RGkALQs_y1zvAmW7WejRMZbvp9kJ"
  },
  "BR SRL": {
    "email": "email9@example.com",
    "number": "14fC0o84TfqKIDSYhe4iFVDJ59eceeL68"
  },
  "MN": {
    "email": "email10@example.com",
    "number": "13mlor8JTTIFnmTwzh07HqlFM5h1pSsZ4"
  },
  "SA SRL": {
    "email": ["email11@example.com", "email12@example.com"],
    "number": "1F0AIqIFVxrf6g-08I9rF5QEwe-I4-R5m"
  },
  "NSA SAS": {
    "email": "email13@example.com",
    "number": "1aWwcUI-oOBsvFhMENdqm0qNizHkO0epQ"
  },
  "MIL M": {
    "email": "email14@example.com",
    "number": "1dAFPbH68j7Zfam2OrtqP2WWIl55L_pdB"
  },
  "SL LG SRL": {
    "email": "email15@example.com",
    "number": "1nYZjeqNspuxKwWiyiFBtY6T4bagNbKId"
  },
  "SO SRL": {
    "email": "email16@example.com",
    "number": "180yVvLO6bFlpX1GggdQ54Sc47T7b5MnS"
  },
  "CM": {
    "email": "email17@example.com",
    "number": "136vj_SYN1FZ4_QLAyLvtmfg7gNniO5IG"
  },
  "BD": {
    "email": "email18@example.com",
    "number": "1DZDZfvfwF2LFUKNC6IvEvmzUu_rznvHD"
  },
  "SR.L.": {
    "email": "email19@example.com",
    "number": "1bBSSJWFX4-zoMEFxqh5VfltpdrCB4UsB"
  },
  "MBBA": {
    "email": "email20@example.com",
    "number": "1t-Sm-trlhfuFwuZLoUWyu0ClbpZwrnxx"
  }
}


function copyandpaste() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var ss = sheet.getActiveSheet()
  //var ss = SpreadsheetApp.setActiveSheet(sheet.getSheets()[1])
  var headers = ss.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
  var headerRange = ss.getSheetValues(1, 1, -1, -1);
  var prestS = sheet.getSheetName().split("-");
  const subject = ss.getSheetName()
  var lastR = ss.getLastRow()
  var data = headers;

var wPrest = subject.split("-");
var pName = wPrest[1].trim();

Logger.log(wPrest)
  var spreadsheet = SpreadsheetApp.create(subject);
  var tempSheet = spreadsheet.getActiveSheet();

var spreadsheetId = spreadsheet.getId()

  tempSheet.getRange(1, 1, data.length, data[0].length).setValues(data);

  headerRange.forEach(function(x){
     var tmprango = headerRange.indexOf(x)+1;
   
      if(x[8]!= "INFO EXTRA" && !!x[8] === true){
      tempSheet.getRange("H"+tmprango+":K"+tmprango).activate()
      tempSheet.getActiveRange().setBackground("red")
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
     } else if(x[19]!= "UGL OK?" && x[19]!= "VERDADERO"){
        tempSheet.getRange("H"+tmprango+":K"+tmprango).activate()
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
  tempSheet.getRange('A1:U'+lastR).activate();
  tempSheet.getActiveRangeList().setFontWeight('bold')
  .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  tempSheet.getActiveRangeList().setHorizontalAlignment('center')
  .setVerticalAlignment('middle');
  tempSheet.setColumnWidth(8, 200);
  tempSheet.setColumnWidth(13, 200);
  tempSheet.getRange('E2:E').activate();
  tempSheet.getActiveRangeList().setNumberFormat('[h]:mm:ss');


moveFileToFolder(spreadsheetId, yellowPages[pName].number)

  return sendMail(spreadsheetId, subject, yellowPages[pName].email)
}

function sendMail(id, subj, recipient){
  

  var recipiente = recipient
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

//  SpreadsheetApp.getActiveSpreadsheet().deleteFile(spreadsheet);
//  Drive.Files.remove(fileId);
}


function moveFileToFolder(fileId, folderId) {

    var file = DriveApp.getFileById(fileId);
  var folder = DriveApp.getFolderById(folderId);
  
  file.moveTo(folder)
}




