//tst 1 14segs/it 

var a1Q1 = ["N° de Visita"," Estado","Fecha comienzo","Fecha fin", "Duración", "Celular persona afiliada", "N° afiliación", "Nombre persona afiliada"," DNI del responsable de la visita", "Responsable de la visita"," N° matrícula responsable de la visita"," Tipo servicio prestador"," UGLtr"," UGLc"," Prestador"," Resultado", , ," Resultado" ," status",]

function extraerPrestadores2(){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("CONTROL")
  var mySet1 = new Set(sheet.getRange("O2:O").getValues().flat());
  var comSet = Array.from(mySet1.values()).filter(x => x != "")
  comSet.push("BECSA")
  spreadsheet.getActiveSheet().sort(15, true);

  return comSet
}

function creadorHoja2(datos) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const [prestador, rango] = datos;
  const diaHoy = Utilities.formatDate(new Date(), "GMT-3", "dd.MM");
  const controlSheet = spreadsheet.getSheetByName('CONTROL');
  const rangeToCopy = controlSheet.getRange(rango);
  const newRowHeight = 35;
  const lastRow = rangeToCopy.getLastRow() - 1;
  const newSheet = spreadsheet.insertSheet(`Control diario de prestadores - ${prestador} - ${diaHoy}`);
  const newRange = newSheet.getRange('A1:Q1');


  newRange.copyTo(newSheet.getRange('A1:Q1'), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  rangeToCopy.copyValuesToRange(newSheet, 1, 16, 2, lastRow + 1);
  rangeToCopy.copyFormatToRange(newSheet, 1, 16, 2, lastRow + 1);

  newSheet.hideColumns(4);
  newSheet.setFrozenRows(1);
  newSheet.sort({column: 8, ascending: false});

  newSheet.setColumnWidth(3, 146);
  newSheet.setColumnWidth(10, 98);
  newSheet.setColumnWidth(16, 125);
  newSheet.getRange('C2:C'+(lastRow+1)).setNumberFormat('#,##0.00');
  newSheet.getRange('G2:G'+(lastRow+1)).setNumberFormat('#,##0.00');
  newSheet.getRange('H2:H'+(lastRow+1)).setNumberFormat('#,##0.00');
  newSheet.getRange('J2:J'+(lastRow+1)).setNumberFormat('#,##0.00');
  newSheet.getRange('P2:P'+(lastRow+1)).setNumberFormat('dd/MM/yyyy');
  newSheet.getRange('A1:P'+lastRow).setVerticalAlignment('middle');
  newSheet.getRange('A1:P'+lastRow).setRowHeight(newRowHeight);

  return;
}


function extraerRango2(prestador){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("CONTROL")
  var rngCtrl = sheet.getSheetValues(1, 1, -1, -1)  
  var tempRng = [];
  var ultimaCol = sheet.getLastRow()

  if(prestador!="BECSA"){
  rngCtrl.forEach(function(x){
      if(x[14] === prestador){
      tempRng.push(rngCtrl.indexOf(x)+1)
      }
  })
  } else {
      tempRng.push(2)
      tempRng.push(ultimaCol)
    }

  return [prestador, `A${tempRng[0]}:P${tempRng.pop()}`]
}

function generarHojas2() {
  var prestadores = extraerPrestadores2();
  const rngPrestadores = prestadores.map(extraerRango2);

  const promises = rngPrestadores.map(creadorHoja);

  return Promise.all(promises);
}

