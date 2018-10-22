//Aplicación para probar la extensión de GAS con GitHub
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // Or DocumentApp or FormApp. -> Creamos el menu con submenu
  ui.createMenu('Identificar Duplicados')
      .addItem('Hoja completa', 'TodaLaHoja')
      .addItem('Seleccionar columnas', 'seleccionColumnas')
      .addItem('Eliminar duplicados desde tabla', 'eliminarDesdeDistintaTabla')
      .addToUi();
}

/**
 * Busca duplicado en toda la hoja comparando por filas.
 */
function TodaLaHoja() {
  var sheet = SpreadsheetApp.getActiveSheet();
  
  //Creamos nueva hoja para alojar los datos sin duplicados.
  var newSheet = sheet.getParent().insertSheet(sheet.getName()+"Sin duplicados");
  
  Logger.log(sheet);
  var data = sheet.getDataRange().getValues();
  Logger.log(data);
  var newData = [];
  for (i in data) {
    var row = data[i];
    Logger.log(row);
    var duplicate = false;
    
    for (j in newData) {
      Logger.log(j);
      if (row.toString() == newData[j].toString()) {
        duplicate = true;
      }
    }
    if (!duplicate) {
      newData.push(row);
    }
  }
  
  newSheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
 //sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
}


function seleccionColumnas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  /*
  var filas = ss.getSelection().getActiveRangeList().getRanges().map(function(e){
    return e.getRow();
  });
  var filas = ss.getActiveRangeList().getRanges().map(function(e){
    return e.getRow();
  });
  
  ss.getSelection().getActiveRangeList().getRanges().forEach(function (e){
  Logger.log(e.getA1Notation());
  });
  */
  
   var selec = ss.getActiveRangeList().getRanges().map(function(e){
    return e.getA1Notation();
  });
  Logger.log(selec);
  var columnas = [];
  
  for (var i = 0; i < selec.length; i++) {
    var col = selec[i].split(":");
    
    for (var p = 0; p < col.length; p++) {
      
      if (columnas.indexOf(col[p].substring(0,1)) < 0) {
        columnas.push(col[p].substring(0,1));
        Logger.log(columnas);
      }
    }   
  }
  
  //Creamos nueva hoja para alojar los datos sin duplicados.
  var newSheet = sheet.getParent().insertSheet(sheet.getName()+"SD_por columnas");

}


