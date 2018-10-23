function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // Or DocumentApp or FormApp. -> Creamos el menu con submenu
  ui.createMenu('Identificar Duplicados')
      .addItem('Hoja completa', 'TodaLaHoja')
      .addItem('Seleccionar columnas', 'seleccionColumnas')
      .addItem('Eliminar duplicados desde tabla', 'eliminarDesdeDistintaTabla2')
      .addToUi();
}


function TodaLaHoja() {
  var sheet = SpreadsheetApp.getActiveSheet();
  
  //Creamos nueva hoja para alojar los datos sin duplicados.
  var newSheet = sheet.getParent().insertSheet(sheet.getName()+" - sin duplicados");

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
  var sheet = SpreadsheetApp.getActiveSheet();
  var selec = sheet.getActiveRangeList().getRanges().map(function(e){
    return e.getA1Notation();
  });
  var columnas = [];
  Logger.log(selec);
  /** FOR para sacar las columnas seleccionadas de la hoja. **/
  for (var i = 0; i < selec.length; i++) {
    var col = selec[i].split(":");  
    for (var p = 0; p < col.length; p++) { //Este FOR es para sacar la letra referente a la columna seleccionada.
      if (columnas.indexOf(col[p].substring(0,1)) < 0) { 
        columnas.push(col[p].substring(0,1));
        Logger.log(columnas);
      }
    }   
  }
  
  var numeroColumnas = [];
  Logger.log(columnas);
  
  /** FOR para recorrer el Array "columnas" que contiene las letras de las columnas seleccionadas y sacar el numero correspondiente a esas letras. **/
  for (var cont = 0; cont < columnas.length; cont++) {
    var numCol = columnas[cont].toString().charCodeAt(0); //Nos devuelve el codigo ASCII de la letra de las columnas seleccionadas.
    var resta = (parseInt(numCol) - 64).toPrecision(1); //Aqui sacamos el numero de columna correspondiente a la letra de las columnas seleccionadas.
    numeroColumnas.push(resta); //Añadimos en "numeroColumnas" el resultado de la operacion anterior que es el numero de las columnas.
    Logger.log(numeroColumnas);
  }
  
  var data = sheet.get; //Aqui tenemos que decirle que trabaje con la columnas que tenemos en el array columnas.
  Logger.log(data);
  
  var newData = [];
  for (i in data) {
    var row = data[i];
    Logger.log(row);
    var duplicate = false;
    
    for (j in newData) {
      if (row.toString() == newData[j].toString()) {
        duplicate = true;
      }
    }
    if (!duplicate) {
      newData.push(row);
    }
  }

}


