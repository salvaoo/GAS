
 function eliminarDesdeDistintaTabla () {
   var hojaA = SpreadsheetApp.getActive().getActiveSheet();
   //var titleRows = Browser.inputBox("Por favor, introduce numero de filas a saltar(encabezados):");
   var hojaACol = Browser.inputBox("Por favor, indica la columna donde se encuentran los datos de trabajo:");
   var hojaB = Browser.inputBox("Por favor, introduce el nombre de la hoja con los datos de referencia:");
   hojaB = SpreadsheetApp.getActive().getSheetByName(hojaB);
   if (!hojaB){ // Validación de la hojaB, si no existe, abortamos
      Browser.msgBox("La hoja especificada no existe");
      
   } else {
     var hojaBCol = Browser.inputBox("Por favor, indica la columna donde se encuentran los datos de referencia:");
     
     rmDuplicadosDistintaHoja(hojaB, Number(hojaBCol), Number(hojaACol), 0, hojaA);
   }  
 }

/********************
*********************/

function rmDuplicadosDistintaHoja(hojaB, hojaBCol, hojaACol, jumpRowsA, hojaA){
  
  //Declaración de variables
    //Datos de trabajo
  var dataA = hojaA.getRange(1+jumpRowsA, 1, hojaA.getMaxRows()-jumpRowsA, hojaA.getMaxColumns()).getValues();
  
  var dataB = hojaB.getRange(1, 1, hojaB.getMaxRows(), 1).getValues().map(function (e){
    return e.toString()
  });
  
  
  
  /*
  Realizamos un filtrado de los datos de A, 
  de modo que sólo nos quedamos con aquellos que no existan en B.
  */
  newSheetData = dataA.filter(function (e){
    
  })
    
  //Si todo ha ido bien, Insertamos los datos generados en una nueva hoja.
  if (newSheetData){
    hojaA.getParent().insertSheet(hojaA.getName()+"_UPDATED_FS").getRange(1, 1, newSheetData.length, newSheetData[0].length).setValues(newSheetData);
  }
} 