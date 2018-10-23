function eliminarDesdeDistintaTabla2 () {
   var hojaA = SpreadsheetApp.getActive().getActiveSheet();
   //var titleRows = Browser.inputBox("Por favor, introduce numero de filas a saltar(encabezados):");
   //var hojaACol = Browser.inputBox("Por favor, indica la columna donde se encuentran los datos de trabajo:");
   var hojaB = Browser.inputBox("Por favor, introduce el nombre de la hoja con los datos de referencia:");
   hojaB = SpreadsheetApp.getActive().getSheetByName(hojaB);
   if (!hojaB){ // Validación de la hojaB, si no existe, abortamos
      Browser.msgBox("La hoja especificada no existe");
     return
   } else {
     var columnasDeA=[];
     var seleccion = hojaA.getSelection().getActiveRangeList().getRanges().map(function (e){
       return e.getA1Notation();
     });
     
     seleccion.forEach(function(e){
       var rango = e.split(',');
       rango.forEach(function(elem){
         elem.split(':').forEach(function(element){
           var col = element.charCodeAt(0)-64;
           if(columnasDeA.indexOf(col) < 0){
             columnasDeA.push(col);
           }
         })
       })      
     });
     
     eliminarDuplicados2Hojas(hojaB, columnasDeA, 1, hojaA);
   }  
 }

/********************
*********************/

function eliminarDuplicados2Hojas(hojaB, columnasDeA, jumpRowsA, hojaA){
  
  //Declaración de variables
    //Datos de trabajo
  var dataA;
  var dataB;
    //
  var newSheetData=[];
  
  dataA = hojaA.getDataRange().getValues();
  dataB = hojaB.getDataRange().getValues().map(function (e){
    return e.filter(function (elem,index){
      return columnasDeA.indexOf(index+1) >= 0;
    });
  });
    
  if(jumpRowsA > 0){
    dataA = dataA.slice(jumpRowsA);
    dataB = dataB.slice(jumpRowsA);
    
    for(var i=0; i < jumpRowsA;i++){
       newSheetData.push(hojaA.getDataRange().getValues()[i]);
    }
  }
  
  
  /*
  Realizamos un filtrado de los datos de A, 
  de modo que sólo nos quedamos con aquellos que no existan en B.
  */
  
  
  for (var i=0; i < dataA.length;i++){
    var encontrado = false;
    var refATemp = dataA[i].filter(function (elem, index){
      return columnasDeA.indexOf(index+1) >= 0;
    });
    
    for(var j=0; j < dataB.length && !encontrado;j++){      
      if(refATemp.toString() === dataB[j].toString()){
        encontrado = true;
      }
    }
    if(!encontrado){
      newSheetData.push(dataA[i])
    }
  }
  
  
    
  //Si todo ha ido bien, Insertamos los datos generados en una nueva hoja.
  if (newSheetData){
    hojaA.getParent().insertSheet(hojaA.getName()+"_SIN_DUPLICADOS").getRange(1, 1, newSheetData.length, newSheetData[0].length).setValues(newSheetData);
  }
} 