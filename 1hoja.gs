/************************

 APP ELIMINAR DUPLICADOS

************************/

/*********************
  Funcion Interfaz
*********************/
/*
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // Or DocumentApp or FormApp. -> Creamos el menu con submenu
  ui.createMenu('Identificar Duplicados')
      .addItem('Hoja completa', 'TodaLaHoja')
      .addItem('Eliminar duplicados misma tabla', 'seleccionColumnas')
      .addItem('Eliminar duplicados desde tabla', 'eliminarDesdeDistintaTabla2')
      .addToUi();
}
*/

function onOpen() {
  SpreadsheetApp.getUi()
    .createAddonMenu()  //Add a new option in the Google Docs Add-ons Menu
    .addItem("Buscar Duplicados", "abrimenu")
    .addToUi();  //Run the showSidebar function when someone clicks the menu
}

function abrimenu() {
  var html = HtmlService.createTemplateFromFile("duplicados")
    .evaluate()
    .setTitle("Gestión de duplicados");  //The title shows in the sidebar
  SpreadsheetApp.getUi().showSidebar(html);
}

/**************************************************************
  Funcion que elimina duplicados para un mismo documento,
  utilizando como referencia todas las columnas de cada fila.
**************************************************************/
function TodaLaHoja() {
  //Definicion de variables.
  //Hoja activa de trabajo.
  var sheet = SpreadsheetApp.getActiveSheet();
  //Creamos nueva hoja para alojar los datos sin duplicados.
  var newSheet = sheet.getParent().insertSheet(sheet.getName() + " - sin duplicados");
  //Datos de trabajo.
  var data = sheet.getDataRange().getValues();
  //Nueva tabla sin duplicados.
  var newData = [];
  /**
  * Recorremos todas las filas de "data",para cada fila comparamos con
  * todas las filas de newData,si la fila actual ya existe en newData
  * no la incorporamos.
  */
  for (i in data) {
    var row = data[i];  //Fila actual
    var duplicate = false;  //Control de duplicados
    for (j in newData) {
      if (row.toString() == newData[j].toString()) {  //Si la fila existe --> está duplicado.
        duplicate = true;
      }
    }
    if (!duplicate) {  //Mientras no sea duplicado, lo incorporamos a la nueva tabla.
      newData.push(row);
    }
  }
  //Insertamos los nuevos datos sin los duplicados en la nueva hoja.
  newSheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
}

/********************************************************************
  Funcion que permite seleccionar las columnas que se tomarán como
  referencia para establecer se una fila está duplicada.
********************************************************************/
function seleccionColumnas() {
  //Definicion de variables.
  //Hoja activa.
  var sheet = SpreadsheetApp.getActiveSheet();
  //Rango (Array) de valores seleccionados en notación A1.
  var selec = sheet.getActiveRangeList().getRanges().map(function(e) {
    return e.getA1Notation();
  });
  //Preparamos el array que contrendrá los valores de las columnas seleccionadas.
  var columnas = [];
  /**
  *  Recorremos selec para obtener solo información referente a las columnas y
  *  descartando aquellas que estén repetidas. De este modo obtenemos todas
  *  las columnas implicadas en la seleccion, indistintamente de como estén seleccionadas
  */
  for (var i = 0; i < selec.length; i++) {
    var col = selec[i].split(":");
    //FOR para sacar la letra referente a la columna seleccionada.
    for (var p = 0; p < col.length; p++) {
      if (columnas.indexOf(col[p].substring(0, 1)) < 0) {
        columnas.push(col[p].substring(0, 1));
      }
    }
  }
  //variable numeroColumnas contine las columnas seleccionadas en valores numericos.
  var numeroColumnas = [];
  /**
  *  FOR para recorrer el Array "columnas" que contiene las letras de las columnas
  *  seleccionadas y sacar el numero correspondiente a esas letras.
  */
  for (var cont = 0; cont < columnas.length; cont++) {
    //Nos devuelve el codigo ASCII de la letra de las columnas seleccionadas.
    var numCol = columnas[cont].toString().charCodeAt(0);
    //Aqui sacamos el numero de columna correspondiente a la letra de las columnas seleccionadas.
    var resta = parseInt(numCol) - 64;
    //Añadimos en "numeroColumnas" el resultado de la operacion anterior que es el numero de las columnas.
    numeroColumnas.push(resta);
  }
  //Aqui guardamos en la variable data los valores de las celdas.
  var data = sheet.getDataRange().getValues();
  //Creamos nueva variable newData.
  var newData = [];
  for (i in data) {
    var duplicados = false; //Inicializamos la variable duplicados a false.
    var row = data[i]; //Toma la primera fila.
    //FOR que entramos si es menor j que el valor de newData y duplicados es igual a false.
    for (var j = 0; j < newData.length && !duplicados; j++) {
      var sonIguales = true; //Inicializamos una variable sonIguales a true.
      for (var p = 0; p < numeroColumnas.length && sonIguales; p++) {
        //SI para comprobar si las filas de las columnas seleccionadas son diferentes.
        if (row[numeroColumnas[p] - 1].toString() != newData[j][numeroColumnas[p] - 1].toString()) {
          sonIguales = false;
        }
      }
      //Aqui le damos el valor de resultado de sonIguales a duplicados.
      duplicados = sonIguales;
    }
    //Si no esta duplicado lo añade en newData.
    if (!duplicados) {
      newData.push(row);
    }
  }
  //Creamos una nueva hoja y pasamos los datos de newData.
  var newSheet = sheet.getParent().insertSheet(sheet.getName() + " - sin duplicados_col");
  newSheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
}

/**************************************************************************************

  DUPLICADOS 2 HOJAS:
  Con esta funcion podemos eliminar aquellas filas de la hoja de trabajo que estén
  presente en una segunda hoja. Los valores de referencia(columnas) para determinar
  si una fila es duplicada o no, se tomarán de la selección previa en la propia hoja.

**************************************************************************************/

/********************************************************************************
  Funcion de control que solicita el input del usuario, comprueba las columnas
  seleccionadas y ejecuta la funcion para eliminar duplicados
********************************************************************************/
function eliminarDesdeDistintaTabla2(nombreHoja) {
  //Declaración de variables.
  //Hoja activa de trabajo.
  var hojaA = SpreadsheetApp.getActive().getActiveSheet();
  //Representa el nº de filas que saltaremos, representan los encabezados. Por defecto 1.
  var nFilasEncabezado = 1;
  //Hoja con los datos de referencia.
  //var hojaB = Browser.inputBox("Por favor, introduce el nombre de la hoja con los datos de referencia:");
  var hojaB = SpreadsheetApp.getActive().getSheetByName(nombreHoja);
  if (!hojaB) {  // Validación de la hojaB, si no existe, abortamos
    Browser.msgBox("La hoja especificada no existe");
    return
  }else {
    //Estructura donde guardaremos las columnas implicadas en la selección.
    var columnasDeA = [];
    //Rango(Array) de las celdas seleccionadas.
    var seleccion = hojaA.getSelection().getActiveRangeList().getRanges().map(function(e) {
      return e.getA1Notation();
    });
    /**
    *  Recorremos el array seleccion, quedandonos solo con los datos de las columnas
    *  implicadas, luego convertimos el valor de la columna de caracter a nº de columna
    *  mediante la conversion ASCII.
    */
    seleccion.forEach(function(e) {  //Ej. seleccion -> [A1:A5, C2]
      var rango = e.split(',');
      //Para cada rango (en caso de selección múltiple no consecutiva). Ej. rango[0]-> A1:A5
      rango.forEach(function(elem) {
        //Para cada celda (ColumnaFila), miramos qué letra(columna) es. Ej. elem[0]-> A1, elem[1]->A5
        elem.split(':').forEach(function(element) {
          /**
          *  Sacamos de elem[0] el primer caracter (la columna) y la convertimos
          *  en su nº de columna.  Ej. col = A
          */
          var col = element.charCodeAt(0) - 64;
          //Si en nuestro array de columnas, no tenemos ya esta columna, la agregamos al array.
          if (columnasDeA.indexOf(col) < 0) {
            columnasDeA.push(col);
          }
        })
      })
    });
    //Ejecutamos nuestra funcion con los datos generados.
    eliminarDuplicados2Hojas(hojaB, columnasDeA, hojaA, nFilasEncabezado);
  }
}


function eliminarDuplicados2Hojas(hojaB, columnasDeA, hojaA, nFilasEncabezado) {

  //Declaración de variables.
  //Datos de trabajo.
  var dataA;
  var dataB;
  //Preparamos la estructira que contendrá la nueva tabla sin duplicados.
  var newSheetData = [];
  //Cargamos los datos de trabajo en RAM.
  dataA = hojaA.getDataRange().getValues();
  //En este caso sólo incorporamos las columnas implicadas en la selección
  dataB = hojaB.getDataRange().getValues().map(function(e) {
    return e.filter(function(elem, index) { //Es decir, las columnas que existan en columnasDeA
      return columnasDeA.indexOf(index + 1) >= 0;
    });
  });
  /**
  *  Si existe(n) encabezado(s),los sacamos de nuestras tablas de trabajo y
  *  los incorporamos este lo primero a la nueva tabla.
  */
  if (nFilasEncabezado > 0) {
    dataA = dataA.slice(nFilasEncabezado);
    dataB = dataB.slice(nFilasEncabezado);
    for (var i = 0; i < nFilasEncabezado; i++) {
      newSheetData.push(hojaA.getDataRange().getValues()[i]);
    }
  }
  /**
  *  Recorremos todas las filas en nuestra tablaA, y para cada una de ellas las
  *  comparamos con todas las filas de la tablaB, si existe alguna coincidencia,
  *  significa que está duplicada en la segunda tabla y no nos la traemos a la
  *  nueva tabla.
  */
  for (var i = 0; i < dataA.length; i++) {
    var encontrado = false;
    var refATemp = dataA[i].filter(function(elem, index) {
      return columnasDeA.indexOf(index + 1) >= 0;
    });

    for (var j = 0; j < dataB.length && !encontrado; j++) {
      if (refATemp.toString() === dataB[j].toString()) {
        encontrado = true;
      }
    }
    if (!encontrado) {
      newSheetData.push(dataA[i])
    }
  }
  //Si todo ha ido bien, Insertamos los datos generados en una nueva hoja.
  if (newSheetData) {
    hojaA.getParent().insertSheet(hojaA.getName() + "_SIN_DUPLICADOS")
      .getRange(1, 1, newSheetData.length, newSheetData[0].length).setValues(newSheetData);
  }
}
