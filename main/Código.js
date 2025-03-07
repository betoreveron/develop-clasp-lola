
/**
 * Maneja las solicitudes GET y devuelve una plantilla HTML evaluada.
 *
 * @return {HtmlOutput} La salida HTML evaluada de la plantilla "home".
 */
function doGet() {
  var template = HtmlService.createTemplateFromFile("home");
  template.clientes = getOptions(); //Obtengo las opciones para el dropdown

  return template.evaluate();

}

function getOptions() {

  return [
    { email: "cliente1@example.com" },
    { email: "cliente2@example.com" },
    { email: "cliente3@example.com" }
  ];

}

//esta funcion es para traer los parametros de home.html (css y js)
function include(filename){
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function onOpen(){
  LolaFunctions.setMenuLola();
  SpreadsheetApp.flush();
  recordatorioEstatusPago();
}
/* ****** ********************** ****** */
/* ****** FUNCIONES PARA EL MENU ***** */
/* ****** ********************** **** */
function setActiveSheet(sheetName) {
  SpreadsheetApp.getActive().getSheetByName(sheetName).activate();
}

function goToAccesorios(){
  setActiveSheet('Accesorios')
}
function goToEntradas(){
  setActiveSheet('Entradas')
}
function goToSalidas(){
  setActiveSheet('Salidas')
}
function goToInventario(){
  setActiveSheet('Inventario')
}
function goToVistaGlobal(){
  setActiveSheet('Vista Global')
}
function goToReporteGlobal(){
  setActiveSheet('Reporte Global')
}
function goToImportRange(){
  setActiveSheet('Import Range')
}
function goToBD(){
  setActiveSheet('BD')
}
function goToActivosFijos(){
  setActiveSheet('Activos Fijos')
}
function goToGastosTotales(){
  setActiveSheet('Gastos Totales')
}
function goToCalculadora(){
  setActiveSheet('Calculadora')
}
function goToFlujodeCaja(){
  setActiveSheet('Flujo de Caja')
}
/* ****** *************** ****** */
/* ****** *************** **** */


//Funcion para ocultar una hoja
function hideSheet() {
  const file = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = file.getActiveSheet();
  sheet.hideSheet();
}

// Funcion para obetener la cantidad de pagos pendientes por cobrar
function getCantEstatusPago() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Salidas');
  var estatus = sheet.getRange('J:J').getValues().flat();
  var contador = 0;

  for (let i = 0; i < estatus.length; i++) {

    if (estatus[i] == "Pendiente por Cobrar") {

      contador = contador + 1;

    }

  }

  return contador
  
}

// funcion recordatorio al abrir el archivo sheet
function recordatorioEstatusPago() {
  const file = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = file.getSheetByName('Salidas');
  const ui = SpreadsheetApp.getUi();
  var cantidad = getCantEstatusPago(); //obtengo la cantidad de pagos pendientes por cobrar

  let result = ui.alert(
    'Que, si confiesas con tu boca que Jesús es el Señor y crees en tu corazón que Dios lo levantó de entre los muertos, serás salvo. Porque con el corazón se cree para ser justificado, pero con la boca se confiesa para ser salvo. \n' + 
    '(Romanos 10: 9 - 10)',
    ui.ButtonSet.OK
  );
  
  if (result == ui.Button.OK) {

    if (cantidad > 0) {

      ui.alert(
        '💡Recordatorio',
        'Tienes (' + cantidad + ') salidas de accesorios pendientes por cobrar.' +
        '\n\n Recuerda actualizar el estatus de pago en la tabla y el reporte de ventas.',
        ui.ButtonSet.OK
      );

      SpreadsheetApp.flush();
      sheet.activate()

    }

  }

}

// Funciones para borrar contenido de formularios //
function clearFormAccesorios() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Formularios');
  sheet.getRange('D3:D19').clearContent();
}
function clearFormEntradas() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Formularios');
  sheet.getRange('I3:I27').clearContent();
}
function clearFormSalidas() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Formularios');
  sheet.getRange('N3:N23').clearContent();
}
// ------------------------------------------- //

//Funcion para crear un accesorio
function crearAccesorio() {

  const archivo = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = archivo.getSheetByName('Accesorios');
  const form = archivo.getSheetByName('Formularios');
  const vistaGlobal = archivo.getSheetByName('Vista Global');
  const ui = SpreadsheetApp.getUi();
  var campos = form.getRange('D5:D19').getValues();

  //Campos requeridos para registrar un accesorio
  if ((campos[0][0] && campos[2][0] && campos[4][0]) != "") {

    let confirm = ui.alert(
      '⏳Confirmación',
      '¿Desea registrar el accesorio?',
      ui.ButtonSet.YES_NO
    );

    if (confirm == ui.Button.YES) {

      vistaGlobal.appendRow([""]); //Inserto una nueva fila vacia para no afectar el Query

      hoja.getRange(hoja.getLastRow() + 1, 2).setValue(campos[0][0]);  //inserto marca en una nueva fila
      hoja.getRange(hoja.getLastRow(), 3).setValue(campos[2][0]);   //inserto categoria accesorio

      hoja.getRange(hoja.getLastRow(), 4).setValue(campos[4][0]);  //inserto descripcion accesorio
      var celda = hoja.getRange(hoja.getLastRow(), 4).getA1Notation(); // Almaceno la celda
      insertarURL(celda, campos[6][0], campos[4][0]) // Inserto el Link del accesorio

      hoja.getRange(hoja.getLastRow(), 5).setValue(campos[8][0]);  //inserto tipo accesorio
      hoja.getRange(hoja.getLastRow(), 6).setValue(campos[10][0]);  //inserto tamano accesorio

      //Fijo los costos, precio estimado y precio en 0
      for (var i = 7; i <= 12; i++) {
        hoja.getRange(hoja.getLastRow(), i).setValue(0);
      }

      form.getRange('D5:D19').clearContent();

      let result = ui.alert(
        '🎉¡Accesorio creado!',
        'El código del accesorio es: ' + hoja.getRange(hoja.getLastRow(), 1).getValue(),
        ui.ButtonSet.OK
      );

      if (result == ui.Button.OK) {
        ui.alert(
          '💡Recordatorio',
          'Debes registrar la entrada del accesorio al Inventario de Lola',
          ui.ButtonSet.OK
        );
      }

    } else {
      form.getRange('D5:D19').clearContent();
    }

    //Mensaje si los campos categoria y descripcion estan vacios
  } else {
    ui.alert(
      '⛔¡Error!',
      'Los campos obligatorios para registrar un accesorio son: \n\n 1. Marca \n 2.Categoría \n 3. Descripción del accesorio',
      ui.ButtonSet.OK
    );
    form.getRange('D5:D19').clearContent();
  }

}

// Funcion para insertar URL a un accesorio
function insertarURL(cell, url, text) {

  var file = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = file.getSheetByName('Accesorios');
  var formula = '=HYPERLINK("' + url + '"; "' + text + '")';

  sheet.getRange(cell).setFormula(formula);

}

// Funcion para extraer el URL y mostrarlo en la busqueda de accesorio
function extraerURL(row) {

  let file = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = file.getSheetByName('Accesorios');
  let cell = sheet.getRange(row, 4).getA1Notation();
  let range = sheet.getRange(cell);
  const url = range.getRichTextValue().getLinkUrl();
  return url;

}

//Funcion para validar que los campos codigo, precio estimado y precio de venta esten vacios
function confirmFormAccesorio() {
  const archivo = SpreadsheetApp.getActiveSpreadsheet();
  const form = archivo.getSheetByName('Formularios');
  const ui = SpreadsheetApp.getUi();
  var codigo = form.getRange('D3').getValue();
  var precioEstimado = form.getRange('D17').getValue();
  var precioVenta = form.getRange('D19').getValue();

  if ((codigo || precioEstimado || precioVenta) != "") {

    let result = ui.alert(
      '⛔¡Error!',
      'Para registrar un accesorio solo debes completar los campos:\n\n' +
      '1. Marca \n' +
      '2. Categoría \n' +
      '3. Descripción \n' +
      '4. Tipo (si aplica) \n' +
      '5. Tamaño',
      ui.ButtonSet.OK
    );

    if (result == ui.Button.OK) {

      let confirm = ui.alert(
        '💡Aviso',
        'Para registrar el accesorio, será eliminado \n cualquier valor en los campos:\n\n' +
        '1. Código \n' +
        '2. Precio Estimado de Venta \n' +
        '3. Precio de Venta \n',
        ui.ButtonSet.OK
      );

      if (confirm == ui.Button.OK) {
        form.getRange('D3').clearContent();
        form.getRange('D17').clearContent();
        form.getRange('D19').clearContent();
      }

    }
  }
  crearAccesorio()

}

//Funcion para buscar los valores del accesorio
function buscarAccesorio() {

  const archivo = SpreadsheetApp.getActiveSpreadsheet();
  const hojaAccesorios = archivo.getSheetByName('Accesorios');
  const hojaForm = archivo.getSheetByName('Formularios');
  var codigo = hojaForm.getRange('D3').getValue(); //Obtengo el codigo a buscar
  var campos = hojaForm.getRange('D5:D19').getValues();

  var accesorios = hojaAccesorios.getDataRange().getValues(); //Tabla accesorios en arreglo
  var listaCodigos = accesorios.map(fila => fila[0]); //Todos los codigos de accesorios en un arreglo
  var indice = listaCodigos.indexOf(codigo);
  var accesorioURL = extraerURL(indice + 1) //Obtengo la URL de la imagen del accesorio
  const ui = SpreadsheetApp.getUi();

  if ((campos[0][0] || campos[2][0] || campos[4][0] || campos[6][0] || campos[8][0] || campos[10][0]
    || campos[12][0] || campos[14][0]) == "") {

    //Busco los valores de cada campo y los muestro en el formulario
    if (indice != -1 && codigo >= 1) {

      ui.alert(
        '🎉¡Búsqueda exitosa!',
        'El accesorio ha sido encontrado',
        ui.ButtonSet.OK
      );
      hojaForm.getRange('D5').setValue(accesorios[indice][1]);  //marca
      hojaForm.getRange('D7').setValue(accesorios[indice][2]);  //categoria
      hojaForm.getRange('D9').setValue(accesorios[indice][3]);  //descripcion
      hojaForm.getRange('D11').setValue(accesorioURL);  //URL
      hojaForm.getRange('D13').setValue(accesorios[indice][4]);  //tipo
      hojaForm.getRange('D15').setValue(accesorios[indice][5]);  //tamano
      hojaForm.getRange('D17').setValue(accesorios[indice][10]);  //precio estimado de venta
      hojaForm.getRange('D19').setValue(accesorios[indice][11]);  //precio de venta
      //SpreadsheetApp.getActive().toast("Búsqueda exitosa", "✅Éxito");

    } else if (codigo != "") {

      ui.alert(
        '⛔¡Error!',
        'No existe un accesorio con el código: ' + codigo,
        ui.ButtonSet.OK
      );
      hojaForm.getRange('D3').clearContent();

    } else if (codigo == "") {

      ui.alert(
        '⛔¡Error!',
        'Debes indicar el código del accesorio que deseas buscar',
        ui.ButtonSet.OK
      );
    }

  } else {
    ui.alert(
      '⛔¡Error!',
      'Solo debe indicar el código a buscar',
      ui.ButtonSet.OK
    );
    hojaForm.getRange('D5:D19').clearContent();
  }

}

//Funcion para editar un accesorio
function editarAccesorio() {

  const archivo = SpreadsheetApp.getActiveSpreadsheet();
  const hojaAccesorios = archivo.getSheetByName('Accesorios');
  const hojaForm = archivo.getSheetByName('Formularios');
  const hojaEntradas = archivo.getSheetByName('Entradas');
  var codigo = hojaForm.getRange('D3').getValue();
  var campos = hojaForm.getRange('D5:D19').getValues();

  var entradas = hojaEntradas.getDataRange().getValues(); //Tabla entradas en arreglo
  var codigosEntrada = entradas.map(entry => entry[0]);  //Todos los codigos de accesorios en un arreglo
  var registroEntrada = codigosEntrada.indexOf(codigo) + 1; //Fila en la tabla (sin tomar el encabezado)

  var accesorios = hojaAccesorios.getDataRange().getValues(); //Tabla accesorios en arreglo
  var listaCodigos = accesorios.map(fila => fila[0]); //Todos los codigos de accesorios en un arreglo
  var indice = listaCodigos.indexOf(codigo) + 1; //Fila en la tabla (sin tomar el encabezado)
  const ui = SpreadsheetApp.getUi();

  //Campos requeridos para registrar un accesorio
  if (indice != 0 && codigo >= 1) {

    var precioEstimadoReal = hojaAccesorios.getRange(indice, 11).getValue();  //Si el codigo existe, precioEstimadoReal
    var precioVentaReal = hojaAccesorios.getRange(indice, 12).getValue();  //Si el codigo existe, precioVentaReal

    //El usuario no puede cambiar el precio estimado de venta
    if (precioEstimadoReal != campos[12][0]) {

      hojaForm.getRange('D17').setValue(precioEstimadoReal); //Reasigno el precio real
      SpreadsheetApp.flush();
      ui.alert(
        '⛔Acción Denegada',
        'No puedes cambiar el precio estimado de venta',
        ui.ButtonSet.OK
      );

      //el precio de venta no puede ser menor al precio estimado
    } else if (campos[14][0] < campos[12][0]) {

      ui.alert(
        '💡Recuerda',
        'El precio de venta debe ser mayor o igual \n al precio estimado de venta',
        ui.ButtonSet.OK
      );

    } else if (campos[14][0] != 0 && registroEntrada == 0) {

      hojaForm.getRange('D19').setValue(precioVentaReal); //Reasigno el precio real
      SpreadsheetApp.flush();

      ui.alert(
        '⛔Acción Denegada',
        'No puedes editar el precio de venta del accesorio porque \n no existe registro(s) de entrada(s) al Inventario de Lola',
        ui.ButtonSet.OK
      );

    } else {

      hojaAccesorios.getRange(indice, 2).setValue(campos[0][0]); //actualizo marca accesorio
      hojaAccesorios.getRange(indice, 3).setValue(campos[2][0]); //actualizo categoria accesorio

      hojaAccesorios.getRange(indice, 4).setValue(campos[4][0]); //actualizo descripcion accesorio

      var celda = hojaAccesorios.getRange(indice, 4).getA1Notation(); // Obtengo la celda
      insertarURL(celda, campos[6][0], campos[4][0]) // Inserto el Link del accesorio al editar

      hojaAccesorios.getRange(indice, 5).setValue(campos[8][0]); //actualizo tipo accesorio
      hojaAccesorios.getRange(indice, 6).setValue(campos[10][0]); //actualizo tamano accesorio
      hojaAccesorios.getRange(indice, 12).setValue(campos[14][0]);  //actualizo precio de venta

      syncEdicionAccesorio(codigo, campos) //Envio el codigo y campos para actualizar

      hojaForm.getRange('D3:D19').clearContent();
      //SpreadsheetApp.getActive().toast("Búsqueda exitosa", "✅Éxito");

      SpreadsheetApp.flush();
      ui.alert(
        '🎉¡Éxito!',
        'El accesorio ha sido actualizado',
        ui.ButtonSet.OK
      );

    }

    //Si el codigo no existe
  } else if (indice == 0) {

    ui.alert(
      '⛔¡Error!',
      'El código del accesorio que desea editar no existe',
      ui.ButtonSet.OK
    );

    //Si no hay un codigo
  } else if (codigo == "") {

    ui.alert(
      '⛔¡Error!',
      'Debe indicar el código del accesorio que deseas editar',
      ui.ButtonSet.OK
    );

  }

}

// Funcion para actualizar valores de accesorios en Entradas y Salidas
function syncEdicionAccesorio(code, data) {

  const file = SpreadsheetApp.getActiveSpreadsheet();
  const hojaEntradas = file.getSheetByName('Entradas');
  const hojaSalidas = file.getSheetByName('Salidas');

  var salidas = hojaSalidas.getRange('A2:A');
  var codSalidas = salidas.getValues();

  var entradas = hojaEntradas.getRange('A2:A');
  var codEntradas = entradas.getValues();

  data.flat();
  var marca = data[0][0];
  var categoria = data[2][0];
  var descripcion = data[4][0];
  var tipo = data[8][0];
  var tamano = data[10][0];

  // Recorro las Entradas para obtener las filas donde se encuentra el codigo
  for (let i = 0; i < codEntradas.length; i++) {

    if (codEntradas[i][0] == code) {

      hojaEntradas.getRange(i + 2, 2).setValue(marca);
      hojaEntradas.getRange(i + 2, 3).setValue(categoria);
      hojaEntradas.getRange(i + 2, 4).setValue(descripcion);
      hojaEntradas.getRange(i + 2, 5).setValue(tipo);
      hojaEntradas.getRange(i + 2, 6).setValue(tamano);

    }

  }

  // Recorro las Salidas para obtener las filas donde se encuentra el codigo
  for (let j = 0; j < codSalidas.length; j++) {

    if (codSalidas[j][0] == code) {

      hojaEntradas.getRange(j + 2, 2).setValue(marca);
      hojaEntradas.getRange(j + 2, 3).setValue(categoria);
      hojaEntradas.getRange(j + 2, 4).setValue(descripcion);
      hojaEntradas.getRange(j + 2, 5).setValue(tipo);
      hojaEntradas.getRange(j + 2, 6).setValue(tamano);

    }

  }


}

//Funcion para eliminar datos de un accesorio
function eliminarAccesorio() {

  const ui = SpreadsheetApp.getUi();
  const archivo = SpreadsheetApp.getActiveSpreadsheet();
  const hojaAccesorios = archivo.getSheetByName('Accesorios');
  const hojaForm = archivo.getSheetByName('Formularios');
  var codigo = hojaForm.getRange('D3').getValue(); //Obtengo el codigo a eliminar
  var campos = hojaForm.getRange('D5:D19').getValues();

  var accesorios = hojaAccesorios.getDataRange().getValues(); //Tabla accesorios en arreglo
  var listaCodigos = accesorios.map(fila => fila[0]); //Todos los codigos de accesorios en un arreglo
  var indice = listaCodigos.indexOf(codigo);

  //---------------------- SALIDAS --------------------------// 
  var hojaSalidas = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Salidas');
  var salidas = hojaSalidas.getRange('A2:A');
  var codSalidas = salidas.getValues();
  var filasSalidas = [];

  //---------------------- ENTRADAS --------------------------// 
  var hojaEntradas = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Entradas');
  var entradas = hojaEntradas.getRange('A2:A');
  var codEntradas = entradas.getValues();
  var filasEntradas = [];

  if ((campos[0][0] || campos[2][0] || campos[4][0] || campos[6][0] || campos[8][0] || campos[10][0] || campos[12][0] || campos[14][0]) == "") {

    //Busco los valores de cada campo y los muestro en el formulario
    if (indice != -1 && codigo >= 1) {

      let result = ui.alert(
        '⏳Confirmación',
        '¿Está seguro de querer eliminar este accesorio?',
        ui.ButtonSet.YES_NO
      );

      if (result == ui.Button.YES) {

        //recorro el rango para obtener las filas donde se encuentra el codigo
        for (var i = 0; i < codSalidas.length; i++) {

          if (codSalidas[i][0] == codigo) {

            filasSalidas.push(i + 2); // Ajusta por el encabezado - Indico las filas donde se encuentra el codigo

          }

        }
        //Si esta el codigo en Entradas no puede eliminarlo
        if (filasSalidas.length >= 1) {

          ui.alert(
            '⚠️Acción Denegada',
            'No puedes eliminar un accesorio que tiene salidas del Inventario de Lola \n\n' +
            'Si necesitas realizar esta acción ponte en contacto con tu Papi Rico y Delicioso',
            ui.ButtonSet.OK
          );
          hojaForm.getRange('D3').clearContent();

        } else {

          //recorro el rango para obtener las filas donde se encuentra el codigo
          for (var i = 0; i < codEntradas.length; i++) {

            if (codEntradas[i][0] == codigo) {

              filasEntradas.push(i + 2); // Ajusta por el encabezado - Indico las filas donde se encuentra el codigo

            }

          }

          //Si el arreglo no tiene nada, es que no existe registro en Entradas
          if (filasEntradas.length == 0) {

            ui.alert(
              '💡Importante',
              'No existe registro de entradas del accesorio \n al Inventario de Lola',
              ui.ButtonSet.OK
            );
            confirmDelete(hojaAccesorios, hojaForm, indice, ui);

            //Si solo existe un registro en Entradas
          } else if (filasEntradas.length == 1) {

            //recorro el rango y elimino las filas donde esta el codigo
            for (var j = filasEntradas.length - 1; j >= 0; j--) {

              hojaEntradas.deleteRow(filasEntradas[j]);

            }

            ui.alert(
              '✅¡Éxito!',
              'Se ha eliminado' + ' ' + '(' + filasEntradas.length + ')' + ' ' + 'registro en la tabla Entradas',
              ui.ButtonSet.OK
            );
            confirmDelete(hojaAccesorios, hojaForm, indice, ui);

            //Si hay mas de un registro en Entradas
          } else if (filasEntradas.length > 1) {

            //recorro el rango y elimino las filas donde esta el codigo
            for (var j = filasEntradas.length - 1; j >= 0; j--) {

              hojaEntradas.deleteRow(filasEntradas[j]);

            }

            ui.alert(
              '✅¡Éxito!',
              'Se han eliminado' + ' ' + '(' + filasEntradas.length + ')' + ' ' + 'registros en la tabla Entradas',
              ui.ButtonSet.OK
            );
            confirmDelete(hojaAccesorios, hojaForm, indice, ui);

          }
        }

      } else {
        // User clicked "No" or X in the title bar.
        hojaForm.getRange('D3:D19').clearContent();
      }

      //Validaciones formulario
    } else if (codigo != "") {

      ui.alert(
        '⛔¡Error!',
        'No existe un accesorio con el código: ' + codigo,
        ui.ButtonSet.OK
      );
      hojaForm.getRange('D3').clearContent();

    } else if (codigo == "") {

      ui.alert(
        '⛔¡Error!',
        'Debes indicar el código del accesorio que deseas eliminar.',
        ui.ButtonSet.OK
      );
    }

  } else {
    /*ui.alert(
      '💡Recuerda',
      'Para eliminar un accesorio solo debes indicar su código',
      ui.ButtonSet.OK
    );*/
    hojaForm.getRange('D5:D19').clearContent();
    SpreadsheetApp.flush();
    eliminarAccesorio();
  }

}

//Funcion para eliminar los valores correspondientes en la tabla Accesorios
function confirmDelete(hojaAccesorios, hojaForm, indice, ui) {

  ui.alert(
    '✅¡Éxito!',
    'El accesorio ha sido eliminado',
    ui.ButtonSet.OK
  );

  /*hojaAccesorios.getRange(indice+1,2).clearContent();
  hojaAccesorios.getRange(indice+1,3).clearContent();
  hojaAccesorios.getRange(indice+1,4).clearContent();
  hojaAccesorios.getRange(indice+1,5).clearContent();
  hojaAccesorios.getRange(indice+1,10).clearContent();
  hojaForm.getRange('D3').clearContent();*/

  //Codigo para borrar contenido de una fila
  for (var i = 2; i <= 12; i++) {
    hojaAccesorios.getRange(indice + 1, i).clearContent();
  }

  borrarUltimaFilaVistaGlobal()

  hojaForm.getRange('D3').clearContent();

}

//Funcion para eliminar ultima fila de Vista Global
function borrarUltimaFilaVistaGlobal() {

  const file = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = file.getSheetByName('Vista Global');
  const codes = sheet.getRange('A2:A').getValues();
  var lastCode = codes[codes.length - 1]; //Ultimo codigo del arreglo
  var previousCode = sheet.getLastRow() - 1;


  // Si el ultimo codigo es vacio
  if (lastCode == "") {

    sheet.deleteRow(previousCode + 1); //Elimino la ultima fila

  }

}

//Funcion para buscar el accesorio en el formulario Salidas
function buscarEntradas() {
  const code = 'I3'; //campo codigo en formulario 
  const array = 'I5:I27'; //los otros campos en formulario
  const form = ['I5', 'I7', 'I9', 'I11', 'I13']; //casillas donde quiero mostrar la busqueda
  buscarInventario(code, array, form)
}

//Funcion para buscar el accesorio en el formulario Salidas
function buscarSalidas() {
  const code = 'N3'; //campo codigo en formulario
  const array = 'N5:N21'; //los otros campos en formulario
  const form = ['N5', 'N7', 'N9', 'N11', 'N13']; //casillas donde quiero mostrar la busqueda
  buscarInventario(code, array, form)
}

// Funcion Global que busca el accesorio y lo muestra en los formularios de Entradas y Salidas
function buscarInventario(code, array, form) {

  const archivo = SpreadsheetApp.getActiveSpreadsheet();
  const hojaAccesorios = archivo.getSheetByName('Accesorios');
  const hojaForm = archivo.getSheetByName('Formularios');
  var codigo = hojaForm.getRange(code).getValue(); //Obtengo el codigo a buscar
  var campos = hojaForm.getRange(array).getValues();
  var accesorios = hojaAccesorios.getDataRange().getValues(); //Tabla accesorios en arreglo
  var listaCodigos = accesorios.map(fila => fila[0]); //Todos los codigos de accesorios en un arreglo
  var indice = listaCodigos.indexOf(codigo);
  const ui = SpreadsheetApp.getUi();

  if ((campos[0][0] || campos[2][0] || campos[4][0] || campos[6][0] || campos[8][0]) == "") {

    //Busco los valores de cada campo y los muestro en el formulario
    if (indice != -1 && codigo >= 1) {

      hojaForm.getRange(form[0]).setValue(accesorios[indice][1]);  //marca
      hojaForm.getRange(form[1]).setValue(accesorios[indice][2]);  //categoria
      hojaForm.getRange(form[2]).setValue(accesorios[indice][3]);  //descripcion
      hojaForm.getRange(form[3]).setValue(accesorios[indice][4]);  //tipo
      hojaForm.getRange(form[4]).setValue(accesorios[indice][5]);  //tamano
      //SpreadsheetApp.getActive().toast("Búsqueda exitosa", "✅Éxito");

      SpreadsheetApp.flush();

      ui.alert(
        '🎉¡Búsqueda exitosa!',
        'El accesorio ha sido encontrado',
        ui.ButtonSet.OK
      );

    } else if (codigo != "") {

      hojaForm.getRange(code).clearContent();

      SpreadsheetApp.flush();

      ui.alert(
        '⛔¡Error!',
        'No existe un accesorio con el código: ' + codigo,
        ui.ButtonSet.OK
      );

    } else if (codigo == "") {

      ui.alert(
        '⛔¡Error!',
        'Debes indicar el código del accesorio que deseas buscar.',
        ui.ButtonSet.OK
      );
    }

  } else {

    hojaForm.getRange(form[0]).clearContent();
    hojaForm.getRange(form[1]).clearContent();
    hojaForm.getRange(form[2]).clearContent();
    hojaForm.getRange(form[3]).clearContent();
    hojaForm.getRange(form[4]).clearContent();

    SpreadsheetApp.flush();

    ui.alert(
      '⛔¡Error!',
      'Solo debe indicar el código a buscar',
      ui.ButtonSet.OK
    );

  }

}


/* ********* FUNCIONES PARA CONTROL DE ENTRADAS Y COSTOS MAXIMOS ********* */
/* ********************************************************************** */

//Funcion para calcular los costos maximos historicos de un accesorio
function calcularCostosMaximos(code) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Entradas');

  var columnaCodigo = 0;
  var columnaProduccion = 7;
  var columnaManoObra = 8;
  var columnaEmpaquetado = 9;
  var columnaEnvio = 10;

  var maxProduccion = 0;
  var maxManoObra = 0;
  var maxEmpaquetado = 0;
  var maxEnvio = 0;

  var filasCodigos = [];

  var datos = sheet.getDataRange().getValues();

  //Recorrer cada fila (omitimos la primera fila de encabezados)
  for (var i = 1; i < datos.length; i++) {

    var codigo = datos[i][columnaCodigo];

    // Si encuentra el codigo en Entradas
    if (code == codigo) {
      var produccion = datos[i][columnaProduccion];
      var manoObra = datos[i][columnaManoObra];
      var empaquetado = datos[i][columnaEmpaquetado];
      var envio = datos[i][columnaEnvio];

      filasCodigos.push(i + 1); //para almacenar las filas donde se encuentra el codigo

      // Actualizar los máximos
      maxProduccion = Math.max(maxProduccion, produccion);
      maxManoObra = Math.max(maxManoObra, manoObra);
      maxEmpaquetado = Math.max(maxEmpaquetado, empaquetado);
      maxEnvio = Math.max(maxEnvio, envio);
    }

  }

  /*Logger.log("Costo máximo de producción: " + maxProduccion);
  Logger.log("Costo máximo de mano de obra: " + maxManoObra);
  Logger.log("Costo máximo de empaquetado: " + maxEmpaquetado);
  Logger.log("Costo máximo de envío: " + maxEnvio);
  Logger.log(cantCodigos);*/

  //actualizo los costos en la tabla accesorios
  actualizarCostosMaximos(code, maxProduccion, maxManoObra, maxEmpaquetado, maxEnvio)

}

//Funcion para actualizar los costos maximos de un accesorio
function actualizarCostosMaximos(code, maxProduccion, maxManoObra, maxEmpaquetado, maxEnvio) {

  const hojaAccesorios = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Accesorios');
  var accesorios = hojaAccesorios.getDataRange().getValues(); //Tabla accesorios en arreglo
  var listaCodigos = accesorios.map(fila => fila[0]); //Todos los codigos de accesorios en un arreglo
  var indice = listaCodigos.indexOf(code); //Fila donde se encuentra el codigo
  const ui = SpreadsheetApp.getUi();
  var precioVenta = hojaAccesorios.getRange(indice + 1, 12).getValue();
  const descripcion = hojaAccesorios.getRange(indice + 1, 4).getValue(); //nombre del accesorio
  var precioEstimado = ((maxProduccion + maxManoObra + maxEmpaquetado + maxEnvio) * 1.3).toFixed(2).replace(".", ",");

  hojaAccesorios.getRange(indice + 1, 7).setValue(maxProduccion);
  hojaAccesorios.getRange(indice + 1, 8).setValue(maxManoObra);
  hojaAccesorios.getRange(indice + 1, 9).setValue(maxEmpaquetado);
  hojaAccesorios.getRange(indice + 1, 10).setValue(maxEnvio);
  hojaAccesorios.getRange(indice + 1, 11).setValue(precioEstimado);

  let result = ui.alert(
    '💡Información',
    'El accesorio "' + descripcion + '" tiene un \n precio estimado de venta de: $' + precioEstimado,
    ui.ButtonSet.OK
  );

  if (result == ui.Button.OK && precioVenta == "") {

    ui.alert(
      '🔔Importante',
      'Debes definir el precio de venta del accesorio',
      ui.ButtonSet.OK
    );

  }

}

//Funcion para registrar la entrada de un accesorio al inventario
function confirmFormEntrada() {
  const archivo = SpreadsheetApp.getActiveSpreadsheet();
  const hojaAccesorios = archivo.getSheetByName('Accesorios');
  const hoja = archivo.getSheetByName('Entradas');
  const form = archivo.getSheetByName('Formularios');
  const ui = SpreadsheetApp.getUi();
  var campos = form.getRange('I3:I27').getValues();

  var accesorios = archivo.getSheetByName('Accesorios').getDataRange().getValues(); //Tabla accesorios en arreglo
  var listaCodigos = accesorios.map(fila => fila[0]); //Todos los codigos de accesorios en un arreglo
  var indice = listaCodigos.indexOf(campos[0][0]) + 1; //Fila en la tabla (sin tomar el encabezado)*/

  if (indice - 1 != -1 && campos[0][0] >= 1) {

    //Campos requeridos para registrar una entrada
    if ((campos[0][0] && campos[12][0] && campos[14][0] && campos[18][0] && campos[22][0]) != "") {

      //Para validar si el usuario cambia algún valor del accesorio
      var marcaReal = hojaAccesorios.getRange(indice, 2).getValue();
      var categoriaReal = hojaAccesorios.getRange(indice, 3).getValue();
      var descripcionReal = hojaAccesorios.getRange(indice, 4).getValue();
      var tipoReal = hojaAccesorios.getRange(indice, 5).getValue();
      var tamanoReal = hojaAccesorios.getRange(indice, 6).getValue();

      if (marcaReal != campos[2][0] ||
        categoriaReal != campos[4][0] ||
        descripcionReal != campos[6][0] ||
        tipoReal != campos[8][0] ||
        tamanoReal != campos[10][0]) {

        form.getRange('I5').setValue(marcaReal);
        form.getRange('I7').setValue(categoriaReal);
        form.getRange('I9').setValue(descripcionReal);
        form.getRange('I11').setValue(tipoReal);
        form.getRange('I13').setValue(tamanoReal);
        SpreadsheetApp.flush();
        crearEntrada(form, hoja)

      } else {
        crearEntrada(form, hoja)
      }

      //Campos obligatorios
    } else {
      ui.alert(
        '💡Instrucciones',
        'Para registrar la entrada de un accesorio al Inventario de Lola:\n\n' +
        '1. Debes buscar el código del accesorio.\n' +
        '2. No debes editar ningún valor de los campos de dicho accesorio.\n' +
        '3. Los campos obligatorios para registrar una entrada son: \n\n' +
        '   - Código \n - Cantidad \n - Costo Producto \n - Costo de Empaque \n - Fecha',
        ui.ButtonSet.OK
      );
    }

  } else {
    form.getRange('I3:I27').clearContent();
    SpreadsheetApp.flush();
    ui.alert(
      '⛔Error!',
      'No puedes registrar una entrada al inventario de \n un accesorio que no existe.',
      ui.ButtonSet.OK
    );
  }

}

function crearEntrada(form, hoja) {

  var campos = form.getRange('I3:I27').getValues();
  Logger.log(campos);
  const ui = SpreadsheetApp.getUi();

  let confirm = ui.alert(
    '⏳Confirmación',
    '¿Desea registrar una entrada del accesorio \n' + '"' + campos[6][0] + '" al Inventario de Lola?',
    ui.ButtonSet.YES_NO
  );

  if (confirm == ui.Button.YES) {

    hoja.getRange(hoja.getLastRow() + 1, 1).setValue(campos[0][0]);  //inserto codigo del accesorio
    hoja.getRange(hoja.getLastRow(), 2).setValue(campos[2][0]);  //inserto marca del accesorio
    hoja.getRange(hoja.getLastRow(), 3).setValue(campos[4][0]);  //inserto categoria del accesorio
    hoja.getRange(hoja.getLastRow(), 4).setValue(campos[6][0]);  //inserto descripcion del accesorio
    hoja.getRange(hoja.getLastRow(), 5).setValue(campos[8][0]);  //inserto tipo del accesorio
    hoja.getRange(hoja.getLastRow(), 6).setValue(campos[10][0]);  //inserto tamano del accesorio

    hoja.getRange(hoja.getLastRow(), 7).setValue(campos[12][0]);  //inserto cantidad
    hoja.getRange(hoja.getLastRow(), 8).setValue(campos[14][0]);  //inserto costo producto

    if (campos[16][0] == "") {
      hoja.getRange(hoja.getLastRow(), 9).setValue(0);  //inserto costo mano de obra
    } else {
        hoja.getRange(hoja.getLastRow(), 9).setValue(campos[16][0]);  //inserto costo mano de obra
    }

    hoja.getRange(hoja.getLastRow(), 10).setValue(campos[18][0]);  //inserto costo empaque

    if (campos[20][0] == "") {
      hoja.getRange(hoja.getLastRow(), 11).setValue(0);  //inserto costo de envio
    } else {
        hoja.getRange(hoja.getLastRow(), 11).setValue(campos[20][0]);  //inserto costo de envio
    }

    hoja.getRange(hoja.getLastRow(), 12).setValue(campos[22][0]);  //inserto fecha
    hoja.getRange(hoja.getLastRow(), 13).setValue(campos[24][0]);  //inserto la nota

    form.getRange('I3:I27').clearContent();
    SpreadsheetApp.flush();

    ui.alert(
      '🎉¡Entrada exitosa!',
      'El accesorio ha ingresado al inventario de Lola',
      ui.ButtonSet.OK
    );

    calcularCostosMaximos(campos[0][0])

  } else {
    form.getRange('I3:I27').clearContent();
  }

}

// Funcion para establecer los costos en 0 de un accesorio que no tiene entradas
function validarEntradas() {

  var file = SpreadsheetApp.getActiveSpreadsheet();
  var accesorios = file.getSheetByName('Accesorios');
  var entradas = file.getSheetByName('Entradas');

  var tablaAccesorios = accesorios.getDataRange().getValues();
  var tablaEntradas = entradas.getDataRange().getValues();

  // Columnas de costos a setear
  var costoProducto = 7;
  var costoManoObra = 8;
  var costoEmpaque = 9;
  var costoEnvio = 10;
  var precioEstimado = 11;
  var precioVenta = 12;

  for (var i = 1; i < tablaAccesorios.length; i++) {

    var codigo = tablaAccesorios[i][0];
    var hayEntradas = false;

    // Verificamos si hay entradas en el inventario para este accesorio
    for (var j = 1; j < tablaEntradas.length; j++) {

      if (tablaEntradas[j][0] === codigo) {
        hayEntradas = true;
        break;
      }

    }

    // Si no hay entradas, establecemos los costos en cero
    if (!hayEntradas) {

      accesorios.getRange(i + 1, costoProducto).setValue(0);
      accesorios.getRange(i + 1, costoManoObra).setValue(0);
      accesorios.getRange(i + 1, costoEmpaque).setValue(0);
      accesorios.getRange(i + 1, costoEnvio).setValue(0);
      accesorios.getRange(i + 1, precioEstimado).setValue(0);
      accesorios.getRange(i + 1, precioVenta).setValue(0);

    }

  }

}

// Funcion para validar el formulario de Salidas
function confirmFormSalida() {

  const archivo = SpreadsheetApp.getActiveSpreadsheet();
  const hojaAccesorios = archivo.getSheetByName('Accesorios');
  const hoja = archivo.getSheetByName('Salidas');
  const form = archivo.getSheetByName('Formularios');
  const ui = SpreadsheetApp.getUi();
  var campos = form.getRange('N3:N23').getValues();

  var accesorios = archivo.getSheetByName('Accesorios').getDataRange().getValues(); //Tabla accesorios en arreglo
  var listaCodigos = accesorios.map(fila => fila[0]); //Todos los codigos de accesorios en un arreglo
  var indice = listaCodigos.indexOf(campos[0][0]) + 1; //Fila en la tabla (sin tomar el encabezado)*/

  if (indice - 1 != -1 && campos[0][0] >= 1) {

    //Campos requeridos para registrar una entrada
    if ((campos[0][0] && campos[12][0] && campos[14][0] && campos[16][0] && campos[18][0]) != "") {

      //Para validar si el usuario cambia algún valor del accesorio
      var marcaReal = hojaAccesorios.getRange(indice, 2).getValue();
      var categoriaReal = hojaAccesorios.getRange(indice, 3).getValue();
      var descripcionReal = hojaAccesorios.getRange(indice, 4).getValue();
      var tipoReal = hojaAccesorios.getRange(indice, 5).getValue();
      var tamanoReal = hojaAccesorios.getRange(indice, 6).getValue();
      var precioVenta = hojaAccesorios.getRange(indice, 12).getValue();
      var stock = hojaAccesorios.getRange(indice, 13).getValue();

      if (marcaReal != campos[2][0] ||
        categoriaReal != campos[4][0] ||
        descripcionReal != campos[6][0] ||
        tipoReal != campos[8][0] ||
        tamanoReal != campos[10][0]) {

        form.getRange('N5').setValue(marcaReal);
        form.getRange('N7').setValue(categoriaReal);
        form.getRange('N9').setValue(descripcionReal);
        form.getRange('N11').setValue(tipoReal);
        form.getRange('N13').setValue(tamanoReal);
        SpreadsheetApp.flush();
        crearSalida(form, hoja, precioVenta)

      } else if (stock == 0) {

        ui.alert(
          '⛔Prohibido',
          'No puedes vender el accesorio porque no hay stock',
          ui.ButtonSet.OK
        );

        SpreadsheetApp.flush();

        form.getRange('N3:N23').clearContent();

      } else if (campos[12][0] > stock){

        ui.alert(
          '⛔Prohibido',
          'No puedes sobrevender el accesorio. \n\n La cantidad disponible es:  ' + stock,
          ui.ButtonSet.OK
        );

      } else {
        crearSalida(form, hoja, precioVenta)
      }

      //Campos obligatorios
    } else {
      ui.alert(
        '💡Instrucciones',
        'Para registrar la salida de un accesorio del Inventario de Lola:\n\n' +
        '1. Debes buscar el código del accesorio.\n' +
        '2. No debes editar ningún valor de los campos de dicho accesorio.\n' +
        '3. Los campos obligatorios para registrar una entrada son: \n\n' +
        '   - Código \n - Cantidad \n - Fecha \n - Estatus de Pago \n - Canal de Venta',
        ui.ButtonSet.OK
      );
    }

  } else {
    form.getRange('N3:N23').clearContent();
    SpreadsheetApp.flush();
    ui.alert(
      '⛔¡Error!',
      'No puedes registrar la salida del inventario de \n un accesorio que no existe.',
      ui.ButtonSet.OK
    );
  }

}

// Funcion para registrar la salida de un accesorio al inventario 
function crearSalida(form, hoja, precio) {

  var campos = form.getRange('N3:N23').getValues();
  const ui = SpreadsheetApp.getUi();

  let confirm = ui.alert(
    '⏳Confirmación',
    '¿Desea registrar una salida del accesorio \n' + '"' + campos[6][0] + '" del Inventario de Lola?',
    ui.ButtonSet.YES_NO
  );

  if (confirm == ui.Button.YES) {

    hoja.getRange(hoja.getLastRow() + 1, 1).setValue(campos[0][0]);  //inserto codigo del accesorio
    hoja.getRange(hoja.getLastRow(), 2).setValue(campos[2][0]);  //inserto marca del accesorio
    hoja.getRange(hoja.getLastRow(), 3).setValue(campos[4][0]);  //inserto categoria del accesorio
    hoja.getRange(hoja.getLastRow(), 4).setValue(campos[6][0]);  //inserto descripcion del accesorio
    hoja.getRange(hoja.getLastRow(), 5).setValue(campos[8][0]);  //inserto tipo del accesorio
    hoja.getRange(hoja.getLastRow(), 6).setValue(campos[10][0]);  //inserto tamano del accesorio
    hoja.getRange(hoja.getLastRow(), 7).setValue(campos[12][0]);  //inserto cantidad
    hoja.getRange(hoja.getLastRow(), 8).setValue(precio);     //inserto precio de venta
    hoja.getRange(hoja.getLastRow(), 9).setValue(campos[14][0]);  //inserto fecha
    hoja.getRange(hoja.getLastRow(), 10).setValue(campos[16][0]);  //inserto estatus pago
    hoja.getRange(hoja.getLastRow(), 11).setValue(campos[18][0]);  //inserto canal de venta
    hoja.getRange(hoja.getLastRow(), 12).setValue(campos[20][0]);  //inserto nota

    form.getRange('N3:N23').clearContent();
    SpreadsheetApp.flush();

    ui.alert(
      '🎉¡Venta exitosa!',
      'El accesorio ha salido del inventario de Lola',
      ui.ButtonSet.OK
    );

  } else {
    form.getRange('N3:N23').clearContent();
  }

}

// ------------------------------------- EN DESARROLLO ------------------------------------- //
function eliminarEntrada() {

  const file = SpreadsheetApp.getActiveSpreadsheet();
  const form = file.getSheetByName('Formularios');
  const ui = SpreadsheetApp.getUi();
  const entradas = file.getSheetByName('Entradas');
  var data = entradas.getDataRange().getValues();
  var entradasArray = [];

  var codigo = Browser.inputBox(
    'Eliminar entrada del Inventario de Lola',
    'Ingrese el codigo del accesorio:',
    Browser.Buttons.OK_CANCEL
  );

  if (codigo === 'cancel') {
    // El usuario canceló la operación
    return;
  }

  // Busco el codigo en Entradas
  for (let i = 1; i < data.length; i++) {

    if (data[i][0] == codigo) {

      entradasArray.push(
        data[i][6],
        data[i][11]
      ); // Almacenamos la entrada encontrada

    }

  }

  if (entradasArray.length == 0) {

    ui.alert(
      '⛔¡Error!',
      'No se encontró ninguna entrada con ese código',
      ui.ButtonSet.OK
    );

  }

  // Mostramos las entradas encontradas al usuario y le pedimos que elija una
  var entradaSeleccionada = Browser.inputBox(
    'Se encontraron las siguientes entradas:\n\n' + entradasArray.join('\n'),
    'Seleccione una entrada por número (1, 2, 3, ...)',
    Browser.Buttons.OK_CANCEL
  );

  if (entradaSeleccionada === 'cancel') {
    // El usuario canceló la operación
    return;
  }

  // Eliminamos la entrada seleccionada
  var filaAEliminar = parseInt(entradaSeleccionada) - 1; // Restamos 1 para ajustar el índice de fila
  entradas.deleteRow(filaAEliminar + 2); // Sumamos 2 para considerar la fila de encabezados

  Browser.msgBox('Entrada eliminada correctamente.');
}

// Funcion para obtener el mes, url y rango para el Import Range
function getImportRangeData() {

  const file = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = file.getSheetByName('Import Range');
  var dataTable = sheet.getDataRange().getValues();
  var data = [];

  for (let i = 1; i < dataTable.length; i++) {

    data.push([
      dataTable[i][0], // almaceno el mes
      dataTable[i][1], // almaceno el url
      dataTable[i][2], // almaceno el rango
    ]);

  }

  return data;

}

// Funcion para automatizar formula Import Range
function setImportRangeFormula() {

  const file = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = file.getSheetByName('Reporte Global');
  const values = getImportRangeData();
  var row = 10;
  var column = 3;
  const ui = SpreadsheetApp.getUi();

  Logger.log(values[0])
  for (let i = 0; i < values.length; i++) {

    if (values[i][1] != "") {

      var url = values[i][1];
      var cadena = values[i][2];
      var formula = '=IMPORTRANGE("' + url + '"; "' + cadena + '")';
      sheet.getRange(row, column + i).setFormula(formula);
      //Logger.log('La columna seria ' + i + '+3' + 'y la url es:' + url + '      cadena' + cadena);

    }

  }

  ui.alert(
    '🎉¡Éxito!',
    'Tabla de Estado Resultado actualizada',
    ui.ButtonSet.OK
  )

  /*try {

    var importedData = SpreadsheetApp.openById(sourceSheetId).getSheetByName(sheet).getRange(rango).getValues();
    // Aquí puedes hacer lo que necesites con importedData (por ejemplo, escribirlo en otra hoja).
    Logger.log('Datos importados desde ' + sourceSheetId + ', rango ' + rango);

  } catch (error) {

    Logger.log('Error al importar desde ' + sourceSheetId + ', rango ' + rango + ': ' + error.message);

  }*/


}

// Funcion para permitir acceso a los archivos de reporte
function addImportRangePermission() {

  const file = SpreadsheetApp.getActiveSpreadsheet();
  const ssId = file.getId(); // id of the spreadsheet to add permission to import
  var values = getImportRangeData();
  //const donorId = '1GrELZHlEKu_QbBVqv...';  // donor or source spreadsheet id, you should get it somewhere

  const token = ScriptApp.getOAuthToken();

  const params = {
    method: 'post',
    headers: {
      Authorization: 'Bearer ' + token,
    },
    muteHttpExceptions: true
  };

  // Autorizo a cada archivo
  for (let i = 0; i < values.length; i++) {

    // Validamos que haya URL
    if (values[i][1] != "") {

      const donorId = values[i][1];

      // adding permission by fetching this url
      const url = `https://docs.google.com/spreadsheets/d/${ssId}/externaldata/addimportrangepermissions?donorDocId=${donorId}`;

      UrlFetchApp.fetch(url, params);

    }

  }

  setImportRangeFormula()

}


// Funcion para 
function setQueryGastosTotales() {

  const file = SpreadsheetApp.getActiveSpreadsheet();
  const ssId = file.getId(); // id of the spreadsheet to add permission to import
  var values = getImportRangeData();

  //const url = `https://docs.google.com/spreadsheets/d/${ssId}/externaldata/addimportrangepermissions?donorDocId=${donorId}`;

  for (let i = 0; i < values.length; i++) {

    // Validamos que haya URL
    if (values[i][1] != "") {

    }

  }
  var formula = '={QUERY(IMPORTRANGE("' + datos + '"; "' + consulta + '"))';


  //sheet.getRange(row, column + i).setFormula(formula);

}


