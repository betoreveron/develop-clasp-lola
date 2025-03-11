/**
 * Maneja las solicitudes GET y devuelve una plantilla HTML evaluada.
 *
 * @return {HtmlOutput} La salida HTML evaluada de la plantilla "home".
 */
function doGet(e) {

  if (e.parameter.view === "register") {

    return goToRegisterView();

  } else {

    return HtmlService.createTemplateFromFile("home").evaluate();

  }

}

function goToRegisterView() {
  var template = HtmlService.createTemplateFromFile("register");
  template.clientes = getOptions(); //Obtengo las opciones para el dropdown
  return template.evaluate();
}

//esta funcion es para traer los parametros de home.html (css y js)
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
