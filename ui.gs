/**
 * ui.gs
 * Funciones de interfaz: renderizado de plantillas HTML para la aplicación web.
 */

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Bolsa de horas')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getIndexHtml() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .getContent();
}

/**
 * Permite incluir fragmentos HTML mediante <?!= include('nombre'); ?>.
 * @param {string} filename
 * @return {string}
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
