/**
 * @file trigger_locker.js
 * @description Contiene el activador basado en tiempo para la unificación de datos en 'Datos Locker'.
 */

/**
 * Clave utilizada para guardar el estado (conteo de filas) en PropertiesService.
 */
const LOCKER_STATE_KEY = 'LOCKER_DATA_STATE';

/**
 * Esta función se ejecuta periódicamente mediante un activador de tiempo (ej. cada 5 minutos).
 * Comprueba si el número de filas en las hojas de BBDD ha cambiado desde la última ejecución.
 * Si hay cambios, ejecuta `unificarDatosLocker()` y actualiza el estado.
 *
 * Para que funcione, este activador debe ser instalado manualmente desde el editor de Apps Script:
 * 1. Ir a "Activadores" (icono del reloj).
 * 2. Hacer clic en "Añadir activador".
 * 3. Elegir la función: verificarCambiosYUnificar
 * 4. Elegir la fuente del evento: "Según el tiempo".
 * 5. Elegir el tipo de activador: "Temporizador por minutos".
 * 6. Elegir el intervalo: "Cada 5 minutos" (o el intervalo que prefieras).
 * 7. Guardar.
 */
function verificarCambiosYUnificar() {
  try {
    // ID de la hoja de cálculo. Es más robusto usar openById para activadores.
    const SPREADSHEET_ID = '1g3K7lGBzD-KfAg8xpyW0RtbkfSYd-sAwcGo-zXSIDC0';
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    if (!ss) throw new Error(`No se pudo abrir la hoja de cálculo con ID: ${SPREADSHEET_ID}`);

    const cobrarCandidates = Array.from(new Set([
      'bbdd reservas horas complementarias',
      (typeof SHEET_RES_COBRAR_NAME === 'string' && SHEET_RES_COBRAR_NAME) ? SHEET_RES_COBRAR_NAME : '',
      SHEET_RES_COBRAR,
      typeof SHEET_RES_COBRAR_ALT === 'undefined' ? '' : SHEET_RES_COBRAR_ALT
    ].filter(Boolean)));

    const TARGET_SHEETS = [
      "bbdd reservas horas trabajar",
      ...cobrarCandidates,
      "bbdd reservas horas librar"
    ];

    // 1. Obtener el estado guardado (conteo de filas anterior)
    const scriptProperties = PropertiesService.getScriptProperties();
    const estadoGuardado = JSON.parse(scriptProperties.getProperty(LOCKER_STATE_KEY) || '{}');

    // 2. Obtener el estado actual (conteo de filas actual)
    const estadoActual = {};
    let totalFilasActual = 0;
    TARGET_SHEETS.forEach(sheetName => {
      const sh = ss.getSheetByName(sheetName);
      const rowCount = sh ? sh.getLastRow() : 0;
      estadoActual[sheetName] = rowCount;
      totalFilasActual += rowCount;
    });

    // 3. Comparar con el estado guardado
    let totalFilasGuardado = 0;
    Object.values(estadoGuardado).forEach(count => totalFilasGuardado += count);

    if (totalFilasActual !== totalFilasGuardado) {
      Logger.log(`Cambio detectado. Filas antes: ${totalFilasGuardado}, ahora: ${totalFilasActual}. Ejecutando unificación.`);
      
      // Ejecutar la función principal de unificación
      unificarDatosLocker();

      // Guardar el nuevo estado para la próxima comprobación
      scriptProperties.setProperty(LOCKER_STATE_KEY, JSON.stringify(estadoActual));
      Logger.log('Unificación completada y nuevo estado guardado.');

    } else {
      Logger.log('No se detectaron cambios en el número de filas. No se ejecuta la unificación.');
    }
  } catch (err) {
    Logger.log('Error en el activador verificarCambiosYUnificar: ' + err.toString());
  }
}
