/**
 * config.gs
 * Configuración centralizada de IDs, constantes globales y referencias compartidas.
 * Otros módulos (usuarios, reservas, reportes, etc.) dependen de estas constantes.
 */

/** Identificador de la hoja principal de la aplicación. */
const SPREADSHEET_ID = '1g3K7lGBzD-KfAg8xpyW0RtbkfSYd-sAwcGo-zXSIDC0';

/** Nombres de hojas dentro del libro principal. */
const SHEET_HORAS          = 'Horas trabajar';
const SHEET_RES            = 'bbdd reservas horas trabajar';
const SHEET_RES_COBRAR     = 'bbdd reservas horas complementarias';
const SHEET_RES_COBRAR_ALT = 'bbdd reserva horas complementarias';
const SHEET_HORAS_LIBRAR   = 'Horas librar';
const SHEET_RES_LIBRAR     = 'bbdd reservas horas librar';
const SHEET_DATOS_LOCKER   = 'Datos Locker';
const SHEET_ADMIN          = 'admin';
const SHEET_EMPLEADOS      = 'Empleados';
const SHEET_CAMPANIAS      = 'Campañas';
const SHEET_LOGS           = 'logs';

/** Propiedad global utilizada para generar IDs correlativos de reserva. */
const RESERVA_ID_PROP_KEY = 'RESERVA_ID_COUNTER';

/** Definición de roles disponibles en la aplicación. */
const ROLE = {
  ADMIN: 'admin',
  GESTOR: 'gestor',
  USER : 'usuario',
  USUARIO: 'usuario'
};

/** Columnas base utilizadas en las hojas de disponibilidad. */
const COL = { CAMPAÑA: 0, FECHA: 1, FRANJA: 3, DISPON: 5 };

/** Recursos estáticos utilizados en correos y vistas. */
const URL_LOGO_INTELCIA = 'https://drive.google.com/uc?export=view&id=1vNS8n_vYYJL9VQKx7jZxjZmMXUz0uECG';
const URL_APP           = 'https://script.google.com/a/macros/intelcia.com/s/AKfycbwMMz5v1uFYTRIboy3rahRv91cfAGXyRbWP_UdlaZr_44LZq9LwyB7TemfGToytCHdmgw/exec';
const PLANTILLAS_FOLDER_ID = '1i_8bEEuTzYU30q2nILSyYn33eRAqLa2C';

/** Conexión principal a la hoja de cálculo y referencias cacheadas. */
const ss               = SpreadsheetApp.openById(SPREADSHEET_ID);
const sheetHoras       = ss.getSheetByName(SHEET_HORAS);
const sheetHorasLibrar = ss.getSheetByName(SHEET_HORAS_LIBRAR);
const sheetResTrabajar = ss.getSheetByName(SHEET_RES) || createReservaSheet(ss, SHEET_RES);
const sheetResCobrar   = getSheetByPossibleNames(
  [SHEET_RES_COBRAR, SHEET_RES_COBRAR_ALT],
  () => createReservaSheet(ss, SHEET_RES_COBRAR)
);
const sheetResLibrar   = ss.getSheetByName(SHEET_RES_LIBRAR) || createReservaSheet(ss, SHEET_RES_LIBRAR);
const sheetDatosLocker = ss.getSheetByName(SHEET_DATOS_LOCKER);
const sheetAdmin       = ss.getSheetByName(SHEET_ADMIN);
const sheetEmpleados   = ss.getSheetByName(SHEET_EMPLEADOS);
const sheetCampanias   = ss.getSheetByName(SHEET_CAMPANIAS);
const sheetLogs        = ss.getSheetByName(SHEET_LOGS) || createLogsSheet(ss, SHEET_LOGS);

/** Nombre efectivo de la hoja de reservas de cobro, respetando alias previos. */
const SHEET_RES_COBRAR_NAME = sheetResCobrar ? sheetResCobrar.getName() : SHEET_RES_COBRAR;

/** Zona horaria base de la aplicación. */
const SCRIPT_PROPS           = PropertiesService.getScriptProperties();
const TIMEZONE_OVERRIDE_PROP = SCRIPT_PROPS ? SCRIPT_PROPS.getProperty('APP_TIMEZONE_OVERRIDE') : '';
const SCRIPT_TIME_ZONE       = Session.getScriptTimeZone();
const SPREADSHEET_TIME_ZONE  = ss.getSpreadsheetTimeZone();
const DEFAULT_TIME_ZONE      = 'Europe/Madrid';

const APP_TIMEZONE = TIMEZONE_OVERRIDE_PROP ||
  SCRIPT_TIME_ZONE ||
  SPREADSHEET_TIME_ZONE ||
  DEFAULT_TIME_ZONE;

/**
 * Devuelve la zona horaria configurada para la app.
 * @return {string}
 */
function getAppTimeZone() {
  return APP_TIMEZONE || Session.getScriptTimeZone();
}

/**
 * Crea una hoja de reservas con la estructura estándar.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet
 * @param {string} name
 * @return {GoogleAppsScript.Spreadsheet.Sheet}
 */
function createReservaSheet(spreadsheet, name) {
  const sh = spreadsheet.insertSheet(name);
  sh.appendRow(['Campaña', 'Fecha', 'HORAS', 'FRANJA', 'key', 'Correo', 'Tipo petición']);
  sh.setFrozenRows(1);
  sh.getRange(1, 11).setValue('ID reserva');
  return sh;
}

/**
 * Crea la hoja de logs con el encabezado estándar.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet
 * @param {string} name
 * @return {GoogleAppsScript.Spreadsheet.Sheet}
 */
function createLogsSheet(spreadsheet, name) {
  const sh = spreadsheet.insertSheet(name);
  sh.appendRow(['Fecha', 'Hora', 'Correo', 'Tipo de evento']);
  sh.setFrozenRows(1);
  return sh;
}

/**
 * Localiza una hoja respetando posibles alias; si no existe y se proporciona
 * un generador, lo ejecuta.
 *
 * @param {string|string[]} names Lista de nombres objetivo o alias.
 * @param {Function} [fallbackFactory] Función que crea la hoja en caso de ausencia.
 * @return {GoogleAppsScript.Spreadsheet.Sheet|null}
 */
function getSheetByPossibleNames(names, fallbackFactory) {
  const list = Array.isArray(names) ? names : [names];

  for (let i = 0; i < list.length; i++) {
    const name = list[i];
    if (!name) continue;
    const sheet = ss.getSheetByName(name);
    if (sheet) return sheet;
  }

  const lowerTargets = list
    .map(name => (name == null ? '' : String(name).toLowerCase()))
    .filter(Boolean);

  if (lowerTargets.length) {
    const sheets = ss.getSheets();
    for (let i = 0; i < sheets.length; i++) {
      const sheet = sheets[i];
      const sheetName = sheet.getName();
      if (lowerTargets.includes(sheetName.toLowerCase())) {
        return sheet;
      }
    }
  }

  return typeof fallbackFactory === 'function' ? fallbackFactory() : null;
}
