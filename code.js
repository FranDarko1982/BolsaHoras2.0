/************************************************************
 *  CONFIGURACIÓNAKA
 ************************************************************/
const SPREADSHEET_ID    = '1g3K7lGBzD-KfAg8xpyW0RtbkfSYd-sAwcGo-zXSIDC0';
const SHEET_HORAS        = 'Horas trabajar';
const SHEET_RES          = 'bbdd reservas horas trabajar';
const SHEET_HORAS_LIBRAR = 'Horas librar';
const SHEET_RES_LIBRAR   = 'bbdd reservas horas librar';
const SHEET_DATOS_LOCKER = 'Datos Locker';
const SHEET_ADMIN        = 'admin';

const RESERVA_ID_PROP_KEY = 'RESERVA_ID_COUNTER';

const ROLE = {
  ADMIN: 'admin',
  USER : 'user'
};

// ÍNDICES
const COL = { CAMPAÑA: 0, FECHA: 1, FRANJA: 3, DISPON: 5 };

// PON AQUÍ LA URL DEL LOGO Y EL ENLACE A TU APP
const URL_LOGO_INTELCIA = "https://drive.google.com/uc?export=view&id=1vNS8n_vYYJL9VQKx7jZxjZmMXUz0uECG";
const URL_APP           = "https://script.google.com/a/macros/intelcia.com/s/AKfycbwMMz5v1uFYTRIboy3rahRv91cfAGXyRbWP_UdlaZr_44LZq9LwyB7TemfGToytCHdmgw/exec"; 

// ——————————————————————————————————————————————
// Cachear referencias a las hojas para no repetir getSheetByName
// ——————————————————————————————————————————————
const ss               = SpreadsheetApp.openById(SPREADSHEET_ID);
const sheetHoras       = ss.getSheetByName(SHEET_HORAS);
const sheetHorasLibrar = ss.getSheetByName(SHEET_HORAS_LIBRAR);
const sheetResTrabajar = ss.getSheetByName(SHEET_RES)        || createReservaSheet(ss, SHEET_RES);
const sheetResLibrar   = ss.getSheetByName(SHEET_RES_LIBRAR) || createReservaSheet(ss, SHEET_RES_LIBRAR);
const sheetDatosLocker = ss.getSheetByName(SHEET_DATOS_LOCKER);

const sheetAdmin       = ss.getSheetByName(SHEET_ADMIN);

/************************************************************
 *  CONFIGURACIÓN: campañas que pueden COBRAR
 ************************************************************/
const SHEET_CONFIG_COBRAR = 'CONFIG_cobrar';

function puedeCobrarCampania(campania) {
  if (!campania) return false;

  const sh = ss.getSheetByName(SHEET_CONFIG_COBRAR);
  if (!sh) return false;

  // Obtiene todas las campañas desde la columna A (desde fila 2)
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return false;

  const valores = sh.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  const listaCampanias = valores
    .map(v => String(v || '').trim().toLowerCase())
    .filter(Boolean);

  return listaCampanias.includes(String(campania).trim().toLowerCase());
}

/************************************************************
 *  Helper: contexto de usuario y roles
 ************************************************************/
function getCurrentUserEmail() {
  try {
    return (Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  } catch (err) {
    return '';
  }
}

function getAdminEmails() {
  if (!sheetAdmin) return [];
  const lastRow = sheetAdmin.getLastRow();
  if (lastRow < 2) return [];
  const values = sheetAdmin.getRange(2, 1, lastRow - 1, 1).getValues();
  return values
    .map(row => String(row[0] || '').trim().toLowerCase())
    .filter(Boolean);
}

function userHasAdminRole(email) {
  if (!email) return false;
  const admins = getAdminEmails();
  return admins.includes(email.toLowerCase());
}

function getUserContext() {
  const email = getCurrentUserEmail();
  const isAdmin = userHasAdminRole(email);
  return {
    email,
    isAdmin,
    role: isAdmin ? ROLE.ADMIN : ROLE.USER
  };
}

function getAccessContext() {
  const { role, email } = getUserContext();
  const baseSections = ['inicio', 'calendario', 'mis-reservas', 'ayuda'];
  const sections = role === ROLE.ADMIN
    ? [...baseSections.slice(0, 3), 'reportes', ...baseSections.slice(3)]
    : baseSections;

  return {
    role,
    email,
    sections
  };
}


/************************************************************
 *  1) SERVIR EL HTML
 ************************************************************/
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


/************************************************************
 *  Helper: Primera fila libre (columna A) después del encabezado
 ************************************************************/
function getFirstFreeRow(sh) {
  const values = sh.getRange(2, 1, sh.getMaxRows() - 1, 1).getValues().flat();
  const idx = values.findIndex(v => v === '' || v == null);
  return (idx >= 0) ? idx + 2 : sh.getLastRow() + 1;
}

/************************************************************
 *  2) LISTA ÚNICA DE CAMPAÑAS
 ************************************************************/
function getCampanias() {
  const sh = sheetHoras;
  if (!sh) throw new Error("No se ha encontrado la hoja 'Horas trabajar'");
  const vals = sh.getRange(2, COL.CAMPAÑA + 1,
              sh.getLastRow() - 1, 1)
    .getValues()
    .flat();
  return [...new Set(vals)].filter(c => c).sort();
}

function getCampaniasLibrar() {
  const sh = sheetHorasLibrar;
  if (!sh) throw new Error("No se ha encontrado la hoja 'Horas librar'");
  const vals = sh.getRange(2, COL.CAMPAÑA + 1,
              sh.getLastRow() - 1, 1)
    .getValues()
    .flat();
  return [...new Set(vals)].filter(c => c).sort();
}

/************************************************************
 *  3) HORAS DISPONIBLES - TRABAJAR
 ************************************************************/
function getHorasDisponibles(campania, startISO, endISO) {
  return _getHorasDisponibles(sheetHoras, sheetResTrabajar, campania, startISO, endISO);
}

/************************************************************
 *  3b) HORAS DISPONIBLES - LIBRAR
 ************************************************************/
function getHorasDisponiblesLibrar(campania, startISO, endISO) {
  return _getHorasDisponibles(sheetHorasLibrar, sheetResLibrar, campania, startISO, endISO);
}

// FUNCIÓN AUXILIAR REUTILIZABLE
function _getHorasDisponibles(sheetHorasObj, sheetResObj, campania, startISO, endISO) {
  const tz    = Session.getScriptTimeZone();
  const start = new Date(startISO);
  const end   = new Date(endISO);

  const shBase = sheetHorasObj;
  const rows   = shBase.getDataRange().getValues().slice(1);
  const eventos = [];

  rows.forEach(r => {
    if (campania && r[COL.CAMPAÑA] !== campania) return;

    const fechaStr    = r[COL.FECHA];    // "dd/MM/yyyy" o Date
    const franja      = r[COL.FRANJA];   // "08:00-09:00"
    const disponible  = +r[COL.DISPON] || 0;
    if (!fechaStr || !franja || disponible <= 0) return;

    // 1) Parsear fecha
    let yyyy, mm, dd;
    if (typeof fechaStr === 'string') {
      [dd, mm, yyyy] = fechaStr.split('/').map(Number);
    } else {
      dd = fechaStr.getDate();
      mm = fechaStr.getMonth() + 1;
      yyyy = fechaStr.getFullYear();
    }

    // 2) Hora de inicio
    const [hIniStr] = franja.split('-');
    const [h0, m0]   = hIniStr.split(':').map(Number);

    // 3) Construir Date y descartar fuera de rango
    const inicio = new Date(yyyy, mm - 1, dd, h0, m0, 0, 0);
    if (inicio < start || inicio >= end) return;

    // 4) Pushear evento usando “disponible”
    eventos.push({
      start    : inicio.toISOString(),
      end      : new Date(inicio.getTime() + 3600000).toISOString(),
      title    : disponible + ' h',
      className: [
        sheetHorasObj === sheetHorasLibrar
          ? 'fc-event-horas-librar'
          : 'fc-event-horas'
      ]
    });
  });

  return eventos;
}

/************************************************************
 *  FUNCIÓN: Comprobar disponibilidad continua
 ************************************************************/

function comprobarDisponibilidadContinua(sheetHorasObj, campania, startISO, horas) {
  const tz  = Session.getScriptTimeZone();
  const ini = new Date(startISO);
  const fechaStr = Utilities.formatDate(ini, tz, 'dd/MM/yyyy');

  // Construimos un SET con todas las franjas disponibles NORMALIZADAS
  const datos = sheetHorasObj.getDataRange().getValues();
  const disponibles = new Set();

  for (let i = 1; i < datos.length; i++) { // saltar encabezado
    const r = datos[i];

    if (campania && r[COL.CAMPAÑA] !== campania) continue;

    // Normalizar la fecha de la celda a dd/MM/yyyy
    const celdaFecha = r[COL.FECHA];
    const fechaCelda = (celdaFecha instanceof Date)
      ? Utilities.formatDate(celdaFecha, tz, 'dd/MM/yyyy')
      : String(celdaFecha).trim();

    if (fechaCelda !== fechaStr) continue;

    const disp = +r[COL.DISPON] || 0;
    if (disp <= 0) continue;

    const franjaNorm = _normalizarFranja(r[COL.FRANJA]);
    if (franjaNorm) disponibles.add(franjaNorm); // p.ej. "13:00-14:00"
  }

  // Comprobar que existen todas las franjas consecutivas solicitadas
  for (let k = 0; k < horas; k++) {
    const dIni = new Date(ini.getTime() + k * 3600000);
    const dFin = new Date(dIni.getTime() + 3600000);
    const esperada = _normalizarFranja(
      Utilities.formatDate(dIni, tz, 'HH:mm') + '-' +
      Utilities.formatDate(dFin, tz, 'HH:mm')
    );

    if (!disponibles.has(esperada)) {
      Logger.log('Falta franja: ' + esperada);
      Logger.log('Disponibles: ' + Array.from(disponibles).join(', '));
      return false;
    }
  }

  return true;
}

/* Helper: normaliza "8:00 - 9:00" -> "08:00-09:00" (sin espacios) */
function _normalizarFranja(franja) {
  const s = String(franja || '').replace(/\s+/g, '');
  const m = s.match(/^(\d{1,2}):(\d{2})-(\d{1,2}):(\d{2})$/);
  if (!m) return null;
  const h1 = ('0' + m[1]).slice(-2), min1 = m[2];
  const h2 = ('0' + m[3]).slice(-2), min2 = m[4];
  return `${h1}:${min1}-${h2}:${min2}`;
}



/************************************************************
 *  4) RESERVAR UNA O VARIAS HORAS (TRABAJAR)
 ************************************************************/
function reservarVariasHoras(campania, startISO, horas) {
  return _reservarVariasHoras('Trabajar', sheetResTrabajar, campania, startISO, horas);
}

/************************************************************
 *  4c) RESERVAR VARIAS HORAS (COBRAR)
 ************************************************************/
function reservarVariasHorasCobrar(campania, startISO, horas) {
  return _reservarVariasHoras('Cobrar', sheetResTrabajar, campania, startISO, horas);
}

/************************************************************
 *  4b) RESERVAR UNA O VARIAS HORAS (LIBRAR)
 ************************************************************/
function reservarHoraLibrar(campania, startISO) {
  return _reservarHora('Librar', sheetResLibrar, campania, startISO);
}
function reservarVariasHorasLibrar(campania, startISO, horas) {
  return _reservarVariasHoras('Librar', sheetResLibrar, campania, startISO, horas);
}

// FUNCIÓN AUXILIAR RESERVA 1
function _reservarHora(tipo, sheetResObj, campania, startISO) {
  const tz   = Session.getScriptTimeZone();
  const ini  = new Date(startISO);
  const fin  = new Date(ini.getTime() + 3600000);
  const franja = Utilities.formatDate(ini, tz, 'HH:00') + '-' +
                 Utilities.formatDate(fin, tz, 'HH:00');
  const fechaSerial = Math.floor(
    (Date.UTC(ini.getFullYear(), ini.getMonth(), ini.getDate()) -
     Date.UTC(1899, 11, 30)) / 86400000
  );
  const email = Session.getActiveUser().getEmail() || '';
  const fechaSoloTexto = Utilities.formatDate(ini, tz, 'dd/MM/yyyy');
  const keyReserva = campania + fechaSerial + franja + email + tipo;
  const reservaId = generateReservaId();

  // Hoja de reservas (usamos la hoja cacheada)
  let sh = sheetResObj;
  ensureReservaIdColumn(sh);

  // Si usamos la versión que crea hoja automáticamente, 
  // asegúrate de que createReservaSheet soporta sólo nombre, 
  // así que aquí asumimos sheetResObj ya existe.

  // Insertar en la primera fila libre
  const targetRow = getFirstFreeRow(sh);
  sh.getRange(targetRow, 1, 1, 7).setValues([[
    campania,
    fechaSoloTexto,
    1,
    franja,
    keyReserva,
    email,
    tipo
  ]]);
  sh.getRange(targetRow, 11).setValue(reservaId);

  // Envío de email de confirmación
  sendConfirmationEmail(tipo, campania, ini, franja.split('-')[0], franja.split('-')[1], 1, email, reservaId);

  return `Solicitud registrada para ${campania} ${franja} (${fechaSoloTexto}) [${tipo}]. ID de reserva: ${reservaId}`;
}

/************************************************************
 *  4b) RESERVAR VARIAS HORAS (LIBRAR/TRABAJAR/COBRAR)
 ************************************************************/
function _reservarVariasHoras(tipo, sheetResObj, campania, startISO, horas) {
  const tz    = Session.getScriptTimeZone();
  const ini   = new Date(startISO);
  const email = Session.getActiveUser().getEmail() || '';
  const sh    = sheetResObj;
  ensureReservaIdColumn(sh);
  const reservaId = generateReservaId();

  // *** NUEVO: Validar disponibilidad continua antes de reservar ***
  let sheetHorasObj = (tipo === 'Librar') ? sheetHorasLibrar : sheetHoras;

  // Validar si la campaña permite cobrar
  if (tipo === 'Cobrar' && !puedeCobrarCampania(campania)) {
    throw new Error(`La campaña "${campania}" no tiene habilitada la opción de COBRAR horas.`);
  }

  if (!comprobarDisponibilidadContinua(sheetHorasObj, campania, startISO, horas)) {
    throw new Error(
      "¡Ay miarma! El tramo que has elegido no está completamente disponible, hay horas intermedias pilladas. " +
      "Prueba con otro horario o menos horas seguidas. No te me vengas arriba solicitando de gratis, ¿eh?"
    );
  }

  const firstRow = getFirstFreeRow(sh);
  const rows     = [];
  let primerFranja, ultimaFranja, fechaTexto;

  for (let h = 0; h < horas; h++) {
    const inicio      = new Date(ini.getTime() + h * 3600000);
    const strFranja   = formatFranja(inicio, tz);
    const fechaSolo   = Utilities.formatDate(inicio, tz, 'dd/MM/yyyy');
    const fechaSerial = toSerialDate(inicio);
    const keyReserva  = `${campania}${fechaSerial}${strFranja}${email}${tipo}`;

    rows.push([
      campania,
      fechaSolo,
      1,
      strFranja,
      keyReserva,
      email,
      tipo
    ]);

    if (h === 0)         primerFranja = strFranja.split('-')[0];
    if (h === horas - 1) ultimaFranja = strFranja.split('-')[1];
    fechaTexto = fechaSolo;
  }

  // Inserción en bloque
  sh.getRange(firstRow, 1, rows.length, 7).setValues(rows);
  const reservaIdValues = rows.map(() => [reservaId]);
  sh.getRange(firstRow, 11, reservaIdValues.length, 1).setValues(reservaIdValues);

  // Envío de email
  sendConfirmationEmail(tipo, campania, ini, primerFranja, ultimaFranja, horas, email, reservaId);

  return `Solicitud registrada para ${campania} ${fechaTexto} de ${primerFranja} a ${ultimaFranja} (${horas}h) [${tipo}]. ID de reserva: ${reservaId}. Se ha enviado un email de confirmación.`;
}

// --- HELPERS EXTRAÍDOS PARA LA FUNCIÓN ANTERIOR ---

function createReservaSheet(ss, name) {
  const sh = ss.insertSheet(name);
  sh.appendRow(['Campaña','Fecha','HORAS','FRANJA','key','Correo','Tipo petición']);
  sh.setFrozenRows(1);
  sh.getRange(1, 11).setValue('ID reserva');
  return sh;
}

function ensureReservaIdColumn(sheet) {
  if (!sheet) return;
  const headerCell = sheet.getRange(1, 11);
  if (!String(headerCell.getValue() || '').trim()) {
    headerCell.setValue('ID reserva');
  }
}

function generateReservaId() {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const props = PropertiesService.getScriptProperties();
    let current = parseInt(props.getProperty(RESERVA_ID_PROP_KEY), 10);
    if (isNaN(current) || current < 0) {
      current = computeHighestReservaId();
    }
    const next = (current || 0) + 1;
    props.setProperty(RESERVA_ID_PROP_KEY, String(next));
    return formatReservaIdNumber(next);
  } finally {
    lock.releaseLock();
  }
}

function computeHighestReservaId() {
  const highest = Math.max(
    getHighestReservaIdFromSheet(sheetResTrabajar),
    getHighestReservaIdFromSheet(sheetResLibrar)
  );
  const props = PropertiesService.getScriptProperties();
  props.setProperty(RESERVA_ID_PROP_KEY, String(highest || 0));
  return highest || 0;
}

function getHighestReservaIdFromSheet(sheet) {
  if (!sheet) return 0;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  const values = sheet.getRange(2, 11, lastRow - 1, 1).getValues().flat();
  let max = 0;
  values.forEach(id => {
    const parsed = parseReservaIdNumber(id);
    if (parsed != null && parsed > max) max = parsed;
  });
  return max;
}

function parseReservaIdNumber(value) {
  if (value == null) return null;
  const match = String(value).trim().match(/^BH(\d{8})$/i);
  if (!match) return null;
  return parseInt(match[1], 10);
}

function formatReservaIdNumber(number) {
  return `BH${String(number).padStart(8, '0')}`;
}

function formatFranja(date, tz) {
  const inicio = Utilities.formatDate(date, tz, 'HH:00');
  const fin    = Utilities.formatDate(new Date(date.getTime() + 3600000), tz, 'HH:00');
  return `${inicio}-${fin}`;
}

function toSerialDate(date) {
  return Math.floor(
    (Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()) -
     Date.UTC(1899, 11, 30)) / 86400000
  );
}

function sendConfirmationEmail(tipo, campania, ini, primer, ultima, horas, email, reservaId) {
  if (!email || email.indexOf('@') < 0) return;

  const tz         = Session.getScriptTimeZone();
  const asunto     = `Solicitud registrada – Bolsa de horas (${tipo})`;
  const fechaLarga = `Madrid, a ${Utilities.formatDate(ini, tz, 'd')} de ${getNombreMes(ini.getMonth())} de ${ini.getFullYear()}`;
  const horario    = `${primer} - ${ultima}`;
  const idHtml     = reservaId ? `<li><b>ID de reserva:</b> ${reservaId}</li>` : '';

  const body = `
    <div style="font-family: Calibri, Arial, sans-serif; max-width:650px; margin:0 auto; background:#fff;">
      <div style="width:100%; text-align:center; margin:24px 0 32px;">
        <img src="${URL_LOGO_INTELCIA}" alt="Intelcia" style="width:100%; max-width:500px; border-radius:8px;">
      </div>
      <div style="color:#222; font-size:16px; margin-bottom:24px;">
        En ${fechaLarga}
      </div>
      <div style="margin-bottom:16px;">
        Desde <span style="color:#C9006C; font-weight:bold;">Intelcia Spanish Region</span> te informamos que tu solicitud ha sido aceptada. Te recordamos la petición que has realizado:
      </div>
      <ul style="color:#C9006C; font-size:16px; margin-left:32px;">
        ${idHtml}
        <li><b>Campaña:</b> ${campania}</li>
        <li><b>Fecha:</b> ${Utilities.formatDate(ini, tz, 'dd/MM/yyyy')}</li>
        <li><b>Horario:</b> ${horario}</li>
        <li><b>Horas:</b> ${horas}</li>
      </ul>
      <div style="font-size:15px; margin:14px 0 16px; color:#222;">
        Puedes cancelar esta solicitud hasta un día antes de la fecha que solicitaste trabajar o librar. Te recordamos que puedes realizar una nueva petición desde la aplicación 
        <span style="color:#C9006C; font-weight:bold;">Bolsa de horas</span> 
        <a href="${URL_APP}" style="color:#C9006C; text-decoration:underline;">(enlace a la aplicación)</a>.
      </div>
      <div style="margin-top:24px; font-size:15px;">Gracias</div>
    </div>
  `;

  MailApp.sendEmail({ to: email, subject: asunto, htmlBody: body });
}


/************************************************************
 *  5) CANCELAR RESERVA (TRABAJAR o LIBRAR)
 ************************************************************/
function cancelarReserva(key) {
  let ok = _cancelarReservaEnHoja(sheetResTrabajar, key);
  if (!ok) ok = _cancelarReservaEnHoja(sheetResLibrar, key);
  return ok;
}

function _cancelarReservaEnHoja(sheetResObj, key) {
  const sh = sheetResObj;
  if (!sh) return false;

  // Usamos TextFinder para localizar la fila de la reserva
  const finder = sh.createTextFinder(key).matchEntireCell(true).findNext();
  if (!finder) return false;

  const row = finder.getRow();
  // Obtener email antes de borrar
  const emailUsuario = sh.getRange(row, /* columna "Correo" es la 6ª */ 6).getValue();

  sh.deleteRow(row);

  // Envío de email de cancelación
  if (emailUsuario && String(emailUsuario).indexOf('@') > -1) {
    const hoy        = new Date();
    const tz         = Session.getScriptTimeZone();
    const fechaLarga = `Madrid, a ${Utilities.formatDate(hoy, tz, 'd')} de ${getNombreMes(hoy.getMonth())} de ${hoy.getFullYear()}`;
    const asunto     = "Solicitud cancelada – Bolsa de horas";

    const body = `
      <div style="font-family: Calibri, Arial, sans-serif; max-width:650px; margin:0 auto; background:#fff;">
        <div style="width:100%; text-align:center; margin:24px 0 32px;">
          <img src="${URL_LOGO_INTELCIA}" alt="Intelcia" style="width:100%; max-width:500px; border-radius:8px;">
        </div>
        <div style="color:#222; font-size:16px; margin-bottom:24px;">
          En ${fechaLarga}
        </div>
        <div style="margin-bottom:16px;">
          Desde <span style="color:#C9006C; font-weight:bold;">Intelcia Spanish Region</span> te informamos que tu solicitud ha sido cancelada.<br>
          Te recordamos que puedes realizar una nueva petición desde la aplicación 
          <span style="color:#C9006C; font-weight:bold;">Bolsa de horas</span>
          <a href="${URL_APP}" style="color:#C9006C; text-decoration:underline;">(enlace a la aplicación)</a>.
        </div>
        <div style="margin-top:24px; font-size:15px;">Gracias</div>
      </div>
    `;

    MailApp.sendEmail({ to: emailUsuario, subject: asunto, htmlBody: body });
  }

  return true;
}

/************************************************************
 *  5c) CANCELAR MÚLTIPLES RESERVAS (TRABAJAR o LIBRAR)
 ************************************************************/
function cancelarMultiplesReservas(keys) {
  if (!Array.isArray(keys) || keys.length === 0) {
    throw new Error('No se proporcionaron identificadores de reserva para cancelar.');
  }

  keys.forEach(key => {
    // Intenta cancelar en ambas hojas, no importa si falla en una.
    _cancelarReservaEnHoja(sheetResTrabajar, key);
    _cancelarReservaEnHoja(sheetResLibrar, key);
  });

  return true; // Asumimos éxito si no hay errores.
}

/************************************************************
 * 5b) ACTUALIZAR SOLICITUD EXISTENTE (SOLO ADMIN)
 ************************************************************/
function actualizarSolicitud(payload) {
  const { isAdmin } = getUserContext();
  if (!isAdmin) {
    throw new Error('No tienes permisos para editar solicitudes.');
  }

  const data = _normalizeEditarPayload(payload);

  const updatedTrabajar = _actualizarSolicitudEnHoja(sheetResTrabajar, data);
  if (updatedTrabajar) {
    return {
      message: 'Solicitud actualizada correctamente.',
      key: updatedTrabajar.key,
      sheet: updatedTrabajar.sheet
    };
  }

  const updatedLibrar = _actualizarSolicitudEnHoja(sheetResLibrar, data);
  if (updatedLibrar) {
    return {
      message: 'Solicitud actualizada correctamente.',
      key: updatedLibrar.key,
      sheet: updatedLibrar.sheet
    };
  }

  throw new Error('No se encontró la solicitud que intentas editar.');
}

function _normalizeEditarPayload(payload) {
  if (!payload || typeof payload !== 'object') {
    throw new Error('Datos de edición no válidos.');
  }

  const key = String(payload.key || '').trim();
  if (!key) {
    throw new Error('Falta el identificador de la solicitud.');
  }

  const campania = String(payload.campania || '').trim();
  if (!campania) {
    throw new Error('La campaña es obligatoria.');
  }

  const tipo = String(payload.tipo || '').trim() || 'Trabajar';
  const correo = String(payload.correo || '').trim();
  const horas = Number(payload.horas);
  if (!Number.isFinite(horas) || horas <= 0) {
    throw new Error('Las horas proporcionadas no son válidas.');
  }

  let start = null;
  if (payload.fechaISO) {
    const parsedISO = new Date(payload.fechaISO);
    if (!Number.isNaN(parsedISO.getTime())) {
      start = parsedISO;
    }
  }

  if ((!start || Number.isNaN(start.getTime())) && payload.fechaDisplay && payload.horaInicio) {
    const [dd, mm, yyyy] = String(payload.fechaDisplay).split('/');
    const [hh, minute] = String(payload.horaInicio).split(':');
    if (dd && mm && yyyy && hh && minute) {
      start = new Date(
        Number(yyyy),
        Number(mm) - 1,
        Number(dd),
        Number(hh),
        Number(minute),
        0,
        0
      );
    }
  }

  if (!start || Number.isNaN(start.getTime())) {
    throw new Error('No se pudo interpretar la fecha de la solicitud.');
  }

  const tz = Session.getScriptTimeZone();
  const totalMinutes = Math.round(horas * 60);
  const end = new Date(start.getTime() + totalMinutes * 60000);

  let franja = String(payload.franja || '').trim();
  const franjaNormalizada = _normalizarFranja(franja);
  if (franjaNormalizada) {
    franja = franjaNormalizada;
  } else {
    const inicioStr = Utilities.formatDate(start, tz, 'HH:mm');
    const finStr = Utilities.formatDate(end, tz, 'HH:mm');
    franja = `${inicioStr}-${finStr}`;
  }

  const fechaDisplay = payload.fechaDisplay
    ? String(payload.fechaDisplay).trim()
    : Utilities.formatDate(start, tz, 'dd/MM/yyyy');

  const estadoSolicitud = payload && payload.estadoSolicitud != null
    ? String(payload.estadoSolicitud).trim()
    : '';

  return {
    key,
    campania,
    tipo,
    correo,
    horas,
    start,
    end,
    franja,
    fechaDisplay,
    estadoSolicitud
  };
}

function _actualizarSolicitudEnHoja(sheet, data) {
  if (!sheet) return null;

  const finder = sheet.createTextFinder(data.key).matchEntireCell(true).findNext();
  if (!finder) return null;

  const row = finder.getRow();
  const headerValues = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
  if (!headerValues.length) return null;

  const headerMap = headerValues[0].reduce((map, header, index) => {
    const key = String(header || '').trim();
    if (key) map[key] = index;
    return map;
  }, {});

  const colCampania = _findFirstIndex(headerMap, ['Campaña', 'Campana']);
  const colFecha = _findFirstIndex(headerMap, ['Fecha', 'Fecha reserva', 'Fecha inicio', 'Fecha Inicio']);
  const colHoras = _findFirstIndex(headerMap, ['HORAS', 'Horas', 'Horas solicitadas', 'Horas reservadas']);
  const colFranja = _findFirstIndex(headerMap, ['FRANJA', 'Franja', 'Franja horaria', 'Horario', 'Horario reserva']);
  const colCorreo = _findFirstIndex(headerMap, ['Correo', 'Email', 'Mail']);
  const colTipo = _findFirstIndex(headerMap, ['Tipo', 'Tipo petición', 'Tipo Peticion', 'Tipo solicitud']);
  const colEstado = _findFirstIndex(headerMap, ['Estado solicitud', 'Estado', 'Estado Solicitud']);
  const colKey = _findFirstIndex(headerMap, ['key', 'Key', 'KEY']);

  let correo = data.correo;
  if (!correo && colCorreo >= 0) {
    correo = String(sheet.getRange(row, colCorreo + 1).getValue() || '').trim();
  }

  let tipoValue = data.tipo;
  if (!tipoValue && colTipo >= 0) {
    tipoValue = String(sheet.getRange(row, colTipo + 1).getValue() || '').trim();
  }

  const franjaParaGuardar = _normalizarFranja(data.franja) || data.franja;
  const fechaSerial = toSerialDate(data.start);
  const newKey = `${data.campania}${fechaSerial}${franjaParaGuardar}${correo}${tipoValue}`;

  if (colCampania >= 0) sheet.getRange(row, colCampania + 1).setValue(data.campania);
  if (colFecha >= 0) sheet.getRange(row, colFecha + 1).setValue(data.fechaDisplay);
  if (colHoras >= 0) sheet.getRange(row, colHoras + 1).setValue(data.horas);
  if (colFranja >= 0) sheet.getRange(row, colFranja + 1).setValue(franjaParaGuardar);
  if (colCorreo >= 0 && correo) sheet.getRange(row, colCorreo + 1).setValue(correo);
  if (colTipo >= 0 && tipoValue) sheet.getRange(row, colTipo + 1).setValue(tipoValue);
  if (colEstado >= 0 && data.estadoSolicitud) sheet.getRange(row, colEstado + 1).setValue(data.estadoSolicitud);
  if (colKey >= 0) sheet.getRange(row, colKey + 1).setValue(newKey);

  return {
    updated: true,
    key: newKey,
    row,
    sheet: sheet.getName()
  };
}

/************************************************************
 *  6) MOSTRAR SOLO LAS RESERVAS DEL USUARIO LOGADO (AMBOS)
 ************************************************************/
function getMisReservas() {
  const { email } = getUserContext();
  if (!email) return [];
  const reservasTrabajar = _getMisReservasDeHoja(sheetResTrabajar, email, 'Trabajar');
  const reservasLibrar   = _getMisReservasDeHoja(sheetResLibrar,   email, 'Librar');
  return [...reservasTrabajar, ...reservasLibrar].sort((a, b) => {
    const fA = parseDateDDMMYYYY(a.fecha);
    const fB = parseDateDDMMYYYY(b.fecha);
    if (fA < fB) return -1;
    if (fA > fB) return 1;
    if (a.tipo < b.tipo) return -1;
    if (a.tipo > b.tipo) return 1;
    return 0;
  });
}

function getSolicitudesAcumuladas() {
  const tz = Session.getScriptTimeZone();
  const { email, isAdmin } = getUserContext();

  if (!email && !isAdmin) {
    return [];
  }

  const normalizar = valor => String(valor || '').trim().toLowerCase();
  const todasSolicitudes = (() => {
    if (sheetDatosLocker) {
      const lockerSolicitudes = _getSolicitudesDeHoja(sheetDatosLocker, '', tz);
      if (Array.isArray(lockerSolicitudes) && lockerSolicitudes.length) {
        return lockerSolicitudes;
      }
    }
    const solicitudesTrabajar = _getSolicitudesDeHoja(sheetResTrabajar, 'Trabajar', tz);
    const solicitudesLibrar = _getSolicitudesDeHoja(sheetResLibrar, 'Librar', tz);
    return [...solicitudesTrabajar, ...solicitudesLibrar];
  })();

  const visibles = isAdmin
    ? todasSolicitudes
    : todasSolicitudes.filter(item => normalizar(item.correo) === email);

  return visibles
    .sort((a, b) => b._timestamp - a._timestamp || a.tipo.localeCompare(b.tipo))
    .map(({ _timestamp, ...rest }) => rest);
}

function _getMisReservasDeHoja(sh, email, tipoPeticion) {
  if (!sh) return [];
  const datos     = sh.getDataRange().getValues();
  const headerMap = datos[0].reduce((m, h, i) => (m[String(h).trim()] = i, m), {});
  const {
    Campaña: C_CAM,
    Fecha: C_FEC,
    HORAS: C_HOR,
    FRANJA: C_FRA,
    Correo: C_COR,
    key: C_KEY,
    ["Tipo petición"]: C_TIP,
    ["Tipo"]: C_TIP2,
    ["Estado solicitud"]: C_EST
  } = headerMap;

  const COL_TIPO = (C_TIP >= 0 ? C_TIP : C_TIP2);

  return datos.slice(1)
    .filter(r => String(r[C_COR] || '').trim().toLowerCase() === email)
    .map(r => ({
      campania       : r[C_CAM],
      fecha          : (r[C_FEC] instanceof Date)
                        ? Utilities.formatDate(r[C_FEC], Session.getScriptTimeZone(), 'dd/MM/yyyy')
                        : String(r[C_FEC] || ''),
      horas          : r[C_HOR],
      franja         : r[C_FRA],
      estado         : (C_EST >= 0 ? r[C_EST] : r[headerMap["Validación"]] || ''),
      key            : String(r[C_KEY] || ''),
      tipo           : String(r[COL_TIPO] || tipoPeticion),
      estadoSolicitud: C_EST >= 0 ? String(r[C_EST] || '') : ''
    }));
}

function _getSolicitudesDeHoja(sheet, defaultTipo, tz) {
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  if (!values.length) return [];

  const headerMap = values[0].reduce((map, header, index) => {
    const key = String(header || '').trim();
    if (key) map[key] = index;
    return map;
  }, {});

  const colCampania = _findFirstIndex(headerMap, ['Campaña', 'Campana']);
  const colCorreo = _findFirstIndex(headerMap, ['Correo', 'Email', 'Mail']);
  const colFecha = _findFirstIndex(headerMap, ['Fecha', 'Fecha reserva', 'Fecha inicio', 'Fecha Inicio']);
  const colFranja = _findFirstIndex(headerMap, ['FRANJA', 'Franja', 'Franja horaria', 'Horario', 'Horario reserva']);
  const colHoras = _findFirstIndex(headerMap, ['HORAS', 'Horas', 'Horas solicitadas', 'Horas reservadas']);
  const colTipo = _findFirstIndex(headerMap, ['Tipo', 'Tipo petición', 'Tipo Peticion', 'Tipo solicitud']);
  const colEstado = _findFirstIndex(headerMap, ['Estado solicitud', 'Estado', 'Estado Solicitud']);
  const colNumeroEmpleado = _findFirstIndex(headerMap, ['NºEmpleado', 'Nº Empleado', 'Num empleado', 'Numero empleado', 'NumeroEmpleado', 'NEmpleado']);
  const colKey = _findFirstIndex(headerMap, ['key', 'Key', 'KEY']);
  const colReservaId = _findFirstIndex(headerMap, ['ID reserva', 'ID Reserva', 'Id reserva', 'Reserva ID', 'ID']);

  const registros = [];

  values.slice(1).forEach(row => {
    if (!row.some(cell => cell !== '' && cell != null)) {
      return;
    }

    const fechaRaw = colFecha >= 0 ? row[colFecha] : '';
    const fechaDate = _parseToDate(fechaRaw);
    const fechaDisplay = fechaDate
      ? Utilities.formatDate(fechaDate, tz, 'dd/MM/yyyy')
      : String(fechaRaw || '').trim();

    const horasRaw = colHoras >= 0 ? row[colHoras] : '';
    let horasDisplay;
    if (typeof horasRaw === 'number') {
      horasDisplay = Number.isFinite(horasRaw) ? horasRaw : '';
    } else {
      horasDisplay = String(horasRaw || '').trim();
    }

    const tipoRaw = colTipo >= 0 ? row[colTipo] : defaultTipo;
    const tipoFormatted = (() => {
      const valor = String(tipoRaw || defaultTipo || '').trim();
      if (!valor) return defaultTipo || '';
      return valor.charAt(0).toUpperCase() + valor.slice(1).toLowerCase();
    })();

    const estadoFormatted = colEstado >= 0
      ? String(row[colEstado] != null ? row[colEstado] : '').trim()
      : '';
    const keyValue = colKey >= 0 ? String(row[colKey] != null ? row[colKey] : '').trim() : '';

    registros.push({
      campania: colCampania >= 0 ? String(row[colCampania] || '').trim() : '',
      correo: colCorreo >= 0 ? String(row[colCorreo] || '').trim() : '',
      fecha: fechaDisplay,
      franja: colFranja >= 0 ? String(row[colFranja] || '').trim() : '',
      horas: horasDisplay,
      tipo: tipoFormatted,
      estadoSolicitud: estadoFormatted,
      key: keyValue,
      numeroEmpleado: colNumeroEmpleado >= 0 ? String(row[colNumeroEmpleado] || '').trim() : '',
      reservaId: colReservaId >= 0 ? String(row[colReservaId] || '').trim() : '',
      _timestamp: fechaDate ? fechaDate.getTime() : 0
    });
  });

  return registros;
}

// Utilidad para ordenar fechas DD/MM/YYYY
function parseDateDDMMYYYY(str) {
  if (!str || typeof str !== "string") return new Date('2100-01-01');
  const [dd, mm, yyyy] = str.split("/");
  return new Date(Number(yyyy), Number(mm) - 1, Number(dd));
}

function _findFirstIndex(map, keys) {
  for (let i = 0; i < keys.length; i++) {
    const key = keys[i];
    if (Object.prototype.hasOwnProperty.call(map, key)) {
      return map[key];
    }
  }
  return -1;
}

// Utilidad para nombre del mes en español
function getNombreMes(numeroMes) {
  const meses = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
  return meses[numeroMes] || '';
}

function getInicioDashboardData() {
  const tz = Session.getScriptTimeZone();
  const hoy = _normalizeDate(new Date());
  const { email, isAdmin, role } = getUserContext();

  const chartMonths = [];
  const chartHours = {};
  for (let offset = 5; offset >= 0; offset--) {
    const ref = new Date(hoy.getFullYear(), hoy.getMonth() - offset, 1);
    const key = `${ref.getFullYear()}-${ref.getMonth()}`;
    chartMonths.push({ key, label: _formatMonthLabel(ref) });
    chartHours[key] = 0;
  }

  const resultadoBase = {
    role,
    usuarioNombre: _obtenerNombreUsuario(),
    lastUpdated: Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm'),
    metrics: {
      totalSolicitudes: 0,
      proximasSolicitudes: 0,
      totalHorasTrabajar: 0,
      totalHorasLibrar: 0,
      saldo: 0
    },
    proximasSolicitudes: [],
    chart: {
      labels: chartMonths.map(item => item.label),
      values: chartMonths.map(item => chartHours[item.key])
    }
  };

  if (!email && !isAdmin) {
    return resultadoBase;
  }

  const reservasVisibles = [];
  let horasTrabajar = 0;
  let horasLibrar = 0;

  [
    { sheet: sheetResTrabajar, tipo: 'Trabajar' },
    { sheet: sheetResLibrar, tipo: 'Librar' }
  ].forEach(({ sheet, tipo }) => {
    if (!sheet) return;
    const values = sheet.getDataRange().getValues();
    if (!values.length) return;

    const headerMap = values[0].reduce((map, header, idx) => {
      const key = String(header || '').trim();
      if (key) map[key] = idx;
      return map;
    }, {});

    values.slice(1).forEach(row => {
      if (!row.some(cell => cell !== '' && cell != null)) return;
      const reserva = _mapReservaDashboard(row, headerMap, tipo);
      if (!reserva) return;
      if (!isAdmin && reserva.correo !== email) return;

      reservasVisibles.push(reserva);
      if (reserva.tipo === 'Librar') {
        horasLibrar += reserva.horasNumero;
      } else {
        horasTrabajar += reserva.horasNumero;
      }

      if (reserva.fechaInicio) {
        const diffMonths = (hoy.getFullYear() - reserva.fechaInicio.getFullYear()) * 12 + (hoy.getMonth() - reserva.fechaInicio.getMonth());
        if (diffMonths >= 0 && diffMonths <= 5) {
          const key = `${reserva.fechaInicio.getFullYear()}-${reserva.fechaInicio.getMonth()}`;
          chartHours[key] = (chartHours[key] || 0) + reserva.horasNumero;
        }
      }
    });
  });

  const proximas = reservasVisibles
    .filter(reserva => reserva.fechaInicio && reserva.fechaInicio >= hoy)
    .sort((a, b) => a.fechaInicio - b.fechaInicio);

  const proximasFormateadas = proximas.slice(0, 5).map(reserva => {
    const fechaInicioStr = reserva.fechaInicio
      ? Utilities.formatDate(reserva.fechaInicio, tz, 'dd/MM/yyyy')
      : '';
    const fechaFinBase = reserva.fechaFin || reserva.fechaInicio;
    const fechaFinStr = fechaFinBase
      ? Utilities.formatDate(fechaFinBase, tz, 'dd/MM/yyyy')
      : '';
    const fechaDisplay = (fechaInicioStr && fechaFinStr && fechaInicioStr !== fechaFinStr)
      ? `${fechaInicioStr} - ${fechaFinStr}`
      : fechaInicioStr || fechaFinStr;

    return {
      campania: reserva.campania || reserva.sala,
      tipo: reserva.tipo,
      fecha: fechaDisplay,
      horario: reserva.franja || reserva.horario || reserva.horas
    };
  });

  return {
    role,
    usuarioNombre: resultadoBase.usuarioNombre,
    lastUpdated: resultadoBase.lastUpdated,
    metrics: {
      totalSolicitudes: reservasVisibles.length,
      proximasSolicitudes: proximas.length,
      totalHorasTrabajar: horasTrabajar,
      totalHorasLibrar: horasLibrar,
      saldo: horasLibrar - horasTrabajar
    },
    proximasSolicitudes: proximasFormateadas,
    chart: {
      labels: chartMonths.map(item => item.label),
      values: chartMonths.map(item => Number(chartHours[item.key] || 0))
    }
  };
}

function _mapReservaDashboard(row, headerMap, defaultTipo) {
  const getValue = (names) => {
    for (let i = 0; i < names.length; i++) {
      const name = names[i];
      if (name in headerMap) {
        return row[headerMap[name]];
      }
    }
    return '';
  };

  const fechaInicio = _parseToDate(getValue(['Fecha inicio', 'Fecha Inicio', 'Fecha', 'Fecha reserva']));
  const fechaFin = _parseToDate(getValue(['Fecha fin', 'Fecha Fin', 'Fecha final', 'Fecha fin reserva']));

  const campania = String(getValue(['Campaña', 'Campana']) || '').trim();
  const sala = String(getValue(['Sala', 'Salas', 'Campaña', 'Campana']) || '').trim();
  const centro = String(getValue(['Centro', 'Site', 'Oficina', 'Sede']) || '').trim();
  const ciudad = String(getValue(['Ciudad', 'Localidad', 'Ubicación', 'Ubicacion']) || '').trim();
  const franja = String(getValue(['FRANJA', 'Franja', 'Franja horaria', 'Horario', 'Horario reserva']) || '').trim();
  const horario = franja || String(getValue(['Horario', 'Horario reserva']) || '').trim();
  const horasRaw = getValue(['HORAS', 'Horas', 'Horas solicitadas', 'Horas reservadas']);
  const horas = String(horasRaw != null ? horasRaw : '').trim();
  let horasNumero = (() => {
    if (typeof horasRaw === 'number') {
      return Number.isFinite(horasRaw) ? horasRaw : 0;
    }
    const parsed = parseFloat(String(horasRaw || '').replace(',', '.'));
    return Number.isFinite(parsed) ? parsed : 0;
  })();
  if (horasNumero <= 0) {
    horasNumero = 1;
  }
  const correo = String(getValue(['Correo', 'Email', 'Mail']) || '').trim().toLowerCase();
  const tipo = (() => {
    const raw = String(getValue(['Tipo', 'Tipo petición', 'Tipo Peticion', 'Tipo solicitud']) || defaultTipo || '').trim();
    if (!raw) return defaultTipo || '';
    return raw.charAt(0).toUpperCase() + raw.slice(1).toLowerCase();
  })();

  if (!fechaInicio) return null;

  return {
    campania: campania || sala,
    sala,
    centro,
    ciudad,
    franja,
    horario,
    horas,
    horasNumero,
    tipo,
    correo,
    fechaInicio,
    fechaFin: fechaFin || null
  };
}

function _parseToDate(value) {
  if (!value && value !== 0) return null;
  if (value instanceof Date) {
    return _normalizeDate(value);
  }
  if (typeof value === 'number') {
    const millis = Math.round((value - 25569) * 86400000);
    return _normalizeDate(new Date(millis));
  }
  const str = String(value).trim();
  if (!str) return null;

  const match = str.match(/^(\d{1,4})[\/-](\d{1,2})[\/-](\d{1,4})/);
  if (!match) return null;

  let first = Number(match[1]);
  let second = Number(match[2]);
  let third = Number(match[3]);

  if (first > 1900) {
    return _normalizeDate(new Date(first, second - 1, third));
  }

  if (third > 1900) {
    return _normalizeDate(new Date(third, second - 1, first));
  }

  if (third < 100) {
    third += 2000;
  }
  return _normalizeDate(new Date(third, second - 1, first));
}

function _normalizeDate(date) {
  const normalized = new Date(date);
  normalized.setHours(0, 0, 0, 0);
  return normalized;
}

function _formatMonthLabel(date) {
  const formatter = new Intl.DateTimeFormat('es-ES', { month: 'short' });
  const label = formatter.format(date);
  return label.charAt(0).toUpperCase() + label.slice(1);
}

function cargarHorasTrabajarDesdeExcel(payload) {
  return _cargarHorasDesdeExcel(payload, {
    sheet: sheetHoras,
    sheetName: SHEET_HORAS,
    requireAdmin: true
  });
}

function cargarHorasLibrarDesdeExcel(payload) {
  return _cargarHorasDesdeExcel(payload, {
    sheet: sheetHorasLibrar,
    sheetName: SHEET_HORAS_LIBRAR,
    requireAdmin: true
  });
}

function _cargarHorasDesdeExcel(payload, { sheet, sheetName, requireAdmin = true }) {
  const context = getUserContext();
  if (requireAdmin && (!context || !context.isAdmin)) {
    throw new Error('Solo los administradores pueden cargar horas.');
  }

  if (!payload || !payload.base64) {
    throw new Error('No se ha recibido ningún archivo para procesar.');
  }

  const fileName = payload.fileName || 'import.xlsx';
  const mimeType = _resolveImportMimeType(fileName, payload.mimeType);
  const blob = Utilities.newBlob(Utilities.base64Decode(payload.base64), mimeType, fileName);

  let tempFile = null;
  let convertedFileId = null;

  try {
    tempFile = DriveApp.createFile(blob);
    if (!tempFile) {
      throw new Error('No se pudo crear el archivo temporal en Drive.');
    }

    const copy = Drive.Files.copy(
      { title: `tmp-horas-${Date.now()}`, mimeType: MimeType.GOOGLE_SHEETS },
      tempFile.getId(),
      { convert: true }
    );

    if (!copy || !copy.id) {
      throw new Error('No se pudo convertir el archivo a hoja de cálculo.');
    }

    convertedFileId = copy.id;
    const tempSpreadsheet = SpreadsheetApp.openById(convertedFileId);
    const sourceSheet = tempSpreadsheet.getSheets()[0];
    if (!sourceSheet) {
      throw new Error('El archivo no contiene hojas para procesar.');
    }

    const data = sourceSheet.getDataRange().getValues();
    if (!data || data.length < 2) {
      return { importedRows: 0, message: 'El archivo no contiene registros para importar.' };
    }

    const targetSheet = sheet;
    if (!targetSheet) {
      throw new Error(`No se ha encontrado la hoja '${sheetName || 'destino'}'.`);
    }

    const targetHeader = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
    const sourceHeader = data[0];
    const mapping = _buildHorasSheetHeaderMapping(sourceHeader, targetHeader);
    if (!mapping.targetLength) {
      throw new Error('No se pudo determinar la estructura de la hoja destino.');
    }

    const rowsToInsert = [];
    for (let i = 1; i < data.length; i++) {
      const sourceRow = data[i];
      if (_isRowCompletelyEmpty(sourceRow)) continue;
      const transformed = _transformHorasSheetRow(sourceRow, mapping);
      if (transformed) {
        rowsToInsert.push(transformed);
      }
    }

    if (!rowsToInsert.length) {
      return { importedRows: 0, message: 'No se encontraron registros válidos en el archivo.' };
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const startRow = getFirstFreeRow(targetSheet);
      const requiredRows = startRow + rowsToInsert.length - 1;
      if (requiredRows > targetSheet.getMaxRows()) {
        const rowsToAdd = requiredRows - targetSheet.getMaxRows();
        targetSheet.insertRowsAfter(targetSheet.getMaxRows(), rowsToAdd);
      }
      targetSheet
        .getRange(startRow, 1, rowsToInsert.length, mapping.targetLength)
        .setValues(rowsToInsert);
    } finally {
      lock.releaseLock();
    }

    return {
      importedRows: rowsToInsert.length,
      message: `Archivo procesado correctamente. Se importaron ${rowsToInsert.length} registros.`
    };
  } catch (error) {
    const message = error && error.message ? error.message : 'Error desconocido al importar las horas.';
    throw new Error(message);
  } finally {
    if (tempFile) {
      try {
        tempFile.setTrashed(true);
      } catch (err) {}
    }
    if (convertedFileId) {
      try {
        DriveApp.getFileById(convertedFileId).setTrashed(true);
      } catch (err) {}
    }
  }
}

function _buildHorasSheetHeaderMapping(sourceHeader, targetHeader) {
  const sourceMap = {};
  const mapping = [];

  (sourceHeader || []).forEach((value, index) => {
    const key = _normalizeHeaderName(value);
    if (key && !(key in sourceMap)) {
      sourceMap[key] = index;
    }
  });

  (targetHeader || []).forEach((value, index) => {
    const key = _normalizeHeaderName(value);
    mapping.push({
      targetIndex: index,
      targetName: key,
      sourceIndex: Object.prototype.hasOwnProperty.call(sourceMap, key)
        ? sourceMap[key]
        : null
    });
  });

  const horasSourceIndex = _getFirstMatchingIndex(sourceMap, [
    'horas',
    'totalhoras',
    'horastrabajar',
    'cantidadhoras'
  ]);

  return {
    mapping,
    sourceMap,
    horasSourceIndex,
    targetLength: (targetHeader && targetHeader.length) || 0
  };
}

function _transformHorasSheetRow(sourceRow, mappingInfo) {
  if (!mappingInfo || !mappingInfo.mapping) return null;
  const { mapping, horasSourceIndex, targetLength } = mappingInfo;
  const result = new Array(targetLength).fill('');
  let hasValue = false;

  mapping.forEach(entry => {
    const { targetIndex, targetName, sourceIndex } = entry;
    if (targetIndex == null || targetIndex < 0 || targetIndex >= targetLength) return;

    let value = sourceIndex != null ? sourceRow[sourceIndex] : '';

    if (
      (targetName === 'disponible' ||
        targetName === 'disponibles' ||
        targetName === 'disponibilidad' ||
        targetName === 'dispon') &&
      sourceIndex == null &&
      horasSourceIndex != null
    ) {
      value = sourceRow[horasSourceIndex];
    }

    switch (targetName) {
      case 'campana':
        value = value != null ? String(value).trim() : '';
        break;
      case 'fecha':
        value = _parseImportDate(value);
        break;
      case 'franja':
        value = _parseFranjaValue(value);
        break;
      case 'horas':
      case 'totalhoras':
      case 'horastrabajar':
      case 'cantidadhoras':
        value = _parseNumericValue(value);
        break;
      case 'disponible':
      case 'disponibles':
      case 'disponibilidad':
      case 'dispon':
        value = _parseNumericValue(value);
        break;
      default:
        if (value instanceof Date) {
          value = _normalizeDate(value);
        } else if (typeof value === 'string') {
          value = value.trim();
        } else if (value == null) {
          value = '';
        }
        break;
    }

    if (value !== '' && value != null) {
      hasValue = true;
    }

    result[targetIndex] = value == null ? '' : value;
  });

  return hasValue ? result : null;
}

function _parseImportDate(value) {
  if (value instanceof Date) {
    return _normalizeDate(value);
  }

  if (typeof value === 'number' && Number.isFinite(value)) {
    const millis = Math.round((value - 25569) * 86400000);
    return _normalizeDate(new Date(millis));
  }

  const str = String(value || '').trim();
  if (!str) return '';

  const iso = str.match(/^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})$/);
  if (iso) {
    const year = Number(iso[1]);
    const month = Number(iso[2]);
    const day = Number(iso[3]);
    return _normalizeDate(new Date(year, month - 1, day));
  }

  const euro = str.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{2,4})$/);
  if (euro) {
    const day = Number(euro[1]);
    const month = Number(euro[2]);
    let year = Number(euro[3]);
    if (year < 100) {
      year += year >= 50 ? 1900 : 2000;
    }
    return _normalizeDate(new Date(year, month - 1, day));
  }

  const parsed = Date.parse(str);
  if (!isNaN(parsed)) {
    return _normalizeDate(new Date(parsed));
  }

  return '';
}

function _parseNumericValue(value) {
  if (value == null || value === '') return '';
  if (typeof value === 'number' && Number.isFinite(value)) {
    return value;
  }
  const str = String(value).trim().replace(',', '.');
  if (!str) return '';
  const num = parseFloat(str);
  return Number.isFinite(num) ? num : '';
}

function _parseFranjaValue(value) {
  if (value == null || value === '') return '';
  const str = String(value).trim();
  const normalized = _normalizarFranja(str);
  return normalized || str;
}

function _resolveImportMimeType(fileName, providedMime) {
  if (providedMime) return providedMime;
  const extension = String(fileName || '')
    .split('.')
    .pop()
    .toLowerCase();
  switch (extension) {
    case 'csv':
      return 'text/csv';
    case 'xls':
      return MimeType.MICROSOFT_EXCEL;
    case 'xlsx':
      return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    default:
      return 'application/octet-stream';
  }
}

function _normalizeHeaderName(value) {
  if (value == null) return '';
  return String(value)
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]+/g, '');
}

function _getFirstMatchingIndex(map, keys) {
  for (let i = 0; i < keys.length; i++) {
    const key = keys[i];
    if (Object.prototype.hasOwnProperty.call(map, key)) {
      return map[key];
    }
  }
  return null;
}

function _isRowCompletelyEmpty(row) {
  if (!row || !row.length) return true;
  return row.every(cell => cell === '' || cell == null || (typeof cell === 'string' && cell.trim() === ''));
}

function _obtenerNombreUsuario() {
  try {
    const email = (Session.getActiveUser().getEmail() || '').trim();
    if (!email) return '';
    const localPart = email.split('@')[0] || '';
    if (!localPart) return '';
    return localPart
      .split(/[._\s-]+/)
      .filter(Boolean)
      .map(fragment => fragment.charAt(0).toUpperCase() + fragment.slice(1).toLowerCase())
      .join(' ');
  } catch (err) {
    return '';
  }
}
