/**
 * logs.gs
 * Utilidades para registrar eventos de la aplicación en la hoja `logs`.
 */

/**
 * Registra un evento en la hoja de logs.
 * @param {string} tipoEvento Descripción del evento realizado.
 * @param {string} [correoActor] Correo del usuario que ejecuta la acción.
 */
function registrarLogEvento(tipoEvento, correoActor) {
  try {
    const sheet = sheetLogs || (ss && ss.getSheetByName && ss.getSheetByName(SHEET_LOGS));
    if (!sheet) return;

    const tz = getAppTimeZone();
    const now = new Date();
    const fecha = Utilities.formatDate(now, tz, 'dd/MM/yyyy');
    const hora = Utilities.formatDate(now, tz, 'HH:mm:ss');

    let correo = correoActor;
    if (!correo) {
      correo = getCurrentUserEmail ? getCurrentUserEmail() : '';
    }
    if (correo) {
      correo = String(correo).trim().toLowerCase();
    } else {
      correo = '';
    }

    const evento = String(tipoEvento || '').trim() || 'Evento';
    sheet.appendRow([fecha, hora, correo, evento]);
  } catch (err) {
    console.error('No se pudo registrar el log de evento:', err);
  }
}
