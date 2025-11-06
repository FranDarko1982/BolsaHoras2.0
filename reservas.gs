/**
 * reservas.gs
 * Lógica de disponibilidad y creación de reservas (trabajar, librar, cobrar).
 */

function getFirstFreeRow(sh) {
  const values = sh.getRange(2, 1, sh.getMaxRows() - 1, 1).getValues().flat();
  const idx = values.findIndex(v => v === '' || v == null);
  return (idx >= 0) ? idx + 2 : sh.getLastRow() + 1;
}

function getHorasDisponibles(campania, startISO, endISO) {
  const context = getUserContext();
  ensureAuthorizedContext(context);
  const resolvedCampania = resolveCampaniaForContext(campania, context, { allowEmpty: true });
  return _getHorasDisponibles(sheetHoras, sheetResTrabajar, resolvedCampania, startISO, endISO, context);
}

function getHorasDisponiblesLibrar(campania, startISO, endISO) {
  const context = getUserContext();
  ensureAuthorizedContext(context);
  const resolvedCampania = resolveCampaniaForContext(campania, context, { allowEmpty: true });
  return _getHorasDisponibles(sheetHorasLibrar, sheetResLibrar, resolvedCampania, startISO, endISO, context);
}

function _getHorasDisponibles(sheetHorasObj, sheetResObj, campania, startISO, endISO, context) {
  const ctx = context || getUserContext();
  ensureAuthorizedContext(ctx);
  const normalizedCampania = normalizeCampaniaValue(campania);

  const start = new Date(startISO);
  const end   = new Date(endISO);

  const shBase = sheetHorasObj;
  const rows   = shBase.getDataRange().getValues().slice(1);
  const eventos = [];

  rows.forEach(r => {
    const rawCampania = String(r[COL.CAMPAÑA] || '').trim();
    const rowCampaniaNormalized = normalizeCampaniaValue(rawCampania);

    if (normalizedCampania) {
      if (rowCampaniaNormalized !== normalizedCampania) return;
    } else if (!ctx.isAdmin) {
      if (!isCampaniaAllowedForContext(ctx, rawCampania)) {
        return;
      }
    }

    const fechaStr    = r[COL.FECHA];
    const franja      = r[COL.FRANJA];
    const disponible  = +r[COL.DISPON] || 0;
    if (!fechaStr || !franja || disponible <= 0) return;

    const baseDate = _parseToDate(fechaStr);
    if (!baseDate) return;

    const [hIniStr] = String(franja).split('-');
    if (!hIniStr) return;
    const [h0Raw, m0Raw] = hIniStr.split(':');
    const h0 = Number(h0Raw);
    const m0 = Number(m0Raw);

    const inicio = new Date(baseDate);
    inicio.setHours(
      Number.isFinite(h0) ? h0 : 0,
      Number.isFinite(m0) ? m0 : 0,
      0,
      0
    );
    if (inicio < start || inicio >= end) return;

    const fin = new Date(inicio.getTime() + 3600000);

    eventos.push({
      start    : toLocalIso(inicio),
      end      : toLocalIso(fin),
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

function comprobarDisponibilidadContinua(sheetHorasObj, campania, startISO, horas) {
  const context = getUserContext();
  ensureAuthorizedContext(context);
  const resolvedCampania = resolveCampaniaForContext(campania, context, {
    allowEmpty: context.isAdmin
  });
  const normalizedCampania = normalizeCampaniaValue(resolvedCampania);

  const tz  = getAppTimeZone();
  const ini = new Date(startISO);
  const fechaStr = Utilities.formatDate(ini, tz, 'dd/MM/yyyy');

  const datos = sheetHorasObj.getDataRange().getValues();
  const disponibles = new Set();

  for (let i = 1; i < datos.length; i++) {
    const r = datos[i];

    if (normalizedCampania) {
      const rowCampaniaNormalized = normalizeCampaniaValue(r[COL.CAMPAÑA]);
      if (rowCampaniaNormalized !== normalizedCampania) continue;
    } else if (!context.isAdmin) {
      const rawCampania = String(r[COL.CAMPAÑA] || '').trim();
      if (!isCampaniaAllowedForContext(context, rawCampania)) continue;
    }

    const celdaFecha = r[COL.FECHA];
    const fechaCelda = (celdaFecha instanceof Date)
      ? Utilities.formatDate(celdaFecha, tz, 'dd/MM/yyyy')
      : String(celdaFecha).trim();

    if (fechaCelda !== fechaStr) continue;

    const disp = +r[COL.DISPON] || 0;
    if (disp <= 0) continue;

    const franjaNorm = _normalizarFranja(r[COL.FRANJA]);
    if (franjaNorm) disponibles.add(franjaNorm);
  }

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

function reservarVariasHoras(campania, startISO, horas) {
  const context = getUserContext();
  ensureAuthorizedContext(context);
  let startDate;
  if (typeof startISO === 'number') {
    startDate = new Date(startISO);
  } else if (typeof startISO === 'string') {
    startDate = new Date(startISO);
  } else {
    throw new Error('Parámetro startISO inválido');
  }
  if (Number.isNaN(startDate.getTime())) {
    throw new Error('Parámetro startISO inválido');
  }
  const resolvedCampania = resolveCampaniaForContext(campania, context);
  return _reservarVariasHoras('Trabajar', sheetResTrabajar, resolvedCampania, startDate, horas, context);
}

function reservarVariasHorasCobrar(campania, startISO, horas) {
  const context = getUserContext();
  ensureAuthorizedContext(context);
  let startDate;
  if (typeof startISO === 'number') {
    startDate = new Date(startISO);
  } else if (typeof startISO === 'string') {
    startDate = new Date(startISO);
  } else {
    throw new Error('Parámetro startISO inválido');
  }
  if (Number.isNaN(startDate.getTime())) {
    throw new Error('Parámetro startISO inválido');
  }
  const resolvedCampania = resolveCampaniaForContext(campania, context);
  return _reservarVariasHoras('Complementaria', sheetResCobrar, resolvedCampania, startDate, horas, context);
}

function reservarHoraLibrar(campania, startISO) {
  const context = getUserContext();
  ensureAuthorizedContext(context);
  const resolvedCampania = resolveCampaniaForContext(campania, context);
  return _reservarHora('Librar', sheetResLibrar, resolvedCampania, startISO, context);
}

function reservarVariasHorasLibrar(campania, startISO, horas) {
  const context = getUserContext();
  ensureAuthorizedContext(context);
  let startDate;
  if (typeof startISO === 'number') {
    startDate = new Date(startISO);
  } else if (typeof startISO === 'string') {
    startDate = new Date(startISO);
  } else {
    throw new Error('Parámetro startISO inválido');
  }
  if (Number.isNaN(startDate.getTime())) {
    throw new Error('Parámetro startISO inválido');
  }
  const resolvedCampania = resolveCampaniaForContext(campania, context);
  return _reservarVariasHoras('Librar', sheetResLibrar, resolvedCampania, startDate, horas, context);
}

function _reservarHora(tipo, sheetResObj, campania, startISO, context) {
  const ctx = context || getUserContext();
  ensureAuthorizedContext(ctx);
  const resolvedCampania = resolveCampaniaForContext(campania, ctx);

  const tz   = getAppTimeZone();
  const ini  = new Date(startISO);
  const fin  = new Date(ini.getTime() + 3600000);
  const franja = Utilities.formatDate(ini, tz, 'HH:00') + '-' +
                 Utilities.formatDate(fin, tz, 'HH:00');
  const fechaSerial = Math.floor(
    (Date.UTC(ini.getFullYear(), ini.getMonth(), ini.getDate()) -
     Date.UTC(1899, 11, 30)) / 86400000
  );
  const email = ctx.email || Session.getActiveUser().getEmail() || '';
  const fechaSoloTexto = Utilities.formatDate(ini, tz, 'dd/MM/yyyy');
  const tipoUpper = String(tipo || '').toUpperCase();
  const esComplementaria = tipoUpper === 'COMPLEMENTARIA' || tipoUpper === 'COBRAR';
  const tipoFinal = esComplementaria ? 'Complementaria' : tipo;
  const tipoKey = esComplementaria ? 'Trabajar' : tipo;
  const keyReserva = resolvedCampania + fechaSerial + franja + email + tipoKey;
  const reservaId = generateReservaId();

  let sh = sheetResObj;
  ensureReservaIdColumn(sh);

  const targetRow = getFirstFreeRow(sh);
  sh.getRange(targetRow, 1, 1, 7).setValues([[
    resolvedCampania,
    fechaSoloTexto,
    1,
    franja,
    keyReserva,
    email,
    tipoFinal
  ]]);
  sh.getRange(targetRow, 11).setValue(reservaId);

  sendConfirmationEmail(tipoFinal, resolvedCampania, ini, franja.split('-')[0], franja.split('-')[1], 1, email, reservaId);

  registrarLogEvento(
    `Reserva creada [${tipoFinal}] - ${resolvedCampania} - ${reservaId}`,
    ctx.email || email
  );

  if (esComplementaria) {
    return [
      'Tu solicitud se ha enviado correctamente.',
      `ID de la solicitud: ${reservaId}`,
      'Debe esperar la confirmación por parte de su campaña antes de que la solicitud sea efectiva.'
    ].join('\n');
  }

  return `Solicitud registrada para ${campania} ${franja} (${fechaSoloTexto}) [${tipoFinal}]. ID de reserva: ${reservaId}`;
}

function _reservarVariasHoras(tipo, sheetResObj, campania, startDate, horas, context) {
  const ctx = context || getUserContext();
  ensureAuthorizedContext(ctx);
  const resolvedCampania = resolveCampaniaForContext(campania, ctx);
  const tipoUpper = String(tipo || '').toUpperCase();
  const esComplementaria = tipoUpper === 'COMPLEMENTARIA' || tipoUpper === 'COBRAR';
  const tipoFinal = esComplementaria ? 'Complementaria' : tipo;

  const tz    = getAppTimeZone();
  const ini   = new Date(startDate);
  if (Number.isNaN(ini.getTime())) {
    throw new Error('Fecha de inicio inválida');
  }
  const email = ctx.email || Session.getActiveUser().getEmail() || '';
  const sh    = sheetResObj;
  ensureReservaIdColumn(sh);
  const reservaId = generateReservaId();

  let sheetHorasObj = (tipo === 'Librar') ? sheetHorasLibrar : sheetHoras;

  if (esComplementaria && !puedeCobrarCampania(resolvedCampania)) {
    throw new Error(`La campaña "${resolvedCampania}" no tiene habilitada la opción de horas COMPLEMENTARIAS.`);
  }

  if (!comprobarDisponibilidadContinua(sheetHorasObj, resolvedCampania, ini, horas)) {
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
    const tipoKey     = esComplementaria ? 'Trabajar' : tipo;
    const keyReserva  = `${resolvedCampania}${fechaSerial}${strFranja}${email}${tipoKey}`;

    rows.push([
      resolvedCampania,
      fechaSolo,
      1,
      strFranja,
      keyReserva,
      email,
      tipoFinal
    ]);

    if (h === 0)         primerFranja = strFranja.split('-')[0];
    if (h === horas - 1) ultimaFranja = strFranja.split('-')[1];
    fechaTexto = fechaSolo;
  }

  sh.getRange(firstRow, 1, rows.length, 7).setValues(rows);
  const reservaIdValues = rows.map(() => [reservaId]);
  sh.getRange(firstRow, 11, reservaIdValues.length, 1).setValues(reservaIdValues);

  sendConfirmationEmail(tipoFinal, resolvedCampania, ini, primerFranja, ultimaFranja, horas, email, reservaId);

  registrarLogEvento(
    `Reserva creada [${tipoFinal}] - ${resolvedCampania} - ${reservaId}`,
    ctx.email || email
  );

  if (esComplementaria) {
    return [
      'Tu solicitud se ha enviado correctamente.',
      `ID de la solicitud: ${reservaId}`,
      'Debe esperar la confirmación por parte de su campaña antes de que la solicitud sea efectiva.'
    ].join('\n');
  }

  return `Solicitud registrada para ${resolvedCampania} ${fechaTexto} de ${primerFranja} a ${ultimaFranja} (${horas}h) [${tipoFinal}]. ID de reserva: ${reservaId}. Se ha enviado un email de confirmación.`;
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
    getHighestReservaIdFromSheet(sheetResCobrar),
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

  const tz         = getAppTimeZone();
  const tipoUpper  = String(tipo || '').toUpperCase();
  const esComplementaria = tipoUpper === 'COMPLEMENTARIA' || tipoUpper === 'COBRAR';
  const tipoDisplay = esComplementaria ? 'Complementaria' : (tipo || '');
  const asunto     = `Solicitud registrada – Bolsa de horas (${tipoDisplay})`;
  const fechaLarga = `Madrid, a ${Utilities.formatDate(ini, tz, 'd')} de ${getNombreMes(ini.getMonth())} de ${ini.getFullYear()}`;
  const horario    = `${primer} - ${ultima}`;
  const idHtml     = reservaId ? `<li><b>ID de reserva:</b> ${reservaId}</li>` : '';

  const introduccionHtml = esComplementaria
    ? `
      <div style="margin-bottom:16px;">
        Desde <span style="color:#C9006C; font-weight:bold;">Intelcia Spanish Region</span> te informamos que tu solicitud se ha registrado correctamente.
      </div>
      <div style="margin-bottom:16px;">
        <b>ID de la solicitud:</b> ${reservaId}.
      </div>
      <div style="margin-bottom:16px;">
        Debe esperar la confirmación por parte de su campaña antes de que la solicitud sea efectiva.
      </div>
      <div style="margin-bottom:16px;">
        Te recordamos la petición que has realizado:
      </div>`
    : `
      <div style="margin-bottom:16px;">
        Desde <span style="color:#C9006C; font-weight:bold;">Intelcia Spanish Region</span> te informamos que tu solicitud ha sido aceptada. Te recordamos la petición que has realizado:
      </div>`;

  const notaFinalHtml = esComplementaria
    ? `
      <div style="font-size:15px; margin:14px 0 16px; color:#222;">
        Puedes cancelar esta solicitud hasta un día antes de la fecha que solicitaste trabajar o librar. Te recordamos que debes esperar la confirmación por parte de tu campaña antes de que la solicitud sea efectiva.
        Además, puedes realizar una nueva petición desde la aplicación 
        <span style="color:#C9006C; font-weight:bold;">Bolsa de horas</span> 
        <a href="${URL_APP}" style="color:#C9006C; text-decoration:underline;">(enlace a la aplicación)</a>.
      </div>`
    : `
      <div style="font-size:15px; margin:14px 0 16px; color:#222;">
        Puedes cancelar esta solicitud hasta un día antes de la fecha que solicitaste trabajar o librar. Te recordamos que puedes realizar una nueva petición desde la aplicación 
        <span style="color:#C9006C; font-weight:bold;">Bolsa de horas</span> 
        <a href="${URL_APP}" style="color:#C9006C; text-decoration:underline;">(enlace a la aplicación)</a>.
      </div>`;

  const body = `
    <div style="font-family: Calibri, Arial, sans-serif; max-width:650px; margin:0 auto; background:#fff;">
      <div style="width:100%; text-align:center; margin:24px 0 32px;">
        <img src="${URL_LOGO_INTELCIA}" alt="Intelcia" style="width:100%; max-width:500px; border-radius:8px;">
      </div>
      <div style="color:#222; font-size:16px; margin-bottom:24px;">
        En ${fechaLarga}
      </div>
      ${introduccionHtml}
      <ul style="color:#C9006C; font-size:16px; margin-left:32px;">
        ${idHtml}
        <li><b>Campaña:</b> ${campania}</li>
        <li><b>Fecha:</b> ${Utilities.formatDate(ini, tz, 'dd/MM/yyyy')}</li>
        <li><b>Horario:</b> ${horario}</li>
        <li><b>Horas:</b> ${horas}</li>
      </ul>
      ${notaFinalHtml}
      <div style="margin-top:24px; font-size:15px;">Gracias</div>
    </div>
  `;

  MailApp.sendEmail({ to: email, subject: asunto, htmlBody: body });
}
