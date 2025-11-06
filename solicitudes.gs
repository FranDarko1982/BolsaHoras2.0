/**
 * solicitudes.gs
 * Operaciones sobre solicitudes existentes: cancelación, actualización y lecturas.
 */

function cancelarReserva(key) {
  const context = getUserContext();
  ensureAuthorizedContext(context);

  let ok = _cancelarReservaEnHoja(sheetResTrabajar, key);
  if (!ok) ok = _cancelarReservaEnHoja(sheetResCobrar, key);
  if (!ok) ok = _cancelarReservaEnHoja(sheetResLibrar, key);
  return ok;
}

function _cancelarReservaEnHoja(sheetResObj, key) {
  const sh = sheetResObj;
  if (!sh) return false;

  const context = getUserContext();
  ensureAuthorizedContext(context);

  const finder = sh.createTextFinder(key).matchEntireCell(true).findNext();
  if (!finder) return false;

  const row = finder.getRow();
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] || [];
  const headerMap = headers.reduce((map, header, index) => {
    const keyHeader = String(header || '').trim();
    if (keyHeader) map[keyHeader] = index;
    return map;
  }, {});

  const colCampania = _findFirstIndex(headerMap, ['Campaña', 'Campana']);
  const colCorreo = _findFirstIndex(headerMap, ['Correo', 'Email', 'Mail']);

  const campania = colCampania >= 0
    ? String(sh.getRange(row, colCampania + 1).getValue() || '').trim()
    : '';
  const emailUsuario = colCorreo >= 0
    ? sh.getRange(row, colCorreo + 1).getValue()
    : sh.getRange(row, 6).getValue();

  assertReservaAccess(context, campania, emailUsuario);

  sh.deleteRow(row);

  if (emailUsuario && String(emailUsuario).indexOf('@') > -1) {
    const hoy        = new Date();
    const tz         = getAppTimeZone();
    const fechaLarga = `Madrid, a ${Utilities.formatDate(hoy, tz, 'd')} de ${getNombreMes(hoy.getMonth())} de ${hoy.getFullYear()}`;
    const asunto     = 'Solicitud cancelada – Bolsa de horas';

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

function cancelarMultiplesReservas(keys) {
  if (!Array.isArray(keys) || keys.length === 0) {
    throw new Error('No se proporcionaron identificadores de reserva para cancelar.');
  }

  const context = getUserContext();
  ensureAuthorizedContext(context);

  keys.forEach(key => {
    _cancelarReservaEnHoja(sheetResTrabajar, key);
    _cancelarReservaEnHoja(sheetResCobrar, key);
    _cancelarReservaEnHoja(sheetResLibrar, key);
  });

  return true;
}

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

  const updatedCobrar = _actualizarSolicitudEnHoja(sheetResCobrar, data);
  if (updatedCobrar) {
    return {
      message: 'Solicitud actualizada correctamente.',
      key: updatedCobrar.key,
      sheet: updatedCobrar.sheet
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

  const tz = getAppTimeZone();
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

function tramitarSolicitudCobrar(options) {
  const context = getUserContext();
  ensureAuthorizedContext(context);

  if (!context.isAdmin && !context.isGestor) {
    throw new Error('No tienes permisos para tramitar solicitudes.');
  }

  if (!sheetResCobrar) {
    throw new Error('No se encontró la hoja de reservas de horas complementarias.');
  }

  const accionRaw = options && typeof options.accion === 'string'
    ? options.accion.trim().toLowerCase()
    : '';
  if (accionRaw !== 'aceptar' && accionRaw !== 'denegar') {
    throw new Error('Acción de tramitación no válida.');
  }

  const key = options && typeof options.key === 'string'
    ? options.key.trim()
    : '';
  const reservaId = options && typeof options.reservaId === 'string'
    ? options.reservaId.trim()
    : '';

  if (!key && !reservaId) {
    throw new Error('Falta el identificador de la solicitud.');
  }

  const dataRange = sheetResCobrar.getDataRange();
  const values = dataRange.getValues();
  if (!values.length) {
    throw new Error('No hay datos disponibles en la hoja de reservas complementarias.');
  }

  const headerRow = values[0];
  const headerMap = headerRow.reduce((map, header, index) => {
    const name = String(header || '').trim();
    if (name) map[name] = index;
    return map;
  }, {});

  const colKey = _findFirstIndex(headerMap, ['key', 'Key', 'KEY']);
  const colReservaId = _findFirstIndex(headerMap, [
    'ID reserva',
    'ID Reserva',
    'Id reserva',
    'Reserva ID',
    'ID'
  ]);
  const colEstado = _findFirstIndex(headerMap, [
    'Estado solic',
    'Estado solicitud',
    'Estado Solicitud',
    'Estado'
  ]);
  const colTipo = _findFirstIndex(headerMap, [
    'Tipo',
    'Tipo petición',
    'Tipo Peticion',
    'Tipo solicitud'
  ]);
  const colValidacion = _findFirstIndex(headerMap, [
    'Validación',
    'Validacion',
    'Validación solicitud',
    'Validacion solicitud'
  ]);

  const bodyRange = sheetResCobrar.getRange(2, 1, sheetResCobrar.getLastRow() - 1, sheetResCobrar.getLastColumn());
  const bodyValues = bodyRange.getValues();
  const estadosRange = colEstado >= 0
    ? sheetResCobrar.getRange(2, colEstado + 1, bodyValues.length, 1)
    : null;
  const estadosValues = estadosRange ? estadosRange.getValues() : null;

  let targetRowIndex = -1;
  for (let i = 0; i < bodyValues.length; i++) {
    const row = bodyValues[i];
    const keyValue = colKey >= 0 ? String(row[colKey] || '').trim() : '';
    const reservaIdValue = colReservaId >= 0 ? String(row[colReservaId] || '').trim() : '';

    if ((key && keyValue === key) || (reservaId && reservaIdValue === reservaId)) {
      targetRowIndex = i;
      break;
    }
  }

  if (targetRowIndex === -1) {
    throw new Error('No se encontró la solicitud indicada.');
  }

  const rowValues = bodyValues[targetRowIndex];
  const estadoFinal = accionRaw === 'aceptar' ? 'ACEPTADA' : 'DENEGADA';
  const esAceptacion = accionRaw === 'aceptar';

  if (esAceptacion) {
    if (colValidacion >= 0) {
      sheetResCobrar
        .getRange(targetRowIndex + 2, colValidacion + 1)
        .setValue('OK');
      rowValues[colValidacion] = 'OK';
    }
    if (colTipo >= 0) {
      sheetResCobrar
        .getRange(targetRowIndex + 2, colTipo + 1)
        .setValue('Complementaria aceptada');
      rowValues[colTipo] = 'Complementaria aceptada';
    }
  } else if (estadosRange) {
    estadosRange.getCell(targetRowIndex + 1, 1).setValue(estadoFinal);
    if (colEstado >= 0) {
      rowValues[colEstado] = estadoFinal;
    }
  }

  const correoIndex = _findFirstIndex(headerMap, ['Correo', 'Email', 'Mail']);
  const campaniaIndex = _findFirstIndex(headerMap, ['Campaña', 'Campana']);
  const fechaIndex = _findFirstIndex(headerMap, ['Fecha', 'Fecha reserva', 'Fecha inicio', 'Fecha Inicio']);
  const franjaIndex = _findFirstIndex(headerMap, ['FRANJA', 'Franja', 'Franja horaria', 'Horario', 'Horario reserva']);
  const horasIndex = _findFirstIndex(headerMap, ['HORAS', 'Horas', 'Horas solicitadas', 'Horas reservadas']);
  const tipoIndex = _findFirstIndex(headerMap, ['Tipo', 'Tipo petición', 'Tipo Peticion', 'Tipo solicitud']);

  const correoDestino = correoIndex >= 0 ? String(rowValues[correoIndex] || '').trim() : '';
  if (correoDestino && correoDestino.indexOf('@') > -1) {
    const tz = getAppTimeZone();
    const campania = campaniaIndex >= 0 ? String(rowValues[campaniaIndex] || '').trim() : '';
    const fechaRaw = fechaIndex >= 0 ? rowValues[fechaIndex] : '';
    const fechaDisplay = fechaRaw instanceof Date
      ? Utilities.formatDate(fechaRaw, tz, 'dd/MM/yyyy')
      : String(fechaRaw || '').trim();
    const franja = franjaIndex >= 0 ? String(rowValues[franjaIndex] || '').trim() : '';
    const horas = horasIndex >= 0 ? rowValues[horasIndex] : '';
    const tipo = tipoIndex >= 0 ? String(rowValues[tipoIndex] || '').trim() : '';
    const tipoAsunto = (typeof normalizeTipoReserva === 'function')
      ? normalizeTipoReserva(tipo || 'Complementaria')
      : (tipo || 'Complementaria');

    const asunto = `Solicitud ${estadoFinal.toLowerCase()} – Bolsa de horas (${tipoAsunto})`;
    const body = `
      <div style="font-family: Calibri, Arial, sans-serif; max-width:650px; margin:0 auto; background:#fff;">
        <div style="width:100%; text-align:center; margin:24px 0 32px;">
          <img src="${URL_LOGO_INTELCIA}" alt="Intelcia" style="width:100%; max-width:500px; border-radius:8px;">
        </div>
        <div style="color:#222; font-size:16px; margin-bottom:24px;">
          Tu solicitud ha sido <strong>${estadoFinal.toLowerCase()}</strong>.
        </div>
        <ul style="color:#C9006C; font-size:16px; margin-left:32px;">
          <li><b>Campaña:</b> ${campania}</li>
          <li><b>Fecha:</b> ${fechaDisplay}</li>
          <li><b>Horario:</b> ${franja}</li>
          <li><b>Horas:</b> ${horas}</li>
        </ul>
        <div style="margin-top:24px; font-size:15px;">Gracias</div>
      </div>
    `;

    MailApp.sendEmail({ to: correoDestino, subject: asunto, htmlBody: body });
  }

  const response = {
    message: `Solicitud ${estadoFinal.toLowerCase()} correctamente.`,
    estadoSolicitud: estadoFinal
  };

  if (colKey >= 0) {
    response.key = String(rowValues[colKey] || '').trim();
  } else if (key) {
    response.key = key;
  }

  if (colReservaId >= 0) {
    response.reservaId = String(rowValues[colReservaId] || '').trim();
  } else if (reservaId) {
    response.reservaId = reservaId;
  }

  if (colEstado >= 0) {
    response.estadoSolicitud = String(rowValues[colEstado] || '').trim();
  }

  return response;
}

function getMisReservas() {
  const context = getUserContext();
  ensureAuthorizedContext(context);

  const reservasTrabajar = _getMisReservasDeHoja(sheetResTrabajar, context, 'Trabajar');
  const reservasCobrar   = _getMisReservasDeHoja(sheetResCobrar,   context, 'Complementaria');
  const reservasLibrar   = _getMisReservasDeHoja(sheetResLibrar,   context, 'Librar');
  return [...reservasTrabajar, ...reservasCobrar, ...reservasLibrar].sort((a, b) => {
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
  const tz = getAppTimeZone();
  const context = getUserContext();
  ensureAuthorizedContext(context);

  const todasSolicitudes = (() => {
    if (sheetDatosLocker) {
      const lockerSolicitudes = _getSolicitudesDeHoja(sheetDatosLocker, '', tz);
      if (Array.isArray(lockerSolicitudes) && lockerSolicitudes.length) {
        return lockerSolicitudes;
      }
    }
    const solicitudesTrabajar = _getSolicitudesDeHoja(sheetResTrabajar, 'Trabajar', tz);
    const solicitudesCobrar = _getSolicitudesDeHoja(sheetResCobrar, 'Complementaria', tz);
    const solicitudesLibrar = _getSolicitudesDeHoja(sheetResLibrar, 'Librar', tz);
    return [...solicitudesTrabajar, ...solicitudesCobrar, ...solicitudesLibrar];
  })();

  const visibles = todasSolicitudes.filter(item => {
    if (!item || typeof item !== 'object') return false;
    const campaniaItem = item.campania || item.sala || '';
    return canAccessReserva(context, {
      campania: campaniaItem,
      correo: item.correo
    });
  });

  return visibles
    .sort((a, b) => b._timestamp - a._timestamp || a.tipo.localeCompare(b.tipo))
    .map(({ _timestamp, ...rest }) => rest);
}

function _getMisReservasDeHoja(hoja, context, tipoReserva) {
  if (!hoja) return [];

  const ctx = context || getUserContext();
  ensureAuthorizedContext(ctx);

  const tz = getAppTimeZone();
  const data = hoja.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0];
  const C_CAMP = headers.indexOf('Campaña');
  const C_FEC = headers.indexOf('Fecha');
  const C_HORAS = headers.indexOf('HORAS');
  const C_FRANJA = headers.indexOf('FRANJA');
  const C_CORREO = headers.indexOf('Correo');
  const C_TIPO = headers.indexOf('Tipo');
  const C_VALID = headers.indexOf('Validación');
  const C_CLAVE = headers.indexOf('Clave');
  const C_ESTADO = headers.indexOf('Estado solic');
  const C_EMPLEADO = headers.indexOf('NºEmpleado');
  const C_IDRES = headers.indexOf('ID reserva');
  const C_KEY = headers.indexOf('key');

  const reservas = [];

  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const correo = (r[C_CORREO] || '').toString().trim();
    const campania = (r[C_CAMP] || '').toString().trim();

    if (!canAccessReserva(ctx, { campania, correo })) continue;

    const fechaRaw = C_FEC >= 0 ? r[C_FEC] : '';
    const fechaFormateada =
      fechaRaw instanceof Date
        ? Utilities.formatDate(_normalizeDate(fechaRaw), tz, 'dd/MM/yyyy')
        : String(fechaRaw || '');

    const tipoSheet = C_TIPO >= 0 ? String(r[C_TIPO] || '').trim() : '';
    const tipoRaw = tipoSheet || tipoReserva || '';
    const tipoFinal = (typeof normalizeTipoReserva === 'function')
      ? normalizeTipoReserva(tipoRaw)
      : tipoRaw;

    reservas.push({
      id: r[C_IDRES] || '',
      campania,
      correo,
      empleado: r[C_EMPLEADO] || '',
      fecha: fechaFormateada,
      franja: r[C_FRANJA] || '',
      horas: r[C_HORAS] || 0,
      tipo: tipoFinal,
      estado: r[C_ESTADO] || '',
      validacion: r[C_VALID] || '',
      clave: r[C_CLAVE] || '',
      key: C_KEY >= 0 ? String(r[C_KEY] || '').trim() : (r[C_CLAVE] || '')
    });
  }

  return reservas;
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
      if (typeof normalizeTipoReserva === 'function') {
        return valor ? normalizeTipoReserva(valor) : normalizeTipoReserva(defaultTipo || '');
      }
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
