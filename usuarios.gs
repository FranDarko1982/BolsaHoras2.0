/**
 * usuarios.gs
 * Gestión de usuarios, roles y contexto de acceso para la aplicación.
 * Depende de config.gs (constantes y hojas) y utilidades.gs (helpers de texto).
 */

let empleadosHeaderInfoCache = null;

function getEmpleadosHeaderInfo() {
  if (!sheetEmpleados) return null;
  if (empleadosHeaderInfoCache) return empleadosHeaderInfoCache;

  const lastColumn = sheetEmpleados.getLastColumn();
  if (!lastColumn) return null;

  const headers = sheetEmpleados.getRange(1, 1, 1, lastColumn).getValues()[0] || [];
  const headerMap = headers.reduce((acc, header, index) => {
    const key = _normalizeHeaderName(header);
    if (key) acc[key] = index;
    return acc;
  }, {});

  const info = {
    headers,
    totalColumns: lastColumn,
    colEmpleado: _getFirstMatchingIndex(headerMap, [
      'empleado',
      'numempleado',
      'numeroempleado',
      'nºempleado',
      'nempleado',
      'idempleado'
    ]),
    colCorreo: _getFirstMatchingIndex(headerMap, ['correo', 'email', 'mail']),
    colCampania: _getFirstMatchingIndex(headerMap, ['campana', 'campaña']),
    colRol: _getFirstMatchingIndex(headerMap, ['rol', 'role'])
  };

  empleadosHeaderInfoCache = info;
  return info;
}

function findEmpleadoRecordByEmail(email) {
  if (!sheetEmpleados) return null;
  const normalizedEmail = normalizeString(email);
  if (!normalizedEmail) return null;

  const info = getEmpleadosHeaderInfo();
  if (!info || info.colCorreo == null || info.colCorreo < 0) return null;

  const lastRow = sheetEmpleados.getLastRow();
  if (lastRow < 2) return null;

  const values = sheetEmpleados
    .getRange(2, 1, lastRow - 1, info.totalColumns)
    .getValues();

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const rawCorreo = row[info.colCorreo];
    const correo = normalizeString(rawCorreo);
    if (!correo) continue;
    if (correo === normalizedEmail) {
      return {
        empleado: info.colEmpleado >= 0 ? row[info.colEmpleado] : '',
        correo: rawCorreo,
        campania: info.colCampania >= 0 ? row[info.colCampania] : '',
        rol: info.colRol >= 0 ? row[info.colRol] : ''
      };
    }
  }

  return null;
}

function normalizeRoleValue(value) {
  const normalized = normalizeString(value);
  if (!normalized) return '';
  if (
    normalized === 'admin' ||
    normalized === 'administrador' ||
    normalized === 'administrator'
  ) {
    return ROLE.ADMIN;
  }
  if (
    normalized === 'gestor' ||
    normalized === 'manager' ||
    normalized === 'coordinador'
  ) {
    return ROLE.GESTOR;
  }
  return ROLE.USER;
}

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
    .map(row => normalizeString(row[0]))
    .filter(Boolean);
}

function userHasAdminRole(email) {
  if (!email) return false;
  const admins = getAdminEmails();
  return admins.includes(normalizeString(email));
}

function getUserContext() {
  const email = getCurrentUserEmail();
  const normalizedEmail = normalizeString(email);
  const adminByList = userHasAdminRole(email);

  const empleadoRecord = normalizedEmail ? findEmpleadoRecordByEmail(normalizedEmail) : null;
  const roleFromEmpleado = empleadoRecord ? normalizeRoleValue(empleadoRecord.rol) : '';
  const assignedCampaniasRaw = empleadoRecord ? parseCampaniaList(empleadoRecord.campania) : [];

  const uniqueCampaniasMap = new Map();
  assignedCampaniasRaw.forEach(campania => {
    const normalized = normalizeCampaniaValue(campania);
    if (!normalized) return;
    if (!uniqueCampaniasMap.has(normalized)) {
      uniqueCampaniasMap.set(normalized, campania);
    }
  });

  const assignedCampanias = Array.from(uniqueCampaniasMap.values());
  const normalizedCampanias = Array.from(uniqueCampaniasMap.keys());

  let role = ROLE.USER;
  if (adminByList || roleFromEmpleado === ROLE.ADMIN) {
    role = ROLE.ADMIN;
  } else if (roleFromEmpleado === ROLE.GESTOR) {
    role = ROLE.GESTOR;
  }

  const isAdmin = role === ROLE.ADMIN;
  const context = {
    email: normalizedEmail,
    isAdmin,
    role,
    isGestor: role === ROLE.GESTOR,
    campania: assignedCampanias[0] || '',
    campanias: assignedCampanias,
    campaniasNormalizadas: normalizedCampanias,
    empleadoId: empleadoRecord ? String(empleadoRecord.empleado || '').trim() : '',
    authorized: Boolean(normalizedEmail && (isAdmin || assignedCampanias.length > 0))
  };

  return context;
}

function ensureAuthorizedContext(context) {
  if (!context || (!context.authorized && !context.isAdmin)) {
    throw new Error('No tienes permisos para acceder a esta aplicación. Contacta con tu responsable.');
  }
}

function getContextAllowedCampanias(context) {
  if (!context || context.isAdmin) return [];
  if (Array.isArray(context.campanias) && context.campanias.length) {
    return context.campanias.slice();
  }
  if (context.campania) {
    return [context.campania];
  }
  return [];
}

function getContextAllowedCampaniasNormalized(context) {
  if (!context || context.isAdmin) return [];
  if (Array.isArray(context.campaniasNormalizadas) && context.campaniasNormalizadas.length) {
    return context.campaniasNormalizadas.slice();
  }
  const single = normalizeCampaniaValue(context.campania);
  return single ? [single] : [];
}

function resolveCampaniaForContext(campania, context, options) {
  const opts = options || {};
  const requestedNormalized = normalizeCampaniaValue(campania);
  if (!context || context.isAdmin) {
    return campania;
  }

  const allowed = getContextAllowedCampanias(context);
  const allowedNormalized = getContextAllowedCampaniasNormalized(context);

  if (!allowedNormalized.length) {
    if (opts.allowEmpty && !requestedNormalized) {
      return '';
    }
    throw new Error('No tienes ninguna campaña asignada. Contacta con tu responsable.');
  }

  if (!requestedNormalized) {
    return allowed[0];
  }

  const idx = allowedNormalized.indexOf(requestedNormalized);
  if (idx === -1) {
    throw new Error('No tienes permisos para operar sobre la campaña seleccionada.');
  }
  return allowed[idx] || allowed[0];
}

function isCampaniaAllowedForContext(context, campania) {
  if (!context) return false;
  if (context.isAdmin) return true;
  const normalized = normalizeCampaniaValue(campania);
  if (!normalized) return false;
  const allowedNormalized = getContextAllowedCampaniasNormalized(context);
  return allowedNormalized.includes(normalized);
}

function canAccessReserva(context, reserva) {
  if (!context || !reserva) return false;
  if (context.isAdmin) return true;
  const correoNormalizado = normalizeString(reserva.correo);
  const campaniaNormalizada = normalizeCampaniaValue(reserva.campania);
  const tieneCampania = Boolean(campaniaNormalizada);
  const campaniaPermitida = tieneCampania
    ? isCampaniaAllowedForContext(context, reserva.campania)
    : !context.isGestor;

  if (!campaniaPermitida) return false;

  if (context.isGestor) {
    return true;
  }

  return !!correoNormalizado && correoNormalizado === context.email;
}

function assertReservaAccess(context, campania, correo) {
  if (!canAccessReserva(context, { campania, correo })) {
    throw new Error('No tienes permisos para operar sobre esta solicitud.');
  }
}

function getAccessContext() {
  const context = getUserContext();
  let sections = ['inicio', 'calendario', 'mis-reservas', 'ayuda'];
  if (context.isAdmin) {
    sections = ['inicio', 'calendario', 'mis-reservas', 'reportes', 'ayuda', 'panel-control'];
  } else if (context.isGestor) {
    sections = ['inicio', 'calendario', 'mis-reservas', 'reportes', 'ayuda', 'panel-control'];
  } else {
    sections = ['inicio', 'calendario', 'mis-reservas', 'ayuda'];
  }

  return {
    role: context.role,
    email: context.email,
    isAdmin: context.isAdmin,
    isGestor: context.isGestor,
    campania: context.campania,
    campanias: context.campanias,
    authorized: context.authorized,
    sections
  };
}
