/**
 * dashboard.gs
 * Construcción del panel de inicio y métricas asociadas a las reservas.
 */

function getInicioDashboardData() {
  const tz = getAppTimeZone();
  const hoy = _normalizeDate(new Date());
  const context = getUserContext();
  ensureAuthorizedContext(context);
  const { email, isAdmin, role } = context;

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
    { sheet: sheetResCobrar, tipo: 'Cobrar' },
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
      const campaniaReserva = reserva.campania || reserva.sala || '';
      if (!canAccessReserva(context, { campania: campaniaReserva, correo: reserva.correo })) return;

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
    email,
    campania: context.campania,
    campanias: Array.isArray(context.campanias) ? context.campanias : (context.campania ? [context.campania] : []),
    authorized: context.authorized,
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
