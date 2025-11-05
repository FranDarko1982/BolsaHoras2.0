/**
 * utilidades.gs
 * Helpers genéricos reutilizados en distintos módulos del proyecto.
 */

function normalizeString(value) {
  return String(value || '').trim().toLowerCase();
}

function normalizeCampaniaValue(value) {
  return normalizeString(value);
}

function parseCampaniaList(value) {
  if (value == null) return [];
  return String(value)
    .split(/[\n,;|]+/)
    .map(part => part.trim())
    .filter(Boolean);
}

function toLocalIso(dt) {
  return Utilities.formatDate(new Date(dt), getAppTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
}

function parseDateDDMMYYYY(str) {
  if (!str || typeof str !== 'string') return new Date('2100-01-01');
  const [dd, mm, yyyy] = str.split('/');
  return new Date(Number(yyyy), Number(mm) - 1, Number(dd));
}

function getNombreMes(numeroMes) {
  const meses = [
    'enero',
    'febrero',
    'marzo',
    'abril',
    'mayo',
    'junio',
    'julio',
    'agosto',
    'septiembre',
    'octubre',
    'noviembre',
    'diciembre'
  ];
  return meses[numeroMes] || '';
}

function _normalizarFranja(franja) {
  const s = String(franja || '').replace(/\s+/g, '');
  const m = s.match(/^(\d{1,2}):(\d{2})-(\d{1,2}):(\d{2})$/);
  if (!m) return null;
  const h1 = ('0' + m[1]).slice(-2), min1 = m[2];
  const h2 = ('0' + m[3]).slice(-2), min2 = m[4];
  return `${h1}:${min1}-${h2}:${min2}`;
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
  const tz = getAppTimeZone();
  const y = Number(Utilities.formatDate(date, tz, 'yyyy'));
  const m = Number(Utilities.formatDate(date, tz, 'MM'));
  const d = Number(Utilities.formatDate(date, tz, 'dd'));
  return new Date(y, m - 1, d, 0, 0, 0, 0);
}

function _formatMonthLabel(date) {
  const formatter = new Intl.DateTimeFormat('es-ES', { month: 'short' });
  const label = formatter.format(date);
  return label.charAt(0).toUpperCase() + label.slice(1);
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
