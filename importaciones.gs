/**
 * importaciones.gs
 * Procesa archivos Excel/CSV para cargar disponibilidades en las hojas maestras.
 */

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
