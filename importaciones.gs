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

function cargarEmpleadosDesdeExcel(payload) {
  const context = getUserContext();
  if (!context || !context.isAdmin) {
    throw new Error('Solo los administradores pueden cargar empleados.');
  }

  if (!sheetEmpleados) {
    throw new Error("No se ha encontrado la hoja 'Empleados'.");
  }

  if (!payload || !payload.base64) {
    throw new Error('No se ha recibido ningún archivo para procesar.');
  }

  const fileName = payload.fileName || 'empleados.xlsx';
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
      { title: `tmp-empleados-${Date.now()}`, mimeType: MimeType.GOOGLE_SHEETS },
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

    const sourceHeader = data[0];
    const sourceHeaderMap = {};
    (sourceHeader || []).forEach((value, index) => {
      const key = _normalizeHeaderName(value);
      if (key && !(key in sourceHeaderMap)) {
        sourceHeaderMap[key] = index;
      }
    });

    const sourceColumns = {
      empleado: _getFirstMatchingIndex(sourceHeaderMap, [
        'empleado',
        'numempleado',
        'numeroempleado',
        'nempleado',
        'idempleado'
      ]),
      correo: _getFirstMatchingIndex(sourceHeaderMap, ['correo', 'email', 'mail']),
      campania: _getFirstMatchingIndex(sourceHeaderMap, ['campana', 'campaña']),
      rol: _getFirstMatchingIndex(sourceHeaderMap, ['rol', 'role'])
    };

    if (sourceColumns.empleado == null && sourceColumns.correo == null) {
      throw new Error('El archivo debe incluir al menos las columnas de empleado o correo.');
    }

    const targetInfo = getEmpleadosHeaderInfo();
    if (!targetInfo || !targetInfo.totalColumns) {
      throw new Error('No se pudo leer la estructura de la hoja Empleados.');
    }

    const lastRow = sheetEmpleados.getLastRow();
    const existingValues =
      lastRow > 1
        ? sheetEmpleados.getRange(2, 1, lastRow - 1, targetInfo.totalColumns).getValues()
        : [];

    const existingEmployeeIds = new Set();
    const existingEmails = new Set();

    existingValues.forEach(row => {
      if (targetInfo.colEmpleado != null && targetInfo.colEmpleado >= 0) {
        const value = normalizeString(row[targetInfo.colEmpleado]);
        if (value) existingEmployeeIds.add(value);
      }
      if (targetInfo.colCorreo != null && targetInfo.colCorreo >= 0) {
        const value = normalizeString(row[targetInfo.colCorreo]);
        if (value) existingEmails.add(value);
      }
    });

    const rowsToInsert = [];
    let skippedDuplicates = 0;
    let skippedInvalid = 0;

    for (let i = 1; i < data.length; i++) {
      const sourceRow = data[i];
      if (_isRowCompletelyEmpty(sourceRow)) continue;

      const empleadoRaw = sourceColumns.empleado != null ? sourceRow[sourceColumns.empleado] : '';
      const correoRaw = sourceColumns.correo != null ? sourceRow[sourceColumns.correo] : '';
      const campaniaRaw = sourceColumns.campania != null ? sourceRow[sourceColumns.campania] : '';
      const rolRaw = sourceColumns.rol != null ? sourceRow[sourceColumns.rol] : '';

      const empleado = empleadoRaw != null ? String(empleadoRaw).trim() : '';
      const correo = correoRaw != null ? String(correoRaw).trim() : '';
      const normalizedEmpleado = normalizeString(empleado);
      const normalizedCorreo = normalizeString(correo);

      if (!normalizedEmpleado && !normalizedCorreo) {
        skippedInvalid++;
        continue;
      }

      if (
        (normalizedEmpleado && existingEmployeeIds.has(normalizedEmpleado)) ||
        (normalizedCorreo && existingEmails.has(normalizedCorreo))
      ) {
        skippedDuplicates++;
        continue;
      }

      const targetRow = new Array(targetInfo.totalColumns).fill('');
      if (targetInfo.colEmpleado != null && targetInfo.colEmpleado >= 0) {
        targetRow[targetInfo.colEmpleado] = empleado;
      }
      if (targetInfo.colCorreo != null && targetInfo.colCorreo >= 0) {
        targetRow[targetInfo.colCorreo] = correo;
      }
      if (targetInfo.colCampania != null && targetInfo.colCampania >= 0) {
        const campaniaValue = campaniaRaw != null ? String(campaniaRaw).trim() : '';
        targetRow[targetInfo.colCampania] = campaniaValue;
      }
      if (targetInfo.colRol != null && targetInfo.colRol >= 0) {
        const rolValue = normalizeRoleValue(rolRaw) || (rolRaw != null ? String(rolRaw).trim() : '');
        targetRow[targetInfo.colRol] = rolValue;
      }

      rowsToInsert.push(targetRow);

      if (normalizedEmpleado) {
        existingEmployeeIds.add(normalizedEmpleado);
      }
      if (normalizedCorreo) {
        existingEmails.add(normalizedCorreo);
      }
    }

    if (!rowsToInsert.length) {
      let message = 'No hay registros nuevos para importar.';
      if (skippedDuplicates > 0) {
        message += ` Se omitieron ${skippedDuplicates} duplicados.`;
      }
      if (skippedInvalid > 0) {
        message += ` Se omitieron ${skippedInvalid} filas sin datos identificativos.`;
      }
      return { importedRows: 0, message };
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const startRow = getFirstFreeRow(sheetEmpleados);
      const requiredRows = startRow + rowsToInsert.length - 1;
      if (requiredRows > sheetEmpleados.getMaxRows()) {
        const rowsToAdd = requiredRows - sheetEmpleados.getMaxRows();
        sheetEmpleados.insertRowsAfter(sheetEmpleados.getMaxRows(), rowsToAdd);
      }
      sheetEmpleados
        .getRange(startRow, 1, rowsToInsert.length, targetInfo.totalColumns)
        .setValues(rowsToInsert);
    } finally {
      lock.releaseLock();
    }

    let message = `Archivo procesado correctamente. Se importaron ${rowsToInsert.length} registros.`;
    if (skippedDuplicates > 0) {
      message += ` Se omitieron ${skippedDuplicates} duplicados.`;
    }
    if (skippedInvalid > 0) {
      message += ` Se omitieron ${skippedInvalid} filas sin datos identificativos.`;
    }

    return { importedRows: rowsToInsert.length, message };
  } catch (error) {
    const message = error && error.message ? error.message : 'Error desconocido al importar empleados.';
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
