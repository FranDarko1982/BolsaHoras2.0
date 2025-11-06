/**
 * plantillas.gs
 * Funciones para consultar las plantillas de importación almacenadas en Drive.
 */

function obtenerPlantillasExcel() {
  const context = getUserContext();
  ensureAuthorizedContext(context);

  if (!context.isAdmin && !context.isGestor) {
    throw new Error('Solo los administradores o gestores pueden consultar las plantillas.');
  }

  if (!PLANTILLAS_FOLDER_ID) {
    return { files: [] };
  }

  let folder;
  try {
    folder = DriveApp.getFolderById(PLANTILLAS_FOLDER_ID);
  } catch (error) {
    throw new Error('No se pudo acceder a la carpeta de plantillas configurada.');
  }

  if (!folder) {
    return { files: [] };
  }

  const MIME = typeof MimeType === 'object' && MimeType ? MimeType : {};
  const allowedExtensions = ['xlsx', 'xls', 'xlsm', 'xlsb', 'csv'];
  const allowedMimeTypes = [
    MIME.MICROSOFT_EXCEL,
    MIME.MICROSOFT_EXCEL_LEGACY,
    MIME.CSV,
    'application/vnd.ms-excel.sheet.macroEnabled.12',
    'application/vnd.ms-excel.sheet.binary.macroEnabled.12',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'application/vnd.ms-excel'
  ]
    .map(function (value) {
      return typeof value === 'string' ? value.toLowerCase() : '';
    })
    .filter(function (value, index, array) {
      return value && array.indexOf(value) === index;
    });

  const files = [];
  const iterator = folder.getFiles();

  while (iterator.hasNext()) {
    const file = iterator.next();
    if (!file) continue;

    const name = String(file.getName() || '').trim();
    const id = file.getId();
    if (!id || !name) continue;

    const extension = _extractPlantillaExtension(name);
    const mimeType = typeof file.getMimeType === 'function' ? String(file.getMimeType() || '') : '';
    const extensionLower = extension ? extension.toLowerCase() : '';
    const mimeTypeLower = mimeType.toLowerCase();
    const hasAllowedExtension = extensionLower ? allowedExtensions.includes(extensionLower) : false;
    const hasAllowedMime = mimeTypeLower ? allowedMimeTypes.includes(mimeTypeLower) : false;

    if (!hasAllowedExtension && !hasAllowedMime) {
      continue;
    }

    const updatedAt = typeof file.getLastUpdated === 'function' ? file.getLastUpdated() : null;
    const sizeBytes = typeof file.getSize === 'function' ? Number(file.getSize()) : null;
    const url = typeof file.getUrl === 'function' ? file.getUrl() : '';

    files.push({
      id: id,
      name: name,
      mimeType: mimeType,
      extension: extension ? '.' + extension.toLowerCase() : '',
      updatedAt: updatedAt instanceof Date ? updatedAt.toISOString() : '',
      sizeBytes: Number.isFinite(sizeBytes) && sizeBytes >= 0 ? sizeBytes : null,
      url: url,
      downloadUrl: _buildPlantillaDownloadUrl(id),
      typeLabel: _resolvePlantillaTypeLabel(extension, mimeType)
    });
  }

  files.sort((a, b) => {
    const dateA = a && a.updatedAt ? new Date(a.updatedAt).getTime() : 0;
    const dateB = b && b.updatedAt ? new Date(b.updatedAt).getTime() : 0;
    if (dateA !== dateB) {
      return dateB - dateA;
    }
    const nameA = a && a.name ? a.name.toLowerCase() : '';
    const nameB = b && b.name ? b.name.toLowerCase() : '';
    return nameA.localeCompare(nameB);
  });

  return { files: files };
}

function _extractPlantillaExtension(name) {
  const normalized = typeof name === 'string' ? name.trim() : '';
  if (!normalized) return '';
  const lastDot = normalized.lastIndexOf('.');
  if (lastDot === -1 || lastDot === normalized.length - 1) return '';
  return normalized.substring(lastDot + 1);
}

function _buildPlantillaDownloadUrl(fileId) {
  const id = String(fileId || '').trim();
  if (!id) return '';
  return 'https://drive.google.com/uc?export=download&id=' + encodeURIComponent(id);
}

function _resolvePlantillaTypeLabel(extension, mimeType) {
  const ext = (extension || '').toLowerCase();
  if (ext === 'csv') return 'CSV';
  if (ext === 'xls' || ext === 'xlsx' || ext === 'xlsm' || ext === 'xlsb') return 'Excel';
  const mime = (mimeType || '').toLowerCase();
  if (mime.indexOf('spreadsheetml') !== -1 || mime.indexOf('excel') !== -1) return 'Excel';
  if (mime.indexOf('csv') !== -1) return 'CSV';
  return 'Archivo';
}
