function unificarDatosLocker() {
  // === Config ===
  const SOURCE_SHEETS = ["bbdd reservas horas trabajar", "bbdd reservas horas librar"];
  const DEST_SHEET = "Datos Locker";
  const EXTERNAL_ID = "1zn02OCArJNv7zCpTSER4MJzu6Rzt_VgyJLqV63Pw7Ko"; // A:B => NºEmpleado | email
  const EMAIL_COL_INDEX = 5;    // Columna F (0-based)
  const EMP_COL_INSERT_AT = 10; // Insertar NºEmpleado en K (0-based)
  const CAMPAÑA_COL_INDEX = 0;  // Columna A (0-based), usada para filtrar "teletrabajo"

  // ID de la hoja de cálculo. Es más robusto usar openById para activadores.
  const SPREADSHEET_ID = '1g3K7lGBzD-KfAg8xpyW0RtbkfSYd-sAwcGo-zXSIDC0';
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  if (!ss) throw new Error(`No se pudo abrir la hoja de cálculo con ID: ${SPREADSHEET_ID}`);

  // === 1) Mapa email -> NºEmpleado desde la hoja externa ===
  const externalSS = SpreadsheetApp.openById(EXTERNAL_ID);
  const externalSheet = externalSS.getSheets()[0];
  const extRange = externalSheet.getRange(1, 1, externalSheet.getLastRow(), 2).getValues();
  const emailToEmp = {}; // email (lower) -> NºEmpleado (string)

  for (let i = 0; i < extRange.length; i++) {
    const emp = extRange[i][0];   // Col A
    const email = extRange[i][1]; // Col B
    if (email && typeof email === "string" && email.indexOf("@") !== -1) {
      emailToEmp[email.trim().toLowerCase()] = (emp == null ? "" : String(emp).trim());
    }
  }

  // === 2) Crear/obtener hoja destino ===
  let dest = ss.getSheetByName(DEST_SHEET);
  if (!dest) dest = ss.insertSheet(DEST_SHEET);

  // Limpia contenidos (mantén formato) a partir de la fila 2
  if (dest.getMaxRows() > 1) {
    const toClearRows = Math.max(1, dest.getMaxRows() - 1);
    dest.getRange(2, 1, toClearRows, dest.getMaxColumns()).clearContent();
  }

  // === 3) Unificar datos (omite cabecera, filas vacías y filas con Campaña = teletrabajo) ===
  let header = null;
  let unified = [];
  let maxCols = 0;

  SOURCE_SHEETS.forEach((name) => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;

    const values = sh.getDataRange().getValues();
    if (!values || values.length < 2) return;

    if (!header) header = values[0];
    maxCols = Math.max(maxCols, values[0].length);

    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      // ✅ Saltar filas vacías
      if (isEmptyRow_(row)) continue;
      // ✅ Saltar filas con Campaña = teletrabajo (sin importar mayúsculas/minúsculas)
      const campañaVal = row[CAMPAÑA_COL_INDEX];
      if (campañaVal && String(campañaVal).trim().toLowerCase() === "teletrabajo") continue;

      maxCols = Math.max(maxCols, row.length);
      unified.push(row);
    }
  });

  if (!header) {
    dest.getRange(1, 1, 1, 1).setValue("Sin datos en hojas origen.");
    return;
  }

  // === 4) Normalizar ancho de columnas ===
  header = padToLen_(header, maxCols);
  unified = unified.map(row => padToLen_(row, maxCols));

  // === 5) Insertar columna "NºEmpleado" en K (índice 10) y cruzar por email (F) ===
  const NEW_COL_NAME = "NºEmpleado";
  header = padToLen_(header, EMP_COL_INSERT_AT);
  header.splice(EMP_COL_INSERT_AT, 0, NEW_COL_NAME);

  for (let i = 0; i < unified.length; i++) {
    let row = unified[i];
    row = padToLen_(row, EMP_COL_INSERT_AT);
    const rawEmail = row[EMAIL_COL_INDEX];
    const emailKey = (rawEmail && typeof rawEmail === "string") ? rawEmail.trim().toLowerCase() : "";
    const empNum = emailKey ? (emailToEmp[emailKey] || "") : "";
    row.splice(EMP_COL_INSERT_AT, 0, empNum);
    unified[i] = row;
  }

  // === 6) Salida final y pegado ===
  const out = [header].concat(unified);
  ensureSheetSize_(dest, out.length, out[0].length);
  dest.getRange(1, 1, out.length, out[0].length).setValues(out);
}

/** Devuelve true si TODAS las celdas de la fila están vacías. */
function isEmptyRow_(row) {
  return row.every(cell => cell == null || String(cell).trim() === "");
}

/** Rellena el array hasta length con "". */
function padToLen_(arr, length) {
  return arr.length >= length ? arr : arr.concat(Array(length - arr.length).fill(""));
}

/** Asegura tamaño mínimo de la hoja destino. */
function ensureSheetSize_(sheet, rows, cols) {
  const maxRows = sheet.getMaxRows();
  if (maxRows < rows) sheet.insertRowsAfter(maxRows, rows - maxRows);
  else if (maxRows > rows + 50) sheet.deleteRows(rows + 1, maxRows - rows);

  const maxCols = sheet.getMaxColumns();
  if (maxCols < cols) sheet.insertColumnsAfter(maxCols, cols - maxCols);
  else if (maxCols > cols + 10) sheet.deleteColumns(cols + 1, maxCols - cols);
}
