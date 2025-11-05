/**
 * campanias.gs
 * Funciones relacionadas con campañas y permisos de acceso.
 */

function puedeCobrarCampania(campania) {
  const context = getUserContext();
  ensureAuthorizedContext(context);

  let targetCampania;
  try {
    targetCampania = resolveCampaniaForContext(campania, context);
  } catch (err) {
    if (context.isAdmin) {
      throw err;
    }
    return false;
  }

  if (!targetCampania) return false;

  const sh = sheetCampanias || ss.getSheetByName(SHEET_CAMPANIAS);
  if (!sh) return false;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return false;

  const normalizedTarget = normalizeCampaniaValue(targetCampania);
  if (!normalizedTarget) return false;

  const registros = sh.getRange(2, 1, lastRow - 1, 2).getValues();
  return registros.some(([campaniaNombre, banderaCobrar]) => {
    const campaniaNormalizada = normalizeCampaniaValue(campaniaNombre);
    if (!campaniaNormalizada || campaniaNormalizada !== normalizedTarget) return false;

    const bandera = normalizeString(banderaCobrar);
    return bandera === 'si' || bandera === 'sí';
  });
}

function getCampanias() {
  const context = getUserContext();
  ensureAuthorizedContext(context);

  const sh = sheetHoras;
  if (!sh) throw new Error("No se ha encontrado la hoja 'Horas trabajar'");
  const lastRow = sh.getLastRow();
  let vals = [];
  if (lastRow >= 2) {
    vals = sh
      .getRange(2, COL.CAMPAÑA + 1, lastRow - 1, 1)
      .getValues()
      .flat()
      .map(value => String(value || '').trim())
      .filter(Boolean);
  }

  const uniqueSheetValues = Array.from(new Set(vals));
  if (context.isAdmin) {
    return uniqueSheetValues.sort();
  }

  const allowed = getContextAllowedCampanias(context);
  const allowedNormalized = getContextAllowedCampaniasNormalized(context);
  if (!allowedNormalized.length) {
    return [];
  }

  const result = [];
  allowed.forEach(campaniaPermitida => {
    const normalizedPermitida = normalizeCampaniaValue(campaniaPermitida);
    if (!normalizedPermitida) return;
    const match = uniqueSheetValues.find(
      item => normalizeCampaniaValue(item) === normalizedPermitida
    );
    const valueToUse = match || campaniaPermitida;
    const alreadyIncluded = result.some(
      existing => normalizeCampaniaValue(existing) === normalizedPermitida
    );
    if (!alreadyIncluded) {
      result.push(valueToUse);
    }
  });

  return result;
}

function getCampaniasLibrar() {
  const context = getUserContext();
  ensureAuthorizedContext(context);

  const sh = sheetHorasLibrar;
  if (!sh) throw new Error("No se ha encontrado la hoja 'Horas librar'");
  const lastRow = sh.getLastRow();
  let vals = [];
  if (lastRow >= 2) {
    vals = sh
      .getRange(2, COL.CAMPAÑA + 1, lastRow - 1, 1)
      .getValues()
      .flat()
      .map(value => String(value || '').trim())
      .filter(Boolean);
  }

  const uniqueSheetValues = Array.from(new Set(vals));
  if (context.isAdmin) {
    return uniqueSheetValues.sort();
  }

  const allowed = getContextAllowedCampanias(context);
  const allowedNormalized = getContextAllowedCampaniasNormalized(context);
  if (!allowedNormalized.length) {
    return [];
  }

  const result = [];
  allowed.forEach(campaniaPermitida => {
    const normalizedPermitida = normalizeCampaniaValue(campaniaPermitida);
    if (!normalizedPermitida) return;
    const match = uniqueSheetValues.find(
      item => normalizeCampaniaValue(item) === normalizedPermitida
    );
    const valueToUse = match || campaniaPermitida;
    const alreadyIncluded = result.some(
      existing => normalizeCampaniaValue(existing) === normalizedPermitida
    );
    if (!alreadyIncluded) {
      result.push(valueToUse);
    }
  });

  return result;
}
