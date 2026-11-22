/**
 * ============================================================================
 * OVERCOMERS 2027 (IPTA) — Automatización de Autorizaciones Firmadas en PDF
 * Archivo: Code.gs (Google Apps Script)
 * ============================================================================
 *
 * INSTRUCCIONES DE INSTALACIÓN (PASO A PASO):
 *
 * 1. Abrí el Google Sheets vinculado a tu Google Form de inscripción al
 *    Campamento de Verano 2027 "OVERCOMERS".
 *
 * 2. En el menú superior, andá a:
 *      Extensiones > Apps Script
 *
 * 3. Si el proyecto está vacío, borrá el contenido inicial y pegá TODO este archivo.
 *    Si el proyecto YA tiene otros archivos (por ejemplo, un sistema de recibos),
 *    NO los toques: creá un archivo nuevo con "+" > "Secuencia de comandos",
 *    llamalo "Overcomers" y pegá este código ahí. Todos los nombres internos
 *    llevan el prefijo "ovc"/"Overcomers" para no chocar con otros archivos.
 *    Si otro archivo ya tiene una función `doPost`, renombrala a `doPostRecibos`:
 *    el doPost de este archivo le reenvía todos los pedidos que no sean firmas.
 *    Guardá con Ctrl+S (o el ícono de disquete).
 *
 * 4. Verificá las constantes en `OVERCOMERS_CONFIG` más abajo:
 *    - `MAIN_SHEET_NAME`: debe coincidir con el nombre de la pestaña donde
 *      caen las respuestas del Google Form (en esta planilla: 'Respuestas de formulario').
 *    - `TARGET_COL_INDEX`: 52 corresponde a la columna AZ ("Adjuntar / 附上").
 *    - `DNI_COL_INDEX`: dejalo en `null` para detectar automáticamente la columna
 *      de DNI/Pasaporte del inscripto por el encabezado de la fila 1, o poné el
 *      número exacto de columna (ej. 5 para columna E) si preferís fijarlo manualmente.
 *
 * 5. DESPLEGAR COMO WEB APP (para recibir las firmas desde index.html):
 *    a) Hacé clic en el botón azul "Implementar" > "Nueva implementación".
 *    b) Tipo: "Aplicación web".
 *    c) Descripción: "OVERCOMERS 2027 - Autorizaciones".
 *    d) Ejecutar como: "Yo (tu correo)".
 *    e) Quién tiene acceso: "Cualquier persona" (NO "cualquier persona con cuenta de Google").
 *    f) Hacé clic en "Implementar" y aceptá los permisos de Google Drive y Sheets.
 *    g) Copiá la "URL de la aplicación web" (termina en `/exec`) y pegala en la
 *       constante `APPS_SCRIPT_URL` dentro de `index.html`.
 *
 *    IMPORTANTE: Si más adelante modificás `Code.gs`, tenés que ir a
 *    Implementar > Gestionar implementaciones > Editar (lápiz) > Versión:
 *    "Nueva versión" > Implementar para que los cambios entren en vigencia.
 *
 * 6. CONFIGURAR ACTIVADOR `onFormSubmitOvercomers` (para vincular firmas previas):
 *    a) En el menú izquierdo de Apps Script, hacé clic en el ícono del reloj
 *       ("Activadores" / Triggers).
 *    b) Hacé clic en "+ Añadir activador" (abajo a la derecha).
 *    c) Función a ejecutar: `onFormSubmitOvercomers`
 *    d) Implementación: `Principal` (Head)
 *    e) Origen del evento: `De una hoja de cálculo`
 *    f) Tipo de evento: `Al enviarse el formulario`
 *    g) Guardá el activador.
 *
 * ============================================================================
 */

const OVERCOMERS_CONFIG = {
  // Carpeta en Google Drive donde se guardarán todos los PDFs firmados
  FOLDER_NAME: 'Autorizaciones OVERCOMERS 2027',

  // Pestaña principal con las respuestas del Google Form
  MAIN_SHEET_NAME: 'Respuestas de formulario',

  // Pestaña auxiliar creada automáticamente para quienes firman antes del Google Form
  PENDING_SHEET_NAME: 'Firmas_Pendientes',

  // Columna AZ = columna 52 ("Adjuntar / 附上"). Aunque se borre esa pregunta del Form,
  // la columna queda en la planilla y acá se sigue pegando el link del PDF firmado.
  TARGET_COL_INDEX: 52,

  // Dejar en null para detectar automáticamente la columna de DNI/Pasaporte del solicitante,
  // o indicar un número de columna fijo (1 = A, 2 = B, etc.) si se desea forzar una columna específica.
  DNI_COL_INDEX: 5, // Columna E = DNI del inscripto (igual que en EmitirRecibo.gs)

  // Columnas del nombre en "Respuestas de formulario" (C = "Apellido / 姓", D = "Nombre / 名").
  // Se usan para unir firma e inscripción por nombre cuando el DNI no coincide.
  APELLIDO_COL_INDEX: 3,
  NOMBRE_COL_INDEX: 4,

  // Columna AY = columna 51 ("¿serás mayor de edad?"). Si la pregunta se saca del Form,
  // se completa sola con lo que la persona eligió al firmar (solo si la celda está vacía).
  MAYOR_COL_INDEX: 51,
  MAYOR_SI: 'Si / 是',
  MAYOR_NO: 'No / 否'
};

/**
 * ÚNICO doPost del proyecto. Reparte cada pedido al sistema que corresponde:
 *   - Firmas de OVERCOMERS (index.html manda origen: 'overcomers') -> ovcProcesarFirma
 *   - Cualquier otro pedido -> doPostRecibos (el doPost original de EmitirRecibo.gs)
 *
 * IMPORTANTE: en EmitirRecibo.gs la función original tiene que llamarse
 * `doPostRecibos` (no `doPost`), porque en un proyecto solo puede haber un doPost.
 */
function doPost(e) {
  let datos = null;
  try {
    datos = JSON.parse(e.postData.contents);
  } catch (_) {}

  const esFirmaOvercomers =
    datos && (datos.origen === 'overcomers' || (datos.pdfBase64 && datos.dniSolicitante));

  if (esFirmaOvercomers) {
    return ovcProcesarFirma(e);
  }

  if (typeof doPostRecibos === 'function') {
    return doPostRecibos(e);
  }

  return ovcBuildJsonResponse({ ok: false, error: 'Solicitud no reconocida.' });
}

/**
 * Procesa la firma digital enviada por `index.html`.
 * Recibe JSON con:
 *   - origen: 'overcomers'
 *   - nombreSolicitante
 *   - dniSolicitante
 *   - pdfBase64
 *   - (opcionales) esMayor, nombreTutor, dniTutor, fileName
 */
function ovcProcesarFirma(e) {
  const lock = LockService.getScriptLock();
  try {
    // Evita conflictos si varios acampantes envían su firma al mismo tiempo
    lock.waitLock(30000);

    if (!e || !e.postData || !e.postData.contents) {
      return ovcBuildJsonResponse({ ok: false, error: 'No se recibieron datos en la solicitud.' });
    }

    const data = JSON.parse(e.postData.contents);
    const nombreSolicitante = String(data.nombreSolicitante || '').trim();
    const dniRaw = String(data.dniSolicitante || '').trim();
    const dniLimpio = ovcNormalizeDni(dniRaw);
    const pdfBase64Raw = String(data.pdfBase64 || '').trim();
    const mayorTexto = ovcMayorTexto(data.esMayor);

    if (!nombreSolicitante || !dniLimpio || !pdfBase64Raw) {
      return ovcBuildJsonResponse({
        ok: false,
        error: 'Faltan campos obligatorios (nombreSolicitante, dniSolicitante o pdfBase64).'
      });
    }

    // 1. Limpiar posible prefijo data URI y decodificar el PDF desde Base64
    const cleanBase64 = pdfBase64Raw.replace(/^data:application\/pdf;base64,/, '');
    const pdfBytes = Utilities.base64Decode(cleanBase64);

    // 2. Construir el nombre estándar: Autorizacion_[Nombre]_[DNI].pdf
    const nombreArchivoLimpio = ovcSanitizeFileNamePart(nombreSolicitante);
    const finalFileName = 'Autorizacion_' + nombreArchivoLimpio + '_' + dniLimpio + '.pdf';
    const pdfBlob = Utilities.newBlob(pdfBytes, 'application/pdf', finalFileName);

    // 3. Guardar en la carpeta "Autorizaciones OVERCOMERS 2027" en Google Drive
    const folder = ovcGetOrCreateFolder(OVERCOMERS_CONFIG.FOLDER_NAME);
    const file = folder.createFile(pdfBlob);

    // Configurar permisos de lectura mediante enlace
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const driveUrl = 'https://drive.google.com/open?id=' + file.getId();

    // 4. Buscar la fila del inscripto en la hoja principal de respuestas
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = ovcGetMainResponsesSheet(ss);
    const dniColIndex = ovcFindDniColumnIndex(mainSheet);

    let matchedRow = -1;
    let porNombre = false;
    if (dniColIndex > 0) {
      matchedRow = ovcFindRowByDni(mainSheet, dniColIndex, dniLimpio);
    }
    if (matchedRow <= 0) {
      // Respaldo: el DNI no coincide (mal tipeado) -> buscar por nombre y apellido
      matchedRow = ovcFindRowByName(mainSheet, ovcNameKey(nombreSolicitante));
      porNombre = matchedRow > 0;
    }

    if (matchedRow > 0) {
      // Escribir el enlace directo de Drive en la columna AZ (columna 52)
      mainSheet.getRange(matchedRow, OVERCOMERS_CONFIG.TARGET_COL_INDEX).setValue(driveUrl);
      ovcFillMayorIfEmpty(mainSheet, matchedRow, mayorTexto);

      // Si existía algún registro previo en Firmas_Pendientes para este DNI, marcarlo como Vinculado
      ovcMarkPendingAsLinkedIfExists(ss, dniLimpio);

      // Si se unió por nombre (DNI distinto), dejarlo anotado para que la Coordinación lo revise
      if (porNombre) {
        ovcGetOrCreatePendingSheet(ss).appendRow([
          new Date(), nombreSolicitante, dniLimpio, driveUrl, 'Vinculado (por nombre)', mayorTexto
        ]);
      }

      return ovcBuildJsonResponse({
        ok: true,
        status: porNombre ? 'linked_by_name' : 'linked',
        matchedRow: matchedRow,
        fileName: finalFileName,
        driveUrl: driveUrl
      });
    } else {
      // 5. El participante firmó ANTES de enviar el Google Form -> guardar en Firmas_Pendientes
      const pendingSheet = ovcGetOrCreatePendingSheet(ss);
      pendingSheet.appendRow([
        new Date(),
        nombreSolicitante,
        dniLimpio,
        driveUrl,
        'Pendiente',
        mayorTexto
      ]);

      return ovcBuildJsonResponse({
        ok: true,
        status: 'pending',
        fileName: finalFileName,
        driveUrl: driveUrl
      });
    }
  } catch (err) {
    return ovcBuildJsonResponse({
      ok: false,
      error: err && err.message ? err.message : String(err)
    });
  } finally {
    try {
      lock.releaseLock();
    } catch (_) {}
  }
}

/**
 * Activador instalable: "Al enviarse el formulario" (On form submit).
 * Cuando entra una nueva inscripción al Google Sheets, revisa si el DNI ya
 * tenía una firma guardada en `Firmas_Pendientes`. Si coincide, pega el enlace
 * de Google Drive en la columna AZ (columna 52) y la marca como "Vinculado".
 */
function onFormSubmitOvercomers(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const pendingSheet = ss.getSheetByName(OVERCOMERS_CONFIG.PENDING_SHEET_NAME);
    if (!pendingSheet || pendingSheet.getLastRow() < 2) {
      return; // No hay firmas pendientes por vincular
    }

    // Determinar hoja y fila recién ingresada
    const targetSheet = (e && e.range) ? e.range.getSheet() : ovcGetMainResponsesSheet(ss);
    const newRow = (e && e.range) ? e.range.getRow() : targetSheet.getLastRow();
    if (!targetSheet || newRow < 2) return;

    const dniColIndex = ovcFindDniColumnIndex(targetSheet);
    const submittedDni = dniColIndex > 0
      ? ovcNormalizeDni(targetSheet.getRange(newRow, dniColIndex).getValue())
      : '';
    const submittedNameKey = ovcNameKey(
      targetSheet.getRange(newRow, OVERCOMERS_CONFIG.NOMBRE_COL_INDEX).getValue() + ' ' +
      targetSheet.getRange(newRow, OVERCOMERS_CONFIG.APELLIDO_COL_INDEX).getValue()
    );
    if (!submittedDni && !submittedNameKey) return;

    // Firmas_Pendientes -> [0] Fecha, [1] Nombre, [2] DNI, [3] URL_Drive, [4] Estado, [5] Mayor de edad
    const pendingLastRow = pendingSheet.getLastRow();
    const pendingValues = pendingSheet.getRange(2, 1, pendingLastRow - 1, 6).getValues();

    // 1) Por DNI; 2) si no hay, por nombre y apellido. Siempre la firma más reciente.
    let idx = ovcBuscarPendiente(pendingValues, (row) => submittedDni && ovcNormalizeDni(row[2]) === submittedDni);
    let porNombre = false;
    if (idx < 0 && ovcNameKeyUsable(submittedNameKey)) {
      idx = ovcBuscarPendiente(pendingValues, (row) => ovcNameKey(row[1]) === submittedNameKey);
      porNombre = idx >= 0;
    }
    if (idx < 0) return;

    // Pegar el enlace de Google Drive en la columna AZ (52) de la nueva fila
    targetSheet.getRange(newRow, OVERCOMERS_CONFIG.TARGET_COL_INDEX).setValue(String(pendingValues[idx][3]).trim());
    ovcFillMayorIfEmpty(targetSheet, newRow, pendingValues[idx][5]);

    // Marcar la firma usada y dejar como "Reemplazada" las firmas viejas de la misma persona
    pendingSheet.getRange(idx + 2, 5).setValue(porNombre ? 'Vinculado (por nombre)' : 'Vinculado');
    const mismaPersona = (row) =>
      (submittedDni && ovcNormalizeDni(row[2]) === submittedDni) ||
      (porNombre && ovcNameKey(row[1]) === submittedNameKey);
    for (let i = 0; i < pendingValues.length; i++) {
      if (i !== idx && ovcPendienteDisponible(pendingValues[i]) && mismaPersona(pendingValues[i])) {
        pendingSheet.getRange(i + 2, 5).setValue('Reemplazada');
      }
    }
  } finally {
    try {
      lock.releaseLock();
    } catch (_) {}
  }
}

/**
 * Función utilitaria opcional (ejecutable manualmente desde el editor de Apps Script):
 * Recorre todas las filas con estado "Pendiente" en `Firmas_Pendientes` e intenta
 * vincularlas con `Respuestas de formulario` si el participante ya se inscribió.
 */
function sincronizarPendientesManualmente() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ovcGetMainResponsesSheet(ss);
  const pendingSheet = ss.getSheetByName(OVERCOMERS_CONFIG.PENDING_SHEET_NAME);
  if (!mainSheet || !pendingSheet || pendingSheet.getLastRow() < 2) return;

  const dniColIndex = ovcFindDniColumnIndex(mainSheet);
  if (dniColIndex <= 0) {
    throw new Error('No se pudo detectar la columna de DNI en ' + mainSheet.getName());
  }

  const pendingLastRow = pendingSheet.getLastRow();
  const pendingValues = pendingSheet.getRange(2, 1, pendingLastRow - 1, 6).getValues();
  let vinculados = 0;
  const filasVinculadas = [];

  // De la más reciente a la más vieja: si una persona firmó dos veces, se usa la última
  for (let i = pendingValues.length - 1; i >= 0; i--) {
    const driveUrl = String(pendingValues[i][3] || '').trim();
    if (!driveUrl || !ovcPendienteDisponible(pendingValues[i])) continue;

    const pendingDni = ovcNormalizeDni(pendingValues[i][2]);
    let matchedRow = pendingDni ? ovcFindRowByDni(mainSheet, dniColIndex, pendingDni) : -1;
    let porNombre = false;
    if (matchedRow <= 0) {
      matchedRow = ovcFindRowByName(mainSheet, ovcNameKey(pendingValues[i][1]));
      porNombre = matchedRow > 0;
    }
    if (matchedRow > 0) {
      const yaTieneLink = String(mainSheet.getRange(matchedRow, OVERCOMERS_CONFIG.TARGET_COL_INDEX).getValue() || '').trim();
      if (yaTieneLink && yaTieneLink !== driveUrl && vinculadoEnEstaPasada(matchedRow)) {
        pendingSheet.getRange(i + 2, 5).setValue('Reemplazada');
        continue;
      }
      mainSheet.getRange(matchedRow, OVERCOMERS_CONFIG.TARGET_COL_INDEX).setValue(driveUrl);
      ovcFillMayorIfEmpty(mainSheet, matchedRow, pendingValues[i][5]);
      pendingSheet.getRange(i + 2, 5).setValue(porNombre ? 'Vinculado (por nombre)' : 'Vinculado');
      filasVinculadas.push(matchedRow);
      vinculados++;
    }
  }

  function vinculadoEnEstaPasada(row) {
    return filasVinculadas.indexOf(row) !== -1;
  }

  Logger.log('Sincronización completada. Firmas vinculadas: ' + vinculados);
}

/* ============================================================================
 * FUNCIONES AUXILIARES INTERNAS
 * ============================================================================ */

/**
 * Clave para comparar nombres sin importar orden, mayúsculas, tildes ni signos:
 * "Facundo Ho", "HO, facundo" y "Ho Facúndo" -> "FACUNDO HO".
 */
function ovcNameKey(text) {
  return String(text || '')
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .toUpperCase()
    .replace(/[^A-Z0-9一-鿿\s]/g, ' ')
    .split(/\s+/)
    .filter(Boolean)
    .sort()
    .join(' ');
}

/** Solo se une por nombre si hay al menos nombre y apellido (2 palabras). */
function ovcNameKeyUsable(key) {
  return Boolean(key) && key.split(' ').length >= 2;
}

/** Una firma pendiente se puede usar si todavía no fue vinculada ni reemplazada. */
function ovcPendienteDisponible(row) {
  const estado = String(row[4] || '').trim();
  return Boolean(String(row[3] || '').trim()) && (estado === '' || estado === 'Pendiente');
}

/** Índice de la firma pendiente más reciente que cumple la condición, o -1. */
function ovcBuscarPendiente(pendingValues, condicion) {
  for (let i = pendingValues.length - 1; i >= 0; i--) {
    if (ovcPendienteDisponible(pendingValues[i]) && condicion(pendingValues[i])) return i;
  }
  return -1;
}

/**
 * Busca la inscripción más reciente cuyo "Nombre" + "Apellido" coincida con la clave.
 * Respaldo para cuando el DNI se tipeó mal. Devuelve la fila (1-based) o -1.
 */
function ovcFindRowByName(sheet, nameKey) {
  if (!ovcNameKeyUsable(nameKey)) return -1;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const colAp = OVERCOMERS_CONFIG.APELLIDO_COL_INDEX;
  const colNo = OVERCOMERS_CONFIG.NOMBRE_COL_INDEX;
  const first = Math.min(colAp, colNo);
  const width = Math.abs(colAp - colNo) + 1;
  const values = sheet.getRange(2, first, lastRow - 1, width).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    const key = ovcNameKey(values[i][colNo - first] + ' ' + values[i][colAp - first]);
    if (key === nameKey) return i + 2;
  }
  return -1;
}

/** true -> "Si / 是", false -> "No / 否", cualquier otra cosa -> "" (igual que las opciones del Form) */
function ovcMayorTexto(esMayor) {
  if (esMayor === true) return OVERCOMERS_CONFIG.MAYOR_SI;
  if (esMayor === false) return OVERCOMERS_CONFIG.MAYOR_NO;
  return '';
}

/** Completa la columna "¿serás mayor de edad?" (AY) solo si está vacía. */
function ovcFillMayorIfEmpty(sheet, row, mayorTexto) {
  const texto = String(mayorTexto || '').trim();
  if (!texto || !OVERCOMERS_CONFIG.MAYOR_COL_INDEX) return;
  const cell = sheet.getRange(row, OVERCOMERS_CONFIG.MAYOR_COL_INDEX);
  if (String(cell.getValue() || '').trim() === '') {
    cell.setValue(texto);
  }
}

function ovcNormalizeDni(value) {
  return String(value || '')
    .trim()
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, '');
}

function ovcSanitizeFileNamePart(str) {
  return String(str || '')
    .trim()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-zA-Z0-9\u4e00-\u9fff\s_-]/g, '')
    .replace(/\s+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_|_$/g, '') || 'Participante';
}

function ovcGetOrCreateFolder(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(folderName);
}

function ovcGetMainResponsesSheet(ss) {
  const sheet = ss.getSheetByName(OVERCOMERS_CONFIG.MAIN_SHEET_NAME);
  if (sheet) return sheet;
  // Respaldo: devuelve la primera pestaña que no sea Firmas_Pendientes
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getName() !== OVERCOMERS_CONFIG.PENDING_SHEET_NAME) {
      return sheets[i];
    }
  }
  return sheets[0];
}

function ovcGetOrCreatePendingSheet(ss) {
  const headers = ['Fecha', 'Nombre', 'DNI', 'URL_Drive', 'Estado', 'Mayor de edad'];
  let sheet = ss.getSheetByName(OVERCOMERS_CONFIG.PENDING_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(OVERCOMERS_CONFIG.PENDING_SHEET_NAME);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  } else if (String(sheet.getRange(1, 6).getValue() || '').trim() === '') {
    // Pestaña creada con la versión anterior (5 columnas): agregar el título de la 6.ª
    sheet.getRange(1, 6).setValue(headers[5]).setFontWeight('bold');
  }
  return sheet;
}

/**
 * Detecta la columna de DNI / Pasaporte del participante en la fila 1 de encabezados.
 * Prioriza encabezados que contengan DNI/Pasaporte/Documento/身分證/護照
 * evitando columnas que correspondan al padre/madre/tutor.
 */
function ovcFindDniColumnIndex(sheet) {
  if (OVERCOMERS_CONFIG.DNI_COL_INDEX && OVERCOMERS_CONFIG.DNI_COL_INDEX > 0) {
    return OVERCOMERS_CONFIG.DNI_COL_INDEX;
  }

  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return -1;

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  let fallbackMatch = -1;

  for (let c = 0; c < headers.length; c++) {
    const h = String(headers[c] || '').toLowerCase();
    const isDniLike =
      h.indexOf('dni') !== -1 ||
      h.indexOf('pasaporte') !== -1 ||
      h.indexOf('documento') !== -1 ||
      h.indexOf('身分證') !== -1 ||
      h.indexOf('護照') !== -1;

    if (!isDniLike) continue;

    const isTutorLike =
      h.indexOf('tutor') !== -1 ||
      h.indexOf('padre') !== -1 ||
      h.indexOf('madre') !== -1 ||
      h.indexOf('emergencia') !== -1 ||
      h.indexOf('監護人') !== -1 ||
      h.indexOf('父母') !== -1;

    if (!isTutorLike) {
      return c + 1;
    } else if (fallbackMatch === -1) {
      fallbackMatch = c + 1;
    }
  }

  return fallbackMatch;
}

/**
 * Busca la fila más reciente cuyo DNI coincida (ignorando puntos, espacios o guiones).
 * Retorna el número de fila (1-based) o -1 si no existe.
 */
function ovcFindRowByDni(sheet, dniColIndex, targetDniNormalized) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;

  const values = sheet.getRange(2, dniColIndex, lastRow - 1, 1).getValues();
  // Recorre de abajo hacia arriba para priorizar la inscripción más reciente si hubo duplicados
  for (let i = values.length - 1; i >= 0; i--) {
    if (ovcNormalizeDni(values[i][0]) === targetDniNormalized) {
      return i + 2;
    }
  }
  return -1;
}

function ovcMarkPendingAsLinkedIfExists(ss, dniLimpio) {
  const pendingSheet = ss.getSheetByName(OVERCOMERS_CONFIG.PENDING_SHEET_NAME);
  if (!pendingSheet || pendingSheet.getLastRow() < 2) return;

  const lastRow = pendingSheet.getLastRow();
  const values = pendingSheet.getRange(2, 1, lastRow - 1, 5).getValues();
  for (let i = 0; i < values.length; i++) {
    if (ovcNormalizeDni(values[i][2]) === dniLimpio && values[i][4] !== 'Vinculado') {
      pendingSheet.getRange(i + 2, 5).setValue('Vinculado');
    }
  }
}

function ovcBuildJsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
