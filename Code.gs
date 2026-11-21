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
 * 3. Borrá el contenido inicial del editor y pegá TODO este archivo `Code.gs`.
 *    Guardá con Ctrl+S (o el ícono de disquete).
 *
 * 4. Verificá las constantes en `CONFIG` más abajo:
 *    - `MAIN_SHEET_NAME`: debe coincidir con el nombre de la pestaña donde
 *      caen las respuestas del Google Form (por defecto: 'Respuestas de formulario 1').
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
 *    e) Quién tiene acceso: "Cualquier usuario".
 *    f) Hacé clic en "Implementar" y aceptá los permisos de Google Drive y Sheets.
 *    g) Copiá la "URL de la aplicación web" (termina en `/exec`) y pegala en la
 *       constante `APPS_SCRIPT_URL` dentro de `index.html`.
 *
 *    IMPORTANTE: Si más adelante modificás `Code.gs`, tenés que ir a
 *    Implementar > Gestionar implementaciones > Editar (lápiz) > Versión:
 *    "Nueva versión" > Implementar para que los cambios entren en vigencia.
 *
 * 6. CONFIGURAR ACTIVADOR `onFormSubmit` (para vincular firmas previas):
 *    a) En el menú izquierdo de Apps Script, hacé clic en el ícono del reloj
 *       ("Activadores" / Triggers).
 *    b) Hacé clic en "+ Añadir activador" (abajo a la derecha).
 *    c) Función a ejecutar: `onFormSubmit`
 *    d) Implementación: `Principal` (Head)
 *    e) Origen del evento: `De una hoja de cálculo`
 *    f) Tipo de evento: `Al enviarse el formulario`
 *    g) Guardá el activador.
 *
 * ============================================================================
 */

const CONFIG = {
  // Carpeta en Google Drive donde se guardarán todos los PDFs firmados
  FOLDER_NAME: 'Autorizaciones OVERCOMERS 2027',

  // Pestaña principal con las respuestas del Google Form
  MAIN_SHEET_NAME: 'Respuestas de formulario 1',

  // Pestaña auxiliar creada automáticamente para quienes firman antes del Google Form
  PENDING_SHEET_NAME: 'Firmas_Pendientes',

  // Columna AZ = columna 52 ("Adjuntar / 附上")
  TARGET_COL_INDEX: 52,

  // Dejar en null para detectar automáticamente la columna de DNI/Pasaporte del solicitante,
  // o indicar un número de columna fijo (1 = A, 2 = B, etc.) si se desea forzar una columna específica.
  DNI_COL_INDEX: null
};

/**
 * Endpoint GET simple para verificar en el navegador que la Web App está activa.
 */
function doGet() {
  return buildJsonResponse({
    ok: true,
    service: 'OVERCOMERS 2027 - Autorizaciones IPTA',
    timestamp: new Date().toISOString()
  });
}

/**
 * Endpoint POST invocado por `index.html` al confirmar la firma digital.
 * Recibe JSON con:
 *   - nombreSolicitante
 *   - dniSolicitante
 *   - pdfBase64
 *   - (opcionales) esMayor, nombreTutor, dniTutor, fileName
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    // Evita conflictos si varios acampantes envían su firma al mismo tiempo
    lock.waitLock(30000);

    if (!e || !e.postData || !e.postData.contents) {
      return buildJsonResponse({ ok: false, error: 'No se recibieron datos en la solicitud.' });
    }

    const data = JSON.parse(e.postData.contents);
    const nombreSolicitante = String(data.nombreSolicitante || '').trim();
    const dniRaw = String(data.dniSolicitante || '').trim();
    const dniLimpio = normalizeDni(dniRaw);
    const pdfBase64Raw = String(data.pdfBase64 || '').trim();

    if (!nombreSolicitante || !dniLimpio || !pdfBase64Raw) {
      return buildJsonResponse({
        ok: false,
        error: 'Faltan campos obligatorios (nombreSolicitante, dniSolicitante o pdfBase64).'
      });
    }

    // 1. Limpiar posible prefijo data URI y decodificar el PDF desde Base64
    const cleanBase64 = pdfBase64Raw.replace(/^data:application\/pdf;base64,/, '');
    const pdfBytes = Utilities.base64Decode(cleanBase64);

    // 2. Construir el nombre estándar: Autorizacion_[Nombre]_[DNI].pdf
    const nombreArchivoLimpio = sanitizeFileNamePart(nombreSolicitante);
    const finalFileName = 'Autorizacion_' + nombreArchivoLimpio + '_' + dniLimpio + '.pdf';
    const pdfBlob = Utilities.newBlob(pdfBytes, 'application/pdf', finalFileName);

    // 3. Guardar en la carpeta "Autorizaciones OVERCOMERS 2027" en Google Drive
    const folder = getOrCreateFolder(CONFIG.FOLDER_NAME);
    const file = folder.createFile(pdfBlob);

    // Configurar permisos de lectura mediante enlace
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const driveUrl = 'https://drive.google.com/open?id=' + file.getId();

    // 4. Buscar la fila del inscripto en la hoja principal de respuestas
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = getMainResponsesSheet(ss);
    const dniColIndex = findDniColumnIndex(mainSheet);

    let matchedRow = -1;
    if (dniColIndex > 0) {
      matchedRow = findRowByDni(mainSheet, dniColIndex, dniLimpio);
    }

    if (matchedRow > 0) {
      // Escribir el enlace directo de Drive en la columna AZ (columna 52)
      mainSheet.getRange(matchedRow, CONFIG.TARGET_COL_INDEX).setValue(driveUrl);

      // Si existía algún registro previo en Firmas_Pendientes para este DNI, marcarlo como Vinculado
      markPendingAsLinkedIfExists(ss, dniLimpio);

      return buildJsonResponse({
        ok: true,
        status: 'linked',
        matchedRow: matchedRow,
        fileName: finalFileName,
        driveUrl: driveUrl
      });
    } else {
      // 5. El participante firmó ANTES de enviar el Google Form -> guardar en Firmas_Pendientes
      const pendingSheet = getOrCreatePendingSheet(ss);
      pendingSheet.appendRow([
        new Date(),
        nombreSolicitante,
        dniLimpio,
        driveUrl,
        'Pendiente'
      ]);

      return buildJsonResponse({
        ok: true,
        status: 'pending',
        fileName: finalFileName,
        driveUrl: driveUrl
      });
    }
  } catch (err) {
    return buildJsonResponse({
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
function onFormSubmit(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const pendingSheet = ss.getSheetByName(CONFIG.PENDING_SHEET_NAME);
    if (!pendingSheet || pendingSheet.getLastRow() < 2) {
      return; // No hay firmas pendientes por vincular
    }

    // Determinar hoja y fila recién ingresada
    const targetSheet = (e && e.range) ? e.range.getSheet() : getMainResponsesSheet(ss);
    const newRow = (e && e.range) ? e.range.getRow() : targetSheet.getLastRow();
    if (!targetSheet || newRow < 2) return;

    const dniColIndex = findDniColumnIndex(targetSheet);
    if (dniColIndex <= 0) return;

    const submittedDniRaw = targetSheet.getRange(newRow, dniColIndex).getValue();
    const submittedDni = normalizeDni(submittedDniRaw);
    if (!submittedDni) return;

    // Buscar en Firmas_Pendientes (de abajo hacia arriba para tomar la firma más reciente)
    const pendingLastRow = pendingSheet.getLastRow();
    const pendingValues = pendingSheet.getRange(2, 1, pendingLastRow - 1, 5).getValues();
    // Columnas: [0] Fecha, [1] Nombre, [2] DNI, [3] URL_Drive, [4] Estado

    for (let i = pendingValues.length - 1; i >= 0; i--) {
      const pendingDni = normalizeDni(pendingValues[i][2]);
      const driveUrl = String(pendingValues[i][3] || '').trim();
      const estado = String(pendingValues[i][4] || '').trim();

      if (pendingDni === submittedDni && driveUrl && estado !== 'Vinculado') {
        // Pegar el enlace de Google Drive en la columna AZ (52) de la nueva fila
        targetSheet.getRange(newRow, CONFIG.TARGET_COL_INDEX).setValue(driveUrl);

        // Marcar el registro en Firmas_Pendientes como "Vinculado"
        const pendingRowIndex = i + 2;
        pendingSheet.getRange(pendingRowIndex, 5).setValue('Vinculado');
        break;
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
 * vincularlas con `Respuestas de formulario 1` si el participante ya se inscribió.
 */
function sincronizarPendientesManualmente() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = getMainResponsesSheet(ss);
  const pendingSheet = ss.getSheetByName(CONFIG.PENDING_SHEET_NAME);
  if (!mainSheet || !pendingSheet || pendingSheet.getLastRow() < 2) return;

  const dniColIndex = findDniColumnIndex(mainSheet);
  if (dniColIndex <= 0) {
    throw new Error('No se pudo detectar la columna de DNI en ' + mainSheet.getName());
  }

  const pendingLastRow = pendingSheet.getLastRow();
  const pendingValues = pendingSheet.getRange(2, 1, pendingLastRow - 1, 5).getValues();
  let vinculados = 0;

  for (let i = 0; i < pendingValues.length; i++) {
    const pendingDni = normalizeDni(pendingValues[i][2]);
    const driveUrl = String(pendingValues[i][3] || '').trim();
    const estado = String(pendingValues[i][4] || '').trim();

    if (!pendingDni || !driveUrl || estado === 'Vinculado') continue;

    const matchedRow = findRowByDni(mainSheet, dniColIndex, pendingDni);
    if (matchedRow > 0) {
      mainSheet.getRange(matchedRow, CONFIG.TARGET_COL_INDEX).setValue(driveUrl);
      pendingSheet.getRange(i + 2, 5).setValue('Vinculado');
      vinculados++;
    }
  }

  Logger.log('Sincronización completada. Firmas vinculadas: ' + vinculados);
}

/* ============================================================================
 * FUNCIONES AUXILIARES INTERNAS
 * ============================================================================ */

function normalizeDni(value) {
  return String(value || '')
    .trim()
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, '');
}

function sanitizeFileNamePart(str) {
  return String(str || '')
    .trim()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-zA-Z0-9\u4e00-\u9fff\s_-]/g, '')
    .replace(/\s+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_|_$/g, '') || 'Participante';
}

function getOrCreateFolder(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(folderName);
}

function getMainResponsesSheet(ss) {
  const sheet = ss.getSheetByName(CONFIG.MAIN_SHEET_NAME);
  if (sheet) return sheet;
  // Respaldo: devuelve la primera pestaña que no sea Firmas_Pendientes
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getName() !== CONFIG.PENDING_SHEET_NAME) {
      return sheets[i];
    }
  }
  return sheets[0];
}

function getOrCreatePendingSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.PENDING_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.PENDING_SHEET_NAME);
    const headers = ['Fecha', 'Nombre', 'DNI', 'URL_Drive', 'Estado'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Detecta la columna de DNI / Pasaporte del participante en la fila 1 de encabezados.
 * Prioriza encabezados que contengan DNI/Pasaporte/Documento/身分證/護照
 * evitando columnas que correspondan al padre/madre/tutor.
 */
function findDniColumnIndex(sheet) {
  if (CONFIG.DNI_COL_INDEX && CONFIG.DNI_COL_INDEX > 0) {
    return CONFIG.DNI_COL_INDEX;
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
function findRowByDni(sheet, dniColIndex, targetDniNormalized) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;

  const values = sheet.getRange(2, dniColIndex, lastRow - 1, 1).getValues();
  // Recorre de abajo hacia arriba para priorizar la inscripción más reciente si hubo duplicados
  for (let i = values.length - 1; i >= 0; i--) {
    if (normalizeDni(values[i][0]) === targetDniNormalized) {
      return i + 2;
    }
  }
  return -1;
}

function markPendingAsLinkedIfExists(ss, dniLimpio) {
  const pendingSheet = ss.getSheetByName(CONFIG.PENDING_SHEET_NAME);
  if (!pendingSheet || pendingSheet.getLastRow() < 2) return;

  const lastRow = pendingSheet.getLastRow();
  const values = pendingSheet.getRange(2, 1, lastRow - 1, 5).getValues();
  for (let i = 0; i < values.length; i++) {
    if (normalizeDni(values[i][2]) === dniLimpio && values[i][4] !== 'Vinculado') {
      pendingSheet.getRange(i + 2, 5).setValue('Vinculado');
    }
  }
}

function buildJsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
