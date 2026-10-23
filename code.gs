
/**
 * AGURIZA TRAINING — biblioteca de entrenamientos y registro de sensaciones.
 * Conserva las funciones existentes y añade una API para Netlify.
 * IMPORTANTE: no ejecutar setupProyecto() para esta migración.
 */


function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({
      ok: true,
      service: 'Aguriza Training API',
      message: 'API disponible. Utiliza POST para consultar o guardar.'
    }))
    .setMimeType(ContentService.MimeType.JSON);
}


/** Ejecutar solo si necesitas volver a verificar/reparar la estructura de las hojas. */
function setupProyecto() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Abre este proyecto desde Extensiones > Apps Script en tu Google Sheet.');
  const cambios = [];
  const planA = ss.getSheetByName('Plan A');
  const planB = ss.getSheetByName('Plan B');
  const diario = ss.getSheetByName('Diario de Sensaciones');
  if (!planA) throw new Error('No se encontró la hoja "Plan A".');
  if (!diario) throw new Error('No se encontró la hoja "Diario de Sensaciones".');
  [planA, planB, diario].filter(Boolean).forEach(sheet => backupSheetOnce_(ss, sheet));

  const lastRowA = planA.getLastRow();
  if (lastRowA >= 2) {
    const rows = planA.getRange(2, 1, lastRowA - 1, Math.max(8, planA.getLastColumn())).getValues();
    rows.forEach((row, i) => {
      const option = String(row[1] || '').trim().toLowerCase();
      if (!(row[2] instanceof Date)) return;
      let replacement = null;
      if (option.includes('fondo/montaña')) replacement = '3-4';
      else if (option.includes('grupal')) replacement = '3-5';
      else if (option.includes('hiit ultra cortos')) replacement = '10-12';
      else if (option.includes('over-under')) replacement = '4-5';
      if (replacement !== null) {
        planA.getRange(i + 2, 3).setNumberFormat('@').setValue(replacement);
        cambios.push('Plan A: intervalos corregidos en ' + row[1]);
      }
    });
  }

  if (planB && planB.getLastRow() >= 2) {
    const bData = planB.getDataRange().getValues();
    for (let i = 1; i < bData.length; i++) {
      const row = bData[i];
      if (
        String(row[1] || '').trim().toLowerCase() === 'domingo' &&
        String(row[2] || '').trim().toLowerCase() === 'salida grupal'
      ) {
        const description = String(row[3] || 'Salida grupal ritmo constante').trim();
        planB.getRange(i + 1, 1, 1, 8).clearContent();
        planB.getRange(i + 1, 1, 1, 8).setValues([
          ['Domingo', 'Salida Grupal', '', '', '', 'Ruta', description, 'Z2-Z4']
        ]);
        cambios.push('Plan B: fila de salida grupal reorganizada.');
      }
    }
  }

  const headers = [
    'Fecha', 'Entrenamiento', 'Tiempo Real', 'Distancia (km)',
    'RPE (1-10)', 'Fatiga', 'Notas', 'Comentarios / Notas'
  ];
  diario.getRange(1, 1, 1, headers.length).setValues([headers]);
  diario.setFrozenRows(1);
  diario.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm');
  diario.getRange('D:D').setNumberFormat('0.0');
  diario.getRange('E:E').setNumberFormat('0');
  [planA, planB].filter(Boolean).forEach(sheet => sheet.setFrozenRows(1));

  return 'Configuración verificada. ' +
    (cambios.length ? cambios.join(' ') : 'No se detectaron correcciones pendientes.') +
    ' Las copias de seguridad existentes se conservaron.';
}

/**
 * Devuelve la biblioteca organizada si existe;
 * de lo contrario, utiliza Plan A y Plan B.
 */
function getPlanOptions() {
  const ss = getWorkbook_();
  const library = ss.getSheetByName('Biblioteca');
  const options = [];

  if (library && library.getLastRow() >= 2) {
    const data = library.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const opcion = cleanCell_(row[1]);
      if (!opcion) continue;

      const workout = {
        id: cleanCell_(row[0]) || ('Biblioteca-' + i),
        fuente: cleanCell_(row[17]) || 'Biblioteca',
        dia: cleanCell_(row[5]) || 'Flexible',
        opcion: opcion,
        categoria: cleanCell_(row[2]) || 'Mixto / técnica',
        objetivo: cleanCell_(row[3]) || cleanCell_(row[2]),
        nivel: cleanCell_(row[4]) || 'Ajustar según experiencia',
        duracion: cleanCell_(row[6]) || 'No especificada',
        intervalos: cleanCell_(row[7]) || 'Según protocolo',
        esfuerzo: cleanCell_(row[8]) || 'Según protocolo',
        recuperacion: cleanCell_(row[9]) || 'Según protocolo',
        terreno: cleanCell_(row[10]) || 'Variable',
        zonas: cleanCell_(row[11]) || 'Ajustar por percepción de esfuerzo',
        calentamiento: cleanCell_(row[12]) || '',
        bloquePrincipal: cleanCell_(row[13]) || 'Sin protocolo definido',
        recuperacionBloques: cleanCell_(row[14]) || '',
        enfriamiento: cleanCell_(row[15]) || '',
        seguridad: cleanCell_(row[16]) || '',
        descripcion: cleanCell_(row[13]) || 'Sin descripción'
      };

      options.push(workout);
    }

    return options;
  }

  // Compatibilidad con la estructura anterior.
  ['Plan A', 'Plan B'].forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) return;

    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const opcion = cleanCell_(row[1]);
      if (!opcion) continue;

      const workout = {
        id: sheetName + '-' + i,
        fuente: sheetName,
        dia: cleanCell_(row[0]) || 'Flexible',
        opcion: opcion,
        intervalos: cleanCell_(row[2]) || 'No especificado',
        duracion: cleanCell_(row[3]) || 'No especificada',
        esfuerzo: cleanCell_(row[3]) || 'Según protocolo',
        recuperacion: cleanCell_(row[4]) || 'No especificada',
        terreno: cleanCell_(row[5]) || 'No especificado',
        descripcion: cleanCell_(row[6]) || 'Sin descripción',
        bloquePrincipal: cleanCell_(row[6]) || 'Sin descripción',
        zonas: cleanCell_(row[7]) || 'No especificadas'
      };

      workout.categoria = categorizeWorkout_(workout);
      workout.objetivo = workout.categoria;
      workout.nivel = 'Ajustar según experiencia';
      workout.calentamiento = '15–20 min progresivos.';
      workout.enfriamiento = '10–15 min suaves.';
      workout.recuperacionBloques = '';
      workout.seguridad =
        'Elige un tramo seguro y adapta la carga a tu experiencia y recuperación.';

      options.push(workout);
    }
  });

  return options;
}

function categorizeWorkout_(w) {
  const text = [w.opcion, w.descripcion, w.zonas].join(' ').toLowerCase();

  if (/explosiv|aceleración|aceleracion|arranque/.test(text)) return 'Explosividad';
  if (/fuerza|cadencia baja|atrancad|resistencia subida/.test(text)) return 'Fuerza';
  if (/sprint|velocidad|potencia corta|llano rápido|llano rapido|vo2|max|hiit|micro-burst|microburst|over-under|piramidal|umbral/.test(text)) {
    return 'Velocidad e intervalos';
  }
  if (/grupal|fondo|montaña|montana|resistencia|sustentación|sustentacion|larga|rodaje/.test(text)) {
    return 'Resistencia';
  }
  if (/tempo|ritmo/.test(text)) return 'Tempo';
  if (/mixto|repecho|cortos\+largos|técnica|tecnica/.test(text)) return 'Mixto / técnica';

  return 'Mixto / técnica';
}

/** Guarda una sesión en la pestaña existente Diario de Sensaciones. */
function saveSession(sessionData) {
  if (!sessionData || typeof sessionData !== 'object') {
    throw new Error('No se recibieron los datos de la sesión.');
  }

  const entrenamiento = String(sessionData.entrenamiento || '').trim();
  if (!entrenamiento) throw new Error('Selecciona un entrenamiento.');

  const rpeRaw = String(sessionData.rpe == null ? '' : sessionData.rpe).trim();
  const rpe = rpeRaw === '' ? '' : Number(rpeRaw);

  if (rpeRaw !== '' && (!Number.isFinite(rpe) || rpe < 1 || rpe > 10)) {
    throw new Error('El RPE debe ser un número entre 1 y 10.');
  }

  const distanciaRaw = String(sessionData.distancia == null ? '' : sessionData.distancia).trim();
  const distancia = distanciaRaw === '' ? '' : Number(distanciaRaw);

  if (distanciaRaw !== '' && (!Number.isFinite(distancia) || distancia < 0)) {
    throw new Error('La distancia debe ser un número igual o mayor que cero.');
  }

  const ss = getWorkbook_();
  const sheet = ss.getSheetByName('Diario de Sensaciones');

  if (!sheet) throw new Error('No se encontró la hoja "Diario de Sensaciones".');

  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);

  try {
    sheet.appendRow([
      new Date(),
      entrenamiento,
      String(sessionData.tiempoReal || '').trim(),
      distancia,
      rpe,
      String(sessionData.fatiga || '').trim(),
      String(sessionData.notas || '').trim(),
      String(sessionData.comentarios || '').trim()
    ]);
  } finally {
    lock.releaseLock();
  }

  return 'Sesión guardada con éxito.';
}

function backupSheetOnce_(ss, sourceSheet) {
  const backupName = 'BACKUP_' + sourceSheet.getName().replace(/[^A-Za-z0-9_]/g, '_');
  if (ss.getSheetByName(backupName)) return;
  sourceSheet.copyTo(ss).setName(backupName).hideSheet();
}

function cleanCell_(value) {
  if (value === null || value === undefined || value === '') return '';
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(value).trim();
}

/* ==========================================================
   API para Netlify
   Requiere la propiedad de script NETLIFY_API_TOKEN.
   ========================================================== */

function doPost(e) {
  try {
    const rawBody = e && e.postData && e.postData.contents
      ? e.postData.contents
      : '{}';

    const request = JSON.parse(rawBody);
    const configuredToken =
      PropertiesService.getScriptProperties().getProperty('NETLIFY_API_TOKEN');

    if (!configuredToken) {
      return jsonResponse_({
        ok: false,
        error: 'Falta configurar la propiedad NETLIFY_API_TOKEN en Apps Script.'
      });
    }

    if (
      !request.token ||
      !constantTimeTokenMatch_(String(request.token), String(configuredToken))
    ) {
      return jsonResponse_({ ok: false, error: 'No autorizado.' });
    }

    let data;

    if (request.action === 'getPlanOptions') {
      data = getPlanOptions();
    } else if (request.action === 'saveSession') {
      data = saveSession(request.payload);
    } else {
      return jsonResponse_({ ok: false, error: 'Acción no reconocida.' });
    }

    return jsonResponse_({ ok: true, data: data });
  } catch (error) {
    return jsonResponse_({
      ok: false,
      error: error && error.message
        ? error.message
        : 'Error interno al procesar la solicitud.'
    });
  }
}

/** Devuelve JSON para las llamadas servidor a servidor desde Netlify. */
function jsonResponse_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Compara el token recibido con el configurado. */
function constantTimeTokenMatch_(provided, expected) {
  if (provided.length !== expected.length) return false;

  let difference = 0;

  for (let i = 0; i < provided.length; i++) {
    difference |= provided.charCodeAt(i) ^ expected.charCodeAt(i);
  }

  return difference === 0;
}

/**
 * Usa el archivo original de Aguriza Training.
 * El ID corresponde a la hoja de cálculo existente.
 */
function getWorkbook_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;

  return SpreadsheetApp.openById(
    '1l-P1tfx7B65N301oFGbGnCjIUbTHJvNI3InaWUR3UUc'
  );
}
