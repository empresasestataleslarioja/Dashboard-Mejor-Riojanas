// ══════════════════════════════════════════════════════════════════
//  lr_bridge_v5.gs — Google Apps Script
//  Dashboard BI La Rioja v6
//  Cambios vs v5 + lock (diagnóstico flujos):
//    · upsertFlujos ahora devuelve {dedupCount, existentesPrevios,
//      filasEscritas, motivoNoEscritura} en vez de nada. doPost vuelca eso
//      en resultados.flujos_* — antes el log en 'OK' solo mostraba el
//      tamaño del array recibido, no si la hoja realmente se escribió.
//  Cambios vs v5 original (fix conflictos):
//    · doPost: LockService.getScriptLock() serializa las escrituras.
//      Sin esto, dos publicaciones cercanas en el tiempo (p.ej. guardar dos
//      flujos seguidos) podían solaparse: la segunda leía la hoja antes de
//      que la primera terminara de escribir, y al guardar pisaba por
//      completo los cambios de la primera — altas y bajas se perdían en
//      silencio aunque ambos requests devolvieran ok:true.
//  Cambios vs v4:
//    · getFile: leer archivo de Drive por ID (para importar Excel desde Drive)
//  Cambios vs v3:
//    · Soporte hojas nuevas: mensual, er_mensual, er_detalle,
//      flujos, flujos_config
//    · doGet devuelve las 5 hojas nuevas
//    · doPost escribe las 5 hojas nuevas
//    · flujos: deduplicación por empresa+periodo+tipo
//    · Auditoría ampliada con 5 columnas nuevas
// ══════════════════════════════════════════════════════════════════

const SHEET_ID    = '1WCs8iiEnGsa641l2vd5yTy2mczM96-KLVbrscD_Rw40';

const TAB = {
  fact:           'fact',
  resultado:      'resultado',
  personal:       'personal',
  transferencias: 'transferencias',
  notas:          'notas',
  presupuesto:    'presupuesto',
  balances:       'balances',
  balances_rubros:'balances_rubros',
  rubros:         'rubros',
  rubros_lista:   'rubros_lista',
  mensual:        'mensual',
  er_mensual:     'er_mensual',
  er_detalle:     'er_detalle',
  flujos:         'flujos',
  flujos_config:  'flujos_config',
  meta:           'meta',
  auditoria:      'auditoria',
  log:            '_log',
};

// ── GET ────────────────────────────────────────────────────────────
function doGet(e) {
  try {
  if (!e || !e.parameter) return jsonResp({ ok: false, error: 'Sin parámetros' });
  const tab    = (e.parameter.tab    || 'all').toLowerCase();
  const action = (e.parameter.action || '').toLowerCase();
  const ss     = SpreadsheetApp.openById(SHEET_ID);

  // ── Leer archivo de Drive ─────────────────────────────────────────
  if (action === 'getfile') {
    const fileId = e.parameter.fileId || '';
    if (!fileId) return jsonResp({ ok: false, error: 'fileId requerido' });
    try {
      const file  = DriveApp.getFileById(fileId);
      const blob  = file.getBlob();
      const bytes = blob.getBytes();
      const b64   = Utilities.base64Encode(bytes);
      return jsonResp({ ok: true, filename: file.getName(), mimeType: blob.getContentType(), data: b64 });
    } catch(err) {
      return jsonResp({ ok: false, error: err.message });
    }
  }

  if (tab === 'meta') return jsonResp({ meta: leerMeta(ss) });
  return jsonResp({
    fact:            leerHoja(ss, TAB.fact),
    resultado:       leerHoja(ss, TAB.resultado),
    personal:        leerHoja(ss, TAB.personal),
    transferencias:  leerHoja(ss, TAB.transferencias),
    notas:           leerHoja(ss, TAB.notas),
    presupuesto:     leerHoja(ss, TAB.presupuesto),
    balances:        leerHoja(ss, TAB.balances),
    balances_rubros: leerHoja(ss, TAB.balances_rubros),
    rubros:          leerHoja(ss, TAB.rubros),
    rubros_lista:    leerHoja(ss, TAB.rubros_lista),
    mensual:         leerHoja(ss, TAB.mensual),
    er_mensual:      leerHoja(ss, TAB.er_mensual),
    er_detalle:      leerHoja(ss, TAB.er_detalle),
    flujos:          leerHoja(ss, TAB.flujos),
    flujos_config:   leerHoja(ss, TAB.flujos_config),
    meta:            leerMeta(ss),
  });
  } catch(err) {
    return jsonResp({ ok: false, error: 'doGet: ' + err.message });
  }
}

// ── POST ───────────────────────────────────────────────────────────
function doPost(e) {
  const ss      = SpreadsheetApp.openById(SHEET_ID);
  const savedAt = new Date().toISOString();

  // Serializar escrituras: cada request hace leer-modificar-escribir sobre las
  // mismas hojas (clearContents + setValues). Sin lock, dos publicaciones
  // cercanas en el tiempo (p.ej. guardar dos flujos seguidos, cada uno dispara
  // su propio syncToCloud) pueden solaparse: la segunda lee la hoja antes de
  // que la primera termine de escribir, y al guardar pisa por completo los
  // cambios de la primera — altas y bajas se pierden en silencio aunque
  // ambos requests devuelvan ok:true.
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (lockErr) {
    escribirLog(ss, savedAt, 'ERROR_LOCK', 'No se pudo obtener el lock en 30s: ' + lockErr.message, '');
    return jsonResp({ ok: false, error: 'Servidor ocupado procesando otra publicación — reintentá en unos segundos' });
  }

  try {
    let body;
    const rawContents = (e.postData && e.postData.contents) ? e.postData.contents.trim() : '';
    const rawParam    = (e.parameter && e.parameter.json)   ? e.parameter.json.trim()    : '';
    const rawSrc      = rawParam || rawContents;

    if (!rawSrc) {
      escribirLog(ss, savedAt, 'ERROR_VACIO', 'Sin datos recibidos', JSON.stringify(e.parameter||{}));
      return jsonResp({ ok: false, error: 'Sin datos en el request' });
    }

    try { body = JSON.parse(rawSrc); }
    catch(parseErr) {
      escribirLog(ss, savedAt, 'ERROR_PARSE', parseErr.message, rawSrc.slice(0, 300));
      return jsonResp({ ok: false, error: 'JSON inválido: ' + parseErr.message });
    }

    const payload = body.data  || {};
    const updBy   = body.updatedBy || 'usuario';
    const nota    = body.nota || '';


    const resultados = {};

    // ── Secciones estándar (reemplazo total) ──────────────────────
    ['fact','resultado','personal','notas','presupuesto'].forEach(sec => {
      if (Array.isArray(payload[sec])) {
        escribirHoja(ss, TAB[sec], payload[sec]);
        resultados[sec] = payload[sec].length;
      }
    });

    // ── TRANSFERENCIAS ────────────────────────────────────────────
    // Reemplazo total con deduplicación previa en el payload
    // El dashboard ya envía datos deduplicados — el bridge solo escribe
    if (Array.isArray(payload.transferencias)) {
      const seen = new Set();
      const dedup = payload.transferencias.filter(r => {
        const k = [String(r.beneficiario||'').trim(), String(r.año||''), String(r.mes||''), String(Math.round(parseFloat(r.importe||0)))].join('|');
        if (seen.has(k)) return false;
        seen.add(k); return true;
      });
      if (payload._trans_modo === 'reemplazar') {
        escribirHoja(ss, TAB.transferencias, dedup);
      } else {
        upsertTransferencias(ss, dedup);
      }
      resultados.transferencias = dedup.length;
      if (dedup.length < payload.transferencias.length) {
        resultados.transferencias_dedup = payload.transferencias.length - dedup.length;
      }
    }

    // ── BALANCES ──────────────────────────────────────────────────
    // Reemplazo total con deduplicación previa (empresa+periodo)
    if (Array.isArray(payload.balances)) {
      const seen = new Set();
      const dedup = payload.balances.filter(r => {
        const k = (r.empresa || '') + '_' + (r.periodo || '');
        if (seen.has(k)) return false;
        seen.add(k); return true;
      });
      reemplazarHojaConHeaders(ss, TAB.balances, dedup);
      resultados.balances = dedup.length;
    } else {
      resultados.balances_status = 'NO_RECIBIDO';
    }

    // ── BALANCES_RUBROS ───────────────────────────────────────────
    if (Array.isArray(payload.balances_rubros)) {
      escribirHoja(ss, TAB.balances_rubros, payload.balances_rubros);
      resultados.balances_rubros = payload.balances_rubros.length;
    }

    // ── RUBROS ────────────────────────────────────────────────────
    // Merge inteligente: si el dashboard envía rubros, reemplazar
    // Si el payload es vacío, NO borrar los existentes en Sheets
    if (Array.isArray(payload.rubros) && payload.rubros.length > 0) {
      // Deduplicar por empresa
      const seen = new Set();
      const dedup = payload.rubros.filter(r => {
        if (seen.has(r.empresa)) return false;
        seen.add(r.empresa); return true;
      });
      escribirHoja(ss, TAB.rubros, dedup);
      resultados.rubros = dedup.length;
    }

    // ── RUBROS_LISTA ──────────────────────────────────────────────
    if (Array.isArray(payload.rubros_lista) && payload.rubros_lista.length > 0) {
      escribirHoja(ss, TAB.rubros_lista, payload.rubros_lista);
      resultados.rubros_lista = payload.rubros_lista.length;
    }

    // ── MENSUAL (fact + resultado mensual) ─────────────────────────────
    if (Array.isArray(payload.mensual)) {
      escribirHoja(ss, TAB.mensual, payload.mensual);
      resultados.mensual = payload.mensual.length;
    }

    // ── ER_MENSUAL (estados de resultados provisorios) ────────────────
    if (Array.isArray(payload.er_mensual)) {
      escribirHoja(ss, TAB.er_mensual, payload.er_mensual);
      resultados.er_mensual = payload.er_mensual.length;
    }

    // ── ER_DETALLE (partidas Excel ER formal) ───────────────────────
    if (Array.isArray(payload.er_detalle)) {
      escribirHoja(ss, TAB.er_detalle, payload.er_detalle);
      resultados.er_detalle = payload.er_detalle.length;
    }

    // ── FLUJOS ────────────────────────────────────────────────────
    if (Array.isArray(payload.flujos) || Array.isArray(payload.flujos_a_eliminar)) {
      const flDiag = upsertFlujos(ss, payload.flujos || [], payload.flujos_a_eliminar || []);
      // flujos_recibidos: tamaño del array que llegó en el payload (antes del dedup).
      // flujos_escritos: filas realmente grabadas en la hoja tras el merge — esta es
      // la que hay que mirar para saber si el guardado surtió efecto de verdad.
      resultados.flujos_recibidos  = (payload.flujos || []).length;
      resultados.flujos_dedup      = flDiag.dedupCount;
      resultados.flujos_existentes_previos = flDiag.existentesPrevios;
      resultados.flujos_escritos   = flDiag.filasEscritas;
      resultados.flujos_eliminados = (payload.flujos_a_eliminar || []).length;
      if (flDiag.motivoNoEscritura) resultados.flujos_motivo_no_escritura = flDiag.motivoNoEscritura;
    }

    // ── FLUJOS_CONFIG ────────────────────────────────────────────
    if (Array.isArray(payload.flujos_config) && payload.flujos_config.length > 0) {
      escribirHoja(ss, TAB.flujos_config, payload.flujos_config);
      resultados.flujos_config = payload.flujos_config.length;
    }

    escribirMeta(ss, savedAt, updBy, nota);
    registrarAuditoria(ss, savedAt, updBy, nota, resultados);
    escribirLog(ss, savedAt, 'OK', JSON.stringify(resultados), '');

    return jsonResp({ ok: true, savedAt, resultados });

  } catch(err) {
    escribirLog(ss, savedAt, 'ERROR', err.message, err.stack ? err.stack.slice(0,300) : '');
    return jsonResp({ ok: false, error: err.message });
  } finally {
    lock.releaseLock();
  }
}

// ── Reemplazar hoja con headers explícitos (balances) ─────────────
function reemplazarHojaConHeaders(ss, nombre, rows) {
  let hoja = ss.getSheetByName(nombre);
  if (!hoja) {
    if (!rows || rows.length === 0) return;
    hoja = ss.insertSheet(nombre);
  }
  const maxRow = hoja.getMaxRows();
  const maxCol = hoja.getMaxColumns();
  if (maxRow > 0 && maxCol > 0) {
    hoja.getRange(1, 1, maxRow, maxCol).clearContent();
  }
  if (!rows || rows.length === 0) return;
  const headers = Object.keys(rows[0]);
  const matrix  = [
    headers,
    ...rows.map(r => headers.map(h => {
      const v = r[h];
      return (v === null || v === undefined) ? '' : v;
    }))
  ];
  hoja.getRange(1, 1, matrix.length, headers.length).setValues(matrix);
}

// ── Escribir hoja genérica con reemplazo completo ─────────────────
function escribirHoja(ss, nombre, rows) {
  let hoja = ss.getSheetByName(nombre);
  if (!hoja) hoja = ss.insertSheet(nombre);
  hoja.clearContents();
  if (!rows || rows.length === 0) return;
  const headers = Object.keys(rows[0]);
  const matrix  = [
    headers,
    ...rows.map(r => headers.map(h => {
      const v = r[h];
      return (v === null || v === undefined) ? '' : v;
    }))
  ];
  hoja.getRange(1, 1, matrix.length, headers.length).setValues(matrix);
}

// ── Leer hoja → array de objetos ──────────────────────────────────
function leerHoja(ss, nombre) {
  const hoja = ss.getSheetByName(nombre);
  if (!hoja) return [];
  const lastRow = hoja.getLastRow();
  const lastCol = hoja.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];
  const datos   = hoja.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = datos[0].map(String);
  return datos.slice(1)
    .filter(row => row.some(c => c !== '' && c !== null && c !== undefined))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i] !== undefined ? row[i] : ''; });
      return obj;
    });
}

// ── Meta ──────────────────────────────────────────────────────────
function leerMeta(ss) {
  const hoja = ss.getSheetByName(TAB.meta);
  if (!hoja || hoja.getLastRow() < 2) return {};
  const vals = hoja.getRange(2, 1, 1, 3).getValues()[0];
  return { savedAt: vals[0] || '', updatedBy: vals[1] || '', nota: vals[2] || '' };
}

function escribirMeta(ss, savedAt, updBy, nota) {
  let hoja = ss.getSheetByName(TAB.meta);
  if (!hoja) hoja = ss.insertSheet(TAB.meta);
  hoja.clearContents();
  hoja.getRange(1, 1, 2, 3).setValues([
    ['savedAt','updatedBy','nota'],
    [savedAt,  updBy,      nota ]
  ]);
}

// ── Log ───────────────────────────────────────────────────────────
function escribirLog(ss, fecha, estado, mensaje, detalle) {
  try {
    let hoja = ss.getSheetByName(TAB.log);
    if (!hoja) {
      hoja = ss.insertSheet(TAB.log);
      hoja.appendRow(['fecha','estado','mensaje','detalle']);
    }
    hoja.appendRow([fecha, estado, String(mensaje).slice(0,500), String(detalle).slice(0,500)]);
    const last = hoja.getLastRow();
    if (last > 101) hoja.deleteRows(2, last - 101);
  } catch(e) {}
}

// ── Auditoría ─────────────────────────────────────────────────────
function registrarAuditoria(ss, savedAt, updBy, nota, resultados) {
  try {
    let hoja = ss.getSheetByName(TAB.auditoria);
    if (!hoja) {
      hoja = ss.insertSheet(TAB.auditoria);
      hoja.appendRow(['fecha','usuario','nota','fact','resultado','personal',
                      'transferencias','notas','presupuesto','balances','rubros',
                      'mensual','er_mensual','er_detalle','flujos','flujos_config']);
    }
    hoja.appendRow([
      savedAt, updBy, nota,
      resultados.fact           || 0,
      resultados.resultado      || 0,
      resultados.personal       || 0,
      resultados.transferencias || 0,
      resultados.notas          || 0,
      resultados.presupuesto    || 0,
      resultados.balances       !== undefined ? resultados.balances : (resultados.balances_status || '?'),
      resultados.rubros         || 0,
      resultados.mensual        || 0,
      resultados.er_mensual     || 0,
      resultados.er_detalle     || 0,
      resultados.flujos         || 0,
      resultados.flujos_config  || 0,
    ]);
    const last = hoja.getLastRow();
    if (last > 201) hoja.deleteRows(2, last - 201);
  } catch(e) {}
}

// ── JSON response ─────────────────────────────────────────────────
function upsertTransferencias(ss, nuevas) {
  if (!nuevas || nuevas.length === 0) return;
  let hoja = ss.getSheetByName(TAB.transferencias);
  if (!hoja) { hoja = ss.insertSheet(TAB.transferencias); }
  const lastRow = hoja.getLastRow();
  const lastCol = hoja.getLastColumn();
  let existentes = {};
  let headers = [];
  if (lastRow >= 2 && lastCol >= 1) {
    const datos = hoja.getRange(1, 1, lastRow, lastCol).getValues();
    headers = datos[0].map(String);
    const bi = headers.indexOf('beneficiario'), ai = headers.indexOf('año'),
          mi = headers.indexOf('mes'), ii = headers.indexOf('importe');
    datos.slice(1).forEach(function(row) {
      if (!row.some(function(c){return c!=='';})) return;
      const k = [row[bi]||'',row[ai]||'',row[mi]||'',row[ii]||''].join('|');
      if (k !== '|||' && !existentes[k]) {
        const obj = {}; headers.forEach(function(h,i){obj[h]=row[i]!==undefined?row[i]:'';}); existentes[k]=obj;
      }
    });
  }
  if (headers.length === 0) headers = Object.keys(nuevas[0]);
  nuevas.forEach(function(r) {
    const k = [r.beneficiario||'',r.año||'',r.mes||'',r.importe||''].join('|');
    if (k !== '|||') existentes[k] = r;
  });
  hoja.clearContents();
  const rows = Object.values(existentes);
  if (rows.length === 0) return;
  const matrix = [headers, ...rows.map(function(r){return headers.map(function(h){const v=r[h];return(v===null||v===undefined)?'':v;});})];
  hoja.getRange(1,1,matrix.length,headers.length).setValues(matrix);
}

function upsertFlujos(ss, nuevos, aEliminar) {
  // Devuelve siempre un diagnóstico real de lo que hizo — antes la función no
  // devolvía nada y doPost sólo logueaba el tamaño del payload de entrada,
  // así que un log en 'OK' no probaba que se hubiera escrito una sola fila.
  const diag = { dedupCount: 0, existentesPrevios: 0, filasEscritas: 0, motivoNoEscritura: '' };
  const elimSet = new Set(aEliminar || []);
  let hoja = ss.getSheetByName(TAB.flujos);
  if (!hoja) hoja = ss.insertSheet(TAB.flujos);
  const lastRow = hoja.getLastRow();
  const lastCol = hoja.getLastColumn();
  let existentes = {};
  let headers = [];
  if (lastRow >= 2 && lastCol >= 1) {
    const datos = hoja.getRange(1, 1, lastRow, lastCol).getValues();
    headers = datos[0].map(String);
    const claveIdx = headers.indexOf('clave');
    datos.slice(1).forEach((row, i) => {
      if (row.some(c => c !== '')) {
        const clave = claveIdx >= 0 ? String(row[claveIdx]) : '';
        if (clave) existentes[clave] = { row: i + 2, data: row };
      }
    });
  }
  diag.existentesPrevios = Object.keys(existentes).length;
  const seen = new Set();
  const dedupNuevos = (nuevos || []).filter(r => {
    const k = (r.empresa||'') + '_' + (r.periodo||'') + '_' + (r.tipo||'');
    if (seen.has(k)) return false; seen.add(k); return true;
  });
  diag.dedupCount = dedupNuevos.length;
  if (dedupNuevos.length === 0 && elimSet.size === 0) {
    diag.motivoNoEscritura = 'sin_datos_ni_eliminaciones';
    return diag;
  }
  if (headers.length === 0 && dedupNuevos.length > 0) {
    headers = Object.keys(dedupNuevos[0]);
  }
  if (headers.length === 0) {
    diag.motivoNoEscritura = 'sin_headers';
    return diag;
  }
  const finalMap = {};
  Object.entries(existentes).forEach(([clave, entry]) => {
    if (!elimSet.has(clave)) {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = entry.data[i] !== undefined ? entry.data[i] : ''; });
      finalMap[clave] = obj;
    }
  });
  dedupNuevos.forEach(r => {
    const k = r.clave || ((r.empresa||'') + '_' + (r.periodo||'') + '_' + (r.tipo||''));
    if (!elimSet.has(k)) finalMap[k] = r;
  });
  hoja.clearContents();
  const rows = Object.values(finalMap);
  if (rows.length === 0) {
    diag.motivoNoEscritura = 'finalMap_vacio';
    return diag;
  }
  const allHeaders = headers.length > 0 ? headers : Object.keys(rows[0]);
  const matrix = [allHeaders, ...rows.map(r => allHeaders.map(h => { const v = r[h]; return (v === null || v === undefined) ? '' : v; }))];
  hoja.getRange(1, 1, matrix.length, allHeaders.length).setValues(matrix);
  diag.filasEscritas = rows.length;
  return diag;
}

function jsonResp(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
