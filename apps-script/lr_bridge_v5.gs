// ══════════════════════════════════════════════════════════════════
//  lr_bridge_v5.gs — Google Apps Script
//  Dashboard BI La Rioja v6
//  Cambios vs v5.9:
//    · Nueva sección "inflacion": la inflación mensual cargada desde el
//      dashboard (solo el año en curso es editable ahí) no viajaba a
//      Sheets bajo NINGUNA circunstancia — ni con "Publicar Todo" — porque
//      el payload nunca incluía esta sección y el bridge no tenía ninguna
//      hoja ni lectura para ella. Se agrega TAB.inflacion, su lectura en
//      doGet y su escritura (reemplazo total simple — es una sola fila
//      por año, sin el riesgo multi-dispositivo que justifica upsert en
//      otras hojas) en doPost.
//  Cambios vs v5.9 (anterior):
//    · upsertRubros: headers ya no quedaba fijo en lo que la hoja tenía
//      ANTES del merge (o el fallback hardcodeado) — un campo nuevo que
//      llegara en `nuevas` pero no existiera todavía como columna (ej.
//      "cuit", "observaciones") se descartaba en silencio al escribir,
//      aunque el objeto fusionado en memoria sí lo tuviera. Ahora,
//      después de fusionar, headers se recalcula como la unión de TODAS
//      las claves presentes en `rows` — mismo criterio que ya usa
//      escribirHoja para hojas de reemplazo total, aplicado acá donde
//      faltaba por ser upsert. Caso real reportado: "cuit"/"observaciones"
//      nunca llegaban a la pestaña "rubros" pese a que el dashboard los
//      mandaba bien y el servidor respondía éxito.
//  Cambios vs v5.8:
//    · Nueva hoja empresas_eliminadas_excepciones: protecciones puntuales
//      entre un nombre actual y una tumba histórica no relacionada (ver
//      _agregarExcepcionTumba) — a diferencia de empresas_eliminadas_
//      restauradas (que anula la tumba ENTERA), esto deja la tumba
//      intacta para el resto de los casos y solo excluye un nombre
//      puntual. Caso real: "FEDERAL SE" quedó eliminada por una fusión
//      real con "FEDERAL EE" (tumba correcta), pero otra empresa sin
//      relación renombrada a "FEDERAL" (sin sufijo) matcheaba igual
//      contra "FEDERAL SE" porque el comparador de nombres ignora a
//      propósito el sufijo legal "SE".
//  Cambios vs v5.7:
//    · Nueva hoja empresas_eliminadas_restauradas: tumbas de
//      empresas_eliminadas deshechas a mano desde el dashboard (ver
//      _restaurarEmpresaEliminada) — hasta ahora no había forma de anular
//      una tumba: empresas_eliminadas es upsert por unión (nunca se pierde
//      una entrada, a propósito, para sobrevivir multi-dispositivo), así
//      que una empresa histórica que vuelve a existir con un nombre que
//      matchea por variante de grafía (ej. "FEDERAL" vs. la tumba
//      "FEDERAL SE") quedaba bloqueada para siempre. upsertEmpresasEliminadas
//      ahora recibe ese set y excluye esas filas antes de escribir.
//  Cambios vs v5.6 (nota: el log interno de versiones saltea algunos
//  números — se preserva tal cual, sin renumerar):
//    · upsertEmpresasEliminadas ahora lee y escribe fusionada_en — antes
//      lo descartaba en cada publicación (solo guardaba empresa/fecha), así
//      que la hoja empresas_eliminadas nunca llegó a tener esa columna por
//      más fusiones que se hicieran desde el dashboard. Sin fusionada_en,
//      _coincideEmpresaTumba no puede proteger a la empresa sobreviviente
//      de una fusión (ver excluir ahí mismo) — causa real de que
//      sobrevivientes de fusiones viejas (ej. VALLESOL, PUERTAS DEL SOL,
//      VIVERO DEL OESTE RIOJANO) se trataran como tumbadas por variante de
//      grafía con su propio nombre previo ya eliminado.
//  Cambios vs v5.6:
//    · rubros (la lista maestra de empresas, Gestión de Empresas y
//      Rubros) pasa de reemplazo total a merge por upsert (empresa/_ts,
//      ver upsertRubros) — misma causa que el fix de autoridades de
//      v5.5, reportada ahora como "empresas activas desaparecen y se
//      desactualizan" después de sincronizar entre varios dispositivos.
//  Cambios vs v5.5:
//    · autoridades pasa de reemplazo total a merge por upsert (_id/_ts,
//      ver upsertAutoridades) y el servidor ahora procesa
//      autoridades_a_eliminar al escribir (antes solo se usaba del lado
//      del cliente) — reportado como "modifican autoridades desde otro
//      dispositivo, publican en Sheets, pero el cambio no se sostiene".
//  Cambios vs v5.4:
//    · Protección multi-dispositivo, a raíz de un caso real: una carga de
//      ER mensual hecha desde un dispositivo fue pisada por el auto-sync
//      de otro que todavía tenía la versión anterior en memoria.
//      - er_mensual pasa de reemplazo total a merge por upsert
//        (empresa+año, ver upsertErMensual) — un dispositivo con una
//        copia local vieja ya no puede borrar el ER mensual de una
//        empresa/año que él mismo no tocó.
//      - Respaldo automático (_backupHojaAntesDeEscribir) antes de CUALQUIER
//        reemplazo total de hoja (escribirHoja/reemplazarHojaConHeaders/
//        upsertFlujos/upsertTransferencias): guarda el contenido anterior
//        en una hoja "_bak_<nombre>" (un respaldo rotativo, no historial
//        completo) para poder deshacer manualmente la última sobrescritura
//        si hizo falta.
//  Cambios vs v5.3:
//    · Nueva hoja beneficiarios_excluidos: terceros de Transferencias a
//      ignorar para siempre (merge por upsert). doPost filtra cualquier
//      fila de transferencias cuyo beneficiario esté excluido, ANTES de
//      escribirla — misma protección que empresas_eliminadas pero para
//      beneficiarios que no son empresas del portafolio.
//  Cambios vs v5.2:
//    · Nueva hoja empresas_eliminadas: "tumbas" de empresas borradas,
//      fusionadas o renombradas (merge por upsert, nunca se pierde una
//      entrada). doPost filtra cualquier fila de fact/resultado/personal/
//      autoridades/notas/presupuesto/balances/balances_rubros/rubros/
//      mensual/er_mensual/er_detalle/flujos que nombre a una empresa ya
//      eliminada, ANTES de escribirla — así una pestaña vieja que todavía
//      tiene esa empresa cargada no puede resucitarla al sincronizar.
//  Cambios vs v5.1:
//    · escribirHoja/reemplazarHojaConHeaders ya NO borran una pestaña
//      cuando el payload llega vacío (antes hoja.clearContents() corría
//      ANTES de chequear si había filas — un envío vacío, por el motivo
//      que fuera, borraba los datos existentes en Sheets igual)
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
  autoridades:    'autoridades',
  transferencias: 'transferencias',
  notas:          'notas',
  presupuesto:    'presupuesto',
  inflacion:      'inflacion',
  balances:       'balances',
  balances_rubros:'balances_rubros',
  rubros:         'rubros',
  rubros_lista:   'rubros_lista',
  mensual:        'mensual',
  er_mensual:     'er_mensual',
  er_detalle:     'er_detalle',
  flujos:         'flujos',
  flujos_config:  'flujos_config',
  empresas_eliminadas: 'empresas_eliminadas',
  empresas_eliminadas_restauradas: 'empresas_eliminadas_restauradas',
  empresas_eliminadas_excepciones: 'empresas_eliminadas_excepciones',
  beneficiarios_excluidos: 'beneficiarios_excluidos',
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
    autoridades:     leerHoja(ss, TAB.autoridades),
    transferencias:  leerHoja(ss, TAB.transferencias),
    notas:           leerHoja(ss, TAB.notas),
    presupuesto:     leerHoja(ss, TAB.presupuesto),
    inflacion:       leerHoja(ss, TAB.inflacion),
    balances:        leerHoja(ss, TAB.balances),
    balances_rubros: leerHoja(ss, TAB.balances_rubros),
    rubros:          leerHoja(ss, TAB.rubros),
    rubros_lista:    leerHoja(ss, TAB.rubros_lista),
    mensual:         leerHoja(ss, TAB.mensual),
    er_mensual:      leerHoja(ss, TAB.er_mensual),
    er_detalle:      leerHoja(ss, TAB.er_detalle),
    flujos:          leerHoja(ss, TAB.flujos),
    flujos_config:   leerHoja(ss, TAB.flujos_config),
    empresas_eliminadas: leerHoja(ss, TAB.empresas_eliminadas),
    empresas_eliminadas_restauradas: leerHoja(ss, TAB.empresas_eliminadas_restauradas),
    empresas_eliminadas_excepciones: leerHoja(ss, TAB.empresas_eliminadas_excepciones),
    beneficiarios_excluidos: leerHoja(ss, TAB.beneficiarios_excluidos),
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

    // ── EMPRESAS_ELIMINADAS ("tumbas") ─────────────────────────────
    // Merge por upsert (nunca se pierde una entrada aunque este envío
    // venga de una pestaña vieja que no las conoce todas) ANTES de
    // procesar el resto de las secciones, para poder filtrarlas de acá
    // en más — es la protección real contra una pestaña desactualizada
    // que todavía tenga cargada una empresa ya borrada/fusionada y la
    // vuelva a publicar sin saberlo.
    const restauradas = upsertEmpresasRestauradas(ss, payload.empresas_eliminadas_restauradas || []);
    const tumbas = upsertEmpresasEliminadas(ss, payload.empresas_eliminadas || [], restauradas);
    resultados.empresas_eliminadas = tumbas.size;
    resultados.empresas_eliminadas_restauradas = restauradas.size;
    resultados.empresas_eliminadas_excepciones = upsertEmpresasExcepciones(ss, payload.empresas_eliminadas_excepciones || []);

    // ── BENEFICIARIOS_EXCLUIDOS ─────────────────────────────────────
    // Mismo criterio de merge por upsert que las tumbas de empresas —
    // nunca se pierde una exclusión, y se usa para filtrar transferencias
    // ANTES de escribirlas.
    const benExcluidos = upsertBeneficiariosExcluidos(ss, payload.beneficiarios_excluidos || []);
    resultados.beneficiarios_excluidos = benExcluidos.size;

    // ── Secciones estándar (reemplazo total) ──────────────────────
    ['fact','resultado','personal','notas','presupuesto'].forEach(sec => {
      if (Array.isArray(payload[sec])) {
        const filtradas = filtrarEliminadas(payload[sec], tumbas);
        escribirHoja(ss, TAB[sec], filtradas);
        resultados[sec] = filtradas.length;
      }
    });

    // ── INFLACIÓN ──────────────────────────────────────────────────
    // Reemplazo total simple (no upsert): es una sola fila por año,
    // publicada siempre entera desde el dashboard — no tiene el riesgo
    // multi-dispositivo de "pisar una edición de otro dispositivo que
    // todavía no bajé" que justifica upsert en autoridades/rubros/
    // er_mensual (ahí cada dispositivo edita registros distintos a la
    // vez; acá cada publicación ya trae el año completo tal cual quedó).
    if (Array.isArray(payload.inflacion)) {
      escribirHoja(ss, TAB.inflacion, payload.inflacion);
      resultados.inflacion = payload.inflacion.length;
    }

    // ── AUTORIDADES ─────────────────────────────────
    // Merge por upsert (_id/_ts), NO reemplazo total como el resto de
    // las secciones de arriba (ver upsertAutoridades) — misma clase de
    // bug ya corregida para ER mensual/Flujos/Transferencias: con
    // reemplazo total, un dispositivo con una copia local vieja (por
    // ejemplo con un cambio pendiente en otra sección, que dispara un
    // publish de TODO el payload) borraba al publicar una edición de
    // autoridades hecha por otro dispositivo que ya había llegado a
    // Sheets — reportado como "modifican autoridades desde otro
    // dispositivo, publican, pero el cambio no se sostiene". También
    // procesa autoridades_a_eliminar (bajas), que antes el servidor
    // nunca leía al escribir.
    if (Array.isArray(payload.autoridades)) {
      const filtradasAut = filtrarEliminadas(payload.autoridades, tumbas);
      resultados.autoridades = upsertAutoridades(ss, filtradasAut, payload.autoridades_a_eliminar || [], tumbas);
    }

    // ── TRANSFERENCIAS ────────────────────────────────────────────
    // Reemplazo total con deduplicación previa en el payload
    // El dashboard ya envía datos deduplicados — el bridge solo escribe
    if (Array.isArray(payload.transferencias)) {
      const seen = new Set();
      const dedup = filtrarBeneficiariosExcluidos(payload.transferencias, benExcluidos).filter(r => {
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
      const dedup = filtrarEliminadas(payload.balances, tumbas).filter(r => {
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
      const filtradas = filtrarEliminadas(payload.balances_rubros, tumbas);
      escribirHoja(ss, TAB.balances_rubros, filtradas);
      resultados.balances_rubros = filtradas.length;
    }

    // ── RUBROS ──────────────────────────────────
    // Merge por upsert (empresa + _ts), NO reemplazo total — ver
    // upsertRubros. Misma clase de bug que autoridades (y ya corregida
    // antes para ER mensual/Flujos/Transferencias): con reemplazo total,
    // cualquier dispositivo con una copia local apenas desactualizada
    // (por ejemplo, sin la última empresa agregada o editada desde otro
    // dispositivo) borraba esas filas al publicar — "rubros" es la lista
    // maestra de empresas (Gestión de Empresas y Rubros), así que el
    // efecto visible era justamente empresas activas que desaparecían
    // o volvían a un estado viejo después de sincronizar entre varios
    // dispositivos.
    if (Array.isArray(payload.rubros)) {
      const seen = new Set();
      const dedup = filtrarEliminadas(payload.rubros, tumbas).filter(r => {
        if (seen.has(r.empresa)) return false;
        seen.add(r.empresa); return true;
      });
      resultados.rubros = upsertRubros(ss, dedup, tumbas);
    }

    // ── RUBROS_LISTA ──────────────────────────────────────────────
    if (Array.isArray(payload.rubros_lista) && payload.rubros_lista.length > 0) {
      escribirHoja(ss, TAB.rubros_lista, payload.rubros_lista);
      resultados.rubros_lista = payload.rubros_lista.length;
    }

    // ── MENSUAL (fact + resultado mensual) ─────────────────────────────
    if (Array.isArray(payload.mensual)) {
      const filtradas = filtrarEliminadas(payload.mensual, tumbas);
      escribirHoja(ss, TAB.mensual, filtradas);
      resultados.mensual = filtradas.length;
    }

    // ── ER_MENSUAL (estados de resultados provisorios) ────────────────
    // Merge por upsert (empresa+año), NO reemplazo total: un dispositivo
    // que publica su propia copia local (que puede no conocer una carga
    // reciente hecha desde otro dispositivo) no debe borrar el ER mensual
    // de una empresa/año que él mismo no tocó. Caso real: una carga de ER
    // mensual hecha desde un dispositivo fue pisada por el auto-sync de
    // otro que todavía tenía la versión anterior en memoria.
    if (Array.isArray(payload.er_mensual)) {
      const filtradas = filtrarEliminadas(payload.er_mensual, tumbas);
      const total = upsertErMensual(ss, filtradas, tumbas);
      resultados.er_mensual = total;
    }

    // ── ER_DETALLE (partidas Excel ER formal) ───────────────────────
    if (Array.isArray(payload.er_detalle)) {
      const filtradas = filtrarEliminadas(payload.er_detalle, tumbas);
      escribirHoja(ss, TAB.er_detalle, filtradas);
      resultados.er_detalle = filtradas.length;
    }

    // ── FLUJOS ────────────────────────────────────────────────────
    if (Array.isArray(payload.flujos) || Array.isArray(payload.flujos_a_eliminar)) {
      const flujosFiltrados = filtrarEliminadas(payload.flujos || [], tumbas);
      upsertFlujos(ss, flujosFiltrados, payload.flujos_a_eliminar || []);
      resultados.flujos = flujosFiltrados.length;
      resultados.flujos_eliminados = (payload.flujos_a_eliminar || []).length;
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
  }
}

// ── Respaldo automático antes de cualquier reemplazo total de hoja ─────
// Un solo respaldo rotativo por hoja (no un historial completo): alcanza
// para deshacer manualmente la ÚLTIMA sobrescritura si hizo falta (abrir
// "_bak_<hoja>" en el Spreadsheet y copiar de vuelta), sin que la planilla
// crezca sin límite. Se llama SIEMPRE antes de vaciar una hoja que se
// publica con reemplazo total — cuidado de la información barato y
// automático, independiente de si el reemplazo terminó pisando algo que
// no debía (ver también: er_mensual ahora usa upsert en vez de reemplazo
// total, que ataca la causa en vez de solo dar in "deshacer").
function _backupHojaAntesDeEscribir(ss, nombre) {
  try {
    const hoja = ss.getSheetByName(nombre);
    if (!hoja) return; // no existe todavía — nada que respaldar
    const lastRow = hoja.getLastRow(), lastCol = hoja.getLastColumn();
    if (lastRow < 1 || lastCol < 1) return; // hoja vacía — nada que respaldar
    const datos = hoja.getRange(1, 1, lastRow, lastCol).getValues();
    const nombreBak = '_bak_' + nombre;
    let hojaBak = ss.getSheetByName(nombreBak);
    if (!hojaBak) hojaBak = ss.insertSheet(nombreBak);
    hojaBak.clearContents();
    hojaBak.getRange(1, 1, datos.length, datos[0].length).setValues(datos);
    hojaBak.getRange(datos.length + 2, 1).setValue('Respaldo automático antes de sobrescribir — ' + new Date().toISOString());
  } catch(e) {
    // Un fallo al respaldar no debe impedir la publicación real.
  }
}

// ── Reemplazar hoja con headers explícitos (balances) ─────────────
function reemplazarHojaConHeaders(ss, nombre, rows) {
  // No borrar la hoja si no llegaron filas (ver misma nota en escribirHoja).
  if (!rows || rows.length === 0) return;
  let hoja = ss.getSheetByName(nombre);
  if (!hoja) {
    hoja = ss.insertSheet(nombre);
  }
  _backupHojaAntesDeEscribir(ss, nombre);
  const maxRow = hoja.getMaxRows();
  const maxCol = hoja.getMaxColumns();
  if (maxRow > 0 && maxCol > 0) {
    hoja.getRange(1, 1, maxRow, maxCol).clearContent();
  }
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
  // No borrar la hoja si no llegaron filas: un payload vacio no distingue
  // "se querian borrar todos los datos" de "el dashboard fallo al leer su
  // propio estado antes de publicar" -- vaciar la hoja en ese segundo caso
  // destruye datos reales sin que nadie lo haya pedido.
  if (!rows || rows.length === 0) return;
  _backupHojaAntesDeEscribir(ss, nombre);
  hoja.clearContents();
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

// ── Empresas eliminadas/fusionadas/renombradas ("tumbas") ──────────
// Merge por upsert: nunca se pierde una entrada, aunque este envío venga
// de una pestaña que solo conoce una parte de las tumbas. Devuelve un Set
// con TODOS los nombres tumba (existentes + nuevas), para filtrar el
// resto de las secciones de este mismo request.
// `restauradas`: Set de nombres (ver upsertEmpresasRestauradas) que el
// dashboard deshizo explícitamente — se excluyen antes de escribir, para
// que la hoja deje de tener esa fila y de reenviarla en la próxima
// descarga. Sin esto, empresas_eliminadas (upsert por unión, a propósito
// para nunca perder una entrada) no tenía forma de anular una tumba.
function upsertEmpresasEliminadas(ss, nuevas, restauradas) {
  let hoja = ss.getSheetByName(TAB.empresas_eliminadas);
  if (!hoja) hoja = ss.insertSheet(TAB.empresas_eliminadas);
  const lastRow = hoja.getLastRow();
  const lastCol = hoja.getLastColumn();
  const existentes = {};
  let headers = ['empresa', 'fecha', 'fusionada_en'];
  if (lastRow >= 2 && lastCol >= 1) {
    const datos = hoja.getRange(1, 1, lastRow, lastCol).getValues();
    headers = datos[0].map(String);
    if (headers.indexOf('fusionada_en') === -1) headers.push('fusionada_en'); // columna agregada después de crearse la hoja — hojas viejas no la tenían
    const ei = headers.indexOf('empresa'), fi = headers.indexOf('fecha'), fei = headers.indexOf('fusionada_en');
    datos.slice(1).forEach(row => {
      const emp = ei >= 0 ? String(row[ei] || '').trim() : '';
      if (emp) existentes[emp] = { empresa: emp, fecha: fi >= 0 ? row[fi] : '', fusionada_en: fei >= 0 ? (row[fei] || '') : '' };
    });
  }
  (nuevas || []).forEach(r => {
    const emp = String(r.empresa || '').trim();
    if (!emp) return;
    const fusionadaEn = r.fusionada_en || r.fusionadaEn || '';
    if (!existentes[emp]) {
      existentes[emp] = { empresa: emp, fecha: r.fecha || '', fusionada_en: fusionadaEn };
    } else if (fusionadaEn && !existentes[emp].fusionada_en) {
      existentes[emp].fusionada_en = fusionadaEn; // completar el destino si faltaba
    }
  });
  (restauradas || new Set()).forEach(emp => { delete existentes[emp]; });
  const rows = Object.values(existentes);
  // A diferencia del resto de las hojas de este tipo, acá SÍ hay que poder
  // escribir con rows.length === 0 (cuando se restauran todas las tumbas
  // existentes) — de lo contrario la hoja conserva filas que ya deberían
  // haberse ido. clearContents antes de escribir, por si la lista encogió.
  hoja.clearContents();
  const matrix = [headers, ...rows.map(r => headers.map(h => (r[h] === null || r[h] === undefined) ? '' : r[h]))];
  hoja.getRange(1, 1, matrix.length, headers.length).setValues(matrix);
  return new Set(Object.keys(existentes));
}

// Lista simple de tumbas de empresas_eliminadas deshechas a mano desde el
// dashboard (ver _restaurarEmpresaEliminada) — unión pura de empresa+fecha,
// sin otro campo que fusionar. upsertEmpresasEliminadas usa el resultado
// para excluir esas filas antes de escribir.
function upsertEmpresasRestauradas(ss, nuevas) {
  let hoja = ss.getSheetByName(TAB.empresas_eliminadas_restauradas);
  if (!hoja) hoja = ss.insertSheet(TAB.empresas_eliminadas_restauradas);
  const lastRow = hoja.getLastRow();
  const existentes = {};
  if (lastRow >= 2) {
    const datos = hoja.getRange(1, 1, lastRow, 2).getValues();
    datos.slice(1).forEach(row => {
      const emp = String(row[0] || '').trim();
      if (emp) existentes[emp] = { empresa: emp, fecha: row[1] || '' };
    });
  }
  (nuevas || []).forEach(r => {
    const emp = String(r.empresa || '').trim();
    if (emp && !existentes[emp]) existentes[emp] = { empresa: emp, fecha: r.fecha || '' };
  });
  const rows = Object.values(existentes);
  if (rows.length > 0) {
    const matrix = [['empresa', 'fecha'], ...rows.map(r => [r.empresa, r.fecha])];
    hoja.getRange(1, 1, matrix.length, 2).setValues(matrix);
  }
  return new Set(Object.keys(existentes));
}

// Protecciones puntuales entre un nombre actual y una tumba histórica no
// relacionada (ver _agregarExcepcionTumba en el dashboard) — unión pura
// por empresa_tumbada+nombre_excepcion, nunca se pierde una. El dashboard
// las consulta del lado del cliente (_coincideEmpresaTumba); el bridge acá
// solo las persiste y reenvía, no cambia nada de upsertEmpresasEliminadas.
function upsertEmpresasExcepciones(ss, nuevas) {
  let hoja = ss.getSheetByName(TAB.empresas_eliminadas_excepciones);
  if (!hoja) hoja = ss.insertSheet(TAB.empresas_eliminadas_excepciones);
  const lastRow = hoja.getLastRow();
  const existentes = {};
  if (lastRow >= 2) {
    const datos = hoja.getRange(1, 1, lastRow, 3).getValues();
    datos.slice(1).forEach(row => {
      const empTumbada = String(row[0] || '').trim();
      const nombreExc  = String(row[1] || '').trim();
      if (empTumbada && nombreExc) existentes[empTumbada + '|' + nombreExc] = { empresa_tumbada: empTumbada, nombre_excepcion: nombreExc, fecha: row[2] || '' };
    });
  }
  (nuevas || []).forEach(r => {
    const empTumbada = String(r.empresa_tumbada || r.empresaTumbada || '').trim();
    const nombreExc  = String(r.nombre_excepcion || r.nombreExcepcion || '').trim();
    if (!empTumbada || !nombreExc) return;
    const key = empTumbada + '|' + nombreExc;
    if (!existentes[key]) existentes[key] = { empresa_tumbada: empTumbada, nombre_excepcion: nombreExc, fecha: r.fecha || '' };
  });
  const rows = Object.values(existentes);
  if (rows.length > 0) {
    const matrix = [['empresa_tumbada', 'nombre_excepcion', 'fecha'], ...rows.map(r => [r.empresa_tumbada, r.nombre_excepcion, r.fecha])];
    hoja.getRange(1, 1, matrix.length, 3).setValues(matrix);
  }
  return Object.keys(existentes).length;
}

// Filtra de `rows` cualquier fila cuyo campo "empresa" esté en `tumbas`.
function filtrarEliminadas(rows, tumbas) {
  if (!rows || !tumbas || tumbas.size === 0) return rows;
  return rows.filter(r => !tumbas.has(String(r.empresa || '').trim()));
}

// Normaliza un beneficiario para comparar sin importar mayúsculas, acentos
// o puntuación — mismo criterio que _normBeneficiarioExcluido del dashboard,
// para que "Arch T S.A." y "ARCH T SA" se traten como el mismo excluido.
// Sin espacios (no solo colapsados): un punto reemplazado por espacio
// ("S.A." → "S A") no puede quedar separado de "SA" (sin puntos).
function _normBeneficiarioExcluido(s) {
  return String(s || '').toUpperCase().normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/[^A-Z0-9]/g, '');
}

// ── Beneficiarios de Transferencias excluidos permanentemente ─────
// Merge por upsert: nunca se pierde una exclusión. Devuelve un Set con
// las formas normalizadas de TODOS los excluidos (existentes + nuevos),
// para filtrar transferencias de este mismo request.
function upsertBeneficiariosExcluidos(ss, nuevos) {
  let hoja = ss.getSheetByName(TAB.beneficiarios_excluidos);
  if (!hoja) hoja = ss.insertSheet(TAB.beneficiarios_excluidos);
  const lastRow = hoja.getLastRow();
  const lastCol = hoja.getLastColumn();
  const existentes = {}; // normalizado -> texto original
  let headers = ['beneficiario'];
  if (lastRow >= 2 && lastCol >= 1) {
    const datos = hoja.getRange(1, 1, lastRow, lastCol).getValues();
    headers = datos[0].map(String);
    const bi = headers.indexOf('beneficiario');
    datos.slice(1).forEach(row => {
      const ben = bi >= 0 ? String(row[bi] || '').trim() : '';
      const n = _normBeneficiarioExcluido(ben);
      if (ben && n && !existentes[n]) existentes[n] = ben;
    });
  }
  (nuevos || []).forEach(r => {
    const ben = String((r && r.beneficiario) || r || '').trim();
    const n = _normBeneficiarioExcluido(ben);
    if (ben && n && !existentes[n]) existentes[n] = ben;
  });
  const rows = Object.values(existentes).map(ben => ({ beneficiario: ben }));
  if (rows.length > 0) {
    const matrix = [headers, ...rows.map(r => headers.map(h => (r[h] === null || r[h] === undefined) ? '' : r[h]))];
    hoja.getRange(1, 1, matrix.length, headers.length).setValues(matrix);
  }
  return new Set(Object.keys(existentes));
}

// Filtra de `rows` cualquier fila cuyo "beneficiario" (normalizado) esté
// en `excluidos` (Set de formas normalizadas, ver upsertBeneficiariosExcluidos).
function filtrarBeneficiariosExcluidos(rows, excluidos) {
  if (!rows || !excluidos || excluidos.size === 0) return rows;
  return rows.filter(r => !excluidos.has(_normBeneficiarioExcluido(r.beneficiario)));
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
      hoja.appendRow(['fecha','usuario','nota','fact','resultado','personal','autoridades',
                      'transferencias','notas','presupuesto','balances','rubros',
                      'mensual','er_mensual','er_detalle','flujos','flujos_config',
                      'empresas_eliminadas','beneficiarios_excluidos']);
    }
    hoja.appendRow([
      savedAt, updBy, nota,
      resultados.fact           || 0,
      resultados.resultado      || 0,
      resultados.personal       || 0,
      resultados.autoridades    || 0,
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
      resultados.empresas_eliminadas || 0,
      resultados.beneficiarios_excluidos || 0,
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
  _backupHojaAntesDeEscribir(ss, TAB.transferencias);
  hoja.clearContents();
  const rows = Object.values(existentes);
  if (rows.length === 0) return;
  const matrix = [headers, ...rows.map(function(r){return headers.map(function(h){const v=r[h];return(v===null||v===undefined)?'':v;});})];
  hoja.getRange(1,1,matrix.length,headers.length).setValues(matrix);
}

// ── ER mensual: merge por upsert (empresa+año) ─────────────────────
// A diferencia de escribirHoja (reemplazo total), acá cada fila entrante
// solo reemplaza la fila EXISTENTE con la misma empresa+año — cualquier
// otra empresa/año que ya estuviera en la hoja se conserva tal cual,
// aunque el payload que llega no la mencione. Esto es lo que evita que
// un dispositivo con una copia local desactualizada borre el ER mensual
// de una empresa/año cargada desde otro dispositivo que todavía no bajó.
// Devuelve la cantidad total de filas después del merge (para el log de
// auditoría — antes era simplemente el largo del payload entrante, que ya
// no representa el total real de la hoja).
function upsertErMensual(ss, nuevas, tumbas) {
  let hoja = ss.getSheetByName(TAB.er_mensual);
  if (!hoja) hoja = ss.insertSheet(TAB.er_mensual);
  const lastRow = hoja.getLastRow();
  const lastCol = hoja.getLastColumn();
  const existentes = {}; // 'empresa|anio' -> objeto fila
  let headers = ['empresa', 'anio', 'filas_json', 'conceptos_json', 'es_trimestral'];
  if (lastRow >= 2 && lastCol >= 1) {
    const datos = hoja.getRange(1, 1, lastRow, lastCol).getValues();
    headers = datos[0].map(String);
    const ei = headers.indexOf('empresa'), ai = headers.indexOf('anio');
    datos.slice(1).forEach(row => {
      if (!row.some(c => c !== '')) return;
      const emp = ei >= 0 ? String(row[ei] || '').trim() : '';
      const anio = ai >= 0 ? String(row[ai] || '').trim() : '';
      if (!emp || !anio) return;
      if (tumbas && tumbas.has(emp)) return; // empresa eliminada — no la resucita
      const obj = {}; headers.forEach((h, i) => { obj[h] = row[i] !== undefined ? row[i] : ''; });
      existentes[emp + '|' + anio] = obj;
    });
  }
  (nuevas || []).forEach(r => {
    const emp = String(r.empresa || '').trim(), anio = String(r.anio || '').trim();
    if (!emp || !anio) return;
    existentes[emp + '|' + anio] = r;
  });
  const rows = Object.values(existentes);
  _backupHojaAntesDeEscribir(ss, TAB.er_mensual);
  hoja.clearContents();
  if (rows.length === 0) return 0;
  if (headers.length === 0) headers = Object.keys(rows[0]);
  const matrix = [headers, ...rows.map(r => headers.map(h => { const v = r[h]; return (v === null || v === undefined) ? '' : v; }))];
  hoja.getRange(1, 1, matrix.length, headers.length).setValues(matrix);
  return rows.length;
}

function upsertFlujos(ss, nuevos, aEliminar) {
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
  const seen = new Set();
  const dedupNuevos = (nuevos || []).filter(r => {
    const k = (r.empresa||'') + '_' + (r.periodo||'') + '_' + (r.tipo||'');
    if (seen.has(k)) return false; seen.add(k); return true;
  });
  if (dedupNuevos.length === 0 && elimSet.size === 0) return;
  if (headers.length === 0 && dedupNuevos.length > 0) {
    headers = Object.keys(dedupNuevos[0]);
  }
  if (headers.length === 0) return;
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
  _backupHojaAntesDeEscribir(ss, TAB.flujos);
  hoja.clearContents();
  const rows = Object.values(finalMap);
  if (rows.length === 0) return;
  const allHeaders = headers.length > 0 ? headers : Object.keys(rows[0]);
  const matrix = [allHeaders, ...rows.map(r => allHeaders.map(h => { const v = r[h]; return (v === null || v === undefined) ? '' : v; }))];
  hoja.getRange(1, 1, matrix.length, allHeaders.length).setValues(matrix);
}

// ── Autoridades: merge por upsert (_id) ───────────────
// Mismo principio que upsertFlujos/upsertErMensual: cada fila entrante
// solo reemplaza la existente con el mismo _id si su _ts es igual o más
// nuevo; si no, la fila ya en Sheets (más nueva, bajada todavía por
// ningún otro dispositivo) prevalece. Una fila tumbada en aEliminar (con
// deleted_at >= su propio _ts) se borra, tanto si viene en `nuevas` como
// si ya estaba en la hoja de una publicación anterior. Devuelve la
// cantidad total de filas después del merge (para el log de auditoría).
function upsertAutoridades(ss, nuevas, aEliminar, tumbas) {
  let hoja = ss.getSheetByName(TAB.autoridades);
  if (!hoja) hoja = ss.insertSheet(TAB.autoridades);
  const lastRow = hoja.getLastRow();
  const lastCol = hoja.getLastColumn();
  const existentes = {}; // _id -> objeto fila
  let headers = ['empresa', 'nombre', 'cargo', 'desde', 'hasta', '_id', '_ts'];
  let sinId = 0;
  if (lastRow >= 2 && lastCol >= 1) {
    const datos = hoja.getRange(1, 1, lastRow, lastCol).getValues();
    headers = datos[0].map(String);
    const idIdx = headers.indexOf('_id'), empIdx = headers.indexOf('empresa');
    datos.slice(1).forEach(row => {
      if (!row.some(c => c !== '')) return;
      const emp = empIdx >= 0 ? String(row[empIdx] || '').trim() : '';
      if (tumbas && tumbas.has(emp)) return; // empresa eliminada — no la resucita
      const obj = {}; headers.forEach((h, i) => { obj[h] = row[i] !== undefined ? row[i] : ''; });
      const id = idIdx >= 0 ? String(row[idIdx] || '') : '';
      if (id) existentes[id] = obj; else existentes['__sinid_' + (sinId++)] = obj;
    });
  }
  const delMap = {};
  (aEliminar || []).forEach(t => {
    if (t && t.id) delMap[String(t.id)] = parseInt(t.deleted_at, 10) || 0;
  });
  (nuevas || []).forEach(r => {
    const id = String(r._id || '');
    if (!id) { existentes['__sinid_' + (sinId++)] = r; return; }
    const tsNuevo = parseInt(r._ts, 10) || 0;
    if (Object.prototype.hasOwnProperty.call(delMap, id) && delMap[id] >= tsNuevo) return; // baja más nueva — no resucitarla
    const actual = existentes[id];
    if (!actual || (parseInt(actual._ts, 10) || 0) <= tsNuevo) existentes[id] = r;
    // si no, la fila ya en Sheets es más nueva — se conserva tal cual
  });
  // Bajas publicadas por otro dispositivo sobre filas que YA estaban en
  // Sheets y no vinieron en este envío (esta pestaña aún no las conoce).
  Object.keys(delMap).forEach(id => {
    const fila = existentes[id];
    if (fila && delMap[id] >= (parseInt(fila._ts, 10) || 0)) delete existentes[id];
  });
  const rows = Object.values(existentes);
  _backupHojaAntesDeEscribir(ss, TAB.autoridades);
  hoja.clearContents();
  if (rows.length === 0) return 0;
  if (headers.length === 0) headers = Object.keys(rows[0]);
  const matrix = [headers, ...rows.map(r => headers.map(h => { const v = r[h]; return (v === null || v === undefined) ? '' : v; }))];
  hoja.getRange(1, 1, matrix.length, headers.length).setValues(matrix);
  return rows.length;
}

// ── Rubros (lista maestra de empresas): merge por upsert (empresa) ─
// Mismo principio que upsertAutoridades, pero la clave es "empresa"
// (no hay _id separado: una empresa es su propia clave natural en esta
// hoja). Cada fila entrante solo reemplaza la existente de la misma
// empresa si su _ts es igual o más nuevo; si no, la fila ya en Sheets
// prevalece. Después de fusionar, headers se recalcula como la unión de
// TODAS las claves presentes en `rows` (no solo las que la hoja ya tenía
// antes del merge, ni el fallback hardcodeado de arriba) — sin esto, un
// campo nuevo que llegara en `nuevas` (ej. "cuit", "observaciones") se
// descartaba en silencio al escribir aunque el objeto fusionado en
// memoria sí lo tuviera. Devuelve la cantidad total de filas después
// del merge.
function upsertRubros(ss, nuevas, tumbas) {
  let hoja = ss.getSheetByName(TAB.rubros);
  if (!hoja) hoja = ss.insertSheet(TAB.rubros);
  const lastRow = hoja.getLastRow();
  const lastCol = hoja.getLastColumn();
  const existentes = {}; // empresa -> objeto fila
  let headers = ['empresa', 'rubro', 'activo', 'fecha_inactiva', 'fecha_constitucion', 'decreto', 'actividad', 'tipo_societario', 'participacion_accionaria', '_ts'];
  if (lastRow >= 2 && lastCol >= 1) {
    const datos = hoja.getRange(1, 1, lastRow, lastCol).getValues();
    headers = datos[0].map(String);
    const empIdx = headers.indexOf('empresa');
    datos.slice(1).forEach(row => {
      if (!row.some(c => c !== '')) return;
      const emp = empIdx >= 0 ? String(row[empIdx] || '').trim() : '';
      if (!emp) return;
      if (tumbas && tumbas.has(emp)) return; // empresa eliminada — no la resucita
      const obj = {}; headers.forEach((h, i) => { obj[h] = row[i] !== undefined ? row[i] : ''; });
      existentes[emp] = obj;
    });
  }
  (nuevas || []).forEach(r => {
    const emp = String(r.empresa || '').trim();
    if (!emp) return;
    const tsNuevo = parseInt(r._ts, 10) || 0;
    const actual = existentes[emp];
    if (!actual || (parseInt(actual._ts, 10) || 0) <= tsNuevo) existentes[emp] = r;
    // si no, la fila ya en Sheets es más nueva — se conserva tal cual
  });
  const rows = Object.values(existentes);
  const headersSet = new Set(headers);
  rows.forEach(r => Object.keys(r).forEach(k => headersSet.add(k)));
  headers = [...headersSet];
  _backupHojaAntesDeEscribir(ss, TAB.rubros);
  hoja.clearContents();
  if (rows.length === 0) return 0;
  const matrix = [headers, ...rows.map(r => headers.map(h => { const v = r[h]; return (v === null || v === undefined) ? '' : v; }))];
  hoja.getRange(1, 1, matrix.length, headers.length).setValues(matrix);
  return rows.length;
}

function jsonResp(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
