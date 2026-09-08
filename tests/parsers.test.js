'use strict';
// ═══════════════════════════════════════════════════════════════════════
//  Regresión de los parsers de importación de Excel (Ventas/Gastos,
//  ER columnas de meses, ER matricial) contra la forma real que rompió
//  cada uno en producción.
//
//  No lee ni escribe datos de usuario: corre en Node, fuera del navegador,
//  y construye los workbooks de prueba en memoria (XLSX.utils.aoa_to_sheet)
//  con datos ficticios que reproducen la ESTRUCTURA de cada bug real, no
//  los números reales de ninguna empresa.
//
//  Uso: npm test  (o) node tests/parsers.test.js
// ═══════════════════════════════════════════════════════════════════════
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const INDEX_HTML = path.join(__dirname, '..', 'index.html');

function extractFn(html, name) {
  const startIdx = html.indexOf('function ' + name + '(');
  if (startIdx < 0) {
    throw new Error(`No se encontró function ${name}() en index.html — este test quedó desactualizado respecto al código y necesita revisión, no solo un re-run.`);
  }
  let depth = 0, i = html.indexOf('{', startIdx);
  for (; i < html.length; i++) {
    if (html[i] === '{') depth++;
    else if (html[i] === '}') { depth--; if (depth === 0) { i++; break; } }
  }
  return html.slice(startIdx, i);
}

const html = fs.readFileSync(INDEX_HTML, 'utf8');

// new Function en vez de eval: arma un factory con el código real de
// index.html y devuelve las funciones pedidas — no depende de que este
// archivo corra en modo no-estricto (eval directo pierde las
// declaraciones dentro de un módulo en modo strict).
function loadFns(names) {
  const src = names.map(n => extractFn(html, n)).join('\n');
  const factory = new Function('XLSX', src + '\nreturn { ' + names.join(', ') + ' };');
  return factory(XLSX);
}

const {
  _esIngExcluido,
  _matchEmpresaPortafolio,
  tryParseVentasGastos,
  tryParseERColumnasMeses,
  tryParseERMatricial,
  tryParseERLibroMayor,
  tryParseERLibroMayorResumen,
  _mismoNombreEmpresa,
} = loadFns([
  '_esIngExcluido',
  '_matchEmpresaPortafolio',
  'tryParseVentasGastos',
  'tryParseERColumnasMeses',
  '_parsearPestanasMensuales',
  'tryParseERMatricial',
  'tryParseERLibroMayor',
  'tryParseERLibroMayorResumen',
  '_mismoNombreEmpresa',
]);

// _erDesdeMonthly lee/escribe sobre el global DATA y usa las consts
// ER_LABELS/ER_ORDER — se arma un factory aparte que expone un setter de
// DATA en vez de pasarlo como parámetro, para no tocar la firma real.
function extractConstBlock(html, startMarker, endMarker) {
  const s = html.indexOf(startMarker);
  if (s < 0) throw new Error(`No se encontró "${startMarker}" en index.html`);
  const e = html.indexOf(endMarker, s);
  if (e < 0) throw new Error(`No se encontró "${endMarker}" después de "${startMarker}"`);
  return html.slice(s, e + endMarker.length);
}
const erConsts = extractConstBlock(html, 'const ER_LABELS = {', 'const ER_ORDER = Object.keys(ER_LABELS);');
const erDesdeMonthlySrc = extractFn(html, '_erDesdeMonthly');
const crossCheckSrc = extractFn(html, '_crossCheckEREmpresaAnio');
const repararSrc = extractFn(html, '_repararResultadoOperativoLibroMayor');
const erStdAnualSrc = extractFn(html, '_erStdAnual');
const { _erDesdeMonthly, _crossCheckEREmpresaAnio, _repararResultadoOperativoLibroMayor, _erStdAnual, setData } = new Function(
  erConsts + '\n' + erDesdeMonthlySrc + '\n' + crossCheckSrc + '\n' + repararSrc + '\n' + erStdAnualSrc + '\n' +
  'let DATA = {};\nfunction setData(d) { DATA = d; }\nfunction marcarPendienteGuardar(){}\n' +
  'return { _erDesdeMonthly, _crossCheckEREmpresaAnio, _repararResultadoOperativoLibroMayor, _erStdAnual, setData };'
)();

// tryParseERNativo también escribe sobre el global DATA y llama a funciones
// de refresco de UI (computeTotals/populateSelects/rebuildActive) que no
// existen fuera del dashboard — se stubean como no-op, solo interesa qué
// devuelve el parser y qué queda en DATA.er.
const tryParseERNativoSrc = extractFn(html, 'tryParseERNativo');
const extraerERMensualSrc = extractFn(html, '_extraerERMensualHojaEmpresa');
const resolveHojaSrc = extractFn(html, '_resolveHojaEmpresa');
const { tryParseERNativo, _extraerERMensualHojaEmpresa, _resolveHojaEmpresa, getERNativoData, resetERNativoData } = new Function('XLSX',
  'let DATA = {fact:{},sit:{},er:{},er_mensual:{},fact_mensual:{},res_mensual:{}};\n' +
  'function computeTotals(){}\nfunction populateSelects(){}\nfunction rebuildActive(){}\n' +
  extraerERMensualSrc + '\n' +
  resolveHojaSrc + '\n' +
  tryParseERNativoSrc + '\n' +
  'function getERNativoData(){ return DATA; }\n' +
  'function resetERNativoData(){ DATA.fact={}; DATA.sit={}; DATA.er={}; DATA.er_mensual={}; DATA.fact_mensual={}; DATA.res_mensual={}; }\n' +
  'return { tryParseERNativo, _extraerERMensualHojaEmpresa, _resolveHojaEmpresa, getERNativoData, resetERNativoData };'
)(XLSX);

let pass = 0, fail = 0;
function assert(cond, msg) {
  if (cond) { pass++; console.log('  ✓ ' + msg); }
  else { fail++; console.error('  ✗ ' + msg); }
}
function group(name, fn) { console.log('\n' + name); fn(); }

function wbFromSheets(sheets) {
  const wb = XLSX.utils.book_new();
  Object.entries(sheets).forEach(([name, rows]) =>
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), name));
  return wb;
}

// ── tryParseERColumnasMeses: hoja RESUMEN con encabezado de contribuyente
//    (CUIT, dirección, actividad) antes de la fila de meses. Bug real:
//    Cerámica Riojana, sep/2026 — el archivo no se reconocía como ER
//    provisorio porque la búsqueda de la fila de meses solo miraba las
//    primeras 5 filas de la hoja. ──────────────────────────────────────
group('tryParseERColumnasMeses — encabezado de contribuyente antes de los meses', () => {
  const rows = [
    ['CONTRIBUYENTE:', 'EMPRESA DE PRUEBA'],
    ['CUIT N°:', '30-00000000-0'],
    ['IIBB N°:', '000-000000-0'],
    ['DIRECCION:', 'CALLE FALSA 123'],
    ['ACTIVIDAD:', 'PRUEBA'],
    [],
    ['RENDICIONES MENSUALES'],
    ['', 'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto'],
    ['INGRESOS'],
    ['Ventas netas - Aportes del Estado', 1000, 1100, 1200, 1300, 1400, 1500, 1600, 1700],
    ['Gastos de Administracion', 300, 310, 320, 330, 340, 350, 360, 370],
    ['Rtdo. Antes del impuesto a las Ganancias', 700, 790, 880, 970, 1060, 1150, 1240, 1330],
    ['Rtdo. Neto por operaciones continuas', 700, 790, 880, 970, 1060, 1150, 1240, 1330],
  ];
  const r = tryParseERColumnasMeses(wbFromSheets({ RESUMEN: rows }), 'archivo de prueba 2026.xlsx');
  assert(!!r, 'detecta la hoja pese al encabezado de 5 filas antes de la fila de meses');
  if (!r) return;
  assert(r.anio === '2026', 'toma el año del nombre de archivo cuando la hoja no trae uno');
  const ventas = r.filas.find(f => f.conceptoStd === 'ventas');
  assert(!!ventas, '"Ventas netas - Aportes del Estado" se clasifica como ventas (no se anula por mencionar al Estado)');
  assert(!!ventas && ventas.valores[1] === 1000, 'toma el valor de enero en la columna correcta');
  const gastoAdm = r.filas.find(f => f.lbl === 'Gastos de Administracion');
  assert(!!gastoAdm && gastoAdm.conceptoStd === 'gastoAdm', '"Gastos de Administracion" (con "de") se clasifica como gastoAdm');
  assert(!!r.filas.find(f => f.conceptoStd === 'resAntesImp'), '"Rtdo. Antes..." abreviado se reconoce como resAntesImp');
  assert(!!r.filas.find(f => f.conceptoStd === 'resEjercicio'), '"Rtdo. Neto..." abreviado se reconoce como resEjercicio');
});

// ── tryParseVentasGastos: nombre de archivo genérico. Bug real: Cerdo de
//    los Llanos, el parser sugirió "ACUMULADO" como empresa y (antes del
//    selector agregado en el preview) se guardó así sin confirmar. Este
//    test fija el contrato correcto: el parser solo SUGIERE el nombre
//    crudo — validarlo contra el portafolio es responsabilidad del
//    preview (_mostrarPreviewVentasGastos), no del parser. ─────────────
group('tryParseVentasGastos — nombre de archivo genérico no se resuelve solo', () => {
  const rows = [
    ['Detalle Ventas y Gastos'],
    ['', 'Enero', 'Febrero', 'Marzo'],
    ['DETALLE DE VENTAS'],
    ['Ventas', 500, 520, 540],
    ['TOTALES', 500, 520, 540],
    ['DETALLE DE GASTOS'],
    ['Sueldos', 200, 210, 220],
    ['TOTALES', 200, 210, 220],
  ];
  const r = tryParseVentasGastos(wbFromSheets({ Hoja1: rows }), 'ACUMULADO Ventas y Gastos al 31-03-2026.xlsx');
  assert(!!r, 'detecta el formato Ventas/Gastos (DETALLE DE VENTAS/GASTOS + TOTALES)');
  assert(!!r && r.empresa === 'ACUMULADO', 'el parser sigue sugiriendo el nombre crudo del archivo — si esto cambia, revisar que _mostrarPreviewVentasGastos siga exigiendo confirmación contra el portafolio antes de aplicar');
});

// ── tryParseERMatricial: multi-empresa genuino + fallback de año por
//    nombre de archivo. Bug real: en Cerámica Riojana este parser
//    interpretaba mal filas de encabezado (CUIT como si fuera período,
//    sin año reconocible) y _aplicarERMatricial() descartaba la fila
//    en silencio sin avisar nada. ───────────────────────────────────────
group('tryParseERMatricial — multi-empresa genérico + fallback de año', () => {
  const rows = [
    ['', 'EMPRESA A', 'EMPRESA B'],
    ['', 'Sin período reconocible', 'Sin período reconocible'],
    ['Ventas Netas', 1000, 2000],
    ['Costo', -400, -900],
    ['Resultado Operativo', 100, -50],
  ];
  const r = tryParseERMatricial(wbFromSheets({ RESUMEN: rows }), 'ER 2025.xlsx');
  assert(!!r, 'detecta el formato matricial multi-empresa');
  if (!r) return;
  assert(r.anio === 2025, 'usa el año del nombre de archivo cuando ninguna columna trae uno (antes quedaba null y la fila se perdía sin aviso)');
  assert(r.empresas.length === 2, 'detecta las dos empresas de las columnas');
  const a = r.empresas.find(e => e.empresa === 'EMPRESA A');
  assert(!!a && a.facturacion === 1000 && a.resultado === 100, 'toma facturación y resultado de EMPRESA A de las filas correctas');
});

// ── _matchEmpresaPortafolio: acotado al portafolio de Rubros Comerciales,
//    tolerante a sufijos societarios pero sin inventar coincidencias. ────
group('_matchEmpresaPortafolio', () => {
  const candidatas = ['CERAMICA RIOJANA', 'CERDO DE LOS LLANOS'];
  assert(_matchEmpresaPortafolio('CERAMICA RIOJANA SAPEM', candidatas) === 'CERAMICA RIOJANA',
    'reconoce la empresa pese al sufijo societario ("SAPEM") ausente del portafolio');
  assert(_matchEmpresaPortafolio('UNA EMPRESA QUE NO EXISTE', candidatas) === '',
    'no inventa una coincidencia para un nombre fuera del portafolio');
});

// ── _beneficiarioCanonico: pedido real del usuario — un cambio de forma
//    societaria (ej. "Alfa SAPEM" → "Alfa SAU") hace que las
//    transferencias de antes y después del cambio queden reportadas con
//    cada nombre. Los rankings/totales por beneficiario de Transferencias
//    agrupaban por el texto crudo, así que aparecían como dos
//    beneficiarios distintos, partiendo el total. Esta función decide
//    bajo qué nombre agrupar (la empresa del portafolio, si matchea) SIN
//    tocar el registro original de la transferencia — eso es a propósito
//    responsabilidad de quien llama a la función, no de esta. ─────────
group('_beneficiarioCanonico — agrupa transferencias por empresa del portafolio, no por el texto crudo', () => {
  const src = extractFn(html, '_beneficiarioCanonico');
  const srcMatch = extractFn(html, '_matchEmpresaPortafolio');
  const { _beneficiarioCanonico, setRubros } = new Function(
    'let rubros = {};\n' + srcMatch + '\n' + src + '\n' +
    'function setRubros(r){ rubros = r; }\n' +
    'return { _beneficiarioCanonico, setRubros };'
  )();
  setRubros({ 'ALFA': 'Industrias', 'CERAMICA RIOJANA': 'Industrias' });

  assert(_beneficiarioCanonico('Alfa SAPEM') === 'ALFA',
    'agrupa "Alfa SAPEM" bajo la empresa del portafolio "ALFA" (cambio de forma societaria)');
  assert(_beneficiarioCanonico('Alfa SAU') === 'ALFA',
    'agrupa "Alfa SAU" bajo la misma empresa "ALFA" — mismo total que "Alfa SAPEM", no partido en dos');
  assert(_beneficiarioCanonico('Ceramica Riojana SAPEM') === 'CERAMICA RIOJANA',
    'funciona igual para cualquier empresa del portafolio, no solo el caso de ejemplo');
  assert(_beneficiarioCanonico('Municipalidad de Chilecito') === 'Municipalidad de Chilecito',
    'un beneficiario que no es una empresa del portafolio (un tercero real) se deja tal cual, sin inventar una empresa');
});

// ── _mismoNombreEmpresa: usada por la alerta "Datos faltantes" (¿esta
//    empresa tiene un balance cargado?) en vez del matcheo por prefijo
//    fijo que había antes. Bugs reales reportados por el usuario:
//    - LA RIOJA VITICOLA aparecía con balances que en realidad eran de
//      LA RIOJA TELECOMUNICACIONES — ambas comparten los primeros 8
//      caracteres ("LA RIOJA"), que era el largo de prefijo usado.
//    - VALLESOL (8 caracteres exactos) nunca podía reconciliarse con
//      "VALLE SOL" (con espacio) porque el código exigía más de 8
//      caracteres para siquiera intentar la comparación difusa. ────────
group('_mismoNombreEmpresa — reemplaza el matcheo por prefijo fijo (causaba falsos positivos y negativos)', () => {
  assert(_mismoNombreEmpresa('LA RIOJA VITICOLA', 'LA RIOJA TELECOMUNICACIONES') === false,
    'NO confunde dos empresas distintas que comparten un prefijo largo ("LA RIOJA ...") — bug real: LA RIOJA VITICOLA mostraba balances de LA RIOJA TELECOMUNICACIONES');
  assert(_mismoNombreEmpresa('VALLESOL', 'VALLE SOL') === true,
    'reconcilia un nombre corto (8 caracteres) con espacio de más/menos — bug real: VALLESOL nunca reconciliaba con "VALLE SOL" por un piso de longitud mínima');
  assert(_mismoNombreEmpresa('PUERTAS DE SOL', 'PUERTAS DEL SOL') === true,
    'tolera artículos distintos ("DE" vs "DEL")');
  assert(_mismoNombreEmpresa('VIENTOS ARAUCO RENOVABLE', 'VIENTOS DE ARAUCO RENOVABLES') === true,
    'tolera singular/plural y un artículo faltante');
  assert(_mismoNombreEmpresa('CERAMICA RIOJANA', 'CERAMICA RIOJANA SAPEM') === true,
    'ignora el sufijo societario');
  assert(_mismoNombreEmpresa('RIOJA VIAL', 'RIOJA BUS') === false,
    'no inventa una coincidencia entre dos empresas distintas que comparten una sola palabra');
  assert(_mismoNombreEmpresa('AGUAS RIOJANAS', 'RIOJA BUS') === false,
    'no inventa una coincidencia por compartir la raíz "RIOJA"');
});

// ── _confirmarRenombre: pedido real del usuario — "la unificación de
//    empresas no debe eliminar registros pero sí la denominación
//    elegida... debería desaparecer de ese reporte". Bug real: al
//    renombrar una empresa (doble clic en el editor de rubros) a un
//    nombre que YA existe, la función solo mostraba "Ya existe una
//    empresa con ese nombre" y no hacía nada — a diferencia de
//    Gestión → Renombrar/Fusionar, nunca ofrecía combinar los datos. La
//    empresa vieja quedaba entonces con su rubro en blanco ("Sin rubro")
//    y sus datos sueltos, sin que nada la borre del portafolio. ────────
group('_confirmarRenombre — unificar denominación (renombrar a un nombre ya existente) fusiona en vez de bloquear', () => {
  const srcRenombrar   = extractFn(html, '_renombrarEmpresa');
  const srcFusionar    = extractFn(html, '_fusionarEmpresas');
  const srcConfirmar   = extractFn(html, '_confirmarRenombre');
  const factory = new Function(
    'let rubros = {}; let DATA = {fact:{},sit:{},personal:[],notas:{},presupuesto:{}}; let _balances = {};\n' +
    'let confirmResult = true;\n' +
    'function confirm(msg){ return confirmResult; }\n' +
    'function saveRubros(){} function computeTotals(){} function saveDataToLocalCache(){}\n' +
    'function marcarPendienteGuardar(){} function populateSelects(){} function rebuildActive(){} function buildRubros(){}\n' +
    srcRenombrar + '\n' + srcFusionar + '\n' + srcConfirmar + '\n' +
    'function setState(r,d,b){ rubros=r; DATA=d; _balances=b; }\n' +
    'function getState(){ return { rubros, DATA, _balances }; }\n' +
    'function setConfirmResult(v){ confirmResult = v; }\n' +
    'return { _confirmarRenombre, setState, getState, setConfirmResult };'
  );
  const { _confirmarRenombre, setState, getState, setConfirmResult } = factory();

  setState(
    { 'LRT': 'Sin rubro', 'LA RIOJA TELECOMUNICACIONES': 'Servicios' },
    {
      fact: { 'LRT': { '2020': 1000000 }, 'LA RIOJA TELECOMUNICACIONES': { '2023': 5000000 } },
      sit:  { 'LRT': { '2020': 50000 },   'LA RIOJA TELECOMUNICACIONES': { '2023': 200000 } },
      personal: [], notas: {}, presupuesto: {},
    },
    { b1: { empresa: 'LRT', periodo: '2020' } }
  );
  setConfirmResult(true);
  _confirmarRenombre('LRT', 'LA RIOJA TELECOMUNICACIONES');
  const st = getState();
  assert(!('LRT' in st.rubros), 'la empresa vieja ("LRT") desaparece del portafolio (rubros) tras unificar');
  assert(st.DATA.fact['LA RIOJA TELECOMUNICACIONES']['2020'] === 1000000 && st.DATA.fact['LA RIOJA TELECOMUNICACIONES']['2023'] === 5000000,
    'no se pierde ningún registro — la facturación de ambos nombres queda combinada bajo el elegido');
  assert(st.DATA.sit['LA RIOJA TELECOMUNICACIONES']['2020'] === 50000 && st.DATA.sit['LA RIOJA TELECOMUNICACIONES']['2023'] === 200000,
    'lo mismo para el resultado — se combinan, no se sobrescriben');
  assert(st.rubros['LA RIOJA TELECOMUNICACIONES'] === 'Servicios',
    'conserva el rubro de la empresa elegida (no lo pisa con "Sin rubro" del nombre viejo)');
  assert(st._balances.b1.empresa === 'LA RIOJA TELECOMUNICACIONES',
    'los balances de la empresa vieja quedan re-vinculados al nombre elegido, no huérfanos');

  // Si el usuario cancela la confirmación, no debe tocar nada.
  setState(
    { 'PEA': 'Sin rubro', 'PARQUE EOLICO ARAUCO': 'Servicios' },
    { fact: { 'PEA': { '2020': 1 } }, sit: {}, personal: [], notas: {}, presupuesto: {} },
    {}
  );
  setConfirmResult(false);
  _confirmarRenombre('PEA', 'PARQUE EOLICO ARAUCO');
  const st2 = getState();
  assert('PEA' in st2.rubros, 'si el usuario cancela la confirmación de unificar, no borra ni fusiona nada');
});

// ── _erDesdeMonthly: el panel "ER Provisorio — <empresa>" leía solo
//    DATA.er_mensual (detalle mes a mes, cargado individualmente desde
//    2025). Bug real: empresas cargadas antes vía RESUMEN histórico
//    multi-empresa (tryParseERNativo → DATA.er[emp][año], mismo desglose
//    de partidas pero sin apertura mensual) mostraban "Sin ER provisorio
//    cargado" pese a tener el dato — la vista nunca consultaba DATA.er. ──
group('_erDesdeMonthly — usa DATA.er (anual nativo) cuando no hay detalle mensual', () => {
  setData({
    er_mensual: {
      'CERDO DE LOS LLANOS': {
        '2025': { filas: [
          { conceptoStd: 'ventas', valores: { 1: 100, 2: 120 } },
          { conceptoStd: 'resOperativo', valores: { 1: 10, 2: 12 } },
        ] },
      },
    },
    er: {
      'CERDO DE LOS LLANOS': {
        // 2025 también existe en er_mensual (arriba) — el detalle mensual
        // tiene que ganar y este valor claramente distinto no debe aparecer
        '2025': { ventas_netas: -1, resultado_operativo: -1 },
        '2022': { ventas_netas: 1127073623.87, costo_ventas: -1101850481.26,
                   utilidad_bruta: 25223142.61, gastos_admin: -83362268.76,
                   gastos_comercializacion: -111320403.98, resultado_operativo: 119884283.08 },
      },
    },
    fact: {}, sit: {},
  });
  const er = _erDesdeMonthly('CERDO DE LOS LLANOS');
  assert(!!er && er['2022'] && er['2022'].costo_ventas === -1101850481.26,
    'un año sin detalle mensual pero con RESUMEN histórico (DATA.er) trae el desglose completo de partidas');
  assert(er['2022'].gastos_admin === -83362268.76 && er['2022'].gastos_comercializacion === -111320403.98,
    'incluye partidas que el fallback anterior (solo fact/sit) no tenía — mismo formato de reporte que el ER mensual');
  assert(er['2025'].ventas_netas === 220, 'si el año ya tiene detalle mensual, el ER anual nativo (DATA.er) no lo pisa');

  const empEr = _erDesdeMonthly('CERAMICA RIOJANA'); // no cargada en absoluto
  assert(empEr === null, 'una empresa sin ningún dato (ni mensual, ni DATA.er, ni fact/sit) sigue devolviendo null');
});

// ── _erDesdeMonthly: bug real — al pedirle a los 4 excel reales el
//    detalle mensual extraído por anclas, Resultado Operativo salía casi
//    el DOBLE del valor real para varias empresas (ej. VALLE SOL 2022:
//    855.6M mostrado vs. 427.8M real). Causa: 'resAntesImp' y
//    'resEjercicio' mapeaban a la MISMA clave 'resultado_operativo' que
//    'resOperativo' y el agregador las SUMABA todas juntas — una empresa
//    con fila 'resOperativo' (el ancla real) Y fila 'resEjercicio' (el
//    resultado del ejercicio, un concepto distinto que el extractor por
//    anclas siempre agrega por separado) terminaba con ambos valores
//    sumados en vez de solo el operativo. ─────────────────────────────
group('_erDesdeMonthly — no duplica ni aproxima Resultado Operativo con Resultado del Ejercicio', () => {
  setData({
    er_mensual: {
      'EMPRESA CON AMBAS FILAS': { '2022': { filas: [
        { conceptoStd: 'ventas', valores: { 1: 1000000 } },
        { conceptoStd: 'resOperativo', valores: { 1: 300000 } },     // el ancla real
        { conceptoStd: 'resEjercicio', valores: { 1: 280000 } },     // resultado final, concepto distinto
      ] } },
      'EMPRESA SOLO CON RESULTADO DEL EJERCICIO': { '2022': { filas: [
        { conceptoStd: 'ventas', valores: { 1: 1000000 } },
        { conceptoStd: 'resEjercicio', valores: { 1: 280000 } },     // sin Resultado Operativo propio
      ] } },
    },
    er: {},
    fact: {},
    sit: { 'EMPRESA SOLO CON RESULTADO DEL EJERCICIO': { '2022': 305000 } }, // valor anual "oficial"
  });
  const conAmbas = _erDesdeMonthly('EMPRESA CON AMBAS FILAS');
  assert(conAmbas['2022'].resultado_operativo === 300000,
    'usa el valor de la fila "resOperativo" tal cual, sin sumarle la fila "resEjercicio" (que es un concepto distinto)');

  const soloEjercicio = _erDesdeMonthly('EMPRESA SOLO CON RESULTADO DEL EJERCICIO');
  assert(soloEjercicio['2022'].resultado_operativo === 305000,
    'sin una fila "resOperativo" propia, cae al valor anual "oficial" (DATA.sit) en vez de aproximar con "resEjercicio" (puede diferir mucho: son conceptos distintos)');
});

// ── _crossCheckEREmpresaAnio: pedido real del usuario — "control cruzado
//    entre los reportes de EVOLUCION EMPRESA y de RESULTADO OPERATIVO para
//    detectar diferencias entre los importes expuestos". Ambos módulos
//    muestran el ER Provisorio vía _erDesdeMonthly (suma de los anclas
//    mensuales), que puede diferir de Facturación/Resultado Operativo
//    "oficiales" (DATA.fact/DATA.sit) cuando el detalle mensual se extrajo
//    por anclas y validó contra el RESUMEN dentro de una tolerancia, no de
//    forma exacta. Antes esa diferencia quedaba invisible. ──────────────
group('_crossCheckEREmpresaAnio — detecta diferencias entre el ER Provisorio y Facturación/Resultado "oficiales"', () => {
  setData({
    er_mensual: {
      'EMPRESA CON DIFERENCIA': {
        '2021': { filas: [
          { conceptoStd: 'ventas', valores: { 1: 500000, 2: 520000 } },       // suma 1.020.000
          { conceptoStd: 'resOperativo', valores: { 1: 40000, 2: 42000 } },   // suma 82.000
        ] },
      },
      'EMPRESA SIN DIFERENCIA': {
        '2021': { filas: [
          { conceptoStd: 'ventas', valores: { 1: 500000, 2: 520000 } },
          { conceptoStd: 'resOperativo', valores: { 1: 40000, 2: 42000 } },
        ] },
      },
    },
    er: {},
    fact: {
      'EMPRESA CON DIFERENCIA': { '2021': 1100000 }, // 7.8% más que el ER Provisorio (1.020.000)
      'EMPRESA SIN DIFERENCIA': { '2021': 1020000 }, // coincide exacto
    },
    sit: {
      'EMPRESA CON DIFERENCIA': { '2021': 82000 },   // coincide (solo Facturación difiere)
      'EMPRESA SIN DIFERENCIA': { '2021': 82000 },
    },
  });
  const diffs = _crossCheckEREmpresaAnio('EMPRESA CON DIFERENCIA', '2021');
  assert(!!diffs && diffs.length === 1 && diffs[0].label === 'Facturación',
    'detecta que Facturación difiere entre el ER Provisorio (suma mensual) y el valor "oficial" (DATA.fact)');
  assert(diffs[0].erProvisorio === 1020000 && diffs[0].oficial === 1100000,
    'informa ambos valores (el del ER Provisorio y el oficial) para poder comparar');

  const sinDiff = _crossCheckEREmpresaAnio('EMPRESA SIN DIFERENCIA', '2021');
  assert(sinDiff === null, 'no marca nada cuando el ER Provisorio y los valores oficiales coinciden');

  // Diferencia menor a la tolerancia (0.5%) — no debe generar ruido por
  // redondeos triviales.
  setData({
    er_mensual: { 'EMPRESA DIFERENCIA MINIMA': { '2021': { filas: [
      { conceptoStd: 'ventas', valores: { 1: 1000000 } },
    ] } } },
    er: {}, fact: { 'EMPRESA DIFERENCIA MINIMA': { '2021': 1002000 } }, sit: {}, // 0.2%
  });
  assert(_crossCheckEREmpresaAnio('EMPRESA DIFERENCIA MINIMA', '2021') === null,
    'no marca diferencias menores a la tolerancia (ruido de redondeo, no una discrepancia real)');
});

// ── _erDesdeMonthly: bug real — Parque Eólico Arauco (ER 2026, formato
//    "libro mayor"). El panel Evolución Empresa (cascada propia) SÍ sumaba
//    Otros Ingresos al derivar Resultado Operativo; _erDesdeMonthly (usado
//    por Resultado Operativo → Desglose ER, y por el cruce contra
//    Facturación/Resultado "oficiales") no lo hacía — mismos datos,
//    Resultado Operativo salía distinto según qué pantalla lo mostrara,
//    sin que hubiera ningún error real en los datos. ────────────────────
group('_erDesdeMonthly — deriva Resultado Operativo incluyendo Otros Ingresos/Egresos y Gastos Operativos (igual que Evolución Empresa)', () => {
  setData({
    er_mensual: {
      'PEA': { '2026': { filas: [
        { conceptoStd: 'ventas',        valores: { 1: 60714891 } },
        { conceptoStd: 'costo',         valores: { 1: -313252065 } },
        { conceptoStd: 'gastoAdm',      valores: { 1: -705811310 } },
        { conceptoStd: 'gastoCom',      valores: { 1: -21442246 } },
        { conceptoStd: 'resFinanciero', valores: { 1: -1036821249 } },
        { conceptoStd: 'otrosIngresos', valores: { 1: 404881000 } },
        { conceptoStd: 'resEjercicio',  valores: { 1: -1611730980 } },
      ] } },
    },
    er: {}, fact: {}, sit: {},
  });
  const er = _erDesdeMonthly('PEA')['2026'];
  // Utilidad Bruta = 60.714.891 - 313.252.065 = -252.537.174
  // Resultado Operativo = Utilidad Bruta + Gastos Adm + Gastos Com + Otros Ingresos
  //                      = -252.537.174 - 705.811.310 - 21.442.246 + 404.881.000 = -574.909.730
  assert(er.resultado_operativo === -574909730,
    'incluye Otros Ingresos en la derivación de Resultado Operativo (antes quedaba afuera: daba -979.790.730)');
  assert(er.otros_ingresos === 404881000, 'expone Otros Ingresos como partida propia (antes se perdía: no había clave en el mapeo)');
  assert(er.gastos_financieros === -1036821249, 'Resultado Financiero sigue expuesto aparte, sin entrar en Resultado Operativo (va después en la cascada)');
});

group('_erDesdeMonthly — con Gastos Operativos como única partida de gasto (sin Admin ni Comercialización) también deriva Resultado Operativo', () => {
  setData({
    er_mensual: {
      'EMPRESA SOLO GASTOS OPERATIVOS': { '2022': { filas: [
        { conceptoStd: 'ventas',   valores: { 1: 1000000 } },
        { conceptoStd: 'costo',    valores: { 1: -400000 } },
        { conceptoStd: 'gastoOper', valores: { 1: -100000 } },
      ] } },
    },
    er: {}, fact: {}, sit: {},
  });
  const er = _erDesdeMonthly('EMPRESA SOLO GASTOS OPERATIVOS')['2022'];
  assert(er.resultado_operativo === 500000,
    'deriva Resultado Operativo (Utilidad Bruta + Gastos Operativos) aunque no haya Gastos de Admin/Comercialización propios');
});

// ── _repararResultadoOperativoLibroMayor: bug real — PEA, AR, Kallpa y
//    AESA ya habían sido cargados (con _aplicarERColumnas) ANTES de
//    unificar la fórmula de Resultado Operativo, así que quedaron con el
//    resultado FINAL del ejercicio guardado en DATA.sit — la corrección
//    de la fórmula (arriba) no alcanza para ellos porque DATA.sit ya
//    estaba guardado en Sheets con el valor viejo, y no se recalcula solo
//    al refrescar. Esta reparación corre en cada sync y corrige esos
//    casos sin pedirle al usuario que vuelva a subir el archivo. ───────
group('_repararResultadoOperativoLibroMayor — corrige en Sheets el Resultado Operativo de empresas cargadas antes de unificar la fórmula', () => {
  setData({
    er_mensual: {
      'PEA': { '2026': { filas: [
        { conceptoStd: 'ventas',        valores: { 1: 1000000 } },
        { conceptoStd: 'costo',         valores: { 1: -400000 } },
        { conceptoStd: 'gastoAdm',      valores: { 1: -100000 } },
        { conceptoStd: 'otrosIngresos', valores: { 1: 50000 } },
        { conceptoStd: 'resFinanciero', valores: { 1: -300000 } },
        { conceptoStd: 'resEjercicio',  valores: { 1: -150000 } }, // Total General: lo que había quedado en DATA.sit
      ] } },
      // Empresa con fila 'resOperativo' propia (ancla real, ej. RESUMEN
      // histórico 2021-2024) — no debe tocarse: no tiene la firma del bug.
      'EMPRESA CON ANCLA PROPIA': { '2022': { filas: [
        { conceptoStd: 'ventas',       valores: { 1: 1000000 } },
        { conceptoStd: 'resOperativo', valores: { 1: 120000 } },
        { conceptoStd: 'resEjercicio', valores: { 1: 90000 } },
      ] } },
      // Ya reparada en una corrida anterior — no debe volver a marcarse.
      'EMPRESA YA CORRECTA': { '2026': { filas: [
        { conceptoStd: 'ventas',       valores: { 1: 500000 } },
        { conceptoStd: 'costo',        valores: { 1: -200000 } },
        { conceptoStd: 'resEjercicio', valores: { 1: 300000 } },
      ] } },
    },
    er: {}, fact: {},
    sit: {
      'PEA': { '2026': -150000 }, // el resultado final, guardado por error
      'EMPRESA CON ANCLA PROPIA': { '2022': 120000 },
      'EMPRESA YA CORRECTA': { '2026': 300000 }, // Utilidad Bruta = 500000-200000 = 300000, ya coincide
    },
  });

  const reparados = _repararResultadoOperativoLibroMayor();

  assert(reparados.length === 1 && reparados[0].emp === 'PEA' && reparados[0].yr === '2026',
    'detecta y corrige solo la empresa/año con la firma del bug (fila "resEjercicio" sin "resOperativo" propia y DATA.sit desactualizado)');
  // Resultado Operativo = Ventas + Costo + GastoAdm + OtrosIngresos = 1.000.000-400.000-100.000+50.000 = 550.000
  assert(reparados[0].antes === -150000 && reparados[0].despues === 550000,
    'recalcula con la misma fórmula que el resto de los reportes (incluye Otros Ingresos, excluye Resultado Financiero)');

  const er = _erDesdeMonthly('PEA')['2026'];
  assert(er.resultado_operativo === 550000, 'el valor corregido queda disponible para el resto de los reportes');

  const reparados2 = _repararResultadoOperativoLibroMayor();
  assert(reparados2.length === 0, 'no vuelve a marcar nada una vez corregido (idempotente — no genera ruido en cada sync)');
});

// ── _erStdAnual: bug real detectado al verificar la reparación de arriba
//    — clasificaba las filas de Otros Ingresos/Otros Egresos/Gastos
//    Operativos por el TEXTO del rótulo (clasificar(f.lbl)) en vez de por
//    conceptoStd directo, a diferencia de ventas/costo/gastoAdm/etc., que
//    sí se reconocían por conceptoStd sin depender del texto. Con un
//    rótulo que no matcheaba el clasificador de texto, la fila se perdía
//    en silencio — mismo tipo de inconsistencia que el bug de Resultado
//    Operativo ya corregido, esta vez en la vista "Evolución Empresa"
//    (buildEvolucionMensual usa el mismo patrón, corregido en el mismo
//    commit — no expuesta acá por depender del DOM). ───────────────────
group('_erStdAnual — clasifica Otros Ingresos/Egresos y Gastos Operativos por conceptoStd, no por el texto del rótulo', () => {
  setData({
    er_mensual: {
      'EMPRESA CON ROTULOS RAROS': { '2026': { filas: [
        { lbl: 'x1', conceptoStd: 'ventas',        valores: { 1: 1000000 } },
        { lbl: 'x2', conceptoStd: 'costo',         valores: { 1: -400000 } },
        { lbl: 'x3', conceptoStd: 'gastoAdm',      valores: { 1: -100000 } },
        // Rótulos que NO matchean ningún patrón de clasificar(lbl) —
        // antes de este fix, estas dos filas se perdían en silencio.
        { lbl: 'Ajuste 4821-B', conceptoStd: 'otrosIngresos', valores: { 1: 50000 } },
        { lbl: 'Partida 9902-C', conceptoStd: 'otrosGastos',  valores: { 1: -20000 } },
      ] } },
    },
    er: {}, fact: {}, sit: {},
  });
  const er = _erStdAnual('EMPRESA CON ROTULOS RAROS', '2026');
  assert(er.otros_ingresos === 50000, 'reconoce Otros Ingresos por conceptoStd aunque el rótulo no sea reconocible por texto');
  // Resultado Operativo = Utilidad Bruta(600.000) + GastoAdm(-100.000) + OtrosIngresos(50.000) + OtrosGastos(-20.000) = 530.000
  assert(er.resultado_operativo === 530000,
    'incluye Otros Ingresos/Egresos en Resultado Operativo aunque no se hayan podido clasificar por el texto del rótulo');
});

// ── tryParseERNativo: bug real — al recargar el RESUMEN histórico multi-
//    empresa de 2022 (ER_2022_ultimo.xlsx), el dashboard lo tomaba como si
//    fuera un archivo de UNA sola empresa (Agroandina), en vez del desglose
//    completo de las ~30 empresas del RESUMEN. Causa: el dispatcher probaba
//    tryParseERColumnasMeses ANTES que tryParseERNativo, y ese parser recorre
//    TODAS las hojas del libro — la hoja individual de Agroandina (una de las
//    ~30 hojas por empresa que trae el mismo archivo) coincidía por accidente
//    con su heurística de "columnas de meses" antes de llegar a la hoja
//    RESUMEN. Fix: tryParseERNativo se prueba primero en el dispatcher, y se
//    le agregó una verificación estructural (Ventas/Utilidad Bruta/Resultado
//    Operativo en las filas fijas del layout) para que solo dispare con el
//    formato multi-empresa real y no con otra hoja que también se llame
//    "RESUMEN" pero sea de una sola empresa (ej. Cerámica Riojana). ────────
group('tryParseERNativo — RESUMEN multi-empresa real vs. RESUMEN de una sola empresa', () => {
  const rowsMulti = [
    ['FECHA DE PRESENTACIÓN', 'EMPRESA A', 'EMPRESA B'],
    ['', 'A DICIEMBRE 2022', 'A DICIEMBRE 2022'],
    ['INGRESOS', '', ''],
    ['Ventas Netas', 1000000, 2000000],
    ['Costo de Ventas', -400000, -900000],
    ['Utilidad Bruta', 600000, 1100000],
    [],
    ['Gastos Financieros', 0, -10000],
    ['Gastos de Administracion', -100000, -200000],
    ['Gastos Operativos', 0, 0],
    ['Gastos Fiscales', 0, 0],
    ['Gastos Bancarios', 0, 0],
    ['Gastos de Produccion', 0, 0],
    ['Gastos de Comercializacion', -50000, -80000],
    ['Gastos de Obra', 0, 0],
    ['Otros Gastos', 0, 0],
    ['Otros Ingresos', 0, 0],
    ['Otros Egresos', 0, 0],
    ['Resultado Operativo', 450000, 810000],
  ];
  resetERNativoData();
  const wbMulti = wbFromSheets({ RESUMEN: rowsMulti });
  const rMulti = tryParseERNativo(wbMulti, 'ER_2022_ultimo.xlsx');
  assert(!!rMulti, 'detecta el RESUMEN multi-empresa real (empresas en columnas, partidas en filas fijas)');
  const dataMulti = getERNativoData();
  assert(Object.keys(dataMulti.er).length === 2 && dataMulti.er['EMPRESA A'] && dataMulti.er['EMPRESA B'],
    'carga el desglose de TODAS las empresas del RESUMEN, no solo la primera hoja del libro');
  assert(dataMulti.er['EMPRESA B']['2022'].resultado_operativo === 810000,
    'toma el valor de la fila "Resultado Operativo" (fija en el layout) para cada empresa');

  // Bug real: recargar el RESUMEN de 2022 detectaba las 25 empresas
  // correctamente (ver arriba) pero no impactaba en Facturación/Resultado
  // de Evolución por Empresa para las que ya tenían un 0 guardado — el
  // sync solo completaba si el valor existente era undefined/null, y un 0
  // (el estado real de una empresa dada de alta sin datos cargados) no
  // calificaba como "ausente", así que quedaba como mera discrepancia sin
  // aplicarse nunca. ──────────────────────────────────────────────────
  resetERNativoData();
  const dataPre = getERNativoData();
  dataPre.fact['EMPRESA A'] = { '2022': 0 };   // "sin cargar" real
  dataPre.sit['EMPRESA A']  = { '2022': 0 };
  dataPre.fact['EMPRESA B'] = { '2022': 1500000 }; // dato real distinto al ER
  dataPre.sit['EMPRESA B']  = { '2022': 810000 };  // coincide con el ER (sin discrepancia)
  tryParseERNativo(wbFromSheets({ RESUMEN: rowsMulti }), 'ER_2022_ultimo.xlsx');
  assert(dataPre.fact['EMPRESA A']['2022'] === 1000000 && dataPre.sit['EMPRESA A']['2022'] === 450000,
    'un 0 guardado en Facturación/Resultado se trata como ausente y SÍ se completa al recargar el RESUMEN');
  assert(dataPre.fact['EMPRESA B']['2022'] === 1500000,
    'un valor existente distinto de 0 nunca se sobreescribe, aunque difiera del RESUMEN (queda como discrepancia)');
  assert(dataPre.sit['EMPRESA B']['2022'] === 810000,
    'un valor existente que coincide con el RESUMEN se deja intacto');

  // Hoja también llamada "RESUMEN" pero de una sola empresa (contribuyente/
  // período en vez de empresas en columnas) — no debe matchear como si fuera
  // el formato multi-empresa, para no taparle el archivo a tryParseERColumnasMeses.
  // Padeada a ≥19 filas (como un archivo real) para ejercitar la verificación
  // estructural nueva y no solo el chequeo preexistente de "menos de 19 filas".
  const rowsUnaEmpresa = [
    ['CONTRIBUYENTE:', 'CERAMICA RIOJANA SAPEM'],
    ['CUIT N°:', '30-00000000-0'],
    ['IIBB N°:', '000-000000-0'],
    ['DIRECCION:', 'CALLE FALSA 123'],
    ['ACTIVIDAD:', 'INDUSTRIA'],
    [],
    ['RENDICIONES MENSUALES'],
    ['', 'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio'],
    ['INGRESOS'],
    ['Ventas netas', 1000, 1100, 1200, 1300, 1400, 1500],
    ['Gastos de Administracion', -300, -310, -320, -330, -340, -350],
    ['Rtdo. Antes del impuesto a las Ganancias', 700, 790, 880, 970, 1060, 1150],
    ['Rtdo. Neto por operaciones continuas', 700, 790, 880, 970, 1060, 1150],
    [], [], [], [], [], [],
  ];
  resetERNativoData();
  const wbUna = wbFromSheets({ RESUMEN: rowsUnaEmpresa });
  const rUna = tryParseERNativo(wbUna, 'ER 2026.xlsx');
  assert(!rUna, 'una hoja "RESUMEN" de una sola empresa (contribuyente/período, no empresas en columnas) no matchea como multi-empresa, aunque tenga ≥19 filas');
});

// ── _extraerERMensualHojaEmpresa: extrae el detalle mensual real de la
//    hoja individual de cada empresa (dentro del mismo libro del RESUMEN
//    multi-empresa) usando filas "ancla" ya calculadas en el excel
//    (Ventas, Costo, Utilidad Bruta, Resultado Operativo, Resultado del
//    Ejercicio) en vez de reclasificar cada línea de gasto idiosincrática.
//    Se autoverifica contra el total anual ya confiable del RESUMEN (5% de
//    tolerancia) — si no valida, no se aplica nada para esa empresa/año. ──
group('_extraerERMensualHojaEmpresa — anclas + verificación contra el RESUMEN', () => {
  const meses = [1,2,3,4,5,6,7,8,9,10,11,12].map(m => new Date(2022, m - 1, 1));

  const ventas   = [100000, 110000, 90000, 105000, 98000, 120000, 115000, 108000, 99000, 102000, 130000, 140000];
  const costo    = [-40000, -44000, -36000, -42000, -39200, -48000, -46000, -43200, -39600, -40800, -52000, -56000];
  const utilB    = ventas.map((v, i) => v + costo[i]);
  const resOp    = utilB.map(u => Math.round(u * 0.6));
  const resEj    = resOp.map(r => Math.round(r * 0.8));
  const sumVentas = ventas.reduce((a, b) => a + b, 0);
  const sumResOp  = resOp.reduce((a, b) => a + b, 0);

  const rowsOk = [
    ['EMPRESA MENSUAL SAPEM'],
    [null, ...meses],
    ['Ventas Netas', ...ventas],
    ['Costo de Ventas', ...costo],
    ['Utilidad Bruta', ...utilB],
    ['Resultado Operativo', ...resOp],
    ['Resultado del Ejercicio', ...resEj],
  ];
  const wsOk = XLSX.utils.aoa_to_sheet(rowsOk, { cellDates: true });
  const rOk = _extraerERMensualHojaEmpresa(wsOk, sumVentas, sumResOp);
  assert(!!rOk, 'valida cuando la suma mensual de Ventas coincide con el total anual de referencia (RESUMEN)');
  if (rOk) {
    const fVentas = rOk.filas.find(f => f.conceptoStd === 'ventas');
    assert(!!fVentas && fVentas.valores[1] === 100000 && fVentas.valores[12] === 140000,
      'toma los valores mensuales de la fila "Ventas Netas" en el orden correcto de columnas');
    assert(!!rOk.filas.find(f => f.conceptoStd === 'costo'), 'incluye la fila de Costo como ancla directa');
    assert(!!rOk.filas.find(f => f.conceptoStd === 'utilBruta'), 'incluye la fila de Utilidad Bruta como ancla directa');
    const fResOp = rOk.filas.find(f => f.conceptoStd === 'resOperativo');
    assert(!!fResOp && fResOp.valores[1] === resOp[0], 'incluye Resultado Operativo (ancla) con su propio valor, no recalculado');
    assert(!!rOk.filas.find(f => f.conceptoStd === 'gastoOper'),
      'agrupa la diferencia entre Utilidad Bruta y Resultado Operativo en un único renglón "no discriminado" (no reclasifica gasto por gasto)');
    assert(!!rOk.filas.find(f => f.conceptoStd === 'resEjercicio'), 'incluye Resultado del Ejercicio (ancla) cuando difiere del Resultado Operativo');
  }

  // Referencia del RESUMEN muy distinta a la suma real de la hoja → no se
  // aplica nada (se prefiere quedarse sin detalle mensual antes que mostrar
  // un número mal verificado).
  const rBad = _extraerERMensualHojaEmpresa(wsOk, sumVentas * 3, sumResOp);
  assert(rBad === null, 'rechaza (devuelve null) cuando la suma mensual no coincide con la referencia dentro del 5% de tolerancia');

  // Caso real (Rioja Bus 2021): "Ventas Netas" es una sub-partida y aparece
  // ANTES que "Total Ingresos" (el total real, incluye subsidios) — ambas
  // filas matchean como ancla "fuerte" de ventas, pero están antes de la
  // fila de Costo. Debe preferirse la ÚLTIMA coincidencia fuerte antes del
  // límite (Costo/Utilidad Bruta), no la primera. ─────────────────────────
  const ventasNetas   = ventas.map(v => Math.round(v * 0.7));
  const totalIngresos = ventas; // el total real, coincide con la referencia del RESUMEN
  const rowsRB = [
    ['RIOJA BUS SAPEM'],
    [null, ...meses],
    ['Ventas Netas', ...ventasNetas],
    ['Total Ingresos', ...totalIngresos],
    ['Costo de Ventas', ...costo],
    ['Utilidad Bruta', ...utilB],
    ['Resultado Operativo', ...resOp],
    ['Resultado del Ejercicio', ...resEj],
  ];
  const wsRB = XLSX.utils.aoa_to_sheet(rowsRB, { cellDates: true });
  const rRB = _extraerERMensualHojaEmpresa(wsRB, sumVentas, sumResOp);
  assert(!!rRB, 'RIOJA BUS: valida usando la fila correcta de ventas (Total Ingresos), no la sub-partida');
  if (rRB) {
    const fVentasRB = rRB.filas.find(f => f.conceptoStd === 'ventas');
    assert(!!fVentasRB && fVentasRB.lbl === 'Total Ingresos',
      'prefiere la ÚLTIMA coincidencia fuerte de ventas antes del límite (Costo/Utilidad Bruta) — "Total Ingresos", no "Ventas Netas"');
  }
});

// ── _extraerERMensualHojaEmpresa: conciliación exacta contra el RESUMEN, y
//    Resultado Operativo validado por separado. Pedido real del usuario:
//    "los reportes deberían dar los mismos montos por empresa" — antes,
//    el extractor aceptaba a Ventas dentro de una tolerancia del 5% sin
//    corregir el residual (el total mostrado quedaba "cerca" del RESUMEN,
//    no exacto), y Resultado Operativo no se validaba en absoluto: contra
//    los 4 excel reales, 15 de 65 empresas con Resultado Operativo
//    mensual estaban a más del 30% del valor real (algunas con el signo
//    invertido). ──────────────────────────────────────────────────────
group('_extraerERMensualHojaEmpresa — concilia Ventas exacto y valida Resultado Operativo por separado', () => {
  const meses2 = [1,2,3,4,5,6,7,8,9,10,11,12].map(m => new Date(2022, m - 1, 1));
  const ventas2 = [100000, 110000, 90000, 105000, 98000, 120000, 115000, 108000, 99000, 102000, 130000, 140000];
  const resOp2  = [10000, 11000, 9000, 10500, 9800, 12000, 11500, 10800, 9900, 10200, 13000, 14000];
  const sumVentas2 = ventas2.reduce((a, b) => a + b, 0);
  const sumResOp2  = resOp2.reduce((a, b) => a + b, 0);

  const rows2 = [
    ['EMPRESA CONCILIACION'],
    [null, ...meses2],
    ['Ventas Netas', ...ventas2],
    ['Resultado Operativo', ...resOp2],
  ];
  const ws2 = XLSX.utils.aoa_to_sheet(rows2, { cellDates: true });

  // Referencia del RESUMEN levemente distinta a la suma real de la hoja
  // (redondeo/tolerancia, dentro del 5%) — el resultado debe coincidir
  // EXACTO con la referencia, no solo "cerca".
  const refVentasConResidual = sumVentas2 + 850; // ~0.07% de diferencia, dentro del 5%
  const rConResidual = _extraerERMensualHojaEmpresa(ws2, refVentasConResidual, sumResOp2);
  assert(!!rConResidual, 'valida cuando el residual está dentro de la tolerancia');
  if (rConResidual) {
    const fV = rConResidual.filas.find(f => f.conceptoStd === 'ventas');
    const sumaFinal = Object.values(fV.valores).reduce((a, b) => a + b, 0);
    assert(sumaFinal === refVentasConResidual,
      'concilia el residual como ajuste en el último mes — el total coincide EXACTO con el RESUMEN, no solo dentro del 5%');
    assert(fV.valores[1] === ventas2[0] && fV.valores[6] === ventas2[5],
      'no toca los meses que no son el último — el ajuste de conciliación solo se aplica al mes final');
  }

  // Resultado Operativo con signo invertido respecto del RESUMEN (bug real
  // encontrado en varias empresas: AGROANDINA 2024, AGROARAUCO 2023, etc.)
  // — debe rechazarse (la fila no debe aparecer), no conciliarse a la
  // fuerza con un ajuste gigante.
  const refResOpInvertido = -sumResOp2;
  const rResOpMalo = _extraerERMensualHojaEmpresa(ws2, sumVentas2, refResOpInvertido);
  assert(!!rResOpMalo, 'sigue devolviendo el detalle de Ventas aunque Resultado Operativo no valide');
  if (rResOpMalo) {
    assert(!rResOpMalo.filas.find(f => f.conceptoStd === 'resOperativo'),
      'omite la fila "resOperativo" cuando su ancla no coincide con el RESUMEN (signo invertido) — no la fuerza con un ajuste enorme');
    assert(!!rResOpMalo.filas.find(f => f.conceptoStd === 'ventas'),
      'conserva el detalle de Ventas, que sí validó — no rechaza toda la empresa por un problema aislado en Resultado Operativo');
  }

  // Resultado Operativo dentro de tolerancia → se concilia exacto, igual que Ventas.
  const rResOpOk = _extraerERMensualHojaEmpresa(ws2, sumVentas2, sumResOp2 - 500);
  if (rResOpOk) {
    const fR = rResOpOk.filas.find(f => f.conceptoStd === 'resOperativo');
    const sumaR = Object.values(fR.valores).reduce((a, b) => a + b, 0);
    assert(sumaR === sumResOp2 - 500, 'Resultado Operativo también se concilia exacto contra su propia referencia del RESUMEN cuando valida');
  }

  // Solo hay "Resultado del Ejercicio" (sin fila "Resultado Operativo" en
  // la hoja) — no debe hacerse pasar por Resultado Operativo, porque son
  // conceptos distintos que pueden diferir bastante.
  const rowsSoloEjercicio = [
    ['EMPRESA SOLO EJERCICIO'],
    [null, ...meses2],
    ['Ventas Netas', ...ventas2],
    ['Resultado del Ejercicio', ...resOp2],
  ];
  const wsSoloEjercicio = XLSX.utils.aoa_to_sheet(rowsSoloEjercicio, { cellDates: true });
  const rSoloEjercicio = _extraerERMensualHojaEmpresa(wsSoloEjercicio, sumVentas2, sumResOp2);
  assert(!!rSoloEjercicio, 'igual devuelve detalle cuando solo hay Resultado del Ejercicio (sin Resultado Operativo)');
  if (rSoloEjercicio) {
    assert(!rSoloEjercicio.filas.find(f => f.conceptoStd === 'resOperativo'),
      'no etiqueta "Resultado del Ejercicio" como si fuera Resultado Operativo');
    assert(!!rSoloEjercicio.filas.find(f => f.conceptoStd === 'resEjercicio'),
      'lo muestra bajo su propio concepto (resEjercicio), informativo pero no conciliado (no hay referencia de RESUMEN para ese concepto)');
  }
});

// ── _extraerERMensualHojaEmpresa: encabezados de mes como TEXTO y otras
//    variantes reales. Bug real: al recargar los RESUMEN históricos de
//    2021/2022, la enorme mayoría de las hojas por empresa no traían el
//    mes como fecha de Excel (único formato que el extractor reconocía al
//    principio) sino como texto — "ENERO", "01/2021" — o usaban rótulos
//    de resultado distintos ("UTILIDAD OPERATIVA" en vez de "RESULTADO
//    OPERATIVO"), o tenían un renglón de sección vacío ("INGRESOS", sin
//    datos) justo antes de la fila real de ventas. Todo eso hacía que casi
//    ninguna empresa de 2021/2022 mostrara detalle mensual en Evolución
//    Empresa, aunque el excel sí lo traía. ──────────────────────────────
group('_extraerERMensualHojaEmpresa — encabezados de mes como texto y variantes reales de rótulos', () => {
  const ventas = [100000, 110000, 90000, 105000, 98000, 120000, 115000, 108000, 99000, 102000, 130000, 140000];
  const costo  = [-40000, -44000, -36000, -42000, -39200, -48000, -46000, -43200, -39600, -40800, -52000, -56000];
  const utilB  = ventas.map((v, i) => v + costo[i]);
  const resOp  = utilB.map(u => Math.round(u * 0.6));
  const sumVentas = ventas.reduce((a, b) => a + b, 0);
  const sumResOp  = resOp.reduce((a, b) => a + b, 0);

  // Caso real AGROANDINA: los meses vienen como texto ("ENERO", "FEBRERO"...).
  const rowsTexto = [
    ['AGROANDINA'],
    [null, 'ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'],
    ['VENTAS NETAS', ...ventas],
    ['Costo de Ventas', ...costo],
    ['UTILIDAD BRUTA', ...utilB],
    ['UTILIDAD OPERATIVA', ...resOp],
  ];
  const wsTexto = XLSX.utils.aoa_to_sheet(rowsTexto);
  const rTexto = _extraerERMensualHojaEmpresa(wsTexto, sumVentas, sumResOp);
  assert(!!rTexto, 'reconoce un encabezado de meses en texto ("ENERO", "FEBRERO"...), no solo fechas de Excel');
  if (rTexto) {
    const fV = rTexto.filas.find(f => f.conceptoStd === 'ventas');
    assert(!!fV && fV.valores[1] === 100000 && fV.valores[12] === 140000,
      'asigna cada columna al número de mes real (1=enero..12=diciembre), no a su posición');
    assert(!!rTexto.filas.find(f => f.conceptoStd === 'resOperativo'),
      '"UTILIDAD OPERATIVA" se reconoce como Resultado Operativo (no solo "RESULTADO OPERATIVO")');
  }

  // Caso real ALFA SAU: los meses vienen como "MM/AAAA", con una columna de
  // subtotal intercalada a mitad de año que NO es un mes y debe ignorarse.
  const rowsMMYYYY = [
    ['ALFA SAU'],
    ['DETALLE', '01/2021', '02/2021', '03/2021', '04/2021', '05/2021', '06/2021', 'SUBTOTAL AL 30/06/2021', '07/2021', '08/2021', '09/2021', '10/2021', '11/2021', '12/2021'],
    ['TOTAL VENTAS', ventas[0], ventas[1], ventas[2], ventas[3], ventas[4], ventas[5], ventas[0]+ventas[1]+ventas[2]+ventas[3]+ventas[4]+ventas[5], ventas[6], ventas[7], ventas[8], ventas[9], ventas[10], ventas[11]],
    ['Costo de Ventas', costo[0], costo[1], costo[2], costo[3], costo[4], costo[5], null, costo[6], costo[7], costo[8], costo[9], costo[10], costo[11]],
    ['UTILIDAD BRUTA', utilB[0], utilB[1], utilB[2], utilB[3], utilB[4], utilB[5], null, utilB[6], utilB[7], utilB[8], utilB[9], utilB[10], utilB[11]],
    ['RESULTADO OPERATIVO', resOp[0], resOp[1], resOp[2], resOp[3], resOp[4], resOp[5], null, resOp[6], resOp[7], resOp[8], resOp[9], resOp[10], resOp[11]],
  ];
  const wsMMYYYY = XLSX.utils.aoa_to_sheet(rowsMMYYYY);
  const rMMYYYY = _extraerERMensualHojaEmpresa(wsMMYYYY, sumVentas, sumResOp);
  assert(!!rMMYYYY, 'reconoce un encabezado de meses en formato "MM/AAAA" (texto), ignorando la columna de subtotal intercalada');
  if (rMMYYYY) {
    const fV = rMMYYYY.filas.find(f => f.conceptoStd === 'ventas');
    assert(!!fV && fV.valores[1] === ventas[0] && fV.valores[6] === ventas[5] && fV.valores[7] === ventas[6],
      'no desalinea los meses posteriores al subtotal intercalado (julio sigue siendo el mes 7, no el 8vo dato)');
  }

  // Caso real AGROARAUCO: la fila "INGRESOS" es solo un título de sección
  // (sin datos) y aparece antes que la fila real "Ventas Brutas" — debe
  // preferirse la fila con datos, no la primera que matchea el regex.
  const rowsHeaderVacio = [
    ['AGROARAUCO'],
    [null, 'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'],
    ['INGRESOS'], // título de sección, sin ningún valor
    ['Ventas Brutas', ...ventas],
    ['Costo de Ventas', ...costo],
    ['UTILIDAD BRUTA', ...utilB],
    ['UTILIDAD OPERATIVA', ...resOp],
  ];
  const wsHeaderVacio = XLSX.utils.aoa_to_sheet(rowsHeaderVacio);
  const rHeaderVacio = _extraerERMensualHojaEmpresa(wsHeaderVacio, sumVentas, sumResOp);
  assert(!!rHeaderVacio, 'no se rinde cuando la primera fila que menciona "ingresos" es un título de sección sin datos');
  if (rHeaderVacio) {
    const fV = rHeaderVacio.filas.find(f => f.conceptoStd === 'ventas');
    assert(!!fV && fV.lbl === 'Ventas Brutas',
      'prefiere la fila que realmente tiene los montos mensuales ("Ventas Brutas") en vez del título de sección vacío ("INGRESOS")');
  }
});

// ── _resolveHojaEmpresa: el texto del encabezado de columna del RESUMEN no
//    siempre coincide letra por letra con el nombre real de la pestaña de
//    esa empresa en el mismo libro. Bug real: "BR. SERVICIOS FINANCIEROS",
//    "PUERTAS DE SOL" y "VIENTOS ARAUCO RENOVABLE" (texto del RESUMEN) no
//    encontraban su pestaña ("BR SERVICIOS FINANCIEROS", "PUERTAS DEL
//    SOL", "VIENTOS DE ARAUCO RENOVABLES") con el match exacto original,
//    así que esas empresas nunca llegaban a intentar el detalle mensual. ─
group('_resolveHojaEmpresa — nombre de columna del RESUMEN vs. nombre real de la pestaña', () => {
  const wb = wbFromSheets({
    RESUMEN: [['x']],
    'BR SERVICIOS FINANCIEROS': [['dato']],
    'PUERTAS DEL SOL': [['dato']],
    'VIENTOS DE ARAUCO RENOVABLES': [['dato']],
  });
  assert(_resolveHojaEmpresa(wb, 'BR. SERVICIOS FINANCIEROS') === wb.Sheets['BR SERVICIOS FINANCIEROS'],
    'ignora diferencias de puntuación ("BR." vs "BR")');
  assert(_resolveHojaEmpresa(wb, 'PUERTAS DE SOL') === wb.Sheets['PUERTAS DEL SOL'],
    'tolera artículos distintos ("DE SOL" vs "DEL SOL")');
  assert(_resolveHojaEmpresa(wb, 'VIENTOS ARAUCO RENOVABLE') === wb.Sheets['VIENTOS DE ARAUCO RENOVABLES'],
    'tolera singular/plural y un artículo faltante ("VIENTOS ARAUCO RENOVABLE" vs "VIENTOS DE ARAUCO RENOVABLES")');
  assert(!_resolveHojaEmpresa(wb, 'UNA EMPRESA QUE NO EXISTE'),
    'no inventa una coincidencia para un nombre sin ninguna pestaña razonablemente parecida');
});

// ── tryParseERLibroMayor: exportación contable tipo "libro mayor" (formato
//    real: Parque Eólico Arauco, ER al 30/06/2026) — encabezado "Cuentas |
//    Cuenta | <meses>", secciones con código ("411 - VENTAS NETAS DE
//    SERVICIOS"), cuentas de detalle debajo, y una fila "Total <rubro>" que
//    agrega ese detalle, hasta cerrar en "Total General". Usa convención de
//    saldo contable (débito positivo): Ventas queda en negativo y
//    Costos/Gastos en positivo — inversa a la convención de presentación
//    del dashboard — así que el parser debe invertir el signo. ───────────
group('tryParseERLibroMayor — exportación contable con filas "Total <rubro>" y signo débito-positivo', () => {
  const rowsBase = [
    ['', 'Periodo base', '', ':', '06/2026'],
    [],
    ['Cuentas', 'Cuenta', 'Enero', 'Febrero'],
    [],
    ['411 - VENTAS NETAS DE SERVICIOS'],
    [],
    ['', '41100000 - PRESTACION DE SERVI', -1000, -1100],
    [],
    ['Total VENTAS NETAS DE SERVICIOS', '', -1000, -1100],
    [],
    ['412 - COSTOS OPERATIVOS'],
    [],
    ['', '51103002 - Combustibles y lubr', 300, 320],
    [],
    ['Total COSTOS OPERATIVOS', '', 300, 320],
    [],
    ['51102 - GASTOS  DE ADMINISTRACION'],
    [],
    ['', '51101001 - Sueldos y jornales', 400, 410],
    [],
    ['Total GASTOS  DE ADMINISTRACION', '', 400, 410],
    [],
    ['51103 - GASTOS OPERATIVOS'],
    [],
    ['', '51102014 - Repuestos e insumos', 50, 60],
    [],
    ['Total GASTOS OPERATIVOS', '', 50, 60],
    [],
    ['51108 - RESULTADO FINANCIERO Y POR TENENCIA'],
    [],
    ['', '42205000 - Intereses  prestamo', 80, 90],
    [],
    ['Total RESULTADO FINANCIERO Y POR TENENCIA', '', 80, 90],
    [],
    ['51109 - Otros Ingresos- Egresos'],
    [],
    ['', '42600000 - INGRESOS VARIOS', -20, -25],
    [],
    ['Total Otros Ingresos- Egresos', '', -20, -25],
    [],
    ['Total General', '', -190, -245],
  ];
  const wb = wbFromSheets({ 'Pag.1': rowsBase });
  const r = tryParseERLibroMayor(wb, 'PEA__ESTADO_DE_RDO_AL_30062026.xlsx');
  assert(!!r, 'reconoce el formato por el encabezado distintivo "Cuentas | Cuenta"');
  if (r) {
    assert(r.anio === '2026', 'toma el año del período base ("Periodo base : 06/2026"), no del nombre de archivo');
    assert(r.mesesCount === 2, 'detecta los meses presentes aunque sean menos de 8 (carga parcial de año)');
    const fV = r.filas.find(f => f.conceptoStd === 'ventas');
    assert(!!fV && fV.valores[1] === 1000 && fV.valores[2] === 1100,
      'invierte el signo de Ventas: negativo (débito-positivo) en el origen → positivo en el dashboard');
    const fC = r.filas.find(f => f.conceptoStd === 'costo');
    assert(!!fC && fC.valores[1] === -300, 'invierte el signo de Costos: positivo en el origen → negativo (se resta) en el dashboard');
    const fGA = r.filas.find(f => f.conceptoStd === 'gastoAdm');
    assert(!!fGA && fGA.valores[1] === -400, 'clasifica "GASTOS DE ADMINISTRACION" como gastoAdm, con signo invertido');
    const fGO = r.filas.find(f => f.conceptoStd === 'gastoCom');
    assert(!!fGO && fGO.valores[1] === -50, 'clasifica "GASTOS OPERATIVOS" (sección aparte de Costos Operativos) como gastoCom');
    const fRF = r.filas.find(f => f.conceptoStd === 'resFinanciero');
    assert(!!fRF && fRF.valores[1] === -80, 'clasifica "RESULTADO FINANCIERO Y POR TENENCIA" como resFinanciero');
    const fOI = r.filas.find(f => f.conceptoStd === 'otrosIngresos');
    assert(!!fOI && fOI.valores[1] === 20, 'clasifica "Otros Ingresos- Egresos" como otrosIngresos, con signo invertido');
    const fGE = r.filas.find(f => f.conceptoStd === 'resEjercicio');
    assert(!!fGE && fGE.valores[1] === 190, '"Total General" se toma como Resultado del Ejercicio (bottom line), signo invertido');
    assert(r.filas.every(f => !/^\d+\s*-\s*/.test(f.lbl)),
      'no confunde las filas de sección con código ("411 - VENTAS...") ni las cuentas de detalle con las filas "Total" — solo usa estas últimas');
  }
});

group('tryParseERLibroMayor — no reconoce formatos que no calzan (evita falsos positivos)', () => {
  const wbSinTotalGeneral = wbFromSheets({ 'Pag.1': [
    ['Cuentas', 'Cuenta', 'Enero', 'Febrero'],
    ['411 - VENTAS NETAS DE SERVICIOS'],
    ['', '41100000 - PRESTACION DE SERVI', -1000, -1100],
    ['Total VENTAS NETAS DE SERVICIOS', '', -1000, -1100],
  ] });
  assert(!tryParseERLibroMayor(wbSinTotalGeneral, 'x.xlsx'),
    'rechaza si no llega a cerrar en "Total General" (evita falsos positivos con otras planillas contables)');

  assert(!tryParseERLibroMayor(wbFromSheets({ RESUMEN: [['CONCEPTO', 'ENERO', 'FEBRERO']] }), 'x.xlsx'),
    'no reconoce una hoja RESUMEN de ER anual estándar (sin el encabezado "Cuentas | Cuenta")');
});

// ── tryParseERLibroMayorResumen: variante CONDENSADA del libro mayor
//    (formato real: Kallpa, Arauco Energía — ER al 30/06/2026) — no hay
//    detalle de cuentas ni filas "Total X" intermedias, cada renglón bajo
//    el encabezado "Cuentas" ES directamente un rubro, hasta "TOTAL
//    GENERAL". A diferencia del resto de los rubros, el renglón de
//    Ingresos viene en positivo (no en convención débito) — el parser
//    reconstruye "TOTAL GENERAL" a partir de los demás rubros ya
//    convertidos y solo acepta el archivo si concilia exacto. Bug real
//    detectado al construir este parser: "OTROS INGRESOS- EGRESOS"
//    contiene la palabra "INGRESOS" y quedaba mal clasificado como
//    Ventas si esa regla se evaluaba primero. ─────────────────────────
group('tryParseERLibroMayorResumen — variante condensada (un renglón por rubro, sin detalle de cuentas)', () => {
  // Enero: 1000 - (100+200+50+30+10+(-5)) = 615
  // Febrero: 2000 - (150+250+60+(-40)+20+15) = 1545
  const rows = [
    ['', 'Periodo base', '', ':', '06/2026'],
    [],
    ['Cuentas', 'ene-26', new Date(2026, 1, 1)],
    [],
    ['426 - INGESOS VARIOS', 1000, 2000],
    ['412 - COSTOS OPERATIVOS', 100, 150],
    ['51102 - GASTOS  DE ADMINISTRACION', 200, 250],
    ['51103 - GASTOS OPERATIVOS', 50, 60],
    ['51108 - RESULTADO FINANCIERO Y POR TENENCIA', 30, -40],
    ['6 - IMPUESTO A LAS GANANCIAS', 10, 20],
    ['OTROS INGRESOS- EGRESOS', -5, 15],
    ['TOTAL GENERAL', 615, 1545],
  ];
  const wb = wbFromSheets({ 'Pag.1': rows });
  const r = tryParseERLibroMayorResumen(wb, 'KALLPA__ESTADO_DE_RESULTADOS_al_30062026.xls');
  assert(!!r, 'reconoce el formato condensado por el encabezado "Cuentas" seguido directo de los meses (sin columna "Cuenta")');
  if (r) {
    assert(r.anio === '2026', 'toma el año del período base, no del nombre de archivo');
    assert(r.mesesCount === 2, 'reconoce meses en texto abreviado ("ene-26") y como fecha de Excel en la misma fila de encabezado');
    const fV = r.filas.find(f => f.conceptoStd === 'ventas');
    assert(!!fV && fV.valores[1] === 1000 && fV.valores[2] === 2000,
      'toma Ingresos tal cual (sin invertir el signo) — es la excepción: acá sí viene positivo de origen');
    const fC = r.filas.find(f => f.conceptoStd === 'costo');
    assert(!!fC && fC.valores[1] === -100, 'invierte el signo de Costos (positivo en el origen → negativo en el dashboard)');
    const fOI = r.filas.find(f => f.conceptoStd === 'otrosIngresos');
    assert(!!fOI && fOI.valores[1] === 5 && fOI.valores[2] === -15,
      'clasifica "OTROS INGRESOS- EGRESOS" como otrosIngresos (no como ventas, aunque su rótulo contenga "INGRESOS") y le invierte el signo');
    const fGE = r.filas.find(f => f.conceptoStd === 'resEjercicio');
    assert(!!fGE && fGE.valores[1] === 615 && fGE.valores[2] === 1545,
      '"TOTAL GENERAL" se toma tal cual (ya viene en la convención correcta) como Resultado del Ejercicio');
  }
});

group('tryParseERLibroMayorResumen — rechaza si no concilia contra TOTAL GENERAL (evita importar con el signo equivocado)', () => {
  const rows = [
    ['Cuentas', 'ene-26'],
    ['426 - INGESOS VARIOS', 1000],
    ['412 - COSTOS OPERATIVOS', 100],
    ['51102 - GASTOS  DE ADMINISTRACION', 200],
    ['TOTAL GENERAL', 999999], // no coincide con 1000 - (100+200) = 700
  ];
  assert(!tryParseERLibroMayorResumen(wbFromSheets({ 'Pag.1': rows }), 'x.xls'),
    'no importa el archivo si la suma de los rubros no reconstruye "TOTAL GENERAL" — evita adivinar el signo');
});

group('tryParseERLibroMayorResumen — no se confunde con el formato con detalle de cuentas (tryParseERLibroMayor)', () => {
  const rows = [
    ['Cuentas', 'Cuenta', 'Enero', 'Febrero'],
    ['411 - VENTAS NETAS DE SERVICIOS'],
    ['', '41100000 - PRESTACION DE SERVI', -1000, -1100],
    ['Total VENTAS NETAS DE SERVICIOS', '', -1000, -1100],
    ['Total General', '', -1000, -1100],
  ];
  assert(!tryParseERLibroMayorResumen(wbFromSheets({ 'Pag.1': rows }), 'x.xlsx'),
    'no reconoce el formato con columna "Cuenta" (detalle de cuentas) — queda exclusivo de tryParseERLibroMayor');
});

// ── _renombrarEmpresa / _fusionarEmpresas / _eliminarEmpresa: autoridades
//    (director/presidente, síndico, etc.) es un historial por empresa
//    agregado en este mismo turno — tiene que seguir el mismo criterio de
//    migración que ya usan fact/sit/personal/notas/etc. en estos tres
//    flujos: renombrar reapunta, fusionar reapunta (es un historial, no
//    hay nada que "sumar"), eliminar limpia los registros de esa empresa. ─
group('autoridades — sigue el mismo criterio de migración que el resto de los módulos al renombrar/fusionar/eliminar una empresa', () => {
  function nuevoContexto() {
    const src = [
      extractFn(html, '_renombrarEmpresa'),
      extractFn(html, '_fusionarEmpresas'),
      extractFn(html, '_eliminarEmpresa'),
    ].join('\n');
    const factory = new Function(
      'let DATA = {fact:{},sit:{},personal:[],notas:{},presupuesto:{},fact_mensual:{},res_mensual:{},er_mensual:{},autoridades:[]};\n' +
      'let rubros = {};\nlet _balances = {};\n' +
      src + '\n' +
      'return { _renombrarEmpresa, _fusionarEmpresas, _eliminarEmpresa, getDATA: () => DATA };'
    );
    return factory();
  }

  // Renombrar: reapunta el campo empresa de todos los registros del historial
  const ctx1 = nuevoContexto();
  ctx1.getDATA().autoridades = [
    { empresa:'LRT', nombre:'Juan', apellido:'Perez', cargo:'Presidente', desde:'2020-01-01', hasta:'' },
    { empresa:'LRT', nombre:'Ana',  apellido:'Diaz',  cargo:'Sindico',    desde:'2019-01-01', hasta:'' },
  ];
  ctx1._renombrarEmpresa('LRT', 'LA RIOJA TELECOMUNICACIONES');
  assert(ctx1.getDATA().autoridades.every(a => a.empresa === 'LA RIOJA TELECOMUNICACIONES') && ctx1.getDATA().autoridades.length === 2,
    '_renombrarEmpresa reapunta todos los registros de autoridades de la empresa renombrada (se conserva el historial completo)');

  // Fusionar: reapunta los registros del origen al destino — no se "suman"
  // valores (es un historial de personas, no un total), se conservan ambos
  const ctx2 = nuevoContexto();
  ctx2.getDATA().autoridades = [
    { empresa:'LRT', nombre:'Juan', apellido:'Perez', cargo:'Presidente', desde:'2020-01-01', hasta:'' },
    { empresa:'LA RIOJA TELECOMUNICACIONES', nombre:'Pedro', apellido:'Lopez', cargo:'Director', desde:'2021-01-01', hasta:'' },
  ];
  ctx2._fusionarEmpresas('LRT', 'LA RIOJA TELECOMUNICACIONES');
  assert(ctx2.getDATA().autoridades.length === 2 && ctx2.getDATA().autoridades.every(a => a.empresa === 'LA RIOJA TELECOMUNICACIONES'),
    '_fusionarEmpresas reapunta los registros del origen al destino y conserva los del destino — ambos historiales coexisten');

  // Eliminar empresa: limpia sus registros de autoridades, no toca los de otras
  const ctx3 = nuevoContexto();
  ctx3.getDATA().autoridades = [
    { empresa:'LRT', nombre:'Juan', apellido:'Perez', cargo:'Presidente', desde:'2020-01-01', hasta:'' },
    { empresa:'PEA', nombre:'Carlos', apellido:'Lopez', cargo:'Presidente', desde:'2023-01-01', hasta:'' },
  ];
  ctx3._eliminarEmpresa('LRT');
  const restantes = ctx3.getDATA().autoridades;
  assert(restantes.length === 1 && restantes[0].empresa === 'PEA',
    '_eliminarEmpresa borra solo los registros de autoridades de la empresa eliminada, no los de otras');
});

console.log(`\n${pass} OK, ${fail} FALLÓ${fail ? ' — revisar antes de publicar' : ''}`);
process.exit(fail ? 1 : 0);
