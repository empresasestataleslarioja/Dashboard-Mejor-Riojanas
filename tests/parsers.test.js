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
} = loadFns([
  '_esIngExcluido',
  '_matchEmpresaPortafolio',
  'tryParseVentasGastos',
  'tryParseERColumnasMeses',
  '_parsearPestanasMensuales',
  'tryParseERMatricial',
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
const { _erDesdeMonthly, setData } = new Function(
  erConsts + '\n' + erDesdeMonthlySrc + '\n' +
  'let DATA = {};\nfunction setData(d) { DATA = d; }\n' +
  'return { _erDesdeMonthly, setData };'
)();

// tryParseERNativo también escribe sobre el global DATA y llama a funciones
// de refresco de UI (computeTotals/populateSelects/rebuildActive) que no
// existen fuera del dashboard — se stubean como no-op, solo interesa qué
// devuelve el parser y qué queda en DATA.er.
const tryParseERNativoSrc = extractFn(html, 'tryParseERNativo');
const { tryParseERNativo, getERNativoData, resetERNativoData } = new Function('XLSX',
  'let DATA = {fact:{},sit:{},er:{}};\n' +
  'function computeTotals(){}\nfunction populateSelects(){}\nfunction rebuildActive(){}\n' +
  tryParseERNativoSrc + '\n' +
  'function getERNativoData(){ return DATA; }\n' +
  'function resetERNativoData(){ DATA.fact={}; DATA.sit={}; DATA.er={}; }\n' +
  'return { tryParseERNativo, getERNativoData, resetERNativoData };'
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

console.log(`\n${pass} OK, ${fail} FALLÓ${fail ? ' — revisar antes de publicar' : ''}`);
process.exit(fail ? 1 : 0);
