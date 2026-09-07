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
const crossCheckSrc = extractFn(html, '_crossCheckEREmpresaAnio');
const { _erDesdeMonthly, _crossCheckEREmpresaAnio, setData } = new Function(
  erConsts + '\n' + erDesdeMonthlySrc + '\n' + crossCheckSrc + '\n' +
  'let DATA = {};\nfunction setData(d) { DATA = d; }\n' +
  'return { _erDesdeMonthly, _crossCheckEREmpresaAnio, setData };'
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

console.log(`\n${pass} OK, ${fail} FALLÓ${fail ? ' — revisar antes de publicar' : ''}`);
process.exit(fail ? 1 : 0);
