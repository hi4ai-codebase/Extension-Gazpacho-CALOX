/*
 * cierre-core.js — GENERADO por `npm run build:extension`. ¡NO EDITAR A MANO!
 * Fuente: src/transform/ (las mismas reglas que usa `npm run transformar`).
 * Generado: 2026-11-19T13:57:35.107Z
 */
window.CierreCaloxCore = (() => {
'use strict';
// ───── src/transform/build-closing-rows.js ─────
/**
 * build-closing-rows.js — Reglas del cierre: convierte los datos de Gazpacho
 * en las filas de la hoja CIERRE (estructura de la "Hoja1" de referencia).
 *
 * ¿Por qué este archivo NO importa nada (ni ExcelJS, ni fs, ni Node)?
 *   Porque son "reglas de negocio puras": reciben datos y devuelven datos.
 *   Así el MISMO código sirve en dos sitios:
 *     - en Node (run-transform.js), leyendo el Excel descargado;
 *     - en la futura extensión de Chrome, con los datos de la API de Gazpacho.
 *   Y además se puede probar con datos inventados, sin Excel ni Gazpacho.
 *
 * Vocabulario:
 *   - pedido  (order) → una fila de "HOJA DE PEDIDOS CALOX (...)": quién, cuándo, droguería.
 *   - línea   (line)  → una fila de "TABLA DE PEDIDO" o "TABLA DE PEDIDO_3": un producto.
 *   - original        → el informe móvil del vendedor del que salió un pedido web
 *                       (bloque "Información de Informe Original" en Gazpacho).
 *
 * Tipos de pedido: Gazpacho tiene dos formularios y cada uno genera 2 hojas:
 *   - cortos:   "HOJA DE PEDIDOS ... CORTOS"   + "TABLA DE PEDIDO"
 *   - masivos:  "HOJA DE PEDIDOS ... MASIVOS"  + "TABLA DE PEDIDO_3" (plantilla de 316 productos)
 *   Aquí se tratan igual: todos van a la misma hoja CIERRE.
 */

// ── Estructura de la hoja CIERRE ────────────────────────────────────────────
// Cada columna dice su título y de dónde sale su valor:
//   'pedido'   → misma columna en la hoja de pedidos
//   'producto' → misma columna en la tabla de productos
//   'original' → dato del informe original: en pedidos Móvil, del propio pedido;
//                en pedidos Web, lo traerá la extensión (por ahora vacío)
// Para cambiar el orden o añadir una columna, basta con tocar esta lista.
const CLOSING_COLUMNS = [
     { header: 'Report ID', source: 'pedido' },
     { header: 'Open Time', source: 'pedido', isDate: true },
     { header: 'First Name', source: 'pedido' },
     { header: 'Last Name', source: 'pedido' },
     { header: 'ID de Informe', source: 'original' },
     { header: 'Apertura', source: 'original', isDate: true },
     { header: 'Usuario', source: 'original' },
     { header: 'External ID', source: 'pedido' },
     { header: 'DROGUERIA', source: 'pedido' },
     { header: 'Descripción', source: 'producto' },
     { header: 'Material', source: 'producto' },
     { header: 'Linea', source: 'producto' },
     { header: 'Cantidad', source: 'producto' },
     { header: 'Cantidades Pedidas', source: 'producto' },
     { header: 'Descuento del Producto %', source: 'producto' },
     { header: 'Farmacia', source: 'producto' },
     { header: 'Numero de Factura o Pedido', source: 'producto' },
     { header: 'Observación', source: 'pedido' },
];

// Nombres de columnas de Gazpacho que usan las reglas (no se repiten como texto suelto).
const ORDER_ID = 'Report ID';         // en la hoja de pedidos
const ORDER_OPEN_TIME = 'Open Time';
const ORDER_SOURCE = 'Source';        // "Mobile(null)" o "Web" (se usa para decidir, NO sale en el Excel)
const ORDER_TABLE_SUMMARY = 'TABLA DE PEDIDO'; // texto "Table 3 rows"
const LINE_REPORT_ID = 'Report';      // en la tabla de productos
const LINE_QUANTITY = 'Cantidad';

// ── Pequeñas utilidades ─────────────────────────────────────────────────────

/** Convierte 2681765 o "2681765" en el número 2681765; cualquier otra cosa → null. */
function toReportId(value) {
     if (typeof value === 'number' && Number.isInteger(value)) return value;
     if (typeof value === 'string' && /^\d+$/.test(value.trim())) return Number(value.trim());
     return null;
}

/** ¿Tiene cantidad? Regla del cierre: Cantidad NO puede ser null ni texto vacío. */
function hasQuantity(line) {
     const q = line[LINE_QUANTITY];
     return q !== null && q !== undefined && String(q).trim() !== '';
}

/** "Table 3 rows" → 3. Si el texto no tiene ese formato → null. */
function expectedLineCount(order) {
     const match = /Table\s+(\d+)\s+rows?/i.exec(String(order[ORDER_TABLE_SUMMARY] ?? ''));
     return match ? Number(match[1]) : null;
}

// ── Rangos de fechas ────────────────────────────────────────────────────────
// Importante: Excel guarda las fechas SIN zona horaria ("13/11/2026 10:05").
// ExcelJS las entrega como Date cuya hora UTC es esa misma hora de pared.
// Por eso todos los rangos se construyen con Date.UTC: comparamos "hora de
// pared" con "hora de pared" y no hay desfases de -4 h.

/** Mes completo: monthRange(2026, 10) → del 01/10/2026 00:00 al 01/11/2026 00:00 (sin incluir). */
function monthRange(year, month) {
     return { from: new Date(Date.UTC(year, month - 1, 1)), to: new Date(Date.UTC(year, month, 1)) };
}

/** Días completos: dayRange('01/10/2026', '15/10/2026') incluye todo el 15. */
function dayRange(fromDDMMYYYY, toDDMMYYYY = fromDDMMYYYY) {
     const parse = (text) => {
          const m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(text);
          if (!m) throw new Error(`Fecha no válida: "${text}". Formato esperado: DD/MM/AAAA.`);
          return Date.UTC(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
     };
     const from = parse(fromDDMMYYYY);
     const to = parse(toDDMMYYYY) + 24 * 60 * 60 * 1000; // hasta el final de ese día
     if (to <= from) throw new Error('La fecha "hasta" es anterior a la fecha "desde".');
     return { from: new Date(from), to: new Date(to) };
}

const isInRange = (date, range) => !range || (date >= range.from && date < range.to);

// ── Informe original: según la FUENTE (Source) del pedido ───────────────────
// En Gazpacho hay dos tipos de informe:
//   - Fuente "Móvil" (Source "Mobile(null)"): lo hizo el vendedor desde el teléfono.
//     ESE pedido ya ES el informe original.
//   - Fuente "Web" (Source "Web"): lo cargó la oficina (Calox Administrator) a
//     partir del informe móvil de un vendedor. Su Report ID es el del "Informe
//     General"; el vendedor y el "ID de Informe" están en el "Informe Original"
//     (Móvil), que NO viene en el Excel: lo traerá la extensión (`originals`).

/** 'mobile' | 'web' | 'unknown', a partir de la columna Source. */
function orderSource(order) {
     const source = String(order[ORDER_SOURCE] ?? '').trim().toLowerCase();
     if (source.startsWith('mobile') || source.startsWith('móvil') || source.startsWith('movil')) return 'mobile';
     if (source.startsWith('web')) return 'web';
     return 'unknown';
}

/**
 * Datos del informe original de un pedido: { Usuario, 'ID de Informe', Apertura } o null.
 *   - Móvil → del propio pedido (Usuario = First Name + Last Name).
 *   - Web   → del mapa `originals` (extensión); si aún no está, null (columnas vacías).
 */
function originalFor(order, id, originals) {
     if (orderSource(order) === 'mobile') {
          const usuario = [order['First Name'], order['Last Name']]
               .map((part) => String(part ?? '').trim()).filter(Boolean).join(' ');
          return { Usuario: usuario || null, 'ID de Informe': id, Apertura: order[ORDER_OPEN_TIME] ?? null };
     }
     return originals.get(id) ?? null;
}

// ── Regla principal ─────────────────────────────────────────────────────────
/**
 * Construye las filas de la hoja CIERRE.
 *
 * @param {object}   input
 * @param {object[]} input.orders     filas de las hojas de pedidos ({ 'Report ID': 2681765, 'Open Time': Date, ... })
 * @param {object[]} input.lines      filas de las tablas de productos (incluidas las "Table", se descartan aquí)
 * @param {{from: Date, to: Date}} [input.range]  filtro por Open Time; sin rango = todos
 * @param {Map<number, object>} [input.originals] informe original de los pedidos WEB, por
 *                                    Report ID ({ Usuario, 'ID de Informe', Apertura }); opcional.
 *                                    Los pedidos Móvil no lo necesitan (son su propio original).
 * @returns {{ rows: object[], summary: object, problems: string[] }}
 *   - rows:     una por producto con cantidad (o una sola fila si el pedido no tiene ninguno)
 *   - summary:  números para el resumen en pantalla
 *   - problems: incoherencias graves; si hay alguna, NO se debe generar el archivo
 */
function buildClosingRows({ orders, lines, range, originals = new Map() }) {
     const problems = [];

     // 1) Agrupar las líneas por pedido. Las filas "Table" (separadores) no tienen
     //    un Report numérico y se cuentan aparte.
     const linesByOrder = new Map();
     let separatorRows = 0;
     for (const line of lines) {
          const id = toReportId(line[LINE_REPORT_ID]);
          if (id === null) { separatorRows++; continue; }
          if (!linesByOrder.has(id)) linesByOrder.set(id, []);
          linesByOrder.get(id).push(line);
     }

     // 2) Controles de calidad: cada pedido debe tener exactamente las líneas que
     //    anuncia ("Table N rows") y no debe haber líneas sin pedido.
     const orderIds = new Set();
     for (const order of orders) {
          const id = toReportId(order[ORDER_ID]);
          if (id === null) { problems.push(`Pedido sin Report ID válido: ${JSON.stringify(order[ORDER_ID])}`); continue; }
          if (orderIds.has(id)) problems.push(`Pedido ${id} repetido.`);
          orderIds.add(id);

          const expected = expectedLineCount(order);
          const actual = linesByOrder.get(id)?.length ?? 0;
          if (expected !== null && expected !== actual) {
               problems.push(`Pedido ${id}: anuncia ${expected} productos pero la tabla tiene ${actual}.`);
          }
          if (!(order[ORDER_OPEN_TIME] instanceof Date)) {
               problems.push(`Pedido ${id}: "Open Time" no es una fecha.`);
          }
     }
     for (const id of linesByOrder.keys()) {
          if (!orderIds.has(id)) problems.push(`Hay productos del pedido ${id}, que no existe en las hojas de pedidos.`);
     }

     // 3) Construir las filas, pedido a pedido, conservando el orden de Gazpacho.
     const rows = [];
     const summary = {
          ordersTotal: orders.length,
          ordersInRange: 0,
          productLinesKept: 0,
          productLinesWithoutQuantity: 0,
          separatorRows,
          ordersWithoutProducts: [], // pedidos sin ningún producto con cantidad
          ordersMobile: 0,           // fuente Móvil (Usuario relleno desde el propio pedido)
          ordersWeb: 0,              // fuente Web (Calox Administrator)
          ordersWebWithoutOriginal: 0, // Web cuyo informe original aún no tenemos
          webPendingIds: [],         // sus Report ID (la lista que se le pasa a la extensión)
          ordersUnknownSource: [],   // fuente desconocida: se avisa
     };

     for (const order of orders) {
          const id = toReportId(order[ORDER_ID]);
          if (id === null || !isInRange(order[ORDER_OPEN_TIME], range)) continue;
          summary.ordersInRange++;

          const source = orderSource(order);
          if (source === 'mobile') summary.ordersMobile++;
          else if (source === 'web') summary.ordersWeb++;
          else summary.ordersUnknownSource.push(id);

          const original = originalFor(order, id, originals);
          if (source === 'web' && !original) {
               summary.ordersWebWithoutOriginal++;
               summary.webPendingIds.push(id);
          }

          const allLines = linesByOrder.get(id) ?? [];
          const kept = allLines.filter(hasQuantity);
          summary.productLinesKept += kept.length;
          summary.productLinesWithoutQuantity += allLines.length - kept.length;

          if (kept.length === 0) {
               // No perdemos el pedido en silencio: sale una fila sin producto y se avisa.
               summary.ordersWithoutProducts.push(id);
               rows.push(makeRow(order, null, original));
          } else {
               for (const line of kept) rows.push(makeRow(order, line, original));
          }
     }

     return { rows, summary, problems };
}

/** Una fila = objeto { 'Report ID': ..., 'Open Time': ..., ... } en el orden de CLOSING_COLUMNS. */
function makeRow(order, line, original) {
     const row = {};
     for (const { header, source } of CLOSING_COLUMNS) {
          const from = source === 'pedido' ? order : source === 'producto' ? line : original;
          row[header] = from?.[header] ?? null;
     }
     return row;
}

// ───── src/transform/originals.js ─────
/**
 * originals.js — Convierte lo que recoge la extensión de Chrome (el "informe
 * original" de cada pedido Web) en los datos de las columnas Usuario,
 * ID de Informe y Apertura.
 *
 * Igual que build-closing-rows.js, NO importa nada: son reglas puras.
 *
 * Formato de los archivos que intercambian Node y la extensión
 * ─────────────────────────────────────────────────────────────
 * 1) Lista de pedidos pendientes (la escribe `npm run transformar`, la lee la extensión):
 *      { "tipo": "gazpacho-pedidos-web-pendientes", "version": 1,
 *        "periodo": "octubre 2026", "pedidos": [2682106, ...] }
 *
 * 2) Resultado de la extensión (lo descarga la extensión, lo lee `npm run transformar`):
 *      { "tipo": "gazpacho-informes-originales", "version": 1, "generado": "...",
 *        "pedidos": {
 *          "2682106": { "estado": "ok", "original": { "reportId": 2682093, "openTime": 1794589576,
 *                                                    "firstName": "ELVIS", "lastName": "LOPÉZ", "source": 0 } },
 *          "2682000": { "estado": "sin-original" },
 *          "2681998": { "estado": "sin-original-movil", "cadena": [2681990, 2681980] },
 *          "2681999": { "estado": "error", "detalle": "HTTP 500" } } }
 *
 *    "original" es SIEMPRE el informe Móvil del vendedor: si la oficina corrigió el
 *    pedido varias veces, la extensión sube la cadena Web → Web → Móvil ("cadena").
 *    La extensión guarda los datos TAL CUAL vienen de la API (docs/api-gazpacho.md);
 *    las conversiones (fecha, nombre) se hacen aquí, en un solo sitio y con pruebas.
 */

const PENDING_FILE_TYPE = 'gazpacho-pedidos-web-pendientes';
const ORIGINALS_FILE_TYPE = 'gazpacho-informes-originales';
const FILE_VERSION = 1;

/**
 * Venezuela está en UTC−4 todo el año (sin horario de verano).
 * La API da "segundos desde 1970" en hora universal; Gazpacho muestra la hora local.
 */
const VENEZUELA_UTC_OFFSET_HOURS = -4;

/**
 * Segundos UTC de la API → Date con la "hora de pared" de Venezuela guardada como UTC.
 * Es el mismo convenio que usan las fechas leídas del Excel (ver build-closing-rows.js),
 * así todas las fechas de la hoja CIERRE se comparan y se escriben igual.
 * Ej.: 1794589576 → 13/11/2026 13:06:16 (lo que se ve en pantalla).
 */
function apiSecondsToWallDate(seconds, offsetHours = VENEZUELA_UTC_OFFSET_HOURS) {
     return new Date((seconds + offsetHours * 3600) * 1000);
}

/**
 * Junta varios archivos de la extensión (del más antiguo al más reciente) en uno.
 * ¿Por qué? Si la extensión se detiene a mitad (corte, 403...) y se relanza, solo
 * pide los que faltaban y descarga OTRO archivo. Hay que usar todos.
 * Regla: lo más reciente gana, EXCEPTO que un error nunca borra un dato bueno.
 */
/** En la API, source 0 = Móvil (el pedido del vendedor), 1 = Web (la oficina). */
const SOURCE_MOBILE = 0;
const isMobileOriginal = (entry) => entry?.estado === 'ok' && entry.original?.source === SOURCE_MOBILE;

function mergeExtensionFiles(files) {
     const merged = { tipo: ORIGINALS_FILE_TYPE, version: FILE_VERSION, pedidos: {} };
     for (const file of files) {
          if (file?.tipo !== ORIGINALS_FILE_TYPE || file.version !== FILE_VERSION) {
               return { merged: file, invalid: true }; // originalsFromExtensionFile explicará el problema
          }
          for (const [key, entry] of Object.entries(file.pedidos ?? {})) {
               const previous = merged.pedidos[key];
               const previousIsGood = isMobileOriginal(previous) || previous?.estado === 'sin-original';
               if (entry?.estado === 'error' && previousIsGood) continue;
               merged.pedidos[key] = entry;
          }
     }
     return { merged, invalid: false };
}

/** Construye el contenido del archivo de pedidos pendientes para la extensión. */
function buildPendingFile(reportIds, periodLabel) {
     return { tipo: PENDING_FILE_TYPE, version: FILE_VERSION, periodo: periodLabel, pedidos: [...reportIds] };
}

/**
 * Lee el archivo de la extensión.
 * @returns {{ originals: Map<number, {Usuario, 'ID de Informe', Apertura}>,
 *             stats: { ok: number, sinOriginal: number, error: number, noMovil: number },
 *             problems: string[] }}
 */
function originalsFromExtensionFile(data) {
     const problems = [];
     const originals = new Map();
     const stats = { ok: 0, sinOriginal: 0, error: 0, noMovil: 0 };

     if (!data || data.tipo !== ORIGINALS_FILE_TYPE) {
          return { originals, stats, problems: [`El archivo no es de la extensión (tipo "${data?.tipo}").`] };
     }
     if (data.version !== FILE_VERSION) {
          return { originals, stats, problems: [`Versión de archivo ${data.version} no soportada (se esperaba ${FILE_VERSION}).`] };
     }

     for (const [key, entry] of Object.entries(data.pedidos ?? {})) {
          const reportId = Number(key);
          if (!Number.isInteger(reportId)) { problems.push(`Clave de pedido no válida: "${key}".`); continue; }

          if (entry?.estado === 'sin-original' || entry?.estado === 'sin-original-movil') { stats.sinOriginal++; continue; }
          if (entry?.estado !== 'ok') { stats.error++; continue; }

          const o = entry.original ?? {};
          if (!Number.isInteger(o.reportId) || typeof o.openTime !== 'number') {
               problems.push(`Pedido ${reportId}: el original no tiene reportId/openTime válidos.`);
               continue;
          }
          // Solo vale un original MÓVIL (el del vendedor). Un "original" Web es otra
          // versión de la oficina (cadena sin recorrer, p. ej. de una versión antigua
          // de la extensión): no se usa y el pedido vuelve a quedar pendiente.
          if (!isMobileOriginal(entry)) { stats.noMovil++; continue; }
          const usuario = [o.firstName, o.lastName].map((p) => String(p ?? '').trim()).filter(Boolean).join(' ');
          originals.set(reportId, {
               Usuario: usuario || null,
               'ID de Informe': o.reportId,
               Apertura: apiSecondsToWallDate(o.openTime),
          });
          stats.ok++;
     }
     return { originals, stats, problems };
}

// ───── src/transform/excel-sheets.js ─────
/**
 * excel-sheets.js — Leer las hojas de Gazpacho y escribir la hoja CIERRE.
 *
 * Trabaja sobre un libro de ExcelJS YA ABIERTO (no abre ni guarda archivos), así
 * que funciona igual en Node y en el navegador (extensión de Chrome):
 *   - Node:       excel-io.js abre el archivo del disco y llama a estas funciones.
 *   - Extensión:  abre el Excel que descarga de Gazpacho y llama a estas mismas.
 *
 * ¿Cómo se reconocen las hojas?
 *   Por sus ENCABEZADOS, no por su nombre. Gazpacho recorta y numera los nombres
 *   ("HOJA DE PEDIDOS CALOX (PARA PED", "TABLA DE PEDIDO_3"...) y podrían cambiar.
 */

// Una hoja es de PEDIDOS si tiene estas columnas, y de PRODUCTOS si tiene estas otras.
// 'Report Type' y 'TABLA DE PEDIDO' solo existen en las hojas de Gazpacho: así una
// hoja hecha a mano con columnas parecidas (como la "Hoja1" de referencia) no se
// confunde con una hoja de pedidos y no duplica los pedidos.
const ORDER_SHEET_HEADERS = ['Report ID', 'Report Type', 'Open Time', 'TABLA DE PEDIDO'];
const LINE_SHEET_HEADERS = ['Report', 'Index', 'Descripción', 'Cantidad'];
const HEADER_FILL_ARGB = 'FFF7941D';
const HEADER_FONT_ARGB = 'FF000000';
const CP1252 = new Map([
     [0x80, '€'], [0x82, '‚'], [0x83, 'ƒ'], [0x84, '„'], [0x85, '…'],
     [0x86, '†'], [0x87, '‡'], [0x88, 'ˆ'], [0x89, '‰'], [0x8a, 'Š'],
     [0x8b, '‹'], [0x8c, 'Œ'], [0x8e, 'Ž'], [0x91, '‘'], [0x92, '’'],
     [0x93, '“'], [0x94, '”'], [0x95, '•'], [0x96, '–'], [0x97, '—'],
     [0x98, '˜'], [0x99, '™'], [0x9a, 'š'], [0x9b, '›'], [0x9c, 'œ'],
     [0x9e, 'ž'], [0x9f, 'Ÿ'],
]);

function sanitizeText(value) {
     return value
          .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\ufffe\uffff]/g, '')
          .replace(/[\u007f-\u009f]/g, (ch) => CP1252.get(ch.charCodeAt(0)) ?? '')
          .replace(/[\ud800-\udbff](?![\udc00-\udfff])|(?<![\ud800-\udbff])[\udc00-\udfff]/g, '');
}

function sanitizeValue(value) {
     if (typeof value === 'string') {
          const clean = sanitizeText(value);
          return clean === value ? value : clean;
     }
     if (value === null || value === undefined || value instanceof Date || typeof value !== 'object') return value;
     if (Array.isArray(value)) {
          let dirty = false;
          const out = value.map((part) => {
               const clean = sanitizeValue(part);
               if (clean !== part) dirty = true;
               return clean;
          });
          return dirty ? out : value;
     }
     let dirty = false;
     const out = { ...value };
     for (const key of ['text', 'result']) {
          if (typeof out[key] !== 'string') continue;
          const clean = sanitizeText(out[key]);
          if (clean !== out[key]) {
               out[key] = clean;
               dirty = true;
          }
     }
     if (Array.isArray(out.richText)) {
          const clean = sanitizeValue(out.richText);
          if (clean !== out.richText) {
               out.richText = clean;
               dirty = true;
          }
     }
     return dirty ? out : value;
}

function sanitizeWorkbookStrings(workbook) {
     let changed = 0;
     for (const sheet of workbook.worksheets) {
          sheet.eachRow({ includeEmpty: false }, (row) => {
               row.eachCell({ includeEmpty: false }, (cell) => {
                    const clean = sanitizeValue(cell.value);
                    if (clean !== cell.value) {
                         cell.value = clean;
                         changed++;
                    }
               });
          });
     }
     return changed;
}

/**
 * Valor "limpio" de una celda de ExcelJS.
 * ExcelJS no siempre devuelve texto o número; según la celda puede devolver:
 *   - { text, hyperlink }  → enlace (el Report ID de Gazpacho es un enlace)
 *   - { richText: [...] }  → texto con varios formatos
 *   - { formula, result }  → fórmula
 *   - { error }            → error de Excel (#N/A...)
 */
function plainValue(value) {
     if (value === null || value === undefined) return null;
     if (value instanceof Date) return value;
     if (typeof value === 'string') return sanitizeText(value);
     if (typeof value !== 'object') return value;
     if (Array.isArray(value.richText)) return sanitizeText(value.richText.map((part) => part.text).join(''));
     if ('hyperlink' in value) {
          const text = plainValue(value.text);
          return typeof text === 'string' ? sanitizeText(text) : text;
     }
     if ('formula' in value || 'sharedFormula' in value) return plainValue(value.result);
     if ('error' in value) return null;
     return sanitizeText(String(value));
}

/** Lee una hoja como lista de objetos, usando la fila 1 como nombres de columna. */
function sheetToRecords(sheet) {
     const headers = [];
     sheet.getRow(1).eachCell({ includeEmpty: true }, (cell, col) => {
          headers[col] = String(plainValue(cell.value) ?? '').trim();
     });

     const records = [];
     const links = new Map(); // Report ID → URL del pedido en Gazpacho
     sheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
          if (rowNumber === 1) return;
          const record = {};
          headers.forEach((header, col) => {
               if (!header) return;
               const raw = row.getCell(col).value;
               record[header] = plainValue(raw);
               if (raw && typeof raw === 'object' && raw.hyperlink && header === 'Report ID') {
                    links.set(toReportId(record[header]), raw.hyperlink);
               }
          });
          records.push(record);
     });
     return { headers: headers.filter(Boolean), records, links };
}

const hasHeaders = (headers, required) => required.every((h) => headers.includes(h));

/**
 * Separa pedidos y productos de un libro de Gazpacho ya abierto.
 * @returns {{ orders, lines, reportLinks: Map, sheets: {orders: string[], lines: string[]} }}
 */
function extractGazpachoData(workbook) {
     const result = { orders: [], lines: [], reportLinks: new Map(), sheets: { orders: [], lines: [] } };
     for (const sheet of workbook.worksheets) {
          const { headers, records, links } = sheetToRecords(sheet);
          if (hasHeaders(headers, ORDER_SHEET_HEADERS)) {
               result.orders.push(...records);
               links.forEach((url, id) => result.reportLinks.set(id, url));
               result.sheets.orders.push(sheet.name);
          } else if (hasHeaders(headers, LINE_SHEET_HEADERS)) {
               result.lines.push(...records);
               result.sheets.lines.push(sheet.name);
          }
          // Cualquier otra hoja (p. ej. una "Hoja1" hecha a mano) se ignora.
     }

     if (result.sheets.orders.length === 0 || result.sheets.lines.length === 0) {
          throw new Error(
               `El archivo no parece un informe de Gazpacho: hojas de pedidos=${result.sheets.orders.length}, ` +
               `hojas de productos=${result.sheets.lines.length}.`
          );
     }
     return result;
}

/**
 * Añade la hoja CIERRE al libro (si ya existe una con ese nombre, la reemplaza).
 * Formato: encabezado en negrita, fila 1 fija, autofiltro, fechas DD/MM/AAAA HH:mm
 * y el Report ID como enlace al pedido en Gazpacho.
 */
function addClosingSheet(workbook, rows, { sheetName = 'CIERRE', reportLinks = new Map() } = {}) {
     const existing = workbook.getWorksheet(sheetName);
     if (existing) workbook.removeWorksheet(existing.id);

     const sheet = workbook.addWorksheet(sheetName, { views: [{ state: 'frozen', ySplit: 1 }] });
     sheet.columns = CLOSING_COLUMNS.map(({ header, isDate }) => ({
          header,
          key: header,
          width: isDate ? 17 : Math.min(Math.max(header.length + 2, 12), 45),
          style: isDate ? { numFmt: 'dd/mm/yyyy hh:mm' } : {},
     }));
     const header = sheet.getRow(1);
     header.font = { bold: true, color: { argb: HEADER_FONT_ARGB } };
     header.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: HEADER_FILL_ARGB } };

     for (const row of rows) {
          const added = sheet.addRow(row);
          const url = reportLinks.get(toReportId(row['Report ID']));
          if (url) {
               added.getCell('Report ID').value = { text: String(row['Report ID']), hyperlink: url };
          }
     }

     // Columnas de texto largo, más anchas para leerlas bien.
     for (const key of ['DROGUERIA', 'Descripción', 'Farmacia', 'Observación']) sheet.getColumn(key).width = 45;
     sheet.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: CLOSING_COLUMNS.length } };
     return sheet;
}

// ───── src/transform/period.js ─────
/**
 * period.js — El período del cierre: qué fechas entran, cómo se llama el archivo
 * y desde qué fecha se le pide el Excel a Gazpacho.
 *
 * Reglas puras (sin Node): las usan `npm run transformar` y la extensión.
 *
 * Nombres del archivo final:
 *   - mes completo (01/10/2026 → 31/10/2026)  → "CIERRE OCTUBRE CALOX.xlsx"
 *   - un día      (13/11/2026)               → "CIERRE CALOX 13-11-2026.xlsx"
 *   - un rango    (01/11/2026 → 13/11/2026)  → "CIERRE CALOX 01-11-2026 a 13-11-2026.xlsx"
 */


const MONTH_NAMES = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO',
     'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'];

/** "13/11/2026" → { day: 13, month: 11, year: 2026 }. Error claro si no tiene ese formato. */
function parseDDMMYYYY(text) {
     const m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(String(text ?? ''));
     if (!m) throw new Error(`Fecha no válida: "${text}". Formato esperado: DD/MM/AAAA.`);
     const [day, month, year] = [Number(m[1]), Number(m[2]), Number(m[3])];
     const check = new Date(Date.UTC(year, month - 1, day));
     if (check.getUTCMonth() !== month - 1 || check.getUTCDate() !== day) throw new Error(`La fecha ${text} no existe.`);
     return { day, month, year };
}

const lastDayOfMonth = (year, month) => new Date(Date.UTC(year, month, 0)).getUTCDate();
const slug = (ddmmyyyy) => ddmmyyyy.replaceAll('/', '-'); // 13/11/2026 → 13-11-2026

/** Mes completo. */
function monthPeriod(year, month) {
     return {
          range: monthRange(year, month),
          label: `${MONTH_NAMES[month - 1].toLowerCase()} ${year}`,
          outputName: `CIERRE ${MONTH_NAMES[month - 1]} CALOX.xlsx`,
     };
}

/** Mes anterior a `today` (el cierre de siempre): hoy 13/11/2026 → { year: 2026, month: 10 }. */
function previousMonth(today = new Date()) {
     const d = new Date(today.getFullYear(), today.getMonth() - 1, 1);
     return { year: d.getFullYear(), month: d.getMonth() + 1 };
}

/**
 * Período a partir de dos fechas DD/MM/AAAA (ambas incluidas).
 * Si cubren un mes completo, se trata como cierre mensual (nombre con el mes).
 */
function datesPeriod(desde, hasta = desde) {
     const from = parseDDMMYYYY(desde);
     const to = parseDDMMYYYY(hasta);
     const range = dayRange(desde, hasta); // valida además que "hasta" no sea anterior a "desde"

     const isWholeMonth = from.day === 1 && from.month === to.month && from.year === to.year &&
          to.day === lastDayOfMonth(to.year, to.month);
     if (isWholeMonth) return monthPeriod(from.year, from.month);

     return desde === hasta
          ? { range, label: `el ${desde}`, outputName: `CIERRE CALOX ${slug(desde)}.xlsx` }
          : { range, label: `del ${desde} al ${hasta}`, outputName: `CIERRE CALOX ${slug(desde)} a ${slug(hasta)}.xlsx` };
}

/** Sin filtro de fechas. */
function allDatesPeriod() {
     return { range: null, label: 'todas las fechas', outputName: 'CIERRE CALOX (TODO).xlsx' };
}

/**
 * Fecha "desde" para pedirle el Excel a Gazpacho: milisegundos UTC de las 00:00
 * de ese día en Venezuela (UTC−4). Comprobado con una captura real:
 * 01/10/2026 → 1790827200000.
 */
function exportFromDateMs(desde) {
     const { day, month, year } = parseDDMMYYYY(desde);
     return Date.UTC(year, month - 1, day) - VENEZUELA_UTC_OFFSET_HOURS * 3600 * 1000;
}

return Object.freeze({ CLOSING_COLUMNS, toReportId, hasQuantity, expectedLineCount, monthRange, dayRange, orderSource, buildClosingRows, PENDING_FILE_TYPE, ORIGINALS_FILE_TYPE, FILE_VERSION, VENEZUELA_UTC_OFFSET_HOURS, apiSecondsToWallDate, mergeExtensionFiles, buildPendingFile, originalsFromExtensionFile, sanitizeWorkbookStrings, extractGazpachoData, addClosingSheet, MONTH_NAMES, monthPeriod, previousMonth, datesPeriod, allDatesPeriod, exportFromDateMs });
})();
