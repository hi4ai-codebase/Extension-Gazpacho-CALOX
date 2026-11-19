/**
 * cierre-calox.js — Extensión "Cierre CALOX" para Gazpacho.
 *
 * Qué hace (un solo botón: "Generar cierre")
 *   1. Pide a Gazpacho el Excel de pedidos desde la fecha elegida
 *      (lo mismo que "Guardar en Excel → New Excel").
 *   2. Lo lee y aplica las reglas del cierre (lib/cierre-core.js).
 *   3. Para cada pedido Web, busca el informe original MÓVIL del vendedor
 *      (subiendo la cadena Web → Web → Móvil si hace falta).
 *   4. Añade la hoja CIERRE y descarga "CIERRE <MES> CALOX.xlsx".
 *
 * Decisiones de seguridad (léelas antes de cambiar nada)
 *   1. NO inicia ni cierra sesión, ni guarda contraseñas: usa la sesión que el
 *      usuario abrió a mano. Al terminar, el usuario sigue dentro de Gazpacho.
 *   2. Usa los PROPIOS servicios de Gazpacho (exportReports, checkExportFileStatus,
 *      getReport), como si el usuario pulsara los botones. Gazpacho rota sus llaves
 *      de sesión en cada respuesta; así las sigue gestionando él. La descarga del
 *      archivo usa solo las cookies (igual que hace Gazpacho), sin tocar las llaves.
 *   3. Va DESPACIO (un pedido cada vez, con pausa). Se detiene si Gazpacho responde
 *      401/403/429, si hay un error de conexión o 3 errores seguidos.
 *   4. NO usa eval: lee las respuestas con un lector propio.
 *   5. Si algo falla o el usuario pulsa Detener, NO genera un Excel incompleto.
 *      Al volver a pulsar "Generar cierre" continúa donde quedó.
 *
 * Las REGLAS no están aquí: vienen de lib/cierre-core.js, que se genera desde
 * src/transform/ con `npm run build:extension` (las mismas que usa Node).
 * Se ejecuta en el "mundo" de la página (manifest: "world": "MAIN").
 */
(() => {
     'use strict';
     if (window.__cierreCaloxLoaded) return; // por si la página inyecta el script dos veces
     window.__cierreCaloxLoaded = true;

     const Core = window.CierreCaloxCore; // reglas compartidas (lib/cierre-core.js)
     const ExcelJSLib = window.ExcelJS;   // lib/exceljs.bare.min.js

     // ── Configuración ───────────────────────────────────────────────────────
     const CONFIG = {
          pauseMs: 500,                   // pausa entre consultas (≈ 2 por segundo como máximo)
          maxConsecutiveErrors: 3,        // errores seguidos antes de detenerse
          fatalStatuses: [401, 403, 429], // sesión caída / bloqueo / demasiadas peticiones
          maxChainSteps: 5,               // tope de escalones Web → Web → Móvil
          exportTemplateId: 5,            // plantilla "New Excel" (la que usa el cierre)
          exportPollMinMs: 2000,          // espera mínima entre comprobaciones del Excel
          exportPollMaxMs: 10000,         // espera máxima entre comprobaciones
          exportTimeoutMs: 10 * 60 * 1000,// si Gazpacho tarda más de 10 min, se abandona
     };
     const SOURCE_MOBILE = 0;              // en la API: 0 = Móvil, 1 = Web
     const RESOLVED = ['ok', 'sin-original', 'sin-original-movil'];

     // ── Estado ──────────────────────────────────────────────────────────────
     const state = {
          running: false,
          stopRequested: false,
          // Se conserva entre intentos para "continuar donde quedó":
          excel: null,   // { fromDateMs, buffer } último Excel descargado de Gazpacho
          results: {},   // reportId → { estado, original?, cadena? } (los originales no cambian)
     };

     const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
     const isResolved = (id) => RESOLVED.includes(state.results[id]?.estado);

     class StopRequested extends Error {}

     // ── Lector de "JSON" de Gazpacho (sin eval) ─────────────────────────────
     /**
      * Gazpacho devuelve JSON con comas sobrantes: `[1, 2, ]` o `{...},` al final.
      * JSON.parse no lo acepta. Aquí se quitan SOLO las comas que van justo antes
      * de `]`, `}` o del final del texto, respetando el contenido de los textos.
      */
     function parseGazpachoJson(text) {
          if (typeof text !== 'string') return text; // ya es un objeto
          let out = '';
          let inString = false;
          let escaped = false;
          for (let i = 0; i < text.length; i++) {
               const ch = text[i];
               if (inString) {
                    out += ch;
                    if (escaped) escaped = false;
                    else if (ch === '\\') escaped = true;
                    else if (ch === '"') inString = false;
                    continue;
               }
               if (ch === '"') { inString = true; out += ch; continue; }
               if (ch === ',') {
                    let j = i + 1;
                    while (j < text.length && /\s/.test(text[j])) j++;
                    if (j >= text.length || text[j] === ']' || text[j] === '}') continue; // coma sobrante
               }
               out += ch;
          }
          return JSON.parse(out);
     }

     // ── Acceso a Gazpacho (AngularJS) ───────────────────────────────────────
     /** Servicios de Gazpacho. Si la página aún no cargó, lanza un error claro. */
     function gazpacho() {
          const ng = window.angular;
          if (!ng) throw new Error('No se encuentra Gazpacho en esta página.');
          for (const el of [document.querySelector('[ng-app]'), document.documentElement, document.body]) {
               const injector = el && ng.element(el).injector();
               if (injector) {
                    return {
                         getReport: injector.get('getReport'),
                         exportReports: injector.get('exportReports'),
                         checkExportFileStatus: injector.get('checkExportFileStatus'),
                         $rootScope: injector.get('$rootScope'),
                    };
               }
          }
          throw new Error('Gazpacho todavía no terminó de cargar. Espera unos segundos.');
     }

     /** ¿Sigue la sesión abierta? (no estamos en el login y se ve el menú de Informes) */
     function isLoggedIn() {
          return !location.hash.includes('/login') && !!document.querySelector('[i18n-key="navigation.reports"]');
     }

     /**
      * Ejecuta una llamada de Gazpacho DENTRO de AngularJS.
      * AngularJS solo envía las peticiones durante su "ciclo de actualización"
      * (digest); $evalAsync pide uno. Devuelve una Promise normal.
      */
     function callInAngular(fn) {
          const { $rootScope } = gazpacho();
          return new Promise((resolve, reject) => {
               $rootScope.$evalAsync(() => {
                    try { fn().then(resolve, reject); } catch (err) { reject(err); }
               });
          });
     }

     /** Texto legible de un error de AngularJS ($http) o de JavaScript. */
     function describeError(err) {
          if (err === undefined) return 'Gazpacho rechazó la petición (¿se cerró la sesión?)';
          if (typeof err?.status === 'number') return `HTTP ${err.status}${err.statusText ? ' ' + err.statusText : ''}`;
          return err?.message ?? String(err);
     }

     const isFatal = (err) => err === undefined || CONFIG.fatalStatuses.includes(err?.status) || err?.status === 0;

     // ── Paso 1: pedir el Excel a Gazpacho ───────────────────────────────────
     /**
      * Igual que "Guardar en Excel → New Excel": Gazpacho prepara el archivo en
      * su servidor (reportExport), avisa cuando está listo (checkFileStatus) y
      * luego se descarga (download). Devuelve el contenido del .xlsx.
      */
     async function exportExcel(fromDateMs, onWait) {
          const { exportReports, checkExportFileStatus } = gazpacho();
          const body = {
               fromDate: fromDateMs,
               reportStatus: 1,
               pagination: { skipRows: 0, maxRows: 1000 },
               sortOrderKey: 'basic',
               exportTemplateId: CONFIG.exportTemplateId,
          };
          const exportUrl = `${exportReports.config.url}${exportReports.config.path}?mode=2`;
          let status = parseGazpachoJson((await callInAngular(() => exportReports.post(body, null, false, exportUrl))).data);
          if (!status?.processId || !status?.fileId) throw new Error('Gazpacho no aceptó la petición del Excel.');

          const started = Date.now();
          for (let cnt = 1; !status.downloadAvailable; cnt++) {
               if (state.stopRequested) throw new StopRequested();
               if (status.resultCode && status.resultCode !== 0) {
                    throw new Error(`Gazpacho no pudo generar el Excel (código ${status.resultCode}).`);
               }
               if (Date.now() - started > CONFIG.exportTimeoutMs) throw new Error('Gazpacho tardó demasiado en preparar el Excel.');
               onWait(cnt);
               const waitMs = Math.min(Math.max((status.delay ?? 0) * 1000, CONFIG.exportPollMinMs), CONFIG.exportPollMaxMs);
               await sleep(waitMs);
               const params = { processId: status.processId, fileId: status.fileId, delay: status.delay ?? 0, cnt };
               status = parseGazpachoJson((await callInAngular(() => checkExportFileStatus.getWithParameters(params))).data);
          }

          // Descarga: solo cookies, como hace Gazpacho (no toca las llaves de sesión).
          const base = new URL(exportReports.config.url, location.href);
          const url = new URL(`download?processId=${encodeURIComponent(status.processId)}&fileId=${encodeURIComponent(status.fileId)}`, base);
          const response = await fetch(url, { credentials: 'same-origin' });
          if (!response.ok) throw Object.assign(new Error(`No se pudo descargar el Excel (HTTP ${response.status}).`), { status: response.status });
          const buffer = await response.arrayBuffer();
          const bytes = new Uint8Array(buffer, 0, Math.min(2, buffer.byteLength));
          if (buffer.byteLength === 0 || bytes[0] !== 0x50 || bytes[1] !== 0x4b) { // un .xlsx empieza por "PK"
               throw new Error('Lo descargado no es un archivo Excel válido.');
          }
          return buffer;
     }

     // ── Paso 3: informe original MÓVIL de un pedido Web ─────────────────────
     /** Pide el detalle (reportView) de un informe a Gazpacho. */
     async function fetchReport(reportId) {
          const { getReport } = gazpacho();
          const response = await callInAngular(() => getReport.getById(null, reportId));
          let data = parseGazpachoJson(response?.data);
          if (Array.isArray(data)) data = data[0];
          if (!data || data.reportId !== reportId) {
               throw new Error(`respuesta inesperada (reportId ${data?.reportId})`);
          }
          return data;
     }

     /**
      * Busca el informe original MÓVIL (el del vendedor) de un pedido Web.
      * ¡Puede ser una CADENA! Si la oficina corrigió un pedido varias veces:
      *   Móvil (vendedor) → Web (1.ª versión) → Web (corregida, la del Excel)
      * "parentReport" apunta solo al escalón anterior; se sube hasta el Móvil.
      * Solo se guardan 5 datos del Móvil, más la lista de escalones (cadena).
      */
     async function fetchOriginal(reportId) {
          let current = await fetchReport(reportId);
          const cadena = [];
          for (let step = 0; step < CONFIG.maxChainSteps; step++) {
               const parents = Array.isArray(current.parentReport) ? current.parentReport : [];
               if (parents.length === 0) {
                    return cadena.length === 0 ? { estado: 'sin-original' } : { estado: 'sin-original-movil', cadena };
               }
               const p = parents[0];
               cadena.push(p.reportId);
               if (p.source === SOURCE_MOBILE) {
                    return {
                         estado: 'ok',
                         original: { reportId: p.reportId, openTime: p.openTime, firstName: p.firstName, lastName: p.lastName, source: p.source },
                         cadena,
                    };
               }
               await sleep(CONFIG.pauseMs); // misma pausa antes de subir otro escalón
               current = await fetchReport(p.reportId);
          }
          return { estado: 'sin-original-movil', cadena, detalle: `más de ${CONFIG.maxChainSteps} escalones` };
     }

     /** Consulta los pedidos Web que aún no tienen original. Se detiene ante errores graves. */
     async function fetchPendingOriginals(webIds, onProgress) {
          const todo = webIds.filter((id) => !isResolved(id));
          let consecutiveErrors = 0;
          for (const id of todo) {
               if (state.stopRequested) throw new StopRequested();
               if (!isLoggedIn()) throw new Error('La sesión de Gazpacho se cerró. Vuelve a entrar y pulsa "Generar cierre" para continuar.');
               try {
                    state.results[id] = await fetchOriginal(id);
                    consecutiveErrors = 0;
               } catch (err) {
                    const detalle = describeError(err);
                    state.results[id] = { estado: 'error', detalle };
                    ui.log(`Pedido ${id}: ${detalle}`, 'error');
                    if (isFatal(err)) throw new Error(`Gazpacho respondió "${detalle}". Me detengo para no insistir.`);
                    if (++consecutiveErrors >= CONFIG.maxConsecutiveErrors) throw new Error(`${consecutiveErrors} errores seguidos. Me detengo.`);
               }
               onProgress(webIds.filter(isResolved).length);
               await sleep(CONFIG.pauseMs);
          }
     }

     // ── Proceso completo: "Generar cierre" ──────────────────────────────────
     async function generate(desde, hasta) {
          if (state.running) return;
          let period;
          try {
               if (!Core || !ExcelJSLib) throw new Error('Faltan archivos de la extensión (lib/). Pide que la reinstalen.');
               period = Core.datesPeriod(desde, hasta);
               if (!isLoggedIn()) throw new Error('Primero inicia sesión en Gazpacho.');
               gazpacho();
          } catch (err) {
               ui.fail(err.message);
               return;
          }

          state.running = true;
          state.stopRequested = false;
          ui.start(period.label);
          try {
               // 1. Excel de Gazpacho (se reutiliza si es de la misma fecha "desde")
               const fromDateMs = Core.exportFromDateMs(desde);
               if (state.excel?.fromDateMs === fromDateMs) {
                    ui.step(1, 'Uso el Excel que ya pedí a Gazpacho.', 15);
               } else {
                    ui.step(1, 'Pidiendo el Excel a Gazpacho…', 2);
                    const buffer = await exportExcel(fromDateMs, (cnt) => ui.step(1, `Gazpacho está preparando el Excel… (${cnt})`, Math.min(2 + cnt, 14)));
                    state.excel = { fromDateMs, buffer };
                    ui.log(`Excel recibido (${Math.round(buffer.byteLength / 1024)} KB).`, 'ok');
               }

               // 2. Leer y aplicar las reglas
               ui.step(2, 'Leyendo el Excel…', 16);
                const workbook = new ExcelJSLib.Workbook();
                await workbook.xlsx.load(state.excel.buffer.slice(0)); // copia: el original se reutiliza al reintentar
                Core.sanitizeWorkbookStrings(workbook);
                const data = Core.extractGazpachoData(workbook);
               const first = Core.buildClosingRows({ orders: data.orders, lines: data.lines, range: period.range });
               if (first.problems.length > 0) {
                    throw new Error(`Los datos de Gazpacho no cuadran (${first.problems.length} problemas), no genero el archivo. ` +
                         `Ejemplo: ${first.problems[0]}`);
               }
               if (first.rows.length === 0) throw new Error(`No hay pedidos ${period.label}.`);
               const webIds = first.summary.webPendingIds;
               ui.log(`${first.summary.ordersInRange} pedidos ${period.label}: ${first.summary.ordersMobile} Móvil y ${webIds.length} Web.`);

               // 3. Vendedor original de los pedidos Web
               const progress = (done) => ui.step(3, `Buscando vendedores originales… ${done} de ${webIds.length}`,
                    20 + Math.round((done / Math.max(webIds.length, 1)) * 75), `${done}/${webIds.length}`);
               progress(webIds.filter(isResolved).length);
               await fetchPendingOriginals(webIds, progress);
               const failed = webIds.filter((id) => !isResolved(id));
               if (failed.length > 0) {
                    // Errores sueltos: no se genera un archivo con vendedores que faltan por error.
                    throw new Error(`${failed.length} pedido(s) no se pudieron consultar. ` +
                         'Pulsa "Generar cierre" otra vez para reintentar solo esos.');
               }

               // 4. Hoja CIERRE y descarga
               ui.step(4, 'Generando el archivo…', 96);
               const { originals } = Core.originalsFromExtensionFile({
                    tipo: Core.ORIGINALS_FILE_TYPE, version: Core.FILE_VERSION, pedidos: state.results,
               });
               const final = Core.buildClosingRows({ orders: data.orders, lines: data.lines, range: period.range, originals });
               Core.addClosingSheet(workbook, final.rows, { reportLinks: data.reportLinks });
               const output = await workbook.xlsx.writeBuffer();
               saveFile(output, period.outputName);

               const s = final.summary;
               ui.log(`${s.ordersInRange} pedidos · ${final.rows.length} filas · ${s.productLinesKept} productos.`, 'ok');
               if (s.ordersWebWithoutOriginal > 0) {
                    ui.log(`${s.ordersWebWithoutOriginal} pedido(s) Web sin vendedor original en Gazpacho (columnas vacías): ` +
                         s.webPendingIds.join(', '));
               }
               if (s.ordersWithoutProducts.length > 0) {
                    ui.log(`${s.ordersWithoutProducts.length} pedido(s) sin productos con cantidad: ${s.ordersWithoutProducts.join(', ')}`);
               }
               ui.done(period.outputName);
          } catch (err) {
               if (err instanceof StopRequested) ui.stopped();
               else ui.fail(err.message);
          } finally {
               state.running = false;
               ui.refresh();
          }
     }

     /** Descarga un archivo a la carpeta Descargas del usuario. */
     function saveFile(content, fileName) {
          const blob = new Blob([content], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
          const a = document.createElement('a');
          a.href = URL.createObjectURL(blob);
          a.download = fileName;
          document.body.appendChild(a);
          a.click();
          a.remove();
          setTimeout(() => URL.revokeObjectURL(a.href), 60_000);
     }

     // ── Fechas del formulario ───────────────────────────────────────────────
     // El <input type="date"> usa AAAA-MM-DD; las reglas usan DD/MM/AAAA.
     const toInput = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
     const fromInput = (value) => (value ? value.split('-').reverse().join('/') : '');
     const PRESETS = {
          'mes-anterior': () => { const t = new Date(); return [new Date(t.getFullYear(), t.getMonth() - 1, 1), new Date(t.getFullYear(), t.getMonth(), 0)]; },
          'este-mes': () => { const t = new Date(); return [new Date(t.getFullYear(), t.getMonth(), 1), t]; },
          ayer: () => { const d = new Date(); d.setDate(d.getDate() - 1); return [d, d]; },
          hoy: () => { const d = new Date(); return [d, d]; },
     };

     // ── Interfaz (panel flotante) ───────────────────────────────────────────
     // - Vive dentro de un Shadow DOM: una "burbuja" de estilos. Así el CSS de
     //   Gazpacho no deforma el panel y el nuestro no toca Gazpacho.
     // - La plantilla HTML es FIJA (sin datos). Todo texto variable se pone con
     //   textContent, nunca con innerHTML.
     // - Minimizar NO detiene el trabajo: el botón redondo muestra el avance.
     const ui = (() => {
          const STYLES = `
               :host { all: initial; }
               * { box-sizing: border-box; font-family: "Segoe UI", Roboto, Arial, sans-serif; }
               .fab {
                    position: fixed; right: 16px; bottom: 16px; z-index: 2147483647;
                    display: flex; align-items: center; gap: 8px;
                    padding: 6px 14px 6px 6px; border: 0; border-radius: 999px;
                    background: #fff; color: #c25e00; font-size: 13px; font-weight: 600;
                    box-shadow: 0 4px 14px rgba(0,0,0,.25); cursor: pointer; transition: transform .15s;
               }
               .fab:hover { transform: translateY(-2px); }
               .fab img { width: 34px; height: 34px; }
               .badge { background: #f7941d; color: #fff; border-radius: 999px; padding: 2px 8px; font-size: 11px; }
               .badge:empty { display: none; }
               .panel {
                    position: fixed; right: 16px; bottom: 16px; z-index: 2147483647; width: 380px;
                    background: #fff; color: #333; border-radius: 12px; overflow: hidden;
                    box-shadow: 0 8px 28px rgba(0,0,0,.28); font-size: 13px;
               }
               header {
                    display: flex; align-items: center; gap: 10px; padding: 10px 12px;
                    background: linear-gradient(90deg, #f7941d, #fbb040); color: #fff;
               }
               header img { width: 32px; height: 32px; background: #fff; border-radius: 50%; padding: 2px; }
               header .title { flex: 1; font-weight: 700; font-size: 15px; line-height: 1.1; }
               header small { display: block; font-weight: 400; font-size: 11px; opacity: .9; }
               .icon {
                    width: 28px; height: 28px; border: 0; border-radius: 6px; cursor: pointer;
                    background: rgba(255,255,255,.25); color: #fff; font-size: 18px; line-height: 1;
               }
               .icon:hover { background: rgba(255,255,255,.4); }
               .body { padding: 12px; }
               .label { font-weight: 700; font-size: 12px; color: #666; margin-bottom: 6px; }
               .dates { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; }
               .dates label { font-size: 11px; color: #777; display: grid; gap: 3px; }
               .dates input {
                    padding: 7px 8px; border: 1.5px solid #ddd; border-radius: 8px; font-size: 13px; color: #333;
               }
               .dates input:focus { outline: none; border-color: #f7941d; }
               .presets { display: flex; flex-wrap: wrap; gap: 6px; margin: 8px 0 12px; }
               .chip {
                    padding: 4px 10px; border: 1px solid #f7941d; border-radius: 999px; background: #fff;
                    color: #c25e00; font-size: 11.5px; cursor: pointer;
               }
               .chip:hover:not(:disabled) { background: #fff4e8; }
               .generate {
                    width: 100%; padding: 11px; border: 0; border-radius: 10px; cursor: pointer;
                    background: #f7941d; color: #fff; font-size: 15px; font-weight: 700;
                    box-shadow: 0 2px 6px rgba(247,148,29,.4); transition: background .15s;
               }
               .generate:hover:not(:disabled) { background: #e8850f; }
               button:disabled, input:disabled { opacity: .5; cursor: not-allowed; }
               .stop {
                    width: 100%; margin-top: 8px; padding: 8px; border: 1.5px solid #b00020; border-radius: 10px;
                    background: #fff; color: #b00020; font-weight: 700; cursor: pointer;
               }
               .stop:hover { background: #fdecee; }
               .progress { height: 8px; margin: 12px 0 6px; background: #f1f1f1; border-radius: 999px; overflow: hidden; }
               .bar { height: 100%; width: 0; background: linear-gradient(90deg, #f7941d, #8cc63f); transition: width .3s; }
               .status { font-weight: 600; font-size: 12.5px; min-height: 18px; }
               .status.ok { color: #1b7a2f; }
               .status.error { color: #b00020; }
               details { margin-top: 8px; }
               summary { cursor: pointer; font-size: 11.5px; color: #777; }
               .log {
                    margin-top: 6px; max-height: 160px; overflow: auto; padding: 6px 8px;
                    background: #fafafa; border: 1px solid #eee; border-radius: 8px;
                    font: 11px/1.45 Consolas, monospace; white-space: pre-wrap;
               }
               .log .error { color: #b00020; }
               .log .ok { color: #1b7a2f; }
               .note { margin-top: 8px; font-size: 11px; color: #888; }
               [hidden] { display: none !important; }
          `;
          // Plantilla fija (no lleva datos de Gazpacho).
          const TEMPLATE = `
               <button class="fab" title="Abrir Cierre CALOX">
                    <img alt=""><span>Cierre CALOX</span><span class="badge"></span>
               </button>
               <section class="panel" hidden>
                    <header>
                         <img alt="">
                         <div class="title">Cierre CALOX<small>Genera el Excel del cierre</small></div>
                         <button class="icon" data-action="minimize" title="Minimizar (sigue trabajando)">–</button>
                    </header>
                    <div class="body">
                         <div class="label">Período</div>
                         <div class="dates">
                              <label>Desde<input type="date" data-field="desde"></label>
                              <label>Hasta<input type="date" data-field="hasta"></label>
                         </div>
                         <div class="presets">
                              <button class="chip" data-preset="mes-anterior">Mes anterior</button>
                              <button class="chip" data-preset="este-mes">Este mes</button>
                              <button class="chip" data-preset="ayer">Ayer</button>
                              <button class="chip" data-preset="hoy">Hoy</button>
                         </div>
                         <button class="generate" data-action="generate">Generar cierre</button>
                         <button class="stop" data-action="stop" hidden>■ Detener</button>
                         <div class="progress"><div class="bar"></div></div>
                         <div class="status"></div>
                         <details><summary>Ver detalles</summary><div class="log"></div></details>
                         <div class="note">No cierres ni recargues esta pestaña mientras trabaja. Tu sesión no se cierra.</div>
                    </div>
               </section>`;

          const host = document.createElement('div');
          host.id = 'cierre-calox-extension';
          const root = host.attachShadow({ mode: 'open' });
          const style = document.createElement('style');
          style.textContent = STYLES;
          root.append(style);
          root.append(document.createRange().createContextualFragment(TEMPLATE));

          const $ = (selector) => root.querySelector(selector);
          const fab = $('.fab');
          const panel = $('.panel');
          const badge = $('.badge');
          const bar = $('.bar');
          const status = $('.status');
          const logBox = $('.log');
          const inputs = { desde: $('[data-field="desde"]'), hasta: $('[data-field="hasta"]') };
          const action = (name) => $(`[data-action="${name}"]`);

          // Logo: su dirección la deja puente.js (ver ese archivo). Si no está, se oculta.
          const logoUrl = document.documentElement.dataset.cierreCaloxLogo;
          root.querySelectorAll('img').forEach((img) => {
               if (logoUrl) img.src = logoUrl; else img.hidden = true;
          });

          const setDates = ([from, to]) => { inputs.desde.value = toInput(from); inputs.hasta.value = toInput(to); };
          setDates(PRESETS['mes-anterior']()); // por defecto: el cierre del mes anterior

          const setOpen = (open) => { panel.hidden = !open; fab.hidden = open; };
          fab.onclick = () => setOpen(true);
          action('minimize').onclick = () => setOpen(false);
          root.querySelectorAll('[data-preset]').forEach((chip) => {
               chip.onclick = () => setDates(PRESETS[chip.dataset.preset]());
          });
          action('generate').onclick = () => generate(fromInput(inputs.desde.value), fromInput(inputs.hasta.value));
          action('stop').onclick = () => {
               state.stopRequested = true;
               setStatus('Deteniendo…');
          };

          function setStatus(text, kind) {
               status.textContent = text;
               status.className = `status${kind ? ' ' + kind : ''}`;
          }

          function log(message, kind) {
               const line = document.createElement('div');
               if (kind) line.className = kind;
               line.textContent = `[${new Date().toLocaleTimeString('es-VE')}] ${message}`;
               logBox.append(line);
               logBox.scrollTop = logBox.scrollHeight;
          }

          function refresh() {
               const busy = state.running;
               action('generate').disabled = busy;
               action('stop').hidden = !busy;
               inputs.desde.disabled = inputs.hasta.disabled = busy;
               root.querySelectorAll('[data-preset]').forEach((chip) => { chip.disabled = busy; });
          }

          return {
               refresh,
               log,
               start(label) {
                    bar.style.width = '0';
                    badge.textContent = '';
                    log(`Inicio: cierre ${label}.`);
                    refresh();
               },
               step(n, text, percent, badgeText) {
                    setStatus(`Paso ${n} de 4 · ${text}`);
                    bar.style.width = `${percent}%`;
                    badge.textContent = badgeText ?? `${percent}%`;
               },
               done(fileName) {
                    bar.style.width = '100%';
                    badge.textContent = '✔';
                    setStatus(`✔ Listo: "${fileName}" está en tu carpeta Descargas.`, 'ok');
                    log(`Descargado ${fileName}.`, 'ok');
               },
               stopped() {
                    badge.textContent = '';
                    setStatus('Detenido. Pulsa "Generar cierre" para continuar donde quedó.');
                    log('Detenido por el usuario. Lo ya consultado se conserva.');
               },
               fail(message) {
                    badge.textContent = '!';
                    setStatus(`✘ ${message}`, 'error');
                    log(message, 'error');
               },
               mount() {
                    document.body.append(host);
                    setOpen(false);
                    refresh();
               },
          };
     })();

     ui.mount();

     // Si el usuario sale de la página con el proceso en marcha, avisar.
     window.addEventListener('beforeunload', (event) => {
          if (state.running) { event.preventDefault(); event.returnValue = ''; }
     });
})();
