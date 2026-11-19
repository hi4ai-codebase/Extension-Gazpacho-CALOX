/**
 * puente.js — Le dice al panel (cierre-calox.js) dónde está el logo.
 *
 * ¿Por qué hace falta?
 *   cierre-calox.js se ejecuta "dentro" de la página de Gazpacho (mundo MAIN)
 *   para poder usar sus servicios, y desde ahí NO existe `chrome.runtime`, que
 *   es lo que da la dirección de los archivos de la extensión. Este script sí
 *   corre en el mundo de la extensión (ISOLATED), así que obtiene la dirección
 *   del logo y la deja anotada en la página para que el panel la lea.
 *   No hace nada más.
 */
document.documentElement.dataset.cierreCaloxLogo = chrome.runtime.getURL('logo_gazpacho.png');
