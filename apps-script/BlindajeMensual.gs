/************************************************************************************************
 *  BLINDAJE AUTOMÁTICO DE OBJETIVOS DEL MES  —  Camping 44
 * ---------------------------------------------------------------------------------------------
 *  QUÉ HACE
 *  Cada día guarda ("blinda") los objetivos del mes que representa el CONFIG, dentro del propio
 *  CONFIG (filas HISTORICO_B64_n, col A/B, en base64). Así, cuando el mes cierra y cargás los
 *  objetivos del mes siguiente, el mes anterior YA quedó archivado con sus objetivos finales.
 *  Esto es lo que evita que un mes pasado aparezca en ₲0 (como pasó con octubre).
 *
 *  Calcula los objetivos con EXACTAMENTE la misma lógica que el tablero (limpiarNombre, handleNum,
 *  bloques de marca con "la última columna pisa"), para que lo archivado = lo que muestra el panel.
 *
 *  CÓMO INSTALAR (una sola vez)
 *   1) En la planilla: Extensiones → Apps Script.
 *   2) + (Secuencia de comandos) → nombralo  BlindajeMensual  → pegá TODO esto → Guardar 💾.
 *   3) Arriba, elegí la función  instalarBlindajeMensual  → ▶ Ejecutar.
 *      (La primera vez pide permisos → Revisar → tu cuenta → Avanzado → Permitir.)
 *   4) Vas a ver "✅ Blindaje mensual instalado". Listo, corre solo todos los días a las 22:00.
 *
 *  PARA PROBAR AHORA (sin esperar): elegí  probarBlindajeMensual  → ▶ Ejecutar. Te muestra qué mes
 *  y qué totales archivaría, SIN escribir nada.
 *  PARA BLINDAR YA (escribe): elegí  blindarMesActual_  → ▶ Ejecutar.
 *  PARA APAGARLO: elegí  desinstalarBlindajeMensual  → ▶ Ejecutar.
 *
 *  SEGURIDAD: si el histórico existente no se puede leer, o alguna fila estuviera ocupada por otro
 *  contenido, ABORTA y manda mail (nunca pisa ni borra meses ya guardados).
 ************************************************************************************************/

var BM_HOJA_CONFIG = "CONFIG";
var BM_PREFIJO_B64 = "HISTORICO_B64_";
var BM_CHUNK = 45000;           // chars por celda (bajo el límite de 50.000 de Google Sheets)
var BM_HORA_BLINDAJE = 22;      // corre todos los días a esta hora (zona horaria del proyecto)

/** Crea (o recrea) el disparador diario. Ejecutá esta función UNA vez. */
function instalarBlindajeMensual() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "blindarMesActual_") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("blindarMesActual_").timeBased().atHour(BM_HORA_BLINDAJE).everyDays(1).create();
  try { SpreadsheetApp.getActive().toast("✅ Blindaje mensual instalado: corre diario ~" + BM_HORA_BLINDAJE + ":00.", "Camping 44", 8); } catch (e) {}
}

/** Apaga el blindaje (borra el disparador). */
function desinstalarBlindajeMensual() {
  var n = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "blindarMesActual_") { ScriptApp.deleteTrigger(t); n++; }
  });
  try { SpreadsheetApp.getActive().toast("🛑 Blindaje mensual apagado (" + n + " disparador/es).", "Camping 44", 8); } catch (e) {}
}

/** DRY-RUN: muestra qué archivaría, sin escribir. Útil para verificar antes de confiar. */
function probarBlindajeMensual() {
  var sh = SpreadsheetApp.getActive().getSheetByName(BM_HOJA_CONFIG);
  if (!sh) { SpreadsheetApp.getUi().alert("No existe la hoja CONFIG"); return; }
  var filas = sh.getDataRange().getValues();
  var mes = bmMesDelConfig_(filas);
  var m = bmConstruirMetas_(filas);
  var h = bmLeerHistorico_(sh);
  var sumV = 0; Object.keys(m.vendedoresTotales).forEach(function (k) { sumV += m.vendedoresTotales[k]; });
  var msg = "Mes detectado (Fecha Fin): " + mes +
    "\nGlobal: ₲ " + m.global.toLocaleString("es-PY") +
    "\nSuma vendedores: ₲ " + sumV.toLocaleString("es-PY") + "  (cuadra: " + (sumV === m.global) + ")" +
    "\nCanales: " + Object.keys(m.canales).length +
    "\nObjetivos por marca: " + Object.keys(m.vendedorMarca).length +
    "\ndiasMes/diasTrans: " + m.diasMes + " / " + m.diasTranscurridos +
    "\n\nHistórico actual: meses [" + Object.keys(h.obj).sort(function(a,b){return a-b;}).join(", ") + "]" +
    (h.hadChunks && !h.ok ? "\n⚠️ OJO: el histórico existente NO se pudo leer (no se tocaría)." : "") +
    "\n\n(Esto es una PRUEBA: no se escribió nada.)";
  try { SpreadsheetApp.getUi().alert("🔎 Prueba de blindaje", msg, SpreadsheetApp.getUi().ButtonSet.OK); } catch (e) { Logger.log(msg); }
}

/* ============================ HELPERS (réplica del tablero) ============================ */

function bmLimpiar_(t) {
  if (!t) return "";
  t = String(t).normalize("NFD").replace(/[̀-ͯ]/g, "").toUpperCase();
  t = t.replace(/\s*-\s*(SALON|ONLINE|E-COMMERCE|MAYORISTAS|VENTA EXTERNA|REPARACIONES|DIRECTORIO)\s*$/g, "");
  return t.trim();
}
function bmHandleNum_(v) {
  if (!v) return 0;
  var s = String(v).replace(/\./g, "");
  if (s.indexOf(",") >= 0) s = s.split(",")[0];
  s = s.replace(/[^0-9-]/g, "");
  var n = parseInt(s, 10);
  return isNaN(n) ? 0 : n;
}
function bmCanalNorm_(c) {
  var u = String(c || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toUpperCase().trim();
  if (u === "SALON") return "Salon";
  if (u === "ONLINE") return "Online";
  if (u === "E-COMMERCE" || u === "ECOMMERCE") return "E-commerce";
  if (u === "MAYORISTAS" || u === "MAYORISTA") return "Mayoristas";
  if (u === "VENTA EXTERNA") return "Venta Externa";
  if (u === "DIRECTORIO") return "Directorio";
  if (u === "REPARACIONES") return "Reparaciones";
  return null;
}

/** Mes (1-12) que representa el CONFIG, leído de "Fecha Fin:" (o "Fecha Inicio:"). */
function bmMesDelConfig_(filas) {
  var fin = "", ini = "";
  filas.forEach(function (row) {
    var c0 = String(row[0] || "").trim();
    if (c0 === "Fecha Fin:" && row[1] !== "" && row[1] != null) fin = row[1];
    if (c0 === "Fecha Inicio:" && row[1] !== "" && row[1] != null) ini = row[1];
  });
  var src = (fin !== "" ? fin : ini);
  if (src === "" || src == null) return 0;
  if (Object.prototype.toString.call(src) === "[object Date]") return src.getMonth() + 1; // Apps Script devuelve Date
  var mm = String(src).match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);  // o texto dd/mm/aaaa
  return mm ? parseInt(mm[2], 10) : 0;
}

/** Construye la estructura de metas del mes a partir de las filas del CONFIG (igual que el tablero). */
function bmConstruirMetas_(filas) {
  var m = { diasMes: 25, diasTranscurridos: 25, global: 0, canales: {}, vendedoresTotales: {}, vendedorMarca: {} };

  filas.forEach(function (row, idx) {
    if (!row || !row.length) return;
    var col0 = String(row[0] || "").trim();
    if (col0 === "Días Hábiles del Mes:") { var v = parseFloat(String(row[1]).replace(",", ".")); if (v > 0 && v <= 31) m.diasMes = v; }
    if (idx === 3 && row[3]) m.global = bmHandleNum_(row[3]);                 // D4 = global
    if (row[5]) { var cn = bmCanalNorm_(row[5]); if (cn) m.canales[cn] = bmHandleNum_(row[6]); }  // F/G = metas por canal
    if (row[8] && idx >= 3 && String(row[8]).trim() !== "Vendedor") {       // I/J = metas por vendedor
      var vClean = bmLimpiar_(row[8]);
      var meta = bmHandleNum_(row[9]);
      if (meta > 0) m.vendedoresTotales[vClean] = meta;
    }
  });

  // Bloques "Metas individuales por vendedor y marca" (misma lógica que procesarMetasConfigActuales)
  var bh = -1;
  for (var i = 0; i < filas.length && i < 20; i++) {
    if (filas[i]) {
      for (var c = 0; c < filas[i].length; c++) {
        var t = String(filas[i][c] || "").toUpperCase().trim();
        if (t === "MARCA" || t === "MARCAS") { bh = i; break; }
      }
      if (bh !== -1) break;
    }
  }
  if (bh > 0) {
    var hr = filas[bh];
    for (var col = 0; col < hr.length; col++) {
      var h = String(hr[col] || "").toUpperCase().trim();
      if (h === "MARCA" || h === "MARCAS") {
        var vend = "";
        for (var rUp = bh - 1; rUp >= 0; rUp--) {
          var pn = String((filas[rUp] && filas[rUp][col]) || "").trim();
          if (pn && pn.indexOf("---") !== 0 && pn.toUpperCase().indexOf("METAS INDIVIDUALES") < 0) { vend = pn; break; }
        }
        if (vend) {
          for (var r = bh + 1; r < filas.length; r++) {
            if (!filas[r]) continue;
            var marc = String(filas[r][col] || "").trim();
            if (marc && marc.indexOf("---") !== 0 && marc.indexOf("[") !== 0) {
              var meta2 = bmHandleNum_(filas[r][col + 1]);
              if (meta2 > 0) m.vendedorMarca[bmLimpiar_(vend) + "|||" + bmLimpiar_(marc)] = meta2;
            }
          }
        }
      }
    }
  }

  // El mes que se archiva se guarda como COMPLETO: transcurridos = hábiles del mes
  // (así, al ver un mes cerrado, el % logrado se calcula contra el objetivo total).
  m.diasTranscurridos = m.diasMes;
  return m;
}

/* ============================ LECTURA / ESCRITURA DEL HISTÓRICO ============================ */

function bmLeerHistorico_(sh) {
  var vals = sh.getRange(1, 1, sh.getLastRow(), 2).getValues();
  var chunks = [];
  for (var r = 0; r < vals.length; r++) {
    var a = String(vals[r][0] || "").trim();
    if (a.indexOf(BM_PREFIJO_B64) === 0) chunks.push({ idx: parseInt(a.replace(BM_PREFIJO_B64, ""), 10) || 0, val: String(vals[r][1] || "") });
  }
  if (!chunks.length) return { obj: {}, ok: true, hadChunks: false };
  chunks.sort(function (a, b) { return a.idx - b.idx; });
  var b64 = chunks.map(function (c) { return c.val; }).join("");
  try {
    var json = Utilities.newBlob(Utilities.base64Decode(b64)).getDataAsString("UTF-8");
    return { obj: JSON.parse(json), ok: true, hadChunks: true };
  } catch (e) {
    return { obj: {}, ok: false, hadChunks: true };
  }
}

function bmEscribirHistorico_(sh, obj) {
  var json = JSON.stringify(obj);
  var b64 = Utilities.base64Encode(Utilities.newBlob(json).getBytes());  // UTF-8
  var partes = [];
  for (var i = 0; i < b64.length; i += BM_CHUNK) partes.push(b64.substr(i, BM_CHUNK));

  var vals = sh.getRange(1, 1, sh.getLastRow(), 2).getValues();
  var filasB64 = [];
  for (var r = 0; r < vals.length; r++) if (String(vals[r][0] || "").indexOf(BM_PREFIJO_B64) === 0) filasB64.push(r + 1); // 1-based
  var inicio = filasB64.length ? filasB64[0] : (sh.getLastRow() + 2);
  var finViejo = filasB64.length ? filasB64[filasB64.length - 1] : inicio - 1;
  var totalFilas = Math.max(partes.length, finViejo - inicio + 1);

  // Seguridad: no pisar filas que NO sean HISTORICO_B64 (col A con otro contenido).
  for (var k = 0; k < totalFilas; k++) {
    var fila = inicio + k;
    var aActual = (fila - 1 < vals.length) ? String(vals[fila - 1][0] || "").trim() : "";
    if (aActual && aActual.indexOf(BM_PREFIJO_B64) !== 0)
      throw new Error("Fila " + fila + " ocupada por '" + aActual + "'; abortado para no corromper el CONFIG.");
  }

  var out = [];
  for (var k2 = 0; k2 < totalFilas; k2++)
    out.push([k2 < partes.length ? (BM_PREFIJO_B64 + k2) : "", k2 < partes.length ? partes[k2] : ""]);
  sh.getRange(inicio, 1, totalFilas, 2).setValues(out);
}

/* ============================ FUNCIÓN QUE CORRE EL DISPARADOR ============================ */

function blindarMesActual_() {
  try {
    var sh = SpreadsheetApp.getActive().getSheetByName(BM_HOJA_CONFIG);
    if (!sh) throw new Error("No existe la hoja CONFIG");
    var filas = sh.getDataRange().getValues();

    var mes = bmMesDelConfig_(filas);
    if (!mes) throw new Error("No pude determinar el mes (Fecha Fin/Inicio vacíos o ilegibles)");

    var metas = bmConstruirMetas_(filas);
    // CONFIG vacío / a medio cargar: no blindar (evita pisar un mes bueno con algo incompleto).
    if (metas.global <= 0 && !Object.keys(metas.vendedoresTotales).length) return;

    var h = bmLeerHistorico_(sh);
    if (h.hadChunks && !h.ok) throw new Error("El histórico existente no se pudo leer (corrupto); abortado para NO perder lo ya guardado.");

    h.obj[String(mes)] = metas;          // agrega/actualiza SOLO el mes actual; conserva los demás
    bmEscribirHistorico_(sh, h.obj);
  } catch (e) {
    try {
      MailApp.sendEmail(Session.getEffectiveUser().getEmail(),
        "⚠️ Blindaje mensual falló - Camping 44",
        "El blindaje automático de objetivos falló.\n\n" + (e && e.message ? e.message : e));
    } catch (e2) {}
  }
}
