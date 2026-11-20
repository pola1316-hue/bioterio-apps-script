// ================== CONSTANTES (solo en este archivo) ==================
// Aquí viven todas las constantes de despachos. watcher_bioterio.gs NO las declara.
const HOJA_DESPACHO = "Despacho IA";
const HOJA_RESUMEN = "Resumen mensual";
const HOJA_TABLA = "Tabla final";
const HOJA_PESOS = "Pesos promedio";

// Peso promedio por individuo (g) para pasar a gramos lo que se despacha contado.
// Se crean en la hoja 'Pesos promedio' la primera vez; desde ahí el zootecnista los cambia.
const PESOS_DEFECTO = [
  ["Codornices", 120, ""],
  ["Ratones", 30, ""],
  ["Ratas", 150, "Promedio: salen ratas de 60 g hasta de 300-500 g."],
  ["Conejos", 200, "Pesar los conejos nuevos: los que están llegando son más grandes."],
  ["Curies", "", "Sin peso promedio todavía."],
  ["Pollos", "", "Sin peso promedio todavía."],
  ["Gallina", "", "Sin peso promedio todavía."],
  ["Zophobas", 0.75, "Promedio aproximado de larva; se usa cuando se anotan individuos."],
  ["Dubias", 2, "Promedio aproximado; se usa cuando se anotan individuos."],
  ["Grillos", 0.4, "Promedio aproximado; se usa cuando se anotan individuos."],
  ["Tenebrios", 0.15, "Promedio aproximado de larva; se usa cuando se anotan individuos."]
];
const VALOR_TACHON = "tachón";

const COL = {
  FECHA: 0, ESPECIE: 1, CANTIDAD: 2, PESO: 3, DESTINO: 4,
  ENC_BIO: 5, ENC_ZONA: 6, ESTADO: 7, REQUIERE: 8, ALERTAS: 9,
  OBS: 10, ESPECIE_ORIG: 11, ARCHIVO: 12, IMAGE_ID: 13, FECHA_PROC: 14,
  DESTINO_ORIG: 15
};

const ENCABEZADOS_DESPACHO = [
  "Fecha", "Producción a despachar", "Cantidad de individuos", "Peso (g)",
  "Destino", "Encargado Bioterio", "Encargado Zona",
  "Estado revisión", "Requiere revisión", "Alertas", "Observaciones",
  "Producción (lectura original)", "Archivo origen", "Image ID", "Fecha procesamiento",
  "Destino (lectura original)"
];
const NUM_COLUMNAS = ENCABEZADOS_DESPACHO.length;

const ENCABEZADOS_RESUMEN = [
  "Año", "Mes", "Producción", "Total individuos", "Peso pesado (g)", "Peso estimado (g)",
  "Peso total (kg)", "Registros validados", "Última actualización"
];

const ENCABEZADOS_RESUMEN_DESTINO = [
  "Año", "Mes", "Destino", "Total individuos", "Peso pesado (g)", "Peso estimado (g)",
  "Peso total (kg)", "Registros validados", "Última actualización"
];
// El bloque de destino empieza dos columnas después del de producción.
const COL_RESUMEN_DESTINO = ENCABEZADOS_RESUMEN.length + 2;

// Orden fijo con el que aparecen las producciones y destinos en los resúmenes.
const ORDEN_ESPECIES = [
  "Tenebrios", "Zophobas", "Grillos", "Dubias", "Gallina", "Pollos",
  "Ratones", "Ratas", "Codornices", "Conejos", "Curies"
];
const ORDEN_DESTINOS = [
  "Neonatos", "Ocelote", "Biología", "Cocina", "Rapaces", "Zorro", "Hospital", "Reptiles", "Otros"
];

const ESPECIES_CATALOGO = {
  "Tenebrios":  ["tenebrio", "tenebrios", "tenebrioo", "tenebrioss"],
  "Zophobas":   ["zophoba", "zophobas", "sophoba", "sophobas", "zofoba", "zofobas",
                 "sofoba", "sofobas", "zofovas", "zofova", "sofovas", "zophovas", "zophabas"],
  "Grillos":    ["grillo", "grillos", "grilo", "grilos", "griyo", "griyos", "grilloss"],
  "Dubias":     ["dubia", "dubias", "dubi", "dubis", "dubya", "dubyas",
                 "cucaracha", "cucarachas", "cucarach", "cucarahc"],
  "Gallina":    ["gallina", "gallinas", "gallin", "galina", "galinas"],
  "Pollos":     ["pollo", "pollos", "polo", "polos", "pollito", "pollitos"],
  "Ratones":    ["raton", "ratones", "ratn", "ratonez"],
  "Ratas":      ["rata", "ratas", "rrata", "ratass"],
  "Codornices": ["codorniz", "codornices", "codornicez", "codornidez", "codornis",
                 "condorniz", "condornices", "codorniss", "codornizs"],
  "Conejos":    ["conejo", "conejos", "conejoo", "conexo", "conexos"],
  "Curies":     ["curi", "curies", "curis", "cuy", "cuye", "cuyes",
                 "conejillo de indias", "conejillos de indias", "conejillo indias",
                 "conejillo", "conejillos", "curi de indias"]
};

const DESTINOS_CATALOGO = {
  "Neonatos":  ["neo", "neos", "neonato", "neonatos", "neonatoss", "neonatoz",
                 "neonatu", "neonatus", "neotpo", "neotpos", "neonatal", "neonatales"],
  "Ocelote":   ["ocelote", "ocelotes", "ocelot"],
  "Biología":  ["biologia", "biologias", "biolgia", "biologa", "biologías", "biologica"],
  "Cocina":    ["cocina", "cocinas", "cosina", "cocinna"],
  "Rapaces":   ["rapaz", "rapaces", "rapas", "rapazes", "rapax"],
  "Zorro":     ["zorro", "zorros", "sorro"],
  "Hospital":  ["hospital", "hospitales", "ospital"],
  "Reptiles":  ["reptil", "reptiles", "reptill"],
  "Otros":     ["otro", "otros", "ortos"]
};

// ================== MENÚ ==================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🐭 Bioterio")
    .addItem("1. ▶ Leer fotos nuevas", "leerFotosNuevas")
    .addSeparator()
    .addItem("2. ▶ Generar tabla resumen (después de revisar)", "generarTablaResumen")
    .addToUi();
}

// Paso 1: lee las fotos nuevas y llena 'Despacho IA'. NO genera tablas.
function leerFotosNuevas() {
  const ui = SpreadsheetApp.getUi();
  try {
    revisarCarpetasBioterio();
  } catch (e) {
    ui.alert("No se pudo leer la foto: " + e.message);
    return;
  }
  SpreadsheetApp.getActive().toast("Fotos leídas. Revisa y corrige 'Despacho IA'; luego usa 'Generar tabla resumen'.", "Bioterio", 10);
}

// Paso 2: con los datos revisados, genera 'Tabla final' (según las fechas de su B2/B3) y el 'Resumen mensual' histórico.
function generarTablaResumen() {
  const ui = SpreadsheetApp.getUi();
  if (leerFilasDespacho_().length === 0) {
    ui.alert("'Despacho IA' está vacía. Primero usa 'Leer fotos nuevas'.");
    return;
  }
  try {
    generarTablaFinal();
    generarResumenMensual();
  } catch (e) {
    ui.alert("No se pudo generar la tabla: " + e.message);
    return;
  }
  SpreadsheetApp.getActive().toast("Tabla final y resumen mensual actualizados.", "Bioterio", 8);
}

// ================== RANGO DE FECHAS (en 'Tabla final') ==================
// B2 = fecha inicio, B3 = fecha fin (en 'Tabla final'). Vacías = todo el registro.
// Debe leerse ANTES de limpiar la hoja.
function leerRangoTabla_() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_TABLA);
  if (!hoja) return null;
  const ini = hoja.getRange("B2").getValue();
  const fin = hoja.getRange("B3").getValue();
  if (!(ini instanceof Date) || !(fin instanceof Date)) return null;
  return { inicio: soloDia_(ini), fin: soloDia_(fin) };
}

function soloDia_(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function fechaEnRango_(valor, rango) {
  if (!rango) return true;
  const d = valor instanceof Date ? soloDia_(valor) : fechaDesdeTexto_(valor);
  if (!d || isNaN(d)) return false;
  return d >= rango.inicio && d <= rango.fin;
}

// ================== INTERPRETACIÓN DE FECHAS ==================
// Acepta dd-mm-aaaa, dd/mm/aaaa, dd-mm-aa, dd/mm/aa, dd mm aa, dd/mm y dd-mm.
// Si no trae año, usa el año actual. También acepta objetos Date de Sheets.
function interpretarFecha_(valor) {
  if (valor instanceof Date && !isNaN(valor)) {
    return soloDia_(valor);
  }
  const partes = String(valor || "").trim().split(/[-\/ .]+/);
  if (partes.length < 2) return null;

  const dia = parseInt(partes[0], 10);
  const mes = parseInt(partes[1], 10);
  let anio = partes.length >= 3 ? parseInt(partes[2], 10) : new Date().getFullYear();
  if (!Number.isFinite(dia) || !Number.isFinite(mes) || !Number.isFinite(anio)) return null;
  if (anio < 100) anio += 2000;
  if (mes < 1 || mes > 12 || dia < 1 || dia > 31) return null;

  const d = new Date(anio, mes - 1, dia);
  // Evita fechas imposibles como 31/02.
  if (d.getDate() !== dia) return null;
  return d;
}

function fechaDesdeTexto_(valor) {
  return interpretarFecha_(valor);
}

function obtenerAnioMes_(valor) {
  const d = interpretarFecha_(valor);
  if (!d) return null;
  return { anio: d.getFullYear(), mes: d.getMonth() + 1 };
}

// ================== UTILIDADES DE TEXTO ==================
function normalizarTexto(texto) {
  return String(texto || "")
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9~ ]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

// Detecta tachones. Nunca se analiza ni se adivina el valor de un campo tachado.
function esTachado(texto) {
  const raw = String(texto || "").trim();
  if (raw === "") return false;
  if (raw === "~" || raw === "-" || raw === "--") return true;
  return normalizarTexto(raw).indexOf("tach") === 0;
}

function aNumero_(valor) {
  if (valor === "" || valor === null || valor === undefined) return null;
  if (esTachado(valor)) return null;
  const n = parseFloat(String(valor).replace(",", "."));
  return Number.isFinite(n) ? n : null;
}

function distanciaEdicion(a, b) {
  const m = a.length, n = b.length;
  const d = [];
  for (let i = 0; i <= m; i++) d[i] = [i];
  for (let j = 0; j <= n; j++) d[0][j] = j;
  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      const costo = a[i - 1] === b[j - 1] ? 0 : 1;
      d[i][j] = Math.min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + costo);
    }
  }
  return d[m][n];
}

// Devuelve { valor, estado, original }
// estado: "ok" | "aproximado" | "tachado" | "vacio" | "no_reconocido"
function clasificarValor(valor, catalogo) {
  const original = String(valor || "").trim();
  if (original === "") return { valor: "", estado: "vacio", original: original };
  if (esTachado(original)) return { valor: VALOR_TACHON, estado: "tachado", original: original };

  const n = normalizarTexto(original);

  // 'Pinky' no es una especie: se quita el prefijo y se clasifica solo lo que sigue (rata o ratón).
  const pinky = n.match(/^(pinky|pinkys|pinkies|pinkie|pinki|pinkis|pinks)\s+(.+)$/);
  if (pinky) {
    const resto = pinky[2];
    const coincidencias = ["Ratones", "Ratas"].filter(function(k) {
      return catalogo[k] && catalogo[k].indexOf(resto) !== -1;
    });
    if (coincidencias.length === 1) return { valor: coincidencias[0], estado: "ok", original: original };
    return { valor: original, estado: "no_reconocido", original: original };
  }

  for (const oficial of Object.keys(catalogo)) {
    if (catalogo[oficial].indexOf(n) !== -1) return { valor: oficial, estado: "ok", original: original };
  }

  let mejor = null, mejorDist = Infinity, empate = false;
  for (const oficial of Object.keys(catalogo)) {
    for (const alias of catalogo[oficial]) {
      const d = distanciaEdicion(n, alias);
      if (d < mejorDist) { mejorDist = d; mejor = oficial; empate = false; }
      else if (d === mejorDist && mejor !== oficial) { empate = true; }
    }
  }
  const umbral = n.length >= 6 ? 2 : 1;
  if (mejor && !empate && mejorDist <= umbral) {
    return { valor: mejor, estado: "aproximado", original: original };
  }
  return { valor: original, estado: "no_reconocido", original: original };
}

// Separa cantidad (número sin unidad) y peso (valor con gramos).
function separarCantidadPeso(cantidadTexto, pesoTexto) {
  const c = String(cantidadTexto || "").trim();
  const p = String(pesoTexto || "").trim();
  let cantidad = "";
  let peso = "";

  if (esTachado(c)) cantidad = VALOR_TACHON;
  if (esTachado(p)) peso = VALOR_TACHON;

  if (cantidad === "" && c !== "") {
    if (esPesoTexto_(c)) {
      peso = peso || extraerNumero_(c);
    } else {
      cantidad = c;
    }
  }
  if (peso === "" && p !== "") peso = extraerNumero_(p);

  return { cantidad: cantidad, peso: peso };
}

function esPesoTexto_(s) {
  return /\d\s*(g|gr|grs|gramos?)\.?\s*$/i.test(s) || /\b(gr|grs|gramos?)\b/i.test(s);
}

function extraerNumero_(s) {
  const m = String(s).match(/\d+(?:[.,]\d+)?/);
  return m ? m[0] : String(s).trim();
}

// ================== VALIDACIÓN DE UN REGISTRO ==================
function validarDespacho(r) {
  const alertas = [];

  const fecha = String(r.fecha || "").trim();
  if (esTachado(fecha)) {
    r.fecha = VALOR_TACHON;
  } else if (fecha === "") {
    alertas.push({ tipo: "fecha_faltante", mensaje: "No se leyó la fecha." });
  } else if (!interpretarFecha_(fecha)) {
    alertas.push({ tipo: "fecha_no_interpretable", mensaje: "Fecha no interpretable: '" + fecha + "'." });
  }

  const esp = clasificarValor(r.produccion_a_despachar, ESPECIES_CATALOGO);
  r.produccion_a_despachar = esp.valor;
  r.produccion_lectura_original = esp.original;
  if (esp.estado === "aproximado") {
    alertas.push({ tipo: "especie_aproximada", mensaje: "Especie '" + esp.original + "' interpretada como '" + esp.valor + "'. Confirmar." });
  } else if (esp.estado === "no_reconocido") {
    alertas.push({ tipo: "especie_no_reconocida", mensaje: "Especie no reconocida: '" + esp.original + "'." });
  } else if (esp.estado === "vacio") {
    alertas.push({ tipo: "especie_faltante", mensaje: "No se leyó la especie." });
  }

  const des = clasificarValor(r.destino, DESTINOS_CATALOGO);
  r.destino_lectura_original = des.original;
  r.destino = des.valor;
  if (des.estado === "aproximado") {
    alertas.push({ tipo: "destino_aproximado", mensaje: "Destino '" + des.original + "' interpretado como '" + des.valor + "'. Confirmar." });
  } else if (des.estado === "no_reconocido") {
    alertas.push({ tipo: "destino_no_reconocido", mensaje: "Destino no reconocido: '" + des.original + "'." });
  } else if (des.estado === "vacio") {
    alertas.push({ tipo: "destino_faltante", mensaje: "No se leyó el destino." });
  }

  const sep = separarCantidadPeso(r.cantidad_individuos, r.peso);
  r.cantidad_individuos = sep.cantidad;
  r.peso = sep.peso;

  // A crossed-out row often arrives with almost every field empty; if it stays, the rows below look shifted.
  const vacioOTachon = function(v) { return String(v || "").trim() === "" || esTachado(v); };
  const huecos = [fecha, r.produccion_a_despachar, r.cantidad_individuos + r.peso, r.destino].filter(vacioOTachon).length;
  if (huecos >= 3) {
    alertas.push({ tipo: "posible_fila_tachada", mensaje: "Casi toda la fila está vacía o tachada: puede ser una fila tachada que se debe borrar. Si es así, revise también que el destino y los encargados de las filas de abajo no hayan quedado corridos una fila." });
  }

  if (r.cantidad_individuos === "" && r.peso === "") {
    alertas.push({ tipo: "cantidad_faltante", mensaje: "No se leyó cantidad ni peso." });
  } else {
    if (r.cantidad_individuos !== "" && !esTachado(r.cantidad_individuos)) {
      const n = aNumero_(r.cantidad_individuos);
      if (n === null || n <= 0) alertas.push({ tipo: "cantidad_invalida", mensaje: "Cantidad no válida: '" + r.cantidad_individuos + "'." });
    }
    if (r.peso !== "" && !esTachado(r.peso)) {
      const n = aNumero_(r.peso);
      if (n === null || n <= 0) alertas.push({ tipo: "peso_invalido", mensaje: "Peso no válido: '" + r.peso + "'." });
    }
  }

  // El encargado de bioterio y el de zona NO generan alerta si están vacíos.

  const obs = String(r.observaciones || "").toLowerCase();
  if (["sobreescrit", "borron", "ilegible", "duda"].some(function(p) { return obs.indexOf(p) !== -1; })) {
    alertas.push({ tipo: "lectura_dudosa", mensaje: "La observación indica posible error de lectura." });
  }

  r.alertas = alertas;
  r.requiere_revision = alertas.length > 0;
  return r;
}

// ================== DETECCIÓN DE DUPLICADOS (por fila) ==================
function claveFila_(fecha, especie, cantidad, peso, destino, encargado) {
  const f = fecha instanceof Date
    ? Utilities.formatDate(fecha, Session.getScriptTimeZone(), "dd/MM/yyyy")
    : String(fecha || "").trim();
  return [f, especie, cantidad, peso, destino, encargado]
    .map(function(x) { return String(x || "").trim().toLowerCase(); }).join("|");
}

// ================== HOJA DESPACHO IA ==================
function guardarDespachoIA(registros, nombreArchivo, imageId) {
  if (!registros.length) return;

  const libro = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = libro.getSheetByName(HOJA_DESPACHO);
  if (!hoja) hoja = libro.insertSheet(HOJA_DESPACHO);

  if (hoja.getLastRow() === 0) {
    hoja.getRange(1, 1, 1, NUM_COLUMNAS).setValues([ENCABEZADOS_DESPACHO]).setFontWeight("bold");
    hoja.setFrozenRows(1);
  } else {
    verificarEncabezadosDespacho_(hoja);
  }

  const claves = new Set();
  leerFilasDespacho_().forEach(function(f) {
    claves.add(claveFila_(f[COL.FECHA], f[COL.ESPECIE], f[COL.CANTIDAD], f[COL.PESO], f[COL.DESTINO], f[COL.ENC_BIO]));
  });

  const ahora = new Date();
  const filas = registros.map(function(r) {
    const clave = claveFila_(r.fecha, r.produccion_a_despachar, r.cantidad_individuos, r.peso, r.destino, r.encargado_bioterio);
    if (claves.has(clave)) {
      r.alertas = (r.alertas || []).concat([{ tipo: "posible_duplicado", mensaje: "Posible duplicado de una fila ya registrada." }]);
      r.requiere_revision = true;
    }
    claves.add(clave);

    const fila = new Array(NUM_COLUMNAS).fill("");
    fila[COL.FECHA] = r.fecha || "";
    fila[COL.ESPECIE] = r.produccion_a_despachar || "";
    fila[COL.CANTIDAD] = r.cantidad_individuos || "";
    fila[COL.PESO] = r.peso || "";
    fila[COL.DESTINO] = r.destino || "";
    fila[COL.ENC_BIO] = r.encargado_bioterio || "";
    fila[COL.ENC_ZONA] = r.encargado_zona || "";
    fila[COL.ESTADO] = r.requiere_revision ? "Pendiente" : "Revisado";
    fila[COL.REQUIERE] = r.requiere_revision ? "Sí" : "No";
    fila[COL.ALERTAS] = (r.alertas || []).map(function(a) { return a.mensaje; }).join(" | ");
    fila[COL.OBS] = r.observaciones || "";
    fila[COL.ESPECIE_ORIG] = r.produccion_lectura_original || "";
    fila[COL.ARCHIVO] = nombreArchivo;
    fila[COL.IMAGE_ID] = imageId;
    fila[COL.FECHA_PROC] = ahora;
    fila[COL.DESTINO_ORIG] = r.destino_lectura_original || "";
    return fila;
  });

  hoja.getRange(hoja.getLastRow() + 1, 1, filas.length, NUM_COLUMNAS).setValues(filas);
}

// Headers added in later versions at the end of the sheet are written when missing; any other difference stops with
// the exact column, so existing data is never read into the wrong column.
function verificarEncabezadosDespacho_(hoja) {
  if (hoja.getMaxColumns() < NUM_COLUMNAS) hoja.insertColumnsAfter(hoja.getMaxColumns(), NUM_COLUMNAS - hoja.getMaxColumns());
  const actual = hoja.getRange(1, 1, 1, NUM_COLUMNAS).getValues()[0].map(function(v) { return String(v).trim(); });
  let ultima = -1;
  actual.forEach(function(v, i) { if (v !== "") ultima = i; });
  const diferencias = [];
  ENCABEZADOS_DESPACHO.forEach(function(h, i) {
    if (i > ultima) return;
    if (actual[i] !== h) diferencias.push("columna " + String.fromCharCode(65 + i) + " dice '" + actual[i] + "' y debe decir '" + h + "'");
  });
  if (diferencias.length) {
    throw new Error("La fila 1 de 'Despacho IA' no coincide: " + diferencias.join("; ") + ". Corrija esos encabezados y vuelva a intentar.");
  }
  if (ultima < NUM_COLUMNAS - 1) {
    hoja.getRange(1, ultima + 2, 1, NUM_COLUMNAS - ultima - 1)
      .setValues([ENCABEZADOS_DESPACHO.slice(ultima + 1)]).setFontWeight("bold");
  }
}

function leerFilasDespacho_() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_DESPACHO);
  if (!hoja || hoja.getLastRow() < 2) return [];
  verificarEncabezadosDespacho_(hoja);
  return hoja.getRange(2, 1, hoja.getLastRow() - 1, NUM_COLUMNAS).getValues();
}

// Válida si fue revisada (o no requiere revisión), sin tachón en fecha/producción/destino
// y con cantidad o peso numérico.
function esFilaValida_(f) {
  const estadoOk = f[COL.ESTADO] === "Revisado" || f[COL.REQUIERE] === "No";
  if (!estadoOk) return false;
  if ([f[COL.FECHA], f[COL.ESPECIE], f[COL.DESTINO]].some(esTachado)) return false;
  return aNumero_(f[COL.CANTIDAD]) !== null || aNumero_(f[COL.PESO]) !== null;
}

// ================== PESOS PROMEDIO ==================
// Lee 'Pesos promedio' (Producción | Peso promedio (g) | Nota). Si no existe, la crea con los valores por defecto.
function leerPesosPromedio_() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = libro.getSheetByName(HOJA_PESOS);
  if (!hoja) {
    hoja = libro.insertSheet(HOJA_PESOS);
    hoja.getRange(1, 1, 1, 3).setValues([["Producción", "Peso promedio (g)", "Nota"]])
      .setFontWeight("bold").setBackground("#1F6F4A").setFontColor("#FFFFFF");
    hoja.getRange(2, 1, PESOS_DEFECTO.length, 3).setValues(PESOS_DEFECTO);
    hoja.getRange(2, 2, PESOS_DEFECTO.length, 1).setBackground("#FFF2CC");
    hoja.setFrozenRows(1);
    hoja.autoResizeColumns(1, 3);
  }
  // Completa los pesos por defecto que falten en una hoja creada con una versión anterior; no cambia los que ya tienen valor.
  const filas = hoja.getLastRow() >= 2 ? hoja.getRange(2, 1, hoja.getLastRow() - 1, 3).getValues() : [];
  PESOS_DEFECTO.forEach(function(def) {
    if (def[1] === "") return;
    const i = filas.findIndex(function(f) { return especieOficial_(f[0]) === def[0]; });
    if (i === -1) {
      hoja.appendRow(def);
      filas.push(def);
    } else if (aNumero_(filas[i][1]) === null) {
      hoja.getRange(i + 2, 2, 1, 2).setValues([[def[1], def[2]]]);
      filas[i] = [filas[i][0], def[1], def[2]];
    }
  });
  const pesos = {};
  filas.forEach(function(f) {
    const especie = especieOficial_(f[0]);
    const g = aNumero_(f[1]);
    if (especie && g !== null && g > 0) pesos[especie] = g;
  });
  return pesos;
}

// Gramos de una fila: el peso anotado si lo hay; si no, individuos x peso promedio.
function pesoFila_(f, pesos) {
  const especie = especieOficial_(f[COL.ESPECIE]);
  const pesado = aNumero_(f[COL.PESO]);
  const ind = aNumero_(f[COL.CANTIDAD]);
  if (pesado !== null) return { pesado: pesado, estimado: 0, total: pesado, origen: "Pesado" };
  if (ind !== null && pesos[especie]) {
    const g = ind * pesos[especie];
    return { pesado: 0, estimado: g, total: g, origen: "Estimado (" + pesos[especie] + " g c/u)" };
  }
  return { pesado: 0, estimado: 0, total: 0, origen: "Sin peso promedio" };
}

// ================== TABLA FINAL (detalle del rango) ==================
// Layout: A1 título | A2/B2 fecha inicio | A3/B3 fecha fin | A4 periodo | fila 6 encabezados | datos desde fila 7
function generarTablaFinal() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const tz = Session.getScriptTimeZone();

  // Leer las fechas escritas por el zootecnista ANTES de limpiar la hoja.
  const rango = leerRangoTabla_();
  const filas = leerFilasDespacho_();
  if (!filas.length) throw new Error("No hay datos en 'Despacho IA'.");
  const pesos = leerPesosPromedio_();

  const validas = filas.filter(function(f) {
    return esFilaValida_(f) && fechaEnRango_(f[COL.FECHA], rango);
  });
  const pendientes = filas.filter(function(f) { return !esFilaValida_(f); }).length;

  let hoja = libro.getSheetByName(HOJA_TABLA);
  if (!hoja) hoja = libro.insertSheet(HOJA_TABLA);
  hoja.clear();

  hoja.getRange("A1").setValue("Producción del bioterio - despachos validados")
    .setFontSize(14).setFontWeight("bold");

  hoja.getRange("A2").setValue("Fecha inicio");
  hoja.getRange("A3").setValue("Fecha fin");
  hoja.getRange("B2").setValue(rango ? rango.inicio : "");
  hoja.getRange("B3").setValue(rango ? rango.fin : "");
  formatearSeguro_(hoja.getRange("B2:B3"), "dd/mm/yyyy");
  hoja.getRange("B2:B3")
    .setBackground("#FFF2CC").setBorder(true, true, true, true, true, true);
  hoja.getRange("C2").setValue("Escribe las fechas y pulsa '2. Generar tabla resumen'. Vacías = todo el registro.")
    .setFontStyle("italic");

  const periodo = rango
    ? "Periodo: " + Utilities.formatDate(rango.inicio, tz, "dd/MM/yyyy") +
      " a " + Utilities.formatDate(rango.fin, tz, "dd/MM/yyyy")
    : "Periodo: todo el registro";
  hoja.getRange("A4").setValue(periodo + " | Registros incluidos: " + validas.length +
    " | Pendientes (no incluidos): " + pendientes);

  const encabezados = ["Fecha", "Producción a despachar", "Cantidad de individuos", "Peso (g)", "Destino",
    "Peso estimado (g)", "Peso total (g)", "Origen del peso"];
  hoja.getRange(6, 1, 1, encabezados.length).setValues([encabezados])
    .setFontWeight("bold").setBackground("#1F6F4A").setFontColor("#FFFFFF");

  if (validas.length) {
    const cuerpo = validas.map(function(f) {
      const p = pesoFila_(f, pesos);
      return [f[COL.FECHA], especieOficial_(f[COL.ESPECIE]), f[COL.CANTIDAD], f[COL.PESO], f[COL.DESTINO],
        p.estimado || "", p.total || "", p.origen];
    });
    hoja.getRange(7, 1, cuerpo.length, encabezados.length).setValues(cuerpo)
      .setBorder(true, true, true, true, true, true);
  }

  const porEspecie = {};
  const porDestino = {};
  validas.forEach(function(f) {
    acumular_(porEspecie, especieOficial_(f[COL.ESPECIE]), f, pesos);
    acumular_(porDestino, f[COL.DESTINO], f, pesos);
  });

  const inicioResumen = 7 + validas.length + 2;
  const siguiente = escribirBloqueResumen_(hoja, "Total por producción", "Producción", porEspecie, inicioResumen, ORDEN_ESPECIES);
  escribirBloqueResumen_(hoja, "Total por destino", "Destino", porDestino, siguiente, ORDEN_DESTINOS);

  hoja.setFrozenRows(6);
  hoja.autoResizeColumns(1, encabezados.length);
}

// ================== RESUMEN MENSUAL (histórico acumulado) ==================
// Hoja 'Resumen mensual':
//   A-G  -> Producción por mes (histórico)
//   I-O  -> Destino por mes (histórico)
// Solo se recalculan los meses presentes en 'Despacho IA'; los demás se conservan.
function generarResumenMensual() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const filas = leerFilasDespacho_();
  if (!filas.length) throw new Error("No hay datos en 'Despacho IA'.");

  const pesos = leerPesosPromedio_();
  const nuevosEsp = {};
  const nuevosDes = {};
  const mesesRecalculados = {};
  const ahora = new Date();

  filas.forEach(function(f) {
    if (!esFilaValida_(f)) return;
    const ym = obtenerAnioMes_(f[COL.FECHA]);
    if (!ym) return;

    const mesClave = ym.anio + "-" + ym.mes;
    mesesRecalculados[mesClave] = true;

    const especie = especieOficial_(f[COL.ESPECIE]);
    const destino = String(f[COL.DESTINO] || "").trim();

    if (especie) acumularMes_(nuevosEsp, mesClave + "|" + especie, ym, especie, f, pesos);
    if (destino) acumularMes_(nuevosDes, mesClave + "|" + destino, ym, destino, f, pesos);
  });

  let hoja = libro.getSheetByName(HOJA_RESUMEN);
  if (!hoja) hoja = libro.insertSheet(HOJA_RESUMEN);

  // 1. Leer el histórico existente antes de limpiar la hoja.
  // La versión anterior (v21) tenía 7 columnas por bloque y el de destino en la columna I; se convierte al formato nuevo.
  const conservEsp = [];
  const conservDes = [];
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila >= 3) {
    const formatoViejo = String(hoja.getRange(2, 5).getValue()).trim() === "Total peso (g)";
    const ancho = formatoViejo ? 7 : ENCABEZADOS_RESUMEN.length;
    const colDestino = formatoViejo ? 9 : COL_RESUMEN_DESTINO;
    const aNuevo = function(f, esProduccion) {
      if (!formatoViejo) return f;
      // [año, mes, nombre, ind, peso, n, fecha] -> [año, mes, nombre, ind, pesado, estimado, kg, n, fecha]
      const pp = esProduccion ? pesos[especieOficial_(f[2])] : null;
      const pesado = aNumero_(f[4]) || 0;
      const estimado = !pesado && pp ? (aNumero_(f[3]) || 0) * pp : 0;
      return [f[0], f[1], f[2], f[3], pesado, estimado, (pesado + estimado) / 1000, f[5], f[6]];
    };
    const leerBloque = function(col, destino, esProduccion) {
      if (hoja.getLastColumn() < col) return;
      hoja.getRange(3, col, ultimaFila - 2, ancho).getValues().forEach(function(f) {
        if (f[0] === "" || f[1] === "" || f[2] === "") return;
        if (!mesesRecalculados[f[0] + "-" + Number(f[1])]) destino.push(aNuevo(f, esProduccion));
      });
    };
    leerBloque(1, conservEsp, true);
    leerBloque(colDestino, conservDes, false);
  }

  // 2. Construir filas nuevas.
  const filaMes = function(n) {
    return [n.anio, n.mes, n.nombre, n.ind, n.pesado, n.estimado, Math.round(n.pesado + n.estimado) / 1000, n.n, ahora];
  };
  const filasEsp = conservEsp.concat(Object.keys(nuevosEsp).map(function(k) {
    return filaMes(nuevosEsp[k]);
  })).sort(function(a, b) {
    if (a[0] !== b[0]) return a[0] - b[0];
    if (a[1] !== b[1]) return a[1] - b[1];
    return indiceOrden_(a[2], ORDEN_ESPECIES) - indiceOrden_(b[2], ORDEN_ESPECIES) ||
      String(a[2]).localeCompare(String(b[2]));
  });

  const filasDes = conservDes.concat(Object.keys(nuevosDes).map(function(k) {
    return filaMes(nuevosDes[k]);
  })).sort(function(a, b) {
    if (a[0] !== b[0]) return a[0] - b[0];
    if (a[1] !== b[1]) return a[1] - b[1];
    return indiceOrden_(a[2], ORDEN_DESTINOS) - indiceOrden_(b[2], ORDEN_DESTINOS) ||
      String(a[2]).localeCompare(String(b[2]));
  });

  // 3. Reescribir la hoja.
  hoja.clear();

  // Títulos en la fila 1 y encabezados en la fila 2; los datos empiezan en la fila 3.
  hoja.getRange(1, 1).setValue("Resumen mensual - Producción").setFontSize(12).setFontWeight("bold");
  hoja.getRange(2, 1, 1, ENCABEZADOS_RESUMEN.length)
    .setValues([ENCABEZADOS_RESUMEN])
    .setFontWeight("bold").setBackground("#1F6F4A").setFontColor("#FFFFFF");
  if (filasEsp.length) {
    hoja.getRange(3, 1, filasEsp.length, ENCABEZADOS_RESUMEN.length)
      .setValues(filasEsp).setBorder(true, true, true, true, true, true);
    formatearSeguro_(hoja.getRange(3, 7, filasEsp.length, 1), "0.00");
    formatearSeguro_(hoja.getRange(3, 9, filasEsp.length, 1), "dd/mm/yyyy hh:mm");
  }

  hoja.getRange(1, COL_RESUMEN_DESTINO).setValue("Resumen mensual - Destino").setFontSize(12).setFontWeight("bold");
  hoja.getRange(2, COL_RESUMEN_DESTINO, 1, ENCABEZADOS_RESUMEN_DESTINO.length)
    .setValues([ENCABEZADOS_RESUMEN_DESTINO])
    .setFontWeight("bold").setBackground("#1F6F4A").setFontColor("#FFFFFF");
  if (filasDes.length) {
    hoja.getRange(3, COL_RESUMEN_DESTINO, filasDes.length, ENCABEZADOS_RESUMEN_DESTINO.length)
      .setValues(filasDes).setBorder(true, true, true, true, true, true);
    formatearSeguro_(hoja.getRange(3, COL_RESUMEN_DESTINO + 6, filasDes.length, 1), "0.00");
    formatearSeguro_(hoja.getRange(3, COL_RESUMEN_DESTINO + 8, filasDes.length, 1), "dd/mm/yyyy hh:mm");
  }

  hoja.setFrozenRows(2);
  hoja.autoResizeColumns(1, ENCABEZADOS_RESUMEN.length);
  hoja.autoResizeColumns(COL_RESUMEN_DESTINO, ENCABEZADOS_RESUMEN_DESTINO.length);
}

// Acumula por mes y nombre (producción o destino).
function acumularMes_(mapa, clave, ym, nombre, f, pesos) {
  if (!mapa[clave]) mapa[clave] = { anio: ym.anio, mes: ym.mes, nombre: nombre, ind: 0, pesado: 0, estimado: 0, n: 0 };
  const ind = aNumero_(f[COL.CANTIDAD]);
  const p = pesoFila_(f, pesos);
  if (ind !== null) mapa[clave].ind += ind;
  mapa[clave].pesado += p.pesado;
  mapa[clave].estimado += p.estimado;
  mapa[clave].n += 1;
}

// ================== AUXILIARES ==================
// Aplica formato de número/fecha sin detener la generación de datos si Sheets lo rechaza.
function formatearSeguro_(rango, formato) {
  try {
    rango.setNumberFormat(formato);
  } catch (e) {
    Logger.log("Formato no aplicado en " + rango.getA1Notation() + ": " + e.message);
  }
}
// Posición de un valor en una lista de orden; los desconocidos quedan al final.
function indiceOrden_(valor, orden) {
  const i = orden.indexOf(valor);
  return i === -1 ? orden.length : i;
}

// Devuelve la producción oficial aunque la celda tenga un alias antiguo (p. ej. "Pinky ratones" -> "Ratones").
function especieOficial_(valor) {
  const v = String(valor || "").trim();
  if (!v) return "";
  if (ESPECIES_CATALOGO[v]) return v;
  const r = clasificarValor(v, ESPECIES_CATALOGO);
  return (r.estado === "ok" || r.estado === "aproximado") ? r.valor : v;
}

function acumular_(mapa, clave, f, pesos) {
  if (!mapa[clave]) mapa[clave] = { ind: 0, pesado: 0, estimado: 0 };
  const ind = aNumero_(f[COL.CANTIDAD]);
  const p = pesoFila_(f, pesos);
  if (ind !== null) mapa[clave].ind += ind;
  mapa[clave].pesado += p.pesado;
  mapa[clave].estimado += p.estimado;
}

function escribirBloqueResumen_(hoja, titulo, etiqueta, mapa, filaInicio, orden) {
  hoja.getRange(filaInicio, 1).setValue(titulo).setFontWeight("bold");
  const claves = Object.keys(mapa).sort(function(a, b) {
    return indiceOrden_(a, orden) - indiceOrden_(b, orden) || a.localeCompare(b);
  });
  if (!claves.length) return filaInicio + 2;

  hoja.getRange(filaInicio + 1, 1, 1, 5)
    .setValues([[etiqueta, "Total individuos", "Peso pesado (g)", "Peso estimado (g)", "Peso total (kg)"]])
    .setFontWeight("bold");

  const filas = claves.map(function(k) {
    const m = mapa[k];
    return [k, m.ind, m.pesado, m.estimado, Math.round(m.pesado + m.estimado) / 1000];
  });
  const totales = filas.reduce(function(a, f) { return [a[0], a[1] + f[1], a[2] + f[2], a[3] + f[3], Math.round((a[2] + f[2] + a[3] + f[3])) / 1000]; },
    ["Total", 0, 0, 0, 0]);
  filas.push(totales);
  hoja.getRange(filaInicio + 2, 1, filas.length, 5).setValues(filas)
    .setBorder(true, true, true, true, true, true);
  hoja.getRange(filaInicio + 1 + filas.length, 1, 1, 5).setFontWeight("bold");
  formatearSeguro_(hoja.getRange(filaInicio + 2, 5, filas.length, 1), "0.00");

  return filaInicio + filas.length + 3;
}
