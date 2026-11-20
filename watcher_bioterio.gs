// ================== CONFIGURACIÓN (solo en este archivo) ==================
// ID de la carpeta raíz "Fotos registros bioterio" (no el de una subcarpeta).
const CARPETA_RAIZ_ID = "1woCX2jOUWZVRlqWh1MFzIQ8nppI-r2oO";
const HOJA_CONTROL = "Control procesados";

// Modelos en orden de preferencia. Si el principal responde 503 o agota su cuota,
// se usa el siguiente. Cada modelo tiene su propia cuota.
const MODELOS = ["gemini-3.6-flash", "gemini-flash-lite-latest"];
const INTENTOS_POR_MODELO = 3;

// Cada subcarpeta se asocia a un procesador. Null = pendiente de configurar.
const TIPOS = {
  "Roedores": null,
  "Despachos": procesarDespachos,
  "Grillos": null
};

const PROMPT_ROEDORES = `
Analiza cuidadosamente la fotografía de esta tabla de registros de roedores.
Extrae todos los registros visibles y responde ÚNICAMENTE con JSON válido, sin Markdown.

Estructura exacta:
{
  "registros": [
    {
      "fecha_nacimiento": "",
      "total_crias": null,
      "total_natimuertos": null,
      "fecha_destete": "",
      "total_hembras": null,
      "total_machos": null,
      "observaciones": ""
    }
  ]
}

Reglas:
- Incluye una entrada por cada fila o registro visible.
- No inventes datos. Si un dato no se puede leer, usa "" para texto o null para números.
- Conserva las fechas tal como aparecen, pero no incluyas símbolos ♀ o ♂ dentro de las fechas.
- Si hay tachones, sobreescritura, borrones, guiones o dudas, indícalos en observaciones.
- No corrijas silenciosamente datos ni determines la causa de una inconsistencia.
- Devuelve siempre un objeto JSON con la propiedad "registros".
`;

const PROMPT_DESPACHO = `
Analiza la fotografía de un formato de DESPACHO DE PRODUCCIÓN DEL BIOTERIO.
Extrae cada fila diligenciada y responde ÚNICAMENTE con JSON válido, sin Markdown.

Estructura exacta:
{
  "registros": [
    {
      "fecha": "",
      "produccion_a_despachar": "",
      "cantidad_individuos": "",
      "peso": "",
      "destino": "",
      "encargado_bioterio": "",
      "encargado_zona": "",
      "observaciones": ""
    }
  ]
}

Reglas:
- Incluye una entrada por cada fila con datos. Ignora filas vacías.
- No inventes datos. Si una celda está vacía, usa "".
- Si una FILA COMPLETA está tachada, no la incluyas. El operario corrige en la fila siguiente.
- La hoja puede estar inclinada y lo escrito en "Destino" y en los encargados puede quedar un poco más arriba o más abajo de su línea. Para armar cada fila, cuenta las filas desde arriba en cada columna: la fila 1 de "Fecha" va con la fila 1 de "Destino" y de los encargados, la fila 2 con la fila 2, y así. Una fila tachada ocupa su lugar en todas las columnas aunque no la incluyas, para que las filas de abajo no se corran.
- Si UNA CELDA está tachada, escribe exactamente "tachón" en ese campo. No leas ni adivines el valor tachado. No agregues observaciones sobre tachones.
- "cantidad_individuos": solo si la celda es un número sin unidad de peso (por ejemplo "12" o "1"). Escríbelo tal como aparece.
- "peso": si la celda trae una referencia a gramos (g, gr, grs, gramos o similar, por ejemplo "180 gr" o "110g"), escribe ese valor tal como aparece en "peso" y deja "cantidad_individuos" vacío.
- "produccion_a_despachar" y "destino": escribe exactamente lo que aparece, sin corregir ortografía, mayúsculas, singular o plural.
- No escribas observaciones sobre ortografía, mayúsculas, plural o nombre de especie o destino.
- Si hay sobreescritura, borrones o dudas en una celda no tachada, indícalo en observaciones.
- No corrijas silenciosamente ningún dato.
`;

// Error especial: cuota agotada en todos los modelos. No seguir con más fotos hoy.
class CuotaAgotadaError extends Error {
  constructor(mensaje) {
    super(mensaje);
    this.name = "CuotaAgotadaError";
  }
}

// ================== FUNCIÓN PRINCIPAL ==================
// Es la función que el activador ejecuta cada 15 minutos y que usa el botón "Leer fotos nuevas".
function revisarCarpetasBioterio() {
  const raiz = DriveApp.getFolderById(CARPETA_RAIZ_ID);
  const procesados = obtenerProcesados();
  const subcarpetas = raiz.getFolders();

  while (subcarpetas.hasNext()) {
    const carpeta = subcarpetas.next();
    const tipo = carpeta.getName();

    if (!(tipo in TIPOS) || !TIPOS[tipo]) {
      Logger.log("Carpeta sin procesador: " + tipo);
      continue;
    }
    const procesar = TIPOS[tipo];

    const archivos = carpeta.getFiles();
    while (archivos.hasNext()) {
      const archivo = archivos.next();
      if (!archivo.getMimeType().startsWith("image/")) continue;
      if (procesados.has(archivo.getId())) continue;

      try {
        procesar(archivo);
        // Se registra inmediatamente después de guardar los datos.
        registrarProcesado(archivo, tipo);
        procesados.add(archivo.getId());
        Logger.log("Procesado: " + tipo + " / " + archivo.getName());
      } catch (e) {
        Logger.log("Pendiente " + archivo.getName() + ": " + e.message);
        if (e.name === "CuotaAgotadaError") {
          Logger.log("Revisión detenida: cuota de Gemini agotada en todos los modelos.");
          return;
        }
      }
    }
  }
}

// ================== PROCESADORES ==================
function procesarRoedores(archivo) {
  const datos = analizarImagen(archivo, PROMPT_ROEDORES);
  datos.registros = datos.registros.map(function(r) {
    const v = validarRegistro(r);
    r.requiere_revision = v.requiere_revision;
    r.alertas = v.alertas;
    return r;
  });
  guardarEnRevisionIA(datos, archivo.getName(), archivo.getId());
}

function procesarDespachos(archivo) {
  const datos = analizarImagen(archivo, PROMPT_DESPACHO);
  const registros = datos.registros.map(validarDespacho);
  guardarDespachoIA(registros, archivo.getName(), archivo.getId());
}

// ================== GEMINI (con respaldo de modelo) ==================
function analizarImagen(archivo, prompt) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  if (!apiKey) throw new Error("No se encontró GEMINI_API_KEY.");

  const blob = archivo.getBlob();
  const payload = {
    contents: [{
      parts: [
        { text: prompt },
        { inline_data: { mime_type: blob.getContentType(), data: Utilities.base64Encode(blob.getBytes()) } }
      ]
    }],
    generationConfig: { temperature: 0, responseMimeType: "application/json" }
  };

  let ultimoError = "";
  let cuotaAgotada = 0;

  for (const modelo of MODELOS) {
    const url = "https://generativelanguage.googleapis.com/v1beta/models/" + modelo +
      ":generateContent?key=" + encodeURIComponent(apiKey);

    for (let intento = 1; intento <= INTENTOS_POR_MODELO; intento++) {
      const r = UrlFetchApp.fetch(url, {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });
      const codigo = r.getResponseCode();
      const texto = r.getContentText();
      Logger.log("Modelo " + modelo + " | Intento " + intento + " | HTTP " + codigo);

      if (codigo === 200) {
        const salida = JSON.parse(texto).candidates?.[0]?.content?.parts?.[0]?.text;
        if (!salida) throw new Error("Respuesta sin texto utilizable.");
        const datos = JSON.parse(limpiarJsonGenerico(salida));
        if (!Array.isArray(datos.registros)) throw new Error("JSON sin arreglo 'registros'.");
        Logger.log("Respondió el modelo: " + modelo);
        return datos;
      }

      ultimoError = "HTTP " + codigo + " (" + modelo + "): " + texto;

      // Cuota agotada en este modelo: no reintentar; probar el siguiente modelo.
      if (codigo === 429 && (texto.indexOf("RESOURCE_EXHAUSTED") !== -1 || texto.indexOf("exceeded your current quota") !== -1)) {
        cuotaAgotada++;
        Logger.log("Cuota agotada en " + modelo + ". Probando siguiente modelo.");
        break;
      }

      // Demanda temporal: esperar y reintentar el mismo modelo.
      if ((codigo === 503 || codigo === 429) && intento < INTENTOS_POR_MODELO) {
        Utilities.sleep(Math.min(Math.pow(2, intento) * 5000, 30000));
        continue;
      }

      // Demanda temporal agotada en este modelo: probar el siguiente.
      if (codigo === 503 || codigo === 429) {
        Logger.log("Modelo " + modelo + " saturado. Probando siguiente modelo.");
        break;
      }

      // Otros errores (400, 404, etc.) no se reintentan.
      throw new Error(ultimoError);
    }
  }

  if (cuotaAgotada === MODELOS.length) {
    throw new CuotaAgotadaError("Cuota diaria de Gemini agotada en todos los modelos. La foto quedará pendiente.");
  }
  throw new Error("Ningún modelo respondió. Último error: " + ultimoError);
}

function limpiarJsonGenerico(texto) {
  const limpio = texto.trim().replace(/^```json\s*/i, "").replace(/^```\s*/i, "").replace(/\s*```$/i, "");
  const inicio = limpio.indexOf("{");
  const fin = limpio.lastIndexOf("}");
  if (inicio === -1 || fin <= inicio) throw new Error("No se encontró JSON válido.");
  return limpio.substring(inicio, fin + 1);
}

// ================== HOJA DE CONTROL ==================
// Lee IDs de Control procesados y, como respaldo, de Despacho IA (columna N).
function obtenerProcesados() {
  const ids = new Set();

  const control = obtenerHojaControl();
  if (control.getLastRow() > 1) {
    control.getRange(2, 1, control.getLastRow() - 1, 1).getValues()
      .forEach(function(f) { if (f[0]) ids.add(String(f[0])); });
  }

  const despacho = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Despacho IA");
  if (despacho && despacho.getLastRow() > 1) {
    despacho.getRange(2, 14, despacho.getLastRow() - 1, 1).getValues()
      .forEach(function(f) { if (f[0]) ids.add(String(f[0])); });
  }
  return ids;
}

function registrarProcesado(archivo, tipo) {
  obtenerHojaControl().appendRow([archivo.getId(), archivo.getName(), tipo, new Date()]);
}

function obtenerHojaControl() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = libro.getSheetByName(HOJA_CONTROL);
  if (!hoja) {
    hoja = libro.insertSheet(HOJA_CONTROL);
    hoja.appendRow(["image_id", "nombre_archivo", "tipo", "fecha_procesamiento"]);
    hoja.setFrozenRows(1);
  }
  return hoja;
}

// Ejecutar una vez si Despacho IA tiene fotos que no están en Control procesados.
function reconstruirControlDesdeDespacho() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const despacho = libro.getSheetByName("Despacho IA");
  if (!despacho || despacho.getLastRow() < 2) {
    Logger.log("Despacho IA está vacía o no existe en ESTE archivo.");
    return;
  }

  const control = obtenerHojaControl();
  const yaRegistrados = new Set();
  if (control.getLastRow() > 1) {
    control.getRange(2, 1, control.getLastRow() - 1, 1).getValues()
      .forEach(function(f) { if (f[0]) yaRegistrados.add(String(f[0])); });
  }

  const filas = despacho.getRange(2, 1, despacho.getLastRow() - 1, 15).getValues();
  const nuevos = {};
  filas.forEach(function(f) {
    const id = String(f[13] || "").trim();
    if (!id || yaRegistrados.has(id) || nuevos[id]) return;
    nuevos[id] = [id, String(f[12] || ""), "Despachos", f[14] || new Date()];
  });

  const lista = Object.keys(nuevos).map(function(k) { return nuevos[k]; });
  if (lista.length) {
    control.getRange(control.getLastRow() + 1, 1, lista.length, 4).setValues(lista);
  }
  Logger.log("Fotos agregadas al control: " + lista.length);
}

// ================== ACTIVADOR ==================
// Ejecutar una sola vez. Crea la revisión automática cada 15 minutos.
function crearActivador() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === "revisarCarpetasBioterio") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("revisarCarpetasBioterio").timeBased().everyMinutes(15).create();
  Logger.log("Activador creado: revisarCarpetasBioterio cada 15 minutos.");
}
