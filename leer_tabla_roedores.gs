function leerTablaRoedores() {
  const apiKey = PropertiesService
    .getScriptProperties()
    .getProperty("GEMINI_API_KEY");

  if (!apiKey) {
    throw new Error("No se encontró GEMINI_API_KEY en Script Properties.");
  }

  // Debe ser el ID del archivo JPEG/PNG, no el ID de la carpeta.
  const imageId = "1x8c6kZW9k9RYe0VRwQWpsTOSK85tLtQi";
  const modelo = "gemini-3.6-flash";

  const archivo = DriveApp.getFileById(imageId);
  const blob = archivo.getBlob();
  const imagenBase64 = Utilities.base64Encode(blob.getBytes());
  const mimeType = blob.getContentType();

  if (!mimeType || !mimeType.startsWith("image/")) {
    throw new Error("El archivo indicado no parece ser una imagen. Tipo detectado: " + mimeType);
  }

  const prompt = `
Analiza cuidadosamente la fotografía de esta tabla de registros de roedores.

Extrae todos los registros visibles y responde ÚNICAMENTE con JSON válido,
sin explicaciones, sin Markdown y sin bloques de código.

Usa exactamente esta estructura:
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
- No inventes datos.
- Si un dato no se puede leer, usa "" para texto o null para números.
- Conserva las fechas tal como aparecen, pero no incluyas símbolos ♀ o ♂ dentro de las fechas.
- Si una fecha es claramente legible como día/mes/año, puedes escribirla como aparece en la imagen.
- Si hay tachones, sobreescritura, borrones, guiones o dudas, indícalos en observaciones.
- No corrijas silenciosamente datos ni determines la causa de una inconsistencia.
- La validación matemática será realizada por Apps Script.
- Devuelve siempre un objeto JSON con la propiedad "registros".
`;

  const url =
    "https://generativelanguage.googleapis.com/v1beta/models/" +
    modelo +
    ":generateContent?key=" +
    encodeURIComponent(apiKey);

  const payload = {
    contents: [{
      parts: [
        { text: prompt },
        { inline_data: { mime_type: mimeType, data: imagenBase64 } }
      ]
    }],
    generationConfig: {
      temperature: 0,
      responseMimeType: "application/json"
    }
  };

  const maxIntentos = 3;
  let ultimoError = "";

  for (let intento = 1; intento <= maxIntentos; intento++) {
    try {
      const respuesta = UrlFetchApp.fetch(url, {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });

      const codigo = respuesta.getResponseCode();
      const texto = respuesta.getContentText();
      Logger.log("Intento " + intento + " | Código HTTP: " + codigo);

      if (codigo === 200) {
        const datos = JSON.parse(texto);
        const salida = datos.candidates?.[0]?.content?.parts?.[0]?.text;

        if (!salida) {
          throw new Error("La respuesta no contiene texto utilizable.");
        }

        const registros = JSON.parse(limpiarRespuestaJson(salida));

        if (!registros.registros || !Array.isArray(registros.registros)) {
          throw new Error("El JSON no contiene un arreglo válido llamado 'registros'.");
        }

        registros.registros = registros.registros.map(function(registro) {
          const validacion = validarRegistro(registro);
          registro.requiere_revision = validacion.requiere_revision;
          registro.alertas = validacion.alertas;
          return registro;
        });

        Logger.log("JSON EXTRAÍDO Y VALIDADO:");
        Logger.log(JSON.stringify(registros, null, 2));

        // Guarda el resultado en una hoja de revisión, no en la base definitiva.
        guardarEnRevisionIA(registros, archivo.getName(), imageId);
        return registros;
      }

      ultimoError = texto;

      if (codigo === 503 || codigo === 429) {
        if (intento < maxIntentos) {
          const espera = Math.pow(2, intento) * 5000;
          Logger.log("Servicio temporalmente ocupado. Esperando " + espera / 1000 + " segundos...");
          Utilities.sleep(espera);
          continue;
        }
        break;
      }

      throw new Error("Error HTTP " + codigo + ": " + texto);

    } catch (error) {
      ultimoError = error.message;
      Logger.log("Error en intento " + intento + ": " + error.message);
      throw error;
    }
  }

  throw new Error("No fue posible obtener el JSON después de " + maxIntentos + " intentos. Último error: " + ultimoError);
}


function guardarEnRevisionIA(datos, nombreArchivo, imageId) {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const nombreHoja = "Revisión IA";
  let hoja = libro.getSheetByName(nombreHoja);

  if (!hoja) {
    hoja = libro.insertSheet(nombreHoja);
  }

  const encabezados = [
    "fecha_procesamiento", "archivo_origen", "image_id", "indice_registro",
    "fecha_nacimiento", "total_crias", "total_natimuertos", "fecha_destete",
    "total_hembras", "total_machos", "observaciones", "requiere_revision",
    "tipo_alerta", "detalle_alerta", "estado_revision", "comentario_responsable"
  ];

  if (hoja.getLastRow() === 0) {
    hoja.appendRow(encabezados);
    hoja.setFrozenRows(1);
  }

  const ahora = new Date();

  datos.registros.forEach(function(registro, indice) {
    const tipos = (registro.alertas || []).map(a => a.tipo).join(" | ");
    const detalles = (registro.alertas || []).map(a => a.mensaje).join(" | ");

    hoja.appendRow([
      ahora,
      nombreArchivo,
      imageId,
      indice + 1,
      registro.fecha_nacimiento || "",
      registro.total_crias ?? "",
      registro.total_natimuertos ?? "",
      registro.fecha_destete || "",
      registro.total_hembras ?? "",
      registro.total_machos ?? "",
      registro.observaciones || "",
      registro.requiere_revision ? "Sí" : "No",
      tipos,
      detalles,
      "Pendiente",
      ""
    ]);
  });

  hoja.autoResizeColumns(1, encabezados.length);
}


function limpiarRespuestaJson(texto) {
  let limpio = texto.trim();
  limpio = limpio.replace(/^```json\s*/i, "");
  limpio = limpio.replace(/^```\s*/i, "");
  limpio = limpio.replace(/\s*```$/i, "");

  const inicio = limpio.indexOf("{");
  const final = limpio.lastIndexOf("}");

  if (inicio === -1 || final === -1 || final <= inicio) {
    throw new Error("La respuesta no contiene un objeto JSON válido.");
  }

  return limpio.substring(inicio, final + 1);
}


function validarRegistro(registro) {
  const alertas = [];
  const totalCrias = numeroValido(registro.total_crias);
  const totalNatimuertos = numeroValido(registro.total_natimuertos);
  const hembras = numeroValido(registro.total_hembras);
  const machos = numeroValido(registro.total_machos);
  const observaciones = String(registro.observaciones || "").toLowerCase();

  if (totalCrias !== null && hembras !== null && machos !== null) {
    const sumaSexos = hembras + machos;
    if (sumaSexos !== totalCrias) {
      alertas.push({
        tipo: "inconsistencia_cantidad",
        mensaje: "El total de crías es " + totalCrias + ", pero hembras + machos suma " + sumaSexos + ".",
        posibles_explicaciones: [
          "Error de captura o conteo",
          "Nacieron crías después del conteo inicial",
          "No se actualizó el total de crías",
          "Los conteos se realizaron en momentos diferentes"
        ],
        accion: "Verificar contra el registro físico"
      });
    }
  }

  if (totalNatimuertos !== null && totalCrias !== null && totalNatimuertos > totalCrias) {
    alertas.push({
      tipo: "natimuertos_mayor_que_crias",
      mensaje: "El número de natimuertos es mayor que el total de crías.",
      accion: "Verificar el registro físico"
    });
  }

  [
    ["total_crias", totalCrias],
    ["total_natimuertos", totalNatimuertos],
    ["total_hembras", hembras],
    ["total_machos", machos]
  ].forEach(function(campo) {
    if (campo[1] !== null && campo[1] < 0) {
      alertas.push({
        tipo: "valor_negativo",
        mensaje: "El campo " + campo[0] + " contiene un valor negativo.",
        accion: "Verificar el registro físico"
      });
    }
  });

  // Datos faltantes o señales de lectura dudosa siempre requieren revisión.
  if (registro.total_crias === null || registro.total_crias === undefined || registro.total_crias === "") {
    alertas.push({ tipo: "dato_faltante", mensaje: "No se pudo leer el total de crías.", accion: "Verificar el registro físico" });
  }

  if (hembras === null || machos === null) {
    alertas.push({ type: "dato_faltante", tipo: "sexo_faltante", mensaje: "Falta el total de hembras o machos; no se puede comprobar la distribución por sexo.", accion: "Verificar el registro físico" });
  }

  const palabrasDudosas = ["tach", "sobreescrit", "borron", "borra", "guion", "guión", "duda", "ilegible", "no se puede leer"];
  if (palabrasDudosas.some(palabra => observaciones.includes(palabra))) {
    alertas.push({
      tipo: "lectura_dudosa",
      mensaje: "La observación indica tachaduras, sobreescritura, guiones o dudas de lectura.",
      accion: "Verificar el registro físico"
    });
  }

  return { requiere_revision: alertas.length > 0, alertas: alertas };
}


function numeroValido(valor) {
  if (valor === null || valor === undefined || valor === "") return null;
  const numero = Number(valor);
  return Number.isFinite(numero) ? numero : null;
}
