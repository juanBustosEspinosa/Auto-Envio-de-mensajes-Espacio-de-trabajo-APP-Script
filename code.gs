const NOMBRE_ESPACIO_ORIGEN = "Prueba"; //nombre del espacio de trabajo
const ID_CARPETA_DESTINO = ""; //carpeta del drive 

function datosSheet(number) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
  if (!sheet) {
    console.error('[ERROR datosSheet] No se encontró la hoja "Sheet1"');
    return [];
  }
  
  var ultimaFila = sheet.getLastRow();
  if (ultimaFila < 2) {
    return [];
  }
  
  var datosA = sheet.getRange(2, number, ultimaFila - 1, 1).getValues();
  
  var arrayDeStrings = datosA
    .flat()
    .map(function(item) {
      return String(item).trim();
    })
    .filter(function(item) {
      return item !== "";
    });
  
  return arrayDeStrings;
}

var REGLAS_RUTEO = [
  {
    keywords: [],
    webhooks: [
      ''
    ]
  }
];
//webhooks son los enlaces de los espacios de trabajo de destino
//keywords las recoge del sheet pero principalmento son por lo que va a filtrar el codigo


function ejecutarProcesoContinuo() {
  for (var i = 0; i < REGLAS_RUTEO.length; i++) {
    REGLAS_RUTEO[i].keywords = datosSheet(i + 1);
  }

  const token = ScriptApp.getOAuthToken();
  const headers = {
    "Authorization": "Bearer " + token,
    "Content-Type": "application/json"
  };

  try {
    const resEspacios = UrlFetchApp.fetch("https://chat.googleapis.com/v1/spaces", { headers, muteHttpExceptions: true });
    const dataEspacios = JSON.parse(resEspacios.getContentText());

    if (dataEspacios.error) {
      console.error("[ERROR API Google Chat]: " + JSON.stringify(dataEspacios.error));
      return;
    }

    const espacioOrigen = (dataEspacios.spaces || []).find(s =>
      s.displayName && s.displayName.trim().toLowerCase() === NOMBRE_ESPACIO_ORIGEN.trim().toLowerCase()
    );

    if (!espacioOrigen) {
      console.error(`[ERROR] No se encontró el espacio "${NOMBRE_ESPACIO_ORIGEN}"`);
      return;
    }

    const props = PropertiesService.getScriptProperties();
    let ultimaFechaProcesadaMs = parseInt(props.getProperty('ULTIMA_FECHA_MS'), 10);

    if (isNaN(ultimaFechaProcesadaMs)) {
      ultimaFechaProcesadaMs = Date.now() - (5 * 60 * 1000); 
    }

    let mensajesEvaluados = [];
    let pageToken = "";
    let seguirBuscando = true;

    while (seguirBuscando) {
      let urlMensajes = `https://chat.googleapis.com/v1/${espacioOrigen.name}/messages?pageSize=100&orderBy=createTime%20desc`;
      if (pageToken) {
        urlMensajes += `&pageToken=${pageToken}`;
      }

      const resMensajes = UrlFetchApp.fetch(urlMensajes, { headers, muteHttpExceptions: true });
      const dataMensajes = JSON.parse(resMensajes.getContentText());

      if (dataMensajes.error) {
        console.error("[ERROR al leer mensajes]: " + dataMensajes.error.message);
        break;
      }

      const paginaMensajes = dataMensajes.messages || [];
      if (paginaMensajes.length === 0) break;

      for (const msg of paginaMensajes) {
        const fechaCreacionMs = Date.parse(msg.createTime);

        if (fechaCreacionMs <= ultimaFechaProcesadaMs) {
          seguirBuscando = false;
          break;
        }

        mensajesEvaluados.push(msg);
      }

      pageToken = dataMensajes.nextPageToken || "";
      if (!pageToken || !seguirBuscando) {
        break;
      }
    }

    if (mensajesEvaluados.length === 0) {
      return;
    }

    mensajesEvaluados.reverse();

    let masRecienteMs = ultimaFechaProcesadaMs;

    for (const msg of mensajesEvaluados) {
      const fechaMsgMs = Date.parse(msg.createTime);
      const texto = msg.text || '';

      // --- FILTRO DE RESPUESTAS Y CITAS ---
      let esRespuesta = false;

      // 1. Detectar si es una cita en línea
      if (msg.quotedMessageMetadata || (msg.threadReply !== undefined && msg.threadReply !== null)) {
        esRespuesta = true;
      }

      // 2. Detectar si es respuesta en hilo
      if (!esRespuesta && msg.thread && msg.thread.name) {
        const idHilo = msg.thread.name.split('/').pop();
        const idMensaje = msg.name.split('/').pop();
        const identificadorPrincipal = `${idHilo}.${idHilo}`;

        if (idMensaje !== identificadorPrincipal && idMensaje !== idHilo) {
          esRespuesta = true;
        }
      }

      if (esRespuesta) {
        console.log(`[OMITIDO - RESPUESTA O CITA]: "${texto}"`);
        if (fechaMsgMs > masRecienteMs) masRecienteMs = fechaMsgMs;
        continue;
      }

      // --- EVALUACIÓN DE KEYWORDS ---
      const webhooksDestino = new Set();

      for (const regla of REGLAS_RUTEO) {
        let coincide = false;

        if (regla.keywords.length === 0) {
          coincide = true;
        } else {
          coincide = regla.keywords.some(k => {
            const regex = new RegExp(`\\b${k.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\b`, 'i');
            return regex.test(texto);
          });
        }

        if (coincide) {
          regla.webhooks.forEach(w => webhooksDestino.add(w));
        }
      }

      if (webhooksDestino.size > 0) {
        console.log(`[REENVIANDO MENSAJE PRINCIPAL]: "${texto || '(mensaje con adjunto)'}"`);

        const payload = construirPayloadWebhook(msg, token);

        for (const urlWebhook of webhooksDestino) {
          UrlFetchApp.fetch(urlWebhook, {
            method: "post",
            contentType: "application/json; charset=UTF-8",
            payload: JSON.stringify(payload),
            muteHttpExceptions: true
          });
        }
      }

      if (fechaMsgMs > masRecienteMs) {
        masRecienteMs = fechaMsgMs;
      }
    }

    props.setProperty('ULTIMA_FECHA_MS', masRecienteMs.toString());

  } catch (e) {
    console.error("Error general: " + e.message);
  }
}

function obtenerCarpetaPorId(idCarpeta) {
  try {
    return DriveApp.getFolderById(idCarpeta);
  } catch (e) {
    console.error(`[ERROR] No se pudo acceder a la carpeta de Drive con ID "${idCarpeta}": ` + e.message);
    return null;
  }
}

function subirAdjuntoADrive(adjunto, token) {
  try {
    let urlDescarga = null;

    if (adjunto.attachmentDataRef && adjunto.attachmentDataRef.resourceName) {
      urlDescarga = `https://chat.googleapis.com/v1/media/${adjunto.attachmentDataRef.resourceName}?alt=media`;
    } else if (adjunto.name) {
      urlDescarga = `https://chat.googleapis.com/v1/media/${adjunto.name}?alt=media`;
    } else if (adjunto.downloadUri) {
      urlDescarga = adjunto.downloadUri;
    }

    if (!urlDescarga) return null;

    const res = UrlFetchApp.fetch(urlDescarga, {
      headers: { "Authorization": "Bearer " + token },
      muteHttpExceptions: true
    });

    if (res.getResponseCode() !== 200) {
      console.error(`[ERROR Descarga HTTP ${res.getResponseCode()}]: ${res.getContentText()}`);
      return null;
    }

    const blob = res.getBlob();

    const nombreOriginal = adjunto.contentName || "";
    let extension = "";
    if (nombreOriginal.includes(".")) {
      extension = nombreOriginal.substring(nombreOriginal.lastIndexOf("."));
    } else {
      const mime = blob.getContentType();
      if (mime.includes("png")) extension = ".png";
      else if (mime.includes("jpeg") || mime.includes("jpg")) extension = ".jpg";
      else if (mime.includes("gif")) extension = ".gif";
    }

    const fechaHora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HH-mm-ss");
    const nuevoNombreArchivo = `Imagen_${fechaHora}${extension}`;
    blob.setName(nuevoNombreArchivo);

    const carpeta = obtenerCarpetaPorId(ID_CARPETA_DESTINO);
    if (!carpeta) return null;

    const archivoDrive = carpeta.createFile(blob);

    try {
      archivoDrive.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (errPerm) {
      try {
        archivoDrive.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (e) {
        // Ignorar
      }
    }

    return {
      url: archivoDrive.getUrl(),
      nombre: nuevoNombreArchivo
    };
  } catch (e) {
    console.error("[ERROR subirAdjuntoADrive]: " + e.message);
    return null;
  }
}

function construirPayloadWebhook(msg, token) {
  const textoOriginal = msg.text || '';
  const adjuntos = msg.attachment || msg.attachments || [];

  const archivosDrive = [];

  for (const adjunto of adjuntos) {
    const info = subirAdjuntoADrive(adjunto, token);
    if (info) archivosDrive.push(info);
  }

  let mensajeFinal = `*Mensaje redirigido desde ${NOMBRE_ESPACIO_ORIGEN}:*\n`;

  if (textoOriginal) {
    mensajeFinal += `${textoOriginal}\n`;
  }

  if (archivosDrive.length > 0) {
    mensajeFinal += `\n📎 *Archivos adjuntos en Google Drive:*`;
    archivosDrive.forEach(archivo => {
      mensajeFinal += `\n• [${archivo.nombre}](${archivo.url})`;
    });
  }

  return { text: mensajeFinal };
}

function resetearMarcaTiempo() {
  PropertiesService.getScriptProperties().deleteProperty('ULTIMA_FECHA_MS');
  console.log("Marca de tiempo reiniciada.");
}
