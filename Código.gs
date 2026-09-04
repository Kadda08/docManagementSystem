/************************************************************
 * GESTOR DOCUMENTAL
 * Google Drive + Apps Script + Gemini + Google Sheets
 *
 * VERSIÓN 1.0
 *
 * FUNCIONES PRINCIPALES:
 *
 * 0. Autorizacion user 
 * 1. Interfaz Web
 * 2. Subir documentos
 * 3. Procesar PDF / JPG / JPEG / PNG
 * 4. Procesar DOCX
 * 5. OCR mediante Gemini
 * 6. Crear índice documental
 * 7. Buscar coincidencias
 * 8. Abrir documentos originales de Drive
 ************************************************************/


/************************************************************
 * CONFIGURACIÓN
 ************************************************************/

const CONFIG = {
  // ========================================================
  // CORREOS PERMITIDOS
  // ========================================================
  ALLOWED_USERS: [
    'zu.joel@gmail.com',
    'adaportich@gmail.com',
    'lo0ky.448@gmail.com'
  ],


  // ========================================================
  // ID DE LA CARPETA PRINCIPAL DE GOOGLE DRIVE
  // ========================================================

  FOLDER_ID: '1vyElmafAN0boahlM7My-wCW-k5oqGgJH',


  // ========================================================
  // MODELO GEMINI
  // ========================================================

  GEMINI_MODEL: 'gemini-3.7-flash',


  // ========================================================
  // NOMBRE DEL GOOGLE SHEET
  // ========================================================

  SHEET_NAME: 'INDICE_DOCUMENTOS',


  // ========================================================
  // NOMBRE DEL ARCHIVO GOOGLE SHEETS
  // ========================================================

  SPREADSHEET_NAME: 'INDICE - GESTOR DOCUMENTAL',


  // ========================================================
  // TIPOS DE ARCHIVO PERMITIDOS
  // ========================================================

  ALLOWED_MIME_TYPES: [

    'application/pdf',

    'image/jpeg',

    'image/png',

    'image/webp',

    'application/vnd.openxmlformats-officedocument.wordprocessingml.document'

  ]

};



/************************************************************
 * 0. VERIFICAR USUARIO AUTORIZADO
 ************************************************************/

function verificarUsuarioAutorizado() {

  const email =
    Session
      .getActiveUser()
      .getEmail()
      .toLowerCase()
      .trim();

  if (!email) {

    throw new Error(
      'No se pudo identificar tu cuenta de Google. ' +
      'Debes iniciar sesión con una cuenta autorizada.'
    );

  }

  const usuarios =
    CONFIG.ALLOWED_USERS.map(
      function(usuario) {
        return usuario.toLowerCase().trim();
      }
    );

  if (
    usuarios.indexOf(email) === -1
  ) {

    throw new Error(
      'Acceso denegado.\n\n' +
      'La cuenta ' +
      email +
      ' no está autorizada para utilizar este sistema.'
    );

  }

  return email;

}

/************************************************************
 * 1. INTERFAZ WEB
 ************************************************************/

function doGet() {

  // ======================================================
  // SEGURIDAD: no servir la interfaz a usuarios no
  // autorizados. Antes esta verificación no existía en
  // doGet(), por lo que cualquiera podía abrir la app.
  // ======================================================
  try {

    verificarUsuarioAutorizado();

  } catch (error) {

    return HtmlService
      .createHtmlOutput(
        '<p style="font-family:sans-serif;color:#b00020;">' +
        error.message.replace(/\n/g, '<br>') +
        '</p>'
      )
      .setTitle('Acceso denegado');

  }

  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('Gestor Documental')
    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    );

}


/************************************************************
 * 2. OBTENER CONFIGURACIÓN
 ************************************************************/

function obtenerConfiguracion() {

  // Antes no se verificaba el usuario aquí: cualquiera
  // que llamara a esta función desde el cliente podía
  // leer el FOLDER_ID y el modelo configurado.
  verificarUsuarioAutorizado();

  verificarConfiguracion();

  return {

    sistema: 'Gestor Documental',

    estado: 'ACTIVO',

    carpeta: CONFIG.FOLDER_ID,

    modelo: CONFIG.GEMINI_MODEL

  };

}


/************************************************************
 * 3. SUBIR DOCUMENTO DESDE LA INTERFAZ
 ************************************************************/

function subirDocumento(formObject) {

  // ======================================================
  // SEGURIDAD
  // ======================================================
  verificarUsuarioAutorizado();
  verificarConfiguracion();

  // ======================================================
  // VALIDAR ARCHIVO
  // ======================================================
  if (
    !formObject ||
    !formObject.archivo
  ) {

    throw new Error(
      'No se seleccionó ningún archivo.'
    );

  }


  const blob =
    formObject.archivo;


  const nombre =
    blob.getName();


  const mimeType =
    blob.getContentType();


  // ========================================================
  // VALIDAR FORMATO
  // ========================================================

  if (
    CONFIG.ALLOWED_MIME_TYPES
      .indexOf(mimeType) === -1
  ) {

    throw new Error(
      'Formato no permitido: ' +
      mimeType +
      '\n\n' +
      'Formatos permitidos:\n' +
      'PDF, JPG, JPEG, PNG, WEBP y DOCX.'
    );

  }

  // ========================================================
  // OBTENER SOLO LA CARPETA
  // ========================================================

  const folder =
    DriveApp.getFolderById(
      CONFIG.FOLDER_ID
    );


  // ========================================================
  // GUARDAR ARCHIVO ORIGINAL
  // ========================================================

  const file =
    folder.createFile(blob);


  Logger.log(
    'Archivo subido: ' +
    file.getName()
  );


  // ========================================================
  // PROCESAR DOCUMENTO
  // ========================================================

  let resultado;
  try {

    resultado =
      procesarArchivo(file);


  } catch (error) {

    // Si falla el procesamiento,
    // el archivo original permanece en Drive.

    registrarError(
      file,
      error
    );

    throw error;

  }


  return {

    success: true,

    id: file.getId(),

    nombre: file.getName(),

    tipo: file.getMimeType(),

    url: file.getUrl(),

    mensaje:
      'Documento almacenado correctamente.',

    textoExtraido:
      resultado.texto,

    coincidencias: 0

  };

}

/************************************************************
 * 4. PROCESAR ARCHIVO
 ************************************************************/

function procesarArchivo(file) {

  const mimeType =
    file.getMimeType();


  Logger.log(
    'Procesando: ' +
    file.getName()
  );


  // ========================================================
  // PDF / IMAGEN
  // ========================================================

  if (

    mimeType ===
      'application/pdf'

    ||

    mimeType ===
      'image/jpeg'

    ||

    mimeType ===
      'image/png'

    ||

    mimeType ===
      'image/webp'

  ) {

    const texto =
      procesarDocumentoConGemini(
        file
      );


    const palabrasClave =
      extraerPalabrasClave(
        texto
      );


    guardarDocumento(
      file,
      texto,
      palabrasClave
    );


    return {

      texto: texto,

      palabrasClave:
        palabrasClave

    };

  }


  // ========================================================
  // DOCX
  // ========================================================

  if (

    mimeType ===
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document'

  ) {

    const texto =
      procesarDOCX(
        file
      );


    const palabrasClave =
      extraerPalabrasClave(
        texto
      );


    guardarDocumento(
      file,
      texto,
      palabrasClave
    );


    return {

      texto: texto,

      palabrasClave:
        palabrasClave

    };

  }


  throw new Error(
    'Tipo de archivo no soportado.'
  );

}


/************************************************************
 * 5. PROCESAR PDF / IMAGEN CON GEMINI
 ************************************************************/

function procesarDocumentoConGemini(file) {

  const apiKey =
    PropertiesService
      .getScriptProperties()
      .getProperty(
        'GEMINI_API_KEY'
      );


  if (!apiKey) {

    throw new Error(
      'No se encontró GEMINI_API_KEY en las propiedades del proyecto.'
    );

  }


  const blob =
    file.getBlob();


  const base64 =
    Utilities.base64Encode(
      blob.getBytes()
    );


  const mimeType =
    file.getMimeType();


  const prompt = `

Eres un sistema OCR especializado
en digitalización y análisis documental.

Analiza completamente el documento.

EXTRAE TODO EL TEXTO QUE PUEDAS LEER.

Reglas:

1. Conserva títulos.
2. Conserva nombres y apellidos.
3. Conserva fechas.
4. Conserva números.
5. Conserva números de documentos.
6. Conserva direcciones.
7. Conserva nombres de instituciones.
8. Conserva tablas cuando sea posible.
9. Conserva párrafos.
10. No inventes información.
11. No resumas.
12. No expliques el documento.
13. Devuelve solamente el texto extraído.
14. Respeta en lo posible el orden del documento.

El objetivo es crear un índice de búsqueda
para posteriormente encontrar palabras
dentro del documento.

`;


  const url =
    'https://generativelanguage.googleapis.com/v1beta/models/' +
    CONFIG.GEMINI_MODEL +
    ':generateContent';


  const payload = {

    contents: [

      {

        role: 'user',

        parts: [

          {

            text: prompt

          },

          {

            inlineData: {

              mimeType: mimeType,

              data: base64

            }

          }

        ]

      }

    ],

    generationConfig: {

      thinkingConfig: {

        thinkingLevel: 'low'

      },

      maxOutputTokens: 65000

    }

  };


  const options = {

    method: 'POST',

    contentType:
      'application/json',

    headers: {

      'x-goog-api-key':
        apiKey

    },

    payload:
      JSON.stringify(payload),

    muteHttpExceptions: true

  };


  const response =
    UrlFetchApp.fetch(
      url,
      options
    );


  const status =
    response.getResponseCode();


  const responseText =
    response.getContentText();


  if (
    status < 200 ||
    status >= 300
  ) {

    throw new Error(

      'Gemini API error ' +
      status +
      ':\n\n' +
      responseText

    );

  }


  const data =
    JSON.parse(
      responseText
    );


  if (
    !data.candidates ||
    data.candidates.length === 0
  ) {

    throw new Error(
      'Gemini no devolvió ningún resultado.'
    );

  }


  let texto = '';


  const parts =
    data.candidates[0]
      .content
      .parts;


  parts.forEach(
    function(part) {

      if (part.text) {

        texto +=
          part.text;

      }

    }
  );


  texto =
    texto.trim();


  if (!texto) {

    throw new Error(
      'Gemini no pudo extraer texto del documento.'
    );

  }


  return texto;

}


/************************************************************
 * 6. PROCESAR DOCX
 *
 * Convierte temporalmente el DOCX en Google Docs
 * para poder extraer su contenido.
 ************************************************************/

function procesarDOCX(file) {

  Logger.log(
    'Procesando DOCX: ' +
    file.getName()
  );


  const blob =
    file.getBlob();


  const metadata = {

    name:
      'TEMP_' +
      file.getName(),

    mimeType:
      'application/vnd.google-apps.document'

  };


  let tempDoc;


  try {

    // ======================================================
    // REQUIERE SERVICIO AVANZADO DE DRIVE
    // ======================================================

    tempDoc =
      Drive.Files.create(
        metadata,
        blob,
        {
          fields: 'id,name'
        }
      );


  } catch (error) {

    throw new Error(

      'No se pudo convertir el DOCX.\n\n' +

      'Debes habilitar el Servicio avanzado de Drive ' +
      'en Apps Script.\n\n' +

      'Error original:\n' +
      error.message

    );

  }


  const tempId =
    tempDoc.id;


  try {

    Utilities.sleep(1500);


    const documento =
      DocumentApp.openById(
        tempId
      );


    const texto =
      documento
        .getBody()
        .getText();


    if (!texto) {

      throw new Error(
        'El DOCX no contiene texto legible.'
      );

    }


    return texto.trim();


  } finally {

    // ======================================================
    // ELIMINAR DOCUMENTO TEMPORAL
    // ======================================================

    try {

      DriveApp
        .getFileById(tempId)
        .setTrashed(true);

    } catch (error) {

      Logger.log(
        'No se pudo eliminar temporal: ' +
        error.message
      );

    }

  }

}


/************************************************************
 * 7. EXTRAER PALABRAS CLAVE
 ************************************************************/

function extraerPalabrasClave(texto) {

  if (!texto) {

    return '';

  }


  const palabras =
    texto

      .toLowerCase()

      .replace(
        /[^a-záéíóúüñ0-9\s]/gi,
        ' '
      )

      .split(/\s+/);


  const stopWords = [

    'para',
    'como',
    'entre',
    'desde',
    'hasta',
    'sobre',
    'este',
    'esta',
    'estos',
    'estas',
    'donde',
    'cuando',
    'tiene',
    'tienen',
    'también',
    'porque',
    'documento',
    'documentos',
    'según',
    'siendo',
    'dicho',
    'dicha',
    'todos',
    'todas'

  ];


  const contador = {};


  palabras.forEach(
    function(palabra) {

      if (
        palabra.length < 5
      ) {

        return;

      }


      if (
        stopWords.indexOf(
          palabra
        ) !== -1
      ) {

        return;

      }


      contador[palabra] =
        (
          contador[palabra] ||
          0
        ) + 1;

    }
  );


  return Object
    .keys(contador)

    .sort(
      function(a, b) {

        return (
          contador[b] -
          contador[a]
        );

      }
    )

    .slice(0, 30)

    .join(', ');

}


/************************************************************
 * 8. GUARDAR DOCUMENTO EN GOOGLE SHEETS
 ************************************************************/

function guardarDocumento(
  file,
  texto,
  palabrasClave
) {

  const spreadsheetId =
    obtenerSpreadsheetId();


  const spreadsheet =
    SpreadsheetApp.openById(
      spreadsheetId
    );


  let sheet =
    spreadsheet.getSheetByName(
      CONFIG.SHEET_NAME
    );


  // ========================================================
  // CREAR HOJA SI NO EXISTE
  // ========================================================

  if (!sheet) {

    sheet =
      spreadsheet.insertSheet(
        CONFIG.SHEET_NAME
      );


    sheet.appendRow([

      'ID_DRIVE',

      'NOMBRE',

      'TIPO',

      'FECHA',

      'TAMAÑO',

      'TEXTO_EXTRAIDO',

      'PALABRAS_CLAVE',

      'URL_DRIVE',

      'ESTADO',

      'FECHA_PROCESAMIENTO'

    ]);

  }


  // ========================================================
  // EVITAR DUPLICADOS
  // ========================================================

  eliminarRegistroAnterior(
    sheet,
    file.getId()
  );


  // ========================================================
  // INSERTAR DOCUMENTO
  // ========================================================

  sheet.appendRow([

    file.getId(),

    file.getName(),

    file.getMimeType(),

    file.getDateCreated(),

    file.getSize(),

    texto,

    palabrasClave,

    file.getUrl(),

    'PROCESADO',

    new Date()

  ]);

}


/************************************************************
 * 9. OBTENER ID DEL GOOGLE SHEETS
 ************************************************************/

function obtenerSpreadsheetId() {

  const properties =
    PropertiesService
      .getScriptProperties();


  let spreadsheetId =
    properties.getProperty(
      'SPREADSHEET_ID'
    );


  if (spreadsheetId) {

    return spreadsheetId;

  }


  // ========================================================
  // SI NO ESTÁ GUARDADO, BUSCAR POR NOMBRE
  // ========================================================

  const files =
    DriveApp
      .getFilesByName(
        CONFIG.SPREADSHEET_NAME
      );


  if (!files.hasNext()) {

    throw new Error(

      'No se encontró el Google Sheet:\n\n' +

      CONFIG.SPREADSHEET_NAME

    );

  }


  const file =
    files.next();


  spreadsheetId =
    file.getId();


  properties.setProperty(

    'SPREADSHEET_ID',

    spreadsheetId

  );


  return spreadsheetId;

}


/************************************************************
 * 10. ELIMINAR REGISTRO ANTERIOR
 ************************************************************/

function eliminarRegistroAnterior(
  sheet,
  fileId
) {

  const lastRow =
    sheet.getLastRow();


  if (lastRow <= 1) {

    return;

  }


  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        1
      )
      .getValues();


  for (
    let i = values.length - 1;
    i >= 0;
    i--
  ) {

    if (
      String(values[i][0]) ===
      String(fileId)
    ) {

      sheet.deleteRow(
        i + 2
      );

    }

  }

}


/************************************************************
 * 11. BUSCADOR
 ************************************************************/

function buscarDocumentos(
  consulta
) {

  verificarUsuarioAutorizado();
  verificarConfiguracion();


  if (
    consulta === null ||
    consulta === undefined
  ) {

    throw new Error(
      'Debes escribir algo para buscar.'
    );

  }


  consulta =
    String(
      consulta
    ).trim();


  if (!consulta) {

    throw new Error(
      'La búsqueda está vacía.'
    );

  }


  const spreadsheetId =
    obtenerSpreadsheetId();


  const spreadsheet =
    SpreadsheetApp.openById(
      spreadsheetId
    );


  const sheet =
    spreadsheet.getSheetByName(
      CONFIG.SHEET_NAME
    );


  if (!sheet) {

    throw new Error(
      'No existe la hoja ' +
      CONFIG.SHEET_NAME
    );

  }


  const lastRow =
    sheet.getLastRow();


  if (lastRow <= 1) {

    return [];

  }


  const lastColumn =
    sheet.getLastColumn();


  const rows =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        lastColumn
      )
      .getValues();


  const resultados = [];


  const busqueda =
    consulta.toLowerCase();


  rows.forEach(
    function(row) {

      const id =
        row[0];

      const nombre =
        row[1] || '';

      const tipo =
        row[2] || '';

      const texto =
        row[5] || '';

      const palabrasClave =
        row[6] || '';

      const url =
        row[7] || '';


      const contenido =

        String(texto) +
        ' ' +
        String(palabrasClave);


      const contenidoLower =
        contenido.toLowerCase();


      if (
        contenidoLower
          .indexOf(busqueda) === -1
      ) {

        return;

      }


      const cantidad =
        contarCoincidencias(
          contenidoLower,
          busqueda
        );


      const fragmento =
        crearFragmento(
          texto,
          consulta
        );


      resultados.push({

        id: id,

        nombre: nombre,

        tipo: tipo,

        url: url,

        coincidencias:
          cantidad,

        fragmento:
          fragmento

      });

    }
  );


  // ========================================================
  // ORDENAR POR MAYOR CANTIDAD DE COINCIDENCIAS
  // ========================================================

  resultados.sort(
    function(a, b) {

      return (
        b.coincidencias -
        a.coincidencias
      );

    }
  );


  return resultados;

}


/************************************************************
 * 12. CONTAR COINCIDENCIAS
 ************************************************************/

function contarCoincidencias(
  texto,
  consulta
) {

  if (!consulta) {

    return 0;

  }


  let posicion = 0;

  let contador = 0;


  while (true) {

    posicion =
      texto.indexOf(
        consulta,
        posicion
      );


    if (posicion === -1) {

      break;

    }


    contador++;

    posicion +=
      consulta.length;

  }


  return contador;

}


/************************************************************
 * 13. CREAR FRAGMENTO DEL DOCUMENTO
 ************************************************************/

function crearFragmento(
  texto,
  consulta
) {

  if (!texto) {

    return '';

  }


  const textoLower =
    texto.toLowerCase();


  const consultaLower =
    consulta.toLowerCase();


  const posicion =
    textoLower.indexOf(
      consultaLower
    );


  if (posicion === -1) {

    return texto.substring(
      0,
      200
    );

  }


  const inicio =
    Math.max(
      0,
      posicion - 100
    );


  const fin =
    Math.min(
      texto.length,
      posicion +
      consulta.length +
      150
    );


  let fragmento =
    texto.substring(
      inicio,
      fin
    );


  if (inicio > 0) {

    fragmento =
      '... ' +
      fragmento;

  }


  if (
    fin <
    texto.length
  ) {

    fragmento +=
      ' ...';

  }


  return fragmento;

}


/************************************************************
 * 14. REGISTRAR ERROR
 ************************************************************/

function registrarError(
  file,
  error
) {

  Logger.log(
    'ERROR: ' +
    file.getName()
  );


  Logger.log(
    error.message
  );


  try {

    const spreadsheetId =
      obtenerSpreadsheetId();


    const spreadsheet =
      SpreadsheetApp.openById(
        spreadsheetId
      );


    const sheet =
      spreadsheet.getSheetByName(
        CONFIG.SHEET_NAME
      );


    if (!sheet) {

      return;

    }


    sheet.appendRow([

      file.getId(),

      file.getName(),

      file.getMimeType(),

      file.getDateCreated(),

      file.getSize(),

      '',

      '',

      file.getUrl(),

      'ERROR: ' +
        error.message,

      new Date()

    ]);

  } catch (e) {

    Logger.log(
      'No se pudo registrar el error.'
    );

  }

}


/************************************************************
 * 15. VERIFICAR CONFIGURACIÓN
 ************************************************************/

function verificarConfiguracion() {

  if (!CONFIG.FOLDER_ID) {

    throw new Error(

      'NO SE HA CONFIGURADO EL ID DE LA CARPETA DRIVE.'

    );

  }


  const apiKey =
    PropertiesService
      .getScriptProperties()
      .getProperty(
        'GEMINI_API_KEY'
      );


  if (!apiKey) {

    throw new Error(

      'No existe GEMINI_API_KEY en las propiedades del proyecto.'

    );

  }

}


/************************************************************
 * 16. PROBAR GEMINI
 ************************************************************/

function probarGemini() {

  const apiKey =
    PropertiesService
      .getScriptProperties()
      .getProperty(
        'GEMINI_API_KEY'
      );


  if (!apiKey) {

    throw new Error(
      'No existe GEMINI_API_KEY.'
    );

  }


  const url =
    'https://generativelanguage.googleapis.com/v1beta/models/' +
    CONFIG.GEMINI_MODEL +
    ':generateContent';


  const payload = {

    contents: [

      {

        parts: [

          {

            text:
              'Responde únicamente: GEMINI FUNCIONA CORRECTAMENTE'

          }

        ]

      }

    ]

  };


  const options = {

    method: 'POST',

    contentType:
      'application/json',

    headers: {

      'x-goog-api-key':
        apiKey

    },

    payload:
      JSON.stringify(payload),

    muteHttpExceptions: true

  };


  const response =
    UrlFetchApp.fetch(
      url,
      options
    );


  Logger.log(
    response.getContentText()
  );

}


/************************************************************
 * 17. PROCESAR TODOS LOS ARCHIVOS EXISTENTES
 *
 * Esta función NO la utiliza la interfaz.
 *
 * Sirve para procesar documentos que ya estaban
 * en Drive antes de crear la interfaz.
 ************************************************************/

function procesarDocumentos() {

  verificarConfiguracion();


  const folder =
    DriveApp.getFolderById(
      CONFIG.FOLDER_ID
    );


  const files =
    folder.getFiles();


  let contador = 0;


  while (
    files.hasNext()
  ) {

    const file =
      files.next();


    const mimeType =
      file.getMimeType();


    if (

      CONFIG.ALLOWED_MIME_TYPES
        .indexOf(mimeType) === -1

    ) {

      continue;

    }


    if (
      documentoYaProcesado(
        file.getId()
      )
    ) {

      continue;

    }


    try {

      procesarArchivo(
        file
      );


      contador++;

    } catch (error) {

      registrarError(
        file,
        error
      );

    }

  }


  Logger.log(
    'Documentos procesados: ' +
    contador
  );

}


/************************************************************
 * 18. COMPROBAR SI YA ESTÁ INDEXADO
 ************************************************************/

function documentoYaProcesado(
  fileId
) {

  const spreadsheetId =
    obtenerSpreadsheetId();


  const spreadsheet =
    SpreadsheetApp.openById(
      spreadsheetId
    );


  const sheet =
    spreadsheet.getSheetByName(
      CONFIG.SHEET_NAME
    );


  if (!sheet) {

    return false;

  }


  const lastRow =
    sheet.getLastRow();


  if (lastRow <= 1) {

    return false;

  }


  const ids =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        1
      )
      .getValues();


  for (
    let i = 0;
    i < ids.length;
    i++
  ) {

    if (
      String(ids[i][0]) ===
      String(fileId)
    ) {

      return true;

    }

  }


  return false;

}


/************************************************************
 * 19. CREAR TRIGGER AUTOMÁTICO
 ************************************************************/

function crearTriggerAutomatico() {

  const triggers =
    ScriptApp.getProjectTriggers();


  triggers.forEach(
    function(trigger) {

      if (

        trigger
          .getHandlerFunction() ===
        'procesarDocumentos'

      ) {

        ScriptApp.deleteTrigger(
          trigger
        );

      }

    }
  );


  ScriptApp
    .newTrigger(
      'procesarDocumentos'
    )

    .timeBased()

    .everyMinutes(5)

    .create();


  Logger.log(
    'Trigger creado correctamente.'
  );

}
