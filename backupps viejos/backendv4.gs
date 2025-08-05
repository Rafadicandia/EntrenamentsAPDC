/**
 * Constantes globales para los nombres de las hojas y columnas.
 * Modifica estos valores si cambias los nombres en tu Google Sheet.
 */
const SESSIONS_SHEET_NAME = "Sesiones";
const REGISTRATIONS_SHEET_NAME = "Registros";
const ACTIVE_STATUS = "Activo";
const CANCELLED_STATUS = "Cancelado";

// Dirección de correo electrónico a la que se enviarán los listados diarios
const DAILY_LIST_RECIPIENT = "rdicandia@gmail.com"; // ¡IMPORTANTE: Cambia esto por el correo real!

// --- Funciones Principales (Web App Endpoints) ---

/**
 * Función principal que maneja las solicitudes GET.
 * - Sin parámetros: Devuelve la lista de sesiones disponibles.
 * - Con parámetros 'action=cancel', 'registrationId', 'token': Procesa una solicitud de cancelación.
 * @param {object} e - El objeto de evento de la solicitud GET.
 * @returns {object} - ContentService output (JSON para sesiones, HTML para cancelación).
 */
function doGet(e) {
  Logger.log("INFO: La función doGet se ha iniciado.");
  try {
    // Verificar si es una solicitud de cancelación
    if (e.parameter.action === "cancel" && e.parameter.registrationId && e.parameter.token) {
      Logger.log("INFO: Solicitud de cancelación detectada.");
      return handleCancellationRequest(e.parameter.registrationId, e.parameter.token);
    }
    // Si no, es una solicitud para obtener las sesiones
    else {
      Logger.log("INFO: Solicitud para obtener sesiones detectada.");
      const sessions = getAvailableSessions();
      // Devolver las sesiones como JSON
      return ContentService
        .createTextOutput(JSON.stringify({ sessions: sessions }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  } catch (error) {
    Logger.log("ERROR: Error en doGet: " + error.message);
    return ContentService
      .createTextOutput(JSON.stringify({ error: "Error interno del servidor: " + error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Función principal que maneja las solicitudes POST (registros).
 * @param {object} e - El objeto de evento de la solicitud POST.
 * @returns {object} - ContentService output (JSON con resultado).
 */
function doPost(e) {
  Logger.log("INFO: La función doPost se ha iniciado.");
  try {
    // Parsear los datos JSON enviados desde el frontend
    const requestData = JSON.parse(e.postData.contents);

    Logger.log("Request Data (POST): " + JSON.stringify(requestData));

    // Validar datos básicos (se pueden añadir más validaciones)
    if (!requestData.sessionId || !requestData.name || !requestData.phone || !requestData.email || requestData.isMember === undefined) { // isMember puede ser 'true' o 'false'
      throw new Error("Faltan datos requeridos en la solicitud.");
    }

    const result = registerUserForSession(requestData);

    Logger.log("Registration Result (from registerUserForSession): " + JSON.stringify(result));

    // Devolver resultado como JSON
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    Logger.log("ERROR: Error capturado en doPost: " + error);
    // Devolver error como JSON
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, message: "Error al procesar el registro: " + error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// --- Lógica de Negocio ---

/**
 * Obtiene la lista de sesiones disponibles con el número de plazas restantes.
 * @returns {Array<object>} - Array de objetos de sesión con 'remainingSpots'.
 */
function getAvailableSessions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sessionsSheet = ss.getSheetByName(SESSIONS_SHEET_NAME);
  const registrationsSheet = ss.getSheetByName(REGISTRATIONS_SHEET_NAME);

  if (!sessionsSheet || !registrationsSheet) {
    throw new Error(`Una o ambas hojas ('${SESSIONS_SHEET_NAME}', '${REGISTRATIONS_SHEET_NAME}') no se encontraron.`);
  }

  const sessionsData = sessionsSheet.getDataRange().getValues();
  const registrationsData = registrationsSheet.getDataRange().getValues();

  Logger.log("Registrations Data Length (in getAvailableSessions): " + registrationsData.length);

  // Obtener cabeceras para mapear columnas por nombre
  const sessionsHeader = sessionsData[0];
  const registrationsHeader = registrationsData[0];

  // Encontrar índices de columnas relevantes
  const sessionCol = {
    id: sessionsHeader.indexOf("SessionID"),
    name: sessionsHeader.indexOf("Name"),
    instructor: sessionsHeader.indexOf("Instructor"),
    location: sessionsHeader.indexOf("Location"),
    dateTime: sessionsHeader.indexOf("DateTime"),
    capacity: sessionsHeader.indexOf("Capacity")
  };

  const regCol = {
    registrationId: registrationsHeader.indexOf("RegistrationID"),
    timestamp: registrationsHeader.indexOf("Timestamp"),
    sessionId: registrationsHeader.indexOf("SessionID"),
    userName: registrationsHeader.indexOf("UserName"), // CAMBIADO: Asegúrate que coincide con el nombre de la columna en tu hoja
    phone: registrationsHeader.indexOf("UserPhone"),   // CAMBIADO: Asegúrate que coincide con el nombre de la columna en tu hoja
    email: registrationsHeader.indexOf("UserEmail"),   // CAMBIADO: Asegúrate que coincide con el nombre de la columna en tu hoja
    isMember: registrationsHeader.indexOf("IsMember"),
    status: registrationsHeader.indexOf("Status"),
    cancellationToken: registrationsHeader.indexOf("CancellationToken")
  };

  // Validar que todas las columnas necesarias existen
  if (Object.values(sessionCol).some(index => index === -1)) {
     const missingSessionCols = Object.keys(sessionCol).filter(key => sessionCol[key] === -1);
     throw new Error(`Faltan columnas requeridas en la hoja '${SESSIONS_SHEET_NAME}': ${missingSessionCols.join(', ')}. Revisa las cabeceras.`);
  }
   if (Object.values(regCol).some(index => index === -1)) {
     const missingRegCols = Object.keys(regCol).filter(key => regCol[key] === -1);
     throw new Error(`Faltan columnas requeridas en la hoja '${REGISTRATIONS_SHEET_NAME}': ${missingRegCols.join(', ')}. Revisa las cabeceras.`);
  }


  // Contar registros activos por SessionID
  const activeRegistrationsCount = {};
  // Empezar desde 1 para saltar la cabecera
  for (let i = 1; i < registrationsData.length; i++) {
    const row = registrationsData[i];
    let sessionIdFromRegistrations = row[regCol.sessionId]; // Valor original de la hoja

    // CORRECCIÓN CLAVE AQUÍ: Formatear la SessionID de Registros para la comparación
    if (sessionIdFromRegistrations instanceof Date) {
      sessionIdFromRegistrations = Utilities.formatDate(sessionIdFromRegistrations, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "dd/MM/yyyy");
    }
    // Si ya es un string, asumimos que está en el formato correcto o lo manejamos si es necesario.
    // El log indica que es un objeto Date.

    const status = row[regCol.status];
    if (status === ACTIVE_STATUS) {
      activeRegistrationsCount[sessionIdFromRegistrations] = (activeRegistrationsCount[sessionIdFromRegistrations] || 0) + 1;
    }
  }
  Logger.log("Active Registrations Count (in getAvailableSessions): " + JSON.stringify(activeRegistrationsCount));

  // Construir la lista de sesiones con plazas restantes
  const availableSessions = [];
  // Empezar desde 1 para saltar la cabecera
  for (let i = 1; i < sessionsData.length; i++) {
    const row = sessionsData[i];
    // Usamos el formato de fecha para el ID de sesión como en el frontend
    const sessionIdFormatted = Utilities.formatDate(new Date(row[sessionCol.id]), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "dd/MM/yyyy");
    Logger.log(`Session ID from Sessions Sheet (raw): '${row[sessionCol.id]}', Formatted for comparison: '${sessionIdFormatted}'`);

    const capacity = parseInt(row[sessionCol.capacity], 10) || 0; // Asegurar que sea número
    const registered = activeRegistrationsCount[sessionIdFormatted] || 0; // Usar el ID formateado para contar
    const remainingSpots = capacity - registered;

    availableSessions.push({
      id: sessionIdFormatted, // ID formateado
      name: row[sessionCol.name],
      instructor: row[sessionCol.instructor],
      location: row[sessionCol.location],
      // Formatear solo la hora y minutos (HH:mm)
      dateTime: Utilities.formatDate(new Date(row[sessionCol.dateTime]), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "HH:mm"),
      capacity: capacity,
      registered: registered,
      remainingSpots: remainingSpots // Añadido para conveniencia del frontend
    });
  }

  return availableSessions;
}

/**
 * Registra un usuario para una sesión específica.
 * @param {object} data - Datos del registro (sessionId, name, phone, email, isMember).
 * @returns {object} - Objeto con { success: boolean, message: string }.
 */
function registerUserForSession(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sessionsSheet = ss.getSheetByName(SESSIONS_SHEET_NAME);
  const registrationsSheet = ss.getSheetByName(REGISTRATIONS_SHEET_NAME);

  if (!sessionsSheet || !registrationsSheet) {
    throw new Error(`Hojas '${SESSIONS_SHEET_NAME}' o '${REGISTRATIONS_SHEET_NAME}' no encontradas.`);
  }

  // --- 1. Verificar Plazas Disponibles ---
  const sessions = getAvailableSessions(); // Reutilizamos la función para obtener datos actualizados
  const targetSession = sessions.find(s => s.id === data.sessionId);

  if (targetSession) {
      Logger.log(`Remaining spots for session ${targetSession.id} BEFORE appending new row: ${targetSession.remainingSpots}`);
  }
  Logger.log(`Session ID received from frontend to register: '${data.sessionId}'`);


  if (!targetSession) {
    return { success: false, message: "La sesión seleccionada no existe." };
  }

  if (targetSession.remainingSpots <= 0) {
    return { success: false, message: "Lo sentimos, no quedan plazas libres para esta sesión." };
  }

  // --- 2. Añadir Registro ---
  const registrationId = Utilities.getUuid(); // Generar ID único para el registro
  const cancellationToken = Utilities.getUuid(); // Generar token único para cancelación
  const timestamp = new Date();
  const isMemberText = data.isMember === 'true' ? 'Sí' : 'No'; // Convertir string 'true'/'false' a texto 'Sí'/'No'

  // Obtener cabeceras para encontrar el orden correcto de las columnas
  const registrationsHeader = registrationsSheet.getDataRange().getValues()[0];
   const regCol = {
    registrationId: registrationsHeader.indexOf("RegistrationID"),
    timestamp: registrationsHeader.indexOf("Timestamp"),
    sessionId: registrationsHeader.indexOf("SessionID"),
    userName: registrationsHeader.indexOf("UserName"), // CAMBIADO: Asegúrate que coincide con el nombre de la columna en tu hoja
    phone: registrationsHeader.indexOf("UserPhone"),   // CAMBIADO: Asegúrate que coincide con el nombre de la columna en tu hoja
    email: registrationsHeader.indexOf("UserEmail"),   // CAMBIADO: Asegúrate que coincide con el nombre de la columna en tu hoja
    isMember: registrationsHeader.indexOf("IsMember"),
    status: registrationsHeader.indexOf("Status"),
    cancellationToken: registrationsHeader.indexOf("CancellationToken"),
    pago: registrationsHeader.indexOf("Pago"),
    asistencia: registrationsHeader.indexOf("Asistencia")
  };

   // Validar que todas las columnas necesarias existen antes de appendRow
   if (Object.values(regCol).some(index => index === -1)) {
     const missingRegCols = Object.keys(regCol).filter(key => regCol[key] === -1);
     throw new Error(`Faltan columnas requeridas en la hoja '${REGISTRATIONS_SHEET_NAME}' para el registro: ${missingRegCols.join(', ')}. Revisa las cabeceras.`);
  }


  // Crear un array para la nueva fila en el orden correcto de las columnas
  const newRow = new Array(registrationsHeader.length).fill(''); // Inicializar con celdas vacías
  newRow[regCol.registrationId] = registrationId;
  newRow[regCol.timestamp] = timestamp;
  newRow[regCol.sessionId] = data.sessionId; // Usar el ID de sesión tal cual viene del frontend (dd/MM/yyyy)
  newRow[regCol.userName] = data.name;
  newRow[regCol.phone] = data.phone;
  newRow[regCol.email] = data.email;
  newRow[regCol.isMember] = isMemberText;
  newRow[regCol.status] = ACTIVE_STATUS; // Estado inicial
  newRow[regCol.cancellationToken] = cancellationToken;
  // Las columnas "Pago" y "Asistencia" se dejan vacías o con valor por defecto si existen


  // Añadir fila a la hoja de Registros
  registrationsSheet.appendRow(newRow);
  SpreadsheetApp.flush(); // Asegurar que los cambios se guarden inmediatamente
  Logger.log(`INFO: Nuevo registro añadido a la hoja '${REGISTRATIONS_SHEET_NAME}' para SessionID: '${data.sessionId}'`);

  // --- 3. Enviar Email de Confirmación (Opcional pero recomendado) ---
  try {
    sendConfirmationEmail(data.email, data.name, targetSession, registrationId, cancellationToken);
  } catch (emailError) {
    Logger.log("Error en enviar el correu electrònic de confirmació per a " + data.email + ": " + emailError.message);
  }

  return { success: true, message: `¡Registre completat per a ${targetSession.name}! Revisa el teu correu electrònic per a la confirmació. Si no el trobes, mira a la bústia de correu brossa.` };
}

/**
 * Maneja una solicitud de cancelación verificando el token.
 * @param {string} registrationId - ID del registro a cancelar.
 * @param {string} token - Token de cancelación proporcionado.
 * @returns {object} - ContentService output (HTML con mensaje de confirmación/error).
 */
function handleCancellationRequest(registrationId, token) {
  Logger.log(`INFO: Iniciando solicitud de cancelación para RegistrationID: ${registrationId}`);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const registrationsSheet = ss.getSheetByName(REGISTRATIONS_SHEET_NAME);

  if (!registrationsSheet) {
    Logger.log("ERROR: handleCancellationRequest - Hoja de registros no encontrada.");
    return ContentService.createTextOutput("Error interno: Hoja de registros no encontrada.").setMimeType(ContentService.MimeType.HTML);
  }

  const data = registrationsSheet.getDataRange().getValues();
  const header = data[0];
  const idCol = header.indexOf("RegistrationID");
  const tokenCol = header.indexOf("CancellationToken");
  const statusCol = header.indexOf("Status");
  const sessionCol = header.indexOf("SessionID");
  const nameCol = header.indexOf("UserName"); // CAMBIADO: Asegúrate que coincide con el nombre de la columna en tu hoja

  // Validar columnas
  if ([idCol, tokenCol, statusCol, sessionCol, nameCol].some(index => index === -1)) {
    Logger.log(`ERROR: Error de cancelación: Faltan columnas requeridas en la hoja '${REGISTRATIONS_SHEET_NAME}'.`);
    return ContentService.createTextOutput("Error interno del servidor al procesar la cancelación.").setMimeType(ContentService.MimeType.HTML);
  }


  let message = "Error: No se pudo encontrar tu registro o el enlace de cancelación es inválido.";
  let found = false;
  let alreadyCancelled = false;

  // Buscar el registro por ID y Token (empezar desde 1 para saltar cabecera)
  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === registrationId && data[i][tokenCol] === token) {
      found = true;
      if (data[i][statusCol] === ACTIVE_STATUS) {
        registrationsSheet.getRange(i + 1, statusCol + 1).setValue(CANCELLED_STATUS);
        message = `El teu registre per a la sessió ha estat cancel·lat correctament. Esperem veure't aviat!`;
        Logger.log(`SUCCESS: Cancelación correcta: ID ${registrationId}, Usuario ${data[i][nameCol]}, Sesión ${data[i][sessionCol]}`);
        SpreadsheetApp.flush(); // Asegurar que los cambios se guarden
      } else if (data[i][statusCol] === CANCELLED_STATUS) {
        alreadyCancelled = true;
        message = "Aquest registre ja havia estat cancel·lat prèviament.";
        Logger.log(`INFO: Intento de cancelación repetido: ID ${registrationId}`);
      }
      break;
    }
  }

  if (!found) {
    Logger.log(`WARN: Intento de cancelación fallido: ID ${registrationId} o Token inválido.`);
  }

  // 1. Carga el archivo HTML como una plantilla (asumiendo que tienes un archivo CancelPage.html)
  try {
      const template = HtmlService.createTemplateFromFile('CancelPage');
      // 2. Pasa variables del script (.gs) a la plantilla (.html)
      template.message = message;
      template.scriptUrl = ScriptApp.getService().getUrl();
      // 3. Evalúa la plantilla para obtener el resultado HTML final
      const htmlOutput = template.evaluate();
      // 4. Devuelve el objeto HtmlOutput
      return htmlOutput;
  } catch (e) {
      Logger.log("ERROR: Error al cargar o procesar la plantilla CancelPage.html: " + e.message);
      return HtmlService.createHtmlOutput(`
          <!DOCTYPE html>
          <html>
          <head>
              <title>Estado de la Cancelación</title>
              <meta name="viewport" content="width=device-width, initial-scale=1.0">
              <style> body { font-family: sans-serif; text-align: center; margin-top: 50px; } </style>
          </head>
          <body>
              <h1>Estado de la Cancelación</h1>
              <p>${message}</p>
              <p><a href="${ScriptApp.getService().getUrl()}">Volver al formulario de registro</a></p>
          </body>
          </html>
      `).setTitle('Estado de la Cancelación');
  }
}


// --- Funciones Auxiliares ---

/**
 * Envía un email de confirmación al usuario con el enlace de cancelación.
 * @param {string} recipientEmail - Email del destinatario.
 * @param {string} recipientName - Nombre del destinatario.
 * @param {object} session - Objeto de la sesión registrada.
 * @param {string} registrationId - ID único del registro.
 * @param {string} cancellationToken - Token único de cancelación.
 */
function sendConfirmationEmail(recipientEmail, recipientName, session, registrationId, cancellationToken) {
  const subject = `Confirmació de Registre: ${session.name}`;
  const scriptUrl = ScriptApp.getService().getUrl();
  const cancellationLink = `${scriptUrl}?action=cancel&registrationId=${registrationId}&token=${cancellationToken}`;

  const body = `
    Hola ${recipientName},<br><br>
    Has estat registrat/ada correctament per a la següent sessió:<br><br>
    <b>Sessió:</b>Entrenaments professionals de contemporani<br>
    <b>Instructora:</b> ${session.instructor}<br>
    <b>Lloc:</b> ${session.location}<br>
    <b>Data:</b> ${session.id}<br>
    <b>Horari:</b> ${session.dateTime}<br><br>
    <a href="https://www.dansacat.org/lassociacio/serveis-associades/entrenaments-professionals-contemporani/">Aquí</a> podeu trobar tota la informació referent als entrenadors/es, horaris,
    adreces, tarifes i mètodes de pagament.<br><br>
    Si has de cancel·lar la teva plaça, si us plau, fes clic al següent enllaç:<br>
    <a href="${cancellationLink}">Cancel·lar el meu registre</a><br><br>
    Si no pots fer clic a l'enllaç, copia i enganxa la següent URL al teu navegador:<br>
    ${cancellationLink}<br><br>
    Ens veiem als entrenaments!<br><br>
    <em>(Aquest és un missatge automàtic, si us plau, no responguis directament.)</em>
  `;

  try {
    MailApp.sendEmail({
      to: recipientEmail,
      subject: subject,
      htmlBody: body,
    });
    Logger.log(`INFO: Email de confirmación enviado a ${recipientEmail} para el registro ${registrationId}`);
  } catch (emailError) {
    Logger.log(`ERROR: Error al enviar el email de confirmación a ${recipientEmail} para el registro ${registrationId}: ${emailError.message}`);
  }
}

// --- Nuevas Funciones para Tareas Programadas ---

/**
 * Obtiene todos los registros activos de la hoja.
 * @returns {Array<object>} - Array de objetos de registro.
 */
function getAllActiveRegistrations() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const registrationsSheet = ss.getSheetByName(REGISTRATIONS_SHEET_NAME);

    if (!registrationsSheet) {
        Logger.log(`ERROR: getAllActiveRegistrations - Hoja '${REGISTRATIONS_SHEET_NAME}' no encontrada.`);
        return [];
    }

    const data = registrationsSheet.getDataRange().getValues();
    if (data.length <= 1) {
        Logger.log(`INFO: La hoja '${REGISTRATIONS_SHEET_NAME}' no tiene registros activos (solo cabecera).`);
        return [];
    }

    const header = data[0];
    const regCol = {
        registrationId: header.indexOf("RegistrationID"),
        timestamp: header.indexOf("Timestamp"),
        sessionId: header.indexOf("SessionID"),
        userName: header.indexOf("UserName"), // CAMBIADO: Asegúrate que coincide con el nombre de la columna en tu hoja
        phone: header.indexOf("UserPhone"),   // CAMBIADO: Asegúrate que coincide con el nombre de la columna en tu hoja
        email: header.indexOf("UserEmail"),   // CAMBIADO: Asegúrate que coincide con el nombre de la columna en tu hoja
        isMember: header.indexOf("IsMember"),
        status: header.indexOf("Status"),
        cancellationToken: header.indexOf("CancellationToken"),
        pago: header.indexOf("Pago"),
        asistencia: header.indexOf("Asistencia")
    };

    if (Object.values(regCol).some(index => index === -1)) {
         const missingRegCols = Object.keys(regCol).filter(key => regCol[key] === -1);
         Logger.log(`ERROR: getAllActiveRegistrations - Faltan columnas requeridas en la hoja '${REGISTRATIONS_SHEET_NAME}': ${missingRegCols.join(', ')}. Revisa las cabeceras.`);
         return [];
    }

    const activeRegistrations = [];
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (row[regCol.status] === ACTIVE_STATUS) {
            let sessionIdFromSheet = row[regCol.sessionId];

            // CORRECCIÓN CLAVE AQUÍ: Formatear la SessionID de Registros
            // Si es un objeto Date (lo más probable), formatéalo.
            if (sessionIdFromSheet instanceof Date) {
                sessionIdFromSheet = Utilities.formatDate(sessionIdFromSheet, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "dd/MM/yyyy");
            } else {
                // Si ya es un string, y no es el formato deseado (por ejemplo, 'Sat May 10...'),
                // intenta parsearlo y luego formatéalo.
                try {
                    const parsedDate = new Date(sessionIdFromSheet);
                    if (!isNaN(parsedDate.getTime())) { // Comprueba si el parseo fue exitoso (no es Invalid Date)
                        sessionIdFromSheet = Utilities.formatDate(parsedDate, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "dd/MM/yyyy");
                    }
                } catch (e) {
                    Logger.log(`WARN: No se pudo parsear el SessionID '${sessionIdFromSheet}' a una fecha en getAllActiveRegistrations. Error: ${e.message}`);
                }
            }

            activeRegistrations.push({
                registrationId: row[regCol.registrationId],
                timestamp: row[regCol.timestamp],
                sessionId: sessionIdFromSheet, // Ahora está correctamente formateado
                userName: row[regCol.userName],
                phone: row[regCol.phone],
                email: row[regCol.email],
                isMember: row[regCol.isMember],
                status: row[regCol.status],
                cancellationToken: row[regCol.cancellationToken],
                pago: regCol.pago !== -1 ? row[regCol.pago] : '',
                asistencia: regCol.asistencia !== -1 ? row[regCol.asistencia] : ''
            });
        }
    }
    Logger.log(`INFO: Se encontraron ${activeRegistrations.length} registros activos en total.`);
    return activeRegistrations;
}


/**
 * Envía un email con el listado de registros activos para un día específico.
 * Esta función será llamada por el activador diario.
 */
function sendDailyLists() {
    const today = new Date();
    const formattedDate = Utilities.formatDate(today, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "dd/MM/yyyy");
    Logger.log(`INFO: Generando listado para el día: ${formattedDate}`);

    const allRegistrations = getAllActiveRegistrations();
    Logger.log(`INFO: Total active registrations found by getAllActiveRegistrations: ${allRegistrations.length}`);

    const dailyRegistrations = allRegistrations.filter(reg => {
        Logger.log(`DEBUG: Comparando registration sessionId '${reg.sessionId}' con formattedDate '${formattedDate}'`);
        return reg.sessionId === formattedDate;
    });

    Logger.log(`INFO: Daily registrations found for ${formattedDate}: ${dailyRegistrations.length}`);

    if (dailyRegistrations.length === 0) {
        Logger.log(`INFO: No hay registros activos para el día ${formattedDate}. No se envía email.`);
        return;
    }

    // Ordenar por hora de registro (timestamp)
    dailyRegistrations.sort((a, b) => a.timestamp.getTime() - b.timestamp.getTime());

    // Prepare data for Excel
    const excelHeaders = ["Nombre", "Teléfono", "Email", "Socio/a", "Pago", "Asistencia"];
    const excelData = dailyRegistrations.map(reg => {
        const registrationTime = Utilities.formatDate(reg.timestamp, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "HH:mm:ss");
        return [
            reg.userName,
            reg.phone,
            reg.email,
            reg.isMember,
            reg.pago || '', // Handle empty values
            reg.asistencia || '' // Handle empty values
        ];
    });

    const fileName = `Llistat_Registres_${formattedDate}.xlsx`;
    let attachmentBlob = null;
    let tempSpreadsheet = null; // Declarar fuera del try para que sea accesible en finally

    try {
        // Crear una hoja de cálculo temporal en Google Drive
        tempSpreadsheet = SpreadsheetApp.create("Temp_" + fileName.replace(".xlsx", ""));
        const tempSheet = tempSpreadsheet.getSheets()[0];

        // Escribir cabeceras
        tempSheet.getRange(1, 1, 1, excelHeaders.length).setValues([excelHeaders]);
        // Escribir datos
        tempSheet.getRange(2, 1, excelData.length, excelData[0].length).setValues(excelData);

        // Asegurarse de que todos los cambios se guarden antes de exportar
        SpreadsheetApp.flush();

        // Exportar la hoja de cálculo temporal como XLSX
        const url = `https://docs.google.com/spreadsheets/d/${tempSpreadsheet.getId()}/export?format=xlsx`;
        const options = {
            headers: {
                Authorization: `Bearer ${ScriptApp.getOAuthToken()}`
            },
            muteHttpExceptions: true // Evita que un error de HTTP lance una excepción, permitiendo manejar el código de respuesta
        };
        const response = UrlFetchApp.fetch(url, options);

        if (response.getResponseCode() === 200) {
            attachmentBlob = response.getBlob().setName(fileName);
            Logger.log(`INFO: Archivo Excel '${fileName}' generado con éxito.`);
        } else {
            Logger.log(`ERROR: Fallo al generar archivo Excel. Código de respuesta: ${response.getResponseCode()}, Mensaje: ${response.getContentText()}`);
            // Fallback a enviar email HTML si la generación de Excel falla
            sendHtmlEmail(formattedDate, dailyRegistrations);
            return; // Detener la ejecución aquí si Excel falló y se envió el fallback
        }

        // Enviar email con adjunto
       const subject = `Llistat de Registres Actius - ${formattedDate}`;
        const emailBodyHtml = `
            <html>
            <body>
                <h2>Llistat de Registres Actius per al ${formattedDate}</h2>
                <p>S'adjunta el llistat de registres actius.</p>
                <p>Total de registres actius per al ${formattedDate}: ${dailyRegistrations.length}</p>
                <br>
                <p><em>Aquest és un correu electrònic automàtic.</em></p>
            </body>
            </html>
        `;

        MailApp.sendEmail({
            to: DAILY_LIST_RECIPIENT,
            subject: subject,
            htmlBody: emailBodyHtml,
            attachments: [attachmentBlob]
        });
        Logger.log(`INFO: Email con archivo Excel enviado a ${DAILY_LIST_RECIPIENT} para el ${formattedDate}.`);

    } catch (e) {
        Logger.log(`ERROR: Error general al generar o enviar el archivo Excel: ${e.message}`);
        // Fallback a enviar email HTML si algo sale mal con la generación/envío de Excel
        sendHtmlEmail(formattedDate, dailyRegistrations);
    } finally {
        // Limpiar: Borrar la hoja de cálculo temporal
        if (tempSpreadsheet) {
            try {
                // Mover a la papelera (no elimina permanentemente de inmediato)
                DriveApp.getFileById(tempSpreadsheet.getId()).setTrashed(true);
                Logger.log(`INFO: Hoja de cálculo temporal '${tempSpreadsheet.getName()}' eliminada.`);
            } catch (deleteError) {
                Logger.log(`WARN: No se pudo eliminar la hoja de cálculo temporal '${tempSpreadsheet.getName()}': ${deleteError.message}. Puede que ya no exista o haya un problema de permisos.`);
            }
        }
    }
}

// Función auxiliar para enviar email HTML (como fallback)
function sendHtmlEmail(formattedDate, dailyRegistrations) {
    let emailBody = `
        <html>
        <body>
            <h2>Listado de Registros Activos para el ${formattedDate}</h2>
            <table border="1" cellpadding="5" cellspacing="0">
                <thead>
                    <tr>
                        <th>Hora Registro</th>
                        <th>Nombre</th>
                        <th>Teléfono</th>
                        <th>Email</th>
                        <th>Socio/a</th>
                        <th>Pago</th>
                        <th>Asistencia</th>
                    </tr>
                </thead>
                <tbody>
    `;

    dailyRegistrations.forEach(reg => {
        const registrationTime = Utilities.formatDate(reg.timestamp, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "HH:mm:ss");
        emailBody += `
            <tr>
                <td>${registrationTime}</td>
                <td>${reg.userName}</td>
                <td>${reg.phone}</td>
                <td>${reg.email}</td>
                <td>${reg.isMember}</td>
                <td>${reg.pago || ''}</td>
                <td>${reg.asistencia || ''}</td>
            </tr>
        `;
    });

    emailBody += `
                </tbody>
            </table>
            <br>
            <p>Total de registros activos para el ${formattedDate}: ${dailyRegistrations.length}</p>
            <br>
            <p><em>Este es un email automático.</em></p>
        </body>
        </html>
    `;

    const subject = `Listado de Registros Activos - ${formattedDate} (HTML - Fallback)`;
    try {
        MailApp.sendEmail({
            to: DAILY_LIST_RECIPIENT,
            subject: subject,
            htmlBody: emailBody,
        });
        Logger.log(`INFO: Email de listado diario HTML (fallback) enviado a ${DAILY_LIST_RECIPIENT} para el ${formattedDate}.`);
    } catch (emailError) {
        Logger.log(`ERROR: Error al enviar el email de listado diario HTML (fallback) a ${DAILY_LIST_RECIPIENT}: ${emailError.message}`);
    }
}


/**
 * Borra todos los registros de la hoja "Registros" excepto la cabecera.
 * Esta función será llamada por el activador semanal.
 */
/**
 * Borra todos los registros de la hoja "Registros" excepto la cabecera.
 * Esta función será llamada por el activador semanal.
 */
function resetWeeklyRegistrations() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const registrationsSheet = ss.getSheetByName(REGISTRATIONS_SHEET_NAME);

    if (!registrationsSheet) {
        Logger.log(`ERROR: resetWeeklyRegistrations - Hoja '${REGISTRATIONS_SHEET_NAME}' no encontrada.`);
        return;
    }

    const lastRow = registrationsSheet.getLastRow();
    const lastColumn = registrationsSheet.getLastColumn(); // Obtenemos la última columna para limpiar todo el rango de datos

    if (lastRow > 1) {
        // Obtenemos el rango de datos desde la fila 2 (después de la cabecera inmovilizada)
        // hasta la última fila y última columna, y borramos solo su contenido.
        const dataRange = registrationsSheet.getRange(2, 1, lastRow - 1, lastColumn);
        dataRange.clearContent(); // Borrar solo el contenido, no las filas.
        Logger.log(`INFO: Se ha borrado el contenido de ${lastRow - 1} registros de la hoja '${REGISTRATIONS_SHEET_NAME}'.`);
        SpreadsheetApp.flush();
    } else {
        Logger.log(`INFO: La hoja '${REGISTRATIONS_SHEET_NAME}' ya está vacía (solo cabecera).`);
    }
}

// --- Función de prueba para activar manualmente ---
function testDailyList() {
  Logger.log("INFO: Ejecutando testDailyList manualmente.");
  sendDailyLists();
}

function testWeeklyReset() {
  Logger.log("INFO: Ejecutando testWeeklyReset manualmente.");
  resetWeeklyRegistrations();
}