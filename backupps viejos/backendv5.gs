/**
 * Constantes globales para los nombres de las hojas y columnas.
 * Modifica estos valores si cambias los nombres en tu Google Sheet.
 */
const SESSIONS_SHEET_NAME = "Sesiones";
const REGISTRATIONS_SHEET_NAME = "Registros";
const ACTIVE_STATUS = "Activo";
const CANCELLED_STATUS = "Cancelado";

// --- Nuevas Constantes ---
const SESSION_LOCATION_COL_HEADER = "Location";
const SESSION_DATE_COL_HEADER = "SessionID"; // **REEMPLAZA** con la cabecera exacta de la columna con la fecha/ID de la sesión en la hoja Sesiones
const LOCATION_RECIPIENT_CONFIG_SHEET_NAME = "Control_Correo";
const CONFIG_LOCATION_COL_HEADER = "Lugar"; // Cabecera para la columna del lugar en la hoja de configuración
const CONFIG_RECIPIENTS_COL_HEADER = "Destinatarios"; // Cabecera para la columna de destinatarios en la hoja de configuración

// Opcional: Un destinatario por defecto si no se encuentra un responsable para la sesión
const DEFAULT_RECIPIENT_IF_LOCATION_NOT_FOUND = "rdicandia@gmail.com"; // **REEMPLAZA** o pon null si no quieres fallback


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
    observaciones: registrationsHeader.indexOf("Observaciones")
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
        observaciones: header.indexOf("Observaciones")
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
                observaciones: regCol.observaciones !== -1 ? row[regCol.observaciones] : ''
            });
        }
    }
    Logger.log(`INFO: Se encontraron ${activeRegistrations.length} registros activos en total.`);
    return activeRegistrations;
}

/**
 * Busca el/los destinatario(s) para un lugar específico en la hoja de configuración por lugar.
 * @param {string} locationKey - El nombre del lugar a buscar.
 * @returns {string|null} - Las direcciones de correo (separadas por comas) o null si no se encuentra.
 */
function getRecipientEmailsByLocation(locationKey) {
    if (!locationKey || locationKey.trim() === "") {
        Logger.log(`WARN: getRecipientEmailsByLocation - Clave de lugar vacía.`);
        return null;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName(LOCATION_RECIPIENT_CONFIG_SHEET_NAME);

    if (!configSheet) {
        Logger.log(`ERROR: getRecipientEmailsByLocation - Hoja de configuración por lugar '${LOCATION_RECIPIENT_CONFIG_SHEET_NAME}' no encontrada.`);
        return null;
    }

    const data = configSheet.getDataRange().getValues();
    if (data.length <= 1) {
        Logger.log(`WARN: Hoja de configuración por lugar '${LOCATION_RECIPIENT_CONFIG_SHEET_NAME}' está vacía o solo tiene cabecera.`);
        return null;
    }

    const header = data[0];
    const locationColIndex = header.indexOf(CONFIG_LOCATION_COL_HEADER);
    const recipientsColIndex = header.indexOf(CONFIG_RECIPIENTS_COL_HEADER);

    if (locationColIndex === -1 || recipientsColIndex === -1) {
        Logger.log(`ERROR: getRecipientEmailsByLocation - Faltan columnas requeridas ('${CONFIG_LOCATION_COL_HEADER}' o '${CONFIG_RECIPIENTS_COL_HEADER}') en la hoja '${LOCATION_RECIPIENT_CONFIG_SHEET_NAME}'. Revisa las cabeceras.`);
        return null;
    }

    // Buscar el lugar en la hoja de configuración
    for (let i = 1; i < data.length; i++) { // Empezar desde 1 para saltar la cabecera
        const row = data[i];
        // Comparar el lugar de la fila con la clave buscada, limpiando espacios
        if (String(row[locationColIndex]).trim().toLowerCase() === locationKey.trim().toLowerCase()) { // Comparar sin distinguir mayúsculas/minúsculas
            // Se encontró el lugar, devolver el/los correo(s)
            const recipients = row[recipientsColIndex];
            return (recipients === null || recipients === undefined) ? null : String(recipients).trim();
        }
    }

    Logger.log(`WARN: No se encontró el lugar '${locationKey}' en la hoja de configuración por lugar '${LOCATION_RECIPIENT_CONFIG_SHEET_NAME}'.`);
    return null; // No se encontró el lugar en la tabla de configuración
}


/**
 * Envía un email con un enlace a un Google Sheet con el listado
 * de registros activos para un día específico, buscando el destinatario
 * basado en el lugar de la sesión.
 * Esta función será llamada por el activador diario.
 */
function sendDailyLists() {
    const today = new Date();
    const formattedDate = Utilities.formatDate(today, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "dd/MM/yyyy");
    Logger.log(`INFO: Generando listado para el día: ${formattedDate}`);

    const allRegistrations = getAllActiveRegistrations();
    Logger.log(`INFO: Total active registrations found by getAllActiveRegistrations: ${allRegistrations.length}`);

    const dailyRegistrations = allRegistrations.filter(reg => {
         const regSessionIdStr = (reg.sessionId instanceof Date)
             ? Utilities.formatDate(reg.sessionId, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "dd/MM/yyyy")
             : String(reg.sessionId);

         return regSessionIdStr === formattedDate;
    });

    Logger.log(`INFO: Daily registrations found for ${formattedDate}: ${dailyRegistrations.length}`);

    if (dailyRegistrations.length === 0) {
        Logger.log(`INFO: No hay registros activos para el día ${formattedDate}. No se crea Sheets ni se envía email.`);
        return;
    }

    // --- OBTENER EL LUGAR DE LA SESIÓN DEL DÍA ---
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sessionsSheet = ss.getSheetByName(SESSIONS_SHEET_NAME);
     let sessionLocation = null;

     if (!sessionsSheet) {
         Logger.log(`ERROR: sendDailyLists - Hoja de sesiones '${SESSIONS_SHEET_NAME}' no encontrada para obtener el lugar.`);
          const fallbackRecipientsIfSheetsError = DEFAULT_RECIPIENT_IF_LOCATION_NOT_FOUND ? String(DEFAULT_RECIPIENT_IF_LOCATION_NOT_FOUND).trim() : null;
           if (fallbackRecipientsIfSheetsError) {
               Logger.log(`INFO: Error al obtener lugar (hoja Sesiones no encontrada). Intentando enviar a destinatario de fallback: ${fallbackRecipientsIfSheetsError}`);
           }
         return;
     }

     const sessionsData = sessionsSheet.getDataRange().getValues();
     const sessionsHeader = sessionsData[0];
     const sessionDateColIndex = sessionsHeader.indexOf(SESSION_DATE_COL_HEADER);
     const sessionLocationColIndex = sessionsHeader.indexOf(SESSION_LOCATION_COL_HEADER);

     if (sessionDateColIndex === -1 || sessionLocationColIndex === -1) {
          Logger.log(`ERROR: sendDailyLists - Faltan columnas requeridas ('${SESSION_DATE_COL_HEADER}' o '${SESSION_LOCATION_COL_HEADER}') en la hoja '${SESSIONS_SHEET_NAME}' para obtener el lugar.`);
         return;
     }

     let sessionRow = null;
     for (let i = 1; i < sessionsData.length; i++) {
         const row = sessionsData[i];
         let rowSessionDateStr = "";
         const rowSessionDate = row[sessionDateColIndex];

          if (rowSessionDate instanceof Date) {
              rowSessionDateStr = Utilities.formatDate(rowSessionDate, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "dd/MM/yyyy");
          } else {
              rowSessionDateStr = String(rowSessionDate).trim();
          }

         if (rowSessionDateStr === formattedDate) {
             sessionLocation = String(row[sessionLocationColIndex]).trim();
             sessionRow = row;
             break;
         }
     }

     if (!sessionLocation) {
          Logger.log(`WARN: No se encontró una sesión en la hoja '${SESSIONS_SHEET_NAME}' con la fecha '${formattedDate}' o no se pudo determinar el lugar.`);
         return;
     }
     Logger.log(`INFO: Lugar de la sesión para ${formattedDate}: '${sessionLocation}'`);

    // --- OBTENER LOS DESTINATARIOS BASADO EN EL LUGAR ---
    let recipients = getRecipientEmailsByLocation(sessionLocation);

    if (!recipients && DEFAULT_RECIPIENT_IF_LOCATION_NOT_FOUND) {
         Logger.log(`INFO: No se encontraron destinatarios configurados para el lugar '${sessionLocation}'. Usando destinatario de fallback: ${DEFAULT_RECIPIENT_IF_LOCATION_NOT_FOUND}`);
         recipients = DEFAULT_RECIPIENT_IF_LOCATION_NOT_FOUND;
    }

    if (!recipients) {
        Logger.log(`WARN: No se encontraron destinatarios configurados para el lugar '${sessionLocation}' ni hay un destinatario de fallback válido. No se envía email.`);
        return;
    }

    const emailRecipients = String(recipients).trim();
     if (emailRecipients === "") {
         Logger.log(`WARN: La configuración para el lugar '${sessionLocation}' (o el fallback) resultó en una lista de destinatarios vacía. No se envía email.`);
         return;
     }


    // --- Preparar datos para la nueva hoja de Google Sheets ---
    dailyRegistrations.sort((a, b) => a.timestamp.getTime() - b.timestamp.getTime());

    const sheetHeaders = ["Hora Registro", "Nombre", "Teléfono", "Email", "Socio/a", "Pago", "Observaciones"];

    const sheetData = dailyRegistrations.map(reg => {
        const registrationTime = Utilities.formatDate(reg.timestamp, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "HH:mm:ss");
        return [
            registrationTime,
            reg.userName,
            reg.phone,
            reg.email,
            reg.isMember,
            reg.pago || '',
            ''
        ];
    });

    const newSheetTitle = `Registres_${formattedDate}_${sessionLocation.replace(/[^a-zA-Z0-9 ]/g, '_').replace(/ /g, '_')}`;
    let newSpreadsheet = null;
    let sharedFileUrl = null;

    try {
        // --- 1. Crear un nuevo documento de Google Sheets ---
        newSpreadsheet = SpreadsheetApp.create(newSheetTitle);
        const newSheet = newSpreadsheet.getSheets()[0];

        // --- 2. Escribir cabeceras y datos ---
        if (sheetHeaders.length > 0) {
            newSheet.getRange(1, 1, 1, sheetHeaders.length).setValues([sheetHeaders]);
        }
        if (sheetData.length > 0 && sheetData[0].length > 0) {
            newSheet.getRange(2, 1, sheetData.length, sheetData[0].length).setValues(sheetData);
        }

        if (sheetHeaders.length > 0) {
             newSheet.autoResizeColumns(1, sheetHeaders.length);
        }


        // --- AGREGAR VALIDACIÓN DE DATOS (DESPLEGABLE) A LA COLUMNA DE PAGO ---
        const pagoColHeader = "Pago";
        const pagoColIndex = sheetHeaders.indexOf(pagoColHeader);

        if (pagoColIndex !== -1 && sheetData.length > 0) {
            const startRow = 2;
            const numRows = sheetData.length;
            const targetColumn = pagoColIndex + 1;

            const opcionesPago = ['Sí', 'No', 'Pendiente', 'N/A']; // **TUS OPCIONES**

            // **** BLOQUE DE VALIDACIÓN SIMPLIFICADO (SIN setAllowInvalidData ni setHelpText) ****
            const builder = SpreadsheetApp.newDataValidation();
            builder.requireValueInList(opcionesPago);
            // Eliminamos: builder.setAllowInvalidData(true);
            // Eliminamos: builder.setHelpText("Selecciona el estado del pago");
            const rule = builder.build();
            // **** FIN DEL BLOQUE SIMPLIFICADO ****


            const pagoRange = newSheet.getRange(startRow, targetColumn, numRows, 1);
            pagoRange.setDataValidation(rule);

            Logger.log(`INFO: Validación de datos (desplegable) aplicada a la columna '${pagoColHeader}'.`);

        } else if (pagoColIndex === -1) {
             Logger.log(`WARN: La columna con cabecera '${pagoColHeader}' no se encontró en los encabezados definidos para el nuevo Sheet. No se aplicará validación de datos.`);
        } else { // sheetData.length === 0
             Logger.log(`INFO: No hay datos de registros activos para el día ${formattedDate}, no se aplica validación de datos.`);
        }

        SpreadsheetApp.flush();

        // --- 3. Obtener el archivo de Google Drive y establecer permisos de compartido ---
        const sharedFile = DriveApp.getFileById(newSpreadsheet.getId());

         const recipientEmailsArray = emailRecipients.split(',').map(email => email.trim()).filter(email => email !== '');

         recipientEmailsArray.forEach(email => {
             try {
                 if (/\S+@\S+\.\S+/.test(email)) {
                     sharedFile.addEditor(email);
                     Logger.log(`INFO: Permiso de edición otorgado a ${email}`);
                 } else {
                     Logger.log(`WARN: Dirección de correo inválida en la lista de destinatarios: '${email}'. Saltando la acción de compartir para esta dirección.`);
                 }

             } catch(e) {
                 Logger.log(`WARN: Error al otorgar permiso de compartido a ${email}: ${e.message}.`);
             }
         });

        // --- 4. Obtener la URL de compartido del nuevo documento ---
        sharedFileUrl = sharedFile.getUrl();
        Logger.log(`INFO: Documento de Sheets '${newSheetTitle}' creado y compartido. URL: ${sharedFileUrl}`);

        // --- 5. Enviar email a los destinatarios con el enlace ---
        const subject = `Llistat de Registres Actius - ${formattedDate} - ${sessionLocation}`;
        // Cuerpo del email (mantenido según tu última instrucción de no cambiarlo)
        const emailBodyHtml = `
            <html>
            <body>
                <h2>Llistat de Registres Actius per al ${formattedDate} (${sessionLocation})</h2>
                <p>Aquí tienes el enlace al documento de Google Sheets con el listado de registros activos para hoy en ${sessionLocation}. Puedes actualizar las columnas de Pago y Observaciones en este documento compartido:</p>
                <p><a href="${sharedFileUrl}">${newSheetTitle}</a></p>
                <p>Total de registres actius per al ${formattedDate} (${sessionLocation}): ${dailyRegistrations.length}</p>
                <br>
                <p><em>Aquest és un correu electrònic automàtic.</em></p>
            </body>
            </html>
        `;

        MailApp.sendEmail({
            to: emailRecipients,
            subject: subject,
            htmlBody: emailBodyHtml,
        });
        Logger.log(`INFO: Email con enlace a Google Sheet enviado a ${emailRecipients} para el ${formattedDate} (${sessionLocation}).`);

    } catch (e) {
        Logger.log(`ERROR: Error general al procesar, crear, compartir o enviar el enlace del Sheets: ${e.message}`);
        SpreadsheetApp.getUi().alert(`Error al procesar y enviar el listado para el ${formattedDate}: ${e.message}`);
    } finally {
        // --- 11. Opcional: Limpiar el documento temporal ---
        if (newSpreadsheet) {
            try {
                 DriveApp.getFileById(newSpreadsheet.getId()).setTrashed(true);
                 Logger.log(`INFO: Documento de Sheets temporal '${newSpreadsheet.getName()}' enviado a la papelera.`);
            } catch (deleteError) {
                 Logger.log(`WARN: No se pudo enviar el documento de Sheets temporal '${newSpreadsheet.getName()}' a la papelera: ${deleteError.message}.`);
            }
        }
    }
}
// --- Puedes mantener estas funciones si las necesitas ---
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
    const lastColumn = registrationsSheet.getLastColumn();

    if (lastRow > 1) {
        const dataRange = registrationsSheet.getRange(2, 1, lastRow - 1, lastColumn);
        dataRange.clearContent();
        Logger.log(`INFO: Se ha borrado el contenido de ${lastRow - 1} registros de la hoja '${REGISTRATIONS_SHEET_NAME}'.`);
        SpreadsheetApp.flush();
    } else {
        Logger.log(`INFO: La hoja '${REGISTRATIONS_SHEET_NAME}' ya está vacía (solo cabecera).`);
    }
}

// --- Funciones de prueba ---
function testDailyList() {
  Logger.log("INFO: Ejecutando testDailyList manualmente.");
  sendDailyLists();
}

function testWeeklyReset() {
  Logger.log("INFO: Ejecutando testWeeklyReset manualmente.");
  resetWeeklyRegistrations();
}

// --- NOTA: Elimina o comenta la función getResponsibleEmailForDate si todavía la tienes en tu script, ya no se usa. ---
// function getResponsibleEmailForDate(...) { ... }
