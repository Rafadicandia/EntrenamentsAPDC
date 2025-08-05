/**
 * Constantes globales para los nombres de las hojas y columnas.
 * Modifica estos valores si cambias los nombres en tu Google Sheet.
 */
const SESSIONS_SHEET_NAME = "Sesiones";
const REGISTRATIONS_SHEET_NAME = "Registros";
const ACTIVE_STATUS = "Activo";
const CANCELLED_STATUS = "Cancelado";

// --- Funciones Principales (Web App Endpoints) ---

/**
 * Función principal que maneja las solicitudes GET.
 * - Sin parámetros: Devuelve la lista de sesiones disponibles.
 * - Con parámetros 'action=cancel', 'registrationId', 'token': Procesa una solicitud de cancelación.
 * @param {object} e - El objeto de evento de la solicitud GET.
 * @returns {object} - ContentService output (JSON para sesiones, HTML para cancelación).
 */
function doGet(e) {
  try {
    // Verificar si es una solicitud de cancelación
    if (e.parameter.action === "cancel" && e.parameter.registrationId && e.parameter.token) {
      return handleCancellationRequest(e.parameter.registrationId, e.parameter.token);
    }
    // Si no, es una solicitud para obtener las sesiones
    else {
      const sessions = getAvailableSessions();
      // Devolver las sesiones como JSON
      return ContentService
        .createTextOutput(JSON.stringify({ sessions: sessions }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  } catch (error) {
    Logger.log("Error en doGet: " + error);
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
  try {
    // Parsear los datos JSON enviados desde el frontend
    const requestData = JSON.parse(e.postData.contents);

    // Validar datos básicos (se pueden añadir más validaciones)
    if (!requestData.sessionId || !requestData.name || !requestData.phone || !requestData.email || !requestData.isMember) {
      throw new Error("Faltan datos requeridos en la solicitud.");
    }

    const result = registerUserForSession(requestData);

    // Devolver resultado como JSON
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    Logger.log("Error en doPost: " + error);
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
    throw new Error("Una o ambas hojas ('Sesiones', 'Registros') no se encontraron.");
  }

  const sessionsData = sessionsSheet.getDataRange().getValues();
  const registrationsData = registrationsSheet.getDataRange().getValues();

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
    sessionId: registrationsHeader.indexOf("SessionID"),
    status: registrationsHeader.indexOf("Status")
  };

  // Validar que todas las columnas necesarias existen
  if (Object.values(sessionCol).some(index => index === -1) || Object.values(regCol).some(index => index === -1)) {
    throw new Error("Faltan columnas requeridas en las hojas 'Sesiones' o 'Registros'. Revisa las cabeceras.");
  }


  // Contar registros activos por SessionID
  const activeRegistrationsCount = {};
  // Empezar desde 1 para saltar la cabecera
  for (let i = 1; i < registrationsData.length; i++) {
    const row = registrationsData[i];
    const sessionId = row[regCol.sessionId];
    const status = row[regCol.status];
    if (status === ACTIVE_STATUS) {
      activeRegistrationsCount[sessionId] = (activeRegistrationsCount[sessionId] || 0) + 1;
    }
  }

  // Construir la lista de sesiones con plazas restantes
  const availableSessions = [];
  // Empezar desde 1 para saltar la cabecera
  for (let i = 1; i < sessionsData.length; i++) {
    const row = sessionsData[i];
    const sessionId = row[sessionCol.id];
    const capacity = parseInt(row[sessionCol.capacity], 10) || 0; // Asegurar que sea número
    const registered = activeRegistrationsCount[sessionId] || 0;
    const remainingSpots = capacity - registered;

    availableSessions.push({
      id: Utilities.formatDate(row[sessionCol.id], SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "dd/MM/yyyy"),
      name: row[sessionCol.name],
      instructor: row[sessionCol.instructor],
      location: row[sessionCol.location],
      // Formatear solo la hora y minutos (HH:mm)
      dateTime: Utilities.formatDate(row[sessionCol.dateTime], SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "HH:mm"),
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
    throw new Error("Hojas 'Sesiones' o 'Registros' no encontradas.");
  }

  // --- 1. Verificar Plazas Disponibles ---
  const sessions = getAvailableSessions(); // Reutilizamos la función para obtener datos actualizados
  const targetSession = sessions.find(s => s.id === data.sessionId);

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
  const isMemberText = data.isMember === 'true' ? 'Sí' : 'No'; // Convertir boolean/string a texto

  // Añadir fila a la hoja de Registros
  registrationsSheet.appendRow([
    registrationId,
    timestamp,
    data.sessionId,
    data.name,
    data.phone,
    data.email,
    isMemberText,
    ACTIVE_STATUS, // Estado inicial
    cancellationToken
  ]);

  // --- 3. Enviar Email de Confirmación (Opcional pero recomendado) ---
  try {
    sendConfirmationEmail(data.email, data.name, targetSession, registrationId, cancellationToken);
  } catch (emailError) {
    Logger.log("Error al enviar email de confirmación para " + data.email + ": " + emailError);
    // No consideramos esto un fallo crítico del registro, pero lo registramos.
    // Podrías devolver un mensaje indicando que el registro fue exitoso pero el email falló.
  }

  return { success: true, message: `¡Registro completado para ${targetSession.name}! Revisa tu email para la confirmación.` };
}

/**
 * Maneja una solicitud de cancelación verificando el token.
 * @param {string} registrationId - ID del registro a cancelar.
 * @param {string} token - Token de cancelación proporcionado.
 * @returns {object} - ContentService output (HTML con mensaje de confirmación/error).
 */
function handleCancellationRequest(registrationId, token) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const registrationsSheet = ss.getSheetByName(REGISTRATIONS_SHEET_NAME);

  if (!registrationsSheet) {
    return ContentService.createTextOutput("Error interno: Hoja de registros no encontrada.").setMimeType(ContentService.MimeType.HTML);
  }

  const data = registrationsSheet.getDataRange().getValues();
  const header = data[0];
  const idCol = header.indexOf("RegistrationID");
  const tokenCol = header.indexOf("CancellationToken");
  const statusCol = header.indexOf("Status");
  const sessionCol = header.indexOf("SessionID"); // Para loggear
  const nameCol = header.indexOf("UserName"); // Para loggear

  // Validar columnas
  if ([idCol, tokenCol, statusCol, sessionCol, nameCol].some(index => index === -1)) {
    Logger.log("Error de cancelación: Faltan columnas en la hoja 'Registros'.");
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
        // Actualizar el estado a Cancelado en la fila encontrada (i+1 porque los índices de hoja son base 1)
        registrationsSheet.getRange(i + 1, statusCol + 1).setValue(CANCELLED_STATUS);
        // (Opcional) Limpiar el token para que no se pueda reusar
        // registrationsSheet.getRange(i + 1, tokenCol + 1).setValue('');
        message = `El teu registre per a la sessió ha estat cancel·lat correctament. Esperem veure't aviat!`;
        Logger.log(`Cancel·lació correcta: ID ${registrationId}, Usuario ${data[i][nameCol]}, Sesión ${data[i][sessionCol]}`);
        SpreadsheetApp.flush(); // Asegurar que los cambios se guarden
      } else if (data[i][statusCol] === CANCELLED_STATUS) {
        alreadyCancelled = true;
        message = "Aquest registre ja havia estat cancel·lat prèviament.";
        Logger.log(`Intent de cancel·lació repetit: ID ${registrationId}`);
      }
      break; // Salir del bucle una vez encontrado
    }
  }

  if (!found) {
    Logger.log(`Intent de cancel·lació fallit: ID ${registrationId} o Token inválido.`);
  }

  // 1. Carga el archivo HTML como una plantilla
  const template = HtmlService.createTemplateFromFile('CancelPage');

  // 2. Pasa variables del script (.gs) a la plantilla (.html)
  template.message = message;
  template.scriptUrl = ScriptApp.getService().getUrl(); // Pasamos la URL del script

  // 3. Evalúa la plantilla para obtener el resultado HTML final
  const htmlOutput = template.evaluate();

  // Opcional: Puedes establecer el título de la página aquí también si quieres
  // htmlOutput.setTitle('Estado de la Cancelación');

  // 4. Devuelve el objeto HtmlOutput
  return htmlOutput;
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
  // Construir la URL base del script desplegado
  const scriptUrl = ScriptApp.getService().getUrl();
  // Crear el enlace de cancelación
  const cancellationLink = `${scriptUrl}?action=cancel&registrationId=${registrationId}&token=${cancellationToken}`;

  const body = `
    Hola ${recipientName},<br><br>
    Has estat registrat/ada correctament per a la següent sessió:<br>
    <b>Sessió:</b> ${session.name}<br>
    <b>Instructora:</b> ${session.instructor}<br>
    <b>Lloc:</b> ${session.location}<br>
    <b>Data:</b> ${session.id}<br>
    <b>Horari:</b> ${session.dateTime}<br><br>
    Si necessites cancel·lar la teva plaça, si us plau, fes clic al següent enllaç:<br>
    <a href="${cancellationLink}">Cancel·lar el meu registre</a><br><br>
    Si no pots fer clic a l'enllaç, copia i enganxa la següent URL al teu navegador:<br>
    ${cancellationLink}<br><br>
    Ens veiem a la classe!<br><br>
    <em>(Aquest és un missatge automàtic, si us plau, no responguis directament.)</em>
  `;

  // Enviar el email usando el servicio MailApp de Google
  MailApp.sendEmail({
    to: recipientEmail,
    subject: subject,
    htmlBody: body, // Usar htmlBody para que el enlace sea clickeable
    // Opcional: nombre del remitente
    // name: "Sistema de Registro de Entrenamientos"
  });

  Logger.log(`Email de confirmación enviado a ${recipientEmail} para el registro ${registrationId}`);
}
