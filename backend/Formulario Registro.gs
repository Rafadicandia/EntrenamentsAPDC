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
