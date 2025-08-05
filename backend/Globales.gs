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



