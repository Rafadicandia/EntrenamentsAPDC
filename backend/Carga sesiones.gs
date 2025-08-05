/**
 * @fileoverview Este script automatiza el cambio de estado de sesiones en Google Sheets
 * basándose en la fecha/hora actual, utilizando el backend de carga de sesiones.
 */





// Nombre de la nueva columna de estado que DEBES AÑADIR a tu hoja de SESIONES
const SESSION_STATUS_COLUMN_NAME = 'EstadoSesion'; 


// --- FUNCIÓN DE TU BACKEND PARA OBTENER SESIONES DISPONIBLES ---
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

  // Encontrar índices de columnas relevantes para SESIONES
  const sessionCol = {
    id: sessionsHeader.indexOf("SessionID"),
    name: sessionsHeader.indexOf("Name"),
    instructor: sessionsHeader.indexOf("Instructor"),
    location: sessionsHeader.indexOf("Location"),
    dateTime: sessionsHeader.indexOf(SESSION_DATETIME_COLUMN_NAME), // Usamos la constante aquí
    capacity: sessionsHeader.indexOf("Capacity"),
    status: sessionsHeader.indexOf(SESSION_STATUS_COLUMN_NAME) // CORREGIDO: Usando 'sessionsHeader.indexOf' y la constante
  };

  // Encontrar índices de columnas relevantes para REGISTROS
  const regCol = {
    registrationId: registrationsHeader.indexOf("RegistrationID"),
    timestamp: registrationsHeader.indexOf("Timestamp"),
    sessionId: registrationsHeader.indexOf("SessionID"),
    userName: registrationsHeader.indexOf("UserName"),
    phone: registrationsHeader.indexOf("UserPhone"),
    email: registrationsHeader.indexOf("UserEmail"),
    isMember: registrationsHeader.indexOf("IsMember"),
    status: registrationsHeader.indexOf("Status"),
    cancellationToken: registrationsHeader.indexOf("CancellationToken")
  };

  // Validar que todas las columnas necesarias existen en Sesiones
  if (Object.values(sessionCol).some(index => index === -1)) {
     const missingSessionCols = Object.keys(sessionCol).filter(key => sessionCol[key] === -1);
     throw new Error(`Faltan columnas requeridas en la hoja '${SESSIONS_SHEET_NAME}': ${missingSessionCols.join(', ')}. Revisa las cabeceras.`);
  }
  // Validar que todas las columnas necesarias existen en Registros
   if (Object.values(regCol).some(index => index === -1)) {
     const missingRegCols = Object.keys(regCol).filter(key => regCol[key] === -1);
     throw new Error(`Faltan columnas requeridas en la hoja '${REGISTRATIONS_SHEET_NAME}': ${missingRegCols.join(', ')}. Revisa las cabeceras.`);
  }

  // Contar registros activos por SessionID
  const activeRegistrationsCount = {};
  // Empezar desde 1 para saltar la cabecera
  for (let i = 1; i < registrationsData.length; i++) {
    const row = registrationsData[i];
    let sessionIdFromRegistrations = row[regCol.sessionId]; 

    if (sessionIdFromRegistrations instanceof Date) {
      sessionIdFromRegistrations = Utilities.formatDate(sessionIdFromRegistrations, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "dd/MM/yyyy");
    }
    
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
    const sessionIdFormatted = Utilities.formatDate(new Date(row[sessionCol.id]), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "dd/MM/yyyy");
    Logger.log(`Session ID from Sessions Sheet (raw): '${row[sessionCol.id]}', Formatted for comparison: '${sessionIdFormatted}'`);

    const capacity = parseInt(row[sessionCol.capacity], 10) || 0;
    const registered = activeRegistrationsCount[sessionIdFormatted] || 0;
    const remainingSpots = capacity - registered;
    const sessionStatus = row[sessionCol.status]; // AÑADIDO: Obtener el estado de la sesión

    availableSessions.push({
      id: sessionIdFormatted,
      name: row[sessionCol.name],
      instructor: row[sessionCol.instructor],
      location: row[sessionCol.location],
      dateTime: Utilities.formatDate(new Date(row[sessionCol.dateTime]), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "HH:mm"),
      capacity: capacity,
      registered: registered,
      remainingSpots: remainingSpots,
      sessionStatus: sessionStatus // CORREGIDO: Usando 'sessionStatus' para el estado de la sesión
    });
  }

  return availableSessions;
}


// --- NUEVA FUNCIÓN PARA ACTUALIZAR EL ESTADO DE LAS SESIONES ---
/**
 * Actualiza el estado de las sesiones a "Finalizado" si su fecha y hora
 * en la columna 'DateTime' es igual o anterior a la fecha y hora actual,
 * y su estado actual en 'EstadoSesion' no es ya "Finalizado".
 *
 * Esta función opera directamente sobre la hoja 'SESSIONS_SHEET_NAME'.
 */
function actualizarEstadoSesionPorFecha() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sessionsSheet = spreadsheet.getSheetByName(SESSIONS_SHEET_NAME);

  if (!sessionsSheet) {
    Logger.log('Error: La hoja de sesiones "' + SESSIONS_SHEET_NAME + '" no fue encontrada.');
    return;
  }

  // Obtener los valores de la cabecera para encontrar los índices de las columnas
  const headers = sessionsSheet.getRange(1, 1, 1, sessionsSheet.getLastColumn()).getValues()[0];

  const dateTimeColumnIndex = headers.indexOf(SESSION_DATETIME_COLUMN_NAME);
  const statusColumnIndex = headers.indexOf(SESSION_STATUS_COLUMN_NAME);

  // Verificar que las columnas necesarias existan
  if (dateTimeColumnIndex === -1) {
    Logger.log('Error: La columna "' + SESSION_DATETIME_COLUMN_NAME + '" no fue encontrada en la hoja "' + SESSIONS_SHEET_NAME + '". Asegúrate de que el nombre coincida exactamente.');
    return;
  }
  if (statusColumnIndex === -1) {
    Logger.log('Error: La columna "' + SESSION_STATUS_COLUMN_NAME + '" no fue encontrada en la hoja "' + SESSIONS_SHEET_NAME + '". Asegúrate de que has añadido esta columna y el nombre coincide exactamente.');
    return;
  }

  // Obtener todos los datos de la hoja de sesiones, excluyendo la fila de encabezado
  const lastRow = sessionsSheet.getLastRow();
  if (lastRow < 2) { 
    Logger.log('No hay datos de sesiones para procesar en la hoja "' + SESSIONS_SHEET_NAME + '".');
    return;
  }
  const dataRange = sessionsSheet.getRange(2, 1, lastRow - 1, sessionsSheet.getLastColumn());
  const values = dataRange.getValues();

  const now = new Date(); // Obtener la fecha y hora actual

  let updatedRowsCount = 0;

  // Iterar sobre cada fila de datos de sesiones
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const sessionDateTimeValue = row[dateTimeColumnIndex]; // Valor de la fecha/hora de la sesión
    const currentSessionStatus = row[statusColumnIndex]; // Estado actual de la sesión

    // Asegurarse de que el valor de la fecha/hora sea una instancia de Date.
    // Si la celda está vacía o no es una fecha válida, se ignora.
    if (!(sessionDateTimeValue instanceof Date)) {
        Logger.log(`Advertencia: La celda de fecha/hora en la fila ${i + 2} no contiene una fecha válida en la columna '${SESSION_DATETIME_COLUMN_NAME}'. Saltando.`);
        continue;
    }

    // Comprobar si la fecha/hora de la sesión ya ha pasado
    // y si el estado actual NO es "Finalizado"
    if (sessionDateTimeValue <= now && currentSessionStatus !== 'Finalizado') {
      // Actualizar el estado a "Finalizado"
      // La fila en la hoja de cálculo es `i + 2` (índice de matriz + fila de encabezado)
      // La columna es `statusColumnIndex + 1` (índice de matriz + 1 para Sheets)
      sessionsSheet.getRange(i + 2, statusColumnIndex + 1).setValue('Finalizado');
      updatedRowsCount++;
      Logger.log(`Fila ${i + 2}: Estado de sesión cambiado a "Finalizado" para la sesión con fecha ${sessionDateTimeValue.toLocaleString()}`);
    }
  }

  if (updatedRowsCount > 0) {
    Logger.log(`Proceso de actualización de estado de sesiones completado. Se actualizaron ${updatedRowsCount} sesiones.`);
  } else {
    Logger.log('Proceso de actualización de estado de sesiones completado. No se encontraron sesiones para actualizar.');
  }
}
