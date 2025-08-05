// --- NUEVA FUNCIÓN PARA ACTUALIZAR EL ESTADO DE LAS SESIONES ---
/**
 * Actualiza el estado de las sesiones a "Finalizado" si su fecha y hora
 * en la columna 'DateTime' es igual o anterior a la fecha y hora actual,
 * y su estado actual en 'EstadoSesion' no es ya "Finalizado".
 *
 * Esta función opera directamente sobre la hoja 'SESSIONS_SHEET_NAME'.
 */

// Nombre de la nueva columna de estado que DEBES AÑADIR a tu hoja de SESIONES

// Nombre de la columna que contiene la fecha y hora de la sesión
const SESSION_DATETIME_COLUMN_NAME = 'SessionID'; 

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