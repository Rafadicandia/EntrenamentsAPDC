
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



function testWeeklyReset() {
  Logger.log("INFO: Ejecutando testWeeklyReset manualmente.");
  resetWeeklyRegistrations();
}
