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

    const sheetHeaders = ["Nombre", "Teléfono", "Email", "Socio/a", "Pago", "Observaciones"];

    const sheetData = dailyRegistrations.map(reg => {
        const registrationTime = Utilities.formatDate(reg.timestamp, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "HH:mm:ss");
        return [
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

  // --- AGREGAR FORMATO AL SHEETS DE SALIDA ---
        if (sheetHeaders.length > 0) {
            const headerRange = newSheet.getRange(1, 1, 1, sheetHeaders.length);
            const tableRange = newSheet.getRange(1, 1, 1 + sheetData.length, sheetHeaders.length); // Incluye cabecera y datos

            // 1. Color de fondo para las cabeceras
            const headerBackgroundColor = "#F48FB1"; // **Define tu color aquí (ej: un azul claro)**
            headerRange.setBackground(headerBackgroundColor);
             Logger.log(`INFO: Color de fondo '${headerBackgroundColor}' aplicado a las cabeceras.`);

            // 2. Bordes para toda la tabla
            // Aplica bordes a todos los lados (top, left, bottom, right, vertical, horizontal)
            tableRange.setBorder(true, true, true, true, true, true, null, SpreadsheetApp.BorderStyle.SOLID); // null para color por defecto, SOLID para estilo de línea
             Logger.log("INFO: Bordes aplicados a la tabla completa.");

            // Opcional: Negrita para las cabeceras
            headerRange.setFontWeight("bold");
             Logger.log("INFO: Cabeceras en negrita.");
        }
        // --- AGREGAR VALIDACIÓN DE DATOS (DESPLEGABLE) A LA COLUMNA DE PAGO ---
        const pagoColHeader = "Pago";
        const pagoColIndex = sheetHeaders.indexOf(pagoColHeader);

        if (pagoColIndex !== -1 && sheetData.length > 0) {
            const startRow = 2;
            const numRows = sheetData.length;
            const targetColumn = pagoColIndex + 1;

            const opcionesPago = ['1 Sessió', 'Bono 5', 'Bono Anual']; // **TUS OPCIONES**

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
                <p>Aquí tens l'enllaç al document de Google Sheets amb el llistat de registres actius per avui: ${sessionLocation}.</p>
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
    } /**finally {
        // --- 11. Opcional: Limpiar el documento temporal ---
        if (newSpreadsheet) {
            try {
                 DriveApp.getFileById(newSpreadsheet.getId()).setTrashed(true);
                 Logger.log(`INFO: Documento de Sheets temporal '${newSpreadsheet.getName()}' enviado a la papelera.`);
            } catch (deleteError) {
                 Logger.log(`WARN: No se pudo enviar el documento de Sheets temporal '${newSpreadsheet.getName()}' a la papelera: ${deleteError.message}.`);
            }
        }
    } */
}
// --- Funciones de prueba ---
function testDailyList() {
  Logger.log("INFO: Ejecutando testDailyList manualmente.");
  sendDailyLists();
}
