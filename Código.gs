// Code.gs (Este es el archivo principal de tu proyecto de Apps Script)

// ====================================================================================================================
// =============================== 1. CONFIGURACIÓN GLOBAL Y VARIABLES ==============================================
// ====================================================================================================================

// ID de la hoja de cálculo de Google
const SPREADSHEET_ID = "1h4TxPJHZ8pynph3J6q2h4FnDyOnG0Uye3VrYqFriPCg"; 

// Información para conectar con la base de datos SQL Server
const DB_ADDRESS = 'gw.hemoeco.com:5300';
const DB_USER = 'caleb.aranda';
const DB_PWD = '5AA3hmq8BfJFkrISTWgJsA==';
const DB_NAME = 'Soporte_Pruebas';
const DB_URL = 'jdbc:sqlserver://' + DB_ADDRESS + ';databaseName=' + DB_NAME;

// URL para la redirección a la aplicación "Principal" (tu menú principal)
const REDIRECTION_URL = "https://script.google.com/a/macros/hemoeco.com/s/AKfycbzQpqK85N-s6l4Qz0dMQGqD1ePr8O-PT5eJ87wYyXcLsuK_GU_lJmv-j38-xp8kbKOxnQ/exec";

// Mapeo de nombres de sucursales (formulario) a códigos (base de datos - INT)
const SUCURSAL_MAP = {
    "ADM": 1, // Asumiendo que "ADM" se mapea a INT 1
    "GDL": 5,
    "MXL": 6,
    "MEX": 8,
    "TIJ": 4,
    "CAN": 3,
    "SJD": 7,
    "MTY": 9
};

// Correos de los administradores para validaciones
const ADMIN_EMAILS = ["caleb.aranda@hemoeco.com"];

// --- ENCABEZADOS ESPECÍFICOS PARA CADA HOJA DE RESPUESTAS ---
// ¡IMPORTANTE! Asegúrate de que las columnas en tus hojas de cálculo coincidan EXACTAMENTE con estos arrays, incluyendo el orden.

// Encabezados para la hoja "RECU" (Registrar Equipo Usado)
// Alineados con el orden de las columnas en la hoja de cálculo
const RECU_SHEET_HEADERS = [
    "Marca temporal",               
    "Dirección de correo electrónico",
    "Costo del Equipo",             
    "Fecha de compra de Equipo",    
    "Fecha de Recolección",        
    "Fecha de Reasignacion",        
    "Estado del equipo",            
    "Observaciones",                
    "Marca",                        
    "Modelo",                       
    "Memoria RAM",                  
    "Almacenamiento (Memoria ROM)", 
    "IMEI",                         
    "Sucursal",                     
    "IDEquipo",                     
    "Error",                        
    "IDSucursal",                   
    "EJECUTADO",                    
    "Comentarios"                  
];

// Encabezados para la hoja "ALyE" (Dar de Alta Línea y Equipo)
const ALyE_SHEET_HEADERS = [
    "Marca temporal",              
    "Dirección de correo electrónico", 
    "Id_tel",                       
    "Región",                       
    "Cuenta_padre",                
    "Cuenta",                      
    "Teléfono",                   
    "Clave_plan",                  
    "Nombre_plan",              
    "Minutos",                     
    "Mensajes",                  
    "Monto_renta",                 
    "Equipo_ilimitado",           
    "Duracion_plan",              
    "Fecha_inicio",               
    "Fecha_termino",               
    "Marca_linea",
    "Modelo_linea",
    "IMEI_linea",
    "SIM",                         
    "Tipo",                        
    "Responsable_Linea",                  
    "Notas",                       
    "IDEMPLEADO_Linea",            
    "Sucursal_Linea",
    "IDSUCURSAL_Telcel",            
    "Datos",                        
    "Extensión",                    
    "ID_Equipo_Nuevo",             
    "Error_Linea",                      
    "EJECUTADO_Linea",                  

    "ID_Equipo",
    "Costo_Equipo",                 
    "Fecha_Compra_Equipo",         
    "Estado_Equipo",               
    "Observaciones_Equipo_Nuevo",   
    "Marca_Equipo_Nuevo",           
    "Modelo_Equipo_Nuevo",         
    "RAM_Equipo_Nuevo",             
    "ROM_Equipo_Nuevo",             
    "IMEI_Equipo_Nuevo",            
    "IDEMPLEADO_Equipo",            
    "Responsable_Equipo",                 
    "Sucursal_Equipo",
    "IDSUCURSAL_Equipo",            
    "Error_Equipo",                        
    "EJECUTADO_Equipo"                     
];


// Placeholder para los encabezados de las otras hojas.
const RLyE_SHEET_HEADERS = ["Marca temporal", "Campo1_RLyE", "Campo2_RLyE", "EJECUTADO"];
const ML_SHEET_HEADERS = ["Marca temporal", "Campo1_ML", "Campo2_ML", "EJECUTADO"];
const MEU_SHEET_HEADERS = ["Marca temporal", "Campo1_MEU", "Campo2_MEU", "EJECUTADO"];


// ====================================================================================================================
// ====================================== 2. FUNCIONES GENERALES DEL PROYECTO =======================================
// ====================================================================================================================

/**
 * Sirve el archivo HTML del formulario según el parámetro 'form' en la URL.
 * Se invoca automáticamente cuando se accede a la URL de la aplicación web de Apps Script.
 * Ejemplo de URL: https://script.google.com/macros/s/.../exec?form=registrarEquipoUsado
 * O para acciones de correo: https://script.google.com/macros/s/.../exec?action=aprobarBajaEquipo&idEquipo=...
 * @param {GoogleAppsScript.Events.DoGet} e El objeto de evento que contiene los parámetros de la URL.
 */
function doGet(e) {
    // --- Manejo de acciones desde botones de correo ---
    if (e.parameter.action) {
        const action = e.parameter.action;
        const params = e.parameter; // Todos los parámetros se pasan a la función de acción

        let htmlOutput;
        switch (action) {
            case 'aprobarBajaEquipo':
                htmlOutput = aprobarBajaEquipo(
                    parseInt(params.idEquipo), 
                    params.solicitanteEmail, 
                    params.razonBaja, 
                    params.sucursal,
                    params.imei 
                );
                break;
            case 'denegarBajaEquipo':
                htmlOutput = denegarBajaEquipo(
                    parseInt(params.idEquipo), 
                    params.solicitanteEmail, 
                    params.sucursal,
                    params.imei 
                );
                break;
            case 'aprobarVentaEquipoStep1':
                // Para step1, se redirige a un formulario, no se genera una página de confirmación directa aquí
                const scriptUrlBase = ScriptApp.getService().getUrl();
                const redirectUrl = `${scriptUrlBase}?form=aprobarVentaForm&idEquipo=${params.idEquipo}&solicitanteEmail=${encodeURIComponent(params.solicitanteEmail)}&personaVende=${encodeURIComponent(params.personaVende)}&sucursal=${encodeURIComponent(params.sucursal)}&imei=${encodeURIComponent(params.imei)}`;
                return HtmlService.createHtmlOutput(`<script>window.top.location.href = '${redirectUrl}';</script>`);
            case 'denegarVentaEquipo':
                htmlOutput = denegarVentaEquipo(
                    parseInt(params.idEquipo), 
                    params.solicitanteEmail, 
                    params.sucursal,
                    params.imei 
                );
                break;
            case 'aprobarVentaEquipoStep2': // Este es llamado por el formulario de monto
                htmlOutput = aprobarVentaEquipoStep2(params);
                break;
            default:
                logMessage(`Acción no reconocida: ${action}`);
                // Retorna una página de error con estilo simple
                htmlOutput = generateConfirmationPage(
                    'Error de Acción',
                    'Acción no reconocida o inválida.',
                    true // isError
                );
                break;
        }
        return htmlOutput; // Retorna el HtmlOutput generado por la función de acción o error
    }

    // --- Manejo de carga de formularios HTML ---
    let formName = e.parameter.form || 'registrarEquipoUsado'; 

    let htmlFileToServe;
    switch (formName) {
        case 'registrarEquipoUsado':
            htmlFileToServe = 'registrarEquipoUsado';
            break;
        case 'darAltaLineaEquipo':
            htmlFileToServe = 'darAltaLineaEquipo';
            break;
        case 'renovacionLineaEquipo':
            htmlFileToServe = 'renovacionLineaEquipo';
            break;
        case 'modificarLinea':
            htmlFileToServe = 'modificarLinea';
            break;
        case 'modificarEquipoUsado':
            htmlFileToserve = 'modificarEquipoUsado';
            break;
        case 'aprobarVentaForm': // Nuevo caso para el formulario de aprobación de venta
            htmlFileToServe = 'aprobarVentaForm';
            break;
        default:
            // Si el nombre del formulario es inválido, se carga el formulario predeterminado y se registra una advertencia.
            htmlFileToServe = 'registrarEquipoUsado';
            logMessage("Advertencia: Nombre de formulario inválido recibido: '" + formName + "'. Se carga 'registrarEquipoUsado'.");
            break;
    }

    // Importante: Usar createTemplateFromFile().evaluate() para procesar <?!= ... ?>
    const template = HtmlService.createTemplateFromFile(htmlFileToServe);
    
    // Pasar parámetros del evento 'e' a la plantilla HTML si es necesario
    // Esto es útil para pasar IDs u otros datos a los formularios de aprobación
    if (e.parameter) {
        for (let param in e.parameter) {
            template[param] = e.parameter[param];
        }
    }

    return template.evaluate()
        .setTitle('Formularios Hemoeco')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // Importante para que funcione en un iframe
}

/**
 * Redirige a la URL de la aplicación web "Principal".
 * Llamado desde el botón 'Volver al Menú'.
 * @returns {string} La URL de redirección.
 */
function redirigirOtraApp() {
    Logger.log("Redirigiendo al menú principal: " + REDIRECTION_URL);
    return REDIRECTION_URL;
}

/**
 * Abre la hoja de cálculo específica por su ID.
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet} El objeto Spreadsheet.
 */
function getSpreadsheet() {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
}

/**
 * Obtiene la hoja (pestaña) específica del libro de cálculo.
 * @param {string} sheetName El nombre de la pestaña a obtener.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} El objeto Sheet.
 */
function getSheet(sheetName) {
    const spreadsheet = getSpreadsheet();
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
        throw new Error(`La hoja '${sheetName}' no se encontró en el libro de cálculo. Por favor, asegúrate de que exista y el nombre sea exacto.`);
    }
    return sheet;
}

/**
 * Registra un mensaje en el log de Apps Script y, opcionalmente, en una hoja de log.
 * Se asume la existencia de una hoja llamada "Log" en el mismo Spreadsheet.
 */
function logMessage(message) {
    Logger.log(message);
    try {
        const logSheet = getSpreadsheet().getSheetByName("Log"); // Asume que hay una hoja "Log"
        if (logSheet) {
            logSheet.appendRow([new Date(), Session.getActiveUser().getEmail(), message]);
        }
    } catch (e) {
        Logger.log("Error al escribir en la hoja de Log: " + e.message);
    }
}

/**
 * Establece una conexión a la base de datos SQL Server.
 * @returns {JdbcConnection} La conexión JDBC.
 */
function getJdbcConnection() {
    try {
        return Jdbc.getConnection(DB_URL, DB_USER, DB_PWD);
    } catch (e) {
        logMessage("Error al conectar con la base de datos SQL Server: " + e.message);
        throw new Error("No se pudo conectar a la base de datos: " + e.message);
    }
}

/**
 * Función auxiliar para formatear fechas para SQL Server DATETIME2.
 * @param {Date|string} dateValue El valor de fecha a formatear. Puede ser un objeto Date o un string.
 * @returns {string|null} La fecha formateada o null si es inválida/vacía.
 */
function formatDateForSql(dateValue) {
    if (!dateValue || (typeof dateValue === 'string' && dateValue.trim() === '')) return null;
    let date;
    if (dateValue instanceof Date) {
        date = dateValue;
    } else {
        date = new Date(dateValue);
    }

    if (isNaN(date.getTime())) {
        logMessage(`Advertencia: Fecha inválida detectada en formatDateForSql: ${dateValue}`);
        return null;
    }
    return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss.SSS");
}

/**
 * Obtiene opciones para un desplegable desde una hoja de cálculo.
 * @param {string} sheetName El nombre de la hoja donde están las opciones (ej. "Listados").
 * @param {string} rangeA1 La notación A1 del rango que contiene las opciones (ej. "A2:A").
 * @returns {Object} Un objeto con 'success' (boolean) y 'data' (array de strings) o 'message' de error.
 */
function getDropdownOptions(sheetName, rangeA1) {
  try {
    const sheet = getSpreadsheet().getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`La hoja '${sheetName}' no se encontró.`);
    }
    const values = sheet.getRange(rangeA1).getValues();
    // Flatten the array of arrays into a single array of values and filter out any empty strings
    const flatList = values.flat().filter(String); 
    logMessage(`Opciones obtenidas de ${sheetName}!${rangeA1}: ${JSON.stringify(flatList)}`);
    return { success: true, data: flatList };
  } catch (e) {
    logMessage("Error al obtener opciones del desplegable: " + e.message);
    return { success: false, message: e.message };
  }
}

/**
 * Envía un correo electrónico con botones de acción que llaman a funciones de Apps Script.
 * @param {string|string[]} recipient El/los correo(s) electrónico(s) del destinatario.
 * @param {string} subject El asunto del correo.
 * @param {string} body El cuerpo del correo (texto plano).
 * @param {Array<Object>} buttons Un array de objetos { text: 'Texto del botón', action: 'nombreFuncionAppsScript', params: {param1: 'valor'} }.
 */
function sendEmailWithButtons(recipient, subject, body, buttons) {
    let htmlBody = `<p>${body.replace(/\n/g, '<br>')}</p>`;
    htmlBody += '<p style="margin-top: 20px;">';

    const scriptUrlBase = ScriptApp.getService().getUrl(); // Obtiene la URL de la aplicación web actual

    buttons.forEach(button => {
        // Construye la URL para el botón, incluyendo la acción y los parámetros
        let actionUrl = `${scriptUrlBase}?action=${button.action}`;
        if (button.params) {
            for (let param in button.params) {
                actionUrl += `&${param}=${encodeURIComponent(button.params[param])}`;
            }
        }
        htmlBody += `<a href="${actionUrl}" style="display: inline-block; padding: 10px 20px; margin-right: 10px; border-radius: 5px; text-decoration: none; color: white; background-color: ${button.color || '#4CAF50'};">`;
        htmlBody += `${button.text}</a>`;
    });
    htmlBody += '</p>';

    MailApp.sendEmail({
        to: Array.isArray(recipient) ? recipient.join(',') : recipient, // Corregido: Asegura que 'to' sea una cadena separada por comas
        subject: subject,
        htmlBody: htmlBody,
    });
    logMessage(`Correo enviado a ${recipient} con asunto: "${subject}"`);
}

/**
 * Genera una página de confirmación HTML usando reemplazo de cadenas.
 * @param {string} title El título de la página.
 * @param {string} message El mensaje a mostrar.
 * @param {boolean} isError Si la página es para un error.
 * @returns {GoogleAppsScript.HTML.HtmlOutput} El objeto HtmlOutput.
 */
function generateConfirmationPage(title, message, isError) {
    const messageStyle = isError ? 'color: #dc3545;' : 'color: #5cb85c;';
    const titleStyle = isError ? 'color: #dc3545;' : 'color: #2d3748;';
    const containerClass = isError ? 'error-state' : '';

    let htmlContent = `
        <!DOCTYPE html>
        <html>
        <head>
            <base target="_top">
            <title>${title}</title>
            <style>
                body {
                    font-family: sans-serif;
                    display: flex;
                    justify-content: center;
                    align-items: center;
                    min-height: 100vh;
                    background-color: #f0f0f0;
                    margin: 0;
                }
                .container {
                    background-color: #ffffff;
                    padding: 30px;
                    border-radius: 8px;
                    box-shadow: 0 2px 10px rgba(0,0,0,0.1);
                    text-align: center;
                    max-width: 500px;
                    width: 90%;
                }
                h1 {
                    font-size: 1.8em;
                    margin-bottom: 15px;
                    ${titleStyle}
                }
                p {
                    font-size: 1.1em;
                    line-height: 1.5;
                    margin-bottom: 0;
                    ${messageStyle}
                }
            </style>
        </head>
        <body>
            <div class="container ${containerClass}">
                <h1>${title}</h1>
                <p>${message}</p>
            </div>
        </body>
        </html>
    `;
    return HtmlService.createHtmlOutput(htmlContent)
        .setTitle(title)
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


// ====================================================================================================================
// ====================================== 3. FUNCIONES DE INTERACCIÓN CON SQL SERVER ==================================
// ====================================================================================================================

/**
 * Obtiene el último ID_Equipo de la tabla Equipo_Usado y devuelve el siguiente consecutivo.
 * Si la tabla está vacía, devuelve 1.
 * @param {JdbcConnection} conn La conexión JDBC ya establecida.
 * @returns {number} El siguiente ID_Equipo consecutivo.
 */
function obtenerUltimoIDEquipoSQL(conn) {
    let stmt = null;
    let results = null;
    let nextId = 1;

    try {
        stmt = conn.createStatement();
        const query = "SELECT MAX(ID_Equipo) AS MaxID FROM Equipo_Usado";
        results = stmt.executeQuery(query);

        if (results.next()) {
            const maxId = results.getInt("MaxID");
            if (!results.wasNull()) {
                nextId = maxId + 1;
            }
        }
        logMessage("Siguiente ID_Equipo para generar (Equipo_Usado): " + nextId);
        return nextId;
    } catch (e) {
        logMessage("Error al obtener el último ID_Equipo de SQL Server (Equipo_Usado): " + e.message);
        throw new Error("Error al obtener ID de equipo (Equipo_Usado): " + e.message);
    } finally {
        if (results) results.close();
        if (stmt) stmt.close();
    }
}

/**
 * Obtiene el último ID_Equipo de la tabla Equipo_Nuevo y devuelve el siguiente consecutivo.
 * Si la tabla está vacía, devuelve 1.
 * @param {JdbcConnection} conn La conexión JDBC ya establecida.
 * @returns {number} El siguiente ID_Equipo consecutivo.
 */
function obtenerUltimoIDEquipoNuevoSQL(conn) {
    let stmt = null;
    let results = null;
    let nextId = 1;

    try {
        stmt = conn.createStatement();
        const query = "SELECT MAX(ID_Equipo) AS MaxID FROM Equipo_Nuevo";
        results = stmt.executeQuery(query);

        if (results.next()) {
            const maxId = results.getInt("MaxID");
            if (!results.wasNull()) {
                nextId = maxId + 1;
            }
        }
        logMessage("Siguiente ID_Equipo para generar (Equipo_Nuevo): " + nextId);
        return nextId;
    } catch (e) {
        logMessage("Error al obtener el último ID_Equipo de SQL Server (Equipo_Nuevo): " + e.message);
        throw new Error("Error al obtener ID de equipo (Equipo_Nuevo): " + e.message);
    } finally {
        if (results) results.close();
        if (stmt) stmt.close();
    }
}


/**
 * Valida si un IMEI ya existe en la tabla Equipo_Usado.
 * @param {string} imei El IMEI a validar.
 * @param {JdbcConnection} conn La conexión JDBC ya establecida.
 * @returns {boolean} True si el IMEI es único, false si ya existe.
 */
function validarIMEIUnicoSQL(imei, conn) {
    let pstmt = null;
    let results = null;
    let isUnique = true;

    try {
        const query = "SELECT COUNT(*) AS CountIMEI FROM Equipo_Usado WHERE IMEI = ?";
        pstmt = conn.prepareStatement(query);
        pstmt.setString(1, imei);
        results = pstmt.executeQuery();

        if (results.next()) {
            if (results.getInt("CountIMEI") > 0) {
                isUnique = false;
            }
        }
        logMessage(`IMEI ${imei} es único en Equipo_Usado: ${isUnique}`);
        return isUnique;
    } catch (e) {
        logMessage("Error al validar IMEI en SQL Server (Equipo_Usado): " + e.message);
        throw new Error("Error al validar IMEI (Equipo_Usado): " + e.message);
    } finally {
        if (results) results.close();
        if (pstmt) pstmt.close();
    }
}

/**
 * Valida si un IMEI ya existe en la tabla Equipo_Nuevo, Equipo_Usado o Telefonía_Telcel.
 * @param {string} imei El IMEI a validar.
 * @param {JdbcConnection} conn La conexión JDBC ya establecida.
 * @returns {boolean} True si el IMEI es único, false si ya existe.
 */
function validarIMEINuevoUnicoSQL(imei, conn) {
    let pstmt = null;
    let results = null;
    let isUnique = true;

    try {
        // Verificar en Equipo_Nuevo
        let query = "SELECT COUNT(*) AS CountIMEI FROM Equipo_Nuevo WHERE IMEI = ?";
        pstmt = conn.prepareStatement(query);
        pstmt.setString(1, imei);
        results = pstmt.executeQuery();
        if (results.next() && results.getInt("CountIMEI") > 0) {
            isUnique = false;
        }
        results.close();
        pstmt.close();

        if (!isUnique) {
            logMessage(`IMEI ${imei} ya existe en Equipo_Nuevo.`);
            return false;
        }


        //verificar en Equipo_Usado
        query = "SELECT COUNT(*) AS CountIMEI FROM Equipo_Usado WHERE IMEI = ?";
        pstmt = conn.prepareStatement(query);
        pstmt.setString(1, imei);
        results = pstmt.executeQuery();
        if (results.next() && results.getInt("CountIMEI") > 0) {
            isUnique = false;
        }

        if (!isUnique) {
            logMessage(`IMEI ${imei} ya existe en Equipo_Usado.`);
            return false;
        }

        // Verificar en Telefonía_Telcel
        query = "SELECT COUNT(*) AS CountIMEI FROM Telefonía_Telcel WHERE IMEI = ?";
        pstmt = conn.prepareStatement(query);
        pstmt.setString(1, imei);
        results = pstmt.executeQuery();
        if (results.next() && results.getInt("CountIMEI") > 0) {
            isUnique = false;
        }
        logMessage(`IMEI ${imei} es único en Equipo_Nuevo, Equipo_Usado y Telefonía_Telcel: ${isUnique}`);
        return isUnique;
    } catch (e) {
        logMessage("Error al validar IMEI en SQL Server (Equipo_Nuevo/Telefonía_Telcel): " + e.message);
        throw new Error("Error al validar IMEI (Equipo_Nuevo/Telefonía_Telcel): " + e.message);
    } finally {
        if (results) results.close();
        if (pstmt) pstmt.close();
    }
}

/**
 * Valida si un número de teléfono ya existe en la tabla Telefonía_Telcel.
 * @param {string} telefono El número de teléfono a validar.
 * @param {JdbcConnection} conn La conexión JDBC ya establecida.
 * @returns {boolean} True si el teléfono es único, false si ya existe.
 */
function validarTelefonoUnicoSQL(telefono, conn) {
    let pstmt = null;
    let results = null;
    let isUnique = true;

    try {
        const query = "SELECT COUNT(*) AS CountTelefono FROM Telefonía_Telcel WHERE Teléfono = ?";
        pstmt = conn.prepareStatement(query);
        pstmt.setString(1, telefono);
        results = pstmt.executeQuery();

        if (results.next()) {
            if (results.getInt("CountTelefono") > 0) {
                isUnique = false;
            }
        }
        logMessage(`Teléfono ${telefono} es único: ${isUnique}`);
        return isUnique;
    } catch (e) {
        logMessage("Error al validar Teléfono en SQL Server: " + e.message);
        throw new Error("Error al validar Teléfono: " + e.message);
    } finally {
        if (results) results.close();
        if (pstmt) pstmt.close();
    }
}

/**
 * Obtiene el ID de Empleado de la hoja "BD" basado en el nombre del responsable.
 * @param {string} responsableName El nombre del responsable a buscar.
 * @returns {string|null} El ID del empleado o null si no se encuentra.
 */
function getResponsableID(responsableName) {
    if (!responsableName) return null;
    try {
        const sheetBD = getSheet("BD"); // Hoja "BD"
        const range = sheetBD.getDataRange();
        const values = range.getValues(); // Obtiene todos los valores de la hoja

        const nombreColIndex = 0; // Columna A (0-indexed)
        const idEmpleadoColIndex = 1; // Columna B (0-indexed)

        for (let i = 0; i < values.length; i++) {
            if (values[i][nombreColIndex] && values[i][nombreColIndex].toString().trim() === responsableName.trim()) {
                return values[i][idEmpleadoColIndex] ? values[i][idEmpleadoColIndex].toString() : null;
            }
        }
        logMessage(`ID de Empleado no encontrado para Responsable: ${responsableName}`);
        return null;
    } catch (e) {
        logMessage("Error al obtener ID de Responsable de la hoja 'BD': " + e.message);
        return null;
    }
}

/**
 * Obtiene el valor de IDRESGUARDO de la base de datos basado en el IDSUCURSAL.
 * ESTA ES UNA FUNCIÓN PLACEHOLDER.
 * DEBERÁS IMPLEMENTAR LA LÓGICA PARA CONSULTAR TU BD Y OBTENER EL IDRESGUARDO.
 * Ejemplo: Podrías tener una tabla de configuración o mapeo en tu SQL Server.
 * @param {number} idSucursal El ID de la sucursal (INT).
 * @param {JdbcConnection} conn La conexión JDBC.
 * @returns {string|null} El ID de resguardo (NVARCHAR) o null si no se encuentra.
 */
function getResguardoPorSucursalSQL(idSucursal, conn) {
    let pstmt = null;
    let results = null;
    try {
        // Ejemplo de consulta (AJUSTA ESTO A TU ESQUEMA REAL DE DB)
        // Asumo una tabla 'ConfiguracionSucursales' con 'IDSUCURSAL' (INT) y 'IDRESGUARDO' (NVARCHAR)
        const query = "SELECT IDRESGUARDO FROM ConfiguracionSucursales WHERE IDSUCURSAL = ?";
        pstmt = conn.prepareStatement(query);
        pstmt.setInt(1, idSucursal); // idSucursal es INT
        results = pstmt.executeQuery();
        if (results.next()) {
            return results.getString("IDRESGUARDO");
        }
        logMessage(`ID de Resguardo no encontrado en DB para IDSUCURSAL: ${idSucursal}`);
        return null;
    } catch (e) {
        logMessage("Error al obtener ID de Resguardo por Sucursal de SQL Server: " + e.message);
        return null;
    } finally {
        if (results) results.close();
        if (pstmt) pstmt.close();
    }
}


// ====================================================================================================================
// ====================================== 4. FUNCIONES DE PROCESAMIENTO DE FORMULARIOS ==============================
// ====================================================================================================================

// --- 4.1. Formulario: Registrar Equipo Usado (RECU) ---
/**
 * Procesa los datos enviados desde el formulario "Registrar Equipo Usado".
 * Primero inserta los datos directamente en SQL Server, luego los registra en la hoja "RECU".
 * @param {Object} formData Los datos del formulario como un objeto JavaScript.
 * @returns {Object} Un objeto con 'success' (boolean) y 'message' (string).
 */
function procesarRECUFormulario(formData) {
    let response = { success: false, message: "" };
    logMessage("Datos recibidos del formulario RECU: " + JSON.stringify(formData));

    let conn = null;
    let idEquipo = null; // Declarar aquí e inicializar a null

    try {
        const sheet = getSheet("RECU"); // Hoja específica para RECU
        
        conn = getJdbcConnection();

        const fechaFormulario = new Date(); // Fecha actual del envío del formulario
        const solicitanteEmail = Session.getActiveUser().getEmail(); // Correo del usuario que envía el formulario

        // --- Validaciones de campos (lado del servidor) ---
        const costoEquipo = parseFloat(formData.costoEquipo) || 0;
        if (isNaN(costoEquipo) || costoEquipo < 1) {
            response.message = "El costo del equipo debe ser un número mayor o igual a $1.00.";
            logMessage("Error: " + response.message);
            return response;
        }

        const fechaCompra = formData.fechaCompra ? new Date(formData.fechaCompra) : null;
        const today = new Date();
        today.setHours(0,0,0,0); // Normalizar a inicio del día para comparación
        if (!fechaCompra || isNaN(fechaCompra.getTime()) || fechaCompra >= today) {
            response.message = "La fecha de compra debe ser una fecha válida y menor a la fecha actual.";
            logMessage("Error: " + response.message);
            return response;
        }

        const fechaRecoleccion = formData.fechaRecoleccion ? new Date(formData.fechaRecoleccion) : null;
        if (!fechaRecoleccion || isNaN(fechaRecoleccion.getTime()) || fechaRecoleccion > today) {
            response.message = "La fecha de recolección debe ser una fecha válida y no puede ser futura.";
            logMessage("Error: " + response.message);
            return response;
        }

        const imei = formData.imei ? formData.imei.trim() : '';
        if (!imei || !/^\d{15,16}$/.test(imei)) {
            response.message = "IMEI inválido (debe ser numérico, 15 o 16 dígitos).";
            logMessage("Error: " + response.message);
            return response;
        }
        // Validar IMEI en ambas tablas
        if (!validarIMEIUnicoSQL(imei, conn)) { // Valida en Equipo_Usado
            response.message = `IMEI '${imei}' ya existe en la base de datos (Equipo_Usado).`;
            logMessage("Error: " + response.message);
            return response;
        }
        if (!validarIMEINuevoUnicoSQL(imei, conn)) { // Valida también en Equipo_Nuevo y Telefonía_Telcel
            response.message = `IMEI '${imei}' ya existe en la base de datos (Equipo_Nuevo o Telefonía_Telcel).`;
            logMessage("Error: " + response.message);
            return response;
        }

        const numeroTelefono = formData.numeroTelefono ? formData.numeroTelefono.trim() : null;
        if (numeroTelefono && !/^\d{10}$/.test(numeroTelefono)) {
            response.message = "Número de Teléfono inválido (debe ser numérico, 10 dígitos).";
            logMessage("Error: " + response.message);
            return response;
        }

        const idSucursalBD = SUCURSAL_MAP[formData.idSucursal] || null; // Usar null si no se encuentra
        if (idSucursalBD === null) {
            response.message = "ID de Sucursal inválido.";
            logMessage("Error: " + response.message);
            return response;
        }
        const sucursalCodigo3Letras = formData.idSucursal; // Para el asunto del correo

        const responsableName = formData.responsable || null;
        let idEmpleado = null;
        if (responsableName) {
            idEmpleado = getResponsableID(responsableName);
            if (!idEmpleado) {
                response.message = `No se encontró ID de Empleado para el responsable '${responsableName}'.`;
                logMessage("Error: " + response.message);
                return response;
            }
        }
        // ID Empleado es obligatorio si Responsable está seleccionado
        if (responsableName && !idEmpleado) {
            response.message = `No se pudo obtener el ID de Empleado para el responsable '${responsableName}'.`;
            logMessage("Error: " + response.message);
            return response;
        }


        let idResguardo = null; // Default a NULL
        let estadoFinalDB = formData.estado; // Estado que se guardará en la DB/Hoja por defecto
        let fechaReasignacionDB = null; // Default a NULL

        // Campos que se envían como NULL por defecto o condicionalmente
        const idAutoriza = null; 
        const documentacion = null;
        let comentariosFinal = formData.comentarios || null;

        // --- Lógica Condicional basada en el Estado del Equipo ---
        switch (formData.estado) {
            case 'Baja':
                estadoFinalDB = "Validación"; // Se guarda como "Validación" en BD/Hoja
                const nombreSolicitanteBaja = formData.nombreSolicitanteBaja || '';
                const razonBaja = formData.razonBaja || '';

                if (!nombreSolicitanteBaja || !razonBaja) {
                    response.message = "Para la Baja, el nombre del solicitante y la razón son obligatorios.";
                    logMessage("Error: " + response.message);
                    return response;
                }

                // Obtener el ID de Equipo_Usado (se hace aquí para poder pasarlo al correo)
                idEquipo = obtenerUltimoIDEquipoSQL(conn); 

                // Enviar correo de validación de Baja
                const bajaSubject = `VALIDACIÓN DE BAJA DE EQUIPO CELULAR DE LA SUCURSAL '${sucursalCodigo3Letras}'`;
                const bajaBody = `Buen día Jorge Fernández.\n\n` +
                                 `${nombreSolicitanteBaja} con correo electrónico ${solicitanteEmail} está solicitando que se valide la baja del siguiente equipo celular:\n` +
                                 `Marca: ${formData.marca}\n` +
                                 `Modelo: ${formData.modelo}\n` +
                                 `IMEI: ${imei}\n` +
                                 `Fecha de compra: ${formData.fechaCompra}\n` +
                                 `El costo del equipo es: $${costoEquipo.toFixed(2)}\n\n` +
                                 `La razón de la baja es la siguiente: ${razonBaja}\n\n` +
                                 `Si aprueba la baja del equipo celular, seleccione la opción de "Aceptar" de lo contrario seleccione la opción de "Denegar" y póngase en contacto con la persona que realizó la solicitud para auditar la baja del equipo.`;
                
                const bajaButtons = [
                    { text: 'Aceptar', action: 'aprobarBajaEquipo', params: { idEquipo: idEquipo, solicitanteEmail: solicitanteEmail, razonBaja: razonBaja, sucursal: sucursalCodigo3Letras, imei: imei }, color: '#28a745' }, // Green
                    { text: 'Denegar', action: 'denegarBajaEquipo', params: { idEquipo: idEquipo, solicitanteEmail: solicitanteEmail, sucursal: sucursalCodigo3Letras, imei: imei }, color: '#dc3545' } // Red
                ];
                sendEmailWithButtons(ADMIN_EMAILS, bajaSubject, bajaBody, bajaButtons);
                response.message = `Solicitud de baja enviada para validación. ID Equipo: ${idEquipo}.`;
                break;

            case 'Stock':
            
            case 'Robado':
                // Numero_Telefono y Responsable se envían como NULL desde el formulario (disabled)
                // IDRESGUARDO se queda en NULL para estos casos (no se obtiene por sucursal)
                idResguardo = null; 
                if (formData.estado === 'Robado' && !comentariosFinal) {
                    comentariosFinal = "Favor de agregar anotaciones del Robo";
                }
                // Obtener el ID de Equipo_Usado para la inserción
                idEquipo = obtenerUltimoIDEquipoSQL(conn); 
                break;

            case 'Vendido':
                estadoFinalDB = "Validación"; // Se guarda como "Validación" en BD/Hoja
                const nombreSolicitanteVenta = formData.nombreSolicitanteVenta || '';
                const personaVende = formData.personaVende || '';

                if (!nombreSolicitanteVenta || !personaVende) {
                    response.message = "Para la Venta, el nombre del solicitante y la persona a vender son obligatorios.";
                    logMessage("Error: " + response.message);
                    return response;
                }

                // Obtener el ID de Equipo_Usado (se hace aquí para poder pasarlo al correo)
                idEquipo = obtenerUltimoIDEquipoSQL(conn); 

                // Enviar correo de validación de Venta
                const ventaSubject = `VALIDACIÓN DE VENTA DE EQUIPO CELULAR DE LA SUCURSAL '${sucursalCodigo3Letras}'`;
                const ventaBody = `Buen día Jorge Fernández.\n\n` +
                                 `${nombreSolicitanteVenta} con correo electrónico ${solicitanteEmail} está solicitando que se valide la venta del siguiente equipo celular:\n` +
                                 `Marca: ${formData.marca}\n` +
                                 `Modelo: ${formData.modelo}\n` +
                                 `IMEI: ${imei}\n` +
                                 `Fecha de compra: ${formData.fechaCompra}\n` +
                                 `El costo del equipo es: $${costoEquipo.toFixed(2)}\n\n` +
                                 `El equipo se propone vender a: ${personaVende}\n\n` +
                                 `Si aprueba la venta del equipo celular, seleccione la opción de "Aceptar", de lo contrario seleccione la opción de "Denegar".`;
                
                const ventaButtons = [
                    { text: 'Aceptar', action: 'aprobarVentaEquipoStep1', params: { idEquipo: idEquipo, solicitanteEmail: solicitanteEmail, personaVende: personaVende, sucursal: sucursalCodigo3Letras, imei: imei }, color: '#28a745' }, // Green
                    { text: 'Denegar', action: 'denegarVentaEquipo', params: { idEquipo: idEquipo, solicitanteEmail: solicitanteEmail, sucursal: sucursalCodigo3Letras, imei: imei }, color: '#dc3545' } // Red
                ];
                sendEmailWithButtons(ADMIN_EMAILS, ventaSubject, ventaBody, ventaButtons);
                response.message = `Solicitud de venta enviada para validación. ID Equipo: ${idEquipo}.`;
                break;

            case 'Reasignado':
                fechaReasignacionDB = formData.fechaReasignacion ? new Date(formData.fechaReasignacion) : null;
                if (!fechaReasignacionDB || isNaN(fechaReasignacionDB.getTime()) || fechaReasignacionDB > today) {
                    response.message = "Para Reasignación, la fecha de reasignación debe ser válida y no futura.";
                    logMessage("Error: " + response.message);
                    return response;
                }
                if (!responsableName) { // Responsable es obligatorio para Reasignado
                    response.message = "Para Reasignación, el campo Responsable es obligatorio.";
                    logMessage("Error: " + response.message);
                    return response;
                }
                // Si hay responsable, IDRESGUARDO es NULL, de lo contrario se busca (pero aquí responsable es obligatorio)
                idResguardo = null; 
                // Obtener el ID de Equipo_Usado para la inserción
                idEquipo = obtenerUltimoIDEquipoSQL(conn); 
                break;

            default: // Activo, En Mantenimiento, Desecho (y cualquier otro estado no manejado específicamente)
                // Si Responsable está vacío, obtener Resguardo por Sucursal
                if (!responsableName) { 
                    idResguardo = getResguardoPorSucursalSQL(idSucursalBD, conn);
                    if (!idResguardo) {
                        response.message = `No se pudo obtener el ID de Resguardo para la sucursal '${formData.idSucursal}'.`;
                        logMessage("Error: " + response.message);
                        return response;
                    }
                }
                // Obtener el ID de Equipo_Usado para la inserción
                idEquipo = obtenerUltimoIDEquipoSQL(conn); 
                break;
        }
        
        // Si idEquipo aún es null, significa que no se generó en los casos condicionales (ej. si el estado no requiere email)
        // Esto asegura que idEquipo siempre tenga un valor antes de intentar insertarlo.
        if (idEquipo === null) {
            idEquipo = obtenerUltimoIDEquipoSQL(conn);
        }

        // --- INSERCIÓN DIRECTA EN SQL SERVER ---
        const insertQuery = `
            INSERT INTO Equipo_Usado (
                ID_Equipo, Costo_Equipo, Fecha_Compra, Fecha_Recoleccion, Fecha_Reasignacion,
                Fecha_Formulario, Estado, Observaciones, Marca, Modelo, RAM, ROM, IMEI,
                Numero_Telefono, IDSUCURSAL, IDEMPLEADO, Responsable, IDRESGUARDO, IDAUTORIZA,
                Comentarios, Documentacion
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        `;
        let pstmt = null;

        try {
            pstmt = conn.prepareStatement(insertQuery);
            
            pstmt.setInt(1, idEquipo);
            pstmt.setObject(2, costoEquipo); // MONEY
            pstmt.setString(3, formatDateForSql(fechaCompra));
            pstmt.setString(4, formatDateForSql(fechaRecoleccion));
            pstmt.setString(5, formatDateForSql(fechaReasignacionDB)); // Fecha_Reasignacion (puede ser NULL o una fecha)
            pstmt.setString(6, formatDateForSql(fechaFormulario));
            pstmt.setString(7, estadoFinalDB || null); // Estado final para la DB
            pstmt.setString(8, formData.observaciones || null);
            pstmt.setString(9, formData.marca || null);
            pstmt.setString(10, formData.modelo || null);
            pstmt.setString(11, formData.ram || null);
            pstmt.setString(12, formData.rom || null);
            pstmt.setString(13, imei);
            pstmt.setString(14, numeroTelefono); // Puede ser NULL
            pstmt.setObject(15, idSucursalBD); // IDSUCURSAL (INT)
            pstmt.setString(16, idEmpleado); // IDEMPLEADO (NVARCHAR)
            pstmt.setString(17, responsableName); // Responsable (NVARCHAR)
            pstmt.setString(18, idResguardo); // IDRESGUARDO (NVARCHAR)
            pstmt.setString(19, idAutoriza); // IDAUTORIZA (NVARCHAR)
            pstmt.setString(20, comentariosFinal); // Comentarios (NVARCHAR)
            pstmt.setString(21, documentacion); // Documentacion (NVARCHAR)

            pstmt.executeUpdate();
            logMessage(`Equipo con ID ${idEquipo} insertado exitosamente en SQL Server.`);
            
            // Si la inserción en SQL es exitosa, la respuesta ya es success, solo se actualiza el mensaje si no es validación
            if (estadoFinalDB !== "Validación") {
                response.message = `Equipo con ID ${idEquipo} insertado en SQL y registrado en hoja de cálculo.`;
            }

        } catch (sqlError) {
            logMessage("Error al insertar en SQL Server (RECU): " + sqlError.message + " Stack: " + sqlError.stack);
            response.message = "Error al guardar en la base de datos (RECU): " + sqlError.message;
            response.success = false;
            return response;
        } finally {
            if (pstmt) {
                try { pstmt.close(); } catch (e) { logMessage("Error al cerrar pstmt en procesarRECUFormulario: " + e.message); }
            }
        }

        // --- REGISTRO EN GOOGLE SHEET (SÓLO SI LA INSERCIÓN EN SQL FUE EXITOSA) ---
        const rowData = [];
        RECU_SHEET_HEADERS.forEach(header => {
            switch (header) {
                case "Marca temporal": rowData.push(fechaFormulario); break;
                case "Dirección de correo electrónico": rowData.push(solicitanteEmail); break;
                case "Costo del Equipo": rowData.push(costoEquipo || ''); break;
                case "Fecha de compra de Equipo": rowData.push(fechaCompra || ''); break;
                case "Fecha de Recolección": rowData.push(fechaRecoleccion || ''); break;
                case "Fecha de Reasignacion": rowData.push(fechaReasignacionDB || ''); break; // Columna F
                case "Estado del equipo": rowData.push(estadoFinalDB || ''); break; // Columna G
                case "Observaciones": rowData.push(formData.observaciones || ''); break; // Columna H
                case "Marca": rowData.push(formData.marca || ''); break; // Columna I
                case "Modelo": rowData.push(formData.modelo || ''); break; // Columna J
                case "Memoria RAM": rowData.push(formData.ram || ''); break; // Columna K
                case "Almacenamiento (Memoria ROM)": rowData.push(formData.rom || ''); break; // Columna L
                case "IMEI": rowData.push(imei); break; // Columna M
                case "Numero_Telefono": rowData.push(numeroTelefono || ''); break; // Columna N
                case "Sucursal": rowData.push(formData.idSucursal || ''); break; // Columna O
                case "IDEquipo": rowData.push(idEquipo); break; // Columna P
                case "Error": rowData.push(""); break; // Columna Q
                case "IDSucursal": rowData.push(idSucursalBD || ''); break; // Columna R
                case "EJECUTADO": rowData.push("SI"); break; // Columna S (MOVIDO)
                case "Comentarios": rowData.push(comentariosFinal || ''); break; // Columna T (MOVIDO)
                case "IDEMPLEADO": rowData.push(idEmpleado || ''); break; // Columna U (REUBICADO)
                case "Responsable": rowData.push(responsableName || ''); break; // Columna V (REUBICADO)
                case "IDRESGUARDO": rowData.push(idResguardo || ''); break; // Columna W (REUBICADO)
                case "IDAUTORIZA": rowData.push(idAutoriza || ''); break; // Columna X (REUBICADO)
                case "Documentacion": rowData.push(documentacion || ''); break; // Columna Y (REUBICADO)
                default: rowData.push(""); // Para cualquier otra columna no mapeada explícitamente
            }
        });

        sheet.appendRow(rowData);
        logMessage(`Equipo con ID ${idEquipo} registrado en hoja 'RECU' después de SQL.`);

    } catch (e) {
        response.message = "Error general en el procesamiento del formulario RECU: " + e.message;
        logMessage("Error en procesarRECUFormulario (general): " + e.message + " Stack: " + e.stack);
        response.success = false;
    } finally {
        if (conn) {
            try { conn.close(); } catch (e) { logMessage("Error al cerrar conexión en procesarRECUFormulario finally block: " + e.message); }
        }
    }
    return response;
}

// --- 4.2. Formulario: Dar de Alta Línea y Equipo (ALyE) ---
/**
 * Procesa los datos enviados desde el formulario "Dar de alta Línea y Equipo".
 * Inserta los datos en las tablas Telefonía_Telcel y Equipo_Nuevo en SQL Server,
 * luego los registra en la hoja "ALyE".
 * @param {Object} formData Los datos del formulario como un objeto JavaScript.
 * @returns {Object} Un objeto con 'success' (boolean) y 'message' (string).
 */
function procesarALyEFormulario(formData) {
    let response = { success: false, message: "" };
    logMessage("Datos recibidos del formulario ALyE: " + JSON.stringify(formData));

    let conn = null;
    let pstmtTelcel = null;
    let pstmtEquipoNuevo = null;
    let idEquipoNuevo = null; // Declarar aquí e inicializar a null
    let idEmpleadoInt = null;
    const estadoEquipoNuevo = 'Asignado';
    const imeiEquipoNuevo = formData.imei_equipo_nuevo ? formData.imei_equipo_nuevo.trim() : '';
    const idSucursalInt = SUCURSAL_MAP[formData.idsucursal] || null;

    try {
        const sheet = getSheet("ALyE"); // Hoja específica para ALyE
        conn = getJdbcConnection();

        const fechaFormulario = new Date(); // Fecha de registro del formulario

        // --- Validaciones y Mapeos Comunes ---
        const telefono = formData.telefono ? formData.telefono.trim() : '';
        if (!telefono) {
            response.message = "El campo Teléfono es obligatorio.";
            logMessage("Error: " + response.message);
            return response;
        }
        if (!validarTelefonoUnicoSQL(telefono, conn)) {
            response.message = `El Teléfono '${telefono}' ya existe en la base de datos.`;
            logMessage("Error: " + response.message);
            return response;
        }

        
        if (!imeiEquipoNuevo) {
            response.message = "El campo IMEI del Equipo Nuevo es obligatorio.";
            logMessage("Error: " + response.message);
            return response;
        }
        if (!validarIMEINuevoUnicoSQL(imeiEquipoNuevo, conn)) {
            response.message = `El IMEI '${imeiEquipoNuevo}' ya existe en la base de datos (Equipo Nuevo o Telefonía Telcel).`;
            logMessage("Error: " + response.message);
            return response;
        }

        
        if (idSucursalInt === null) {
            response.message = "ID de Sucursal inválido.";
            logMessage("Error: " + response.message);
            return response;
        }

        
        if (formData.responsable) {
            idEmpleadoInt = getResponsableID(formData.responsable);
            if (!idEmpleadoInt) {
                response.message = `No se encontró ID de Empleado para el responsable '${formData.responsable}'.`;
                logMessage("Error: " + response.message);
                return response;
            }
        }


        // --- Generar ID para Equipo_Nuevo ---
        idEquipoNuevo = obtenerUltimoIDEquipoNuevoSQL(conn); // Asignación del valor

        // --- INSERCIÓN EN Equipo_Nuevo ---
        const insertEquipoNuevoQuery = `
            INSERT INTO Equipo_Nuevo (
                ID_Equipo, Costo_Equipo, Fecha_Compra, Estado, Observaciones,
                Marca, Modelo, RAM, ROM, IMEI, IDEMPLEADO, Responsable, IDSUCURSAL
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        `;
        try {
            pstmtEquipoNuevo = conn.prepareStatement(insertEquipoNuevoQuery);
            
            // Campos de Equipo_Nuevo no presentes en este formulario, se envían como NULL o valores por defecto si la DB lo permite
            const fechaCompraEquipoNuevo = formatDateForSql(formData.fecha_inicio); 

            pstmtEquipoNuevo.setInt(1, idEquipoNuevo);
            pstmtEquipoNuevo.setObject(2, parseFloat(formData.equipo_ilimitado) || null); 
            pstmtEquipoNuevo.setString(3, fechaCompraEquipoNuevo); // DATETIME2, puede ser NULL
            pstmtEquipoNuevo.setString(4, estadoEquipoNuevo); // NVARCHAR(50), puede ser NULL
            pstmtEquipoNuevo.setString(5, formData.observaciones_equipo_nuevo || null);
            pstmtEquipoNuevo.setString(6, formData.marca_nuevo || null);
            pstmtEquipoNuevo.setString(7, formData.modelo_nuevo || null);
            pstmtEquipoNuevo.setString(8, formData.ram_nuevo || null);
            pstmtEquipoNuevo.setString(9, formData.rom_nuevo || null);
            pstmtEquipoNuevo.setString(10, imeiEquipoNuevo);
            pstmtEquipoNuevo.setObject(11, idEmpleadoInt); // IDEMPLEADO de Equipo_Nuevo (asumo INT)
            pstmtEquipoNuevo.setString(12, formData.responsable || null); // Responsable de Equipo_Nuevo
            pstmtEquipoNuevo.setObject(13, idSucursalInt); // IDSUCURSAL de Equipo_Nuevo (asumo INT)

            pstmtEquipoNuevo.executeUpdate();
            logMessage(`Equipo_Nuevo con ID ${idEquipoNuevo} insertado exitosamente.`);

        } catch (sqlError) {
            logMessage("Error al insertar en Equipo_Nuevo: " + sqlError.message + " Stack: " + sqlError.stack);
            response.message = "Error al guardar el equipo nuevo: " + sqlError.message;
            response.success = false;
            return response;
        } finally {
            if (pstmtEquipoNuevo) pstmtEquipoNuevo.close();
        }

        // --- INSERCIÓN EN Telefonía_Telcel ---
        // id_tel es IDENTITY(1,1), no se incluye en el INSERT
        const insertTelcelQuery = `
            INSERT INTO Telefonía_Telcel (
                Región, Cuenta_padre, Cuenta, Teléfono, Clave_plan, Nombre_plan, Minutos, Mensajes,
                Monto_renta, Equipo_ilimitado, Servicio_a_la_carta, Servicio_blackberry, Duracion_plan,
                Fecha_inicio, Fecha_termino, Estatus_adendum, Meses_restantes, Marca, Modelo, IMEI,
                SIM, Tipo, Responsable, Notas, IDEMPLEADO, IDSUCURSAL, Datos, Extensión,
                IDEquipoUSado, IDEquipoNuevo
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        `;
        try {
            pstmtTelcel = conn.prepareStatement(insertTelcelQuery);

            pstmtTelcel.setString(1, formData.region || null);
            pstmtTelcel.setString(2, formData.cuenta_padre || null);
            pstmtTelcel.setString(3, formData.cuenta || null);
            pstmtTelcel.setString(4, telefono);
            pstmtTelcel.setString(5, formData.clave_plan || null);
            pstmtTelcel.setString(6, formData.nombre_plan || null);
            pstmtTelcel.setString(7, formData.minutos || null);
            pstmtTelcel.setString(8, formData.mensajes || null);
            pstmtTelcel.setObject(9, parseFloat(formData.monto_renta) || null); // MONEY
            pstmtTelcel.setObject(10, parseFloat(formData.equipo_ilimitado) || null); // MONEY
            pstmtTelcel.setObject(11, parseFloat(formData.servicio_a_la_carta) || null); // MONEY
            pstmtTelcel.setObject(12, parseFloat(formData.servicio_blackberry) || null); // MONEY
            pstmtTelcel.setString(13, formData.duracion_plan || null);
            pstmtTelcel.setString(14, formatDateForSql(formData.fecha_inicio));
            pstmtTelcel.setString(15, formatDateForSql(formData.fecha_termino));
            pstmtTelcel.setString(16, formData.estatus_adendum || null);
            pstmtTelcel.setString(17, formData.meses_restantes || null);
            pstmtTelcel.setString(18, formData.marca_nuevo); // Marca del equipo nuevo
            pstmtTelcel.setString(19, formData.modelo_nuevo); // Modelo del equipo nuevo
            pstmtTelcel.setString(20, imeiEquipoNuevo); // IMEI del equipo nuevo
            pstmtTelcel.setString(21, formData.sim || null);
            pstmtTelcel.setString(22, formData.tipo || null);
            pstmtTelcel.setString(23, formData.responsable || null); // Responsable de Telefonía_Telcel
            pstmtTelcel.setString(24, formData.notas || null);
            pstmtTelcel.setObject(25, idEmpleadoInt); // IDEMPLEADO de Telefonía_Telcel (asumo INT)
            pstmtTelcel.setObject(26, idSucursalInt); // IDSUCURSAL de Telefonía_Telcel (asumo INT)
            pstmtTelcel.setObject(27, parseInt(formData.datos) || null); // Datos (GB) (asumo INT)
            pstmtTelcel.setString(28, formData.extension || null);
            pstmtTelcel.setString(29, null); // IDEquipoUSado (omitido en este formulario)
            pstmtTelcel.setObject(30, idEquipoNuevo); // IDEquipoNuevo (vinculado al ID generado)

            pstmtTelcel.executeUpdate();
            logMessage(`Telefonía_Telcel para Teléfono ${telefono} insertado exitosamente.`);
            
            response.success = true;
            response.message = `Línea y Equipo registrados exitosamente. ID Equipo Nuevo: ${idEquipoNuevo}.`;

        } catch (sqlError) {
            logMessage("Error al insertar en Telefonía_Telcel: " + sqlError.message + " Stack: " + sqlError.stack);
            response.message = "Error al guardar la línea telefónica: " + sqlError.message;
            response.success = false;
            // Si falla la segunda inserción, podrías considerar hacer un rollback de la primera si la DB lo permite.
            // Para Apps Script JDBC, esto requeriría manejar transacciones manualmente.
            return response;
        } finally {
            if (pstmtTelcel) pstmtTelcel.close();
        }

        // --- REGISTRO EN GOOGLE SHEET (SÓLO SI LAS INSERCIONES EN SQL FUERON EXITOSAS) ---
        const rowData = [];
        ALyE_SHEET_HEADERS.forEach(header => {
            switch (header) {
                case "Marca temporal": rowData.push(fechaFormulario); break;
                case "Dirección de correo electrónico": rowData.push(Session.getActiveUser().getEmail()); break;
                case "Id_tel": rowData.push("Generado por DB"); break; // id_tel es IDENTITY
                case "Región": rowData.push(formData.region || ''); break;
                case "Cuenta_padre": rowData.push(formData.cuenta_padre || ''); break;
                case "Cuenta": rowData.push(formData.cuenta || ''); break;
                case "Teléfono": rowData.push(telefono); break;
                case "Clave_plan": rowData.push(formData.clave_plan || ''); break;
                case "Nombre_plan": rowData.push(formData.nombre_plan || ''); break;
                case "Minutos": rowData.push(formData.minutos || ''); break;
                case "Mensajes": rowData.push(formData.mensajes || ''); break;
                case "Monto_renta": rowData.push(parseFloat(formData.monto_renta) || ''); break;
                case "Equipo_ilimitado": rowData.push(parseFloat(formData.equipo_ilimitado) || ''); break;
                case "Duracion_plan": rowData.push(formData.duracion_plan || ''); break;
                case "Fecha_inicio": rowData.push(formData.fecha_inicio ? new Date(formData.fecha_inicio) : ''); break;
                case "Fecha_termino": rowData.push(formData.fecha_termino ? new Date(formData.fecha_termino) : ''); break;
                case "Marca_linea": rowData.push(formData.marca_nuevo || ''); break;
                case "Modelo_linea": rowData.push(formData.modelo_nuevo || ''); break;
                case "IMEI_linea": rowData.push(formData.imeiEquipoNuevo || ''); break;
                case "SIM": rowData.push(formData.sim || ''); break;
                case "Tipo": rowData.push(formData.tipo || ''); break;
                case "Responsable_Linea": rowData.push(formData.responsable || ''); break;
                case "Notas": rowData.push(formData.notas || '')
                case "IDEMPLEADO_Linea": rowData.push(idEmpleadoInt || ''); break;
                case "Sucursal_Linea" : rowData.push (formData.idsucursal || ''); break;
                case "IDSUCURSAL_Telcel": rowData.push(idSucursalInt || ''); break;
                case "Datos": rowData.push(parseInt(formData.datos) || ''); break;
                case "Extensión": rowData.push(formData.extension || ''); break;
                case "ID_Equipo_Nuevo": rowData.push(idEquipoNuevo); break;
                case "Error_Linea": rowData.push(""); break;
                case "EJECUTADO_Linea": rowData.push("SI"); break;

                case "ID_Equipo": rowData.push(idEquipoNuevo); break;
                case "Costo_Equipo": rowData.push(parseFloat(formData.equipo_ilimitado) || ''); break;
                case "Fecha_Compra_Equipo": rowData.push(formData.fecha_inicio ? new Date(formData.fecha_inicio) : ''); break;
                case "Estado_Equipo": rowData.push(estadoEquipoNuevo || '');break;
                case "Observaciones_Equipo_Nuevo": rowData.push(formData.observaciones_equipo_nuevo || ''); break;
                case "Marca_Equipo_Nuevo": rowData.push(formData.marca_nuevo || ''); break;
                case "Modelo_Equipo_Nuevo": rowData.push(formData.modelo_nuevo || ''); break;
                case "RAM_Equipo_Nuevo": rowData.push(formData.ram_nuevo || ''); break;
                case "ROM_Equipo_Nuevo": rowData.push(formData.rom_nuevo || ''); break;
                case "IMEI_Equipo_Nuevo": rowData.push(imeiEquipoNuevo); break;
                case "IDEMPLEADO_Equipo": rowData.push(idEmpleadoInt || ''); break;
                case "Responsable_Equipo": rowData.push(formData.responsable || ''); break;
                case "Sucursal_Equipo" : rowData.push (formData.idsucursal || ''); break;
                case "IDSUCURSAL_Equipo": rowData.push(idSucursalInt || ''); break;
                case "Error_Equipo": rowData.push(""); break;
                case "EJECUTADO_Equipo": rowData.push("SI"); break;
                default: rowData.push(""); 
            }
        });

        sheet.appendRow(rowData);
        logMessage(`Datos de ALyE registrados en hoja 'ALyE' después de SQL.`);

    } catch (e) {
        response.message = "Error general en el procesamiento del formulario ALyE: " + e.message;
        logMessage("Error en procesALyEFormulario (general): " + e.message + " Stack: " + e.stack);
        response.success = false;
    } finally {
        if (conn) {
            try { conn.close(); } catch (e) { logMessage("Error al cerrar conexión en procesALyEFormulario finally block: " + e.message); }
        }
    }
    return response;
}


// --- 4.3. Formulario: Renovación de Línea y Equipo (RLyE) ---
/**
 * Placeholder para procesar el formulario "Renovación de Línea y Equipo".
 * DEBERÁS IMPLEMENTAR LA LÓGICA ESPECÍFICA PARA ESTE FORMULARIO, INCLUYENDO INTERACCIÓN CON SQL SI ES NECESARIO.
 * Asegúrate de usar `getSheet("RLyE")` y `RLyE_SHEET_HEADERS`.
 */
function procesarRLyEFormulario(formData) {
    let response = { success: false, message: "" };
    logMessage("Datos recibidos del formulario RLyE: " + JSON.stringify(formData));
    try {
        const sheet = getSheet("RLyE"); // Hoja específica para RLyE
        // Implementa aquí la lógica para guardar en SQL y luego en la hoja "RLyE"
        // Ejemplo simplificado para demostración:
        const rowData = [];
        RLyE_SHEET_HEADERS.forEach(header => {
            switch (header) {
                case "Marca temporal": rowData.push(new Date()); break;
                case "EJECUTADO": rowData.push("SI"); break;
                default: rowData.push(`Valor para ${header}: ${formData[header] || ''}`); // Ajusta según tus campos reales
            }
        });
        sheet.appendRow(rowData);
        response.success = true;
        response.message = "Datos de Renovación de Línea y Equipo procesados (simulado).";
        logMessage(response.message);
    } catch (e) {
        response.message = "Error al procesar formulario RLyE: " + e.message;
        logMessage("Error en procesarRLyEFormulario: " + e.message);
    }
    return response;
}

// --- 4.4. Formulario: Modificar Línea (ML) ---
/**
 * Placeholder para procesar el formulario "Modificar Línea".
 * DEBERÁS IMPLEMENTAR LA LÓGICA ESPECÍFICA PARA ESTE FORMULARIO, INCLUYENDO INTERACCIÓN CON SQL SI ES NECESARIO.
 * Asegúrate de usar `getSheet("ML")` y `ML_SHEET_HEADERS`.
 */
function procesarMLFormulario(formData) {
    let response = { success: false, message: "" };
    logMessage("Datos recibidos del formulario ML: " + JSON.stringify(formData));
    try {
        const sheet = getSheet("ML"); // Hoja específica para ML
        // Implementa aquí la lógica para guardar en SQL y luego en la hoja "ML"
        // Ejemplo simplificado para demostración:
        const rowData = [];
        ML_SHEET_HEADERS.forEach(header => {
            switch (header) {
                case "Marca temporal": rowData.push(new Date()); break;
                case "EJECUTADO": rowData.push("SI"); break;
                default: rowData.push(`Valor para ${header}: ${formData[header] || ''}`); // Ajusta según tus campos reales
            }
        });
        sheet.appendRow(rowData);
        response.success = true;
        response.message = "Datos de Modificar Línea procesados (simulado).";
        logMessage(response.message);
    } catch (e) {
        response.message = "Error al procesar formulario ML: " + e.message;
        logMessage("Error en procesarMLFormulario: " + e.message);
    }
    return response;
}

// --- 4.5. Formulario: Modificar Equipo Usado (MEU) ---
/**
 * Placeholder para procesar el formulario "Modificar Equipo Usado".
 * DEBERÁS IMPLEMENTAR LA LÓGICA ESPECÍFICA PARA ESTE FORMULARIO, INCLUYENDO INTERACCIÓN CON SQL SI ES NECESARIO.
 * Asegúrate de usar `getSheet("MEU")` y `MEU_SHEET_HEADERS`.
 */
function procesarMEUFormulario(formData) {
    let response = { success: false, message: "" };
    logMessage("Datos recibidos del formulario MEU: " + JSON.stringify(formData));
    try {
        const sheet = getSheet("MEU"); // Hoja específica para MEU
        // Implementa aquí la lógica para guardar en SQL y luego en la hoja "MEU"
        // Ejemplo simplificado para demostración:
        const rowData = [];
        MEU_SHEET_HEADERS.forEach(header => {
            switch (header) {
                case "Marca temporal": rowData.push(new Date()); break;
                case "EJECUTADO": rowData.push("SI"); break;
                default: rowData.push(`Valor para ${header}: ${formData[header] || ''}`); // Ajusta según tus campos reales
            }
        });
        sheet.appendRow(rowData);
        response.success = true;
        response.message = "Datos de Modificar Equipo Usado procesados (simulado).";
        logMessage(response.message);
    } catch (e) {
        response.message = "Error al procesar formulario MEU: " + e.message;
        logMessage("Error en procesarMEUFormulario: " + e.message);
    }
    return response;
}


// ====================================================================================================================
// ========================= 5. FUNCIONES DE PROCESAMIENTO SECUNDARIO (EJ. TRIGGERS) ================================
// ====================================================================================================================

/**
 * Función principal para leer filas de la hoja de cálculo
 * y enviarlas a la base de datos SQL Server.
 *
 * NOTA: Con la nueva lógica, las nuevas entradas del formulario se insertan directamente en SQL.
 * Esta función es ahora principalmente para:
 * 1. Reprocesar filas que no fueron insertadas correctamente al principio (marcadas con 'ERROR').
 * 2. Procesar manualmente filas añadidas o modificadas directamente en la hoja (marcadas con 'NO').
 * 3. Mantener un trigger de respaldo si se desea una doble verificación.
 *
 * Esta es la función que se configuraría con un trigger (ej. de tiempo o al enviar formulario de Google).
 * Actualmente, esta función solo procesa la hoja "RECU".
 */
function enviarHojaASQL() {
    logMessage("Iniciando 'enviarHojaASQL'. (Rol Secundario)");
    let conn = null;
    try {
        conn = getJdbcConnection(); // Obtener la conexión al inicio

        const sheet = getSheet("RECU"); // Procesar específicamente la hoja "RECU"
        
        const range = sheet.getDataRange();
        const values = range.getValues();

        if (values.length < 2) {
            logMessage("No hay datos nuevos en la hoja 'RECU' para enviar a SQL.");
            return;
        }

        const headerRow = values[0];
        const getColIndex = (headerName) => {
            const index = headerRow.indexOf(headerName);
            if (index === -1) {
                logMessage(`Advertencia: Columna '${headerName}' no encontrada en los encabezados de la hoja 'RECU'.`);
            }
            return index;
        };

        const colMap = {
            marcaTemporal: getColIndex("Marca temporal"),
            direccionEmail: getColIndex("Dirección de correo electrónico"),
            costoEquipo: getColIndex("Costo del Equipo"),
            fechaCompra: getColIndex("Fecha de compra de Equipo"),
            fechaRecoleccion: getColIndex("Fecha de Recolección"),
            fechaReasignacion: getColIndex("Fecha de Reasignacion"), // Columna F
            estado: getColIndex("Estado del equipo"),
            observaciones: getColIndex("Observaciones"),
            marca: getColIndex("Marca"),
            modelo: getColIndex("Modelo"),
            ram: getColIndex("Memoria RAM"),
            rom: getColIndex("Almacenamiento (Memoria ROM)"),
            imei: getColIndex("IMEI"),
            numeroTelefono: getColIndex("Numero_Telefono"), 
            sucursalNombre: getColIndex("Sucursal"),
            idEquipo: getColIndex("IDEquipo"),
            errorCol: getColIndex("Error"),
            idSucursalBD: getColIndex("IDSucursal"),
            idEmpleado: getColIndex("IDEMPLEADO"), 
            responsable: getColIndex("Responsable"), 
            idResguardo: getColIndex("IDRESGUARDO"), 
            idAutoriza: getColIndex("IDAUTORIZA"), 
            comentarios: getColIndex("Comentarios"),
            documentacion: getColIndex("Documentacion"), 
            ejecutadoCol: getColIndex("EJECUTADO")
        };
        
        // Preparar la consulta INSERT para SQL Server
        const insertQuery = `
            INSERT INTO Equipo_Usado (
                ID_Equipo, Costo_Equipo, Fecha_Compra, Fecha_Recoleccion, Fecha_Reasignacion,
                Fecha_Formulario, Estado, Observaciones, Marca, Modelo, RAM, ROM, IMEI,
                Numero_Telefono, IDSUCURSAL, IDEMPLEADO, Responsable, IDRESGUARDO, IDAUTORIZA,
                Comentarios, Documentacion
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        `;

        let pstmt = null;
        
        // Iterar sobre todas las filas de datos (excepto el encabezado)
        for (let i = 1; i < values.length; i++) { // Empezar desde la segunda fila (índice 1)
            let row = values[i];
            let sheetRowNumber = i + 1; // Número de fila real en la hoja (1-based)
            let rowError = "";
            let status = "NO"; // Default status

            // Solo procesar si no ha sido ejecutado (estado "NO" o "ERROR")
            if (row[colMap.ejecutadoCol] === "SI") {
                continue; // Saltar filas ya procesadas exitosamente
            }

            try {
                // Verificar si los índices de columna son válidos antes de acceder a ellos
                if (colMap.idEquipo === -1 || colMap.costoEquipo === -1 || colMap.imei === -1 || colMap.idSucursalBD === -1) {
                    throw new Error("Una o más columnas requeridas no se encontraron en la hoja de cálculo. Por favor, verifique los encabezados.");
                }

                const idEquipo = parseInt(row[colMap.idEquipo]);
                const costoEquipo = parseFloat(row[colMap.costoEquipo]);

                // Función auxiliar para formatear fechas
                const formatDateForSql = (dateValue) => {
                    if (!dateValue || !(dateValue instanceof Date)) return null;
                    return Utilities.formatDate(dateValue, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss.SSS");
                };

                const fechaCompra = formatDateForSql(row[colMap.fechaCompra]);
                const fechaRecoleccion = formatDateForSql(row[colMap.fechaRecoleccion]);
                const fechaReasignacion = formatDateForSql(row[colMap.fechaReasignacion] || null); // Puede ser NULL
                const fechaFormulario = formatDateForSql(row[colMap.marcaTemporal]);

                const estado = row[colMap.estado] || '';
                const observaciones = row[colMap.observaciones] || '';
                const marca = row[colMap.marca] || '';
                const modelo = row[colMap.modelo] || '';
                const ram = row[colMap.ram] || '';
                const rom = row[colMap.rom] || '';
                const imei = row[colMap.imei] || '';
                const numeroTelefono = row[colMap.numeroTelefono] || null; 
                const idSucursal = SUCURSAL_MAP[row[colMap.sucursalNombre]] || null; // Mapear el nombre de sucursal a su INT
                const idEmpleado = row[colMap.idEmpleado] || null; 
                const responsable = row[colMap.responsable] || null; 
                const idResguardo = row[colMap.idResguardo] || null; 
                const idAutoriza = row[colMap.idAutoriza] || null; 
                const comentarios = row[colMap.comentarios] || ''; 
                const documentacion = row[colMap.documentacion] || null; 


                pstmt = conn.prepareStatement(insertQuery);
                pstmt.setInt(1, idEquipo);
                pstmt.setObject(2, costoEquipo);
                pstmt.setString(3, fechaCompra);
                pstmt.setString(4, fechaRecoleccion);
                pstmt.setString(5, fechaReasignacion);
                pstmt.setString(6, fechaFormulario);
                pstmt.setString(7, estado);
                pstmt.setString(8, observaciones);
                pstmt.setString(9, marca);
                pstmt.setString(10, modelo);
                pstmt.setString(11, ram);
                pstmt.setString(12, rom);
                pstmt.setString(13, imei);
                pstmt.setString(14, numeroTelefono);
                pstmt.setObject(15, idSucursal);
                pstmt.setString(16, idEmpleado);
                pstmt.setString(17, responsable);
                pstmt.setString(18, idResguardo);
                pstmt.setString(19, idAutoriza);
                pstmt.setString(20, comentarios);
                pstmt.setString(21, documentacion);

                pstmt.executeUpdate(); // Ejecutar la inserción para la fila actual
                status = "SI";
                logMessage(`Fila ${sheetRowNumber} (ID_Equipo: ${idEquipo}) insertada exitosamente en SQL.`);

            } catch (e) {
                rowError = `Error al insertar fila ${sheetRowNumber} en SQL: ${e.message}`;
                logMessage(rowError);
                status = "ERROR";
            } finally {
                if (pstmt) {
                    try { pstmt.close(); } catch (e) { logMessage("Error al cerrar pstmt: " + e.message); }
                }
                // Actualizar el estado y el mensaje de error en la hoja de cálculo
                // sheetRowNumber es 1-based, y colMap.ejecutadoCol es 0-based, por eso +1
                if (colMap.ejecutadoCol !== -1) {
                    sheet.getRange(sheetRowNumber, colMap.ejecutadoCol + 1).setValue(status);
                }
                if (colMap.errorCol !== -1) {
                    sheet.getRange(sheetRowNumber, colMap.errorCol + 1).setValue(rowError);
                }
            }
        }

    } catch (e) {
        logMessage("Error general en 'enviarHojaASQL': " + e.message + " Stack: " + e.stack);
    } finally {
        if (conn) {
            try { conn.close(); } catch (e) { logMessage("Error al cerrar conexión: " + e.message); }
        }
    }
    logMessage("Finalizando 'enviarHojaASQL'.");
}

// ====================================================================================================================
// ====================================== 6. FUNCIONES DE ACCIÓN DE CORREO (CALLBACKS) ==============================
// ====================================================================================================================

/**
 * Aprueba la baja de un equipo, actualizando su estado en BD y hoja, y notificando al solicitante.
 * Llamada desde el botón "Aceptar" en el correo de validación de baja.
 * @param {number} idEquipo El ID del equipo a aprobar (Apps Script generado ID).
 * @param {string} solicitanteEmail El correo del solicitante original.
 * @param {string} razonBaja La razón de la baja proporcionada.
 * @param {string} sucursal La sucursal del equipo.
 * @param {string} imei El IMEI del equipo (para la actualización en DB).
 */
function aprobarBajaEquipo(idEquipo, solicitanteEmail, razonBaja, sucursal, imei) {
    let conn = null;
    let confirmationTitle = "Baja Aprobada";
    let confirmationMessage = `El equipo con IMEI ${imei} ha sido marcado como "Baja" y se ha enviado una notificación al solicitante.`;
    let success = true;

    try {
        conn = getJdbcConnection();
        const sheet = getSheet("RECU");

        // 1. Actualizar en SQL Server usando IMEI
        const updateSql = `UPDATE Equipo_Usado SET Estado = ?, IDAUTORIZA = ?, Comentarios = ? WHERE IMEI = ?`;
        let pstmt = conn.prepareStatement(updateSql);
        pstmt.setString(1, "Baja");
        pstmt.setString(2, "645"); // IDAUTORIZA fijo
        pstmt.setString(3, `Baja aprobada. Razón: ${razonBaja}`);
        pstmt.setString(4, imei); // Usar IMEI para el WHERE
        pstmt.executeUpdate();
        pstmt.close();
        logMessage(`Baja aprobada en SQL para Equipo IMEI: ${imei}`);

        // 2. Actualizar en Google Sheet
        const range = sheet.getDataRange();
        const values = range.getValues();
        const headerRow = values[0];
        const imeiColIndex = headerRow.indexOf("IMEI"); // Obtener índice de IMEI
        const estadoColIndex = headerRow.indexOf("Estado del equipo");
        const idAutorizaColIndex = headerRow.indexOf("IDAUTORIZA");
        const comentariosColIndex = headerRow.indexOf("Comentarios");
        const ejecutadoColIndex = headerRow.indexOf("EJECUTADO");

        for (let i = 1; i < values.length; i++) {
            // Buscar por IMEI y por estado "Validación"
            if (values[i][imeiColIndex] == imei && values[i][estadoColIndex] === "Validación") {
                sheet.getRange(i + 1, estadoColIndex + 1).setValue("Baja");
                sheet.getRange(i + 1, idAutorizaColIndex + 1).setValue("645");
                sheet.getRange(i + 1, comentariosColIndex + 1).setValue(`Baja aprobada. Razón: ${razonBaja}`);
                sheet.getRange(i + 1, ejecutadoColIndex + 1).setValue("SI"); // Marcar como ejecutado
                logMessage(`Baja aprobada en Sheet para Equipo IMEI: ${imei}`);
                break;
            }
        }

        // 3. Enviar correo de confirmación al solicitante
        const confirmSubject = `Solicitud de Baja de Equipo Celular APROBADA (${sucursal})`;
        const confirmBody = `Estimado(a) ${solicitanteEmail.split('@')[0]},\n\n` +
                            `Su solicitud de baja para el equipo con IMEI ${imei} de la sucursal ${sucursal} ha sido APROBADA.\n` +
                            `Razón de la baja: ${razonBaja}\n\n` +
                            `El estado del equipo ha sido actualizado a "Baja" en el sistema.\n\n` +
                            `Saludos cordiales,\nSistema de Gestión de Equipos.`;
        MailApp.sendEmail(solicitanteEmail, confirmSubject, confirmBody);
        
    } catch (e) {
        logMessage("Error al aprobar baja de equipo: " + e.message + " Stack: " + e.stack);
        confirmationTitle = "Error al Aprobar Baja";
        confirmationMessage = `Ha ocurrido un error al aprobar la baja del equipo con IMEI ${imei}: ${e.message}`;
        success = false;
    } finally {
        if (conn) conn.close();
    }

    return generateConfirmationPage(confirmationTitle, confirmationMessage, !success);
}

/**
 * Deniega la baja de un equipo, y notifica al solicitante.
 * Llamada desde el botón "Denegar" en el correo de validación de baja.
 * @param {number} idEquipo El ID del equipo a denegar.
 * @param {string} solicitanteEmail El correo del solicitante original.
 * @param {string} sucursal La sucursal del equipo.
 * @param {string} imei El IMEI del equipo.
 */
function denegarBajaEquipo(idEquipo, solicitanteEmail, sucursal, imei) {
    let confirmationTitle = "Baja Denegada";
    let confirmationMessage = `Se ha enviado una notificación al solicitante sobre la denegación de la baja del equipo con IMEI ${imei}.`;
    let success = true;

    try {
        // No se actualiza el estado en BD/Hoja, ya que se mantiene en "Validación" para auditoría.

        // Enviar correo de denegación al solicitante
        const denySubject = `Solicitud de Baja de Equipo Celular DENEGADA (${sucursal})`;
        const denyBody = `Estimado(a) ${solicitanteEmail.split('@')[0]},\n\n` +
                         `Su solicitud de baja para el equipo con IMEI ${imei} de la sucursal ${sucursal} ha sido DENEGADA.\n` +
                         `Por favor, póngase en contacto con el administrador para más detalles.\n\n` +
                         `Saludos cordiales,\nSistema de Gestión de Equipos.`;
        MailApp.sendEmail(solicitanteEmail, denySubject, denyBody);
        logMessage(`Baja denegada para Equipo IMEI: ${imei}. Notificación enviada a ${solicitanteEmail}`);

    } catch (e) {
        logMessage("Error al denegar baja de equipo: " + e.message + " Stack: " + e.stack);
        confirmationTitle = "Error al Denegar Baja";
        confirmationMessage = `Ha ocurrido un error al denegar la baja del equipo con IMEI ${imei}: ${e.message}`;
        success = false;
    }

    return generateConfirmationPage(confirmationTitle, confirmationMessage, !success);
}

/**
 * Primera etapa de aprobación de venta: redirige a un formulario para solicitar el monto.
 * Llamada desde el botón "Aceptar" en el correo de validación de venta.
 * @param {number} idEquipo El ID del equipo a vender (Apps Script generado ID).
 * @param {string} solicitanteEmail El correo del solicitante original.
 * @param {string} personaVende La persona a la que se vende el equipo.
 * @param {string} sucursal La sucursal del equipo.
 * @param {string} imei El IMEI del equipo (para la actualización en DB).
 */
function aprobarVentaEquipoStep1(idEquipo, solicitanteEmail, personaVende, sucursal, imei) {
    // Redirige a un formulario HTML simple para pedir el monto de venta
    const scriptUrlBase = ScriptApp.getService().getUrl();
    // Asegurarse de pasar el IMEI a la siguiente etapa
    const redirectUrl = `${scriptUrlBase}?form=aprobarVentaForm&idEquipo=${idEquipo}&solicitanteEmail=${encodeURIComponent(solicitanteEmail)}&personaVende=${encodeURIComponent(personaVende)}&sucursal=${encodeURIComponent(sucursal)}&imei=${encodeURIComponent(imei)}`;
    
    return HtmlService.createHtmlOutput(`<script>window.top.location.href = '${redirectUrl}';</script>`);
}

/**
 * Segunda etapa de aprobación de venta: actualiza el estado, comentarios y notifica al solicitante.
 * Llamada desde el formulario HTML de aprobación de venta.
 * @param {Object} formData Datos del formulario, incluyendo idEquipo, montoVenta, imei, etc.
 */
function aprobarVentaEquipoStep2(formData) {
    const idEquipo = parseInt(formData.idEquipo);
    const montoVenta = parseFloat(formData.montoVenta) || 0;
    const solicitanteEmail = formData.solicitanteEmail;
    const personaVende = formData.personaVende;
    const sucursal = formData.sucursal;
    const imei = formData.imei; // Recibir IMEI

    let confirmationTitle = "Venta Aprobada";
    let confirmationMessage = `El equipo con IMEI ${imei} ha sido marcado como "Vendido" y se ha enviado una notificación al solicitante.`;
    let success = true;

    if (isNaN(montoVenta) || montoVenta <= 0) {
        // Reconstruir la URL con todos los parámetros para volver al formulario con error
        const scriptUrlBase = ScriptApp.getService().getUrl();
        const returnUrl = `${scriptUrlBase}?form=aprobarVentaForm&idEquipo=${idEquipo}&solicitanteEmail=${encodeURIComponent(solicitanteEmail)}&personaVende=${encodeURIComponent(personaVende)}&sucursal=${encodeURIComponent(sucursal)}&imei=${encodeURIComponent(imei)}`;
        return HtmlService.createHtmlOutput(`<h1>Error</h1><p>Monto de venta inválido. Por favor, ingrese un valor numérico mayor a cero.</p><p><a href="${returnUrl}">Volver al formulario de venta</a></p>`);
    }

    let conn = null;
    try {
        conn = getJdbcConnection();
        const sheet = getSheet("RECU");

        // 1. Actualizar en SQL Server usando IMEI
        const updateSql = `UPDATE Equipo_Usado SET Estado = ?, Comentarios = ? WHERE IMEI = ?`;
        let pstmt = conn.prepareStatement(updateSql);
        pstmt.setString(1, "Vendido");
        pstmt.setString(2, `El equipo fue vendido a ${personaVende} con costo de $${montoVenta.toFixed(2)} pesos.`);
        pstmt.setString(3, imei); // Usar IMEI para el WHERE
        pstmt.executeUpdate();
        pstmt.close();
        logMessage(`Venta aprobada en SQL para Equipo IMEI: ${imei}`);

        // 2. Actualizar en Google Sheet
        const range = sheet.getDataRange();
        const values = range.getValues();
        const headerRow = values[0];
        const imeiColIndex = headerRow.indexOf("IMEI"); // Obtener índice de IMEI
        const estadoColIndex = headerRow.indexOf("Estado del equipo");
        const comentariosColIndex = headerRow.indexOf("Comentarios");
        const ejecutadoColIndex = headerRow.indexOf("EJECUTADO");

        for (let i = 1; i < values.length; i++) {
            // Buscar por IMEI y por estado "Validación"
            if (values[i][imeiColIndex] == imei && values[i][estadoColIndex] === "Validación") {
                sheet.getRange(i + 1, estadoColIndex + 1).setValue("Vendido");
                sheet.getRange(i + 1, comentariosColIndex + 1).setValue(`El equipo fue vendido a ${personaVende} con costo de $${montoVenta.toFixed(2)} pesos.`);
                sheet.getRange(i + 1, ejecutadoColIndex + 1).setValue("SI"); // Marcar como ejecutado
                logMessage(`Venta aprobada en Sheet para Equipo IMEI: ${imei}`);
                break;
            }
        }

        // 3. Enviar correo de confirmación al solicitante
        const confirmSubject = `Solicitud de Venta de Equipo Celular APROBADA (${sucursal})`;
        const confirmBody = `Estimado(a) ${solicitanteEmail.split('@')[0]},\n\n` +
                            `Su solicitud de venta para el equipo con IMEI ${imei} de la sucursal ${sucursal} ha sido APROBADA.\n` +
                            `El equipo fue vendido a ${personaVende} con un costo de $${montoVenta.toFixed(2)} pesos.\n\n` +
                            `El estado del equipo ha sido actualizado a "Vendido" en el sistema.\n\n` +
                            `Saludos cordiales,\nSistema de Gestión de Equipos.`;
        MailApp.sendEmail(solicitanteEmail, confirmSubject, confirmBody);
        
    } catch (e) {
        logMessage("Error al aprobar venta de equipo (Step 2): " + e.message + " Stack: " + e.stack);
        confirmationTitle = "Error al Aprobar Venta";
        confirmationMessage = `Ha ocurrido un error al aprobar la venta del equipo con IMEI ${imei}: ${e.message}`;
        success = false;
    } finally {
        if (conn) conn.close();
    }

    return generateConfirmationPage(confirmationTitle, confirmationMessage, !success);
}

/**
 * Deniega la venta de un equipo, y notifica al solicitante.
 * Llamada desde el botón "Denegar" en el correo de validación de venta.
 * @param {number} idEquipo El ID del equipo a denegar.
 * @param {string} solicitanteEmail El correo del solicitante original.
 * @param {string} sucursal La sucursal del equipo.
 * @param {string} imei El IMEI del equipo.
 */
function denegarVentaEquipo(idEquipo, solicitanteEmail, sucursal, imei) {
    let confirmationTitle = "Venta Denegada";
    let confirmationMessage = `Se ha enviado una notificación al solicitante sobre la denegación de la venta del equipo con IMEI ${imei}.`;
    let success = true;

    try {
        // No se actualiza el estado en BD/Hoja, ya que se mantiene en "Validación" para auditoría.

        // Enviar correo de denegación al solicitante
        const denySubject = `Solicitud de Venta de Equipo Celular DENEGADA (${sucursal})`;
        const denyBody = `Estimado(a) ${solicitanteEmail.split('@')[0]},\n\n` +
                         `Su solicitud de venta para el equipo con IMEI ${imei} de la sucursal ${sucursal} ha sido DENEGADA.\n` +
                         `Por favor, póngase en contacto con el administrador para más detalles.\n\n` +
                         `Saludos cordiales,\nSistema de Gestión de Equipos.`;
        MailApp.sendEmail(solicitanteEmail, denySubject, denyBody);
        logMessage(`Venta denegada para Equipo IMEI: ${imei}. Notificación enviada a ${solicitanteEmail}`);

    } catch (e) {
        logMessage("Error al denegar venta de equipo: " + e.message + " Stack: " + e.stack);
        confirmationTitle = "Error al Denegar Venta";
        confirmationMessage = `Ha ocurrido un error al denegar la venta del equipo con IMEI ${imei}: ${e.message}`;
        success = false;
    }

    return generateConfirmationPage(confirmationTitle, confirmationMessage, !success);
}
