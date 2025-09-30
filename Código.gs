// Code.gs (Este es el archivo principal de tu proyecto de Apps Script)

// ====================================================================================================================
// =============================== 1. CONFIGURACIÓN GLOBAL Y VARIABLES ==============================================
// ====================================================================================================================

// ID de la hoja de cálculo de Google
const SPREADSHEET_ID = "1h4TxPJHZ8pynph3J6q2h4FnDyOnG0Uye3VrYqFriPCg"; 

// Información para conectar con la base de datos SQL Server
const DB_ADDRESS = 'gw.hemoeco.com:5300';
const DB_USER = '';
const DB_PWD = '';
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

// Mapeo inverso de ID de sucursal a nombre de 3 letras (para mostrar en el formulario)
const SUCURSAL_ID_TO_NAME_MAP = Object.fromEntries(
    Object.entries(SUCURSAL_MAP).map(([name, id]) => [id, name])
);

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

const MEU_SHEET_HEADERS = [
    "Marca temporal",
    "Dirección de correo electrónico",
    "IMEI del equipo",
    "Campo Modificado",
    "Valor Anterior",
    "Nuevo Valor",
    "Tipo de Operación",
    "Costo del Equipo",
    "Fecha de compra de Equipo",
    "Fecha de Recolección",
    "Fecha de Reasignacion",
    "Estado Anterior",
    "Nuevo Estado",
    "Observaciones",
    "Marca",
    "Modelo",
    "Numero_Telefono",
    "Sucursal Anterior",
    "Nueva Sucursal",
    "IDEquipo",
    "Error",
    "IDSucursal Anterior",
    "Nuevo IDSucursal",
    "IDEMPLEADO Anterior",
    "Nuevo IDEMPLEADO",
    "Responsable Anterior",
    "Nuevo Responsable",
    "IDRESGUARDO Anterior",
    "Nuevo IDRESGUARDO",
    "IDAUTORIZA Anterior",
    "Nuevo IDAUTORIZA",
    "Comentarios",
    "Documentacion",
    "EJECUTADO"
];

// Placeholder para los encabezados de las otras hojas.
const RLyE_SHEET_HEADERS = ["Marca temporal", "Campo1_RLyE", "Campo2_RLyE", "EJECUTADO"];
const ML_SHEET_HEADERS = ["Marca temporal", "Campo1_ML", "Campo2_ML", "EJECUTADO"];

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
                const scriptUrlBase = ScriptApp.getService().getUrl();
                // CORRECCIÓN: Usar '?form=' en lugar de '?action=' para redirigir a un formulario HTML
                Logger.log(scriptUrlBase);
                const redirectUrl = `${scriptUrlBase}?form=aprobarVentaForm&idEquipo=${params.idEquipo}&solicitanteEmail=${encodeURIComponent(params.solicitanteEmail)}&personaVende=${encodeURIComponent(params.personaVende)}&sucursal=${encodeURIComponent(params.sucursal)}&imei=${encodeURIComponent(params.imei)}`;
                Logger.log(redirectUrl);

                return HtmlService.createHtmlOutput(`<script>window.top.location.href = '${redirectUrl}';</script>`);
                
            case 'denegarVentaEquipo':
                htmlOutput = denegarVentaEquipo(
                    parseInt(params.idEquipo), 
                    params.solicitanteEmail, 
                    params.sucursal,
                    params.imei 
                );
                break;
            case 'aprobarVentaEquipoStep2':
                htmlOutput = aprobarVentaEquipoStep2(params);
                break;
            default:
                logMessage(`Acción no reconocida: ${action}`);
                htmlOutput = generateConfirmationPage(
                    'Error de Acción',
                    'Acción no reconocida o inválida.',
                    true 
                );
                break;
        }
        return htmlOutput;
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
            htmlFileToServe = 'modificarEquipoUsado';
            break;
        case 'aprobarVentaForm':
            htmlFileToServe = 'aprobarVentaForm';
            break;
        default:
            htmlFileToServe = 'registrarEquipoUsado';
            logMessage("Advertencia: Nombre de formulario inválido recibido: '" + formName + "'. Se carga 'registrarEquipoUsado'.");
            break;
    }

    const template = HtmlService.createTemplateFromFile(htmlFileToServe);
    
    if (e.parameter) {
        for (let param in e.parameter) {
            template[param] = e.parameter[param];
        }
    }

    return template.evaluate()
        .setTitle('Formularios Hemoeco')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
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
 * CORREGIDA para manejar correctamente las zonas horarias.
 * @param {Date|string} dateValue El valor de fecha a formatear.
 * @returns {string|null} La fecha formateada o null si es inválida/vacía.
 */
function formatDateForSql(dateValue) {
    if (!dateValue || (typeof dateValue === 'string' && dateValue.trim() === '')) return null;

    let date;

    // Si el valor es una cadena de texto en formato 'YYYY-MM-DD', la procesamos
    // manualmente para evitar que JavaScript la interprete como UTC.
    if (typeof dateValue === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(dateValue)) {
        const parts = dateValue.split('-');
        const year = parseInt(parts[0], 10);
        const month = parseInt(parts[1], 10) - 1; // El mes en JavaScript es 0-indexado (Ene=0)
        const day = parseInt(parts[2], 10);
        date = new Date(year, month, day);
    } else {
        // Para cualquier otro caso (si ya es un objeto Date o tiene otro formato),
        // usamos el método anterior.
        date = new Date(dateValue);
    }

    if (isNaN(date.getTime())) {
        logMessage(`Advertencia: Fecha inválida detectada en formatDateForSql: ${dateValue}`);
        return null;
    }
    
    // La salida sigue siendo la misma, pero ahora la fecha de entrada es la correcta.
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
 * Envía un correo electrónico con opciones avanzadas como cuerpo HTML, botones y archivos adjuntos.
 * @param {Object} options Objeto de configuración del correo.
 * @param {string|string[]} options.to El/los destinatario(s).
 * @param {string} options.subject El asunto del correo.
 * @param {string} options.body El cuerpo del correo en texto plano.
 * @param {Array<Object>} [options.buttons] Botones de acción opcionales.
 * @param {Array<GoogleAppsScript.Base.Blob>} [options.attachments] Archivos adjuntos opcionales.
 */
function sendEmailWithOptions(options) {
    const scriptUrlBase = ScriptApp.getService().getUrl();
    let htmlBody = `<p style="font-family: sans-serif;">${options.body.replace(/\n/g, '<br>')}</p>`;

    if (options.buttons && options.buttons.length > 0) {
        htmlBody += '<p style="margin-top: 20px;">';
        options.buttons.forEach(button => {
            let actionUrl = `${scriptUrlBase}?action=${button.action}`;
            if (button.params) {
                for (let param in button.params) {
                    actionUrl += `&${param}=${encodeURIComponent(button.params[param])}`;
                }
            }
            htmlBody += `<a href="${actionUrl}" style="display: inline-block; padding: 10px 20px; margin-right: 10px; border-radius: 5px; text-decoration: none; color: white; background-color: ${button.color || '#4CAF50'}; font-family: sans-serif;">${button.text}</a>`;
        });
        htmlBody += '</p>';
    }

    const mailOptions = {
        to: Array.isArray(options.to) ? options.to.join(',') : options.to,
        subject: options.subject,
        htmlBody: htmlBody
    };

    if (options.attachments && options.attachments.length > 0) {
        mailOptions.attachments = options.attachments;
    }

    MailApp.sendEmail(mailOptions);
    logMessage(`Correo enviado a ${options.to} con asunto: "${options.subject}"`);
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


function buscarEquipoIMEI() {
  buscarEquipoPorIMEI("111111111111131");
}

/**
 * Busca un equipo usado por su IMEI y devuelve sus datos.
 * @param {string} imei El IMEI del equipo a buscar.
 * @returns {Object} Un objeto con 'success' (boolean) y 'data' (objeto con los datos del equipo) o 'message' de error.
 */
function buscarEquipoPorIMEI(imei) {
    let conn = null;
    let pstmt = null;
    let results = null;
    try {
        conn = getJdbcConnection();
        
        // CORRECCIÓN: Se añaden RAM y ROM a la consulta SQL
        const query = `
            SELECT 
                ID_Equipo, Costo_Equipo, Fecha_Compra, Fecha_Recoleccion, Fecha_Reasignacion, 
                Estado, Observaciones, Marca, Modelo, RAM, ROM, Numero_Telefono, IDSUCURSAL, 
                Responsable, Comentarios, Documentacion 
            FROM Equipo_Usado 
            WHERE IMEI = ? AND Estado NOT IN ('Baja','Vendido', 'Validación','Robado')
            ORDER BY ID_Equipo DESC`;
            
        pstmt = conn.prepareStatement(query);
        pstmt.setString(1, imei);
        results = pstmt.executeQuery();

        if (results.next()) {
            const idSucursal = results.getInt("IDSUCURSAL");
            const data = {
                ID_Equipo: results.getInt("ID_Equipo"),
                Costo_Equipo: results.getDouble("Costo_Equipo"),
                Fecha_Compra: getFechaBD(results.getObject("Fecha_Compra")),
                Fecha_Recoleccion: getFechaBD(results.getObject("Fecha_Recoleccion")),
                Fecha_Reasignacion: getFechaBD(results.getObject("Fecha_Reasignacion")),
                Estado: results.getString("Estado"),
                Observaciones: results.getString("Observaciones"),
                Marca: results.getString("Marca"),
                Modelo: results.getString("Modelo"),
                RAM: results.getString("RAM"),
                ROM: results.getString("ROM"),
                Numero_Telefono: results.getString("Numero_Telefono"),
                IDSUCURSAL: idSucursal,
                IDSUCURSAL_Name: SUCURSAL_ID_TO_NAME_MAP[idSucursal] || String(idSucursal),
                Responsable: results.getString("Responsable"),
                Comentarios: results.getString("Comentarios"),
                Documentacion: results.getString("Documentacion")
            };
            
            logMessage(`Equipo encontrado para IMEI ${imei}: ${JSON.stringify(data)}`);
            return { success: true, data: data };
        } else {
            return { success: false, message: "No se encontró un equipo activo con ese IMEI o su estado no permite modificación." };
        }
        
    } catch (e) {
        logMessage("Error al buscar equipo por IMEI en SQL Server: " + e.message);
        return { success: false, message: "Error al buscar equipo: " + e.message };
    } finally {
        if (results) results.close();
        if (pstmt) pstmt.close();
        if (conn) conn.close();
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

/**}
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
 * Obtiene el nombre y ID de empleado de la hoja "BD" basado en IDPUESTO y SUCURSALNOMINA.
 * Esto es para el caso de "Stock" en Modificar Equipo Usado.
 * @param {number} idSucursalBD El ID numérico de la sucursal.
 * @param {JdbcConnection} conn La conexión JDBC ya establecida.
 * @returns {Object|null} Un objeto {nombre: string, id: string} o null si no se encuentra.
 */
function getResponsableAndIdFromBDByPuesto(idSucursalBD, conn) {
    try {
        const sheetBD = getSheet("BD"); // Hoja "BD"
        const range = sheetBD.getDataRange();
        const values = range.getValues(); // Obtiene todos los valores de la hoja

        // Columnas en la hoja "BD"
        const nombreColIndex = 0; // NOMBRECOMPLETO (Col A)
        const idEmpleadoColIndex = 1; // IDEMPLEADO (Col B)
        const idPuestoColIndex = 2; // IDPUESTO (Col C)
        const sucursalNominaColIndex = 5; // SUCURSALNOMINA (Col F)
        const idSucursalColIndex = 6; // IDSUCURSAL (Col G) - Esto es el ID numérico de la sucursal

        let matchingRecords = [];

        for (let i = 1; i < values.length; i++) { // Empezar desde la fila 1 (después de encabezados)
            const row = values[i];
            const currentIdPuesto = row[idPuestoColIndex] ? String(row[idPuestoColIndex]).trim() : '';
            const currentSucursalNomina = row[sucursalNominaColIndex] ? String(row[sucursalNominaColIndex]).trim() : '';
            const currentIdSucursal = row[idSucursalColIndex]; // El ID numérico de la sucursal en la hoja BD

            // Verificar si IDPUESTO es '6' o '47' y si SUCURSALNOMINA coincide con el nombre de la sucursal
            // Y si el IDSUCURSAL en la hoja BD coincide con el ID numérico de la sucursal del equipo
            if ((currentIdPuesto === '6' || currentIdPuesto === '47') && currentIdSucursal == idSucursalBD) {
                matchingRecords.push({
                    nombre: row[nombreColIndex] ? String(row[nombreColIndex]).trim() : null,
                    id: row[idEmpleadoColIndex] ? String(row[idEmpleadoColIndex]).trim() : null,
                    idEmpleadoNum: parseInt(row[idEmpleadoColIndex]) || Infinity // Para ordenar por IDEMPLEADO menor
                });
            }else if((currentIdPuesto === '1') && currentIdSucursal == idSucursalBD){
                matchingRecords.push({
                    nombre: row[nombreColIndex] ? String(row[nombreColIndex]).trim() : null,
                    id: row[idEmpleadoColIndex] ? String(row[idEmpleadoColIndex]).trim() : null,
                    idEmpleadoNum: parseInt(row[idEmpleadoColIndex]) || Infinity // Para ordenar por IDEMPLEADO menor
                });
            }
        }

        if (matchingRecords.length > 0) {
            // Ordenar por IDEMPLEADO menor
            matchingRecords.sort((a, b) => a.idEmpleadoNum - b.idEmpleadoNum);
            logMessage(`Responsable encontrado para Stock: ${matchingRecords[0].nombre} (ID: ${matchingRecords[0].id})`);
            return { nombre: matchingRecords[0].nombre, id: matchingRecords[0].id };
        }

        logMessage(`No se encontró un responsable con IDPUESTO '6' o '47' para la sucursal ID: ${idSucursalBD}.`);
        return null;
    } catch (e) {
        logMessage("Error al obtener Responsable por Puesto y Sucursal de la hoja 'BD': " + e.message);
        return null;
    }
}

/**
 * Extraer la consulta de base de datos de manera correcta
 * @param {string} imei El IMEI a validar.
 * @param {JdbcConnection} conn La conexión JDBC ya establecida.
 * @returns {boolean} True si el IMEI es único, false si ya existe.
 */

function getFechaBD(fecha_bd) {
  if(fecha_bd) {
    const fecha = new Date(fecha_bd);
    if (isNaN(fecha.getTime())) {
      return ''; // Devuelve vacío si la fecha es inválida
    }
    const dia = String(fecha.getDate()).padStart(2, '0');
    const mes = String(fecha.getMonth() + 1).padStart(2, '0'); // Los meses son de 0 a 11
    const anio = fecha.getFullYear();

    // CORRECCIÓN: Devolver en formato YYYY-MM-DD
    return `${anio}-${mes}-${dia}`;
  }
  return fecha_bd;
}

/**
 * Función de ayuda para buscar un equipo por ID en una tabla específica.
 * @private
 */
function _buscarEquipoPorId(tabla, id, conn) {
    let equipo = {};
    const query = `SELECT IMEI, Marca, Modelo, Estado, Fecha_Compra, Costo_Equipo, Observaciones FROM ${tabla} WHERE ID_Equipo = ?`;
    const stmt = conn.prepareStatement(query);
    stmt.setInt(1, id);
    const rs = stmt.executeQuery();
    if (rs.next()) {
        equipo = {
            IMEI: rs.getString("IMEI"),
            Marca: rs.getString("Marca"),
            Modelo: rs.getString("Modelo"),
            Estado: rs.getString("Estado"),
            Fecha_Compra: getFechaBD(rs.getObject("Fecha_Compra")),
            Costo_Equipo: rs.getDouble("Costo_Equipo"),
            Observaciones: rs.getString("Observaciones")
        };
    }
    rs.close();
    stmt.close();
    return equipo;
}

/**
 * Busca una línea telefónica por su número y el equipo vinculado.
 * VERSIÓN FINAL con depuración y lógica corregida.
 * @param {string} telefono El número de 10 dígitos a buscar.
 * @returns {object} Un objeto con el resultado de la búsqueda.
 */
function buscarLineaPorTelefono(telefono) {
    let conn;
    try {
        conn = getJdbcConnection();
        
        const lineaQuery = `
            SELECT 
                Región, Cuenta_padre, Cuenta, Teléfono, Clave_plan, Nombre_plan, Minutos, Mensajes, 
                Monto_renta, Duracion_plan, Fecha_inicio, Fecha_termino, SIM, Tipo, Responsable, 
                IDSUCURSAL, Extensión, Datos, Notas, IDEquipoNuevo, IDEquipoUsado, IMEI, Marca, Modelo, Equipo_ilimitado
            FROM Telefonía_Telcel 
            WHERE Teléfono = ? AND Tipo = 'SmartPhone'`;
        
        const stmt = conn.prepareStatement(lineaQuery);
        stmt.setString(1, telefono);
        const results = stmt.executeQuery();

        if (results.next()) {
            const idSucursal = results.getInt("IDSUCURSAL");
            const sucursalNombre = SUCURSAL_ID_TO_NAME_MAP[idSucursal] || '';
            
            const linea = { /* ... Se omite por brevedad, el contenido es el mismo ... */ };
            linea.region = results.getString("Región");
            linea.cuenta_padre = results.getString("Cuenta_padre");
            linea.cuenta = results.getString("Cuenta");
            linea.telefono = results.getString("Teléfono");
            linea.clave_plan = results.getString("Clave_plan");
            linea.nombre_plan = results.getString("Nombre_plan");
            linea.minutos = results.getString("Minutos");
            linea.mensajes = results.getString("Mensajes");
            linea.monto_renta = results.getDouble("Monto_renta");
            linea.duracion_plan = results.getString("Duracion_plan");
            linea.fecha_inicio = getFechaBD(results.getObject("Fecha_inicio"));
            linea.fecha_termino = getFechaBD(results.getObject("Fecha_termino"));
            linea.sim = results.getString("SIM");
            linea.tipo = results.getString("Tipo");
            linea.responsable = results.getString("Responsable");
            linea.idsucursal = sucursalNombre;
            linea.extension = results.getString("Extensión");
            linea.datos = results.getDouble("Datos");
            linea.notas = results.getString("Notas");

            let equipoVinculado = null;
            
            // --- LÓGICA DE BÚSQUEDA CORREGIDA Y CON DEPURACIÓN ---
            
            // 1. Intentar con Equipo Nuevo
            const idEquipoNuevo = results.getInt("IDEquipoNuevo");
            if (!results.wasNull()) {
                logMessage(`Línea tiene IDEquipoNuevo: ${idEquipoNuevo}. Buscando en tabla Equipo_Nuevo...`);
                const equipoNuevoStmt = conn.prepareStatement("SELECT * FROM Equipo_Nuevo WHERE ID_Equipo = ?");
                equipoNuevoStmt.setInt(1, idEquipoNuevo);
                const rsNuevo = equipoNuevoStmt.executeQuery();

                if (rsNuevo.next()) {
                    logMessage(`Éxito: Se encontró el equipo en Equipo_Nuevo.`);
                    equipoVinculado = {
                        IMEI: rsNuevo.getString("IMEI"), Marca: rsNuevo.getString("Marca"), Modelo: rsNuevo.getString("Modelo"),
                        Estado: rsNuevo.getString("Estado"), Fecha_Compra: getFechaBD(rsNuevo.getObject("Fecha_Compra")),
                        Costo_Equipo: rsNuevo.getDouble("Costo_Equipo"), Observaciones: rsNuevo.getString("Observaciones"),
                        TipoEquipo: 'Nuevo'
                    };
                } else {
                    logMessage(`ADVERTENCIA: No se encontró equipo en Equipo_Nuevo con ID ${idEquipoNuevo}, aunque el ID existía en la línea.`);
                }
                rsNuevo.close();
                equipoNuevoStmt.close();
            }

            // 2. Si no se encontró nada, intentar con Equipo Usado
            if (!equipoVinculado) {
                const idEquipoUsado = results.getInt("IDEquipoUsado");
                if (!results.wasNull()) {
                    logMessage(`Línea tiene IDEquipoUsado: ${idEquipoUsado}. Buscando en tabla Equipo_Usado...`);
                    const equipoUsadoStmt = conn.prepareStatement("SELECT * FROM Equipo_Usado WHERE ID_Equipo = ?");
                    equipoUsadoStmt.setInt(1, idEquipoUsado);
                    const rsUsado = equipoUsadoStmt.executeQuery();

                    if (rsUsado.next()) {
                         logMessage(`Éxito: Se encontró el equipo en Equipo_Usado.`);
                         equipoVinculado = {
                            IMEI: rsUsado.getString("IMEI"), Marca: rsUsado.getString("Marca"), Modelo: rsUsado.getString("Modelo"),
                            Estado: rsUsado.getString("Estado"), Fecha_Compra: getFechaBD(rsUsado.getObject("Fecha_Compra")),
                            Costo_Equipo: rsUsado.getDouble("Costo_Equipo"), Observaciones: rsUsado.getString("Observaciones"),
                            TipoEquipo: 'Usado'
                        };
                    } else {
                         logMessage(`ADVERTENCIA: No se encontró equipo en Equipo_Usado con ID ${idEquipoUsado}.`);
                    }
                    rsUsado.close();
                    equipoUsadoStmt.close();
                }
            }

            // 3. Si sigue sin haber equipo, usar la información de la línea como último recurso
            if (!equipoVinculado) {
                logMessage(`No se encontró equipo vinculado en tablas de equipos. Usando datos de la línea como fallback.`);
                equipoVinculado = {
                    IMEI: results.getString("IMEI"), Marca: results.getString("Marca"), Modelo: results.getString("Modelo"),
                    Estado: "N/A (Info. de Línea)", Fecha_Compra: getFechaBD(results.getObject("Fecha_inicio")),
                    Costo_Equipo: results.getDouble("Equipo_ilimitado"), Observaciones: "Datos tomados del registro de la línea.",
                    TipoEquipo: 'Linea'
                };
            }
            
            return { success: true, data: { linea: linea, equipoVinculado: equipoVinculado } };
        } else {
            return { success: false, message: "No se encontró una línea activa de SmartPhone con ese número." };
        }
    } catch (e) {
        logMessage("Error en buscarLineaPorTelefono: " + e.message);
        return { success: false, message: e.message };
    } finally {
        if (conn) conn.close();
    }
}

/**
 * Función de ayuda para buscar todos los datos de un equipo por ID en una tabla específica.
 * CORREGIDA para usar una consulta explícita y ser más robusta.
 * @private
 */
function _buscarEquipoPorId(tabla, id, conn) {
    let equipo = null;
    const query = `
        SELECT 
            ID_Equipo, Costo_Equipo, Fecha_Compra, Estado, Observaciones, 
            Marca, Modelo, RAM, ROM, IMEI, IDSUCURSAL 
        FROM ${tabla} 
        WHERE ID_Equipo = ?`;

    const stmt = conn.prepareStatement(query);
    stmt.setInt(1, id);
    const rs = stmt.executeQuery();

    if (rs.next()) {
        equipo = {
            ID_Equipo: rs.getInt("ID_Equipo"),
            IMEI: rs.getString("IMEI"),
            Marca: rs.getString("Marca"),
            Modelo: rs.getString("Modelo"),
            RAM: rs.getString("RAM"),
            ROM: rs.getString("ROM"),
            Estado: rs.getString("Estado"),
            Fecha_Compra: getFechaBD(rs.getObject("Fecha_Compra")),
            Costo_Equipo: rs.getDouble("Costo_Equipo"),
            Observaciones: rs.getString("Observaciones"),
            IDSUCURSAL: rs.getInt("IDSUCURSAL")
        };
    }
    rs.close();
    stmt.close();
    
    return equipo; // Si no encuentra nada, devuelve null correctamente.
}

/**
 * Devuelve una lista de responsables filtrada por sucursal.
 * @param {string} sucursal Las iniciales de la sucursal seleccionada (ej. "GDL").
 * @returns {Object} Un objeto con la lista de responsables.
 */
function getResponsablesPorSucursal(sucursal) {
  try {
    const sheet = getSheet("BD");
    const values = sheet.getDataRange().getValues();
    
    values.shift(); // Quita el encabezado para no procesarlo
    const nombreColIndex = 0;   // Columna A: NOMBRECOMPLETO
    const sucursalColIndex = 7; // Columna H: SUCURSAL (0-indexed)

    const responsablesFiltrados = values
      // CORRECCIÓN: Se añade .toString().trim() para hacer la comparación más robusta.
      .filter(row => row[sucursalColIndex] && row[sucursalColIndex].toString().trim() === sucursal)
      .map(row => row[nombreColIndex]);

    logMessage(`Buscando responsables para sucursal '${sucursal}'. Encontrados: ${responsablesFiltrados.length}`);
    return { success: true, data: responsablesFiltrados };
    
  } catch (e) {
    logMessage("Error en getResponsablesPorSucursal: " + e.message);
    return { success: false, message: e.message };
  }
}

/**
 * Devuelve una lista de modelos de celular filtrada por marca.
 * @param {string} marca La marca seleccionada en el formulario.
 * @returns {Object} Un objeto con la lista de modelos.
 */
function getModelosPorMarca(marca) {
  try {
    let columna;
    // Mapeo de la marca a la columna correspondiente en la hoja "Listas2"
    switch (marca) {
      case "Samsung":   columna = "H2:H"; break;
      case "Xiaomi":    columna = "I2:I"; break;
      case "iPhone":    columna = "J2:J"; break;
      case "Motorola":  columna = "K2:K"; break;
      case "Oppo":      columna = "L2:L"; break;
      default:
        return { success: true, data: [] }; // Si no coincide, devuelve una lista vacía
    }
    return getDropdownOptions('Listas2', columna);
  } catch (e) {
    logMessage("Error en getModelosPorMarca: " + e.message);
    return { success: false, message: e.message };
  }
}

/**
 * Devuelve una lista de números de teléfono filtrada por sucursal.
 * @param {string} sucursal Las iniciales de la sucursal seleccionada (ej. "GDL").
 * @returns {Object} Un objeto con la lista de teléfonos.
 */
function getTelefonosPorSucursal(sucursal) {
    try {
        const sheet = getSheet("BD");
        const values = sheet.getDataRange().getValues();
        
        values.shift(); // Quita el encabezado
        const telefonoColIndex = 10; // Columna K: TELEFONO
        const sucursalColIndex = 13; // Columna N: SUCURSAL_LINEA

        const telefonosFiltrados = values
            .filter(row => row[sucursalColIndex] && row[sucursalColIndex].toString().trim() === sucursal)
            .map(row => row[telefonoColIndex])
            .filter(telefono => telefono); // Elimina posibles valores vacíos

        return { success: true, data: telefonosFiltrados };
    } catch(e) {
        logMessage("Error en getTelefonosPorSucursal: " + e.message);
        return { success: false, message: e.message };
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
    let idEquipo = null;
    
    try {
        const sheet = getSheet("RECU"); // Hoja específica para RECU
        
        conn = getJdbcConnection();

        const fechaFormulario = new Date(); // Fecha actual del envío del formulario
        const solicitanteEmail = Session.getActiveUser().getEmail(); // Correo del usuario que envía el formulario

        // --- Validaciones de campos (lado del servidor) ---
        const costoEquipo = parseFloat(formObject.costoEquipo) || 0;
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

        let responsableName = formData.responsable || null;
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
                estadoFinalDB = "Validación";
                responsableName = null;
                idEmpleado = null;
                idResguardo = null;

                const nombreSolicitanteBaja = formData.nombreSolicitanteBaja || '';
                const razonBaja = formData.razonBaja || '';

                if (!nombreSolicitanteBaja || !razonBaja) {
                    throw new Error("Para la Baja, el nombre del solicitante y la razón son obligatorios.");
                }

                idEquipo = obtenerUltimoIDEquipoSQL(conn);

                // --- NUEVA LÓGICA PARA MANEJAR ARCHIVOS ---
                let attachments = [];
                let folderUrl = '';

                if (formData.archivos && formData.archivos.length > 0) {
                    const parentFolderId = "1zJ3p-eaQb7NyQtOn8k8gG8pjyAIlhf1i"; // ID de la carpeta principal "EQUIPO USADO"
                    const parentFolder = DriveApp.getFolderById(parentFolderId);
                    const newFolderName = `BAJA${formData.imei}`;
                    const newFolder = parentFolder.createFolder(newFolderName);
                    folderUrl = newFolder.getUrl();

                    formData.archivos.forEach(fileInfo => {
                        const decodedData = Utilities.base64Decode(fileInfo.data);
                        const blob = Utilities.newBlob(decodedData, fileInfo.mimeType, fileInfo.fileName);
                        newFolder.createFile(blob);
                        attachments.push(blob);
                    });
                    logMessage(`Archivos para IMEI ${formData.imei} subidos a Drive: ${folderUrl}`);
                }
                // --- FIN DE LA LÓGICA DE ARCHIVOS ---

                const bajaSubject = `VALIDACIÓN DE BAJA DE EQUIPO CELULAR DE LA SUCURSAL '${sucursalCodigo3Letras}'`;
                let bajaBody = `Buen día Jorge Fernández.\n\n` +
                              `${nombreSolicitanteBaja} con correo electrónico ${solicitanteEmail} está solicitando que se valide la baja del siguiente equipo celular:\n` +
                              `Marca: ${formData.marca}\n` +
                              `Modelo: ${formData.modelo}\n` +
                              `IMEI: ${imei}\n` +
                              `Fecha de compra: ${formData.fechaCompra}\n` +
                              `El costo del equipo es: $${costoEquipo.toFixed(2)}\n\n` +
                              `La razón de la baja es la siguiente: ${razonBaja}\n\n`;

                if (folderUrl) {
                    bajaBody += `Se ha creado una carpeta con la evidencia en Google Drive: ${folderUrl}\n\n`;
                }
                bajaBody += `Si aprueba la baja del equipo celular, seleccione la opción de "Aceptar" de lo contrario seleccione la opción de "Denegar" y póngase en contacto con la persona que realizó la solicitud para auditar la baja del equipo.`;

                const bajaButtons = [
                    { text: 'Aceptar', action: 'aprobarBajaEquipo', params: { idEquipo: idEquipo, solicitanteEmail: solicitanteEmail, razonBaja: razonBaja, sucursal: sucursalCodigo3Letras, imei: imei }, color: '#28a745' },
                    { text: 'Denegar', action: 'denegarBajaEquipo', params: { idEquipo: idEquipo, solicitanteEmail: solicitanteEmail, sucursal: sucursalCodigo3Letras, imei: imei }, color: '#dc3545' }
                ];

                // Se usa la nueva función de correo que soporta adjuntos
                sendEmailWithOptions({
                    to: ADMIN_EMAILS,
                    subject: bajaSubject,
                    body: bajaBody,
                    buttons: bajaButtons,
                    attachments: attachments
                });

                response.message = `Solicitud de baja enviada para validación. ID Equipo: ${idEquipo}.`;
                break;

            case 'Stock':
                const sucursalDBID = idSucursalBD; 
                const resguardoInfo = getResponsableAndIdFromBDByPuesto(sucursalDBID, conn); 
                
                if (resguardoInfo && resguardoInfo.nombre && resguardoInfo.id) {
                    // CORRECCIÓN: Se usan las variables correctas (responsableName, idEmpleado, etc.)
                    responsableName = resguardoInfo.nombre;
                    idEmpleado = resguardoInfo.id;
                    idResguardo = resguardoInfo.id; 
                } else {
                    response.message = `No se encontró un responsable válido (puesto 6 o 47) para Stock en la sucursal ID: ${sucursalDBID}.`;
                    logMessage("Error: " + response.message);
                    return response;
                }
                break;
            
            case 'Robado':
                // Numero_Telefono y Responsable se envían como NULL desde el formulario (disabled)
                // IDRESGUARDO se queda en NULL para estos casos (no se obtiene por sucursal)
                responsableName = null;
                idEmpleado = null;
                idResguardo = null;

                idResguardo = null; 
                if (formData.estado === 'Robado' && !comentariosFinal) {
                    comentariosFinal = "Favor de agregar anotaciones del Robo";
                }
                // Obtener el ID de Equipo_Usado para la inserción
                idEquipo = obtenerUltimoIDEquipoSQL(conn); 
                break;

            case 'Vendido':
                estadoFinalDB = "Validación"; // Se guarda como "Validación" en BD/Hoja

                responsableName = null;
                idEmpleado = null;
                idResguardo = null;

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

            default: 
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
            response.success = true;
            if (estadoFinalDB !== "Validación") {
                response.message = `Equipo con ID ${idEquipo} insertado en SQL y registrado en hoja de cálculo.`;
            }
            //----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                if (formData.estado === 'Nuevo') {

                    idEquipoNuevo = obtenerUltimoIDEquipoNuevoSQL(conn);
                    const insertEquipoQuery = `
                        INSERT INTO Equipo_Nuevo (ID_Equipo, Costo_Equipo, Fecha_Compra, Estado, Marca, Modelo, RAM, ROM, IMEI, Observaciones, IDEMPLEADO, Responsable, IDSUCURSAL)
                        VALUES (?, ?, ?, 'Nuevo', ?, ?, ?, ?, ?, ?, ?, ?, ?)`;
                    
                    const pstmtEquipo = conn.prepareStatement(insertEquipoQuery);
                    pstmtEquipo.setInt(1, idEquipoNuevo);
                    pstmt.setObject(2, costoEquipo);
                    pstmt.setString(3, formatDateForSql(fechaCompra));
                    pstmt.setString(4, formData.marca || null);
                    pstmt.setString(5, formData.modelo || null);
                    pstmt.setString(6, formData.ram || null);
                    pstmt.setString(7, formData.rom || null);
                    pstmt.setString(8, imei);
                    pstmtEquipo.setString(9, 'EQUIPO NUEVO REGISTRADO MANUALMENTE SIN LINEA ASÍGNADA');
                    pstmtEquipo.setString(10, idEmpleado);
                    pstmtEquipo.setString(11, responsableName);
                    pstmtEquipo.setInt(12, idSucursalBD);
                    pstmtEquipo.executeUpdate();
                    pstmtEquipo.close();
                    logMessage(`Equipo nuevo creado con ID: ${idEquipoNuevo}`);
                }
        //----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

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
// En tu archivo Code.gs

/**
 * Procesa los datos del formulario "Dar de Alta Línea y Equipo",
 * con lógica condicional para registrar el equipo solo si es necesario.
 */
function procesarALyEFormulario(formData) {
    let response = { success: false, message: "" };
    logMessage("Datos recibidos de ALyE: " + JSON.stringify(formData));
    let conn = null;

    try {
        conn = getJdbcConnection();
        conn.setAutoCommit(false); // Iniciar transacción para asegurar la integridad de los datos

        // --- 1. VALIDACIONES BÁSICAS DE LA LÍNEA ---
        const telefono = formData.telefono ? formData.telefono.trim() : '';
        if (!telefono || !/^\d{10}$/.test(telefono)) {
            throw new Error("El Teléfono debe ser un número válido de 10 dígitos.");
        }
        if (!validarTelefonoUnicoSQL(telefono, conn)) {
            throw new Error(`El Teléfono '${telefono}' ya existe en la base de datos.`);
        }
        
        let idEquipoNuevo = null;
        let idSucursalInt = SUCURSAL_MAP[formData.idsucursal] || null;
        if (idSucursalInt === null) {
            throw new Error("La Sucursal seleccionada es inválida.");
        }
        const idEmpleadoInt = getResponsableID(formData.responsable);

        // --- 2. LÓGICA CONDICIONAL PARA EL EQUIPO CELULAR ---
        if (formData.tipo === 'Smartphone' && !formData.lineaSinequipo) {
            
            const imeiEquipoNuevo = formData.imei_equipo_nuevo ? formData.imei_equipo_nuevo.trim() : '';
            if (!imeiEquipoNuevo || !/^\d{15,16}$/.test(imeiEquipoNuevo)) {
                throw new Error("El IMEI del equipo es obligatorio y debe ser válido para un Smartphone.");
            }
            if (!validarIMEINuevoUnicoSQL(imeiEquipoNuevo, conn)) {
                throw new Error(`El IMEI '${imeiEquipoNuevo}' ya existe en la base de datos.`);
            }

            idEquipoNuevo = obtenerUltimoIDEquipoNuevoSQL(conn);
            const insertEquipoQuery = `
                INSERT INTO Equipo_Nuevo (ID_Equipo, Costo_Equipo, Fecha_Compra, Estado, Marca, Modelo, RAM, ROM, IMEI, Observaciones, IDEMPLEADO, Responsable, IDSUCURSAL)
                VALUES (?, ?, ?, 'Asignado', ?, ?, ?, ?, ?, ?, ?, ?, ?)`;
            
            const pstmtEquipo = conn.prepareStatement(insertEquipoQuery);
            pstmtEquipo.setInt(1, idEquipoNuevo);
            pstmtEquipo.setDouble(2, parseFloat(formData.costo_equipo_nuevo) || 0);
            pstmtEquipo.setString(3, formatDateForSql(formData.fecha_inicio));
            pstmtEquipo.setString(4, formData.marca_nuevo);
            pstmtEquipo.setString(5, formData.modelo_nuevo);
            pstmtEquipo.setString(6, formData.ram_nuevo);
            pstmtEquipo.setString(7, formData.rom_nuevo);
            pstmtEquipo.setString(8, imeiEquipoNuevo);
            pstmtEquipo.setString(9, formData.observaciones_equipo_nuevo);
            pstmtEquipo.setString(10, idEmpleadoInt);
            pstmtEquipo.setString(11, formData.responsable);
            pstmtEquipo.setInt(12, idSucursalInt);
            pstmtEquipo.executeUpdate();
            pstmtEquipo.close();
            logMessage(`Equipo nuevo creado con ID: ${idEquipoNuevo}`);
        }

        // --- 3. INSERCIÓN EN LA TABLA Telefonía_Telcel ---
        const insertTelcelQuery = `
            INSERT INTO Telefonía_Telcel (
                Región, Cuenta_padre, Cuenta, Teléfono, Clave_plan, Nombre_plan, Minutos, Mensajes,
                Monto_renta, Duracion_plan, Fecha_inicio, Fecha_termino, SIM, Tipo, Responsable, 
                Notas, IDEMPLEADO, IDSUCURSAL, Datos, Extensión, IDEquipoNuevo, 
                Marca, Modelo, IMEI, Equipo_ilimitado, Estado
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Activa')`;

        const pstmtTelcel = conn.prepareStatement(insertTelcelQuery);
        pstmtTelcel.setString(1, formData.region);
        pstmtTelcel.setString(2, formData.cuenta_padre);
        pstmtTelcel.setString(3, formData.cuenta);
        pstmtTelcel.setString(4, telefono);
        pstmtTelcel.setString(5, formData.clave_plan);
        pstmtTelcel.setString(6, formData.nombre_plan);
        pstmtTelcel.setString(7, formData.minutos);
        pstmtTelcel.setString(8, formData.mensajes);
        pstmtTelcel.setDouble(9, parseFloat(formData.monto_renta) || 0.0);
        pstmtTelcel.setString(10, formData.duracion_plan);
        pstmtTelcel.setString(11, formatDateForSql(formData.fecha_inicio));
        pstmtTelcel.setString(12, formatDateForSql(formData.fecha_termino));
        pstmtTelcel.setString(13, formData.sim);
        pstmtTelcel.setString(14, formData.tipo);
        pstmtTelcel.setString(15, formData.responsable);
        pstmtTelcel.setString(16, null);
        pstmtTelcel.setString(17, idEmpleadoInt);
        pstmtTelcel.setInt(18, idSucursalInt);
        pstmtTelcel.setDouble(19, parseFloat(formData.datos) || null);
        pstmtTelcel.setString(20, formData.extension);
        
        if (idEquipoNuevo) {
            pstmtTelcel.setInt(21, idEquipoNuevo);
            pstmtTelcel.setString(22, formData.marca_nuevo);
            pstmtTelcel.setString(23, formData.modelo_nuevo);
            pstmtTelcel.setString(24, formData.imei_equipo_nuevo);
            pstmtTelcel.setDouble(25, parseFloat(formData.costo_equipo_nuevo) || 0.0);
        } else {
            pstmtTelcel.setObject(21, null);
            pstmtTelcel.setObject(22, null);
            pstmtTelcel.setObject(23, null);
            pstmtTelcel.setObject(24, null);
            pstmtTelcel.setObject(25, null);
        }

        pstmtTelcel.executeUpdate();
        pstmtTelcel.close();
        
        conn.commit();
        
        response.success = true;
        response.message = `Línea registrada exitosamente.` + (idEquipoNuevo ? ` Equipo Nuevo ID: ${idEquipoNuevo}.` : '');

    } catch (e) {
        if (conn) conn.rollback();
        response.message = e.message;
        logMessage("Error en procesarALyEFormulario: " + e.message + " Stack: " + e.stack);
    } finally {
        if (conn) conn.close();
    }
    return response;
}


// --- 4.3. Formulario: Renovación de Línea y Equipo (RLyE) ---

/**
 * Procesa el formulario de Renovación de Línea y Equipo.
 * Utiliza la lógica robusta de archivado de equipo de 'modificarLinea'.
 */
function procesarRenovacionFormulario(formData) {
    let conn;
    try {
        logMessage("Iniciando procesamiento de Renovación: " + JSON.stringify(formData));
        conn = getJdbcConnection();
        conn.setAutoCommit(false); // Iniciar transacción

        // --- 1. GESTIONAR EL EQUIPO VINCULADO ANTERIOR (Lógica corregida) ---
        const lineaStmt = conn.prepareStatement("SELECT IDEquipoNuevo, IDEquipoUsado FROM Telefonía_Telcel WHERE Teléfono = ?");
        lineaStmt.setString(1, formData.telefono);
        const lineaRs = lineaStmt.executeQuery();

        if (lineaRs.next()) {
            const idEquipoNuevoOriginal = lineaRs.getInt("IDEquipoNuevo");
            if (idEquipoNuevoOriginal && !lineaRs.wasNull()) {
                // CASO: El equipo anterior era NUEVO. Se archiva como USADO.
                logMessage(`Archivando equipo NUEVO original con ID: ${idEquipoNuevoOriginal}`);
                const equipoNuevoOriginal = _buscarEquipoPorId('Equipo_Nuevo', idEquipoNuevoOriginal, conn);
                
                if (equipoNuevoOriginal) {
                    // a) Actualizar estado del equipo en Equipo_Nuevo a "Usado"
                    const updateNuevoStmt = conn.prepareStatement("UPDATE Equipo_Nuevo SET Estado = 'Usado' WHERE ID_Equipo = ?");
                    updateNuevoStmt.setInt(1, idEquipoNuevoOriginal);
                    updateNuevoStmt.executeUpdate();
                    updateNuevoStmt.close();
                    logMessage(`Equipo Nuevo ${idEquipoNuevoOriginal} actualizado a estado 'Usado'.`);

                    // b) Crear nuevo registro en Equipo_Usado con los datos del equipo nuevo
                    const nuevoIdUsado = obtenerUltimoIDEquipoSQL(conn);
                    const resguardoInfo = getResponsableAndIdFromBDByPuesto(equipoNuevoOriginal.IDSUCURSAL, conn);
                    
                    const insertStmt = conn.prepareStatement(
                      `INSERT INTO Equipo_Usado (ID_Equipo, Costo_Equipo, Fecha_Compra, Fecha_Formulario, Estado, Observaciones, Marca, Modelo, RAM, ROM, IMEI, IDSUCURSAL, IDEMPLEADO, Responsable, IDRESGUARDO, Comentarios) 
                       VALUES (?, ?, ?, ?, 'Stock', ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'ENVIADO DESDE RENOVACIÓN')`
                    );
                    insertStmt.setInt(1, nuevoIdUsado);
                    insertStmt.setDouble(2, equipoNuevoOriginal.Costo_Equipo);
                    insertStmt.setString(3, formatDateForSql(new Date(equipoNuevoOriginal.Fecha_Compra)));
                    insertStmt.setString(4, formatDateForSql(new Date()));
                    insertStmt.setString(5, "ENVIADO DESDE RENOVACIÓN");
                    insertStmt.setString(6, equipoNuevoOriginal.Marca);
                    insertStmt.setString(7, equipoNuevoOriginal.Modelo);
                    insertStmt.setString(8, equipoNuevoOriginal.RAM);
                    insertStmt.setString(9, equipoNuevoOriginal.ROM);
                    insertStmt.setString(10, equipoNuevoOriginal.IMEI);
                    insertStmt.setInt(11, equipoNuevoOriginal.IDSUCURSAL);
                    insertStmt.setString(12, resguardoInfo ? resguardoInfo.id : null);
                    insertStmt.setString(13, resguardoInfo ? resguardoInfo.nombre : null);
                    insertStmt.setString(14, resguardoInfo ? resguardoInfo.id : null);
                    insertStmt.executeUpdate();
                    insertStmt.close();
                    logMessage(`Equipo Nuevo ${idEquipoNuevoOriginal} archivado como Equipo Usado ${nuevoIdUsado}.`);
                }
            } else {
                const idEquipoUsadoOriginal = lineaRs.getInt("IDEquipoUsado");
                if (idEquipoUsadoOriginal && !lineaRs.wasNull()) {
                   // CASO: El equipo anterior era USADO. Se actualiza a "Stock".
                   logMessage(`Actualizando equipo USADO original con ID: ${idEquipoUsadoOriginal} a Stock.`);
                   const equipoUsadoOriginal = _buscarEquipoPorId('Equipo_Usado', idEquipoUsadoOriginal, conn);
                   if (equipoUsadoOriginal) {
                       const resguardoInfo = getResponsableAndIdFromBDByPuesto(equipoUsadoOriginal.IDSUCURSAL, conn);
                       const updateStmt = conn.prepareStatement(
                           `UPDATE Equipo_Usado SET Estado = 'Stock', IDEMPLEADO = ?, Responsable = ?, IDRESGUARDO = ?, Comentarios = 'ENVIADO DESDE RENOVACIÓN', Numero_Telefono = NULL WHERE ID_Equipo = ?`
                       );
                       updateStmt.setString(1, resguardoInfo ? resguardoInfo.id : null);
                       updateStmt.setString(2, resguardoInfo ? resguardoInfo.nombre : null);
                       updateStmt.setString(3, resguardoInfo ? resguardoInfo.id : null);
                       updateStmt.setInt(4, idEquipoUsadoOriginal);
                       updateStmt.executeUpdate();
                       updateStmt.close();
                   }
                }
            }
        }
        lineaRs.close();
        lineaStmt.close();

        // --- 2. CREAR EL NUEVO EQUIPO DE LA RENOVACIÓN ---
        const idSucursalNueva = SUCURSAL_MAP[formData.idsucursal];
        const idEmpleadoNuevo = getResponsableID(formData.responsable);
        const nuevoIdEquipoRenovado = obtenerUltimoIDEquipoNuevoSQL(conn);
        const insertEquipoSql = `
            INSERT INTO Equipo_Nuevo (ID_Equipo, Costo_Equipo, Fecha_Compra, Estado, Marca, Modelo, RAM, ROM, IMEI, Observaciones, IDEMPLEADO, Responsable, IDSUCURSAL)
            VALUES (?, ?, ?, 'Asignado', ?, ?, ?, ?, ?, ?, ?, ?, ?)`;
        
        let pstmtEquipo = conn.prepareStatement(insertEquipoSql);
        pstmtEquipo.setInt(1, nuevoIdEquipoRenovado);
        pstmtEquipo.setDouble(2, parseFloat(formData.costo_equipo_nuevo));
        pstmtEquipo.setString(3, formatDateForSql(formData.fecha_inicio));
        pstmtEquipo.setString(4, formData.marca_nuevo);
        pstmtEquipo.setString(5, formData.modelo_nuevo);
        pstmtEquipo.setString(6, formData.ram_nuevo);
        pstmtEquipo.setString(7, formData.rom_nuevo);
        pstmtEquipo.setString(8, formData.imei_equipo_nuevo);
        pstmtEquipo.setString(9, formData.observaciones_equipo_nuevo);
        pstmtEquipo.setString(10, idEmpleadoNuevo);
        pstmtEquipo.setString(11, formData.responsable);
        pstmtEquipo.setInt(12, idSucursalNueva);
        pstmtEquipo.executeUpdate();
        pstmtEquipo.close();

        // --- 3. ACTUALIZAR LA LÍNEA TELEFÓNICA CON LOS NUEVOS DATOS ---
        const updateLineaSql = `
            UPDATE Telefonía_Telcel SET 
            Clave_plan = ?, Nombre_plan = ?, Minutos = ?, Mensajes = ?, Monto_renta = ?, Duracion_plan = ?,
            Fecha_inicio = ?, Fecha_termino = ?, Responsable = ?, IDEMPLEADO = ?, IDSUCURSAL = ?, Extensión = ?,
            Datos = ?, Notas = ?, IDEquipoNuevo = ?, IDEquipoUsado = NULL, IMEI = ?, Marca = ?, Modelo = ?, Equipo_ilimitado = ?
            WHERE Teléfono = ?`;

        let pstmtLinea = conn.prepareStatement(updateLineaSql);
        pstmtLinea.setString(1, formData.clave_plan);
        pstmtLinea.setString(2, formData.nombre_plan);
        pstmtLinea.setString(3, formData.minutos);
        pstmtLinea.setString(4, formData.mensajes);
        pstmtLinea.setDouble(5, parseFloat(formData.monto_renta));
        pstmtLinea.setString(6, formData.duracion_plan);
        pstmtLinea.setString(7, formatDateForSql(formData.fecha_inicio));
        pstmtLinea.setString(8, formatDateForSql(formData.fecha_termino));
        pstmtLinea.setString(9, formData.responsable);
        pstmtLinea.setString(10, idEmpleadoNuevo);
        pstmtLinea.setInt(11, idSucursalNueva);
        pstmtLinea.setString(12, formData.extension);
        pstmtLinea.setDouble(13, parseFloat(formData.datos));
        pstmtLinea.setString(14, formData.notas);
        pstmtLinea.setInt(15, nuevoIdEquipoRenovado);
        pstmtLinea.setString(16, formData.imei_equipo_nuevo);
        pstmtLinea.setString(17, formData.marca_nuevo);
        pstmtLinea.setString(18, formData.modelo_nuevo);
        pstmtLinea.setDouble(19, parseFloat(formData.costo_equipo_nuevo));
        pstmtLinea.setString(20, formData.telefono);
        pstmtLinea.executeUpdate();
        pstmtLinea.close();

        conn.commit();
        return { success: true, message: "Renovación procesada y equipo nuevo registrado exitosamente." };

    } catch (e) {
        if (conn) conn.rollback();
        logMessage("Error en procesarRenovacionFormulario: " + e.message + " Stack: " + e.stack);
        return { success: false, message: e.message };
    } finally {
        if (conn) conn.close();
    }
}
// --- 4.4. Formulario: Modificar Línea (ML) ---
/**
 * Procesa el formulario de Modificación de Línea con lógica de archivado de equipo.
 */
function procesarModificacionLineaFormulario(formData) {
    let conn;
    try {
        logMessage("Iniciando procesamiento de Modificación de Línea: " + JSON.stringify(formData));
        conn = getJdbcConnection();
        conn.setAutoCommit(false); // Iniciar transacción para asegurar la integridad de los datos
        
        const idSucursal = SUCURSAL_MAP[formData.idsucursal];
        const idEmpleado = getResponsableID(formData.responsable);
        const equipoAVincular = formData.vincularEquipo;

        // 1. SI SE VA A REEMPLAZAR UN EQUIPO NUEVO, LO ARCHIVAMOS EN EQUIPO_USADO
        if (equipoAVincular.id && formData.estadoEquipoAnterior) {
            const lineaOriginalStmt = conn.prepareStatement("SELECT IDEquipoNuevo, IDSUCURSAL FROM Telefonía_Telcel WHERE Teléfono = ?");
            lineaOriginalStmt.setString(1, formData.telefono);
            const lineaRs = lineaOriginalStmt.executeQuery();

            if (lineaRs.next()) {
                const idEquipoNuevoOriginal = lineaRs.getInt("IDEquipoNuevo");
                if (idEquipoNuevoOriginal && !lineaRs.wasNull()) {
                    // Si efectivamente había un equipo nuevo, lo buscamos para obtener todos sus datos
                    const equipoNuevoOriginal = _buscarEquipoPorId('Equipo_Nuevo', idEquipoNuevoOriginal, conn);
                    
                    if (equipoNuevoOriginal) {
                        // A. El equipo nuevo original se marca como "Usado" para conservarlo en el historial de activos.
                        const updateNuevoStmt = conn.prepareStatement("UPDATE Equipo_Nuevo SET Estado = 'Usado' WHERE ID_Equipo = ?");
                        updateNuevoStmt.setInt(1, idEquipoNuevoOriginal);
                        updateNuevoStmt.executeUpdate();
                        updateNuevoStmt.close();
                        logMessage(`Equipo Nuevo ${idEquipoNuevoOriginal} actualizado a estado 'Usado'.`);

                        // B. Se crea una copia de este equipo en la tabla de Equipos Usados
                        const nuevoIdUsado = obtenerUltimoIDEquipoSQL(conn);
                        const resguardoInfo = getResponsableAndIdFromBDByPuesto(equipoNuevoOriginal.IDSUCURSAL, conn);

                        const insertStmt = conn.prepareStatement(
                          `INSERT INTO Equipo_Usado (ID_Equipo, Costo_Equipo, Fecha_Compra, Fecha_Formulario, Estado, Marca, Modelo, RAM, ROM, IMEI, IDSUCURSAL, IDEMPLEADO, Responsable, IDRESGUARDO) 
                           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
                        );
                        insertStmt.setInt(1, nuevoIdUsado);
                        insertStmt.setDouble(2, equipoNuevoOriginal.Costo_Equipo);
                        insertStmt.setString(3, formatDateForSql(new Date(equipoNuevoOriginal.Fecha_Compra)));
                        insertStmt.setString(4, formatDateForSql(new Date()));
                        insertStmt.setString(5, formData.estadoEquipoAnterior); // "Stock" o "Robado"
                        insertStmt.setString(6, equipoNuevoOriginal.Marca);
                        insertStmt.setString(7, equipoNuevoOriginal.Modelo);
                        insertStmt.setString(8, equipoNuevoOriginal.RAM);
                        insertStmt.setString(9, equipoNuevoOriginal.ROM);
                        insertStmt.setString(10, equipoNuevoOriginal.IMEI);
                        insertStmt.setInt(11, equipoNuevoOriginal.IDSUCURSAL);
                        insertStmt.setString(12, resguardoInfo ? resguardoInfo.id : null);
                        insertStmt.setString(13, resguardoInfo ? resguardoInfo.nombre : null);
                        insertStmt.setString(14, resguardoInfo ? resguardoInfo.id : null);
                        insertStmt.executeUpdate();
                        insertStmt.close();
                        logMessage(`Equipo Nuevo ${idEquipoNuevoOriginal} archivado como Equipo Usado ${nuevoIdUsado}.`);
                    }
                }
            }
            lineaRs.close();
            lineaOriginalStmt.close();
        }

        // 2. Actualizar los datos básicos de la línea celular
        const updateLineaSql = `UPDATE Telefonía_Telcel SET Tipo = ?, Responsable = ?, IDEMPLEADO = ?, IDSUCURSAL = ?, Extensión = ?, Notas = ? WHERE Teléfono = ?`;
        let pstmtLinea = conn.prepareStatement(updateLineaSql);
        pstmtLinea.setString(1, formData.tipo);
        pstmtLinea.setString(2, formData.responsable);
        pstmtLinea.setString(3, idEmpleado);
        pstmtLinea.setInt(4, idSucursal);
        pstmtLinea.setString(5, formData.extension);
        pstmtLinea.setString(6, formData.notas);
        pstmtLinea.setString(7, formData.telefono);
        pstmtLinea.executeUpdate();
        pstmtLinea.close();

        // 3. Si se vinculó un equipo usado nuevo, actualizar ese equipo y la línea
        if (equipoAVincular.id) {
            // 3a. Actualizar el registro del Equipo Usado que se está vinculando
            const updateEquipoSql = `UPDATE Equipo_Usado SET Estado = 'Reasignado', Responsable = ?, IDEMPLEADO = ?, IDSUCURSAL = ?, Fecha_Reasignacion = ?, Numero_Telefono = ?, Observaciones = ?, IDRESGUARDO = NULL WHERE ID_Equipo = ?`;
            let pstmtEquipo = conn.prepareStatement(updateEquipoSql);
            pstmtEquipo.setString(1, formData.responsable);
            pstmtEquipo.setString(2, idEmpleado);
            pstmtEquipo.setInt(3, idSucursal);
            pstmtEquipo.setString(4, formatDateForSql(equipoAVincular.fechaReasignacion));
            pstmtEquipo.setString(5, formData.telefono);
            pstmtEquipo.setString(6, equipoAVincular.observaciones);
            pstmtEquipo.setInt(7, parseInt(equipoAVincular.id));
            pstmtEquipo.executeUpdate();
            pstmtEquipo.close();

            // 3b. Obtener los datos completos del equipo recién vinculado para actualizar la línea
            const equipoVinculadoInfo = _buscarEquipoPorId('Equipo_Usado', equipoAVincular.id, conn);
            
            // 3c. Vincular este equipo a la línea en Telefonía_Telcel y actualizar sus datos
            const linkEquipoSql = `UPDATE Telefonía_Telcel SET IDEquipoUsado = ?, IDEquipoNuevo = NULL, IMEI = ?, Marca = ?, Modelo = ?, Equipo_ilimitado = ? WHERE Teléfono = ?`;
            let pstmtLink = conn.prepareStatement(linkEquipoSql);
            pstmtLink.setInt(1, parseInt(equipoAVincular.id));
            pstmtLink.setString(2, equipoAVincular.imei);
            pstmtLink.setString(3, equipoVinculadoInfo.Marca);
            pstmtLink.setString(4, equipoVinculadoInfo.Modelo);
            pstmtLink.setDouble(5, equipoVinculadoInfo.Costo_Equipo); // Actualizar costo
            pstmtLink.setString(6, formData.telefono);
            pstmtLink.executeUpdate();
            pstmtLink.close();
        }
        
        conn.commit(); // Si todo sale bien, confirmar todos los cambios en la BD
        return { success: true, message: "Línea modificada exitosamente." };

    } catch (e) {
        if (conn) conn.rollback(); // Si algo falla, revertir todos los cambios
        logMessage("Error en procesarModificacionLineaFormulario: " + e.message);
        return { success: false, message: e.message };
    } finally {
        if (conn) conn.close();
    }
}

// --- 4.5. Formulario: Modificar Equipo Usado (MEU) ---
/**
 * Procesa los datos enviados desde el formulario "Modificar Equipo Usado".
 * Inicia flujos de aprobación o actualiza directamente el registro en la BD.
 * @param {Object} formData Los datos del formulario como un objeto JavaScript.
 * @returns {Object} Un objeto con 'success' (boolean) y 'message' (string).
 */
function procesarMEUFormulario(formData) {
    let response = { success: false, message: "" };
    logMessage("Datos recibidos del formulario MEU: " + JSON.stringify(formData));

    let conn = null;

    try {
        conn = getJdbcConnection();
        // --- INICIO DE LA TRANSACCIÓN ---
        conn.setAutoCommit(false);

        const solicitanteEmail = Session.getActiveUser().getEmail();
        const today = new Date();
        today.setHours(0, 0, 0, 0);

        // 1. Obtener datos originales del equipo desde la BD
        const originalEquipoData = buscarEquipoPorIMEI(formData.imei);
        if (!originalEquipoData.success || !originalEquipoData.data) {
            throw new Error(originalEquipoData.message || `No se encontró el equipo con IMEI ${formData.imei}.`);
        }
        const originalData = originalEquipoData.data;
        const idEquipo = originalData.ID_Equipo; // El ID del equipo nunca cambia

        // 2. Definir variables finales que se guardarán en la BD, partiendo de los datos originales
        let estadoFinalDB = formData.nuevoEstado;
        let idEmpleadoFinal = originalData.IDEMPLEADO || null;
        let responsableFinal = originalData.Responsable || null;
        let idResguardoFinal = originalData.IDRESGUARDO || null;
        let idSucursalFinal = originalData.IDSUCURSAL;
        let comentariosFinal = formData.comentarios;
        let marcaFinal = originalData.Marca;
        let modeloFinal = originalData.Modelo;
        let numeroTelefonoFinal = formData.numeroTelefono || null;
        let fechaReasignacionFinal = formData.fechaReasignacion ? new Date(formData.fechaReasignacion) : null;

        // --- 3. LÓGICA CONDICIONAL POR NUEVO ESTADO ---
        switch (formData.nuevoEstado) {
            case 'Baja':
                estadoFinalDB = "Validación";
                responsableFinal = null;
                idEmpleadoFinal = null;
                idResguardoFinal = null;
                numeroTelefonoFinal = null;

                if (!formData.nombreSolicitanteBaja || !formData.razonBaja) {
                    throw new Error("El solicitante y la razón son obligatorios para la baja.");
                }
                
                const bajaSubject = `VALIDACIÓN DE BAJA DE EQUIPO (MODIFICACIÓN) - SUC: '${originalData.IDSUCURSAL_Name}'`;
                const bajaBody = `Buen día.\n\n` +
                               `${formData.nombreSolicitanteBaja} (${solicitanteEmail}) solicita validar la BAJA del siguiente equipo:\n\n` +
                               `Marca: ${originalData.Marca}\n` +
                               `Modelo: ${originalData.Modelo}\n` +
                               `IMEI: ${formData.imei}\n\n` +
                               `Razón de la Baja: ${formData.razonBaja}\n\n` +
                               `Por favor, Acepte o Deniegue la solicitud.`;

                const bajaButtons = [
                    { text: 'Aceptar', action: 'aprobarBajaEquipo', params: { idEquipo: idEquipo, solicitanteEmail: solicitanteEmail, razonBaja: formData.razonBaja, sucursal: originalData.IDSUCURSAL_Name, imei: formData.imei }, color: '#28a745' },
                    { text: 'Denegar', action: 'denegarBajaEquipo', params: { idEquipo: idEquipo, solicitanteEmail: solicitanteEmail, sucursal: originalData.IDSUCURSAL_Name, imei: formData.imei }, color: '#dc3545' }
                ];
                sendEmailWithButtons(ADMIN_EMAILS, bajaSubject, bajaBody, bajaButtons);
                response.message = `Solicitud de baja enviada para validación.`;
                break;

            case 'Vendido':
                estadoFinalDB = "Validación";
                responsableFinal = null;
                idEmpleadoFinal = null;
                idResguardoFinal = null;
                numeroTelefonoFinal = null;

                if (!formData.nombreSolicitanteVenta || !formData.personaVende) {
                    throw new Error("El solicitante y la persona a vender son obligatorios.");
                }

                const ventaSubject = `VALIDACIÓN DE VENTA DE EQUIPO (MODIFICACIÓN) - SUC: '${originalData.IDSUCURSAL_Name}'`;
                const ventaBody = `Buen día.\n\n` +
                                `${formData.nombreSolicitanteVenta} (${solicitanteEmail}) solicita validar la VENTA del siguiente equipo:\n\n` +
                                `Marca: ${originalData.Marca}\n` +
                                `Modelo: ${originalData.Modelo}\n` +
                                `IMEI: ${formData.imei}\n\n` +
                                `Se propone vender a: ${formData.personaVende}\n\n` +
                                `Por favor, Acepte o Deniegue la solicitud.`;

                const ventaButtons = [
                    { text: 'Aceptar', action: 'aprobarVentaEquipoStep1', params: { idEquipo: idEquipo, solicitanteEmail: solicitanteEmail, personaVende: formData.personaVende, sucursal: originalData.IDSUCURSAL_Name, imei: formData.imei }, color: '#28a745' },
                    { text: 'Denegar', action: 'denegarVentaEquipo', params: { idEquipo: idEquipo, solicitanteEmail: solicitanteEmail, sucursal: originalData.IDSUCURSAL_Name, imei: formData.imei }, color: '#dc3545' }
                ];
                sendEmailWithButtons(ADMIN_EMAILS, ventaSubject, ventaBody, ventaButtons);
                response.message = `Solicitud de venta enviada para validación.`;
                break;

            case 'Robado':
                responsableFinal = null;
                idEmpleadoFinal = null;
                idResguardoFinal = null;
                numeroTelefonoFinal = null;
                if (!comentariosFinal) {
                    comentariosFinal = "Favor de agregar anotaciones del Robo";
                }
                break;

            case 'Stock':
                responsableFinal = null;
                idEmpleadoFinal = null;
                idResguardoFinal = null;
                numeroTelefonoFinal = null;
                
                const resguardoInfo = getResponsableAndIdFromBDByPuesto(idSucursalFinal, conn);
                if (resguardoInfo && resguardoInfo.id) {
                    responsableFinal = resguardoInfo.nombre;
                    idEmpleadoFinal = resguardoInfo.id;
                    idResguardoFinal = resguardoInfo.id;
                } else {
                    throw new Error(`No se encontró un responsable de Stock en la sucursal.`);
                }
                
                if (formData.marca !== originalData.Marca || formData.modelo !== originalData.Modelo) {
                   if (!formData.comentarios) {
                     throw new Error("Debe justificar el cambio de marca y/o modelo en los comentarios.");
                   }
                   marcaFinal = formData.marca;
                   modeloFinal = formData.modelo;
                }
                break;
                
            case 'Reasignado':
                if (!formData.nuevaSucursal || !formData.nuevoResponsable || !fechaReasignacionFinal || !numeroTelefonoFinal) {
                    throw new Error("La nueva sucursal, el nuevo responsable, la fecha y el teléfono son obligatorios.");
                }
                responsableFinal = formData.nuevoResponsable;
                idEmpleadoFinal = getResponsableID(responsableFinal);
                idSucursalFinal = SUCURSAL_MAP[formData.nuevaSucursal];
                idResguardoFinal = null;
                if (!idEmpleadoFinal) { throw new Error(`No se encontró ID para el responsable '${responsableFinal}'.`); }
                if (!idSucursalFinal) { throw new Error(`La sucursal '${formData.nuevaSucursal}' es inválida.`); }
                if (isNaN(fechaReasignacionFinal.getTime()) || fechaReasignacionFinal > today) {
                    throw new Error("La fecha de reasignación debe ser válida y no futura.");
                }
                break;
        }

        // --- 4. ACTUALIZACIÓN EN SQL SERVER ---
        const updateSql = `
            UPDATE Equipo_Usado SET 
            Estado = ?, Observaciones = ?, Comentarios = ?, Numero_Telefono = ?, 
            IDSUCURSAL = ?, IDEMPLEADO = ?, Responsable = ?, IDRESGUARDO = ?,
            Marca = ?, Modelo = ?, Fecha_Reasignacion = ?
            WHERE IMEI = ?`;
        const pstmt = conn.prepareStatement(updateSql);
        
        pstmt.setString(1, estadoFinalDB);
        pstmt.setString(2, formData.observaciones);
        pstmt.setString(3, comentariosFinal);
        pstmt.setString(4, numeroTelefonoFinal);
        pstmt.setObject(5, idSucursalFinal);
        pstmt.setString(6, idEmpleadoFinal);
        pstmt.setString(7, responsableFinal);
        pstmt.setString(8, idResguardoFinal);
        pstmt.setString(9, marcaFinal);
        pstmt.setString(10, modeloFinal);
        pstmt.setString(11, formatDateForSql(fechaReasignacionFinal));
        pstmt.setString(12, formData.imei);

        pstmt.executeUpdate();
        
        conn.commit();
        logMessage(`Transacción completada. Equipo IMEI ${formData.imei} actualizado en SQL.`);

        // (Aquí puedes añadir la lógica para registrar el cambio en la hoja "MEU")

        response.success = true;
        if (!response.message) {
            response.message = `Registro actualizado exitosamente para IMEI: ${formData.imei}.`;
        }

    } catch (e) {
        if (conn) { try { conn.rollback(); } catch (rollError) { logMessage("Error al revertir: " + rollError.message); }}
        response.message = e.message;
        logMessage("Error en procesarMEUFormulario: " + e.message + " Stack: " + e.stack);
        response.success = false;
    } finally {
        if (conn) { try { conn.setAutoCommit(true); conn.close(); } catch (finalError) { logMessage("Error al cerrar conexión: " + finalError.message); }}
    }
    return response;
}

// Helper function to find a row in a sheet by IMEI (for MEU log updates)
function findSheetRowByIMEI(sheet, imei) {
    const range = sheet.getDataRange();
    const values = range.getValues();
    const headerRow = values[0];
    const imeiColIndex = headerRow.indexOf("IMEI del equipo"); 
    
    if (imeiColIndex === -1) {
        logMessage("La hoja MEU no tiene la columna 'IMEI del equipo'. No se puede buscar la fila.");
        return null;
    }

    for (let i = 1; i < values.length; i++) {
        if (values[i][imeiColIndex] == imei) {
            // Return the entire row data and its 1-based index
            return { data: values[i], rowIndex: i + 1 };
        }
    }
    return null; // Not found
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

/**
function aprobarVentaEquipoStep1(idEquipo, solicitanteEmail, personaVende, sucursal, imei) {
    // Redirige a un formulario HTML simple para pedir el monto de venta
    const scriptUrlBase = ScriptApp.getService().getUrl();
    // Asegurarse de pasar el IMEI a la siguiente etapa
    const redirectUrl = `${scriptUrlBase}?form=aprobarVentaForm&idEquipo=${idEquipo}&solicitanteEmail=${encodeURIComponent(solicitanteEmail)}&personaVende=${encodeURIComponent(personaVende)}&sucursal=${encodeURIComponent(sucursal)}&imei=${encodeURIComponent(imei)}`;
    
    return HtmlService.createHtmlOutput(`<script>window.top.location.href = '${redirectUrl}';</script>`);
} */

/**
 * Segunda etapa de aprobación de venta: actualiza el estado, comentarios y notifica al solicitante.
 * Llamada desde el formulario HTML de aprobación de venta.
 * @param {Object} formData Datos del formulario, incluyendo idEquipo, montoVenta, imei, etc.
 */
function aprobarVentaEquipoStep2(formData) {
    // Se extraen los parámetros del objeto formData
    const idEquipo = parseInt(formData.idEquipo);
    const montoVenta = parseFloat(formData.montoVenta) || 0;
    const solicitanteEmail = formData.solicitanteEmail;
    const personaVende = formData.personaVende;
    const sucursal = formData.sucursal;
    const imei = formData.imei; // CORRECCIÓN: Ahora se recibe el IMEI correctamente

    let confirmationTitle = "Venta Aprobada";
    let confirmationMessage = `El equipo con IMEI ${imei} ha sido marcado como "Vendido" y se ha enviado notificación.`;
    let success = true;
    let conn = null;

    try {
        // Validación crucial del IMEI
        if (!imei) {
            throw new Error("El IMEI del equipo no fue proporcionado. No se puede procesar la venta.");
        }

        if (isNaN(montoVenta) || montoVenta <= 0) {
            throw new Error("Monto de venta inválido. Debe ser un número mayor a cero.");
        }

        conn = getJdbcConnection();
        const sheet = getSheet("RECU");

        // 1. Actualizar en SQL Server usando IMEI
        const updateSql = `UPDATE Equipo_Usado SET Estado = ?, Comentarios = ? WHERE IMEI = ? AND Estado = 'Validación'`;
        const pstmt = conn.prepareStatement(updateSql);
        pstmt.setString(1, "Vendido");
        pstmt.setString(2, `El equipo fue vendido a ${personaVende} con costo de $${montoVenta.toFixed(2)} pesos.`);
        pstmt.setString(3, imei);
        
        const rowsAffected = pstmt.executeUpdate();
        pstmt.close();

        if (rowsAffected > 0) {
            logMessage(`Venta aprobada en SQL para Equipo IMEI: ${imei}`);
        } else {
            logMessage(`La venta para el IMEI ${imei} ya había sido procesada en la BD o no se encontró en estado 'Validación'.`);
            confirmationMessage = `La venta para el equipo con IMEI ${imei} ya fue procesada anteriormente.`;
        }

        // 2. Actualizar en Google Sheet
        const range = sheet.getDataRange();
        const values = range.getValues();
        const headerRow = values[0];
        const imeiColIndex = headerRow.indexOf("IMEI");
        const estadoColIndex = headerRow.indexOf("Estado del equipo");
        const comentariosColIndex = headerRow.indexOf("Comentarios");

        for (let i = 1; i < values.length; i++) {
            if (values[i][imeiColIndex] == imei && values[i][estadoColIndex] === "Validación") {
                sheet.getRange(i + 1, estadoColIndex + 1).setValue("Vendido");
                sheet.getRange(i + 1, comentariosColIndex + 1).setValue(`El equipo fue vendido a ${personaVende} con costo de $${montoVenta.toFixed(2)} pesos.`);
                break;
            }
        }

        // 3. Enviar correo de confirmación
        const confirmSubject = `Solicitud de Venta de Equipo APROBADA (${sucursal})`;
        const confirmBody = `Su solicitud de venta para el equipo con IMEI ${imei} ha sido APROBADA.\n\n` +
                          `El equipo fue vendido a ${personaVende} por $${montoVenta.toFixed(2)} pesos.`;
        MailApp.sendEmail(solicitanteEmail, confirmSubject, confirmBody);

    } catch (e) {
        logMessage("Error al aprobar venta (Step 2): " + e.message);
        confirmationTitle = "Error al Aprobar Venta";
        confirmationMessage = `Ha ocurrido un error: ${e.message}`;
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
