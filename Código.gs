// Code.gs (Este es el archivo principal de tu proyecto de Apps Script)

// --- CONFIGURACIÓN DE LA HOJA DE CÁLCULO Y BASE DE DATOS ---

// ID de la hoja de cálculo de Google (Nueva URL proporcionada)
const SPREADSHEET_ID = "1h4TxPJHZ8pynph3J6q2h4FnDyOnG0Uye3VrYqFriPCg"; 

// Información para conectar con la base de datos SQL Server
const DB_ADDRESS = 'gw.hemoeco.com:5300';
const DB_USER = 'caleb.aranda';
const DB_PWD = '5AA3hmq8BfJFkrISTWgJsA==';
const DB_NAME = 'Soporte_Pruebas';
const DB_URL = 'jdbc:sqlserver://' + DB_ADDRESS + ';databaseName=' + DB_NAME;

// URL para la redirección a la aplicación "Principal" (tu menú principal)
const REDIRECTION_URL = "https://script.google.com/a/macros/hemoeco.com/s/AKfycbzQpqK85N-s6l4Qz0dMQGqD1ePr8O-PT5eJ87wYyXcLsuK_GU_lJmv-j38-xp8kbKOxnQ/exec";

// Mapeo de nombres de sucursales (formulario) a códigos (base de datos)
const SUCURSAL_MAP = {
    "ADM": "01",
    "GDL": "05",
    "MXL": "06",
    "MEX": "08",
    "TIJ": "04",
    "CAN": "03",
    "SJD": "07",
    "MTY": "09"
};

// --- ENCABEZADOS ESPECÍFICOS PARA CADA HOJA DE RESPUESTAS ---
// ¡IMPORTANTE! Asegúrate de que las columnas en tus hojas de cálculo coincidan EXACTAMENTE con estos arrays, incluyendo el orden.

// Encabezados para la hoja "RECU" (Registrar Equipo Usado)
const RECU_SHEET_HEADERS = [
    "Marca temporal",
    "Dirección de correo electrónico",
    "Costo del Equipo",
    "Fecha de compra de Equipo",
    "Fecha de Recolección",
    "Fecha de Reasignación",
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

// Placeholder para los encabezados de las otras hojas.
// DEBERÁS DEFINIR ESTOS ARRAYS CON LAS COLUMNAS EXACTAS DE CADA HOJA CUANDO DESARROLLES ESOS FORMULARIOS.
// He añadido al menos "Marca temporal" y "EJECUTADO" como base.
const ALyE_SHEET_HEADERS = ["Marca temporal", "Campo1_ALyE", "Campo2_ALyE", "EJECUTADO"];
const RLyE_SHEET_HEADERS = ["Marca temporal", "Campo1_RLyE", "Campo2_RLyE", "EJECUTADO"];
const ML_SHEET_HEADERS = ["Marca temporal", "Campo1_ML", "Campo2_ML", "EJECUTADO"];
const MEU_SHEET_HEADERS = ["Marca temporal", "Campo1_MEU", "Campo2_MEU", "EJECUTADO"];


// --- FUNCIONES BÁSICAS DE APPS SCRIPT ---

/**
 * Sirve el archivo HTML del formulario según el parámetro 'form' en la URL.
 * Se invoca automáticamente cuando se accede a la URL de la aplicación web de Apps Script.
 * Ejemplo de URL: https://script.google.com/macros/s/.../exec?form=registrarEquipoUsado
 * @param {GoogleAppsScript.Events.DoGet} e El objeto de evento que contiene los parámetros de la URL.
 */
function doGet(e) {
    // Si no se especifica un parámetro 'form', por defecto carga 'registrarEquipoUsado'
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
        default:
            // Si el nombre del formulario es inválido, se carga el formulario predeterminado y se registra una advertencia.
            htmlFileToServe = 'registrarEquipoUsado';
            logMessage("Advertencia: Nombre de formulario inválido recibido: '" + formName + "'. Se carga 'registrarEquipoUsado'.");
            break;
    }

    return HtmlService.createHtmlOutputFromFile(htmlFileToServe)
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


// --- FUNCIONES DE INTERACCIÓN CON SQL SERVER ---

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
        logMessage("Siguiente ID_Equipo para generar: " + nextId);
        return nextId;
    } catch (e) {
        logMessage("Error al obtener el último ID_Equipo de SQL Server: " + e.message);
        throw new Error("Error al obtener ID de equipo: " + e.message);
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
        logMessage(`IMEI ${imei} es único: ${isUnique}`);
        return isUnique;
    } catch (e) {
        logMessage("Error al validar IMEI en SQL Server: " + e.message);
        throw new Error("Error al validar IMEI: " + e.message);
    } finally {
        if (results) results.close();
        if (pstmt) pstmt.close();
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


// --- PROCESAMIENTO ESPECÍFICO DE CADA FORMULARIO ---

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

    try {
        const sheet = getSheet("RECU"); // Hoja específica para RECU
        
        conn = getJdbcConnection();

        const fechaFormulario = new Date();
        const idEquipo = obtenerUltimoIDEquipoSQL(conn);

        const imei = formData.imei ? formData.imei.trim() : '';
        if (!imei) {
            response.message = "El campo IMEI es obligatorio.";
            logMessage("Error: " + response.message);
            return response;
        }
        if (!validarIMEIUnicoSQL(imei, conn)) {
            response.message = `IMEI '${imei}' ya existe en la base de datos.`;
            logMessage("Error: " + response.message);
            return response;
        }

        const idSucursalBD = SUCURSAL_MAP[formData.idSucursal] || '';
        if (!idSucursalBD) {
            response.message = "ID de Sucursal inválido.";
            logMessage("Error: " + response.message);
            return response;
        }

        // --- INSERCIÓN DIRECTA EN SQL SERVER ---
        // Se asegura que la consulta INSERT coincida con la tabla `Equipo_Usado` proporcionada.
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
            
            const costoEquipo = parseFloat(formData.costoEquipo) || 0;
            const fechaCompra = formatDateForSql(formData.fechaCompra);
            const fechaRecoleccion = formatDateForSql(formData.fechaRecoleccion);
            const fechaReasignacion = formatDateForSql(formData.fechaReasignacion);
            const fechaFormularioSql = formatDateForSql(fechaFormulario);

            pstmt.setInt(1, idEquipo);
            pstmt.setDouble(2, costoEquipo);
            pstmt.setString(3, fechaCompra);
            pstmt.setString(4, fechaRecoleccion);
            pstmt.setString(5, fechaReasignacion);
            pstmt.setString(6, fechaFormularioSql);
            pstmt.setString(7, formData.estado);
            pstmt.setString(8, formData.observaciones);
            pstmt.setString(9, formData.marca);
            pstmt.setString(10, formData.modelo);
            pstmt.setString(11, formData.ram);
            pstmt.setString(12, formData.rom);
            pstmt.setString(13, imei);
            pstmt.setString(14, null); // Numero_Telefono (omitido en este formulario)
            pstmt.setString(15, idSucursalBD);
            pstmt.setString(16, null); // IDEMPLEADO (omitido en este formulario)
            pstmt.setString(17, null); // Responsable (omitido en este formulario)
            pstmt.setString(18, null); // IDRESGUARDO (se determinará más adelante, actualmente NULL)
            pstmt.setString(19, null); // IDAUTORIZA (se determinará más adelante, actualmente NULL)
            pstmt.setString(20, formData.comentarios);
            pstmt.setString(21, null); // Documentacion (omitido en este formulario)

            pstmt.executeUpdate();
            logMessage(`Equipo con ID ${idEquipo} insertado exitosamente en SQL Server.`);
            response.success = true;
            response.message = `Equipo con ID ${idEquipo} insertado en SQL y registrado en hoja de cálculo.`;

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
                case "Dirección de correo electrónico": rowData.push(Session.getActiveUser().getEmail()); break;
                case "Costo del Equipo": rowData.push(parseFloat(formData.costoEquipo) || 0); break;
                case "Fecha de compra de Equipo": rowData.push(formData.fechaCompra ? new Date(formData.fechaCompra) : ''); break;
                case "Fecha de Recolección": rowData.push(formData.fechaRecoleccion ? new Date(formData.fechaRecoleccion) : ''); break;
                case "Fecha de Reasignación": rowData.push(formData.fechaReasignacion ? new Date(formData.fechaReasignacion) : ''); break;
                case "Estado del equipo": rowData.push(formData.estado); break;
                case "Observaciones": rowData.push(formData.observaciones); break;
                case "Marca": rowData.push(formData.marca); break;
                case "Modelo": rowData.push(formData.modelo); break;
                case "Memoria RAM": rowData.push(formData.ram); break;
                case "Almacenamiento (Memoria ROM)": rowData.push(formData.rom); break;
                case "IMEI": rowData.push(imei); break;
                case "Sucursal": rowData.push(formData.idSucursal); break;
                case "IDEquipo": rowData.push(idEquipo); break;
                case "Error": rowData.push(""); break;
                case "IDSucursal": rowData.push(idSucursalBD); break;
                case "EJECUTADO": rowData.push("SI"); break;
                case "Comentarios": rowData.push(formData.comentarios); break;
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

/**
 * Placeholder para procesar el formulario "Dar de alta Línea y Equipo".
 * DEBERÁS IMPLEMENTAR LA LÓGICA ESPECÍFICA PARA ESTE FORMULARIO, INCLUYENDO INTERACCIÓN CON SQL SI ES NECESARIO.
 * Asegúrate de usar `getSheet("ALyE")` y `ALyE_SHEET_HEADERS`.
 */
function procesarALyEFormulario(formData) {
    let response = { success: false, message: "" };
    logMessage("Datos recibidos del formulario ALyE: " + JSON.stringify(formData));
    try {
        const sheet = getSheet("ALyE"); // Hoja específica para ALyE
        // Implementa aquí la lógica para guardar en SQL y luego en la hoja "ALyE"
        // Ejemplo simplificado para demostración:
        const rowData = [];
        ALyE_SHEET_HEADERS.forEach(header => {
            switch (header) {
                case "Marca temporal": rowData.push(new Date()); break;
                case "EJECUTADO": rowData.push("SI"); break;
                default: rowData.push(`Valor para ${header}: ${formData[header] || ''}`); // Ajusta según tus campos reales
            }
        });
        sheet.appendRow(rowData);
        response.success = true;
        response.message = "Datos de Dar de alta Línea y Equipo procesados (simulado).";
        logMessage(response.message);
    } catch (e) {
        response.message = "Error al procesar formulario ALyE: " + e.message;
        logMessage("Error en procesarALyEFormulario: " + e.message);
    }
    return response;
}

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


// --- FUNCIÓN PARA ENVIAR DATOS DE LA HOJA DE CÁLCULO A SQL SERVER (Rol Secundario) ---
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
            fechaReasignacion: getColIndex("Fecha de Reasignación"),
            estado: getColIndex("Estado del equipo"),
            observaciones: getColIndex("Observaciones"),
            marca: getColIndex("Marca"),
            modelo: getColIndex("Modelo"),
            ram: getColIndex("Memoria RAM"),
            rom: getColIndex("Almacenamiento (Memoria ROM)"),
            imei: getColIndex("IMEI"),
            sucursalNombre: getColIndex("Sucursal"),
            idEquipo: getColIndex("IDEquipo"),
            errorCol: getColIndex("Error"),
            idSucursalBD: getColIndex("IDSucursal"),
            ejecutadoCol: getColIndex("EJECUTADO"),
            comentarios: getColIndex("Comentarios")
        };
        
        // Preparar la consulta INSERT para SQL Server
        const insertQuery = `
            INSERT INTO Equipo_Usado (
                ID_Equipo, Costo_Equipo, Fecha_Compra, Fecha_Recoleccion, Fecha_Reasignacion,
                Fecha_Formulario, Estado, Observaciones, Marca, Modelo, RAM, ROM, IMEI,
                Numero_Telefono, IDSUCURSAL, IDEMPLEADO, Responsable, Resguardo, Autoriza,
                Comentarios, Documentacion, TEXTO
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,?)
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
                const fechaReasignacion = formatDateForSql(row[colMap.fechaReasignacion]);
                const fechaFormulario = formatDateForSql(row[colMap.marcaTemporal]);

                const estado = row[colMap.estado] || '';
                const observaciones = row[colMap.observaciones] || '';
                const marca = row[colMap.marca] || '';
                const modelo = row[colMap.modelo] || '';
                const ram = row[colMap.ram] || '';
                const rom = row[colMap.rom] || '';
                const imei = row[colMap.imei] || '';
                const idSucursal = row[colMap.idSucursalBD] || '';
                const comentarios = row[colMap.comentarios] || ''; // Obtener comentarios de la columna mapeada

                // Campos que se omiten en el formulario, se insertan como NULL
                const numeroTelefono = null;
                const idEmpleado = null;
                const responsable = null;
                const idResguardo = null;
                const idAutoriza = null;
                const documentacion = null;

                // Setear parámetros para la consulta preparada
                pstmt.setInt(1, idEquipo);
                pstmt.setDouble(2, costoEquipo);
                pstmt.setString(3, fechaCompra);
                pstmt.setString(4, fechaRecoleccion);
                pstmt.setString(5, fechaReasignacion);
                pstmt.setString(6, fechaFormularioSql);
                pstmt.setString(7, formData.estado);
                pstmt.setString(8, formData.observaciones);
                pstmt.setString(9, formData.marca);
                pstmt.setString(10, formData.modelo);
                pstmt.setString(11, formData.ram);
                pstmt.setString(12, formData.rom);
                pstmt.setString(13, imei);
                pstmt.setString(14, null); // Numero_Telefono (omitido en este formulario)
                pstmt.setString(15, idSucursalBD);
                pstmt.setString(16, null); // IDEMPLEADO (omitido en este formulario)
                pstmt.setString(17, null); // Responsable (omitido en este formulario)
                pstmt.setString(18, null); // IDRESGUARDO (se determinará más adelante)
                pstmt.setString(19, null); // IDAUTORIZA (se determinará más adelante)
                pstmt.setString(20, formData.comentarios);
                pstmt.setString(21, null); // Documentacion (omitido en este formulario)
                pstmt.setString(22, null);

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
