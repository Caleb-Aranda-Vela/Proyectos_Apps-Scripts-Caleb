//BD.gs

//
function spreadsheet() {
  return SpreadsheetApp.openById("1h4TxPJHZ8pynph3J6q2h4FnDyOnG0Uye3VrYqFriPCg");
}

//
function sheetMain() {
  return spreadsheet().getSheetByName("BD");
}

//
function sheetLog() {
  return spreadsheet().getSheetByName("Log");
}

//
function importar() {
  Logger.log("importar");
  // Hacer la conexion a SQL Server.
  // Replace the variables in this block with real values.
  var address = 'gw.hemoeco.com:5300'; // '187.189.51.7:5300';
  var user = '';
  var pwd = '';
  var db = 'IT_Rentas';

  var dbUrl = 'jdbc:sqlserver://' + address + ';databaseName=' + db;
  
  try {
    // Leer hasta 1000 registros de la tabla
    leerTabla(dbUrl, user, pwd);
  }
  catch(ex) {
    Logger.log(ex.message);
    writeToLog();  // Dejar huella de cualquier error que haya acontecido.
  }
  
  // Este log era sólo para verificar que estuviera actualizando correctamente. (Se pidió activarlo para monitorear que si se esté actualizando)
  writeToLog();  // Por el momento escribimos siempre el log para monitorear las actualizaciones
}

/**
 * Se conecta a la base de datos, ejecuta dos consultas y escribe los resultados
 * en la misma hoja de cálculo. La primera consulta se escribe a partir de la
 * columna A y la segunda a partir de la columna J.
 */
function leerTabla(dbUrl, user, pwd) {
  var conn = Jdbc.getConnection(dbUrl, user, pwd);
  var start = new Date();
  var sheet = sheetMain();
  
  // Limpiar todo el contenido de la hoja antes de escribir nuevos datos.
  sheet.clearContents();

  // --- INICIO: PROCESO PARA LA PRIMERA CONSULTA ---
  var stmt1 = conn.createStatement();
  var query1 = "SELECT CE.NOMBRECOMPLETO, CE.IDEMPLEADO, CE.IDPUESTO, CE.ACTIVO, CE.TELEFONOOFICINA, CE.SUCURSALNOMINA, CE.IDSUCURSAL, CS.INICIALES AS [Sucursal] from IT_Rentas.dbo.CataEmpleados as CE LEFT JOIN IT_Rentas.dbo.CataSucursales CS ON CE.SUCURSALNOMINA = CS.idSucursal WHERE CE.ACTIVO = 1 AND CE.NOMBRECOMPLETO NOT LIKE '%ADM%'AND CE.NOMBRECOMPLETO NOT LIKE '%CAN%'AND CE.NOMBRECOMPLETO NOT LIKE '%TIJ%'AND CE.NOMBRECOMPLETO NOT LIKE '%GDL%'AND CE.NOMBRECOMPLETO NOT LIKE '%MXL%'AND CE.NOMBRECOMPLETO NOT LIKE '%SJD%'AND CE.NOMBRECOMPLETO NOT LIKE '%MEX%'AND CE.NOMBRECOMPLETO NOT LIKE '%MTY%' ORDER BY CE.NOMBRECOMPLETO";
  
  var results1 = stmt1.executeQuery(query1);
  var metaData1 = results1.getMetaData();
  var numCols1 = metaData1.getColumnCount();
  
  // Escribir los títulos de la primera consulta
  var titulos1 = [];
  for (var col = 0; col < numCols1; col++) {
    titulos1.push(metaData1.getColumnName(col + 1));
  }
  sheet.getRange(1, 1, 1, titulos1.length).setValues([titulos1]);

  // Leer y preparar los datos de la primera consulta
  var rows1 = [];
  while (results1.next()) {
    var row = [];
    for (var col = 0; col < numCols1; col++) {
      row.push(results1.getString(col + 1));
    }
    rows1.push(row);
  }

  // Escribir los datos de la primera consulta en la hoja
  if (rows1.length > 0) {
      sheet.getRange(2, 1, rows1.length, numCols1).setValues(rows1);
  }

  results1.close();
  stmt1.close();
  // --- FIN: PROCESO PARA LA PRIMERA CONSULTA ---


  // --- INICIO: PROCESO PARA LA SEGUNDA CONSULTA ---
  var stmt2 = conn.createStatement();
  var query2 = "select TT.Región, TT.Teléfono, TT.Estado, TT.IDSUCURSAL, CS.INICIALES AS [Sucursal] from Soporte.dbo.Telefonía_Telcel AS TT LEFT JOIN IT_Rentas.dbo.CataSucursales CS ON TT.IDSUCURSAL = CS.idSucursal where TT.Estado = 'Activo' and TT.Tipo = 'SmartPhone'";
  
  var results2 = stmt2.executeQuery(query2);
  var metaData2 = results2.getMetaData();
  var numCols2 = metaData2.getColumnCount();

  // Escribir los títulos de la segunda consulta a partir de la columna J (columna 10)
  var titulos2 = [];
  for (var col = 0; col < numCols2; col++) {
    titulos2.push(metaData2.getColumnName(col + 1));
  }
  sheet.getRange(1, 10, 1, titulos2.length).setValues([titulos2]);

  // Leer y preparar los datos de la segunda consulta
  var rows2 = [];
  while (results2.next()) {
    var row = [];
    for (var col = 0; col < numCols2; col++) {
      row.push(results2.getString(col + 1));
    }
    rows2.push(row);
  }

  // Escribir los datos de la segunda consulta en la hoja a partir de la columna J
  if (rows2.length > 0) {
    sheet.getRange(2, 10, rows2.length, numCols2).setValues(rows2);
  }

  results2.close();
  stmt2.close();
  // --- FIN: PROCESO PARA LA SEGUNDA CONSULTA ---

  conn.close();
  
  var end = new Date();
  Logger.log('Time elapsed: %sms', end - start);
}

//
function writeToLog(){
  // write the Drive file link to the Drive File Report Tab for safe keeping/logging purposes
  var reportSheet = sheetLog();
  
  // Error log
  reportSheet.appendRow([new Date(),Session.getActiveUser().getEmail(), Logger.getLog()]);
}

//
function onOpen() {
  // Esta función se manda ejecutar desde los activadores del proyecto cada vez que se abre la hoja
  var spreadsheet = SpreadsheetApp.getActive();
  if (spreadsheet != null) {
    var menuItems = [
      {name: '¡Actualizar ahora!', functionName: 'importar'},
    ];
    spreadsheet.addMenu('Acciones especiales', menuItems);
  }
}
