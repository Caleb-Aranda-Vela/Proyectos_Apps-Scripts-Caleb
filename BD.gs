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
  var user = 'reportes';
  var pwd = 'R3p0rt3s';
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

// Read up to 1000 rows of data from the table and log them.
function leerTabla(dbUrl, user, pwd) {
  var conn = Jdbc.getConnection(dbUrl, user, pwd);

  var start = new Date();
  
  var stmt = conn.createStatement();
  
  var query = "SELECT NOMBRECOMPLETO, IDEMPLEADO, IDPUESTO, ACTIVO, TELEFONOOFICINA,SUCURSALNOMINA, IDSUCURSAL from IT_Rentas.dbo.CataEmpleados WHERE ACTIVO = 1 AND NOMBRECOMPLETO NOT LIKE '%ADM%'AND NOMBRECOMPLETO NOT LIKE '%CAN%'AND NOMBRECOMPLETO NOT LIKE '%TIJ%'AND NOMBRECOMPLETO NOT LIKE '%GDL%'AND NOMBRECOMPLETO NOT LIKE '%MXL%'AND NOMBRECOMPLETO NOT LIKE '%SJD%'AND NOMBRECOMPLETO NOT LIKE '%MEX%'AND NOMBRECOMPLETO NOT LIKE '%MTY%' ORDER BY NOMBRECOMPLETO";     
  // stmt.setMaxRows(1000);  
  var results = stmt.executeQuery(query);
  var metaData = results.getMetaData();
  var numCols = metaData.getColumnCount();
  var sheet = sheetMain();
  var titulos = [];
  
  // Llenar la columna de titulos
  for (var col = 0; col < numCols; col++) {
    titulos.push(metaData.getColumnName(col+1));
  }
  
  // Limpiar e lcontenido de la hoja
  sheet.clearContents();
  
  // Poner los titulos de las columnas
  var destRange = sheet.getRange(1, 1, 1, titulos.length);
  destRange.setValues([titulos]);
  
  // Leer los datos
  var numRows = 0;
  var rows = [];
  var row = [];
  while (results.next()) {
    row = [];
    for (var col = 0; col < numCols; col++) {
      row.push(results.getString(col + 1));
    }
    rows.push(row);
    numRows++;
  }

  // Guardar los datos en la hoja.
  destRange = sheet.getRange(2, 1, numRows, titulos.length);
  destRange.setValues(rows);

  results.close();
  stmt.close();

  // Logger.log(rows);
  
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
