/**
 * Ejecuta esta función manualmente una vez para actualizar la columna TD101_NOMBRE_C
 */
function ejecutarActualizacionManualNombres() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('TD101_MAIN');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // Localizar índices de columnas
  const idx1 = headers.indexOf('TD101NOMBRE1');
  const idx2 = headers.indexOf('TD101NOMBRE2');
  const idxC = headers.indexOf('TD101NOMBREC');

  if (idx1 === -1 || idx2 === -1 || idxC === -1) {
    Logger.log("Error: No se encontraron las columnas necesarias.");
    return;
  }

  const updates = [];

  // Procesar cada fila (saltando el encabezado)
  for (let i = 1; i < data.length; i++) {
    const val1 = data[i][idx1];
    const val2 = data[i][idx2];

    const n1Procesado = transformarNombreEspecial(val1, false);
    const n2Procesado = transformarNombreEspecial(val2, true);
    
    const nombreCombinado = (n1Procesado + " " + n2Procesado).trim();
    
    // Guardamos la actualización para la celda específica [fila, columna]
    // i + 1 porque las filas en Sheets empiezan en 1
    // idxC + 1 porque las columnas empiezan en 1
    sheet.getRange(i + 1, idxC + 1).setValue(nombreCombinado);
  }
  
  Logger.log("Proceso completado con éxito.");
}

/**
 * Tu lógica de transformación con excepciones
 */
function transformarNombreEspecial(valor, esNombre2) {
  if (!valor) return "";
  let str = String(valor).trim();
  
  // Regla especial para apóstrofe en apellidos (Ej: O'connor)
  if (esNombre2 && str.includes("'")) {
    return str.split("'").map(p => p.charAt(0).toUpperCase() + p.slice(1).toLowerCase()).join("'");
  }

  let partes = str.split(/\s+/);
  if (partes.length === 0) return "";

  const conectores = ["de", "del", "los", "la", "las"];
  let resultado = "";
  let i = 0;

  // Primer nombre
  resultado += partes[i].charAt(0).toUpperCase() + partes[i].slice(1).toLowerCase();
  i++;

  // Conectores e inicial
  while (i < partes.length) {
    let palabraActual = partes[i].toLowerCase();
    if (conectores.includes(palabraActual)) {
      resultado += " " + palabraActual;
      i++;
    } else {
      resultado += " " + partes[i].charAt(0).toUpperCase() + ".";
      break; 
    }
  }
  return resultado;
}

/**
 * Procesa todas las hojas del documento DATA_SS_ID y extrae los encabezados
 * para consolidarlos en una tabla dentro de 'hojaX'.
 */
function generateHeadersInventory() {
  const TARGET_SHEET_NAME = 'hojaX';
  
  try {
    const ss = SpreadsheetApp.openById(DATA_SS_ID);
    const sheets = ss.getSheets();
    let inventoryData = [];

    // Iterar por cada hoja del documento
    sheets.forEach(sheet => {
      const sheetName = sheet.getName();
      
      // Evitar procesar la hoja de destino para no crear bucles de datos
      if (sheetName === TARGET_SHEET_NAME) return;

      // Obtener la primera fila (encabezados)
      // getRange(fila, columna, numFilas, numColumnas)
      const lastColumn = sheet.getLastColumn();
      
      if (lastColumn > 0) {
        const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
        
        headers.forEach(header => {
          if (header !== "") {
            inventoryData.push([sheetName, header]);
          }
        });
      }
    });

    // Gestión de la hoja de destino 'hojaX'
    let targetSheet = ss.getSheetByName(TARGET_SHEET_NAME);
    if (!targetSheet) {
      targetSheet = ss.insertSheet(TARGET_SHEET_NAME);
    } else {
      targetSheet.clearContents(); // Limpiar contenido previo
    }

    // Insertar encabezados de la nueva tabla
    targetSheet.getRange(1, 1, 1, 2).setValues([["Hoja", "Columna"]]);
    targetSheet.getRange(1, 1, 1, 2).setFontWeight("bold");

    // Insertar los datos recolectados
    if (inventoryData.length > 0) {
      targetSheet.getRange(2, 1, inventoryData.length, 2).setValues(inventoryData);
    }

    Logger.log("Inventario generado con éxito en " + TARGET_SHEET_NAME);
    
  } catch (e) {
    Logger.log("Error: " + e.toString());
  }
}
