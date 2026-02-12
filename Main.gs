/**
 * @fileoverview Main.gs - Lógica de Configuración y Campos Maestros
 */

const DATA_SS_ID = '1tREeWG6QugdcGFfC8uy3vSG7Q6DSjVpuVBEtdR094eQ'; 
const CONFIG_SS_ID = '1s4N_pwkwPHMWXlNqcG9dQXm9_yg2jdKImkZdmghKIbs'; 
const CONFIG_SHEET_NAME = 'ConfigViewTB';

function doGet(e) {
  try {
    const action = e.parameter.action;
    const appTienda = e.parameter.appTienda;
    const userTienda = e.parameter.userTienda || 'DEFAULT';

    if (action === "getTableFriendlyNames") {
      const result = getTableFriendlyNames(appTienda || userTienda);
      return createJsonResponse(result);
    }

    const tableName = e.parameter.tableName || e.parameter.sheet;
    if (!tableName) throw new Error("Parámetro 'tableName' omitido.");

    const ignoreVisibility = e.parameter.ignoreVisibility === 'true'; 

    const ssData = SpreadsheetApp.openById(DATA_SS_ID);
    const ssConfig = SpreadsheetApp.openById(CONFIG_SS_ID);
    
    const configSheet = ssConfig.getSheetByName(CONFIG_SHEET_NAME);
    const dataSheet = ssData.getSheetByName(tableName);

    if (!dataSheet) throw new Error("La tabla '" + tableName + "' no existe.");
    if (!configSheet) throw new Error("La hoja de configuración no existe.");

    // 1. Obtener Configuración de Columnas
    const configRows = configSheet.getDataRange().getValues();
    const configData = configRows.slice(1);
    const configMap = {};
    const availableTables = [];
    const fullConfigForFrontend = [];

    configData.forEach(row => {
      const rowAppTienda = String(row[0]).trim();
      const nombreTabla = String(row[1]).trim();
      
      if (rowAppTienda === userTienda && !availableTables.includes(nombreTabla)) {
        availableTables.push(nombreTabla);
      }

      if (nombreTabla === tableName) {
        const idColumna = String(row[2]).trim();
        const upperColId = idColumna.toUpperCase();
        const tablePrefix = tableName.split('_')[0].toUpperCase();
        const esID = (upperColId === `${tablePrefix}ID`);

        const configObj = {
          ID_Columna: idColumna,
          Nombre_Encabezado: String(row[3] || idColumna).trim(),
          Visible_Encabezado: String(row[4] || "").trim(),
          Justificado_Campo: String(row[5] || "left").trim().toLowerCase(),
          Es_Obligatorio: !esID && String(row[6] || "").trim().toLowerCase() === "x" 
        };
        configMap[idColumna] = configObj;
        fullConfigForFrontend.push(configObj);
      }
    }); 

    // 2. Procesar Datos de la Tabla
    const fullData = dataSheet.getDataRange().getValues();
    if (fullData.length === 0) throw new Error("La tabla está vacía.");
    
    const originalHeaders = fullData[0];
    const tablePrefix = tableName.split('_')[0].toUpperCase();
    const finalHeaders = [];
    const finalDisplayMap = {};
    const finalAlignMap = {};
    const colIndexesToFetch = [];

    originalHeaders.forEach((headerName, index) => {
      const cleanH = String(headerName).trim();
      const upperH = cleanH.toUpperCase();
      const config = configMap[cleanH];
      
      const isPK = upperH === `${tablePrefix}ID`;
      const isAuditField = upperH.endsWith("REGISTROUSER") || upperH.endsWith("REGISTRODATA");
      const isTypeReg = upperH.endsWith("TYPEREG");
      
      if (ignoreVisibility || (config && config.Visible_Encabezado !== "") || isPK || isAuditField || isTypeReg) {
        finalHeaders.push(cleanH);
        finalDisplayMap[cleanH] = (config && config.Nombre_Encabezado) ? config.Nombre_Encabezado : cleanH;
        finalAlignMap[cleanH] = (config && config.Justificado_Campo) ? config.Justificado_Campo : 'left';
        colIndexesToFetch.push(index);
      }
    });

    const jsonData = fullData.slice(1).map(row => {
      const obj = {};
      colIndexesToFetch.forEach((colIdx, i) => { obj[finalHeaders[i]] = row[colIdx]; });
      return obj;
    });

    // 3. Sincronización Maestra
    const masterFields = typeof syncAndGetMasterFields === "function" ? syncAndGetMasterFields(ssData) : []; 

    return createJsonResponse({
      success: true,
      data: jsonData,
      columnOrder: finalHeaders,
      displayMap: finalDisplayMap,
      alignMap: finalAlignMap,
      fullConfig: fullConfigForFrontend,
      availableTables: availableTables,
      masterFields: masterFields 
    });

  } catch (err) {
    return createJsonResponse({ success: false, message: err.toString() });
  }
}

function doPost(e) {
  try {
    const params = e.parameter;
    const action = params.action;
    let result;

    if (action === "registerDynamicDataTD") result = handleDynamicDataTD(params, "REGISTER");
    else if (action === "editDynamicDataTD") result = handleDynamicDataTD(params, "EDIT");
    else if (action === "deleteDynamicDataTD") result = handleDynamicDataTD(params, "DELETE"); // Nueva acción
    else throw new Error("Acción desconocida");

    syncAndGetMasterFields();
    return result;
  } catch (err) {
    return createJsonResponse({ success: false, message: err.toString() });
  }
}

function syncAndGetMasterFields() {
  const ssConfig = SpreadsheetApp.openById(CONFIG_SS_ID);
  const ssData = SpreadsheetApp.openById(DATA_SS_ID);
  const configData = ssConfig.getSheetByName(CONFIG_SHEET_NAME).getDataRange().getValues().slice(1);

  const fieldCounts = {};
  configData.forEach(row => {
    const field = String(row[2]).trim();
    if (field) fieldCounts[field] = (fieldCounts[field] || 0) + 1;
  });

  const sharedFields = Object.keys(fieldCounts).filter(f => fieldCounts[f] > 1);
  const masterStructure = [];

  sharedFields.forEach(field => {
    const fieldPrefix = field.substring(0, 5).toUpperCase();
    const baseRow = configData.find(row => {
      const tName = String(row[1]).toUpperCase();
      return String(row[2]) === field && tName.startsWith(fieldPrefix);
    });

    if (baseRow) {
      const baseTableName = baseRow[1];
      const values = extractUniqueValues(ssData, baseTableName, field);
      masterStructure.push({
        Nombre_Tabla: baseTableName,
        Encabezado_Tabla: field,
        Valores_Encabezado: values
      });
    }
  });

  saveToValoresTable(ssData, masterStructure);
  return masterStructure;
}

function saveToValoresTable(ss, data) {
  let sheet = ss.getSheetByName('TV001_VALORES') || ss.insertSheet('TV001_VALORES');
  
  // 1. Preparar la matriz completa empezando por los encabezados
  const output = [['Nombre_Tabla', 'Encabezado_Tabla', 'Valores_Encabezado']];
  
  // 2. Si hay datos, transformarlos y agregarlos a la matriz
  if (data && data.length > 0) {
    const rows = data.map(item => [
      item.Nombre_Tabla, 
      item.Encabezado_Tabla, 
      JSON.stringify(item.Valores_Encabezado)
    ]);
    output.push(...rows);
  }

  // 3. Limpiar el contenido previo (solo valores, mantiene formatos si los hay)
  sheet.clearContents();

  // 4. Escribir todo el bloque desde la fila 1, columna 1
  // Esto garantiza que los encabezados se escriban siempre en la línea 1
  sheet.getRange(1, 1, output.length, 3).setValues(output);
  
  // 5. Opcional: Proteger los encabezados (Negrita y Congelar fila)
  sheet.getRange(1, 1, 1, 3).setFontWeight("bold");
  if (sheet.getFrozenRows() === 0) sheet.setFrozenRows(1);
}

function extractUniqueValues(ss, sheetName, colName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const idx = data[0].indexOf(colName);
  if (idx === -1) return [];
  return [...new Set(data.slice(1).map(r => r[idx]).filter(c => c !== ""))].sort();
}

function createJsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

