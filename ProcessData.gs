/**
 * Procesa registros (Crear/Editar) de forma dinámica preservando la integridad
 * de todas las tablas y normalizando formatos numéricos.
 */
function handleDynamicDataTD(params, mode) {
  const ssData = SpreadsheetApp.openById(DATA_SS_ID);
  const sheet = ssData.getSheetByName(params.TABLA_DESTINO);
  
  if (!sheet) return createJsonResponse({ success: false, message: 'Tabla no encontrada: ' + params.TABLA_DESTINO });

  const fullData = sheet.getDataRange().getValues();
  const headers = fullData[0];
  const timestamp = Utilities.formatDate(new Date(), "GMT-4", "dd/MM/yyyy HH:mm:ss");
  
  const tablePrefix = params.TABLA_DESTINO.split('_')[0].toUpperCase();
  const campoClave = params.CAMPO_CLAVE || (tablePrefix + "ID");
  
  const idColIndex = headers.indexOf(campoClave);
  if (idColIndex === -1) return createJsonResponse({ success: false, message: 'Falta columna clave: ' + campoClave });

  let rowIndex = -1;
  let newGeneratedId = null;

  // 1. LOCALIZACIÓN O GENERACIÓN DE ID
  if (mode === "REGISTER") {
      newGeneratedId = generateNextIDInternal(fullData, tablePrefix);
    } else {
      // LOCALIZACIÓN (Común para EDIT y DELETE)
      const rawIdValue = params[campoClave] || params.ID_VALUE;
      const valorBusqueda = Number(String(rawIdValue).replace(/[.,\s]/g, ''));

      for (let i = 1; i < fullData.length; i++) {
        const cellValue = Number(String(fullData[i][idColIndex]).replace(/[.,\s]/g, ''));
        if (cellValue === valorBusqueda) {
          rowIndex = i + 1;
          break;
        }
      }
      if (rowIndex === -1) return createJsonResponse({ success: false, message: 'ID no hallado.' });
      
      // LÓGICA DE BORRADO
      if (mode === "DELETE") {
        return moveRowToHistory(ssData, sheet, rowIndex, headers, params);
      }
    }

  // 2. PREPARACIÓN DE FILA (INTEGRIDAD TOTAL)
  // Si es EDIT, cargamos los valores actuales de la fila para no perder campos omitidos por el frontend
  const rowValues = (mode === "EDIT") ? [...fullData[rowIndex - 1]] : new Array(headers.length).fill("");

  headers.forEach((header, index) => {
    const cleanH = header.trim();
    const upperH = cleanH.toUpperCase();
    
    // Identificadores de propósito del campo (Universal para cualquier tabla)
    const isRegistroUser = upperH.endsWith("REGISTROUSER");
    const isRegistroData = upperH.endsWith("REGISTRODATA");
    const isIdentityName = upperH.endsWith("IDNOMBRE");

    // A. Asignación de Llave Primaria
    if (cleanH === campoClave && mode === "REGISTER") {
      rowValues[index] = newGeneratedId;
    } 
    // B. Metadatos de Auditoría (Se activan por sufijo)
    else if (isRegistroUser) {
      rowValues[index] = params[cleanH] || params.currentUser || "UserSys";
    } 
    else if (isRegistroData) {
      rowValues[index] = timestamp;
    } 
    // C. Datos Dinámicos provenientes del Frontend
    else if (params[cleanH] !== undefined) {
      let val = params[cleanH];

      // --- CONVERSIÓN NUMÉRICA UNIVERSAL ---
      if (typeof val === "string" && val.trim() !== "") {
        // No intentamos convertir campos de Identidad/Nombre
        if (!isIdentityName) {
          // Reemplazamos coma decimal por punto (Estándar Anglo-Sajón de Google Sheets)
          // Esto evita multiplicar por 100 valores como "25,50"
          let normalizedVal = val.replace(/,/g, '.');

          // Validamos si es un número real y no un código con ceros a la izquierda (ej. "00123")
          if (!isNaN(normalizedVal) && !val.startsWith('0')) {
            val = Number(normalizedVal);
          }
        }
      }
      rowValues[index] = (val === null || val === undefined) ? "" : val;
    }
    // NOTA IMPORTANTE: Si params[cleanH] es undefined, rowValues[index] mantiene
    // el valor que ya tenía la fila, protegiendo campos como TD101DAT01.
  });

  // 3. PERSISTENCIA EN HOJA DE CÁLCULO
  try {
    if (mode === "REGISTER") {
      sheet.appendRow(rowValues);
    } else {
      sheet.getRange(rowIndex, 1, 1, headers.length).setValues([rowValues]);
    }

    const responseObj = {};
    headers.forEach((h, i) => responseObj[h.trim()] = rowValues[i]);

    return createJsonResponse({ 
      success: true, 
      message: mode === "EDIT" ? 'Registro actualizado correctamente.' : 'Nuevo registro creado.',
      data: responseObj 
    });

  } catch (e) {
    return createJsonResponse({ success: false, message: 'Error al escribir en la hoja: ' + e.toString() });
  }
}

function moveRowToHistory(ss, sourceSheet, rowIndex, headers, params) {
  const historySheet = ss.getSheetByName("TD999_BORRADOS");
  if (!historySheet) return createJsonResponse({ success: false, message: 'Tabla TD999_BORRADOS no hallada.' });

  // 1. Obtener datos actuales antes de borrar
  const rowDataArray = sourceSheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
  const rowDataObj = {};
  headers.forEach((h, i) => rowDataObj[h.trim()] = rowDataArray[i]);

  const timestamp = Utilities.formatDate(new Date(), "GMT-4", "dd/MM/yyyy HH:mm:ss");
  const tablePrefix = params.TABLA_DESTINO.split('_')[0].toUpperCase();
  const idDocumento = rowDataObj[tablePrefix + "ID"] || "N/A";

  // 2. Generar ID Correlativo para TD999ID (Busca el máximo para permitir orden descendente)
  const lastRow = historySheet.getLastRow();
  let nextId = 999001;
  if (lastRow > 1) {
    // Obtenemos todos los IDs de la primera columna
    const allIds = historySheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    const maxId = Math.max(...allIds.filter(id => !isNaN(id)));
    if (maxId >= 999001) nextId = maxId + 1;
  }

  // 3. Estructura: TD999ID, TD999IDDOC, TD999DATAJSON, TD999RegistroUser, TD999RegistroData
  const historyRow = [
    nextId,           // TD999ID
    idDocumento,      // TD999IDDOC
    JSON.stringify(rowDataObj), 
    params.usuario_id || "User", 
    timestamp
  ];

  try {
    // 4. Insertar el registro
    historySheet.appendRow(historyRow);

    // 5. ORDENAR DESCENDENTE (Por la columna 1: TD999ID)
    const newLastRow = historySheet.getLastRow();
    if (newLastRow > 1) {
      const lastCol = historySheet.getLastColumn();
      // Aplicamos el sort a todo el rango de datos (excluyendo encabezado)
      historySheet.getRange(2, 1, newLastRow - 1, lastCol)
                  .sort({ column: 1, ascending: false });
    }

    // 6. Eliminar de la hoja original
    sourceSheet.deleteRow(rowIndex);

    return createJsonResponse({ 
      success: true, 
      message: 'Registro eliminado correctamente.' 
    });
  } catch (e) {
    return createJsonResponse({ success: false, message: 'Error en archivo: ' + e.toString() });
  }
}

function generateNextIDInternal(fullData, prefix) {
  const numericPrefix = prefix.replace(/\D/g, "");
  const rangeStart = parseInt(numericPrefix + "1001");
  const rangeEnd = parseInt(numericPrefix + "9999");
  
  const ids = fullData.slice(1).map(row => {
    if (!row[0]) return null;
    const cleanId = String(row[0]).replace(/[.,\s]/g, "");
    const numId = parseInt(cleanId);
    return isNaN(numId) ? null : numId;
  }).filter(id => id !== null && id >= rangeStart && id <= rangeEnd);

  const maxId = ids.length === 0 ? rangeStart - 1 : Math.max(...ids);
  const nextId = maxId + 1;

  if (nextId > rangeEnd) throw new Error("Rango agotado para " + prefix);
  return nextId;
}

/**
 * Obtiene la configuración de nombres amigables de las tablas.
 * Filtra por AppTienda para devolver solo lo relevante.
 */
function getTableFriendlyNames(appTienda) {
  try {
    const ssConfig = SpreadsheetApp.openById(CONFIG_SS_ID);
    const sheet = ssConfig.getSheetByName("ConfigTB");
    if (!sheet) return { success: false, message: "Hoja ConfigTB no encontrada" };

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1);

    // Mapeamos los datos para que el frontend reciba un diccionario útil
    // Filtramos por AppTienda (si se proporciona)
    const configMap = {};
    
    rows.forEach(row => {
      const tienda = row[0]; // AppTienda
      const nombreTecnico = row[1]; // Nombre_Tabla (ej: TD101_BASIC)
      const nombreAmigable = row[2]; // Descripción_Tabla (ej: PRINCIPAL)
      
      if (!appTienda || tienda === appTienda) {
        configMap[nombreTecnico] = {
          label: nombreAmigable,
          c1: row[3], // ConfgTB01 (opcional para usos futuros)
          c2: row[4]  // ConfgTB02
        };
      }
    });

    return { success: true, data: configMap };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}
