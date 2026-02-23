/**
 * ---------------------------------------------------------
 * ⚔️ LÓGICA DE CÁLCULO - MODO SPARTAN
 * ---------------------------------------------------------
 */

function ajustarDiferenciaEfectivo() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  var lastRow = sheet.getLastRow();

  // --- 1. UBICAR SEMANA ACTUAL PRIMERO ---
  var today = new Date();
  today.setHours(0,0,0,0);
  var dataFechas = sheet.getRange(2, 1, lastRow - 1, 1).getValues(); 
  var targetRowIndex = -1;
  for (var i = 0; i < dataFechas.length; i++) {
    var rowDate = new Date(dataFechas[i][0]);
    rowDate.setHours(0,0,0,0);
    if (rowDate <= today) {
      targetRowIndex = i + 2;
    } else {
      break; 
    }
  }

  if (targetRowIndex == -1) return;

  // --- 2. CALCULAR DINERO TEÓRICO (SOLO HASTA SEMANA ACTUAL) ---
  var numFilas = targetRowIndex - 1;
  var valuesN = sheet.getRange(2, 14, numFilas, 1).getValues();
  
  var teorico = 0;
  for (var i = 0; i < valuesN.length; i++) {
    var val = valuesN[i][0];
    if (typeof val === 'number') {
      teorico += val;
    }
  }
  teorico = Math.round(teorico * 100) / 100;

  // --- 3. SOLICITAR DINERO FÍSICO ---
  var promptMessage = 'Teórico: $' + teorico.toFixed(2) + '\n\n' +
                      '¿Cuánto dinero físico tienes realmente?';
  var response = ui.prompt('Ajuste Spartan', promptMessage, ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() != ui.Button.OK) {
    return;
  }

  var inputUsuario = response.getResponseText();
  var fisico = parseFloat(inputUsuario.replace(',', '.').trim());

  if (isNaN(fisico)) {
    return;
  }

  // --- 4. CALCULAR DIFERENCIA ---
  var diferencia = teorico - fisico;
  diferencia = Math.round(diferencia * 100) / 100;

  if (diferencia === 0) {
    return;
  }

  // --- 5. EJECUTAR AJUSTE EN COMIDA (Columna K) ---
  var cellK = sheet.getRange(targetRowIndex, 11);
  var formula = cellK.getFormula();
  var value = cellK.getValue();
  var nuevaFormula = "";

  if (formula === "" && (value === "" || value === null)) {
    nuevaFormula = "=(" + diferencia + ")";
  } 
  else if (formula === "" && typeof value === 'number') {
    nuevaFormula = "=" + value + "+(" + diferencia + ")";
  }
  else {
    var regex = /\+\(([^)]+)\)$/;
    var match = formula.match(regex);
    if (match) {
      var contenidoPrevio = match[1];
      var nuevoContenido = contenidoPrevio + "+" + diferencia;
      nuevaFormula = formula.replace(regex, "+(" + nuevoContenido + ")");
    } else {
      nuevaFormula = formula + "+(" + diferencia + ")";
    }
  }

  cellK.setFormula(nuevaFormula);
}

/**
 * ---------------------------------------------------------
 * 🟢 VISOR TÁCTICO - ACTUALIZAR SEMANA VISUAL
 * ---------------------------------------------------------
 * Colorea la fila de la semana actual en verde claro (#d9ead3).
 * RESPETA las celdas que ya tengan otro color.
 * Limpia el rastro de semanas anteriores.
 * EJECUCIÓN SILENCIOSA.
 */
function resaltarSemanaSpartan() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  
  // Asumimos que los datos van de A a P (Columna 1 a 16).
  var range = sheet.getRange(2, 1, lastRow - 1, 16); 
  var backgrounds = range.getBackgrounds(); 
  var dates = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

  var COLOR_SEMANA = "#d9ead3"; // Verde Claro
  var COLOR_FONDO = "#ffffff";  // Blanco

  // 1. ENCONTRAR LA FILA DE LA SEMANA ACTUAL
  var today = new Date();
  today.setHours(0,0,0,0);
  var targetRowIndex = -1; 

  for (var i = 0; i < dates.length; i++) {
    var rowDate = new Date(dates[i][0]);
    rowDate.setHours(0,0,0,0);
    if (rowDate <= today) {
      targetRowIndex = i; 
    } else {
      break; 
    }
  }

  if (targetRowIndex == -1) {
    return; // Si falla, abortamos en silencio
  }

  // 2. BARRIDO DE COLORES
  for (var r = 0; r < backgrounds.length; r++) {
    var isTargetRow = (r == targetRowIndex);

    for (var c = 0; c < 16; c++) { 
       var currentColor = backgrounds[r][c];

       if (isTargetRow) {
          if (currentColor === COLOR_FONDO || currentColor === COLOR_SEMANA) {
             backgrounds[r][c] = COLOR_SEMANA;
          }
       } else {
          if (currentColor === COLOR_SEMANA) {
             backgrounds[r][c] = COLOR_FONDO;
          }
       }
    }
  }

  // 3. APLICAR CAMBIOS
  range.setBackgrounds(backgrounds);
  // Silencio total.
}