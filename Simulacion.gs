/**
 * ---------------------------------------------------------
 * 🔮 MOTOR DE SIMULACIÓN -> REALIDAD (CONTROL) [MODO SILENCIOSO]
 * ---------------------------------------------------------
 */

function confirmarSimulacionAControl() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetSim = ss.getSheetByName("SIMULACION");
  var sheetControl = ss.getSheetByName("CONTROL");

  // Validación mínima de seguridad (sin alertas)
  if (!sheetSim || !sheetControl) return;

  // 1. CALCULAR LA FILA DE LA SEMANA ACTUAL (Matemática Pura)
  var filaActual = calcularFilaPorMatematica();
  
  // 2. OBTENER DATOS (Columnas E a M)
  // E=5, M=13. Total de columnas a copiar = 9
  var rangoGastos = sheetSim.getRange(filaActual, 5, 1, 9); 
  var valoresGastos = rangoGastos.getValues();

  // 3. INYECTAR DATOS EN CONTROL
  sheetControl.getRange(filaActual, 5, 1, 9).setValues(valoresGastos);

  // CERO AVISOS, CERO TOASTS. MISIÓN CUMPLIDA.
}

/**
 * 🧮 CALCULADORA DE FILAS (Versión Simulacion.gs)
 * Base absoluta: 02/01/2026 = Fila 2.
 */
function calcularFilaPorMatematica() {
    var fechaBase = new Date(2026, 0, 2); 
    var hoy = new Date();
    
    fechaBase.setHours(0,0,0,0);
    hoy.setHours(0,0,0,0);
    
    var diferenciaTiempo = hoy.getTime() - fechaBase.getTime();
    var diferenciaDias = Math.floor(diferenciaTiempo / (1000 * 3600 * 24));
    var semanas = Math.floor(diferenciaDias / 7);
    
    return 2 + semanas;
} //d