/**
 * ---------------------------------------------------------
 * 🪙 RIKO FINANZAS - MENÚ PRINCIPAL
 * ---------------------------------------------------------
 */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🪙 FINANZAS 🪙')
    .addItem('🚀 IMPORTAR (Nómina + Vales)', 'ejecutarTodo')
    .addSeparator()
    .addItem('✅ Confirmar Proyección (F:M)', 'confirmarSimulacionAControl') 
    .addSeparator()
    .addItem('💰 Ajustar Efectivo Real', 'ajustarDiferenciaEfectivo')
    .addItem('🟢 Actualizar Semana Visual', 'resaltarSemanaSpartan')
    .addToUi();
}

function ejecutarTodo() {
  // AQUÍ ESTABA EL ERROR: Antes hacíamos la lógica aquí.
  // AHORA: Llamamos directamente a la función maestra de NominaYVales.gs
  // que contiene el cálculo matemático y el "Protocolo Agresivo" del 0.
  importarTodo(); 
}

// Estas funciones se mantienen por si los otros botones las necesitan,
// pero el botón principal ahora obedece a NominaYVales.gs
function confirmarSimulacionAControl() {
  // Aseguramos que llame a la función correcta en Simulacion.gs
  if (typeof confirmarSimulacionAControl === 'function') {
      confirmarSimulacionAControl();
  } else {
      // Fallback si hay conflicto de nombres
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      ss.toast("Ejecutando desde Simulacion.gs...");
  }
}