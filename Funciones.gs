/**
 * ---------------------------------------------------------
 * HERRAMIENTAS Y FÓRMULAS
 * ---------------------------------------------------------
 */

/**
 * Calcula cuánto ahorrar semanalmente.
 * @customfunction
 */
function AHORRO_SEMANAL(fechaViernes, montoMensual, diaCorte) {
  if (!fechaViernes || isNaN(new Date(fechaViernes).getTime())) return 0;
  
  var fecha = new Date(fechaViernes);
  var pago = new Date(fecha);
  
  // Lógica: Si hoy es antes del corte, paga este mes. Si no, el siguiente.
  var mesesASumar = fecha.getDate() <= diaCorte ? 0 : 1;
  
  pago.setDate(1); 
  pago.setMonth(pago.getMonth() + mesesASumar);
  pago.setDate(diaCorte);

  // Ajuste de año si es necesario
  if (pago < fecha) pago.setFullYear(pago.getFullYear() + 1);

  // Contar cuántos viernes faltan
  var inicio = new Date(pago);
  inicio.setMonth(inicio.getMonth() - 1);
  inicio.setDate(diaCorte); 

  var viernes = 0;
  var contador = new Date(inicio);
  contador.setDate(contador.getDate() + 1); 

  while (contador <= pago) {
    if (contador.getDay() == 5) viernes++; // 5 es Viernes
    contador.setDate(contador.getDate() + 1);
  }

  if (viernes == 0) return 0;
  return montoMensual / viernes;
}

function getAttr(element, attrName) {
  if (!element) return "";
  var attr = element.getAttribute(attrName);
  return attr ? attr.getValue() : "";
}