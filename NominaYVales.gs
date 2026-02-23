/**
 * ---------------------------------------------------------
 * FUNCIÓN PRINCIPAL: IMPORTAR TODO (NÓMINA + VALES)
 * ---------------------------------------------------------
 */
function importarTodo() {
  var ui = SpreadsheetApp.getUi();
  var sheetControl = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CONTROL");
  var sheetSim = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SIMULACION");

  if (!sheetControl || !sheetSim) {
    ui.alert("❌ Error: Faltan hojas CONTROL o SIMULACION.");
    return;
  }

  // 1. Ejecutar motor de Nómina 
  var resultadoNomina = procesarNomina();
  
  // 2. Ejecutar motor de Vales (Devuelve null si no hay correos)
  var resultadoVales = procesarVales(true);
  
  // --- PROTOCOLO AGRESIVO: SI NO HAY VALES, PONER 0 ---
  if (resultadoVales === null) {
    
    // USAMOS CÁLCULO MATEMÁTICO (Base 02/01/2026)
    var filaDestino = calcularFilaPorMatematica();
    
    if (filaDestino > 0) {
        
        // --- ACCIÓN 1: SIMULACIÓN (BORRAR EL 2000) ---
        sheetSim.getRange(filaDestino, 3).setValue(0); 
        
        // --- ACCIÓN 2: CONTROL (LLENAR VACÍO) ---
        var celdaValesControl = sheetControl.getRange(filaDestino, 3);
        var valorActual = celdaValesControl.getValue();

        // Si está vacío, nulo o tiene basura, ponemos 0.
        if (valorActual === "" || valorActual === null) {
           celdaValesControl.setValue(0);
        }

        resultadoVales = "ℹ️ Vales: No hubo correo. Se puso $0.00 en la fila " + filaDestino + " (Calculada Matemáticamente).";
    } else {
        resultadoVales = "⚠️ Alerta: No hubo vales y no pude calcular la fila.";
    }
  }

  ui.alert(resultadoNomina + "\n" + resultadoVales);
}

/**
 * MOTOR DE NÓMINA (SEAH)
 */
function procesarNomina() {
  var sheetNomina = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("NOMINAS");
  if (!sheetNomina) return "❌ Falta hoja NOMINAS.";

  var query = 'from:recibos_nominasseah@m.seahprecision.com.mx has:attachment filename:xml after:2025/12/31';
  var hilos = GmailApp.search(query); 
  
  if (hilos.length == 0) return "⚠️ No encontré correos de nómina.";

  var contadorImportados = 0;
  var detalles = "";

  for (var i = 0; i < hilos.length; i++) {
    var mensajes = hilos[i].getMessages();
    var mensaje = mensajes[mensajes.length - 1]; 
    var adjuntos = mensaje.getAttachments();
    var xmlData = null;

    for (var j = 0; j < adjuntos.length; j++) {
      if (adjuntos[j].getName().toLowerCase().endsWith(".xml")) {
        xmlData = adjuntos[j].getDataAsString();
        break;
      }
    }

    if (!xmlData) continue; 

    try {
      var doc = XmlService.parse(xmlData);
      var root = doc.getRootElement();
      var nsNomina = XmlService.getNamespace("http://www.sat.gob.mx/nomina12");
      var nsCfdi = root.getNamespace();
      var complemento = root.getChild("Complemento", nsCfdi);
      
      if (!complemento) {
        var hijos = root.getChildren();
        for (var h=0; h<hijos.length; h++) { if (hijos[h].getName() == "Complemento") complemento = hijos[h]; }
      }
      if (!complemento) continue;

      var nomina = complemento.getChild("Nomina", nsNomina);
      if (!nomina) continue;

      var fechaPago = obtenerValorXML(nomina, "FechaPago") || obtenerValorXML(root, "Fecha");
      if (fechaPago.indexOf("T") > -1) fechaPago = fechaPago.split("T")[0];

      // Verificar duplicados
      var duplicado = false;
      var datos = sheetNomina.getDataRange().getValues();
      for (var d = 0; d < datos.length; d++) {
        var fechaCelda = datos[d][0];
        var fechaString = (fechaCelda instanceof Date) ? Utilities.formatDate(fechaCelda, Session.getScriptTimeZone(), "yyyy-MM-dd") : fechaCelda.toString();
        if (fechaString.indexOf(fechaPago) > -1) { duplicado = true; break; }
      }
      if (duplicado) continue; 

      var totalPercepciones = parseFloat(obtenerValorXML(nomina, "TotalPercepciones") || 0);
      var totalDeducciones = parseFloat(obtenerValorXML(nomina, "TotalDeducciones") || 0);
      var isr = 0, imss = 0, fondoAhorro = 0, fonacot = 0, sindicato = 0, ajuste = 0;

      var deducciones = nomina.getChild("Deducciones", nsNomina);
      if (deducciones) {
        var lista = deducciones.getChildren("Deduccion", nsNomina);
        for (var k = 0; k < lista.length; k++) {
          var tipo = obtenerValorXML(lista[k], "TipoDeduccion");
          var importe = parseFloat(obtenerValorXML(lista[k], "Importe") || 0);
          var concepto = obtenerValorXML(lista[k], "Concepto").toLowerCase();

          if (tipo == "002") isr += importe;
          else if (tipo == "001") imss += importe;
          else if (concepto.includes("fondo") && concepto.includes("ahorro")) fondoAhorro += importe;
          else if (concepto.includes("fonacot")) fonacot += importe;
          else if (concepto.includes("sindical") || concepto.includes("cuota")) sindicato += importe;
          else if (concepto.includes("ajuste")) ajuste += importe;
        }
      }
      var otrosPagos = nomina.getChild("OtrosPagos", nsNomina);
      var totalOtros = 0;
      if (otrosPagos) totalOtros = parseFloat(obtenerValorXML(otrosPagos, "TotalOtrosPagos") || 0);
      
      var netoReal = totalPercepciones + totalOtros - totalDeducciones;

      // INSERTAR EN NOMINAS
      sheetNomina.insertRowBefore(2);
      sheetNomina.getRange(2, 1, 1, 9).setValues([[fechaPago, totalPercepciones, isr, imss, fondoAhorro, fonacot, sindicato, ajuste, netoReal]]);
      
      // ACTUALIZAR SIMULACION (Sueldo y Pensión)
      var filaSim = calcularFilaPorFechaInput(fechaPago); 
      if (filaSim > 0) {
        var sheetSim = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SIMULACION");
        
        // 1. Sueldo (Valor directo)
        sheetSim.getRange(filaSim, 2).setValue(netoReal); 
        
        // 2. Pensión (Fórmula con VALORES)
        // Construimos la fórmula con los números exactos: =(4500-200-100)*0.15
        var formulaPension = "=(" + totalPercepciones + "-" + isr + "-" + imss + ")*0.15";
        sheetSim.getRange(filaSim, 4).setFormula(formulaPension);
      }

      contadorImportados++;
      detalles += "\n✅ Nómina: " + fechaPago + " ($" + netoReal + ")";

    } catch (e) { continue; }
  }

  return contadorImportados > 0 ? "📥 REPORTE:\n" + detalles : "✅ Nómina al día.";
}

/**
 * MOTOR DE VALES
 */
function procesarVales(silencioso) {
  var sheetControl = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CONTROL"); 
  if (!sheetControl) return null;

  var query = 'from:info@mercadopago.com subject:"Tu transferencia fue enviada" newer_than:5d';
  var hilos = GmailApp.search(query);
  
  if (hilos.length == 0) return null; // Retorna NULL para activar protocolo 0

  var mensaje = hilos[0].getMessages().pop();
  var cuerpo = mensaje.getPlainBody();
  var fechaCorreo = mensaje.getDate();
  var regex = /transferencia de\s*\$\s*([0-9,]+\.[0-9]{2})/;
  var coincidencia = cuerpo.match(regex);
  if (!coincidencia) return null;

  var monto = parseFloat(coincidencia[1].replace(',', ''));
  
  // Usamos CÁLCULO MATEMÁTICO con la fecha del correo
  var filaActual = calcularFilaPorFechaInput(fechaCorreo);
  
  if (filaActual > 0) {
      var celdaVales = sheetControl.getRange(filaActual, 3); 
      if (celdaVales.getValue() === "") {
         celdaVales.setValue(monto);
         
         // Actualizar Simulación también
         var sheetSim = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SIMULACION");
         if(sheetSim) sheetSim.getRange(filaActual, 3).setValue(monto);
         
         return "✅ VALES: $" + monto;
      } else {
         return "⚠️ Vales ya registrados.";
      }
  }
  return null;
}

/**
 * 🧮 CALCULADORA DE FILAS SPARTAN (Matemática Pura - Base Hardcodeada)
 * Ignora lo que diga la celda A2. Asume 02/01/2026 como inicio absoluto.
 */
function calcularFilaPorMatematica() {
    // FECHA BASE ABSOLUTA: 2 de Enero de 2026
    var fechaBase = new Date(2026, 0, 2); // Mes 0 es Enero
    var hoy = new Date();
    
    fechaBase.setHours(0,0,0,0);
    hoy.setHours(0,0,0,0);
    
    var diferenciaTiempo = hoy.getTime() - fechaBase.getTime();
    var diferenciaDias = Math.floor(diferenciaTiempo / (1000 * 3600 * 24));
    
    // Semanas completas (Usamos floor para quedarnos en el viernes pasado)
    var semanas = Math.floor(diferenciaDias / 7);
    
    // La fila es: Base(2) + Semanas
    return 2 + semanas;
}

/**
 * Versión para fechas específicas (Nómina/Vales)
 */
function calcularFilaPorFechaInput(fechaInput) {
    var fechaBase = new Date(2026, 0, 2); // Base absoluta
    var fechaObj = (fechaInput instanceof Date) ? new Date(fechaInput) : new Date(fechaInput.replace(/-/g, '/'));
    
    fechaBase.setHours(0,0,0,0);
    fechaObj.setHours(0,0,0,0);
    
    var diffDias = Math.floor((fechaObj.getTime() - fechaBase.getTime()) / (1000 * 3600 * 24));
    var semanas = Math.round(diffDias / 7);
    
    return 2 + semanas;
}

// Helper para procesarNomina
function obtenerValorXML(nodo, attr) {
  var a = nodo.getAttribute(attr);
  return a ? a.getValue() : null;
}