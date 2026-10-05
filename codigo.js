/**
 * Crea el menú personalizado en la hoja de cálculo al abrir el documento.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Recibos')
    .addItem('Generar recibos', 'mostrarDialogoSeleccion')
    .addToUi();
}

/**
 * Muestra el modal con filtros (Período, Días, Tipo Cliente, Producto y Opción de Reprocesar).
 */
function mostrarDialogoSeleccion() {
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html>' +
    '<html>' +
    '<head>' +
      '<style>' +
        'body { font-family: Arial, sans-serif; padding: 15px; background-color: #f8f9fa; color: #333; font-size: 12px; }' +
        '.form-group { margin-bottom: 10px; }' +
        'label { font-weight: bold; display: block; margin-bottom: 3px; }' +
        'select, input[type="number"], input[type="text"] { width: 100%; padding: 6px 8px; box-sizing: border-box; border: 1px solid #ccc; border-radius: 4px; font-size: 12px; }' +
        '.row { display: flex; gap: 10px; }' +
        '.col { flex: 1; }' +
        '.checkbox-group { margin-top: 10px; display: flex; align-items: center; gap: 6px; }' +
        'button { margin-top: 15px; width: 100%; padding: 10px; background-color: #003366; color: white; border: none; border-radius: 4px; font-weight: bold; cursor: pointer; font-size: 13px; }' +
        'button:hover { background-color: #002244; }' +
      '</style>' +
    '</head>' +
    '<body>' +
      '<div class="row">' +
        '<div class="col form-group">' +
          '<label for="mes">Mes:</label>' +
          '<select id="mes">' +
            '<option value="1">Enero</option>' +
            '<option value="2">Febrero</option>' +
            '<option value="3">Marzo</option>' +
            '<option value="4">Abril</option>' +
            '<option value="5">Mayo</option>' +
            '<option value="6">Junio</option>' +
            '<option value="7">Julio</option>' +
            '<option value="8">Agosto</option>' +
            '<option value="9">Septiembre</option>' +
            '<option value="10">Octubre</option>' +
            '<option value="11">Noviembre</option>' +
            '<option value="12">Diciembre</option>' +
          '</select>' +
        '</div>' +
        '<div class="col form-group">' +
          '<label for="anio">Año:</label>' +
          '<input type="number" id="anio" value="2026" min="2000" max="2100">' +
        '</div>' +
      '</div>' +

      '<div class="form-group">' +
        '<label>Días del mes (Opcional):</label>' +
        '<div class="row">' +
          '<div class="col"><input type="number" id="diaDesde" placeholder="Desde día (ej: 1)" min="1" max="31"></div>' +
          '<div class="col"><input type="number" id="diaHasta" placeholder="Hasta día (ej: 31)" min="1" max="31"></div>' +
        '</div>' +
      '</div>' +

      '<div class="form-group">' +
        '<label for="tipoCliente">Tipo de Cliente (Opcional):</label>' +
        '<input type="text" id="tipoCliente" placeholder="Ej: VIP, Regular, Nuevo (dejar vacío = todos)">' +
      '</div>' +

      '<div class="form-group">' +
        '<label for="filtroProducto">Categoría / Producto (Opcional):</label>' +
        '<input type="text" id="filtroProducto" placeholder="Ej: Carnes, Lácteos (dejar vacío = todos)">' +
      '</div>' +

      '<div class="checkbox-group">' +
        '<input type="checkbox" id="reprocesar" id="reprocesar">' +
        '<label for="reprocesar" style="font-weight: normal;">Regenerar aunque ya tengan fecha en Recibo_Generado</label>' +
      '</div>' +

      '<button onclick="ejecutar()">Generar Recibos Filtrados</button>' +

      '<script>' +
        'document.getElementById("mes").value = "1";' +
        'function ejecutar() {' +
          'var filtros = {' +
            'mes: document.getElementById("mes").value,' +
            'anio: document.getElementById("anio").value,' +
            'diaDesde: document.getElementById("diaDesde").value,' +
            'diaHasta: document.getElementById("diaHasta").value,' +
            'tipoCliente: document.getElementById("tipoCliente").value,' +
            'producto: document.getElementById("filtroProducto").value,' +
            'reprocesar: document.getElementById("reprocesar").checked' +
          '};' +
          'google.script.run.withSuccessHandler(function() { google.script.host.close(); }).generarRecibosConFiltros(filtros);' +
        '}' +
      '</script>' +
    '</body>' +
    '</html>'
  )
  .setWidth(360)
  .setHeight(390)
  .setTitle('Generar Recibos con Filtros');

  SpreadsheetApp.getUi().showModalDialog(html, 'Filtros de Recibos');
}

/**
 * Función principal para procesar ventas en lote.
 */
function generarRecibosConFiltros(filtros) {
  var tiempoInicio = new Date().getTime();
  var maxTiempoMs = 4.5 * 60 * 1000;

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getActiveSheet();

    var ultimoRango = hoja.getDataRange();
    var datosVisibles = ultimoRango.getDisplayValues();

    if (datosVisibles.length <= 1) {
      SpreadsheetApp.getUi().alert('La hoja no contiene datos suficientes.');
      return;
    }

    var headerRowIdx = 0;
    for (var r = 0; r < Math.min(datosVisibles.length, 5); r++) {
      var filaStr = datosVisibles[r].join(' ').toLowerCase();
      if (filaStr.indexOf('fecha') !== -1) {
        headerRowIdx = r;
        break;
      }
    }

    var cabecera = datosVisibles[headerRowIdx];

    // Identificar columna Recibo_Generado
    var colReciboIdx = -1;
    for (var c = 0; c < cabecera.length; c++) {
      var colNombre = String(cabecera[c]).toLowerCase().trim();
      if (colNombre === 'recibo_generado' || colNombre === 'recibo generado') {
        colReciboIdx = c;
        break;
      }
    }

    if (colReciboIdx === -1) {
      colReciboIdx = cabecera.length;
      hoja.getRange(headerRowIdx + 1, colReciboIdx + 1).setValue('Recibo_Generado');
      cabecera.push('Recibo_Generado');
      for (var i = headerRowIdx + 1; i < datosVisibles.length; i++) {
        datosVisibles[i].push('');
      }
    }

    // Mapeo exacto de columnas especificas
    var idxFecha = buscarColumnaExacta(cabecera, ['fecha']);
    var idxSucursal = buscarColumnaExacta(cabecera, ['sucursal']);
    var idxVendedor = buscarColumnaExacta(cabecera, ['vendedor']);
    var idxCliente = buscarColumnaCliente(cabecera); // Busca prioritariamente la columna con nombre del cliente
    var idxTipoCliente = buscarColumnaExacta(cabecera, ['tipo cliente', 'tipo de cliente']);
    var idxCategoria = buscarColumnaExacta(cabecera, ['categoría', 'categoria']);
    var idxProducto = buscarColumnaExacta(cabecera, ['producto']);
    var idxCantidad = buscarColumnaExacta(cabecera, ['cantidad']);
    var idxTotal = buscarColumnaExacta(cabecera, ['total venta ($)', 'total ventas ($)', 'total venta', 'total ventas']);
    var idxMetodoPago = buscarColumnaExacta(cabecera, ['método pago', 'metodo pago']);

    var nombresMeses = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];
    var nombreMesStr = nombresMeses[parseInt(filtros.mes, 10) - 1];
    var nombreCarpeta = 'Recibos - ' + nombreMesStr + ' ' + filtros.anio;
    var carpeta = obtenerOCrearCarpeta(nombreCarpeta);

    var mesTarget = parseInt(filtros.mes, 10);
    var anioTarget = parseInt(filtros.anio, 10);
    var diaDesdeNum = filtros.diaDesde ? parseInt(filtros.diaDesde, 10) : 1;
    var diaHastaNum = filtros.diaHasta ? parseInt(filtros.diaHasta, 10) : 31;
    var tipoClienteFiltro = String(filtros.tipoCliente || '').toLowerCase().trim();
    var productoFiltro = String(filtros.producto || '').toLowerCase().trim();
    var reprocesar = filtros.reprocesar === true;

    var ventasPorCliente = {};

    for (var rowIdx = headerRowIdx + 1; rowIdx < datosVisibles.length; rowIdx++) {
      var fila = datosVisibles[rowIdx];
      var marcaRecibo = String(fila[colReciboIdx] || '').trim();

      // Saltear si ya fue procesado EXCEPTO si el usuario activó la casilla "reprocesar"
      if (marcaRecibo !== '' && !reprocesar) {
        continue;
      }

      var textoFecha = fila[idxFecha];
      if (!textoFecha) continue;

      var diaFila = 0, mesFila = 0, anioFila = 0;
      var partes = textoFecha.split('T')[0].split(/[-\/]/);

      if (partes.length === 3) {
        if (partes[0].length === 4) { // YYYY-MM-DD
          anioFila = parseInt(partes[0], 10);
          mesFila = parseInt(partes[1], 10);
          diaFila = parseInt(partes[2], 10);
        } else { // DD-MM-YYYY
          diaFila = parseInt(partes[0], 10);
          mesFila = parseInt(partes[1], 10);
          anioFila = parseInt(partes[2], 10);
        }
      }

      if (!mesFila || !anioFila) continue;

      // 1. Filtrar por Mes y Año
      if (mesFila !== mesTarget || anioFila !== anioTarget) {
        continue;
      }

      // 2. Filtrar por Rango de Días
      if (diaFila < diaDesdeNum || diaFila > diaHastaNum) {
        continue;
      }

      // 3. Filtrar por Tipo de Cliente
      if (tipoClienteFiltro !== '') {
        var tipoCliVal = idxTipoCliente !== -1 ? String(fila[idxTipoCliente] || '').toLowerCase().trim() : '';
        if (tipoCliVal.indexOf(tipoClienteFiltro) === -1) continue;
      }

      // 4. Filtrar por Producto/Categoría
      if (productoFiltro !== '') {
        var prodVal = idxProducto !== -1 ? String(fila[idxProducto] || '').toLowerCase().trim() : '';
        var catVal = idxCategoria !== -1 ? String(fila[idxCategoria] || '').toLowerCase().trim() : '';
        if (prodVal.indexOf(productoFiltro) === -1 && catVal.indexOf(productoFiltro) === -1) continue;
      }

      var cliente = String(fila[idxCliente] || 'Cliente_Desconocido').trim();

      if (!ventasPorCliente[cliente]) {
        ventasPorCliente[cliente] = [];
      }

      var totalMonto = 0;
      if (idxTotal !== -1 && fila[idxTotal] !== undefined) {
        var cleanTotal = String(fila[idxTotal]).replace(/[^0-9\.,]/g, '').replace(',', '.');
        totalMonto = parseFloat(cleanTotal) || 0;
      }

      var diaStr = ('0' + diaFila).slice(-2);
      var mesStr = ('0' + mesFila).slice(-2);

      ventasPorCliente[cliente].push({
        filaIndex: rowIdx,
        fecha: diaStr + '/' + mesStr + '/' + anioFila,
        sucursal: idxSucursal !== -1 ? String(fila[idxSucursal] || '').trim() : '',
        vendedor: idxVendedor !== -1 ? String(fila[idxVendedor] || '').trim() : '',
        producto: idxProducto !== -1 ? String(fila[idxProducto] || '').trim() : '',
        cantidad: idxCantidad !== -1 ? fila[idxCantidad] || 0 : 0,
        total: totalMonto,
        metodoPago: idxMetodoPago !== -1 ? String(fila[idxMetodoPago] || '').trim() : ''
      });
    }

    var clientesKeys = Object.keys(ventasPorCliente);
    if (clientesKeys.length === 0) {
      SpreadsheetApp.getUi().alert('No se encontraron ventas pendientes de recibo para ' + nombreMesStr + ' ' + anioTarget + '. Si ya generaste recibos antes, marca la casilla "Regenerar aunque ya tengan fecha".');
      return;
    }

    var recibosGeneradosCount = 0;
    var fechaHoraActual = new Date();
    var fechaHoraStr = formatearFechaHora(fechaHoraActual);
    var seCortoPorTiempo = false;

    var sufijoPeriodo = nombreMesStr + ' ' + anioTarget;
    if (filtros.diaDesde || filtros.diaHasta) {
      sufijoPeriodo += ' (Días ' + diaDesdeNum + '-' + diaHastaNum + ')';
    }

    // Generar PDFs por cliente
    for (var k = 0; k < clientesKeys.length; k++) {
      if (new Date().getTime() - tiempoInicio > maxTiempoMs) {
        seCortoPorTiempo = true;
        break;
      }

      var clienteNom = clientesKeys[k];
      var listaCompras = ventasPorCliente[clienteNom];

      var htmlContent = construirHtmlRecibo(clienteNom, listaCompras, sufijoPeriodo);

      var clienteLimpio = limpiarNombreArchivo(clienteNom);
      var nombrePDF = 'Recibo - ' + clienteLimpio + ' - ' + sufijoPeriodo;

      var blobHtml = Utilities.newBlob(htmlContent, MimeType.HTML, nombrePDF + '.html');
      var pdfBlob = blobHtml.getAs(MimeType.PDF);
      pdfBlob.setName(nombrePDF + '.pdf');

      carpeta.createFile(pdfBlob);
      recibosGeneradosCount++;

      for (var m = 0; m < listaCompras.length; m++) {
        var rIdx = listaCompras[m].filaIndex;
        datosVisibles[rIdx][colReciboIdx] = fechaHoraStr;
      }
    }

    // Actualización masiva de marcas de tiempo
    var columnaMarcas = [];
    for (var r = 0; r < datosVisibles.length; r++) {
      columnaMarcas.push([datosVisibles[r][colReciboIdx]]);
    }
    hoja.getRange(1, colReciboIdx + 1, columnaMarcas.length, 1).setValues(columnaMarcas);

    mostrarModalResultado(carpeta.getUrl(), recibosGeneradosCount, seCortoPorTiempo);

  } catch (error) {
    SpreadsheetApp.getUi().alert('Ocurrió un error durante la ejecución: ' + error.toString());
  }
}

/**
 * Busca específicamente la columna que contiene el Nombre del Cliente (evitando "Cliente ID").
 */
function buscarColumnaCliente(cabecera) {
  for (var i = 0; i < cabecera.length; i++) {
    var colTexto = String(cabecera[i]).toLowerCase().trim();
    if (colTexto === 'cliente' || colTexto === 'cliente ') {
      return i;
    }
  }
  for (var j = 0; j < cabecera.length; j++) {
    var colTexto2 = String(cabecera[j]).toLowerCase().trim();
    if (colTexto2.indexOf('cliente') !== -1 && colTexto2.indexOf('id') === -1) {
      return j;
    }
  }
  return 0;
}

/**
 * Busca el índice exacto o aproximado de una columna.
 */
function buscarColumnaExacta(cabecera, opciones) {
  for (var i = 0; i < cabecera.length; i++) {
    var colTexto = String(cabecera[i]).toLowerCase().trim();
    for (var j = 0; j < opciones.length; j++) {
      if (colTexto === opciones[j] || colTexto.indexOf(opciones[j]) !== -1) {
        return i;
      }
    }
  }
  return -1;
}

/**
 * Limpia caracteres no válidos para nombres de archivos.
 */
function limpiarNombreArchivo(texto) {
  return texto.replace(/[\/\\:\*\?"<>\|]/g, '_');
}

/**
 * Obtiene o crea la carpeta de destino en Drive.
 */
function obtenerOCrearCarpeta(nombreCarpeta) {
  var carpetas = DriveApp.getFoldersByName(nombreCarpeta);
  if (carpetas.hasNext()) {
    return carpetas.next();
  } else {
    return DriveApp.createFolder(nombreCarpeta);
  }
}

/**
 * Formatea un objeto Date a dd/mm/yyyy HH:mm:ss.
 */
function formatearFechaHora(fecha) {
  var dia = ('0' + fecha.getDate()).slice(-2);
  var mes = ('0' + (fecha.getMonth() + 1)).slice(-2);
  var anio = fecha.getFullYear();
  var hh = ('0' + fecha.getHours()).slice(-2);
  var mm = ('0' + fecha.getMinutes()).slice(-2);
  var ss = ('0' + fecha.getSeconds()).slice(-2);
  return dia + '/' + mes + '/' + anio + ' ' + hh + ':' + mm + ':' + ss;
}

/**
 * Construye el HTML del recibo usando tablas y estilos en línea.
 */
function construirHtmlRecibo(cliente, compras, periodoStr) {
  var totalGeneral = 0;
  for (var i = 0; i < compras.length; i++) {
    totalGeneral += compras[i].total;
  }

  var html = '<!DOCTYPE html>' +
  '<html>' +
  '<head>' +
    '<meta charset="utf-8">' +
    '<style>' +
      'body { font-family: Arial, sans-serif; color: #111111; margin: 20px; font-size: 12px; }' +
      'table { width: 100%; border-collapse: collapse; }' +
      '.header-table td { padding: 4px; vertical-align: top; }' +
      '.details-table th { background-color: #003366; color: #ffffff; padding: 8px; text-align: left; font-size: 11px; font-weight: bold; }' +
      '.details-table td { border-bottom: 1px solid #dddddd; padding: 7px; text-align: left; font-size: 11px; }' +
      '.total-row td { border-top: 2px solid #003366; font-weight: bold; font-size: 13px; color: #003366; }' +
      '.footer-text { font-size: 10px; color: #555555; text-align: center; margin-top: 25px; border-top: 1px dashed #cccccc; padding-top: 10px; }' +
    '</style>' +
  '</head>' +
  '<body>' +
    '<table class="header-table" style="margin-bottom: 20px;">' +
      '<tr>' +
        '<td style="width: 60%;">' +
          '<h1 style="color: #003366; margin: 0 0 5px 0; font-size: 20px;">TODO MAYOR</h1>' +
          '<p style="margin: 2px 0; color: #444444;">Aragua, Palo negro, municipio Libertador</p>' +
        '</td>' +
        '<td style="width: 40%; text-align: right;">' +
          '<h2 style="color: #111111; margin: 0 0 5px 0; font-size: 16px;">RECIBO DE COMPRA</h2>' +
          '<p style="margin: 2px 0; font-weight: bold;">Período: ' + periodoStr + '</p>' +
        '</td>' +
      '</tr>' +
    '</table>' +

    '<table class="header-table" style="margin-bottom: 20px; background-color: #f2f5f8; border-radius: 4px; padding: 8px;">' +
      '<tr>' +
        '<td>' +
          '<strong style="color: #003366;">Cliente:</strong> ' + cliente +
        '</td>' +
      '</tr>' +
    '</table>' +

    '<table class="details-table">' +
      '<thead>' +
        '<tr>' +
          '<th>Fecha</th>' +
          '<th>Sucursal</th>' +
          '<th>Vendedor</th>' +
          '<th>Método Pago</th>' +
          '<th>Producto</th>' +
          '<th style="text-align: center;">Cant.</th>' +
          '<th style="text-align: right;">Monto Total</th>' +
        '</tr>' +
      '</thead>' +
      '<tbody>';

  for (var j = 0; j < compras.length; j++) {
    var c = compras[j];
    html += '<tr>' +
      '<td>' + c.fecha + '</td>' +
      '<td>' + c.sucursal + '</td>' +
      '<td>' + c.vendedor + '</td>' +
      '<td>' + c.metodoPago + '</td>' +
      '<td>' + c.producto + '</td>' +
      '<td style="text-align: center;">' + c.cantidad + '</td>' +
      '<td style="text-align: right;">$' + c.total.toFixed(2) + '</td>' +
    '</tr>';
  }

  html += '<tr class="total-row">' +
        '<td colspan="6" style="text-align: right; padding-top: 10px;">TOTAL GENERAL:</td>' +
        '<td style="text-align: right; padding-top: 10px;">$' + totalGeneral.toFixed(2) + '</td>' +
      '</tr>' +
      '</tbody>' +
    '</table>' +

    '<div class="footer-text">' +
      '<p style="margin: 3px 0; font-weight: bold;">¡Gracias por su compra!</p>' +
      '<p style="margin: 3px 0;">Este documento es un comprobante de pago no fiscal</p>' +
    '</div>' +
  '</body>' +
  '</html>';

  return html;
}

/**
 * Muestra el modal final de resultados.
 */
function mostrarModalResultado(urlCarpeta, cantidad, huboCorte) {
  var mensajeCorte = huboCorte ? '<p style="color: #d9534f; font-weight: bold;">El proceso superó los 4.5 minutos y se pausó. Vuelve a ejecutarlo para continuar con el resto.</p>' : '';

  var htmlOutput = HtmlService.createHtmlOutput(
    '<!DOCTYPE html>' +
    '<html>' +
    '<head>' +
      '<style>' +
        'body { font-family: Arial, sans-serif; padding: 15px; text-align: center; color: #333; }' +
        '.btn { display: inline-block; padding: 10px 18px; background-color: #003366; color: white; text-decoration: none; border-radius: 4px; font-weight: bold; margin-top: 15px; }' +
        '.btn:hover { background-color: #002244; }' +
      '</style>' +
    '</head>' +
    '<body>' +
      '<h3>¡Proceso Completado!</h3>' +
      '<p>Se generaron <strong>' + cantidad + '</strong> recibos correctamente.</p>' +
      mensajeCorte +
      '<a href="' + urlCarpeta + '" target="_blank" class="btn">Abrir carpeta en Drive</a>' +
    '</body>' +
    '</html>'
  )
  .setWidth(360)
  .setHeight(220)
  .setTitle('Resultado');

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Resultado');
}