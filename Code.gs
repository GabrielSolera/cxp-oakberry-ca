const SHEET_NAME_FACTURAS    = 'Facturas';
const SHEET_NAME_PROVEEDORES = 'Proveedores';
const SHEET_NAME_PAISES      = 'Países';
const SHEET_NAME_PAGOS       = 'Pagos';
const SHEET_NAME_CONFIG      = 'Config';
const SHEET_NAME_MOVIMIENTOS = 'Movimientos';
const DRIVE_FOLDER_NAME      = 'CXP_CA_Documentos';

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('CXP Centroamérica')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getSpreadsheet() {
  return SpreadsheetApp.openById('1MYl1cIpE1QfM2tuXu2cUlD27HL-jOoKWv3pZf0GCAVY');
}

function registrarMovimiento(tipo, idFactura, detalle, monto, moneda, justificacion) {
  const sheet = getSpreadsheet().getSheetByName(SHEET_NAME_MOVIMIENTOS);
  const id    = generarId('MOV');
  sheet.appendRow([
    id,
    new Date(),
    tipo,
    idFactura,
    detalle,
    monto || '',
    moneda || '',
    justificacion || ''
  ]);
}

function getMovimientos(idFactura) {
  const sheet = getSpreadsheet().getSheetByName(SHEET_NAME_MOVIMIENTOS);
  const data  = sheet.getDataRange().getValues();
  const movs  = [];
  for (let i = 1; i < data.length; i++) {
    if (!idFactura || data[i][3] === idFactura) {
      movs.push({
        id            : data[i][0],
        fecha         : data[i][1] ? formatDate(data[i][1]) : '',
        fechaHora     : data[i][1] ? formatDateTime(data[i][1]) : '',
        tipo          : data[i][2],
        idFactura     : data[i][3],
        detalle       : data[i][4],
        monto         : parseFloat(data[i][5]) || 0,
        moneda        : data[i][6],
        justificacion : data[i][7]
      });
    }
  }
  return movs.reverse();
}

function getFacturas(filtros) {
  const ss    = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_FACTURAS);
  const data  = sheet.getDataRange().getValues();
  const hoy   = new Date();
  hoy.setHours(0, 0, 0, 0);
  const facturas = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;

    const factura = {
      id            : row[0],
      proveedor     : row[1],
      pais          : row[2],
      moneda        : row[3],
      montoTotal    : parseFloat(row[4]) || 0,
      montoPagado   : parseFloat(row[5]) || 0,
      fechaEmision  : row[6] ? formatDate(row[6]) : '',
      fechaVence    : row[7] ? formatDate(row[7]) : '',
      notas         : row[8] || '',
      docUrl        : row[9] || '',
      numeroFactura : row[10] || '',
      fila          : i + 1
    };

    factura.saldo  = factura.montoTotal - factura.montoPagado;
    factura.estado = calcularEstado(factura, hoy);
    factura.dias   = calcularDias(row[7], hoy);

    if (filtros) {
      if (filtros.pais       && factura.pais      !== filtros.pais)      continue;
      if (filtros.proveedor  && factura.proveedor !== filtros.proveedor) continue;
      if (filtros.moneda     && factura.moneda    !== filtros.moneda)    continue;
      if (filtros.fechaDesde && new Date(row[6])  < new Date(filtros.fechaDesde)) continue;
      if (filtros.fechaHasta && new Date(row[6])  > new Date(filtros.fechaHasta)) continue;
    }

    facturas.push(factura);
  }
  return facturas;
}

function crearFactura(data) {
  const ss    = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_FACTURAS);
  const id    = generarId('FAC');
  sheet.appendRow([
    id, data.proveedor, data.pais, data.moneda,
    parseFloat(data.montoTotal), 0,
    new Date(data.fechaEmision), new Date(data.fechaVence),
    data.notas || '', '', data.numeroFactura || ''
  ]);
  registrarMovimiento('Creación', id,
    'Factura creada: ' + (data.numeroFactura || id) + ' — ' + data.proveedor,
    parseFloat(data.montoTotal), data.moneda, '');
  return { ok: true, id };
}

function editarFactura(id, data, justificacion) {
  const ss    = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_FACTURAS);
  const rows  = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === id) {
      const montoAnterior = rows[i][4];
      sheet.getRange(i + 1, 2).setValue(data.proveedor);
      sheet.getRange(i + 1, 3).setValue(data.pais);
      sheet.getRange(i + 1, 4).setValue(data.moneda);
      sheet.getRange(i + 1, 5).setValue(parseFloat(data.montoTotal));
      sheet.getRange(i + 1, 7).setValue(new Date(data.fechaEmision));
      sheet.getRange(i + 1, 8).setValue(new Date(data.fechaVence));
      sheet.getRange(i + 1, 9).setValue(data.notas || '');
      sheet.getRange(i + 1, 11).setValue(data.numeroFactura || '');
      registrarMovimiento('Edición factura', id,
        'Monto anterior: ' + montoAnterior + ' → nuevo: ' + data.montoTotal,
        parseFloat(data.montoTotal), data.moneda, justificacion || '');
      return { ok: true };
    }
  }
  return { ok: false, error: 'Factura no encontrada' };
}

function eliminarFactura(id) {
  const ss    = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_FACTURAS);
  const rows  = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === id) {
      registrarMovimiento('Eliminación', id,
        'Factura eliminada: ' + (rows[i][10] || id) + ' — ' + rows[i][1],
        rows[i][4], rows[i][3], '');
      sheet.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Factura no encontrada' };
}

function registrarPago(idFactura, monto, fecha) {
  const ss            = getSpreadsheet();
  const sheetFacturas = ss.getSheetByName(SHEET_NAME_FACTURAS);
  const sheetPagos    = ss.getSheetByName(SHEET_NAME_PAGOS);
  const data          = sheetFacturas.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === idFactura) {
      const nuevoMonto = (parseFloat(data[i][5]) || 0) + parseFloat(monto);
      sheetFacturas.getRange(i + 1, 6).setValue(nuevoMonto);
      const idPago = generarId('PAG');
      sheetPagos.appendRow([idPago, idFactura, parseFloat(monto), new Date(fecha), '', '']);
      registrarMovimiento('Pago', idFactura,
        'Pago registrado: ' + idPago + ' — Fecha: ' + fecha,
        parseFloat(monto), data[i][3], '');
      return { ok: true, idPago };
    }
  }
  return { ok: false, error: 'Factura no encontrada' };
}

function editarPago(idPago, nuevoMonto, nuevaFecha, justificacion) {
  const ss            = getSpreadsheet();
  const sheetPagos    = ss.getSheetByName(SHEET_NAME_PAGOS);
  const sheetFacturas = ss.getSheetByName(SHEET_NAME_FACTURAS);
  const rows          = sheetPagos.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === idPago) {
      const montoAnterior = parseFloat(rows[i][2]) || 0;
      const idFactura     = rows[i][1];
      const diferencia    = parseFloat(nuevoMonto) - montoAnterior;

      sheetPagos.getRange(i + 1, 3).setValue(parseFloat(nuevoMonto));
      sheetPagos.getRange(i + 1, 4).setValue(new Date(nuevaFecha));
      sheetPagos.getRange(i + 1, 6).setValue(justificacion || '');

      const rowsFact = sheetFacturas.getDataRange().getValues();
      for (let j = 1; j < rowsFact.length; j++) {
        if (rowsFact[j][0] === idFactura) {
          const montoPagadoActual = parseFloat(rowsFact[j][5]) || 0;
          sheetFacturas.getRange(j + 1, 6).setValue(montoPagadoActual + diferencia);
          registrarMovimiento('Edición pago', idFactura,
            'Pago ' + idPago + ' editado: ' + montoAnterior + ' → ' + nuevoMonto,
            parseFloat(nuevoMonto), rowsFact[j][3], justificacion);
          break;
        }
      }
      return { ok: true };
    }
  }
  return { ok: false, error: 'Pago no encontrado' };
}

function getPagosPorFactura(idFactura) {
  const sheet = getSpreadsheet().getSheetByName(SHEET_NAME_PAGOS);
  const data  = sheet.getDataRange().getValues();
  const pagos = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === idFactura) {
      pagos.push({
        id            : data[i][0],
        monto         : parseFloat(data[i][2]) || 0,
        fecha         : data[i][3] ? formatDate(data[i][3]) : '',
        fechaRaw      : data[i][3] ? formatDateISO(data[i][3]) : '',
        docUrl        : data[i][4] || '',
        justificacion : data[i][5] || ''
      });
    }
  }
  return pagos;
}

function getProveedores() {
  const sheet = getSpreadsheet().getSheetByName(SHEET_NAME_PROVEEDORES);
  const data  = sheet.getDataRange().getValues();
  return data.slice(1).filter(r => r[1]).map(r => ({ id: r[0], nombre: r[1], pais: r[2] }));
}

function getPaises() {
  const sheet = getSpreadsheet().getSheetByName(SHEET_NAME_PAISES);
  const data  = sheet.getDataRange().getValues();
  return data.slice(1).filter(r => r[0]).map(r => r[0]);
}

function getResumen() {
  const facturas = getFacturas(null);
  const resumen  = {};
  facturas.forEach(f => {
    if (!resumen[f.moneda]) resumen[f.moneda] = { pendiente: 0, vencido: 0, pagado: 0, total: 0 };
    resumen[f.moneda].total += f.montoTotal;
    if (f.estado === 'Pagada')   resumen[f.moneda].pagado    += f.montoPagado;
    if (f.estado === 'Vencida')  resumen[f.moneda].vencido   += f.saldo;
    if (f.estado === 'Pendiente' || f.estado === 'Parcial') resumen[f.moneda].pendiente += f.saldo;
  });
  return resumen;
}

function getRootFolder() {
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(DRIVE_FOLDER_NAME);
}

function getPaisFolder(rootFolder, pais) {
  const folders = rootFolder.getFoldersByName(pais);
  if (folders.hasNext()) return folders.next();
  return rootFolder.createFolder(pais);
}

function getProveedorFolder(paisFolder, idProveedor, nombreProveedor) {
  const nombre  = idProveedor + '_' + nombreProveedor;
  const folders = paisFolder.getFoldersByName(nombre);
  if (folders.hasNext()) return folders.next();
  return paisFolder.createFolder(nombre);
}

function subirDocumento(idFactura, nombreArchivo, base64Data, mimeType) {
  const ss    = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_FACTURAS);
  const rows  = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === idFactura) {
      const pais = rows[i][2], proveedor = rows[i][1];
      const sheetProv = ss.getSheetByName(SHEET_NAME_PROVEEDORES);
      const provData  = sheetProv.getDataRange().getValues();
      let idProveedor = 'PROV';
      for (let j = 1; j < provData.length; j++) {
        if (provData[j][1] === proveedor) { idProveedor = provData[j][0] || 'PROV-' + j; break; }
      }
      const root    = getRootFolder();
      const blob    = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, idFactura + '_' + nombreArchivo);
      const archivo = getProveedorFolder(getPaisFolder(root, pais), idProveedor, proveedor).createFile(blob);
      archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      const url = archivo.getUrl();
      sheet.getRange(i + 1, 10).setValue(url);
      registrarMovimiento('Documento', idFactura, 'Documento adjuntado: ' + nombreArchivo, 0, '', '');
      return { ok: true, url };
    }
  }
  return { ok: false, error: 'Factura no encontrada' };
}

function eliminarDocumento(idFactura) {
  const ss    = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_FACTURAS);
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === idFactura) {
      sheet.getRange(i + 1, 10).setValue('');
      registrarMovimiento('Documento eliminado', idFactura, 'Documento eliminado de factura', 0, '', '');
      return { ok: true };
    }
  }
  return { ok: false };
}

function subirDocumentoPago(idPago, nombreArchivo, base64Data, mimeType, idFactura) {
  const ss        = getSpreadsheet();
  const sheet     = ss.getSheetByName(SHEET_NAME_PAGOS);
  const sheetFact = ss.getSheetByName(SHEET_NAME_FACTURAS);
  const rows      = sheet.getDataRange().getValues();
  const rowsFact  = sheetFact.getDataRange().getValues();

  let pais = 'General', proveedor = 'General', idProveedor = 'PROV';
  for (let i = 1; i < rowsFact.length; i++) {
    if (rowsFact[i][0] === idFactura) {
      pais = rowsFact[i][2]; proveedor = rowsFact[i][1];
      const sheetProv = ss.getSheetByName(SHEET_NAME_PROVEEDORES);
      const provData  = sheetProv.getDataRange().getValues();
      for (let j = 1; j < provData.length; j++) {
        if (provData[j][1] === proveedor) { idProveedor = provData[j][0] || 'PROV-' + j; break; }
      }
      break;
    }
  }

  const root    = getRootFolder();
  const blob    = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, idPago + '_comprobante_' + nombreArchivo);
  const archivo = getProveedorFolder(getPaisFolder(root, pais), idProveedor, proveedor).createFile(blob);
  archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const url = archivo.getUrl();

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === idPago) {
      sheet.getRange(i + 1, 5).setValue(url);
      registrarMovimiento('Comprobante pago', idFactura, 'Comprobante adjuntado al pago ' + idPago, 0, '', '');
      return { ok: true, url };
    }
  }
  return { ok: false, error: 'Pago no encontrado' };
}

function calcularEstado(factura, hoy) {
  if (factura.montoPagado >= factura.montoTotal) return 'Pagada';
  if (!factura.fechaVence) return 'Pendiente';
  const vence = new Date(factura.fechaVence);
  vence.setHours(0, 0, 0, 0);
  if (vence < hoy)             return 'Vencida';
  if (factura.montoPagado > 0) return 'Parcial';
  return 'Pendiente';
}

function calcularDias(fechaVence, hoy) {
  if (!fechaVence) return null;
  const vence = new Date(fechaVence);
  vence.setHours(0, 0, 0, 0);
  return Math.round((vence - hoy) / (1000 * 60 * 60 * 24));
}

function formatDate(date) {
  if (!date) return '';
  const d = new Date(date);
  return d.toLocaleDateString('es-CR', { day: '2-digit', month: '2-digit', year: 'numeric' });
}

function formatDateTime(date) {
  if (!date) return '';
  const d = new Date(date);
  return d.toLocaleDateString('es-CR', { day: '2-digit', month: '2-digit', year: 'numeric' }) +
    ' ' + d.toLocaleTimeString('es-CR', { hour: '2-digit', minute: '2-digit' });
}

function formatDateISO(date) {
  if (!date) return '';
  const d = new Date(date);
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return y + '-' + m + '-' + day;
}

function generarId(prefijo) {
  const ts = new Date().getTime().toString().slice(-6);
  return prefijo + '-' + ts;
}
