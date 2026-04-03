
const SHEET_NAME_FACTURAS    = 'Facturas';
const SHEET_NAME_PROVEEDORES = 'Proveedores';
const SHEET_NAME_PAISES      = 'Países';
const SHEET_NAME_PAGOS       = 'Pagos';
const SHEET_NAME_CONFIG      = 'Config';

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
      id          : row[0],
      proveedor   : row[1],
      pais        : row[2],
      moneda      : row[3],
      montoTotal  : parseFloat(row[4]) || 0,
      montoPagado : parseFloat(row[5]) || 0,
      fechaEmision: row[6] ? formatDate(row[6]) : '',
      fechaVence  : row[7] ? formatDate(row[7]) : '',
      notas       : row[8] || '',
      fila        : i + 1
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
  sheet.appendRow([id, data.proveedor, data.pais, data.moneda, parseFloat(data.montoTotal), 0, new Date(data.fechaEmision), new Date(data.fechaVence), data.notas || '']);
  return { ok: true, id };
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
      sheetPagos.appendRow([idPago, idFactura, parseFloat(monto), new Date(fecha)]);
      return { ok: true, idPago };
    }
  }
  return { ok: false, error: 'Factura no encontrada' };
}

function getProveedores() {
  const sheet = getSpreadsheet().getSheetByName(SHEET_NAME_PROVEEDORES);
  const data  = sheet.getDataRange().getValues();
  return data.slice(1).filter(r => r[0]).map(r => ({ id: r[0], nombre: r[1], pais: r[2] }));
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

function generarId(prefijo) {
  const ts = new Date().getTime().toString().slice(-6);
  return prefijo + '-' + ts;
}
