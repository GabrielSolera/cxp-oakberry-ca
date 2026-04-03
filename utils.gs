// ============================================================
// utils.js — Helpers y funciones de soporte
// ============================================================

function formatCurrency(monto, moneda) {
  const simbolos = { USD: '$', CRC: '₡', PAB: 'B/.' };
  const simbolo = simbolos[moneda] || '';
  return simbolo + monto.toLocaleString('es-CR', { minimumFractionDigits: 2 });
}

function getDiasLabel(dias) {
  if (dias === null) return '';
  if (dias < 0)  return `Vencida hace ${Math.abs(dias)}d`;
  if (dias === 0) return 'Vence hoy';
  return `Vence en ${dias}d`;
}

function getEstadoColor(estado) {
  const colores = {
    'Pagada'    : 'success',
    'Pendiente' : 'warning',
    'Vencida'   : 'danger',
    'Parcial'   : 'info'
  };
  return colores[estado] || 'secondary';
}

function validarFactura(data) {
  const errores = [];
  if (!data.proveedor)    errores.push('Proveedor requerido');
  if (!data.pais)         errores.push('País requerido');
  if (!data.moneda)       errores.push('Moneda requerida');
  if (!data.montoTotal || data.montoTotal <= 0) errores.push('Monto inválido');
  if (!data.fechaEmision) errores.push('Fecha de emisión requerida');
  if (!data.fechaVence)   errores.push('Fecha de vencimiento requerida');
  return errores;
}