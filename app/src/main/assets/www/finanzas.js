// Reglas de negocio: saldos, reportes, deudas, cuotas y metas.
//
// Tipos de movimiento:
//   ingreso        +monto en la cuenta (cuenta como ingreso en los reportes)
//   egreso         −monto en la cuenta (cuenta como gasto en los reportes)
//   transferencia  −monto en cuenta origen, +monto en cuenta destino (no es ingreso ni gasto)
//   interno        +/−delta en la cuenta (préstamos, cobros, aportes a metas…): mueve el saldo
//                  pero NO cuenta como ingreso ni gasto, porque la plata sigue siendo tuya
//                  o es un préstamo que se devuelve.

import { datos, guardar, borrar, nuevoId } from './db.js';

/* ======================= fechas y formato ======================= */

export const hoy = () => fechaISO(new Date());
export function fechaISO(d) {
  const p = (n) => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
}
export const aFecha = (iso) => {
  const [y, m, d] = iso.split('-').map(Number);
  return new Date(y, m - 1, d);
};
export const mesDe = (iso) => iso.slice(0, 7); // "2026-11"

export function sumarPeriodo(iso, frecuencia, n) {
  const d = aFecha(iso);
  if (frecuencia === 'semanal') d.setDate(d.getDate() + 7 * n);
  else if (frecuencia === 'quincenal') d.setDate(d.getDate() + 15 * n);
  else {
    const dia = d.getDate();
    d.setDate(1);
    d.setMonth(d.getMonth() + n);
    const ultimo = new Date(d.getFullYear(), d.getMonth() + 1, 0).getDate();
    d.setDate(Math.min(dia, ultimo)); // 31 de enero -> 28/29 de febrero
  }
  return fechaISO(d);
}

export const diasEntre = (desde, hasta) => Math.round((aFecha(hasta) - aFecha(desde)) / 86400000);

const fmtCOP = new Intl.NumberFormat('es-CO', { style: 'currency', currency: 'COP', maximumFractionDigits: 0 });
export const dinero = (v) => fmtCOP.format(Math.round(Number(v) || 0)).replace(/ /g, ' ');
export const dineroCorto = (v) => {
  const a = Math.abs(v), s = v < 0 ? '−' : '';
  if (a >= 1e9) return `${s}$${(a / 1e9).toFixed(1).replace('.', ',')} mil M`;
  if (a >= 1e6) return `${s}$${(a / 1e6).toFixed(1).replace('.', ',').replace(',0', '')} M`;
  if (a >= 1e3) return `${s}$${Math.round(a / 1e3)} mil`;
  return `${s}$${Math.round(a)}`;
};

const MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
export const nombreMes = (ym, corto = false) => {
  const [y, m] = ym.split('-').map(Number);
  const n = MESES[m - 1];
  return corto ? n.slice(0, 3) : `${n[0].toUpperCase()}${n.slice(1)} ${y}`;
};
export const fechaLarga = (iso) => {
  const d = aFecha(iso);
  const dif = diasEntre(hoy(), iso);
  if (dif === 0) return 'Hoy';
  if (dif === -1) return 'Ayer';
  if (dif === 1) return 'Mañana';
  return d.toLocaleDateString('es-CO', { weekday: 'short', day: 'numeric', month: 'short', year: d.getFullYear() === new Date().getFullYear() ? undefined : 'numeric' });
};
export const sumarMeses = (ym, n) => {
  const [y, m] = ym.split('-').map(Number);
  const d = new Date(y, m - 1 + n, 1);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
};

/* ======================= datos iniciales ======================= */

export const CATEGORIAS_INICIALES = [
  ['egreso', 'Comida', '🍽️', '#E07A3F'], ['egreso', 'Mercado', '🛒', '#D4A017'], ['egreso', 'Transporte', '🚌', '#3D7DD8'],
  ['egreso', 'Vivienda', '🏠', '#7A5AC8'], ['egreso', 'Servicios', '💡', '#2FA4A9'], ['egreso', 'Salud', '🩺', '#D2475B'],
  ['egreso', 'Educación', '📚', '#5468D4'], ['egreso', 'Ropa', '👕', '#C15DA4'], ['egreso', 'Entretenimiento', '🎬', '#8E44AD'],
  ['egreso', 'Celular e internet', '📱', '#4A90A4'], ['egreso', 'Otros gastos', '📦', '#7D8590'],
  ['ingreso', 'Salario', '💼', '#2E9E62'], ['ingreso', 'Ventas', '🧾', '#1F8A70'], ['ingreso', 'Proyectos', '💻', '#4B3FA8'],
  ['ingreso', 'Regalos', '🎁', '#C9A227'], ['ingreso', 'Otros ingresos', '💵', '#5E8C61'],
];
export const CUENTAS_INICIALES = [['Efectivo', 'efectivo', '#2E9E62'], ['Banco', 'banco', '#3D7DD8'], ['Nequi', 'billetera', '#B83B8F']];

export async function sembrarSiVacio() {
  if (datos.categorias.length || datos.cuentas.length) return;
  for (const [tipo, nombre, icono, color] of CATEGORIAS_INICIALES) await guardar('categorias', { tipo, nombre, icono, color, presupuesto: 0 });
  for (const [nombre, tipo, color] of CUENTAS_INICIALES) await guardar('cuentas', { nombre, tipo, color, saldoInicial: 0, archivada: false });
}

/* ======================= consultas ======================= */

export const categoria = (id) => datos.categorias.find((c) => c.id === id);
export const cuenta = (id) => datos.cuentas.find((c) => c.id === id);
export const cuentasActivas = () => datos.cuentas.filter((c) => !c.archivada);

/** Efecto de un movimiento en el saldo de una cuenta. */
export function efecto(m, cuentaId) {
  if (m.tipo === 'ingreso') return m.cuentaId === cuentaId ? m.monto : 0;
  if (m.tipo === 'egreso') return m.cuentaId === cuentaId ? -m.monto : 0;
  if (m.tipo === 'transferencia') return (m.cuentaId === cuentaId ? -m.monto : 0) + (m.cuentaDestinoId === cuentaId ? m.monto : 0);
  if (m.tipo === 'interno') return m.cuentaId === cuentaId ? m.delta : 0;
  return 0;
}

export function saldoCuenta(id) {
  const c = cuenta(id);
  let s = Number(c?.saldoInicial) || 0;
  for (const m of datos.movimientos) s += efecto(m, id);
  return s;
}

export const saldoTotal = () => cuentasActivas().reduce((t, c) => t + saldoCuenta(c.id), 0);

export function resumenMes(ym) {
  let ingresos = 0, egresos = 0;
  const porCategoria = new Map();
  for (const m of datos.movimientos) {
    if (mesDe(m.fecha) !== ym) continue;
    if (m.tipo === 'ingreso') ingresos += m.monto;
    if (m.tipo === 'egreso') {
      egresos += m.monto;
      porCategoria.set(m.categoriaId, (porCategoria.get(m.categoriaId) || 0) + m.monto);
    }
  }
  const categorias = [...porCategoria.entries()]
    .map(([id, total]) => ({ cat: categoria(id) || { nombre: 'Sin categoría', icono: '❔', color: '#999' }, id, total }))
    .sort((a, b) => b.total - a.total);
  return { ingresos, egresos, balance: ingresos - egresos, categorias };
}

export function serieMeses(n = 6, hasta = mesDe(hoy())) {
  const out = [];
  for (let i = n - 1; i >= 0; i--) {
    const ym = sumarMeses(hasta, -i);
    const r = resumenMes(ym);
    out.push({ ym, ingresos: r.ingresos, egresos: r.egresos });
  }
  return out;
}

export function gastoCategoriaMes(catId, ym) {
  return datos.movimientos.reduce((t, m) => t + (m.tipo === 'egreso' && m.categoriaId === catId && mesDe(m.fecha) === ym ? m.monto : 0), 0);
}

/* ======================= movimientos ======================= */

export async function guardarMovimiento(m) {
  m.monto = Math.abs(Number(m.monto) || 0);
  if (m.tipo === 'interno') m.delta = Number(m.delta) || 0;
  return guardar('movimientos', m);
}

/** Borra un movimiento y, si venía de una deuda, cuota o meta, también ese registro. */
export async function borrarMovimiento(id) {
  const m = datos.movimientos.find((x) => x.id === id);
  if (!m) return;
  if (m.origen) {
    const { tipo, id: oid, sub } = m.origen;
    const reg = datos[tipo === 'cuota' ? 'cuotas' : tipo === 'meta' ? 'metas' : 'deudas'].find((x) => x.id === oid);
    if (reg) {
      if (tipo === 'deuda') {
        if (sub === 'inicial') reg.movimientoId = null;
        else reg.abonos = (reg.abonos || []).filter((a) => a.movimientoId !== id);
        await guardar('deudas', reg);
      } else if (tipo === 'cuota') {
        reg.pagos = (reg.pagos || []).filter((p) => p.movimientoId !== id);
        await guardar('cuotas', reg);
      } else if (tipo === 'meta') {
        reg.aportes = (reg.aportes || []).filter((a) => a.movimientoId !== id);
        await guardar('metas', reg);
      }
    }
  }
  await borrar('movimientos', id);
}

/* ======================= deudas (me deben / debo) ======================= */

export const abonado = (d) => (d.abonos || []).reduce((t, a) => t + a.monto, 0);
export const pendienteDeuda = (d) => Math.max(0, d.monto - abonado(d));

/**
 * Crea o actualiza una deuda. Si se elige una cuenta, registra la salida (si prestaste)
 * o la entrada (si te prestaron) como movimiento interno.
 */
export async function guardarDeuda(d, cuentaId) {
  d.abonos = d.abonos || [];
  d.monto = Math.abs(Number(d.monto) || 0);
  const nueva = !d.id;
  if (nueva) d.id = nuevoId();
  if (nueva && cuentaId) {
    const mov = await guardarMovimiento({
      tipo: 'interno', cuentaId, fecha: d.fecha, monto: d.monto,
      delta: d.direccion === 'me_deben' ? -d.monto : d.monto,
      nota: d.direccion === 'me_deben' ? `Préstamo a ${d.persona}` : `Préstamo de ${d.persona}`,
      origen: { tipo: 'deuda', id: d.id, sub: 'inicial' },
    });
    d.movimientoId = mov.id;
  } else if (!nueva && d.movimientoId) {
    // al editar la deuda, el movimiento del préstamo se ajusta al nuevo valor
    const mov = datos.movimientos.find((x) => x.id === d.movimientoId);
    if (mov) {
      mov.monto = d.monto;
      mov.fecha = d.fecha;
      mov.delta = d.direccion === 'me_deben' ? -d.monto : d.monto;
      await guardar('movimientos', mov);
    }
  }
  d.cerrada = pendienteDeuda(d) <= 0;
  return guardar('deudas', d);
}

export async function abonarDeuda(d, monto, fecha, cuentaId, nota = '') {
  monto = Math.abs(Number(monto) || 0);
  const abono = { id: nuevoId(), monto, fecha, nota, movimientoId: null };
  if (cuentaId) {
    const mov = await guardarMovimiento({
      tipo: 'interno', cuentaId, fecha, monto,
      delta: d.direccion === 'me_deben' ? monto : -monto,
      nota: d.direccion === 'me_deben' ? `${d.persona} me pagó` : `Pago a ${d.persona}`,
      origen: { tipo: 'deuda', id: d.id, sub: 'abono' },
    });
    abono.movimientoId = mov.id;
  }
  d.abonos = [...(d.abonos || []), abono];
  d.cerrada = pendienteDeuda(d) <= 0;
  return guardar('deudas', d);
}

export async function borrarDeuda(d) {
  const ids = [d.movimientoId, ...(d.abonos || []).map((a) => a.movimientoId)].filter(Boolean);
  for (const id of ids) await borrar('movimientos', id);
  await borrar('deudas', d.id);
}

/* ======================= compras a cuotas ======================= */

export const fechaCuota = (c, n) => sumarPeriodo(c.primeraFecha, c.frecuencia, n - 1);
export const cuotasPagadas = (c) => (c.pagos || []).length;
export const pagadoCuotas = (c) => (c.pagos || []).reduce((t, p) => t + p.monto, 0);
export const totalCuotas = (c) => c.valorCuota * c.numCuotas;
export const saldoCuotas = (c) => Math.max(0, totalCuotas(c) - pagadoCuotas(c));
export function siguienteCuota(c) {
  const n = cuotasPagadas(c) + 1;
  if (n > c.numCuotas) return null;
  return { n, fecha: fechaCuota(c, n), monto: c.valorCuota };
}

/** Paga la siguiente cuota: queda como un gasto de la categoría elegida. */
export async function pagarCuota(c, { fecha, monto, cuentaId }) {
  const sig = siguienteCuota(c);
  if (!sig) throw new Error('Ya pagaste todas las cuotas');
  monto = Math.abs(Number(monto) || sig.monto);
  const mov = await guardarMovimiento({
    tipo: 'egreso', categoriaId: c.categoriaId, cuentaId, fecha, monto,
    nota: `${c.nombre} · cuota ${sig.n} de ${c.numCuotas}`,
    origen: { tipo: 'cuota', id: c.id },
  });
  c.pagos = [...(c.pagos || []), { n: sig.n, fecha, monto, movimientoId: mov.id }];
  return guardar('cuotas', c);
}

export async function borrarPlanCuotas(c) {
  for (const p of c.pagos || []) if (p.movimientoId) await borrar('movimientos', p.movimientoId);
  await borrar('cuotas', c.id);
}

/* ======================= metas de ahorro ======================= */

export const ahorrado = (m) => (m.aportes || []).reduce((t, a) => t + a.monto, 0);

export async function aportarMeta(meta, monto, fecha, cuentaId, retiro = false) {
  monto = Math.abs(Number(monto) || 0);
  const valor = retiro ? -monto : monto;
  const aporte = { id: nuevoId(), monto: valor, fecha, movimientoId: null };
  if (cuentaId) {
    const mov = await guardarMovimiento({
      tipo: 'interno', cuentaId, fecha, monto, delta: -valor,
      nota: retiro ? `Retiro de la meta "${meta.nombre}"` : `Ahorro para "${meta.nombre}"`,
      origen: { tipo: 'meta', id: meta.id },
    });
    aporte.movimientoId = mov.id;
  }
  meta.aportes = [...(meta.aportes || []), aporte];
  return guardar('metas', meta);
}

/** Cuánto ahorrar por mes para llegar a tiempo. */
export function ahorroSugerido(meta) {
  const falta = meta.objetivo - ahorrado(meta);
  if (falta <= 0 || !meta.fechaLimite) return null;
  const meses = Math.max(1, Math.ceil(diasEntre(hoy(), meta.fechaLimite) / 30.44));
  return { mensual: falta / meses, meses };
}

export async function borrarMeta(meta) {
  for (const a of meta.aportes || []) if (a.movimientoId) await borrar('movimientos', a.movimientoId);
  await borrar('metas', meta.id);
}

/* ======================= próximos pagos ======================= */

export function proximosVencimientos(dias = 30) {
  const limite = fechaISO(new Date(Date.now() + dias * 86400000));
  const out = [];
  for (const c of datos.cuotas) {
    const s = siguienteCuota(c);
    if (s && s.fecha <= limite) out.push({ tipo: 'cuota', fecha: s.fecha, titulo: c.nombre, detalle: `Cuota ${s.n} de ${c.numCuotas}`, monto: s.monto, ref: c });
  }
  for (const d of datos.deudas) {
    if (d.cerrada || !d.vence || d.vence > limite) continue;
    out.push({
      tipo: d.direccion, fecha: d.vence, titulo: d.persona,
      detalle: d.direccion === 'debo' ? `Le debes · ${d.concepto || 'deuda'}` : `Te debe · ${d.concepto || 'préstamo'}`,
      monto: pendienteDeuda(d), ref: d,
    });
  }
  return out.sort((a, b) => a.fecha.localeCompare(b.fecha));
}
