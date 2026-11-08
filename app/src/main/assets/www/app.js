import * as db from './db.js';
import { datos } from './db.js';
import * as F from './finanzas.js';
import { exportarExcel, guardarRespaldo, leerRespaldo } from './export.js';

const $ = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => [...r.querySelectorAll(s)];
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

const COLORES = ['#4B3FA8', '#3D7DD8', '#2FA4A9', '#1C9A62', '#5E8C61', '#D4A017', '#E07A3F', '#D2453F', '#C15DA4', '#8E44AD', '#7D8590', '#17181D'];
const EMOJIS = ['🍽️', '🛒', '🚌', '🚗', '⛽', '🏠', '💡', '💧', '📱', '🩺', '💊', '📚', '👕', '💇', '🎬', '🎮', '✈️', '🐶', '👶', '🎁', '🏋️', '☕', '🍺', '🧾',
  '💼', '💻', '💵', '🏦', '📈', '🛠️', '🧹', '🎓', '❤️', '⭐', '📦', '🔁'];

const ICO = {
  gear: '<svg viewBox="0 0 24 24"><g fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.7 1.7 0 0 0 .3 1.8l.1.1a2 2 0 1 1-2.8 2.8l-.1-.1a1.7 1.7 0 0 0-1.8-.3 1.7 1.7 0 0 0-1 1.5V21a2 2 0 1 1-4 0v-.1a1.7 1.7 0 0 0-1.1-1.5 1.7 1.7 0 0 0-1.8.3l-.1.1a2 2 0 1 1-2.8-2.8l.1-.1a1.7 1.7 0 0 0 .3-1.8 1.7 1.7 0 0 0-1.5-1H3a2 2 0 1 1 0-4h.1a1.7 1.7 0 0 0 1.5-1.1 1.7 1.7 0 0 0-.3-1.8l-.1-.1a2 2 0 1 1 2.8-2.8l.1.1a1.7 1.7 0 0 0 1.8.3H9a1.7 1.7 0 0 0 1-1.5V3a2 2 0 1 1 4 0v.1a1.7 1.7 0 0 0 1 1.5 1.7 1.7 0 0 0 1.8-.3l.1-.1a2 2 0 1 1 2.8 2.8l-.1.1a1.7 1.7 0 0 0-.3 1.8V9a1.7 1.7 0 0 0 1.5 1H21a2 2 0 1 1 0 4h-.1a1.7 1.7 0 0 0-1.5 1z"/></g></svg>',
  back: '<svg viewBox="0 0 24 24"><path d="M15 5l-7 7 7 7" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"/></svg>',
  left: '<svg viewBox="0 0 24 24" width="20" height="20"><path d="M15 5l-7 7 7 7" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"/></svg>',
  right: '<svg viewBox="0 0 24 24" width="20" height="20"><path d="M9 5l7 7-7 7" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"/></svg>',
  search: '<svg viewBox="0 0 24 24"><g fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="11" cy="11" r="7"/><path d="M20 20l-4-4"/></g></svg>',
  transfer: '🔁', internal: '🤝', goal: '🎯',
};

/* ======================= estado ======================= */

const st = {
  vista: 'inicio',
  mes: F.mesDe(F.hoy()),
  filtroTipo: 'todos',
  filtroCuenta: null,
  filtroCategoria: null,
  busqueda: '',
  deudasTab: 'me_deben',
  verCerradas: false,
  catTab: 'egreso',
};

/* ======================= utilidades de interfaz ======================= */

let toastTimer;
function toast(msg) {
  const t = $('#toast');
  t.textContent = msg;
  t.hidden = false;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => (t.hidden = true), 2600);
}

let confirmResolve = null;
function confirmar(texto, si = 'Eliminar', peligro = true) {
  return new Promise((resolve) => {
    $('#confirmText').textContent = texto;
    $('#confirmYes').textContent = si;
    $('#confirmYes').className = 'btn ' + (peligro ? 'danger' : 'primary');
    $('#confirm').hidden = false;
    confirmResolve = resolve;
  });
}
const cerrarConfirm = (v) => {
  $('#confirm').hidden = true;
  const r = confirmResolve;
  confirmResolve = null;
  r && r(v);
};
$('#confirmYes').addEventListener('click', () => cerrarConfirm(true));
$('#confirmNo').addEventListener('click', () => cerrarConfirm(false));

/* Panel inferior con historial: un detalle puede abrir un formulario y al cerrarlo se vuelve al detalle. */
const pila = [];
function abrirPanel(render) {
  pila.push({ render, mem: {} }); // mem conserva lo escrito si el panel se vuelve a pintar
  pintarPanel();
}
function pintarPanel() {
  const actual = pila[pila.length - 1];
  if (!actual) {
    $('#sheetWrap').hidden = true;
    return;
  }
  $('#sheetWrap').hidden = false;
  const sheet = $('#sheet');
  sheet.innerHTML = '<div class="grip"></div>';
  actual.render(sheet, actual.mem);
  sheet.scrollTop = 0;
}
function atrasPanel() {
  pila.pop();
  pintarPanel();
}
function cerrarPaneles() {
  pila.length = 0;
  pintarPanel();
}
$('#sheetWrap').addEventListener('click', (e) => {
  if (e.target.id === 'sheetWrap') atrasPanel();
});

/** Guarda, cierra el panel actual y refresca la pantalla. */
async function listo(msg, cerrarTodo = false) {
  if (cerrarTodo) cerrarPaneles(); else atrasPanel();
  render();
  if (msg) toast(msg);
}

/* ---------- campos ---------- */

const parseMonto = (s) => Number(String(s ?? '').replace(/[^\d]/g, '')) || 0;
const formatoMiles = (n) => (n ? Math.round(n).toLocaleString('es-CO') : '');

function campoMonto(id, valor = 0, clase = '') {
  return `<div class="money ${clase}"><span>$</span><input class="input num" id="${id}" inputmode="numeric" autocomplete="off" placeholder="0" value="${formatoMiles(valor)}"></div>`;
}
function activarMontos(raiz) {
  $$('.money input', raiz).forEach((inp) => {
    inp.addEventListener('input', () => {
      const n = parseMonto(inp.value);
      inp.value = formatoMiles(n);
    });
  });
}
function opcionesCuentas(sel, { ninguna = false, textoNinguna = 'No mover ninguna cuenta' } = {}) {
  const ops = F.cuentasActivas().map((c) => `<option value="${c.id}" ${c.id === sel ? 'selected' : ''}>${esc(c.nombre)} · ${F.dinero(F.saldoCuenta(c.id))}</option>`);
  if (ninguna) ops.unshift(`<option value="" ${!sel ? 'selected' : ''}>${textoNinguna}</option>`);
  return ops.join('');
}
const cuentaDefecto = () => db.ajuste('cuentaDefecto') || F.cuentasActivas()[0]?.id || '';
const inicial = (s) => (String(s || '?').trim()[0] || '?').toUpperCase();
const colorPersona = (s) => COLORES[[...String(s)].reduce((h, c) => h + c.charCodeAt(0), 0) % 9];

/* ======================= gráficos (SVG) ======================= */

function dona(partes, total) {
  const r = 52, C = 2 * Math.PI * r;
  let acum = 0;
  const arcos = partes.map((p) => {
    const frac = total ? p.total / total : 0;
    const s = `<circle r="${r}" cx="66" cy="66" fill="none" stroke="${p.cat.color}" stroke-width="20" stroke-dasharray="${(frac * C).toFixed(2)} ${C.toFixed(2)}" stroke-dashoffset="${(-acum * C).toFixed(2)}" transform="rotate(-90 66 66)"/>`;
    acum += frac;
    return s;
  }).join('');
  return `<svg class="donut" viewBox="0 0 132 132">${total ? arcos : `<circle r="${r}" cx="66" cy="66" fill="none" stroke="#E7E3DA" stroke-width="20"/>`}
    <text x="66" y="62" text-anchor="middle" font-size="11" fill="#6E6C66">Gastos</text>
    <text x="66" y="80" text-anchor="middle" font-size="15" font-weight="700" fill="#17181D">${esc(F.dineroCorto(total))}</text></svg>`;
}

function barras(serie) {
  const W = 320, H = 150, top = 10, base = 124, max = Math.max(1, ...serie.flatMap((s) => [s.ingresos, s.egresos]));
  const ancho = W / serie.length;
  const bw = Math.min(16, ancho / 3.2);
  const y = (v) => base - (v / max) * (base - top);
  const cuerpo = serie.map((s, i) => {
    const cx = ancho * i + ancho / 2;
    const actual = s.ym === F.mesDe(F.hoy());
    return `<rect x="${cx - bw - 2}" y="${y(s.ingresos)}" width="${bw}" height="${base - y(s.ingresos)}" rx="4" fill="#1C9A62"/>
      <rect x="${cx + 2}" y="${y(s.egresos)}" width="${bw}" height="${base - y(s.egresos)}" rx="4" fill="#D2453F"/>
      <text x="${cx}" y="142" text-anchor="middle" font-size="11" fill="${actual ? '#17181D' : '#6E6C66'}" font-weight="${actual ? 700 : 500}">${F.nombreMes(s.ym, true)}</text>`;
  }).join('');
  return `<svg class="bars" viewBox="0 0 ${W} ${H}" preserveAspectRatio="none"><line x1="0" x2="${W}" y1="${base}" y2="${base}" stroke="#E7E3DA"/>${cuerpo}</svg>`;
}

function anillo(frac, color) {
  const r = 24, C = 2 * Math.PI * r, f = Math.max(0, Math.min(1, frac));
  return `<svg class="goal-ring" viewBox="0 0 58 58"><circle r="${r}" cx="29" cy="29" fill="none" stroke="#ECE8DF" stroke-width="7"/>
    <circle r="${r}" cx="29" cy="29" fill="none" stroke="${color}" stroke-width="7" stroke-linecap="round" stroke-dasharray="${(f * C).toFixed(2)} ${C}" transform="rotate(-90 29 29)"/>
    <text x="29" y="33" text-anchor="middle" font-size="12" font-weight="700" fill="#17181D">${Math.round(f * 100)}%</text></svg>`;
}

/* ======================= filas ======================= */

function filaMovimiento(m) {
  let ico, color, titulo, sub, monto, clase;
  const c = F.cuenta(m.cuentaId);
  if (m.tipo === 'transferencia') {
    ico = ICO.transfer; color = '#7D8590';
    titulo = m.nota || 'Transferencia';
    sub = `${c?.nombre || '—'} → ${F.cuenta(m.cuentaDestinoId)?.nombre || '—'}`;
    monto = F.dinero(m.monto); clase = '';
  } else if (m.tipo === 'interno') {
    const meta = m.origen?.tipo === 'meta';
    ico = meta ? ICO.goal : ICO.internal; color = meta ? '#4B3FA8' : '#C98A0B';
    titulo = m.nota || (meta ? 'Ahorro' : 'Préstamo');
    sub = `${meta ? 'Meta de ahorro' : 'Deuda'} · ${c?.nombre || '—'}`;
    monto = (m.delta >= 0 ? '+' : '−') + F.dinero(Math.abs(m.delta)); clase = m.delta >= 0 ? 'up' : 'down';
  } else {
    const cat = F.categoria(m.categoriaId);
    ico = cat?.icono || '❔'; color = cat?.color || '#999';
    titulo = m.nota || cat?.nombre || 'Sin categoría';
    sub = `${m.nota ? (cat?.nombre || 'Sin categoría') + ' · ' : ''}${c?.nombre || '—'}`;
    monto = (m.tipo === 'ingreso' ? '+' : '−') + F.dinero(m.monto); clase = m.tipo === 'ingreso' ? 'up' : 'down';
  }
  return `<button class="row" data-mov="${m.id}">
    <span class="chip-ico" style="--c:${color}">${ico}</span>
    <span class="row-main"><span class="row-title">${esc(titulo)}</span><span class="row-sub">${esc(sub)}</span></span>
    <span class="row-amt ${clase}">${monto}</span></button>`;
}

function enlazarMovimientos(raiz) {
  $$('[data-mov]', raiz).forEach((b) => b.addEventListener('click', () => {
    const m = datos.movimientos.find((x) => x.id === b.dataset.mov);
    if (m) abrirPanel((s, mem) => formMovimiento(s, m, {}, mem));
  }));
}

const switchMes = () => `<div class="month-switch"><button data-mes="-1" aria-label="Mes anterior">${ICO.left}</button><b>${F.nombreMes(st.mes)}</b><button data-mes="1" aria-label="Mes siguiente">${ICO.right}</button></div>`;
function enlazarMes(raiz) {
  $$('[data-mes]', raiz).forEach((b) => b.addEventListener('click', () => {
    st.mes = F.sumarMeses(st.mes, Number(b.dataset.mes));
    render();
  }));
}

/* ======================= pantalla: inicio ======================= */

function vistaInicio(v) {
  const r = F.resumenMes(st.mes);
  const total = F.saldoTotal();
  const proximos = F.proximosVencimientos(30);
  const presup = datos.categorias.filter((c) => c.tipo === 'egreso' && c.presupuesto > 0);
  const metas = datos.metas.filter((m) => F.ahorrado(m) < m.objetivo).slice(0, 3);
  const recientes = datos.movimientos.slice().sort((a, b) => b.fecha.localeCompare(a.fecha) || (b.creado || 0) - (a.creado || 0)).slice(0, 5);
  const meDeben = datos.deudas.filter((d) => d.direccion === 'me_deben').reduce((t, d) => t + F.pendienteDeuda(d), 0);
  const deboPrestamos = datos.deudas.filter((d) => d.direccion === 'debo').reduce((t, d) => t + F.pendienteDeuda(d), 0);
  const deboCuotas = datos.cuotas.reduce((t, c) => t + F.saldoCuotas(c), 0);
  const debo = deboPrestamos + deboCuotas;
  const propio = total - debo;

  v.innerHTML = `
    <div class="head">
      <div><h1>Mis cuentas</h1><p class="sub">${new Date().toLocaleDateString('es-CO', { weekday: 'long', day: 'numeric', month: 'long' })}</p></div>
      <button class="icon-btn" id="btnAjustes" aria-label="Ajustes">${ICO.gear}</button>
    </div>
    <div class="hero">
      <p class="hero-label">SALDO DISPONIBLE · EN MIS CUENTAS</p>
      <p class="hero-amount num">${F.dinero(total)}</p>
      <button class="hero-propio" id="btnPropio">
        <span>Lo que es mío<small>${debo ? 'descontando lo que debo' : 'no tienes deudas'}</small></span>
        <b class="num">${F.dinero(propio)}</b><i>›</i>
      </button>
      <div class="hero-row">
        <div class="hero-box"><span><span class="arrow">▲</span>Ingresos del mes</span><b class="num">${F.dinero(r.ingresos)}</b></div>
        <div class="hero-box"><span><span class="arrow">▼</span>Gastos del mes</span><b class="num">${F.dinero(r.egresos)}</b></div>
      </div>
    </div>
    ${switchMes()}

    <div class="section">
      <div class="totals" style="margin:0">
        <button class="card" data-ir="deudas:me_deben" style="text-align:left"><span>Me deben</span><b class="up">${F.dinero(meDeben)}</b></button>
        <button class="card" data-ir="deudas:debo" style="text-align:left"><span>Debo (incluye cuotas)</span><b class="down">${F.dinero(debo)}</b></button>
      </div>

      ${proximos.length ? `
      <div class="section-title"><h2>Próximos pagos y cobros</h2><span class="muted" style="font-size:13px">30 días</span></div>
      <div class="card" style="padding:4px 14px">
        ${proximos.slice(0, 5).map((p) => {
          const dias = F.diasEntre(F.hoy(), p.fecha);
          const tag = dias < 0 ? `<span class="tag down">Vencido</span>` : dias <= 3 ? `<span class="tag warn">${F.fechaLarga(p.fecha)}</span>` : `<span class="muted" style="font-size:12px">${F.fechaLarga(p.fecha)}</span>`;
          return `<button class="row" data-prox="${p.tipo}:${p.ref.id}">
            <span class="chip-ico" style="--c:${p.tipo === 'me_deben' ? '#1C9A62' : '#D2453F'}">${p.tipo === 'cuota' ? '💳' : p.tipo === 'debo' ? '📤' : '📥'}</span>
            <span class="row-main"><span class="row-title">${esc(p.titulo)}</span><span class="row-sub">${esc(p.detalle)}</span></span>
            <span class="row-amt ${p.tipo === 'me_deben' ? 'up' : ''}">${F.dinero(p.monto)}<small>${tag}</small></span></button>`;
        }).join('')}
      </div>` : ''}

      <div class="section-title"><h2>Gastos de ${F.nombreMes(st.mes).split(' ')[0].toLowerCase()}</h2></div>
      <div class="card">
        ${r.egresos ? `<div class="donut-wrap">${dona(r.categorias, r.egresos)}
          <div class="legend">${r.categorias.slice(0, 6).map((c) => `<div><i style="--c:${c.cat.color}"></i><span>${esc(c.cat.icono)} ${esc(c.cat.nombre)}</span><b>${Math.round(c.total / r.egresos * 100)}%</b></div>`).join('')}
          ${r.categorias.length > 6 ? `<div class="muted">+ ${r.categorias.length - 6} más</div>` : ''}</div></div>`
          : `<div class="empty"><b>Sin gastos este mes</b>Registra uno con el botón ＋</div>`}
      </div>

      ${presup.length ? `
      <div class="section-title"><h2>Presupuestos</h2><button class="link" data-ir="ajustes">Editar</button></div>
      <div class="card">
        ${presup.map((c) => {
          const g = F.gastoCategoriaMes(c.id, st.mes), f = g / c.presupuesto;
          return `<div style="margin-bottom:12px"><div style="display:flex;justify-content:space-between;font-size:14px"><span>${esc(c.icono)} ${esc(c.nombre)}</span>
            <span class="num ${f > 1 ? 'down' : ''}">${F.dinero(g)} <span class="muted">de ${F.dinero(c.presupuesto)}</span></span></div>
            <div class="bar ${f > 1 ? 'over' : ''}" style="--c:${c.color}"><i style="width:${Math.min(100, f * 100)}%"></i></div></div>`;
        }).join('')}
      </div>` : ''}

      <div class="section-title"><h2>Últimos 6 meses</h2></div>
      <div class="card">${barras(F.serieMeses(6, st.mes))}
        <div class="chart-key"><span><i style="background:#1C9A62"></i>Ingresos</span><span><i style="background:#D2453F"></i>Gastos</span></div></div>

      <div class="section-title"><h2>Mis cuentas</h2><button class="link" data-ir="ajustes">Administrar</button></div>
      <div class="card" style="padding:4px 14px">
        ${F.cuentasActivas().map((c) => `<button class="row" data-cuenta="${c.id}">
          <span class="chip-ico" style="--c:${c.color}">${c.tipo === 'banco' ? '🏦' : c.tipo === 'billetera' ? '📲' : c.tipo === 'tarjeta' ? '💳' : '💵'}</span>
          <span class="row-main"><span class="row-title">${esc(c.nombre)}</span></span>
          <span class="row-amt ${F.saldoCuenta(c.id) < 0 ? 'down' : ''}">${F.dinero(F.saldoCuenta(c.id))}</span></button>`).join('') || '<div class="empty">Agrega una cuenta en Ajustes</div>'}
      </div>

      ${metas.length ? `
      <div class="section-title"><h2>Metas de ahorro</h2><button class="link" data-ir="metas">Ver todas</button></div>
      <div class="card" style="padding:4px 14px">
        ${metas.map((m) => `<button class="row" data-meta="${m.id}">${anillo(F.ahorrado(m) / m.objetivo, m.color)}
          <span class="row-main"><span class="row-title">${esc(m.icono)} ${esc(m.nombre)}</span><span class="row-sub">${F.dinero(F.ahorrado(m))} de ${F.dinero(m.objetivo)}</span></span></button>`).join('')}
      </div>` : ''}

      <div class="section-title"><h2>Últimos movimientos</h2><button class="link" data-ir="movimientos">Ver todos</button></div>
      <div class="card" style="padding:4px 14px">${recientes.map(filaMovimiento).join('') || `<div class="empty"><b>Aún no hay movimientos</b>Toca ＋ para registrar tu primer ingreso o gasto.</div>`}</div>
    </div>`;

  $('#btnAjustes', v).addEventListener('click', () => irA('ajustes'));
  $('#btnPropio', v).addEventListener('click', () => abrirPanel((s) => detallePropio(s, { total, deboPrestamos, deboCuotas, meDeben })));
  enlazarMes(v);
  enlazarMovimientos(v);
  $$('[data-ir]', v).forEach((b) => b.addEventListener('click', () => {
    const [vista, tab] = b.dataset.ir.split(':');
    if (tab) st.deudasTab = tab;
    irA(vista);
  }));
  $$('[data-cuenta]', v).forEach((b) => b.addEventListener('click', () => {
    st.filtroCuenta = b.dataset.cuenta;
    st.filtroCategoria = null;
    irA('movimientos');
  }));
  $$('[data-meta]', v).forEach((b) => b.addEventListener('click', () => abrirPanel((s) => detalleMeta(s, b.dataset.meta))));
  $$('[data-prox]', v).forEach((b) => b.addEventListener('click', () => {
    const [tipo, id] = b.dataset.prox.split(':');
    if (tipo === 'cuota') abrirPanel((s) => detalleCuota(s, id));
    else abrirPanel((s) => detalleDeuda(s, id));
  }));
}

function detallePropio(s, { total, deboPrestamos, deboCuotas, meDeben }) {
  const propio = total - deboPrestamos - deboCuotas;
  const linea = (t, v, clase = '') => `<div class="row"><span class="row-main"><span class="row-title" style="font-weight:500">${t}</span></span><span class="row-amt ${clase}">${v}</span></div>`;
  s.innerHTML += `<h3>¿Cuánto es mío?</h3>
    <p class="hint" style="margin:0 0 12px">El saldo disponible es la plata que hay en tus cuentas, incluida la que te prestaron. Lo que es tuyo es eso menos lo que debes.</p>
    <div class="detail-list">
      ${linea('En mis cuentas', F.dinero(total))}
      ${linea('Préstamos que debo', '−' + F.dinero(deboPrestamos), 'down')}
      ${linea('Cuotas por pagar', '−' + F.dinero(deboCuotas), 'down')}
      <div class="row" style="border-top:2px solid var(--ink)"><span class="row-main"><span class="row-title">Lo que es mío</span></span><span class="row-amt ${propio < 0 ? 'down' : 'up'}" style="font-size:18px">${F.dinero(propio)}</span></div>
    </div>
    ${meDeben ? `<p class="hint" style="margin-top:12px">Además te deben ${F.dinero(meDeben)}. Si te pagan todo, tendrías ${F.dinero(propio + meDeben)}.</p>` : ''}
    <div class="sheet-actions"><button class="btn ghost" id="pdCerrar">Cerrar</button><button class="btn soft" id="pdDeudas">Ver deudas</button></div>`;
  $('#pdCerrar', s).addEventListener('click', () => atrasPanel());
  $('#pdDeudas', s).addEventListener('click', () => { st.deudasTab = 'debo'; cerrarPaneles(); irA('deudas'); });
}

/* ======================= pantalla: movimientos ======================= */

function vistaMovimientos(v) {
  const q = st.busqueda.trim().toLowerCase();
  let lista = datos.movimientos.filter((m) => F.mesDe(m.fecha) === st.mes);
  if (st.filtroTipo !== 'todos') lista = lista.filter((m) => m.tipo === st.filtroTipo);
  if (st.filtroCuenta) lista = lista.filter((m) => m.cuentaId === st.filtroCuenta || m.cuentaDestinoId === st.filtroCuenta);
  if (st.filtroCategoria) lista = lista.filter((m) => m.categoriaId === st.filtroCategoria);
  if (q) {
    lista = lista.filter((m) => [m.nota, F.categoria(m.categoriaId)?.nombre, F.cuenta(m.cuentaId)?.nombre, String(m.monto)]
      .some((t) => String(t || '').toLowerCase().includes(q)));
  }
  lista.sort((a, b) => b.fecha.localeCompare(a.fecha) || (b.creado || 0) - (a.creado || 0));
  const ing = lista.reduce((t, m) => t + (m.tipo === 'ingreso' ? m.monto : 0), 0);
  const egr = lista.reduce((t, m) => t + (m.tipo === 'egreso' ? m.monto : 0), 0);

  const dias = new Map();
  for (const m of lista) {
    if (!dias.has(m.fecha)) dias.set(m.fecha, []);
    dias.get(m.fecha).push(m);
  }
  const pills = [['todos', 'Todos'], ['egreso', 'Gastos'], ['ingreso', 'Ingresos'], ['transferencia', 'Transferencias'], ['interno', 'Préstamos y ahorro']];
  const fc = st.filtroCuenta && F.cuenta(st.filtroCuenta), fcat = st.filtroCategoria && F.categoria(st.filtroCategoria);

  v.innerHTML = `
    <div class="head"><div><h1>Movimientos</h1></div>
      <button class="icon-btn" id="btnExportarMes" aria-label="Exportar a Excel"><svg viewBox="0 0 24 24"><g fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 4v11M7 10l5 5 5-5M5 20h14"/></g></svg></button></div>
    ${switchMes()}
    <div class="search">${ICO.search}<input id="busca" placeholder="Buscar por nota, categoría o valor" value="${esc(st.busqueda)}"></div>
    <div class="filters">
      ${fc ? `<button class="pill on" data-quitar="cuenta">${esc(fc.nombre)} ✕</button>` : ''}
      ${fcat ? `<button class="pill on" data-quitar="cat">${esc(fcat.icono)} ${esc(fcat.nombre)} ✕</button>` : ''}
      ${pills.map(([k, l]) => `<button class="pill ${st.filtroTipo === k ? 'on' : ''}" data-tipo="${k}">${l}</button>`).join('')}
    </div>
    <div class="totals">
      <div class="card"><span>Ingresos</span><b class="up">${F.dinero(ing)}</b></div>
      <div class="card"><span>Gastos</span><b class="down">${F.dinero(egr)}</b></div>
    </div>
    <div class="section">
      ${[...dias.entries()].map(([fecha, movs]) => {
        const neto = movs.reduce((t, m) => t + (m.tipo === 'ingreso' ? m.monto : m.tipo === 'egreso' ? -m.monto : 0), 0);
        return `<div class="day-label"><span>${F.fechaLarga(fecha)}</span><span class="num">${neto ? (neto > 0 ? '+' : '−') + F.dinero(Math.abs(neto)) : ''}</span></div>
          <div class="card" style="padding:2px 14px">${movs.map(filaMovimiento).join('')}</div>`;
      }).join('') || `<div class="empty"><b>No hay movimientos</b>${q || st.filtroTipo !== 'todos' || fc || fcat ? 'Prueba quitando los filtros.' : 'Registra uno con el botón ＋.'}</div>`}
    </div>`;

  enlazarMes(v);
  enlazarMovimientos(v);
  const busca = $('#busca', v);
  busca.addEventListener('input', () => {
    st.busqueda = busca.value;
    const pos = busca.selectionStart;
    render();
    const nb = $('#busca');
    nb.focus();
    nb.setSelectionRange(pos, pos);
  });
  $$('[data-tipo]', v).forEach((b) => b.addEventListener('click', () => { st.filtroTipo = b.dataset.tipo; render(); }));
  $$('[data-quitar]', v).forEach((b) => b.addEventListener('click', () => {
    if (b.dataset.quitar === 'cuenta') st.filtroCuenta = null; else st.filtroCategoria = null;
    render();
  }));
  $('#btnExportarMes', v).addEventListener('click', () => exportar());
}

/* ======================= formulario de movimiento ======================= */

function formMovimiento(s, mov = null, preset = {}, mem = {}) {
  if (mov && mov.tipo === 'interno') return detalleInterno(s, mov);
  const m = mem.m || (mem.m = mov ? { ...mov } : { tipo: preset.tipo || 'egreso', monto: 0, fecha: F.hoy(), cuentaId: cuentaDefecto(), categoriaId: null, nota: '' });
  const pintar = () => {
    const cats = datos.categorias.filter((c) => c.tipo === m.tipo).sort((a, b) => usoCategoria(b.id) - usoCategoria(a.id));
    if (m.tipo !== 'transferencia' && !cats.some((c) => c.id === m.categoriaId)) m.categoriaId = null;
    s.innerHTML = `<div class="grip"></div>
      <h3>${mov ? 'Editar movimiento' : 'Nuevo movimiento'}</h3>
      <div class="segment" style="margin:0 0 14px">
        ${[['egreso', 'Gasto'], ['ingreso', 'Ingreso'], ['transferencia', 'Transferencia']].map(([k, l]) => `<button data-t="${k}" class="${m.tipo === k ? 'on' : ''}">${l}</button>`).join('')}
      </div>
      <div class="field">${campoMonto('fMonto', m.monto, m.tipo === 'ingreso' ? 'in' : m.tipo === 'egreso' ? 'out' : '')}</div>
      ${m.tipo === 'transferencia' ? `
        <div class="grid2">
          <div class="field"><label>Desde</label><select class="input" id="fCuenta">${opcionesCuentas(m.cuentaId)}</select></div>
          <div class="field"><label>Hacia</label><select class="input" id="fDestino">${opcionesCuentas(m.cuentaDestinoId || F.cuentasActivas().find((c) => c.id !== m.cuentaId)?.id)}</select></div>
        </div>` : `
        <div class="field"><label>Categoría</label>
          <div class="cat-grid">
            ${cats.map((c) => `<button class="cat-opt ${m.categoriaId === c.id ? 'on' : ''}" data-cat="${c.id}"><span class="chip-ico" style="--c:${c.color}">${esc(c.icono)}</span>${esc(c.nombre)}</button>`).join('')}
            <button class="cat-opt add" id="fNuevaCat"><span class="chip-ico" style="--c:#4B3FA8">＋</span>Nueva</button>
          </div></div>
        <div class="field"><label>${m.tipo === 'ingreso' ? 'Entra a' : 'Sale de'}</label><select class="input" id="fCuenta">${opcionesCuentas(m.cuentaId)}</select></div>`}
      <div class="grid2">
        <div class="field"><label>Fecha</label><input class="input" type="date" id="fFecha" value="${m.fecha}"></div>
        <div class="field"><label>Nota</label><input class="input" id="fNota" maxlength="80" placeholder="Opcional" value="${esc(m.nota)}"></div>
      </div>
      ${m.origen?.tipo === 'cuota' ? '<p class="hint">Este gasto es el pago de una cuota.</p>' : ''}
      <div class="sheet-actions">
        ${mov ? '<button class="btn ghost" id="fBorrar" style="flex:0 0 auto;color:var(--down)">Eliminar</button>' : ''}
        <button class="btn primary" id="fGuardar">Guardar</button>
      </div>`;
    activarMontos(s);
    const leer = () => {
      m.monto = parseMonto($('#fMonto', s).value);
      m.fecha = $('#fFecha', s).value || F.hoy();
      m.nota = $('#fNota', s).value.trim();
      m.cuentaId = $('#fCuenta', s)?.value || '';
      if (m.tipo === 'transferencia') m.cuentaDestinoId = $('#fDestino', s).value;
    };
    $$('[data-t]', s).forEach((b) => b.addEventListener('click', () => { leer(); m.tipo = b.dataset.t; pintar(); }));
    $$('[data-cat]', s).forEach((b) => b.addEventListener('click', () => {
      m.categoriaId = b.dataset.cat;
      $$('[data-cat]', s).forEach((x) => x.classList.toggle('on', x === b));
    }));
    $('#fNuevaCat', s)?.addEventListener('click', () => {
      leer();
      abrirPanel((s2) => formCategoria(s2, null, m.tipo, (cat) => { m.categoriaId = cat.id; }));
    });
    $('#fBorrar', s)?.addEventListener('click', async () => {
      if (!(await confirmar('¿Eliminar este movimiento?'))) return;
      await F.borrarMovimiento(mov.id);
      listo('Movimiento eliminado');
    });
    $('#fGuardar', s).addEventListener('click', async () => {
      leer();
      if (!m.monto) return toast('Escribe el valor');
      if (!F.cuentasActivas().length) return toast('Primero crea una cuenta en Ajustes');
      if (m.tipo === 'transferencia') {
        if (m.cuentaId === m.cuentaDestinoId) return toast('Elige dos cuentas distintas');
        m.categoriaId = null;
      } else {
        delete m.cuentaDestinoId;
        if (!m.categoriaId) return toast('Elige una categoría');
      }
      if (!m.creado) m.creado = Date.now();
      await F.guardarMovimiento(m);
      // si es el pago de una cuota, el plan se actualiza con el valor real pagado
      if (m.origen?.tipo === 'cuota') {
        const c = datos.cuotas.find((x) => x.id === m.origen.id);
        const p = c?.pagos.find((x) => x.movimientoId === m.id);
        if (p) { p.monto = m.monto; p.fecha = m.fecha; await db.guardar('cuotas', c); }
      }
      await db.fijarAjuste('cuentaDefecto', m.tipo === 'transferencia' ? cuentaDefecto() : m.cuentaId);
      listo(mov ? 'Movimiento actualizado' : 'Movimiento guardado');
    });
    if (!mov) setTimeout(() => $('#fMonto', s)?.focus(), 80);
  };
  pintar();
}

const usoCategoria = (id) => datos.movimientos.reduce((t, m) => t + (m.categoriaId === id ? 1 : 0), 0);

function detalleInterno(s, m) {
  const c = F.cuenta(m.cuentaId);
  const deMeta = m.origen?.tipo === 'meta';
  s.innerHTML += `<h3>${esc(m.nota || 'Movimiento')}</h3>
    <div class="big-amt ${m.delta >= 0 ? 'up' : 'down'}">${m.delta >= 0 ? '+' : '−'}${F.dinero(Math.abs(m.delta))}</div>
    <p class="muted" style="margin:0 0 14px">${F.fechaLarga(m.fecha)} · ${esc(c?.nombre || 'Sin cuenta')}</p>
    <p class="hint" style="margin-bottom:16px">Este movimiento viene de ${deMeta ? 'una meta de ahorro' : 'una deuda'}. Mueve el saldo de tu cuenta, pero no cuenta como ingreso ni gasto porque ${deMeta ? 'la plata sigue siendo tuya' : 'es un préstamo'}.</p>
    <div class="sheet-actions">
      <button class="btn ghost" id="dBorrar" style="color:var(--down)">Eliminar</button>
      <button class="btn soft" id="dVer">Ver ${deMeta ? 'meta' : 'deuda'}</button>
    </div>`;
  $('#dBorrar', s).addEventListener('click', async () => {
    if (!(await confirmar(`¿Eliminar este movimiento? También se quita el ${deMeta ? 'aporte de la meta' : 'registro en la deuda'}.`))) return;
    await F.borrarMovimiento(m.id);
    listo('Movimiento eliminado');
  });
  $('#dVer', s).addEventListener('click', () => {
    const id = m.origen?.id;
    if (deMeta) abrirPanel((s2) => detalleMeta(s2, id)); else abrirPanel((s2) => detalleDeuda(s2, id));
  });
}

/* ======================= pantalla: deudas y cuotas ======================= */

function vistaDeudas(v) {
  const tab = st.deudasTab;
  let cuerpo = '', totales = '';
  if (tab === 'cuotas') {
    const planes = datos.cuotas.filter((c) => st.verCerradas || F.siguienteCuota(c));
    const saldo = datos.cuotas.reduce((t, c) => t + F.saldoCuotas(c), 0);
    const mesCuotas = datos.cuotas.reduce((t, c) => {
      const sig = F.siguienteCuota(c);
      return t + (sig && F.mesDe(sig.fecha) <= F.mesDe(F.hoy()) ? sig.monto : 0);
    }, 0);
    totales = `<div class="card"><span>Saldo por pagar</span><b class="down">${F.dinero(saldo)}</b></div><div class="card"><span>Por pagar este mes</span><b>${F.dinero(mesCuotas)}</b></div>`;
    cuerpo = planes.map((c) => {
      const sig = F.siguienteCuota(c), pag = F.cuotasPagadas(c), cat = F.categoria(c.categoriaId);
      const dias = sig ? F.diasEntre(F.hoy(), sig.fecha) : null;
      return `<button class="card" data-cuota="${c.id}" style="display:block;width:100%;text-align:left">
        <div class="debt-top"><span class="chip-ico" style="--c:${cat?.color || '#4B3FA8'}">${esc(cat?.icono || '💳')}</span>
          <span class="row-main"><span class="row-title">${esc(c.nombre)}</span><span class="row-sub">${pag} de ${c.numCuotas} cuotas · ${F.dinero(c.valorCuota)} c/u</span></span>
          <span class="row-amt">${F.dinero(F.saldoCuotas(c))}<small>por pagar</small></span></div>
        <div class="bar" style="--c:${cat?.color || '#4B3FA8'}"><i style="width:${pag / c.numCuotas * 100}%"></i></div>
        <div class="debt-meta"><span>${sig ? `Próxima: ${F.fechaLarga(sig.fecha)}` : '¡Pagada!'}</span>
          ${sig ? (dias < 0 ? '<span class="tag down">Vencida</span>' : dias <= 5 ? '<span class="tag warn">Pronto</span>' : '') : '<span class="tag up">Completa</span>'}</div>
      </button>`;
    }).join('') || `<div class="empty"><b>Sin compras a cuotas</b>Registra créditos, compras con tarjeta o cualquier cosa que pagues por partes.</div>`;
  } else {
    const lista = datos.deudas.filter((d) => d.direccion === tab && (st.verCerradas || !d.cerrada))
      .sort((a, b) => (a.cerrada - b.cerrada) || (a.vence || '9999').localeCompare(b.vence || '9999'));
    const pendiente = datos.deudas.filter((d) => d.direccion === tab).reduce((t, d) => t + F.pendienteDeuda(d), 0);
    const personas = new Set(datos.deudas.filter((d) => d.direccion === tab && !d.cerrada).map((d) => d.persona.trim().toLowerCase())).size;
    totales = `<div class="card"><span>${tab === 'me_deben' ? 'Te deben en total' : 'Debes en total'}</span><b class="${tab === 'me_deben' ? 'up' : 'down'}">${F.dinero(pendiente)}</b></div>
      <div class="card"><span>Personas</span><b>${personas}</b></div>`;
    cuerpo = lista.map((d) => {
      const pend = F.pendienteDeuda(d), frac = d.monto ? F.abonado(d) / d.monto : 0;
      const dias = d.vence ? F.diasEntre(F.hoy(), d.vence) : null;
      return `<button class="card" data-deuda="${d.id}" style="display:block;width:100%;text-align:left;${d.cerrada ? 'opacity:.6' : ''}">
        <div class="debt-top"><span class="avatar" style="--c:${colorPersona(d.persona)}">${esc(inicial(d.persona))}</span>
          <span class="row-main"><span class="row-title">${esc(d.persona)}</span><span class="row-sub">${esc(d.concepto || (tab === 'me_deben' ? 'Préstamo' : 'Deuda'))}</span></span>
          <span class="row-amt ${tab === 'me_deben' ? 'up' : 'down'}">${F.dinero(pend)}<small>de ${F.dinero(d.monto)}</small></span></div>
        <div class="bar" style="--c:${tab === 'me_deben' ? '#1C9A62' : '#D2453F'}"><i style="width:${frac * 100}%"></i></div>
        <div class="debt-meta"><span>${d.cerrada ? 'Saldada' : d.vence ? `Vence ${F.fechaLarga(d.vence).toLowerCase()}` : 'Sin fecha límite'}</span>
          ${d.cerrada ? '<span class="tag up">Pagada</span>' : dias !== null && dias < 0 ? '<span class="tag down">Vencida</span>' : dias !== null && dias <= 5 ? '<span class="tag warn">Pronto</span>' : ''}</div>
      </button>`;
    }).join('') || `<div class="empty"><b>${tab === 'me_deben' ? 'Nadie te debe' : 'No debes nada'}</b>${tab === 'me_deben' ? 'Anota aquí la plata que prestas.' : 'Anota aquí lo que te prestan.'}</div>`;
  }

  v.innerHTML = `
    <div class="head"><div><h1>Deudas</h1><p class="sub">Préstamos y compras a cuotas</p></div></div>
    <div class="segment">
      <button data-dt="me_deben" class="${tab === 'me_deben' ? 'on' : ''}">Me deben</button>
      <button data-dt="debo" class="${tab === 'debo' ? 'on' : ''}">Debo</button>
      <button data-dt="cuotas" class="${tab === 'cuotas' ? 'on' : ''}">Cuotas</button>
    </div>
    <div class="totals">${totales}</div>
    <div class="section">
      <button class="btn primary block" id="btnNuevaDeuda" style="margin:4px 0 14px">＋ ${tab === 'cuotas' ? 'Nueva compra a cuotas' : tab === 'me_deben' ? 'Presté plata' : 'Me prestaron plata'}</button>
      ${cuerpo}
      <div style="text-align:center;margin-top:10px"><button class="link" id="btnCerradas">${st.verCerradas ? 'Ocultar las ya pagadas' : 'Mostrar también las ya pagadas'}</button></div>
    </div>`;

  $$('[data-dt]', v).forEach((b) => b.addEventListener('click', () => { st.deudasTab = b.dataset.dt; render(); }));
  $('#btnCerradas', v).addEventListener('click', () => { st.verCerradas = !st.verCerradas; render(); });
  $('#btnNuevaDeuda', v).addEventListener('click', () => {
    if (tab === 'cuotas') abrirPanel((s) => formCuota(s));
    else abrirPanel((s) => formDeuda(s, null, tab));
  });
  $$('[data-deuda]', v).forEach((b) => b.addEventListener('click', () => abrirPanel((s) => detalleDeuda(s, b.dataset.deuda))));
  $$('[data-cuota]', v).forEach((b) => b.addEventListener('click', () => abrirPanel((s) => detalleCuota(s, b.dataset.cuota))));
}

function formDeuda(s, deuda = null, direccion = 'me_deben') {
  const d = deuda ? { ...deuda } : { direccion, persona: '', concepto: '', monto: 0, fecha: F.hoy(), vence: '', notas: '' };
  const meDeben = d.direccion === 'me_deben';
  const personas = [...new Set(datos.deudas.map((x) => x.persona))];
  s.innerHTML += `<h3>${deuda ? 'Editar' : meDeben ? 'Presté plata' : 'Me prestaron plata'}</h3>
    <div class="field"><label>${meDeben ? '¿A quién?' : '¿Quién te prestó?'}</label><input class="input" id="dPersona" list="dPersonas" maxlength="60" value="${esc(d.persona)}" placeholder="Nombre"></div>
    <datalist id="dPersonas">${personas.map((p) => `<option value="${esc(p)}">`).join('')}</datalist>
    <div class="field"><label>Valor</label>${campoMonto('dMonto', d.monto, meDeben ? 'in' : 'out')}</div>
    <div class="field"><label>Concepto</label><input class="input" id="dConcepto" maxlength="80" value="${esc(d.concepto)}" placeholder="Ej: préstamo para el arriendo"></div>
    <div class="grid2">
      <div class="field"><label>Fecha</label><input class="input" type="date" id="dFecha" value="${d.fecha}"></div>
      <div class="field"><label>Fecha límite</label><input class="input" type="date" id="dVence" value="${d.vence || ''}"></div>
    </div>
    ${deuda ? '' : `<div class="field"><label>${meDeben ? '¿De dónde salió la plata?' : '¿A qué cuenta entró?'}</label>
      <select class="input" id="dCuenta">${opcionesCuentas('', { ninguna: true, textoNinguna: 'No mover ninguna cuenta (solo anotar)' })}</select>
      <p class="hint">Si eliges una cuenta, su saldo ${meDeben ? 'baja' : 'sube'}. No cuenta como ${meDeben ? 'gasto' : 'ingreso'}.</p></div>`}
    <div class="field"><label>Notas</label><textarea class="input" id="dNotas" maxlength="300">${esc(d.notas || '')}</textarea></div>
    <div class="sheet-actions"><button class="btn primary" id="dGuardar">Guardar</button></div>`;
  activarMontos(s);
  $('#dGuardar', s).addEventListener('click', async () => {
    d.persona = $('#dPersona', s).value.trim();
    d.monto = parseMonto($('#dMonto', s).value);
    d.concepto = $('#dConcepto', s).value.trim();
    d.fecha = $('#dFecha', s).value || F.hoy();
    d.vence = $('#dVence', s).value || '';
    d.notas = $('#dNotas', s).value.trim();
    if (!d.persona) return toast('Escribe el nombre');
    if (!d.monto) return toast('Escribe el valor');
    if (deuda && d.monto < F.abonado(d)) return toast('El valor no puede ser menor que lo ya abonado');
    await F.guardarDeuda(d, deuda ? null : $('#dCuenta', s).value);
    listo(deuda ? 'Deuda actualizada' : 'Deuda guardada');
  });
  if (!deuda) setTimeout(() => $('#dPersona', s).focus(), 80);
}

function detalleDeuda(s, id) {
  const d = datos.deudas.find((x) => x.id === id);
  if (!d) return atrasPanel();
  const meDeben = d.direccion === 'me_deben', pend = F.pendienteDeuda(d);
  s.innerHTML += `
    <div class="debt-top" style="margin-bottom:12px"><span class="avatar" style="--c:${colorPersona(d.persona)}">${esc(inicial(d.persona))}</span>
      <div class="row-main"><h3 style="margin:0">${esc(d.persona)}</h3><div class="row-sub">${meDeben ? 'Te debe' : 'Le debes'} · ${esc(d.concepto || '')}</div></div></div>
    <div class="card">
      <div class="muted" style="font-size:13px">Pendiente</div>
      <div class="big-amt ${meDeben ? 'up' : 'down'}">${F.dinero(pend)}</div>
      <div class="bar" style="--c:${meDeben ? '#1C9A62' : '#D2453F'}"><i style="width:${d.monto ? F.abonado(d) / d.monto * 100 : 0}%"></i></div>
      <div class="debt-meta"><span>Abonado ${F.dinero(F.abonado(d))}</span><span>Total ${F.dinero(d.monto)}</span></div>
      <div class="debt-meta"><span>Desde ${F.fechaLarga(d.fecha).toLowerCase()}</span><span>${d.vence ? 'Vence ' + F.fechaLarga(d.vence).toLowerCase() : ''}</span></div>
      ${d.notas ? `<p class="hint" style="margin-top:10px">${esc(d.notas)}</p>` : ''}
    </div>
    ${pend > 0 ? `<button class="btn primary block" id="ddAbonar" style="margin-top:12px">${meDeben ? 'Registrar pago que me hicieron' : 'Registrar pago que hice'}</button>` : '<p class="hint" style="text-align:center">✅ Deuda saldada</p>'}
    <div class="section-title"><h2>Abonos</h2></div>
    <div class="detail-list">${(d.abonos || []).slice().reverse().map((a) => `<div class="row"><span class="row-main"><span class="row-title">${F.fechaLarga(a.fecha)}</span>
      <span class="row-sub">${esc(a.nota || (a.movimientoId ? F.cuenta(datos.movimientos.find((m) => m.id === a.movimientoId)?.cuentaId)?.nombre || '' : 'Sin mover cuenta'))}</span></span>
      <span class="row-amt">${F.dinero(a.monto)}</span><button class="link" data-quitar-abono="${a.id}" style="color:var(--down);margin-left:6px">✕</button></div>`).join('') || '<div class="empty" style="padding:14px">Todavía no hay abonos</div>'}</div>
    <div class="sheet-actions">
      <button class="btn ghost" id="ddBorrar" style="color:var(--down)">Eliminar</button>
      <button class="btn ghost" id="ddEditar">Editar</button>
    </div>`;
  $('#ddAbonar', s)?.addEventListener('click', () => abrirPanel((s2) => formAbono(s2, d)));
  $('#ddEditar', s).addEventListener('click', () => abrirPanel((s2) => formDeuda(s2, d)));
  $('#ddBorrar', s).addEventListener('click', async () => {
    if (!(await confirmar(`¿Eliminar la deuda con ${d.persona}? También se borran sus movimientos en las cuentas.`))) return;
    await F.borrarDeuda(d);
    listo('Deuda eliminada', true);
  });
  $$('[data-quitar-abono]', s).forEach((b) => b.addEventListener('click', async () => {
    if (!(await confirmar('¿Quitar este abono?', 'Quitar'))) return;
    const a = d.abonos.find((x) => x.id === b.dataset.quitarAbono);
    if (a.movimientoId) await db.borrar('movimientos', a.movimientoId);
    d.abonos = d.abonos.filter((x) => x.id !== a.id);
    d.cerrada = F.pendienteDeuda(d) <= 0;
    await db.guardar('deudas', d);
    pintarPanel();
    render();
  }));
}

function formAbono(s, d) {
  const meDeben = d.direccion === 'me_deben', pend = F.pendienteDeuda(d);
  s.innerHTML += `<h3>${meDeben ? `${esc(d.persona)} te pagó` : `Le pagaste a ${esc(d.persona)}`}</h3>
    <div class="field"><label>Valor del pago</label>${campoMonto('aMonto', pend, meDeben ? 'in' : 'out')}
      <p class="hint">Pendiente: ${F.dinero(pend)}. Puedes registrar un pago parcial.</p></div>
    <div class="grid2">
      <div class="field"><label>Fecha</label><input class="input" type="date" id="aFecha" value="${F.hoy()}"></div>
      <div class="field"><label>${meDeben ? 'Entró a' : 'Salió de'}</label><select class="input" id="aCuenta">${opcionesCuentas(cuentaDefecto(), { ninguna: true, textoNinguna: 'Ninguna' })}</select></div>
    </div>
    <div class="field"><label>Nota</label><input class="input" id="aNota" maxlength="80" placeholder="Opcional"></div>
    <div class="sheet-actions"><button class="btn primary" id="aGuardar">Registrar pago</button></div>`;
  activarMontos(s);
  $('#aGuardar', s).addEventListener('click', async () => {
    const monto = parseMonto($('#aMonto', s).value);
    if (!monto) return toast('Escribe el valor');
    if (monto > pend) return toast(`El pago no puede ser mayor que lo pendiente (${F.dinero(pend)})`);
    await F.abonarDeuda(d, monto, $('#aFecha', s).value || F.hoy(), $('#aCuenta', s).value, $('#aNota', s).value.trim());
    listo(F.pendienteDeuda(d) <= 0 ? '¡Deuda saldada! 🎉' : 'Pago registrado');
  });
}

/* ---------- compras a cuotas ---------- */

function formCuota(s, plan = null) {
  const c = plan ? { ...plan } : { nombre: '', valorCuota: 0, numCuotas: 12, frecuencia: 'mensual', primeraFecha: F.sumarPeriodo(F.hoy(), 'mensual', 1), categoriaId: null, cuentaId: cuentaDefecto(), notas: '', pagos: [] };
  const cats = datos.categorias.filter((x) => x.tipo === 'egreso');
  s.innerHTML += `<h3>${plan ? 'Editar compra a cuotas' : 'Nueva compra a cuotas'}</h3>
    <div class="field"><label>¿Qué compraste o qué crédito es?</label><input class="input" id="cNombre" maxlength="60" value="${esc(c.nombre)}" placeholder="Ej: Celular, crédito moto, tarjeta"></div>
    <div class="grid2">
      <div class="field"><label>Valor de cada cuota</label>${campoMonto('cValor', c.valorCuota, 'out')}</div>
      <div class="field"><label>Número de cuotas</label><input class="input num" id="cNum" type="number" min="1" max="360" inputmode="numeric" value="${c.numCuotas}" style="height:64px;font-size:26px;font-weight:800"></div>
    </div>
    <p class="hint" id="cTotal" style="margin:-6px 2px 12px"></p>
    <div class="grid2">
      <div class="field"><label>Se paga cada</label><select class="input" id="cFrec">
        ${[['mensual', 'Mes'], ['quincenal', 'Quincena'], ['semanal', 'Semana']].map(([k, l]) => `<option value="${k}" ${c.frecuencia === k ? 'selected' : ''}>${l}</option>`).join('')}</select></div>
      <div class="field"><label>${plan && F.cuotasPagadas(plan) ? 'Fecha de la cuota 1' : 'Primera cuota'}</label><input class="input" type="date" id="cFecha" value="${c.primeraFecha}"></div>
    </div>
    <div class="field"><label>Categoría del gasto</label><select class="input" id="cCat">
      ${cats.map((x) => `<option value="${x.id}" ${x.id === c.categoriaId ? 'selected' : ''}>${esc(x.icono)} ${esc(x.nombre)}</option>`).join('')}</select></div>
    <div class="field"><label>Normalmente se paga desde</label><select class="input" id="cCuenta">${opcionesCuentas(c.cuentaId)}</select></div>
    ${plan ? '' : `<div class="field"><label>¿Ya pagaste algunas?</label><input class="input num" type="number" id="cYaPagadas" min="0" value="0" inputmode="numeric">
      <p class="hint">Se marcan como pagadas sin crear gastos (para créditos que ya venías pagando).</p></div>`}
    <div class="sheet-actions"><button class="btn primary" id="cGuardar">Guardar</button></div>`;
  activarMontos(s);
  const total = () => {
    const v = parseMonto($('#cValor', s).value), n = Number($('#cNum', s).value) || 0;
    $('#cTotal', s).textContent = v && n ? `Total a pagar: ${F.dinero(v * n)}` : '';
  };
  $('#cValor', s).addEventListener('input', total);
  $('#cNum', s).addEventListener('input', total);
  total();
  $('#cGuardar', s).addEventListener('click', async () => {
    c.nombre = $('#cNombre', s).value.trim();
    c.valorCuota = parseMonto($('#cValor', s).value);
    c.numCuotas = Math.max(1, Math.min(360, parseInt($('#cNum', s).value, 10) || 0));
    c.frecuencia = $('#cFrec', s).value;
    c.primeraFecha = $('#cFecha', s).value || F.hoy();
    c.categoriaId = $('#cCat', s).value;
    c.cuentaId = $('#cCuenta', s).value;
    if (!c.nombre) return toast('Escribe el nombre');
    if (!c.valorCuota) return toast('Escribe el valor de la cuota');
    if (plan && c.numCuotas < F.cuotasPagadas(plan)) return toast(`Ya hay ${F.cuotasPagadas(plan)} cuotas pagadas`);
    if (!plan) {
      const ya = Math.min(c.numCuotas, Math.max(0, parseInt($('#cYaPagadas', s).value, 10) || 0));
      c.pagos = Array.from({ length: ya }, (_, i) => ({ n: i + 1, fecha: F.fechaCuota(c, i + 1), monto: c.valorCuota, movimientoId: null }));
    }
    await db.guardar('cuotas', c);
    listo(plan ? 'Compra actualizada' : 'Compra a cuotas guardada');
  });
}

function detalleCuota(s, id) {
  const c = datos.cuotas.find((x) => x.id === id);
  if (!c) return atrasPanel();
  const sig = F.siguienteCuota(c), cat = F.categoria(c.categoriaId);
  s.innerHTML += `
    <div class="debt-top" style="margin-bottom:12px"><span class="chip-ico" style="--c:${cat?.color || '#4B3FA8'}">${esc(cat?.icono || '💳')}</span>
      <div class="row-main"><h3 style="margin:0">${esc(c.nombre)}</h3><div class="row-sub">${F.cuotasPagadas(c)} de ${c.numCuotas} cuotas pagadas</div></div></div>
    <div class="card">
      <div class="muted" style="font-size:13px">Saldo por pagar</div>
      <div class="big-amt down">${F.dinero(F.saldoCuotas(c))}</div>
      <div class="bar" style="--c:${cat?.color || '#4B3FA8'}"><i style="width:${F.cuotasPagadas(c) / c.numCuotas * 100}%"></i></div>
      <div class="debt-meta"><span>Pagado ${F.dinero(F.pagadoCuotas(c))}</span><span>Total ${F.dinero(F.totalCuotas(c))}</span></div>
    </div>
    ${sig ? `<button class="btn primary block" id="dcPagar" style="margin-top:12px">Pagar cuota ${sig.n} · ${F.dinero(sig.monto)}</button>` : '<p class="hint" style="text-align:center">✅ Todas las cuotas pagadas</p>'}
    <div class="section-title"><h2>Calendario</h2></div>
    <div class="detail-list schedule">${Array.from({ length: c.numCuotas }, (_, i) => {
      const n = i + 1, p = (c.pagos || []).find((x) => x.n === n);
      return `<div class="${p ? 'paid' : sig && sig.n === n ? 'next' : ''}"><span>Cuota ${n} · ${F.fechaLarga(p ? p.fecha : F.fechaCuota(c, n))}</span><span class="num">${F.dinero(p ? p.monto : c.valorCuota)}</span></div>`;
    }).join('')}</div>
    <div class="sheet-actions">
      <button class="btn ghost" id="dcBorrar" style="color:var(--down)">Eliminar</button>
      ${F.cuotasPagadas(c) ? '<button class="btn ghost" id="dcDeshacer">Deshacer último pago</button>' : ''}
      <button class="btn ghost" id="dcEditar">Editar</button>
    </div>`;
  $('#dcPagar', s)?.addEventListener('click', () => abrirPanel((s2) => formPagoCuota(s2, c)));
  $('#dcEditar', s).addEventListener('click', () => abrirPanel((s2) => formCuota(s2, c)));
  $('#dcDeshacer', s)?.addEventListener('click', async () => {
    if (!(await confirmar('¿Deshacer el último pago? Se borra también el gasto registrado.', 'Deshacer'))) return;
    const p = c.pagos[c.pagos.length - 1];
    if (p.movimientoId) await db.borrar('movimientos', p.movimientoId);
    c.pagos = c.pagos.slice(0, -1);
    await db.guardar('cuotas', c);
    pintarPanel();
    render();
  });
  $('#dcBorrar', s).addEventListener('click', async () => {
    if (!(await confirmar(`¿Eliminar "${c.nombre}"? También se borran los gastos de las cuotas pagadas.`))) return;
    await F.borrarPlanCuotas(c);
    listo('Compra eliminada', true);
  });
}

function formPagoCuota(s, c) {
  const sig = F.siguienteCuota(c);
  s.innerHTML += `<h3>Pagar cuota ${sig.n} de ${c.numCuotas}</h3>
    <div class="field"><label>Valor pagado</label>${campoMonto('pMonto', sig.monto, 'out')}<p class="hint">Cámbialo si pagaste distinto (intereses, abono extra).</p></div>
    <div class="grid2">
      <div class="field"><label>Fecha</label><input class="input" type="date" id="pFecha" value="${F.hoy()}"></div>
      <div class="field"><label>Pagaste desde</label><select class="input" id="pCuenta">${opcionesCuentas(c.cuentaId || cuentaDefecto())}</select></div>
    </div>
    <p class="hint">Se registra como un gasto de ${esc(F.categoria(c.categoriaId)?.nombre || 'la categoría elegida')}.</p>
    <div class="sheet-actions"><button class="btn primary" id="pGuardar">Registrar pago</button></div>`;
  activarMontos(s);
  $('#pGuardar', s).addEventListener('click', async () => {
    const monto = parseMonto($('#pMonto', s).value);
    if (!monto) return toast('Escribe el valor');
    if (!$('#pCuenta', s).value) return toast('Primero crea una cuenta en Ajustes');
    await F.pagarCuota(c, { fecha: $('#pFecha', s).value || F.hoy(), monto, cuentaId: $('#pCuenta', s).value });
    listo(F.siguienteCuota(c) ? 'Cuota pagada' : '¡Terminaste de pagar! 🎉');
  });
}

/* ======================= pantalla: metas ======================= */

function vistaMetas(v) {
  const metas = datos.metas.slice().sort((a, b) => (F.ahorrado(a) >= a.objetivo) - (F.ahorrado(b) >= b.objetivo) || (a.fechaLimite || '9').localeCompare(b.fechaLimite || '9'));
  const totalAhorrado = metas.reduce((t, m) => t + F.ahorrado(m), 0);
  v.innerHTML = `
    <div class="head"><div><h1>Metas de ahorro</h1><p class="sub">Aparta plata para lo que quieres</p></div></div>
    <div class="totals"><div class="card"><span>Ahorrado en metas</span><b class="up">${F.dinero(totalAhorrado)}</b></div>
      <div class="card"><span>Metas activas</span><b>${metas.filter((m) => F.ahorrado(m) < m.objetivo).length}</b></div></div>
    <div class="section">
      <button class="btn primary block" id="btnNuevaMeta" style="margin:4px 0 14px">＋ Nueva meta</button>
      ${metas.map((m) => {
        const ah = F.ahorrado(m), sug = F.ahorroSugerido(m), cumplida = ah >= m.objetivo;
        return `<button class="card" data-meta="${m.id}" style="display:block;width:100%;text-align:left">
          <div class="debt-top">${anillo(ah / m.objetivo, m.color)}
            <span class="row-main"><span class="row-title">${esc(m.icono)} ${esc(m.nombre)}</span>
            <span class="row-sub">${F.dinero(ah)} de ${F.dinero(m.objetivo)}</span>
            <span class="row-sub">${cumplida ? '🎉 ¡Meta cumplida!' : sug ? `Ahorra ${F.dinero(sug.mensual)} al mes · ${sug.meses} ${sug.meses === 1 ? 'mes' : 'meses'}` : m.fechaLimite ? 'Fecha límite cumplida' : `Faltan ${F.dinero(m.objetivo - ah)}`}</span></span></div>
        </button>`;
      }).join('') || `<div class="empty"><b>Sin metas todavía</b>Crea una para un viaje, un fondo de emergencia, un computador…</div>`}
    </div>`;
  $('#btnNuevaMeta', v).addEventListener('click', () => abrirPanel((s) => formMeta(s)));
  $$('[data-meta]', v).forEach((b) => b.addEventListener('click', () => abrirPanel((s) => detalleMeta(s, b.dataset.meta))));
}

function formMeta(s, meta = null) {
  const m = meta ? { ...meta } : { nombre: '', objetivo: 0, fechaLimite: '', icono: '🎯', color: COLORES[0], aportes: [] };
  s.innerHTML += `<h3>${meta ? 'Editar meta' : 'Nueva meta de ahorro'}</h3>
    <div class="field"><label>Nombre</label><input class="input" id="mNombre" maxlength="50" value="${esc(m.nombre)}" placeholder="Ej: Viaje a Cartagena"></div>
    <div class="field"><label>¿Cuánto quieres ahorrar?</label>${campoMonto('mObjetivo', m.objetivo, 'in')}</div>
    <div class="field"><label>¿Para cuándo? (opcional)</label><input class="input" type="date" id="mFecha" value="${m.fechaLimite || ''}"></div>
    <div class="field"><label>Ícono</label><div class="emojis">${['🎯', '✈️', '🏠', '🚗', '💻', '📱', '🎓', '💍', '🏥', '🛟', '🎁', '🏖️', '🐶', '👶', '🏍️', '💰'].map((e) => `<button data-emoji="${e}" class="${m.icono === e ? 'on' : ''}">${e}</button>`).join('')}</div></div>
    <div class="field"><label>Color</label><div class="swatches">${COLORES.slice(0, 10).map((c) => `<button data-color="${c}" style="--c:${c}" class="${m.color === c ? 'on' : ''}" aria-label="Color"></button>`).join('')}</div></div>
    <div class="sheet-actions"><button class="btn primary" id="mGuardar">Guardar</button></div>`;
  activarMontos(s);
  $$('[data-emoji]', s).forEach((b) => b.addEventListener('click', () => { m.icono = b.dataset.emoji; $$('[data-emoji]', s).forEach((x) => x.classList.toggle('on', x === b)); }));
  $$('[data-color]', s).forEach((b) => b.addEventListener('click', () => { m.color = b.dataset.color; $$('[data-color]', s).forEach((x) => x.classList.toggle('on', x === b)); }));
  $('#mGuardar', s).addEventListener('click', async () => {
    m.nombre = $('#mNombre', s).value.trim();
    m.objetivo = parseMonto($('#mObjetivo', s).value);
    m.fechaLimite = $('#mFecha', s).value || '';
    if (!m.nombre) return toast('Escribe el nombre');
    if (!m.objetivo) return toast('Escribe cuánto quieres ahorrar');
    await db.guardar('metas', m);
    listo(meta ? 'Meta actualizada' : 'Meta creada');
  });
}

function detalleMeta(s, id) {
  const m = datos.metas.find((x) => x.id === id);
  if (!m) return atrasPanel();
  const ah = F.ahorrado(m), sug = F.ahorroSugerido(m);
  s.innerHTML += `
    <div class="debt-top" style="margin-bottom:12px">${anillo(ah / m.objetivo, m.color)}
      <div class="row-main"><h3 style="margin:0">${esc(m.icono)} ${esc(m.nombre)}</h3><div class="row-sub">${m.fechaLimite ? 'Para ' + F.fechaLarga(m.fechaLimite).toLowerCase() : 'Sin fecha límite'}</div></div></div>
    <div class="card">
      <div class="muted" style="font-size:13px">Ahorrado</div>
      <div class="big-amt up">${F.dinero(ah)}</div>
      <div class="bar" style="--c:${m.color}"><i style="width:${Math.min(100, ah / m.objetivo * 100)}%"></i></div>
      <div class="debt-meta"><span>Faltan ${F.dinero(Math.max(0, m.objetivo - ah))}</span><span>Meta ${F.dinero(m.objetivo)}</span></div>
      ${sug ? `<p class="hint" style="margin-top:10px">Para llegar a tiempo, ahorra <b>${F.dinero(sug.mensual)}</b> al mes durante ${sug.meses} ${sug.meses === 1 ? 'mes' : 'meses'}.</p>` : ''}
    </div>
    <div class="card-actions">
      <button class="btn primary" id="dmAportar">＋ Ahorrar</button>
      ${ah > 0 ? '<button class="btn soft" id="dmRetirar">Sacar plata</button>' : ''}
    </div>
    <div class="section-title"><h2>Historial</h2></div>
    <div class="detail-list">${(m.aportes || []).slice().reverse().map((a) => `<div class="row"><span class="row-main"><span class="row-title">${a.monto >= 0 ? 'Ahorro' : 'Retiro'}</span>
      <span class="row-sub">${F.fechaLarga(a.fecha)}${a.movimientoId ? ' · ' + esc(F.cuenta(datos.movimientos.find((x) => x.id === a.movimientoId)?.cuentaId)?.nombre || '') : ''}</span></span>
      <span class="row-amt ${a.monto >= 0 ? 'up' : 'down'}">${a.monto >= 0 ? '+' : '−'}${F.dinero(Math.abs(a.monto))}</span>
      <button class="link" data-quitar-aporte="${a.id}" style="color:var(--down);margin-left:6px">✕</button></div>`).join('') || '<div class="empty" style="padding:14px">Aún no has ahorrado para esta meta</div>'}</div>
    <div class="sheet-actions">
      <button class="btn ghost" id="dmBorrar" style="color:var(--down)">Eliminar</button>
      <button class="btn ghost" id="dmEditar">Editar</button>
    </div>`;
  $('#dmAportar', s).addEventListener('click', () => abrirPanel((s2) => formAporte(s2, m, false)));
  $('#dmRetirar', s)?.addEventListener('click', () => abrirPanel((s2) => formAporte(s2, m, true)));
  $('#dmEditar', s).addEventListener('click', () => abrirPanel((s2) => formMeta(s2, m)));
  $('#dmBorrar', s).addEventListener('click', async () => {
    if (!(await confirmar(`¿Eliminar la meta "${m.nombre}"? Lo ahorrado vuelve a sus cuentas.`))) return;
    await F.borrarMeta(m);
    listo('Meta eliminada', true);
  });
  $$('[data-quitar-aporte]', s).forEach((b) => b.addEventListener('click', async () => {
    if (!(await confirmar('¿Quitar este registro?', 'Quitar'))) return;
    const a = m.aportes.find((x) => x.id === b.dataset.quitarAporte);
    if (a.movimientoId) await db.borrar('movimientos', a.movimientoId);
    m.aportes = m.aportes.filter((x) => x.id !== a.id);
    await db.guardar('metas', m);
    pintarPanel();
    render();
  }));
}

function formAporte(s, m, retiro) {
  const ah = F.ahorrado(m);
  s.innerHTML += `<h3>${retiro ? 'Sacar plata de la meta' : `Ahorrar para ${esc(m.nombre)}`}</h3>
    <div class="field"><label>Valor</label>${campoMonto('apMonto', retiro ? 0 : Math.round(F.ahorroSugerido(m)?.mensual || 0), retiro ? 'out' : 'in')}
      ${retiro ? `<p class="hint">Disponible en la meta: ${F.dinero(ah)}</p>` : ''}</div>
    <div class="grid2">
      <div class="field"><label>Fecha</label><input class="input" type="date" id="apFecha" value="${F.hoy()}"></div>
      <div class="field"><label>${retiro ? 'Vuelve a' : 'Sale de'}</label><select class="input" id="apCuenta">${opcionesCuentas(cuentaDefecto(), { ninguna: true, textoNinguna: 'Ninguna' })}</select></div>
    </div>
    <p class="hint">${retiro ? 'La plata vuelve a la cuenta elegida.' : 'El saldo disponible de la cuenta baja; la plata queda apartada en la meta. No cuenta como gasto.'}</p>
    <div class="sheet-actions"><button class="btn primary" id="apGuardar">${retiro ? 'Sacar' : 'Ahorrar'}</button></div>`;
  activarMontos(s);
  $('#apGuardar', s).addEventListener('click', async () => {
    const monto = parseMonto($('#apMonto', s).value);
    if (!monto) return toast('Escribe el valor');
    if (retiro && monto > ah) return toast(`Solo hay ${F.dinero(ah)} en la meta`);
    await F.aportarMeta(m, monto, $('#apFecha', s).value || F.hoy(), $('#apCuenta', s).value, retiro);
    listo(!retiro && F.ahorrado(m) >= m.objetivo ? '¡Cumpliste la meta! 🎉' : retiro ? 'Retiro registrado' : 'Ahorro registrado');
  });
}

/* ======================= pantalla: ajustes ======================= */

function vistaAjustes(v) {
  const cats = datos.categorias.filter((c) => c.tipo === st.catTab).sort((a, b) => a.nombre.localeCompare(b.nombre, 'es'));
  const pin = !!db.ajuste('pinHash');
  v.innerHTML = `
    <div class="head"><button class="icon-btn" id="btnVolver" aria-label="Volver">${ICO.back}</button><div style="flex:1"><h1>Ajustes</h1></div></div>
    <div class="section">
      <div class="section-title"><h2>Cuentas</h2><button class="link" id="btnNuevaCuenta">＋ Agregar</button></div>
      <div class="card" style="padding:4px 14px">
        ${datos.cuentas.map((c) => `<button class="row" data-ecuenta="${c.id}" style="${c.archivada ? 'opacity:.5' : ''}">
          <span class="chip-ico" style="--c:${c.color}">${c.tipo === 'banco' ? '🏦' : c.tipo === 'billetera' ? '📲' : c.tipo === 'tarjeta' ? '💳' : '💵'}</span>
          <span class="row-main"><span class="row-title">${esc(c.nombre)}</span><span class="row-sub">${c.archivada ? 'Archivada' : { banco: 'Banco', billetera: 'Billetera digital', tarjeta: 'Tarjeta', efectivo: 'Efectivo' }[c.tipo] || ''}</span></span>
          <span class="row-amt">${F.dinero(F.saldoCuenta(c.id))}</span></button>`).join('')}
      </div>

      <div class="section-title"><h2>Categorías</h2><button class="link" id="btnNuevaCat">＋ Agregar</button></div>
      <div class="segment" style="margin:0 0 10px">
        <button data-ct="egreso" class="${st.catTab === 'egreso' ? 'on' : ''}">En qué gasto</button>
        <button data-ct="ingreso" class="${st.catTab === 'ingreso' ? 'on' : ''}">De dónde entra</button>
      </div>
      <div class="card" style="padding:4px 14px">
        ${cats.map((c) => `<button class="row" data-ecat="${c.id}"><span class="chip-ico" style="--c:${c.color}">${esc(c.icono)}</span>
          <span class="row-main"><span class="row-title">${esc(c.nombre)}</span><span class="row-sub">${c.tipo === 'egreso' && c.presupuesto ? 'Presupuesto ' + F.dinero(c.presupuesto) + ' al mes' : usoCategoria(c.id) + ' movimientos'}</span></span></button>`).join('')}
      </div>

      <div class="section-title"><h2>Exportar y respaldo</h2></div>
      <div class="card" style="display:flex;flex-direction:column;gap:10px">
        <button class="btn soft block" id="btnExcel">Exportar todo a Excel</button>
        <button class="btn ghost block" id="btnRespaldo">Guardar copia de seguridad</button>
        <button class="btn ghost block" id="btnRestaurar">Restaurar copia de seguridad</button>
        <p class="hint" style="margin:0">La copia guarda todo en un archivo. Úsala para pasar tus cuentas a otro teléfono o por si cambias de celular.</p>
      </div>

      <div class="section-title"><h2>Seguridad</h2></div>
      <div class="card" style="display:flex;flex-direction:column;gap:10px">
        <button class="btn ${pin ? 'ghost' : 'soft'} block" id="btnPin">${pin ? 'Cambiar PIN' : 'Proteger con PIN'}</button>
        ${pin ? '<button class="btn ghost block" id="btnQuitarPin" style="color:var(--down)">Quitar PIN</button>' : ''}
        <p class="hint" style="margin:0">${pin ? 'La app pide el PIN al abrirse y al volver después de un minuto.' : 'Pide un PIN de 4 dígitos al abrir la app.'}</p>
      </div>

      <p class="hint" style="text-align:center;margin:26px 0 8px">Cuentas DECOM · tus datos se guardan solo en este teléfono, sin internet.</p>
    </div>`;

  $('#btnVolver', v).addEventListener('click', () => irA('inicio'));
  $('#btnNuevaCuenta', v).addEventListener('click', () => abrirPanel((s) => formCuenta(s)));
  $$('[data-ecuenta]', v).forEach((b) => b.addEventListener('click', () => abrirPanel((s) => formCuenta(s, F.cuenta(b.dataset.ecuenta)))));
  $('#btnNuevaCat', v).addEventListener('click', () => abrirPanel((s) => formCategoria(s, null, st.catTab)));
  $$('[data-ecat]', v).forEach((b) => b.addEventListener('click', () => abrirPanel((s) => formCategoria(s, F.categoria(b.dataset.ecat)))));
  $$('[data-ct]', v).forEach((b) => b.addEventListener('click', () => { st.catTab = b.dataset.ct; render(); }));
  $('#btnExcel', v).addEventListener('click', () => exportar());
  $('#btnRespaldo', v).addEventListener('click', async () => {
    try { await guardarRespaldo(datos); } catch (e) { toast('No se pudo guardar: ' + e.message); }
  });
  $('#btnRestaurar', v).addEventListener('click', () => $('#fileInput').click());
  $('#btnPin', v).addEventListener('click', () => pedirPin('crear'));
  $('#btnQuitarPin', v)?.addEventListener('click', async () => {
    if (!(await confirmar('¿Quitar el PIN? Cualquiera con tu teléfono podrá ver tus cuentas.', 'Quitar PIN'))) return;
    await db.fijarAjuste('pinHash', null);
    render();
    toast('PIN quitado');
  });
}

$('#fileInput').addEventListener('change', async (e) => {
  const f = e.target.files[0];
  e.target.value = '';
  if (!f) return;
  try {
    const nuevos = await leerRespaldo(f);
    const n = nuevos.movimientos.length;
    if (!(await confirmar(`¿Reemplazar todos los datos de este teléfono por la copia? Trae ${n} movimientos, ${nuevos.deudas.length} deudas y ${nuevos.metas.length} metas.`, 'Restaurar', false))) return;
    const pinActual = db.ajuste('pinHash');
    await db.reemplazarTodo(nuevos);
    if (pinActual && !db.ajuste('pinHash')) await db.fijarAjuste('pinHash', pinActual);
    render();
    toast('Copia restaurada');
  } catch (err) {
    toast(err.message);
  }
});

function formCuenta(s, cta = null) {
  const c = cta ? { ...cta } : { nombre: '', tipo: 'banco', color: COLORES[1], saldoInicial: 0, archivada: false };
  const saldoActual = cta ? F.saldoCuenta(cta.id) : 0;
  s.innerHTML += `<h3>${cta ? 'Editar cuenta' : 'Nueva cuenta'}</h3>
    <div class="field"><label>Nombre</label><input class="input" id="ctNombre" maxlength="40" value="${esc(c.nombre)}" placeholder="Ej: Bancolombia, Daviplata, Alcancía"></div>
    <div class="field"><label>Tipo</label><select class="input" id="ctTipo">
      ${[['efectivo', 'Efectivo'], ['banco', 'Banco'], ['billetera', 'Billetera digital (Nequi, Daviplata…)'], ['tarjeta', 'Tarjeta de crédito']].map(([k, l]) => `<option value="${k}" ${c.tipo === k ? 'selected' : ''}>${l}</option>`).join('')}</select></div>
    <div class="field"><label>${cta ? 'Saldo actual' : '¿Cuánto tienes hoy en esta cuenta?'}</label>${campoMonto('ctSaldo', Math.abs(cta ? saldoActual : c.saldoInicial))}
      <label class="check" style="margin-top:8px"><input type="checkbox" id="ctNegativo" ${(cta ? saldoActual : c.saldoInicial) < 0 ? 'checked' : ''}> Es un saldo negativo (debes en esta cuenta)</label>
      ${cta ? '<p class="hint">Si lo cambias, se ajusta el saldo de partida. Tus movimientos no se tocan.</p>' : ''}</div>
    <div class="field"><label>Color</label><div class="swatches">${COLORES.slice(0, 10).map((x) => `<button data-color="${x}" style="--c:${x}" class="${c.color === x ? 'on' : ''}" aria-label="Color"></button>`).join('')}</div></div>
    ${cta ? `<label class="check"><input type="checkbox" id="ctArchivada" ${c.archivada ? 'checked' : ''}> Archivar (ocultar y no sumarla al saldo)</label>` : ''}
    <div class="sheet-actions">
      ${cta ? '<button class="btn ghost" id="ctBorrar" style="flex:0 0 auto;color:var(--down)">Eliminar</button>' : ''}
      <button class="btn primary" id="ctGuardar">Guardar</button></div>`;
  activarMontos(s);
  $$('[data-color]', s).forEach((b) => b.addEventListener('click', () => { c.color = b.dataset.color; $$('[data-color]', s).forEach((x) => x.classList.toggle('on', x === b)); }));
  $('#ctGuardar', s).addEventListener('click', async () => {
    c.nombre = $('#ctNombre', s).value.trim();
    c.tipo = $('#ctTipo', s).value;
    if (!c.nombre) return toast('Escribe el nombre');
    let saldo = parseMonto($('#ctSaldo', s).value);
    if ($('#ctNegativo', s).checked) saldo = -saldo;
    // el saldo de partida es el que hace que el saldo actual quede igual al escrito
    c.saldoInicial = cta ? (Number(cta.saldoInicial) || 0) + (saldo - saldoActual) : saldo;
    if (cta) c.archivada = $('#ctArchivada', s).checked;
    await db.guardar('cuentas', c);
    listo(cta ? 'Cuenta actualizada' : 'Cuenta creada');
  });
  $('#ctBorrar', s)?.addEventListener('click', async () => {
    const usos = datos.movimientos.filter((m) => m.cuentaId === cta.id || m.cuentaDestinoId === cta.id).length;
    if (usos) return toast(`Tiene ${usos} movimientos. Mejor archívala para no perder el historial.`);
    if (!(await confirmar(`¿Eliminar la cuenta "${cta.nombre}"?`))) return;
    await db.borrar('cuentas', cta.id);
    listo('Cuenta eliminada');
  });
}

function formCategoria(s, cat = null, tipo = 'egreso', alGuardar = null) {
  const c = cat ? { ...cat } : { nombre: '', tipo, icono: tipo === 'ingreso' ? '💵' : '📦', color: COLORES[2], presupuesto: 0 };
  s.innerHTML += `<h3>${cat ? 'Editar categoría' : c.tipo === 'ingreso' ? 'Nueva fuente de ingreso' : 'Nueva categoría de gasto'}</h3>
    <div class="field"><label>Nombre</label><input class="input" id="caNombre" maxlength="30" value="${esc(c.nombre)}" placeholder="${c.tipo === 'ingreso' ? 'Ej: Arriendo que recibo' : 'Ej: Gimnasio'}"></div>
    <div class="field"><label>Ícono</label><div class="emojis">${EMOJIS.map((e) => `<button data-emoji="${e}" class="${c.icono === e ? 'on' : ''}">${e}</button>`).join('')}</div></div>
    <div class="field"><label>Color</label><div class="swatches">${COLORES.slice(0, 11).map((x) => `<button data-color="${x}" style="--c:${x}" class="${c.color === x ? 'on' : ''}" aria-label="Color"></button>`).join('')}</div></div>
    ${c.tipo === 'egreso' ? `<div class="field"><label>Presupuesto mensual (opcional)</label>${campoMonto('caPresu', c.presupuesto)}
      <p class="hint">Te muestra en Inicio cuánto llevas gastado de este presupuesto cada mes.</p></div>` : ''}
    <div class="sheet-actions">
      ${cat ? '<button class="btn ghost" id="caBorrar" style="flex:0 0 auto;color:var(--down)">Eliminar</button>' : ''}
      ${cat ? '<button class="btn ghost" id="caVer">Ver movimientos</button>' : ''}
      <button class="btn primary" id="caGuardar">Guardar</button></div>`;
  activarMontos(s);
  $$('[data-emoji]', s).forEach((b) => b.addEventListener('click', () => { c.icono = b.dataset.emoji; $$('[data-emoji]', s).forEach((x) => x.classList.toggle('on', x === b)); }));
  $$('[data-color]', s).forEach((b) => b.addEventListener('click', () => { c.color = b.dataset.color; $$('[data-color]', s).forEach((x) => x.classList.toggle('on', x === b)); }));
  $('#caGuardar', s).addEventListener('click', async () => {
    c.nombre = $('#caNombre', s).value.trim();
    if (!c.nombre) return toast('Escribe el nombre');
    if (datos.categorias.some((x) => x.id !== c.id && x.tipo === c.tipo && x.nombre.toLowerCase() === c.nombre.toLowerCase())) return toast('Ya existe una categoría con ese nombre');
    if (c.tipo === 'egreso') c.presupuesto = parseMonto($('#caPresu', s).value);
    const guardada = await db.guardar('categorias', c);
    alGuardar && alGuardar(guardada);
    listo(cat ? 'Categoría actualizada' : 'Categoría creada');
  });
  $('#caVer', s)?.addEventListener('click', () => {
    st.filtroCategoria = cat.id;
    st.filtroCuenta = null;
    st.filtroTipo = 'todos';
    cerrarPaneles();
    irA('movimientos');
  });
  $('#caBorrar', s)?.addEventListener('click', async () => {
    const usos = usoCategoria(cat.id) + datos.cuotas.filter((x) => x.categoriaId === cat.id).length;
    if (usos) return toast(`La usan ${usos} registros; no se puede eliminar.`);
    if (!(await confirmar(`¿Eliminar la categoría "${cat.nombre}"?`))) return;
    await db.borrar('categorias', cat.id);
    listo('Categoría eliminada');
  });
}

async function exportar() {
  try {
    toast('Generando Excel…');
    await exportarExcel(datos, F);
  } catch (e) {
    console.error(e);
    toast('No se pudo exportar: ' + e.message);
  }
}

/* ======================= PIN ======================= */

const hashPin = async (pin) => {
  const buf = await crypto.subtle.digest('SHA-256', new TextEncoder().encode('cuentas-decom:' + pin));
  return [...new Uint8Array(buf)].map((b) => b.toString(16).padStart(2, '0')).join('');
};

let pinModo = null, pinTexto = '', pinPrimero = '', pinResolver = null;
function pedirPin(modo) {
  pinModo = modo;
  pinTexto = '';
  pinPrimero = '';
  $('#lockTitle').textContent = modo === 'crear' ? 'Elige un PIN de 4 dígitos' : 'Ingresa tu PIN';
  $('#lockError').textContent = '';
  $('#lockCancel').style.visibility = modo === 'crear' ? 'visible' : 'hidden';
  pintarPuntos();
  $('#lock').hidden = false;
  return new Promise((r) => { pinResolver = r; });
}
const pintarPuntos = () => $$('#pinDots i').forEach((d, i) => d.classList.toggle('on', i < pinTexto.length));
const errorPin = (msg) => {
  $('#lockError').textContent = msg;
  const dots = $('#pinDots');
  dots.classList.remove('shake');
  void dots.offsetWidth;
  dots.classList.add('shake');
  pinTexto = '';
  pintarPuntos();
};
const cerrarLock = (ok) => {
  $('#lock').hidden = true;
  pinResolver && pinResolver(ok);
  pinResolver = null;
};
$('#keypad').addEventListener('click', async (e) => {
  const b = e.target.closest('button');
  if (!b) return;
  const k = b.dataset.k;
  if (k === 'cancel') return cerrarLock(false);
  if (k === 'del') { pinTexto = pinTexto.slice(0, -1); return pintarPuntos(); }
  if (pinTexto.length >= 4) return;
  pinTexto += k;
  pintarPuntos();
  if (pinTexto.length < 4) return;
  await new Promise((r) => setTimeout(r, 120));
  if (pinModo === 'entrar') {
    if ((await hashPin(pinTexto)) === db.ajuste('pinHash')) cerrarLock(true);
    else errorPin('PIN incorrecto');
  } else if (!pinPrimero) {
    pinPrimero = pinTexto;
    pinTexto = '';
    $('#lockTitle').textContent = 'Repite el PIN';
    pintarPuntos();
  } else if (pinTexto === pinPrimero) {
    await db.fijarAjuste('pinHash', await hashPin(pinTexto));
    cerrarLock(true);
    render();
    toast('PIN activado');
  } else {
    pinPrimero = '';
    $('#lockTitle').textContent = 'Elige un PIN de 4 dígitos';
    errorPin('No coinciden, inténtalo de nuevo');
  }
});

let ocultaDesde = 0;
document.addEventListener('visibilitychange', () => {
  if (document.hidden) ocultaDesde = Date.now();
  else if (db.ajuste('pinHash') && ocultaDesde && Date.now() - ocultaDesde > 60000 && $('#lock').hidden) pedirPin('entrar');
});

/* ======================= navegación ======================= */

const VISTAS = { inicio: vistaInicio, movimientos: vistaMovimientos, deudas: vistaDeudas, metas: vistaMetas, ajustes: vistaAjustes };

function render() {
  const v = $('#view');
  VISTAS[st.vista](v);
  $$('.tabbar [data-tab]').forEach((b) => b.classList.toggle('on', b.dataset.tab === st.vista));
}

function irA(vista) {
  if (vista !== 'movimientos' && st.vista === 'movimientos') { st.filtroCuenta = null; st.filtroCategoria = null; st.busqueda = ''; }
  st.vista = vista;
  render();
  window.scrollTo(0, 0);
}

$$('.tabbar [data-tab]').forEach((b) => b.addEventListener('click', () => irA(b.dataset.tab)));
$('#fab').addEventListener('click', () => {
  const tipo = st.vista === 'movimientos' && st.filtroTipo === 'ingreso' ? 'ingreso' : 'egreso';
  abrirPanel((s, mem) => formMovimiento(s, null, { tipo }, mem));
});

// Botón "atrás" de Android
window.handleBack = () => {
  if (!$('#lock').hidden) {
    if (pinModo === 'crear') { cerrarLock(false); return true; }
    return false; // con la app bloqueada, "atrás" sale de la app
  }
  if (!$('#confirm').hidden) { cerrarConfirm(false); return true; }
  const listoPanel = $('#archivoListo');
  if (listoPanel && !listoPanel.hidden) { listoPanel.hidden = true; return true; }
  if (pila.length) { atrasPanel(); return true; }
  if (st.vista !== 'inicio') { irA('inicio'); return true; }
  return false;
};

/* ======================= arranque ======================= */

(async () => {
  try {
    await db.abrir();
    await F.sembrarSiVacio();
  } catch (e) {
    toast('No se pudieron leer los datos: ' + e.message);
  }
  render();
  setTimeout(() => {
    $('#splash').classList.add('gone');
    setTimeout(() => $('#splash').remove(), 500);
  }, 1100);
  if (db.ajuste('pinHash')) pedirPin('entrar');
})();
