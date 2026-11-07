// Exportar a Excel (.xlsx armado a mano con JSZip) y copia de seguridad en JSON.
import { COLECCIONES } from './db.js';

const native = window.Android;
const FORMATO_RESPALDO = 'cuentas-decom';

/* ======================= XLSX mínimo ======================= */

const escXml = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]))
  // caracteres de control no permitidos en XML
  .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');

const colLetra = (i) => {
  let s = '';
  for (i += 1; i > 0; i = Math.floor((i - 1) / 26)) s = String.fromCharCode(65 + ((i - 1) % 26)) + s;
  return s;
};

// estilos: 0 normal · 1 encabezado · 2 dinero · 3 total (negrita dinero) · 4 título · 5 negrita · 6 porcentaje
const ESTILOS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="2"><numFmt numFmtId="164" formatCode="&quot;$&quot;\\ #,##0;[Red]\\-&quot;$&quot;\\ #,##0"/><numFmt numFmtId="165" formatCode="0%"/></numFmts>
<fonts count="4"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="15"/><color rgb="FF1B1838"/><name val="Calibri"/></font></fonts>
<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF4B3FA8"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="2"><border/><border><top style="thin"><color rgb="FF999999"/></top></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="7">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="164" fontId="2" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1"/>
<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="0" fontId="2" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1"/>
<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;

/** celda: valor simple, o {v, s} con estilo. Números se guardan como número. */
function hojaXml({ filas, anchos = [], congelar = 1 }) {
  const cols = anchos.length ? `<cols>${anchos.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('')}</cols>` : '';
  const vista = congelar ? `<sheetViews><sheetView workbookViewId="0"><pane ySplit="${congelar}" topLeftCell="A${congelar + 1}" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>` : '';
  const cuerpo = filas.map((fila, r) => `<row r="${r + 1}">${fila.map((celda, c) => {
    if (celda === null || celda === undefined || celda === '') return '';
    const { v, s = 0 } = typeof celda === 'object' ? celda : { v: celda };
    const ref = colLetra(c) + (r + 1);
    if (typeof v === 'number' && Number.isFinite(v)) return `<c r="${ref}" s="${s}"><v>${v}</v></c>`;
    return `<c r="${ref}" s="${s}" t="inlineStr"><is><t xml:space="preserve">${escXml(v)}</t></is></c>`;
  }).join('')}</row>`).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${vista}${cols}<sheetData>${cuerpo}</sheetData></worksheet>`;
}

async function armarXlsx(hojas) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
${hojas.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}</Types>`);
  zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  zip.file('xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>
${hojas.map((h, i) => `<sheet name="${escXml(h.nombre.slice(0, 31))}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets></workbook>`);
  zip.file('xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${hojas.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')}
<Relationship Id="rId${hojas.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`);
  zip.file('xl/styles.xml', ESTILOS);
  hojas.forEach((h, i) => zip.file(`xl/worksheets/sheet${i + 1}.xml`, hojaXml(h)));
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

/* ======================= contenido del Excel ======================= */

const H = (t) => ({ v: t, s: 1 });
const $ = (n) => ({ v: Math.round(Number(n) || 0), s: 2 });
const T = (n) => ({ v: Math.round(Number(n) || 0), s: 3 });
const B = (t) => ({ v: t, s: 5 });

const TIPOS = { ingreso: 'Ingreso', egreso: 'Gasto', transferencia: 'Transferencia', interno: 'Préstamo / ahorro' };

export async function exportarExcel(datos, F) {
  const movs = datos.movimientos.slice().sort((a, b) => b.fecha.localeCompare(a.fecha) || (b.creado || 0) - (a.creado || 0));
  const nombreCuenta = (id) => F.cuenta(id)?.nombre || '';

  // 1. Movimientos
  const hMovs = [[H('Fecha'), H('Tipo'), H('Categoría'), H('Cuenta'), H('Cuenta destino'), H('Nota'), H('Ingreso'), H('Gasto'), H('Efecto en saldo')]];
  for (const m of movs) {
    const cat = F.categoria(m.categoriaId);
    const efecto = m.tipo === 'ingreso' ? m.monto : m.tipo === 'egreso' ? -m.monto : m.tipo === 'interno' ? m.delta : 0;
    hMovs.push([m.fecha, TIPOS[m.tipo] || m.tipo, cat?.nombre || (m.origen?.tipo === 'meta' ? 'Meta de ahorro' : m.origen?.tipo === 'deuda' ? 'Deuda' : ''),
      nombreCuenta(m.cuentaId), nombreCuenta(m.cuentaDestinoId), m.nota || '',
      m.tipo === 'ingreso' ? $(m.monto) : '', m.tipo === 'egreso' ? $(m.monto) : '', m.tipo === 'transferencia' ? '' : $(efecto)]);
  }
  const ingTot = movs.reduce((t, m) => t + (m.tipo === 'ingreso' ? m.monto : 0), 0);
  const egrTot = movs.reduce((t, m) => t + (m.tipo === 'egreso' ? m.monto : 0), 0);
  hMovs.push(['', '', '', '', '', B('Total'), T(ingTot), T(egrTot), '']);

  // 2. Resumen mensual y por categoría
  const meses = [...new Set(movs.map((m) => F.mesDe(m.fecha)))].sort().reverse();
  const hRes = [[{ v: 'Resumen por mes', s: 4 }], [H('Mes'), H('Ingresos'), H('Gastos'), H('Balance')]];
  for (const ym of meses) {
    const r = F.resumenMes(ym);
    hRes.push([F.nombreMes(ym), $(r.ingresos), $(r.egresos), $(r.balance)]);
  }
  hRes.push([], [{ v: 'Gastos por categoría (todo el historial)', s: 4 }], [H('Categoría'), H('Total'), H('% del gasto'), H('Presupuesto mensual')]);
  const porCat = new Map();
  for (const m of movs) if (m.tipo === 'egreso') porCat.set(m.categoriaId, (porCat.get(m.categoriaId) || 0) + m.monto);
  [...porCat.entries()].sort((a, b) => b[1] - a[1]).forEach(([id, total]) => {
    const c = F.categoria(id);
    hRes.push([c?.nombre || 'Sin categoría', $(total), { v: egrTot ? total / egrTot : 0, s: 6 }, c?.presupuesto ? $(c.presupuesto) : '']);
  });
  hRes.push([], [{ v: 'Ingresos por fuente', s: 4 }], [H('Fuente'), H('Total')]);
  const porFuente = new Map();
  for (const m of movs) if (m.tipo === 'ingreso') porFuente.set(m.categoriaId, (porFuente.get(m.categoriaId) || 0) + m.monto);
  [...porFuente.entries()].sort((a, b) => b[1] - a[1]).forEach(([id, total]) => hRes.push([F.categoria(id)?.nombre || 'Sin categoría', $(total)]));

  // 3. Cuentas
  const hCtas = [[H('Cuenta'), H('Tipo'), H('Saldo inicial'), H('Saldo actual'), H('Estado')]];
  for (const c of datos.cuentas) hCtas.push([c.nombre, c.tipo, $(c.saldoInicial), $(F.saldoCuenta(c.id)), c.archivada ? 'Archivada' : 'Activa']);
  hCtas.push([B('Total disponible'), '', '', T(F.saldoTotal()), '']);

  // 4. Deudas
  const hDeu = [[H('Tipo'), H('Persona'), H('Concepto'), H('Fecha'), H('Vence'), H('Valor'), H('Abonado'), H('Pendiente'), H('Estado')]];
  for (const d of datos.deudas.slice().sort((a, b) => a.direccion.localeCompare(b.direccion) || a.fecha.localeCompare(b.fecha))) {
    hDeu.push([d.direccion === 'me_deben' ? 'Me debe' : 'Le debo', d.persona, d.concepto || '', d.fecha, d.vence || '',
      $(d.monto), $(F.abonado(d)), $(F.pendienteDeuda(d)), d.cerrada ? 'Pagada' : 'Pendiente']);
  }
  const pend = (dir) => datos.deudas.filter((d) => d.direccion === dir).reduce((t, d) => t + F.pendienteDeuda(d), 0);
  hDeu.push([], [B('Total que me deben'), '', '', '', '', '', '', T(pend('me_deben'))], [B('Total que debo'), '', '', '', '', '', '', T(pend('debo'))]);
  hDeu.push([], [{ v: 'Abonos', s: 4 }], [H('Persona'), H('Fecha'), H('Valor'), H('Nota')]);
  for (const d of datos.deudas) for (const a of d.abonos || []) hDeu.push([d.persona, a.fecha, $(a.monto), a.nota || '']);

  // 5. Cuotas
  const hCuo = [[H('Compra / crédito'), H('Valor cuota'), H('Cuotas'), H('Pagadas'), H('Frecuencia'), H('Próxima'), H('Total'), H('Pagado'), H('Saldo')]];
  for (const c of datos.cuotas) {
    const sig = F.siguienteCuota(c);
    hCuo.push([c.nombre, $(c.valorCuota), c.numCuotas, F.cuotasPagadas(c), c.frecuencia, sig ? sig.fecha : 'Completa',
      $(F.totalCuotas(c)), $(F.pagadoCuotas(c)), $(F.saldoCuotas(c))]);
  }
  hCuo.push([B('Total'), '', '', '', '', '', T(datos.cuotas.reduce((t, c) => t + F.totalCuotas(c), 0)),
    T(datos.cuotas.reduce((t, c) => t + F.pagadoCuotas(c), 0)), T(datos.cuotas.reduce((t, c) => t + F.saldoCuotas(c), 0))]);

  // 6. Metas
  const hMet = [[H('Meta'), H('Objetivo'), H('Ahorrado'), H('Falta'), H('% logrado'), H('Fecha límite'), H('Sugerido al mes')]];
  for (const m of datos.metas) {
    const ah = F.ahorrado(m), sug = F.ahorroSugerido(m);
    hMet.push([m.nombre, $(m.objetivo), $(ah), $(Math.max(0, m.objetivo - ah)), { v: m.objetivo ? Math.min(1, ah / m.objetivo) : 0, s: 6 }, m.fechaLimite || '', sug ? $(sug.mensual) : '']);
  }

  const bytes = await armarXlsx([
    { nombre: 'Movimientos', filas: hMovs, anchos: [12, 16, 20, 14, 14, 30, 14, 14, 15] },
    { nombre: 'Resumen', filas: hRes, anchos: [30, 16, 14, 18], congelar: 0 },
    { nombre: 'Cuentas', filas: hCtas, anchos: [22, 12, 15, 15, 12] },
    { nombre: 'Deudas', filas: hDeu, anchos: [18, 20, 26, 12, 12, 14, 14, 14, 11] },
    { nombre: 'Cuotas', filas: hCuo, anchos: [26, 14, 9, 9, 12, 12, 15, 15, 15] },
    { nombre: 'Metas', filas: hMet, anchos: [26, 15, 15, 15, 11, 13, 16] },
  ]);
  await entregar(bytes, `Cuentas DECOM ${F.hoy()}.xlsx`, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
}

/* ======================= copia de seguridad ======================= */

export async function guardarRespaldo(datos) {
  const copia = { formato: FORMATO_RESPALDO, version: 1, creado: new Date().toISOString() };
  for (const c of COLECCIONES) copia[c] = datos[c].filter((r) => c !== 'ajustes' || r.id !== 'pinHash');
  const bytes = new TextEncoder().encode(JSON.stringify(copia));
  const fecha = copia.creado.slice(0, 10);
  await entregar(bytes, `Respaldo Cuentas DECOM ${fecha}.json`, 'application/json');
}

export async function leerRespaldo(file) {
  let obj;
  try { obj = JSON.parse(await file.text()); } catch { throw new Error('El archivo no es una copia válida'); }
  if (!obj || obj.formato !== FORMATO_RESPALDO) throw new Error('Ese archivo no es una copia de Cuentas DECOM');
  const nuevos = {};
  for (const c of COLECCIONES) {
    const lista = Array.isArray(obj[c]) ? obj[c] : [];
    nuevos[c] = lista.filter((r) => r && typeof r === 'object' && typeof r.id === 'string');
  }
  return nuevos;
}

/* ======================= entregar el archivo ======================= */

const aBase64 = (bytes) => {
  let bin = '';
  for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  return btoa(bin);
};

function enviarANativo(bytes, nombre, mime, destino) {
  native.beginFile(nombre);
  const TROZO = 384 * 1024;
  for (let i = 0; i < bytes.length; i += TROZO) native.appendFile(aBase64(bytes.subarray(i, i + TROZO)));
  return native.endFile(mime, destino);
}

function panelArchivo(nombre, uri, mime, bytes) {
  let p = document.getElementById('archivoListo');
  if (!p) {
    p = document.createElement('div');
    p.id = 'archivoListo';
    p.className = 'confirm';
    p.innerHTML = `<div class="confirm-card"><p style="margin-bottom:4px"><b>Archivo guardado</b></p>
      <p class="hint" id="archivoNombre" style="margin:0 0 16px"></p>
      <div class="confirm-actions" style="flex-wrap:wrap">
        <button class="btn ghost" id="archivoCerrar">Cerrar</button>
        <button class="btn soft" id="archivoCompartir">Compartir</button>
        <button class="btn primary" id="archivoAbrir">Abrir</button></div></div>`;
    document.body.appendChild(p);
    p.addEventListener('click', (e) => { if (e.target === p) p.hidden = true; });
    p.querySelector('#archivoCerrar').addEventListener('click', () => (p.hidden = true));
  }
  p.querySelector('#archivoNombre').textContent = `${nombre} · en Descargas › Cuentas DECOM`;
  p.querySelector('#archivoAbrir').onclick = () => { p.hidden = true; native.openUri(uri, mime); };
  p.querySelector('#archivoCompartir').onclick = () => {
    p.hidden = true;
    if (enviarANativo(bytes, nombre, mime, 'compartir')) native.shareQueued(mime);
  };
  p.hidden = false;
}

async function entregar(bytes, nombre, mime) {
  if (native && native.beginFile) {
    const uri = enviarANativo(bytes, nombre, mime, 'dispositivo');
    if (!uri) throw new Error('No se pudo guardar el archivo en el teléfono');
    panelArchivo(nombre, uri, mime, bytes);
    return;
  }
  // en un navegador normal (pruebas): descarga directa
  const url = URL.createObjectURL(new Blob([bytes], { type: mime }));
  const a = Object.assign(document.createElement('a'), { href: url, download: nombre });
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 10000);
}
