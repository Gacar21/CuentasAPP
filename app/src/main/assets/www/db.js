// Almacenamiento en el teléfono (IndexedDB). Todo se carga en memoria al abrir la app:
// para finanzas personales son pocos miles de registros y así todo es instantáneo.

const DB_NAME = 'cuentas-decom';
const DB_VERSION = 1;
export const COLECCIONES = ['cuentas', 'categorias', 'movimientos', 'deudas', 'cuotas', 'metas', 'ajustes'];

let db = null;
export const datos = Object.fromEntries(COLECCIONES.map((c) => [c, []]));

const promesa = (req) => new Promise((resolve, reject) => {
  req.onsuccess = () => resolve(req.result);
  req.onerror = () => reject(req.error);
});

export async function abrir() {
  db = await new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = () => {
      for (const c of COLECCIONES) if (!req.result.objectStoreNames.contains(c)) req.result.createObjectStore(c, { keyPath: 'id' });
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
  const t = db.transaction(COLECCIONES, 'readonly');
  for (const c of COLECCIONES) datos[c] = await promesa(t.objectStore(c).getAll());
}

export const nuevoId = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 8);

/** Guarda (crea o reemplaza) un registro en memoria y en el teléfono. */
export async function guardar(coleccion, reg) {
  if (!reg.id) reg.id = nuevoId();
  const lista = datos[coleccion];
  const i = lista.findIndex((x) => x.id === reg.id);
  if (i >= 0) lista[i] = reg; else lista.push(reg);
  await new Promise((resolve, reject) => {
    const t = db.transaction(coleccion, 'readwrite');
    t.objectStore(coleccion).put(reg);
    t.oncomplete = resolve;
    t.onerror = () => reject(t.error);
  });
  return reg;
}

export async function borrar(coleccion, id) {
  datos[coleccion] = datos[coleccion].filter((x) => x.id !== id);
  await new Promise((resolve, reject) => {
    const t = db.transaction(coleccion, 'readwrite');
    t.objectStore(coleccion).delete(id);
    t.oncomplete = resolve;
    t.onerror = () => reject(t.error);
  });
}

/** Reemplaza todos los datos (al restaurar una copia de seguridad). */
export async function reemplazarTodo(nuevos) {
  await new Promise((resolve, reject) => {
    const t = db.transaction(COLECCIONES, 'readwrite');
    for (const c of COLECCIONES) {
      const s = t.objectStore(c);
      s.clear();
      for (const r of nuevos[c] || []) s.put(r);
    }
    t.oncomplete = resolve;
    t.onerror = () => reject(t.error);
  });
  for (const c of COLECCIONES) datos[c] = (nuevos[c] || []).slice();
}

export const ajuste = (clave, defecto = null) => datos.ajustes.find((a) => a.id === clave)?.valor ?? defecto;
export const fijarAjuste = (clave, valor) => guardar('ajustes', { id: clave, valor });
