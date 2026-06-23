/* ============================================================
   ASEMA · Módulo SL (doble partida) — datos y lógica pura
   Sin React. Se importa desde App.jsx para la pestaña "SL".

   Cada contabilidad trae su listado de subcuentas (extraído del
   "LISTADO DE CUENTAS" de Monitor). Las tres subcuentas del Conversor:
     - cliente/proveedor  → rango 430 / 400
     - IVA                → 477 (repercutido) / 472 (soportado), por tipo
     - gasto/ingreso      → 700 / 600, GEMELO de la de cliente/proveedor
                            (mismo sufijo: 400→600, 430→700)
   ============================================================ */

import bodegon1 from "./sl-seeds/bodegon1.json";
import bodegon2 from "./sl-seeds/bodegon2.json";
import matilde from "./sl-seeds/matilde.json";
import carpydekor from "./sl-seeds/carpydekor.json";
import bruzon from "./sl-seeds/bruzon.json";

/* Contabilidades disponibles (5 sociedades; El Bodegón tiene 2 centros).
   Falta por añadir la 5ª empresa nueva sin listado (numeración desde raíz). */
export const CONTABILIDADES_SL = [bodegon1, bodegon2, matilde, carpydekor, bruzon];

/* Formas societarias y ruido que se ignoran al emparejar por nombre */
const FORMAS = /\b(S\.?L\.?U?|S\.?A\.?U?|S\.?C\.?P?|S\.?L\.?L|S\.?L\.?N\.?E|C\.?B|S\.?COOP)\b/g;

/* Normaliza un nombre para comparar: mayúsculas, sin tildes, sin forma
   societaria ni puntuación. "Caracoles Gutiérrez, S.L." → "CARACOLES GUTIERREZ" */
export function normNombre(s) {
  return String(s || "")
    .toUpperCase()
    .normalize("NFD").replace(/[̀-ͯ]/g, "")
    .replace(/\./g, "")   // quita puntos SIN separar: "S.L." → "SL" (luego se elimina como forma)
    .replace(/,/g, " ")
    .replace(FORMAS, " ")
    .replace(/\s+/g, " ")
    .trim();
}

/* Busca una subcuenta en un rango (código→nombre) por nombre.
   Devuelve el código del mejor candidato o null. Match exacto preferente;
   si no, contención (uno contiene al otro). Puede ser ambiguo si hay
   nombres repetidos: por eso la fila se marca en ámbar para revisar. */
export function buscaPorNombre(mapa, nombre) {
  const t = normNombre(nombre);
  if (!t || t.length < 3) return null;
  let contiene = null;
  for (const code in mapa) {
    const n = normNombre(mapa[code]);
    if (!n) continue;
    if (n === t) return code;
    if (!contiene && n.length > 2 && (n.includes(t) || t.includes(n))) contiene = code;
  }
  return contiene;
}

/* ---------- Concepto del Conversor de Monitor ----------
   Monitor SOLO admite estos 4 conceptos en doble partida (confirmado con Monitor):
   COMPRAS, VENTAS, GASTOS y ENERGIA (sin tilde). Se derivan de la cuenta que
   asigna la IA (mismo plan que autónomos: 1=ventas, 2=compras, 5=energía, resto=gastos). */
export const CONCEPTOS_SL = ["COMPRAS", "VENTAS", "GASTOS", "ENERGIA"];
export function conceptoConversor(cuenta) {
  const c = Number(cuenta);
  if (c === 1) return "VENTAS";
  if (c === 2) return "COMPRAS";
  if (c === 5) return "ENERGIA";
  return "GASTOS"; // alquileres, primas, gastos diversos y cualquier otro gasto
}

/* Subcuenta de gasto/ingreso GEMELA de la de proveedor/cliente:
   400→600, 430→700, conservando el sufijo. El código es determinista. */
const PREFIJO_GEMELO = { 400: "600", 430: "700" };
export function codigoGemelo(sub) {
  const p = String(sub).slice(0, 3);
  const g = PREFIJO_GEMELO[p];
  return g ? g + String(sub).slice(3) : "";
}

/* Tipo de IVA → subcuenta, por sentido. tipo numérico (21, 10, 4, 7.5...). */
export function subIva(tipo, sentido, ent) {
  const t = Number(String(tipo).replace(",", "."));
  if (!isFinite(t) || t === 0) return ""; // exento / sin IVA: sin subcuenta
  const key = String(t);
  const mapa = sentido === "compra" ? ent.ivaCompraPorTipo : ent.ivaVentaPorTipo;
  return (mapa && mapa[key]) || "";
}

/* Siguiente código libre y correlativo dentro de una raíz (p.ej. "40000000").
   Solo mira los códigos de esa raíz densa, así no salta a subcuentas genéricas
   del PGC (40040000000 "moneda extranjera", 40090000000 "pendientes"...). */
export function siguienteCodigo(mapa, raiz) {
  let max = 0;
  for (const c in mapa) {
    if (c.startsWith(raiz)) { const n = Number(c); if (n > max) max = n; }
  }
  const base = max || Number(raiz.padEnd(11, "0"));
  return String(base + 1).padStart(11, "0");
}

/* ---------- Persistencia NIF→subcuenta por contabilidad ----------
   Así el mismo NIF recibe SIEMPRE la misma subcuenta. */
const claveLS = (entId) => `asema_sl_sub_${entId}`;
export function cargarMapaNif(entId) {
  try { return JSON.parse(localStorage.getItem(claveLS(entId)) || "{}"); } catch { return {}; }
}
export function guardarMapaNif(entId, mapa) {
  try { localStorage.setItem(claveLS(entId), JSON.stringify(mapa)); } catch { /* opcional */ }
}

/* ---------- Asignador con memoria de sesión ----------
   Crea un asignador que recuerda los NIF ya vistos y va dando números
   correlativos a los nuevos (sin repetir dentro del mismo lote).
   asigna({nif, nombre, sentido}) → datos de las dos subcuentas. */
export function crearAsignador(ent, mapaNifInicial) {
  const mapaNif = { ...(mapaNifInicial || {}) };
  // Copias locales de los rangos para ir dando de alta los autonumerados
  const prov = { ...ent.proveedores };
  const cli = { ...ent.clientes };
  const gasto = { ...ent.gasto };
  const ingreso = { ...ent.ingreso };

  function asigna({ nif, nombre, sentido }) {
    const esCompra = sentido === "compra";
    const mapaCP = esCompra ? prov : cli;
    const raiz = esCompra ? (ent.raizProveedor || "400") : (ent.raizCliente || "430");
    const k = nif && nif !== "0" ? nif : null;

    let sub = null;
    let nueva = false;  // emparejada por nombre (existe en el listado, falta confirmar el NIF)
    let creada = false; // NUEVA de verdad: no estaba en ningún listado, se autonumera correlativa

    if (k && mapaNif[k]) {
      sub = mapaNif[k].sub; // ya emparejado y revisado antes
    }
    if (!sub) {
      const porNombre = buscaPorNombre(mapaCP, nombre);
      if (porNombre) { sub = porNombre; nueva = true; } // empareja por nombre: a revisar
    }
    if (!sub) {
      sub = siguienteCodigo(mapaCP, raiz);
      mapaCP[sub] = nombre; // alta provisional, correlativa al último
      creada = true;        // proveedor/cliente que la app ha tenido que crear
    }
    const nombreSub = mapaCP[sub] || nombre;
    if (k) mapaNif[k] = { sub, nombre: nombreSub };

    // Gemelo de gasto/ingreso (código determinista; nombre del listado real)
    const giCode = codigoGemelo(sub);
    const mapaGI = esCompra ? gasto : ingreso;
    let giNombre = mapaGI[giCode] || "";
    let giNueva = false;
    if (giCode && !giNombre) { mapaGI[giCode] = nombre; giNombre = nombre; giNueva = true; }
    // Descuadre: el gemelo existe pero su nombre no se parece a la contraparte
    const giDescuadre = !!giNombre && normNombre(giNombre) !== normNombre(nombreSub);

    return { sub, nombreSub, nueva, creada, giCode, giNombre, giNueva, giDescuadre };
  }

  return { asigna, dump: () => mapaNif };
}

/* ---------- Validación de fila SL (cuadre del asiento) ---------- */
const _n = (v) => {
  if (typeof v === "number") return v;
  let s = String(v ?? "").trim();
  if (s === "") return NaN;
  if (s.includes(",")) s = s.replace(/\./g, "").replace(",", ".");
  const x = parseFloat(s);
  return isFinite(x) ? x : NaN;
};
const _r2 = (n) => Math.round((n + Number.EPSILON) * 100) / 100;

/* ¿La fecha dd/mm/aaaa cae en el trimestre (1-4) + año seleccionados? */
function _enTrimestre(fechaStr, tri, anio) {
  const m = String(fechaStr || "").match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (!m) return true; // sin fecha válida no aplicamos este aviso (ya hay otro)
  return Number(m[3]) === Number(anio) && Math.floor((Number(m[2]) - 1) / 3) + 1 === Number(tri);
}
function _rangoTri(tri, anio) {
  const ini = (tri - 1) * 3 + 1, fin = tri * 3;
  const d = new Date(anio, fin, 0).getDate();
  return `01/${String(ini).padStart(2, "0")} a ${String(d).padStart(2, "0")}/${String(fin).padStart(2, "0")}`;
}

export function validarFilaSL(row, tri, anio) {
  const issues = [];
  const base = _n(row.base), iva = _n(row.cuotaIva), tipo = _n(row.tipoIva), total = _n(row.total);
  if (!row.contraparte) issues.push({ lv: "err", msg: "Falta el nombre de cliente/proveedor" });
  if (!row.numero) issues.push({ lv: "warn", msg: "Falta el número de factura" });
  if (!isFinite(base)) issues.push({ lv: "err", msg: "Base imponible no numérica" });
  if (!isFinite(iva)) issues.push({ lv: "err", msg: "Importe de IVA no numérico" });
  if (!row.subCP) issues.push({ lv: "err", msg: "Sin subcuenta de cliente/proveedor" });
  if (!row.subGI) issues.push({ lv: "err", msg: "Sin subcuenta de gasto/ingreso" });
  if (isFinite(tipo) && tipo !== 0 && !row.subIva) issues.push({ lv: "warn", msg: "Sin subcuenta de IVA para ese tipo" });
  if (row.subCPCreada) {
    const tipo = row.sentido === "venta" ? "CLIENTE" : "PROVEEDOR";
    issues.push({ lv: "warn", msg: `⚠ ${tipo} NUEVO: no estaba en el listado. Se han creado las subcuentas ${row.subCP} (y gasto/ingreso ${row.subGI}) correlativas a la última. Revísalas antes de exportar.` });
  } else if (row.subCPNueva) {
    issues.push({ lv: "warn", msg: `Subcuenta ${row.subCP} emparejada por nombre: confirma que es la correcta` });
  }
  if (row.subGIDescuadre) issues.push({ lv: "warn", msg: `La cuenta de gasto/ingreso (${row.subGINombre}) no coincide con la contraparte` });
  if (isFinite(base) && isFinite(iva) && isFinite(tipo)) {
    const esp = _r2((base * tipo) / 100);
    if (Math.abs(esp - iva) > 0.02) issues.push({ lv: "warn", msg: `IVA no cuadra: ${base} × ${row.tipoIva}% = ${esp}` });
  }
  if (isFinite(base) && isFinite(iva) && isFinite(total)) {
    const re = _n(row.cuotaRe) || 0, ret = _n(row.cuotaRet) || 0;
    const esp = _r2(base + iva + re - ret);
    if (Math.abs(esp - total) > 0.02) issues.push({ lv: "warn", msg: `Total no cuadra: base + IVA + RE − ret = ${esp}` });
  }
  if (tri && anio && /^\d{2}\/\d{2}\/\d{4}$/.test(row.fechaFactura) && !_enTrimestre(row.fechaFactura, tri, anio)) {
    issues.push({ lv: "warn", msg: `Fecha fuera del trimestre seleccionado (${tri}T ${anio}: ${_rangoTri(tri, anio)})` });
  }
  return issues;
}

/* Cabeceras del Excel que importa el Conversor de Monitor.
   Orden alineado con la plantilla del Conversor del despacho: A=Nombre, B=NIF,
   C=Nº, D=Fecha (asiento y factura), E=Base, F=Importe IVA, H=%IVA, J=Concepto
   ya quedan donde la plantilla los espera; solo hay que mapear las 3 subcuentas
   nuevas (I=cliente/prov., K=IVA, L=gasto/ingreso). */
export const SL_HEADERS = [
  "Nombre", "NIF", "Nº factura", "Fecha",
  "Base imponible", "Importe IVA", "Total", "% IVA",
  "Subcuenta cliente/proveedor", "Concepto", "Subcuenta de IVA", "Subcuenta de gasto/ingreso",
  "Subcuenta recargo", "Subcuenta retención",
];
