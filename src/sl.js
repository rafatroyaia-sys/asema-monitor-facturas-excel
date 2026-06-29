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
    const mapaGI = esCompra ? gasto : ingreso;
    const nifSeed = esCompra ? (ent.nifProveedor || {}) : (ent.nifCliente || {}); // NIF→subcuenta del listado
    const k = nif && nif !== "0" ? nif : null;

    let sub = null;
    let nueva = false;  // emparejada por nombre (existe en el listado, falta confirmar el NIF)
    let creada = false; // NO está en el listado → subcuenta EN BLANCO + aviso (Monitor la crea al importar)

    if (k && mapaNif[k]) sub = mapaNif[k].sub;            // ya emparejado y revisado antes (localStorage)
    if (!sub && k && nifSeed[k]) sub = nifSeed[k];        // NIF está en el listado → confirmado, no es nuevo
    if (!sub) { const p = buscaPorNombre(mapaCP, nombre); if (p) { sub = p; nueva = true; } } // por nombre
    if (!sub) { sub = ""; creada = true; }                // nuevo: NO se crea subcuenta (Monitor da error si la inventa la app)

    const nombreSub = (sub && mapaCP[sub]) || nombre;
    if (k && sub) mapaNif[k] = { sub, nombre: nombreSub };

    // Contrapartida de gasto/ingreso: se BUSCA en el listado por nombre, NO por el sufijo
    // (la terminación no siempre coincide). Si no aparece, en blanco + aviso.
    const giCode = sub ? (buscaPorNombre(mapaGI, nombreSub) || "") : "";
    const giNombre = giCode ? mapaGI[giCode] : "";
    const giNueva = !!sub && !giCode; // existe el cliente/proveedor pero su contrapartida no está en el listado

    return { sub, nombreSub, nueva, creada, giCode, giNombre, giNueva, giDescuadre: false };
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
  if (!row.subCP && !row.subCPCreada) issues.push({ lv: "err", msg: "Sin subcuenta de cliente/proveedor" });
  if (!row.subGI && !row.subCPCreada && !row.subGINueva) issues.push({ lv: "err", msg: "Sin subcuenta de gasto/ingreso" });
  if (isFinite(tipo) && tipo !== 0 && !row.subIva) issues.push({ lv: "warn", msg: "Sin subcuenta de IVA para ese tipo" });
  if (!row.subBanco) issues.push({ lv: "warn", msg: "Sin subcuenta de banco (cobro/pago)" });
  if (row.subCPCreada) {
    const quien = row.sentido === "venta" ? "CLIENTE" : "PROVEEDOR";
    issues.push({ lv: "warn", msg: `⚠ ${quien} NUEVO: no está en el listado. Deja la subcuenta y su contrapartida EN BLANCO — Monitor las crea al importar. (Si la app pusiera una subcuenta inventada, daría error.)` });
  } else if (row.subCPNueva) {
    issues.push({ lv: "warn", msg: `Subcuenta ${row.subCP} emparejada por nombre: confirma que es la correcta` });
  }
  if (row.subGINueva && !row.subCPCreada) issues.push({ lv: "warn", msg: "Contrapartida de gasto/ingreso no encontrada en el listado: déjala en blanco (Monitor) o complétala a mano" });
  // IVA cero: Monitor NO admite IVA 0 → error en ROJO (no se puede exportar así)
  if (isFinite(base) && base !== 0 && ((isFinite(tipo) && tipo === 0) || (isFinite(iva) && iva === 0))) {
    issues.push({ lv: "err", msg: "IVA 0: Monitor no admite IVA cero. Pon el tipo de IVA correcto antes de exportar." });
  }
  // IVA que no cuadra con la base → aviso ámbar
  if (isFinite(base) && isFinite(iva) && isFinite(tipo) && tipo !== 0) {
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
/* Cabeceras del Excel siguiendo las plantillas oficiales de Monitor. Se exporta
   COMPRAS y VENTAS POR SEPARADO (el despacho las contabiliza por separado), cada
   una con su cabecera. Misma estructura (27 columnas A-AA); el banco va en
   DEBE/HABER (pago/cobro) para que Monitor genere el asiento de banco. Las
   columnas opcionales que no automatizamos (domicilio, R.E., IRPF, rectificativa)
   van en blanco. Concepto = el que clasifica la app desde la factura. */
export const SL_HEADERS_COMPRAS = [
  "FECHA ASIENTO", "FECHA FACTURA", "Nº FACTURA", "CONCEPTO",
  "SUBCUENTA PROVEEDOR", "CIF/NIF", "NOMBRE", "DOMICILIO", "LOCALIDAD", "PROVINCIA", "C.P.",
  "BASE", "%IVA/IGIC", "CUOTA IVA/IGIC", "SUBCUENTA IVA/IGIC",
  "% R.E.", "IMPORTE R.E.", "SUBCUENTA R.E.", "% IRPF", "CUOTA IRPF", "SUBCUENTA IRPF",
  "RECTIFICATIVA", "SUBCUENTA GASTO", "IMPORTE GASTO/INGRESO", "DEBE", "HABER", "TOTAL FRA.",
];
export const SL_HEADERS_VENTAS = [
  "FECHA ASIENTO", "FECHA FACTURA", "Nº FACTURA", "CONCEPTO",
  "SUBCUENTA CLIENTE", "CIF/NIF CLIENTE", "NOMBRE", "DOMICILIO", "LOCALIDAD", "PROVINCIA", "C.P. CLIENTE",
  "BASE", "% IVA/IGIC", "CUOTA IVA/IGIC", "SUBCUENTA IVA/IGIC",
  "% RE", "CUOTA RE", "SUBCUENTA RE", "% RETENCIÓN", "IMPORTE RETEN", "SUBCUENTA RETENCION",
  "RECTIFICATIVA", "SUBCUENTA INGRESO", "BASE", "DEBE", "HABER", "TOTAL",
];

/* ---------- Banco (contrapartida de cobro/pago) por contabilidad ----------
   Todas las SL cobran/pagan por banco (subcuenta 572…). Se elige por entidad y
   se recuerda; por defecto, la 57200000000 que traen todos los listados. */
const claveBanco = (entId) => `asema_sl_banco_${entId}`;
export function cargarBanco(ent) {
  if (!ent) return "";
  try { return localStorage.getItem(claveBanco(ent.id)) || ent.banco || ""; } catch { return ent.banco || ""; }
}
export function guardarBanco(entId, code) {
  try { localStorage.setItem(claveBanco(entId), code); } catch { /* opcional */ }
}
