import React, { useState, useEffect, useRef, useMemo } from "react";
import * as XLSX from "xlsx";
import mammoth from "mammoth";
import { PDFDocument } from "pdf-lib";
import {
  Upload, FileText, Trash2, Download, Calculator, AlertTriangle,
  CheckCircle2, XCircle, Loader2, Building2, Save, Plus, FileSpreadsheet,
  Lock, LogOut, CalendarCheck,
} from "lucide-react";
import {
  CONTABILIDADES_SL, crearAsignador, cargarMapaNif, guardarMapaNif,
  subIva, validarFilaSL, SL_HEADERS_COMPRAS, SL_HEADERS_VENTAS, conceptoConversor, CONCEPTOS_SL,
  cargarBanco, guardarBanco,
} from "./sl";

/* ============================================================
   ASEMA ADVISORY · Facturas → Excel Monitor (EOS Autónomos)
   Versión de despliegue (Vercel) · v1.0
   - La extracción IA se hace en el backend /api/extraer
   - Acceso protegido con la clave del despacho (ASEMA_PASSWORD)
   - Clientes guardados en el navegador de cada PC (localStorage)
   ============================================================ */

/* ---------- Paleta de marca (logo ASEMA) ---------- */
const C = {
  vino: "#7A1230",
  vinoOscuro: "#570C22",
  vinoSuave: "#9A2C4A",
  crema: "#F7F1E2",
  cremaOscura: "#EDE2C9",
  linea: "#E0D5BC",
  tinta: "#2B141C",
  papel: "#FFFFFF",
  ok: "#2F7D4F",
  okBg: "#EAF4EE",
  warn: "#A8731B",
  warnBg: "#FBF3E2",
  err: "#B3261E",
  errBg: "#FBEAE8",
  gris: "#7A6E5C",
};

const MONO = "ui-monospace, 'Cascadia Mono', 'Segoe UI Mono', Menlo, monospace";
const MAX_MB = 4; // límite de Vercel por petición (~4,5 MB); escanea a 150-200 ppp

/* Elimina caracteres invisibles (BOM, espacios de ancho cero, control)
   que se cuelan al copiar y pegar desde Word, PDF o notas. */
const limpiaClave = (s) =>
  String(s || "")
    .replace(/[\u0000-\u001F\u007F\u200B-\u200D\u2060\uFEFF]/g, "")
    .trim();

/* ---------- Cuentas de Monitor (módulo autónomos / EOS) ---------- */
const CUENTAS = [
  { n: 1, t: "VENTAS" },
  { n: 2, t: "COMPRAS" },
  { n: 3, t: "PERSONAL" },
  { n: 4, t: "SEG. SOCIAL" },
  { n: 5, t: "ENERGÍA" },
  { n: 6, t: "ALQUILERES" },
  { n: 7, t: "INTERESES" },
  { n: 8, t: "PRIMAS SEG." },
  { n: 9, t: "TRIB. NO EST." },
  { n: 10, t: "GASTOS DIV." },
  { n: 11, t: "AMORTIZAC." },
  { n: 14, t: "PAGOS" },
];

/* ---------- Letra de cuenta que valida Monitor en la importación ----------
   Monitor NO admite el número de cuenta en el Excel, sino una letra:
   V=Ventas, C=Compras, E=Energía, A=Alquileres, G=Gastos (resto).
   Internamente seguimos usando los números (más cómodos para revisar y
   para el resumen), y solo traducimos a letra al exportar. */
const cuentaLetra = (n) => {
  const c = Number(n);
  if (c === 1) return "V";   // VENTAS (emitidas)
  if (c === 2) return "C";   // COMPRAS
  if (c === 5) return "E";   // ENERGÍA
  if (c === 6) return "A";   // ALQUILERES
  return "G";                // GASTOS (10 gastos div., 8 primas seg., y resto)
};

/* ---------- Concepto del apunte (campo "Concepto" de Monitor) ----------
   Monitor SOLO admite 4 conceptos: COMPRAS, VENTAS, GASTOS y ENERGIA (sin
   tilde). Cuenta 1 → VENTAS, 2 → COMPRAS, 5 → ENERGIA, y cualquier otro gasto
   (personal, seg. social, alquileres, primas, tributos, gastos diversos...) →
   GASTOS. Mismo criterio que la pestaña SL (reutiliza conceptoConversor).
   El detalle fino se conserva en la columna CUENTA (letra) y en la CLAVE. */
const conceptoPorCuenta = (n) => conceptoConversor(n);

/* ---------- Claves de gastos diversos (cuenta 10) ---------- */
const CLAVES = [
  { k: "", t: "— sin clave —" },
  { k: "R", t: "R · Reparaciones y conservación" },
  { k: "P", t: "P · Profesionales independientes" },
  { k: "E", t: "E · Otros servicios exteriores" },
  { k: "A", t: "A · Suministro de agua" },
  { k: "G", t: "G · Suministro de gas" },
  { k: "I", t: "I · Telefonía e internet" },
  { k: "S", t: "S · Otros suministros" },
  { k: "O", t: "O · Otros fiscalmente deducibles" },
];

/* ---------- Helpers numéricos (formato español) ---------- */
const num = (v) => {
  if (typeof v === "number") return isFinite(v) ? v : NaN;
  if (v === null || v === undefined) return NaN;
  let s = String(v).trim();
  if (s === "") return NaN;
  if (s.includes(",")) s = s.replace(/\./g, "").replace(",", ".");
  const n = parseFloat(s);
  return isFinite(n) ? n : NaN;
};
const r2 = (n) => Math.round((n + Number.EPSILON) * 100) / 100;
const fmtES = (n) =>
  isFinite(n)
    ? n.toLocaleString("es-ES", { minimumFractionDigits: 2, maximumFractionDigits: 2 })
    : "—";
const toInput = (n) => {
  const x = num(n);
  return isFinite(x) ? x.toFixed(2).replace(".", ",") : "";
};
const toInputPct = (n) => {
  const x = num(n);
  if (!isFinite(x)) return "0";
  return String(x).replace(".", ",");
};

/* ---------- Validación NIF/CIF/NIE ---------- */
const validNif = (nifRaw) => {
  if (!nifRaw) return false;
  const nif = String(nifRaw).toUpperCase().replace(/[\s\-\.]/g, "");
  if (nif === "0") return true; // particulares sin NIF: criterio del despacho
  const letras = "TRWAGMYFPDXBNJZSQVHLCKE";
  if (/^\d{8}[A-Z]$/.test(nif)) return nif[8] === letras[parseInt(nif.slice(0, 8), 10) % 23];
  if (/^[XYZ]\d{7}[A-Z]$/.test(nif)) {
    const map = { X: "0", Y: "1", Z: "2" };
    const n = parseInt(map[nif[0]] + nif.slice(1, 8), 10);
    return nif[8] === letras[n % 23];
  }
  if (/^[ABCDEFGHJNPQRSUVW]\d{7}[0-9A-J]$/.test(nif)) {
    const digits = nif.slice(1, 8);
    let suma = 0;
    for (let i = 0; i < 7; i++) {
      let d = parseInt(digits[i], 10);
      if (i % 2 === 0) {
        d *= 2;
        if (d > 9) d = Math.floor(d / 10) + (d % 10);
      }
      suma += d;
    }
    const control = (10 - (suma % 10)) % 10;
    const letraCtrl = "JABCDEFGHI"[control];
    return nif[8] === String(control) || nif[8] === letraCtrl;
  }
  return false;
};

/* ---------- Fechas ---------- */
const normFecha = (s) => {
  if (!s) return "";
  const str = String(s).trim();
  let m = str.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2,4})$/);
  if (m) {
    const yyyy = m[3].length === 2 ? "20" + m[3] : m[3];
    return `${m[1].padStart(2, "0")}/${m[2].padStart(2, "0")}/${yyyy}`;
  }
  m = str.match(/^(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})$/);
  if (m) return `${m[3].padStart(2, "0")}/${m[2].padStart(2, "0")}/${m[1]}`;
  return str;
};
const fechaValida = (s) => {
  const m = String(s || "").match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (!m) return false;
  const d = parseInt(m[1], 10), mo = parseInt(m[2], 10), y = parseInt(m[3], 10);
  if (mo < 1 || mo > 12 || d < 1) return false;
  const dias = new Date(y, mo, 0).getDate();
  return d <= dias && y >= 2000 && y <= 2100;
};
const fechaKey = (s) => {
  const m = String(s || "").match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  return m ? `${m[3]}${m[2]}${m[1]}` : "99999999";
};

let SEQ = 1;
const uid = () => `r${Date.now().toString(36)}${(SEQ++).toString(36)}`;

/* ---------- Lectura de fichero a base64 ---------- */
const toB64 = (file) =>
  new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(String(r.result).split(",")[1]);
    r.onerror = () => rej(new Error("No se pudo leer el archivo"));
    r.readAsDataURL(file);
  });

/* ---------- Llamada al backend con un bloque (base64) ---------- */
async function pedirBloque(b64, media, empresa, claveDespacho) {
  const resp = await fetch("/api/extraer", {
    method: "POST",
    headers: { "Content-Type": "application/json", "x-asema-key": limpiaClave(claveDespacho) },
    body: JSON.stringify({ media_type: media, data: b64, empresa }),
  });
  let data = {};
  try { data = await resp.json(); } catch { /* respuesta sin cuerpo */ }
  if (resp.status === 401) { const e = new Error(data.error || "Clave del despacho incorrecta"); e.auth = true; throw e; }
  if (!resp.ok) throw new Error(data.error || `Error del servidor (HTTP ${resp.status})`);
  if (!Array.isArray(data.facturas)) throw new Error("Respuesta inesperada del servidor.");
  return data.facturas;
}

/* Convierte un ArrayBuffer (PDF) a base64 sin reventar la pila con archivos grandes */
function abToB64(bytes) {
  let bin = "";
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) {
    bin += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
  }
  return btoa(bin);
}

const PAGINAS_POR_BLOQUE = 8; // tamaño de cada lote al trocear PDFs grandes

/* ---------- Extracción de facturas (trocea PDFs grandes por páginas) ----------
   - Imagen o PDF de pocas páginas → una sola llamada.
   - PDF con muchas páginas → se parte en bloques de PAGINAS_POR_BLOQUE,
     se procesa cada bloque y se juntan los resultados. Sin límite práctico. */
async function extraerFacturas(file, empresa, claveDespacho, onProgreso) {
  const esPdf = file.type === "application/pdf" || /\.pdf$/i.test(file.name);

  if (!esPdf) {
    const b64 = await toB64(file);
    return pedirBloque(b64, file.type || "image/jpeg", empresa, claveDespacho);
  }

  // Cargar el PDF para contar páginas
  const buf = await file.arrayBuffer();
  let doc;
  try {
    doc = await PDFDocument.load(buf, { ignoreEncryption: true });
  } catch {
    // Si no se puede partir, intentar la vía clásica de una sola llamada
    const b64 = await toB64(file);
    return pedirBloque(b64, "application/pdf", empresa, claveDespacho);
  }
  const nPag = doc.getPageCount();

  // PDF pequeño: una sola llamada
  if (nPag <= PAGINAS_POR_BLOQUE) {
    const b64 = await toB64(file);
    return pedirBloque(b64, "application/pdf", empresa, claveDespacho);
  }

  // PDF grande: trocear en bloques de páginas
  const todas = [];
  const nBloques = Math.ceil(nPag / PAGINAS_POR_BLOQUE);
  for (let b = 0; b < nBloques; b++) {
    if (onProgreso) onProgreso(b + 1, nBloques);
    const sub = await PDFDocument.create();
    const ini = b * PAGINAS_POR_BLOQUE;
    const fin = Math.min(ini + PAGINAS_POR_BLOQUE, nPag);
    const indices = [];
    for (let p = ini; p < fin; p++) indices.push(p);
    const paginas = await sub.copyPages(doc, indices);
    paginas.forEach((pg) => sub.addPage(pg));
    const bytes = await sub.save();
    const b64 = abToB64(bytes);
    const facturas = await pedirBloque(b64, "application/pdf", empresa, claveDespacho);
    todas.push(...facturas);
  }
  return todas;
}

/* ---------- Factura en Word (.docx/.doc) → texto → IA ---------- */
async function extraerFacturaWord(file, empresa, claveDespacho) {
  const esDoc = /\.doc$/i.test(file.name);
  let texto = "";
  try {
    const arrayBuffer = await file.arrayBuffer();
    const result = await mammoth.extractRawText({ arrayBuffer });
    texto = (result && result.value ? result.value : "").trim();
  } catch (e) {
    texto = "";
  }
  if (!texto || texto.length < 15) {
    if (esDoc) {
      throw new Error("No se pudo leer este .doc antiguo. Ábrelo en Word y usa «Guardar como → Word (.docx)», y vuelve a subirlo.");
    }
    throw new Error("El Word no contiene texto legible (¿es una factura escaneada pegada como imagen? En ese caso, súbela como PDF o foto).");
  }
  const resp = await fetch("/api/extraer", {
    method: "POST",
    headers: { "Content-Type": "application/json", "x-asema-key": limpiaClave(claveDespacho) },
    body: JSON.stringify({ textoDocumento: texto, media_type: "text/plain", data: "x", empresa }),
  });
  let data = {};
  try { data = await resp.json(); } catch { /* sin cuerpo */ }
  if (resp.status === 401) { const e = new Error(data.error || "Clave del despacho incorrecta"); e.auth = true; throw e; }
  if (!resp.ok) throw new Error(data.error || `Error del servidor (HTTP ${resp.status})`);
  if (!Array.isArray(data.facturas)) throw new Error("Respuesta inesperada del servidor.");
  return data.facturas;
}

/* ---------- Listado PDF (Aaron) por IA → filas de la tabla ---------- */
async function extraerListadoPDF(file, claveDespacho) {
  const b64 = await toB64(file);
  const resp = await fetch("/api/extraer", {
    method: "POST",
    headers: { "Content-Type": "application/json", "x-asema-key": limpiaClave(claveDespacho) },
    body: JSON.stringify({ media_type: "application/pdf", data: b64, promptOverride: promptListadoPDF }),
  });
  let data = {};
  try { data = await resp.json(); } catch { /* sin cuerpo */ }
  if (resp.status === 401) { const e = new Error(data.error || "Clave del despacho incorrecta"); e.auth = true; throw e; }
  if (!resp.ok) throw new Error(data.error || `Error del servidor (HTTP ${resp.status})`);
  if (!Array.isArray(data.facturas)) throw new Error("Respuesta inesperada del servidor.");
  return data.facturas;
}

/* ---------- Listado de INGRESOS (Matilde, doble partida) por IA ---------- */
async function extraerListadoIngresos(file, claveDespacho) {
  const b64 = await toB64(file);
  const resp = await fetch("/api/extraer", {
    method: "POST",
    headers: { "Content-Type": "application/json", "x-asema-key": limpiaClave(claveDespacho) },
    body: JSON.stringify({ media_type: "application/pdf", data: b64, promptOverride: promptListadoIngresos }),
  });
  let data = {};
  try { data = await resp.json(); } catch { /* sin cuerpo */ }
  if (resp.status === 401) { const e = new Error(data.error || "Clave del despacho incorrecta"); e.auth = true; throw e; }
  if (!resp.ok) throw new Error(data.error || `Error del servidor (HTTP ${resp.status})`);
  if (!Array.isArray(data.facturas)) throw new Error("Respuesta inesperada del servidor.");
  return data.facturas;
}

/* Convierte las líneas del listado PDF en filas de la tabla */
function listadoPDFARows(items) {
  return items.map((f) => {
    const cuenta = f.sentido === "compra" ? 2 : 1;
    const base = num(f.base);
    const iva = num(f.cuota_iva);
    const re = num(f.recargo) || 0;
    let total = num(f.total);
    if (!isFinite(total)) total = r2((isFinite(base) ? base : 0) + (isFinite(iva) ? iva : 0) + re);
    return filaListado({
      contraparte: String(f.contraparte || "").toUpperCase().trim(),
      nif: limpiaNif(f.nif),
      numero: String(f.numero ?? "").trim(),
      fecha: normFecha(f.fecha),
      base, tipoIva: f.tipo_iva ?? 10, cuotaIva: iva, total, cuenta,
      fileName: "Listado PDF (Aaron)",
    });
  });
}

/* ---------- Facturas extraídas → filas de la tabla ---------- */
function facturasARows(facturas, fileName, tri, anio) {
  const rows = [];
  facturas.forEach((f) => {
    const invoiceId = uid();
    const lineas = Array.isArray(f.lineas) && f.lineas.length ? f.lineas : [{}];
    const ret = num(f.cuota_ret) || 0;
    lineas.forEach((l, idx) => {
      const base = num(l.base);
      const iva = num(l.cuota_iva);
      const re = num(l.cuota_re) || 0;
      const retFila = idx === 0 ? ret : 0;
      let total;
      if (lineas.length === 1 && isFinite(num(f.total))) total = num(f.total);
      else total = r2((isFinite(base) ? base : 0) + (isFinite(iva) ? iva : 0) + re - retFila);
      rows.push({
        id: uid(),
        invoiceId,
        fileName,
        contraparte: (f.contraparte || "").toUpperCase().trim(),
        nif: (f.nif || "0").toUpperCase().replace(/[\s\-\.]/g, ""),
        numero: String(f.numero ?? "").trim(),
        fecha: ajustaFechaATrimestre(f.fecha, tri, anio),
        base: toInput(base),
        tipoIva: toInputPct(l.tipo_iva ?? 0),
        cuotaIva: toInput(iva),
        tipoRe: toInputPct(l.tipo_re ?? 0),
        cuotaRe: toInput(re),
        tipoRet: toInputPct(idx === 0 ? f.tipo_ret ?? 0 : 0),
        cuotaRet: toInput(retFila),
        total: toInput(total),
        cuenta: Number(f.cuenta) || 10,
        concepto: conceptoPorCuenta(f.cuenta),
        clave: Number(f.cuenta) === 10 ? String(f.clave || "").toUpperCase().slice(0, 1) : "",
        confianza: f.confianza || "media",
        obs: f.obs || "",
      });
    });
  });
  return rows;
}

/* ---------- Facturas extraídas → filas de la tabla SL (doble partida) ----------
   Reutiliza la lectura IA de la pestaña de facturas y añade las tres subcuentas
   del Conversor. El asignador (creado por lote) empareja por NIF/nombre y
   autonumera las nuevas, recordando NIF→subcuenta. */
function facturasASLRows(facturas, fileName, ent, asignador, tri, anio, banco) {
  const rows = [];
  facturas.forEach((f) => {
    const invoiceId = uid();
    // Entidades "solo ingresos" (p.ej. Matilde Mateos): todo se registra como venta
    const sentido = ent.soloIngresos ? "venta" : (Number(f.cuenta) === 1 ? "venta" : "compra");
    const nif = (f.nif || "0").toUpperCase().replace(/[\s\-\.]/g, "");
    const contraparte = (f.contraparte || "").toUpperCase().trim();
    const fechaAj = ajustaFechaATrimestre(f.fecha, tri, anio); // traslada al trimestre si cae fuera
    const asg = asignador.asigna({ nif, nombre: contraparte, sentido });
    const lineas = Array.isArray(f.lineas) && f.lineas.length ? f.lineas : [{}];
    const ret = num(f.cuota_ret) || 0;
    lineas.forEach((l, idx) => {
      const base = num(l.base);
      const iva = num(l.cuota_iva);
      const re = num(l.cuota_re) || 0;
      const retFila = idx === 0 ? ret : 0;
      let total;
      if (lineas.length === 1 && isFinite(num(f.total))) total = num(f.total);
      else total = r2((isFinite(base) ? base : 0) + (isFinite(iva) ? iva : 0) + re - retFila);
      const tipoIva = l.tipo_iva ?? 0;
      rows.push({
        id: uid(), invoiceId, fileName, sentido,
        contraparte, nif,
        numero: String(f.numero ?? "").trim(),
        fechaFactura: fechaAj,
        fechaAsiento: fechaAj, // decisión: fecha del asiento = fecha de factura
        base: toInput(base),
        tipoIva: toInputPct(tipoIva),
        cuotaIva: toInput(iva),
        cuotaRe: toInput(re),
        cuotaRet: toInput(retFila),
        total: toInput(total),
        concepto: ent.soloIngresos ? "VENTAS" : conceptoConversor(f.cuenta), // 4 conceptos del Conversor
        subCP: asg.sub, subCPNombre: asg.nombreSub, subCPNueva: asg.nueva, subCPCreada: asg.creada,
        subGI: asg.giCode, subGINombre: asg.giNombre, subGINueva: asg.giNueva, subGIDescuadre: asg.giDescuadre,
        subIva: subIva(tipoIva, sentido, ent),
        subBanco: banco || "",
        subRe: "", subRet: "",
        confianza: f.confianza || "media", obs: f.obs || "",
      });
    });
  });
  return rows;
}

/* ---------- Listado de ingresos (Matilde) → filas SL ----------
   Cada línea es una venta. Empareja el cliente por CIF si el listado lo trae
   (2T en adelante), o por nombre/razón si no (1T). IVA y total del propio listado. */
function listadoMatildeASLRows(items, fileName, ent, asignador, tri, anio, banco) {
  const rows = [];
  items.forEach((f) => {
    const razon = String(f.razon || f.contraparte || "").toUpperCase().trim();
    const nif = limpiaNif(f.cif); // "0" si el listado aún no trae CIF
    const asg = asignador.asigna({ nif, nombre: razon, sentido: "venta" });
    const base = num(f.base);
    const iva = num(f.cuota_iva);
    let tipoIva = num(f.tipo_iva);
    if (!isFinite(tipoIva) || tipoIva === 0)
      tipoIva = (isFinite(base) && base !== 0 && isFinite(iva)) ? Math.round((iva / base) * 100) : 21;
    let total = num(f.total);
    if (!isFinite(total)) total = r2((isFinite(base) ? base : 0) + (isFinite(iva) ? iva : 0));
    const fechaAj = ajustaFechaATrimestre(f.fecha, tri, anio);
    rows.push({
      id: uid(), invoiceId: uid(), fileName, sentido: "venta",
      contraparte: razon, nif,
      numero: String(f.numero ?? "").trim(),
      fechaFactura: fechaAj, fechaAsiento: fechaAj,
      base: toInput(base), tipoIva: toInputPct(tipoIva), cuotaIva: toInput(iva),
      cuotaRe: "0,00", cuotaRet: "0,00",
      total: toInput(total),
      concepto: "VENTAS",
      subCP: asg.sub, subCPNombre: asg.nombreSub, subCPNueva: asg.nueva, subCPCreada: asg.creada,
      subGI: asg.giCode, subGINombre: asg.giNombre, subGINueva: asg.giNueva, subGIDescuadre: asg.giDescuadre,
      subIva: subIva(tipoIva, "venta", ent),
      subBanco: banco || "",
      subRe: "", subRet: "",
      confianza: f.confianza || "media", obs: f.obs || "",
    });
  });
  return rows;
}

/* ---------- Validación de cada fila ---------- */
function validarFila(row, rows, tri, anio) {
  const issues = [];
  const base = num(row.base), iva = num(row.cuotaIva), tipo = num(row.tipoIva);
  const re = num(row.cuotaRe) || 0, ret = num(row.cuotaRet) || 0, total = num(row.total);

  if (!row.contraparte) issues.push({ lv: "err", msg: "Falta el nombre del cliente/proveedor" });
  if (!row.numero) issues.push({ lv: "warn", msg: "Falta el número de factura" });
  if (!fechaValida(row.fecha)) issues.push({ lv: "err", msg: "Fecha inválida (dd/mm/aaaa)" });
  if (!isFinite(base)) issues.push({ lv: "err", msg: "Base imponible no numérica" });
  if (!isFinite(iva)) issues.push({ lv: "err", msg: "Cuota de IVA no numérica" });
  if (!isFinite(total)) issues.push({ lv: "err", msg: "Total no numérico" });
  if (!validNif(row.nif)) issues.push({ lv: "warn", msg: "NIF con formato dudoso" });
  if (Number(row.cuenta) === 10 && !row.clave) issues.push({ lv: "warn", msg: "Gasto diverso sin clave (R/P/E/A/G/I/S/O)" });
  if (!row.concepto) issues.push({ lv: "warn", msg: "Falta el concepto del apunte" });
  if (row.confianza === "baja") issues.push({ lv: "warn", msg: `Extracción con confianza baja${row.obs ? ": " + row.obs : ""}` });

  if (isFinite(base) && isFinite(iva) && isFinite(tipo)) {
    const esperado = r2((base * tipo) / 100);
    if (Math.abs(esperado - iva) > 0.02)
      issues.push({ lv: "warn", msg: `Cuota IVA no cuadra: ${fmtES(base)} × ${row.tipoIva}% = ${fmtES(esperado)}` });
  }
  if (isFinite(base) && isFinite(iva) && isFinite(total)) {
    const esperado = r2(base + iva + re - ret);
    if (Math.abs(esperado - total) > 0.02)
      issues.push({ lv: "warn", msg: `Total no cuadra: base + IVA + RE − ret. = ${fmtES(esperado)}` });
  }
  const dup = rows.some(
    (o) => o.id !== row.id && o.invoiceId !== row.invoiceId && o.nif === row.nif && o.nif !== "0" && o.numero === row.numero && o.numero !== ""
  );
  if (dup) issues.push({ lv: "warn", msg: "Posible factura duplicada (mismo NIF y nº de factura)" });
  if (tri && anio && fechaValida(row.fecha) && !fechaEnTrimestre(row.fecha, tri, anio)) {
    issues.push({ lv: "warn", msg: `Fecha fuera del trimestre seleccionado (${tri}T ${anio}: ${rangoTrimestre(tri, anio)})` });
  }
  return issues;
}

/* ============================================================ */

/* ============================================================
   PLANTILLAS FIJAS para listados Excel/PDF de clientes concretos.
   No usan IA: leen columnas en posiciones fijas → exactas y gratis.
   Cada plantilla devuelve filas en el formato de la tabla de revisión.
   ============================================================ */

/* Convierte un número de serie de fecha de Excel a dd/mm/aaaa */
const excelFechaANum = (v) => {
  if (v == null || v === "") return "";
  if (typeof v === "string") return normFecha(v);
  if (v instanceof Date) {
    const dd = String(v.getDate()).padStart(2, "0");
    const mm = String(v.getMonth() + 1).padStart(2, "0");
    return `${dd}/${mm}/${v.getFullYear()}`;
  }
  if (typeof v === "number") {
    // Serie de Excel: día 0 = 30/12/1899
    const ms = Math.round((v - 25569) * 86400 * 1000);
    const d = new Date(ms);
    if (isNaN(d.getTime())) return "";
    const dd = String(d.getUTCDate()).padStart(2, "0");
    const mm = String(d.getUTCMonth() + 1).padStart(2, "0");
    return `${dd}/${mm}/${d.getUTCFullYear()}`;
  }
  return "";
};

/* Devuelve el trimestre (1-4) de una fecha dd/mm/aaaa, o 0 si no es válida */
const trimestreDeFecha = (fechaStr) => {
  const m = String(fechaStr || "").match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (!m) return 0;
  const mes = parseInt(m[2], 10);
  return Math.floor((mes - 1) / 3) + 1;
};

/* ---------- Trimestre + año global a contabilizar ---------- */
/* ¿La fecha dd/mm/aaaa cae dentro del trimestre (1-4) + año? Sin fecha válida → true
   (no añadimos este aviso encima de "fecha inválida"). */
const fechaEnTrimestre = (fechaStr, tri, anio) => {
  const m = String(fechaStr || "").match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (!m) return true;
  return parseInt(m[3], 10) === Number(anio) && Math.floor((parseInt(m[2], 10) - 1) / 3) + 1 === Number(tri);
};
/* Texto del rango del trimestre, p.ej. "01/04 a 30/06" */
const rangoTrimestre = (tri, anio) => {
  const ini = (tri - 1) * 3 + 1, fin = tri * 3;
  const diaFin = new Date(anio, fin, 0).getDate();
  return `01/${String(ini).padStart(2, "0")} a ${String(diaFin).padStart(2, "0")}/${String(fin).padStart(2, "0")}`;
};
/* Primer día del trimestre como dd/mm/aaaa (1T→01/01, 2T→01/04, 3T→01/07, 4T→01/10) */
const primerDiaTrimestre = (tri, anio) => {
  const ini = (tri - 1) * 3 + 1;
  return `01/${String(ini).padStart(2, "0")}/${anio}`;
};
/* Ajusta una fecha al trimestre: si es válida pero cae fuera, la traslada al
   primer día del trimestre seleccionado; si no, la deja como está. */
const ajustaFechaATrimestre = (fechaStr, tri, anio) => {
  const f = normFecha(fechaStr);
  if (tri && anio && fechaValida(f) && !fechaEnTrimestre(f, tri, anio)) return primerDiaTrimestre(tri, anio);
  return f;
};

const limpiaNif = (v) => {
  let s = String(v ?? "").toUpperCase().replace(/[\s\-\.\/]/g, "");
  // Prefijo de IVA intracomunitario español: ES + NIF
  if (/^ES[0-9A-Z]/.test(s)) s = s.slice(2);
  return s || "0";
};

/* Lee un File (xlsx/xls) y devuelve la primera hoja como matriz de filas */
async function leerHojaExcel(file) {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array", cellDates: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "" });
}

/* Plantilla VENTAS Aaron - Serie Principal 001
   Datos desde fila 8 (índice 7). B=nº(1) C=fecha(2) D=NIF(3) E=nombre(4)
   L=base(11) M=iva(12) N=total(13) */
function parseVentasSerie001(rows) {
  const out = [];
  for (let i = 7; i < rows.length; i++) {
    const r = rows[i];
    const numero = String(r[1] ?? "").trim();
    const nombre = String(r[4] ?? "").trim();
    const base = num(r[11]);
    if (!nombre || !numero || !isFinite(base)) continue;
    const iva = num(r[12]);
    const total = num(r[13]);
    const tipo = isFinite(base) && base !== 0 && isFinite(iva) ? Math.round((iva / base) * 100) : 21;
    out.push(filaListado({
      contraparte: nombre.toUpperCase(), nif: limpiaNif(r[3]), numero,
      fecha: excelFechaANum(r[2]), base, tipoIva: tipo, cuotaIva: iva, total,
      cuenta: 1, fileName: "Ventas Serie 001",
    }));
  }
  return out;
}

/* Plantilla VENTAS Aaron - TPV Serie 002
   Datos desde fila 8. B=nº(1) C=fecha(2) D=nombre(3) E=NIF(4)
   G=base(6) H=iva(7) I=total(8) */
function parseVentasTPV002(rows) {
  const out = [];
  for (let i = 7; i < rows.length; i++) {
    const r = rows[i];
    const numero = String(r[1] ?? "").trim();
    const nombre = String(r[3] ?? "").trim();
    const base = num(r[6]);
    if (!nombre || !numero || !isFinite(base)) continue;
    const iva = num(r[7]);
    const total = num(r[8]);
    const tipo = isFinite(base) && base !== 0 && isFinite(iva) ? Math.round((iva / base) * 100) : 21;
    out.push(filaListado({
      contraparte: nombre.toUpperCase(), nif: limpiaNif(r[4]), numero,
      fecha: excelFechaANum(r[2]), base, tipoIva: tipo, cuotaIva: iva, total,
      cuenta: 1, fileName: "Ventas TPV 002",
    }));
  }
  return out;
}

/* Plantilla PROVEEDORES Ferretería (.xls)
   Datos desde fila 8. D=fecha(3) E=código(4) G=proveedor(6) H=CIF(7)
   Hasta 4 tramos IVA: (base,%,iva) en col 10/_/9, 11/12/13, 15/16/17, 19/20/21
   Globales: W=base(23) X=iva(24) AA=total(26).
   Si hay un solo tramo → 1 fila; si hay varios → 1 fila por tramo. */
function parseProveedoresFerreteria(rows) {
  const out = [];
  for (let i = 8; i < rows.length; i++) {
    const r = rows[i];
    const cod = String(r[4] ?? "").trim();
    const prov = String(r[6] ?? "").trim();
    const baseG = num(r[23]);
    if (!prov || !cod || !isFinite(baseG)) continue;
    const nif = limpiaNif(r[7]);
    const fecha = excelFechaANum(r[3]);
    // Detectar tramos rellenos: pares (base, iva)
    const tramos = [];
    const pares = [[10, 9], [11, 13], [15, 17], [19, 21]]; // [colBase, colIva]
    for (const [cb, ci] of pares) {
      const b = num(r[cb]); const q = num(r[ci]);
      if (isFinite(b) && b !== 0) tramos.push({ base: b, iva: isFinite(q) ? q : 0 });
    }
    if (tramos.length === 0) {
      // Sin tramos detallados: usar base/iva globales
      const ivaG = num(r[24]);
      const totalG = num(r[26]);
      const tipo = baseG !== 0 && isFinite(ivaG) ? Math.round((ivaG / baseG) * 100) : 21;
      out.push(filaListado({
        contraparte: prov.toUpperCase(), nif, numero: cod, fecha,
        base: baseG, tipoIva: tipo, cuotaIva: ivaG, total: totalG,
        cuenta: 2, fileName: "Proveedores Ferretería",
      }));
    } else {
      const invoiceId = uid();
      tramos.forEach((t) => {
        const tipo = t.base !== 0 && isFinite(t.iva) ? Math.round((t.iva / t.base) * 100) : 21;
        out.push(filaListado({
          contraparte: prov.toUpperCase(), nif, numero: cod, fecha,
          base: t.base, tipoIva: tipo, cuotaIva: t.iva, total: r2(t.base + t.iva),
          cuenta: 2, fileName: "Proveedores Ferretería", invoiceId,
        }));
      });
    }
  }
  return out;
}

/* Plantilla VENTAS Aaron — por CABECERA (robusta al orden de columnas).
   Localiza las columnas por su nombre (CLIENTE, FACTURA, FECHA, BASE, IVA,
   R.E, IMPORTE, CIF) en lugar de por posición fija: así, aunque Aaron cambie
   el orden de las columnas entre meses, las encuentra igual.
   Tipo de IVA calculado de base/IVA. Cuenta V (ventas). Abonos negativos tal cual. */
function parseVentasAaronCabecera(rows) {
  // Buscar la fila de cabecera: la que contiene "FACTURA" y "BASE"
  let h = -1;
  for (let i = 0; i < Math.min(rows.length, 15); i++) {
    const celdas = (rows[i] || []).map((c) => String(c ?? "").trim().toUpperCase());
    if (celdas.includes("FACTURA") && celdas.includes("BASE")) { h = i; break; }
  }
  if (h === -1) return [];
  const head = (rows[h] || []).map((c) => String(c ?? "").trim().toUpperCase());
  // Localizar índice de cada columna por palabra clave (tolerante a variaciones)
  const col = (claves) => head.findIndex((c) => claves.some((k) => c.includes(k)));
  const iCli = col(["CLIENTE", "NOMBRE", "RAZON"]);
  const iFac = col(["FACTURA"]);
  const iFec = col(["FECHA"]);
  const iBase = head.findIndex((c) => c === "BASE" || c.startsWith("BASE"));
  const iIva = head.findIndex((c) => c === "IVA" || c.startsWith("IVA"));
  const iImp = col(["IMPORTE", "TOTAL"]);
  const iNif = col(["CIF", "NIF", "DNI"]);
  const out = [];
  for (let i = h + 1; i < rows.length; i++) {
    const r = rows[i] || [];
    const nombre = String(r[iCli] ?? "").trim();
    const numero = String(r[iFac] ?? "").trim();
    const base = num(r[iBase]);
    if (!nombre || !numero || !isFinite(base)) continue;
    // Saltar fila de totales (suele no tener nº de factura válido, ya filtrado)
    const iva = num(r[iIva]);
    const total = iImp >= 0 ? num(r[iImp]) : NaN;
    const tipo = isFinite(base) && base !== 0 && isFinite(iva) ? Math.round((iva / base) * 100) : 10;
    out.push(filaListado({
      contraparte: nombre.toUpperCase(),
      nif: limpiaNif(iNif >= 0 ? r[iNif] : ""),
      numero,
      fecha: excelFechaANum(r[iFec]),
      base, tipoIva: tipo, cuotaIva: iva,
      total: isFinite(total) ? total : r2(base + (isFinite(iva) ? iva : 0)),
      cuenta: 1, fileName: "Ventas Aaron",
    }));
  }
  return out;
}

/* Construye una fila de la tabla de revisión a partir de datos de listado */
function filaListado({ contraparte, nif, numero, fecha, base, tipoIva, cuotaIva, total, cuenta, fileName, invoiceId }) {
  return {
    id: uid(), invoiceId: invoiceId || uid(), fileName,
    contraparte, nif, numero, fecha,
    base: toInput(base), tipoIva: toInputPct(tipoIva), cuotaIva: toInput(cuotaIva),
    tipoRe: "0", cuotaRe: "0,00", tipoRet: "0", cuotaRet: "0,00",
    total: toInput(total), cuenta, concepto: conceptoPorCuenta(cuenta),
    clave: "", confianza: "alta", obs: "",
  };
}

/* ---------- Prompt para PDF de LISTADO (doble columna ventas/compras) ---------- */
const promptListadoPDF = `Eres el sistema de extracción contable de una asesoría española. El documento adjunto es un LISTADO mensual de contabilidad con DOS bloques en la misma página:
- IZQUIERDA: VENTAS A CLIENTES (columnas: CLIENTE, FACTURA, FECHA, BASE, IVA, R.E, IMPORTE, ESTADO).
- DERECHA: COMPRAS A PROVEEDORES (columnas: FECHA ALBARÁN, FECHA, PROVEEDOR, FACTURA, BASE, IVA, IMPORTE).
Las dos tablas están alineadas lado a lado pero son INDEPENDIENTES: cada fila de ventas no tiene relación con la fila de compras a su derecha.

Devuelve EXCLUSIVAMENTE un objeto JSON válido, sin markdown ni texto alrededor:
{"facturas":[{"contraparte":"NOMBRE","nif":"0","numero":"25030001","fecha":"01/03/2025","base":666.10,"tipo_iva":10,"cuota_iva":66.61,"recargo":0,"total":732.71,"sentido":"venta"}]}

REGLAS:
- Una entrada por cada línea de venta (sentido:"venta") y una por cada línea de compra (sentido:"compra").
- Ignora filas de cabecera, subtotales y la fila de TOTALES del final.
- Importes con punto decimal y 2 decimales. Si el número usa coma decimal en el documento (666,10), conviértelo a punto (666.10).
- "nif": estos listados normalmente NO traen NIF. Si no aparece, pon "0". Si en alguna fila SÍ hay un NIF/CIF/DNI, ponlo en mayúsculas, sin espacios, guiones ni puntos, y sin el prefijo ES.
- "numero": el número de factura tal cual (en ventas suele ser tipo 25030001).
- "fecha": dd/mm/aaaa.
- "tipo_iva": calcúlalo de base e IVA (suele ser 10 en cárnicas). "recargo": el R.E si lo hay, si no 0.
- Las líneas de ABONO o importes negativos: respétalos con su signo negativo.
- "total": el importe total de esa línea.
Sé exhaustivo: extrae TODAS las líneas de ambas columnas, de todas las páginas.`;

/* ---------- Prompt para LISTADO DE INGRESOS (administración de comunidades) ---------- */
const promptListadoIngresos = `Eres el sistema de extracción contable de una asesoría española. El documento es un LISTADO de FACTURAS EMITIDAS (ingresos) de una administración de fincas y comunidades. Cada fila es una factura de ingreso a una comunidad o cliente.

Columnas habituales: Nº Factura, Fecha, Código, Razón (nombre de la comunidad/cliente), Cuotas (base imponible), IRPF, IVA (importe), Total. ALGUNOS listados traen ADEMÁS una columna con el CIF/NIF del cliente: si existe, extráela; si no, deja "cif":"".

Devuelve EXCLUSIVAMENTE un objeto JSON válido, sin markdown ni texto alrededor:
{"facturas":[{"numero":"000001","fecha":"16/01/2026","razon":"MARQUES DE LA ENSENADA 17","cif":"","base":70.00,"tipo_iva":21,"cuota_iva":14.70,"total":84.70}]}

REGLAS:
- Una entrada en "facturas" por cada fila de factura. Ignora cabeceras, subtotales y la fila de TOTALES del final.
- "razon": el nombre del cliente/comunidad tal cual figura en la columna Razón, en mayúsculas.
- "cif": el CIF/NIF del cliente SOLO si el listado trae esa columna; en mayúsculas, sin espacios ni puntos ni guiones. Si no aparece, "".
- Importes con punto decimal y 2 decimales (convierte la coma decimal a punto: 70,00 → 70.00).
- "base": columna Cuotas. "cuota_iva": columna IVA (importe). "total": columna Total.
- "tipo_iva": calcúlalo de base e IVA y redondea (14.70 / 70.00 ≈ 21). Si IRPF no es 0, no lo restes de la base.
- "fecha": dd/mm/aaaa.
- Sé exhaustivo: TODAS las filas de TODAS las páginas.`;

const PLANTILLAS = [
  { id: "ventas001", nombre: "Excel · Ferretería El Paso — Ventas Serie 001 (comercio mayor)", tipo: "excel", trimestral: true, fn: parseVentasSerie001 },
  { id: "ventasTPV", nombre: "Excel · Ferretería El Paso — Ventas Serie 002 (TPV tienda)", tipo: "excel", trimestral: true, fn: parseVentasTPV002 },
  { id: "provFerre", nombre: "Excel · Ferretería El Paso — Proveedores (compras)", tipo: "excel", trimestral: true, fn: parseProveedoresFerreteria },
  { id: "ventasAaron", nombre: "Excel · Aaron — Ingresos (detecta columnas por cabecera)", tipo: "excel", fn: parseVentasAaronCabecera },
  { id: "pdfAaron", nombre: "PDF · Listado mensual ventas+compras (red de seguridad)", tipo: "pdf" },
];


export default function App() {
  const [claveDespacho, setClaveDespacho] = useState(() => {
    try { return limpiaClave(localStorage.getItem("asema_clave")); } catch { return ""; }
  });
  const [autenticado, setAutenticado] = useState(() => {
    try { return !!localStorage.getItem("asema_clave"); } catch { return false; }
  });
  const [claveInput, setClaveInput] = useState("");
  const [avisoClave, setAvisoClave] = useState("");

  const [modo, setModo] = useState("facturas"); // "facturas" | "listados" | "sl"
  const [contabSelId, setContabSelId] = useState(CONTABILIDADES_SL[0]?.id || "");
  const [rowsSL, setRowsSL] = useState([]);
  const [slMsg, setSlMsg] = useState("");
  const [bancoSel, setBancoSel] = useState(() => cargarBanco(CONTABILIDADES_SL[0]));
  const [plantillaSel, setPlantillaSel] = useState("ventas001");
  const [trimestreSel, setTrimestreSel] = useState(() => Math.floor(new Date().getMonth() / 3) + 1);
  const [anioSel, setAnioSel] = useState(() => new Date().getFullYear());
  const [listadoMsg, setListadoMsg] = useState("");
  const [empresa, setEmpresa] = useState({ nombre: "", nif: "" });
  const [guardadas, setGuardadas] = useState([]);
  const [files, setFiles] = useState([]);
  const [rows, setRows] = useState([]);
  const [procesando, setProcesando] = useState(false);
  const [aviso, setAviso] = useState("");
  const [arrastrando, setArrastrando] = useState(false);
  const inputRef = useRef(null);
  const fileObjs = useRef({});

  /* ----- Clientes guardados (localStorage de cada PC del despacho) ----- */
  useEffect(() => {
    try {
      const r = localStorage.getItem("asema_empresas");
      if (r) setGuardadas(JSON.parse(r));
    } catch { /* sin clientes guardados todavía */ }
  }, []);

  /* Banco (contrapartida) por defecto al cambiar de contabilidad SL */
  useEffect(() => {
    setBancoSel(cargarBanco(CONTABILIDADES_SL.find((e) => e.id === contabSelId)));
  }, [contabSelId]);

  const guardarEmpresa = () => {
    if (!empresa.nombre || !empresa.nif) return;
    const limpia = { nombre: empresa.nombre.toUpperCase().trim(), nif: empresa.nif.toUpperCase().replace(/[\s\-\.]/g, "") };
    const nuevas = [limpia, ...guardadas.filter((e) => e.nif !== limpia.nif)].slice(0, 50);
    setGuardadas(nuevas);
    try { localStorage.setItem("asema_empresas", JSON.stringify(nuevas)); } catch { /* opcional */ }
  };

  /* ----- Acceso ----- */
  const entrar = () => {
    const k = limpiaClave(claveInput);
    if (!k) { setAvisoClave("Escribe la clave del despacho."); return; }
    try { localStorage.setItem("asema_clave", k); } catch { /* seguimos en memoria */ }
    setClaveDespacho(k);
    setAutenticado(true);
    setAvisoClave("");
  };
  const salir = () => {
    try { localStorage.removeItem("asema_clave"); } catch { /* nada */ }
    setClaveDespacho("");
    setClaveInput("");
    setAutenticado(false);
  };

  /* ----- Gestión de archivos ----- */
  const addFiles = (lista) => {
    const nuevos = [];
    Array.from(lista).forEach((f) => {
      const okTipo = f.type === "application/pdf" || /^image\/(jpeg|png|webp)$/.test(f.type) || /\.(pdf|jpe?g|png|webp|docx|doc)$/i.test(f.name);
      if (!okTipo) return;
      const grande = f.size > MAX_MB * 1024 * 1024;
      const id = uid();
      fileObjs.current[id] = f;
      nuevos.push({
        id,
        name: f.name,
        size: f.size,
        status: grande ? "error" : "pendiente",
        msg: grande ? `Archivo demasiado grande (máx. ${MAX_MB} MB) — escanea a 150-200 ppp o divide el PDF` : "",
        nFacturas: 0,
      });
    });
    if (nuevos.length) setFiles((p) => [...p, ...nuevos]);
  };

  const setFile = (id, patch) => setFiles((p) => p.map((f) => (f.id === id ? { ...f, ...patch } : f)));

  const procesarPendientes = async () => {
    const entSL = modo === "sl" ? CONTABILIDADES_SL.find((e) => e.id === contabSelId) : null;
    if (modo === "sl") {
      if (!entSL) { setAviso("Elige una contabilidad SL antes de procesar."); return; }
    } else if (!empresa.nombre || !empresa.nif) {
      setAviso("Antes de procesar, indica el nombre y NIF del cliente: la app lo necesita para distinguir facturas emitidas de recibidas.");
      return;
    }
    setAviso("");
    setProcesando(true);
    if (modo !== "sl") guardarEmpresa();
    // En SL, el titular es la contabilidad elegida; el asignador empareja y autonumera subcuentas
    const titular = entSL ? { nombre: entSL.nombre, nif: entSL.nif || "0" } : empresa;
    const asignador = entSL ? crearAsignador(entSL, cargarMapaNif(entSL.id)) : null;
    const pendientes = files.filter((f) => f.status === "pendiente" || f.status === "error");
    for (const f of pendientes) {
      const obj = fileObjs.current[f.id];
      if (!obj || obj.size > MAX_MB * 1024 * 1024) continue;
      setFile(f.id, { status: "procesando", msg: "" });
      try {
        const esWord = /\.(docx|doc)$/i.test(obj.name);
        const facturas = esWord
          ? await extraerFacturaWord(obj, titular, claveDespacho)
          : await extraerFacturas(obj, titular, claveDespacho, (b, n) => setFile(f.id, { progreso: `bloque ${b}/${n}` }));
        if (entSL) {
          const nuevas = facturasASLRows(facturas, f.name, entSL, asignador, trimestreSel, anioSel, bancoSel);
          setRowsSL((p) => [...p, ...nuevas]);
        } else {
          const nuevas = facturasARows(facturas, f.name, trimestreSel, anioSel);
          setRows((p) => [...p, ...nuevas]);
        }
        setFile(f.id, { status: "ok", nFacturas: facturas.length, msg: "", progreso: "" });
      } catch (e) {
        setFile(f.id, { status: "error", msg: e.message || "Error desconocido" });
        if (e.auth) {
          setProcesando(false);
          salir();
          setAvisoClave("La clave del despacho no es correcta. Vuelve a introducirla.");
          return;
        }
      }
    }
    if (asignador) guardarMapaNif(entSL.id, asignador.dump());
    setProcesando(false);
  };

  const quitarFile = (id) => {
    setFiles((p) => p.filter((f) => f.id !== id));
    delete fileObjs.current[id];
  };

  /* ----- Edición de filas ----- */
  const upd = (id, campo, valor) =>
    setRows((p) =>
      p.map((r) => {
        if (r.id !== id) return r;
        const nr = { ...r, [campo]: valor };
        if (campo === "cuenta") {
          nr.concepto = conceptoPorCuenta(valor);
          if (Number(valor) !== 10) nr.clave = "";
        }
        return nr;
      })
    );

  const recalcular = (id) =>
    setRows((p) =>
      p.map((r) => {
        if (r.id !== id) return r;
        const base = num(r.base);
        if (!isFinite(base)) return r;
        const iva = r2((base * (num(r.tipoIva) || 0)) / 100);
        const re = r2((base * (num(r.tipoRe) || 0)) / 100);
        const ret = r2((base * (num(r.tipoRet) || 0)) / 100);
        return { ...r, cuotaIva: toInput(iva), cuotaRe: toInput(re), cuotaRet: toInput(ret), total: toInput(r2(base + iva + re - ret)) };
      })
    );

  const borrarFila = (id) => setRows((p) => p.filter((r) => r.id !== id));

  const filaManual = () =>
    setRows((p) => [
      ...p,
      {
        id: uid(), invoiceId: uid(), fileName: "manual",
        contraparte: "", nif: "0", numero: "", fecha: "",
        base: "", tipoIva: "21", cuotaIva: "", tipoRe: "0", cuotaRe: "0,00",
        tipoRet: "0", cuotaRet: "0,00", total: "",
        cuenta: 10, concepto: conceptoPorCuenta(10), clave: "", confianza: "alta", obs: "",
      },
    ]);

  /* ----- Edición de filas SL ----- */
  const updSL = (id, campo, valor) =>
    setRowsSL((p) => p.map((r) => (r.id === id ? { ...r, [campo]: valor } : r)));

  const recalcularSL = (id) =>
    setRowsSL((p) =>
      p.map((r) => {
        if (r.id !== id) return r;
        const base = num(r.base);
        if (!isFinite(base)) return r;
        const iva = r2((base * (num(r.tipoIva) || 0)) / 100);
        const re = num(r.cuotaRe) || 0;
        const ret = num(r.cuotaRet) || 0;
        return { ...r, cuotaIva: toInput(iva), total: toInput(r2(base + iva + re - ret)) };
      })
    );

  const borrarFilaSL = (id) => setRowsSL((p) => p.filter((r) => r.id !== id));

  /* Contabilidad SL seleccionada (para los gates de subida y el importador de ingresos) */
  const entSLSel = CONTABILIDADES_SL.find((e) => e.id === contabSelId);

  /* ----- Importación del listado de ingresos (Matilde, solo ingresos) ----- */
  const importarListadoMatilde = async (file) => {
    if (!file) return;
    const ent = entSLSel;
    if (!ent) { setSlMsg("Elige una contabilidad antes de importar."); return; }
    if (!/\.pdf$/i.test(file.name)) { setSlMsg("El listado de ingresos debe ser un PDF."); return; }
    if (file.size > MAX_MB * 1024 * 1024) { setSlMsg(`El PDF supera los ${MAX_MB} MB. Divídelo o redúcelo.`); return; }
    setSlMsg("Leyendo el listado de ingresos con IA… (puede tardar unos segundos)");
    try {
      const items = await extraerListadoIngresos(file, claveDespacho);
      if (!items.length) { setSlMsg("La IA no encontró líneas en el listado. Revisa el PDF."); return; }
      const asignador = crearAsignador(ent, cargarMapaNif(ent.id));
      const nuevas = listadoMatildeASLRows(items, file.name, ent, asignador, trimestreSel, anioSel, bancoSel);
      guardarMapaNif(ent.id, asignador.dump());
      setRowsSL((p) => [...p, ...nuevas]);
      const conCif = nuevas.filter((r) => r.nif && r.nif !== "0").length;
      setSlMsg(`Importadas ${nuevas.length} líneas de ingresos${conCif ? ` (${conCif} con CIF)` : " (sin columna CIF: clientes emparejados por nombre, NIF en blanco)"}. Revísalas abajo antes de exportar.`);
    } catch (e) {
      if (e.auth) { salir(); setAvisoClave("La clave del despacho no es correcta. Vuelve a introducirla."); return; }
      setSlMsg("Error al leer el listado: " + (e.message || "desconocido"));
    }
  };

  /* ----- Importación de listados Excel (plantillas fijas) ----- */
  const importarListado = async (file) => {
    if (!file) return;
    const plantilla = PLANTILLAS.find((p) => p.id === plantillaSel);
    if (!plantilla) return;
    try {
      if (plantilla.tipo === "pdf") {
        // Listado PDF (Aaron) por IA: requiere la clave del despacho
        const esPdf = /\.pdf$/i.test(file.name);
        if (!esPdf) { setListadoMsg("Esta plantilla espera un archivo PDF."); return; }
        if (file.size > MAX_MB * 1024 * 1024) { setListadoMsg(`El PDF supera los ${MAX_MB} MB. Divídelo o reduce su tamaño.`); return; }
        setListadoMsg("Leyendo el listado PDF con IA… (puede tardar unos segundos)");
        const items = await extraerListadoPDF(file, claveDespacho);
        const nuevas = listadoPDFARows(items);
        if (!nuevas.length) { setListadoMsg("La IA no encontró líneas en el PDF. Revisa el documento."); return; }
        setRows((p) => [...p, ...nuevas]);
        const nv = nuevas.filter((r) => Number(r.cuenta) === 1).length;
        const nc = nuevas.length - nv;
        setListadoMsg(`Importadas ${nuevas.length} líneas del PDF (${nv} ventas, ${nc} compras). Revísalas abajo antes de exportar.`);
        return;
      }
      // Plantilla Excel (sin IA)
      const esExcel = /\.(xlsx|xls)$/i.test(file.name);
      if (!esExcel) { setListadoMsg("Esta plantilla espera un archivo Excel (.xlsx/.xls)."); return; }
      setListadoMsg("Leyendo archivo Excel…");
      const filas = await leerHojaExcel(file);
      let nuevas = plantilla.fn(filas);
      if (!nuevas.length) { setListadoMsg("No se han encontrado filas válidas. ¿Es la plantilla correcta para este archivo?"); return; }
      // Filtro por trimestre para clientes que acumulan todo el año (Ferretería)
      if (plantilla.trimestral) {
        const total = nuevas.length;
        const sinFecha = nuevas.filter((r) => trimestreDeFecha(r.fecha) === 0).length;
        nuevas = nuevas.filter((r) => trimestreDeFecha(r.fecha) === Number(trimestreSel));
        if (!nuevas.length) {
          setListadoMsg(`El archivo trae ${total} líneas, pero ninguna del ${trimestreSel}º trimestre. ¿Has elegido el trimestre correcto?`);
          return;
        }
        setRows((p) => [...p, ...nuevas]);
        const aviso = sinFecha > 0 ? ` (${sinFecha} líneas sin fecha válida quedaron fuera del filtro)` : "";
        setListadoMsg(`Importadas ${nuevas.length} líneas del ${trimestreSel}º trimestre (de ${total} totales en el archivo)${aviso}. Revísalas abajo antes de exportar.`);
        return;
      }
      setRows((p) => [...p, ...nuevas]);
      setListadoMsg(`Importadas ${nuevas.length} líneas desde "${file.name}". Revísalas abajo antes de exportar.`);
    } catch (e) {
      if (e.auth) { setListadoMsg("Clave del despacho incorrecta para procesar el PDF con IA."); return; }
      setListadoMsg("Error al procesar el archivo: " + (e.message || "desconocido"));
    }
  };

  /* ----- Validación y resumen ----- */
  const issuesById = useMemo(() => {
    const m = {};
    rows.forEach((r) => { m[r.id] = validarFila(r, rows, trimestreSel, anioSel); });
    return m;
  }, [rows, trimestreSel, anioSel]);

  const nErr = rows.filter((r) => issuesById[r.id]?.some((i) => i.lv === "err")).length;
  const nWarn = rows.filter((r) => issuesById[r.id]?.some((i) => i.lv === "warn") && !issuesById[r.id]?.some((i) => i.lv === "err")).length;

  const issuesSLById = useMemo(() => {
    const m = {};
    rowsSL.forEach((r) => { m[r.id] = validarFilaSL(r, trimestreSel, anioSel); });
    return m;
  }, [rowsSL, trimestreSel, anioSel]);
  const nErrSL = rowsSL.filter((r) => issuesSLById[r.id]?.some((i) => i.lv === "err")).length;
  const nWarnSL = rowsSL.filter((r) => issuesSLById[r.id]?.some((i) => i.lv === "warn") && !issuesSLById[r.id]?.some((i) => i.lv === "err")).length;

  const resumen = useMemo(() => {
    const porCuenta = {};
    let baseVentas = 0, ivaRep = 0, baseGastos = 0, ivaSop = 0;
    rows.forEach((r) => {
      const c = Number(r.cuenta);
      const b = num(r.base), i = num(r.cuotaIva);
      if (!porCuenta[c]) porCuenta[c] = { n: 0, base: 0 };
      porCuenta[c].n += 1;
      if (isFinite(b)) porCuenta[c].base = r2(porCuenta[c].base + b);
      if (c === 1) {
        if (isFinite(b)) baseVentas = r2(baseVentas + b);
        if (isFinite(i)) ivaRep = r2(ivaRep + i);
      } else {
        if (isFinite(b)) baseGastos = r2(baseGastos + b);
        if (isFinite(i)) ivaSop = r2(ivaSop + i);
      }
    });
    return { porCuenta, baseVentas, ivaRep, baseGastos, ivaSop };
  }, [rows]);

  /* ----- Exportación a Excel para Monitor ----- */
  const exportar = () => {
    if (!rows.length) return;
    if (modo === "facturas" && (!empresa.nombre || !empresa.nombre.trim())) {
      setAviso("Antes de exportar, rellena el nombre del cliente del despacho (arriba). Es obligatorio para identificar el archivo.");
      window.scrollTo({ top: 0, behavior: "smooth" });
      return;
    }
    const orden = [...rows].sort((a, b) => {
      const c = Number(a.cuenta) - Number(b.cuenta);
      if (c !== 0) return c;
      const f = fechaKey(a.fecha).localeCompare(fechaKey(b.fecha));
      if (f !== 0) return f;
      return a.numero.localeCompare(b.numero, "es");
    });
    const headers = [
      "NOMBRE CLIENTE-PROVEEDOR", "NIF", "Nº FACTURA", "FECHA",
      "BASE IMPONIBLE", "CANTIDAD IVA", "TOTAL", "TIPO DE IVA", "CUENTA", "CONCEPTO",
      "CLAVE GASTO", "TIPO RETENCION", "CUOTA RETENCION", "TIPO RE", "CUOTA RE",
    ];
    const aoa = [headers];
    orden.forEach((r) => {
      aoa.push([
        r.contraparte, r.nif, r.numero, r.fecha,
        num(r.base), num(r.cuotaIva), num(r.total), num(r.tipoIva), cuentaLetra(r.cuenta),
        r.concepto || conceptoPorCuenta(r.cuenta),
        Number(r.cuenta) === 10 ? r.clave : "",
        num(r.tipoRet) || 0, num(r.cuotaRet) || 0, num(r.tipoRe) || 0, num(r.cuotaRe) || 0,
      ]);
    });
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const range = XLSX.utils.decode_range(ws["!ref"]);
    const colsMoneda = [4, 5, 6, 12, 14]; // base, cuota IVA, total, cuota ret, cuota RE
    for (let R = 1; R <= range.e.r; R++) {
      colsMoneda.forEach((Ccol) => {
        const cell = ws[XLSX.utils.encode_cell({ r: R, c: Ccol })];
        if (cell && cell.t === "n") cell.z = "#,##0.00";
      });
    }
    ws["!cols"] = [
      { wch: 32 }, { wch: 12 }, { wch: 14 }, { wch: 11 },
      { wch: 14 }, { wch: 12 }, { wch: 12 }, { wch: 10 }, { wch: 8 }, { wch: 22 },
      { wch: 11 }, { wch: 13 }, { wch: 14 }, { wch: 8 }, { wch: 10 },
    ];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "FACTURAS");
    const limpio = (s) => String(s || "").toUpperCase().replace(/[^A-Z0-9]+/g, "_").replace(/^_|_$/g, "");
    const ahora = new Date();
    const sello = `${ahora.getFullYear()}-${String(ahora.getMonth() + 1).padStart(2, "0")}-${String(ahora.getDate()).padStart(2, "0")}_${String(ahora.getHours()).padStart(2, "0")}${String(ahora.getMinutes()).padStart(2, "0")}`;
    const cli = limpio(empresa.nombre) || "CLIENTE";
    const nombre = `MONITOR_${cli}_${trimestreSel}T${anioSel}_${sello}.xlsx`;
    XLSX.writeFile(wb, nombre);
  };

  /* ----- Exportación al formato del Conversor (doble partida) ----- */
  const exportarSL = (tipo) => {
    const esCompra = tipo === "compra";
    const filas = rowsSL.filter((r) => (esCompra ? r.sentido !== "venta" : r.sentido === "venta"));
    if (!filas.length) { setSlMsg(`No hay ${esCompra ? "compras" : "ventas"} en la tabla para exportar.`); return; }
    const ent = CONTABILIDADES_SL.find((e) => e.id === contabSelId);
    const orden = [...filas].sort((a, b) => {
      const f = fechaKey(a.fechaFactura).localeCompare(fechaKey(b.fechaFactura));
      if (f !== 0) return f;
      return String(a.numero).localeCompare(String(b.numero), "es");
    });
    const headers = esCompra ? SL_HEADERS_COMPRAS : SL_HEADERS_VENTAS;
    const aoa = [headers];
    orden.forEach((r) => {
      const base = num(r.base);
      // Banco a la manera de Monitor (datos de pago/cobro):
      //   compra → DEBE proveedor / HABER banco · venta → DEBE banco / HABER cliente
      const debe = esCompra ? (r.subCP || "") : (r.subBanco || "");
      const haber = esCompra ? (r.subBanco || "") : (r.subCP || "");
      aoa.push([
        r.fechaAsiento, r.fechaFactura, r.numero,
        CONCEPTOS_SL.includes(r.concepto) ? r.concepto : (esCompra ? "COMPRAS" : "VENTAS"),
        r.subCP || "", r.nif, r.contraparte,
        "", "", "", "",                                  // H domicilio, I localidad, J provincia, K c.p.
        num(r.base), num(r.tipoIva), num(r.cuotaIva), r.subIva || "",
        "", "", "",                                      // P/Q/R recargo (no automatizado)
        "", "", "",                                      // S/T/U IRPF/retención (no automatizado)
        "",                                              // V rectificativa
        r.subGI || "", isFinite(base) ? base : "",       // W subcuenta gasto/ingreso, X importe gasto/ingreso
        debe, haber, num(r.total),                       // Y DEBE, Z HABER, AA TOTAL
      ]);
    });
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const range = XLSX.utils.decode_range(ws["!ref"]);
    const colsMoneda = [11, 13, 23, 26]; // BASE, CUOTA IVA, IMPORTE GASTO, TOTAL
    for (let R = 1; R <= range.e.r; R++) {
      colsMoneda.forEach((Ccol) => {
        const cell = ws[XLSX.utils.encode_cell({ r: R, c: Ccol })];
        if (cell && cell.t === "n") cell.z = "#,##0.00";
      });
    }
    ws["!cols"] = headers.map((h) => ({ wch: Math.max(9, Math.min(20, h.length + 1)) }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, esCompra ? "COMPRAS" : "VENTAS");
    const limpio = (s) => String(s || "").toUpperCase().replace(/[^A-Z0-9]+/g, "_").replace(/^_|_$/g, "");
    const ahora = new Date();
    const sello = `${ahora.getFullYear()}-${String(ahora.getMonth() + 1).padStart(2, "0")}-${String(ahora.getDate()).padStart(2, "0")}_${String(ahora.getHours()).padStart(2, "0")}${String(ahora.getMinutes()).padStart(2, "0")}`;
    const cli = limpio((ent && (ent.centro ? ent.nombre + " " + ent.centro : ent.nombre)) || "SL");
    XLSX.writeFile(wb, `CONVERSOR_${esCompra ? "COMPRAS" : "VENTAS"}_${cli}_${trimestreSel}T${anioSel}_${sello}.xlsx`);
  };

  /* ----- Estilos reutilizables ----- */
  const inputBase = {
    border: `1px solid ${C.linea}`, borderRadius: 6, padding: "6px 8px",
    fontSize: 13, color: C.tinta, background: C.papel, outline: "none", width: "100%",
  };
  const inputNum = { ...inputBase, fontFamily: MONO, textAlign: "right" };
  const th = {
    padding: "8px 6px", fontSize: 10.5, fontWeight: 700, letterSpacing: "0.06em",
    color: C.vino, textTransform: "uppercase", whiteSpace: "nowrap", textAlign: "left",
    borderBottom: `2px solid ${C.vino}`, background: C.crema, position: "sticky", top: 0, zIndex: 1,
  };
  const btn = (primario) => ({
    display: "inline-flex", alignItems: "center", gap: 8, padding: "10px 18px",
    borderRadius: 8, fontSize: 14, fontWeight: 600, cursor: "pointer",
    border: primario ? "none" : `1.5px solid ${C.vino}`,
    background: primario ? C.vino : "transparent", color: primario ? C.crema : C.vino,
  });

  const estadoIcono = (s) => {
    if (s === "procesando") return <Loader2 size={16} className="animate-spin" style={{ color: C.vino }} />;
    if (s === "ok") return <CheckCircle2 size={16} style={{ color: C.ok }} />;
    if (s === "error") return <XCircle size={16} style={{ color: C.err }} />;
    return <FileText size={16} style={{ color: C.gris }} />;
  };

  const pendientes = files.filter((f) => f.status === "pendiente").length;

  /* ============== Pantalla de acceso ============== */
  if (!autenticado) {
    return (
      <div style={{ minHeight: "100vh", background: C.crema, display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "system-ui, -apple-system, 'Segoe UI', sans-serif" }}>
        <div style={{ background: C.papel, border: `1px solid ${C.linea}`, borderTop: `5px solid ${C.vino}`, borderRadius: 14, padding: "36px 32px", width: 380, maxWidth: "92vw", textAlign: "center" }}>
          <div style={{ color: C.vino, fontSize: 34, fontWeight: 800, letterSpacing: "0.18em" }}>ASEMA</div>
          <div style={{ color: C.gris, fontSize: 11, fontWeight: 600, letterSpacing: "0.42em", marginTop: 4 }}>ADVISORY</div>
          <div style={{ marginTop: 18, fontSize: 14, color: C.tinta }}>Facturas → Excel Monitor · Herramienta interna del despacho</div>
          <div style={{ marginTop: 24, textAlign: "left" }}>
            <label style={{ fontSize: 11, fontWeight: 700, color: C.gris, letterSpacing: "0.05em" }}>CLAVE DEL DESPACHO</label>
            <div style={{ display: "flex", gap: 8, marginTop: 6 }}>
              <input
                type="password"
                style={{ border: `1px solid ${C.linea}`, borderRadius: 8, padding: "10px 12px", fontSize: 14, flex: 1, outline: "none", fontFamily: MONO }}
                value={claveInput}
                onChange={(e) => setClaveInput(e.target.value)}
                onKeyDown={(e) => e.key === "Enter" && entrar()}
                placeholder="••••••••"
                autoFocus
              />
              <button style={btn(true)} onClick={entrar}>
                <Lock size={15} /> Entrar
              </button>
            </div>
            {avisoClave && (
              <div style={{ marginTop: 10, fontSize: 12.5, color: C.err, display: "flex", gap: 6, alignItems: "center" }}>
                <AlertTriangle size={14} /> {avisoClave}
              </div>
            )}
            <div style={{ marginTop: 14, fontSize: 11.5, color: C.gris }}>
              La clave la gestiona Rafael (variable ASEMA_PASSWORD en Vercel). Se comprueba al procesar la primera factura.
            </div>
          </div>
        </div>
      </div>
    );
  }

  /* ============================ APP ============================ */
  return (
    <div style={{ minHeight: "100vh", background: C.crema, color: C.tinta, fontFamily: "system-ui, -apple-system, 'Segoe UI', sans-serif" }}>

      {/* ---------- Cabecera de marca ---------- */}
      <header style={{ background: C.vino, borderBottom: `4px solid ${C.vinoOscuro}` }}>
        <div className="max-w-7xl mx-auto px-6 py-5 flex items-end justify-between flex-wrap gap-3">
          <div>
            <div style={{ color: C.crema, fontSize: 30, fontWeight: 800, letterSpacing: "0.18em", lineHeight: 1 }}>
              ASEMA
            </div>
            <div style={{ color: C.cremaOscura, fontSize: 11, fontWeight: 600, letterSpacing: "0.42em", marginTop: 4 }}>
              ADVISORY
            </div>
          </div>
          <div style={{ textAlign: "right" }}>
            <div style={{ color: C.crema, fontSize: 15, fontWeight: 700 }}>Facturas → Excel Monitor</div>
            <div style={{ color: C.cremaOscura, fontSize: 12, display: "flex", gap: 12, alignItems: "center", justifyContent: "flex-end" }}>
              <span>EOS · Autónomos · Estimación directa simplificada</span>
              <button onClick={salir} title="Cambiar la clave del despacho" style={{ border: "none", background: "transparent", color: C.cremaOscura, cursor: "pointer", display: "inline-flex", alignItems: "center", gap: 4, fontSize: 11, padding: 0 }}>
                <LogOut size={12} /> Cambiar clave
              </button>
            </div>
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-6 py-6 flex flex-col gap-6">

        {/* ---------- Selector de modo ---------- */}
        <div style={{ display: "flex", gap: 8 }}>
          {[["facturas", "Facturas escaneadas (IA)"], ["listados", "Listados Excel (clientes)"], ["sl", "SL (doble partida)"]].map(([id, txt]) => (
            <button
              key={id}
              onClick={() => setModo(id)}
              style={{
                padding: "9px 18px", borderRadius: 8, fontSize: 14, fontWeight: 700, cursor: "pointer",
                border: `1.5px solid ${C.vino}`,
                background: modo === id ? C.vino : "transparent",
                color: modo === id ? C.crema : C.vino,
              }}
            >
              {txt}
            </button>
          ))}
        </div>

        {/* ---------- Periodo global (trimestre + año) · referencia de todas las pestañas ---------- */}
        <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap", background: C.papel, border: `1px solid ${C.linea}`, borderLeft: `4px solid ${C.vino}`, borderRadius: 10, padding: "10px 14px" }}>
          <CalendarCheck size={16} style={{ color: C.vino }} />
          <span style={{ fontSize: 11, fontWeight: 800, color: C.vino, letterSpacing: "0.05em" }}>PERIODO A CONTABILIZAR</span>
          <select value={trimestreSel} onChange={(e) => setTrimestreSel(Number(e.target.value))} style={{ ...inputBase, width: "auto", fontWeight: 700, borderColor: C.vino }}>
            <option value={1}>1T · ene-mar</option>
            <option value={2}>2T · abr-jun</option>
            <option value={3}>3T · jul-sep</option>
            <option value={4}>4T · oct-dic</option>
          </select>
          <select value={anioSel} onChange={(e) => setAnioSel(Number(e.target.value))} style={{ ...inputBase, width: "auto", fontWeight: 700, borderColor: C.vino, fontFamily: MONO }}>
            {Array.from({ length: 6 }, (_, i) => new Date().getFullYear() - 3 + i).map((y) => (
              <option key={y} value={y}>{y}</option>
            ))}
          </select>
          <span style={{ fontSize: 12, color: C.gris }}>
            Referencia para las tres pestañas. Las facturas fuera de <b style={{ color: C.tinta }}>{trimestreSel}T {anioSel}</b> ({rangoTrimestre(trimestreSel, anioSel)}) se marcan en ámbar.
          </span>
        </div>


        {/* ---------- SL · Contabilidad (solo modo sl) ---------- */}
        {modo === "sl" && (() => {
          const entSL = CONTABILIDADES_SL.find((e) => e.id === contabSelId);
          const nCtas = entSL ? Object.keys(entSL.gasto || {}).length + Object.keys(entSL.ingreso || {}).length : 0;
          return (
          <section style={{ background: C.papel, border: `1px solid ${C.linea}`, borderRadius: 12, padding: 20 }}>
            <div className="flex items-center gap-2 mb-4">
              <Building2 size={18} style={{ color: C.vino }} />
              <h2 style={{ fontSize: 15, fontWeight: 800, letterSpacing: "0.04em" }}>1 · CONTABILIDAD (DOBLE PARTIDA)</h2>
              <span style={{ fontSize: 12, color: C.gris }}>— sociedad que se lleva por partida doble; se importa con el Conversor de Monitor</span>
            </div>
            <div className="grid grid-cols-1 md:grid-cols-12 gap-3 items-end">
              <div className="md:col-span-5">
                <label style={{ fontSize: 11, fontWeight: 700, color: C.gris, letterSpacing: "0.05em" }}>CONTABILIDAD</label>
                <select style={{ ...inputBase, fontWeight: 600 }} value={contabSelId} onChange={(e) => setContabSelId(e.target.value)}>
                  {CONTABILIDADES_SL.map((e) => (
                    <option key={e.id} value={e.id}>{e.centro ? `${e.nombre} · ${e.centro}` : e.nombre}</option>
                  ))}
                </select>
              </div>
              <div className="md:col-span-4">
                <label style={{ fontSize: 11, fontWeight: 700, color: C.vino, letterSpacing: "0.05em" }}>BANCO · cobros/pagos (572)</label>
                <select style={{ ...inputBase, fontFamily: MONO, fontWeight: 600, borderColor: C.vino }} value={bancoSel} onChange={(e) => { setBancoSel(e.target.value); guardarBanco(entSL.id, e.target.value); }}>
                  {Object.entries(entSL.bancos || { "57200000000": "Banco c/c (euros)" }).map(([c, n]) => (
                    <option key={c} value={c}>{c} · {n}</option>
                  ))}
                </select>
              </div>
              <div className="md:col-span-3" style={{ fontSize: 12, color: C.gris, lineHeight: 1.5 }}>
                {entSL && (
                  <span>
                    Listado: <b style={{ color: C.tinta }}>{Object.keys(entSL.proveedores || {}).length}</b> proveedores ·{" "}
                    <b style={{ color: C.tinta }}>{Object.keys(entSL.clientes || {}).length}</b> clientes ·{" "}
                    <b style={{ color: C.tinta }}>{nCtas}</b> cuentas de gasto/ingreso.
                    {entSL.soloIngresos && <span style={{ color: C.vino, fontWeight: 700 }}> · Solo ingresos: todo se registra como VENTA.</span>}
                    {!entSL.nif && <span style={{ color: C.warn }}> Falta el NIF de la sociedad: añádelo para afinar emitida/recibida.</span>}
                  </span>
                )}
              </div>
            </div>
          </section>
          );
        })()}

        {/* ---------- 1 · Cliente (modos facturas/listados) ---------- */}
        {modo !== "sl" && (
        <section style={{ background: C.papel, border: `1px solid ${C.linea}`, borderRadius: 12, padding: 20 }}>
          <div className="flex items-center gap-2 mb-4">
            <Building2 size={18} style={{ color: C.vino }} />
            <h2 style={{ fontSize: 15, fontWeight: 800, letterSpacing: "0.04em" }}>1 · CLIENTE DEL DESPACHO</h2>
            <span style={{ fontSize: 12, color: C.gris }}>— titular de la contabilidad; imprescindible para distinguir emitidas de recibidas</span>
          </div>
          <div className="grid grid-cols-1 md:grid-cols-12 gap-3 items-end">
            <div className="md:col-span-5">
              <label style={{ fontSize: 11, fontWeight: 700, color: C.gris, letterSpacing: "0.05em" }}>NOMBRE / RAZÓN SOCIAL</label>
              <input
                style={inputBase}
                value={empresa.nombre}
                onChange={(e) => setEmpresa({ ...empresa, nombre: e.target.value })}
                placeholder="PEPITO PÉREZ PÉREZ"
              />
            </div>
            <div className="md:col-span-3">
              <label style={{ fontSize: 11, fontWeight: 700, color: C.gris, letterSpacing: "0.05em" }}>NIF</label>
              <input
                style={{ ...inputBase, fontFamily: MONO }}
                value={empresa.nif}
                onChange={(e) => setEmpresa({ ...empresa, nif: e.target.value.toUpperCase() })}
                placeholder="31XXXXXXX"
              />
            </div>
            <div className="md:col-span-4">
              <button style={{ ...btn(false), width: "100%", justifyContent: "center", padding: "9px 10px" }} onClick={guardarEmpresa}>
                <Save size={15} /> Guardar cliente
              </button>
            </div>
          </div>
          {empresa.nif && !validNif(empresa.nif) && (
            <div style={{ marginTop: 8, fontSize: 12, color: C.warn, display: "flex", alignItems: "center", gap: 6 }}>
              <AlertTriangle size={14} /> El NIF del cliente no pasa la validación de letra de control. Revísalo: la clasificación depende de él.
            </div>
          )}
          {guardadas.length > 0 && (
            <div className="flex flex-wrap gap-2 mt-3">
              {guardadas.map((g) => (
                <button
                  key={g.nif}
                  onClick={() => setEmpresa(g)}
                  style={{
                    fontSize: 12, padding: "5px 12px", borderRadius: 999, cursor: "pointer",
                    border: `1px solid ${empresa.nif === g.nif ? C.vino : C.linea}`,
                    background: empresa.nif === g.nif ? C.vino : C.crema,
                    color: empresa.nif === g.nif ? C.crema : C.tinta, fontWeight: 600,
                  }}
                >
                  {g.nombre}
                </button>
              ))}
            </div>
          )}
        </section>
        )}

        {/* ---------- 2 · Listado de ingresos (SL · solo ingresos, p.ej. Matilde) ---------- */}
        {modo === "sl" && entSLSel?.soloIngresos && (
        <section style={{ background: C.papel, border: `1px solid ${C.linea}`, borderRadius: 12, padding: 20 }}>
          <div className="flex items-center gap-2 mb-4">
            <FileSpreadsheet size={18} style={{ color: C.vino }} />
            <h2 style={{ fontSize: 15, fontWeight: 800, letterSpacing: "0.04em" }}>2 · LISTADO DE INGRESOS (PDF)</h2>
            <span style={{ fontSize: 12, color: C.gris }}>— se lee con IA; cada línea es una venta (concepto VENTAS)</span>
          </div>
          <div className="grid grid-cols-1 md:grid-cols-12 gap-3 items-end">
            <div className="md:col-span-7">
              <label style={{ fontSize: 11, fontWeight: 700, color: C.gris, letterSpacing: "0.05em", display: "block" }}>ARCHIVO PDF DEL LISTADO DE INGRESOS</label>
              <input type="file" accept=".pdf" onChange={(e) => { importarListadoMatilde(e.target.files[0]); e.target.value = ""; }} style={{ ...inputBase, padding: "5px 8px", cursor: "pointer" }} />
            </div>
          </div>
          {slMsg && (
            <div style={{ marginTop: 12, padding: "10px 14px", borderRadius: 8, background: C.crema, color: C.tinta, fontSize: 13, display: "flex", gap: 8, alignItems: "center" }}>
              <FileText size={16} style={{ color: C.vino }} /> {slMsg}
            </div>
          )}
          <div style={{ marginTop: 12, fontSize: 12, color: C.gris, lineHeight: 1.6 }}>
            Sube el PDF de «Facturas emitidas» de la administración. Si el listado trae la columna del <b>CIF</b> de cada cliente, el NIF se rellena solo (columna B del Conversor); si aún no la trae, los clientes se emparejan <b>por nombre</b> y el NIF queda en blanco para que lo completes. Revisa siempre antes de exportar.
          </div>
        </section>
        )}

        {/* ---------- 2 · Facturas (facturas, y SL salvo solo-ingresos) ---------- */}
        {(modo === "facturas" || (modo === "sl" && !entSLSel?.soloIngresos)) && (
        <section style={{ background: C.papel, border: `1px solid ${C.linea}`, borderRadius: 12, padding: 20 }}>
          <div className="flex items-center gap-2 mb-4">
            <Upload size={18} style={{ color: C.vino }} />
            <h2 style={{ fontSize: 15, fontWeight: 800, letterSpacing: "0.04em" }}>2 · FACTURAS ESCANEADAS</h2>
            <span style={{ fontSize: 12, color: C.gris }}>— PDF, JPG o PNG · varios archivos a la vez · puede haber varias facturas por PDF</span>
          </div>

          <div
            onDragOver={(e) => { e.preventDefault(); setArrastrando(true); }}
            onDragLeave={() => setArrastrando(false)}
            onDrop={(e) => { e.preventDefault(); setArrastrando(false); addFiles(e.dataTransfer.files); }}
            onClick={() => inputRef.current && inputRef.current.click()}
            style={{
              border: `2px dashed ${arrastrando ? C.vino : C.linea}`,
              background: arrastrando ? C.crema : "transparent",
              borderRadius: 10, padding: "28px 16px", textAlign: "center", cursor: "pointer",
            }}
          >
            <Upload size={26} style={{ color: C.vino, margin: "0 auto 8px" }} />
            <div style={{ fontSize: 14, fontWeight: 600 }}>Arrastra aquí las facturas o haz clic para seleccionarlas</div>
            <div style={{ fontSize: 12, color: C.gris, marginTop: 4 }}>Máximo {MAX_MB} MB por archivo · escanea a 150-200 ppp en gris para mejores resultados</div>
            <input
              ref={inputRef}
              type="file"
              multiple
              accept=".pdf,.jpg,.jpeg,.png,.webp,.docx,.doc,application/pdf,image/jpeg,image/png,image/webp,application/vnd.openxmlformats-officedocument.wordprocessingml.document,application/msword"
              style={{ display: "none" }}
              onChange={(e) => { addFiles(e.target.files); e.target.value = ""; }}
            />
          </div>

          {files.length > 0 && (
            <div className="mt-4 flex flex-col gap-1">
              {files.map((f) => (
                <div key={f.id} className="flex items-center gap-3 flex-wrap" style={{ padding: "7px 10px", borderRadius: 8, background: f.status === "error" ? C.errBg : C.crema }}>
                  {estadoIcono(f.status)}
                  <span style={{ fontSize: 13, fontWeight: 600, flex: 1, minWidth: 180, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{f.name}</span>
                  <span style={{ fontSize: 11, color: C.gris, fontFamily: MONO }}>{(f.size / 1024 / 1024).toFixed(1)} MB</span>
                  {f.status === "ok" && <span style={{ fontSize: 12, color: C.ok, fontWeight: 700 }}>{f.nFacturas} factura{f.nFacturas !== 1 ? "s" : ""}</span>}
                  {f.status === "error" && <span style={{ fontSize: 12, color: C.err }}>{f.msg}</span>}
                  {f.status === "procesando" && <span style={{ fontSize: 12, color: C.vino }}>{f.progreso ? `leyendo ${f.progreso}…` : "leyendo…"}</span>}
                  {f.status !== "procesando" && (
                    <button onClick={() => quitarFile(f.id)} style={{ border: "none", background: "transparent", cursor: "pointer", color: C.gris, padding: 2 }} aria-label="Quitar archivo">
                      <Trash2 size={15} />
                    </button>
                  )}
                </div>
              ))}
            </div>
          )}

          {aviso && (
            <div style={{ marginTop: 12, padding: "10px 14px", borderRadius: 8, background: C.warnBg, color: C.warn, fontSize: 13, display: "flex", gap: 8, alignItems: "center" }}>
              <AlertTriangle size={16} /> {aviso}
            </div>
          )}

          <div className="mt-4 flex gap-3 flex-wrap">
            <button
              style={{ ...btn(true), opacity: procesando || !files.some((f) => f.status === "pendiente" || f.status === "error") ? 0.5 : 1 }}
              disabled={procesando || !files.some((f) => f.status === "pendiente" || f.status === "error")}
              onClick={procesarPendientes}
            >
              {procesando ? <Loader2 size={16} className="animate-spin" /> : <Calculator size={16} />}
              {procesando ? "Procesando…" : `Procesar ${pendientes || ""} factura${pendientes === 1 ? "" : "s"} con IA`}
            </button>
            {modo !== "sl" && (
              <button style={btn(false)} onClick={filaManual}>
                <Plus size={16} /> Añadir apunte manual
              </button>
            )}
          </div>
        </section>

)}

        {/* ---------- 2b · Listados Excel (solo modo listados) ---------- */}
        {modo === "listados" && (
        <section style={{ background: C.papel, border: `1px solid ${C.linea}`, borderRadius: 12, padding: 20 }}>
          <div className="flex items-center gap-2 mb-4">
            <FileSpreadsheet size={18} style={{ color: C.vino }} />
            <h2 style={{ fontSize: 15, fontWeight: 800, letterSpacing: "0.04em" }}>2 · LISTADO EXCEL DEL CLIENTE</h2>
            <span style={{ fontSize: 12, color: C.gris }}>— para clientes que envían su contabilidad ya en Excel, no facturas sueltas</span>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-12 gap-3 items-end mb-4">
            <div className="md:col-span-7">
              <label style={{ fontSize: 11, fontWeight: 700, color: C.gris, letterSpacing: "0.05em" }}>PLANTILLA (formato del archivo)</label>
              <select style={{ ...inputBase, fontWeight: 600 }} value={plantillaSel} onChange={(e) => setPlantillaSel(e.target.value)}>
                {PLANTILLAS.map((p) => (<option key={p.id} value={p.id}>{p.nombre}</option>))}
              </select>
              {(PLANTILLAS.find((p) => p.id === plantillaSel) || {}).trimestral && (
                <div style={{ fontSize: 11.5, color: C.vino, fontWeight: 600, marginTop: 5 }}>
                  Se importará solo el <b>{trimestreSel}º trimestre</b> (según el selector de periodo de arriba).
                </div>
              )}
            </div>
            <div className="md:col-span-5">
              <label style={{ fontSize: 11, fontWeight: 700, color: C.gris, letterSpacing: "0.05em", display: "block" }}>
                {(PLANTILLAS.find((p) => p.id === plantillaSel) || {}).tipo === "pdf" ? "ARCHIVO PDF" : "ARCHIVO EXCEL (.xlsx / .xls)"}
              </label>
              <input
                type="file"
                accept={(PLANTILLAS.find((p) => p.id === plantillaSel) || {}).tipo === "pdf" ? ".pdf" : ".xlsx,.xls"}
                onChange={(e) => { importarListado(e.target.files[0]); e.target.value = ""; }}
                style={{ ...inputBase, padding: "5px 8px", cursor: "pointer" }}
              />
            </div>
          </div>

          {listadoMsg && (
            <div style={{ padding: "10px 14px", borderRadius: 8, background: C.crema, color: C.tinta, fontSize: 13, display: "flex", gap: 8, alignItems: "center" }}>
              <FileText size={16} style={{ color: C.vino }} /> {listadoMsg}
            </div>
          )}

          <div style={{ marginTop: 14, fontSize: 12, color: C.gris, lineHeight: 1.6 }}>
            Para Ferretería El Paso (que acumula todo el año en el mismo archivo), se importan solo las líneas del <b>trimestre seleccionado arriba</b> (el selector de periodo global). Para Aaron no se filtra porque ya envía cada trimestre por separado. Las ventas se asignan a la cuenta <b>V</b> y las compras a la cuenta <b>C</b>. Los abonos (importes negativos) se respetan con su signo; las facturas con varios tipos de IVA se separan en una línea por tipo. El NIF se normaliza solo (quita ES, guiones y puntos). El listado PDF de Aaron se lee con IA (extrae ventas y compras en una pasada) y, como aún no trae NIF, esas líneas saldrán con NIF 0 y aviso ámbar: es normal. Revisa siempre el resultado antes de exportar. Consejo: rellena arriba el nombre del cliente para que el Excel se descargue con su nombre y la fecha, no como "CLIENTE".
          </div>
        </section>
        )}

                {/* ---------- 3 · Revisión (facturas/listados) ---------- */}
        {modo !== "sl" && rows.length > 0 && (
          <section style={{ background: C.papel, border: `1px solid ${C.linea}`, borderRadius: 12, padding: 20 }}>
            <div className="flex items-center gap-2 mb-1 flex-wrap">
              <FileSpreadsheet size={18} style={{ color: C.vino }} />
              <h2 style={{ fontSize: 15, fontWeight: 800, letterSpacing: "0.04em" }}>3 · REVISIÓN ANTES DE EXPORTAR</h2>
              <span style={{ fontSize: 12, color: C.gris }}>— {rows.length} apunte{rows.length !== 1 ? "s" : ""}</span>
              <div className="flex gap-2 ml-auto">
                {nErr > 0 && (
                  <span style={{ fontSize: 12, fontWeight: 700, color: C.err, background: C.errBg, padding: "4px 10px", borderRadius: 999 }}>
                    {nErr} con errores
                  </span>
                )}
                {nWarn > 0 && (
                  <span style={{ fontSize: 12, fontWeight: 700, color: C.warn, background: C.warnBg, padding: "4px 10px", borderRadius: 999 }}>
                    {nWarn} con avisos
                  </span>
                )}
                {nErr === 0 && nWarn === 0 && (
                  <span style={{ fontSize: 12, fontWeight: 700, color: C.ok, background: C.okBg, padding: "4px 10px", borderRadius: 999 }}>
                    Todo cuadra
                  </span>
                )}
              </div>
            </div>

            {/* Resumen por cuenta — espejo del panel CUENTAS de Monitor */}
            <div className="flex flex-wrap gap-2 my-4">
              {CUENTAS.filter((c) => resumen.porCuenta[c.n]).map((c) => (
                <div key={c.n} style={{ border: `1px solid ${C.linea}`, borderLeft: `4px solid ${C.vino}`, borderRadius: 8, padding: "6px 12px", background: C.crema }}>
                  <div style={{ fontSize: 10, fontWeight: 800, color: C.vino, letterSpacing: "0.08em" }}>
                    {String(c.n).padStart(2, "0")} {c.t}
                  </div>
                  <div style={{ fontSize: 13, fontFamily: MONO }}>
                    {resumen.porCuenta[c.n].n} apt. · {fmtES(resumen.porCuenta[c.n].base)} €
                  </div>
                </div>
              ))}
              <div style={{ marginLeft: "auto", textAlign: "right", fontSize: 12, color: C.gris }}>
                <div>IVA repercutido: <b style={{ fontFamily: MONO, color: C.tinta }}>{fmtES(resumen.ivaRep)} €</b></div>
                <div>IVA soportado: <b style={{ fontFamily: MONO, color: C.tinta }}>{fmtES(resumen.ivaSop)} €</b></div>
              </div>
            </div>

            {/* Tabla editable */}
            <div className="overflow-x-auto" style={{ border: `1px solid ${C.linea}`, borderRadius: 8 }}>
              <table style={{ borderCollapse: "collapse", width: "100%", minWidth: 1580 }}>
                <thead>
                  <tr>
                    <th style={{ ...th, width: 36 }}></th>
                    <th style={{ ...th, minWidth: 190 }}>Cliente / Proveedor</th>
                    <th style={{ ...th, width: 110 }}>NIF</th>
                    <th style={{ ...th, width: 110 }}>Nº factura</th>
                    <th style={{ ...th, width: 100 }}>Fecha</th>
                    <th style={{ ...th, width: 100, textAlign: "right" }}>Base</th>
                    <th style={{ ...th, width: 64, textAlign: "right" }}>% IVA</th>
                    <th style={{ ...th, width: 95, textAlign: "right" }}>Cuota IVA</th>
                    <th style={{ ...th, width: 60, textAlign: "right" }}>% RE</th>
                    <th style={{ ...th, width: 85, textAlign: "right" }}>Cuota RE</th>
                    <th style={{ ...th, width: 60, textAlign: "right" }}>% Ret.</th>
                    <th style={{ ...th, width: 85, textAlign: "right" }}>Cuota Ret.</th>
                    <th style={{ ...th, width: 100, textAlign: "right" }}>Total</th>
                    <th style={{ ...th, width: 130 }}>Cuenta</th>
                    <th style={{ ...th, width: 160 }}>Concepto</th>
                    <th style={{ ...th, width: 175 }}>Clave gasto</th>
                    <th style={{ ...th, width: 92, right: 0, zIndex: 2, borderLeft: `1px solid ${C.linea}` }}>Acciones</th>
                  </tr>
                </thead>
                <tbody>
                  {rows.map((r, idx) => {
                    const iss = issuesById[r.id] || [];
                    const hayErr = iss.some((i) => i.lv === "err");
                    const hayWarn = iss.some((i) => i.lv === "warn");
                    const fondo = hayErr ? C.errBg : hayWarn ? C.warnBg : idx % 2 ? "#FCFAF4" : C.papel;
                    const fueraTrim = fechaValida(r.fecha) && !fechaEnTrimestre(r.fecha, trimestreSel, anioSel);
                    return (
                      <React.Fragment key={r.id}>
                        <tr style={{ background: fondo }}>
                          <td style={{ padding: "4px 6px", textAlign: "center" }}>
                            {hayErr ? <XCircle size={15} style={{ color: C.err }} />
                              : hayWarn ? <AlertTriangle size={15} style={{ color: C.warn }} />
                              : <CheckCircle2 size={15} style={{ color: C.ok }} />}
                          </td>
                          <td style={{ padding: 3 }}>
                            <input style={inputBase} value={r.contraparte} onChange={(e) => upd(r.id, "contraparte", e.target.value.toUpperCase())} />
                          </td>
                          <td style={{ padding: 3 }}>
                            <input style={{ ...inputBase, fontFamily: MONO }} value={r.nif} onChange={(e) => upd(r.id, "nif", e.target.value.toUpperCase().replace(/[\s\-\.]/g, ""))} />
                          </td>
                          <td style={{ padding: 3 }}>
                            <input style={{ ...inputBase, fontFamily: MONO }} value={r.numero} onChange={(e) => upd(r.id, "numero", e.target.value)} />
                          </td>
                          <td style={{ padding: 3 }}>
                            <input style={{ ...inputBase, fontFamily: MONO }} value={r.fecha} onChange={(e) => upd(r.id, "fecha", e.target.value)} onBlur={(e) => upd(r.id, "fecha", normFecha(e.target.value))} placeholder="dd/mm/aaaa" />
                          </td>
                          <td style={{ padding: 3 }}><input style={inputNum} value={r.base} onChange={(e) => upd(r.id, "base", e.target.value)} /></td>
                          <td style={{ padding: 3 }}><input style={inputNum} value={r.tipoIva} onChange={(e) => upd(r.id, "tipoIva", e.target.value)} /></td>
                          <td style={{ padding: 3 }}><input style={inputNum} value={r.cuotaIva} onChange={(e) => upd(r.id, "cuotaIva", e.target.value)} /></td>
                          <td style={{ padding: 3 }}><input style={inputNum} value={r.tipoRe} onChange={(e) => upd(r.id, "tipoRe", e.target.value)} /></td>
                          <td style={{ padding: 3 }}><input style={inputNum} value={r.cuotaRe} onChange={(e) => upd(r.id, "cuotaRe", e.target.value)} /></td>
                          <td style={{ padding: 3 }}><input style={inputNum} value={r.tipoRet} onChange={(e) => upd(r.id, "tipoRet", e.target.value)} /></td>
                          <td style={{ padding: 3 }}><input style={inputNum} value={r.cuotaRet} onChange={(e) => upd(r.id, "cuotaRet", e.target.value)} /></td>
                          <td style={{ padding: 3 }}><input style={{ ...inputNum, fontWeight: 700 }} value={r.total} onChange={(e) => upd(r.id, "total", e.target.value)} /></td>
                          <td style={{ padding: 3 }}>
                            <select style={{ ...inputBase, fontWeight: 600 }} value={r.cuenta} onChange={(e) => upd(r.id, "cuenta", Number(e.target.value))}>
                              {CUENTAS.map((c) => (
                                <option key={c.n} value={c.n}>{c.n} · {c.t}</option>
                              ))}
                            </select>
                          </td>
                          <td style={{ padding: 3 }}>
                            <select style={{ ...inputBase, fontWeight: 600 }} value={CONCEPTOS_SL.includes(r.concepto) ? r.concepto : "GASTOS"} onChange={(e) => upd(r.id, "concepto", e.target.value)}>
                              {CONCEPTOS_SL.map((c) => (<option key={c} value={c}>{c}</option>))}
                            </select>
                          </td>
                          <td style={{ padding: 3 }}>
                            <select
                              style={{ ...inputBase, opacity: Number(r.cuenta) === 10 ? 1 : 0.35 }}
                              disabled={Number(r.cuenta) !== 10}
                              value={r.clave}
                              onChange={(e) => upd(r.id, "clave", e.target.value)}
                            >
                              {CLAVES.map((c) => (
                                <option key={c.k} value={c.k}>{c.t}</option>
                              ))}
                            </select>
                          </td>
                          <td style={{ padding: 3, whiteSpace: "nowrap", position: "sticky", right: 0, background: fondo, zIndex: 1, borderLeft: `1px solid ${C.linea}` }}>
                            {fueraTrim && (
                              <button onClick={() => upd(r.id, "fecha", primerDiaTrimestre(trimestreSel, anioSel))} title={`Ajustar al ${trimestreSel}T ${anioSel} (${primerDiaTrimestre(trimestreSel, anioSel)})`} style={{ border: "none", background: "transparent", cursor: "pointer", color: C.warn, padding: 4 }} aria-label="Ajustar fecha al trimestre">
                                <CalendarCheck size={15} />
                              </button>
                            )}
                            <button onClick={() => recalcular(r.id)} title="Recalcular cuotas y total desde la base" style={{ border: "none", background: "transparent", cursor: "pointer", color: C.vino, padding: 4 }} aria-label="Recalcular">
                              <Calculator size={15} />
                            </button>
                            <button onClick={() => borrarFila(r.id)} title="Eliminar este apunte" style={{ border: "none", background: "transparent", cursor: "pointer", color: C.err, padding: 4 }} aria-label="Eliminar apunte">
                              <Trash2 size={16} />
                            </button>
                          </td>
                        </tr>
                        {iss.length > 0 && (
                          <tr style={{ background: fondo }}>
                            <td></td>
                            <td colSpan={16} style={{ padding: "0 6px 7px", fontSize: 11.5, color: hayErr ? C.err : C.warn }}>
                              {iss.map((i, k) => (
                                <span key={k} style={{ marginRight: 14 }}>• {i.msg}</span>
                              ))}
                              <span style={{ color: C.gris }}>· origen: {r.fileName}</span>
                            </td>
                          </tr>
                        )}
                      </React.Fragment>
                    );
                  })}
                </tbody>
              </table>
            </div>

            {/* Barra de exportación */}
            <div className="mt-5 flex items-center gap-4 flex-wrap">
              <button style={{ ...btn(true), opacity: nErr > 0 ? 0.85 : 1 }} onClick={exportar}>
                <Download size={16} /> Exportar Excel para Monitor
              </button>
              {nErr > 0 && (
                <span style={{ fontSize: 12.5, color: C.err, display: "flex", alignItems: "center", gap: 6 }}>
                  <AlertTriangle size={15} /> Hay {nErr} apunte{nErr !== 1 ? "s" : ""} con errores: corrígelos antes de importar en Monitor o la importación fallará.
                </span>
              )}
              <span style={{ fontSize: 12, color: C.gris, marginLeft: "auto" }}>
                Columnas: NOMBRE · NIF · Nº FACTURA · FECHA · BASE · CANTIDAD IVA · TOTAL · TIPO IVA · CUENTA · CONCEPTO · CLAVE · RET. · RE
              </span>
            </div>
          </section>
        )}

        {/* ---------- 3 · Revisión SL (doble partida) ---------- */}
        {modo === "sl" && rowsSL.length > 0 && (
          <section style={{ background: C.papel, border: `1px solid ${C.linea}`, borderRadius: 12, padding: 20 }}>
            <div className="flex items-center gap-2 mb-1 flex-wrap">
              <FileSpreadsheet size={18} style={{ color: C.vino }} />
              <h2 style={{ fontSize: 15, fontWeight: 800, letterSpacing: "0.04em" }}>3 · REVISIÓN — CONVERSOR (DOBLE PARTIDA)</h2>
              <span style={{ fontSize: 12, color: C.gris }}>— {rowsSL.length} línea{rowsSL.length !== 1 ? "s" : ""}</span>
              <div className="flex gap-2 ml-auto">
                {nErrSL > 0 && (<span style={{ fontSize: 12, fontWeight: 700, color: C.err, background: C.errBg, padding: "4px 10px", borderRadius: 999 }}>{nErrSL} con errores</span>)}
                {nWarnSL > 0 && (<span style={{ fontSize: 12, fontWeight: 700, color: C.warn, background: C.warnBg, padding: "4px 10px", borderRadius: 999 }}>{nWarnSL} con avisos</span>)}
                {nErrSL === 0 && nWarnSL === 0 && (<span style={{ fontSize: 12, fontWeight: 700, color: C.ok, background: C.okBg, padding: "4px 10px", borderRadius: 999 }}>Todo cuadra</span>)}
              </div>
            </div>
            <div style={{ fontSize: 12, color: C.gris, marginBottom: 12, lineHeight: 1.5 }}>
              Subcuentas <b style={{ color: C.warn }}>en ámbar</b>: o se han <b>emparejado por nombre</b> (confirma que es la correcta), o el cliente/proveedor es <b>● NUEVO</b> y no estaba en ningún listado, así que la app le ha creado una subcuenta correlativa a la última. Revísalas antes de exportar: el mismo NIF recibirá siempre esa subcuenta a partir de ahora.
            </div>

            <div className="overflow-x-auto" style={{ border: `1px solid ${C.linea}`, borderRadius: 8 }}>
              <table style={{ borderCollapse: "collapse", width: "100%", minWidth: 1680 }}>
                <thead>
                  <tr>
                    <th style={{ ...th, width: 36 }}></th>
                    <th style={{ ...th, width: 72 }}>Sentido</th>
                    <th style={{ ...th, minWidth: 180 }}>Nombre</th>
                    <th style={{ ...th, width: 105 }}>NIF</th>
                    <th style={{ ...th, width: 100 }}>Nº factura</th>
                    <th style={{ ...th, width: 100 }}>Fecha</th>
                    <th style={{ ...th, width: 95, textAlign: "right" }}>Base</th>
                    <th style={{ ...th, width: 56, textAlign: "right" }}>% IVA</th>
                    <th style={{ ...th, width: 92, textAlign: "right" }}>Importe IVA</th>
                    <th style={{ ...th, width: 95, textAlign: "right" }}>Total</th>
                    <th style={{ ...th, minWidth: 150 }}>Concepto</th>
                    <th style={{ ...th, width: 132 }}>Subcta. cliente/prov.</th>
                    <th style={{ ...th, width: 120 }}>Subcta. IVA</th>
                    <th style={{ ...th, width: 132 }}>Subcta. gasto/ingreso</th>
                    <th style={{ ...th, width: 120 }}>Subcta. banco</th>
                    <th style={{ ...th, width: 92, right: 0, zIndex: 2, borderLeft: `1px solid ${C.linea}` }}>Acciones</th>
                  </tr>
                </thead>
                <tbody>
                  {rowsSL.map((r, idx) => {
                    const iss = issuesSLById[r.id] || [];
                    const hayErr = iss.some((i) => i.lv === "err");
                    const hayWarn = iss.some((i) => i.lv === "warn");
                    const fondo = hayErr ? C.errBg : hayWarn ? C.warnBg : idx % 2 ? "#FCFAF4" : C.papel;
                    const subCPbg = (r.subCPNueva || r.subCPCreada) ? C.warnBg : C.papel;
                    const subGIbg = r.subGINueva ? C.warnBg : C.papel;
                    const fueraTrimSL = fechaValida(r.fechaFactura) && !fechaEnTrimestre(r.fechaFactura, trimestreSel, anioSel);
                    return (
                      <React.Fragment key={r.id}>
                        <tr style={{ background: fondo }}>
                          <td style={{ padding: "4px 6px", textAlign: "center" }}>
                            {hayErr ? <XCircle size={15} style={{ color: C.err }} /> : hayWarn ? <AlertTriangle size={15} style={{ color: C.warn }} /> : <CheckCircle2 size={15} style={{ color: C.ok }} />}
                          </td>
                          <td style={{ padding: 3 }}>
                            <select style={{ ...inputBase, fontWeight: 600 }} value={r.sentido} onChange={(e) => updSL(r.id, "sentido", e.target.value)}>
                              <option value="compra">Compra</option>
                              <option value="venta">Venta</option>
                            </select>
                          </td>
                          <td style={{ padding: 3 }}><input style={inputBase} value={r.contraparte} onChange={(e) => updSL(r.id, "contraparte", e.target.value.toUpperCase())} /></td>
                          <td style={{ padding: 3 }}><input style={{ ...inputBase, fontFamily: MONO }} value={r.nif} onChange={(e) => updSL(r.id, "nif", e.target.value.toUpperCase().replace(/[\s\-\.]/g, ""))} /></td>
                          <td style={{ padding: 3 }}><input style={{ ...inputBase, fontFamily: MONO }} value={r.numero} onChange={(e) => updSL(r.id, "numero", e.target.value)} /></td>
                          <td style={{ padding: 3 }}><input style={{ ...inputBase, fontFamily: MONO }} value={r.fechaFactura} onChange={(e) => updSL(r.id, "fechaFactura", e.target.value)} onBlur={(e) => { const v = normFecha(e.target.value); updSL(r.id, "fechaFactura", v); updSL(r.id, "fechaAsiento", v); }} placeholder="dd/mm/aaaa" /></td>
                          <td style={{ padding: 3 }}><input style={inputNum} value={r.base} onChange={(e) => updSL(r.id, "base", e.target.value)} /></td>
                          <td style={{ padding: 3 }}><input style={inputNum} value={r.tipoIva} onChange={(e) => updSL(r.id, "tipoIva", e.target.value)} /></td>
                          <td style={{ padding: 3 }}><input style={inputNum} value={r.cuotaIva} onChange={(e) => updSL(r.id, "cuotaIva", e.target.value)} /></td>
                          <td style={{ padding: 3 }}><input style={{ ...inputNum, fontWeight: 700 }} value={r.total} onChange={(e) => updSL(r.id, "total", e.target.value)} /></td>
                          <td style={{ padding: 3 }}>
                            <select style={{ ...inputBase, fontWeight: 600 }} value={CONCEPTOS_SL.includes(r.concepto) ? r.concepto : "GASTOS"} onChange={(e) => updSL(r.id, "concepto", e.target.value)}>
                              {CONCEPTOS_SL.map((c) => (<option key={c} value={c}>{c}</option>))}
                            </select>
                          </td>
                          <td style={{ padding: 3 }} title={r.subCPNombre || ""}>
                            <input style={{ ...inputBase, fontFamily: MONO, background: subCPbg, borderColor: r.subCPCreada ? C.warn : C.linea }} value={r.subCP} onChange={(e) => updSL(r.id, "subCP", e.target.value.replace(/\D/g, ""))} />
                            {r.subCPCreada && (
                              <div style={{ fontSize: 9.5, fontWeight: 800, color: C.warn, letterSpacing: "0.03em", marginTop: 2, whiteSpace: "nowrap" }}>
                                ● NUEVO · no estaba en el listado
                              </div>
                            )}
                          </td>
                          <td style={{ padding: 3 }}>
                            <input style={{ ...inputBase, fontFamily: MONO }} value={r.subIva} onChange={(e) => updSL(r.id, "subIva", e.target.value.replace(/\D/g, ""))} />
                          </td>
                          <td style={{ padding: 3 }} title={r.subGINombre || ""}>
                            <input style={{ ...inputBase, fontFamily: MONO, background: subGIbg }} value={r.subGI} onChange={(e) => updSL(r.id, "subGI", e.target.value.replace(/\D/g, ""))} />
                          </td>
                          <td style={{ padding: 3 }} title="Subcuenta de banco (cobro/pago)">
                            <input style={{ ...inputBase, fontFamily: MONO }} value={r.subBanco || ""} onChange={(e) => updSL(r.id, "subBanco", e.target.value.replace(/\D/g, ""))} />
                          </td>
                          <td style={{ padding: 3, whiteSpace: "nowrap", position: "sticky", right: 0, background: fondo, zIndex: 1, borderLeft: `1px solid ${C.linea}` }}>
                            {fueraTrimSL && (
                              <button onClick={() => { const f = primerDiaTrimestre(trimestreSel, anioSel); updSL(r.id, "fechaFactura", f); updSL(r.id, "fechaAsiento", f); }} title={`Ajustar al ${trimestreSel}T ${anioSel} (${primerDiaTrimestre(trimestreSel, anioSel)})`} style={{ border: "none", background: "transparent", cursor: "pointer", color: C.warn, padding: 4 }} aria-label="Ajustar fecha al trimestre"><CalendarCheck size={15} /></button>
                            )}
                            <button onClick={() => recalcularSL(r.id)} title="Recalcular IVA y total desde la base" style={{ border: "none", background: "transparent", cursor: "pointer", color: C.vino, padding: 4 }} aria-label="Recalcular"><Calculator size={15} /></button>
                            <button onClick={() => borrarFilaSL(r.id)} title="Eliminar esta línea" style={{ border: "none", background: "transparent", cursor: "pointer", color: C.err, padding: 4 }} aria-label="Eliminar línea"><Trash2 size={16} /></button>
                          </td>
                        </tr>
                        {iss.length > 0 && (
                          <tr style={{ background: fondo }}>
                            <td></td>
                            <td colSpan={15} style={{ padding: "0 6px 7px", fontSize: 11.5, color: hayErr ? C.err : C.warn }}>
                              {iss.map((i, k) => (<span key={k} style={{ marginRight: 14 }}>• {i.msg}</span>))}
                              <span style={{ color: C.gris }}>· origen: {r.fileName}</span>
                            </td>
                          </tr>
                        )}
                      </React.Fragment>
                    );
                  })}
                </tbody>
              </table>
            </div>

            <div className="mt-5 flex items-center gap-4 flex-wrap">
              <button
                style={{ ...btn(true), opacity: rowsSL.some((r) => r.sentido !== "venta") ? 1 : 0.4 }}
                disabled={!rowsSL.some((r) => r.sentido !== "venta")}
                onClick={() => exportarSL("compra")}
                title="Exporta solo las compras, con la plantilla COMPRAS de Monitor"
              >
                <Download size={16} /> Exportar COMPRAS
              </button>
              <button
                style={{ ...btn(true), opacity: rowsSL.some((r) => r.sentido === "venta") ? 1 : 0.4 }}
                disabled={!rowsSL.some((r) => r.sentido === "venta")}
                onClick={() => exportarSL("venta")}
                title="Exporta solo las ventas, con la plantilla VENTAS de Monitor"
              >
                <Download size={16} /> Exportar VENTAS
              </button>
              {nErrSL > 0 && (
                <span style={{ fontSize: 12.5, color: C.err, display: "flex", alignItems: "center", gap: 6 }}>
                  <AlertTriangle size={15} /> Hay {nErrSL} línea{nErrSL !== 1 ? "s" : ""} con errores: corrígelas antes de importar en el Conversor.
                </span>
              )}
              <span style={{ fontSize: 12, color: C.gris, marginLeft: "auto" }}>
                Excel con el layout de las plantillas de Monitor (Fecha asiento · Fecha factura · Nº · Concepto · Subcta. cliente/prov. · NIF · Nombre · … · Base · %IVA · Cuota IVA · Subcta. IVA · … · Subcta. gasto/ingreso · Importe gasto · DEBE · HABER · Total). El <b>banco va en DEBE/HABER</b> (pago/cobro) para que Monitor genere también el asiento de banco. Las subcuentas en blanco (cliente/proveedor nuevo) las crea Monitor al importar.
              </span>
            </div>
          </section>
        )}

        <footer style={{ textAlign: "center", fontSize: 11.5, color: C.gris, paddingBottom: 16 }}>
          ASEMA Advisory · Chiclana de la Frontera · Herramienta interna del despacho · v2.8 — revisa siempre los apuntes antes de importar en Monitor.
        </footer>
      </main>
    </div>
  );
}
