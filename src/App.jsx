import React, { useState, useEffect, useRef, useMemo } from "react";
import * as XLSX from "xlsx";
import {
  Upload, FileText, Trash2, Download, Calculator, AlertTriangle,
  CheckCircle2, XCircle, Loader2, Building2, Save, Plus, FileSpreadsheet,
  Lock, LogOut,
} from "lucide-react";

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

/* ---------- Concepto del apunte (campo "Concepto" de Monitor) ---------- */
const conceptoPorCuenta = (n) => {
  const c = Number(n);
  if (c === 1) return "INGRESOS POR VENTAS";
  if (c === 2) return "GASTOS DE COMPRAS";
  return "GASTOS VARIOS";
};

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

/* ---------- Llamada al backend del despacho ---------- */
async function extraerFacturas(file, empresa, claveDespacho) {
  const b64 = await toB64(file);
  const esPdf = file.type === "application/pdf" || /\.pdf$/i.test(file.name);
  const media = esPdf ? "application/pdf" : file.type || "image/jpeg";

  const resp = await fetch("/api/extraer", {
    method: "POST",
    headers: { "Content-Type": "application/json", "x-asema-key": limpiaClave(claveDespacho) },
    body: JSON.stringify({ media_type: media, data: b64, empresa }),
  });
  let data = {};
  try { data = await resp.json(); } catch { /* respuesta sin cuerpo */ }
  if (resp.status === 401) {
    const e = new Error(data.error || "Clave del despacho incorrecta");
    e.auth = true;
    throw e;
  }
  if (!resp.ok) throw new Error(data.error || `Error del servidor (HTTP ${resp.status})`);
  if (!Array.isArray(data.facturas)) throw new Error("Respuesta inesperada del servidor.");
  return data.facturas;
}

/* ---------- Facturas extraídas → filas de la tabla ---------- */
function facturasARows(facturas, fileName) {
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
        fecha: normFecha(f.fecha),
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

/* ---------- Validación de cada fila ---------- */
function validarFila(row, rows) {
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
  return issues;
}

/* ============================================================ */
export default function App() {
  const [claveDespacho, setClaveDespacho] = useState(() => {
    try { return limpiaClave(localStorage.getItem("asema_clave")); } catch { return ""; }
  });
  const [autenticado, setAutenticado] = useState(() => {
    try { return !!localStorage.getItem("asema_clave"); } catch { return false; }
  });
  const [claveInput, setClaveInput] = useState("");
  const [avisoClave, setAvisoClave] = useState("");

  const [empresa, setEmpresa] = useState({ nombre: "", nif: "" });
  const [periodo, setPeriodo] = useState("");
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
      const okTipo = f.type === "application/pdf" || /^image\/(jpeg|png|webp)$/.test(f.type) || /\.(pdf|jpe?g|png|webp)$/i.test(f.name);
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
    if (!empresa.nombre || !empresa.nif) {
      setAviso("Antes de procesar, indica el nombre y NIF del cliente: la app lo necesita para distinguir facturas emitidas de recibidas.");
      return;
    }
    setAviso("");
    setProcesando(true);
    guardarEmpresa();
    const pendientes = files.filter((f) => f.status === "pendiente" || f.status === "error");
    for (const f of pendientes) {
      const obj = fileObjs.current[f.id];
      if (!obj || obj.size > MAX_MB * 1024 * 1024) continue;
      setFile(f.id, { status: "procesando", msg: "" });
      try {
        const facturas = await extraerFacturas(obj, empresa, claveDespacho);
        const nuevas = facturasARows(facturas, f.name);
        setRows((p) => [...p, ...nuevas]);
        setFile(f.id, { status: "ok", nFacturas: facturas.length, msg: "" });
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
        cuenta: 10, concepto: "GASTOS VARIOS", clave: "", confianza: "alta", obs: "",
      },
    ]);

  /* ----- Validación y resumen ----- */
  const issuesById = useMemo(() => {
    const m = {};
    rows.forEach((r) => { m[r.id] = validarFila(r, rows); });
    return m;
  }, [rows]);

  const nErr = rows.filter((r) => issuesById[r.id]?.some((i) => i.lv === "err")).length;
  const nWarn = rows.filter((r) => issuesById[r.id]?.some((i) => i.lv === "warn") && !issuesById[r.id]?.some((i) => i.lv === "err")).length;

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
    const nombre = `MONITOR_${limpio(empresa.nombre) || "CLIENTE"}${periodo ? "_" + limpio(periodo) : ""}.xlsx`;
    XLSX.writeFile(wb, nombre);
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

        {/* ---------- 1 · Cliente ---------- */}
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
            <div className="md:col-span-2">
              <label style={{ fontSize: 11, fontWeight: 700, color: C.gris, letterSpacing: "0.05em" }}>PERIODO (opcional)</label>
              <input style={inputBase} value={periodo} onChange={(e) => setPeriodo(e.target.value)} placeholder="3T 2025" />
            </div>
            <div className="md:col-span-2">
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

        {/* ---------- 2 · Facturas ---------- */}
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
              accept=".pdf,.jpg,.jpeg,.png,.webp,application/pdf,image/jpeg,image/png,image/webp"
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
                  {f.status === "procesando" && <span style={{ fontSize: 12, color: C.vino }}>leyendo…</span>}
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
            <button style={btn(false)} onClick={filaManual}>
              <Plus size={16} /> Añadir apunte manual
            </button>
          </div>
        </section>

        {/* ---------- 3 · Revisión ---------- */}
        {rows.length > 0 && (
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
                    <th style={{ ...th, width: 76 }}></th>
                  </tr>
                </thead>
                <tbody>
                  {rows.map((r, idx) => {
                    const iss = issuesById[r.id] || [];
                    const hayErr = iss.some((i) => i.lv === "err");
                    const hayWarn = iss.some((i) => i.lv === "warn");
                    const fondo = hayErr ? C.errBg : hayWarn ? C.warnBg : idx % 2 ? "#FCFAF4" : C.papel;
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
                            <input style={inputBase} value={r.concepto || ""} onChange={(e) => upd(r.id, "concepto", e.target.value.toUpperCase())} />
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
                          <td style={{ padding: 3, whiteSpace: "nowrap" }}>
                            <button onClick={() => recalcular(r.id)} title="Recalcular cuotas y total desde la base" style={{ border: "none", background: "transparent", cursor: "pointer", color: C.vino, padding: 4 }} aria-label="Recalcular">
                              <Calculator size={15} />
                            </button>
                            <button onClick={() => borrarFila(r.id)} title="Eliminar apunte" style={{ border: "none", background: "transparent", cursor: "pointer", color: C.gris, padding: 4 }} aria-label="Eliminar">
                              <Trash2 size={15} />
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

        <footer style={{ textAlign: "center", fontSize: 11.5, color: C.gris, paddingBottom: 16 }}>
          ASEMA Advisory · Chiclana de la Frontera · Herramienta interna del despacho · v1.2 — revisa siempre los apuntes antes de importar en Monitor.
        </footer>
      </main>
    </div>
  );
}
