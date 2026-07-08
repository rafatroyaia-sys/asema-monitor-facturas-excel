/* ============================================================
   ASEMA · Backend de extracción de facturas
   Función serverless de Vercel: /api/extraer
   - Valida la clave del despacho (variable ASEMA_PASSWORD)
   - Llama a la API de Anthropic (variable ANTHROPIC_API_KEY)
   - Devuelve { facturas: [...] } ya parseado
   La clave API NUNCA llega al navegador.
   ============================================================ */

/* Limpia secretos (clave API, contraseña): elimina BOM, saltos de linea,
   espacios y cualquier caracter invisible que se cuele al pegar el valor
   en las variables de entorno de Vercel. */
const limpiaSecreto = (s) => String(s || "").replace(/[^\x21-\x7E]/g, "");

const buildPrompt = (empresa) => `Eres el sistema de extracción contable de una asesoría española (Graduado Social). Analiza el documento adjunto, que contiene una o varias facturas escaneadas, y devuelve EXCLUSIVAMENTE un objeto JSON válido. Sin markdown, sin comentarios, sin texto antes ni después.

TITULAR DE LA CONTABILIDAD (cliente de la asesoría, régimen de estimación directa simplificada):
- Nombre: ${empresa.nombre}
- NIF: ${empresa.nif}

CLASIFICACIÓN — campo "cuenta" (plan de cuentas del programa Monitor para autónomos):
- 1 (VENTAS): el titular es el EMISOR de la factura → ingreso. "contraparte" = el cliente que recibe la factura.
- 2 (COMPRAS): factura RECIBIDA de mercancía, género o materias primas propias de su actividad. "contraparte" = proveedor.
- 5 (ENERGÍA): suministro de electricidad.
- 6 (ALQUILERES): solo si es claramente alquiler del local de negocio (suele llevar retención IRPF).
- 8 (PRIMAS SEG.): solo si es claramente una prima de seguro.
- 10 (GASTOS DIVERSOS): cualquier otro gasto. Asigna además "clave":
  "A"=suministro de agua · "G"=suministro de gas · "I"=telefonía e internet · "S"=otros suministros · "P"=servicios de profesionales independientes (abogado, notario, registrador, asesor, arquitecto, técnico...) · "R"=reparaciones y conservación · "E"=otros servicios exteriores · "O"=otros conceptos fiscalmente deducibles.
- "clave" solo se rellena cuando cuenta=10; en el resto, cadena vacía "".
- Para decidir el sentido (emitida/recibida) compara el NIF del titular con emisor y destinatario. Si el NIF no consta, usa la similitud del nombre. Si aun así hay duda, "confianza":"baja" y explícalo en "obs".

REGLAS DE EXTRACCIÓN:
- Una entrada en "facturas" por cada factura distinta del documento (puede haber varias por PDF). También tickets/facturas simplificadas.
- "lineas": una por cada tipo de IVA distinto que tenga la factura (21, 10, 4, 5, 2, 0...). Operación exenta o sin IVA → tipo_iva 0 y cuota_iva 0.
- Importes: número con punto decimal y 2 decimales. Porcentajes: número (21, 10, 4, 0, 5.2, 1.4, 0.5, 15, 19, 7...).
- "nif": mayúsculas, sin espacios, puntos ni guiones. Si no consta en el documento → "0".
- "fecha": fecha de expedición de la factura en formato dd/mm/aaaa.
- Retención IRPF (típica de profesionales y alquileres): "tipo_ret" (%) y "cuota_ret" (importe, positivo). Si no hay, 0 y 0.
- Recargo de equivalencia, si aparece, por línea: "tipo_re" (%) y "cuota_re". Si no hay, 0 y 0.
- SUPLIDOS: si la factura refactura "suplidos" o "gastos y suplidos" (importes pagados por cuenta del cliente: tasas, comidas, entradas, desplazamientos, museos, mercado...), NO llevan IVA ni retención y NO forman parte de la base imponible ni de "lineas". Pon la SUMA total de esos suplidos en el campo "suplidos" (número). Si no hay, 0.
- "total": importe total a pagar/cobrar de la factura tal como figura en el documento.
- "numero": número de factura tal como figura (con su serie si la tiene).
- Si un dato es ilegible o no consta: usa null en ese campo, marca "confianza":"baja" y explica en "obs" qué falta.
- "confianza": "alta", "media" o "baja" según la legibilidad del escaneo y la seguridad de la clasificación.

FORMATO EXACTO DE RESPUESTA (solo este JSON):
{"facturas":[{"contraparte":"NOMBRE","nif":"B11111111","numero":"A-123","fecha":"01/09/2025","lineas":[{"base":100.00,"tipo_iva":21,"cuota_iva":21.00,"tipo_re":0,"cuota_re":0}],"tipo_ret":0,"cuota_ret":0,"suplidos":0,"total":121.00,"cuenta":10,"clave":"O","confianza":"alta","obs":""}]}`;

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Método no permitido" });
  }

  // --- Control de acceso del despacho ---
  // .trim() en ambos lados: elimina espacios y saltos de linea invisibles
  // que a veces se cuelan al definir variables de entorno desde la CLI.
  const clave = limpiaSecreto(req.headers["x-asema-key"]);
  const esperada = limpiaSecreto(process.env.ASEMA_PASSWORD);
  if (!esperada) {
    return res.status(500).json({ error: "El servidor no tiene configurada la variable ASEMA_PASSWORD" });
  }
  if (!clave || clave !== esperada) {
    return res.status(401).json({ error: "Clave del despacho incorrecta" });
  }
  if (!process.env.ANTHROPIC_API_KEY) {
    return res.status(500).json({ error: "El servidor no tiene configurada la variable ANTHROPIC_API_KEY" });
  }

  const { media_type, data, empresa, promptOverride, textoDocumento } = req.body || {};
  // Debe venir un documento (archivo en base64) o el texto de un Word ya extraído.
  if (!textoDocumento && (!data || !media_type)) {
    return res.status(400).json({ error: "Faltan datos: archivo" });
  }
  // Si no hay prompt alternativo (listados), se exige el cliente del despacho.
  if (!promptOverride && (!empresa || !empresa.nombre || !empresa.nif)) {
    return res.status(400).json({ error: "Faltan datos: cliente del despacho" });
  }

  // Bloque de contenido: texto (Word), PDF, o imagen.
  let bloque;
  if (textoDocumento) {
    bloque = { type: "text", text: "CONTENIDO DEL DOCUMENTO WORD (factura en texto):\n\n" + String(textoDocumento).slice(0, 60000) };
  } else if (media_type === "application/pdf") {
    bloque = { type: "document", source: { type: "base64", media_type: "application/pdf", data } };
  } else {
    bloque = { type: "image", source: { type: "base64", media_type, data } };
  }

  try {
    const r = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "content-type": "application/json",
        "x-api-key": limpiaSecreto(process.env.ANTHROPIC_API_KEY),
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify({
        model: "claude-sonnet-4-6",
        max_tokens: 16000,
        messages: [{ role: "user", content: [bloque, { type: "text", text: promptOverride || buildPrompt(empresa) }] }],
      }),
    });

    const dataResp = await r.json();
    if (!r.ok) {
      const msg = dataResp?.error?.message || `Error de la API de Anthropic (HTTP ${r.status})`;
      return res.status(502).json({ error: msg });
    }

    const texto = (dataResp.content || [])
      .filter((b) => b.type === "text")
      .map((b) => b.text)
      .join("\n");
    const ini = texto.indexOf("{");
    const fin = texto.lastIndexOf("}");
    if (ini === -1 || fin === -1) {
      return res.status(500).json({ error: "La IA no devolvió JSON. Reintenta o revisa la calidad del escaneo." });
    }
    let parsed;
    try {
      parsed = JSON.parse(texto.slice(ini, fin + 1));
    } catch {
      return res.status(500).json({ error: "Respuesta JSON incompleta. Divide el PDF en lotes con menos facturas." });
    }
    if (!parsed.facturas || !Array.isArray(parsed.facturas)) {
      return res.status(500).json({ error: "Formato de respuesta inesperado de la IA." });
    }
    return res.status(200).json({ facturas: parsed.facturas });
  } catch (e) {
    return res.status(500).json({ error: e.message || "Error interno del servidor" });
  }
}
