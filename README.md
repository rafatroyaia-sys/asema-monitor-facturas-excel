# ASEMA · Facturas → Excel Monitor (EOS Autónomos)

Herramienta interna del despacho: sube facturas escaneadas (PDF/JPG/PNG), la IA las lee
y clasifica, las revisas en una tabla editable y exportas el Excel normalizado para
importarlo en Monitor Informática (módulo de autónomos / estimación directa simplificada).

**Versión 1.0** · Frontend React (Vite) + función serverless de Vercel que llama a la API
de Anthropic. La clave API nunca llega al navegador. El acceso está protegido con una
clave del despacho.

---

## 1 · Requisitos previos (una sola vez)

1. **Cuenta API de Anthropic**: entra en https://console.anthropic.com, crea cuenta,
   añade tarjeta en *Billing* y carga crédito (con 5-10 $ empiezas de sobra).
   En *API Keys* crea una clave nueva (empieza por `sk-ant-...`) y guárdala en sitio
   seguro: **solo se muestra una vez**.
   - Ojo: esta cuenta y su facturación son **independientes** de tu suscripción de
     claude.ai. Aquí pagas por uso (~1 céntimo por factura, aprox.).
2. **Cuenta de GitHub** (gratuita): https://github.com
3. **Cuenta de Vercel** (gratuita): https://vercel.com — regístrate con tu cuenta de GitHub.

## 2 · Despliegue paso a paso

### Opción A — Con Claude Code (recomendada para ti)

Descomprime este proyecto en una carpeta, abre Claude Code dentro de ella y dile:

> "Quiero desplegar este proyecto en Vercel. Crea el repositorio en GitHub, conéctalo a
> Vercel y configura las variables de entorno ANTHROPIC_API_KEY y ASEMA_PASSWORD.
> Te iré dando los valores cuando me los pidas."

Claude Code te guiará (instalación de la CLI de Vercel incluida si hace falta).

### Opción B — Desde la web de Vercel (sin terminal)

1. Crea un repositorio nuevo en GitHub (privado) y sube el contenido de esta carpeta.
2. En https://vercel.com → **Add New → Project** → importa ese repositorio.
3. Vercel detecta Vite automáticamente. Antes de pulsar *Deploy*, abre
   **Environment Variables** y añade:

   | Nombre              | Valor                                          |
   |---------------------|------------------------------------------------|
   | `ANTHROPIC_API_KEY` | tu clave `sk-ant-...` de console.anthropic.com |
   | `ASEMA_PASSWORD`    | la clave del despacho que tú decidas           |

4. Pulsa **Deploy**. En un minuto tendrás la URL fija, p. ej.
   `https://asema-facturas.vercel.app`.
5. En cada PC del despacho: abre la URL en Chrome/Edge → menú ⋮ →
   *Guardar y compartir / Crear acceso directo* → marca "Abrir como ventana".
   Queda como un programa más en el escritorio.

## 3 · Uso diario

1. Abrir la app → introducir la **clave del despacho** (se recuerda en ese PC).
2. Seleccionar o dar de alta el **cliente** (nombre + NIF). Los clientes se guardan en
   el navegador de cada PC.
3. Arrastrar las facturas escaneadas y pulsar **Procesar**.
4. Revisar la tabla: filas en rojo = errores que impedirían la importación; en ámbar =
   avisos (descuadres, NIF dudoso, confianza baja...). Todo es editable; la calculadora
   de cada fila recalcula cuotas y total desde la base.
5. **Exportar Excel para Monitor** → importar el .xlsx en Monitor con la plantilla.

### Columnas del Excel (fila 1 = cabecera, datos desde la fila 2)

NOMBRE CLIENTE-PROVEEDOR · NIF · Nº FACTURA · FECHA (dd/mm/aaaa) · BASE IMPONIBLE ·
CANTIDAD IVA · TOTAL · TIPO DE IVA · CUENTA (número: 1, 2, 5, 10...) · CONCEPTO ·
CLAVE GASTO (R/P/E/A/G/I/S/O) · TIPO RETENCION · CUOTA RETENCION · TIPO RE · CUOTA RE

## 4 · Actualizaciones (misma URL siempre)

Cuando haya que corregir o mejorar algo: se actualiza el código, se sube el cambio al
repositorio (`git push` o desde Claude Code) y **Vercel redespliega solo sobre la misma
URL**. Nadie tiene que cambiar accesos directos.

## 5 · Límites y consejos

- **Máx. 4 MB por archivo** (límite de la petición en Vercel). Escanea a 150-200 ppp en
  gris: sobra para que la IA lea bien y los archivos quedan pequeños.
- Si un PDF trae muchas facturas y la respuesta llega incompleta, divide el lote.
- Coste orientativo: ~1 céntimo por factura de una página (modelo claude-sonnet-4-6,
  3 $/M tokens entrada y 15 $/M salida). En console.anthropic.com → *Usage* ves el gasto
  real, y en *Limits* puedes fijar un tope mensual de gasto. Hazlo.
- La tabla de revisión NO se guarda: procesa → revisa → exporta en la misma sesión.
- RGPD: las facturas se envían a la API de Anthropic para su lectura. Documéntalo en el
  registro de actividades de tratamiento del despacho como cualquier encargado de
  tratamiento cloud.

## 6 · Problemas frecuentes

| Síntoma | Causa probable | Solución |
|---|---|---|
| "Clave del despacho incorrecta" | ASEMA_PASSWORD no coincide | Revisar la variable en Vercel → Settings → Environment Variables (tras cambiarla, *Redeploy*) |
| Error 502 con mensaje de crédito/billing | Sin crédito en la cuenta API | Cargar crédito en console.anthropic.com |
| "El servidor no tiene configurada la variable..." | Variables de entorno sin definir | Añadirlas en Vercel y redesplegar |
| "Respuesta JSON incompleta" | PDF con demasiadas facturas | Dividir el lote en PDFs más pequeños |
| Archivo rechazado por tamaño | Escaneo a demasiada resolución | Reescanear a 150-200 ppp o dividir el PDF |

---

ASEMA Advisory · Chiclana de la Frontera · Uso interno.
