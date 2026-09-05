# Golfito IA (golfito-whatsapp)

> Parte de **Golfito (B2C)** dentro del ecosistema GolfitoGolf.IA — ver [`../CLAUDE.md`](../CLAUDE.md) para el contexto de negocio completo. Este es el producto en sí (no confundir con `../golfito-social/`, que es solo su marketing).

Chatbot "profesor de golf" sobre WhatsApp: recibe videos de swing, los analiza con IA, da ejercicios y planes de entrenamiento personalizados, y cobra por análisis/planes vía MercadoPago. Incluye un panel web para que los profesores supervisen y respondan manualmente.

> Carpeta movida desde `C:\Users\marti\Golfito-ia` el 2026-09-03 (repo git intacto, `.git` incluido). Si alguna doc vieja o algún script todavía referencia esa ruta, actualizarla al toparse con ella.

## Negocio (B2C)

Embudo: ejercicio gratis automatizado (lead gen, sin costo) → análisis de swing con Gemini (pago) → plan de entrenamiento personalizado (pago). Generación de ejercicios vía **GPT-4o-mini**; análisis de swing vía **Gemini 2.5-flash** (feedback cualitativo + scoring de 7 dimensiones — ver `_analizarSwingConGemini` / Índice Golfito en "Entry points" abajo). Cobros de ambos pasos pagos vía MercadoPago (ver `_crearPreferenciaPago` etc.).

- **WhatsApp Business (Meta):** número `922201076`, App ID `1008146235206568`, Business Manager "Servicios Informáticos Martín Schroeder SpA". Esta misma app/Business Manager la reusa `golfito-social` para Facebook/Instagram — no crear una app nueva para nada relacionado a Golfito.
- **Cloud Run `generar-imagen-swing`** (proyecto `golfito-prod`, `southamerica-west1`): genera las imágenes comparativas de swing (antes/después o vs. referencia) que se mandan como parte del análisis. Vive fuera de este repo (es un servicio aparte en GCP), pero es parte del flujo de análisis de swing.
- **Deploy vía clasp (confirmado funcionando 2026-09-05):** `clasp push` sube `codigo.gs`/`appsscript.json`/`webapp.html` al proyecto de Apps Script (ya logueado como `martinmdp93@gmail.com`, ver `.clasp.json`). Ojo: el webhook de WhatsApp y de MercadoPago (`MP_WEBHOOK_URL` en `codigo.gs`) apuntan a un **deployment versionado**, no a `@HEAD` — `clasp push` por sí solo NO actualiza lo que corre en producción. Para que el cambio quede realmente live hay que correr `clasp deploy -i <deploymentId>` (ver `MP_WEBHOOK_URL` para el ID) para apuntar ese mismo deployment a la nueva versión. El problema de TLS documentado antes ya no reproduce (probablemente resuelto por una actualización de Node/clasp).

## Stack

- **Backend:** Google Apps Script (`codigo.gs`, ~1900 líneas, un solo archivo). Runtime V8.
- **Frontend:** un único `webapp.html` servido por `doGet` (panel de entrenadores), con JS embebido que llama al backend vía `google.script.run`.
- **Config:** `appsscript.json` — timeZone America/Santiago, acceso webapp `ANYONE_ANONYMOUS` ejecutando como `USER_DEPLOYING`.
- **Datos:** Google Sheets como base de datos (hojas: `Leads`, `Sesiones`, `Excercises_Gemini`, `Conversations`, `Consultas`, `IndiceGolfito`, `ChatLog`).
- **Storage de video:** Google Drive (carpeta `Golfito_Videos`).
- **Integraciones externas:** WhatsApp Business (Meta Graph API), OpenAI (generación de planes), Google Gemini (análisis de video de swing y del Índice Golfito), MercadoPago (cobros y reintegros).
- **Secrets:** todos vía `PropertiesService.getScriptProperties()` (`OPENAI_API_KEY`, `MP_ACCESS_TOKEN`, `MP_ACCESS_TOKEN_AR`, `PHONE_NUMBER_ID`, `META_TOKEN`, `GEMINI_API_KEY`, `PANEL_PASSWORD`, `WEBHOOK_SECRET`, flags `MODO_TEST_ANALISIS`/`MODO_TEST_PLAN`). Nunca hardcodeados en el código (excepto `VERIFY_TOKEN` del webhook de Meta, que está literal en el código — es el único pendiente del audit original, de severidad menor).

## Entry points

- `doGet(e)` — maneja: confirmación de pago MercadoPago (`type=payment`/`source=mp`), verificación de webhook de Meta (`hub.verify_token`), y sirve el panel HTML por default. Los dos primeros validan `_secretWebhookValido(e)` antes de procesar.
- `doPost(e)` — webhook único para: notificaciones de pago de MercadoPago y mensajes entrantes de WhatsApp (texto o video). También valida `_secretWebhookValido(e)`. Hace dedup de mensajes por `message.id` (CacheService, 10 min) y de videos por `mediaId` (ScriptProperties) para evitar reprocesar reintentos de Meta.
- `_procesarMensajeEntrante(from, text)` — máquina de estados de la conversación por WhatsApp (la función más grande del archivo). Corre bajo `LockService.getUserLock()`.
- `_procesarVideoEntrante` / `_procesarAnalisisVideo` / `_analizarSwingConGemini` — flujo de análisis de swing. `_procesarIndiceGolfito` / `_analizarIndiceConGemini` — flujo del Índice Golfito (score 1-100 sobre 7 dimensiones técnicas).
- `generarYEnviarPlanDesdeWeb` / `_generarHTMLPlan` / `generarPlanConIA` — generación de planes (vía OpenAI + banco de ejercicios en Sheets), expuestas al panel (requieren token).
- `_crearPreferenciaPago` / `_procesarPagoMP` / `_procesarReintegroMP` — cobros y reintegros vía MercadoPago. `_procesarPagoMP` corre bajo el mismo `LockService.getUserLock()` que el chat para evitar carrera sobre el saldo de wallet.

La mayoría de las funciones internas llevan prefijo `_` (Apps Script las esconde de `google.script.run`, evitando que cualquier visitante anónimo del panel las invoque directo). Quedan sin prefijo: `doGet`/`doPost`, las funciones "WEB APP FUNCTIONS" que el panel llama legítimamente (protegidas con `_autorizadoPanel(token)`), `validarPasswordPanel`/`validarTokenPanel`, y los 3 triggers de tiempo (`procesarSesionesPendientes`, `enviarFeedbackPendiente`, `limpiarCacheAnalisis` — Apps Script no permite que un trigger apunte a una función con `_`).

## Panel web (auth)

- Login por password único (`PANEL_PASSWORD` en Script Properties) vía `validarPasswordPanel`. Rate limit: bloquea el login tras 5 intentos fallidos en una ventana de 15 min (`ScriptProperties` clave `panel_login_fails`).
- Token = `Utilities.getUuid()` random por login (no derivado del password), guardado en `ScriptProperties` con expiración de 8h (`validarTokenPanel`).
- **Toda función expuesta al panel via `google.script.run` debe validar el token server-side ella misma** (patrón `_autorizadoPanel(token)`) — el login en `webapp.html` es solo UI, cualquier función global de Apps Script es invocable directamente sin pasar por el panel.
- Los webhooks (Meta y MercadoPago) usan un mecanismo aparte: `WEBHOOK_SECRET` viaja como query param `?secret=...` en la URL, vía `_secretWebhookValido(e)`. Es necesario porque Apps Script no expone headers HTTP custom (no se puede leer `X-Hub-Signature-256` ni `x-signature`). Diseño fail-open: si `WEBHOOK_SECRET` no está configurado en Script Properties, no bloquea nada.

## Convenciones del código existente

- Todo en un solo archivo `codigo.gs`, organizado por comentarios de sección (`// ====...`), no por módulos.
- Nombres de funciones y variables en español.
- Convención de prefijo `_` = función interna, no expuesta a `google.script.run` (ver sección "Entry points" arriba).
- `_safeString`, `_normalizeText`, `_truncateForWhatsApp` son helpers usados en todo el archivo para sanitizar inputs antes de guardarlos en Sheets o mandarlos por WhatsApp.
- Los videos entrantes se rechazan si duran más de 10s (`_obtenerDuracionMp4`, parsea el box `mvhd` del MP4; fail-open si no lo puede leer).
- Cualquier paso de la máquina de estados que solo repite un mensaje fijo sin avanzar el `paso` necesita una vía de salida (palabra clave *saltar*, o un contador tipo `intentos_video`) — si no, el alumno queda en loop infinito si nunca manda lo que se le pide.
- Hay varias funciones `_test*` al final del archivo (`_testDrive`, `_testEnvioMeta`, `_listarModelosGemini`, etc.) pensadas para ejecutarse manualmente desde el editor de Apps Script, no parte del flujo de producción.

## Notas para trabajar acá

- El archivo es grande (~197KB) — al leerlo con herramientas, las líneas 29-30 son los logos en base64 (decenas de miles de caracteres cada una); conviene leer por rangos de línea o vía grep en vez de cargar el archivo completo.
- No hay tests automatizados ni linter configurado.
- El historial de git tiene commits frecuentes tipo "Update codigo.gs" — no hay changelog ni mensajes descriptivos, así que para entender una decisión conviene mirar el código actual, no el historial de commits.
