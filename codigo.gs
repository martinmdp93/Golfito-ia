// ============================================
// CONFIGURACIÓN GENERAL
// ============================================
const LEADS_SHEET = "Leads";
const SESIONES_SHEET = "Sesiones";
const EXERCISES_SHEET = "Excercises_Gemini";
const CONVERSATIONS_SHEET = "Conversations";
const CONSULTAS_SHEET = "Consultas";

const OPENAI_API_KEY = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
const MP_ACCESS_TOKEN = PropertiesService.getScriptProperties().getProperty("MP_ACCESS_TOKEN");
const MP_WEBHOOK_URL = "https://script.google.com/macros/s/AKfycbydtxoAOR2ZHUwBfOsp5kehU7OgtB5RgUh_EtcUo8r4Hya3M6Z89B09CX4N-KEH_6Wfjw/exec";
const PHONE_NUMBER_ID = PropertiesService.getScriptProperties().getProperty("PHONE_NUMBER_ID");
const META_TOKEN = PropertiesService.getScriptProperties().getProperty("META_TOKEN");
const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");

const FREE_EXERCISE_LIMIT = 1;
const VERIFY_TOKEN = "golfito_verify_123";
const MODO_TEST_ANALISIS = PropertiesService.getScriptProperties().getProperty("MODO_TEST_ANALISIS") === "true";
const MODO_TEST_PLAN = PropertiesService.getScriptProperties().getProperty("MODO_TEST_PLAN") === "true";

const ENTRADA_CALOR_STD = "Empeza con 5 minutos de movilidad articular (hombros, caderas y munecas). Tira 10-15 chips cortos para activar el tacto. Luego hace 5-8 swings completos a medio ritmo antes de arrancar con los ejercicios.";
const CONSIDERACIONES_STD = "Animo! Intenta este plan en 2 a 4 sesiones de entrenamiento y comentanos tus avances o cualquier duda adicional. Recorda siempre tirar algunas bolas de forma natural y sin pensamientos tecnicos antes de dejar el driving, y evita pensamientos complejos al competir.";

const COL = {
  TIMESTAMP: 1, WHATSAPP: 2, ASPECTO: 3, NIVEL: 4, EJVSPLAN: 5,
  STATUS: 6, CONTENIDO_ENVIADO: 7, EJERCICIO_GRATIS_ID: 8, VIDEO_URL1: 9,
  ANALISIS_ERROR_IA1: 10, ANALISIS_AREA_IA1: 11, ANALISIS_SEVERIDAD_IA1: 12,
  ANALISIS_RECOMENDACION_IA1: 13, ANALISIS_ANGULO_IA1: 14, VIDEO_URL2: 15,
  ANALISIS_ERROR_IA2: 16, ANALISIS_AREA_IA2: 17, ANALISIS_SEVERIDAD_IA2: 18,
  ANALISIS_RECOMENDACION_IA2: 19, ANALISIS_ANGULO_IA2: 20,
  PAGO_STATUS: 21, PAGO_MONTO: 22, PAGO_FECHA: 23, CODIGO_PLAN: 24,
  NUM_PLAN_ALUMNO: 25, NOTA_COACH: 26, DIAGNOSTICO: 27, ENTRADA_EN_CALOR: 28,
  EJERCICIO_APPROACH: 29, EJERCICIO_FULLSWING_HIERROS: 30,
  EJERCICIO_FULLSWING_MADERAS: 31, EJERCICIO_PUTTER: 32,
  CONSIDERACIONES: 33, FEEDBACK_ALUMNO: 34, FECHA_FEEDBACK: 35,
  COMENTARIOS_ALUMNO: 36, ANALISIS_MANUAL_1: 37, ANALISIS_MANUAL_2: 38,
  FEEDBACK_SCORE: 39, MP_PAYMENT_ID: 40
};

const SESIONES_HEADERS = [
  "timestamp","whatsapp","aspecto","nivel","ejvsplan",
  "status","contenido_enviado","ejercicio_gratis_id","video_url1",
  "analisis_error_IA1","analisis_area_IA1","analisis_severidad_IA1","analisis_recomendacion_IA1","analisis_angulo_IA1",
  "video_url2",
  "analisis_error_IA2","analisis_area_IA2","analisis_severidad_IA2","analisis_recomendacion_IA2","analisis_angulo_IA2",
  "pago_status","pago_monto","pago_fecha","codigo_plan","num_plan_alumno",
  "nota_coach","diagnostico","entrada_en_calor",
  "ejercicio_approach","ejercicio_fullswing_hierros","ejercicio_fullswing_maderas","ejercicio_putter",
  "consideraciones","feedback_alumno","fecha_feedback",
  "comentarios_alumno","analisis_manual_1","analisis_manual_2","feedback_score","mp_payment_id"
];

// ============================================
// WEBHOOK - GET
// ============================================
function doGet(e) {
  if (e.parameter?.type === "payment" && e.parameter?.["data.id"]) {
    procesarPagoMP(e.parameter["data.id"]);
    return ContentService.createTextOutput("OK");
  }
  if (e.parameter?.source === "mp" && e.parameter?.id) {
    procesarPagoMP(e.parameter["id"]);
    return ContentService.createTextOutput("OK");
  }
  if (e.parameter["hub.verify_token"] === VERIFY_TOKEN && e.parameter["hub.mode"] === "subscribe") {
    return ContentService.createTextOutput(e.parameter["hub.challenge"]);
  }
  return HtmlService.createHtmlOutputFromFile('webapp')
    .setTitle('Golfito — Panel de Entrenadores')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================
// WEBHOOK - POST
// ============================================
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.type === "payment" && body.data?.id) { procesarPagoMP(body.data.id); return okResponse(); }
    if (e.parameter?.type === "payment" && e.parameter?.["data.id"]) { procesarPagoMP(e.parameter["data.id"]); return okResponse(); }
    if (!body.entry?.[0]?.changes?.[0]?.value?.messages) return okResponse();
    const message = body.entry[0].changes[0].value.messages[0];
    const from = message.from;
    if (message.type === "video") {
      const mediaId = message.video.id;
      const props = PropertiesService.getScriptProperties();
      const guardKey = "media_processed_" + mediaId;
      if (props.getProperty(guardKey)) { Logger.log("Video duplicado ignorado: " + mediaId); return okResponse(); }
      props.setProperty(guardKey, "1");
      logMensaje(from, "entrante", "video", mediaId);
      procesarVideoEntrante(from, mediaId);
      return okResponse();
    }
    const text = safeString(message.text?.body).trim();
    if (!text) return okResponse();
    logMensaje(from, "entrante", "texto", text);
    procesarMensajeEntrante(from, text);
  } catch (err) {
    Logger.log("Error en doPost: " + err + " | Stack: " + err.stack);
    try { SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ChatLog").appendRow([new Date(),"ERROR","doPost",err.toString()]); } catch(e2) {}
  }
  return okResponse();
}
function okResponse() { return ContentService.createTextOutput("OK"); }

// ============================================
// LOGIN
// ============================================
function validarPasswordPanel(password) {
  try {
    const correcta = PropertiesService.getScriptProperties().getProperty("PANEL_PASSWORD");
    if (!correcta) return { ok: false, error: "Password no configurado (clave: PANEL_PASSWORD)" };
    if (password !== correcta) return { ok: false, error: "Contrasena incorrecta" };
    const token = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, password + "_" + new Date().toDateString()));
    PropertiesService.getScriptProperties().setProperty("panel_token_" + token, String(Date.now()));
    return { ok: true, token };
  } catch(err) { return { ok: false, error: err.toString() }; }
}
function validarTokenPanel(token) {
  try {
    if (!token) return { ok: false };
    const ts = PropertiesService.getScriptProperties().getProperty("panel_token_" + token);
    if (!ts) return { ok: false };
    if (Date.now() - parseInt(ts) > 8*60*60*1000) { PropertiesService.getScriptProperties().deleteProperty("panel_token_" + token); return { ok: false }; }
    return { ok: true };
  } catch(err) { return { ok: false }; }
}

// ============================================
// HELPERS DE CONVERSACION
// ============================================
function esUsuarioConocido(from) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LEADS_SHEET);
  if (!sheet) return false;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) { if (safeString(data[i][0]) === from) return true; }
  return false;
}
function obtenerNombreLead(from) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LEADS_SHEET);
  if (!sheet) return "";
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) { if (safeString(data[i][0]) === from) return safeString(data[i][1]); }
  return "";
}
function obtenerHandicapLead(from) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LEADS_SHEET);
  if (!sheet) return "";
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) { if (safeString(data[i][0]) === from) return safeString(data[i][3]); }
  return "";
}

// ============================================
// MENU PRINCIPAL
// ============================================
function enviarMenuPrincipal(from, nombre) {
  enviarMensajeWhatsApp(from,
    "\u00a1Hola de nuevo *" + nombre + "*! \u26f3 \u00bfCon qu\u00e9 te puedo ayudar?\n\n" +
    "1\ufe0f\u20e3 *Ejercicio gratis*\n" +
    "2\ufe0f\u20e3 *An\u00e1lisis de video* \u2014 $ 5.000\n" +
    "3\ufe0f\u20e3 *Plan personalizado* \u2014 $ 15.000\n" +
    "4\ufe0f\u20e3 *Actualizar mis datos*\n" +
    "5\ufe0f\u20e3 *Otras consultas*\n" +
    "6\ufe0f\u20e3 *Validar mi pago*"
  );
}

// ============================================
// MOTOR CONVERSACIONAL
// ============================================
function procesarMensajeEntrante(from, text) {
  const lock = LockService.getUserLock();
  if (!lock.tryLock(5000)) { Logger.log("Lock no obtenido: " + from); return; }
  try {
    const conv = obtenerConversacion(from);
    const paso = conv.paso || "inicio";
    const textLower = text.toLowerCase();

    if (textLower === "hola" || textLower === "inicio" || paso === "inicio") {
      if (esUsuarioConocido(from)) {
        const nombre = obtenerNombreLead(from);
        enviarMenuPrincipal(from, nombre);
        guardarConversacion(from, { ...conv, paso: "esperando_menu_principal", nombre });
      } else {
        enviarMensajeWhatsApp(from, "\u00a1Hola! \ud83c\udfcc\ufe0f Soy *Golfito*, tu coach de golf por WhatsApp.\n\n\u00bfCu\u00e1l es tu nombre?");
        guardarConversacion(from, { paso: "esperando_nombre", nombre: "", handicap: "", aspecto: "", ejvsplan: "", video_url1: "", video_url2: "" });
      }
      return;
    }

    if (paso === "completo" || paso === "plan_solicitado" || paso === "esperando_menu_principal") {
      if (paso === "completo") {
        if (text.toUpperCase() === "PLAN") {
          const nombre = conv.nombre || obtenerNombreLead(from);
          enviarMensajeWhatsApp(from, "Perfecto " + nombre + " \u26f3 Enviame un video de tu swing \ud83c\udfa5 _(menos de 7 segundos)_\n\n_(Si no ten\u00e9s uno, cualquier otra consulta escribinos \u26f3)_");
          guardarConversacion(from, { ...conv, paso: "esperando_video_plan", ejvsplan: "3" });
          return;
        }
        const nombre = conv.nombre || obtenerNombreLead(from);
        enviarMenuPrincipal(from, nombre);
        guardarConversacion(from, { ...conv, paso: "esperando_menu_principal" });
        return;
      }

      const v = text.trim();
      const nombre = conv.nombre || obtenerNombreLead(from);

      if (v === "1") {
        const handicap = conv.handicap || obtenerHandicapLead(from);
        if (conv.aspecto) {
          enviarMensajeWhatsApp(from, "Perfecto \u26f3 Estoy preparando tu ejercicio...");
          const datos = { ...conv, paso: "completo", ejvsplan: "1", nombre, handicap };
          guardarConversacion(from, datos); registrarSesion(from, datos);
        } else {
          enviarMensajeWhatsApp(from, "\u00bfQu\u00e9 aspecto quer\u00e9s trabajar?\n\n1\ufe0f\u20e3 Driver\n2\ufe0f\u20e3 Hierros\n3\ufe0f\u20e3 Approach\n4\ufe0f\u20e3 Putting\n5\ufe0f\u20e3 Bunker\n6\ufe0f\u20e3 Primera vez en el golf");
          guardarConversacion(from, { ...conv, paso: "esperando_aspecto_menu", ejvsplan: "1", nombre, handicap });
        }
      } else if (v === "2") {
        enviarMensajeWhatsApp(from, "Genial \ud83c\udfa5 Enviame un video de tu swing _(menos de 7 segundos)_ para analizarlo.");
        guardarConversacion(from, { ...conv, paso: "esperando_video_analisis", ejvsplan: "2", nombre, video_url1: "", video_url2: "" });
      } else if (v === "3") {
        const vids = obtenerUltimosVideosSesion(from);
        if (vids.url1 || vids.url2) {
          const cuantos = (vids.url1 && vids.url2) ? "dos videos" : "un video";
          enviarMensajeWhatsApp(from, "Perfecto " + nombre + " \u26f3 Veo que ya enviaste " + cuantos + " anteriormente.\n\n\u00bfUsamos esos para armar tu plan, o quer\u00e9s enviar nuevos?\n\n1\ufe0f\u20e3 Usar los videos que ya envi\u00e9\n2\ufe0f\u20e3 Enviar videos nuevos");
          guardarConversacion(from, { ...conv, paso: "esperando_reusar_videos", ejvsplan: "3", nombre, video_url1_prev: vids.url1, video_url2_prev: vids.url2 });
        } else {
          enviarMensajeWhatsApp(from, "Perfecto \u26f3 Para armar tu plan necesito videos de tu swing \ud83c\udfa5\n\nEnviame un video *de perfil* _(c\u00e1mara detr\u00e1s tuyo, menos de 7 segundos)_");
          guardarConversacion(from, { ...conv, paso: "esperando_video_plan_1", ejvsplan: "3", nombre, video_url1: "", video_url2: "" });
        }
      } else if (v === "4") {
        enviarMensajeWhatsApp(from, "\u00bfQu\u00e9 quer\u00e9s actualizar?\n\n1\ufe0f\u20e3 Mi nombre\n2\ufe0f\u20e3 Mi handicap");
        guardarConversacion(from, { ...conv, paso: "esperando_actualizar_datos" });
      } else if (v === "5") {
        enviarMensajeWhatsApp(from, "\u00a1Claro! Escrib\u00ed tu consulta o comentario y te respondemos a la brevedad \ud83d\udcdd");
        guardarConversacion(from, { ...conv, paso: "esperando_consulta" });
      } else if (v === "6") {
        enviarMensajeWhatsApp(from, "\u23f3 Verificando tu pago...");
        const externalRef = conv.mp_external_ref || conv.mp_codigo_plan || "";
        if (!externalRef) {
          enviarMensajeWhatsApp(from, "No encontramos ning\u00fan pago pendiente asociado a tu cuenta \u26f3");
          enviarMenuPrincipal(from, nombre); return;
        }
        const verificacion = verificarPagoAprobado(externalRef);
        if (verificacion.ok) {
          try {
            const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SESIONES_SHEET);
            if (sh) {
              const d = sh.getDataRange().getValues();
              for (let i = d.length - 1; i >= 1; i--) {
                if (safeString(d[i][COL.WHATSAPP-1]) === from && safeString(d[i][COL.PAGO_STATUS-1]) !== "pagado") {
                  sh.getRange(i+1, COL.PAGO_STATUS).setValue("pagado");
                  sh.getRange(i+1, COL.PAGO_FECHA).setValue(new Date());
                  sh.getRange(i+1, COL.MP_PAYMENT_ID).setValue(String(verificacion.paymentId));
                  if (safeString(d[i][COL.EJVSPLAN-1]) === "3") sh.getRange(i+1, COL.STATUS).setValue("pendiente_manual");
                  break;
                }
              }
            }
          } catch(se) { Logger.log("Error validando pago manual: " + se); }
          const ejvsplan = conv.ejvsplan || "";
          if (ejvsplan === "2") {
            enviarMensajeWhatsApp(from, "\u2705 Pago confirmado " + nombre + ". Analizando tu swing con IA, dame un momento...");
            const datosAnal = { ...conv, paso: "analizando_video" };
            guardarConversacion(from, datosAnal); procesarAnalisisVideo(from, datosAnal, false);
          } else if (ejvsplan === "3") {
            enviarMensajeWhatsApp(from, "\u2705 Pago confirmado " + nombre + " \u26f3 Estamos preparando tu plan y te lo enviamos pronto por ac\u00e1.");
            const perfil = { nivel: mapNivel(conv.handicap||"", mapAspectoLead(conv.aspecto||"")), aspecto: mapAspectoLead(conv.aspecto||""), tiempo: "45 minutos" };
            notificarNuevoPlan(nombre, from, perfil, conv.video_url1||"", conv.comentarios_alumno||"");
            guardarConversacion(from, { ...conv, paso: "completo" });
          } else {
            enviarMensajeWhatsApp(from, "\u2705 Pago confirmado " + nombre + " \u26f3");
            guardarConversacion(from, { ...conv, paso: "completo" }); enviarMenuPrincipal(from, nombre);
          }
        } else {
          enviarMensajeWhatsApp(from, "No encontramos un pago aprobado todav\u00eda. Si ya pagaste, esper\u00e1 unos minutos e intent\u00e1 de nuevo \u26f3");
          enviarMenuPrincipal(from, nombre);
        }
      } else {
        enviarMenuPrincipal(from, nombre);
      }
      return;
    }

    if (paso === "esperando_aspecto_menu") {
      enviarMensajeWhatsApp(from, "Perfecto \u26f3 Estoy preparando tu ejercicio...");
      const datos = { ...conv, paso: "completo", aspecto: text };
      guardarConversacion(from, datos); registrarSesion(from, datos); return;
    }

    if (paso === "esperando_actualizar_datos") {
      if (text === "1") {
        enviarMensajeWhatsApp(from, "\u00bfCu\u00e1l es tu nombre?");
        guardarConversacion(from, { ...conv, paso: "actualizando_nombre" });
      } else if (text === "2") {
        enviarMensajeWhatsApp(from, "\u00bfCu\u00e1l es tu handicap actual?\n\n_(Si est\u00e1s empezando, escrib\u00ed *no tengo*)_");
        guardarConversacion(from, { ...conv, paso: "actualizando_handicap" });
      } else {
        enviarMensajeWhatsApp(from, "Respond\u00e9 1 para nombre o 2 para handicap.");
      }
      return;
    }
    if (paso === "actualizando_nombre") {
      const nombre = sanitizarNombre(text);
      const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LEADS_SHEET);
      if (ss) { const data = ss.getDataRange().getValues(); for (let i=1;i<data.length;i++) { if (safeString(data[i][0])===from) { ss.getRange(i+1,2).setValue(nombre); break; } } }
      enviarMensajeWhatsApp(from, "\u2705 Nombre actualizado a *" + nombre + "*.\n\nCualquier otra consulta escribinos \u26f3");
      enviarMenuPrincipal(from, nombre); guardarConversacion(from, { ...conv, paso: "esperando_menu_principal", nombre }); return;
    }
    if (paso === "actualizando_handicap") {
      actualizarHandicapLead(from, text);
      const nombre = conv.nombre || obtenerNombreLead(from);
      enviarMensajeWhatsApp(from, "\u2705 Handicap actualizado a *" + text + "*.\n\nCualquier otra consulta escribinos \u26f3");
      enviarMenuPrincipal(from, nombre); guardarConversacion(from, { ...conv, paso: "esperando_menu_principal", handicap: text }); return;
    }
    if (paso === "esperando_consulta") {
      guardarConsulta(from, conv.nombre || obtenerNombreLead(from), text);
      enviarMensajeWhatsApp(from, "\u00a1Gracias por tu consulta! Te respondemos a la brevedad \u26f3");
      guardarConversacion(from, { ...conv, paso: "completo" }); return;
    }
    if (paso === "esperando_nombre") {
      const nombre = sanitizarNombre(text);
      enviarMensajeWhatsApp(from, "Hola *" + nombre + "* \ud83d\udc4b\n\n\u00bfCu\u00e1l es tu handicap?\n\n_(Si est\u00e1s empezando, escrib\u00ed *no tengo*)_");
      guardarConversacion(from, { ...conv, paso: "esperando_handicap", nombre }); return;
    }
    if (paso === "esperando_handicap") {
      enviarMenuPrincipal(from, conv.nombre);
      guardarConversacion(from, { ...conv, paso: "esperando_menu_principal", handicap: text }); return;
    }
    if (paso === "esperando_aspecto") {
      enviarMensajeWhatsApp(from, "\u00bfQu\u00e9 quer\u00e9s hacer?\n\n1\ufe0f\u20e3 *Ejercicio gratis* \u2014 te mando uno ahora\n2\ufe0f\u20e3 *An\u00e1lisis de video* \u2014 $ 5.000\n3\ufe0f\u20e3 *Plan personalizado* \u2014 $ 15.000");
      guardarConversacion(from, { ...conv, paso: "esperando_ejvsplan", aspecto: text }); return;
    }
    if (paso === "esperando_ejvsplan") {
      const tipo = mapTipoSolicitud(text);
      if (tipo === "ejercicio_gratis") { enviarMensajeWhatsApp(from, "Perfecto \u26f3 Estoy preparando tu ejercicio..."); const datos = { ...conv, paso: "completo", ejvsplan: "1" }; guardarConversacion(from, datos); registrarSesion(from, datos); }
      else if (tipo === "analisis_video") { enviarMensajeWhatsApp(from, "Genial \ud83c\udfa5 Enviame un video de tu swing _(menos de 7 segundos)_ para analizarlo."); guardarConversacion(from, { ...conv, paso: "esperando_video_analisis", ejvsplan: "2", video_url1: "", video_url2: "" }); }
      else if (tipo === "plan_personalizado") { enviarMensajeWhatsApp(from, "Perfecto \u26f3 Enviame un video de tu swing \ud83c\udfa5 _(menos de 7 segundos)_\n\n_(Si no ten\u00e9s uno a mano, cualquier otra consulta escribinos \u26f3)_"); guardarConversacion(from, { ...conv, paso: "esperando_video_plan", ejvsplan: "3" }); }
      return;
    }
    if (paso === "esperando_video_analisis") { enviarMensajeWhatsApp(from, "Para enviar el video us\u00e1 el clip \ud83d\udcce de WhatsApp."); return; }

    if (paso === "esperando_pago_analisis") {
      const nombre = conv.nombre || obtenerNombreLead(from);
      enviarMensajeWhatsApp(from, "\u23f3 Verificando tu pago...");
      const externalRef = conv.mp_external_ref || conv.mp_codigo_plan || "";
      const verificacion = verificarPagoAprobado(externalRef);
      if (verificacion.ok) {
        try {
          const sh2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SESIONES_SHEET);
          if (sh2) {
            const d2 = sh2.getDataRange().getValues();
            for (let i2 = d2.length-1; i2 >= 1; i2--) {
              if (safeString(d2[i2][COL.WHATSAPP-1]) === from && safeString(d2[i2][COL.PAGO_STATUS-1]) !== "pagado") {
                sh2.getRange(i2+1,COL.PAGO_STATUS).setValue("pagado"); sh2.getRange(i2+1,COL.PAGO_FECHA).setValue(new Date()); sh2.getRange(i2+1,COL.MP_PAYMENT_ID).setValue(String(verificacion.paymentId)); break;
              }
            }
          }
        } catch(se) { Logger.log("Error guardando pago: " + se); }
        enviarMensajeWhatsApp(from, "\u2705 Pago confirmado " + nombre + ". Analizando tu swing con IA, dame un momento...");
        const datosAnal = { ...conv, paso: "analizando_video" }; guardarConversacion(from, datosAnal); procesarAnalisisVideo(from, datosAnal, false);
      } else if (verificacion.motivo === "no_encontrado" || verificacion.motivo === "pending" || verificacion.motivo === "in_process") {
        enviarMensajeWhatsApp(from, "\u23f3 Todav\u00eda no veo el pago confirmado " + nombre + ". Esper\u00e1 unos minutos y us\u00e1 la opci\u00f3n *6 \u2014 Validar mi pago* del men\u00fa \u26f3");
        guardarConversacion(from, { ...conv, paso: "esperando_menu_principal" }); enviarMenuPrincipal(from, nombre);
      } else {
        enviarMensajeWhatsApp(from, "No encontramos un pago aprobado. Si ya pagaste, esper\u00e1 unos minutos y us\u00e1 la opci\u00f3n *6 \u2014 Validar mi pago* del men\u00fa \u26f3");
        guardarConversacion(from, { ...conv, paso: "esperando_menu_principal" }); enviarMenuPrincipal(from, nombre);
      }
      return;
    }

    if (paso === "esperando_segundo_video") {
      const nombre = conv.nombre || obtenerNombreLead(from);
      guardarConversacion(from, { ...conv, paso: "esperando_menu_principal" }); enviarMenuPrincipal(from, nombre); return;
    }
    if (paso === "esperando_reusar_videos") {
      if (text === "1") {
        const url1 = conv.video_url1_prev || ""; const url2 = conv.video_url2_prev || "";
        if (url1 && url2) { enviarMensajeWhatsApp(from, "Perfecto \u26f3 \u00bfHay algo espec\u00edfico que quer\u00e9s mejorar o en lo que quer\u00e9s enfocarte? _(opcional \u2014 pod\u00e9s escribir *saltar*)_"); guardarConversacion(from, { ...conv, paso: "esperando_comentarios_alumno", video_url1: url1, video_url2: url2 }); }
        else { const tieneUrl1 = !!url1; enviarMensajeWhatsApp(from, "Genial, uso el video que ya enviaste \u2705\n\nPara un an\u00e1lisis m\u00e1s completo, \u00bfme envi\u00e1s un video desde el " + (tieneUrl1 ? "frente" : "costado trasero") + "? _(menos de 7 segundos)_\n\nSi no ten\u00e9s, escrib\u00ed *saltar*"); guardarConversacion(from, { ...conv, paso: "esperando_video_plan_complementario", video_url1: url1||"", video_url2: url2||"" }); }
      } else if (text === "2") { enviarMensajeWhatsApp(from, "Buen\u00edsimo \u26f3 Enviame el primer video de tu swing \ud83c\udfa5\n_(c\u00e1mara trasera \u2014 detr\u00e1s tuyo, menos de 7 segundos)_"); guardarConversacion(from, { ...conv, paso: "esperando_video_plan_1", video_url1: "", video_url2: "" }); }
      else { enviarMensajeWhatsApp(from, "Respond\u00e9 *1* para usar los videos anteriores o *2* para enviar nuevos."); }
      return;
    }
    if (paso === "esperando_video_plan_complementario") {
      if (textLower === "saltar") { enviarMensajeWhatsApp(from, "Ok \u26f3 \u00bfHay algo espec\u00edfico que quer\u00e9s mejorar o en lo que quer\u00e9s enfocarte? _(opcional \u2014 pod\u00e9s escribir *saltar*)_"); guardarConversacion(from, { ...conv, paso: "esperando_comentarios_alumno" }); }
      else { enviarMensajeWhatsApp(from, "Enviam\u00e9 el video usando el clip \ud83d\udcce de WhatsApp _(menos de 7 segundos)_, o escrib\u00ed *saltar* para continuar sin \u00e9l."); }
      return;
    }
    if (paso === "esperando_video_plan_1") { enviarMensajeWhatsApp(from, "Enviam\u00e9 el video usando el clip \ud83d\udcce de WhatsApp _(menos de 7 segundos)_."); return; }
    if (paso === "esperando_video_plan_2") {
      if (textLower === "saltar") { enviarMensajeWhatsApp(from, "Ok \u26f3 \u00bfHay algo espec\u00edfico que quer\u00e9s mejorar o en lo que quer\u00e9s enfocarte? _(opcional \u2014 pod\u00e9s escribir *saltar*)_"); guardarConversacion(from, { ...conv, paso: "esperando_comentarios_alumno" }); }
      else { enviarMensajeWhatsApp(from, "Enviam\u00e9 el segundo video usando el clip \ud83d\udcce _(menos de 7 segundos)_, o escrib\u00ed *saltar* para continuar."); }
      return;
    }

    if (paso === "esperando_comentarios_alumno") {
      const comentarios = textLower === "saltar" ? "" : text.trim();
      const nombre = conv.nombre || obtenerNombreLead(from);
      if (MODO_TEST_PLAN) {
        Logger.log("MODO_TEST_PLAN activo — saltando pago en esperando_comentarios_alumno");
        const datos = { ...conv, paso: "completo", comentarios_alumno: comentarios };
        guardarConversacion(from, datos);
        registrarSesion(from, datos);
        const perfil = { nivel: mapNivel(conv.handicap||"", mapAspectoLead(conv.aspecto||"")), aspecto: mapAspectoLead(conv.aspecto||""), tiempo: "45 minutos" };
        notificarNuevoPlan(nombre, from, perfil, conv.video_url1||"", comentarios);
        enviarMensajeWhatsApp(from, "Perfecto " + nombre + " \u26f3 Estamos preparando tu plan y te lo enviamos pronto por ac\u00e1.");
      } else {
        const codigoPlan = generarCodigoPlan();
        const mpRes = crearPreferenciaPago(from, nombre, "plan", codigoPlan);
        if (mpRes.ok) {
          enviarMensajeWhatsApp(from, "Perfecto " + nombre + " \u26f3 Para confirmar tu plan, realiz\u00e1 el pago de *$ 15.000* ac\u00e1:\n" + mpRes.link + "\n\nUna vez confirmado el pago te avisamos y empezamos con tu plan \ud83c\udfcc\ufe0f");
          const datosConPago = { ...conv, paso: "esperando_pago_plan", comentarios_alumno: comentarios, mp_codigo_plan: codigoPlan };
          guardarConversacion(from, datosConPago);
          registrarSesion(from, { ...datosConPago, ejvsplan: "3" });
        } else {
          Logger.log("MP fallo: " + (mpRes.error||""));
          const datos = { ...conv, paso: "completo", comentarios_alumno: comentarios };
          guardarConversacion(from, datos); registrarSesion(from, datos);
        }
      }
      return;
    }

    if (paso === "esperando_video_plan") {
      if (textLower === "saltar") {
        if (MODO_TEST_PLAN) { enviarMensajeWhatsApp(from, "Ok \u26f3 \u00bfHay algo espec\u00edfico que quer\u00e9s mejorar o en lo que quer\u00e9s enfocarte? _(opcional \u2014 pod\u00e9s escribir *saltar*)_"); guardarConversacion(from, { ...conv, paso: "esperando_comentarios_alumno", video_url1: "" }); }
        else { const cod = safeString(obtenerCodigoPlanPendiente(from)) || generarCodigoPlan(); const mpR = crearPreferenciaPago(from, conv.nombre||obtenerNombreLead(from), "plan", cod); const lnk = mpR.ok ? mpR.link : "https://mpago.la/TU-LINK-PLAN"; enviarMensajeWhatsApp(from, "Para continuar, realiz\u00e1 el pago de *$ 15.000* ac\u00e1:\n" + lnk + "\n\nCualquier otra consulta escribinos \u26f3"); guardarConversacion(from, { ...conv, paso: "esperando_pago_plan", mp_codigo_plan: cod }); }
      } else { enviarMensajeWhatsApp(from, "Para enviar el video us\u00e1 el clip \ud83d\udcce de WhatsApp _(menos de 7 segundos)_, o cualquier otra consulta escribinos \u26f3"); }
      return;
    }

    if (paso === "esperando_pago_plan") {
      const nombrePlan = conv.nombre || obtenerNombreLead(from);
      enviarMensajeWhatsApp(from, "\u23f3 Verificando tu pago...");
      const vPlan = verificarPagoAprobado(conv.mp_codigo_plan || "");
      if (vPlan.ok) {
        try {
          const sh3 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SESIONES_SHEET);
          if (sh3) { const d3 = sh3.getDataRange().getValues(); for (let i3=d3.length-1;i3>=1;i3--) { if (safeString(d3[i3][COL.WHATSAPP-1])===from && safeString(d3[i3][COL.PAGO_STATUS-1])!=="pagado") { sh3.getRange(i3+1,COL.PAGO_STATUS).setValue("pagado"); sh3.getRange(i3+1,COL.PAGO_FECHA).setValue(new Date()); sh3.getRange(i3+1,COL.MP_PAYMENT_ID).setValue(String(vPlan.paymentId)); sh3.getRange(i3+1,COL.STATUS).setValue("pendiente_manual"); break; } } }
        } catch(se2) { Logger.log("Error guardando pago plan: " + se2); }
        enviarMensajeWhatsApp(from, "\u2705 Pago confirmado " + nombrePlan + " \u26f3 Estamos preparando tu plan personalizado y te lo enviamos pronto por ac\u00e1.");
        guardarConversacion(from, { ...conv, paso: "completo" });
        const perfil = { nivel: mapNivel(conv.handicap||"", mapAspectoLead(conv.aspecto||"")), aspecto: mapAspectoLead(conv.aspecto||""), tiempo: "45 minutos" };
        notificarNuevoPlan(nombrePlan, from, perfil, conv.video_url1||"", conv.comentarios_alumno||"");
      } else if (vPlan.motivo === "no_encontrado" || vPlan.motivo === "pending" || vPlan.motivo === "in_process") {
        enviarMensajeWhatsApp(from, "\u23f3 Todav\u00eda no veo el pago confirmado " + nombrePlan + ". Esper\u00e1 unos minutos y us\u00e1 la opci\u00f3n *6 \u2014 Validar mi pago* del men\u00fa \u26f3");
        guardarConversacion(from, { ...conv, paso: "esperando_menu_principal" }); enviarMenuPrincipal(from, nombrePlan);
      } else {
        enviarMensajeWhatsApp(from, "No encontramos un pago aprobado. Si ya pagaste, esper\u00e1 unos minutos y us\u00e1 la opci\u00f3n *6 \u2014 Validar mi pago* del men\u00fa \u26f3");
        guardarConversacion(from, { ...conv, paso: "esperando_menu_principal" }); enviarMenuPrincipal(from, nombrePlan);
      }
      return;
    }

    if (paso === "esperando_feedback") {
      const score = parseInt(text.trim(), 10);
      if (!isNaN(score) && score >= 1 && score <= 5) {
        guardarFeedbackScore(from, score);
        enviarMensajeWhatsApp(from, "\u00a1Gracias " + (conv.nombre||"") + "! \ud83d\ude4f\n\n\u00bfTen\u00e9s alg\u00fan comentario o sugerencia para nosotros? \u26f3 _(opcional \u2014 pod\u00e9s escribir *saltar*)_");
        guardarConversacion(from, { ...conv, paso: "esperando_comentario_feedback", feedback_score: score });
      } else { enviarMensajeWhatsApp(from, "Por favor respond\u00e9 con un n\u00famero del 1 al 5."); }
      return;
    }
    if (paso === "esperando_comentario_feedback") {
      const comentario = text.toLowerCase() === "saltar" ? "" : text.trim();
      if (comentario) guardarFeedback(from, comentario);
      enviarMensajeWhatsApp(from, "Muchas gracias " + (conv.nombre||"") + " \ud83d\ude4f Tu opini\u00f3n nos ayuda a mejorar \u26f3");
      guardarConversacion(from, { ...conv, paso: "completo" });
      return;
    }

    const nombreFallback = conv.nombre || obtenerNombreLead(from);
    if (nombreFallback) { enviarMenuPrincipal(from, nombreFallback); guardarConversacion(from, { ...conv, paso: "esperando_menu_principal" }); }
    else { enviarMensajeWhatsApp(from, "Cualquier otra consulta escribinos \u26f3"); }

  } finally { lock.releaseLock(); }
}

// ============================================
// PROCESAR VIDEO ENTRANTE
// ============================================
function procesarVideoEntrante(from, mediaId) {
  const lock = LockService.getUserLock();
  if (!lock.tryLock(5000)) { Logger.log("Lock no obtenido para video: " + from); return; }
  try {
    const conv = obtenerConversacion(from);
    const paso = conv.paso || "";
    const pasosValidos = ["esperando_video_analisis","esperando_video_plan","esperando_video_plan_1","esperando_video_plan_2","esperando_video_plan_complementario"];
    if (!pasosValidos.includes(paso)) return;
    try {
      const metaRes = UrlFetchApp.fetch("https://graph.facebook.com/v19.0/" + mediaId, { headers: { "Authorization": "Bearer " + META_TOKEN } });
      const mediaData = JSON.parse(metaRes.getContentText());
      const videoRes = UrlFetchApp.fetch(mediaData.url, { headers: { "Authorization": "Bearer " + META_TOKEN } });
      const videoBlob = videoRes.getBlob().setName("swing_" + from + "_" + Date.now() + ".mp4");
      let folder; const folders = DriveApp.getFoldersByName("Golfito_Videos");
      folder = folders.hasNext() ? folders.next() : DriveApp.createFolder("Golfito_Videos");
      const file = folder.createFile(videoBlob);
      const driveUrl = file.getUrl();
      logMensaje(from, "entrante", "video_drive", driveUrl);

      if (paso === "esperando_video_analisis") {
        const nombreAnal = conv.nombre || obtenerNombreLead(from);
        if (MODO_TEST_ANALISIS) {
          enviarMensajeWhatsApp(from, "\u23f3 Recib\u00ed tu video " + nombreAnal + ". Analizando tu swing con IA, dame un momento...");
          const datos = { ...conv, paso: "analizando_video", video_url1: driveUrl, ejvsplan: "2" };
          // FIX BUG 2: registrar sesion ANTES del analisis para que el video quede en Sheets
          registrarSesion(from, datos);
          guardarConversacion(from, datos);
          procesarAnalisisVideo(from, datos, false);
        } else {
          const codigoAnal = "ANAL-" + String(Date.now()).slice(-6);
          const mpResAnal = crearPreferenciaPago(from, nombreAnal, "analisis", codigoAnal);
          if (mpResAnal.ok) {
            enviarMensajeWhatsApp(from, "Recib\u00ed tu video " + nombreAnal + " \u2705\n\nPara analizar tu swing realiz\u00e1 el pago de *$ 5.000* ac\u00e1:\n" + mpResAnal.link + "\n\nUna vez que pagues, escribinos ac\u00e1 y verificamos el pago \u26f3");
            const datosAnal = { ...conv, paso: "esperando_pago_analisis", video_url1: driveUrl, mp_external_ref: mpResAnal.externalRef || codigoAnal, ejvsplan: "2" };
            guardarConversacion(from, datosAnal);
            registrarSesion(from, datosAnal);
          } else {
            Logger.log("MP fallo en analisis, fallback directo: " + (mpResAnal.error||""));
            enviarMensajeWhatsApp(from, "\u23f3 Recib\u00ed tu video " + nombreAnal + ". Analizando tu swing con IA, dame un momento...");
            const datos = { ...conv, paso: "analizando_video", video_url1: driveUrl, ejvsplan: "2" };
            registrarSesion(from, datos);
            guardarConversacion(from, datos);
            procesarAnalisisVideo(from, datos, false);
          }
        }
      } else if (paso === "esperando_video_plan_1") {
        enviarMensajeWhatsApp(from, "Recib\u00ed el primer video \u2705\n\nAhora enviam\u00e9 un segundo video desde el *frente* \ud83c\udfa5 _(menos de 7 segundos)_\n\nSi no ten\u00e9s, escrib\u00ed *saltar*");
        guardarConversacion(from, { ...conv, paso: "esperando_video_plan_2", video_url1: driveUrl });
      } else if (paso === "esperando_video_plan_2") {
        enviarMensajeWhatsApp(from, "Recib\u00ed los dos videos \u2705\n\n\u00bfHay algo espec\u00edfico que quer\u00e9s mejorar o en lo que quer\u00e9s enfocarte? _(opcional \u2014 pod\u00e9s escribir *saltar*)_");
        guardarConversacion(from, { ...conv, paso: "esperando_comentarios_alumno", video_url2: driveUrl });
      } else if (paso === "esperando_video_plan_complementario") {
        const yaTeníaUrl1 = !!conv.video_url1;
        const datos = yaTeníaUrl1 ? { ...conv, video_url2: driveUrl } : { ...conv, video_url1: driveUrl };
        enviarMensajeWhatsApp(from, "Recib\u00ed el video complementario \u2705\n\n\u00bfHay algo espec\u00edfico que quer\u00e9s mejorar o en lo que quer\u00e9s enfocarte? _(opcional \u2014 pod\u00e9s escribir *saltar*)_");
        guardarConversacion(from, { ...datos, paso: "esperando_comentarios_alumno" });
      } else if (paso === "esperando_video_plan") {
        if (MODO_TEST_PLAN) {
          enviarMensajeWhatsApp(from, "Recib\u00ed tu video " + conv.nombre + " \u2705\n\n\u00bfHay algo espec\u00edfico que quer\u00e9s mejorar o en lo que quer\u00e9s enfocarte? _(opcional \u2014 pod\u00e9s escribir *saltar*)_");
          guardarConversacion(from, { ...conv, paso: "esperando_comentarios_alumno", video_url1: driveUrl });
        } else {
          const cod2 = safeString(obtenerCodigoPlanPendiente(from)) || generarCodigoPlan();
          const mpRes2 = crearPreferenciaPago(from, conv.nombre||obtenerNombreLead(from), "plan", cod2);
          const lnk2 = mpRes2.ok ? mpRes2.link : "https://mpago.la/TU-LINK-PLAN";
          enviarMensajeWhatsApp(from, "Recib\u00ed tu video " + conv.nombre + " \u2705\n\nPara continuar realiz\u00e1 el pago de *$ 15.000* ac\u00e1:\n" + lnk2 + "\n\nCualquier otra consulta escribinos \u26f3");
          guardarConversacion(from, { ...conv, paso: "esperando_pago_plan", video_url1: driveUrl, mp_codigo_plan: cod2 });
        }
      }
    } catch (err) {
      const errorMsg = "Error procesando video: " + err.toString();
      Logger.log(errorMsg); logMensaje(from, "error", "video_error", errorMsg);
      const mensajeError = err.toString().includes("unavailable") || err.toString().includes("fbsbx")
        ? "El link del video expir\u00f3 \u23f1 Por favor reenviam\u00e9 el video e intentamos de nuevo."
        : "Hubo un problema al recibir el video. Pod\u00e9s intentar de nuevo.";
      enviarMensajeWhatsApp(from, mensajeError);
    }
  } finally { lock.releaseLock(); }
}

// ============================================
// PROCESAR ANALISIS DE VIDEO CON GEMINI
// ============================================
function procesarAnalisisVideo(from, conv, esSegundoVideo) {
  try {
    const driveUrl = esSegundoVideo ? conv.video_url2 : conv.video_url1;
    if (!driveUrl) throw new Error("No hay video para analizar");
    const match = driveUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
    if (!match) throw new Error("No se pudo obtener el ID del video");
    const resultado = analizarSwingConGemini(match[1]);
    if (!resultado.ok) throw new Error(resultado.error);
    const a = resultado.analisis;
    const titulo = esSegundoVideo ? "\ud83e\udd16 *An\u00e1lisis complementario (" + a.angulo + ")*" : "\ud83e\udd16 *An\u00e1lisis de tu swing (" + a.angulo + ")*";
    enviarMensajeWhatsApp(from, titulo + "\n\n\ud83c\udfaf *Error principal:* " + a.error_principal + "\n\ud83d\udcca *Severidad:* " + a.severidad + "\n\n\ud83d\udcdd *Detalles:*\n" + a.detalles + "\n\n\ud83d\udca1 *Recomendaci\u00f3n:*\n" + a.recomendacion);
    const nombre = conv.nombre || obtenerNombreLead(from);
    if (!esSegundoVideo) {
      enviarMenuPrincipal(from, nombre);
      const datos = { ...conv, paso: "esperando_menu_principal", analisis1: a };
      guardarConversacion(from, datos);
      if (safeString(conv.ejvsplan) === "2") actualizarSesionAnalisis(from, a, null);
    } else {
      enviarMenuPrincipal(from, nombre);
      const datos = { ...conv, paso: "esperando_menu_principal", analisis2: a };
      guardarConversacion(from, datos);
      if (safeString(conv.ejvsplan) !== "3") actualizarSesionAnalisis(from, conv.analisis1 || null, a);
    }
  } catch (err) {
    Logger.log("Error procesarAnalisisVideo: " + err);
    logMensaje(from, "error", "analisis_error", err.toString());
    enviarMensajeWhatsApp(from, "Estamos experimentando alta demanda en este momento. Intent\u00e1 de nuevo en unos minutos \u26f3");
    const nombre = conv.nombre || obtenerNombreLead(from);
    const datos = { ...conv, paso: "esperando_menu_principal" };
    guardarConversacion(from, datos); enviarMenuPrincipal(from, nombre);
  }
}

function guardarFeedbackScore(from, score) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SESIONES_SHEET);
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  for (let i = data.length-1; i >= 1; i--) {
    if (safeString(data[i][COL.WHATSAPP-1]) === from && safeString(data[i][COL.EJVSPLAN-1]) === "3") {
      sheet.getRange(i+1,COL.FEEDBACK_SCORE).setValue(score); sheet.getRange(i+1,COL.FEEDBACK_ALUMNO).setValue(String(score)); sheet.getRange(i+1,COL.FECHA_FEEDBACK).setValue(new Date()); return;
    }
  }
}
function guardarFeedback(from, texto) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SESIONES_SHEET);
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  for (let i = data.length-1; i >= 1; i--) {
    if (safeString(data[i][COL.WHATSAPP-1]) === from && safeString(data[i][COL.EJVSPLAN-1]) === "3") {
      sheet.getRange(i+1,COL.FEEDBACK_ALUMNO).setValue(texto); sheet.getRange(i+1,COL.FECHA_FEEDBACK).setValue(new Date()); return;
    }
  }
}

function enviarFeedbackPendiente() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SESIONES_SHEET); const leadsSheet = ss.getSheetByName(LEADS_SHEET);
  if (!sheet || !leadsSheet) return;
  const data = sheet.getDataRange().getValues(); const leads = leadsSheet.getDataRange().getValues();
  const ahora = new Date(); const dosHoras = 2*60*60*1000;
  for (let i = 1; i < data.length; i++) {
    const ejvsplan = safeString(data[i][COL.EJVSPLAN-1]); const status = safeString(data[i][COL.STATUS-1]).toLowerCase();
    const feedback = safeString(data[i][COL.FEEDBACK_ALUMNO-1]); const fechaFeedback = data[i][COL.FECHA_FEEDBACK-1];
    const timestamp = data[i][COL.TIMESTAMP-1]; const whatsapp = safeString(data[i][COL.WHATSAPP-1]);
    if (ejvsplan !== "3" || status !== "enviado" || feedback || fechaFeedback || !timestamp) continue;
    if (ahora - new Date(timestamp) < dosHoras) continue;
    let nombre = "";
    for (let j = 1; j < leads.length; j++) { if (safeString(leads[j][0]) === whatsapp) { nombre = safeString(leads[j][1]); break; } }
    sheet.getRange(i+1, COL.FECHA_FEEDBACK).setValue(ahora);
    const conv = obtenerConversacion(whatsapp);
    guardarConversacion(whatsapp, { ...conv, paso: "esperando_feedback" });
    enviarMensajeWhatsApp(whatsapp, "\u00a1Hola " + nombre + "! \ud83c\udfcc\ufe0f Del 1 al 5, \u00bfqu\u00e9 tan \u00fatil fue tu plan?\n\n1 = poco \u00fatil \u00b7 5 = muy \u00fatil");
  }
}

function guardarConsulta(from, nombre, texto) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONSULTAS_SHEET);
  if (!sheet) { sheet = ss.insertSheet(CONSULTAS_SHEET); sheet.appendRow(["timestamp","whatsapp","nombre","consulta","respondida"]); }
  sheet.appendRow([new Date(), from, nombre, texto, "no"]);
  try { GmailApp.sendEmail(Session.getActiveUser().getEmail(), "\ud83d\udcac Nueva consulta de " + nombre, "WhatsApp: " + from + "\nNombre: " + nombre + "\n\nConsulta:\n" + texto); } catch(err) { Logger.log("Error email consulta: " + err); }
}

function actualizarHandicapLead(from, handicap) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LEADS_SHEET);
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) { if (safeString(data[i][0]) === from) { sheet.getRange(i+1,4).setValue(handicap); return; } }
}

function logMensaje(from, direccion, tipo, contenido) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("ChatLog");
  if (!sheet) { sheet = ss.insertSheet("ChatLog"); sheet.appendRow(["timestamp","whatsapp","direccion","tipo","contenido"]); }
  sheet.appendRow([new Date(), from, direccion, tipo, contenido]);
}

function registrarOActualizarLead(from, nombre, handicap) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(LEADS_SHEET);
  if (!sheet) { sheet = ss.insertSheet(LEADS_SHEET); sheet.appendRow(["whatsapp","nombre","fecha_registro","handicap","notas"]); }
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) { if (safeString(data[i][0]) === from) { sheet.getRange(i+1,4).setValue(handicap); return; } }
  sheet.appendRow([from, nombre, new Date(), handicap, ""]);
}

function registrarSesion(from, conv, analisis1, analisis2) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SESIONES_SHEET);
  if (!sheet) { sheet = ss.insertSheet(SESIONES_SHEET); sheet.appendRow(SESIONES_HEADERS); }
  const aspecto = mapAspectoLead(conv.aspecto||""); const nivel = mapNivel(conv.handicap||"", aspecto);
  const ejvsplan = conv.ejvsplan||"1"; const a1 = analisis1||conv.analisis1||null; const a2 = analisis2||conv.analisis2||null;
  const codigoPlan = ejvsplan==="3" ? generarCodigoPlan() : ""; const numPlanAlumno = ejvsplan==="3" ? contarPlanesAlumno(from,sheet)+1 : "";
  const pagoStatus = (MODO_TEST_PLAN || MODO_TEST_ANALISIS) ? "pagado" : (ejvsplan==="1" ? "gratis" : "pendiente");
  const pagoMonto = ejvsplan==="1" ? 0 : ejvsplan==="2" ? 5 : 15;
  const row = new Array(SESIONES_HEADERS.length).fill("");
  row[COL.TIMESTAMP-1]=new Date(); row[COL.WHATSAPP-1]=from; row[COL.ASPECTO-1]=aspecto; row[COL.NIVEL-1]=nivel;
  row[COL.EJVSPLAN-1]=ejvsplan; row[COL.STATUS-1]=""; row[COL.VIDEO_URL1-1]=conv.video_url1||"";
  row[COL.ANALISIS_ERROR_IA1-1]=a1?.error_principal||""; row[COL.ANALISIS_AREA_IA1-1]=a1?.area||"";
  row[COL.ANALISIS_SEVERIDAD_IA1-1]=a1?.severidad||""; row[COL.ANALISIS_RECOMENDACION_IA1-1]=a1?.recomendacion||"";
  row[COL.ANALISIS_ANGULO_IA1-1]=a1?.angulo||""; row[COL.VIDEO_URL2-1]=conv.video_url2||"";
  row[COL.ANALISIS_ERROR_IA2-1]=a2?.error_principal||""; row[COL.ANALISIS_AREA_IA2-1]=a2?.area||"";
  row[COL.ANALISIS_SEVERIDAD_IA2-1]=a2?.severidad||""; row[COL.ANALISIS_RECOMENDACION_IA2-1]=a2?.recomendacion||"";
  row[COL.ANALISIS_ANGULO_IA2-1]=a2?.angulo||""; row[COL.PAGO_STATUS-1]=pagoStatus; row[COL.PAGO_MONTO-1]=pagoMonto;
  row[COL.PAGO_FECHA-1]=(MODO_TEST_PLAN || MODO_TEST_ANALISIS) ? new Date() : ""; row[COL.CODIGO_PLAN-1]=codigoPlan; row[COL.NUM_PLAN_ALUMNO-1]=numPlanAlumno;
  row[COL.COMENTARIOS_ALUMNO-1]=conv.comentarios_alumno||"";
  sheet.appendRow(row);
  registrarOActualizarLead(from, conv.nombre||"", conv.handicap||"");
  Utilities.sleep(1000); procesarSesionesPendientes();
}

function generarCodigoPlan() { return "PLAN-" + new Date().getFullYear() + "-" + String(Date.now()).slice(-6); }

function actualizarSesionAnalisis(from, analisis1, analisis2) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SESIONES_SHEET);
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (safeString(data[i][COL.WHATSAPP-1]) !== from) continue;
      if (safeString(data[i][COL.EJVSPLAN-1]) !== "2") continue;
      const rowNum = i + 1;
      if (analisis1) {
        sheet.getRange(rowNum, COL.ANALISIS_ERROR_IA1).setValue(analisis1.error_principal || "");
        sheet.getRange(rowNum, COL.ANALISIS_AREA_IA1).setValue(analisis1.area || "");
        sheet.getRange(rowNum, COL.ANALISIS_SEVERIDAD_IA1).setValue(analisis1.severidad || "");
        sheet.getRange(rowNum, COL.ANALISIS_RECOMENDACION_IA1).setValue(analisis1.recomendacion || "");
        sheet.getRange(rowNum, COL.ANALISIS_ANGULO_IA1).setValue(analisis1.angulo || "");
      }
      if (analisis2) {
        sheet.getRange(rowNum, COL.ANALISIS_ERROR_IA2).setValue(analisis2.error_principal || "");
        sheet.getRange(rowNum, COL.ANALISIS_AREA_IA2).setValue(analisis2.area || "");
        sheet.getRange(rowNum, COL.ANALISIS_SEVERIDAD_IA2).setValue(analisis2.severidad || "");
        sheet.getRange(rowNum, COL.ANALISIS_RECOMENDACION_IA2).setValue(analisis2.recomendacion || "");
        sheet.getRange(rowNum, COL.ANALISIS_ANGULO_IA2).setValue(analisis2.angulo || "");
      }
      sheet.getRange(rowNum, COL.STATUS).setValue("enviado");
      Logger.log("actualizarSesionAnalisis OK fila " + rowNum);
      return;
    }
    Logger.log("actualizarSesionAnalisis: no se encontro fila para " + from);
  } catch(err) { Logger.log("Error actualizarSesionAnalisis: " + err); }
}

function obtenerCodigoPlanPendiente(from) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SESIONES_SHEET);
  if (!sheet) return "";
  const data = sheet.getDataRange().getValues();
  for (let i = data.length-1; i >= 1; i--) { if (safeString(data[i][COL.WHATSAPP-1])===from && safeString(data[i][COL.EJVSPLAN-1])==="3" && !safeString(data[i][COL.PAGO_STATUS-1]).includes("pagado")) return safeString(data[i][COL.CODIGO_PLAN-1]); }
  return "";
}
function contarPlanesAlumno(from, sheet) {
  const data = sheet.getDataRange().getValues(); let count = 0;
  for (let i = 1; i < data.length; i++) { if (safeString(data[i][COL.WHATSAPP-1])===from && safeString(data[i][COL.EJVSPLAN-1])==="3") count++; }
  return count;
}

function obtenerConversacion(from) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONVERSATIONS_SHEET);
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) { if (safeString(data[i][0])===from) { try { return JSON.parse(safeString(data[i][1]))||{}; } catch(e) { return {}; } } }
  return {};
}
function guardarConversacion(from, estado) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONVERSATIONS_SHEET);
  if (!sheet) { sheet = ss.insertSheet(CONVERSATIONS_SHEET); sheet.appendRow(["whatsapp","estado","ultima_actualizacion"]); }
  const data = sheet.getDataRange().getValues(); const timestamp = new Date();
  for (let i = 1; i < data.length; i++) { if (safeString(data[i][0])===from) { sheet.getRange(i+1,2).setValue(JSON.stringify(estado)); sheet.getRange(i+1,3).setValue(timestamp); return; } }
  sheet.appendRow([from, JSON.stringify(estado), timestamp]);
}

function enviarMensajeWhatsApp(telefono, mensaje) {
  UrlFetchApp.fetch("https://graph.facebook.com/v19.0/" + PHONE_NUMBER_ID + "/messages", {
    method: "POST", headers: { "Authorization": "Bearer " + META_TOKEN, "Content-Type": "application/json" },
    payload: JSON.stringify({ messaging_product: "whatsapp", to: telefono, type: "text", text: { body: mensaje } })
  });
  logMensaje(telefono, "saliente", "texto", mensaje);
}
function enviarDocumentoWhatsApp(telefono, fileId, fileName) {
  const file = DriveApp.getFileById(fileId); const blob = file.getBlob();
  const uploadRes = UrlFetchApp.fetch("https://graph.facebook.com/v19.0/" + PHONE_NUMBER_ID + "/media", { method: "POST", headers: { "Authorization": "Bearer " + META_TOKEN }, payload: { messaging_product: "whatsapp", type: "application/pdf", file: blob } });
  const uploadData = JSON.parse(uploadRes.getContentText());
  if (!uploadData.id) throw new Error("Error subiendo media: " + uploadRes.getContentText());
  UrlFetchApp.fetch("https://graph.facebook.com/v19.0/" + PHONE_NUMBER_ID + "/messages", {
    method: "POST", headers: { "Authorization": "Bearer " + META_TOKEN, "Content-Type": "application/json" },
    payload: JSON.stringify({ messaging_product: "whatsapp", to: telefono, type: "document", document: { id: uploadData.id, filename: fileName, caption: "\u26f3 Ac\u00e1 est\u00e1 tu plan de entrenamiento personalizado. \u00a1Cualquier consulta escribinos \u26f3!" } })
  });
}

// ============================================
// FIX BUG 1: procesarSesionesPendientes
// Saltear planes (ejvsplan=3) con pago_status=pagado — ya manejados por el flujo
// ============================================
function procesarSesionesPendientes() {
  const lock = LockService.getScriptLock(); if (!lock.tryLock(3000)) return;
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sesionesSheet = ss.getSheetByName(SESIONES_SHEET); const leadsSheet = ss.getSheetByName(LEADS_SHEET);
    if (!sesionesSheet || !leadsSheet) return;
    const sesiones = sesionesSheet.getDataRange().getValues(); const leads = leadsSheet.getDataRange().getValues();
    if (sesiones.length < 2) return;
    for (let i = 1; i < sesiones.length; i++) {
      const rowNumber = i+1; const row = sesiones[i];
      const whatsapp = safeString(row[COL.WHATSAPP-1]); const aspectoRaw = safeString(row[COL.ASPECTO-1]);
      const nivelRaw = safeString(row[COL.NIVEL-1]); const ejvsplanRaw = safeString(row[COL.EJVSPLAN-1]);
      const status = safeString(row[COL.STATUS-1]).toLowerCase(); const contenidoEnviado = safeString(row[COL.CONTENIDO_ENVIADO-1]);
      const pagoStatusRow = safeString(row[COL.PAGO_STATUS-1]).toLowerCase();
      if (!whatsapp || status==="enviado" || status==="procesando" || status==="pendiente_manual" || contenidoEnviado) continue;
      if (pagoStatusRow === "pendiente") continue;
      // FIX BUG 1: planes pagados ya fueron manejados por el flujo conversacional
      if (ejvsplanRaw === "3" && pagoStatusRow === "pagado") {
        sesionesSheet.getRange(rowNumber, COL.STATUS).setValue("pendiente_manual");
        continue;
      }
      let nombre = ""; for (let j=1;j<leads.length;j++) { if (safeString(leads[j][0])===whatsapp) { nombre=safeString(leads[j][1]); break; } }
      try {
        sesionesSheet.getRange(rowNumber,COL.STATUS).setValue("procesando");
        const tipoSolicitud = mapTipoSolicitudPorEjvsplan(ejvsplanRaw);
        const perfil = { nivel: nivelRaw, aspecto: aspectoRaw, tiempo: "45 minutos" };
        if (tipoSolicitud === "analisis_video") { sesionesSheet.getRange(rowNumber,COL.STATUS).setValue("enviado"); continue; }
        if (tipoSolicitud === "plan_personalizado") {
          const mensajePlan = construirMensajePendienteManual(nombre, perfil);
          enviarMensajeWhatsApp(whatsapp, mensajePlan);
          notificarNuevoPlan(nombre, whatsapp, perfil, safeString(row[COL.VIDEO_URL1-1]), safeString(row[COL.COMENTARIOS_ALUMNO-1]));
          sesionesSheet.getRange(rowNumber,COL.STATUS).setValue("pendiente_manual");
          sesionesSheet.getRange(rowNumber,COL.CONTENIDO_ENVIADO).setValue(mensajePlan); continue;
        }
        const ejerciciosYaEnviados = obtenerEjerciciosYaEnviados(sesiones, whatsapp);
        if (ejerciciosYaEnviados.length >= FREE_EXERCISE_LIMIT) {
          const mensajeUpsell = construirMensajeUpsell(nombre);
          enviarMensajeWhatsApp(whatsapp, mensajeUpsell);
          sesionesSheet.getRange(rowNumber,COL.STATUS).setValue("enviado");
          sesionesSheet.getRange(rowNumber,COL.CONTENIDO_ENVIADO).setValue(mensajeUpsell);
          sesionesSheet.getRange(rowNumber,COL.EJERCICIO_GRATIS_ID).setValue("UPSELL"); continue;
        }
        const ejercicios = seleccionarEjercicios(perfil, 1, ejerciciosYaEnviados);
        if (ejercicios.length === 0) {
          const mensajeError = "Hola " + nombre + " \u26f3\n\nHubo un inconveniente al preparar tu ejercicio.\nUno de nuestros expertos se va a contactar con vos a la brevedad.";
          enviarMensajeWhatsApp(whatsapp, mensajeError);
          sesionesSheet.getRange(rowNumber,COL.STATUS).setValue("pendiente_manual");
          sesionesSheet.getRange(rowNumber,COL.CONTENIDO_ENVIADO).setValue(mensajeError);
          sesionesSheet.getRange(rowNumber,COL.EJERCICIO_GRATIS_ID).setValue("SIN_EJERCICIO"); continue;
        }
        const prompt = buildPromptConEjercicios({ nombre, perfil, ejercicios });
        const plan = llamarOpenAI(prompt, OPENAI_API_KEY);
        const mensajeFinal = construirMensajeFinal(plan, ejercicios);
        enviarMensajeWhatsApp(whatsapp, truncateForWhatsApp(mensajeFinal, 1400));
        sesionesSheet.getRange(rowNumber,COL.STATUS).setValue("enviado");
        sesionesSheet.getRange(rowNumber,COL.CONTENIDO_ENVIADO).setValue(mensajeFinal);
        sesionesSheet.getRange(rowNumber,COL.EJERCICIO_GRATIS_ID).setValue(ejercicios[0].id);
      } catch (error) { sesionesSheet.getRange(rowNumber,COL.STATUS).setValue("error"); sesionesSheet.getRange(rowNumber,COL.CONTENIDO_ENVIADO).setValue(String(error)); Logger.log("Error sesion fila " + rowNumber + ": " + error); }
    }
  } finally { lock.releaseLock(); }
}

function notificarNuevoPlan(nombre, whatsapp, perfil, videoUrl, comentariosAlumno) {
  try {
    const email = Session.getActiveUser().getEmail();
    GmailApp.sendEmail(email, "\u26f3 Nuevo plan solicitado \u2014 " + nombre,
      "\ud83d\udc64 Nombre: " + nombre + "\n\ud83d\udcf1 WhatsApp: " + whatsapp + "\n\ud83c\udfcc\ufe0f Nivel: " + perfil.nivel + "\n\ud83c\udfaf Aspecto: " + perfil.aspecto + "\n\ud83d\udcac Comentarios: " + (comentariosAlumno||"No indic\u00f3") + "\n\ud83c\udfa5 Video: " + (videoUrl||"No envi\u00f3 video") + "\n\n\ud83d\udcac WhatsApp: https://wa.me/" + whatsapp + "\n\nAbr\u00ed el panel de Golfito para generar y enviar el plan.");
  } catch(err) { Logger.log("Error email: " + err); }
}

// ============================================
// WEB APP FUNCTIONS
// ============================================
function obtenerLeads() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sesionesSheet = ss.getSheetByName(SESIONES_SHEET); const leadsSheet = ss.getSheetByName(LEADS_SHEET);
  if (!sesionesSheet) return { leads: [] };
  const sesiones = sesionesSheet.getDataRange().getValues(); const leadsData = leadsSheet ? leadsSheet.getDataRange().getValues() : [];
  const leadsMap = {}; for (let i=1;i<leadsData.length;i++) { const wa=safeString(leadsData[i][0]); if (wa) leadsMap[wa]={nombre:safeString(leadsData[i][1]),handicap:safeString(leadsData[i][3])}; }
  const convsSheet = ss.getSheetByName(CONVERSATIONS_SHEET); const convsMap = {};
  if (convsSheet) { const cd=convsSheet.getDataRange().getValues(); for (let i=1;i<cd.length;i++) { const wa=safeString(cd[i][0]); if (!wa) continue; try { const e=JSON.parse(safeString(cd[i][1]))||{}; if (e.comentarios_alumno) convsMap[wa]=e.comentarios_alumno; } catch(e){} } }
  const leads = [];
  for (let i=1;i<sesiones.length;i++) {
    const row=sesiones[i]; const whatsapp=safeString(row[COL.WHATSAPP-1]); const ejvsplan=safeString(row[COL.EJVSPLAN-1]); const status=safeString(row[COL.STATUS-1]).toLowerCase();
    if (!whatsapp || (ejvsplan!=="3" && ejvsplan!=="2")) continue;
    const li = leadsMap[whatsapp]||{nombre:"",handicap:""};
    leads.push({ rowIndex:i+1, whatsapp, nombre:li.nombre, handicap:li.handicap, nivel:safeString(row[COL.NIVEL-1]), aspecto:safeString(row[COL.ASPECTO-1]), ejvsplan, status,
      videoUrl1:safeString(row[COL.VIDEO_URL1-1]), videoUrl2:safeString(row[COL.VIDEO_URL2-1]),
      analisisError1:safeString(row[COL.ANALISIS_ERROR_IA1-1]), analisisArea1:safeString(row[COL.ANALISIS_AREA_IA1-1]),
      analisisSeveridad1:safeString(row[COL.ANALISIS_SEVERIDAD_IA1-1]), analisisRecomendacion1:safeString(row[COL.ANALISIS_RECOMENDACION_IA1-1]),
      analisisAngulo1:safeString(row[COL.ANALISIS_ANGULO_IA1-1]), analisisError2:safeString(row[COL.ANALISIS_ERROR_IA2-1]),
      analisisArea2:safeString(row[COL.ANALISIS_AREA_IA2-1]), analisisSeveridad2:safeString(row[COL.ANALISIS_SEVERIDAD_IA2-1]),
      analisisRecomendacion2:safeString(row[COL.ANALISIS_RECOMENDACION_IA2-1]), analisisAngulo2:safeString(row[COL.ANALISIS_ANGULO_IA2-1]),
      analisisManual1:safeString(row[COL.ANALISIS_MANUAL_1-1]), analisisManual2:safeString(row[COL.ANALISIS_MANUAL_2-1]),
      pagoStatus:safeString(row[COL.PAGO_STATUS-1]), pagoMonto:safeString(row[COL.PAGO_MONTO-1]),
      codigoPlan:safeString(row[COL.CODIGO_PLAN-1]), numPlanAlumno:safeString(row[COL.NUM_PLAN_ALUMNO-1]),
      notaCoach:safeString(row[COL.NOTA_COACH-1]), diagnostico:safeString(row[COL.DIAGNOSTICO-1]),
      entradaEnCalor:safeString(row[COL.ENTRADA_EN_CALOR-1]), consideraciones:safeString(row[COL.CONSIDERACIONES-1]),
      comentariosAlumno:safeString(row[COL.COMENTARIOS_ALUMNO-1])||convsMap[whatsapp]||"",
      ejercicioApproach:safeString(row[COL.EJERCICIO_APPROACH-1]), ejercicioFullswingHierros:safeString(row[COL.EJERCICIO_FULLSWING_HIERROS-1]),
      ejercicioFullswingMaderas:safeString(row[COL.EJERCICIO_FULLSWING_MADERAS-1]), ejercicioPutter:safeString(row[COL.EJERCICIO_PUTTER-1]),
      feedbackAlumno:safeString(row[COL.FEEDBACK_ALUMNO-1]), feedbackScore:safeString(row[COL.FEEDBACK_SCORE-1]),
      fecha:row[0] ? Utilities.formatDate(new Date(row[0]),Session.getScriptTimeZone(),"dd/MM/yyyy HH:mm") : "—" });
  }
  leads.sort((a,b) => { if (a.status==='pendiente_manual' && b.status!=='pendiente_manual') return -1; if (a.status!=='pendiente_manual' && b.status==='pendiente_manual') return 1; return 0; });
  return { leads };
}

function obtenerConversacionesActivas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const convsSheet = ss.getSheetByName(CONVERSATIONS_SHEET); const leadsSheet = ss.getSheetByName(LEADS_SHEET);
  if (!convsSheet) return { conversaciones: [] };
  const convsData = convsSheet.getDataRange().getValues(); const leadsData = leadsSheet ? leadsSheet.getDataRange().getValues() : [];
  const leadsMap = {}; for (let i=1;i<leadsData.length;i++) { const wa=safeString(leadsData[i][0]); if (wa) leadsMap[wa]={nombre:safeString(leadsData[i][1])}; }
  const conversaciones = [];
  for (let i=1;i<convsData.length;i++) {
    const wa=safeString(convsData[i][0]); if (!wa) continue;
    let estado={}; try { estado=JSON.parse(safeString(convsData[i][1]))||{}; } catch(e){}
    const ua = convsData[i][2] ? Utilities.formatDate(new Date(convsData[i][2]),Session.getScriptTimeZone(),"dd/MM/yyyy HH:mm") : "—";
    const li = leadsMap[wa]||{nombre:""};
    conversaciones.push({ whatsapp:wa, nombre:li.nombre||estado.nombre||wa, paso:estado.paso||"—", ultimaActualizacion:ua });
  }
  conversaciones.sort((a,b) => b.ultimaActualizacion.localeCompare(a.ultimaActualizacion));
  return { conversaciones };
}

function obtenerChatLogPorWhatsapp(whatsapp) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ChatLog");
    if (!sheet) return { mensajes: [] };
    const data = sheet.getDataRange().getValues();
    const mensajes = [];
    for (let i = 1; i < data.length; i++) {
      const wa = safeString(data[i][1]);
      if (wa !== whatsapp) continue;
      const ts = data[i][0];
      const direccion = safeString(data[i][2]);
      const tipo = safeString(data[i][3]);
      const contenido = safeString(data[i][4]);
      if (tipo !== "texto" && tipo !== "video" && tipo !== "video_drive") continue;
      mensajes.push({
        timestamp: ts ? Utilities.formatDate(new Date(ts), Session.getScriptTimeZone(), "dd/MM HH:mm") : "—",
        direccion, tipo,
        contenido: contenido.length > 300 ? contenido.slice(0, 300) + "…" : contenido
      });
    }
    return { mensajes: mensajes.slice(-60) };
  } catch(err) { Logger.log("Error obtenerChatLogPorWhatsapp: " + err); return { mensajes: [] }; }
}

function enviarMensajeDesdePanel(whatsapp, mensaje) {
  try { if (!whatsapp||!mensaje) return {ok:false,error:"Numero y mensaje requeridos"}; enviarMensajeWhatsApp(whatsapp,mensaje); return {ok:true}; }
  catch(err) { Logger.log("Error enviarMensajeDesdePanel: "+err); return {ok:false,error:err.toString()}; }
}

// ============================================
// FIX BUG 3: obtenerConsultas y marcarConsultaRespondida
// ============================================
function obtenerConsultas() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONSULTAS_SHEET);
    const leadsSheet = ss.getSheetByName(LEADS_SHEET);
    if (!sheet) return { consultas: [] };
    const data = sheet.getDataRange().getValues();
    const leadsData = leadsSheet ? leadsSheet.getDataRange().getValues() : [];
    const leadsMap = {};
    for (let i=1;i<leadsData.length;i++) { const wa=safeString(leadsData[i][0]); if (wa) leadsMap[wa]={nombre:safeString(leadsData[i][1])}; }
    const consultas = [];
    for (let i=1;i<data.length;i++) {
      const whatsapp = safeString(data[i][1]);
      const li = leadsMap[whatsapp]||{nombre:""};
      consultas.push({
        rowIndex: i+1,
        timestamp: data[i][0] ? Utilities.formatDate(new Date(data[i][0]),Session.getScriptTimeZone(),"dd/MM/yyyy HH:mm") : "—",
        whatsapp,
        nombre: li.nombre || safeString(data[i][2]),
        consulta: safeString(data[i][3]),
        respondida: safeString(data[i][4])
      });
    }
    consultas.sort((a,b) => a.respondida==="no" && b.respondida!=="no" ? -1 : 1);
    return { consultas };
  } catch(err) { Logger.log("Error obtenerConsultas: "+err); return { consultas: [] }; }
}

function marcarConsultaRespondida(rowIndex) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONSULTAS_SHEET);
    if (!sheet) return {ok:false,error:"No encontre Consultas"};
    sheet.getRange(rowIndex, 5).setValue("si");
    return {ok:true};
  } catch(err) { return {ok:false,error:err.toString()}; }
}

function obtenerEjerciciosParaSelector() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EXERCISES_SHEET); if (!sheet) return [];
  const data = sheet.getDataRange().getValues(); const ejercicios = [];
  for (let i=1;i<data.length;i++) { const e=rowToEjercicio(data[i]); if (!ejercicioValido(e)) continue; ejercicios.push({id:e.id,nombre:e.nombre,aspecto:e.aspecto,nivel:e.nivel,descripcion:e.instruccion,instruccion:e.instruccion,videoUrl:e.video_url,tipoEjercicio:e.tipo_ejercicio,profesor:e.profesor}); }
  return ejercicios;
}
function obtenerEjerciciosCompletos() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EXERCISES_SHEET); if (!sheet) return {ejercicios:[]};
  const data = sheet.getDataRange().getValues(); const ejercicios = [];
  for (let i=1;i<data.length;i++) { const e=rowToEjercicio(data[i]); if (!e.id) continue; ejercicios.push({rowIndex:i+1,id:e.id,nombre:e.nombre,aspecto:e.aspecto,nivel:e.nivel,videoUrl:e.video_url,instruccion:e.instruccion,focoTecnico:e.foco_tecnico,tipoEjercicio:e.tipo_ejercicio,etiqueta1:e.etiqueta_1,etiqueta2:e.etiqueta_2,etiqueta3:e.etiqueta_3,profesor:e.profesor}); }
  return {ejercicios};
}
function guardarEjercicio(datos) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EXERCISES_SHEET); if (!sheet) throw new Error("No encontre la hoja Excercises");
    const row=[datos.id||"",datos.nombre||"",datos.instruccion||"",(datos.aspecto||"").toLowerCase(),(datos.nivel||"").toLowerCase(),datos.videoUrl||"",datos.tipoEjercicio||"",datos.focoTecnico||"",datos.etiqueta1||"",datos.etiqueta2||"",datos.etiqueta3||"",datos.profesor||""];
    if (datos.rowIndex) { sheet.getRange(datos.rowIndex,1,1,row.length).setValues([row]); }
    else { if (!datos.id) row[0]="EJ-"+String(Date.now()).slice(-6); sheet.appendRow(row); }
    return {ok:true};
  } catch(err) { Logger.log("Error guardarEjercicio: "+err); return {ok:false,error:err.toString()}; }
}
function eliminarEjercicio(rowIndex) {
  try { const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EXERCISES_SHEET); if (!sheet) throw new Error("No encontre la hoja Excercises"); sheet.deleteRow(rowIndex); return {ok:true}; }
  catch(err) { return {ok:false,error:err.toString()}; }
}
function guardarPlanProfe(rowIndex, datos) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SESIONES_SHEET); if (!sheet) throw new Error("No encontre Sesiones");
    sheet.getRange(rowIndex,COL.NOTA_COACH).setValue(datos.notaCoach||""); sheet.getRange(rowIndex,COL.EJERCICIO_APPROACH).setValue(datos.ejercicioApproach||"");
    sheet.getRange(rowIndex,COL.EJERCICIO_FULLSWING_HIERROS).setValue(datos.ejercicioFullswingHierros||""); sheet.getRange(rowIndex,COL.EJERCICIO_FULLSWING_MADERAS).setValue(datos.ejercicioFullswingMaderas||""); sheet.getRange(rowIndex,COL.EJERCICIO_PUTTER).setValue(datos.ejercicioPutter||"");
    return {ok:true};
  } catch(err) { return {ok:false,error:err.toString()}; }
}
function guardarAnalisisEnSheets(rowIndex, a1, a2) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SESIONES_SHEET); if (!sheet) return {ok:false,error:"No encontre Sesiones"};
    if (a1) { sheet.getRange(rowIndex,COL.ANALISIS_ERROR_IA1).setValue(a1.error_principal||""); sheet.getRange(rowIndex,COL.ANALISIS_AREA_IA1).setValue(a1.area||""); sheet.getRange(rowIndex,COL.ANALISIS_SEVERIDAD_IA1).setValue(a1.severidad||""); sheet.getRange(rowIndex,COL.ANALISIS_RECOMENDACION_IA1).setValue(a1.recomendacion||""); sheet.getRange(rowIndex,COL.ANALISIS_ANGULO_IA1).setValue(a1.angulo||""); }
    if (a2) { sheet.getRange(rowIndex,COL.ANALISIS_ERROR_IA2).setValue(a2.error_principal||""); sheet.getRange(rowIndex,COL.ANALISIS_AREA_IA2).setValue(a2.area||""); sheet.getRange(rowIndex,COL.ANALISIS_SEVERIDAD_IA2).setValue(a2.severidad||""); sheet.getRange(rowIndex,COL.ANALISIS_RECOMENDACION_IA2).setValue(a2.recomendacion||""); sheet.getRange(rowIndex,COL.ANALISIS_ANGULO_IA2).setValue(a2.angulo||""); }
    return {ok:true};
  } catch(err) { return {ok:false,error:err.toString()}; }
}
function guardarAnalisisManual(rowIndex, manual1, manual2) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SESIONES_SHEET); if (!sheet) return {ok:false,error:"No encontre Sesiones"};
    if (manual1!==undefined) sheet.getRange(rowIndex,COL.ANALISIS_MANUAL_1).setValue(manual1); if (manual2!==undefined) sheet.getRange(rowIndex,COL.ANALISIS_MANUAL_2).setValue(manual2);
    return {ok:true};
  } catch(err) { return {ok:false,error:err.toString()}; }
}
function generarPlanConIA(lead) {
  try {
    const ctx = lead.analisisError1 ? "\n\nAnalisis swing (" + (lead.analisisAngulo1||"video 1") + "):\n- Error: " + lead.analisisError1 + "\n- Recomendacion: " + lead.analisisRecomendacion1 + (lead.analisisError2 ? "\n\nAnalisis 2 (" + (lead.analisisAngulo2||"video 2") + "):\n- Error: " + lead.analisisError2 + "\n- Recomendacion: " + lead.analisisRecomendacion2 : "") : "";
    const prompt = "Sos un coach profesional de golf. Tu tono es amigable y pedagógico para jugadores amateurs. Analizá los datos del alumno y generá un diagnóstico y recomendaciones.\n\nDatos:\n- Nombre: " + lead.nombre + "\n- Handicap: " + lead.handicap + "\n- Nivel: " + lead.nivel + "\n- Comentarios: " + (lead.comentariosAlumno||"no especificado") + ctx + "\n\nGUÍA DE LENGUAJE:\n✔ SÍ PODÉS usar términos naturales del golf: grip, stance, putt, drive, approach, backswing, downswing, follow through, finish.\n❌ EVITÁ jerga compleja. Usá alternativas en español para: casting (\"soltar las manos antes de tiempo\"), over the top (\"tirar el cuerpo encima\"), chicken wing (\"doblar el codo hacia afuera\"), early extension (\"perder la postura\"), sway (\"balanceo lateral\"), scooping (\"cucharear\"), topping (\"pegarle arriba\"), chunk (\"pegarle pesado a la tierra\").\n\nResponde SOLO en JSON sin texto adicional ni markdown:\n{\"diagnostico\":\"...\",\"entrada_en_calor\":{\"descripcion\":\"...\",\"video_url\":\"...\"},\"consideraciones\":\"...\"}";
    const planJson = llamarOpenAI(prompt, OPENAI_API_KEY);
    const plan = JSON.parse(planJson.replace(/```json|```/g,'').trim());
    return {ok:true,plan};
  } catch(err) { Logger.log("Error generarPlanConIA: "+err); return {ok:false,error:err.toString()}; }
}

function generarYEnviarPlanDesdeWeb(rowIndex, planData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sesionesSheet = ss.getSheetByName(SESIONES_SHEET); const leadsSheet = ss.getSheetByName(LEADS_SHEET);
    const sesionRow = sesionesSheet.getRange(rowIndex,1,1,SESIONES_HEADERS.length).getValues()[0];
    const whatsapp = safeString(sesionRow[COL.WHATSAPP-1]);
    let nombre="", handicap=""; const leadsData = leadsSheet.getDataRange().getValues();
    for (let i=1;i<leadsData.length;i++) { if (safeString(leadsData[i][0])===whatsapp) { nombre=safeString(leadsData[i][1]); handicap=safeString(leadsData[i][3]); break; } }
    const htmlPlan = generarHTMLPlan(planData, nombre, handicap, sesionRow);
    const blob = Utilities.newBlob(htmlPlan,'text/html','plan.html'); const driveFile = DriveApp.createFile(blob);
    const pdfBlob = driveFile.getAs('application/pdf').setName("Plan_Golfito_"+nombre+"_"+Date.now()+".pdf"); driveFile.setTrashed(true);
    let folder; const folders = DriveApp.getFoldersByName("Golfito_Planes"); folder = folders.hasNext() ? folders.next() : DriveApp.createFolder("Golfito_Planes");
    const pdfFile = folder.createFile(pdfBlob); pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    enviarDocumentoWhatsApp(whatsapp, pdfFile.getId(), "Plan_Golfito_"+nombre+".pdf");
    sesionesSheet.getRange(rowIndex,COL.STATUS).setValue("enviado");
    sesionesSheet.getRange(rowIndex,COL.NOTA_COACH).setValue(planData.nota_coach||"");
    sesionesSheet.getRange(rowIndex,COL.CONTENIDO_ENVIADO).setValue(pdfFile.getUrl());
    sesionesSheet.getRange(rowIndex,COL.DIAGNOSTICO).setValue(planData.diagnostico||"");
    sesionesSheet.getRange(rowIndex,COL.ENTRADA_EN_CALOR).setValue(planData.entrada_en_calor?.descripcion||"");
    sesionesSheet.getRange(rowIndex,COL.CONSIDERACIONES).setValue(planData.consideraciones||"");
    sesionesSheet.getRange(rowIndex,COL.EJERCICIO_APPROACH).setValue(planData.approach?.ejercicio||planData.ejercicioApproach||"");
    sesionesSheet.getRange(rowIndex,COL.EJERCICIO_FULLSWING_HIERROS).setValue(planData.fullswing_hierros?.ejercicio||planData.ejercicioFullswingHierros||"");
    sesionesSheet.getRange(rowIndex,COL.EJERCICIO_FULLSWING_MADERAS).setValue(planData.fullswing_maderas?.ejercicio||planData.ejercicioFullswingMaderas||"");
    sesionesSheet.getRange(rowIndex,COL.EJERCICIO_PUTTER).setValue(planData.putter?.ejercicio||planData.ejercicioPutter||"");
    return {ok:true};
  } catch(err) { Logger.log("Error generarYEnviarPlanDesdeWeb: "+err); return {ok:false,error:err.toString()}; }
}

// ============================================
// GENERAR HTML DEL PLAN
// ============================================
function generarHTMLPlan(plan, nombre, handicap, sesionRow) {
  const nivel=safeString(sesionRow[COL.NIVEL-1]); const codigoPlan=safeString(sesionRow[COL.CODIGO_PLAN-1]); const numPlan=safeString(sesionRow[COL.NUM_PLAN_ALUMNO-1]);
  const modoAnalisis=plan.modoAnalisis||'ia';
  const manual1=plan.analisisManual1||""; const manual2=plan.analisisManual2||"";
  const err1=plan.analisisError1||safeString(sesionRow[COL.ANALISIS_ERROR_IA1-1]); const rec1=plan.analisisRecomendacion1||safeString(sesionRow[COL.ANALISIS_RECOMENDACION_IA1-1]); const ang1=plan.analisisAngulo1||safeString(sesionRow[COL.ANALISIS_ANGULO_IA1-1]);
  const err2=plan.analisisError2||safeString(sesionRow[COL.ANALISIS_ERROR_IA2-1]); const rec2=plan.analisisRecomendacion2||safeString(sesionRow[COL.ANALISIS_RECOMENDACION_IA2-1]); const ang2=plan.analisisAngulo2||safeString(sesionRow[COL.ANALISIS_ANGULO_IA2-1]);
  const foco=plan.foco||""; const notaCoach=plan.nota_coach||""; const comentariosAlumno=plan.comentariosAlumno||safeString(sesionRow[COL.COMENTARIOS_ALUMNO-1]);
  const entradaCalorTexto=plan.entrada_en_calor?.descripcion||ENTRADA_CALOR_STD; const consideracionesTexto=plan.consideraciones||CONSIDERACIONES_STD;
  const logoSVG='<svg width="48" height="48" viewBox="0 0 48 48" xmlns="http://www.w3.org/2000/svg"><circle cx="24" cy="24" r="24" fill="#c9a84c"/><text x="24" y="32" text-anchor="middle" font-size="26" fill="white">⛳</text></svg>';
  function bloqueIA(error,rec,angulo,etiqueta) {
    if (!error) return "";
    const al=angulo ? " — "+angulo.charAt(0).toUpperCase()+angulo.slice(1) : "";
    return '<div class="analisis-bloque"><div class="analisis-header">'+etiqueta+al+'</div><div class="analisis-body"><div class="analisis-item"><span class="analisis-label">Error principal</span><span class="analisis-valor">'+error+'</span></div><div class="analisis-item"><span class="analisis-label">Recomendacion</span><span class="analisis-valor">'+rec+'</span></div></div></div>';
  }
  function bloqueManual(texto,etiqueta) {
    if (!texto) return "";
    return '<div class="analisis-bloque"><div class="analisis-header">'+etiqueta+'</div><div class="analisis-body"><div class="analisis-valor" style="line-height:1.7">'+texto+'</div></div></div>';
  }
  let bloquesAnalisis="";
  if (modoAnalisis==='manual') { bloquesAnalisis+=bloqueManual(manual1,"📝 Analisis del profe — Video 1"); bloquesAnalisis+=bloqueManual(manual2,"📝 Analisis del profe — Video 2"); }
  else { bloquesAnalisis+=bloqueIA(err1,rec1,ang1,"🤖 Analisis IA — Video 1"); bloquesAnalisis+=bloqueIA(err2,rec2,ang2,"🤖 Analisis IA — Video 2"); }
  function seccion(header, body) { return '<div class="section"><div class="section-header">'+header+'</div><div class="section-body">'+body+'</div></div>'; }
  function seccionEj(header, ej) { if (!ej) return seccion(header,""); return seccion(header,'<strong>'+(ej.ejercicio||"")+'</strong><br>'+(ej.descripcion||"")+(ej.video_url?'<br><a href="'+ej.video_url+'">'+ej.video_url+'</a>':"")); }
  return '<!DOCTYPE html><html><head><meta charset="UTF-8"><style>body{font-family:Arial,sans-serif;margin:0;padding:20px;color:#1a1a1a;font-size:13px;}.header{background:#1a472a;color:white;padding:18px 24px;border-radius:8px;margin-bottom:18px;display:flex;align-items:center;gap:18px;}.header h1{margin:0 0 4px 0;font-size:24px;letter-spacing:2px;}.header-sub{font-size:12px;color:rgba(255,255,255,0.7);}.data-box{background:#f0ebe0;border-radius:8px;padding:12px 16px;margin-bottom:14px;}.data-box h3{margin:0 0 8px 0;font-size:11px;text-transform:uppercase;color:#1a472a;border-bottom:2px solid #1a472a;padding-bottom:4px;}.data-grid{display:flex;gap:24px;flex-wrap:wrap;}.data-item{font-size:13px;}.data-item strong,.data-item-full strong{color:#1a472a;}.data-item-full{font-size:13px;width:100%;margin-top:6px;}.analisis-bloque{background:#e8f5e9;border-left:4px solid #40916c;border-radius:0 8px 8px 0;margin-bottom:10px;overflow:hidden;}.analisis-header{background:#2d6a4f;color:white;padding:7px 14px;font-size:11px;font-weight:bold;text-transform:uppercase;}.analisis-body{padding:10px 14px;display:grid;gap:6px;}.analisis-item{display:flex;gap:8px;}.analisis-label{font-size:10px;text-transform:uppercase;color:#2d6a4f;font-weight:bold;min-width:110px;}.analisis-valor{font-size:13px;line-height:1.5;}.foco-box{background:#fff8e6;border-left:4px solid #c9a84c;padding:10px 14px;border-radius:0 8px 8px 0;margin-bottom:12px;font-weight:600;color:#1a472a;}.nota-coach{background:#fff8e6;border-left:4px solid #c9a84c;padding:10px 14px;border-radius:0 8px 8px 0;margin-bottom:14px;font-style:italic;line-height:1.6;}.diagnostico{background:#e8f5e9;border-left:4px solid #1a472a;padding:10px 14px;border-radius:0 8px 8px 0;margin-bottom:12px;line-height:1.6;}.section{margin-bottom:12px;border-radius:8px;overflow:hidden;}.section-header{background:#1a472a;color:white;padding:9px 14px;font-weight:bold;font-size:12px;text-transform:uppercase;}.section-body{background:#faf7f0;padding:12px 14px;line-height:1.6;}.section-body a{color:#1565c0;font-size:12px;}.section-body strong{color:#1a472a;}</style></head><body>'
    +'<div class="header"><div>'+logoSVG+'</div><div><h1>⛳ Golfito Plan</h1><div class="header-sub">'+(numPlan?"Plan #"+numPlan:"")+(codigoPlan?" · "+codigoPlan:"")+'</div></div></div>'
    +'<div class="data-box"><h3>Datos del jugador</h3><div class="data-grid"><div class="data-item"><strong>Nombre:</strong> '+nombre+'</div><div class="data-item"><strong>Handicap:</strong> '+handicap+'</div><div class="data-item"><strong>Nivel:</strong> '+nivel+'</div></div>'+(comentariosAlumno?'<div class="data-item-full"><strong>💬 Comentarios:</strong> '+comentariosAlumno+'</div>':'')+'</div>'
    +(foco?'<div class="foco-box">🎯 Foco: '+foco+'</div>':"")
    +(notaCoach?'<div class="nota-coach">"'+notaCoach+'"</div>':"")
    +bloquesAnalisis
    +(plan.diagnostico?'<div class="diagnostico"><strong>Diagnostico:</strong> '+plan.diagnostico+'</div>':"")
    +seccion("🔥 Entrada en calor", entradaCalorTexto)
    +seccionEj("🎯 Approach", plan.approach)
    +seccionEj("🏌️ Full Swing — Hierros", plan.fullswing_hierros)
    +seccionEj("🏌️ Full Swing — Maderas", plan.fullswing_maderas)
    +seccionEj("⛳ Putter", plan.putter)
    +seccion("💡 Consideraciones", consideracionesTexto)
    +'</body></html>';
}

// ============================================
// GEMINI
// ============================================
function subirVideoAGemini(fileId) {
  const file = DriveApp.getFileById(fileId); const blob = file.getBlob(); const mimeType = "video/mp4";
  const initRes = UrlFetchApp.fetch("https://generativelanguage.googleapis.com/upload/v1beta/files?key="+GEMINI_API_KEY, { method:"POST", headers:{"X-Goog-Upload-Protocol":"resumable","X-Goog-Upload-Command":"start","X-Goog-Upload-Header-Content-Type":mimeType,"Content-Type":"application/json"}, payload:JSON.stringify({file:{display_name:file.getName()}}), muteHttpExceptions:true });
  const uploadUrl = initRes.getHeaders()["x-goog-upload-url"];
  if (!uploadUrl) throw new Error("No se obtuvo upload URL: " + initRes.getContentText());
  const uploadRes = UrlFetchApp.fetch(uploadUrl, { method:"POST", headers:{"X-Goog-Upload-Offset":"0","X-Goog-Upload-Command":"upload, finalize","Content-Type":mimeType}, payload:blob, muteHttpExceptions:true });
  const uploadData = JSON.parse(uploadRes.getContentText());
  const fileUri = uploadData.file?.uri; if (!fileUri) throw new Error("No se obtuvo file URI: " + uploadRes.getContentText());
  let estado = uploadData.file?.state; let intentos = 0; const fileName = uploadData.file?.name?.split('/').pop();
  while (estado === "PROCESSING" && intentos < 10) { Utilities.sleep(3000); const sr = UrlFetchApp.fetch("https://generativelanguage.googleapis.com/v1beta/files/"+fileName+"?key="+GEMINI_API_KEY, {muteHttpExceptions:true}); estado = JSON.parse(sr.getContentText()).state; intentos++; }
  if (estado !== "ACTIVE") throw new Error("Video no activo en Gemini. Estado: " + estado);
  return fileUri;
}

function analizarSwingConGemini(driveFileId) {
  try {
    const fileUri = subirVideoAGemini(driveFileId);
    Logger.log("fileUri: " + fileUri);
    const prompt = "Sos un coach profesional de golf con 20 años de experiencia analizando swings amateurs. Tu enfoque es amigable y pedagógico para principiantes: identificás el error más crítico y lo explicás de forma natural, usando el vocabulario habitual del club pero evitando tecnicismos complejos de física o biomecánica que abrumen al jugador.\n\nAnalizá el video del swing amateur siguiendo estrictamente estos pasos:\n\nPASO 1 — CONTEXTO DEL TIRO:\nIdentificá qué palo/área está usando el jugador: Drive, Hierros, Approach, Putt, Bunker o General.\n\nPASO 2 — DETECTAR ÁNGULO DE CÁMARA:\n- \"perfil\" (Cámara en la línea del objetivo).\n- \"frontal\" (Cámara de frente al pecho del jugador).\n- \"desconocido\" (Cámara en movimiento, oblicua o ángulo que impide ver el cuerpo/palo).\n*Nota: Si es \"desconocido\", omití el paso 3 y reportá el error de cámara en los detalles técnicos.*\n\nPASO 3 — EVALUAR ERROR CRÍTICO SEGÚN EL ÁNGULO:\nEvaluá las siguientes fases buscando únicamente el error que cause mayor pérdida de consistencia o dirección:\n- Si es PERFIL: Alineación en el stance, inicio del movimiento (takeaway), plano del backswing, downswing, impacto, follow through y finish.\n- Si es FRONTAL: Postura/grip, posición de la cabeza, transferencia de peso en el backswing, y liberación de manos en el impacto.\n\nGUÍA DE LENGUAJE (Equilibrio para Principiantes):\n✔ SÍ PODÉS usar los términos tradicionales y naturales del golf: grip, stance, lie, spin, smash factor, putt, drive, approach, backswing, downswing, follow through, finish.\n❌ EVITÁ jerga compleja de errores o datos de lanzamiento. Buscá alternativas descriptivas y sencillas en español para conceptos como:\n- Errores/Biomecánica: casting (ej: \"soltar las manos antes de tiempo\"), over the top (\"tirar el cuerpo encima\"), chicken wing (\"doblar el codo hacia afuera\"), early extension (\"perder la postura / enderezarse antes de tiempo\"), sway/slide (\"balanceo lateral\"), wrist hinge (\"quiebre de muñecas\"), scooping (\"cucharear\") o topping (\"pegarle arriba\").\n- Datos de radar/física: launch angle, club path, face angle, attack angle, club speed, apex, chunk/duff (\"pegarle pesado/a la tierra\"), o thin shot (\"pegarle finito\").\n\nRespondé EXCLUSIVAMENTE con un objeto JSON válido, sin texto adicional ni backticks. Respetá este esquema:\n{\"angulo\":\"perfil|frontal|desconocido\",\"area\":\"Drive|Hierros|Approach|Putt|Bunker|General\",\"severidad\":\"leve|moderado|importante\",\"error_principal\":\"...\",\"detalles\":\"...\",\"recomendacion\":\"...\"}";
    const MAX_INTENTOS=5; const ESPERA_MS=[5000,10000,20000,30000,40000]; let lastError="";
    for (let intento=0;intento<MAX_INTENTOS;intento++) {
      if (intento>0) Utilities.sleep(ESPERA_MS[intento-1]);
      const response = UrlFetchApp.fetch("https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key="+GEMINI_API_KEY, { method:"POST", headers:{"Content-Type":"application/json"}, muteHttpExceptions:true, deadline:120, payload:JSON.stringify({contents:[{parts:[{text:prompt},{file_data:{mime_type:"video/mp4",file_uri:fileUri}}]}],generationConfig:{temperature:0.2}}) });
      const statusCode = response.getResponseCode(); const body = response.getContentText();
      if (statusCode===503) { lastError="Error Gemini 503 intento "+(intento+1)+": "+body; Logger.log(lastError); continue; }
      if (statusCode<200||statusCode>=300) throw new Error("Error Gemini "+statusCode+": "+body);
      const data = JSON.parse(body); const text = data.candidates?.[0]?.content?.parts?.[0]?.text||"";
      const analisis = JSON.parse(text.replace(/```json|```/g,"").trim());
      return {ok:true,analisis};
    }
    throw new Error("Gemini no disponible tras "+MAX_INTENTOS+" intentos. "+lastError);
  } catch(err) { Logger.log("Error analizarSwingConGemini: "+err); return {ok:false,error:err.toString()}; }
}

// ============================================
// MERCADOPAGO
// ============================================
function crearPreferenciaPago(whatsapp, nombre, tipo, codigoPlan) {
  try {
    const esAnalisis = tipo==="analisis";
    const monto = esAnalisis ? 1000 : 3000;
    const externalRef = codigoPlan+"_"+whatsapp;
    const payload = { items:[{title:(esAnalisis?"Analisis de swing — ":"Plan personalizado — ")+nombre,quantity:1,unit_price:monto,currency_id:"CLP"}], payer:{name:nombre}, external_reference:externalRef, notification_url:MP_WEBHOOK_URL+"?source=mp", back_urls:{success:"https://wa.me/"+whatsapp,failure:"https://wa.me/"+whatsapp,pending:"https://wa.me/"+whatsapp}, auto_return:"approved", statement_descriptor:"Golfito" };
    const res = UrlFetchApp.fetch("https://api.mercadopago.com/checkout/preferences", { method:"POST", headers:{"Authorization":"Bearer "+MP_ACCESS_TOKEN,"Content-Type":"application/json"}, payload:JSON.stringify(payload), muteHttpExceptions:true });
    const data = JSON.parse(res.getContentText());
    if (!data.init_point) throw new Error("MP no devolvio init_point: "+res.getContentText());
    const link = (MODO_TEST_PLAN || MODO_TEST_ANALISIS) ? (data.sandbox_init_point||data.init_point) : data.init_point;
    return {ok:true,link,preferenceId:data.id,externalRef};
  } catch(err) { Logger.log("Error crearPreferenciaPago: "+err); return {ok:false,error:err.toString()}; }
}

function procesarPagoMP(paymentId) {
  try {
    Logger.log("Procesando pago MP: " + paymentId);
    const props = PropertiesService.getScriptProperties();
    const guardKey = "mp_processed_" + String(paymentId);
    if (props.getProperty(guardKey)) { Logger.log("Pago duplicado ignorado: " + paymentId); return; }
    props.setProperty(guardKey, String(Date.now()));
    try { const ss=SpreadsheetApp.getActiveSpreadsheet(); let l=ss.getSheetByName("ChatLog"); if (!l) l=ss.insertSheet("ChatLog"); l.appendRow([new Date(),"MP_WEBHOOK","entrante","payment",String(paymentId)]); } catch(le){}
    const res = UrlFetchApp.fetch("https://api.mercadopago.com/v1/payments/"+paymentId, {headers:{"Authorization":"Bearer "+MP_ACCESS_TOKEN},muteHttpExceptions:true});
    const pago = JSON.parse(res.getContentText());
    Logger.log("Pago MP status: "+pago.status+" | ref: "+pago.external_reference);
    if (pago.status!=="approved") { Logger.log("Pago no aprobado: "+pago.status); return; }
    const externalRef = pago.external_reference||""; if (!externalRef) { Logger.log("Sin external_reference"); return; }
    const partes = externalRef.split("_"); const whatsapp = partes[partes.length-1]; const codigoPlan = partes.slice(0,partes.length-1).join("_");
    const ss = SpreadsheetApp.getActiveSpreadsheet(); const sheet = ss.getSheetByName(SESIONES_SHEET); if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    for (let i=1;i<data.length;i++) {
      const rowWa=safeString(data[i][COL.WHATSAPP-1]); const rowCodigo=safeString(data[i][COL.CODIGO_PLAN-1]); const rowStatus=safeString(data[i][COL.PAGO_STATUS-1]);
      if (rowWa!==whatsapp) continue; if (rowCodigo && rowCodigo!==codigoPlan) continue; if (rowStatus==="pagado") continue;
      sheet.getRange(i+1,COL.PAGO_STATUS).setValue("pagado"); sheet.getRange(i+1,COL.PAGO_FECHA).setValue(new Date()); sheet.getRange(i+1,COL.PAGO_MONTO).setValue(pago.transaction_amount||""); sheet.getRange(i+1,COL.MP_PAYMENT_ID).setValue(String(paymentId));
      const ejvsplan = safeString(data[i][COL.EJVSPLAN-1]); const conv = obtenerConversacion(whatsapp); const nombre = obtenerNombreLead(whatsapp)||conv.nombre||"";
      if (ejvsplan==="2") {
        enviarMensajeWhatsApp(whatsapp, "\u2705 \u00a1Pago confirmado "+nombre+"! Analizando tu swing con IA, dame un momento...");
        const convConVideo = { ...conv, paso: "analizando_video", video_url1: safeString(data[i][COL.VIDEO_URL1-1]) || conv.video_url1 || "", ejvsplan: "2" };
        guardarConversacion(whatsapp, convConVideo); procesarAnalisisVideo(whatsapp, convConVideo, false);
      } else if (ejvsplan==="3") {
        enviarMensajeWhatsApp(whatsapp, "\u2705 \u00a1Pago confirmado "+nombre+"! \u26f3 Estamos preparando tu plan personalizado y te lo enviamos pronto por ac\u00e1.");
        const perfil={nivel:safeString(data[i][COL.NIVEL-1]),aspecto:safeString(data[i][COL.ASPECTO-1]),tiempo:"45 minutos"};
        notificarNuevoPlan(nombre,whatsapp,perfil,safeString(data[i][COL.VIDEO_URL1-1]),safeString(data[i][COL.COMENTARIOS_ALUMNO-1]));
        sheet.getRange(i+1,COL.STATUS).setValue("pendiente_manual");
      }
      Logger.log("Pago procesado OK para: "+whatsapp); return;
    }
    Logger.log("No se encontro sesion para: "+externalRef);
  } catch(err) { Logger.log("Error procesarPagoMP: "+err); }
}

function procesarReintegroDesdePanel(rowIndex) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SESIONES_SHEET); if (!sheet) return {ok:false,error:"No encontre Sesiones"};
    const row = sheet.getRange(rowIndex,1,1,SESIONES_HEADERS.length).getValues()[0];
    const paymentId = safeString(row[COL.MP_PAYMENT_ID-1]); const whatsapp = safeString(row[COL.WHATSAPP-1]); const nombre = obtenerNombreLead(whatsapp);
    if (!paymentId) return {ok:false,error:"No hay payment_id registrado para esta sesion"};
    const res = procesarReintegroMP(paymentId, null);
    if (res.ok) { sheet.getRange(rowIndex,COL.PAGO_STATUS).setValue("reintegrado"); enviarMensajeWhatsApp(whatsapp,"Hola "+nombre+" \u26f3 Te confirmamos que procesamos el reintegro de tu pago. En breve lo ver\u00e1s reflejado en tu cuenta."); }
    return res;
  } catch(err) { return {ok:false,error:err.toString()}; }
}

function verificarPagoAprobado(externalRef) {
  try {
    const res = UrlFetchApp.fetch("https://api.mercadopago.com/v1/payments/search?external_reference="+encodeURIComponent(externalRef)+"&sort=date_created&criteria=desc&limit=1", {headers:{"Authorization":"Bearer "+MP_ACCESS_TOKEN},muteHttpExceptions:true});
    const data = JSON.parse(res.getContentText()); const resultados = data.results||[];
    if (resultados.length===0) return {ok:false,motivo:"no_encontrado"};
    const pago = resultados[0];
    if (pago.status==="approved") return {ok:true,paymentId:pago.id,monto:pago.transaction_amount};
    return {ok:false,motivo:pago.status};
  } catch(err) { Logger.log("Error verificarPagoAprobado: "+err); return {ok:false,motivo:"error",error:err.toString()}; }
}

function procesarReintegroMP(paymentId, monto) {
  try {
    const payload = monto ? {amount:parseFloat(monto)} : {};
    const res = UrlFetchApp.fetch("https://api.mercadopago.com/v1/payments/"+paymentId+"/refunds", {method:"POST",headers:{"Authorization":"Bearer "+MP_ACCESS_TOKEN,"Content-Type":"application/json"},payload:JSON.stringify(payload),muteHttpExceptions:true});
    const data = JSON.parse(res.getContentText());
    if (data.id) return {ok:true,refundId:data.id,status:data.status};
    return {ok:false,error:data.message||res.getContentText()};
  } catch(err) { return {ok:false,error:err.toString()}; }
}

// ============================================
// HELPERS
// ============================================
// FIX BUG 2: obtenerUltimosVideosSesion busca en cualquier sesion (ejvsplan=2 incluido)
function obtenerUltimosVideosSesion(from) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SESIONES_SHEET); if (!sheet) return {url1:"",url2:""};
  const data = sheet.getDataRange().getValues();
  for (let i=data.length-1;i>=1;i--) {
    if (safeString(data[i][COL.WHATSAPP-1])!==from) continue;
    const url1=safeString(data[i][COL.VIDEO_URL1-1]);
    const url2=safeString(data[i][COL.VIDEO_URL2-1]);
    if (url1||url2) return {url1,url2};
  }
  return {url1:"",url2:""};
}
function sanitizarNombre(nombre) { return safeString(nombre).replace(/[^a-záéíóúüñA-ZÁÉÍÓÚÜÑ\s]/g,'').split(' ')[0].replace(/^\w/,c=>c.toUpperCase()); }
function seleccionarEjercicios(perfil, maxEjercicios, ejerciciosExcluidos=[]) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EXERCISES_SHEET); if (!sheet) throw new Error("No encontre "+EXERCISES_SHEET);
  const data = sheet.getDataRange().getValues(); if (data.length<2) return [];
  const ejercicios=[]; const pa=normalizeText(perfil.aspecto); const pn=normalizeText(perfil.nivel);
  for (let i=1;i<data.length;i++) { const e=rowToEjercicio(data[i]); if (!ejercicioValido(e)) continue; if (ejerciciosExcluidos.includes(e.id)) continue; if (e.aspecto===pa && e.nivel===pn) ejercicios.push(e); }
  return ejercicios.slice(0,maxEjercicios);
}
function rowToEjercicio(row) { return {id:safeString(row[0]),nombre:safeString(row[1]),instruccion:safeString(row[2]),aspecto:normalizeText(row[3]),nivel:normalizeText(row[4]),video_url:safeString(row[5]),tipo_ejercicio:safeString(row[6]),foco_tecnico:safeString(row[7]),etiqueta_1:safeString(row[8]),etiqueta_2:safeString(row[9]),etiqueta_3:safeString(row[10]),profesor:safeString(row[11])}; }
function ejercicioValido(e) { return e.id && e.nombre && e.nombre!=="pendiente" && e.aspecto!=="pendiente" && e.nivel!=="pendiente" && e.video_url && e.video_url!=="pendiente"; }
function buildPromptConEjercicios({nombre,perfil,ejercicios}) {
  const e=ejercicios[0];
  return "Actua como un coach profesional de golf para jugadores amateurs. Tu tono es amigable, directo y pedagógico.\n\nJugador:\n- Nombre: "+(nombre||"Jugador")+"\n- Nivel: "+perfil.nivel+"\n- Aspecto a trabajar: "+perfil.aspecto+"\n\nEl ejercicio asignado es: *"+e.nombre+"*\nInstruccion base: "+e.instruccion+"\n\nEscriba una explicacion simple y clara de como hacer este ejercicio, en exactamente 2 lineas. No repitas el nombre del ejercicio ni la instruccion base. No menciones la duracion. Usa asteriscos simples para negrita (*palabra*). Tono directo y accionable.\n\nGUÍA DE LENGUAJE:\n✔ SÍ PODÉS usar términos naturales del golf: grip, stance, putt, drive, approach, backswing, downswing, follow through, finish.\n❌ EVITÁ jerga compleja. Usá alternativas en español para: casting (\"soltar las manos antes de tiempo\"), over the top (\"tirar el cuerpo encima\"), chicken wing (\"doblar el codo hacia afuera\"), early extension (\"perder la postura\"), sway (\"balanceo lateral\"), scooping (\"cucharear\"), topping (\"pegarle arriba\"), chunk (\"pegarle pesado a la tierra\").";
}
function llamarOpenAI(prompt, apiKey) {
  const response = UrlFetchApp.fetch("https://api.openai.com/v1/responses", {method:"post",muteHttpExceptions:true,headers:{"Authorization":"Bearer "+apiKey,"Content-Type":"application/json"},payload:JSON.stringify({model:"gpt-4o-mini",input:prompt})});
  const statusCode=response.getResponseCode(); const body=response.getContentText();
  if (statusCode<200||statusCode>=300) throw new Error("Error OpenAI "+statusCode+": "+body);
  const data=JSON.parse(body); let text="";
  if (data.output_text) { text=data.output_text; }
  else if (Array.isArray(data.output)) { for (const item of data.output) { if (Array.isArray(item.content)) { for (const part of item.content) { if (part.text) text+=part.text; } } } }
  text=safeString(text).trim().replace(/\*\*([^*]+)\*\*/g,'*$1*');
  if (!text) throw new Error("OpenAI respondio sin texto utilizable.");
  return text;
}
function construirMensajeFinal(plan, ejercicios) {
  const e=ejercicios[0];
  return "\u26f3 *"+e.nombre+"*\n_"+e.instruccion+"_\n\n"+plan+"\n\n\ud83c\udfa5 "+e.video_url+"\n\n\ud83d\udca1 *Tip:* Enfocate en calidad antes que cantidad.\n\ud83c\udf10 _Activa los subtitulos en tu idioma en el video de YouTube._\n\n\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\nCualquier otra consulta escribinos \u26f3";
}
function construirMensajePendienteManual(nombre, perfil) { return "Hola "+(nombre||"")+" \u26f3 Recib\u00ed tu pedido. Estamos preparando tu plan personalizado y te lo enviamos pronto por ac\u00e1."; }
function construirMensajeUpsell(nombre) { return "Hola "+(nombre||"")+" \u26f3\n\nYa usaste tu ejercicio gratuito.\n\nPara seguir mejorando:\n- \ud83c\udfa5 *An\u00e1lisis de video con IA* \u2014 $ 5.000\n- \ud83d\udccb *Plan personalizado* \u2014 $ 15.000\n\nCualquier otra consulta escribinos \u26f3"; }
function obtenerEjerciciosYaEnviados(sesiones, whatsapp) {
  const h=[];
  for (let i=1;i<sesiones.length;i++) { const wa=safeString(sesiones[i][COL.WHATSAPP-1]); const st=safeString(sesiones[i][COL.STATUS-1]).toLowerCase(); const eid=safeString(sesiones[i][COL.EJERCICIO_GRATIS_ID-1]); const ejv=safeString(sesiones[i][COL.EJVSPLAN-1]); if (wa===whatsapp && st==="enviado" && ejv==="1" && eid && eid!=="UPSELL" && eid!=="SIN_EJERCICIO") h.push(eid); }
  return h;
}
function truncateForWhatsApp(text, maxLen) { const clean=safeString(text).trim(); if (clean.length<=maxLen) return clean; return clean.slice(0,maxLen-40)+"\n\n[Mensaje recortado por longitud]"; }
function safeString(value) { if (value===null||value===undefined) return ""; return String(value).trim(); }
function normalizeText(value) { return safeString(value).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g,""); }
function mapNivel(value, aspecto) {
  const v=safeString(value).toLowerCase(); const an=normalizeText(aspecto); const hcp=Number(v);
  if (an==="primeras veces"||v.includes("primeras")||v.includes("no tengo")||v.includes("sin handicap")) return "Primeras veces";
  if (!isNaN(hcp)&&v!=="") { if (hcp<10) return "Avanzado"; if (hcp<=20) return "Intermedio"; return "Principiante"; }
  if (v.includes("avan")) return "Avanzado"; if (v.includes("inter")) return "Intermedio"; return "Principiante";
}
function mapAspectoLead(value) {
  const v=safeString(value).toLowerCase();
  const map={"1":"Driver","2":"Hierros","3":"Approach","4":"Putting","5":"Bunker","6":"Primeras veces","driver":"Driver","hierros":"Hierros","approach":"Approach","putting":"Putting","bunker":"Bunker","primeras veces":"Primeras veces","primeras":"Primeras veces","otro":"Otro"};
  return map[v]||"Otro";
}
function mapTipoSolicitud(value) { const v=safeString(value).toLowerCase(); if (v==="3"||v.includes("plan")) return "plan_personalizado"; if (v==="2"||v.includes("analisis")||v.includes("análisis")||v.includes("video")) return "analisis_video"; return "ejercicio_gratis"; }
function mapTipoSolicitudPorEjvsplan(value) { const v=safeString(value); if (v==="3") return "plan_personalizado"; if (v==="2") return "analisis_video"; return "ejercicio_gratis"; }

// ============================================
// TESTS Y UTILIDADES
// ============================================
function testDrive() { let f; const fs=DriveApp.getFoldersByName("Golfito_Videos"); f=fs.hasNext()?fs.next():DriveApp.createFolder("Golfito_Videos"); Logger.log("Drive OK: "+f.getName()); }
function testEnvioMeta() { enviarMensajeWhatsApp("56975466327","Test desde Apps Script ✅"); }
function listarModelosGemini() { const res=UrlFetchApp.fetch("https://generativelanguage.googleapis.com/v1beta/models?key="+GEMINI_API_KEY,{muteHttpExceptions:true}); const data=JSON.parse(res.getContentText()); Logger.log(data.models?.map(m=>m.name).join("\n")||"Sin modelos"); }
function testLogDirecto() { const sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ChatLog"); Logger.log("Sheet: "+(sheet?"SI":"NO")); sheet.appendRow([new Date(),"test123","saliente","texto","prueba directa"]); Logger.log("OK"); }
function testGeminiV7() { const resultado=analizarSwingConGemini("10EFfw5W9gInxH4tDdPugTNVp2-Fnyztv"); Logger.log(JSON.stringify(resultado)); }
function limpiarCacheAnalisis() {
  const props=PropertiesService.getScriptProperties(); const keys=props.getKeys(); let eliminadas=0;
  keys.forEach(k => {
    if (k.startsWith('analisis_')||k.startsWith('media_processed_')) { props.deleteProperty(k); eliminadas++; }
    if (k.startsWith('mp_processed_')) { const ts=props.getProperty(k); if (!ts||Date.now()-parseInt(ts)>24*60*60*1000) { props.deleteProperty(k); eliminadas++; } }
    if (k.startsWith('panel_token_')) { const ts=props.getProperty(k); if (!ts||Date.now()-parseInt(ts)>8*60*60*1000) { props.deleteProperty(k); eliminadas++; } }
  });
  Logger.log('Limpieza: '+eliminadas+' entradas eliminadas');
}
