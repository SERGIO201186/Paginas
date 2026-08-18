/**
 * PÁGINAS CONMEMORATIVAS - Backend Apps Script
 * =============================================
 * Sistema de páginas de homenaje/condolencias para funerarias.
 *
 * HOJAS DE CÁLCULO REQUERIDAS (en el mismo Google Sheet):
 *  - "Servicios"     -> 1 fila por página/finado creada
 *  - "Condolencias"  -> mensajes de condolencia (con estado de moderación)
 *  - "Fotos"         -> fotos subidas (con estado de moderación)
 *  - "VideoLlamadas" -> videollamadas programadas por servicio
 *  - "Pagos"         -> registro de pagos/transferencia de propiedad
 *  - "Funerarias"    -> catálogo de funerarias clientes (login simple por código)
 *
 * CONFIGURACIÓN REQUERIDA (Apps Script > Configuración > Propiedades del script):
 *  - STRIPE_SECRET_KEY     -> sk_live_... o sk_test_...
 *  - STRIPE_WEBHOOK_SECRET -> token propio (no el de Stripe) para proteger la
 *                             URL del webhook; ver manejarWebhookStripe_ para
 *                             cómo generarlo y dónde configurarlo en Stripe
 *  - DRIVE_FOLDER_ID       -> ID de la carpeta de Drive raíz donde se guardan fotos
 *  - PRECIO_TRANSFERENCIA  -> monto en MXN, ej. "499"
 *  - ADMIN_MASTER_KEY      -> clave maestra para que TÚ (el desarrollador) administres todo
 */

// ============================================================
// CONFIGURACIÓN
// ============================================================

function getConfig_(clave) {
  return PropertiesService.getScriptProperties().getProperty(clave);
}

function getSheet_(nombre) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(nombre);
  if (!sheet) {
    sheet = ss.insertSheet(nombre);
    inicializarEncabezados_(sheet, nombre);
  }
  return sheet;
}

function inicializarEncabezados_(sheet, nombre) {
  const encabezados = {
    "Servicios": ["id", "funerariaId", "nombreFinado", "fechaNacimiento", "fechaDefuncion",
      "fotoUrl", "biografia", "fechaCepelio", "horaCepelio", "ubicacionCepelio",
      "fechaMisa", "horaMisa", "ubicacionMisa", "fechaSalida", "horaSalida", "ubicacionSalida",
      "reglasPublicacion", "estado", "slug", "creadoEn", "transferidoEn", "emailFamilia", "fotoPortadaUrl"],
    "Condolencias": ["id", "servicioId", "nombre", "mensaje", "estadoModeracion", "creadoEn"],
    "Fotos": ["id", "servicioId", "nombrePersona", "driveFileId", "urlPublica", "estadoModeracion", "creadoEn"],
    "VideoLlamadas": ["id", "servicioId", "titulo", "enlace", "fecha", "hora", "creadoEn", "linkActivo", "compradores"],
    "Pagos": ["id", "servicioId", "monto", "moneda", "stripeSessionId", "stripePaymentIntentId", "estado", "creadoEn", "completadoEn", "tipo", "email", "creditos"],
    "Funerarias": ["id", "nombre", "codigoAcceso", "activo", "creadoEn", "plan", "whatsapp", "email", "suscripcionActiva", "venceSuscripcion"],
    "Recuerdos": ["id", "servicioId", "nombre", "mensaje", "pieFoto", "driveFileId", "urlFoto", "estadoModeracion", "creadoEn", "reporteRevisado"],
    "Veladoras": ["id", "servicioId", "nombre", "creadoEn"],
    "Creditos": ["email", "servicioId", "saldo", "actualizadoEn"],
    "Gestos": ["id", "servicioId", "email", "nombre", "tipoGesto", "creditos", "direccion", "creadoEn"],
    "PedidosFlores": ["id", "servicioId", "email", "nombre", "direccion", "estado", "creadoEn"],
    "ArticulosTienda": ["id", "emoji", "nombre", "descripcion", "creditos", "tipo", "activo", "creadoEn"],
    "PaquetesCreditos": ["id", "creditos", "precio", "etiqueta", "activo", "creadoEn"],
    "ConfigMaestro": ["clave", "valor"],
    "Reacciones": ["id", "servicioId", "recuerdoId", "tipo", "nombre", "creadoEn"],
    "Comentarios": ["id", "servicioId", "recuerdoId", "nombre", "texto", "creadoEn"],
    "SolicitudesVL": ["id", "servicioId", "email", "nombre", "precio", "stripeSessionId", "estado", "creadoEn"],
    "ComisionesFuneraria": ["id", "funerariaId", "servicioId", "concepto", "montoTotal", "comision25", "estado", "creadoEn"],
    "SuscripcionesFuneraria": ["id", "funerariaId", "tipo", "monto", "stripeSessionId", "estado", "creadoEn", "venceEn"]
  };
  const fila = encabezados[nombre];
  if (fila) {
    sheet.getRange(1, 1, 1, fila.length).setValues([fila]);
    sheet.setFrozenRows(1);
  }
}

function generarId_() {
  return Utilities.getUuid();
}

// Devuelve el índice 0-based de una columna, agregándola al final de la hoja
// si todavía no existe (para migrar hojas creadas antes de agregar el campo).
function obtenerOCrearColumna_(sheet, nombreCol) {
  const enc = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let idx = enc.indexOf(nombreCol);
  if (idx < 0) {
    idx = enc.length;
    sheet.getRange(1, idx + 1).setValue(nombreCol);
  }
  return idx;
}

function limpiarValor_(val) {
  // Convierte Date objects de Sheets a string limpio antes de devolver al cliente
  if (!val && val !== 0) return '';
  if (val instanceof Date) {
    // Detectar si es solo hora (año 1899 = hora sin fecha en Sheets)
    if (val.getFullYear() === 1899 || val.getFullYear() === 1900) {
      const hh = String(val.getHours()).padStart(2, '0');
      const mm = String(val.getMinutes()).padStart(2, '0');
      return hh + ':' + mm;
    }
    // Es una fecha real
    const y = val.getFullYear();
    const m = String(val.getMonth() + 1).padStart(2, '0');
    const d = String(val.getDate()).padStart(2, '0');
    return y + '-' + m + '-' + d;
  }
  return String(val);
}

// ============================================================
// CICLO DE VIDA DE LA PÁGINA TRANSFERIDA
// ============================================================
// Una página transferida a la familia está "activa" (se puede comentar,
// subir fotos, encender veladoras, enviar gestos) durante los 9 días
// siguientes a la fecha de defunción (el novenario/rosario acostumbrado),
// y vuelve a activarse por una ventana de 9 días alrededor de cada
// aniversario luctuoso. El resto del tiempo queda en "solo_lectura": el
// contenido publicado se sigue mostrando, pero no se puede agregar nada
// nuevo. Los botones de contacto de la funeraria (WhatsApp, "planes a
// futuro") nunca dependen de esta fase — siempre están activos.
const VENTANA_ACTIVA_DIAS = 9;

function calcularFaseVisibilidad_(fechaDefuncionStr, estado) {
  // Solo las páginas ya entregadas a la familia siguen este ciclo. Mientras
  // la funeraria la sigue gestionando (en_curso / pendiente_pago) se
  // considera siempre activa, igual que hasta ahora.
  if (estado !== "transferido" || !fechaDefuncionStr) return "activo";

  const partes = String(fechaDefuncionStr).split("-").map(Number);
  if (partes.length !== 3 || partes.some(isNaN)) return "activo";
  const defuncion = new Date(partes[0], partes[1] - 1, partes[2]);

  const ahora = new Date();
  const hoy = new Date(ahora.getFullYear(), ahora.getMonth(), ahora.getDate());
  if (hoy < defuncion) return "activo";

  // Buscar el aniversario (año 0 = la fecha de defunción misma) más
  // reciente que ya pasó, y ver si "hoy" cae dentro de su ventana activa.
  let aniversario = new Date(defuncion.getFullYear(), defuncion.getMonth(), defuncion.getDate());
  while (true) {
    const siguiente = new Date(aniversario.getFullYear() + 1, aniversario.getMonth(), aniversario.getDate());
    if (siguiente > hoy) break;
    aniversario = siguiente;
  }
  const MS_POR_DIA = 24 * 60 * 60 * 1000;
  const diasTranscurridos = Math.round((hoy - aniversario) / MS_POR_DIA);
  return (diasTranscurridos >= 0 && diasTranscurridos <= VENTANA_ACTIVA_DIAS) ? "activo" : "solo_lectura";
}

// Usado por las acciones que reciben visitantes (condolencias, fotos,
// recuerdos, comentarios, reacciones, veladoras, gestos) para bloquear
// la escritura cuando la página transferida está en modo solo lectura.
function verificarPuedeEscribir_(servicioId) {
  if (!servicioId) return true;
  const sheet = getSheet_("Servicios");
  const filas = sheet.getDataRange().getValues();
  const enc = filas[0];
  const idIdx = enc.indexOf("id");
  const estadoIdx = enc.indexOf("estado");
  const fechaDefIdx = enc.indexOf("fechaDefuncion");
  for (let i = 1; i < filas.length; i++) {
    if (filas[i][idIdx] === servicioId) {
      const estado = filas[i][estadoIdx];
      const fechaDefuncion = limpiarValor_(filas[i][fechaDefIdx]);
      return calcularFaseVisibilidad_(fechaDefuncion, estado) === "activo";
    }
  }
  return true; // servicio no encontrado: no es este el lugar para rechazar
}

const MENSAJE_SOLO_LECTURA_ = "Esta página está en modo de solo lectura por ahora. Los recuerdos, fotos y mensajes publicados se siguen mostrando; vuelve a activarse automáticamente cada aniversario.";

function generarSlug_(nombreFinado) {
  const base = nombreFinado
    .toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9\s]/g, "")
    .trim()
    .replace(/\s+/g, "-");
  const sufijo = Math.random().toString(36).substring(2, 7);
  return `${base}-${sufijo}`;
}

// ============================================================
// ENTRADA PRINCIPAL (Web App)
// ============================================================

function doGet(e) {
  try {
    const accion = e.parameter.accion;
    let resultado;

    switch (accion) {
      case "obtenerServicioPublico":
        resultado = obtenerServicioPublico_(e.parameter.slug);
        break;
      case "obtenerCondolencias":
        resultado = obtenerCondolencias_(e.parameter.servicioId, e.parameter.soloAprobadas !== "false");
        break;
      case "obtenerFotos":
        resultado = obtenerFotos_(e.parameter.servicioId, e.parameter.soloAprobadas !== "false");
        break;
      case "obtenerVideoLlamadas":
        resultado = obtenerVideoLlamadas_(e.parameter.servicioId);
        break;
      case "listarServiciosFuneraria":
        resultado = listarServiciosFuneraria_(e.parameter.funerariaId, e.parameter.codigoAcceso);
        break;
      case "verificarPago":
        resultado = verificarEstadoPago_(e.parameter.servicioId);
        break;
      case "obtenerRecuerdos":
        resultado = obtenerRecuerdos_(e.parameter.servicioId);
        break;
      case "obtenerVeladoras":
        resultado = obtenerVeladoras_(e.parameter.servicioId);
        break;
      case "obtenerComentarios":
        resultado = obtenerComentarios_(e.parameter.recuerdoId);
        break;
      case "obtenerRecuerdosPendientes":
        resultado = obtenerRecuerdosPendientes_(e.parameter.servicioId);
        break;
      case "obtenerRecuerdosReportados":
        resultado = obtenerRecuerdosReportados_(e.parameter.servicioId);
        break;
      case "obtenerSaldo":
        resultado = obtenerSaldo_(e.parameter.email, e.parameter.servicioId);
        break;
      case "obtenerPedidosFlores":
        resultado = obtenerPedidosFlores_(e.parameter.servicioId, e.parameter.funerariaId, e.parameter.codigoAcceso);
        break;
      case "verificarMasterKey":
        resultado = { ok: e.parameter.masterKey === getConfig_("ADMIN_MASTER_KEY") };
        break;
      case "obtenerPlanFuneraria":
        resultado = obtenerPlanFuneraria_(e.parameter.funerariaId, e.parameter.codigoAcceso);
        break;
      case "obtenerComisionesFuneraria":
        resultado = obtenerComisionesFuneraria_(e.parameter.funerariaId, e.parameter.masterKey);
        break;
      case "obtenerSuscripcionesFuneraria":
        resultado = obtenerSuscripcionesFuneraria_(e.parameter.masterKey);
        break;
      case "listarTodasFunerarias":
        resultado = listarTodasFunerarias_(e.parameter.masterKey);
        break;
      case "obtenerArticulosTienda":
        resultado = obtenerArticulosTienda_(e.parameter.masterKey);
        break;
      case "obtenerPaquetesCreditos":
        resultado = obtenerPaquetesCreditos_(e.parameter.masterKey);
        break;
      case "obtenerConfigMaestro":
        resultado = obtenerConfigMaestro_(e.parameter.masterKey);
        break;
      case "obtenerStats":
        resultado = obtenerStats_(e.parameter.masterKey);
        break;
      case "obtenerPagos":
        resultado = obtenerPagos_(e.parameter.masterKey);
        break;
      case "obtenerTodosServicios":
        resultado = obtenerTodosServicios_(e.parameter.masterKey);
        break;
      case "obtenerVideollamadasPorEmail":
        resultado = obtenerVideollamadasPorEmail_(e.parameter.email, e.parameter.servicioId);
        break;
      case "verificarSolicitudVL":
        resultado = verificarSolicitudVL_(e.parameter.servicioId, e.parameter.email);
        break;
      case "obtenerSolicitudesVLMaster":
        resultado = obtenerSolicitudesVLMaster_(e.parameter.masterKey);
        break;
      // Acción especial solo para el administrador del sistema (tú)
      case "crearFuneraria":
        resultado = crearFuneraria_({
          masterKey: e.parameter.masterKey,
          nombre: e.parameter.nombre
        });
        break;
      default:
        resultado = { error: "Acción no reconocida: " + accion };
    }

    return responderJSON_(resultado);
  } catch (err) {
    return responderJSON_({ error: err.message, stack: err.stack });
  }
}

function doPost(e) {
  try {
    // Webhook de Stripe llega con content-type distinto, lo detectamos primero
    if (e.parameter.accion === "stripeWebhook" || e.parameter.accion === "webhookCreditosStripe") {
      return manejarWebhookStripe_(e);
    }

    const datos = JSON.parse(e.postData.contents);
    const accion = datos.accion;
    let resultado;

    switch (accion) {
      case "crearServicio":
        resultado = crearServicio_(datos);
        break;
      case "actualizarServicio":
        resultado = actualizarServicio_(datos);
        break;
      case "publicarCondolencia":
        resultado = publicarCondolencia_(datos);
        break;
      case "moderarCondolencia":
        resultado = moderarCondolencia_(datos);
        break;
      case "subirFoto":
        resultado = subirFoto_(datos);
        break;
      case "moderarFoto":
        resultado = moderarFoto_(datos);
        break;
      case "programarVideoLlamada":
        resultado = programarVideoLlamada_(datos);
        break;
      case "crearSesionPago":
        resultado = crearSesionPagoStripe_(datos);
        break;
      case "loginFuneraria":
        resultado = loginFuneraria_(datos);
        break;
      case "crearFuneraria":
        resultado = crearFuneraria_(datos);
        break;
      case "publicarRecuerdo":
        resultado = publicarRecuerdo_(datos);
        break;
      case "moderarRecuerdo":
        resultado = moderarRecuerdo_(datos);
        break;
      case "encenderVeladora":
        resultado = encenderVeladora_(datos);
        break;
      case "crearSesionCreditos":
        resultado = crearSesionCreditos_(datos);
        break;
      case "enviarGesto":
        resultado = enviarGesto_(datos);
        break;
      case "marcarFloresEntregadas":
        resultado = marcarFloresEntregadas_(datos);
        break;
      case "toggleEstadoServicio":
        resultado = toggleEstadoServicio_(datos);
        break;
      case "actualizarPortada":
        resultado = actualizarPortada_(datos);
        break;
      case "toggleFuneraria":
        resultado = toggleFuneraria_(datos);
        break;
      case "actualizarFuneraria":
        resultado = actualizarFuneraria_(datos);
        break;
      case "asignarPlan":
        resultado = asignarPlan_(datos);
        break;
      case "eliminarFuneraria":
        resultado = eliminarFuneraria_(datos);
        break;
      case "crearSesionSetup":
        resultado = crearSesionSetup_(datos);
        break;
      case "crearSesionMensualidad":
        resultado = crearSesionMensualidad_(datos);
        break;
      case "crearSesionPagina":
        resultado = crearSesionPagina_(datos);
        break;
      case "crearSesionActivacionPagina":
        resultado = crearSesionActivacionPagina_(datos);
        break;
      case "reintentarPagoActivacion":
        resultado = reintentarPagoActivacion_(datos);
        break;
      case "registrarComision":
        resultado = registrarComision_(datos);
        break;
      case "agregarArticuloTienda":
        resultado = agregarArticuloTienda_(datos);
        break;
      case "eliminarArticuloTienda":
        resultado = eliminarArticuloTienda_(datos);
        break;
      case "agregarPaqueteCreditos":
        resultado = agregarPaqueteCreditos_(datos);
        break;
      case "eliminarPaqueteCreditos":
        resultado = eliminarPaqueteCreditos_(datos);
        break;
      case "guardarConfigMaestro":
        resultado = guardarConfigMaestro_(datos);
        break;
      case "crearSesionVideollamada":
        resultado = crearSesionVideollamada_(datos);
        break;
      case "distribuirLinkVL":
        resultado = distribuirLinkVL_(datos);
        break;
      case "reaccionarRecuerdo":
        resultado = reaccionarRecuerdo_(datos);
        break;
      case "comentarRecuerdo":
        resultado = comentarRecuerdo_(datos);
        break;
      case "darCreditosGratis":
        resultado = darCreditosGratisPublico_(datos);
        break;
      case "verificarMasterKey":
        resultado = { ok: datos.masterKey === getConfig_("ADMIN_MASTER_KEY") };
        break;
      default:
        resultado = { error: "Acción no reconocida: " + accion };
    }

    return responderJSON_(resultado);
  } catch (err) {
    return responderJSON_({ error: err.message, stack: err.stack });
  }
}

// CRÍTICO: este header faltante rompe TODAS las respuestas en versiones anteriores.
function responderJSON_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// FUNERARIAS (clientes que usan el sistema)
// ============================================================

function crearFuneraria_(datos) {
  if (!validarMaster_(datos.masterKey)) {
    return { error: "No autorizado." };
  }
  const sheet = getSheet_("Funerarias");
  const id = generarId_();
  const codigoAcceso = Math.random().toString(36).substring(2, 8).toUpperCase();
  // Escribe por nombre de columna (no por posición fija) para no depender del
  // orden exacto del encabezado y no perder whatsapp/email como pasaba antes.
  const enc = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const fila = new Array(enc.length).fill("");
  const setCampo = (nombreCol, valor) => {
    const idx = enc.indexOf(nombreCol);
    if (idx >= 0) fila[idx] = valor;
  };
  setCampo("id", id);
  setCampo("nombre", datos.nombre || "");
  setCampo("codigoAcceso", codigoAcceso);
  setCampo("activo", true);
  setCampo("creadoEn", new Date().toISOString());
  setCampo("plan", "gratis");
  setCampo("whatsapp", datos.whatsapp || "");
  setCampo("email", datos.email || "");
  sheet.appendRow(fila);
  return { ok: true, id, codigoAcceso };
}

function loginFuneraria_(datos) {
  const sheet = getSheet_("Funerarias");
  const filas = sheet.getDataRange().getValues();
  for (let i = 1; i < filas.length; i++) {
    if (filas[i][2] === datos.codigoAcceso && filas[i][3] === true) {
      return { ok: true, funerariaId: filas[i][0], nombre: filas[i][1] };
    }
  }
  return { ok: false, error: "Código de acceso inválido o inactivo." };
}

function listarServiciosFuneraria_(funerariaId, codigoAcceso) {
  if (!validarFuneraria_(funerariaId, codigoAcceso)) {
    return { error: "No autorizado." };
  }
  const sheet = getSheet_("Servicios");
  const filas = sheet.getDataRange().getValues();
  const encabezados = filas[0];
  const servicios = [];
  for (let i = 1; i < filas.length; i++) {
    if (filas[i][1] === funerariaId) {
      const obj = {};
      encabezados.forEach((h, idx) => obj[h] = limpiarValor_(filas[i][idx]));
      servicios.push(obj);
    }
  }
  return { ok: true, servicios };
}

function validarFuneraria_(funerariaId, codigoAcceso) {
  const sheet = getSheet_("Funerarias");
  const filas = sheet.getDataRange().getValues();
  for (let i = 1; i < filas.length; i++) {
    if (filas[i][0] === funerariaId && filas[i][2] === codigoAcceso && filas[i][3] === true) {
      return true;
    }
  }
  return false;
}

// ============================================================
// SERVICIOS (páginas de homenaje)
// ============================================================

function crearServicio_(datos) {
  if (!validarFuneraria_(datos.funerariaId, datos.codigoAcceso)) {
    return { error: "No autorizado." };
  }
  const sheet = getSheet_("Servicios");
  const id = generarId_();
  const slug = generarSlug_(datos.nombreFinado);
  const ahora = new Date().toISOString();

  sheet.appendRow([
    id, datos.funerariaId, datos.nombreFinado, datos.fechaNacimiento || "", datos.fechaDefuncion || "",
    datos.fotoUrl || "", datos.biografia || "",
    datos.fechaCepelio || "", datos.horaCepelio || "", datos.ubicacionCepelio || "",
    datos.fechaMisa || "", datos.horaMisa || "", datos.ubicacionMisa || "",
    datos.fechaSalida || "", datos.horaSalida || "", datos.ubicacionSalida || "",
    datos.reglasPublicacion || "", "en_curso", slug, ahora, "", ""
  ]);

  return { ok: true, id, slug };
}

function actualizarServicio_(datos) {
  if (!validarFuneraria_(datos.funerariaId, datos.codigoAcceso)) {
    return { error: "No autorizado." };
  }
  const sheet = getSheet_("Servicios");
  const filas = sheet.getDataRange().getValues();
  const encabezados = filas[0];
  const idIdx = encabezados.indexOf("id");
  const estadoIdx = encabezados.indexOf("estado");

  for (let i = 1; i < filas.length; i++) {
    if (filas[i][idIdx] === datos.id) {
      // No permitir edición si ya fue transferido a la familia
      if (filas[i][estadoIdx] === "transferido") {
        return { error: "Esta página ya fue transferida a la familia y no puede editarse." };
      }
      const campos = ["nombreFinado", "fechaNacimiento", "fechaDefuncion", "fotoUrl", "biografia",
        "fechaCepelio", "horaCepelio", "ubicacionCepelio", "fechaMisa", "horaMisa", "ubicacionMisa",
        "fechaSalida", "horaSalida", "ubicacionSalida", "reglasPublicacion"];
      campos.forEach(campo => {
        if (datos[campo] !== undefined) {
          const colIdx = encabezados.indexOf(campo);
          sheet.getRange(i + 1, colIdx + 1).setValue(datos[campo]);
        }
      });
      return { ok: true };
    }
  }
  return { error: "Servicio no encontrado." };
}

function obtenerServicioPublico_(slug) {
  const sheet = getSheet_("Servicios");
  const filas = sheet.getDataRange().getValues();
  const encabezados = filas[0];
  for (let i = 1; i < filas.length; i++) {
    if (filas[i][encabezados.indexOf("slug")] === slug) {
      const obj = {};
      encabezados.forEach((h, idx) => obj[h] = limpiarValor_(filas[i][idx]));
      if (obj.estado === "pendiente_pago") {
        return { ok: false, error: "Esta página todavía no ha sido activada." };
      }
      obj.faseVisibilidad = calcularFaseVisibilidad_(obj.fechaDefuncion, obj.estado);
      const likes = contarLikesFotos_(obj.id);
      obj.likesPortada = likes.portada;
      obj.likesPerfil = likes.perfil;
      const funeraria = obtenerFunerariaPublica_(obj.funerariaId);
      obj.funerariaNombre = funeraria.nombre;
      obj.funerariaWhatsapp = funeraria.whatsapp;
      return { ok: true, servicio: obj };
    }
  }
  return { ok: false, error: "Página no encontrada." };
}

// Datos públicos y no sensibles de la funeraria (nombre y WhatsApp) para
// mostrar en la página del finado — nunca el código de acceso.
function obtenerFunerariaPublica_(funerariaId) {
  const sheet = getSheet_("Funerarias");
  const filas = sheet.getDataRange().getValues();
  const enc = filas[0];
  const idIdx = enc.indexOf("id");
  const nombreIdx = enc.indexOf("nombre");
  const whaIdx = enc.indexOf("whatsapp");
  for (let i = 1; i < filas.length; i++) {
    if (filas[i][idIdx] === funerariaId) {
      return {
        nombre: filas[i][nombreIdx] || "",
        whatsapp: whaIdx >= 0 ? (filas[i][whaIdx] || "") : ""
      };
    }
  }
  return { nombre: "", whatsapp: "" };
}

// Cuenta los "me gusta" de la foto de portada y de perfil de un servicio.
// Se guardan en la hoja Reacciones (la misma que usan los recuerdos) con
// recuerdoId sintético "foto_portada:<servicioId>" / "foto_perfil:<servicioId>",
// así se reutiliza reaccionarRecuerdo_ sin tocar el esquema de datos.
function contarLikesFotos_(servicioId) {
  const sheet = getSheet_("Reacciones");
  const filas = sheet.getDataRange().getValues();
  const enc = filas[0];
  const recIdIdx = enc.indexOf("recuerdoId");
  const tipoIdx = enc.indexOf("tipo");
  const idPortada = "foto_portada:" + servicioId;
  const idPerfil = "foto_perfil:" + servicioId;
  let portada = 0, perfil = 0;
  for (let i = 1; i < filas.length; i++) {
    if (filas[i][tipoIdx] !== "like") continue;
    if (filas[i][recIdIdx] === idPortada) portada++;
    else if (filas[i][recIdIdx] === idPerfil) perfil++;
  }
  return { portada, perfil };
}

// ============================================================
// CONDOLENCIAS
// ============================================================

function publicarCondolencia_(datos) {
  if (!datos.nombre || !datos.mensaje || !datos.servicioId) {
    return { error: "Faltan datos requeridos." };
  }
  if (!verificarPuedeEscribir_(datos.servicioId)) {
    return { error: MENSAJE_SOLO_LECTURA_ };
  }
  const sheet = getSheet_("Condolencias");
  const id = generarId_();
  // Toda condolencia entra en estado "pendiente" -> requiere aprobación de la funeraria
  sheet.appendRow([id, datos.servicioId, datos.nombre, datos.mensaje, "pendiente", new Date().toISOString()]);
  return { ok: true, mensaje: "Tu condolencia fue enviada y se publicará tras ser revisada." };
}

function moderarCondolencia_(datos) {
  if (!validarFuneraria_(datos.funerariaId, datos.codigoAcceso)) {
    return { error: "No autorizado." };
  }
  const sheet = getSheet_("Condolencias");
  const filas = sheet.getDataRange().getValues();
  const encabezados = filas[0];
  const idIdx = encabezados.indexOf("id");
  const estadoIdx = encabezados.indexOf("estadoModeracion");

  for (let i = 1; i < filas.length; i++) {
    if (filas[i][idIdx] === datos.id) {
      sheet.getRange(i + 1, estadoIdx + 1).setValue(datos.aprobada ? "aprobada" : "rechazada");
      return { ok: true };
    }
  }
  return { error: "Condolencia no encontrada." };
}

function obtenerCondolencias_(servicioId, soloAprobadas) {
  const sheet = getSheet_("Condolencias");
  const filas = sheet.getDataRange().getValues();
  const encabezados = filas[0];
  const resultado = [];
  for (let i = 1; i < filas.length; i++) {
    const obj = {};
    encabezados.forEach((h, idx) => obj[h] = filas[i][idx]);
    if (obj.servicioId === servicioId) {
      if (!soloAprobadas || obj.estadoModeracion === "aprobada") {
        resultado.push(obj);
      }
    }
  }
  resultado.sort((a, b) => new Date(b.creadoEn) - new Date(a.creadoEn));
  return { ok: true, condolencias: resultado };
}

// ============================================================
// FOTOS (Drive)
// ============================================================

function subirFoto_(datos) {
  if (!datos.servicioId || !datos.imagenBase64) {
    return { error: "Faltan datos requeridos." };
  }
  if (!verificarPuedeEscribir_(datos.servicioId)) {
    return { error: MENSAJE_SOLO_LECTURA_ };
  }
  try {
    const slug = obtenerSlugServicio_(datos.servicioId);
    const folderServicio = carpetaServicioEnDrive_(slug, "Fotos");

    const partesBase64 = datos.imagenBase64.split(",");
    const contenido = partesBase64.length > 1 ? partesBase64[1] : partesBase64[0];
    const tipoMime = datos.tipoMime || "image/jpeg";
    const bytes = Utilities.base64Decode(contenido);
    const blob = Utilities.newBlob(bytes, tipoMime, `foto_${Date.now()}.jpg`);

    const archivo = folderServicio.createFile(blob);
    archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const driveFileId = archivo.getId();
    const urlPublica = `https://drive.google.com/uc?export=view&id=${driveFileId}`;

    const sheet = getSheet_("Fotos");
    const id = generarId_();
    sheet.appendRow([id, datos.servicioId, datos.nombrePersona || "Anónimo", driveFileId, urlPublica, "pendiente", new Date().toISOString()]);

    return { ok: true, id, urlPublica, mensaje: "Tu foto fue enviada y se publicará tras ser revisada." };
  } catch (err) {
    return { error: "Error al subir foto: " + err.message };
  }
}

function moderarFoto_(datos) {
  if (!validarFuneraria_(datos.funerariaId, datos.codigoAcceso)) {
    return { error: "No autorizado." };
  }
  const sheet = getSheet_("Fotos");
  const filas = sheet.getDataRange().getValues();
  const encabezados = filas[0];
  const idIdx = encabezados.indexOf("id");
  const estadoIdx = encabezados.indexOf("estadoModeracion");
  const driveIdIdx = encabezados.indexOf("driveFileId");

  for (let i = 1; i < filas.length; i++) {
    if (filas[i][idIdx] === datos.id) {
      if (datos.aprobada) {
        sheet.getRange(i + 1, estadoIdx + 1).setValue("aprobada");
      } else {
        sheet.getRange(i + 1, estadoIdx + 1).setValue("rechazada");
        // Opcional: eliminar el archivo de Drive si se rechaza
        try {
          DriveApp.getFileById(filas[i][driveIdIdx]).setTrashed(true);
        } catch (e) { /* el archivo ya no existe, ignorar */ }
      }
      return { ok: true };
    }
  }
  return { error: "Foto no encontrada." };
}

function obtenerFotos_(servicioId, soloAprobadas) {
  const sheet = getSheet_("Fotos");
  const filas = sheet.getDataRange().getValues();
  const encabezados = filas[0];
  const resultado = [];
  for (let i = 1; i < filas.length; i++) {
    const obj = {};
    encabezados.forEach((h, idx) => obj[h] = filas[i][idx]);
    if (obj.servicioId === servicioId) {
      if (!soloAprobadas || obj.estadoModeracion === "aprobada") {
        resultado.push(obj);
      }
    }
  }
  resultado.sort((a, b) => new Date(b.creadoEn) - new Date(a.creadoEn));
  return { ok: true, fotos: resultado };
}

// ============================================================
// VIDEOLLAMADAS
// ============================================================

function programarVideoLlamada_(datos) {
  if (!validarFuneraria_(datos.funerariaId, datos.codigoAcceso)) {
    return { error: "No autorizado." };
  }
  const sheet = getSheet_("VideoLlamadas");
  const id = generarId_();
  sheet.appendRow([id, datos.servicioId, datos.titulo, datos.enlace, datos.fecha, datos.hora, new Date().toISOString()]);
  return { ok: true, id };
}

function obtenerVideoLlamadas_(servicioId) {
  const sheet = getSheet_("VideoLlamadas");
  const filas = sheet.getDataRange().getValues();
  const encabezados = filas[0];
  const resultado = [];
  for (let i = 1; i < filas.length; i++) {
    const obj = {};
    encabezados.forEach((h, idx) => obj[h] = filas[i][idx]);
    if (obj.servicioId === servicioId) {
      resultado.push(obj);
    }
  }
  resultado.sort((a, b) => new Date(a.fecha + "T" + a.hora) - new Date(b.fecha + "T" + b.hora));
  return { ok: true, videoLlamadas: resultado };
}

// ============================================================
// PAGOS (Stripe) - venta y transferencia de la página a la familia
// ============================================================

function crearSesionPagoStripe_(datos) {
  if (!datos.servicioId || !datos.emailFamilia) {
    return { error: "Faltan datos requeridos." };
  }

  const stripeKey = getConfig_("STRIPE_SECRET_KEY");
  const precio = getConfig_("PRECIO_TRANSFERENCIA") || "499";
  const urlBase = datos.urlExito || ScriptApp.getService().getUrl();

  const payload = {
    "mode": "payment",
    "payment_method_types[]": "card",
    "line_items[0][price_data][currency]": "mxn",
    "line_items[0][price_data][product_data][name]": "Transferencia de página conmemorativa",
    "line_items[0][price_data][unit_amount]": (parseFloat(precio) * 100).toString(),
    "line_items[0][quantity]": "1",
    "customer_email": datos.emailFamilia,
    "success_url": datos.urlExito + "?pago=exitoso&servicioId=" + datos.servicioId,
    "cancel_url": datos.urlCancelado + "?pago=cancelado",
    "metadata[servicioId]": datos.servicioId
  };

  const opciones = {
    method: "post",
    headers: { Authorization: "Bearer " + stripeKey },
    payload: payload,
    muteHttpExceptions: true
  };

  const respuesta = UrlFetchApp.fetch("https://api.stripe.com/v1/checkout/sessions", opciones);
  const resultado = JSON.parse(respuesta.getContentText());

  if (resultado.error) {
    return { error: "Error de Stripe: " + resultado.error.message };
  }

  // Registrar el intento de pago como "pendiente"
  const sheet = getSheet_("Pagos");
  const id = generarId_();
  sheet.appendRow([id, datos.servicioId, precio, "mxn", resultado.id, "", "pendiente", new Date().toISOString(), "", "transferencia", datos.emailFamilia, ""]);

  // Guardar también el email de la familia en el servicio
  const sheetServicios = getSheet_("Servicios");
  const filas = sheetServicios.getDataRange().getValues();
  const encabezados = filas[0];
  const idIdx = encabezados.indexOf("id");
  const emailIdx = encabezados.indexOf("emailFamilia");
  for (let i = 1; i < filas.length; i++) {
    if (filas[i][idIdx] === datos.servicioId) {
      sheetServicios.getRange(i + 1, emailIdx + 1).setValue(datos.emailFamilia);
      break;
    }
  }

  return { ok: true, urlPago: resultado.url, sessionId: resultado.id };
}

/**
 * Webhook de Stripe: confirma el pago y transfiere la propiedad de la página.
 *
 * IMPORTANTE — Apps Script no expone las cabeceras HTTP de la petición
 * entrante (no hay forma de leer "Stripe-Signature"), así que no es posible
 * validar la firma criptográfica que Stripe genera para cada webhook. Como
 * protección, esta URL exige un token secreto como parámetro — sin él,
 * cualquiera que adivine la URL podría simular un "pago completado" sin
 * pagar nada.
 *
 * Configuración necesaria:
 * 1. En Apps Script > Configuración > Propiedades del script, define
 *    STRIPE_WEBHOOK_SECRET con un valor largo y aleatorio (no tiene que ser
 *    el "Signing secret" de Stripe, es un token propio).
 * 2. En el dashboard de Stripe, configura la URL del webhook con ese mismo
 *    valor como parámetro "secreto":
 *    https://tu-script-url/exec?accion=stripeWebhook&secreto=TU_VALOR
 *
 * Si STRIPE_WEBHOOK_SECRET no está configurado, se acepta cualquier
 * petición (igual que antes) para no romper pagos existentes — pero eso
 * deja la vulnerabilidad abierta, así que se recomienda configurarlo.
 */
function manejarWebhookStripe_(e) {
  try {
    const secretoEsperado = getConfig_("STRIPE_WEBHOOK_SECRET");
    if (secretoEsperado && e.parameter.secreto !== secretoEsperado) {
      return responderJSON_({ error: "No autorizado." });
    }

    const evento = JSON.parse(e.postData.contents);

    if (evento.type === "checkout.session.completed") {
      const session = evento.data.object;
      const tipo = session.metadata && session.metadata.tipo;

      if (tipo === "creditos") {
        // Acreditar FuneralCredits al saldo del comprador
        const email = session.metadata.email;
        const creditos = parseInt(session.metadata.creditos) || 0;
        actualizarSaldo_(email, creditos);
        marcarPagoCompletado_(session);
        return responderJSON_({ ok: true });
      }

      if (tipo === "videollamada") {
        procesarPagoVideollamada_(session);
        return responderJSON_({ ok: true });
      }

      if (tipo === "setup" || tipo === "mensualidad") {
        marcarSuscripcionCompletada_(session);
        return responderJSON_({ ok: true });
      }

      if (tipo === "activacionPagina" || tipo === "pagina") {
        activarPaginaPagada_(session);
        return responderJSON_({ ok: true });
      }

      // Caso por defecto (sin metadata[tipo]): transferencia de la página conmemorativa a la familia
      const servicioId = session.metadata.servicioId;
      marcarPagoCompletado_(session);

      const sheetServicios = getSheet_("Servicios");
      const filasServicios = sheetServicios.getDataRange().getValues();
      const encServicios = filasServicios[0];
      const idIdx = encServicios.indexOf("id");
      const estadoServIdx = encServicios.indexOf("estado");
      const transferidoIdx = encServicios.indexOf("transferidoEn");

      for (let i = 1; i < filasServicios.length; i++) {
        if (filasServicios[i][idIdx] === servicioId) {
          sheetServicios.getRange(i + 1, estadoServIdx + 1).setValue("transferido");
          sheetServicios.getRange(i + 1, transferidoIdx + 1).setValue(new Date().toISOString());
          break;
        }
      }

      // Respaldo inmutable en Drive apenas se entrega la página — para que
      // si algo se rompe en la hoja compartida, la familia no pierda lo
      // que ya pagó por conservar.
      try { respaldarServicioADrive_(servicioId); } catch (errRespaldo) { /* no bloquear la transferencia por un fallo de respaldo */ }
    }

    return responderJSON_({ ok: true });
  } catch (err) {
    return responderJSON_({ error: err.message });
  }
}

// ============================================================
// RESPALDO DE PÁGINAS TRANSFERIDAS (aislar del resto del sistema)
// ============================================================
// Todo el sistema vive en una sola hoja de cálculo compartida entre todas
// las funerarias. Para que un problema general (cuota agotada, una fila
// borrada por error, un bug que escribe donde no debía) no le haga perder
// a una familia lo que ya pagó por conservar, cada página se respalda en
// un archivo JSON independiente dentro de Drive apenas se transfiere, y
// periódicamente después (ver respaldarTodasTransferidasDiario_). El
// respaldo es de solo lectura — no se usa para servir la página pública,
// solo para poder restaurarla a mano si algún día hace falta.

// Estructura de Drive: DRIVE_FOLDER_ID / Páginas / <slug> / {Fotos,Recuerdos,Respaldo}
// El slug ya es el nombre del finado (generarSlug_) más un sufijo corto para
// que sea único, así que la carpeta de cada página es reconocible a simple
// vista en Drive, en vez del id interno (un UUID sin sentido para una
// persona). El id interno se sigue usando como llave primaria en las hojas
// de cálculo — ahí sí necesita ser estable y opaco — pero ya no aparece en
// nombres de carpeta.
function obtenerOCrearSubcarpeta_(carpetaPadre, nombre) {
  const existentes = carpetaPadre.getFoldersByName(nombre);
  return existentes.hasNext() ? existentes.next() : carpetaPadre.createFolder(nombre);
}

function carpetaPaginasRaiz_() {
  const folderRaizId = getConfig_("DRIVE_FOLDER_ID");
  const folderRaiz = DriveApp.getFolderById(folderRaizId);
  return obtenerOCrearSubcarpeta_(folderRaiz, "Páginas");
}

// slugOId: idealmente el slug (nombre-sufijo); si no se tiene a la mano se
// acepta el id como respaldo para no fallar la operación.
function carpetaServicioEnDrive_(slugOId, subcarpeta) {
  const carpetaPagina = obtenerOCrearSubcarpeta_(carpetaPaginasRaiz_(), slugOId);
  return subcarpeta ? obtenerOCrearSubcarpeta_(carpetaPagina, subcarpeta) : carpetaPagina;
}

function obtenerSlugServicio_(servicioId) {
  const sheet = getSheet_("Servicios");
  const filas = sheet.getDataRange().getValues();
  const enc = filas[0];
  const idIdx = enc.indexOf("id");
  const slugIdx = enc.indexOf("slug");
  for (let i = 1; i < filas.length; i++) {
    if (filas[i][idIdx] === servicioId) return filas[i][slugIdx] || servicioId;
  }
  return servicioId;
}

function filasPorServicio_(nombreHoja, servicioId) {
  const sheet = getSheet_(nombreHoja);
  const filas = sheet.getDataRange().getValues();
  const enc = filas[0];
  const idx = enc.indexOf("servicioId");
  if (idx < 0) return [];
  const resultado = [];
  for (let i = 1; i < filas.length; i++) {
    if (filas[i][idx] === servicioId) {
      const obj = {};
      enc.forEach((h, j) => obj[h] = limpiarValor_(filas[i][j]));
      resultado.push(obj);
    }
  }
  return resultado;
}

// Genera y guarda en Drive una copia completa (JSON) de un servicio y todo
// su contenido asociado. Devuelve el ID del archivo creado.
function respaldarServicioADrive_(servicioId) {
  const sheetServicios = getSheet_("Servicios");
  const filas = sheetServicios.getDataRange().getValues();
  const enc = filas[0];
  const idIdx = enc.indexOf("id");
  let servicio = null;
  for (let i = 1; i < filas.length; i++) {
    if (filas[i][idIdx] === servicioId) {
      servicio = {};
      enc.forEach((h, j) => servicio[h] = limpiarValor_(filas[i][j]));
      break;
    }
  }
  if (!servicio) return null;

  const respaldo = {
    generadoEn: new Date().toISOString(),
    servicio: servicio,
    condolencias: filasPorServicio_("Condolencias", servicioId),
    fotos: filasPorServicio_("Fotos", servicioId),
    recuerdos: filasPorServicio_("Recuerdos", servicioId),
    veladoras: filasPorServicio_("Veladoras", servicioId),
    comentarios: (() => {
      // Los comentarios se ligan por recuerdoId, no por servicioId; se filtran
      // a partir de los recuerdos de este servicio.
      const sheet = getSheet_("Comentarios");
      const filasC = sheet.getDataRange().getValues();
      const encC = filasC[0];
      const recIdIdx = encC.indexOf("recuerdoId");
      const idsRecuerdos = new Set(filasPorServicio_("Recuerdos", servicioId).map(r => r.id));
      const resultado = [];
      for (let i = 1; i < filasC.length; i++) {
        if (idsRecuerdos.has(filasC[i][recIdIdx])) {
          const obj = {};
          encC.forEach((h, j) => obj[h] = limpiarValor_(filasC[i][j]));
          resultado.push(obj);
        }
      }
      return resultado;
    })()
  };

  // El respaldo nunca lo lee la página pública — solo sirve para restaurar a
  // mano si algún día hace falta — así que se guarda comprimido (gzip) para
  // ocupar el mínimo de espacio posible sin afectar en nada lo que ven los
  // visitantes.
  const carpeta = carpetaServicioEnDrive_(servicio.slug || servicioId, "Respaldo");
  const nombreArchivo = "respaldo_" + new Date().toISOString().replace(/[:.]/g, "-") + ".json.gz";
  const blobJson = Utilities.newBlob(JSON.stringify(respaldo), "application/json", "respaldo.json");
  const archivo = carpeta.createFile(Utilities.gzip(blobJson, nombreArchivo));

  limpiarRespaldosViejos_(carpeta);
  return archivo.getId();
}

// Conserva solo los respaldos más recientes de cada página para que el
// respaldo diario no crezca sin límite. RETENCION_RESPALDOS controla cuántos.
const RETENCION_RESPALDOS = 5;
function limpiarRespaldosViejos_(carpeta) {
  const archivos = [];
  const it = carpeta.getFiles();
  while (it.hasNext()) archivos.push(it.next());
  if (archivos.length <= RETENCION_RESPALDOS) return;
  archivos.sort((a, b) => b.getDateCreated() - a.getDateCreated());
  archivos.slice(RETENCION_RESPALDOS).forEach(a => a.setTrashed(true));
}

// Repite el respaldo de todas las páginas ya transferidas. Pensada para
// correr sola vía un trigger de tiempo (ver instalarTriggerRespaldoDiario_).
function respaldarTodasTransferidasDiario_() {
  const sheet = getSheet_("Servicios");
  const filas = sheet.getDataRange().getValues();
  const enc = filas[0];
  const idIdx = enc.indexOf("id");
  const estadoIdx = enc.indexOf("estado");
  for (let i = 1; i < filas.length; i++) {
    if (filas[i][estadoIdx] === "transferido") {
      try { respaldarServicioADrive_(filas[i][idIdx]); } catch (err) { /* seguir con las demás */ }
    }
  }
}

// Ejecutar UNA VEZ manualmente desde el editor de Apps Script (no se llama
// desde doGet/doPost) para instalar el respaldo diario automático.
function instalarTriggerRespaldoDiario_() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "respaldarTodasTransferidasDiario_") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("respaldarTodasTransferidasDiario_").timeBased().everyDays(1).atHour(3).create();
}

// Marca como "completado" la fila de Pagos que corresponda a esta sesión de Stripe
function marcarPagoCompletado_(session) {
  const sheetPagos = getSheet_("Pagos");
  const filasPagos = sheetPagos.getDataRange().getValues();
  const encPagos = filasPagos[0];
  const sessionIdIdx = encPagos.indexOf("stripeSessionId");
  const estadoIdx = encPagos.indexOf("estado");
  const completadoIdx = encPagos.indexOf("completadoEn");
  const piIdx = encPagos.indexOf("stripePaymentIntentId");

  for (let i = 1; i < filasPagos.length; i++) {
    if (filasPagos[i][sessionIdIdx] === session.id) {
      sheetPagos.getRange(i + 1, estadoIdx + 1).setValue("completado");
      sheetPagos.getRange(i + 1, completadoIdx + 1).setValue(new Date().toISOString());
      sheetPagos.getRange(i + 1, piIdx + 1).setValue(session.payment_intent);
      break;
    }
  }
}

// Marca en SuscripcionesFuneraria el cobro (setup/mensualidad) como completado
// y, si es mensualidad, refresca la vigencia de la funeraria un mes.
function marcarSuscripcionCompletada_(session) {
  const sheet = getSheet_("SuscripcionesFuneraria");
  const filas = sheet.getDataRange().getValues();
  const enc = filas[0];
  const sessionIdIdx = enc.indexOf("stripeSessionId");
  const estadoIdx = enc.indexOf("estado");
  for (let i = 1; i < filas.length; i++) {
    if (filas[i][sessionIdIdx] === session.id) {
      sheet.getRange(i + 1, estadoIdx + 1).setValue("completado");
      break;
    }
  }

  const tipo = session.metadata && session.metadata.tipo;
  const funerariaId = session.metadata && session.metadata.funerariaId;
  if (!funerariaId) return;

  const sheetF = getSheet_("Funerarias");
  const filasF = sheetF.getDataRange().getValues();
  const encF = filasF[0];
  const idIdxF = encF.indexOf("id");
  const suscActIdx = obtenerOCrearColumna_(sheetF, "suscripcionActiva");
  const venceIdx = obtenerOCrearColumna_(sheetF, "venceSuscripcion");
  for (let i = 1; i < filasF.length; i++) {
    if (String(filasF[i][idIdxF]) === String(funerariaId)) {
      sheetF.getRange(i + 1, suscActIdx + 1).setValue(true);
      if (tipo === "mensualidad") {
        const vence = new Date();
        vence.setMonth(vence.getMonth() + 1);
        sheetF.getRange(i + 1, venceIdx + 1).setValue(vence.toISOString());
      }
      break;
    }
  }
}

// Activa la página que estaba "pendiente_pago" (creada desde el panel de la
// funeraria) en cuanto Stripe confirma el pago de activación, y marca el
// cobro correspondiente como completado en SuscripcionesFuneraria.
function activarPaginaPagada_(session) {
  const servicioId = session.metadata && session.metadata.servicioId;
  if (servicioId) {
    const sheetS = getSheet_("Servicios");
    const filasS = sheetS.getDataRange().getValues();
    const encS = filasS[0];
    const idIdxS = encS.indexOf("id");
    const estadoIdxS = encS.indexOf("estado");
    for (let i = 1; i < filasS.length; i++) {
      if (filasS[i][idIdxS] === servicioId) {
        if (filasS[i][estadoIdxS] === "pendiente_pago") {
          sheetS.getRange(i + 1, estadoIdxS + 1).setValue("en_curso");
        }
        break;
      }
    }
  }

  const sheetSusc = getSheet_("SuscripcionesFuneraria");
  const filasSusc = sheetSusc.getDataRange().getValues();
  const encSusc = filasSusc[0];
  const sessionIdIdx = encSusc.indexOf("stripeSessionId");
  const estadoIdx = encSusc.indexOf("estado");
  for (let i = 1; i < filasSusc.length; i++) {
    if (filasSusc[i][sessionIdIdx] === session.id) {
      sheetSusc.getRange(i + 1, estadoIdx + 1).setValue("completado");
      break;
    }
  }
}

function verificarEstadoPago_(servicioId) {
  const sheet = getSheet_("Servicios");
  const filas = sheet.getDataRange().getValues();
  const encabezados = filas[0];
  const idIdx = encabezados.indexOf("id");
  const estadoIdx = encabezados.indexOf("estado");

  for (let i = 1; i < filas.length; i++) {
    if (filas[i][idIdx] === servicioId) {
      return { ok: true, estado: filas[i][estadoIdx] };
    }
  }
  return { error: "Servicio no encontrado." };
}

// ============================================================
// RECUERDOS (condolencias + fotos unificadas en una sola tabla)
// ============================================================

function publicarRecuerdo_(datos) {
  if (!datos.nombre || (!datos.mensaje && !datos.imagenBase64)) {
    return { error: "Faltan datos requeridos." };
  }
  if (!verificarPuedeEscribir_(datos.servicioId)) {
    return { error: MENSAJE_SOLO_LECTURA_ };
  }

  let urlFoto = "";
  let driveFileId = "";

  // Si viene foto, subirla a Drive
  if (datos.imagenBase64) {
    try {
      const slug = obtenerSlugServicio_(datos.servicioId);
      const folderServicio = carpetaServicioEnDrive_(slug, "Recuerdos");
      const partesBase64 = datos.imagenBase64.split(",");
      const contenido = partesBase64.length > 1 ? partesBase64[1] : partesBase64[0];
      const tipoMime = datos.tipoMime || "image/jpeg";
      const bytes = Utilities.base64Decode(contenido);
      const blob = Utilities.newBlob(bytes, tipoMime, "recuerdo_" + Date.now() + ".jpg");
      const archivo = folderServicio.createFile(blob);
      archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      driveFileId = archivo.getId();
      urlFoto = "https://drive.google.com/uc?export=view&id=" + driveFileId;
    } catch (err) {
      return { error: "Error al subir la foto: " + err.message };
    }
  }

  const sheet = getSheet_("Recuerdos");
  const id = generarId_();
  // Se publica de inmediato, sin aprobación previa. Si otros visitantes lo
  // reportan con suficientes dislikes, la funeraria podrá retirarlo después
  // (ver reaccionarRecuerdo_ y obtenerRecuerdosReportados_).
  sheet.appendRow([
    id, datos.servicioId, datos.nombre,
    datos.mensaje || "", datos.pieFoto || "",
    driveFileId, urlFoto,
    "aprobado", new Date().toISOString()
  ]);

  return { ok: true, mensaje: "Tu recuerdo ya está publicado." };
}

// Los recuerdos se publican solos (ver publicarRecuerdo_), así que esta función
// ya no "aprueba" contenido nuevo: se usa para resolver un reporte por dislikes.
// aprobada:true  -> se descarta el reporte y el recuerdo se queda publicado.
// aprobada:false -> se retira el recuerdo de la página pública.
function moderarRecuerdo_(datos) {
  if (!validarFuneraria_(datos.funerariaId, datos.codigoAcceso)) {
    return { error: "No autorizado." };
  }
  const sheet = getSheet_("Recuerdos");
  const filas = sheet.getDataRange().getValues();
  const enc = filas[0];
  const idIdx = enc.indexOf("id");
  const estadoIdx = enc.indexOf("estadoModeracion");
  const driveIdx = enc.indexOf("driveFileId");
  const pieFotoIdx = enc.indexOf("pieFoto");

  for (let i = 1; i < filas.length; i++) {
    if (filas[i][idIdx] === datos.id) {
      if (datos.aprobada) {
        sheet.getRange(i + 1, estadoIdx + 1).setValue("aprobado");
        const reporteIdx = obtenerOCrearColumna_(sheet, "reporteRevisado");
        sheet.getRange(i + 1, reporteIdx + 1).setValue(true);
      } else {
        sheet.getRange(i + 1, estadoIdx + 1).setValue("rechazado");
        // driveFileId guarda un emoji (no un archivo real) cuando la fila es un
        // gesto/vela premium — solo intentar borrar de Drive para fotos/recuerdos.
        const esGestoOVela = pieFotoIdx >= 0 && (filas[i][pieFotoIdx] === "gesto" || filas[i][pieFotoIdx] === "vela_premium");
        if (!esGestoOVela && filas[i][driveIdx]) {
          try { DriveApp.getFileById(filas[i][driveIdx]).setTrashed(true); } catch(e) {}
        }
      }
      return { ok: true };
    }
  }
  return { error: "Recuerdo no encontrado." };
}

function obtenerRecuerdos_(servicioId) {
  const sheet = getSheet_("Recuerdos");
  const filas = sheet.getDataRange().getValues();
  const enc = filas[0];
  const resultado = [];
  const pieFotoIdx = enc.indexOf("pieFoto");
  const driveFileIdIdx = enc.indexOf("driveFileId");

  // Contar likes/dislikes por recuerdo en un solo recorrido de la hoja Reacciones
  const conteoReacciones = {};
  const sheetReac = getSheet_("Reacciones");
  const filasReac = sheetReac.getDataRange().getValues();
  const encReac = filasReac[0];
  const recIdIdxReac = encReac.indexOf("recuerdoId");
  const tipoIdxReac = encReac.indexOf("tipo");
  for (let i = 1; i < filasReac.length; i++) {
    const rid = filasReac[i][recIdIdxReac];
    if (!conteoReacciones[rid]) conteoReacciones[rid] = { likes: 0, dislikes: 0 };
    if (filasReac[i][tipoIdxReac] === "like") conteoReacciones[rid].likes++;
    else if (filasReac[i][tipoIdxReac] === "dislike") conteoReacciones[rid].dislikes++;
  }

  // Contar comentarios por recuerdo
  const conteoComentarios = {};
  const sheetCom = getSheet_("Comentarios");
  const filasCom = sheetCom.getDataRange().getValues();
  const encCom = filasCom[0];
  const recIdIdxCom = encCom.indexOf("recuerdoId");
  for (let i = 1; i < filasCom.length; i++) {
    const rid = filasCom[i][recIdIdxCom];
    conteoComentarios[rid] = (conteoComentarios[rid] || 0) + 1;
  }

  for (let i = 1; i < filas.length; i++) {
    const obj = {};
    enc.forEach((h, idx) => obj[h] = limpiarValor_(filas[i][idx]));
    if (obj.servicioId === servicioId && obj.estadoModeracion === "aprobado") {
      // pieFoto guarda el tipo cuando es gesto/vela_premium
      // driveFileId guarda el emoji cuando es gesto
      if (obj.pieFoto === "gesto" || obj.pieFoto === "vela_premium") {
        obj.tipo = obj.pieFoto;
        obj.emoji = obj.driveFileId;
        obj.descripcion = obj.mensaje;
        obj.urlFoto = "";
        obj.pieFoto = "";
      } else {
        obj.tipo = "recuerdo";
        // Convertir URL de Drive al formato que funciona en web
        if (obj.urlFoto) {
          const m = obj.urlFoto.match(/[-\w]{25,}/);
          if (m) obj.urlFoto = "https://lh3.googleusercontent.com/d/" + m[0];
        }
      }
      obj.likes = (conteoReacciones[obj.id] && conteoReacciones[obj.id].likes) || 0;
      obj.dislikes = (conteoReacciones[obj.id] && conteoReacciones[obj.id].dislikes) || 0;
      obj.comentarios = conteoComentarios[obj.id] || 0;
      resultado.push(obj);
    }
  }
  resultado.sort((a, b) => new Date(b.creadoEn) - new Date(a.creadoEn));
  return { ok: true, recuerdos: resultado };
}

// Comentarios de un recuerdo (o de una foto de portada/perfil, usando el mismo
// esquema de id sintético "foto_portada:<servicioId>" / "foto_perfil:<servicioId>"
// que ya usa reaccionarRecuerdo_ para los likes de esas fotos).
function obtenerComentarios_(recuerdoId) {
  if (!recuerdoId) return { error: "Falta recuerdoId." };
  const sheet = getSheet_("Comentarios");
  const filas = sheet.getDataRange().getValues();
  const enc = filas[0];
  const comentarios = [];
  for (let i = 1; i < filas.length; i++) {
    const limpio = {};
    enc.forEach((h, idx) => limpio[h] = limpiarValor_(filas[i][idx]));
    if (limpio.recuerdoId === recuerdoId) comentarios.push(limpio);
  }
  comentarios.sort((a, b) => new Date(a.creadoEn) - new Date(b.creadoEn));
  return { ok: true, comentarios: comentarios };
}

function comentarRecuerdo_(datos) {
  if (!datos.recuerdoId || !datos.nombre || !datos.texto) {
    return { error: "Faltan datos." };
  }
  if (!verificarPuedeEscribir_(datos.servicioId)) {
    return { error: MENSAJE_SOLO_LECTURA_ };
  }
  const texto = String(datos.texto).trim().slice(0, 500);
  if (!texto) return { error: "El comentario no puede estar vacío." };
  const sheet = getSheet_("Comentarios");
  const id = generarId_();
  const creadoEn = new Date().toISOString();
  sheet.appendRow([id, datos.servicioId || "", datos.recuerdoId, datos.nombre, texto, creadoEn]);
  return { ok: true, id: id, nombre: datos.nombre, texto: texto, creadoEn: creadoEn };
}

// Listar recuerdos pendientes para el panel admin
function obtenerRecuerdosPendientes_(servicioId) {
  const sheet = getSheet_("Recuerdos");
  const filas = sheet.getDataRange().getValues();
  const enc = filas[0];
  const resultado = [];
  for (let i = 1; i < filas.length; i++) {
    const obj = {};
    enc.forEach((h, idx) => obj[h] = limpiarValor_(filas[i][idx]));
    if (obj.servicioId === servicioId && obj.estadoModeracion === "pendiente") {
      resultado.push(obj);
    }
  }
  return { ok: true, recuerdos: resultado };
}

// Recuerdos ya publicados que acumularon suficientes dislikes (ver
// UMBRAL_DISLIKES_ALERTA) como para que la funeraria los revise y decida si
// los retira. Se excluyen los que la funeraria ya marcó como revisados.
function obtenerRecuerdosReportados_(servicioId) {
  const sheetReac = getSheet_("Reacciones");
  const filasReac = sheetReac.getDataRange().getValues();
  const encReac = filasReac[0];
  const recIdIdxReac = encReac.indexOf("recuerdoId");
  const tipoIdxReac = encReac.indexOf("tipo");
  const dislikesPorRecuerdo = {};
  for (let i = 1; i < filasReac.length; i++) {
    if (filasReac[i][tipoIdxReac] === "dislike") {
      const rid = filasReac[i][recIdIdxReac];
      dislikesPorRecuerdo[rid] = (dislikesPorRecuerdo[rid] || 0) + 1;
    }
  }

  const sheet = getSheet_("Recuerdos");
  const filas = sheet.getDataRange().getValues();
  const enc = filas[0];
  const reporteIdx = enc.indexOf("reporteRevisado");
  const resultado = [];
  for (let i = 1; i < filas.length; i++) {
    const obj = {};
    enc.forEach((h, idx) => obj[h] = limpiarValor_(filas[i][idx]));
    if (obj.servicioId !== servicioId || obj.estadoModeracion !== "aprobado") continue;
    const dislikes = dislikesPorRecuerdo[obj.id] || 0;
    const yaRevisado = reporteIdx >= 0 && filas[i][reporteIdx] === true;
    if (dislikes >= UMBRAL_DISLIKES_ALERTA && !yaRevisado) {
      obj.dislikes = dislikes;
      resultado.push(obj);
    }
  }
  resultado.sort((a, b) => b.dislikes - a.dislikes);
  return { ok: true, recuerdos: resultado };
}

// ============================================================
// VELADORAS
// ============================================================

function encenderVeladora_(datos) {
  if (!datos.servicioId || !datos.nombre) {
    return { error: "Faltan datos." };
  }
  if (!verificarPuedeEscribir_(datos.servicioId)) {
    return { error: MENSAJE_SOLO_LECTURA_ };
  }
  const sheet = getSheet_("Veladoras");
  const id = generarId_();
  sheet.appendRow([id, datos.servicioId, datos.nombre, new Date().toISOString()]);

  // Devolver total actualizado y lista de nombres
  return obtenerVeladoras_(datos.servicioId);
}

function obtenerVeladoras_(servicioId) {
  const sheet = getSheet_("Veladoras");
  const filas = sheet.getDataRange().getValues();
  const enc = filas[0];
  const nombres = [];
  for (let i = 1; i < filas.length; i++) {
    const obj = {};
    enc.forEach((h, idx) => obj[h] = filas[i][idx]);
    if (obj.servicioId === servicioId) {
      nombres.push(String(obj.nombre));
    }
  }
  return { ok: true, total: nombres.length, nombres: nombres };
}

// ============================================================
// CRÉDITOS HUERTA
// ============================================================

function obtenerSaldo_(email, servicioId) {
  if (!email) return { error: "Email requerido." };
  const sheet = getSheet_("Creditos");
  const filas = sheet.getDataRange().getValues();
  const enc = filas[0];
  const emailIdx = enc.indexOf("email");
  const saldoIdx = enc.indexOf("saldo");
  for (let i = 1; i < filas.length; i++) {
    if (filas[i][emailIdx] === email) {
      return { ok: true, saldo: filas[i][saldoIdx] || 0 };
    }
  }
  return { ok: true, saldo: 0 };
}

function actualizarSaldo_(email, delta) {
  const sheet = getSheet_("Creditos");
  const filas = sheet.getDataRange().getValues();
  const enc = filas[0];
  const emailIdx = enc.indexOf("email");
  const saldoIdx = enc.indexOf("saldo");
  const fechaIdx = enc.indexOf("actualizadoEn");
  for (let i = 1; i < filas.length; i++) {
    if (filas[i][emailIdx] === email) {
      const nuevo = Math.max(0, (filas[i][saldoIdx] || 0) + delta);
      sheet.getRange(i + 1, saldoIdx + 1).setValue(nuevo);
      sheet.getRange(i + 1, fechaIdx + 1).setValue(new Date().toISOString());
      return nuevo;
    }
  }
  // No existe, crear fila nueva
  const nuevo = Math.max(0, delta);
  sheet.appendRow([email, "", nuevo, new Date().toISOString()]);
  return nuevo;
}

// Busca el precio/créditos REAL de un paquete en la hoja PaquetesCreditos
// (o en el catálogo por defecto si la hoja está vacía). Nunca hay que
// confiar en los valores de creditos/precio que manda el cliente: por eso
// esta función es la única fuente de verdad para crearSesionCreditos_.
function buscarPaqueteCreditos_(datos) {
  const sheet = getSheet_("PaquetesCreditos");
  const filas = sheet.getDataRange().getValues();
  const enc = filas[0];
  const idIdx = enc.indexOf("id");
  const credIdx = enc.indexOf("creditos");
  const precioIdx = enc.indexOf("precio");
  const activoIdx = enc.indexOf("activo");

  const catalogo = [];
  for (let i = 1; i < filas.length; i++) {
    if (activoIdx >= 0 && filas[i][activoIdx] === false) continue;
    catalogo.push({
      id: filas[i][idIdx],
      creditos: Number(filas[i][credIdx]) || 0,
      precio: Number(filas[i][precioIdx]) || 0
    });
  }
  const fuente = catalogo.length ? catalogo : PAQUETES_DEFAULT;

  if (datos.paqueteId) {
    const porId = fuente.find(p => String(p.id) === String(datos.paqueteId));
    if (porId) return porId;
  }
  // Compatibilidad con clientes que aún no tienen el frontend actualizado y
  // solo mandan creditos+precio (sin paqueteId): solo se acepta si coincide
  // EXACTO con un paquete real y activo — nunca se inventa un precio nuevo.
  return fuente.find(p =>
    Number(p.creditos) === Number(datos.creditos) && Number(p.precio) === Number(datos.precio)
  ) || null;
}

function crearSesionCreditos_(datos) {
  if (!datos.email) {
    return { error: "Faltan datos requeridos." };
  }
  const paquete = buscarPaqueteCreditos_(datos);
  if (!paquete || !paquete.creditos || !paquete.precio) {
    return { error: "Paquete no válido." };
  }
  const stripeKey = getConfig_("STRIPE_SECRET_KEY");
  const payload = {
    "mode": "payment",
    "payment_method_types[]": "card",
    "line_items[0][price_data][currency]": "mxn",
    "line_items[0][price_data][product_data][name]": paquete.creditos + " FuneralCredits",
    "line_items[0][price_data][unit_amount]": Math.round(paquete.precio * 100).toString(),
    "line_items[0][quantity]": "1",
    "customer_email": datos.email,
    "success_url": datos.urlExito,
    "cancel_url": datos.urlCancelado,
    "metadata[tipo]": "creditos",
    "metadata[email]": datos.email,
    "metadata[creditos]": String(paquete.creditos),
    "metadata[servicioId]": datos.servicioId || ""
  };
  const opciones = {
    method: "post",
    headers: { Authorization: "Bearer " + stripeKey },
    payload: payload,
    muteHttpExceptions: true
  };
  const respuesta = UrlFetchApp.fetch("https://api.stripe.com/v1/checkout/sessions", opciones);
  const resultado = JSON.parse(respuesta.getContentText());
  if (resultado.error) return { error: "Error de Stripe: " + resultado.error.message };

  // Registrar el intento de pago como "pendiente" (antes esto no se guardaba)
  const sheetPagos = getSheet_("Pagos");
  const idPago = generarId_();
  sheetPagos.appendRow([idPago, datos.servicioId || "", paquete.precio, "mxn", resultado.id, "", "pendiente", new Date().toISOString(), "", "creditos", datos.email, paquete.creditos]);

  return { ok: true, urlPago: resultado.url };
}

// ============================================================
// GESTOS
// ============================================================

// Alias heredados: navegadores con el frontend viejo en caché todavía pueden
// mandar estos nombres genéricos de categoría en vez del id real del
// artículo. Se mantienen aquí solo por compatibilidad hacia atrás.
const ALIAS_GESTOS_LEGACY = {
  rosas:        { creditos: 10,  tipo: "virtual", emoji: "🌹", nombre: "Rosa virtual" },
  paloma:       { creditos: 10,  tipo: "virtual", emoji: "🕊️", nombre: "Paloma de paz" },
  abrazo:       { creditos: 10,  tipo: "virtual", emoji: "💛", nombre: "Abrazo de luz" },
  corazon:      { creditos: 10,  tipo: "virtual", emoji: "💛", nombre: "Abrazo de luz" },
  arreglo_v:    { creditos: 30,  tipo: "virtual", emoji: "💐", nombre: "Arreglo virtual" },
  arreglo_virtual: { creditos: 30, tipo: "virtual", emoji: "💐", nombre: "Arreglo virtual" },
  vela_p:       { creditos: 25,  tipo: "premium", emoji: "🕯", nombre: "Vela premium" },
  vela_premium: { creditos: 25,  tipo: "premium", emoji: "🕯", nombre: "Vela premium" },
  flores_ch:    { creditos: 250, tipo: "fisico",  emoji: "🌺", nombre: "Arreglo físico chico" },
  flores_med:   { creditos: 400, tipo: "fisico",  emoji: "💐", nombre: "Arreglo físico mediano" },
  flores_gr:    { creditos: 600, tipo: "fisico",  emoji: "🌸", nombre: "Arreglo físico grande" },
  flores_fisicas: { creditos: 250, tipo: "fisico", emoji: "🌺", nombre: "Arreglo físico" }
};

// Fuente de verdad del costo real de un gesto: busca primero en la hoja
// ArticulosTienda (lo que configura el master y lo que ve el visitante en la
// tienda) y solo si no lo encuentra ahí recurre a los alias heredados. Nunca
// hay que confiar en los créditos que manda el cliente.
function buscarArticuloGesto_(tipoGesto) {
  const sheet = getSheet_("ArticulosTienda");
  const filas = sheet.getDataRange().getValues();
  const enc = filas[0];
  const idIdx = enc.indexOf("id");
  for (let i = 1; i < filas.length; i++) {
    if (String(filas[i][idIdx]) === String(tipoGesto)) {
      const obj = {};
      enc.forEach((h, idx) => obj[h] = filas[i][idx]);
      return {
        creditos: Number(obj.creditos) || 0,
        tipo: obj.tipo || "virtual",
        emoji: obj.emoji || "💐",
        nombre: obj.nombre || "",
        descripcion: obj.descripcion || obj.nombre || ""
      };
    }
  }
  return ALIAS_GESTOS_LEGACY[tipoGesto] || null;
}

function enviarGesto_(datos) {
  if (!datos.email || !datos.nombre || !datos.tipoGesto) {
    return { error: "Faltan datos requeridos." };
  }
  if (!verificarPuedeEscribir_(datos.servicioId)) {
    return { error: MENSAJE_SOLO_LECTURA_ };
  }
  const articulo = buscarArticuloGesto_(datos.tipoGesto);
  if (!articulo || !articulo.creditos) {
    return { error: "Artículo no válido." };
  }
  const costo = articulo.creditos; // nunca datos.creditos — ese lo manda el cliente

  const saldoActual = obtenerSaldo_(datos.email, datos.servicioId).saldo;
  if (saldoActual < costo) {
    return { error: "Saldo insuficiente. Tienes " + saldoActual + " créditos." };
  }
  // Descontar créditos
  const nuevoSaldo = actualizarSaldo_(datos.email, -costo);

  // Registrar gesto en hoja Gestos
  const sheetG = getSheet_("Gestos");
  const id = generarId_();
  sheetG.appendRow([
    id, datos.servicioId, datos.email, datos.nombre,
    datos.tipoGesto, costo, datos.direccion || "",
    new Date().toISOString()
  ]);

  // Si es un artículo físico, registrar también en PedidosFlores y notificar
  if (articulo.tipo === "fisico") {
    const sheetP = getSheet_("PedidosFlores");
    sheetP.appendRow([
      generarId_(), datos.servicioId, datos.email, datos.nombre,
      datos.direccion || "", "pendiente", new Date().toISOString()
    ]);
    // Notificar por email al administrador
    try {
      MailApp.sendEmail({
        to: Session.getActiveUser().getEmail(),
        subject: "Nuevo pedido de arreglo floral fisico",
        body: "Servicio: " + datos.servicioId + "\nCliente: " + datos.nombre + " (" + datos.email + ")\nDireccion: " + datos.direccion
      });
    } catch(e) {}
  }

  // Registrar como entrada en Recuerdos para que aparezca en el feed
  const sheetR = getSheet_("Recuerdos");
  const tipoPublico = articulo.tipo === "premium" ? "vela_premium" : "gesto";
  sheetR.appendRow([
    generarId_(), datos.servicioId, datos.nombre,
    articulo.descripcion || articulo.nombre || datos.tipoGesto,
    "", "", "",
    "aprobado", // los gestos se aprueban automáticamente
    new Date().toISOString()
  ]);

  // Actualizar el tipo y emoji en la última fila
  const lastRow = sheetR.getLastRow();
  // Agregar columnas extra si no existen usando una nota temporal
  // En su lugar guardamos tipo y emoji en campos existentes
  // pieFoto -> tipo, driveFileId -> emoji
  const enc = sheetR.getRange(1, 1, 1, sheetR.getLastColumn()).getValues()[0];
  const pieFotoIdx = enc.indexOf("pieFoto");
  const driveFileIdIdx = enc.indexOf("driveFileId");
  if (pieFotoIdx >= 0) sheetR.getRange(lastRow, pieFotoIdx + 1).setValue(tipoPublico);
  if (driveFileIdIdx >= 0) sheetR.getRange(lastRow, driveFileIdIdx + 1).setValue(articulo.emoji || "💐");

  // Registrar comision del 25% para la funeraria
  registrarComisionInterna_(datos.servicioId, costo, datos.tipoGesto);

  return { ok: true, nuevoSaldo };
}

// ============================================================
// PEDIDOS DE FLORES FÍSICAS
// ============================================================

function obtenerPedidosFlores_(servicioId, funerariaId, codigoAcceso) {
  if (!validarFuneraria_(funerariaId, codigoAcceso)) {
    return { error: "No autorizado." };
  }
  const sheet = getSheet_("PedidosFlores");
  const filas = sheet.getDataRange().getValues();
  const enc = filas[0];
  const resultado = [];
  for (let i = 1; i < filas.length; i++) {
    const obj = {};
    enc.forEach((h, idx) => obj[h] = limpiarValor_(filas[i][idx]));
    if (obj.servicioId === servicioId) {
      obj._fila = i + 1;
      resultado.push(obj);
    }
  }
  resultado.sort((a, b) => new Date(b.creadoEn) - new Date(a.creadoEn));
  return { ok: true, pedidos: resultado };
}

function marcarFloresEntregadas_(datos) {
  if (!validarFuneraria_(datos.funerariaId, datos.codigoAcceso)) {
    return { error: "No autorizado." };
  }
  const sheet = getSheet_("PedidosFlores");
  const filas = sheet.getDataRange().getValues();
  const enc = filas[0];
  const idIdx = enc.indexOf("id");
  const estadoIdx = enc.indexOf("estado");
  for (let i = 1; i < filas.length; i++) {
    if (filas[i][idIdx] === datos.id) {
      sheet.getRange(i + 1, estadoIdx + 1).setValue("entregado");
      return { ok: true };
    }
  }
  return { error: "Pedido no encontrado." };
}

// ============================================================
// FOTO DE PORTADA
// ============================================================

function actualizarPortada_(datos) {
  if (!validarFuneraria_(datos.funerariaId, datos.codigoAcceso)) {
    return { error: "No autorizado." };
  }
  const sheet = getSheet_("Servicios");
  const filas = sheet.getDataRange().getValues();
  const enc = filas[0];
  const idIdx = enc.indexOf("id");

  // Asegurar que la columna fotoPortadaUrl existe
  let portadaIdx = enc.indexOf("fotoPortadaUrl");
  if (portadaIdx < 0) {
    // Agregar columna al final
    portadaIdx = enc.length;
    sheet.getRange(1, portadaIdx + 1).setValue("fotoPortadaUrl");
  }

  for (let i = 1; i < filas.length; i++) {
    if (filas[i][idIdx] === datos.id) {
      sheet.getRange(i + 1, portadaIdx + 1).setValue(datos.fotoPortadaUrl || "");
      return { ok: true };
    }
  }
  return { error: "Servicio no encontrado." };
}

// ============================================================
// MASTER KEY HELPERS
// ============================================================

function validarMaster_(masterKey) {
  return masterKey === getConfig_("ADMIN_MASTER_KEY");
}

// ============================================================
// CONFIG MAESTRO (tabla clave-valor)
// ============================================================

function getConfigMaestro_(clave) {
  const sheet = getSheet_("ConfigMaestro");
  const filas = sheet.getDataRange().getValues();
  for (let i = 1; i < filas.length; i++) {
    if (filas[i][0] === clave) return filas[i][1];
  }
  return null;
}

function setConfigMaestro_(clave, valor) {
  const sheet = getSheet_("ConfigMaestro");
  const filas = sheet.getDataRange().getValues();
  for (let i = 1; i < filas.length; i++) {
    if (filas[i][0] === clave) {
      sheet.getRange(i + 1, 2).setValue(valor);
      return;
    }
  }
  sheet.appendRow([clave, valor]);
}

function obtenerConfigMaestro_(masterKey) {
  // __public__ permite leer solo precios públicos sin exponer datos sensibles
  if (masterKey !== "__public__" && !validarMaster_(masterKey)) return { error: "No autorizado." };
  if (masterKey === "__public__") {
    return {
      ok: true,
      precioVideollamada: getConfigMaestro_("precioVideollamada") || "150",
      publicidadActiva: getConfigMaestro_("publicidadActiva") === "true",
      publicidadImagenUrl: getConfigMaestro_("publicidadImagenUrl") || "",
      publicidadEnlace: getConfigMaestro_("publicidadEnlace") || "",
      publicidadTexto: getConfigMaestro_("publicidadTexto") || ""
    };
  }
  return {
    ok: true,
    whatsapp: getConfigMaestro_("whatsapp") || "",
    email: getConfigMaestro_("email") || "",
    precioTransferencia: getConfigMaestro_("precioTransferencia") || "499",
    creditosGratis: getConfigMaestro_("creditosGratis") || "10",
    precioVideollamada: getConfigMaestro_("precioVideollamada") || "150",
    publicidadActiva: getConfigMaestro_("publicidadActiva") === "true",
    publicidadImagenUrl: getConfigMaestro_("publicidadImagenUrl") || "",
    publicidadEnlace: getConfigMaestro_("publicidadEnlace") || "",
    publicidadTexto: getConfigMaestro_("publicidadTexto") || ""
  };
}

function guardarConfigMaestro_(datos) {
  if (!validarMaster_(datos.masterKey)) return { error: "No autorizado." };
  if (datos.whatsapp !== undefined) setConfigMaestro_("whatsapp", datos.whatsapp);
  if (datos.email !== undefined) setConfigMaestro_("email", datos.email);
  if (datos.precioTransferencia !== undefined) setConfigMaestro_("precioTransferencia", datos.precioTransferencia);
  if (datos.creditosGratis !== undefined) setConfigMaestro_("creditosGratis", datos.creditosGratis);
  if (datos.precioVideollamada !== undefined) setConfigMaestro_("precioVideollamada", datos.precioVideollamada);
  if (datos.publicidadActiva !== undefined) setConfigMaestro_("publicidadActiva", datos.publicidadActiva ? "true" : "false");
  if (datos.publicidadImagenUrl !== undefined) setConfigMaestro_("publicidadImagenUrl", datos.publicidadImagenUrl);
  if (datos.publicidadEnlace !== undefined) setConfigMaestro_("publicidadEnlace", datos.publicidadEnlace);
  if (datos.publicidadTexto !== undefined) setConfigMaestro_("publicidadTexto", datos.publicidadTexto);
  return { ok: true };
}

// ============================================================
// FUNERARIAS (funciones maestro)
// ============================================================

function listarTodasFunerarias_(masterKey) {
  if (!validarMaster_(masterKey)) return { error: "No autorizado." };
  const sheet = getSheet_("Funerarias");
  const filas = sheet.getDataRange().getValues();
  const enc = filas[0];
  const funerarias = [];
  // Deduplicar por nombre — solo mostrar la más reciente de cada nombre
  const vistos = {};
  for (let i = filas.length - 1; i >= 1; i--) {
    const obj = {};
    enc.forEach((h, idx) => obj[h] = limpiarValor_(filas[i][idx]));
    const key = obj.nombre + '_' + obj.codigoAcceso;
    if (!vistos[key]) {
      vistos[key] = true;
      obj.activo = filas[i][enc.indexOf("activo")];
      funerarias.unshift(obj);
    }
  }
  return { ok: true, funerarias };
}

function toggleFuneraria_(datos) {
  if (!validarMaster_(datos.masterKey)) return { error: "No autorizado." };
  const sheet = getSheet_("Funerarias");
  const filas = sheet.getDataRange().getValues();
  const enc = filas[0];
  const idIdx = enc.indexOf("id");
  const activoIdx = enc.indexOf("activo");
  for (let i = 1; i < filas.length; i++) {
    if (filas[i][idIdx] === datos.id) {
      sheet.getRange(i + 1, activoIdx + 1).setValue(datos.activo);
      return { ok: true };
    }
  }
  return { error: "No encontrada." };
}

// Edita nombre/whatsapp/email de una funeraria ya creada. Solo actualiza los
// campos que vengan definidos en datos, así se puede llamar para cambiar
// nada más el WhatsApp sin tocar el resto.
function actualizarFuneraria_(datos) {
  if (!validarMaster_(datos.masterKey)) return { error: "No autorizado." };
  if (!datos.id) return { error: "Falta el id de la funeraria." };
  const sheet = getSheet_("Funerarias");
  const filas = sheet.getDataRange().getValues();
  const enc = filas[0];
  const idIdx = enc.indexOf("id");
  for (let i = 1; i < filas.length; i++) {
    if (filas[i][idIdx] === datos.id) {
      // obtenerOCrearColumna_ agrega la columna si la hoja es de antes de que
      // existiera (por eso antes se guardaba "bien" pero nunca aparecía).
      ["nombre", "whatsapp", "email"].forEach(campo => {
        if (datos[campo] !== undefined) {
          const colIdx = obtenerOCrearColumna_(sheet, campo);
          sheet.getRange(i + 1, colIdx + 1).setValue(datos[campo]);
        }
      });
      return { ok: true };
    }
  }
  return { error: "Funeraria no encontrada." };
}

// ============================================================
// ARTÍCULOS TIENDA
// ============================================================

function obtenerArticulosTienda_(masterKey) {
  // Lectura pública permitida (sin masterKey)
  const sheet = getSheet_("ArticulosTienda");
  const filas = sheet.getDataRange().getValues();
  const enc = filas[0];
  const articulos = [];
  for (let i = 1; i < filas.length; i++) {
    const obj = {};
    enc.forEach((h, idx) => obj[h] = filas[i][idx]);
    const activo = obj.activo;
    // Sheets puede mandar boolean true, string "TRUE", o checkbox marcado
    if (activo === true || String(activo).toUpperCase() === "TRUE") {
      const limpio = {};
      enc.forEach((h, idx) => limpio[h] = limpiarValor_(filas[i][idx]));
      articulos.push(limpio);
    }
  }
  return { ok: true, articulos };
}

function agregarArticuloTienda_(datos) {
  if (!validarMaster_(datos.masterKey)) return { error: "No autorizado." };
  const sheet = getSheet_("ArticulosTienda");
  const id = generarId_();
  sheet.appendRow([id, datos.emoji||"🌹", datos.nombre, datos.descripcion||"", datos.creditos, datos.tipo||"virtual", true, new Date().toISOString()]);
  return { ok: true, id };
}

function eliminarArticuloTienda_(datos) {
  if (!validarMaster_(datos.masterKey)) return { error: "No autorizado." };
  const sheet = getSheet_("ArticulosTienda");
  const filas = sheet.getDataRange().getValues();
  const enc = filas[0];
  const idIdx = enc.indexOf("id");
  const activoIdx = enc.indexOf("activo");
  for (let i = 1; i < filas.length; i++) {
    if (String(filas[i][idIdx]) === String(datos.id)) {
      sheet.getRange(i + 1, activoIdx + 1).setValue(false);
      return { ok: true };
    }
  }
  return { error: "No encontrado." };
}

// ============================================================
// PAQUETES CRÉDITOS
// ============================================================

function obtenerPaquetesCreditos_(masterKey) {
  // Lectura pública permitida (sin masterKey)
  const sheet = getSheet_("PaquetesCreditos");
  const filas = sheet.getDataRange().getValues();
  const enc = filas[0];
  const paquetes = [];
  for (let i = 1; i < filas.length; i++) {
    const obj = {};
    enc.forEach((h, idx) => obj[h] = filas[i][idx]);
    const activo = obj.activo;
    if (activo === true || String(activo).toUpperCase() === "TRUE") {
      const limpio = {};
      enc.forEach((h, idx) => limpio[h] = limpiarValor_(filas[i][idx]));
      paquetes.push(limpio);
    }
  }
  return { ok: true, paquetes };
}

function agregarPaqueteCreditos_(datos) {
  if (!validarMaster_(datos.masterKey)) return { error: "No autorizado." };
  const sheet = getSheet_("PaquetesCreditos");
  const id = generarId_();
  sheet.appendRow([id, datos.creditos, datos.precio, datos.etiqueta||"", true, new Date().toISOString()]);
  return { ok: true, id };
}

function eliminarPaqueteCreditos_(datos) {
  if (!validarMaster_(datos.masterKey)) return { error: "No autorizado." };
  const sheet = getSheet_("PaquetesCreditos");
  const filas = sheet.getDataRange().getValues();
  const enc = filas[0];
  const idIdx = enc.indexOf("id");
  const activoIdx = enc.indexOf("activo");
  for (let i = 1; i < filas.length; i++) {
    if (filas[i][idIdx] === datos.id) {
      sheet.getRange(i + 1, activoIdx + 1).setValue(false);
      return { ok: true };
    }
  }
  return { error: "No encontrado." };
}

// ============================================================
// CRÉDITOS GRATIS AL ENTRAR
// ============================================================

function darCreditosGratis_(email, servicioId) {
  const clave = "cg_" + email + "_" + servicioId;
  // Solo dar créditos una vez por email+servicio
  const ya = getConfigMaestro_(clave);
  if (ya) return 0;

  // Tope de seguridad: no hay verificación de correo ni forma de limitar por
  // IP (Apps Script no expone la IP del visitante), así que sin este tope
  // alguien podría generar correos falsos sin límite para acumular créditos
  // gratis y canjearlos por gestos reales (incluyendo flores físicas).
  const claveContador = "cgTotal_" + servicioId;
  const totalRepartido = parseInt(getConfigMaestro_(claveContador) || "0");
  const topeMaximo = parseInt(getConfigMaestro_("topeCreditosGratisPorPagina") || "500");
  if (totalRepartido >= topeMaximo) return 0;

  const cantidad = parseInt(getConfigMaestro_("creditosGratis") || "10");
  actualizarSaldo_(email, cantidad);
  setConfigMaestro_(clave, "1");
  setConfigMaestro_(claveContador, String(totalRepartido + cantidad));
  return cantidad;
}

// ============================================================
// REACCIONES (like / dislike) CON ALERTAS
// ============================================================

// A partir de cuántos dislikes se alerta a la funeraria para que revise un recuerdo.
const UMBRAL_DISLIKES_ALERTA = 3;

// Correo que siempre recibe la alerta de recuerdo reportado, sin depender de que
// el email de la funeraria o el de ConfigMaestro estén configurados.
const EMAIL_ALERTA_RESPALDO = "Ventas@omnia-technology.com";

function reaccionarRecuerdo_(datos) {
  if (!datos.recuerdoId || !datos.tipo || !datos.nombre) {
    return { error: "Faltan datos." };
  }
  if (!verificarPuedeEscribir_(datos.servicioId)) {
    return { error: MENSAJE_SOLO_LECTURA_ };
  }
  // Verificar si ya reaccionó (por nombre + recuerdoId) para evitar duplicados
  const sheet = getSheet_("Reacciones");
  const filas = sheet.getDataRange().getValues();
  const enc = filas[0];
  const recIdIdx = enc.indexOf("recuerdoId");
  const nombreIdx = enc.indexOf("nombre");
  const tipoIdx = enc.indexOf("tipo");
  for (let i = 1; i < filas.length; i++) {
    if (filas[i][recIdIdx] === datos.recuerdoId && filas[i][nombreIdx] === datos.nombre) {
      return { ok: true, yaExistia: true }; // ya reaccionó, no duplicar
    }
  }
  const id = generarId_();
  sheet.appendRow([id, datos.servicioId, datos.recuerdoId, datos.tipo, datos.nombre, new Date().toISOString()]);

  // Al llegar exactamente al umbral, avisar una sola vez (no en cada dislike posterior)
  if (datos.tipo === "dislike") {
    let totalDislikes = 1; // cuenta la reacción que se acaba de agregar
    for (let i = 1; i < filas.length; i++) {
      if (filas[i][recIdIdx] === datos.recuerdoId && filas[i][tipoIdx] === "dislike") totalDislikes++;
    }
    if (totalDislikes === UMBRAL_DISLIKES_ALERTA) {
      alertarRecuerdoReportado_(datos.recuerdoId, datos.servicioId, totalDislikes);
    }
  }
  return { ok: true };
}

// Notifica a la funeraria dueña de la página (y, para supervisión, al administrador
// de la plataforma) que un recuerdo llegó al umbral de dislikes, para que decida si
// lo retira desde "Recuerdos reportados" en su panel.
function alertarRecuerdoReportado_(recuerdoId, servicioId, totalDislikes) {
  try {
    const sheetR = getSheet_("Recuerdos");
    const filasR = sheetR.getDataRange().getValues();
    const encR = filasR[0];
    const idIdxR = encR.indexOf("id");
    const nombreIdxR = encR.indexOf("nombre");
    let nombreAutor = "";
    for (let i = 1; i < filasR.length; i++) {
      if (filasR[i][idIdxR] === recuerdoId) { nombreAutor = filasR[i][nombreIdxR]; break; }
    }

    const sheetS = getSheet_("Servicios");
    const filasS = sheetS.getDataRange().getValues();
    const encS = filasS[0];
    const idIdxS = encS.indexOf("id");
    const funerariaIdIdxS = encS.indexOf("funerariaId");
    const nombreFinadoIdxS = encS.indexOf("nombreFinado");
    let funerariaId = "", nombreFinado = "";
    for (let i = 1; i < filasS.length; i++) {
      if (filasS[i][idIdxS] === servicioId) {
        funerariaId = filasS[i][funerariaIdIdxS];
        nombreFinado = filasS[i][nombreFinadoIdxS];
        break;
      }
    }

    let whatsappFuneraria = "", emailFuneraria = "";
    if (funerariaId) {
      const sheetF = getSheet_("Funerarias");
      const filasF = sheetF.getDataRange().getValues();
      const encF = filasF[0];
      const idIdxF = encF.indexOf("id");
      const whaIdxF = encF.indexOf("whatsapp");
      const emailIdxF = encF.indexOf("email");
      for (let i = 1; i < filasF.length; i++) {
        if (filasF[i][idIdxF] === funerariaId) {
          whatsappFuneraria = whaIdxF >= 0 ? filasF[i][whaIdxF] : "";
          emailFuneraria = emailIdxF >= 0 ? filasF[i][emailIdxF] : "";
          break;
        }
      }
    }

    const msg = "⚠️ Recuerdo reportado en \"" + (nombreFinado || servicioId) + "\"\n" +
      "De: " + (nombreAutor || "Anónimo") + "\n" +
      "Recibió " + totalDislikes + " \"no me gusta\". Revísalo en tu panel (Recuerdos reportados) y decide si lo retiras.";

    const destinatariosEmail = [emailFuneraria, getConfigMaestro_("email"), EMAIL_ALERTA_RESPALDO]
      .filter(Boolean)
      .filter((v, i, arr) => arr.indexOf(v) === i); // sin duplicados si coincide con algún otro correo
    destinatariosEmail.forEach(function(destino) {
      try {
        MailApp.sendEmail({ to: destino, subject: "Funeral360 · Recuerdo reportado - revisar", body: msg });
      } catch(e) {}
    });

    const destinatariosWA = [whatsappFuneraria, getConfigMaestro_("whatsapp")].filter(Boolean);
    destinatariosWA.forEach(function(numero) {
      try {
        UrlFetchApp.fetch("https://api.whatsapp.com/send?phone=" + numero + "&text=" + encodeURIComponent(msg), { muteHttpExceptions: true });
      } catch(e) {}
    });
  } catch (e) {
    // Nunca bloquear la reacción del visitante si falla el envío de alertas
  }
}

// ============================================================
// ESTADÍSTICAS GENERALES
// ============================================================

function obtenerPagos_(masterKey) {
  if (!validarMaster_(masterKey)) return { error: "No autorizado." };

  const sheet = getSheet_("Pagos");
  const filas = sheet.getDataRange().getValues();
  const enc = filas[0];
  const pagos = [];
  for (let i = 1; i < filas.length; i++) {
    const limpio = {};
    enc.forEach((h, idx) => limpio[h] = limpiarValor_(filas[i][idx]));
    pagos.push(limpio);
  }
  pagos.sort((a, b) => new Date(b.creadoEn) - new Date(a.creadoEn));
  return { ok: true, pagos: pagos.slice(0, 200) };
}

function obtenerStats_(masterKey) {
  if (!validarMaster_(masterKey)) return { error: "No autorizado." };

  const funerarias = Math.max(0, getSheet_("Funerarias").getLastRow() - 1);
  const servicios = Math.max(0, getSheet_("Servicios").getLastRow() - 1);
  const gestos = Math.max(0, getSheet_("Gestos").getLastRow() - 1);

  // ===== DESGLOSE DE INGRESOS =====
  var desglose = { paginas: 0, transferencias: 0, creditos: 0, videollamadas: 0 };
  var vlCount = 0;

  // 1. Pagos de Stripe (transferencias de páginas a familias + créditos)
  var ingresoCreditosReal = 0;
  try {
    const sheetPag = getSheet_("Pagos");
    const filasPag = sheetPag.getDataRange().getValues();
    const encPag = filasPag[0];
    const montoIdx = encPag.indexOf("monto");
    const estadoIdx = encPag.indexOf("estado");
    const tipoIdx = encPag.indexOf("tipo");
    for (let i = 1; i < filasPag.length; i++) {
      if (String(filasPag[i][estadoIdx]) === "completado") {
        const tipoPago = tipoIdx > -1 ? String(filasPag[i][tipoIdx] || "") : "";
        const monto = parseFloat(filasPag[i][montoIdx] || 0);
        if (tipoPago === "creditos") ingresoCreditosReal += monto;
        else desglose.transferencias += monto;
      }
    }
  } catch(e) {}

  // 2. Créditos comprados (ingreso real si ya existe historial de Pagos; si no, aproximación anterior)
  try {
    if (ingresoCreditosReal > 0) {
      desglose.creditos = Math.round(ingresoCreditosReal);
    } else {
      const sheetGestos = getSheet_("Gestos");
      const filasG = sheetGestos.getDataRange().getValues();
      const encG = filasG[0];
      const creditosIdx = encG.indexOf("creditos");
      for (let i = 1; i < filasG.length; i++) {
        desglose.creditos += parseFloat(filasG[i][creditosIdx] || 0);
      }
      // 1 crédito = ~$1 MXN promedio (ajuste por paquetes)
      desglose.creditos = Math.round(desglose.creditos * 1.08); // factor promedio de paquetes
    }
  } catch(e) {}

  // 3. Videollamadas pagadas
  try {
    const sheetVL = getSheet_("SolicitudesVL");
    const filasVL = sheetVL.getDataRange().getValues();
    const encVL = filasVL[0];
    const precioVLIdx = encVL.indexOf("precio");
    const estadoVLIdx = encVL.indexOf("estado");
    for (let i = 1; i < filasVL.length; i++) {
      if (String(filasVL[i][estadoVLIdx]) === "completado") {
        desglose.videollamadas += parseFloat(filasVL[i][precioVLIdx] || 0);
        vlCount++;
      }
    }
  } catch(e) {}

  // 4. Páginas activadas (suscripciones tipo pagina)
  try {
    const sheetSus = getSheet_("SuscripcionesFuneraria");
    const filasSus = sheetSus.getDataRange().getValues();
    const encSus = filasSus[0];
    const tipoIdx = encSus.indexOf("tipo");
    const montoIdx = encSus.indexOf("monto");
    const estadoIdx = encSus.indexOf("estado");
    for (let i = 1; i < filasSus.length; i++) {
      if (String(filasSus[i][estadoIdx]) === "completado") {
        const tipo = String(filasSus[i][tipoIdx] || "");
        const monto = parseFloat(filasSus[i][montoIdx] || 0);
        if (tipo === "pagina") desglose.paginas += monto;
        else if (tipo === "setup" || tipo === "mensualidad") desglose.transferencias += monto;
      }
    }
  } catch(e) {}

  // ===== COMISIONES 25% POR FUNERARIA =====
  var totalComisionesPendientes = 0;
  var comisionesPorFuneraria = [];
  try {
    const sheetCom = getSheet_("ComisionesFuneraria");
    const filasCom = sheetCom.getDataRange().getValues();
    const encCom = filasCom[0];
    const funIdx = encCom.indexOf("funerariaId");
    const comIdx = encCom.indexOf("comision25");
    const estadoIdx = encCom.indexOf("estado");
    const creditosComIdx = encCom.indexOf("montoTotal");

    // Agrupar por funeraria
    var agrupado = {};
    for (let i = 1; i < filasCom.length; i++) {
      const funId = String(filasCom[i][funIdx] || "");
      if (!agrupado[funId]) agrupado[funId] = { comisionPendiente: 0, totalGestos: 0, totalCreditos: 0 };
      if (String(filasCom[i][estadoIdx]) === "pendiente") {
        agrupado[funId].comisionPendiente += parseFloat(filasCom[i][comIdx] || 0);
        agrupado[funId].totalGestos++;
        agrupado[funId].totalCreditos += parseFloat(filasCom[i][creditosComIdx] || 0);
        totalComisionesPendientes += parseFloat(filasCom[i][comIdx] || 0);
      }
    }

    // Obtener nombres de funerarias
    const sheetFun = getSheet_("Funerarias");
    const filasFun = sheetFun.getDataRange().getValues();
    const encFun = filasFun[0];
    const idFunIdx = encFun.indexOf("id");
    const nomFunIdx = encFun.indexOf("nombre");
    const nombresMap = {};
    for (let i = 1; i < filasFun.length; i++) {
      nombresMap[String(filasFun[i][idFunIdx])] = String(filasFun[i][nomFunIdx]);
    }

    Object.keys(agrupado).forEach(funId => {
      if (agrupado[funId].comisionPendiente > 0) {
        comisionesPorFuneraria.push({
          funerariaId: funId,
          nombre: nombresMap[funId] || funId.substring(0,8),
          comisionPendiente: Math.round(agrupado[funId].comisionPendiente * 100) / 100,
          totalGestos: agrupado[funId].totalGestos,
          totalCreditos: Math.round(agrupado[funId].totalCreditos)
        });
      }
    });
  } catch(e) {}

  // Pedidos flores pendientes
  var pedidosPendientes = [];
  try {
    const sheetP = getSheet_("PedidosFlores");
    const filasP = sheetP.getDataRange().getValues();
    const encP = filasP[0];
    const estadoPIdx = encP.indexOf("estado");
    for (let i = 1; i < filasP.length; i++) {
      if (String(filasP[i][estadoPIdx]) === "pendiente") {
        const obj = {};
        encP.forEach((h, idx) => obj[h] = limpiarValor_(filasP[i][idx]));
        pedidosPendientes.push(obj);
      }
    }
  } catch(e) {}

  return {
    ok: true,
    funerarias,
    servicios,
    gestos,
    videollamadas: vlCount,
    desglose: {
      paginas: Math.round(desglose.paginas),
      transferencias: Math.round(desglose.transferencias),
      creditos: Math.round(desglose.creditos),
      videollamadas: Math.round(desglose.videollamadas)
    },
    totalComisionesPendientes: Math.round(totalComisionesPendientes * 100) / 100,
    comisionesPorFuneraria,
    pedidosFloresPendientes: pedidosPendientes
  };
}

function darCreditosGratisPublico_(datos) {
  if (!datos.email || !datos.servicioId) return { error: "Faltan datos." };
  const cantidad = darCreditosGratis_(datos.email, datos.servicioId);
  return { ok: true, cantidad };
}

// ============================================================
// VIDEOLLAMADA CON PAGO
// ============================================================

function crearSesionVideollamada_(datos) {
  if (!datos.email || !datos.servicioId) {
    return { error: "Faltan datos requeridos." };
  }
  // Precio fijo del lado del servidor — nunca confiar en datos.precio, que
  // viene del cliente y podría venir alterado.
  const precio = PRECIOS.videollamada;
  const stripeKey = getConfig_("STRIPE_SECRET_KEY");
  const payload = {
    "mode": "payment",
    "payment_method_types[]": "card",
    "line_items[0][price_data][currency]": "mxn",
    "line_items[0][price_data][product_data][name]": "Videollamada conmemorativa - Acceso completo",
    "line_items[0][price_data][unit_amount]": Math.round(precio * 100).toString(),
    "line_items[0][quantity]": "1",
    "customer_email": datos.email,
    "success_url": datos.urlExito,
    "cancel_url": datos.urlCancelado,
    "metadata[tipo]": "videollamada",
    "metadata[email]": datos.email,
    "metadata[nombre]": datos.nombre || "",
    "metadata[servicioId]": datos.servicioId,
    "metadata[precio]": String(precio)
  };
  const opciones = {
    method: "post",
    headers: { Authorization: "Bearer " + stripeKey },
    payload: payload,
    muteHttpExceptions: true
  };
  const respuesta = UrlFetchApp.fetch("https://api.stripe.com/v1/checkout/sessions", opciones);
  const resultado = JSON.parse(respuesta.getContentText());
  if (resultado.error) return { error: "Error de Stripe: " + resultado.error.message };

  // Registrar solicitud pendiente
  const sheet = getSheet_("SolicitudesVL");
  const id = generarId_();
  sheet.appendRow([id, datos.servicioId, datos.email, datos.nombre||"", precio, resultado.id, "pendiente", new Date().toISOString()]);

  return { ok: true, urlPago: resultado.url };
}

// El webhook de Stripe llama a esto cuando se confirma el pago de videollamada
function procesarPagoVideollamada_(session) {
  const servicioId = session.metadata.servicioId;
  const email = session.metadata.email;
  const nombre = session.metadata.nombre || "";
  const precio = session.metadata.precio || "0";

  // Marcar solicitud como completada
  const sheet = getSheet_("SolicitudesVL");
  const filas = sheet.getDataRange().getValues();
  const enc = filas[0];
  const sessionIdx = enc.indexOf("stripeSessionId");
  const estadoIdx = enc.indexOf("estado");
  for (let i = 1; i < filas.length; i++) {
    if (filas[i][sessionIdx] === session.id) {
      sheet.getRange(i + 1, estadoIdx + 1).setValue("completado");
      break;
    }
  }

  // Notificar al administrador
  const whatsapp = getConfigMaestro_("whatsapp");
  const emailAdmin = getConfigMaestro_("email");
  const msg = "Nueva solicitud de videollamada pagada\nServicio: " + servicioId + "\nCliente: " + nombre + " (" + email + ")\nMonto: $" + precio + " MXN\n\nPor favor crea el Meet y distribúyelo desde el panel maestro.";

  if (emailAdmin) {
    try {
      MailApp.sendEmail({
        to: emailAdmin,
        subject: "Funeral360 · Videollamada pagada - Acción requerida",
        body: msg
      });
    } catch(e) {}
  }
  if (whatsapp) {
    try {
      UrlFetchApp.fetch(
        "https://api.whatsapp.com/send?phone=" + whatsapp + "&text=" + encodeURIComponent(msg),
        { muteHttpExceptions: true }
      );
    } catch(e) {}
  }

  // Enviar confirmación al comprador
  try {
    MailApp.sendEmail({
      to: email,
      subject: "Funeral360 · Tu acceso a la videollamada conmemorativa",
      body: "Hola " + nombre + ",\n\nTu pago fue confirmado. Cuando la videollamada esté lista, recibirás el link de conexión en este correo.\n\nGracias por acompañar a la familia en este momento."
    });
  } catch(e) {}
}

// Listar solicitudes de VL por servicio (para el panel admin)
function obtenerVideollamadasPorEmail_(email, servicioId) {
  const sheet = getSheet_("SolicitudesVL");
  const filas = sheet.getDataRange().getValues();
  const enc = filas[0];
  const resultado = [];
  for (let i = 1; i < filas.length; i++) {
    const obj = {};
    enc.forEach((h, idx) => obj[h] = limpiarValor_(filas[i][idx]));
    if (obj.servicioId === servicioId && obj.email === email) resultado.push(obj);
  }
  return { ok: true, solicitudes: resultado };
}

// ============================================================
// TOGGLE ESTADO DE SERVICIO (activo / inactivo / vendido)
// ============================================================

function toggleEstadoServicio_(datos) {
  if (!validarMaster_(datos.masterKey)) return { error: "No autorizado." };
  const sheet = getSheet_("Servicios");
  const filas = sheet.getDataRange().getValues();
  const enc = filas[0];
  const idIdx = enc.indexOf("id");
  const estadoIdx = enc.indexOf("estado");
  for (let i = 1; i < filas.length; i++) {
    if (filas[i][idIdx] === datos.id) {
      sheet.getRange(i + 1, estadoIdx + 1).setValue(datos.nuevoEstado);
      return { ok: true };
    }
  }
  return { error: "Servicio no encontrado." };
}

function obtenerTodosServicios_(masterKey) {
  if (!validarMaster_(masterKey)) return { error: "No autorizado." };
  const sheet = getSheet_("Servicios");
  const filas = sheet.getDataRange().getValues();
  const enc = filas[0];
  const servicios = [];
  for (let i = 1; i < filas.length; i++) {
    const obj = {};
    enc.forEach((h, idx) => obj[h] = limpiarValor_(filas[i][idx]));
    servicios.push(obj);
  }
  servicios.sort((a,b) => new Date(b.creadoEn) - new Date(a.creadoEn));
  return { ok: true, servicios };
}

// ============================================================
// PLANES DE FUNERARIA
// ============================================================
// Plan "gratis": solo crear páginas básicas, sin tienda, sin videollamada de pago
// Plan "pro": acceso completo a todo

function asignarPlan_(datos) {
  if (!validarMaster_(datos.masterKey)) return { error: "No autorizado." };
  const sheet = getSheet_("Funerarias");
  const filas = sheet.getDataRange().getValues();
  const enc = filas[0];
  const idIdx = enc.indexOf("id");

  // Asegurar que existe columna plan
  let planIdx = enc.indexOf("plan");
  if (planIdx < 0) {
    planIdx = enc.length;
    sheet.getRange(1, planIdx + 1).setValue("plan");
  }

  for (let i = 1; i < filas.length; i++) {
    if (String(filas[i][idIdx]) === String(datos.funerariaId)) {
      sheet.getRange(i + 1, planIdx + 1).setValue(datos.plan || "gratis");
      return { ok: true };
    }
  }
  return { error: "Funeraria no encontrada." };
}

function obtenerPlanFuneraria_(funerariaId, codigoAcceso) {
  // Acceso público permitido para mostrar/ocultar features en la página del finado
  if (codigoAcceso !== '__publico__' && !validarFuneraria_(funerariaId, codigoAcceso)) {
    return { error: "No autorizado." };
  }
  const sheet = getSheet_("Funerarias");
  const filas = sheet.getDataRange().getValues();
  const enc = filas[0];
  const idIdx = enc.indexOf("id");
  const planIdx = enc.indexOf("plan");
  for (let i = 1; i < filas.length; i++) {
    if (String(filas[i][idIdx]) === String(funerariaId)) {
      const plan = planIdx >= 0 ? String(filas[i][planIdx] || "gratis") : "gratis";
      return { ok: true, plan };
    }
  }
  return { ok: true, plan: "gratis" };
}

// ============================================================
// MODELO C - SUSCRIPCIONES Y COBROS
// ============================================================

// Precios oficiales del sistema
var PRECIOS = {
  setup: 500,
  mensualidad: 99,
  pagina: 149,
  videollamada: 299,
  comisionFuneraria: 0.25  // 25% para la funeraria
};

// Artículos predeterminados del catálogo (se cargan si ArticulosTienda está vacío)
var CATALOGO_DEFAULT = [
  {emoji:"🌹", nombre:"Rosa virtual", descripcion:"Una rosa en su memoria", creditos:10, tipo:"virtual"},
  {emoji:"🕊️", nombre:"Paloma de paz", descripcion:"Que descanse en paz", creditos:10, tipo:"virtual"},
  {emoji:"💛", nombre:"Abrazo de luz", descripcion:"Un abrazo para la familia", creditos:10, tipo:"virtual"},
  {emoji:"💐", nombre:"Arreglo virtual", descripcion:"Arreglo floral especial en la página", creditos:30, tipo:"virtual"},
  {emoji:"🕯", nombre:"Vela premium", descripcion:"Vela con luz dorada especial", creditos:25, tipo:"premium"},
  {emoji:"🌺", nombre:"Arreglo físico chico", descripcion:"Arreglo real entregado en el velorio (~$300 MXN)", creditos:250, tipo:"fisico"},
  {emoji:"💐", nombre:"Arreglo físico mediano", descripcion:"Arreglo real mediano entregado en el velorio (~$480 MXN)", creditos:400, tipo:"fisico"},
  {emoji:"🌸", nombre:"Arreglo físico grande", descripcion:"Arreglo real grande entregado en el velorio (~$720 MXN)", creditos:600, tipo:"fisico"}
];

var PAQUETES_DEFAULT = [
  {creditos:50, precio:60, etiqueta:""},
  {creditos:120, precio:130, etiqueta:"Más popular"},
  {creditos:300, precio:299, etiqueta:"Mejor valor"}
];

function inicializarCatalogoPredeterminado_() {
  const sheet = getSheet_("ArticulosTienda");
  if (sheet.getLastRow() > 1) return; // ya tiene datos
  CATALOGO_DEFAULT.forEach(a => {
    sheet.appendRow([generarId_(), a.emoji, a.nombre, a.descripcion, a.creditos, a.tipo, true, new Date().toISOString()]);
  });
}

function inicializarPaquetesPredeterminados_() {
  const sheet = getSheet_("PaquetesCreditos");
  if (sheet.getLastRow() > 1) return;
  PAQUETES_DEFAULT.forEach(p => {
    sheet.appendRow([generarId_(), p.creditos, p.precio, p.etiqueta, true, new Date().toISOString()]);
  });
}

// ============================================================
// COBROS MODELO C (Stripe)
// ============================================================

function crearSesionStripe_(descripcion, monto, email, urlExito, urlCancelado, metadata) {
  const stripeKey = getConfig_("STRIPE_SECRET_KEY");
  const payload = {
    "mode": "payment",
    "payment_method_types[]": "card",
    "line_items[0][price_data][currency]": "mxn",
    "line_items[0][price_data][product_data][name]": descripcion,
    "line_items[0][price_data][unit_amount]": (parseFloat(monto) * 100).toString(),
    "line_items[0][quantity]": "1",
    "success_url": urlExito,
    "cancel_url": urlCancelado
  };
  if (email) payload["customer_email"] = email;
  Object.keys(metadata).forEach(k => payload["metadata[" + k + "]"] = String(metadata[k]));
  const resp = UrlFetchApp.fetch("https://api.stripe.com/v1/checkout/sessions", {
    method: "post",
    headers: { Authorization: "Bearer " + stripeKey },
    payload: payload,
    muteHttpExceptions: true
  });
  return JSON.parse(resp.getContentText());
}

function crearSesionSetup_(datos) {
  if (!validarMaster_(datos.masterKey)) return { error: "No autorizado." };
  const r = crearSesionStripe_(
    "Setup Páginas Conmemorativas - " + (datos.nombreFuneraria || ""),
    PRECIOS.setup,
    datos.email || "",
    datos.urlExito || "",
    datos.urlCancelado || "",
    { tipo: "setup", funerariaId: datos.funerariaId, masterKey: datos.masterKey }
  );
  if (r.error) return { error: "Stripe: " + r.error.message };

  const sheet = getSheet_("SuscripcionesFuneraria");
  sheet.appendRow([generarId_(), datos.funerariaId, "setup", PRECIOS.setup, r.id, "pendiente", new Date().toISOString(), ""]);
  return { ok: true, urlPago: r.url };
}

function crearSesionMensualidad_(datos) {
  if (!validarMaster_(datos.masterKey)) return { error: "No autorizado." };
  const r = crearSesionStripe_(
    "Mensualidad Páginas Conmemorativas - " + (datos.nombreFuneraria || ""),
    PRECIOS.mensualidad,
    datos.email || "",
    datos.urlExito || "",
    datos.urlCancelado || "",
    { tipo: "mensualidad", funerariaId: datos.funerariaId }
  );
  if (r.error) return { error: "Stripe: " + r.error.message };

  const sheet = getSheet_("SuscripcionesFuneraria");
  const vence = new Date();
  vence.setMonth(vence.getMonth() + 1);
  sheet.appendRow([generarId_(), datos.funerariaId, "mensualidad", PRECIOS.mensualidad, r.id, "pendiente", new Date().toISOString(), vence.toISOString()]);
  return { ok: true, urlPago: r.url };
}

function crearSesionPagina_(datos) {
  // Cobro manual generado por el master (link para enviar aparte) — se mantiene
  // por compatibilidad; el flujo normal es crearSesionActivacionPagina_, que la
  // propia funeraria dispara desde su panel al crear una página nueva.
  if (!validarMaster_(datos.masterKey)) return { error: "No autorizado." };
  const r = crearSesionStripe_(
    "Activación de página conmemorativa",
    PRECIOS.pagina,
    datos.email || "",
    datos.urlExito || "",
    datos.urlCancelado || "",
    { tipo: "pagina", funerariaId: datos.funerariaId, servicioId: datos.servicioId || "" }
  );
  if (r.error) return { error: "Stripe: " + r.error.message };
  registrarSuscripcionPendiente_(datos.funerariaId, "pagina", PRECIOS.pagina, r.id, datos.servicioId || "");
  return { ok: true, urlPago: r.url };
}

// Cobro automático de la activación de una página nueva, disparado por la
// propia funeraria desde su panel (admin/panel.html) al guardar por primera
// vez los datos del finado. La página se guarda de inmediato con estado
// "pendiente_pago" para no perder lo capturado; el webhook de Stripe la pasa
// a "en_curso" (visible al público) solo cuando el pago se confirma.
function crearSesionActivacionPagina_(datos) {
  if (!validarFuneraria_(datos.funerariaId, datos.codigoAcceso)) {
    return { error: "No autorizado." };
  }
  if (!datos.nombreFinado) {
    return { error: "El nombre del finado es obligatorio." };
  }

  const sheet = getSheet_("Servicios");
  const id = generarId_();
  const slug = generarSlug_(datos.nombreFinado);
  const ahora = new Date().toISOString();

  sheet.appendRow([
    id, datos.funerariaId, datos.nombreFinado, datos.fechaNacimiento || "", datos.fechaDefuncion || "",
    datos.fotoUrl || "", datos.biografia || "",
    datos.fechaCepelio || "", datos.horaCepelio || "", datos.ubicacionCepelio || "",
    datos.fechaMisa || "", datos.horaMisa || "", datos.ubicacionMisa || "",
    datos.fechaSalida || "", datos.horaSalida || "", datos.ubicacionSalida || "",
    datos.reglasPublicacion || "", "pendiente_pago", slug, ahora, "", ""
  ]);

  const r = crearSesionStripe_(
    "Activación de página conmemorativa - " + datos.nombreFinado,
    PRECIOS.pagina,
    emailFuneraria_(datos.funerariaId),
    datos.urlExito || "",
    datos.urlCancelado || "",
    { tipo: "activacionPagina", funerariaId: datos.funerariaId, servicioId: id }
  );
  if (r.error) return { error: "Stripe: " + r.error.message };

  registrarSuscripcionPendiente_(datos.funerariaId, "pagina", PRECIOS.pagina, r.id, id);
  return { ok: true, id: id, slug: slug, urlPago: r.url };
}

// Reintentar el cobro de una página que quedó "pendiente_pago" (por ejemplo si
// la funeraria cerró la ventana de Stripe sin terminar de pagar).
function reintentarPagoActivacion_(datos) {
  if (!validarFuneraria_(datos.funerariaId, datos.codigoAcceso)) {
    return { error: "No autorizado." };
  }
  const sheet = getSheet_("Servicios");
  const filas = sheet.getDataRange().getValues();
  const enc = filas[0];
  const idIdx = enc.indexOf("id");
  const funIdx = enc.indexOf("funerariaId");
  const estadoIdx = enc.indexOf("estado");
  const nombreIdx = enc.indexOf("nombreFinado");
  let nombreFinado = "";
  let encontrado = false;
  for (let i = 1; i < filas.length; i++) {
    if (filas[i][idIdx] === datos.servicioId && String(filas[i][funIdx]) === String(datos.funerariaId)) {
      encontrado = true;
      nombreFinado = filas[i][nombreIdx];
      if (filas[i][estadoIdx] !== "pendiente_pago") {
        return { error: "Esta página ya está activa." };
      }
      break;
    }
  }
  if (!encontrado) return { error: "Página no encontrada." };

  const r = crearSesionStripe_(
    "Activación de página conmemorativa - " + nombreFinado,
    PRECIOS.pagina,
    emailFuneraria_(datos.funerariaId),
    datos.urlExito || "",
    datos.urlCancelado || "",
    { tipo: "activacionPagina", funerariaId: datos.funerariaId, servicioId: datos.servicioId }
  );
  if (r.error) return { error: "Stripe: " + r.error.message };

  registrarSuscripcionPendiente_(datos.funerariaId, "pagina", PRECIOS.pagina, r.id, datos.servicioId);
  return { ok: true, urlPago: r.url };
}

function emailFuneraria_(funerariaId) {
  const sheet = getSheet_("Funerarias");
  const filas = sheet.getDataRange().getValues();
  const enc = filas[0];
  const idIdx = enc.indexOf("id");
  const emailIdx = enc.indexOf("email");
  if (emailIdx < 0) return "";
  for (let i = 1; i < filas.length; i++) {
    if (String(filas[i][idIdx]) === String(funerariaId)) return filas[i][emailIdx] || "";
  }
  return "";
}

// Registra en SuscripcionesFuneraria el intento de cobro (queda "pendiente"
// hasta que el webhook de Stripe confirme el pago). La columna servicioId se
// agrega dinámicamente porque la hoja original no la tenía (cobros de
// setup/mensualidad no están ligados a una página en particular).
function registrarSuscripcionPendiente_(funerariaId, tipo, monto, stripeSessionId, servicioId) {
  const sheet = getSheet_("SuscripcionesFuneraria");
  const servicioIdIdx = obtenerOCrearColumna_(sheet, "servicioId");
  sheet.appendRow([generarId_(), funerariaId, tipo, monto, stripeSessionId, "pendiente", new Date().toISOString(), ""]);
  if (servicioId) {
    sheet.getRange(sheet.getLastRow(), servicioIdIdx + 1).setValue(servicioId);
  }
}

// ============================================================
// COMISIONES 25% PARA FUNERARIA
// ============================================================

function registrarComisionInterna_(servicioId, creditos, tipoGesto) {
  try {
    // Obtener funerariaId del servicio
    const sheetS = getSheet_("Servicios");
    const filasS = sheetS.getDataRange().getValues();
    const encS = filasS[0];
    const idIdx = encS.indexOf("id");
    const funIdx = encS.indexOf("funerariaId");
    let funerariaId = "";
    for (let i = 1; i < filasS.length; i++) {
      if (filasS[i][idIdx] === servicioId) { funerariaId = filasS[i][funIdx]; break; }
    }
    if (!funerariaId) return;

    const montoTotal = parseFloat(creditos) || 0;
    const comision = Math.round(montoTotal * PRECIOS.comisionFuneraria * 100) / 100;

    const sheet = getSheet_("ComisionesFuneraria");
    sheet.appendRow([
      generarId_(), funerariaId, servicioId,
      tipoGesto || "gesto", montoTotal, comision,
      "pendiente", new Date().toISOString()
    ]);
  } catch(e) {}
}

function registrarComision_(datos) {
  if (!validarMaster_(datos.masterKey)) return { error: "No autorizado." };
  registrarComisionInterna_(datos.servicioId, datos.creditos, datos.concepto);
  return { ok: true };
}

function obtenerComisionesFuneraria_(funerariaId, masterKey) {
  if (!validarMaster_(masterKey)) return { error: "No autorizado." };
  const sheet = getSheet_("ComisionesFuneraria");
  const filas = sheet.getDataRange().getValues();
  const enc = filas[0];
  const funIdx = enc.indexOf("funerariaId");
  const resultado = [];
  let totalPendiente = 0;
  for (let i = 1; i < filas.length; i++) {
    if (String(filas[i][funIdx]) === String(funerariaId)) {
      const obj = {};
      enc.forEach((h, idx) => obj[h] = limpiarValor_(filas[i][idx]));
      resultado.push(obj);
      if (obj.estado === "pendiente") totalPendiente += parseFloat(obj.comision25 || 0);
    }
  }
  return { ok: true, comisiones: resultado, totalPendiente: Math.round(totalPendiente * 100) / 100 };
}

function obtenerSuscripcionesFuneraria_(masterKey) {
  if (!validarMaster_(masterKey)) return { error: "No autorizado." };
  const sheet = getSheet_("SuscripcionesFuneraria");
  const filas = sheet.getDataRange().getValues();
  const enc = filas[0];
  const resultado = [];
  for (let i = 1; i < filas.length; i++) {
    const obj = {};
    enc.forEach((h, idx) => obj[h] = limpiarValor_(filas[i][idx]));
    resultado.push(obj);
  }
  resultado.sort((a,b) => new Date(b.creadoEn) - new Date(a.creadoEn));
  return { ok: true, suscripciones: resultado };
}

// Inicializar catálogo y paquetes si están vacíos
function inicializarSistema() {
  inicializarCatalogoPredeterminado_();
  inicializarPaquetesPredeterminados_();
  return { ok: true, mensaje: "Sistema inicializado correctamente." };
}

// ============================================================
// VIDEOLLAMADA - FUNCIONES COMPLETAS
// ============================================================

function verificarSolicitudVL_(servicioId, email) {
  if (!servicioId || !email) return { ok: false };
  const sheet = getSheet_("SolicitudesVL");
  const filas = sheet.getDataRange().getValues();
  const enc = filas[0];
  const svcIdx = enc.indexOf("servicioId");
  const emailIdx = enc.indexOf("email");
  const estadoIdx = enc.indexOf("estado");
  for (let i = filas.length - 1; i >= 1; i--) {
    if (String(filas[i][svcIdx]) === String(servicioId) &&
        String(filas[i][emailIdx]).toLowerCase() === String(email).toLowerCase()) {
      return { ok: true, estado: filas[i][estadoIdx] || "pendiente" };
    }
  }
  return { ok: false };
}

function obtenerSolicitudesVLMaster_(masterKey) {
  if (!validarMaster_(masterKey)) return { error: "No autorizado." };
  const sheet = getSheet_("SolicitudesVL");
  const filas = sheet.getDataRange().getValues();
  const enc = filas[0];
  const resultado = [];
  for (let i = 1; i < filas.length; i++) {
    const obj = {};
    enc.forEach((h, idx) => obj[h] = limpiarValor_(filas[i][idx]));
    resultado.push(obj);
  }
  resultado.sort((a, b) => new Date(b.creadoEn) - new Date(a.creadoEn));
  return { ok: true, solicitudes: resultado };
}

// El master ingresa el link de Meet y el sistema lo envía a todos los que pagaron
function distribuirLinkVL_(datos) {
  if (!validarMaster_(datos.masterKey)) return { error: "No autorizado." };
  if (!datos.servicioId || !datos.enlace) return { error: "Faltan datos." };

  // Guardar el link en VideoLlamadas
  const sheetVL = getSheet_("VideoLlamadas");
  const enc = sheetVL.getRange(1, 1, 1, sheetVL.getLastColumn()).getValues()[0];
  const idExisteIdx = enc.indexOf("id");
  const svcIdx = enc.indexOf("servicioId");
  const filas = sheetVL.getDataRange().getValues();

  let encontrado = false;
  for (let i = 1; i < filas.length; i++) {
    if (String(filas[i][svcIdx]) === String(datos.servicioId)) {
      const enlaceIdx = enc.indexOf("enlace");
      if (enlaceIdx >= 0) sheetVL.getRange(i + 1, enlaceIdx + 1).setValue(datos.enlace);
      encontrado = true;
      break;
    }
  }

  if (!encontrado) {
    sheetVL.appendRow([
      generarId_(), datos.servicioId,
      datos.titulo || "Videollamada conmemorativa",
      datos.enlace,
      datos.fecha || "", datos.hora || "",
      new Date().toISOString(), true, ""
    ]);
  }

  // Obtener nombre del finado
  const sheetS = getSheet_("Servicios");
  const filasS = sheetS.getDataRange().getValues();
  const encS = filasS[0];
  const idSIdx = encS.indexOf("id");
  const nomIdx = encS.indexOf("nombreFinado");
  let nombreFinado = "su ser querido";
  for (let i = 1; i < filasS.length; i++) {
    if (String(filasS[i][idSIdx]) === String(datos.servicioId)) {
      nombreFinado = filasS[i][nomIdx] || nombreFinado;
      break;
    }
  }

  // Enviar link a todos los que pagaron este servicio
  const sheetSVL = getSheet_("SolicitudesVL");
  const filasSVL = sheetSVL.getDataRange().getValues();
  const encSVL = filasSVL[0];
  const svcSVLIdx = encSVL.indexOf("servicioId");
  const emailSVLIdx = encSVL.indexOf("email");
  const nomSVLIdx = encSVL.indexOf("nombre");
  const estadoSVLIdx = encSVL.indexOf("estado");

  let enviados = 0;
  for (let i = 1; i < filasSVL.length; i++) {
    if (String(filasSVL[i][svcSVLIdx]) === String(datos.servicioId) &&
        filasSVL[i][estadoSVLIdx] === "completado") {
      const emailDest = String(filasSVL[i][emailSVLIdx]);
      const nombreDest = String(filasSVL[i][nomSVLIdx] || "");
      try {
        MailApp.sendEmail({
          to: emailDest,
          subject: "Tu link de videollamada - " + nombreFinado,
          body: "Hola " + nombreDest + ",\n\n" +
                "La videollamada conmemorativa de " + nombreFinado + " ya está lista.\n\n" +
                "Únete aquí:\n" + datos.enlace + "\n\n" +
                (datos.fecha ? "Fecha: " + datos.fecha + "\n" : "") +
                (datos.hora ? "Hora: " + datos.hora + " hrs\n\n" : "\n") +
                "Puedes compartir este link con familiares y amigos.\n" +
                "Todos podrán ver, escuchar y hablar en tiempo real.\n\n" +
                "Con cariño,\nPáginas Conmemorativas"
        });
        enviados++;
      } catch(e) {}
    }
  }

  // Notificar al admin
  const emailAdmin = getConfigMaestro_("email");
  const wha = getConfigMaestro_("whatsapp");
  const msg = "Link de videollamada distribuido\nServicio: " + nombreFinado + "\nEnviado a " + enviados + " persona(s)\nLink: " + datos.enlace;
  if (emailAdmin) {
    try { MailApp.sendEmail({ to: emailAdmin, subject: "VL distribuida - " + nombreFinado, body: msg }); } catch(e) {}
  }
  if (wha) {
    try { UrlFetchApp.fetch("https://api.whatsapp.com/send?phone=" + wha + "&text=" + encodeURIComponent(msg), { muteHttpExceptions: true }); } catch(e) {}
  }

  return { ok: true, enviados };
}

// ============================================================
// ELIMINAR FUNERARIA
// ============================================================

function eliminarFuneraria_(datos) {
  if (!validarMaster_(datos.masterKey)) return { error: "No autorizado." };
  if (!datos.id) return { error: "ID requerido." };

  const sheet = getSheet_("Funerarias");
  const filas = sheet.getDataRange().getValues();
  const enc = filas[0];
  const idIdx = enc.indexOf("id");

  for (let i = 1; i < filas.length; i++) {
    if (String(filas[i][idIdx]) === String(datos.id)) {
      sheet.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { error: "Funeraria no encontrada." };
}

