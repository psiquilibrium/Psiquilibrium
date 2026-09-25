// ============================================================
//  PSIQUILIBRIUM — Backend v3 (franjas de 30 min)
// ============================================================

const SHEET_RESERVAS = "Reservas";
const SHEET_USUARIOS = "Usuarios";
const SHEET_BLOQUEOS = "Bloqueos";
const SHEET_AUDITORIA = "Auditoria";
const SHEET_OPERACIONES = "Operaciones";
const NOMBRES_CONSULTORIOS = ["Consultorio 1","Consultorio 2","Consultorio 3","Consultorio 4"];
const ROLES_USUARIO = ["admin", "socio", "asistente", "profesional"];
const USER_HEADERS = ["id", "nombre", "rol", "contraseña", "correo", "activo"];
const CACHE_AGENDA_PREFIX = "agenda_v1_";
const OPERATION_RETENTION_DAYS = 180;
const PERFORMANCE_METRICS_ENABLED = true; // Temporal: desactivar tras cerrar el diagnóstico.

// Franjas: 0=8:00, 1=8:30, 2=9:00 ... 19=17:30, 20=18:00 (no incluida)
// Sábado hasta las 12:00 = franja 8

function doGet(e) { return handle(e); }
function doPost(e) { return handle(e); }
function doOptions(e) { return ContentService.createTextOutput("").setMimeType(ContentService.MimeType.TEXT); }

function handle(e) {
  const startedAt = Date.now();
  let action = "desconocida";
  let requestOperationId = "";
  try {
    const params = e.parameter || {};
    const body = (e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : {};
    const merged = { ...params, ...body };
    action = merged.action || "desconocida";
    requestOperationId = String(merged.operationId || "").trim();
    const token = merged.token;

    if (action === "login") return respMedida(login(merged), action, startedAt);
    if (!validarToken(token)) return respMedida({ ok: false, error: "No autorizado" }, action, startedAt);

    switch (action) {
      case "getAgenda":              return respMedida(getAgenda(merged), action, startedAt);
      case "getAgendaVersion":       return respMedida(getAgendaVersion(merged), action, startedAt);
      case "getReservas":            return respMedida(getReservas(token), action, startedAt);
      case "getReporteReservas":     return respMedida(getReporteReservas(merged, token), action, startedAt);
      case "crearReserva":           return resp(crearReserva(merged, token));
      case "editarReserva":          return resp(editarReserva(merged, token));
      case "eliminarReserva":        return resp(eliminarReserva(merged, token));
      case "cambiarEstado":          return resp(cambiarEstado(merged, token));
      case "moverReserva":           return resp(moverReserva(merged, token));
      case "copiarReserva":          return resp(copiarReserva(merged, token));
      case "crearBloqueo":           return resp(crearBloqueo(merged, token));
      case "moverBloqueo":           return resp(moverBloqueo(merged, token));
      case "eliminarBloqueo":        return resp(eliminarBloqueo(merged, token));
      case "getBloqueos":            return respMedida(getBloqueos(), action, startedAt);
      case "getAuditoria":           return respMedida(getAuditoria(merged, token), action, startedAt);
      case "getUsuariosAdmin":       return respMedida(getUsuariosAdmin(token), action, startedAt);
      case "crearUsuario":           return resp(crearUsuario(merged, token));
      case "editarUsuario":          return resp(editarUsuario(merged, token));
      case "crearRespaldoManual":    return resp(crearRespaldoManual(token));
      case "generarPreestablecidas": return resp(generarPreestablecidas(merged, token));
      case "diagnosticarDatos":      return resp(diagnosticarDatos(token));
      case "migrarFranjas":          return resp(migrarFranjas(token));
      default: return resp({ ok: false, error: "Acción no reconocida" });
    }
  } catch (err) {
    return respMedida({ ok: false, error: err.message, retryable: !!requestOperationId }, action, startedAt);
  }
}

function respMedida(data, action, startedAt) {
  const durationMs = Date.now() - startedAt;
  if (PERFORMANCE_METRICS_ENABLED) {
    if (data && typeof data === "object") data.performanceMs = durationMs;
    console.log(`[Psiquilibrium] ${action}: ${durationMs} ms`);
  }
  return resp(data);
}

// ── Autenticación ────────────────────────────────────────────
function login(body) {
  const registro = leerRegistroUsuarios(SpreadsheetApp.getActiveSpreadsheet());
  const requestedId = String(body.userId || "").trim();
  const requestedKey = normalizarUserId(requestedId);
  const coincidencias = registro.porId[requestedKey] || [];
  if (coincidencias.length > 1) {
    return { ok: false, error: `El ID "${requestedId}" está repetido en Usuarios. Solicita al administrador corregirlo antes de ingresar.` };
  }
  const matched = coincidencias.length === 1 && coincidencias[0].id === requestedId &&
    String(coincidencias[0].password).trim() === String(body.password || "").trim() ? coincidencias[0] : null;
  if (matched && !matched.activo) return { ok: false, error: "Esta cuenta está desactivada. Consulta con administración." };
  const matchedUser = matched ? usuarioPublico(matched) : null;
  if (matchedUser) {
    const users = registro.usuarios
      .filter(u => u.activo && (registro.porId[u.idKey] || []).length === 1)
      .map(usuarioPublico);
    const token = Utilities.base64Encode(`${matchedUser.id}:${matchedUser.rol}:${new Date().toDateString()}`);
    return { ok: true, user: matchedUser, token, users };
  }
  return { ok: false, error: "Credenciales incorrectas" };
}

function validarToken(token) {
  if (!token) return false;
  try {
    const decoded = Utilities.newBlob(Utilities.base64Decode(token)).getDataAsString();
    return decoded.split(":")[2] === new Date().toDateString();
  } catch { return false; }
}

function getUserFromToken(token) {
  const decoded = Utilities.newBlob(Utilities.base64Decode(token)).getDataAsString();
  const parts = decoded.split(":");
  return { id: String(parts[0] || "").trim(), rol: String(parts[1] || "").trim().toLowerCase() };
}

function esRolOperativo(user) {
  return user.rol === "admin" || user.rol === "asistente";
}

function puedeGestionarBloqueos(user) {
  return user.rol === "admin" || user.rol === "asistente";
}

function puedeVerAuditoria(user) {
  return user.rol === "admin" || user.rol === "socio" || user.rol === "asistente";
}

function puedeCrearRespaldo(user) {
  return user.rol === "admin" || user.rol === "socio";
}

function puedeVerReportes(user) {
  return user.rol === "admin" || user.rol === "socio";
}

function puedeEliminarReservaCancelada(user) {
  return user.rol === "admin" || user.rol === "asistente" || user.rol === "profesional";
}

function getUserNameById(userId) {
  const registro = leerRegistroUsuarios(SpreadsheetApp.getActiveSpreadsheet());
  const matches = registro.porId[normalizarUserId(userId)] || [];
  if (matches.length === 1) return matches[0].nombre || matches[0].id;
  return String(userId || "");
}

function normalizarUserId(value) {
  return String(value || "").trim().toLowerCase();
}

function normalizarCorreo(value) {
  return String(value || "").trim().toLowerCase();
}

function normalizarNombre(value) {
  return String(value || "").trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, " ");
}

function usuarioEstaActivo(value) {
  if (value === "" || value === null || typeof value === "undefined") return true;
  return value !== false && !["false", "inactivo", "no", "0"].includes(String(value).trim().toLowerCase());
}

function usuarioPublico(usuario) {
  return { id: usuario.id, nombre: usuario.nombre, rol: usuario.rol };
}

function leerRegistroUsuarios(ss) {
  const sheet = ss.getSheetByName(SHEET_USUARIOS);
  const result = { sheet, usuarios: [], porId: {}, porCorreo: {}, porNombre: {} };
  if (!sheet) return result;
  const data = sheet.getDataRange().getValues();
  const headers = (data[0] || []).map(value => String(value || "").trim().toLowerCase());
  const tieneCorreo = ["correo", "email"].includes(headers[4]);
  const tieneActivo = ["activo", "activa"].includes(headers[5]);
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row.some(value => String(value || "").trim())) continue;
    const usuario = {
      fila: i + 1,
      id: String(row[0] || "").trim(),
      nombre: String(row[1] || "").trim(),
      rol: String(row[2] || "").trim().toLowerCase(),
      password: String(row[3] || ""),
      correo: tieneCorreo ? String(row[4] || "").trim() : "",
      activo: tieneActivo ? usuarioEstaActivo(row[5]) : true
    };
    usuario.idKey = normalizarUserId(usuario.id);
    usuario.nombreKey = normalizarNombre(usuario.nombre);
    usuario.correoKey = normalizarCorreo(usuario.correo);
    result.usuarios.push(usuario);
    if (usuario.idKey) {
      if (!result.porId[usuario.idKey]) result.porId[usuario.idKey] = [];
      result.porId[usuario.idKey].push(usuario);
    }
    if (usuario.nombreKey) {
      if (!result.porNombre[usuario.nombreKey]) result.porNombre[usuario.nombreKey] = [];
      result.porNombre[usuario.nombreKey].push(usuario);
    }
    if (usuario.correoKey) {
      if (!result.porCorreo[usuario.correoKey]) result.porCorreo[usuario.correoKey] = [];
      result.porCorreo[usuario.correoKey].push(usuario);
    }
  }
  return result;
}

function resolverUsuarioUnico(ss, userId, opciones) {
  const opts = opciones || {};
  const registro = leerRegistroUsuarios(ss);
  const matches = registro.porId[normalizarUserId(userId)] || [];
  if (!String(userId || "").trim()) return { ok: false, error: "Selecciona un profesional" };
  if (matches.length === 0) return { ok: false, error: `El profesional con ID "${userId}" no existe. Actualiza la agenda y vuelve a intentarlo.` };
  if (matches.length > 1) return { ok: false, error: `El ID "${userId}" está repetido. Corrígelo en Usuarios antes de continuar.` };
  const usuario = matches[0];
  if (opts.activo !== false && !usuario.activo) return { ok: false, error: `${usuario.nombre || usuario.id} está desactivado.` };
  if (opts.rol && usuario.rol !== opts.rol) return { ok: false, error: `${usuario.nombre || usuario.id} no tiene rol profesional.` };
  return { ok: true, usuario, registro };
}

function ensureUsuariosExtendedSchema(sheet) {
  const header = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 6)).getValues()[0];
  const aliases = [
    ["id", "userid", "usuario"],
    ["nombre"],
    ["rol"],
    ["contraseña", "password", "clave"],
    ["correo", "email"],
    ["activo", "activa"]
  ];
  for (let i = 0; i < USER_HEADERS.length; i++) {
    const actual = String(header[i] || "").trim().toLowerCase();
    if (actual && !aliases[i].includes(actual)) throw new Error(`La columna ${i + 1} de Usuarios contiene "${header[i]}". No se modificó la hoja.`);
    if (!actual) {
      const filas = Math.max(sheet.getLastRow() - 1, 0);
      const contieneDatos = filas > 0 && sheet.getRange(2, i + 1, filas, 1).getValues().some(row => String(row[0] || "").trim());
      if (contieneDatos) throw new Error(`La columna ${i + 1} de Usuarios tiene datos sin encabezado. No se modificó la hoja.`);
      sheet.getRange(1, i + 1).setValue(USER_HEADERS[i]);
    }
  }
}

function usuarioAudit(usuario) {
  return usuario ? { id: usuario.id, nombre: usuario.nombre, rol: usuario.rol, activo: usuario.activo } : null;
}

function getUsuariosAdmin(token) {
  const actor = getUserFromToken(token);
  if (actor.rol !== "admin") return { ok: false, error: "Solo admin" };
  const registro = leerRegistroUsuarios(SpreadsheetApp.getActiveSpreadsheet());
  return {
    ok: true,
    usuarios: registro.usuarios.map(u => ({
      ...usuarioAudit(u),
      correo: u.correo || "",
      fila: u.fila,
      idDuplicado: (registro.porId[u.idKey] || []).length > 1,
      nombreRepetido: (registro.porNombre[u.nombreKey] || []).length > 1
    }))
  };
}

function validarDatosUsuario(body, creando) {
  const id = String(body.id || "").trim();
  const nombre = String(body.nombre || "").trim();
  const rol = String(body.rol || "").trim().toLowerCase();
  const correo = String(body.correo || "").trim();
  const password = String(body.password || "");
  if (!id) return { ok: false, error: "El ID es obligatorio." };
  if (!/^[A-Za-z0-9._-]{3,60}$/.test(id)) return { ok: false, error: "El ID debe tener entre 3 y 60 caracteres y usar solo letras, números, punto, guion o guion bajo." };
  if (!nombre) return { ok: false, error: "El nombre es obligatorio." };
  if (!ROLES_USUARIO.includes(rol)) return { ok: false, error: "El rol no es válido." };
  if (correo && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(correo)) return { ok: false, error: "El correo no es válido." };
  if (creando && !password.trim()) return { ok: false, error: "La contraseña es obligatoria para un usuario nuevo." };
  return { ok: true, datos: { id, nombre, rol, correo, password, activo: usuarioEstaActivo(body.activo) } };
}

function crearUsuario(body, token) {
  return withWriteLock(function() {
    const actor = getUserFromToken(token);
    if (actor.rol !== "admin") return { ok: false, error: "Solo admin" };
    const validacion = validarDatosUsuario(body, true);
    if (!validacion.ok) return validacion;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const registro = leerRegistroUsuarios(ss);
    if (!registro.sheet) return { ok: false, error: "No existe hoja de usuarios" };
    const datos = validacion.datos;
    const existentes = registro.porId[normalizarUserId(datos.id)] || [];
    if (existentes.length === 1) {
      const existente = existentes[0];
      const mismoUsuario = existente.id === datos.id && existente.nombre === datos.nombre && existente.rol === datos.rol &&
        existente.correo === datos.correo && existente.activo === datos.activo && existente.password === datos.password;
      if (mismoUsuario) return { ok: true, idempotent: true, usuario: usuarioAudit(existente) };
    }
    if (existentes.length) return { ok: false, error: `El ID "${datos.id}" ya existe. Usa un ID diferente.` };
    if (datos.correo && (registro.porCorreo[normalizarCorreo(datos.correo)] || []).length) return { ok: false, error: `El correo "${datos.correo}" ya está registrado.` };
    ensureUsuariosExtendedSchema(registro.sheet);
    registro.sheet.appendRow([datos.id, datos.nombre, datos.rol, datos.password, datos.correo, datos.activo]);
    const despues = usuarioAudit(datos);
    registrarAuditoria(actor, "crear", "usuario", datos.id, `creó usuario ${datos.nombre} · ID ${datos.id} · rol ${datos.rol}`, null, despues);
    return { ok: true, usuario: despues, avisoNombre: (registro.porNombre[normalizarNombre(datos.nombre)] || []).length ? "Ya existe otro usuario con el mismo nombre. Sus IDs los mantienen separados." : "" };
  });
}

function editarUsuario(body, token) {
  return withWriteLock(function() {
    const actor = getUserFromToken(token);
    if (actor.rol !== "admin") return { ok: false, error: "Solo admin" };
    const validacion = validarDatosUsuario(body, false);
    if (!validacion.ok) return validacion;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const registro = leerRegistroUsuarios(ss);
    const matches = registro.porId[normalizarUserId(validacion.datos.id)] || [];
    if (matches.length !== 1) return { ok: false, error: matches.length ? "El ID está repetido; corrígelo manualmente antes de editar." : "Usuario no encontrado." };
    const actual = matches[0];
    const datos = validacion.datos;
    datos.id = actual.id;
    if (typeof body.activo === "undefined" || body.activo === "") datos.activo = actual.activo;
    if (actual.id === actor.id && (datos.rol !== actual.rol || datos.activo !== actual.activo)) return { ok: false, error: "No puedes cambiar tu propio rol ni desactivar tu cuenta desde la app." };
    const correoMatches = datos.correo ? (registro.porCorreo[normalizarCorreo(datos.correo)] || []).filter(u => u.fila !== actual.fila) : [];
    if (correoMatches.length) return { ok: false, error: `El correo "${datos.correo}" ya está registrado.` };
    const despues = { ...actual, nombre: datos.nombre, rol: datos.rol, correo: datos.correo, activo: datos.activo };
    const passwordCambia = datos.password.trim() && actual.password !== datos.password;
    const sinCambios = actual.nombre === despues.nombre && actual.rol === despues.rol && actual.correo === despues.correo && actual.activo === despues.activo && !passwordCambia;
    if (sinCambios) return { ok: true, idempotent: true, usuario: usuarioAudit(actual) };
    ensureUsuariosExtendedSchema(registro.sheet);
    registro.sheet.getRange(actual.fila, 2, 1, 2).setValues([[datos.nombre, datos.rol]]);
    if (passwordCambia) registro.sheet.getRange(actual.fila, 4).setValue(datos.password);
    registro.sheet.getRange(actual.fila, 5, 1, 2).setValues([[datos.correo, datos.activo]]);
    const cambios = [];
    if (actual.nombre !== despues.nombre) cambios.push(`nombre: ${actual.nombre} → ${despues.nombre}`);
    if (actual.rol !== despues.rol) cambios.push(`rol: ${actual.rol} → ${despues.rol}`);
    if (actual.activo !== despues.activo) cambios.push(despues.activo ? "activó la cuenta" : "desactivó la cuenta");
    if (actual.correo !== despues.correo) cambios.push("actualizó el correo");
    if (passwordCambia) cambios.push("actualizó la contraseña");
    const accion = actual.activo !== despues.activo ? (despues.activo ? "activar" : "desactivar") : "editar";
    registrarAuditoria(actor, accion, "usuario", actual.id, `actualizó usuario ${despues.nombre} · ID ${actual.id} · ${cambios.join(" · ")}`, usuarioAudit(actual), usuarioAudit(despues));
    return { ok: true, usuario: usuarioAudit(despues), avisoNombre: (registro.porNombre[normalizarNombre(datos.nombre)] || []).filter(u => u.fila !== actual.fila).length ? "Ya existe otro usuario con el mismo nombre. Sus IDs los mantienen separados." : "" };
  });
}

// ── Migración de franjas horarias (ejecutar UNA sola vez) ────
function migrarFranjas(token) {
  const user = getUserFromToken(token);
  if (user.rol !== "admin") return { ok: false, error: "Solo admin" };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let migradas = 0;

  // Migrar Reservas
  const sheetR = ss.getSheetByName(SHEET_RESERVAS);
  const dataR = sheetR.getDataRange().getValues();
  for (let i = 1; i < dataR.length; i++) {
    const franja = Number(dataR[i][4]);
    const duracion = Number(dataR[i][5]);
    // Si franja <= 9 y duracion es múltiplo de 60, es formato anterior (horas enteras)
    if (franja <= 9 && duracion % 60 === 0) {
      sheetR.getRange(i + 1, 5).setValue(franja * 2);
      sheetR.getRange(i + 1, 6).setValue(duracion / 60 * 2);  // convertir horas a franjas de 30min
      // Pero guardamos en minutos para consistencia
      sheetR.getRange(i + 1, 6).setValue(duracion); // duracion ya está en minutos, no cambia
      sheetR.getRange(i + 1, 5).setValue(franja * 2); // solo franja ×2
      migradas++;
    }
  }

  // Migrar Bloqueos
  const sheetB = ss.getSheetByName(SHEET_BLOQUEOS);
  if (sheetB) {
    const dataB = sheetB.getDataRange().getValues();
    for (let i = 1; i < dataB.length; i++) {
      const franja = Number(dataB[i][2]);
      const duracion = Number(dataB[i][4]);
      if (franja <= 9 && duracion % 60 === 0) {
        sheetB.getRange(i + 1, 3).setValue(franja * 2);
        migradas++;
      }
    }
  }

  return { ok: true, migradas, mensaje: `Migración completa. ${migradas} filas actualizadas.` };
}

// ── Reservas ─────────────────────────────────────────────────
function getReservas(token) {
  const user = getUserFromToken(token);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RESERVAS);
  const data = sheet.getDataRange().getValues();
  const reservas = [];

  for (let i = 1; i < data.length; i++) {
    const reserva = reservaFromRow(data[i]);
    if (reserva && !puedeVerListadoCompletoReservas(user) && reserva.userId !== user.id) continue;
    if (reserva) reservas.push(reserva);
  }
  return { ok: true, reservas, version: getAgendaCacheVersion() };
}

function puedeVerListadoCompletoReservas(user) {
  return user.rol === "admin" || user.rol === "socio" || user.rol === "asistente";
}

function getReporteReservas(body, token) {
  const user = getUserFromToken(token);
  if (!puedeVerReportes(user)) return { ok: false, error: "Sin permiso" };

  const desde = String(body.desde || "").slice(0, 10);
  const hasta = String(body.hasta || "").slice(0, 10);
  const profesional = String(body.profesional || "todos");
  if (!esFechaValida(desde) || !esFechaValida(hasta) || desde > hasta) return { ok: false, error: "Rango inválido" };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_RESERVAS);
  if (!sheet) return { ok: false, error: "No existe hoja de reservas" };

  const nombres = getUserNamesMap(ss);
  const data = sheet.getDataRange().getValues();
  const resumen = { confirmadas: 0, canceladas: 0, total: 0 };
  const porProfesional = {};
  const porConsultorio = {};

  for (let i = 1; i < data.length; i++) {
    const reserva = reservaFromRow(data[i]);
    if (!reserva) continue;
    if (reserva.fecha < desde || reserva.fecha > hasta) continue;
    if (profesional !== "todos" && reserva.userId !== profesional) continue;

    const estado = reserva.estado === "cancelada" ? "canceladas" : "confirmadas";
    resumen[estado]++;
    resumen.total++;

    if (!porProfesional[reserva.userId]) {
      porProfesional[reserva.userId] = { id: reserva.userId, nombre: nombres[reserva.userId] || reserva.userId, confirmadas: 0, canceladas: 0, total: 0 };
    }
    porProfesional[reserva.userId][estado]++;
    porProfesional[reserva.userId].total++;

    const consultorioNombre = NOMBRES_CONSULTORIOS[Number(reserva.consultorio)] || String(reserva.consultorio);
    if (!porConsultorio[consultorioNombre]) {
      porConsultorio[consultorioNombre] = { consultorio: consultorioNombre, confirmadas: 0, canceladas: 0, total: 0 };
    }
    porConsultorio[consultorioNombre][estado]++;
    porConsultorio[consultorioNombre].total++;
  }

  return {
    ok: true,
    desde,
    hasta,
    profesional,
    resumen,
    porProfesional: Object.values(porProfesional).sort((a, b) => b.total - a.total || a.nombre.localeCompare(b.nombre)),
    porConsultorio: Object.values(porConsultorio).sort((a, b) => a.consultorio.localeCompare(b.consultorio))
  };
}

function getUserNamesMap(ss) {
  const map = {};
  const registro = leerRegistroUsuarios(ss);
  registro.usuarios.forEach(usuario => {
    if ((registro.porId[usuario.idKey] || []).length === 1) map[usuario.id] = usuario.nombre || usuario.id;
  });
  return map;
}

function getAgenda(body) {
  const desde = String(body.desde || "").slice(0, 10);
  const hasta = String(body.hasta || "").slice(0, 10);
  if (!desde || !hasta) return { ok: false, error: "Rango inválido" };

  const cacheKey = CACHE_AGENDA_PREFIX + getAgendaCacheVersion() + "_" + desde + "_" + hasta;
  const cached = getCacheValue(cacheKey);
  if (cached) return JSON.parse(cached);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetR = ss.getSheetByName(SHEET_RESERVAS);
  const dataR = sheetR.getDataRange().getValues();
  const reservas = [];
  for (let i = 1; i < dataR.length; i++) {
    const reserva = reservaFromRow(dataR[i]);
    if (reserva && reserva.fecha >= desde && reserva.fecha <= hasta) reservas.push(reserva);
  }

  const sheetB = ensureBloqueosSheet(ss);
  const dataB = sheetB.getDataRange().getValues();
  const bloqueos = [];
  for (let i = 1; i < dataB.length; i++) {
    const bloqueo = bloqueoFromRow(dataB[i]);
    if (!bloqueo) continue;
    if (bloqueo.fecha >= desde && bloqueo.fecha <= hasta) {
      bloqueos.push(bloqueo);
      continue;
    }
    if (bloqueo.repeticion === "semanal" && bloqueoAplicaEnRango(bloqueo.fecha, desde, hasta)) {
      bloqueos.push(bloqueo);
    }
  }

  const result = { ok: true, reservas, bloqueos, desde, hasta, version: getAgendaCacheVersion() };
  putCacheValue(cacheKey, JSON.stringify(result), 20);
  return result;
}

function getAgendaVersion(body) {
  const desde = String(body.desde || "").slice(0, 10);
  const hasta = String(body.hasta || "").slice(0, 10);
  if (!desde || !hasta) return { ok: false, error: "Rango inválido" };
  return { ok: true, desde, hasta, version: getAgendaCacheVersion() };
}

function reservaFromRow(row) {
  const [id, consultorio, userId, fecha, franja, duracion, nota, activa, tipo, estado] = row;
  if (!id) return null;
  if (activa === false || String(activa).toUpperCase() === "FALSE") return null;

  const fechaStr = fechaToString(fecha);
  let consultorioIdx = Number(consultorio);
  if (isNaN(consultorioIdx)) {
    consultorioIdx = NOMBRES_CONSULTORIOS.indexOf(String(consultorio));
    if (consultorioIdx === -1) consultorioIdx = 0;
  }

  return {
    id: String(id), consultorio: consultorioIdx, userId: String(userId),
    fecha: fechaStr, franja: Number(franja), duracion: Number(duracion),
    nota: String(nota || ""), tipo: String(tipo || "normal"),
    estado: String(estado || "confirmada").trim().toLowerCase(),
    version: rowVersion("reserva", row)
  };
}

// ── Idempotencia de operaciones de creación ─────────────────
function normalizarOperationId(value) {
  const operationId = String(value || "").trim();
  if (!operationId) return "";
  if (operationId.length < 16 || operationId.length > 120 || !/^[A-Za-z0-9._:-]+$/.test(operationId)) {
    throw new Error("Identificador de operación inválido. Actualiza la app e intenta de nuevo.");
  }
  return operationId;
}

function hashOperacion(value) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(value),
    Utilities.Charset.UTF_8
  );
  return bytes.map(byte => (byte < 0 ? byte + 256 : byte).toString(16).padStart(2, "0")).join("");
}

function ensureOperacionesSheet(ss) {
  const headers = ["operationId", "creadoEn", "userId", "accion", "tipo", "elementoId", "payloadHash", "estado", "actualizadoEn"];
  let sheet = ss.getSheetByName(SHEET_OPERACIONES);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_OPERACIONES);
    sheet.appendRow(headers);
  } else {
    const actual = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
    for (let i = 0; i < headers.length; i++) {
      if (String(actual[i] || "").trim().toLowerCase() !== headers[i].toLowerCase()) {
        throw new Error("La hoja Operaciones ya existe con una estructura diferente. No se modificaron datos.");
      }
    }
  }
  return sheet;
}

function buscarFilaPorId(sheet, id) {
  if (!sheet || !id) return null;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) return { fila: i + 1, row: data[i] };
  }
  return null;
}

function prepararOperacionCreacion(ss, body, user, accion, tipo, payload, prefijo) {
  const operationId = normalizarOperationId(body.operationId);
  if (!operationId) {
    return { ok: false, error: "Esta versión de la app necesita actualizarse antes de crear o copiar. Recarga la aplicación e intenta de nuevo." };
  }

  const payloadHash = hashOperacion(JSON.stringify(payload));
  const elementoId = `${prefijo}_OP_${hashOperacion(operationId).slice(0, 20)}`;
  const sheet = ensureOperacionesSheet(ss);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) !== operationId) continue;
    const mismoContexto = String(data[i][2]) === String(user.id) &&
      String(data[i][3]) === accion && String(data[i][4]) === tipo &&
      String(data[i][5]) === elementoId && String(data[i][6]) === payloadHash;
    if (!mismoContexto) {
      return { ok: false, error: "Ese identificador de operación ya fue usado para una acción diferente. Inicia la acción nuevamente." };
    }
    return { ok: true, legacy: false, operationId, elementoId, payloadHash, existente: true, fila: i + 1, estado: String(data[i][7] || "") };
  }
  return { ok: true, legacy: false, operationId, elementoId, payloadHash, existente: false, sheet };
}

function iniciarOperacionCreacion(ss, operacion, user, accion, tipo) {
  if (operacion.legacy || operacion.existente) return operacion;
  ensureAuditoriaOperationColumn(ensureAuditoriaSheet(ss));
  const sheet = operacion.sheet || ensureOperacionesSheet(ss);
  const now = new Date();
  sheet.appendRow([operacion.operationId, now, user.id, accion, tipo, operacion.elementoId, operacion.payloadHash, "pendiente", now]);
  operacion.existente = true;
  operacion.fila = sheet.getLastRow();
  return operacion;
}

function completarOperacionCreacion(ss, operacion) {
  if (operacion.legacy) return;
  try {
    const sheet = ensureOperacionesSheet(ss);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) !== operacion.operationId) continue;
      sheet.getRange(i + 1, 8, 1, 2).setValues([["completada", new Date()]]);
      limpiarOperacionesAntiguas(ss);
      return;
    }
  } catch (err) {}
}

function limpiarOperacionesAntiguas(ss) {
  try {
    const props = PropertiesService.getScriptProperties();
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    if (props.getProperty("ultimaLimpiezaOperaciones") === today) return;
    props.setProperty("ultimaLimpiezaOperaciones", today);
    const sheet = ss.getSheetByName(SHEET_OPERACIONES);
    if (!sheet || sheet.getLastRow() < 2) return;
    const cutoff = Date.now() - OPERATION_RETENTION_DAYS * 24 * 60 * 60 * 1000;
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
    for (let i = data.length - 1; i >= 0; i--) {
      const createdAt = data[i][1] instanceof Date ? data[i][1].getTime() : new Date(data[i][1]).getTime();
      if (String(data[i][7]) === "completada" && Number.isFinite(createdAt) && createdAt < cutoff) sheet.deleteRow(i + 2);
    }
  } catch (err) {}
}

// ── Auditoría ────────────────────────────────────────────────
function ensureAuditoriaSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_AUDITORIA);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_AUDITORIA);
    sheet.appendRow(["timestamp","userId","nombre","rol","accion","tipo","elementoId","resumen","antes","despues","operationId"]);
  }
  return sheet;
}

function ensureAuditoriaOperationColumn(sheet) {
  const header = String(sheet.getRange(1, 11).getValue() || "").trim();
  if (header && header.toLowerCase() !== "operationid") throw new Error("La columna 11 de Auditoria ya está ocupada.");
  if (!header) {
    const filas = Math.max(sheet.getLastRow() - 1, 0);
    const contieneDatos = filas > 0 && sheet.getRange(2, 11, filas, 1).getValues().some(row => String(row[0] || "").trim());
    if (contieneDatos) throw new Error("La columna 11 de Auditoria tiene datos sin encabezado. No se modificó la hoja.");
    sheet.getRange(1, 11).setValue("operationId");
  }
}

function auditoriaOperacionExiste(sheet, operationId) {
  if (!operationId || sheet.getLastRow() < 2) return false;
  ensureAuditoriaOperationColumn(sheet);
  const values = sheet.getRange(2, 11, sheet.getLastRow() - 1, 1).getValues();
  return values.some(row => String(row[0]) === operationId);
}

function registrarAuditoria(user, accion, tipo, elementoId, resumen, antes, despues, operationId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ensureAuditoriaSheet(ss);
    const opId = String(operationId || "");
    if (opId) ensureAuditoriaOperationColumn(sheet);
    const nombre = getUserNameById(user.id) || user.id;
    const row = [
      new Date(),
      user.id,
      nombre,
      user.rol,
      accion,
      tipo,
      elementoId,
      resumen,
      antes ? JSON.stringify(antes) : "",
      despues ? JSON.stringify(despues) : ""
    ];
    if (opId) row.push(opId);
    sheet.appendRow(row);
    return true;
  } catch (err) {
    return false;
  }
}

function registrarAuditoriaUnaVez(user, accion, tipo, elementoId, resumen, antes, despues, operationId) {
  if (!operationId) return registrarAuditoria(user, accion, tipo, elementoId, resumen, antes, despues);
  try {
    const sheet = ensureAuditoriaSheet(SpreadsheetApp.getActiveSpreadsheet());
    if (auditoriaOperacionExiste(sheet, operationId)) return true;
  } catch (err) {
    return false;
  }
  return registrarAuditoria(user, accion, tipo, elementoId, resumen, antes, despues, operationId);
}

function getAuditoria(body, token) {
  const user = getUserFromToken(token);
  if (!puedeVerAuditoria(user)) return { ok: false, error: "Sin permiso" };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ensureAuditoriaSheet(ss);

  const limit = Math.min(Math.max(Number(body.limit || 100), 1), 200);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: true, eventos: [] };
  const startRow = Math.max(2, lastRow - limit + 1);
  const data = sheet.getRange(startRow, 1, lastRow - startRow + 1, 10).getValues();
  const eventos = data.reverse().map(row => ({
    timestamp: row[0] instanceof Date ? Utilities.formatDate(row[0], Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss") : String(row[0] || ""),
    userId: String(row[1] || ""),
    nombre: String(row[2] || row[1] || ""),
    rol: String(row[3] || ""),
    accion: String(row[4] || ""),
    tipo: String(row[5] || ""),
    elementoId: String(row[6] || ""),
    resumen: String(row[7] || ""),
    antes: String(row[8] || ""),
    despues: String(row[9] || "")
  }));
  return { ok: true, eventos };
}

function reservaAuditFromRow(row) {
  if (!row) return null;
  const [id, consultorio, userId, fecha, franja, duracion, nota, activa, tipo, estado] = row;
  return {
    id: String(id || ""),
    consultorio: String(consultorio || ""),
    userId: String(userId || ""),
    profesional: getUserNameById(userId),
    fecha: fechaToString(fecha),
    franja: Number(franja),
    hora: franjaToHora(franja),
    duracion: Number(duracion),
    nota: String(nota || ""),
    activa: activa !== false && String(activa).toUpperCase() !== "FALSE",
    tipo: String(tipo || "normal"),
    estado: String(estado || "confirmada")
  };
}

function bloqueoAuditFromRow(row) {
  if (!row) return null;
  const [id, consultorio, franja, fecha, duracion, nota, activo, repeticion] = row;
  return {
    id: String(id || ""),
    consultorio: String(consultorio || ""),
    fecha: fechaToString(fecha),
    franja: Number(franja),
    hora: franjaToHora(franja),
    duracion: Number(duracion),
    nota: String(nota || ""),
    activo: activo !== false && String(activo).toUpperCase() !== "FALSE",
    repeticion: String(repeticion || "ninguna")
  };
}

function franjaToHora(franja) {
  const n = Number(franja);
  if (!Number.isFinite(n)) return String(franja || "");
  const h = 8 + Math.floor(n / 2);
  const m = n % 2 === 0 ? "00" : "30";
  return String(h).padStart(2, "0") + ":" + m;
}

function resumenReservaAudit(r) {
  if (!r) return "";
  return `${r.profesional || r.userId} · ${r.consultorio} · ${r.fecha} · ${r.hora}`;
}

function resumenEliminacionReservaAudit(r) {
  if (!r) return "";
  const paciente = r.nota ? ` · Paciente/nota: ${r.nota}` : "";
  return `eliminó reserva cancelada de ${r.profesional || r.userId}${paciente} · ${r.consultorio} · ${r.fecha} · ${r.hora} · estado previo: cancelada`;
}

function resumenBloqueoAudit(b) {
  if (!b) return "";
  return `${b.consultorio} · ${b.fecha} · ${b.hora}${b.nota ? " · " + b.nota : ""}`;
}

// ── Respaldos manuales ───────────────────────────────────────
function crearRespaldoManual(token) {
  return withWriteLock(function() {
    const user = getUserFromToken(token);
    if (!puedeCrearRespaldo(user)) return { ok: false, error: "Solo admin o socio" };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmm");
    const hojas = [SHEET_RESERVAS, SHEET_BLOQUEOS, SHEET_USUARIOS, SHEET_AUDITORIA, SHEET_OPERACIONES];
    const creadas = [];
    const omitidas = [];

    hojas.forEach(nombre => {
      const source = ss.getSheetByName(nombre);
      if (!source) {
        omitidas.push(nombre);
        return;
      }
      const backupName = uniqueSheetName(ss, `Backup_${nombre}_${timestamp}`);
      source.copyTo(ss).setName(backupName);
      creadas.push(backupName);
    });

    const resumen = `creó respaldo manual de ${creadas.length} hojas: ${creadas.join(", ")}`;
    registrarAuditoria(user, "respaldo", "sistema", timestamp, resumen, null, { hojas: creadas, omitidas });
    return { ok: true, timestamp, hojas: creadas, omitidas };
  });
}

function uniqueSheetName(ss, baseName) {
  let name = baseName.slice(0, 99);
  let n = 2;
  while (ss.getSheetByName(name)) {
    const suffix = `_${n}`;
    name = baseName.slice(0, 99 - suffix.length) + suffix;
    n++;
  }
  return name;
}

function crearReserva(body, token) {
  return withWriteLock(function() {
    const user = getUserFromToken(token);
    const { consultorio, fecha, franja, duracion, nota } = body;
    if (user.rol === "asistente" && !body.targetUserId) return { ok: false, error: "Selecciona profesional" };
    const requestedUserId = (esRolOperativo(user) && body.targetUserId) ? body.targetUserId : user.id;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const payload = {
      consultorio: String(consultorio), fecha: fechaToString(fecha), franja: Number(franja),
      duracion: Number(duracion), nota: String(nota || ""), targetUserId: String(requestedUserId), tipo: String(body.tipo || "normal")
    };
    const operacion = prepararOperacionCreacion(ss, body, user, "crearReserva", "reserva", payload, "R");
    if (!operacion.ok) return operacion;
    const sheet = ss.getSheetByName(SHEET_RESERVAS);
    const yaCreada = buscarFilaPorId(sheet, operacion.elementoId);
    if (yaCreada) {
      if (!esRolOperativo(user) && String(yaCreada.row[2]) !== user.id) return { ok: false, error: "Sin permiso" };
      const despuesExistente = reservaAuditFromRow(yaCreada.row);
      registrarAuditoriaUnaVez(user, "crear", "reserva", operacion.elementoId, `creó reserva de ${resumenReservaAudit(despuesExistente)}`, null, despuesExistente, operacion.operationId);
      completarOperacionCreacion(ss, operacion);
      invalidateAgendaCache();
      return { ok: true, id: operacion.elementoId, idempotent: true };
    }

    const identidad = resolverUsuarioUnico(ss, requestedUserId, { rol: "profesional" });
    if (!identidad.ok) return identidad;
    const userId = identidad.usuario.id;

    if (estaBloquado(consultorio, fecha, franja, duracion)) return { ok: false, error: "Franja bloqueada" };
    if (hayConflicto(consultorio, fecha, franja, duracion, null)) return { ok: false, error: "Conflicto de horario" };

    const consultorioNombre = NOMBRES_CONSULTORIOS[Number(consultorio)] || String(consultorio);
    iniciarOperacionCreacion(ss, operacion, user, "crearReserva", "reserva");
    const id = operacion.elementoId;
    const row = [id, consultorioNombre, userId, fecha, franja, duracion, nota || "", true, body.tipo || "normal", "confirmada"];
    sheet.appendRow(row);
    const despues = reservaAuditFromRow(row);
    registrarAuditoriaUnaVez(user, "crear", "reserva", id, `creó reserva de ${resumenReservaAudit(despues)}`, null, despues, operacion.operationId);
    completarOperacionCreacion(ss, operacion);
    invalidateAgendaCache();
    return { ok: true, id };
  });
}

function editarReserva(body, token) {
  return withWriteLock(function() {
    const user = getUserFromToken(token);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RESERVAS);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) !== String(body.id)) continue;
      if (!esRolOperativo(user) && user.id !== String(data[i][2])) return { ok: false, error: "Sin permiso" };
      if (!rowActiva(data[i][7])) return conflictoActualizado("reserva");
      if (Number(data[i][5]) === Number(body.duracion) && String(data[i][6] || "") === String(body.nota || "")) {
        return { ok: true, idempotent: true, version: rowVersion("reserva", data[i]) };
      }
      if (!versionEsperadaCoincide(body, "reserva", data[i])) return conflictoActualizado("reserva");
      if (hayConflicto(data[i][1], data[i][3], data[i][4], body.duracion, body.id)) return { ok: false, error: "Conflicto de horario" };
      const antes = reservaAuditFromRow(data[i]);
      sheet.getRange(i + 1, 6).setValue(body.duracion);
      sheet.getRange(i + 1, 7).setValue(body.nota || "");
      const despues = reservaAuditFromRow([data[i][0],data[i][1],data[i][2],data[i][3],data[i][4],body.duracion,body.nota || "",data[i][7],data[i][8],data[i][9]]);
      registrarAuditoria(user, "editar", "reserva", String(body.id), `editó reserva de ${resumenReservaAudit(despues)}`, antes, despues);
      invalidateAgendaCache();
      return { ok: true, version: rowVersion("reserva", [data[i][0],data[i][1],data[i][2],data[i][3],data[i][4],body.duracion,body.nota || "",data[i][7],data[i][8],data[i][9]]) };
    }
    return conflictoActualizado("reserva");
  });
}

function cambiarEstado(body, token) {
  return withWriteLock(function() {
    const user = getUserFromToken(token);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RESERVAS);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) !== String(body.id)) continue;
      if (!rowActiva(data[i][7])) return conflictoActualizado("reserva");
      const propietario = String(data[i][2]);
      // Admin puede cambiar cualquier estado. Asistente puede cancelar cualquiera. El dueño solo puede cancelar la suya.
      if (user.rol === "asistente") {
        if (body.estado !== "cancelada") return { ok: false, error: "Solo admin puede reconfirmar reservas" };
      } else if (user.rol !== "admin") {
        if (user.id !== propietario) return { ok: false, error: "Sin permiso" };
        if (body.estado !== "cancelada") return { ok: false, error: "Solo puedes cancelar tus propias reservas" };
      }
      const estadoActual = String(data[i][9] || "confirmada").trim().toLowerCase();
      const estadoSolicitado = String(body.estado || "").trim().toLowerCase();
      if (estadoActual === estadoSolicitado) {
        return { ok: true, idempotent: true, version: rowVersion("reserva", data[i]) };
      }
      if (!versionEsperadaCoincide(body, "reserva", data[i])) return conflictoActualizado("reserva");
      const antes = reservaAuditFromRow(data[i]);
      sheet.getRange(i + 1, 10).setValue(body.estado);
      const despues = reservaAuditFromRow([data[i][0],data[i][1],data[i][2],data[i][3],data[i][4],data[i][5],data[i][6],data[i][7],data[i][8],body.estado]);
      const accion = body.estado === "cancelada" ? "cancelar" : "reconfirmar";
      registrarAuditoria(user, accion, "reserva", String(body.id), `${accion === "cancelar" ? "canceló" : "reconfirmó"} reserva de ${resumenReservaAudit(despues)}`, antes, despues);
      invalidateAgendaCache();
      return { ok: true, version: rowVersion("reserva", [data[i][0],data[i][1],data[i][2],data[i][3],data[i][4],data[i][5],data[i][6],data[i][7],data[i][8],body.estado]) };
    }
    return conflictoActualizado("reserva");
  });
}

function moverReserva(body, token) {
  return withWriteLock(function() {
    const user = getUserFromToken(token);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RESERVAS);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) !== String(body.id)) continue;
      if (!esRolOperativo(user) && user.id !== String(data[i][2])) return { ok: false, error: "Sin permiso" };
      if (!rowActiva(data[i][7])) return conflictoActualizado("reserva");
      const dur = Number(data[i][5]);
      const consultorioNombre = NOMBRES_CONSULTORIOS[Number(body.consultorio)] || String(body.consultorio);
      const mismoDestino = normalizarConsultorioIndice(data[i][1]) === normalizarConsultorioIndice(consultorioNombre) &&
        fechaToString(data[i][3]) === fechaToString(body.fecha) && Number(data[i][4]) === Number(body.franja);
      if (mismoDestino) return { ok: true, idempotent: true, version: rowVersion("reserva", data[i]) };
      if (!versionEsperadaCoincide(body, "reserva", data[i])) return conflictoActualizado("reserva");
      if (estaBloquado(body.consultorio, body.fecha, body.franja, dur)) return { ok: false, error: "Franja bloqueada" };
      if (hayConflicto(body.consultorio, body.fecha, body.franja, dur, body.id)) return { ok: false, error: "Conflicto de horario" };
      const antes = reservaAuditFromRow(data[i]);
      sheet.getRange(i + 1, 2).setValue(consultorioNombre);
      sheet.getRange(i + 1, 4).setValue(body.fecha);
      sheet.getRange(i + 1, 5).setValue(body.franja);
      const despues = reservaAuditFromRow([data[i][0],consultorioNombre,data[i][2],body.fecha,body.franja,data[i][5],data[i][6],data[i][7],data[i][8],data[i][9]]);
      registrarAuditoria(user, "mover", "reserva", String(body.id), `movió reserva de ${resumenReservaAudit(antes)} → ${despues.consultorio} · ${despues.fecha} · ${despues.hora}`, antes, despues);
      invalidateAgendaCache();
      return { ok: true, version: rowVersion("reserva", [data[i][0],consultorioNombre,data[i][2],body.fecha,body.franja,data[i][5],data[i][6],data[i][7],data[i][8],data[i][9]]) };
    }
    return conflictoActualizado("reserva");
  });
}

function copiarReserva(body, token) {
  return withWriteLock(function() {
    const user = getUserFromToken(token);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_RESERVAS);
    const data = sheet.getDataRange().getValues();
    const payload = {
      sourceId: String(body.id || ""), consultorio: String(body.consultorio),
      fecha: fechaToString(body.fecha), franja: Number(body.franja)
    };
    const operacion = prepararOperacionCreacion(ss, body, user, "copiarReserva", "reserva", payload, "R");
    if (!operacion.ok) return operacion;
    const yaCreada = buscarFilaPorId(sheet, operacion.elementoId);
    if (yaCreada) {
      if (!esRolOperativo(user) && String(yaCreada.row[2]) !== user.id) return { ok: false, error: "Sin permiso" };
      const source = data.find(row => String(row[0]) === String(body.id));
      const despuesExistente = reservaAuditFromRow(yaCreada.row);
      registrarAuditoriaUnaVez(user, "crear", "reserva", operacion.elementoId, `copió reserva de ${resumenReservaAudit(despuesExistente)}`, source ? reservaAuditFromRow(source) : null, despuesExistente, operacion.operationId);
      completarOperacionCreacion(ss, operacion);
      invalidateAgendaCache();
      return { ok: true, id: operacion.elementoId, idempotent: true };
    }

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) !== String(body.id)) continue;
      const [, , userId, , , duracion, nota, , tipo] = data[i];
      if (!esRolOperativo(user) && user.id !== String(userId)) return { ok: false, error: "Sin permiso" };
      const identidad = resolverUsuarioUnico(ss, userId, { rol: "profesional" });
      if (!identidad.ok) return { ok: false, error: `No se puede copiar esta reserva: ${identidad.error}` };
      const targetUserId = identidad.usuario.id;
      if (estaBloquado(body.consultorio, body.fecha, body.franja, duracion)) return { ok: false, error: "Franja bloqueada" };
      if (hayConflicto(body.consultorio, body.fecha, body.franja, duracion, null)) return { ok: false, error: "Conflicto de horario" };
      const consultorioNombre = NOMBRES_CONSULTORIOS[Number(body.consultorio)] || String(body.consultorio);
      iniciarOperacionCreacion(ss, operacion, user, "copiarReserva", "reserva");
      const newId = operacion.elementoId;
      const row = [newId, consultorioNombre, targetUserId, body.fecha, body.franja, duracion, nota || "", true, tipo || "normal", "confirmada"];
      sheet.appendRow(row);
      const despues = reservaAuditFromRow(row);
      registrarAuditoriaUnaVez(user, "crear", "reserva", newId, `copió reserva de ${resumenReservaAudit(despues)}`, reservaAuditFromRow(data[i]), despues, operacion.operationId);
      completarOperacionCreacion(ss, operacion);
      invalidateAgendaCache();
      return { ok: true, id: newId };
    }
    return { ok: false, error: "Reserva original no encontrada" };
  });
}

function eliminarReserva(body, token) {
  return withWriteLock(function() {
    const user = getUserFromToken(token);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RESERVAS);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) !== String(body.id)) continue;
      const estado = String(data[i][9] || "confirmada").trim().toLowerCase();
      if (estado !== "cancelada") return { ok: false, error: "Solo se pueden eliminar reservas canceladas" };
      if (!puedeEliminarReservaCancelada(user)) return { ok: false, error: "Sin permiso para eliminar reservas canceladas" };
      if (!rowActiva(data[i][7])) return { ok: true, idempotent: true };
      if (!versionEsperadaCoincide(body, "reserva", data[i])) return conflictoActualizado("reserva");
      const antes = reservaAuditFromRow(data[i]);
      sheet.getRange(i + 1, 8).setValue(false);
      const despues = reservaAuditFromRow([data[i][0],data[i][1],data[i][2],data[i][3],data[i][4],data[i][5],data[i][6],false,data[i][8],data[i][9]]);
      registrarAuditoria(user, "eliminar", "reserva", String(body.id), resumenEliminacionReservaAudit(antes), antes, despues);
      invalidateAgendaCache();
      return { ok: true };
    }
    return conflictoActualizado("reserva");
  });
}

// ── Bloqueos ─────────────────────────────────────────────────
function getBloqueos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ensureBloqueosSheet(ss);
  const data = sheet.getDataRange().getValues();
  const bloqueos = [];
  for (let i = 1; i < data.length; i++) {
    const bloqueo = bloqueoFromRow(data[i]);
    if (bloqueo) bloqueos.push(bloqueo);
  }
  return { ok: true, bloqueos, version: getAgendaCacheVersion() };
}

function ensureBloqueosSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_BLOQUEOS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_BLOQUEOS);
    sheet.appendRow(["id","consultorio","franja","fecha","duracion","nota","activo","repeticion"]);
  }
  return sheet;
}

function bloqueoFromRow(row) {
  const [id, consultorio, franja, fecha, duracion, nota, activo, repeticion] = row;
  if (!id || activo === false || String(activo).toUpperCase() === "FALSE") return null;
  let consultorioIdx = consultorio === "todos" ? "todos" : Number(consultorio);
  if (typeof consultorioIdx === "number" && isNaN(consultorioIdx)) {
    consultorioIdx = NOMBRES_CONSULTORIOS.indexOf(String(consultorio));
  }
  return {
    id: String(id), consultorio: consultorioIdx, franja: Number(franja),
    fecha: fechaToString(fecha), duracion: Number(duracion), nota: String(nota || ""),
    repeticion: String(repeticion || "ninguna"),
    version: rowVersion("bloqueo", row)
  };
}

function rowActiva(value) {
  return value !== false && String(value).toUpperCase() !== "FALSE";
}

function normalizarConsultorioBloqueo(consultorio) {
  return String(consultorio).toLowerCase() === "todos" ? "todos" : normalizarConsultorioIndice(consultorio);
}

function rowVersion(tipo, row) {
  const values = tipo === "reserva"
    ? [String(row[0] || ""), normalizarConsultorioIndice(row[1]), String(row[2] || ""), fechaToString(row[3]), Number(row[4]), Number(row[5]), String(row[6] || ""), rowActiva(row[7]), String(row[8] || "normal"), String(row[9] || "confirmada").trim().toLowerCase()]
    : [String(row[0] || ""), normalizarConsultorioBloqueo(row[1]), Number(row[2]), fechaToString(row[3]), Number(row[4]), String(row[5] || ""), rowActiva(row[6]), String(row[7] || "ninguna")];
  const text = JSON.stringify(values);
  let hashA = 2166136261;
  let hashB = 5381;
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    hashA = Math.imul(hashA ^ code, 16777619);
    hashB = Math.imul(hashB, 33) ^ code;
  }
  return (hashA >>> 0).toString(16).padStart(8, "0") + (hashB >>> 0).toString(16).padStart(8, "0");
}

function versionEsperadaCoincide(body, tipo, row) {
  return !body.expectedVersion || String(body.expectedVersion) === rowVersion(tipo, row);
}

function conflictoActualizado(tipo) {
  const nombre = tipo === "bloqueo" ? "bloqueo" : "reserva";
  return {
    ok: false,
    conflict: true,
    error: `Este ${nombre} fue modificado desde otro dispositivo. La agenda se actualizó.`
  };
}

function crearBloqueo(body, token) {
  return withWriteLock(function() {
    const user = getUserFromToken(token);
    if (!puedeGestionarBloqueos(user)) return { ok: false, error: "Solo admin o asistente" };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ensureBloqueosSheet(ss);
    const consultorioVal = body.consultorio === "todos" ? "todos"
      : (NOMBRES_CONSULTORIOS[Number(body.consultorio)] || String(body.consultorio));
    const esCopia = String(body.operationKind || "") === "copiar";
    const accionOperacion = esCopia ? "copiarBloqueo" : "crearBloqueo";
    const payload = {
      sourceId: esCopia ? String(body.sourceId || "") : "",
      consultorio: String(consultorioVal), fecha: fechaToString(body.fecha), franja: Number(body.franja),
      duracion: Number(body.duracion), nota: String(body.nota || ""), repeticion: String(body.repeticion || "ninguna")
    };
    const operacion = prepararOperacionCreacion(ss, body, user, accionOperacion, "bloqueo", payload, "B");
    if (!operacion.ok) return operacion;
    const yaCreado = buscarFilaPorId(sheet, operacion.elementoId);
    if (yaCreado) {
      const despuesExistente = bloqueoAuditFromRow(yaCreado.row);
      const resumen = esCopia ? `copió bloqueo · ${resumenBloqueoAudit(despuesExistente)}` : `creó bloqueo · ${resumenBloqueoAudit(despuesExistente)}`;
      registrarAuditoriaUnaVez(user, "crear", "bloqueo", operacion.elementoId, resumen, null, despuesExistente, operacion.operationId);
      completarOperacionCreacion(ss, operacion);
      invalidateAgendaCache();
      return { ok: true, id: operacion.elementoId, idempotent: true };
    }
    if (esCopia && !buscarFilaPorId(sheet, body.sourceId)) return { ok: false, error: "El bloqueo original ya no está disponible. Cópialo nuevamente." };

    iniciarOperacionCreacion(ss, operacion, user, accionOperacion, "bloqueo");
    const id = operacion.elementoId;
    const row = [id, consultorioVal, body.franja, body.fecha, body.duracion, body.nota || "", true, body.repeticion || "ninguna"];
    sheet.appendRow(row);
    const despues = bloqueoAuditFromRow(row);
    const resumen = esCopia ? `copió bloqueo · ${resumenBloqueoAudit(despues)}` : `creó bloqueo · ${resumenBloqueoAudit(despues)}`;
    registrarAuditoriaUnaVez(user, "crear", "bloqueo", id, resumen, null, despues, operacion.operationId);
    completarOperacionCreacion(ss, operacion);
    invalidateAgendaCache();
    return { ok: true, id };
  });
}

function eliminarBloqueo(body, token) {
  return withWriteLock(function() {
    const user = getUserFromToken(token);
    if (!puedeGestionarBloqueos(user)) return { ok: false, error: "Solo admin o asistente" };
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_BLOQUEOS);
    if (!sheet) return { ok: false, error: "No existe hoja de bloqueos" };
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(body.id)) {
        if (!rowActiva(data[i][6])) return { ok: true, idempotent: true };
        if (!versionEsperadaCoincide(body, "bloqueo", data[i])) return conflictoActualizado("bloqueo");
        const antes = bloqueoAuditFromRow(data[i]);
        sheet.getRange(i + 1, 7).setValue(false);
        const despues = bloqueoAuditFromRow([data[i][0],data[i][1],data[i][2],data[i][3],data[i][4],data[i][5],false,data[i][7]]);
        registrarAuditoria(user, "eliminar", "bloqueo", String(body.id), `eliminó bloqueo · ${resumenBloqueoAudit(antes)}`, antes, despues);
        invalidateAgendaCache();
        return { ok: true };
      }
    }
    return conflictoActualizado("bloqueo");
  });
}

function moverBloqueo(body, token) {
  return withWriteLock(function() {
    const user = getUserFromToken(token);
    if (!puedeGestionarBloqueos(user)) return { ok: false, error: "Solo admin o asistente" };
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_BLOQUEOS);
    if (!sheet) return { ok: false, error: "No existe hoja de bloqueos" };
    const data = sheet.getDataRange().getValues();
    const consultorioVal = body.consultorio === "todos" ? "todos"
      : (NOMBRES_CONSULTORIOS[Number(body.consultorio)] || String(body.consultorio));

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(body.id)) {
        if (!rowActiva(data[i][6])) return conflictoActualizado("bloqueo");
        const mismoDestino = normalizarConsultorioBloqueo(data[i][1]) === normalizarConsultorioBloqueo(consultorioVal) &&
          fechaToString(data[i][3]) === fechaToString(body.fecha) && Number(data[i][2]) === Number(body.franja);
        if (mismoDestino) return { ok: true, idempotent: true, id: String(body.id), version: rowVersion("bloqueo", data[i]) };
        if (!versionEsperadaCoincide(body, "bloqueo", data[i])) return conflictoActualizado("bloqueo");
        const antes = bloqueoAuditFromRow(data[i]);
        sheet.getRange(i + 1, 2).setValue(consultorioVal);
        sheet.getRange(i + 1, 3).setValue(body.franja);
        sheet.getRange(i + 1, 4).setValue(body.fecha);
        const despues = bloqueoAuditFromRow([data[i][0],consultorioVal,body.franja,body.fecha,data[i][4],data[i][5],data[i][6],data[i][7]]);
        registrarAuditoria(user, "mover", "bloqueo", String(body.id), `movió bloqueo · ${resumenBloqueoAudit(antes)} → ${despues.consultorio} · ${despues.fecha} · ${despues.hora}`, antes, despues);
        invalidateAgendaCache();
        return { ok: true, id: String(body.id), version: rowVersion("bloqueo", [data[i][0],consultorioVal,body.franja,body.fecha,data[i][4],data[i][5],data[i][6],data[i][7]]) };
      }
    }
    return conflictoActualizado("bloqueo");
  });
}

// ── Diagnóstico de datos (solo lectura) ──────────────────────
function diagnosticarDatos(token) {
  const user = getUserFromToken(token);
  if (user.rol !== "admin") return { ok: false, error: "Solo admin" };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const usuarios = diagnosticarUsuarios(ss);
  const reservas = diagnosticarReservas(ss);
  const identidadReservas = diagnosticarIdentidadReservas(ss);
  const bloqueos = diagnosticarBloqueos(ss);
  return {
    ok: true,
    generadoEn: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss"),
    resumen: {
      usuariosRegistrados: usuarios.registrados,
      usuariosDuplicados: usuarios.totalDuplicados,
      usuariosInvalidos: usuarios.totalInvalidos,
      nombresRepetidos: usuarios.totalNombresRepetidos,
      nombresSimilares: usuarios.totalNombresSimilares,
      correosDuplicados: usuarios.totalCorreosDuplicados,
      reservasActivas: reservas.activas,
      reservasDuplicadas: reservas.totalDuplicados,
      reservasInvalidas: reservas.totalInvalidos,
      reservasIdentidadAmbigua: identidadReservas.totalProblemas,
      bloqueosActivos: bloqueos.activos,
      bloqueosDuplicados: bloqueos.totalDuplicados,
      bloqueosInvalidos: bloqueos.totalInvalidos
    },
    usuarios,
    reservas,
    identidadReservas,
    bloqueos
  };
}

function diagnosticarUsuarios(ss) {
  const sheet = ss.getSheetByName(SHEET_USUARIOS);
  const result = {
    registrados: 0,
    totalDuplicados: 0,
    totalInvalidos: 0,
    totalNombresRepetidos: 0,
    totalNombresSimilares: 0,
    totalCorreosDuplicados: 0,
    duplicados: [],
    invalidos: [],
    nombresRepetidos: [],
    nombresSimilares: [],
    correosDuplicados: []
  };
  if (!sheet) {
    result.invalidos.push({ fila: null, id: "", problemas: ["No existe hoja de usuarios"] });
    result.totalInvalidos = result.invalidos.length;
    return result;
  }

  const registro = leerRegistroUsuarios(ss);
  const porId = {};
  const porNombre = {};
  const porCorreo = {};
  registro.usuarios.forEach(usuario => {
    const { id: cleanId, nombre, rol: cleanRole, correo, fila } = usuario;
    result.registrados++;
    const problemas = [];
    if (!cleanId) problemas.push("Sin id");
    if (!nombre) problemas.push("Sin nombre");
    if (!ROLES_USUARIO.includes(cleanRole)) problemas.push("Rol inválido");
    if (problemas.length) result.invalidos.push({ fila, id: cleanId, problemas });
    if (cleanId) {
      const key = "id:" + cleanId.toLowerCase();
      if (!porId[key]) porId[key] = [];
      porId[key].push({ fila, id: cleanId, nombre, problemas: [] });
    }
    if (usuario.nombreKey) {
      const key = "nombre:" + usuario.nombreKey;
      if (!porNombre[key]) porNombre[key] = [];
      porNombre[key].push({ fila, id: cleanId, nombre, problemas: [] });
    }
    if (usuario.correoKey) {
      const key = "correo:" + usuario.correoKey;
      if (!porCorreo[key]) porCorreo[key] = [];
      porCorreo[key].push({ fila, id: cleanId, nombre, correo, problemas: [] });
    }
  });

  const duplicados = gruposDuplicados(porId, "ID de usuario repetido");
  const nombresRepetidos = gruposDuplicados(porNombre, "Nombre visible repetido (permitido, revisar IDs)");
  const correosDuplicados = gruposDuplicados(porCorreo, "Correo repetido");
  const similares = detectarNombresSimilares(registro.usuarios);
  result.totalDuplicados = duplicados.length;
  result.totalInvalidos = result.invalidos.length;
  result.totalNombresRepetidos = nombresRepetidos.length;
  result.totalNombresSimilares = similares.length;
  result.totalCorreosDuplicados = correosDuplicados.length;
  result.duplicados = limitarEjemplos(duplicados);
  result.invalidos = limitarEjemplos(result.invalidos);
  result.nombresRepetidos = limitarEjemplos(nombresRepetidos);
  result.nombresSimilares = limitarEjemplos(similares);
  result.correosDuplicados = limitarEjemplos(correosDuplicados);
  return result;
}

function detectarNombresSimilares(usuarios) {
  const ejemplos = [];
  for (let i = 0; i < usuarios.length; i++) {
    for (let j = i + 1; j < usuarios.length; j++) {
      const a = usuarios[i];
      const b = usuarios[j];
      if (!a.nombreKey || !b.nombreKey || a.nombreKey === b.nombreKey) continue;
      const primeroA = a.nombreKey.split(" ")[0];
      const primeroB = b.nombreKey.split(" ")[0];
      if (primeroA.length >= 4 && primeroA === primeroB) {
        ejemplos.push({
          tipo: "Nombres similares (solo aviso)",
          cantidad: 2,
          items: [
            { fila: a.fila, id: a.id, nombre: a.nombre },
            { fila: b.fila, id: b.id, nombre: b.nombre }
          ]
        });
      }
    }
  }
  return ejemplos;
}

function diagnosticarIdentidadReservas(ss) {
  const result = { revisadas: 0, totalProblemas: 0, problemas: [] };
  const sheet = ss.getSheetByName(SHEET_RESERVAS);
  if (!sheet) return result;
  const registro = leerRegistroUsuarios(ss);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    result.revisadas++;
    const userId = String(row[2] || "").trim();
    const problemas = [];
    const matches = registro.porId[normalizarUserId(userId)] || [];
    if (!userId) problemas.push("Reserva sin ID de profesional");
    else if (matches.length > 1) problemas.push(`ID de profesional ambiguo: ${userId}`);
    else if (matches.length === 0) {
      const porNombre = registro.porNombre[normalizarNombre(userId)] || [];
      if (porNombre.length) problemas.push(`El valor "${userId}" coincide con un nombre, no con un ID; requiere revisión manual`);
      else problemas.push(`Profesional no encontrado: ${userId}`);
    } else if (matches[0].id !== userId) {
      problemas.push(`El ID "${userId}" difiere en mayúsculas/minúsculas del ID canónico "${matches[0].id}"; requiere revisión manual`);
    } else if (matches[0].rol !== "profesional") {
      problemas.push(`El ID ${userId} pertenece al rol ${matches[0].rol}, no a un profesional`);
    }
    if (problemas.length) {
      result.problemas.push({
        fila: i + 1,
        id: String(row[0] || ""),
        profesionalId: userId,
        fecha: fechaToString(row[3]),
        problemas
      });
    }
  }
  result.totalProblemas = result.problemas.length;
  result.problemas = limitarEjemplos(result.problemas);
  return result;
}

function diagnosticarReservas(ss) {
  const sheet = ss.getSheetByName(SHEET_RESERVAS);
  const result = { activas: 0, totalDuplicados: 0, totalInvalidos: 0, duplicados: [], invalidos: [] };
  if (!sheet) {
    result.invalidos.push({ fila: null, id: "", problemas: ["No existe hoja de reservas"] });
    result.totalInvalidos = result.invalidos.length;
    return result;
  }

  const data = sheet.getDataRange().getValues();
  const porId = {};
  const porSlot = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const [id, consultorio, userId, fecha, franja, duracion, , activa, , estado] = row;
    if (activa === false || String(activa).toUpperCase() === "FALSE") continue;
    if (String(estado || "confirmada").toLowerCase() === "cancelada") continue;
    result.activas++;

    const problemas = [];
    const consIdx = normalizarConsultorioIndice(consultorio);
    const fechaStr = fechaToString(fecha);
    const franjaNum = Number(franja);
    const duracionNum = Number(duracion);

    if (!id) problemas.push("Sin id");
    if (consIdx < 0 || consIdx >= NOMBRES_CONSULTORIOS.length) problemas.push("Consultorio inválido");
    if (!userId) problemas.push("Sin profesional");
    if (!esFechaValida(fechaStr)) problemas.push("Fecha inválida");
    if (!esFranjaValida(franjaNum)) problemas.push("Franja inválida");
    if (!esDuracionValida(duracionNum, franjaNum)) problemas.push("Duración inválida");

    if (problemas.length) {
      result.invalidos.push(ejemploDiagnostico(i + 1, id, consultorio, fechaStr, franja, duracion, problemas));
    }

    const idKey = String(id || "").trim();
    if (idKey) agregarGrupoDiagnostico(porId, "id:" + idKey, i + 1, id, consultorio, fechaStr, franja, duracion);
    if (idKey && consIdx >= 0 && esFechaValida(fechaStr) && esFranjaValida(franjaNum) && esDuracionValida(duracionNum, franjaNum)) {
      const slotKey = ["slot", consIdx, fechaStr, franjaNum, duracionNum, String(userId || "")].join("|");
      agregarGrupoDiagnostico(porSlot, slotKey, i + 1, id, consultorio, fechaStr, franja, duracion);
    }
  }

  const duplicados = gruposDuplicados(porId, "Id repetido").concat(gruposDuplicados(porSlot, "Misma reserva activa"));
  result.totalDuplicados = duplicados.length;
  result.totalInvalidos = result.invalidos.length;
  result.duplicados = duplicados;
  result.invalidos = limitarEjemplos(result.invalidos);
  result.duplicados = limitarEjemplos(result.duplicados);
  return result;
}

function diagnosticarBloqueos(ss) {
  const sheet = ss.getSheetByName(SHEET_BLOQUEOS);
  const result = { activas: 0, totalDuplicados: 0, totalInvalidos: 0, duplicados: [], invalidos: [] };
  if (!sheet) return result;

  const data = sheet.getDataRange().getValues();
  const porId = {};
  const porBloqueo = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const [id, consultorio, franja, fecha, duracion, , activo, repeticion] = row;
    if (activo === false || String(activo).toUpperCase() === "FALSE") continue;
    result.activas++;

    const problemas = [];
    const consIdx = consultorio === "todos" ? "todos" : normalizarConsultorioIndice(consultorio);
    const fechaStr = fechaToString(fecha);
    const franjaNum = Number(franja);
    const duracionNum = Number(duracion);
    const rep = String(repeticion || "ninguna");

    if (!id) problemas.push("Sin id");
    if (consIdx !== "todos" && (consIdx < 0 || consIdx >= NOMBRES_CONSULTORIOS.length)) problemas.push("Consultorio inválido");
    if (!esFechaValida(fechaStr)) problemas.push("Fecha inválida");
    if (!esFranjaValida(franjaNum)) problemas.push("Franja inválida");
    if (!esDuracionValida(duracionNum, franjaNum)) problemas.push("Duración inválida");
    if (rep !== "ninguna" && rep !== "semanal") problemas.push("Repetición inválida");

    if (problemas.length) {
      result.invalidos.push(ejemploDiagnostico(i + 1, id, consultorio, fechaStr, franja, duracion, problemas));
    }

    const idKey = String(id || "").trim();
    if (idKey) agregarGrupoDiagnostico(porId, "id:" + idKey, i + 1, id, consultorio, fechaStr, franja, duracion);
    if (idKey && (consIdx === "todos" || consIdx >= 0) && esFechaValida(fechaStr) && esFranjaValida(franjaNum) && esDuracionValida(duracionNum, franjaNum)) {
      const bloqueoKey = ["bloqueo", consIdx, fechaStr, franjaNum, duracionNum, rep].join("|");
      agregarGrupoDiagnostico(porBloqueo, bloqueoKey, i + 1, id, consultorio, fechaStr, franja, duracion);
    }
  }

  const duplicados = gruposDuplicados(porId, "Id repetido").concat(gruposDuplicados(porBloqueo, "Mismo bloqueo activo"));
  result.totalDuplicados = duplicados.length;
  result.totalInvalidos = result.invalidos.length;
  result.duplicados = duplicados;
  result.invalidos = limitarEjemplos(result.invalidos);
  result.duplicados = limitarEjemplos(result.duplicados);
  return result;
}

function normalizarConsultorioIndice(consultorio) {
  let idx = Number(consultorio);
  if (isNaN(idx)) idx = NOMBRES_CONSULTORIOS.indexOf(String(consultorio));
  return idx;
}

function esFechaValida(fechaStr) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(fechaStr))) return false;
  const d = parseFechaLocal(fechaStr);
  return !isNaN(d.getTime()) && fechaToString(d) === fechaStr;
}

function esFranjaValida(franja) {
  return Number.isInteger(Number(franja)) && Number(franja) >= 0 && Number(franja) < 20;
}

function esDuracionValida(duracion, franja) {
  const dur = Number(duracion);
  const fra = Number(franja);
  return Number.isFinite(dur) && dur > 0 && dur % 30 === 0 && Number.isFinite(fra) && fra * 30 + dur <= 600;
}

function ejemploDiagnostico(fila, id, consultorio, fecha, franja, duracion, problemas) {
  return {
    fila,
    id: String(id || ""),
    consultorio: String(consultorio || ""),
    fecha: String(fecha || ""),
    franja: String(franja || ""),
    duracion: String(duracion || ""),
    problemas
  };
}

function agregarGrupoDiagnostico(map, key, fila, id, consultorio, fecha, franja, duracion) {
  if (!map[key]) map[key] = [];
  map[key].push(ejemploDiagnostico(fila, id, consultorio, fecha, franja, duracion, []));
}

function gruposDuplicados(map, tipo) {
  return Object.keys(map).filter(k => map[k].length > 1).map(k => ({
    tipo,
    cantidad: map[k].length,
    items: map[k]
  }));
}

function limitarEjemplos(items) {
  return items.slice(0, 10);
}

// ── Reservas preestablecidas ──────────────────────────────────
function generarPreestablecidas(body, token) {
  return withWriteLock(function() {
    const user = getUserFromToken(token);
    if (user.rol !== "admin") return { ok: false, error: "Solo admin" };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const daniela = resolverUsuarioUnico(ss, "daniela", { rol: "profesional" });
    const ramiro = resolverUsuarioUnico(ss, "ramiro", { rol: "profesional" });
    if (!daniela.ok) return { ok: false, error: `No se generaron reservas fijas: ${daniela.error}` };
    if (!ramiro.ok) return { ok: false, error: `No se generaron reservas fijas: ${ramiro.error}` };
    const sheet = ss.getSheetByName(SHEET_RESERVAS);
    const hoy = new Date();
    let creadas = 0;

    for (let semana = 0; semana < 4; semana++) {
      for (let dia = 0; dia < 5; dia++) {
        const fecha = new Date(hoy);
        const diffLunes = (dia + (1 - hoy.getDay() + 7) % 7) + semana * 7;
        fecha.setDate(hoy.getDate() + diffLunes);
        if (fecha < hoy) continue;
        const fechaStr = Utilities.formatDate(fecha, Session.getScriptTimeZone(), "yyyy-MM-dd");

        // Daniela: Consultorio 1, 2pm(franja 12) duración 180min
        if (!hayConflicto(0, fechaStr, 12, 180, null)) {
          sheet.appendRow(["PRE_D_"+fechaStr, "Consultorio 1", daniela.usuario.id, fechaStr, 12, 180, "Reserva fija Daniela", true, "preestablecida", "confirmada"]);
          creadas++;
        }
        // Ramiro: Consultorio 3, 10am(franja 4) duración 120min
        if (!hayConflicto(2, fechaStr, 4, 120, null)) {
          sheet.appendRow(["PRE_R1_"+fechaStr, "Consultorio 3", ramiro.usuario.id, fechaStr, 4, 120, "Reserva fija Ramiro mañana", true, "preestablecida", "confirmada"]);
          creadas++;
        }
        // Ramiro: Consultorio 3, 2pm(franja 12) duración 120min
        if (!hayConflicto(2, fechaStr, 12, 120, null)) {
          sheet.appendRow(["PRE_R2_"+fechaStr, "Consultorio 3", ramiro.usuario.id, fechaStr, 12, 120, "Reserva fija Ramiro tarde", true, "preestablecida", "confirmada"]);
          creadas++;
        }
      }
    }
    if (creadas > 0) invalidateAgendaCache();
    return { ok: true, creadas };
  });
}

// ── Validaciones ─────────────────────────────────────────────
function parseFechaLocal(fechaStr) {
  // Evita desfase UTC: parsea yyyy-MM-dd como fecha local
  const [y, m, d] = String(fechaStr).slice(0, 10).split("-").map(Number);
  return new Date(y, m - 1, d);
}

function fechaToString(fecha) {
  return fecha instanceof Date
    ? Utilities.formatDate(fecha, Session.getScriptTimeZone(), "yyyy-MM-dd")
    : String(fecha).slice(0, 10);
}

function bloqueoAplicaEnRango(fechaBloqueo, desde, hasta) {
  const targetDay = parseFechaLocal(fechaBloqueo).getDay();
  const cursor = parseFechaLocal(desde);
  const end = parseFechaLocal(hasta);
  while (cursor <= end) {
    if (cursor.getDay() === targetDay) return true;
    cursor.setDate(cursor.getDate() + 1);
  }
  return false;
}

function getAgendaCacheVersion() {
  return PropertiesService.getScriptProperties().getProperty("agendaCacheVersion") || "0";
}

function getCacheValue(key) {
  try {
    return CacheService.getScriptCache().get(key);
  } catch (err) {
    return null;
  }
}

function putCacheValue(key, value, ttlSeconds) {
  try {
    CacheService.getScriptCache().put(key, value, ttlSeconds);
  } catch (err) {}
}

function invalidateAgendaCache() {
  PropertiesService.getScriptProperties().setProperty("agendaCacheVersion", String(Date.now()));
}

function withWriteLock(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return { ok: false, error: "Sistema ocupado, intenta de nuevo" };
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function estaBloquado(consultorio, fecha, franja, duracion) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_BLOQUEOS);
  if (!sheet) return false;
  const data = sheet.getDataRange().getValues();
  const startMin = Number(franja) * 30;
  const endMin = startMin + Number(duracion);
  const fechaStr = fechaToString(fecha);
  const diaSemana = parseFechaLocal(fechaStr).getDay();

  for (let i = 1; i < data.length; i++) {
    const [id, cons, fra, fec, dur, , activo, repeticion] = data[i];
    if (!activo || activo === false || String(activo).toUpperCase() === "FALSE") continue;

    if (cons !== "todos") {
      let consIdx = Number(cons);
      if (isNaN(consIdx)) consIdx = NOMBRES_CONSULTORIOS.indexOf(String(cons));
      let inputIdx = Number(consultorio);
      if (isNaN(inputIdx)) inputIdx = NOMBRES_CONSULTORIOS.indexOf(String(consultorio));
      if (consIdx !== inputIdx) continue;
    }

    const bloqueoFecha = fechaToString(fec);

    const rep = String(repeticion || "ninguna");
    const mismaFecha = bloqueoFecha === fechaStr;
    const mismoDia = rep === "semanal" && parseFechaLocal(bloqueoFecha).getDay() === diaSemana;
    if (!mismaFecha && !mismoDia) continue;

    const bStart = Number(fra) * 30;
    const bEnd = bStart + Number(dur);
    if (startMin < bEnd && endMin > bStart) return true;
  }
  return false;
}

function hayConflicto(consultorio, fecha, franja, duracion, excludeId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RESERVAS);
  const data = sheet.getDataRange().getValues();
  const startMin = Number(franja) * 30;
  const endMin = startMin + Number(duracion);
  const targetFecha = fechaToString(fecha);

  for (let i = 1; i < data.length; i++) {
    const [id, cons, , fec, fra, dur, , activa, , estado] = data[i];
    if (activa === false || String(activa).toUpperCase() === "FALSE") continue;
    if (String(estado).toLowerCase() === "cancelada") continue;
    if (String(id) === String(excludeId)) continue;

    let consIdx = Number(cons);
    if (isNaN(consIdx)) consIdx = NOMBRES_CONSULTORIOS.indexOf(String(cons));
    let inputIdx = Number(consultorio);
    if (isNaN(inputIdx)) inputIdx = NOMBRES_CONSULTORIOS.indexOf(String(consultorio));
    if (consIdx !== inputIdx) continue;

    let fechaStr = fechaToString(fec);
    if (fechaStr !== targetFecha) continue;

    const rStart = Number(fra) * 30;
    const rEnd = rStart + Number(dur);
    if (startMin < rEnd && endMin > rStart) return true;
  }
  return false;
}

function resp(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
