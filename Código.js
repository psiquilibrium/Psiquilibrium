// ============================================================
//  PSIQUILIBRIUM — Backend v3 (franjas de 30 min)
// ============================================================

const SHEET_RESERVAS = "Reservas";
const SHEET_USUARIOS = "Usuarios";
const SHEET_BLOQUEOS = "Bloqueos";
const NOMBRES_CONSULTORIOS = ["Consultorio 1","Consultorio 2","Consultorio 3","Consultorio 4"];
const CACHE_AGENDA_PREFIX = "agenda_v1_";

// Franjas: 0=8:00, 1=8:30, 2=9:00 ... 19=17:30, 20=18:00 (no incluida)
// Sábado hasta las 12:00 = franja 8

function doGet(e) { return handle(e); }
function doPost(e) { return handle(e); }
function doOptions(e) { return ContentService.createTextOutput("").setMimeType(ContentService.MimeType.TEXT); }

function handle(e) {
  try {
    const params = e.parameter || {};
    const body = (e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : {};
    const merged = { ...params, ...body };
    const action = merged.action;
    const token = merged.token;

    if (action === "login") return resp(login(merged));
    if (!validarToken(token)) return resp({ ok: false, error: "No autorizado" });

    switch (action) {
      case "getAgenda":              return resp(getAgenda(merged));
      case "getAgendaVersion":       return resp(getAgendaVersion(merged));
      case "getReservas":            return resp(getReservas());
      case "crearReserva":           return resp(crearReserva(merged, token));
      case "editarReserva":          return resp(editarReserva(merged, token));
      case "eliminarReserva":        return resp(eliminarReserva(merged, token));
      case "cambiarEstado":          return resp(cambiarEstado(merged, token));
      case "moverReserva":           return resp(moverReserva(merged, token));
      case "copiarReserva":          return resp(copiarReserva(merged, token));
      case "crearBloqueo":           return resp(crearBloqueo(merged, token));
      case "moverBloqueo":           return resp(moverBloqueo(merged, token));
      case "eliminarBloqueo":        return resp(eliminarBloqueo(merged, token));
      case "getBloqueos":            return resp(getBloqueos());
      case "generarPreestablecidas": return resp(generarPreestablecidas(merged, token));
      case "migrarFranjas":          return resp(migrarFranjas(token));
      default: return resp({ ok: false, error: "Acción no reconocida" });
    }
  } catch (err) {
    return resp({ ok: false, error: err.message });
  }
}

// ── Autenticación ────────────────────────────────────────────
function login(body) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USUARIOS);
  const data = sheet.getDataRange().getValues();
  const users = [];
  let matchedUser = null;

  for (let i = 1; i < data.length; i++) {
    const [id, nombre, rol, pass] = data[i];
    if (!id) continue;
    users.push({ id: String(id), nombre: String(nombre), rol: String(rol) });
    if (String(id).trim() === String(body.userId).trim() &&
        String(pass).trim() === String(body.password).trim()) {
      matchedUser = { id: String(id), nombre: String(nombre), rol: String(rol) };
    }
  }

  if (matchedUser) {
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
  return { id: parts[0], rol: parts[1] };
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
function getReservas() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RESERVAS);
  const data = sheet.getDataRange().getValues();
  const reservas = [];

  for (let i = 1; i < data.length; i++) {
    const reserva = reservaFromRow(data[i]);
    if (reserva) reservas.push(reserva);
  }
  return { ok: true, reservas };
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
    estado: String(estado || "confirmada")
  };
}

function crearReserva(body, token) {
  return withWriteLock(function() {
    const user = getUserFromToken(token);
    const { consultorio, fecha, franja, duracion, nota } = body;
    const userId = (user.rol === "admin" && body.targetUserId) ? body.targetUserId : user.id;

    if (estaBloquado(consultorio, fecha, franja, duracion)) return { ok: false, error: "Franja bloqueada" };
    if (hayConflicto(consultorio, fecha, franja, duracion, null)) return { ok: false, error: "Conflicto de horario" };

    const consultorioNombre = NOMBRES_CONSULTORIOS[Number(consultorio)] || String(consultorio);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RESERVAS);
    const id = "R_" + Date.now();
    sheet.appendRow([id, consultorioNombre, userId, fecha, franja, duracion, nota || "", true, body.tipo || "normal", "confirmada"]);
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
      if (user.rol !== "admin" && user.id !== String(data[i][2])) return { ok: false, error: "Sin permiso" };
      if (hayConflicto(data[i][1], data[i][3], data[i][4], body.duracion, body.id)) return { ok: false, error: "Conflicto de horario" };
      sheet.getRange(i + 1, 6).setValue(body.duracion);
      sheet.getRange(i + 1, 7).setValue(body.nota || "");
      invalidateAgendaCache();
      return { ok: true };
    }
    return { ok: false, error: "Reserva no encontrada" };
  });
}

function cambiarEstado(body, token) {
  return withWriteLock(function() {
    const user = getUserFromToken(token);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RESERVAS);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) !== String(body.id)) continue;
      const propietario = String(data[i][2]);
      // Admin puede cambiar cualquier estado. El dueño solo puede cancelar la suya.
      if (user.rol !== "admin") {
        if (user.id !== propietario) return { ok: false, error: "Sin permiso" };
        if (body.estado !== "cancelada") return { ok: false, error: "Solo puedes cancelar tus propias reservas" };
      }
      sheet.getRange(i + 1, 10).setValue(body.estado);
      invalidateAgendaCache();
      return { ok: true };
    }
    return { ok: false, error: "Reserva no encontrada" };
  });
}

function moverReserva(body, token) {
  return withWriteLock(function() {
    const user = getUserFromToken(token);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RESERVAS);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) !== String(body.id)) continue;
      if (user.rol !== "admin" && user.id !== String(data[i][2])) return { ok: false, error: "Sin permiso" };
      const dur = Number(data[i][5]);
      if (estaBloquado(body.consultorio, body.fecha, body.franja, dur)) return { ok: false, error: "Franja bloqueada" };
      if (hayConflicto(body.consultorio, body.fecha, body.franja, dur, body.id)) return { ok: false, error: "Conflicto de horario" };
      const consultorioNombre = NOMBRES_CONSULTORIOS[Number(body.consultorio)] || String(body.consultorio);
      sheet.getRange(i + 1, 2).setValue(consultorioNombre);
      sheet.getRange(i + 1, 4).setValue(body.fecha);
      sheet.getRange(i + 1, 5).setValue(body.franja);
      invalidateAgendaCache();
      return { ok: true };
    }
    return { ok: false, error: "Reserva no encontrada" };
  });
}

function copiarReserva(body, token) {
  return withWriteLock(function() {
    const user = getUserFromToken(token);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RESERVAS);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) !== String(body.id)) continue;
      const [, , userId, , , duracion, nota, , tipo] = data[i];
      const targetUserId = String(userId);
      if (estaBloquado(body.consultorio, body.fecha, body.franja, duracion)) return { ok: false, error: "Franja bloqueada" };
      if (hayConflicto(body.consultorio, body.fecha, body.franja, duracion, null)) return { ok: false, error: "Conflicto de horario" };
      const consultorioNombre = NOMBRES_CONSULTORIOS[Number(body.consultorio)] || String(body.consultorio);
      const newId = "R_" + Date.now();
      sheet.appendRow([newId, consultorioNombre, targetUserId, body.fecha, body.franja, duracion, nota || "", true, tipo || "normal", "confirmada"]);
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
      if (user.rol !== "admin" && user.id !== String(data[i][2])) return { ok: false, error: "Sin permiso" };
      sheet.getRange(i + 1, 8).setValue(false);
      invalidateAgendaCache();
      return { ok: true };
    }
    return { ok: false, error: "Reserva no encontrada" };
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
  return { ok: true, bloqueos };
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
    repeticion: String(repeticion || "ninguna")
  };
}

function crearBloqueo(body, token) {
  return withWriteLock(function() {
    const user = getUserFromToken(token);
    if (user.rol !== "admin") return { ok: false, error: "Solo admin" };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ensureBloqueosSheet(ss);
    const consultorioVal = body.consultorio === "todos" ? "todos"
      : (NOMBRES_CONSULTORIOS[Number(body.consultorio)] || String(body.consultorio));
    const id = "B_" + Date.now();
    sheet.appendRow([id, consultorioVal, body.franja, body.fecha, body.duracion, body.nota || "", true, body.repeticion || "ninguna"]);
    invalidateAgendaCache();
    return { ok: true, id };
  });
}

function eliminarBloqueo(body, token) {
  return withWriteLock(function() {
    const user = getUserFromToken(token);
    if (user.rol !== "admin") return { ok: false, error: "Solo admin" };
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_BLOQUEOS);
    if (!sheet) return { ok: false, error: "No existe hoja de bloqueos" };
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(body.id)) {
        sheet.getRange(i + 1, 7).setValue(false);
        invalidateAgendaCache();
        return { ok: true };
      }
    }
    return { ok: false, error: "Bloqueo no encontrado" };
  });
}

function moverBloqueo(body, token) {
  return withWriteLock(function() {
    const user = getUserFromToken(token);
    if (user.rol !== "admin") return { ok: false, error: "Solo admin" };
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_BLOQUEOS);
    if (!sheet) return { ok: false, error: "No existe hoja de bloqueos" };
    const data = sheet.getDataRange().getValues();
    const consultorioVal = body.consultorio === "todos" ? "todos"
      : (NOMBRES_CONSULTORIOS[Number(body.consultorio)] || String(body.consultorio));

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(body.id)) {
        sheet.getRange(i + 1, 2).setValue(consultorioVal);
        sheet.getRange(i + 1, 3).setValue(body.franja);
        sheet.getRange(i + 1, 4).setValue(body.fecha);
        invalidateAgendaCache();
        return { ok: true, id: String(body.id) };
      }
    }
    return { ok: false, error: "Bloqueo no encontrado" };
  });
}

// ── Reservas preestablecidas ──────────────────────────────────
function generarPreestablecidas(body, token) {
  return withWriteLock(function() {
    const user = getUserFromToken(token);
    if (user.rol !== "admin") return { ok: false, error: "Solo admin" };

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RESERVAS);
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
          sheet.appendRow(["PRE_D_"+fechaStr, "Consultorio 1", "daniela", fechaStr, 12, 180, "Reserva fija Daniela", true, "preestablecida", "confirmada"]);
          creadas++;
        }
        // Ramiro: Consultorio 3, 10am(franja 4) duración 120min
        if (!hayConflicto(2, fechaStr, 4, 120, null)) {
          sheet.appendRow(["PRE_R1_"+fechaStr, "Consultorio 3", "ramiro", fechaStr, 4, 120, "Reserva fija Ramiro mañana", true, "preestablecida", "confirmada"]);
          creadas++;
        }
        // Ramiro: Consultorio 3, 2pm(franja 12) duración 120min
        if (!hayConflicto(2, fechaStr, 12, 120, null)) {
          sheet.appendRow(["PRE_R2_"+fechaStr, "Consultorio 3", "ramiro", fechaStr, 12, 120, "Reserva fija Ramiro tarde", true, "preestablecida", "confirmada"]);
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
