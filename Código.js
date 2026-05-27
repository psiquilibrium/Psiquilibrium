// ============================================================
//  PSIQUILIBRIUM — Backend v3 (franjas de 30 min)
// ============================================================

const SHEET_RESERVAS = "Reservas";
const SHEET_USUARIOS = "Usuarios";
const SHEET_BLOQUEOS = "Bloqueos";
const SHEET_AUDITORIA = "Auditoria";
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
      case "getAuditoria":           return resp(getAuditoria(merged, token));
      case "crearRespaldoManual":    return resp(crearRespaldoManual(token));
      case "generarPreestablecidas": return resp(generarPreestablecidas(merged, token));
      case "diagnosticarDatos":      return resp(diagnosticarDatos(token));
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

function getUserNameById(userId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USUARIOS);
  if (!sheet) return String(userId || "");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(userId)) return String(data[i][1] || userId);
  }
  return String(userId || "");
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

// ── Auditoría ────────────────────────────────────────────────
function ensureAuditoriaSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_AUDITORIA);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_AUDITORIA);
    sheet.appendRow(["timestamp","userId","nombre","rol","accion","tipo","elementoId","resumen","antes","despues"]);
  }
  return sheet;
}

function registrarAuditoria(user, accion, tipo, elementoId, resumen, antes, despues) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ensureAuditoriaSheet(ss);
    const nombre = getUserNameById(user.id) || user.id;
    sheet.appendRow([
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
    ]);
  } catch (err) {}
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
    const hojas = [SHEET_RESERVAS, SHEET_BLOQUEOS, SHEET_USUARIOS, SHEET_AUDITORIA];
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
    const userId = (esRolOperativo(user) && body.targetUserId) ? body.targetUserId : user.id;

    if (estaBloquado(consultorio, fecha, franja, duracion)) return { ok: false, error: "Franja bloqueada" };
    if (hayConflicto(consultorio, fecha, franja, duracion, null)) return { ok: false, error: "Conflicto de horario" };

    const consultorioNombre = NOMBRES_CONSULTORIOS[Number(consultorio)] || String(consultorio);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RESERVAS);
    const id = "R_" + Date.now();
    const row = [id, consultorioNombre, userId, fecha, franja, duracion, nota || "", true, body.tipo || "normal", "confirmada"];
    sheet.appendRow(row);
    const despues = reservaAuditFromRow(row);
    registrarAuditoria(user, "crear", "reserva", id, `creó reserva de ${resumenReservaAudit(despues)}`, null, despues);
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
      if (hayConflicto(data[i][1], data[i][3], data[i][4], body.duracion, body.id)) return { ok: false, error: "Conflicto de horario" };
      const antes = reservaAuditFromRow(data[i]);
      sheet.getRange(i + 1, 6).setValue(body.duracion);
      sheet.getRange(i + 1, 7).setValue(body.nota || "");
      const despues = reservaAuditFromRow([data[i][0],data[i][1],data[i][2],data[i][3],data[i][4],body.duracion,body.nota || "",data[i][7],data[i][8],data[i][9]]);
      registrarAuditoria(user, "editar", "reserva", String(body.id), `editó reserva de ${resumenReservaAudit(despues)}`, antes, despues);
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
      // Admin puede cambiar cualquier estado. Asistente puede cancelar cualquiera. El dueño solo puede cancelar la suya.
      if (user.rol === "asistente") {
        if (body.estado !== "cancelada") return { ok: false, error: "Solo admin puede reconfirmar reservas" };
      } else if (user.rol !== "admin") {
        if (user.id !== propietario) return { ok: false, error: "Sin permiso" };
        if (body.estado !== "cancelada") return { ok: false, error: "Solo puedes cancelar tus propias reservas" };
      }
      const antes = reservaAuditFromRow(data[i]);
      sheet.getRange(i + 1, 10).setValue(body.estado);
      const despues = reservaAuditFromRow([data[i][0],data[i][1],data[i][2],data[i][3],data[i][4],data[i][5],data[i][6],data[i][7],data[i][8],body.estado]);
      const accion = body.estado === "cancelada" ? "cancelar" : "reconfirmar";
      registrarAuditoria(user, accion, "reserva", String(body.id), `${accion === "cancelar" ? "canceló" : "reconfirmó"} reserva de ${resumenReservaAudit(despues)}`, antes, despues);
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
      if (!esRolOperativo(user) && user.id !== String(data[i][2])) return { ok: false, error: "Sin permiso" };
      const dur = Number(data[i][5]);
      if (estaBloquado(body.consultorio, body.fecha, body.franja, dur)) return { ok: false, error: "Franja bloqueada" };
      if (hayConflicto(body.consultorio, body.fecha, body.franja, dur, body.id)) return { ok: false, error: "Conflicto de horario" };
      const consultorioNombre = NOMBRES_CONSULTORIOS[Number(body.consultorio)] || String(body.consultorio);
      const antes = reservaAuditFromRow(data[i]);
      sheet.getRange(i + 1, 2).setValue(consultorioNombre);
      sheet.getRange(i + 1, 4).setValue(body.fecha);
      sheet.getRange(i + 1, 5).setValue(body.franja);
      const despues = reservaAuditFromRow([data[i][0],consultorioNombre,data[i][2],body.fecha,body.franja,data[i][5],data[i][6],data[i][7],data[i][8],data[i][9]]);
      registrarAuditoria(user, "mover", "reserva", String(body.id), `movió reserva de ${resumenReservaAudit(antes)} → ${despues.consultorio} · ${despues.fecha} · ${despues.hora}`, antes, despues);
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
      if (!esRolOperativo(user) && user.id !== String(userId)) return { ok: false, error: "Sin permiso" };
      const targetUserId = String(userId);
      if (estaBloquado(body.consultorio, body.fecha, body.franja, duracion)) return { ok: false, error: "Franja bloqueada" };
      if (hayConflicto(body.consultorio, body.fecha, body.franja, duracion, null)) return { ok: false, error: "Conflicto de horario" };
      const consultorioNombre = NOMBRES_CONSULTORIOS[Number(body.consultorio)] || String(body.consultorio);
      const newId = "R_" + Date.now();
      const row = [newId, consultorioNombre, targetUserId, body.fecha, body.franja, duracion, nota || "", true, tipo || "normal", "confirmada"];
      sheet.appendRow(row);
      const despues = reservaAuditFromRow(row);
      registrarAuditoria(user, "crear", "reserva", newId, `copió reserva de ${resumenReservaAudit(despues)}`, reservaAuditFromRow(data[i]), despues);
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
      const antes = reservaAuditFromRow(data[i]);
      sheet.getRange(i + 1, 8).setValue(false);
      const despues = reservaAuditFromRow([data[i][0],data[i][1],data[i][2],data[i][3],data[i][4],data[i][5],data[i][6],false,data[i][8],data[i][9]]);
      registrarAuditoria(user, "eliminar", "reserva", String(body.id), `eliminó reserva de ${resumenReservaAudit(antes)}`, antes, despues);
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
    if (!puedeGestionarBloqueos(user)) return { ok: false, error: "Solo admin o asistente" };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ensureBloqueosSheet(ss);
    const consultorioVal = body.consultorio === "todos" ? "todos"
      : (NOMBRES_CONSULTORIOS[Number(body.consultorio)] || String(body.consultorio));
    const id = "B_" + Date.now();
    const row = [id, consultorioVal, body.franja, body.fecha, body.duracion, body.nota || "", true, body.repeticion || "ninguna"];
    sheet.appendRow(row);
    const despues = bloqueoAuditFromRow(row);
    registrarAuditoria(user, "crear", "bloqueo", id, `creó bloqueo · ${resumenBloqueoAudit(despues)}`, null, despues);
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
        const antes = bloqueoAuditFromRow(data[i]);
        sheet.getRange(i + 1, 7).setValue(false);
        const despues = bloqueoAuditFromRow([data[i][0],data[i][1],data[i][2],data[i][3],data[i][4],data[i][5],false,data[i][7]]);
        registrarAuditoria(user, "eliminar", "bloqueo", String(body.id), `eliminó bloqueo · ${resumenBloqueoAudit(antes)}`, antes, despues);
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
    if (!puedeGestionarBloqueos(user)) return { ok: false, error: "Solo admin o asistente" };
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_BLOQUEOS);
    if (!sheet) return { ok: false, error: "No existe hoja de bloqueos" };
    const data = sheet.getDataRange().getValues();
    const consultorioVal = body.consultorio === "todos" ? "todos"
      : (NOMBRES_CONSULTORIOS[Number(body.consultorio)] || String(body.consultorio));

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(body.id)) {
        const antes = bloqueoAuditFromRow(data[i]);
        sheet.getRange(i + 1, 2).setValue(consultorioVal);
        sheet.getRange(i + 1, 3).setValue(body.franja);
        sheet.getRange(i + 1, 4).setValue(body.fecha);
        const despues = bloqueoAuditFromRow([data[i][0],consultorioVal,body.franja,body.fecha,data[i][4],data[i][5],data[i][6],data[i][7]]);
        registrarAuditoria(user, "mover", "bloqueo", String(body.id), `movió bloqueo · ${resumenBloqueoAudit(antes)} → ${despues.consultorio} · ${despues.fecha} · ${despues.hora}`, antes, despues);
        invalidateAgendaCache();
        return { ok: true, id: String(body.id) };
      }
    }
    return { ok: false, error: "Bloqueo no encontrado" };
  });
}

// ── Diagnóstico de datos (solo lectura) ──────────────────────
function diagnosticarDatos(token) {
  const user = getUserFromToken(token);
  if (user.rol !== "admin") return { ok: false, error: "Solo admin" };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reservas = diagnosticarReservas(ss);
  const bloqueos = diagnosticarBloqueos(ss);
  return {
    ok: true,
    generadoEn: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss"),
    resumen: {
      reservasActivas: reservas.activas,
      reservasDuplicadas: reservas.totalDuplicados,
      reservasInvalidas: reservas.totalInvalidos,
      bloqueosActivos: bloqueos.activos,
      bloqueosDuplicados: bloqueos.totalDuplicados,
      bloqueosInvalidos: bloqueos.totalInvalidos
    },
    reservas,
    bloqueos
  };
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
