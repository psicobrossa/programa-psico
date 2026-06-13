// ============================================================
// PSICOGESTIÓN — Código.gs
// Google Apps Script · Backend API
// ============================================================

const SS = SpreadsheetApp.getActiveSpreadsheet();

// ── Hojas ────────────────────────────────────────────────────
const HOJAS = {
  pacientes:     "Pacientes",
  turnos:        "Turnos",
  cobros:        "Cobros",
  obrasSociales: "ObrasSociales",
  gastos:        "Gastos"
};

// ── Headers por hoja ─────────────────────────────────────────
const HEADERS = {
  pacientes:     ["ID","Nombre","DNI","Teléfono","Email","ObraSocial","ValorSesion","Estado"],
  turnos:        ["ID","Fecha","Hora","PacienteID","Nombre","Estado","Notas"],
  cobros:        ["ID","Fecha","PacienteID","Nombre","Importe","MedioPago","Estado"],
  obrasSociales: ["ID","Paciente","ObraSocial","SesionesMes","Facturado","Cobrado"],
  gastos:        ["ID","Fecha","Concepto","Categoría","Importe"]
};

// ============================================================
// INICIALIZACIÓN — crea las hojas y headers si no existen
// ============================================================
function inicializar() {
  Object.entries(HOJAS).forEach(([key, nombre]) => {
    let hoja = SS.getSheetByName(nombre);
    if (!hoja) {
      hoja = SS.insertSheet(nombre);
      hoja.appendRow(HEADERS[key]);
      hoja.getRange(1, 1, 1, HEADERS[key].length)
        .setBackground("#1a1a2e")
        .setFontColor("#E50055")
        .setFontWeight("bold");
      hoja.setFrozenRows(1);
    }
  });
  return { ok: true, mensaje: "Hojas inicializadas correctamente." };
}

// ============================================================
// ROUTER PRINCIPAL
// ============================================================
function doGet(e) {
  return manejarRequest(e);
}

function doPost(e) {
  return manejarRequest(e);
}

function manejarRequest(e) {
  const params = e.parameter || {};
  const body   = e.postData ? JSON.parse(e.postData.contents || "{}") : {};
  const accion = params.accion || body.accion;

  let resultado;
  try {
    switch (accion) {
      // ── Pacientes
      case "listarPacientes":     resultado = listarPacientes(); break;
      case "agregarPaciente":     resultado = agregarPaciente(body); break;
      case "editarPaciente":      resultado = editarPaciente(body); break;

      // ── Turnos
      case "listarTurnos":        resultado = listarTurnos(params); break;
      case "turnosHoy":           resultado = turnosHoy(); break;
      case "agregarTurno":        resultado = agregarTurno(body); break;
      case "actualizarEstadoTurno": resultado = actualizarEstadoTurno(body); break;

      // ── Cobros
      case "listarCobros":        resultado = listarCobros(params); break;
      case "registrarCobro":      resultado = registrarCobro(body); break;
      case "cobrosHoy":           resultado = cobrosHoy(); break;

      // ── Gastos
      case "listarGastos":        resultado = listarGastos(params); break;
      case "agregarGasto":        resultado = agregarGasto(body); break;

      // ── Dashboard
      case "dashboard":           resultado = dashboard(); break;

      // ── Setup
      case "editarTurno":         resultado = editarTurno(body); break;
      case "editarCobro":         resultado = editarCobro(body); break;
      case "editarGasto":         resultado = editarGasto(body); break;
      case "eliminarPaciente":    resultado = eliminarPaciente(body); break;
      case "eliminarTurno":       resultado = eliminarTurno(body); break;
      case "eliminarCobro":       resultado = eliminarCobro(body); break;
      case "eliminarGasto":       resultado = eliminarGasto(body); break;

      case "inicializar":         resultado = inicializar(); break;

      default:
        resultado = { error: "Acción no reconocida: " + accion };
    }
  } catch (err) {
    resultado = { error: err.message };
  }

  return ContentService
    .createTextOutput(JSON.stringify(resultado))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// PACIENTES
// ============================================================
function listarPacientes() {
  const datos = getData("pacientes");
  return datos;
}

function agregarPaciente(d) {
  const hoja = SS.getSheetByName(HOJAS.pacientes);
  const id   = generarID("pacientes");
  hoja.appendRow([
    id, d.nombre, d.dni, d.telefono, d.email,
    d.obraSocial || "Particular",
    d.valorSesion || 0,
    d.estado || "Activo"
  ]);
  return { ok: true, id };
}

function editarPaciente(d) {
  const hoja = SS.getSheetByName(HOJAS.pacientes);
  const fila = encontrarFila(hoja, d.id);
  if (!fila) return { error: "Paciente no encontrado" };
  hoja.getRange(fila, 2, 1, 7).setValues([[
    d.nombre, d.dni, d.telefono, d.email,
    d.obraSocial, d.valorSesion, d.estado
  ]]);
  return { ok: true };
}

// ============================================================
// TURNOS
// ============================================================
function listarTurnos(params) {
  const todos = getData("turnos");
  if (params.fecha) {
    return todos.filter(t => t.Fecha === params.fecha);
  }
  return todos;
}

function turnosHoy() {
  const hoy = formatFecha(new Date());
  return getData("turnos").filter(t => t.Fecha === hoy);
}

function agregarTurno(d) {
  const hoja = SS.getSheetByName(HOJAS.turnos);
  const id   = generarID("turnos");
  // Buscar nombre del paciente
  const paciente = getData("pacientes").find(p => p.ID === d.pacienteId);
  const nombre   = paciente ? paciente.Nombre : d.nombre || "";
  hoja.appendRow([
    id, d.fecha, d.hora, d.pacienteId,
    nombre, d.estado || "Programado", d.notas || ""
  ]);
  return { ok: true, id };
}

function actualizarEstadoTurno(d) {
  const hoja = SS.getSheetByName(HOJAS.turnos);
  const fila = encontrarFila(hoja, d.id);
  if (!fila) return { error: "Turno no encontrado" };
  hoja.getRange(fila, 6).setValue(d.estado);
  return { ok: true };
}

// ============================================================
// COBROS
// ============================================================
function listarCobros(params) {
  const todos = getData("cobros");
  if (params.mes) {
    return todos.filter(c => c.Fecha && c.Fecha.startsWith(params.mes));
  }
  return todos;
}

function cobrosHoy() {
  const hoy = formatFecha(new Date());
  return getData("cobros").filter(c => c.Fecha === hoy);
}

function registrarCobro(d) {
  const hoja = SS.getSheetByName(HOJAS.cobros);
  const id   = generarID("cobros");
  const paciente = getData("pacientes").find(p => p.ID === d.pacienteId);
  const nombre   = paciente ? paciente.Nombre : d.nombre || "";
  hoja.appendRow([
    id,
    d.fecha || formatFecha(new Date()),
    d.pacienteId,
    nombre,
    d.importe,
    d.medioPago || "Efectivo",
    d.estado || "Cobrado"
  ]);
  return { ok: true, id };
}

// ============================================================
// GASTOS
// ============================================================
function listarGastos(params) {
  const todos = getData("gastos");
  if (params.mes) {
    return todos.filter(g => g.Fecha && g.Fecha.startsWith(params.mes));
  }
  return todos;
}

function agregarGasto(d) {
  const hoja = SS.getSheetByName(HOJAS.gastos);
  const id   = generarID("gastos");
  hoja.appendRow([
    id,
    d.fecha || formatFecha(new Date()),
    d.concepto,
    d.categoria || "Otros",
    d.importe
  ]);
  return { ok: true, id };
}

// ============================================================
// DASHBOARD — indicadores del mes actual
// ============================================================
function dashboard() {
  const ahora     = new Date();
  const mesActual = formatFecha(ahora).substring(0, 7); // "YYYY-MM"

  const cobros  = getData("cobros").filter(c => c.Fecha && c.Fecha.startsWith(mesActual));
  const gastos  = getData("gastos").filter(g => g.Fecha && g.Fecha.startsWith(mesActual));
  const turnos  = getData("turnos").filter(t => t.Fecha && t.Fecha.startsWith(mesActual));
  const hoy     = turnosHoy();
  const cobHoy  = cobrosHoy();

  // Ingresos
  const ingresosTotal    = cobros.reduce((s, c) => s + Number(c.Importe || 0), 0);
  const ingresosPendient = cobros.filter(c => c.Estado === "Pendiente")
                                  .reduce((s, c) => s + Number(c.Importe || 0), 0);
  const ingresosCobrados = ingresosTotal - ingresosPendient;

  // Gastos
  const gastosTotal = gastos.reduce((s, g) => s + Number(g.Importe || 0), 0);

  // Resultado
  const resultado = ingresosCobrados - gastosTotal;

  // Turnos
  const asistidos  = turnos.filter(t => t.Estado === "Asistió").length;
  const ausentes   = turnos.filter(t => t.Estado === "Ausente").length;
  const totalT     = turnos.filter(t => ["Asistió","Ausente"].includes(t.Estado)).length;
  const ausentismo = totalT > 0 ? Math.round((ausentes / totalT) * 100) : 0;

  // Pacientes activos
  const pacActivos = getData("pacientes").filter(p => p.Estado === "Activo").length;

  // Distribución medios de pago
  const medios = {};
  cobros.forEach(c => {
    const m = c.MedioPago || "Otros";
    medios[m] = (medios[m] || 0) + Number(c.Importe || 0);
  });

  // Cobros hoy
  const cobrosHoyTotal = cobHoy.reduce((s, c) => s + Number(c.Importe || 0), 0);

  return {
    mes: mesActual,
    ingresos: {
      total:     ingresosTotal,
      cobrados:  ingresosCobrados,
      pendiente: ingresosPendient
    },
    gastos:    gastosTotal,
    resultado,
    pacientesActivos: pacActivos,
    turnos: {
      asistidos,
      ausentes,
      ausentismo,
      total: totalT
    },
    hoy: {
      turnos:  hoy.length,
      cobros:  cobrosHoyTotal,
      lista:   hoy.sort((a, b) => a.Hora > b.Hora ? 1 : -1)
    },
    mediosPago: medios
  };
}

// ============================================================
// UTILIDADES
// ============================================================
function getData(key) {
  const hoja   = SS.getSheetByName(HOJAS[key]);
  if (!hoja) return [];
  const valores = hoja.getDataRange().getValues();
  if (valores.length < 2) return [];
  const headers = valores[0];
  return valores.slice(1).map(fila => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = fila[i]; });
    return obj;
  });
}

function generarID(key) {
  const datos = getData(key);
  return datos.length + 1;
}

function encontrarFila(hoja, id) {
  const vals = hoja.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][0]) === String(id)) return i + 1;
  }
  return null;
}

function formatFecha(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

// ── Función de testing manual desde el editor ────────────────
function test_inicializar() {
  Logger.log(JSON.stringify(inicializar()));
}

function test_dashboard() {
  Logger.log(JSON.stringify(dashboard()));
}

// ============================================================
// EDICIÓN DE REGISTROS
// ============================================================
function editarTurno(d) {
  const hoja = SS.getSheetByName(HOJAS.turnos);
  const fila = encontrarFila(hoja, d.id);
  if (!fila) return { error: "Turno no encontrado" };
  const paciente = getData("pacientes").find(p => String(p.ID) === String(d.pacienteId));
  const nombre   = paciente ? paciente.Nombre : "";
  hoja.getRange(fila, 2, 1, 6).setValues([[
    d.fecha, d.hora, d.pacienteId, nombre, d.estado || "Programado", d.notas || ""
  ]]);
  return { ok: true };
}

function editarCobro(d) {
  const hoja = SS.getSheetByName(HOJAS.cobros);
  const fila = encontrarFila(hoja, d.id);
  if (!fila) return { error: "Cobro no encontrado" };
  const paciente = getData("pacientes").find(p => String(p.ID) === String(d.pacienteId));
  const nombre   = paciente ? paciente.Nombre : "";
  hoja.getRange(fila, 2, 1, 6).setValues([[
    d.fecha, d.pacienteId, nombre, d.importe, d.medioPago || "Efectivo", d.estado || "Cobrado"
  ]]);
  return { ok: true };
}

function editarGasto(d) {
  const hoja = SS.getSheetByName(HOJAS.gastos);
  const fila = encontrarFila(hoja, d.id);
  if (!fila) return { error: "Gasto no encontrado" };
  hoja.getRange(fila, 2, 1, 4).setValues([[
    d.fecha, d.concepto, d.categoria || "Otros", d.importe
  ]]);
  return { ok: true };
}

// ============================================================
// ELIMINACIÓN DE REGISTROS (marca la fila para borrar)
// ============================================================
function eliminarRegistro(hojaKey, id) {
  const hoja = SS.getSheetByName(HOJAS[hojaKey]);
  const fila = encontrarFila(hoja, id);
  if (!fila) return { error: "Registro no encontrado" };
  hoja.deleteRow(fila);
  return { ok: true };
}

function eliminarPaciente(d) {
  // Advertencia: no elimina cobros/turnos relacionados, solo el paciente
  return eliminarRegistro("pacientes", d.id);
}

function eliminarTurno(d)   { return eliminarRegistro("turnos", d.id); }
function eliminarCobro(d)   { return eliminarRegistro("cobros", d.id); }
function eliminarGasto(d)   { return eliminarRegistro("gastos", d.id); }
