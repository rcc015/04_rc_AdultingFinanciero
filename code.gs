/***** CONFIGURACIÓN *****/
const ALLOWED = new Set([
  "rcc015@gmail.com",
  "afibn3255@gmail.com"
]);

const SHEET_NAME = "Movimientos";
const EMAILS_NOTIFY = ["rcc015@gmail.com", "afibn3255@gmail.com"];

/***** WEB APP ENTRY *****/
function doGet() {
  const email = getUserEmail_();
  if (!isAllowed_(email)) return denied_(email);
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("Wallet Pro Final");
}

/***** 1. GESTIÓN DE DATOS (Autocomplete) *****/
function uiGetDataForAutocomplete() {
  const email = getUserEmail_();
  if (!isAllowed_(email)) throw new Error("Acceso denegado");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shCat = ss.getSheetByName("Categorias");
  let cats = [];
  if (shCat && shCat.getLastRow() >= 1) {
    const vals = shCat.getRange(1, 1, shCat.getLastRow(), 1).getValues().flat();
    cats = [...new Set(vals.map(v => String(v||"").trim()).filter(v => v))];
  }

  const shMov = ss.getSheetByName(SHEET_NAME);
  let descs = [];
  if (shMov && shMov.getLastRow() >= 2) {
    const startRow = Math.max(2, shMov.getLastRow() - 500);
    const numRows = shMov.getLastRow() - startRow + 1;
    const data = shMov.getRange(startRow, 5, numRows, 1).getValues().flat(); 
    descs = [...new Set(data.map(v => String(v||"").trim()).filter(v => v.length > 2))].sort();
  }

  return { ok: true, categories: cats, descriptions: descs };
}

/***** 2. GESTIÓN DE MOVIMIENTOS *****/
function uiAdd(data) {
  const email = getUserEmail_();
  if (!isAllowed_(email)) throw new Error("Acceso denegado");

  const {fecha, tipo, categoria, descripcion, monto} = validateData_(data);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME);
  
  const ts = new Date();
  const mes = fecha.slice(0, 7); 

  sh.appendRow([fecha, tipo, categoria, monto, descripcion, email, ts, "'" + mes]);

  // Checar presupuesto con semáforo (80% y 100%)
  if (tipo === 'Gasto') checkBudgetLimit(mes, categoria);

  return uiSummary(mes);
}

function uiEditMovement(rowId, data) {
  const email = getUserEmail_();
  if (!isAllowed_(email)) throw new Error("Acceso denegado");
  
  const {fecha, tipo, categoria, descripcion, monto} = validateData_(data);
  const r = parseInt(rowId, 10);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME);
  
  if (!sh || r < 2 || r > sh.getLastRow()) throw new Error("Movimiento no encontrado");

  const mes = fecha.slice(0, 7);
  const ts = new Date();
  
  sh.getRange(r, 1, 1, 8).setValues([[fecha, tipo, categoria, monto, descripcion, email, ts, "'" + mes]]);
  
  // También checamos límite al editar, por si el cambio dispara la alerta
  if (tipo === 'Gasto') checkBudgetLimit(mes, categoria);

  return { ok: true };
}

function uiDeleteMovement(rowId, ym) {
  const email = getUserEmail_();
  if (!isAllowed_(email)) throw new Error("Acceso denegado");
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME);
  const r = parseInt(rowId, 10);
  
  if (sh && r > 1 && r <= sh.getLastRow()) sh.deleteRow(r);
  return { ok: true };
}

/***** 3. AUTOMATIZACIÓN (Recurrentes) *****/
function uiLoadRecurring(ym) {
  const email = getUserEmail_();
  if (!isAllowed_(email)) throw new Error("Acceso denegado");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let shRec = ss.getSheetByName("Recurrentes");

  if (!shRec) {
    shRec = ss.insertSheet("Recurrentes");
    shRec.appendRow(["Categoria", "Descripcion", "Monto", "Usuario (Email)"]);
    return { ok: false, msg: "Hoja 'Recurrentes' creada. Llénala primero." };
  }

  const lastRow = shRec.getLastRow();
  if (lastRow < 2) return { ok: false, msg: "La hoja 'Recurrentes' está vacía." };

  // 1. Obtener gastos YA registrados este mes para evitar duplicados
  const shMov = ss.getSheetByName(SHEET_NAME);
  const existingMovs = uiFilteredMovements({ ym: ym, limit: 1000 }).items;
  // Creamos una lista de "firmas" (Ej: "Renta Departamento-5000")
  const existingSignatures = new Set(existingMovs.map(m => `${m.descripcion}-${m.monto}`));

  const templates = shRec.getRange(2, 1, lastRow - 1, 4).getValues();
  const ts = new Date();
  const fechaStr = ym + "-01"; 
  
  let count = 0;
  let total = 0;

  templates.forEach(row => {
    const cat = String(row[0]);
    const desc = String(row[1]);
    const monto = Number(row[2]);
    const user = String(row[3] || email).trim();
    
    // Firma del gasto que queremos agregar
    const signature = `${desc}-${monto}`;

    // Solo agregamos si NO existe ya en este mes
    if (cat && monto > 0 && !existingSignatures.has(signature)) {
      shMov.appendRow([fechaStr, "Gasto", cat, monto, desc, user, ts, "'" + ym]);
      count++;
      total += monto;
    }
  });

  if (count === 0) {
    return { ok: true, count: 0, total: 0, msg: "Parece que ya cargaste los recurrentes de este mes." };
  }

  return { ok: true, count: count, total: total };
}

/***** 4. REPORTES *****/
function uiSummary(ym) {
  const email = getUserEmail_();
  if (!isAllowed_(email)) throw new Error("Acceso denegado");
  const month = normalizeDateStr_(ym);
  const data = uiFilteredMovements({ ym: month, limit: 1000 }).items;
  let ing = 0, gas = 0;
  data.forEach(d => { if(d.tipo === 'Ingreso') ing += d.monto; else gas += d.monto; });
  return { ok: true, month, ingresos: ing, gastos: gas, balance: ing - gas };
}

function uiFilteredMovements(opts) {
  const ym = normalizeDateStr_(opts.ym);
  const limit = Number(opts.limit) || 100;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh || sh.getLastRow() < 2) return { ok: true, items: [] };

  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues();
  const items = [];

  for (let i = 0; i < data.length; i++) {
    const r = data[i];
    let rowMonth = r[7] ? normalizeDateStr_(r[7]) : normalizeDateStr_(r[0]);
    if (rowMonth === ym) {
      items.push({
        rowId: i + 2,
        fecha: formatDateDisplay_(r[0]),
        tipo: String(r[1]),
        categoria: String(r[2]),
        monto: Number(r[3]),
        descripcion: String(r[4]),
        usuario: String(r[5])
      });
    }
  }
  items.sort((a, b) => {
    if (b.fecha > a.fecha) return 1;
    if (b.fecha < a.fecha) return -1;
    return b.rowId - a.rowId; 
  });
  return { ok: true, items: items.slice(0, limit) };
}

/***** 5. PRESUPUESTOS Y ALERTAS (Aquí está la magia del 80%) *****/
function uiGetBudgets(ym) {
  const month = normalizeDateStr_(ym); 
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName("Presupuestos");
  if (!sh) { sh = ss.insertSheet("Presupuestos"); sh.appendRow(["Año-Mes", "Categoria", "Presupuesto"]); return { ok: true, month, items: [] }; }
  
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { ok: true, month, items: [] };
  const values = sh.getRange(2, 1, lastRow - 1, 3).getValues();
  const items = [];
  for (let i = 0; i < values.length; i++) {
    if (normalizeDateStr_(values[i][0]) === month) {
      items.push({ ym: month, categoria: String(values[i][1]).trim(), presupuesto: Number(values[i][2]) || 0 });
    }
  }
  return { ok: true, month, items };
}

function uiSetBudget(ym, categoria, monto) {
  const email = getUserEmail_();
  if (!isAllowed_(email)) throw new Error("Acceso denegado");
  const month = normalizeDateStr_(ym);
  const cat = String(categoria || "").trim();
  const val = Number(monto);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName("Presupuestos");
  if(!sh) { sh = ss.insertSheet("Presupuestos"); sh.appendRow(["Año-Mes", "Categoria", "Presupuesto"]); }
  
  const lastRow = sh.getLastRow();
  const monthText = "'" + month; 
  let action = "created";
  let foundRow = -1;
  if (lastRow >= 2) {
    const data = sh.getRange(2, 1, lastRow - 1, 2).getValues();
    for(let i=0; i<data.length; i++) {
      if (normalizeDateStr_(data[i][0]) === month && String(data[i][1]) === cat) { foundRow = i + 2; break; }
    }
  }
  if (foundRow > 0) { sh.getRange(foundRow, 3).setValue(val); action = "overwritten"; }
  else { sh.appendRow([monthText, cat, val]); }

  if (cat !== '_SAVINGS_TARGET_') {
    EMAILS_NOTIFY.forEach(e => MailApp.sendEmail(e, `🎯 Presupuesto: ${cat}`, `Nuevo límite para ${cat}: $${val} (Mes: ${month})`));
  }
  return { ok: true, action: action };
}

function uiBudgetVsActual(ym) {
  const month = normalizeDateStr_(ym);
  const budgets = uiGetBudgets(month).items || [];
  const movs = uiFilteredMovements({ ym: month, limit: 2000 }).items || [];
  const realMap = {};
  movs.forEach(m => { if(m.tipo === 'Gasto') realMap[m.categoria] = (realMap[m.categoria] || 0) + m.monto; });
  return { ok: true, month, items: budgets.map(b => ({ categoria: b.categoria, presupuesto: b.presupuesto, gasto: realMap[b.categoria] || 0 })) };
}

// --- LOGICA DE SEMÁFORO (80% y 100%) ---
function checkBudgetLimit(ym, categoria) {
  try {
    const budgets = uiGetBudgets(ym).items;
    const meta = budgets.find(b => b.categoria === categoria);
    if (meta && meta.presupuesto > 0) {
      // Recalcular gasto total (incluyendo el recién guardado)
      const movs = uiFilteredMovements({ ym: ym, limit: 2000 }).items;
      let gastado = 0;
      movs.forEach(m => { if (m.tipo === 'Gasto' && m.categoria === categoria) gastado += m.monto; });
      
      const porcentaje = gastado / meta.presupuesto;

      // 1. Alerta ROJA (> 100%)
      if (gastado > meta.presupuesto) {
        const asunto = `🚨 CRÍTICO: ${categoria} al ${(porcentaje*100).toFixed(0)}%`;
        const cuerpo = `Has superado el límite de ${categoria}.\n\n` +
                       `Presupuesto: $${meta.presupuesto}\n` +
                       `Gastado: $${gastado}\n` +
                       `Exceso: $${gastado - meta.presupuesto}`;
        EMAILS_NOTIFY.forEach(e => MailApp.sendEmail(e, asunto, cuerpo));
      } 
      // 2. Alerta AMARILLA (>= 80% y < 100%)
      else if (porcentaje >= 0.80) {
        const asunto = `⚠️ ALERTA: ${categoria} al ${(porcentaje*100).toFixed(0)}%`;
        const cuerpo = `Estás por acabarte el presupuesto de ${categoria}.\n\n` +
                       `Presupuesto: $${meta.presupuesto}\n` +
                       `Gastado: $${gastado}\n` +
                       `Te quedan: $${meta.presupuesto - gastado}`;
        EMAILS_NOTIFY.forEach(e => MailApp.sendEmail(e, asunto, cuerpo));
      }
    }
  } catch(e) {}
}

/***** 6. TRIGGERS (Automáticos) *****/
function enviarReporteSemanal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME);
  const hoy = new Date();
  const hace7dias = new Date(hoy.getTime() - 7 * 24 * 60 * 60 * 1000);
  const data = sh.getRange(2, 1, sh.getLastRow()-1, 6).getValues();
  let totalRod = 0, totalBri = 0;
  data.forEach(r => {
    const fecha = new Date(r[0]);
    const tipo = r[1];
    const monto = r[3];
    const user = String(r[5]).toLowerCase();
    if (fecha >= hace7dias && tipo === 'Gasto') {
      if (user.includes("rcc015")) totalRod += monto; else totalBri += monto;
    }
  });
  const total = totalRod + totalBri;
  const cuerpo = `Resumen semanal:\nRodrigo: $${Math.floor(totalRod)}\nBrian: $${Math.floor(totalBri)}\nTOTAL: $${Math.floor(total)}`;
  EMAILS_NOTIFY.forEach(e => MailApp.sendEmail(e, `📊 Semanal: $${Math.floor(total)}`, cuerpo));
}

function cierreMensual() {
  const hoy = new Date();
  const manana = new Date(hoy);
  manana.setDate(hoy.getDate() + 1);
  if (manana.getDate() !== 1) return; 

  const ym = Utilities.formatDate(hoy, Session.getScriptTimeZone(), "yyyy-MM");
  const data = uiSummary(ym);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let shHist = ss.getSheetByName("Historial_Cierres");
  if (!shHist) {
    shHist = ss.insertSheet("Historial_Cierres");
    shHist.appendRow(["Mes", "Ingresos", "Gastos", "Balance", "Fecha"]);
  }
  shHist.appendRow([ym, data.ingresos, data.gastos, data.balance, new Date()]);
  const cuerpo = `Cierre ${ym}:\nIngresos: $${data.ingresos}\nGastos: $${data.gastos}\nBalance: $${data.balance}`;
  EMAILS_NOTIFY.forEach(e => MailApp.sendEmail(e, `📅 Cierre ${ym}`, cuerpo));
}

/***** HELPERS *****/
function validateData_(data) {
  const fecha = (data.fecha || "").trim();
  const tipo = (data.tipo || "").trim();
  const categoria = (data.categoria || "").trim();
  const descripcion = (data.descripcion || "").trim();
  const monto = Number(data.monto);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(fecha) || !categoria || !descripcion || !Number.isFinite(monto) || monto <= 0) throw new Error("Datos inválidos");
  return {fecha, tipo, categoria, descripcion, monto};
}
function normalizeDateStr_(val) {
  if (!val) return "";
  if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM");
  const s = String(val).trim();
  return (s.length >= 7) ? s.slice(0, 7) : s;
}
function formatDateDisplay_(val) {
  if (!val) return "";
  const d = new Date(val);
  return isNaN(d.getTime()) ? String(val) : d.toISOString().slice(0, 10);
}
function getUserEmail_() { return (Session.getActiveUser().getEmail() || "").toLowerCase(); }
function isAllowed_(email) { return email && ALLOWED.has(email); }
function denied_(email) { return HtmlService.createHtmlOutput(`Acceso denegado. ${email}`); }

// ============================================================
//  INTEGRACIÓN BBVA — Agrega esto a tu Code.gs existente
// ============================================================

/***** BBVA: PARSEAR CORREOS Y REGISTRAR MOVIMIENTOS *****/

/**
 * Parsea todos los correos de BBVA no procesados y los registra
 * en la hoja de Movimientos. Llama esto desde un trigger de tiempo.
 */
function parsearCorreosBBVA() {
  const email = getUserEmail_();
  if (!isAllowed_(email)) throw new Error("Acceso denegado");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME);

  // Buscar correos de BBVA no marcados como procesados
  const threads = GmailApp.search('from:clientes@bbva.mx -label:bbva-procesado', 0, 50);

  if (threads.length === 0) return { ok: true, count: 0 };

  // Crear label "bbva-procesado" si no existe
  let label = GmailApp.getUserLabelByName('bbva-procesado');
  if (!label) label = GmailApp.createLabel('bbva-procesado');

  let count = 0;
  const ts = new Date();

  threads.forEach(thread => {
    const messages = thread.getMessages();
    messages.forEach(msg => {
      const asunto = msg.getSubject() || "";
      const body = msg.getBody() || "";
      const fechaMsg = msg.getDate();

      let movimiento = null;

      // --- TIPO 1: Transferencia exitosa ---
      if (asunto.toLowerCase().includes("transferencia") &&
          asunto.toLowerCase().includes("exitosa")) {
        movimiento = _parsearTransferencia(body, fechaMsg);
      }

      // --- TIPO 2: Retiro de apartado ---
      else if (asunto.toLowerCase().includes("retiraste dinero") ||
               asunto.toLowerCase().includes("apartado")) {
        movimiento = _parsearApartado(body, fechaMsg, "retiro");
      }

      // --- TIPO 3: Creación de apartado ---
      else if (asunto.toLowerCase().includes("realizaste un apartado")) {
        movimiento = _parsearApartado(body, fechaMsg, "creacion");
      }

      // Registrar si se parseó correctamente y no es duplicado
      if (movimiento && !_esDuplicado(sh, movimiento)) {
        const mes = movimiento.fecha.slice(0, 7);
        sh.appendRow([
          movimiento.fecha,
          movimiento.tipo,
          movimiento.categoria,
          movimiento.monto,
          movimiento.descripcion,
          email,
          ts,
          "'" + mes,
          "BBVA_AUTO"  // Columna 9: marcador de origen
        ]);
        count++;
      }
    });

    // Marcar thread como procesado
    thread.addLabel(label);
  });

  return { ok: true, count: count };
}

/**
 * Parsea un correo de transferencia BBVA
 */
function _parsearTransferencia(body, fechaMsg) {
  try {
    // Extraer importe: "Importe: $ 1,600.00" o "$ 1,600.00"
    const montoMatch = body.match(/Importe:\s*\$?\s*([\d,]+\.?\d*)/i);
    if (!montoMatch) return null;
    const monto = parseFloat(montoMatch[1].replace(/,/g, ''));
    if (!monto || monto <= 0) return null;

    // Extraer beneficiario
    const benMatch = body.match(/Beneficiario:\s*([A-ZÁÉÍÓÚÑ\s]+?)(?:<br|\\n|\n|Cuenta)/i);
    const beneficiario = benMatch ? benMatch[1].trim() : "Beneficiario BBVA";

    // Fecha del mensaje
    const fecha = Utilities.formatDate(fechaMsg, Session.getScriptTimeZone(), "yyyy-MM-dd");

    return {
      fecha: fecha,
      tipo: "Gasto",
      categoria: "BBVA Transferencia",
      monto: monto,
      descripcion: "→ " + beneficiario
    };
  } catch(e) {
    Logger.log("Error parseando transferencia: " + e);
    return null;
  }
}

/**
 * Parsea un correo de apartado BBVA
 */
function _parsearApartado(body, fechaMsg, accion) {
  try {
    const montoMatch = body.match(/\$\s*([\d,]+\.?\d*)/);
    if (!montoMatch) return null;
    const monto = parseFloat(montoMatch[1].replace(/,/g, ''));
    if (!monto || monto <= 0) return null;

    const fecha = Utilities.formatDate(fechaMsg, Session.getScriptTimeZone(), "yyyy-MM-dd");

    if (accion === "retiro") {
      return {
        fecha: fecha,
        tipo: "Ingreso",
        categoria: "BBVA Apartado",
        monto: monto,
        descripcion: "Retiro de apartado BBVA"
      };
    } else {
      return {
        fecha: fecha,
        tipo: "Gasto",
        categoria: "BBVA Apartado",
        monto: monto,
        descripcion: "Creación de apartado BBVA"
      };
    }
  } catch(e) {
    Logger.log("Error parseando apartado: " + e);
    return null;
  }
}

/**
 * Evita duplicados: verifica si ya existe un movimiento con misma
 * fecha, categoría BBVA y monto en la hoja.
 */
function _esDuplicado(sh, mov) {
  if (!sh || sh.getLastRow() < 2) return false;
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
  return data.some(r => {
    const fecha = Utilities.formatDate(new Date(r[0]), Session.getScriptTimeZone(), "yyyy-MM-dd");
    return fecha === mov.fecha &&
           String(r[2]).includes("BBVA") &&
           Number(r[3]) === mov.monto;
  });
}

/***** BBVA: OBTENER MOVIMIENTOS PARA DASHBOARD *****/

/**
 * Devuelve solo los movimientos con origen BBVA del mes indicado.
 * Usado por la pestaña BBVA en el frontend.
 */
function uiGetBBVAMovements(ym) {
  const email = getUserEmail_();
  if (!isAllowed_(email)) throw new Error("Acceso denegado");

  const month = normalizeDateStr_(ym);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh || sh.getLastRow() < 2) return { ok: true, items: [] };

  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues();
  const items = [];

  for (let i = 0; i < data.length; i++) {
    const r = data[i];
    const cat = String(r[2] || "");
    const origen = String(r[8] || "");
    const esBBVA = cat.includes("BBVA") || origen === "BBVA_AUTO";

    if (!esBBVA) continue;

    let rowMonth = r[7] ? normalizeDateStr_(r[7]) : normalizeDateStr_(r[0]);
    if (rowMonth !== month) continue;

    items.push({
      rowId: i + 2,
      fecha: formatDateDisplay_(r[0]),
      tipo: String(r[1]),
      categoria: String(r[2]),
      monto: Number(r[3]),
      descripcion: String(r[4]),
      usuario: String(r[5])
    });
  }

  items.sort((a, b) => b.fecha > a.fecha ? 1 : -1);

  // Calcular totales por beneficiario/descripción
  const totales = {};
  items.forEach(it => {
    if (it.tipo === "Gasto") {
      const key = it.descripcion;
      totales[key] = (totales[key] || 0) + it.monto;
    }
  });

  const totalGastos = items.filter(i => i.tipo === "Gasto").reduce((s, i) => s + i.monto, 0);
  const totalIngresos = items.filter(i => i.tipo === "Ingreso").reduce((s, i) => s + i.monto, 0);

  return {
    ok: true,
    items: items,
    totalesPorBeneficiario: totales,
    totalGastos: totalGastos,
    totalIngresos: totalIngresos
  };
}

/***** TRIGGER: Configurar revisión automática cada 15 min *****/

/**
 * Ejecuta esto UNA VEZ manualmente desde el editor de Apps Script
 * para activar el monitoreo automático de correos BBVA.
 * 
 * Ve a: Ejecutar > configurarTriggerBBVA
 */
function configurarTriggerBBVA() {
  // Eliminar triggers previos del mismo nombre para evitar duplicados
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'parsearCorreosBBVA') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Crear trigger cada 15 minutos
  ScriptApp.newTrigger('parsearCorreosBBVA')
    .timeBased()
    .everyMinutes(15)
    .create();

  Logger.log("✅ Trigger BBVA configurado: cada 15 minutos");
}

/**
 * Ejecutar manualmente para procesar correos BBVA históricos
 * (útil para importar correos pasados que no tenían el label)
 */
function importarBBVAHistorico() {
  const email = getUserEmail_();
  if (!isAllowed_(email)) throw new Error("Acceso denegado");

  // Buscar TODOS los correos BBVA (con o sin label procesado)
  const threads = GmailApp.search('from:clientes@bbva.mx', 0, 200);
  let label = GmailApp.getUserLabelByName('bbva-procesado');
  if (!label) label = GmailApp.createLabel('bbva-procesado');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME);
  const ts = new Date();
  let count = 0;
  const email_ = email;

  threads.forEach(thread => {
    thread.getMessages().forEach(msg => {
      const asunto = msg.getSubject() || "";
      const body = msg.getBody() || "";
      const fechaMsg = msg.getDate();

      let movimiento = null;

      if (asunto.toLowerCase().includes("transferencia") && asunto.toLowerCase().includes("exitosa")) {
        movimiento = _parsearTransferencia(body, fechaMsg);
      } else if (asunto.toLowerCase().includes("retiraste dinero") || asunto.toLowerCase().includes("apartado")) {
        movimiento = _parsearApartado(body, fechaMsg, "retiro");
      } else if (asunto.toLowerCase().includes("realizaste un apartado")) {
        movimiento = _parsearApartado(body, fechaMsg, "creacion");
      }

      if (movimiento && !_esDuplicado(sh, movimiento)) {
        const mes = movimiento.fecha.slice(0, 7);
        sh.appendRow([movimiento.fecha, movimiento.tipo, movimiento.categoria, movimiento.monto, movimiento.descripcion, email_, ts, "'" + mes, "BBVA_AUTO"]);
        count++;
      }
    });
    thread.addLabel(label);
  });

  return { ok: true, count: count, msg: count + " movimientos BBVA importados" };
}