// =========================
//  Código.gs — Gestión Convenios + Pólizas
//  Clínica Portoazul
// =========================

const ID_DOC        = "1Xhg4ewoAAX0KUGbUyF0izf7Hn6-TCkVnXDvfd81W5ww";
const NOMBRE_HOJA   = "EMPRESAS";
const HOJA_POLIZAS  = "POLIZAS";
const HOJA_DETALLES = "DETALLES_CONTRATOS";  // nueva hoja para detalles

// ──────────────────────────────────────────────────────────────
//  SERVIR HTML
// ──────────────────────────────────────────────────────────────
function doGet() {
  return HtmlService.createHtmlOutputFromFile("convenio")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setTitle("Gestión de Contratos — Clínica Portoazul");
}

// ══════════════════════════════════════════════════════════════
//  EMPRESAS / CONTRATOS
// ══════════════════════════════════════════════════════════════

function asegurarHojaYEncabezado() {
  const ss  = SpreadsheetApp.openById(ID_DOC);
  let hoja  = ss.getSheetByName(NOMBRE_HOJA);
  if (!hoja) hoja = ss.insertSheet(NOMBRE_HOJA);

  if (hoja.getLastRow() === 0) {
    hoja.appendRow([
      "NIT","Nombre","Segmento","Plan","N° Contrato",
      "Fecha Inicio","Tipo fin","Fecha fin","Fecha registro"
    ]);
  }
  return hoja;
}

/** Guardar empresa */
function guardarEmpresa(datos) {
  const hoja    = asegurarHojaYEncabezado();
  const tipoFin = datos.tipoFin || "Renovación automática";
  const fechaFin = tipoFin === "Fecha específica" ? datos.fechaFin : "Renovación automática";

  hoja.appendRow([
    datos.nit     || "",
    datos.nombre  || "",
    datos.segmento|| "",
    datos.plan    || "",
    datos.contrato|| "",
    datos.fecha   || "",
    tipoFin,
    fechaFin,
    new Date()
  ]);
  return "✅ Empresa registrada correctamente";
}

/** Obtener lista de empresas para el consolidado */
function obtenerEmpresas() {
  const ss   = SpreadsheetApp.openById(ID_DOC);
  const hoja = ss.getSheetByName(NOMBRE_HOJA);
  if (!hoja) return [];

  const datos = hoja.getDataRange().getValues();
  if (datos.length <= 1) return [];

  return datos.slice(1).map(row => ({
    nit:        row[0],
    nombre:     row[1],
    segmento:   row[2],
    plan:       row[3],
    contrato:   row[4],
    fechaInicio: formatDateCell(row[5]),
    fechaFin:   formatDateCell(row[7])
  }));
}

// ══════════════════════════════════════════════════════════════
//  DETALLE DE CONTRATO
// ══════════════════════════════════════════════════════════════

/** Asegura que la hoja DETALLES_CONTRATOS exista con encabezados */
function asegurarHojaDetalles() {
  const ss  = SpreadsheetApp.openById(ID_DOC);
  let hoja  = ss.getSheetByName(HOJA_DETALLES);
  if (!hoja) {
    hoja = ss.insertSheet(HOJA_DETALLES);
    hoja.appendRow([
      "NIT","N° Contrato","Objeto","Fecha Inicial","Fecha Final",
      "Valor","Observaciones","Nombre archivo","URL archivo","Última actualización"
    ]);
  }
  return hoja;
}

/**
 * Obtener detalle de un contrato por NIT + N° contrato.
 * Devuelve un objeto con todos los campos o {} si no existe aún.
 */
function obtenerDetalleContrato(nit, contrato) {
  const hoja = asegurarHojaDetalles();
  const datos = hoja.getDataRange().getValues();

  for (let i = 1; i < datos.length; i++) {
    const row = datos[i];
    if (String(row[0]).trim() === String(nit).trim() &&
        String(row[1]).trim() === String(contrato).trim()) {
      return {
        nit:          row[0],
        contrato:     row[1],
        objeto:       row[2],
        fechaInicial: formatDateCell(row[3]),
        fechaFinal:   formatDateCell(row[4]),
        valor:        row[5],
        observaciones:row[6],
        fileName:     row[7],
        fileUrl:      row[8]
      };
    }
  }

  // Si no hay detalle guardado aún, devolver vacío con NIT y contrato
  return { nit, contrato };
}

/**
 * Guardar (o actualizar) el detalle de un contrato.
 * Si ya existe la fila la sobreescribe; si no, la crea.
 */
function guardarDetalleContrato(detalle) {
  const hoja  = asegurarHojaDetalles();
  const datos = hoja.getDataRange().getValues();

  const nit      = String(detalle.nit     || "").trim();
  const contrato = String(detalle.contrato|| "").trim();

  // Buscar fila existente
  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][0]).trim() === nit &&
        String(datos[i][1]).trim() === contrato) {
      // Actualizar columnas 3-7 (objeto, fechas, valor, observaciones)
      hoja.getRange(i + 1, 3, 1, 5).setValues([[
        detalle.objeto       || "",
        detalle.fechaInicial || "",
        detalle.fechaFinal   || "",
        detalle.valor        || "",
        detalle.observaciones|| ""
      ]]);
      hoja.getRange(i + 1, 10).setValue(new Date()); // última actualización
      return "OK";
    }
  }

  // No existe → crear fila nueva
  hoja.appendRow([
    nit,
    contrato,
    detalle.objeto       || "",
    detalle.fechaInicial || "",
    detalle.fechaFinal   || "",
    detalle.valor        || "",
    detalle.observaciones|| "",
    "",   // nombre archivo (se llena con subirContratoFile)
    "",   // url archivo
    new Date()
  ]);
  return "OK";
}

/**
 * Subir un PDF al Drive y guardar el link en DETALLES_CONTRATOS.
 * Recibe: { nit, contrato, name, mimeType, base64 }
 */
function subirContratoFile(info) {
  const ss   = SpreadsheetApp.openById(ID_DOC);

  // Carpeta destino: "Contratos_CPA" en el Drive raíz (se crea si no existe)
  const carpetaNombre = "Contratos_CPA";
  let carpeta;
  const carpetas = DriveApp.getFoldersByName(carpetaNombre);
  if (carpetas.hasNext()) {
    carpeta = carpetas.next();
  } else {
    carpeta = DriveApp.createFolder(carpetaNombre);
  }

  // Decodificar y crear archivo
  const blob = Utilities.newBlob(
    Utilities.base64Decode(info.base64),
    info.mimeType,
    info.name
  );
  const archivo = carpeta.createFile(blob);
  archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const fileUrl  = "https://drive.google.com/file/d/" + archivo.getId() + "/preview";
  const fileName = archivo.getName();

  // Guardar URL en DETALLES_CONTRATOS
  const hoja  = asegurarHojaDetalles();
  const datos = hoja.getDataRange().getValues();
  const nit      = String(info.nit     || "").trim();
  const contrato = String(info.contrato|| "").trim();

  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][0]).trim() === nit &&
        String(datos[i][1]).trim() === contrato) {
      hoja.getRange(i + 1, 8).setValue(fileName);
      hoja.getRange(i + 1, 9).setValue(fileUrl);
      hoja.getRange(i + 1, 10).setValue(new Date());
      return { fileName, fileUrl };
    }
  }

  // Si no existe la fila de detalle, crearla con solo los datos del archivo
  hoja.appendRow([nit, contrato, "","","","","", fileName, fileUrl, new Date()]);
  return { fileName, fileUrl };
}

// ══════════════════════════════════════════════════════════════
//  PÓLIZAS
// ══════════════════════════════════════════════════════════════

/** Tipos únicos desde columna D */
function obtenerTiposPoliza() {
  const ss   = SpreadsheetApp.openById(ID_DOC);
  const hoja = ss.getSheetByName(HOJA_POLIZAS);
  if (!hoja) return [];

  const datos = hoja.getRange(2, 4, Math.max(hoja.getLastRow() - 1, 1)).getValues();
  return [...new Set(
    datos.flat()
      .filter(v => v && v !== "")
      .map(v => String(v).trim())
  )];
}

/** Obtener pólizas filtradas por tipo */
function obtenerPolizas(tipoSeleccionado) {
  const ss   = SpreadsheetApp.openById(ID_DOC);
  const hoja = ss.getSheetByName(HOJA_POLIZAS);
  if (!hoja) return [];

  const datos     = hoja.getDataRange().getValues();
  const resultado = [];

  for (let i = 1; i < datos.length; i++) {
    const fila       = datos[i];
    const nombre     = fila[1];  // B
    const plan       = fila[2];  // C
    const tipo       = fila[3];  // D
    const aseguradora= fila[4];  // E
    const inicio     = fila[5];  // F
    const fin        = fila[6];  // G
    const estado     = fila[9];  // J

    if (!tipoSeleccionado || String(tipo).trim() === String(tipoSeleccionado).trim()) {
      resultado.push({
        id:          i + 1,
        entidad:     nombre       || "",
        plan:        plan         || "",
        aseguradora: aseguradora  || "",
        inicio:      formatearFecha(inicio),
        fin:         formatearFecha(fin),
        inicioRaw:   formatDateInput(inicio),
        finRaw:      formatDateInput(fin),
        estado:      estado       || ""
      });
    }
  }
  return resultado;
}

/**
 * Actualizar fecha de inicio o fin de una póliza.
 * Recalcula el estado automáticamente en la columna J (10).
 */
function actualizarFechaPoliza(fila, campo, valor) {
  const ss   = SpreadsheetApp.openById(ID_DOC);
  const hoja = ss.getSheetByName(HOJA_POLIZAS);

  const columna = campo === "inicio" ? 6 : 7;  // F o G
  hoja.getRange(fila, columna).setValue(valor);

  // Recalcular estado según nueva fecha de fin
  const filaData  = hoja.getRange(fila, 1, 1, hoja.getLastColumn()).getValues()[0];
  const fechaFin  = filaData[6];  // columna G (índice 6)
  const nuevoEstado = calcularEstado(fechaFin);

  hoja.getRange(fila, 10).setValue(nuevoEstado);  // columna J

  return "OK";
}

/** Lógica de estado según fecha de fin */
function calcularEstado(fechaFin) {
  if (!fechaFin || fechaFin === "") return "N/A";

  const hoy     = new Date();
  const fin     = new Date(fechaFin);
  if (isNaN(fin)) return "N/A";

  hoy.setHours(0,0,0,0);
  fin.setHours(0,0,0,0);

  const diffDias = Math.ceil((fin - hoy) / (1000 * 60 * 60 * 24));

  if (diffDias < 0)   return "Vencida";
  if (diffDias <= 60) return "Proxima a vencer";
  return "Vigente";
}

/** Obtener una sola póliza por número de fila (para actualizar la UI) */
function obtenerPolizaPorFila(fila) {
  const ss   = SpreadsheetApp.openById(ID_DOC);
  const hoja = ss.getSheetByName(HOJA_POLIZAS);
  const datos = hoja.getRange(fila, 1, 1, hoja.getLastColumn()).getValues()[0];

  return {
    id:          fila,
    entidad:     datos[1] || "",
    plan:        datos[2] || "",
    aseguradora: datos[4] || "",
    inicio:      formatearFecha(datos[5]),
    fin:         formatearFecha(datos[6]),
    inicioRaw:   formatDateInput(datos[5]),
    finRaw:      formatDateInput(datos[6]),
    estado:      datos[9] || ""
  };
}

// ══════════════════════════════════════════════════════════════
//  ALERTAS POR EMAIL
// ══════════════════════════════════════════════════════════════

const CORREOS_ALERTA = [
  "thalia.eguis@auna.org",
  "camilo.martinez@auna.org",
  "erika.vengoechea@auna.org",
  "auxiliarconvenios.CPA@auna.org",
  "practicante.comercial@auna.org",
  "auxiliarcomercial.CPA@auna.org"
].join(", ");

/** Enviar alerta de pólizas próximas a vencer */
function enviarAlertaPolizas() {
  const lista = _obtenerPolizasPorEstado("proxima a vencer");
  if (lista.length === 0) return;

  const html = _construirTablaEmail(
    "⚠️ PÓLIZAS PRÓXIMAS A VENCER",
    `Buen día equipo,<br><br>
     Por favor revisar las siguientes pólizas próximas a vencer.
     Una vez gestionadas, actualizar la fecha de vigencia en el aplicativo.`,
    "#00B0CA",
    lista
  );

  MailApp.sendEmail({
    to:       CORREOS_ALERTA,
    subject:  "⚠️ Alerta: Pólizas próximas a vencer — Clínica Portoazul",
    htmlBody: html
  });
}

/** Enviar alerta de pólizas vencidas */
function enviarAlertaPolizasVencidas() {
  const lista = _obtenerPolizasPorEstado("vencida");
  if (lista.length === 0) return;

  const html = _construirTablaEmail(
    "🚨 PÓLIZAS VENCIDAS",
    `Buen día,<br><br>
     Las siguientes pólizas se encuentran <b>vencidas</b> y requieren gestión inmediata:`,
    "#8B0000",
    lista
  );

  MailApp.sendEmail({
    to:       CORREOS_ALERTA,
    subject:  "🚨 Alerta: Pólizas vencidas — Clínica Portoazul",
    htmlBody: html
  });
}

// ── helpers privados de alertas ─────────────────────────────

function _obtenerPolizasPorEstado(estadoBuscado) {
  const hoja  = SpreadsheetApp.openById(ID_DOC).getSheetByName(HOJA_POLIZAS);
  const datos = hoja.getDataRange().getValues();
  const lista = [];

  for (let i = 1; i < datos.length; i++) {
    const estado = String(datos[i][9] || "").toLowerCase().trim();
    if (estado === estadoBuscado) {
      const fechaFin = datos[i][6];
      lista.push({
        entidad:     datos[i][1],
        tipo:        datos[i][3],
        plan:        datos[i][2],
        aseguradora: datos[i][4],
        fechaFin:    fechaFin
          ? Utilities.formatDate(new Date(fechaFin), "America/Bogota", "dd/MM/yyyy")
          : ""
      });
    }
  }
  return lista;
}

function _construirTablaEmail(titulo, intro, colorHeader, lista) {
  let filas = lista.map(p => `
    <tr>
      <td style="padding:8px 10px;border-bottom:1px solid #eee">${p.entidad}</td>
      <td style="padding:8px 10px;border-bottom:1px solid #eee">${p.tipo}</td>
      <td style="padding:8px 10px;border-bottom:1px solid #eee">${p.plan}</td>
      <td style="padding:8px 10px;border-bottom:1px solid #eee">${p.aseguradora}</td>
      <td style="padding:8px 10px;border-bottom:1px solid #eee">${p.fechaFin}</td>
    </tr>`).join("");

  return `
    <div style="font-family:'Segoe UI',sans-serif;max-width:700px;margin:0 auto;">
      <div style="background:${colorHeader};padding:20px 24px;border-radius:10px 10px 0 0;">
        <h2 style="color:#fff;margin:0;font-size:1.2rem">${titulo}</h2>
      </div>
      <div style="background:#fff;padding:20px 24px;border:1px solid #eee;">
        <p style="color:#333;line-height:1.6">${intro}</p>
        <p style="color:#333">Total: <b>${lista.length}</b></p>
        <table style="width:100%;border-collapse:collapse;margin-top:12px;font-size:.88rem">
          <thead>
            <tr style="background:${colorHeader};color:#fff">
              <th style="padding:10px;text-align:left">Entidad</th>
              <th style="padding:10px;text-align:left">Tipo</th>
              <th style="padding:10px;text-align:left">Plan</th>
              <th style="padding:10px;text-align:left">Aseguradora</th>
              <th style="padding:10px;text-align:left">Fecha Fin</th>
            </tr>
          </thead>
          <tbody>${filas}</tbody>
        </table>
      </div>
      <div style="background:#f4f8fb;padding:12px 24px;border-radius:0 0 10px 10px;font-size:.78rem;color:#888">
        Clínica Portoazul — Sistema de Gestión de Convenios
      </div>
    </div>`;
}

// ══════════════════════════════════════════════════════════════
//  UTILIDADES
// ══════════════════════════════════════════════════════════════

function formatDateCell(value) {
  if (!value && value !== 0) return "";
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return String(value);
}

function formatearFecha(valor) {
  if (valor === "" || valor === null || valor === undefined) return "";
  if (typeof valor === "string") return valor;
  if (Object.prototype.toString.call(valor) === "[object Date]" && !isNaN(valor)) {
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), "dd/MM/yyyy");
  }
  return valor;
}

function formatDateInput(fecha) {
  if (Object.prototype.toString.call(fecha) === "[object Date]" && !isNaN(fecha)) {
    return Utilities.formatDate(fecha, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return "";
}
function myFunction() {
  
}
