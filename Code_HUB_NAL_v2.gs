// ═══════════════════════════════════════════════════════════════════════════
// HUB OPERATIVO NAL — Google Apps Script v2
// Sheet ID: 1Ozxdr-V67f-IsaBBb7Jqp9LETwRrFm0widhoqt-2axs
// Hojas reales confirmadas del Sheet
// ═══════════════════════════════════════════════════════════════════════════
// INSTRUCCIONES:
// 1. En tu Google Sheet → Extensiones → Apps Script
// 2. Borra todo lo que haya y pega este código completo
// 3. Guardar (Ctrl+S)
// 4. Implementar → Nueva implementación
//    - Tipo: Aplicación web
//    - Ejecutar como: Yo (tu cuenta)
//    - Acceso: Cualquier persona
// 5. Autorizar y copiar la URL que te da
// 6. En el HUB HTML busca: const SHEETS_URL = ''
//    y pega la URL entre las comillas
// ═══════════════════════════════════════════════════════════════════════════

// ── NOMBRES EXACTOS DE TUS HOJAS ──────────────────────────────────────────
const H_ESTATUS    = 'Estatus_Diario';
const H_OPERADORES = 'Control_Operadores';
const H_CAJAS      = 'Control_Cajas';
const H_INCIDENCIAS= 'Incidencias';
const H_RENDIMIENTO= 'Rendimientos';
const H_MTTO       = 'Solicitudes_Mtto';
const H_DATA       = 'Data';
const H_MOVIMIENTOS= 'Movimientos_RH';
const H_WHATSAPP   = 'Whats App';
const H_URGENCIAS  = 'Urgencias';

// ── MAPEO DE MOTIVO → CODE (basado en tu Sheet real) ─────────────────────
function motivoACode(motivo) {
  const m = (motivo || '').trim().toUpperCase();
  if (m.startsWith('TRN')) return 'TRN';
  if (m.startsWith('VTA')) return 'VTA';
  if (m.startsWith('DCO')) return 'DCO';
  if (m.startsWith('DSO')) return 'DSO';
  if (m.startsWith('LIB')) return 'LIB';
  if (m.startsWith('SO'))  return 'SO';
  if (m.startsWith('CP'))  return 'CP';
  if (m.startsWith('SG'))  return 'SG';
  if (m.startsWith('RM'))  return 'RM';
  if (m.startsWith('PER')) return 'PER';
  if (m.startsWith('SGR')) return 'SGR';
  return 'DCO';
}

// ── ENTRY POINT ───────────────────────────────────────────────────────────
function doGet(e) {
  // Permitir CORS para que el HUB pueda llamar desde cualquier origen
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  const action = (e && e.parameter && e.parameter.action) || 'all';
  const result = { ok: false };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (action === 'all' || action === 'estatus') {
      result.units = getEstatus(ss);
    }
    if (action === 'all' || action === 'cajas') {
      result.cajas = getCajas(ss);
    }
    if (action === 'all' || action === 'operadores') {
      result.operadores = getOperadores(ss);
    }
    if (action === 'all' || action === 'urgencias') {
      result.urgencias = getUrgencias(ss);
    }
    if (action === 'all' || action === 'whatsapp') {
      result.whatsapp = getWhatsApp(ss);
    }

    result.fecha = Utilities.formatDate(new Date(), 'America/Monterrey', 'yyyy-MM-dd');
    result.hora  = Utilities.formatDate(new Date(), 'America/Monterrey', 'HH:mm');
    result.ok    = true;

  } catch (err) {
    result.error = err.toString();
    result.stack = err.stack || '';
  }

  output.setContent(JSON.stringify(result));
  return output;
}

// ── ESTATUS DIARIO ────────────────────────────────────────────────────────
// Columnas: Fecha | Unidad | Operador | UnidadDeNegocio | Estatus | Motivo |
//           FechaCompromisoUnidad | FechaCompromisoOperador | NombreRuta |
//           Monto | Coordinador | Comentarios
function getEstatus(ss) {
  const sh = ss.getSheetByName(H_ESTATUS);
  if (!sh) return [];

  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];

  // Índices exactos por nombre de columna
  const heads = data[0].map(h => String(h).trim());
  const idx = {
    fecha:    heads.indexOf('Fecha'),
    unidad:   heads.indexOf('Unidad'),
    operador: heads.indexOf('Operador'),
    negocio:  heads.indexOf('UnidadDeNegocio'),
    estatus:  heads.indexOf('Estatus'),
    motivo:   heads.indexOf('Motivo'),
    fechaCompU: heads.indexOf('FechaCompromisoUnidad'),
    fechaCompO: heads.indexOf('FechaCompromisoOperador'),
    ruta:     heads.indexOf('NombreRuta'),
    monto:    heads.indexOf('Monto'),
    coord:    heads.indexOf('Coordinador'),
    obs:      heads.indexOf('Comentarios'),
  };

  const units = [];
  // Obtener la fecha más reciente del archivo (puede tener histórico)
  let maxFecha = null;
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const uid = String(row[idx.unidad] || '').trim();
    if (!uid) continue;
    const neg = String(row[idx.negocio] || '').toLowerCase();
    if (neg && !neg.includes('nacional') && !neg.includes('nal')) continue;
    const f = row[idx.fecha];
    if (f instanceof Date && (!maxFecha || f > maxFecha)) maxFecha = f;
  }

  // Filtrar solo registros de la fecha más reciente
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const uid = String(row[idx.unidad] || '').trim();
    if (!uid || uid.length < 3) continue;

    // Solo NAL
    const neg = String(row[idx.negocio] || '').toLowerCase();
    if (neg && !neg.includes('nacional') && !neg.includes('nal')) continue;

    // Solo fecha más reciente
    if (maxFecha) {
      const f = row[idx.fecha];
      if (f instanceof Date) {
        if (f.toDateString() !== maxFecha.toDateString()) continue;
      }
    }

    const motivo  = String(row[idx.motivo] || '').trim();
    const code    = motivoACode(motivo);
    const monto   = parseFloat(String(row[idx.monto] || '0').replace(/[$,\s]/g, '')) || 0;
    const fechaCompU = row[idx.fechaCompU] instanceof Date
      ? Utilities.formatDate(row[idx.fechaCompU], 'America/Monterrey', 'dd/MM/yyyy') : '';
    const fechaCompO = row[idx.fechaCompO] instanceof Date
      ? Utilities.formatDate(row[idx.fechaCompO], 'America/Monterrey', 'dd/MM/yyyy') : '';

    units.push({
      u:           uid,
      code:        code,
      motivo:      motivo,
      estatus_raw: String(row[idx.estatus] || '').trim(),
      op:          String(row[idx.operador] || '').trim() || 'SIN ASIGNAR',
      ruta:        String(row[idx.ruta] || '').trim(),
      monto:       monto,
      coord:       String(row[idx.coord] || '').trim(),
      obs:         String(row[idx.obs] || '').substring(0, 100),
      fechaCompU:  fechaCompU,
      fechaCompO:  fechaCompO,
      // Campos que se llenaran con otros archivos
      rv: null, kms: null, lit: null, vmax: null, rel: null,
      rf: 'ok', inc: 0, ilist: [], mtto: [], fact: 0, util: 0,
      circ: '—', tu: 'Tractor'
    });
  }
  return units;
}

// ── CONTROL DE CAJAS ─────────────────────────────────────────────────────
function getCajas(ss) {
  const sh = ss.getSheetByName(H_CAJAS);
  if (!sh) return [];

  const data  = sh.getDataRange().getValues();
  if (data.length < 2) return [];

  const heads = data[0].map(h => String(h).trim());

  // Buscar columnas de forma flexible
  function findCol(keywords) {
    for (const kw of keywords) {
      const i = heads.findIndex(h => h.toLowerCase().includes(kw.toLowerCase()));
      if (i >= 0) return i;
    }
    return -1;
  }

  const iNum   = findCol(['caja', 'numero', 'num', 'id']);
  const iEst   = findCol(['estatus', 'estado', 'status']);
  const iCli   = findCol(['cliente']);
  const iViaje = findCol(['viaje', 'folio', 'orden', 'trip']);
  const iOp    = findCol(['operador', 'conductor']);
  const iUbic  = findCol(['ubicacion', 'ubicación', 'ciudad', 'lugar', 'destino', 'location']);
  const iAct   = findCol(['actualizacion', 'actualización', 'ultima', 'última', 'fecha']);

  if (iNum < 0) return [];

  const cajas = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const num = String(row[iNum] || '').trim();
    if (!num) continue;
    cajas.push({
      num:        num,
      estatus:    iEst   >= 0 ? String(row[iEst]   || '').trim() : '',
      cliente:    iCli   >= 0 ? String(row[iCli]   || '').trim() : '',
      viaje:      iViaje >= 0 ? String(row[iViaje] || '').trim() : '',
      op:         iOp    >= 0 ? String(row[iOp]    || '').trim() : '',
      ubicacion:  iUbic  >= 0 ? String(row[iUbic]  || '').trim() : '',
      actualizado:iAct   >= 0 ? (row[iAct] instanceof Date
        ? Utilities.formatDate(row[iAct], 'America/Monterrey', 'dd/MM HH:mm')
        : String(row[iAct] || '').substring(0, 16)) : ''
    });
  }
  return cajas;
}

// ── CONTROL DE OPERADORES ─────────────────────────────────────────────────
function getOperadores(ss) {
  const sh = ss.getSheetByName(H_OPERADORES);
  if (!sh) return [];

  const data  = sh.getDataRange().getValues();
  if (data.length < 2) return [];

  const heads = data[0].map(h => String(h).trim());

  function findCol(keywords) {
    for (const kw of keywords) {
      const i = heads.findIndex(h => h.toLowerCase().includes(kw.toLowerCase()));
      if (i >= 0) return i;
    }
    return -1;
  }

  const iNom   = findCol(['nombre', 'operador']);
  const iUnid  = findCol(['unidad', 'economico']);
  const iCoord = findCol(['coordinador']);
  const iEst   = findCol(['estatus', 'estado']);
  const iLic   = findCol(['licencia']);
  const iVenc  = findCol(['vencimiento', 'vence', 'vigencia']);
  const iTel   = findCol(['telefono', 'teléfono', 'celular', 'contacto']);
  const iAnt   = findCol(['antiguedad', 'antigüedad', 'ingreso', 'alta']);

  if (iNom < 0) return [];

  const ops = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const nom = String(row[iNom] || '').trim();
    if (!nom) continue;
    const venc = iVenc >= 0 && row[iVenc] instanceof Date
      ? Utilities.formatDate(row[iVenc], 'America/Monterrey', 'dd/MM/yyyy') : String(row[iVenc] || '');
    ops.push({
      nombre:    nom,
      unidad:    iUnid  >= 0 ? String(row[iUnid]  || '').trim() : '',
      coord:     iCoord >= 0 ? String(row[iCoord] || '').trim() : '',
      estatus:   iEst   >= 0 ? String(row[iEst]   || '').trim() : '',
      licencia:  iLic   >= 0 ? String(row[iLic]   || '').trim() : '',
      vencimiento: venc,
      telefono:  iTel   >= 0 ? String(row[iTel]   || '').trim() : '',
      antiguedad:iAnt   >= 0 ? (row[iAnt] instanceof Date
        ? Utilities.formatDate(row[iAnt], 'America/Monterrey', 'dd/MM/yyyy')
        : String(row[iAnt] || '')) : ''
    });
  }
  return ops;
}

// ── URGENCIAS ─────────────────────────────────────────────────────────────
function getUrgencias(ss) {
  const sh = ss.getSheetByName(H_URGENCIAS);
  if (!sh) return [];

  const data  = sh.getDataRange().getValues();
  if (data.length < 2) return [];

  const heads = data[0].map(h => String(h).trim());

  function findCol(keywords) {
    for (const kw of keywords) {
      const i = heads.findIndex(h => h.toLowerCase().includes(kw.toLowerCase()));
      if (i >= 0) return i;
    }
    return -1;
  }

  const iHora    = findCol(['hora', 'time', 'fecha']);
  const iUnidad  = findCol(['unidad', 'economico']);
  const iOp      = findCol(['operador', 'conductor']);
  const iProb    = findCol(['problema', 'descripcion', 'descripción', 'mensaje', 'texto']);
  const iPrior   = findCol(['prioridad', 'nivel', 'urgencia', 'criticidad']);
  const iImpacto = findCol(['impacto']);
  const iAccion  = findCol(['accion', 'acción', 'siguiente']);
  const iResp    = findCol(['responsable']);
  const iEst     = findCol(['estatus', 'estado', 'resuelto']);

  const urgencias = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const prob = String(row[iProb >= 0 ? iProb : 0] || '').trim();
    if (!prob) continue;
    urgencias.push({
      hora:        iHora    >= 0 ? (row[iHora] instanceof Date
        ? Utilities.formatDate(row[iHora], 'America/Monterrey', 'dd/MM HH:mm')
        : String(row[iHora] || '')) : '',
      unidad:      iUnidad  >= 0 ? String(row[iUnidad]  || '').trim() : '',
      operador:    iOp      >= 0 ? String(row[iOp]      || '').trim() : '',
      problema:    prob.substring(0, 200),
      prioridad:   iPrior   >= 0 ? String(row[iPrior]   || '').trim() : '',
      impacto:     iImpacto >= 0 ? String(row[iImpacto] || '').trim() : '',
      accion:      iAccion  >= 0 ? String(row[iAccion]  || '').trim() : '',
      responsable: iResp    >= 0 ? String(row[iResp]    || '').trim() : '',
      estatus:     iEst     >= 0 ? String(row[iEst]     || '').trim() : 'Abierto'
    });
  }
  // Devolver las 50 más recientes
  return urgencias.slice(-50).reverse();
}

// ── WHATSAPP (mensajes del grupo) ─────────────────────────────────────────
function getWhatsApp(ss) {
  const sh = ss.getSheetByName(H_WHATSAPP);
  if (!sh) return [];

  const data  = sh.getDataRange().getValues();
  if (data.length < 2) return [];

  const heads = data[0].map(h => String(h).trim());

  function findCol(keywords) {
    for (const kw of keywords) {
      const i = heads.findIndex(h => h.toLowerCase().includes(kw.toLowerCase()));
      if (i >= 0) return i;
    }
    return -1;
  }

  const iFecha   = findCol(['fecha', 'date', 'hora']);
  const iGrupo   = findCol(['grupo', 'group', 'canal']);
  const iRemite  = findCol(['remite', 'quien', 'de', 'from', 'nombre', 'contacto']);
  const iMensaje = findCol(['mensaje', 'texto', 'message', 'contenido']);

  const msgs = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const msg = String(row[iMensaje >= 0 ? iMensaje : 0] || '').trim();
    if (!msg) continue;
    msgs.push({
      fecha:   iFecha   >= 0 ? (row[iFecha] instanceof Date
        ? Utilities.formatDate(row[iFecha], 'America/Monterrey', 'dd/MM HH:mm')
        : String(row[iFecha] || '')) : '',
      grupo:   iGrupo   >= 0 ? String(row[iGrupo]   || '').trim() : '',
      remite:  iRemite  >= 0 ? String(row[iRemite]  || '').trim() : '',
      mensaje: msg.substring(0, 500)
    });
  }
  return msgs.slice(-100).reverse();
}

// ── FUNCIÓN DE PRUEBA (ejecuta desde Apps Script para verificar) ──────────
function testScript() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('Hojas disponibles: ' + ss.getSheets().map(s => s.getName()).join(', '));
  
  const units = getEstatus(ss);
  Logger.log('Unidades cargadas: ' + units.length);
  if (units.length > 0) Logger.log('Primera unidad: ' + JSON.stringify(units[0]));
  
  const cajas = getCajas(ss);
  Logger.log('Cajas cargadas: ' + cajas.length);
  
  const ops = getOperadores(ss);
  Logger.log('Operadores: ' + ops.length);
}
