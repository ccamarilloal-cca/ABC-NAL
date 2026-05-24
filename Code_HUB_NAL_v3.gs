// ═══════════════════════════════════════════════════════════════════════════
// HUB OPERATIVO NAL — Apps Script v3 COMPLETO
// Sin datos fijos. Todo desde Google Sheets.
// ═══════════════════════════════════════════════════════════════════════════
// HOJAS DEL SHEET:
const H = {
  ESTATUS:    'Estatus_Diario',
  CAJAS:      'Control_Cajas',
  OPERADORES: 'Control_Operadores',
  INCIDENCIAS:'Incidencias',
  RENDIMIENTO:'Rendimientos',
  MTTO:       'Solicitudes_Mtto',
  DATA:       'Data',
  MOVIMIENTOS:'Movimientos_RH',
  WHATSAPP:   'Whats App',
  URGENCIAS:  'Urgencias',
  IN_CAJAS:   'IN CAJAS',
  PROG_VIAJES:'Programación de Viajes',
};

// COORDINADORES ACTIVOS (para detección en WhatsApp y filtros)
const COORDS = [
  { nombre: 'YAMILETH FERNANDEZ MURILLO',  alias: ['YAMILETH','YAMILET','FERNANDEZ'] },
  { nombre: 'JUAN JOSE TELLO LAMAS',       alias: ['TELLO','JUAN JOSE','JJ'] },
  { nombre: 'EDER ZUÑIGA RAMIREZ',         alias: ['EDER','ZUÑIGA','ZUNI'] },
  { nombre: 'JULIO ALEJANDRO HERNANDEZ GALVAN', alias: ['JULIO','HERNANDEZ GALVAN'] },
];

// CÓDIGOS DE MOTIVO
const MOTIVOS = {
  TRN:'En Tránsito', VTA:'Facturando', LIB:'Por Liberar', DCO:'Dispo c/Op',
  DSO:'Dispo s/Op',  SO:'Sin Operador', CP:'Correctivo',  SG:'Siniestro',
  RM:'Rep. Mayor',   PER:'Permiso',     SGR:'Sensores GPS'
};

// ── ENTRY POINT ───────────────────────────────────────────────────────────
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'all';
  const result = { ok: false };
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (action === 'all' || action === 'estatus')    result.units      = getEstatus(ss);
    if (action === 'all' || action === 'cajas')      result.cajas      = getCajasCompleto(ss);
    if (action === 'all' || action === 'data')       result.data       = getData(ss);
    if (action === 'all' || action === 'urgencias')  result.urgencias  = getUrgencias(ss);
    if (action === 'all' || action === 'operadores') result.operadores = getOperadores(ss);
    result.fecha = Utilities.formatDate(new Date(),'America/Monterrey','yyyy-MM-dd');
    result.hora  = Utilities.formatDate(new Date(),'America/Monterrey','HH:mm');
    result.ok    = true;
  } catch(err) {
    result.error = err.toString();
    result.stack = err.stack;
  }
  const cb = e && e.parameter && e.parameter.callback;
  const body = JSON.stringify(result);
  return ContentService
    .createTextOutput(cb ? cb+'('+body+')' : body)
    .setMimeType(cb ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}

// ── HELPERS ───────────────────────────────────────────────────────────────
function sheetToObjects(ss, sheetName) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return [];
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  const heads = data[0].map(h => String(h).trim());
  return data.slice(1).map(row => {
    const obj = {};
    heads.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  }).filter(r => Object.values(r).some(v => v !== '' && v !== null && v !== undefined));
}

function col(heads, keywords) {
  for (const kw of keywords) {
    const i = heads.findIndex(h => h.toLowerCase().includes(kw.toLowerCase()));
    if (i >= 0) return heads[i];
  }
  return null;
}

function fmtDate(val) {
  if (!val) return '';
  if (val instanceof Date) return Utilities.formatDate(val, 'America/Monterrey', 'yyyy-MM-dd');
  return String(val).substring(0, 10);
}

function fmtDateTime(val) {
  if (!val) return '';
  if (val instanceof Date) return Utilities.formatDate(val, 'America/Monterrey', 'dd/MM HH:mm');
  return String(val).substring(0, 16);
}

function toNum(v) {
  const n = parseFloat(String(v || '0').replace(/[$,\s]/g, ''));
  return isNaN(n) ? 0 : n;
}

function motivoToCode(motivo) {
  const m = (motivo || '').trim().toUpperCase();
  for (const code of Object.keys(MOTIVOS)) {
    if (m.startsWith(code + ' ') || m.startsWith(code + '-') || m === code) return code;
  }
  return 'DCO';
}

function detectCoord(texto) {
  const t = (texto || '').toUpperCase();
  for (const c of COORDS) {
    if (c.alias.some(a => t.includes(a))) return c.nombre;
  }
  return '';
}

// ── ESTATUS DIARIO ────────────────────────────────────────────────────────
// Solo la fecha más reciente. Solo unidades NAL activas (~80).
function getEstatus(ss) {
  const rows = sheetToObjects(ss, H.ESTATUS);
  if (!rows.length) return [];
  const heads = Object.keys(rows[0]);

  const cFecha  = col(heads, ['Fecha', 'fecha']);
  const cUnidad = col(heads, ['Unidad', 'economico', 'unidad']);
  const cNeg    = col(heads, ['UnidadDeNegocio', 'negocio', 'empresa']);
  const cOp     = col(heads, ['Operador', 'operador']);
  const cEst    = col(heads, ['Estatus', 'estatus', 'estado']);
  const cMotivo = col(heads, ['Motivo', 'motivo', 'movimiento']);
  const cRuta   = col(heads, ['NombreRuta', 'ruta', 'destino']);
  const cMonto  = col(heads, ['Monto', 'monto', 'importe', 'venta']);
  const cCoord  = col(heads, ['Coordinador', 'coordinador']);
  const cObs    = col(heads, ['Comentarios', 'comentario', 'observacion', 'obs']);
  const cFCompU = col(heads, ['FechaCompromisoUnidad', 'compromiso']);
  const cFCompO = col(heads, ['FechaCompromisoOperador', 'comp.op']);

  // Filtrar solo NAL
  let nalRows = rows.filter(r => {
    const uid = String(r[cUnidad] || '').trim();
    if (!uid || uid.length < 3) return false;
    if (cNeg) {
      const neg = String(r[cNeg] || '').toLowerCase();
      if (neg && !neg.includes('nacional') && !neg.includes('nal')) return false;
    }
    return true;
  });

  // Fecha más reciente
  const fechas = nalRows
    .map(r => { const f = r[cFecha]; return f instanceof Date ? f : null; })
    .filter(Boolean);
  if (fechas.length) {
    const maxF = new Date(Math.max(...fechas.map(d => d.getTime())));
    const maxStr = Utilities.formatDate(maxF, 'America/Monterrey', 'yyyy-MM-dd');
    nalRows = nalRows.filter(r => {
      const f = r[cFecha];
      if (!(f instanceof Date)) return true; // sin fecha: incluir
      return Utilities.formatDate(f, 'America/Monterrey', 'yyyy-MM-dd') === maxStr;
    });
  }

  // Deduplicar por unidad (quedarse con último registro)
  const seen = {};
  nalRows.reverse().forEach(r => {
    const uid = String(r[cUnidad] || '').trim();
    if (!seen[uid]) seen[uid] = r;
  });
  nalRows = Object.values(seen);

  return nalRows.map(r => {
    const mot = String(r[cMotivo] || '').trim();
    const est = String(r[cEst]    || '').trim();
    const code = motivoToCode(mot) || motivoToCode(est);
    const monto = toNum(r[cMonto]);
    return {
      u:           String(r[cUnidad] || '').trim(),
      code,
      motivo:      mot || est,
      estatus_raw: est,
      op:          String(r[cOp]     || '').trim() || 'SIN ASIGNAR',
      ruta:        String(r[cRuta]   || '').trim(),
      monto,
      coord:       String(r[cCoord]  || '').trim(),
      obs:         String(r[cObs]    || '').substring(0, 120),
      fechaCompU:  fmtDate(r[cFCompU]),
      fechaCompO:  fmtDate(r[cFCompO]),
      // financiero (se completa con Data)
      fact: 0, util: 0, rv: null, kms: null, circ: '—',
      inc: 0, ilist: [], mtto: [], tu: 'Tractor',
    };
  });
}

// ── BASE DATA — fuente principal de financieros ───────────────────────────
function getData(ss) {
  const rows = sheetToObjects(ss, H.DATA);
  if (!rows.length) return { circs: [], kpis: {}, byCoord: [] };
  const heads = Object.keys(rows[0]);

  const cCirc   = col(heads, ['Circuito', 'circuito', 'ruta']);
  const cUnidad = col(heads, ['NumeroEconomico', 'Unidad', 'economico', 'unidad']);
  const cCoord  = col(heads, ['Coordinador', 'coordinador']);
  const cFact   = col(heads, ['Facturacion', 'facturacion', 'venta', 'Venta']);
  const cUtil   = col(heads, ['Utilidad', 'utilidad', 'util']);
  const cKmsC   = col(heads, ['KmsCargados', 'km cargado', 'kms carg']);
  const cKmsV   = col(heads, ['KmsVacios', 'km vacio', 'kms vac']);
  const cViajes = col(heads, ['Viajes', 'viajes', 'viaje']);
  const cFecha  = col(heads, ['Fecha', 'fecha', 'period']);

  // Filtrar solo totales (excluir filas vacías o header duplicados)
  const dataRows = rows.filter(r => {
    const u = String(r[cUnidad] || '').trim();
    const f = toNum(r[cFact]);
    return u && u.toUpperCase() !== 'TOTAL' && u.toUpperCase() !== 'UNIDAD' && (f !== 0 || u.length > 2);
  });

  // Acumular por circuito
  const circMap = {};
  dataRows.forEach(r => {
    const c = String(r[cCirc] || 'Sin Circuito').trim();
    if (!circMap[c]) circMap[c] = { circ: c, uds: new Set(), fact: 0, util: 0, kmsC: 0, kmsV: 0, viajes: 0 };
    circMap[c].uds.add(String(r[cUnidad] || '').trim());
    circMap[c].fact  += toNum(r[cFact]);
    circMap[c].util  += toNum(r[cUtil]);
    circMap[c].kmsC  += toNum(r[cKmsC]);
    circMap[c].kmsV  += toNum(r[cKmsV]);
    circMap[c].viajes+= toNum(r[cViajes]) || 1;
  });
  const circs = Object.values(circMap).map(c => ({
    circ: c.circ, uds: c.uds.size,
    fact: Math.round(c.fact), util: Math.round(c.util),
    util_ud: c.uds.size ? Math.round(c.util / c.uds.size) : 0,
    kmsC: Math.round(c.kmsC), kmsV: Math.round(c.kmsV),
    viajes: c.viajes,
    margen: c.fact ? Math.round(c.util / c.fact * 100 * 10) / 10 : 0,
  })).filter(c => c.fact !== 0 || c.util !== 0);

  // Acumular por coordinador (solo los 3 activos)
  const coordMap = {};
  dataRows.forEach(r => {
    const rawCoord = String(r[cCoord] || '').trim();
    const coord = rawCoord || detectCoord(String(r[cUnidad] || ''));
    if (!coord) return;
    // Solo coordinadores activos
    const esActivo = COORDS.some(c => coord.toUpperCase().includes(c.alias[0]));
    if (!esActivo && coord) {} // incluir de todas formas para no perder datos
    if (!coordMap[coord]) coordMap[coord] = {
      coord, fact: 0, util: 0, kmsC: 0, kmsV: 0, viajes: 0,
      uds: new Set(), cajas: new Set()
    };
    coordMap[coord].fact   += toNum(r[cFact]);
    coordMap[coord].util   += toNum(r[cUtil]);
    coordMap[coord].kmsC   += toNum(r[cKmsC]);
    coordMap[coord].kmsV   += toNum(r[cKmsV]);
    coordMap[coord].viajes += toNum(r[cViajes]) || 1;
    const u = String(r[cUnidad] || '').trim();
    if (u) coordMap[coord].uds.add(u);
    const caj = String(r[col(heads, ['caja', 'Caja'])] || '').trim();
    if (caj) coordMap[coord].cajas.add(caj);
  });

  // Totales globales del mes
  const totalFact  = dataRows.reduce((s, r) => s + toNum(r[cFact]),  0);
  const totalUtil  = dataRows.reduce((s, r) => s + toNum(r[cUtil]),  0);
  const totalKmsC  = dataRows.reduce((s, r) => s + toNum(r[cKmsC]),  0);
  const totalKmsV  = dataRows.reduce((s, r) => s + toNum(r[cKmsV]),  0);
  const totalViaj  = dataRows.reduce((s, r) => s + (toNum(r[cViajes]) || 1), 0);

  const byCoord = Object.values(coordMap)
    .filter(c => c.fact !== 0 || c.util !== 0)
    .map(c => ({
      coord:   c.coord,
      fact:    Math.round(c.fact),
      util:    Math.round(c.util),
      margen:  c.fact ? Math.round(c.util / c.fact * 1000) / 10 : 0,
      kmsC:    Math.round(c.kmsC),
      kmsV:    Math.round(c.kmsV),
      viajes:  c.viajes,
      uds:     c.uds.size,
      cajas:   c.cajas.size,
    }))
    .sort((a, b) => b.fact - a.fact);

  // Unidades por unidad (para enriquecer estatus)
  const byUnit = {};
  dataRows.forEach(r => {
    const u = String(r[cUnidad] || '').trim();
    if (!u) return;
    if (!byUnit[u]) byUnit[u] = { u, fact: 0, util: 0, kmsC: 0, circ: '', coord: '' };
    byUnit[u].fact  += toNum(r[cFact]);
    byUnit[u].util  += toNum(r[cUtil]);
    byUnit[u].kmsC  += toNum(r[cKmsC]);
    if (r[cCirc]) byUnit[u].circ = String(r[cCirc]).trim();
    if (r[cCoord]) byUnit[u].coord = String(r[cCoord]).trim();
  });

  return {
    circs,
    byCoord,
    byUnit,
    kpis: {
      fact_mes:  Math.round(totalFact),
      util_mes:  Math.round(totalUtil),
      margen_mes: totalFact ? Math.round(totalUtil / totalFact * 1000) / 10 : 0,
      kmsC_mes:  Math.round(totalKmsC),
      kmsV_mes:  Math.round(totalKmsV),
      viajes_mes: totalViaj,
    }
  };
}

// ── CAJAS COMPLETO — cruza Control_Cajas + IN CAJAS + Programación ────────
function getCajasCompleto(ss) {
  const cajRows  = sheetToObjects(ss, H.CAJAS)       || [];
  const inRows   = sheetToObjects(ss, H.IN_CAJAS)    || [];
  const progRows = sheetToObjects(ss, H.PROG_VIAJES) || [];

  // Indexar programación por caja
  const progMap = {};
  if (progRows.length) {
    const ph = Object.keys(progRows[0]);
    const pCaja  = col(ph, ['caja', 'Caja', 'cajón']);
    const pViaje = col(ph, ['viaje', 'folio', 'orden']);
    const pUnid  = col(ph, ['unidad', 'economico']);
    const pOrig  = col(ph, ['origen', 'carga']);
    const pDest  = col(ph, ['destino', 'entrega']);
    const pCli   = col(ph, ['cliente', 'Client']);
    const pFecha = col(ph, ['fecha', 'programado', 'salida']);
    const pCoord = col(ph, ['coordinador', 'coord']);
    progRows.forEach(r => {
      const k = String(r[pCaja] || '').trim();
      if (!k) return;
      progMap[k] = {
        viaje: String(r[pViaje] || '').trim(),
        unidad: String(r[pUnid] || '').trim(),
        origen: String(r[pOrig] || '').trim(),
        destino: String(r[pDest] || '').trim(),
        cliente: String(r[pCli] || '').trim(),
        fecha: fmtDate(r[pFecha]),
        coord: String(r[pCoord] || '').trim(),
      };
    });
  }

  // Indexar IN CAJAS por número de caja
  const inMap = {};
  if (inRows.length) {
    const ih = Object.keys(inRows[0]);
    const iCaja   = col(ih, ['caja', 'numero', 'num']);
    const iEst    = col(ih, ['estatus', 'estado', 'status']);
    const iUbic   = col(ih, ['ubicacion', 'patio', 'lugar']);
    const iUltMov = col(ih, ['ultima', 'actualizacion', 'fecha']);
    inRows.forEach(r => {
      const k = String(r[iCaja] || '').trim();
      if (!k) return;
      inMap[k] = {
        estatus: String(r[iEst] || '').trim(),
        ubicacion: String(r[iUbic] || '').trim(),
        ultimoMov: fmtDate(r[iUltMov]),
      };
    });
  }

  // Construir cajas finales
  const now = new Date();
  const cajas = [];
  const procesadas = new Set();

  // Desde Control_Cajas
  if (cajRows.length) {
    const ch = Object.keys(cajRows[0]);
    const cNum   = col(ch, ['caja', 'numero', 'num', 'id']);
    const cEst   = col(ch, ['estatus', 'estado', 'status']);
    const cCli   = col(ch, ['cliente', 'Client']);
    const cViaje = col(ch, ['viaje', 'folio']);
    const cOp    = col(ch, ['operador', 'conductor']);
    const cUbic  = col(ch, ['ubicacion', 'ciudad', 'lugar', 'patio']);
    const cAct   = col(ch, ['actualizacion', 'ultima', 'fecha']);
    const cCoord = col(ch, ['coordinador', 'coord']);
    const cUnid  = col(ch, ['unidad', 'tractor']);
    const cOrig  = col(ch, ['origen', 'carga']);
    const cDest  = col(ch, ['destino', 'entrega']);

    cajRows.forEach(r => {
      const num = String(r[cNum] || '').trim();
      if (!num) return;
      procesadas.add(num);
      const inInfo  = inMap[num]  || {};
      const prog    = progMap[num] || {};
      const estatus = String(r[cEst] || inInfo.estatus || '').trim();
      const ultAct  = r[cAct] instanceof Date ? r[cAct] : (inInfo.ultimoMov ? new Date(inInfo.ultimoMov) : null);
      const diasSin = ultAct ? Math.floor((now - ultAct) / 86400000) : null;
      const analisis = generarAnalisisCaja(estatus, diasSin, prog, inInfo);

      cajas.push({
        num,
        estatus,
        cliente:   String(r[cCli]   || prog.cliente || '').trim(),
        viaje:     String(r[cViaje] || prog.viaje   || '').trim(),
        op:        String(r[cOp]    || '').trim(),
        ubicacion: String(r[cUbic]  || inInfo.ubicacion || prog.destino || '').trim(),
        actualizado: fmtDateTime(r[cAct]) || inInfo.ultimoMov || '',
        coord:     String(r[cCoord] || prog.coord   || '').trim(),
        unidad:    String(r[cUnid]  || prog.unidad  || '').trim(),
        origen:    String(r[cOrig]  || prog.origen  || '').trim(),
        destino:   String(r[cDest]  || prog.destino || '').trim(),
        diasSinMov: diasSin,
        analisis,
      });
    });
  }

  // Agregar cajas que solo están en IN CAJAS
  Object.entries(inMap).forEach(([num, info]) => {
    if (procesadas.has(num)) return;
    const prog = progMap[num] || {};
    const diasSin = info.ultimoMov ? Math.floor((now - new Date(info.ultimoMov)) / 86400000) : null;
    cajas.push({
      num, estatus: info.estatus, cliente: prog.cliente || '',
      viaje: prog.viaje || '', op: '', ubicacion: info.ubicacion || '',
      actualizado: info.ultimoMov || '', coord: prog.coord || '',
      unidad: prog.unidad || '', origen: prog.origen || '', destino: prog.destino || '',
      diasSinMov: diasSin, analisis: generarAnalisisCaja(info.estatus, diasSin, prog, info),
    });
    procesadas.add(num);
  });

  return cajas;
}

function generarAnalisisCaja(estatus, diasSin, prog, inInfo) {
  const e = (estatus || '').toUpperCase();
  if (e.includes('TRÁNSITO') || e.includes('TRANSITO')) return 'En tránsito hacia ' + (prog.destino || 'destino');
  if (e.includes('CARGAD')) return 'Cargada — pendiente de salida';
  if (prog.viaje) return 'Programada para viaje ' + prog.viaje + ' → ' + (prog.destino || '');
  if (e.includes('DAÑA') || e.includes('SINIESTRO')) return 'Requiere atención — daño reportado';
  if (e.includes('PATIO')) return 'Disponible en patio' + (inInfo.ubicacion ? ' ' + inInfo.ubicacion : '');
  if (e.includes('DISPONIB')) return 'Disponible para asignación';
  if (diasSin !== null && diasSin > 15) return 'Sin movimiento ' + diasSin + ' días — requiere validación física';
  if (diasSin !== null && diasSin > 7)  return 'Sin movimiento ' + diasSin + ' días — verificar con coordinador';
  if (e.includes('CLIENT') || e.includes('CON CLIENT')) return 'Con cliente — pendiente devolución';
  if (e.includes('NO LOCAL') || e.includes('SIN UBIC'))  return 'Sin ubicación confirmada — localizar urgente';
  return 'Verificar estatus con patio';
}

// ── URGENCIAS / WHATSAPP ──────────────────────────────────────────────────
// Lee la hoja "Urgencias" y también procesa "Whats App" si tiene texto raw
function getUrgencias(ss) {
  const rows = sheetToObjects(ss, H.URGENCIAS);
  if (!rows.length) return [];
  const heads = Object.keys(rows[0]);

  const cFecha  = col(heads, ['fecha', 'Fecha', 'date']);
  const cHora   = col(heads, ['hora', 'Hora', 'time']);
  const cUnidad = col(heads, ['unidad', 'Unidad', 'economico']);
  const cCaja   = col(heads, ['caja', 'Caja']);
  const cOp     = col(heads, ['operador', 'Operador', 'conductor']);
  const cCoord  = col(heads, ['coordinador', 'Coordinador', 'coord']);
  const cCliente= col(heads, ['cliente', 'Cliente']);
  const cTipo   = col(heads, ['tipo', 'Tipo', 'incidencia', 'categoria']);
  const cProb   = col(heads, ['problema', 'descripcion', 'mensaje', 'texto', 'Descripcion']);
  const cEst    = col(heads, ['estatus', 'Estatus', 'estado']);
  const cRes    = col(heads, ['resuelto', 'Resuelto', 'resolved']);
  const cPend   = col(heads, ['pendiente', 'Pendiente', 'pending']);
  const cSeg    = col(heads, ['seguimiento', 'Seguimiento', 'responsable']);
  const cAct    = col(heads, ['actualizacion', 'actualización', 'ultima']);

  return rows.map(r => {
    const texto = String(r[cProb] || '').trim();
    if (!texto) return null;
    const coord = String(r[cCoord] || detectCoord(texto) || '').trim();
    const tipo  = String(r[cTipo] || clasificarIncidencia(texto) || '').trim();
    return {
      fecha:      fmtDate(r[cFecha]),
      hora:       String(r[cHora] || '').substring(0,5),
      unidad:     String(r[cUnidad] || extractUnidad(texto) || '').trim(),
      caja:       String(r[cCaja]  || extractCaja(texto)   || '').trim(),
      operador:   String(r[cOp]    || '').trim(),
      coord:      coord,
      cliente:    String(r[cCliente] || '').trim(),
      tipo,
      problema:   texto.substring(0, 200),
      estatus:    String(r[cEst]  || 'Abierto').trim(),
      resuelto:   String(r[cRes]  || 'No').trim(),
      pendiente:  String(r[cPend] || 'Sí').trim(),
      seguimiento:String(r[cSeg]  || coord || '').trim(),
      actualizado:fmtDateTime(r[cAct]),
    };
  }).filter(Boolean).reverse().slice(0, 100);
}

function extractUnidad(texto) {
  const m = texto.match(/\b(\d{3}-ABC|ABC-\d{3}|\d{3}ABC)\b/i);
  return m ? m[0].toUpperCase() : '';
}

function extractCaja(texto) {
  const m = texto.match(/\bcaja[\s:#-]*([A-Z0-9]{3,10})\b/i);
  return m ? m[1].toUpperCase() : '';
}

function clasificarIncidencia(texto) {
  const t = (texto || '').toLowerCase();
  if (/accidente|choque|volcadura|colisión|siniestro/.test(t))        return 'Siniestro';
  if (/robo|asalto|secuestro|extorsión|levantón/.test(t))             return 'Seguridad';
  if (/falla|motor|frenos|llanta|mecánic|descompuest|varado/.test(t)) return 'Mecánica';
  if (/caja|remolque|pinchazo|daño caja/.test(t))                     return 'Caja';
  if (/cliente|entrega|queja|reclamación|factura/.test(t))            return 'Cliente';
  if (/rescate|grúa|arrastre/.test(t))                                return 'Rescate';
  if (/retraso|demora|ventana|tardó/.test(t))                         return 'Retraso';
  if (/document|licencia|permiso|carta porte/.test(t))                return 'Documentación';
  if (/preventivo|aceite|servicio|taller/.test(t))                    return 'Mantenimiento';
  if (/tractor|cabeza/.test(t))                                       return 'Tractor';
  return 'Operativa';
}

// ── OPERADORES ────────────────────────────────────────────────────────────
function getOperadores(ss) {
  const rows = sheetToObjects(ss, H.OPERADORES);
  if (!rows.length) return [];
  const heads = Object.keys(rows[0]);
  const cNom   = col(heads, ['nombre', 'operador', 'Nombre']);
  const cUnid  = col(heads, ['unidad', 'economico', 'Unidad']);
  const cCoord = col(heads, ['coordinador', 'Coordinador']);
  const cEst   = col(heads, ['estatus', 'estado', 'Estatus']);
  const cLic   = col(heads, ['licencia', 'Licencia']);
  const cVenc  = col(heads, ['vencimiento', 'vigencia', 'vence']);
  const cTel   = col(heads, ['telefono', 'celular', 'Telefono']);
  const cAnt   = col(heads, ['antiguedad', 'ingreso', 'alta']);
  return rows.map(r => {
    const nom = String(r[cNom] || '').trim();
    if (!nom) return null;
    const venc = r[cVenc] instanceof Date
      ? Utilities.formatDate(r[cVenc], 'America/Monterrey', 'dd/MM/yyyy')
      : String(r[cVenc] || '');
    const ant = r[cAnt] instanceof Date
      ? Utilities.formatDate(r[cAnt], 'America/Monterrey', 'dd/MM/yyyy')
      : String(r[cAnt] || '');
    const hoy = new Date();
    const diasVenc = r[cVenc] instanceof Date ? Math.floor((r[cVenc] - hoy) / 86400000) : null;
    return {
      nombre: nom,
      unidad: String(r[cUnid]  || '').trim(),
      coord:  String(r[cCoord] || '').trim(),
      estatus:String(r[cEst]   || '').trim(),
      licencia:String(r[cLic]  || '').trim(),
      vencimiento: venc,
      diasParaVencer: diasVenc,
      alertaLic: diasVenc !== null && diasVenc < 30,
      telefono:  String(r[cTel] || '').trim(),
      antiguedad: ant,
    };
  }).filter(Boolean);
}

// ── FUNCIÓN DE PRUEBA ─────────────────────────────────────────────────────
function testV3() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojas = ss.getSheets().map(s => s.getName());
  Logger.log('Hojas: ' + hojas.join(', '));

  const units = getEstatus(ss);
  Logger.log('Unidades: ' + units.length);
  if (units.length) {
    const codes = {};
    units.forEach(u => codes[u.code] = (codes[u.code]||0)+1);
    Logger.log('Codes: ' + JSON.stringify(codes));
    Logger.log('Ejemplo: ' + JSON.stringify(units[0]));
  }

  const data = getData(ss);
  Logger.log('Circuitos: ' + data.circs.length);
  Logger.log('Por coord: ' + data.byCoord.length);
  Logger.log('KPIs mes: ' + JSON.stringify(data.kpis));

  const cajas = getCajasCompleto(ss);
  Logger.log('Cajas: ' + cajas.length);

  const urg = getUrgencias(ss);
  Logger.log('Urgencias: ' + urg.length);
}
