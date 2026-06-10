const fs = require('fs');
const path = require('path');
const multer = require('multer');
const XLSX = require('xlsx');

const DATA_DIR     = path.join(__dirname, '../data/proyectos');
const UPLOAD_DIR   = path.join(__dirname, '../uploads/contratos');
const upload       = multer({ dest: '/tmp/waller-uploads/' });
// Multer con almacenamiento permanente para contratos PDF
const uploadContrato = multer({
  storage: multer.diskStorage({
    destination: (req, file, cb) => {
      if (!fs.existsSync(UPLOAD_DIR)) fs.mkdirSync(UPLOAD_DIR, { recursive: true });
      cb(null, UPLOAD_DIR);
    },
    filename: (req, file, cb) => {
      const cc  = req.params.cc || 'obra';
      const ext = path.extname(file.originalname) || '.pdf';
      cb(null, `${cc}-contrato${ext}`);
    }
  }),
  fileFilter: (req, file, cb) => cb(null, /pdf/i.test(file.mimetype) || file.originalname.toLowerCase().endsWith('.pdf'))
});

const OBRA_NAMES = {
  '509':'Play and Fun','510':'NHAOS','511':'Hidalma','512':'Casa Velázquez',
  '513':'Trabajos en Planta','514':'Deimare','515':'Praia','516':'Amatitán',
  '517':'Artwalk','518':'IU Life','519':'Matilde','520':'Zendera'
};

const PARTIDA_MAP = {
  '01':'iva_no_acreditable','02':'deductivas',
  '10':'muros_panel','11':'fletes','20':'complementarios',
  '30':'mano_de_obra','40':'herramienta_mano','50':'herramienta_corte',
  '60':'renta_andamios','70':'epp','80':'indirectos_campo',
  '90':'indirectos_central'
};

const PARTIDA_LABELS = {
  muros_panel:'Muros / Panel',fletes:'Fletes',complementarios:'Complementarios',
  mano_de_obra:'Mano de obra',herramienta_mano:'Herramienta de mano',
  herramienta_corte:'Herramienta de corte',renta_andamios:'Renta de andamios',
  epp:'EPP',indirectos_campo:'Indirectos de campo',
  indirectos_central:'Indirectos of. central',iva_no_acreditable:'IVA no acreditable',
  deductivas:'Deductivas'
};

function leerProyecto(cc) {
  const file = path.join(DATA_DIR, cc + '.json');
  if (!fs.existsSync(file)) return null;
  try { return JSON.parse(fs.readFileSync(file, 'utf8')); } catch(e) { return null; }
}
function guardarProyecto(cc, data) {
  if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
  fs.writeFileSync(path.join(DATA_DIR, cc + '.json'), JSON.stringify(data, null, 2), 'utf8');
}

// Guarda metadata del archivo procesado: solo actual + respaldo anterior
function actualizarMetaArchivo(cc, tipo, nombre) {
  const proyecto = leerProyecto(cc) || { cc };
  if (!proyecto._archivos) proyecto._archivos = {};
  const actual = proyecto._archivos[tipo]?.actual;
  proyecto._archivos[tipo] = {
    actual: { nombre, fechaCarga: new Date().toISOString().slice(0, 10), estado: 'ok' },
    ...(actual ? { anterior: actual } : {})
  };
  guardarProyecto(cc, proyecto);
}

function parseLibroAuxiliar(filePath) {
  const wb = XLSX.readFile(filePath, { cellText: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  const result = {};
  let currentCC = null;
  let currentPartida = null;

  for (const row of rows) {
    const col1 = String(row[0] || '').trim();
    const col6 = String(row[5] || '').replace(/,/g, '').trim();

    const acctMatch = col1.match(/^5(\d{2})-(\d{2})-(\d{3})$/);
    if (acctMatch) {
      currentCC = '5' + acctMatch[1];
      const sub = acctMatch[2];
      currentPartida = PARTIDA_MAP[sub] || ('cuenta_' + sub);
      if (!result[currentCC]) {
        result[currentCC] = {
          cc: currentCC,
          nombre: OBRA_NAMES[currentCC] || ('Obra ' + currentCC),
          partidas: {},
          _cargos: {},
          _abonos: {},
          ultimaActualizacion: new Date().toISOString().slice(0, 10)
        };
      }
      continue;
    }

    // Si encontramos cualquier otra cuenta (ej: 102-xx, 602-xx), resetear contexto
    if (col1.match(/^\d{3}-\d{2}-\d{3}$/)) {
      currentCC = null;
      currentPartida = null;
      continue;
    }

    // Ignorar filas de totales y resumen
    if (String(row[4] || '').includes('Total') || col1.startsWith('Total')) {
      continue;
    }

    if (col1.match(/^\d{1,2}\/\w{3}\/\d{4}$/) && currentCC && currentPartida) {
      const cargo = parseFloat(col6) || 0;
      const abono = parseFloat(String(row[6] || '').replace(/,/g, '').trim()) || 0;
      if (cargo > 0) result[currentCC]._cargos[currentPartida] = (result[currentCC]._cargos[currentPartida] || 0) + cargo;
      if (abono > 0) result[currentCC]._abonos[currentPartida] = (result[currentCC]._abonos[currentPartida] || 0) + abono;
    }
  }

  for (const cc of Object.keys(result)) {
    const obra = result[cc];
    obra.partidas = {};
    for (const [k, v] of Object.entries(obra._cargos || {})) {
      obra.partidas[k] = Math.max(0, v - (obra._abonos[k] || 0));
    }
    obra.totalEgresado = Object.values(obra.partidas).reduce((s, v) => s + v, 0);
    obra.partidasDetalle = Object.entries(obra.partidas).map(([k, v]) => ({
      clave: k,
      nombre: PARTIDA_LABELS[k] || k,
      total: Math.round(v * 100) / 100,
      cargos: Math.round((obra._cargos[k] || 0) * 100) / 100,
      abonos: Math.round((obra._abonos[k] || 0) * 100) / 100
    })).sort((a, b) => b.total - a.total);
    delete obra._cargos;
    delete obra._abonos;
  }

  return result;
}

function parseCronograma(filePath) {
  // Leer sin cellText para obtener valores numéricos reales de fórmulas
  const wb = XLSX.readFile(filePath, { cellText: false, cellDates: false });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  const NIVELES = ['NPB','N1','N2','N3','N4','N5','N6','N7','N8','N9','N10','N11','N12','NAZ','S1','S2'];

  const toNum = v => {
    if (typeof v === 'number') return isFinite(v) ? v : 0;
    const s = String(v || '').replace(/,/g,'').replace(/#/g,'').trim();
    const n = parseFloat(s);
    return isFinite(n) ? n : 0;
  };

  // 1. Detectar columna inicial de semanas y cantidad de semanas
  let semanaStartCols = [];
  let semanaRowIdx = -1;
  for (let r = 0; r < Math.min(30, rows.length); r++) {
    const row = rows[r];
    for (let c = 0; c < row.length; c++) {
      if (String(row[c] || '').toLowerCase().includes('semana')) {
        semanaRowIdx = r;
        // Recopilar todas las columnas con "semana"
        for (let cc = c; cc < row.length; cc++) {
          if (String(row[cc] || '').toLowerCase().includes('semana')) {
            semanaStartCols.push(cc);
          }
        }
        break;
      }
    }
    if (semanaRowIdx >= 0) break;
  }

  // 2. Detectar columna inicial de niveles (fila con NPB)
  let nivelesStartCol = -1;
  for (let r = 0; r < Math.min(20, rows.length); r++) {
    const idx = rows[r].findIndex(c => String(c).trim().toUpperCase() === 'NPB');
    if (idx >= 0) { nivelesStartCol = idx; break; }
  }
  if (nivelesStartCol < 0) nivelesStartCol = 13;

  // 3. m2 por nivel — solo filas que mencionan "cm" en col 11 o 10
  const m2PorNivel = {};
  for (let r = 0; r < Math.min(30, rows.length); r++) {
    const row = rows[r];
    const desc = String(row[11] || row[10] || '').toLowerCase();
    if (!desc.includes('cm') || desc.includes('dias') || desc.includes('día')) continue;
    NIVELES.forEach((nv, idx) => {
      const v = toNum(row[nivelesStartCol + idx]);
      if (v > 0) m2PorNivel[nv] = Math.round(((m2PorNivel[nv] || 0) + v) * 1000) / 1000;
    });
  }

  // 4. Programa semanal por nivel
  // El formato real de Waller tiene el nivel en col A (idx 0), no en col B
  const programaSemanal = {};
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r];
    // Buscar nivel en col A o col B
    const colA = String(row[0] || '').trim().toUpperCase();
    const colB = String(row[1] || '').trim().toUpperCase();
    const nivel = NIVELES.includes(colA) ? colA : NIVELES.includes(colB) ? colB : null;
    if (!nivel) continue;
    programaSemanal[nivel] = {};
    semanaStartCols.forEach((startCol, semIdx) => {
      // El valor de la semana está en la primera columna del grupo (no suma días)
      const v = toNum(row[startCol]);
      if (v > 0) programaSemanal[nivel]['S' + (semIdx + 1)] = Math.round(v * 10) / 10;
    });
  }

  // 5. Fecha de inicio — fila de seriales de fecha justo debajo de "Semana 1"
  let fechaInicio = null;
  if (semanaRowIdx >= 0) {
    // Buscar la fila de seriales (números > 40000 = fechas Excel post-2009)
    for (let r = semanaRowIdx + 1; r < Math.min(semanaRowIdx + 5, rows.length); r++) {
      const row = rows[r];
      for (let c = 0; c < row.length; c++) {
        const v = toNum(row[c]);
        if (v > 40000 && v < 60000) {
          // Convertir serial Excel a fecha ISO
          const ms = Math.round((v - 25569) * 86400 * 1000);
          fechaInicio = new Date(ms).toISOString().slice(0, 10);
          break;
        }
      }
      if (fechaInicio) break;
    }
  }

  return {
    m2PorNivel,
    programaSemanal,
    totalM2: Math.round(Object.values(m2PorNivel).reduce((s, v) => s + v, 0) * 100) / 100,
    totalSemanas: semanaStartCols.length || 15,
    ...(fechaInicio ? { fechaInicio } : {})
  };
}

function parseEstimacion(filePath) {
  const wb = XLSX.readFile(filePath, { cellText: false });
  const names = wb.SheetNames;

  const estName = names.find(n => /^ESTIMACI[ÓO]N$/i.test(n.trim()));
  const estRows = XLSX.utils.sheet_to_json(wb.Sheets[estName || names[2]], { header: 1, defval: '' });

  const numMatch = String(estRows[0]?.[8] || '').match(/(\d+)/);
  const numero = numMatch ? parseInt(numMatch[1]) : null;
  const totalPpto = parseFloat(estRows[4]?.[6]) || 0;

  const r6 = estRows[5] || [];
  const acumAnterior   = parseFloat(r6[8])  || 0;
  const estaEstimacion = parseFloat(r6[11]) || 0;
  const acumActual     = parseFloat(r6[14]) || 0;
  const porEjercer     = parseFloat(r6[17]) || 0;

  const conceptos = [];
  let m2Esta = 0, m2Acum = 0;
  for (let i = 10; i < 17 && i < estRows.length; i++) {
    const row = estRows[i];
    const concepto = String(row[2] || '').trim();
    if (!concepto || /^total/i.test(concepto) || concepto === '.') continue;
    const cantTotal    = parseFloat(row[4])  || 0;
    const pu           = parseFloat(row[5])  || 0;
    const importeTotal = parseFloat(row[6])  || 0;
    const cantEsta     = parseFloat(row[11]) || 0;
    const importeEsta  = parseFloat(row[12]) || 0;
    const cantAcum     = parseFloat(row[14]) || 0;
    const importeAcum  = parseFloat(row[15]) || 0;
    if (cantTotal > 0 || importeTotal > 0) {
      conceptos.push({ concepto, cantTotal, pu, importeTotal, cantEsta, importeEsta, cantAcum, importeAcum });
      m2Esta += cantEsta;
      m2Acum += cantAcum;
    }
  }

  let periodoDesde = null, periodoHasta = null;
  try {
    const cRows = XLSX.utils.sheet_to_json(wb.Sheets[names[1]], { header: 1, defval: '' });
    const r16 = cRows[15] || [];
    const toISO = s => new Date(Math.round((parseFloat(s) - 25569) * 86400000)).toISOString().slice(0, 10);
    if (parseFloat(r16[5]) > 40000) { periodoDesde = toISO(r16[5]); periodoHasta = toISO(r16[6]); }
  } catch(e) {}

  let retencionEsta = estaEstimacion * 0.05, importePagar = estaEstimacion * 0.95, montoContrato = totalPpto;
  try {
    const c0 = XLSX.utils.sheet_to_json(wb.Sheets[names[0]], { header: 1, defval: '' });
    montoContrato = parseFloat(c0[12]?.[4]) || totalPpto;
    retencionEsta = parseFloat(c0[23]?.[10]) || retencionEsta;
    importePagar  = parseFloat(c0[30]?.[10]) || importePagar;
  } catch(e) {}

  let cc = null;
  const txt = estRows.slice(0, 15).map(r => (r||[]).join(' ')).join(' ').toUpperCase();
  for (const [code, name] of Object.entries(OBRA_NAMES)) {
    if (name && txt.includes(name.toUpperCase())) { cc = code; break; }
  }

  return {
    numero, cc, montoContrato, totalPpto,
    acumAnterior, estaEstimacion, acumActual, porEjercer,
    periodoDesde, periodoHasta,
    retencionEsta: Math.round(retencionEsta * 100) / 100,
    importePagar:  Math.round(importePagar * 100) / 100,
    m2Esta: Math.round(m2Esta * 100) / 100,
    m2Acum: Math.round(m2Acum * 100) / 100,
    conceptos
  };
}

// ── Extractor de campos de contrato (regex sobre texto real Waller) ───────────
function extraerCamposContrato(texto) {
  const t = texto || '';

  // ── Helpers ────────────────────────────────────────────────────────────────
  const num = pattern => {
    const m = t.match(pattern);
    return m ? parseFloat(String(m[1]).replace(/,/g,'')) : null;
  };
  const str = pattern => {
    const m = t.match(pattern);
    return m ? m[1]?.trim() || null : null;
  };
  const MESES = {enero:'01',febrero:'02',marzo:'03',abril:'04',mayo:'05',junio:'06',
    julio:'07',agosto:'08',septiembre:'09',octubre:'10',noviembre:'11',diciembre:'12'};
  const fecha = (d,mes,y) => {
    const mm = MESES[(mes||'').toLowerCase()];
    return d && mm && y ? `${y}-${mm}-${String(d).padStart(2,'0')}` : null;
  };
  const fechaPattern = txt => {
    // "el día X de MES de/del AÑO"
    let m = txt?.match(/(\d{1,2})\s+(?:de\s+)?(\w+)\s+(?:de[l]?\s+)?(\d{4})/i);
    if (m) return fecha(m[1], m[2], m[3]);
    // DD/MM/YYYY
    m = txt?.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
    if (m) { const meses2 = ['','01','02','03','04','05','06','07','08','09','10','11','12']; return `${m[3]}-${meses2[parseInt(m[2])]||m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}`; }
    return null;
  };

  // ── SECCIÓN 1: Financieros ─────────────────────────────────────────────────

  // Tipo de contrato
  const tipoMatch = t.match(/(?:contrato\s+a\s+)?precio\s+(alzado|unitario(?:\s+con\s+m[aá]ximo\s+garantizado)?)/i);
  const tipoContrato = tipoMatch
    ? (tipoMatch[1].toLowerCase().includes('alzado') ? 'Precio Alzado'
      : tipoMatch[1].toLowerCase().includes('m') ? 'Precio Unitario con Máximo Garantizado'
      : 'Precio Unitario')
    : null;

  // Monto (múltiples patrones del formato Waller)
  const montoPatterns = [
    /cantidad\s+(?:neta\s+)?de\s+\$\s*([\d,]+(?:\.\d+)?)/i,
    /valor\s+de[l]?\s+contrato[^$\n]*\$\s*([\d,]+(?:\.\d+)?)/i,
    /monto\s+(?:total\s+)?(?:del?\s+)?contrato[^$\n]*\$\s*([\d,]+(?:\.\d+)?)/i,
    /importe\s+(?:total\s+)?(?:del?\s+)?contrato[^$\n]*\$\s*([\d,]+(?:\.\d+)?)/i,
    /por\s+la\s+cantidad\s+de\s+\$\s*([\d,]+(?:\.\d+)?)/i,
    /precio\s+(?:total\s+)?pactado[^$\n]*\$\s*([\d,]+(?:\.\d+)?)/i,
  ];
  let montoContrato = null;
  for (const p of montoPatterns) { montoContrato = num(p); if (montoContrato && montoContrato > 10000) break; }
  // Fallback: mayor cantidad > 100,000 en el texto
  if (!montoContrato) {
    const todos = [...t.matchAll(/\$\s*([\d,]+(?:\.\d{2})?)/g)].map(m=>parseFloat(m[1].replace(/,/g,''))).filter(v=>v>100000);
    if (todos.length) montoContrato = Math.max(...todos);
  }

  const anticipoPct      = num(/anticipo[^%\d]*(\d+(?:\.\d+)?)\s*%/i);
  const retencionPct     = num(/(?:retenci[oó]n|fondo\s+de\s+garant[ií]a)[^%\d]*(\d+(?:\.\d+)?)\s*%/i);
  const penalizacionDiaPct = num(/pena[^\d%]*(\d+(?:\.\d+)?)\s*(?:al\s+millar|%)[^d]*d[ií]a/i)
    || num(/(\d+(?:\.\d+)?)\s*%[^d]*por\s+d[ií]a/i);
  const penalRescisionPct = num(/rescisi[oó]n[^%\d]*(\d+(?:\.\d+)?)\s*%/i);
  const topeMaxPenalizacionPct = num(/tope[^%\d]*(\d+(?:\.\d+)?)\s*%/i)
    || num(/hasta\s+(?:un\s+)?(?:m[aá]ximo\s+de\s+)?(\d+(?:\.\d+)?)\s*%/i);

  // Fechas
  const fechaFirmaRaw   = str(/(?:el\s+d[ií]a|firmado\s+(?:el|en))[^:,]?\s+(\d{1,2}\s+\w+\s+\d{4}|\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4})/i);
  const fechaIniRaw     = str(/(?:fecha\s+de\s+inicio|iniciar[aá]\s+(?:los\s+)?trabajos|dar[aá]\s+inicio)[^\d]*(\d{1,2}\s+(?:de\s+)?\w+\s+(?:de[l]?\s+)?\d{4}|\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4})/i);
  const fechaTermRaw    = str(/(?:fecha\s+de\s+terminaci[oó]n|fecha\s+l[ií]mite|concluir[aá]\s+(?:los\s+)?trabajos)[^\d]*(\d{1,2}\s+(?:de\s+)?\w+\s+(?:de[l]?\s+)?\d{4}|\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4})/i);

  const fechaFirma        = fechaPattern(fechaFirmaRaw);
  const fechaInicio       = fechaPattern(fechaIniRaw);
  const fechaTerminacion  = fechaPattern(fechaTermRaw);

  let plazoEjecucionDias = num(/plazo[^:]*?(\d+)\s*d[ií]as\s+calendario/i)
    || num(/(\d+)\s*d[ií]as\s+calendario/i)
    || num(/plazo\s+de\s+ejecuci[oó]n[^:]*:?\s*(\d+)/i);
  // Calcular desde fechas si no se detectó
  if (!plazoEjecucionDias && fechaInicio && fechaTerminacion) {
    plazoEjecucionDias = Math.round((new Date(fechaTerminacion) - new Date(fechaInicio)) / 86400000);
    if (plazoEjecucionDias <= 0) plazoEjecucionDias = null;
  }

  // ── SECCIÓN 2: Partes ──────────────────────────────────────────────────────
  const folioContrato = str(/(?:folio|n[uú]mero\s+de\s+contrato|contrato\s+n[uoú][m.]?)[^\w\n]*([A-Z0-9][\w\-\/]{3,25})/i);

  const nombreProyecto =
    // 1. Nombre comercial — Torre/Edificio/Residencial/etc en el encabezado
    str(/(?:PROYECTO|OBRA|EDIFICIO|TORRE|RESIDENCIAL|DESARROLLO|CONJUNTO)[:\s"«]+([A-ZÁÉÍÓÚÑ][^"»\n,.]{3,60})/i)
    || str(/(?:torre|edificio|residencial|desarrollo|conjunto|plaza)\s+([A-ZÁÉÍÓÚÑ][a-záéíóúñA-Z\s]{2,40})/i)
    // 2. Nombre entre comillas en el encabezado
    || str(/[""«]([A-ZÁÉÍÓÚÑ][^""»\n]{4,60})[""»]/i)
    // 3. Fallback — objeto del contrato
    || str(/(?:proyecto\s+denominado|obra\s+denominada?|proyecto[:\s"«]+)["«]?([^"»\n,]{5,80})["»]?/i)
    || str(/objeto\s+del\s+contrato[:\s]+([^.\n]{10,80})/i);

  const direccionObra = str(/(?:ubicad[ao]|localizado|sito\s+en|ejecutar[aá]\s+en)[:\s]+([^.\n]{15,120})/i)
    || str(/domicilio\s+de\s+la\s+obra[:\s]+([^.\n]{10,100})/i);

  // Razón social cliente — empresa que aparece como CONTRATANTE
  const clienteMatch = t.match(/([A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑa-záéíóúñ\s,.]{5,80}(?:S\.A\.?|S\.A\.\s*DE\s*C\.V\.?|S\.\s*de\s*R\.L\.|A\.C\.)(?:\s*de\s*C\.V\.?)?)[,;\s]+(?:en\s+adelante|denominad[ao]|a\s+quien)[^,;]*(?:el\s+)?(?:contratante|cliente)/i);
  const razonSocialCliente = clienteMatch?.[1]?.trim() || str(/(?:contratante|cliente)[:\s]+([A-ZÁÉÍÓÚÑ][^.\n;,]{10,80}(?:S\.A|S\.\s*de\s*R|A\.C))/i);

  const repLegalCliente = str(/(?:representante\s+legal|apoderado\s+legal|representada\s+por)[:\s]+(?:el\s+)?(?:c\.?|sr\.?|lic\.?|ing\.?|arq\.?)?\s*([A-ZÁÉÍÓÚÑ][a-záéíóúñ]+(?: [A-ZÁÉÍÓÚÑ][a-záéíóúñ]+){1,4})/i);

  const rfcMatch = t.match(/\b([A-ZÑ&]{3,4}\d{6}[A-Z0-9]{3})\b/);
  const rfcCliente = rfcMatch?.[1] || null;

  const empresaSupervisora = str(/(?:empresa\s+supervisora|quien\s+supervisar[aá]|a\s+cargo\s+de\s+la\s+supervisi[oó]n)[:\s]+([A-ZÁÉÍÓÚÑ][^.\n,;]{5,60})/i);
  const repSupervisora = str(/(?:supervisor\s+designado|representante\s+de\s+supervisi[oó]n)[:\s]+(?:el\s+)?([A-ZÁÉÍÓÚÑ][a-záéíóúñ]+(?: [A-ZÁÉÍÓÚÑ][a-záéíóúñ]+){1,4})/i);

  // ── SECCIÓN 3: Estimaciones y pagos ───────────────────────────────────────
  let periodicidadEstimaciones = null;
  if (/estimaciones\s+semanales|cada\s+semana/i.test(t)) periodicidadEstimaciones = 'Semanal';
  else if (/estimaciones\s+quincenales|cada\s+quince/i.test(t)) periodicidadEstimaciones = 'Quincenal';
  else if (/estimaciones\s+mensuales|cada\s+mes/i.test(t)) periodicidadEstimaciones = 'Mensual';

  const diasRevisionSupervisora    = num(/(\d+)\s+d[ií]as?\s+h[aá]biles?[^.]*(?:revisi[oó]n|supervisora)/i)
    || num(/supervisora?\s+(?:tendr[aá]|contar[aá])[^.]*(\d+)\s+d[ií]as/i);
  const diasAprobacionContratante  = num(/(\d+)\s+d[ií]as?\s+h[aá]biles?[^.]*(?:aprobaci[oó]n|contratante)/i)
    || num(/contratante\s+(?:tendr[aá]|contar[aá])[^.]*(\d+)\s+d[ií]as/i);

  const clabeMatch = t.match(/\b(\d{18})\b/);
  const clabe = clabeMatch?.[1] || null;
  const BANCOS = ['BANREGIO','BBVA','BANAMEX','CITIBANAMEX','SANTANDER','HSBC','BANORTE','SCOTIABANK','INBURSA'];
  const banco = BANCOS.find(b => t.toUpperCase().includes(b)) || null;
  const numeroCuenta = str(/(?:cuenta\s+(?:n[uú]mero|no\.?)|n[uú]mero\s+de\s+cuenta)[:\s]+(\d{10,16})/i);

  // ── SECCIÓN 4: Garantías ───────────────────────────────────────────────────
  const fianzaAnticipoPct        = num(/fianza\s+de\s+anticipo[^%\d]*(\d+(?:\.\d+)?)\s*%/i) || (anticipoPct ? 100 : null);
  const fianzaCumplimientoPct    = num(/fianza\s+de\s+cumplimiento[^%\d]*(\d+(?:\.\d+)?)\s*%/i);
  const fianzaPasivosPct         = num(/(?:pasivos\s+contingentes|responsabilidad\s+laboral)[^%\d]*(\d+(?:\.\d+)?)\s*%/i);
  const fianzaPenasPct           = num(/(?:penas\s+convencionales)[^%\d]*(\d+(?:\.\d+)?)\s*%/i);
  const seguroRCPct              = num(/responsabilidad\s+civil[^%\d]*(\d+(?:\.\d+)?)\s*%/i);
  const viciosMatch              = t.match(/vicios\s+ocultos[^\d]*(\d+)\s*(meses|a[ñn]os)/i);
  const garantiaViciosOcultosMeses = viciosMatch
    ? parseInt(viciosMatch[1]) * (viciosMatch[2].startsWith('a') ? 12 : 1)
    : null;

  return {
    tipoContrato, montoContrato, anticipoPct, retencionPct,
    penalizacionDiaPct, penalRescisionPct,
    fechaFirma, fechaInicio, fechaTerminacion, plazoEjecucionDias,
    folioContrato, nombreProyecto, direccionObra,
    razonSocialCliente, repLegalCliente, rfcCliente,
    empresaSupervisora, repSupervisora,
    periodicidadEstimaciones, diasRevisionSupervisora, diasAprobacionContratante,
    banco, numeroCuenta, clabe,
    fianzaAnticipoPct, fianzaCumplimientoPct, fianzaPasivosPct,
    fianzaPenasPct, seguroRCPct, garantiaViciosOcultosMeses, topeMaxPenalizacionPct
  };
}

// ── Parser: Cotización Waller (formato propio) ───────────────────────────────
function parseCotizacionWaller(filePath) {
  const wb = XLSX.readFile(filePath, { cellText: false });

  const toNum = v => { const n = parseFloat(String(v||'').replace(/,/g,'')); return isFinite(n) ? n : 0; };

  // ── Hoja "Waller": paneles por sección ────────────────────────────────────
  const wsW = wb.Sheets['Waller'] || wb.Sheets[wb.SheetNames[0]];
  const rowsW = XLSX.utils.sheet_to_json(wsW, { header:1, defval:'' });

  const secciones = [];
  let seccionActual = null;

  for (let i = 0; i < rowsW.length; i++) {
    const r = rowsW[i];
    const col0 = String(r[0]||'').trim();
    const col1 = String(r[1]||'').trim();

    // Detectar encabezado de sección (ej. "Muros Divisorios", "Muros Fachadas")
    if (/^muros\s+(divisorios|interiores|fachadas|detalles|losas)/i.test(col1) && !col0) {
      seccionActual = { seccion: col1, conceptos: [], totalM2: 0, totalImporte: 0 };
      secciones.push(seccionActual);
      continue;
    }

    // Detectar fila de datos: col0 es número de código, col4 es importe > 0
    if (seccionActual && /^\d+$/.test(col0) && col1 && toNum(r[4]) > 0) {
      const m2      = toNum(r[2]);
      const precioM2 = toNum(r[3]);
      const importe  = toNum(r[4]);
      seccionActual.conceptos.push({ codigo: col0, concepto: col1, m2, precioM2, importe });
      seccionActual.totalM2      += m2;
      seccionActual.totalImporte += importe;
    }

    // Fila de total de sección
    if (/total/i.test(col1) && toNum(r[4]) > 0 && seccionActual) {
      seccionActual.totalImporte = toNum(r[4]); // usar el total calculado en el archivo
    }
  }

  // ── Hoja "COM": complementarios ──────────────────────────────────────────
  const complementarios = [];
  const wsC = wb.Sheets['COM'];
  if (wsC) {
    const rowsC = XLSX.utils.sheet_to_json(wsC, { header:1, defval:'' });
    let headerFila = -1;
    for (let i = 0; i < Math.min(10, rowsC.length); i++) {
      if (/CONCEPTO/i.test(String(rowsC[i][1]||''))) { headerFila = i; break; }
    }
    if (headerFila >= 0) {
      for (let i = headerFila + 1; i < rowsC.length; i++) {
        const r = rowsC[i];
        const concepto  = String(r[1]||'').trim();
        const unidad    = String(r[2]||'').trim();
        const cantidad  = toNum(r[3]);
        const pu        = toNum(r[4]);
        const importe   = toNum(r[5]);
        if (concepto && importe > 0) {
          complementarios.push({ concepto, unidad, cantidad, precioUnitario: pu, importe });
        }
      }
    }
  }

  // ── Hojas de costos adicionales (MO, FL, indirectos, utilidad) ───────────
  const resumenCostos = {};
  const hojasEspeciales = {
    'M O': 'manoDeObra', 'FL': 'fletes',
    'Ind of cen (2)': 'indirectosOficina', 'In cam (2)': 'indirectosCampo',
    'Uti y Res (2)': 'utilidadReserva'
  };
  for (const [hoja, clave] of Object.entries(hojasEspeciales)) {
    const ws = wb.Sheets[hoja];
    if (!ws) continue;
    const rows = XLSX.utils.sheet_to_json(ws, { header:1, defval:'' });
    // Buscar el total más grande en las últimas filas
    let maxVal = 0;
    for (const r of rows.slice(-10)) {
      for (const cell of r) { const v = toNum(cell); if (v > maxVal) maxVal = v; }
    }
    if (maxVal > 0) resumenCostos[clave] = maxVal;
  }

  // ── Totales ───────────────────────────────────────────────────────────────
  const totalPaneles     = secciones.reduce((s, sec) => s + sec.totalImporte, 0);
  const totalCompl       = complementarios.reduce((s, c) => s + c.importe, 0);
  const totalPresupuesto = totalPaneles + totalCompl + Object.values(resumenCostos).reduce((s,v)=>s+v,0);
  const totalM2          = secciones.reduce((s, sec) => s + sec.totalM2, 0);

  return {
    tipo: 'cotizacion',
    secciones,
    complementarios,
    resumenCostos,
    totalPaneles:     Math.round(totalPaneles * 100) / 100,
    totalComplementarios: Math.round(totalCompl * 100) / 100,
    totalPresupuesto: Math.round(totalPresupuesto * 100) / 100,
    totalM2:          Math.round(totalM2 * 100) / 100
  };
}

// ── Auto-detect: detección de tipo por contenido ────────────────────────────
function detectarTipoContenido(texto, headers) {
  const all = ((texto || '') + ' ' + (headers || []).join(' ')).toLowerCase();
  const scores = { contpaq:0, cronograma:0, estimacion:0, preciosUnitarios:0, cotizacion:0, avanceDiario:0, contrato:0, cambioAlcance:0 };

  // Cotización Waller — formato propio con paneles
  if (/panel\s+waller/i.test(all)) scores.cotizacion += 60;
  if (/muros\s+(divisorios|interiores|fachadas)/i.test(all)) scores.cotizacion += 40;
  if (/\$\s*x\s*m²|\$\s*x\s*m2/i.test(all)) scores.cotizacion += 30;
  if (/tc\s*(div|fac)|tarjeta.*costo/i.test(all)) scores.cotizacion += 20;

  // ContPAQ — números de cuenta 5XX-XX-XXX
  if (/\b5\d{2}-\d{2}-\d{3}\b/.test(all)) scores.contpaq += 60;
  if (/\b(cargo|abono|debe|haber|auxiliar)\b/.test(all)) scores.contpaq += 15;

  // Cronograma — semanas + niveles
  if (/\bsemana\s*\d+|\bs\d{1,2}\b/.test(all)) scores.cronograma += 25;
  if (/\b(npb|n\d{1,2}|naz|azotea|s[oó]tano)\b/.test(all)) scores.cronograma += 20;
  if (/\bcronograma\b|\bprograma\b/.test(all)) scores.cronograma += 20;

  // Estimación — keywords financieros de estimación
  if (/estimaci[oó]n/.test(all)) scores.estimacion += 35;
  if (/\b(per[ií]odo|retenci[oó]n|importe a pagar|neto a pagar)\b/.test(all)) scores.estimacion += 20;
  if (/acumulado anterior/.test(all)) scores.estimacion += 25;

  // Precios Unitarios (genérico, no Waller)
  if (/precio\s+unitario|p\.u\b/.test(all)) scores.preciosUnitarios += 35;
  if (/\b(partida|concepto|cantidad|importe)\b/.test(all)) scores.preciosUnitarios += 10;
  if (/presupuesto\s+(de\s+)?obra/.test(all)) scores.preciosUnitarios += 20;

  // Avance Diario — niveles + m2 + fechas
  if (/avance\s+diario|m2\s+instalado|instalado.*m2/.test(all)) scores.avanceDiario += 35;
  if (/\b(lunes|martes|mi[eé]rcoles|jueves|viernes|s[aá]bado)\b/.test(all)) scores.avanceDiario += 15;
  if (/\b(npb|n\d{1,2})\b/.test(all) && /\bm2\b/.test(all)) scores.avanceDiario += 20;

  // Contrato
  if (/\bcontrato\b/.test(all)) scores.contrato += 20;
  if (/\b(anticipo|penalizaci[oó]n|contratante|contratista)\b/.test(all)) scores.contrato += 25;
  if (/precio\s+fijo|plazo\s+de\s+ejecuci[oó]n/.test(all)) scores.contrato += 20;

  // Cambio de alcance (deductiva / aditiva)
  if (/\b(deductiv[ao]|aditiv[ao]|adicional)\b/.test(all)) scores.cambioAlcance += 30;
  if (/cambio\s+de\s+alcance|orden\s+de\s+cambio/.test(all)) scores.cambioAlcance += 30;
  if (/\b(aprobado|pendiente|negociaci[oó]n)\b/.test(all)) scores.cambioAlcance += 15;

  const max = Math.max(...Object.values(scores));
  if (max === 0) return { tipo: null, confianza: 'baja' };
  const tipo = Object.entries(scores).sort((a,b)=>b[1]-a[1])[0][0];
  return { tipo, confianza: max >= 35 ? 'alta' : 'baja' };
}

// Extrae preview (primeras 3 filas como array de objetos)
function extraerPreview(rows, maxRows = 3) {
  if (!rows || rows.length < 2) return [];
  const headers = rows[0].map(h => String(h || '').trim()).filter(Boolean);
  return rows.slice(1, maxRows + 1).map(row =>
    Object.fromEntries(headers.map((h, i) => [h, String(row[i] ?? '')]))
  );
}

// Extrae texto y filas de un Excel para detección
function leerExcelParaDeteccion(filePath) {
  const wb = XLSX.readFile(filePath, { cellText: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  const headers = (rows[0] || []).map(h => String(h || ''));
  const texto = rows.slice(0, 20).map(r => r.join(' ')).join('\n');
  return { rows, headers, texto };
}

// Extrae texto de PDF: pdf-parse para digitales, pdftoppm+tesseract CLI para escaneados
async function extraerTextoPDF(filePath, maxPaginas = 5) {
  // 1. Texto nativo (PDFs digitales — instantáneo)
  try {
    const pdfParse = require('pdf-parse');
    const data = await pdfParse(fs.readFileSync(filePath));
    if (data.text && data.text.replace(/\s/g, '').length > 80) return data.text;
  } catch(e) {}

  // 2. OCR con binarios del sistema (pdftoppm + tesseract CLI)
  const os = require('os');
  const { execFile } = require('child_process');
  const { promisify } = require('util');
  const exec = promisify(execFile);
  const tmpDir = path.join(os.tmpdir(), 'waller-ocr-' + Date.now());
  try {
    fs.mkdirSync(tmpDir, { recursive: true });
    const prefix = path.join(tmpDir, 'pag');

    await exec('pdftoppm', ['-png', '-r', '200', '-l', String(maxPaginas), filePath, prefix]);

    const imagenes = fs.readdirSync(tmpDir)
      .filter(f => f.endsWith('.png'))
      .sort()
      .map(f => path.join(tmpDir, f));

    if (!imagenes.length) return '';

    let texto = '';
    for (const img of imagenes) {
      const { stdout } = await exec('tesseract', [img, 'stdout', '-l', 'spa']);
      texto += stdout + '\n';
    }
    return texto.trim();
  } catch(e) {
    return '';
  } finally {
    try { fs.rmSync(tmpDir, { recursive: true, force: true }); } catch(e) {}
  }
}

// Extrae datos estructurados de PDF por tipo usando regex
function parsearPdfPorTipo(texto, tipo) {
  const t = texto;
  const num = s => { const m = s?.match(/[\d,]+\.?\d*/); return m ? parseFloat(m[0].replace(/,/g,'')) : null; };
  const fecha = s => { const m = s?.match(/\d{1,2}[\/\-]\w{2,9}[\/\-]\d{2,4}/); return m ? m[0] : null; };

  if (tipo === 'contrato') {
    const re = (pattern, flags='i') => { try { return new RegExp(pattern, flags); } catch(e) { return null; } };
    const find = (pattern, group=1) => { const m = t.match(re(pattern)); return m ? m[group]?.trim() || null : null; };
    const findNum = pattern => { const v = find(pattern); return v ? parseFloat(v.replace(/,/g,'')) : null; };

    // Monto — el más grande > 100,000 que aparece cerca de palabras clave
    const montoMatches = [...t.matchAll(/(?:monto|importe|cantidad|valor)[^$\d]*\$?\s*([\d,]+(?:\.\d{2})?)/gi)]
      .map(m => parseFloat(m[1].replace(/,/g,'')))
      .filter(v => v > 50000);
    // Fallback: todos los $ > 100K
    const todosDolares = [...t.matchAll(/\$\s*([\d,]+(?:\.\d{2})?)/g)]
      .map(m => parseFloat(m[1].replace(/,/g,''))).filter(v => v > 100000);
    const montos = [...new Set([...montoMatches, ...todosDolares])].sort((a,b)=>b-a);

    // Fechas en formatos: DD/MM/YYYY, DD-MM-YYYY, "día X del mes Y del año Z"
    const parseFechaES = s => {
      if (!s) return null;
      const meses = {enero:'01',febrero:'02',marzo:'03',abril:'04',mayo:'05',junio:'06',julio:'07',agosto:'08',septiembre:'09',octubre:'10',noviembre:'11',diciembre:'12'};
      // Día X del mes de MMMM del año YYYY
      const m1 = s.match(/(\d{1,2})[^\w]*(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre)[^\d]*(\d{4})/i);
      if (m1) return `${m1[3]}-${meses[m1[2].toLowerCase()]}-${m1[1].padStart(2,'0')}`;
      // DD/MM/YYYY o DD-MM-YYYY
      const m2 = s.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
      if (m2) return `${m2[3]}-${m2[2].padStart(2,'0')}-${m2[1].padStart(2,'0')}`;
      return null;
    };

    const fechaFirmaRaw   = find('(?:fecha[^:]*firma|firma[^:]*fecha|suscrito|celebrado)[^:]*?([\\d]{1,2}[^\\w]+(?:enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre)[^\\d]*[\\d]{4})');
    const fechaIniRaw     = find('(?:fecha[^:]*inicio|inicio[^:]*obra|fecha[^:]*arranque)[^:]*?([\\d]{1,2}[^\\w]+(?:enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre)[^\\d]*[\\d]{4})');
    const fechaTermRaw    = find('(?:fecha[^:]*terminaci[oó]n|fecha[^:]*entrega|plazo.*?hasta)[^:]*?([\\d]{1,2}[^\\w]+(?:enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre)[^\\d]*[\\d]{4})');

    // RFC mexicano
    const rfcMatch = t.match(/\b([A-ZÑ&]{3,4}\d{6}[A-Z0-9]{3})\b/);

    // CLABE — 18 dígitos
    const clabeMatch = t.match(/\b(\d{18})\b/);

    // Banco conocido
    const bancos = ['BANREGIO','BBVA','BANAMEX','CITIBANAMEX','SANTANDER','HSBC','BANORTE','SCOTIABANK','INBURSA'];
    const bancoDetect = bancos.find(b => t.toUpperCase().includes(b)) || null;

    // Plazo en días calendario
    const plazoMatch = t.match(/plazo[^:]*?(\d+)\s*d[ií]as\s+calendario/i) ||
                       t.match(/(\d+)\s*d[ií]as\s+calendario/i);

    // Folio / número contrato
    const folioMatch = t.match(/(?:folio|n[uú]mero\s*de\s*contrato|contrato\s*n[oúu])[^\w]*([A-Z0-9][\w\-\/]{3,20})/i);

    // Tope máximo penalización
    const topeMatch = t.match(/tope[^%\d]*(\d+(?:\.\d+)?)\s*%/i) ||
                      t.match(/hasta\s+(?:un\s+máximo\s+de\s+)?(\d+(?:\.\d+)?)\s*%/i);

    // Vicios ocultos
    const viciosMatch = t.match(/vicios\s+ocultos[^\d]*(\d+)\s*(meses|a[ñn]os)/i);
    let viciosMeses = null;
    if (viciosMatch) {
      viciosMeses = parseInt(viciosMatch[1]) * (viciosMatch[2].startsWith('a') ? 12 : 1);
    }

    // Razón social cliente — empresa antes de "CONTRATANTE" o "CLIENTE"
    const clienteMatch = t.match(/([A-ZÁÉÍÓÚa-záéíóúñÑ][^.\n]{10,80}?)\s*(?:en\s+adelante|denominad[ao])[^"]*(?:contratante|cliente)/i) ||
                         t.match(/(?:contratante|cliente)[:\s]+([A-ZÁÉÍÓÚ][^.\n,;]{5,80}S\.?A\.?\s*(?:de\s*C\.?V\.?|P\.?I\.?|B\.?)?)/i);

    // Representante legal
    const repMatch = t.match(/(?:representante\s+legal|apoderado\s+legal)[^,.\n]*?\b([A-ZÁÉÍÓÚa-záéíóúñÑ][a-záéíóúñÑ]+(?:\s+[A-ZÁÉÍÓÚa-záéíóúñÑ][a-záéíóúñÑ]+){1,4})/i);

    // Supervisor / empresa supervisora
    const supervisorMatch = t.match(/(?:empresa\s+supervisora|supervisor[a]?\s+designad[ao])[:\s]+([A-ZÁÉÍÓÚa-záéíóúñÑ][^.\n,;]{5,60})/i);

    return {
      montoContrato:     montos[0] || null,
      anticipoPct:       findNum('anticipo[^\\d%]*([\\d]+(?:\\.\\d+)?)\\s*%'),
      retencionPct:      findNum('(?:retenci[oó]n|fondo\\s+de\\s+garant[ií]a)[^\\d%]*([\\d]+(?:\\.\\d+)?)\\s*%'),
      penalizacionPct:   findNum('(?:pena\\s+convencional|penalizaci[oó]n)[^\\d%]*([\\d]+(?:\\.\\d+)?)\\s*%'),
      penalRescisionPct: findNum('rescisi[oó]n[^\\d%]*([\\d]+(?:\\.\\d+)?)\\s*%'),
      fechaFirma:        parseFechaES(fechaFirmaRaw),
      fechaInicio:       parseFechaES(fechaIniRaw),
      fechaTerminacion:  parseFechaES(fechaTermRaw),
      plazoEjecucionDias: plazoMatch ? parseInt(plazoMatch[1]) : null,
      folioContrato:     folioMatch?.[1]?.trim() || null,
      cliente:           clienteMatch?.[1]?.trim() || null,
      representanteLegal: repMatch?.[1]?.trim() || null,
      rfcCliente:        rfcMatch?.[1] || null,
      clabe:             clabeMatch?.[1] || null,
      banco:             bancoDetect,
      empresaSupervisora: supervisorMatch?.[1]?.trim() || null,
      garantiaMeses:     viciosMeses,
      topeMaxPenalizacionPct: topeMatch ? parseFloat(topeMatch[1]) : null,
    };
  }
  if (tipo === 'preciosUnitarios') {
    // Buscar líneas con número | descripción | cantidad | PU | importe
    const lineas = t.split('\n').filter(l => /\d/.test(l) && l.length > 10);
    const totalMatch = t.match(/total[^\d]*\$?\s*([\d,]+\.?\d*)/i);
    return { totalPresupuesto: totalMatch ? parseFloat(totalMatch[1].replace(/,/g,'')) : null, lineasDetectadas: lineas.length, _textoParcial: t.slice(0,500) };
  }
  return { _textoParcial: t.slice(0, 600) };
}

// ── Helpers Google Sheets ────────────────────────────────────────────────────
function gsheetsToCsvUrl(url) {
  // Acepta: /edit, /pub, /pubhtml, o URL de exportación directa
  const idMatch = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  if (!idMatch) return null;
  const id = idMatch[1];
  const gidMatch = url.match(/[#&?]gid=(\d+)/);
  const gid = gidMatch ? gidMatch[1] : '0';
  return `https://docs.google.com/spreadsheets/d/${id}/export?format=csv&id=${id}&gid=${gid}`;
}

const NIVELES_CONOCIDOS = ['NPB','N1','N2','N3','N4','N5','N6','N7','N8','N9','N10','N11','N12','NAZ','S1','S2','S3','SOTANO','PB','PLANTA BAJA','AZOTEA'];

function parseAvanceDiarioCsv(csvText) {
  // Normaliza saltos de línea y split
  const lines = csvText.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n').filter(l => l.trim());
  if (lines.length < 2) return null;

  // Parse CSV respetando comillas
  const parseCsvLine = line => {
    const result = []; let cur = ''; let inQ = false;
    for (let i = 0; i < line.length; i++) {
      if (line[i] === '"') { inQ = !inQ; continue; }
      if (line[i] === ',' && !inQ) { result.push(cur.trim()); cur = ''; continue; }
      cur += line[i];
    }
    result.push(cur.trim());
    return result;
  };

  const headers = parseCsvLine(lines[0]).map(h => h.toUpperCase().trim());

  // ¿Los headers son niveles? (formato: Fecha, NPB, N1, N2...)
  const nvelCols = {};
  NIVELES_CONOCIDOS.forEach(nv => {
    const idx = headers.indexOf(nv);
    if (idx >= 0) nvelCols[nv] = idx;
  });

  // O formato: Fecha, Nivel, M2, Notas
  const fechaCol  = headers.findIndex(h => /^(FECHA|DATE|DÍA|DIA|SEMANA|WEEK)$/.test(h));
  const nvelCol   = headers.findIndex(h => /^(NIVEL|LEVEL|PISO|FLOOR)$/.test(h));
  const m2Col     = headers.findIndex(h => /^(M2|M²|METROS|INSTALADO|M2 INST|AVANCE)$/.test(h));
  const notasCol  = headers.findIndex(h => /^(NOTAS|NOTES|OBSERVACION|OBS)$/.test(h));

  const entradas = [];
  const m2PorNivel = {};

  for (let i = 1; i < lines.length; i++) {
    const vals = parseCsvLine(lines[i]);
    if (vals.every(v => !v)) continue;

    if (Object.keys(nvelCols).length >= 2) {
      // Formato columnar: cada columna es un nivel
      const fecha = fechaCol >= 0 ? vals[fechaCol] : (vals[0] || '');
      if (!fecha) continue;
      const entrada = { fecha };
      let hayDatos = false;
      NIVELES_CONOCIDOS.forEach(nv => {
        if (nvelCols[nv] === undefined) return;
        const raw = vals[nvelCols[nv]] || '';
        const v = parseFloat(raw.replace(',', '.')) || 0;
        if (v > 0) { entrada[nv] = v; m2PorNivel[nv] = (m2PorNivel[nv] || 0) + v; hayDatos = true; }
      });
      if (hayDatos) entradas.push(entrada);

    } else if (nvelCol >= 0 && m2Col >= 0) {
      // Formato por fila: Fecha | Nivel | M²
      const fecha  = fechaCol >= 0 ? vals[fechaCol] : (vals[0] || '');
      const nivel  = (vals[nvelCol] || '').trim().toUpperCase();
      const m2     = parseFloat((vals[m2Col] || '').replace(',', '.')) || 0;
      const notas  = notasCol >= 0 ? (vals[notasCol] || '') : '';
      const nvMatch = NIVELES_CONOCIDOS.find(nv => nivel === nv || nivel.startsWith(nv));
      if (nvMatch && m2 > 0) {
        entradas.push({ fecha, nivel: nvMatch, m2, ...(notas ? { notas } : {}) });
        m2PorNivel[nvMatch] = (m2PorNivel[nvMatch] || 0) + m2;
      }
    }
  }

  if (!entradas.length) return null;
  return {
    entradas,
    m2PorNivel,
    totalM2Instalado: Math.round(Object.values(m2PorNivel).reduce((s, v) => s + v, 0) * 100) / 100
  };
}

// ── Helpers PDF / Gemini ─────────────────────────────────────────────────────
function promptParaPdf(tipo, cc, nombreObra) {
  const prompts = {
    contrato: `Eres un asistente de análisis de contratos de construcción. Lee este contrato de obra y extrae en JSON:
{
  "montoContrato": número sin comas,
  "cliente": "nombre del cliente o contratante",
  "contratista": "nombre del contratista",
  "fechaFirma": "YYYY-MM-DD",
  "fechaInicioObra": "YYYY-MM-DD o null",
  "fechaTerminacion": "YYYY-MM-DD o null",
  "plazoSemanas": número o null,
  "anticipoPct": número (porcentaje del anticipo) o null,
  "retencionPct": número o null,
  "penalizacionPct": número por día o null,
  "garantiaMeses": número o null,
  "representanteLegal": "nombre" o null,
  "obra": "${nombreObra} CC ${cc}",
  "notas": "cualquier condición importante en máx 200 caracteres"
}
Solo devuelve el JSON, sin explicaciones.`,

    preciosUnitarios: `Eres un asistente de análisis de presupuestos de obra. Lee este documento de precios unitarios y extrae en JSON:
{
  "conceptos": [
    {
      "descripcion": "descripción breve del concepto",
      "unidad": "m², ml, pza, etc.",
      "cantidad": número,
      "precioUnitario": número,
      "importe": número
    }
  ],
  "totalPresupuesto": número,
  "moneda": "MXN",
  "notas": "condiciones relevantes en máx 150 caracteres"
}
Solo devuelve el JSON.`,

    estimacion: `Eres un asistente de análisis de estimaciones de obra. Lee esta estimación y extrae en JSON:
{
  "numero": número de estimación,
  "periodoDesde": "YYYY-MM-DD",
  "periodoHasta": "YYYY-MM-DD",
  "montoEstimacion": número,
  "retencion": número,
  "importePagar": número,
  "m2Esta": número de m² estimados,
  "m2Acum": número de m² acumulados,
  "cliente": "nombre",
  "notas": "máx 150 caracteres"
}
Solo devuelve el JSON.`,

    avanceDiario: `Eres un asistente de seguimiento de obra. Lee este documento de avance de obra y extrae en JSON:
{
  "entradas": [
    { "fecha": "YYYY-MM-DD", "nivel": "NPB/N1/N2/...", "m2": número, "notas": "opcional" }
  ],
  "m2PorNivel": { "NPB": número, "N1": número, ... },
  "totalM2Instalado": número
}
Solo incluye niveles con m² > 0. Solo devuelve el JSON.`
  };
  return prompts[tipo] || `Lee este documento PDF de la obra ${nombreObra} CC ${cc} y extrae la información relevante en formato JSON. Solo devuelve el JSON.`;
}

function aplicarPdfAlProyecto(proyecto, tipo, extraido) {
  switch (tipo) {
    case 'contrato':
      return {
        ...proyecto,
        ...(extraido.montoContrato ? { montoContrato: extraido.montoContrato } : {}),
        contrato: { ...(proyecto.contrato || {}), ...extraido }
      };
    case 'preciosUnitarios':
      return { ...proyecto, preciosUnitarios: extraido };
    case 'estimacion':
      if (extraido.numero) {
        const ests = proyecto.estimaciones || [];
        const idx = ests.findIndex(e => e.numero === extraido.numero);
        if (idx >= 0) ests[idx] = { ...ests[idx], ...extraido };
        else { ests.push(extraido); ests.sort((a, b) => a.numero - b.numero); }
        return { ...proyecto, estimaciones: ests };
      }
      return proyecto;
    case 'avanceDiario':
      return {
        ...proyecto,
        _avanceDiario: extraido.entradas || [],
        _m2InstalPorNivel: extraido.m2PorNivel || {},
        _totalM2Instalado: extraido.totalM2Instalado || 0,
        _ultimaActAvance: new Date().toISOString().slice(0, 10)
      };
    default:
      return { ...proyecto, [`_pdf_${tipo}`]: extraido };
  }
}

module.exports = function(app) {

  app.get('/api/proyectos', (req, res) => {
    try {
      if (!fs.existsSync(DATA_DIR)) return res.json([]);
      const files = fs.readdirSync(DATA_DIR).filter(f => f.endsWith('.json'));
      const proyectos = files.map(f => {
        try { return JSON.parse(fs.readFileSync(path.join(DATA_DIR, f), 'utf8')); }
        catch(e) { return null; }
      }).filter(Boolean);
      res.json(proyectos);
    } catch(e) { res.status(500).json({ error: e.message }); }
  });

  app.get('/api/proyectos/:cc', (req, res) => {
    const p = leerProyecto(req.params.cc);
    if (!p) return res.status(404).json({ error: 'Proyecto no encontrado. CC: ' + req.params.cc });
    res.json(p);
  });

  app.post('/api/proyectos/:cc', (req, res) => {
    const { cc } = req.params;
    const existente = leerProyecto(cc) || {};
    const data = { ...existente, ...req.body, cc };
    guardarProyecto(cc, data);
    res.json({ ok: true, data });
  });

  // ── Endpoints modulares por dominio ────────────────────────────────────────

  app.get('/api/proyectos/:cc/resumen', (req, res) => {
    const p = leerProyecto(req.params.cc);
    if (!p) return res.status(404).json({ error: 'CC no encontrado: ' + req.params.cc });
    const ests = p.estimaciones || [];
    const ultima = ests[ests.length - 1];
    res.json({
      cc: p.cc,
      nombre: p.nombre,
      cliente: p.cliente || null,
      montoContrato: p.montoContrato || 0,
      totalEgresado: p.totalEgresado || 0,
      totalFacturado: ultima ? (ultima.acumActual || 0) : 0,
      avanceFisico: ultima ? (ultima.m2Acum || 0) : 0,
      estimaciones: ests.length,
      fechaInicio: p.fechaInicio || null,
      fechaFin: p.fechaFin || null,
      ultimaActualizacion: p.ultimaActualizacion || null,
    });
  });

  app.get('/api/proyectos/:cc/estimaciones', (req, res) => {
    const p = leerProyecto(req.params.cc);
    if (!p) return res.status(404).json({ error: 'CC no encontrado: ' + req.params.cc });
    res.json(p.estimaciones || []);
  });

  app.get('/api/proyectos/:cc/cambios', (req, res) => {
    const p = leerProyecto(req.params.cc);
    if (!p) return res.status(404).json({ error: 'CC no encontrado: ' + req.params.cc });
    res.json(p.cambios || []);
  });

  app.get('/api/proyectos/:cc/cronograma', (req, res) => {
    const p = leerProyecto(req.params.cc);
    if (!p) return res.status(404).json({ error: 'CC no encontrado: ' + req.params.cc });
    const cron = p.cronograma || {};
    res.json({
      m2PorNivel: cron.m2PorNivel || {},
      programaSemanal: cron.programaSemanal || {},
      totalM2: cron.totalM2 || 0,
      totalSemanas: cron.totalSemanas || 0,
      fechaInicio: cron.fechaInicio || p.fechaInicio || null,
    });
  });

  app.get('/api/proyectos/:cc/contrato', (req, res) => {
    const p = leerProyecto(req.params.cc);
    if (!p) return res.status(404).json({ error: 'CC no encontrado: ' + req.params.cc });
    const c = p.contrato?.campos || {};
    const campo = (key, fallback) => p[key] ?? c[key]?.valor ?? fallback ?? null;
    res.json({
      cliente:           campo('cliente', c.repLegalCliente?.valor || null),
      rfcCliente:        campo('rfcCliente'),
      montoContrato:     campo('montoContrato'),
      tipoContrato:      campo('tipoContrato'),
      anticipoPct:       campo('anticipoPct'),
      retencionPct:      campo('retencionPct'),
      fechaInicio:       campo('fechaInicio'),
      fechaFin:          p.fechaFin || c.fechaTerminacion?.valor || null,
      plazoEjecucionDias:campo('plazoEjecucionDias'),
      banco:             campo('banco'),
      clabe:             campo('clabe'),
      archivoPdf:        p.contrato?.archivoPdf || null,
    });
  });

  app.get('/api/proyectos/:cc/variaciones', (req, res) => {
    const p = leerProyecto(req.params.cc);
    if (!p) return res.status(404).json({ error: 'CC no encontrado: ' + req.params.cc });
    const partidas = (p.partidasDetalle || []).map(part => {
      const presupuestado = part.presupuesto || 0;
      const ejercido = part.ejercido || part.total || 0;
      const variacion = presupuestado > 0 ? ejercido - presupuestado : null;
      const pctEjecutado = presupuestado > 0 ? (ejercido / presupuestado * 100) : null;
      return { ...part, presupuestado, ejercido, variacion, pctEjecutado };
    });
    res.json(partidas);
  });

  app.get('/api/proyectos/:cc/flujo-caja', (req, res) => {
    const p = leerProyecto(req.params.cc);
    if (!p) return res.status(404).json({ error: 'CC no encontrado: ' + req.params.cc });
    const ests = p.estimaciones || [];
    const porMes = {};
    ests.forEach(e => {
      const mes = (e.periodoHasta || e.periodoDesde || '').slice(0, 7) || 'sin-fecha';
      if (!porMes[mes]) porMes[mes] = { mes, ingresos: 0, estimaciones: [] };
      porMes[mes].ingresos += e.estaEstimacion || 0;
      porMes[mes].estimaciones.push(e.numero);
    });
    res.json({
      porMes: Object.values(porMes),
      totalEgresado: p.totalEgresado || 0,
      ultimaActualizacion: p.ultimaActualizacion || null,
    });
  });

  app.get('/api/proyectos/:cc/bitacora', (req, res) => {
    const p = leerProyecto(req.params.cc);
    if (!p) return res.status(404).json({ error: 'CC no encontrado: ' + req.params.cc });
    res.json(p.bitacora || []);
  });

  app.post('/api/proyectos/:cc/bitacora', (req, res) => {
    const { cc } = req.params;
    const p = leerProyecto(cc);
    if (!p) return res.status(404).json({ error: 'CC no encontrado: ' + cc });
    const entrada = req.body;
    if (!entrada || !entrada.fecha || !entrada.texto)
      return res.status(400).json({ error: 'Se requieren fecha y texto' });
    if (!entrada.id) entrada.id = Date.now().toString(36);
    const bitacora = p.bitacora || [];
    bitacora.unshift(entrada);
    p.bitacora = bitacora;
    guardarProyecto(cc, p);
    res.json({ ok: true, total: bitacora.length });
  });

  app.get('/api/corporativo/resumen', (req, res) => {
    try {
      if (!fs.existsSync(DATA_DIR)) return res.json({ totalObras: 0, listaProyectos: [] });
      const files = fs.readdirSync(DATA_DIR).filter(f => f.endsWith('.json'));
      const proyectos = files.map(f => {
        try { return JSON.parse(fs.readFileSync(path.join(DATA_DIR, f), 'utf8')); }
        catch(e) { return null; }
      }).filter(Boolean);

      let totalContrato = 0, totalEgresado = 0, totalFacturado = 0;
      const listaProyectos = proyectos.map(p => {
        const ests = p.estimaciones || [];
        const ultima = ests[ests.length - 1];
        const facturado = ultima ? (ultima.acumActual || 0) : 0;
        totalContrato  += p.montoContrato || 0;
        totalEgresado  += p.totalEgresado || 0;
        totalFacturado += facturado;
        return {
          cc: p.cc,
          nombre: p.nombre,
          cliente: p.cliente || null,
          montoContrato: p.montoContrato || 0,
          totalEgresado: p.totalEgresado || 0,
          totalFacturado: facturado,
          fechaInicio: p.fechaInicio || null,
          fechaFin: p.fechaFin || null,
          ultimaActualizacion: p.ultimaActualizacion || null,
        };
      });

      res.json({
        totalObras: proyectos.length,
        totalContrato,
        totalEgresado,
        totalFacturado,
        listaProyectos,
      });
    } catch(e) { res.status(500).json({ error: e.message }); }
  });

  // ── Eliminar proyecto completo ─────────────────────────────────────────────
  app.delete('/api/proyectos/:cc', (req, res) => {
    const { cc } = req.params;
    try {
      let eliminados = [];

      // 1. Eliminar JSON del proyecto
      const jsonPath = path.join(DATA_DIR, cc + '.json');
      if (fs.existsSync(jsonPath)) {
        fs.unlinkSync(jsonPath);
        eliminados.push(cc + '.json');
      }

      // 2. Eliminar archivos físicos del proyecto (contratos, estimaciones, etc.)
      const archivosDir = path.join(DATA_DIR, '..', 'uploads', cc);
      if (fs.existsSync(archivosDir)) {
        const files = fs.readdirSync(archivosDir);
        files.forEach(f => {
          try { fs.unlinkSync(path.join(archivosDir, f)); eliminados.push(f); } catch(e) {}
        });
        try { fs.rmdirSync(archivosDir); } catch(e) {}
        eliminados.push('carpeta uploads/' + cc);
      }

      // 3. Eliminar PDFs de contratos si existen
      const contratoPath = path.join(__dirname, '../uploads/contratos');
      if (fs.existsSync(contratoPath)) {
        const files = fs.readdirSync(contratoPath).filter(f => f.startsWith(cc + '-') || f.startsWith('contrato-' + cc));
        files.forEach(f => {
          try { fs.unlinkSync(path.join(contratoPath, f)); eliminados.push('contratos/' + f); } catch(e) {}
        });
      }

      res.json({ ok: true, cc, eliminados, mensaje: `Proyecto CC ${cc} eliminado correctamente` });
    } catch(e) {
      res.status(500).json({ error: 'Error al eliminar proyecto: ' + e.message });
    }
  });


  // Lee el listado de cuentas ContPAQ (.xlsx), detecta cuentas de obra (5XX-YY-ZZZ)
  // y las registra en un archivo catalogo.json global. Las nuevas se agregan,
  // las existentes se usan solo como referencia.
  app.post('/api/upload/catalogo', upload.single('archivo'), (req, res) => {
    if (!req.file) return res.status(400).json({ error: 'No se recibió ningún archivo' });
    try {
      const wb = XLSX.readFile(req.file.path, { cellText: true });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

      const catalogoPath = path.join(DATA_DIR, '../catalogo.json');
      let catalogoExistente = {};
      try { catalogoExistente = JSON.parse(fs.readFileSync(catalogoPath, 'utf8')); } catch(e) {}

      let cuentasNuevas = 0;
      let cuentasExistentes = 0;

      for (const row of rows) {
        const cuenta = String(row[0] || '').trim();
        const nombre = String(row[1] || '').trim();
        if (!cuenta || !nombre) continue;

        // Solo procesar cuentas de obra formato 5XX-YY-ZZZ o 5XXXXXXX
        const esObra = /^5\d{2}-\d{2}-\d{3}$/.test(cuenta) || /^5\d{7}$/.test(cuenta);
        if (!esObra) continue;

        if (catalogoExistente[cuenta]) {
          cuentasExistentes++;
        } else {
          catalogoExistente[cuenta] = { nombre, fechaAlta: new Date().toISOString().slice(0,10) };
          cuentasNuevas++;
        }
      }

      // Guardar catálogo actualizado
      if (!fs.existsSync(path.dirname(catalogoPath))) {
        fs.mkdirSync(path.dirname(catalogoPath), { recursive: true });
      }
      fs.writeFileSync(catalogoPath, JSON.stringify(catalogoExistente, null, 2), 'utf8');
      try { fs.unlinkSync(req.file.path); } catch(e) {}

      res.json({
        ok: true,
        cuentasNuevas,
        cuentasExistentes,
        totalCuentas: Object.keys(catalogoExistente).length,
        mensaje: `Catálogo actualizado: ${cuentasNuevas} cuentas nuevas, ${cuentasExistentes} ya existían`
      });
    } catch(e) {
      try { fs.unlinkSync(req.file.path); } catch(ex) {}
      res.status(500).json({ error: 'Error al procesar catálogo: ' + e.message });
    }
  });

  // GET catálogo completo
  app.get('/api/catalogo', (req, res) => {
    try {
      const catalogoPath = path.join(DATA_DIR, '../catalogo.json');
      if (!fs.existsSync(catalogoPath)) return res.json({});
      res.json(JSON.parse(fs.readFileSync(catalogoPath, 'utf8')));
    } catch(e) { res.status(500).json({ error: e.message }); }
  });

  app.post('/api/upload/contpaq', upload.single('archivo'), (req, res) => {
    if (!req.file) return res.status(400).json({ error: 'No se recibió ningún archivo' });
    try {
      const parsed = parseLibroAuxiliar(req.file.path);
      const obras = Object.values(parsed);
      const guardadas = [];
      for (const obra of obras) {
        if (!obra.cc) continue;
        const existente = leerProyecto(obra.cc) || {};
        const actualizado = {
          ...existente,
          cc: obra.cc,
          nombre: obra.nombre || existente.nombre || ('Obra ' + obra.cc),
          partidas: obra.partidas,
          partidasDetalle: obra.partidasDetalle,
          totalEgresado: obra.totalEgresado,
          ultimaActualizacion: obra.ultimaActualizacion
        };
        guardarProyecto(obra.cc, actualizado);
        guardadas.push({
          cc: obra.cc,
          nombre: actualizado.nombre,
          totalEgresado: obra.totalEgresado,
          partidas: obra.partidasDetalle.length,
          detalle: obra.partidasDetalle
        });
      }
      try { fs.unlinkSync(req.file.path); } catch(e) {}
      guardadas.forEach(o => actualizarMetaArchivo(o.cc, 'contpaq', req.file.originalname || req.file.filename));
      res.json({
        ok: true,
        obras: guardadas,
        mensaje: `${guardadas.length} obras procesadas y actualizadas correctamente`,
        resumen: guardadas.map(o => `${o.nombre} (CC ${o.cc}): $${Math.round(o.totalEgresado).toLocaleString()}`).join(' · ')
      });
    } catch(e) {
      try { fs.unlinkSync(req.file.path); } catch(ex) {}
      res.status(500).json({ error: 'Error al procesar el archivo: ' + e.message });
    }
  });

  app.post('/api/upload/cronograma/:cc', upload.single('archivo'), (req, res) => {
    if (!req.file) return res.status(400).json({ error: 'No se recibió ningún archivo' });
    const { cc } = req.params;
    try {
      const cronograma = parseCronograma(req.file.path);
      const existente = leerProyecto(cc) || { cc, nombre: OBRA_NAMES[cc] || 'Obra ' + cc };
      const actualizado = { ...existente, cronograma, ultimaActCronograma: new Date().toISOString().slice(0,10) };
      guardarProyecto(cc, actualizado);
      try { fs.unlinkSync(req.file.path); } catch(e) {}
      actualizarMetaArchivo(cc, 'planObra', req.file.originalname || req.file.filename);
      res.json({ ok: true, cc, cronograma, mensaje: `Cronograma de ${existente.nombre} actualizado. ${cronograma.totalM2} m² totales, ${cronograma.totalSemanas} semanas.` });
    } catch(e) {
      try { fs.unlinkSync(req.file.path); } catch(ex) {}
      res.status(500).json({ error: 'Error al procesar cronograma: ' + e.message });
    }
  });

  // ── Cronograma vía Google Sheets ──────────────────────────────────────────
  app.post('/api/upload/cronograma-gdocs/:cc', (req, res) => {
    const { cc } = req.params;
    const { url, confirmar } = req.body;
    if (!url) return res.status(400).json({ error: 'URL requerida' });

    const csvUrl = gsheetsToCsvUrl(url);
    if (!csvUrl) return res.status(400).json({ error: 'URL de Google Sheets no válida.' });

    const https = require('https');
    const http = require('http');
    const fetchUrl = (u, redirects = 5) => new Promise((resolve, reject) => {
      if (redirects === 0) return reject(new Error('Demasiadas redirecciones'));
      const mod = u.startsWith('https') ? https : http;
      mod.get(u, { headers: { 'User-Agent': 'WallerDashboard/1.0' } }, r => {
        if (r.statusCode === 301 || r.statusCode === 302) { r.resume(); return resolve(fetchUrl(r.headers.location, redirects - 1)); }
        if (r.statusCode !== 200) { r.resume(); return reject(new Error('HTTP ' + r.statusCode + ' — ¿el Sheet está publicado para todos?')); }
        let data = ''; r.setEncoding('utf8'); r.on('data', c => data += c); r.on('end', () => resolve(data));
      }).on('error', reject);
    });

    fetchUrl(csvUrl)
      .then(csvText => {
        // Parsear CSV como cronograma: Nivel | Semana | M²
        const lines = csvText.replace(/\r/g,'').split('\n').filter(l => l.trim());
        if (lines.length < 2) return res.status(400).json({ error: 'El Sheet no tiene datos suficientes.' });

        const parseLn = l => { const r=[];let cur='',inQ=false; for(const c of l){if(c==='"'){inQ=!inQ;}else if(c===','&&!inQ){r.push(cur.trim());cur='';}else cur+=c;} r.push(cur.trim()); return r; };
        const headers = parseLn(lines[0]).map(h => h.toUpperCase().trim());
        const rows = lines.slice(1).map(parseLn);

        const nivelCol = headers.findIndex(h => /^(NIVEL|LEVEL|PISO)$/.test(h));
        const semanaCol = headers.findIndex(h => /^(SEMANA|WEEK|S)$/.test(h));
        const m2Col = headers.findIndex(h => /^(M2|M²|METROS|AREA|CANTIDAD)$/.test(h));

        let cronograma = { fuente: 'google-sheets', url: csvUrl, fechaActualizacion: new Date().toISOString().slice(0,10), totalM2: 0, totalSemanas: 0, m2PorNivel: {}, semanas: [] };

        if (nivelCol >= 0 && m2Col >= 0) {
          rows.forEach(r => {
            const nivel = (r[nivelCol]||'').trim().toUpperCase();
            const m2 = parseFloat((r[m2Col]||'').replace(',','.')) || 0;
            const semana = semanaCol >= 0 ? parseInt(r[semanaCol]) || 0 : 0;
            if (nivel && m2 > 0) {
              cronograma.m2PorNivel[nivel] = (cronograma.m2PorNivel[nivel] || 0) + m2;
              cronograma.totalM2 += m2;
              if (semana > cronograma.totalSemanas) cronograma.totalSemanas = semana;
            }
          });
        } else {
          // Formato columnar: columnas = niveles, filas = semanas
          const semanas = [];
          rows.forEach((r, si) => {
            const semObj = { semana: si+1 };
            headers.forEach((h, hi) => {
              if (hi === 0) return;
              const v = parseFloat((r[hi]||'').replace(',','.')) || 0;
              if (v > 0) {
                semObj[h] = v;
                cronograma.m2PorNivel[h] = (cronograma.m2PorNivel[h]||0) + v;
                cronograma.totalM2 += v;
              }
            });
            semanas.push(semObj);
          });
          cronograma.semanas = semanas;
          cronograma.totalSemanas = semanas.length;
        }

        if (!confirmar) {
          return res.json({ ok: true, preview: true, cronograma, mensaje: `${cronograma.totalM2.toFixed(1)} m² en ${Object.keys(cronograma.m2PorNivel).length} niveles, ${cronograma.totalSemanas} semanas.` });
        }

        const existente = leerProyecto(cc) || { cc, nombre: OBRA_NAMES[cc] || 'Obra ' + cc };
        const actualizado = { ...existente, cronograma, _cronogramaGdocsUrl: csvUrl, ultimaActCronograma: new Date().toISOString().slice(0,10) };
        guardarProyecto(cc, actualizado);
        actualizarMetaArchivo(cc, 'planObra', 'Google Sheets · ' + new Date().toLocaleDateString('es-MX'));
        res.json({ ok: true, cronograma, mensaje: `Cronograma actualizado desde Google Sheets: ${cronograma.totalM2.toFixed(1)} m² totales, ${cronograma.totalSemanas} semanas.` });
      })
      .catch(e => res.status(500).json({ error: e.message }));
  });

  // Refrescar cronograma desde Sheet guardado
  app.post('/api/upload/cronograma-gdocs/:cc/refresh', (req, res) => {
    const { cc } = req.params;
    const proyecto = leerProyecto(cc);
    if (!proyecto?._cronogramaGdocsUrl) return res.status(400).json({ error: 'No hay URL de Google Sheets guardada para el cronograma de esta obra.' });
    req.body = { url: proyecto._cronogramaGdocsUrl, confirmar: true };
    // Reutilizar el handler — simplificado inline
    const https = require('https');
    const fetchUrl = (u, n=5) => new Promise((ok,err) => {
      if(!n) return err(new Error('Demasiadas redirecciones'));
      const mod = u.startsWith('https') ? https : require('http');
      mod.get(u,{headers:{'User-Agent':'WallerDash/1'}},r=>{
        if(r.statusCode===301||r.statusCode===302){r.resume();return ok(fetchUrl(r.headers.location,n-1));}
        if(r.statusCode!==200){r.resume();return err(new Error('HTTP '+r.statusCode));}
        let d='';r.setEncoding('utf8');r.on('data',c=>d+=c);r.on('end',()=>ok(d));
      }).on('error',err);
    });
    fetchUrl(proyecto._cronogramaGdocsUrl)
      .then(csv => {
        // Parseo simplificado — misma lógica que arriba
        const lines = csv.replace(/\r/g,'').split('\n').filter(l=>l.trim());
        const parseLn = l=>{const r=[];let cur='',inQ=false;for(const c of l){if(c==='"'){inQ=!inQ;}else if(c===','&&!inQ){r.push(cur.trim());cur='';}else cur+=c;}r.push(cur.trim());return r;};
        const headers = parseLn(lines[0]).map(h=>h.toUpperCase().trim());
        const rows = lines.slice(1).map(parseLn);
        const m2Col = headers.findIndex(h=>/^(M2|M²|METROS|AREA)$/.test(h));
        let cronograma = {...(proyecto.cronograma||{}), fuente:'google-sheets', url:proyecto._cronogramaGdocsUrl, fechaActualizacion:new Date().toISOString().slice(0,10), totalM2:0, m2PorNivel:{}, totalSemanas:0};
        rows.forEach((r,si)=>{
          headers.forEach((h,hi)=>{
            if(hi===0)return;
            const v=parseFloat((r[hi]||'').replace(',','.'))||0;
            if(v>0){cronograma.m2PorNivel[h]=(cronograma.m2PorNivel[h]||0)+v;cronograma.totalM2+=v;}
          });
          cronograma.totalSemanas=si+1;
        });
        guardarProyecto(cc,{...proyecto,cronograma,ultimaActCronograma:new Date().toISOString().slice(0,10)});
        actualizarMetaArchivo(cc,'planObra','Google Sheets · '+new Date().toLocaleDateString('es-MX'));
        res.json({ok:true,cronograma,mensaje:`Cronograma actualizado: ${cronograma.totalM2.toFixed(1)} m² totales.`});
      })
      .catch(e=>res.status(500).json({error:e.message}));
  });

  app.post('/api/upload/estimacion', upload.single('archivo'), (req, res) => {
    if (!req.file) return res.status(400).json({ error: 'No se recibió ningún archivo' });
    try {
      const est = parseEstimacion(req.file.path);
      if (!est.cc) return res.status(400).json({ error: 'No se pudo identificar la obra en el archivo.' });
      if (!est.numero) return res.status(400).json({ error: 'No se pudo leer el número de estimación.' });

      const existente = leerProyecto(est.cc) || {};
      const estimaciones = existente.estimaciones || [];
      const alertas = [];

      if (est.acumAnterior > 0) {
        const prev = estimaciones.find(e => e.numero === est.numero - 1);
        if (prev && Math.abs(prev.acumActual - est.acumAnterior) > 500) {
          alertas.push(`⚠ El acumulado anterior ($${Math.round(est.acumAnterior).toLocaleString()}) no coincide con estimación ${est.numero - 1} ($${Math.round(prev.acumActual).toLocaleString()})`);
        }
      }
      if (est.montoContrato > 0 && existente.montoContrato && Math.abs(existente.montoContrato - est.montoContrato) / existente.montoContrato > 0.01) {
        alertas.push(`⚠ El monto de contrato ($${Math.round(est.montoContrato).toLocaleString()}) difiere del registrado ($${Math.round(existente.montoContrato).toLocaleString()})`);
      }

      const idx = estimaciones.findIndex(e => e.numero === est.numero);
      if (idx >= 0) estimaciones[idx] = est;
      else { estimaciones.push(est); estimaciones.sort((a, b) => a.numero - b.numero); }

      const actualizado = {
        ...existente,
        cc: est.cc,
        nombre: existente.nombre || OBRA_NAMES[est.cc] || 'Obra ' + est.cc,
        montoContrato: est.montoContrato || existente.montoContrato,
        estimaciones,
        ultimaActualizacion: new Date().toISOString().slice(0, 10)
      };
      guardarProyecto(est.cc, actualizado);
      try { fs.unlinkSync(req.file.path); } catch(e) {}
      actualizarMetaArchivo(est.cc, 'estimaciones', req.file.originalname || req.file.filename);
      res.json({
        ok: true, cc: est.cc, numero: est.numero, alertas,
        mensaje: `Estimación ${est.numero} de ${actualizado.nombre} procesada · $${Math.round(est.estaEstimacion).toLocaleString()} · ${Math.round(est.m2Esta)} m²`,
        data: est
      });
    } catch(e) {
      try { fs.unlinkSync(req.file.path); } catch(ex) {}
      res.status(500).json({ error: 'Error al procesar estimación: ' + e.message });
    }
  });

  app.post('/api/diagnostico/contpaq', upload.single('archivo'), (req, res) => {
    if (!req.file) return res.status(400).json({ error: 'No se recibió archivo' });
    try {
      const wb = XLSX.readFile(req.file.path, { cellText: true });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

      const cuentas = {};
      let currentCuenta = null;

      for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        const col1 = String(row[0] || '').trim();
        const cargo   = parseFloat(String(row[5] || '').replace(/,/g,'')) || 0;
        const abono   = parseFloat(String(row[6] || '').replace(/,/g,'')) || 0;
        const concepto = String(row[2] || row[3] || row[1] || '').trim().slice(0, 80);

        if (col1.match(/^5\d{2}-\d{2}-\d{3}$/)) {
          currentCuenta = col1;
          if (!cuentas[col1]) cuentas[col1] = { cuenta: col1, cargos: 0, abonos: 0, movs: [] };
          continue;
        }
        if (col1.match(/^\d{1,2}\/\w{3}\/\d{4}$/) && currentCuenta) {
          if (cargo > 0 || abono > 0) {
            cuentas[currentCuenta].cargos += cargo;
            cuentas[currentCuenta].abonos += abono;
            if (cuentas[currentCuenta].movs.length < 5) {
              cuentas[currentCuenta].movs.push({ fecha: col1, cargo, abono, concepto });
            }
          }
        }
      }

      const detalle519 = Object.values(cuentas)
        .filter(c => c.cuenta.startsWith('519-'))
        .sort((a, b) => b.cargos - a.cargos)
        .map(c => ({
          cuenta: c.cuenta,
          subcuenta: c.cuenta.split('-')[1],
          partida: PARTIDA_MAP[c.cuenta.split('-')[1]] || 'desconocida',
          cargos: Math.round(c.cargos * 100) / 100,
          abonos: Math.round(c.abonos * 100) / 100,
          neto: Math.round((c.cargos - c.abonos) * 100) / 100,
          muestra: c.movs
        }));

      try { fs.unlinkSync(req.file.path); } catch(e) {}
      res.json({ ok: true, cc519: detalle519, total: detalle519.reduce((s,c) => s + c.neto, 0) });
    } catch(e) {
      try { fs.unlinkSync(req.file.path); } catch(ex) {}
      res.status(500).json({ error: e.message });
    }
  });

  // ── Registro de archivos por proyecto ────────────────────────────────────
  // GET /api/archivos/:cc — devuelve metadata de archivos procesados
  app.get('/api/archivos/:cc', (req, res) => {
    const p = leerProyecto(req.params.cc);
    const archivos = (p && p._archivos) ? JSON.parse(JSON.stringify(p._archivos)) : {};

    // Si hay datos en el servidor pero no hay metadata, infiere desde los datos
    if (p) {
      if (!archivos.contpaq && p.totalEgresado > 0)
        archivos.contpaq = { actual: { nombre: 'Datos cargados (servidor)', fechaCarga: p.ultimaActualizacion || '—', estado: 'ok', fuente: 'inferido' } };
      if (!archivos.planObra && p.cronograma?.totalM2 > 0)
        archivos.planObra = { actual: { nombre: 'Cronograma en servidor', fechaCarga: p.ultimaActCronograma || '—', estado: 'ok', fuente: 'inferido' } };
      if (!archivos.estimaciones && p.estimaciones?.length > 0)
        archivos.estimaciones = { actual: { nombre: `${p.estimaciones.length} estimación(es) en servidor`, fechaCarga: p.ultimaActualizacion || '—', estado: 'ok', fuente: 'inferido' } };
      if (!archivos.contrato && p.montoContrato > 0)
        archivos.contrato = { actual: { nombre: `Contrato $${(p.montoContrato/1e6).toFixed(2)}M`, fechaCarga: '—', estado: 'ok', fuente: 'inferido' } };
    }
    res.json(archivos);
  });

  // POST /api/archivos/:cc/meta — registra metadata de archivo no procesable (contrato, cotiz, etc.)
  app.post('/api/archivos/:cc/meta', (req, res) => {
    const { cc } = req.params;
    const { tipo, nombre, entradaId } = req.body;
    if (!tipo || !nombre) return res.status(400).json({ error: 'tipo y nombre requeridos' });
    actualizarMetaArchivo(cc, tipo, nombre);
    res.json({ ok: true, tipo, nombre });
  });

  // POST /api/archivos/:cc/meta/multi — agrega entrada a tipo multi (preciosUnitarios, contpaq extra)
  app.post('/api/archivos/:cc/meta/multi', (req, res) => {
    const { cc } = req.params;
    const { tipo, nombre } = req.body;
    if (!tipo || !nombre) return res.status(400).json({ error: 'tipo y nombre requeridos' });
    const proyecto = leerProyecto(cc) || { cc };
    if (!proyecto._archivos) proyecto._archivos = {};
    if (!Array.isArray(proyecto._archivos[tipo])) proyecto._archivos[tipo] = [];
    // Max 2 entradas: nueva + la última
    const nueva = { id: Date.now().toString(36), nombre, fechaCarga: new Date().toISOString().slice(0, 10), estado: 'ok' };
    proyecto._archivos[tipo] = [nueva, ...proyecto._archivos[tipo]].slice(0, 2);
    guardarProyecto(cc, proyecto);
    res.json({ ok: true, tipo, nombre });
  });

  app.get('/api/upload/status', (req, res) => {
    try {
      if (!fs.existsSync(DATA_DIR)) return res.json({ proyectos: 0, archivos: [] });
      const files = fs.readdirSync(DATA_DIR).filter(f => f.endsWith('.json'));
      res.json({ proyectos: files.length, archivos: files });
    } catch(e) { res.json({ proyectos: 0, archivos: [] }); }
  });

  // ── Google Sheets — Avance diario ─────────────────────────────────────────
  // ── Upload Contrato PDF — guarda archivo + extrae campos ─────────────────
  app.post('/api/upload/contrato/:cc', uploadContrato.single('archivo'), async (req, res) => {
    if (!req.file) return res.status(400).json({ error: 'No se recibió el archivo PDF' });
    const { cc } = req.params;
    try {
      // Ruta permanente del PDF
      const archivoPdf = `/uploads/contratos/${req.file.filename}`;

      // Extraer texto completo
      const textoCompleto = await extraerTextoPDF(req.file.path);
      if (!textoCompleto) {
        return res.json({ ok: true, archivoPdf, textoCompleto: '', camposDetectados: 0,
          campos: {}, aviso: 'El PDF no tiene texto extraíble (posiblemente escaneado). Llena los campos manualmente.' });
      }

      // Aplicar regex a los 31 campos
      const raw = extraerCamposContrato(textoCompleto);

      // Construir objeto campos con estructura { valor, noAplica }
      const campos = {};
      let camposDetectados = 0;
      Object.entries(raw).forEach(([k, v]) => {
        const detectado = v !== null && v !== undefined;
        if (detectado) camposDetectados++;
        campos[k] = { valor: detectado ? v : null, noAplica: false };
      });

      // Guardar referencia del contrato en el proyecto
      // Mapear campos detectados directamente al proyecto
      const val = (k) => campos[k]?.valor || null;
      const existente = leerProyecto(cc) || { cc };
      const actualizado = {
        ...existente,
        contrato: { archivoPdf, camposDetectados, campos },
        // Datos principales del proyecto extraídos del contrato
        ...(val('montoContrato')   ? { montoContrato:   parseFloat(String(val('montoContrato')).replace(/[^0-9.]/g,'')) || existente.montoContrato } : {}),
        ...(val('fechaInicio')     ? { fechaInicio:      val('fechaInicio') }     : {}),
        ...(val('fechaTerminacion')? { fechaFin:         val('fechaTerminacion') } : {}),
        ...(val('razonSocialCliente') ? { cliente:       val('razonSocialCliente') } : {}),
        ...(val('repLegalCliente') && !val('razonSocialCliente') ? { cliente: val('repLegalCliente') } : {}),
        ...(val('direccionObra')   ? { ubicacion:        val('direccionObra') }   : {}),
        ...(val('nombreProyecto')  && (!existente.nombre || existente.nombre === 'Obra ' + cc) ? { nombre: val('nombreProyecto') } : {}),
        ...(val('plazoEjecucionDias') ? { plazoEjecucionDias: val('plazoEjecucionDias') } : {}),
        ...(val('tipoContrato')    ? { tipoContrato:     val('tipoContrato') }    : {}),
        ...(val('anticipoPct')     ? { anticipoPct:      val('anticipoPct') }     : {}),
        ...(val('retencionPct')    ? { retencionPct:     val('retencionPct') }    : {}),
        ...(val('folioContrato')   ? { folioContrato:    val('folioContrato') }   : {}),
        ...(val('rfcCliente')      ? { rfcCliente:       val('rfcCliente') }      : {}),
        ...(val('banco')           ? { banco:            val('banco') }           : {}),
        ...(val('clabe')           ? { clabe:            val('clabe') }           : {}),
      };
      guardarProyecto(cc, actualizado);
      actualizarMetaArchivo(cc, 'contrato', req.file.originalname || req.file.filename);

      res.json({ ok: true, archivoPdf, textoCompleto: textoCompleto.slice(0, 20000), camposDetectados, campos,
        mensaje: `Contrato guardado. Se detectaron ${camposDetectados} de 31 campos.` });

    } catch(e) {
      res.status(500).json({ error: 'Error procesando el PDF: ' + e.message });
    }
  });

  // Servir los PDFs guardados
  app.use('/uploads/contratos', require('express').static(path.join(__dirname, '../uploads/contratos')));

  // ── Debug: ver texto extraído de PDF ─────────────────────────────────────
  app.post('/api/debug/pdf', upload.single('archivo'), async (req, res) => {
    if (!req.file) return res.status(400).json({ error: 'No archivo' });
    try {
      const texto = await extraerTextoPDF(req.file.path);
      try { fs.unlinkSync(req.file.path); } catch(e) {}
      res.json({ chars: texto.length, texto: texto.slice(0, 5000), lineas: texto.split('\n').slice(0,80) });
    } catch(e) {
      try { fs.unlinkSync(req.file.path); } catch(ex) {}
      res.status(500).json({ error: e.message });
    }
  });

  // ── Debug: ver filas crudas de cualquier Excel ──────────────────────────
  app.post('/api/debug/excel', upload.single('archivo'), (req, res) => {
    if (!req.file) return res.status(400).json({ error: 'No archivo' });
    try {
      const wb = XLSX.readFile(req.file.path, { cellText: true });
      const resultado = {};
      wb.SheetNames.forEach(name => {
        const ws = wb.Sheets[name];
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
        resultado[name] = rows.slice(0, 25).map((r, i) => ({ fila: i, celdas: r.slice(0, 15).map(String) }));
      });
      try { fs.unlinkSync(req.file.path); } catch(e) {}
      res.json({ hojas: wb.SheetNames, filas: resultado });
    } catch(e) {
      try { fs.unlinkSync(req.file.path); } catch(ex) {}
      res.status(500).json({ error: e.message });
    }
  });

  // ── Upload Universal — detección automática ──────────────────────────────
  app.post('/api/upload/auto/:cc', upload.single('archivo'), async (req, res) => {
    const { cc } = req.params;
    const url      = req.body.url;
    const confirmar = req.body.confirmar === '1';
    const tipoForzado = req.body.tipo || null;

    const limpiar = () => { if (req.file) try { fs.unlinkSync(req.file.path); } catch(e) {} };

    try {
      let rows = [], texto = '', headers = [], formato = '', nombreArchivo = 'archivo';

      // ── 1. Leer contenido ──────────────────────────────────────────────
      if (url) {
        const csvUrl = gsheetsToCsvUrl(url);
        if (!csvUrl) return res.status(400).json({ error: 'URL de Google Sheets no válida.' });
        const https = require('https'), http = require('http');
        const fetchUrl = (u, n=5) => new Promise((ok,err) => {
          if(!n) return err(new Error('Demasiadas redirecciones'));
          const mod = u.startsWith('https') ? https : http;
          mod.get(u,{headers:{'User-Agent':'WallerDash/1'}}, r => {
            if(r.statusCode===301||r.statusCode===302){r.resume();return ok(fetchUrl(r.headers.location,n-1));}
            if(r.statusCode!==200){r.resume();return err(new Error('HTTP '+r.statusCode+' — ¿el Sheet está publicado para todos?'));}
            let d=''; r.setEncoding('utf8'); r.on('data',c=>d+=c); r.on('end',()=>ok(d));
          }).on('error',err);
        });
        const csvText = await fetchUrl(csvUrl);
        texto = csvText;
        formato = 'google-sheets';
        nombreArchivo = 'Google Sheets';
        // Parsear CSV simple
        const lineas = csvText.replace(/\r/g,'').split('\n').filter(l=>l.trim());
        const parseLn = l => { const r=[];let cur='',inQ=false; for(const c of l){if(c==='"'){inQ=!inQ;}else if(c===','&&!inQ){r.push(cur.trim());cur='';}else cur+=c;} r.push(cur.trim()); return r; };
        rows = lineas.map(parseLn);
        headers = rows[0]||[];
      } else if (req.file) {
        const ext = path.extname(req.file.originalname).toLowerCase();
        nombreArchivo = req.file.originalname;
        if (ext==='.xlsx'||ext==='.xls') {
          const leido = leerExcelParaDeteccion(req.file.path);
          rows = leido.rows; headers = leido.headers; texto = leido.texto;
          formato = ext.slice(1);
        } else if (ext==='.csv') {
          texto = fs.readFileSync(req.file.path,'utf8');
          const lineas = texto.replace(/\r/g,'').split('\n').filter(l=>l.trim());
          rows = lineas.map(l=>l.split(',').map(v=>v.trim().replace(/"/g,'')));
          headers = rows[0]||[];
          formato = 'csv';
        } else if (ext==='.pdf') {
          texto = await extraerTextoPDF(req.file.path);
          formato = 'pdf';
          headers = [];
        } else {
          limpiar();
          return res.status(400).json({ error: 'Formato no soportado. Usa .xlsx, .xls, .csv o .pdf' });
        }
      } else {
        return res.status(400).json({ error: 'Envía un archivo o una URL de Google Sheets.' });
      }

      // ── 2. Detectar tipo ───────────────────────────────────────────────
      const deteccion = tipoForzado
        ? { tipo: tipoForzado, confianza: 'alta' }
        : detectarTipoContenido(texto, headers);

      // ── 3. Parsear según tipo ──────────────────────────────────────────
      let datos = {}, registros = 0, preview = [];

      if (formato === 'pdf') {
        datos = parsearPdfPorTipo(texto, deteccion.tipo || 'desconocido');
        registros = Object.keys(datos).filter(k=>datos[k]!==null&&!k.startsWith('_')).length;
        preview = [{ texto: texto.slice(0, 3000) }];
      } else {
        // Excel / CSV / Google Sheets
        if (deteccion.tipo === 'contpaq' && req.file) {
          const parsed = parseLibroAuxiliar(req.file.path);
          datos = parsed;
          registros = Object.keys(parsed).length;
          preview = Object.values(parsed).slice(0,3).map(o=>({ obra: o.nombre, cc: o.cc, totalEgresado: o.totalEgresado }));
        } else if (deteccion.tipo === 'cronograma' && req.file) {
          datos = parseCronograma(req.file.path);
          registros = Object.keys(datos.m2PorNivel||{}).length;
          preview = Object.entries(datos.m2PorNivel||{}).slice(0,3).map(([n,m])=>({ nivel:n, m2:m }));
        } else if (deteccion.tipo === 'estimacion' && req.file) {
          datos = parseEstimacion(req.file.path);
          registros = datos.conceptos?.length || 1;
          preview = [{ 'Estimación': datos.numero, 'Monto': '$'+Math.round(datos.estaEstimacion||0).toLocaleString('es-MX'), 'M²': datos.m2Esta }];
        } else if (deteccion.tipo === 'cotizacion' && req.file) {
          datos = parseCotizacionWaller(req.file.path);
          registros = datos.secciones?.reduce((s,sec)=>s+sec.conceptos.length,0) || 0;
          preview = datos.secciones?.slice(0,3).map(sec=>({
            seccion: sec.seccion,
            'm² total': sec.totalM2,
            importe: '$'+Math.round(sec.totalImporte).toLocaleString('es-MX')
          })) || [];
        } else if (deteccion.tipo === 'avanceDiario') {
          const csvText2 = url ? texto : rows.map(r=>r.join(',')).join('\n');
          const parsed = parseAvanceDiarioCsv(csvText2 || rows.map(r=>r.join(',')).join('\n'));
          datos = parsed || { entradas:[], m2PorNivel:{}, totalM2Instalado:0 };
          registros = datos.entradas?.length || 0;
          preview = Object.entries(datos.m2PorNivel||{}).slice(0,3).map(([n,m])=>({ nivel:n, m2Acum:m }));
        } else if (deteccion.tipo === 'cambioAlcance' && req.file) {
          const wb2 = XLSX.readFile(req.file.path, { cellText: true });
          const ws2 = wb2.Sheets[wb2.SheetNames[0]];
          const rawRows = XLSX.utils.sheet_to_json(ws2, { defval: '' });
          datos = { cambios: rawRows };
          registros = rawRows.length;
          preview = rawRows.slice(0,3);
        } else {
          // Tipo desconocido — mostrar raw para que el usuario confirme
          datos = { _headers: headers };
          registros = rows.length - 1;
          // Buscar fila de encabezado real (primera con más de 2 celdas no vacías)
          const headerRowIdx = rows.findIndex(r => r.filter(c=>String(c).trim()).length >= 3);
          preview = extraerPreview(rows.slice(headerRowIdx >= 0 ? headerRowIdx : 0), 3);
        }
      }

      // ── 4. Si no confirmar → devolver preview ─────────────────────────
      if (!confirmar) {
        limpiar();
        return res.json({
          ok: true,
          preview: true,
          tipo: deteccion.tipo,
          confianza: deteccion.confianza,
          formato,
          nombreArchivo,
          registros,
          preview,
          datos: formato === 'pdf' ? datos : undefined
        });
      }

      // ── 5. Confirmar → guardar en proyecto ────────────────────────────
      const existente = leerProyecto(cc) || { cc, nombre: OBRA_NAMES[cc] || 'Obra '+cc };
      let actualizado = { ...existente };
      const hoy = new Date().toISOString().slice(0,10);

      if (deteccion.tipo === 'contpaq') {
        // ContPAQ actualiza múltiples obras
        const obras = Object.values(datos);
        const guardadas = [];
        for (const obra of obras) {
          if (!obra.cc) continue;
          const ex = leerProyecto(obra.cc) || {};
          guardarProyecto(obra.cc, { ...ex, cc:obra.cc, nombre:obra.nombre||ex.nombre||('Obra '+obra.cc), partidas:obra.partidas, partidasDetalle:obra.partidasDetalle, totalEgresado:obra.totalEgresado, ultimaActualizacion:hoy });
          actualizarMetaArchivo(obra.cc,'contpaq',nombreArchivo);
          guardadas.push(obra.nombre||obra.cc);
        }
        limpiar();
        return res.json({ ok:true, tipo:'contpaq', formato, registros:guardadas.length, mensaje:`ContPAQ procesado: ${guardadas.join(', ')}` });
      } else if (deteccion.tipo === 'cronograma') {
        actualizado = { ...actualizado, cronograma:datos, ultimaActCronograma:hoy };
      } else if (deteccion.tipo === 'estimacion') {
        const ests = actualizado.estimaciones || [];
        const idx = ests.findIndex(e=>e.numero===datos.numero);
        if(idx>=0) ests[idx]=datos; else { ests.push(datos); ests.sort((a,b)=>a.numero-b.numero); }
        actualizado = { ...actualizado, estimaciones:ests, montoContrato:datos.montoContrato||actualizado.montoContrato, ultimaActualizacion:hoy };
      } else if (deteccion.tipo === 'avanceDiario') {
        actualizado = { ...actualizado, _avanceDiario:datos.entradas, _m2InstalPorNivel:datos.m2PorNivel, _totalM2Instalado:datos.totalM2Instalado, _gdocsUrl:url||null, _ultimaActAvance:hoy };
      } else if (deteccion.tipo === 'cotizacion') {
        actualizado = { ...actualizado, cotizacion: datos, ...(datos.totalPresupuesto ? { _totalCotizacion: datos.totalPresupuesto } : {}) };
      } else if (deteccion.tipo === 'contrato') {
        actualizado = { ...actualizado, contrato:{ ...(actualizado.contrato||{}), ...datos }, ...(datos.montoContrato?{montoContrato:datos.montoContrato}:{}) };
      } else if (deteccion.tipo === 'preciosUnitarios') {
        actualizado = { ...actualizado, preciosUnitarios:datos };
      } else if (deteccion.tipo === 'cambioAlcance') {
        const cambiosExist = actualizado.cambiosAlcance || [];
        const nuevos = (datos.cambios || []).filter(c => c);
        actualizado = { ...actualizado, cambiosAlcance: [...cambiosExist, ...nuevos] };
      } else {
        actualizado = { ...actualizado, [`_auto_${deteccion.tipo||'desconocido'}`]:datos };
      }

      guardarProyecto(cc, actualizado);
      actualizarMetaArchivo(cc, deteccion.tipo||'otro', nombreArchivo);
      limpiar();
      res.json({ ok:true, tipo:deteccion.tipo, formato, registros, mensaje:`${deteccion.tipo||'Archivo'} guardado correctamente. ${registros} registros.` });

    } catch(e) {
      limpiar();
      res.status(500).json({ error: e.message });
    }
  });

  app.post('/api/upload/gdocs/:cc', (req, res) => {
    const { cc } = req.params;
    const { url, confirmar } = req.body;
    if (!url) return res.status(400).json({ error: 'URL requerida' });

    const csvUrl = gsheetsToCsvUrl(url);
    if (!csvUrl) return res.status(400).json({ error: 'URL de Google Sheets no válida. Asegúrate de que el enlace sea de un Google Sheet.' });

    const https = require('https');
    const http = require('http');
    const fetchUrl = (u, redirects = 5) => new Promise((resolve, reject) => {
      if (redirects === 0) return reject(new Error('Demasiadas redirecciones'));
      const mod = u.startsWith('https') ? https : http;
      mod.get(u, { headers: { 'User-Agent': 'WallerDashboard/1.0' } }, r => {
        if (r.statusCode === 301 || r.statusCode === 302) {
          r.resume();
          return resolve(fetchUrl(r.headers.location, redirects - 1));
        }
        if (r.statusCode !== 200) { r.resume(); return reject(new Error('HTTP ' + r.statusCode + ' al acceder al Sheet. ¿Está publicado para todos?')); }
        let data = '';
        r.setEncoding('utf8');
        r.on('data', c => data += c);
        r.on('end', () => resolve(data));
      }).on('error', reject);
    });

    fetchUrl(csvUrl)
      .then(csvText => {
        const parsed = parseAvanceDiarioCsv(csvText);
        if (!parsed || !parsed.entradas.length)
          return res.status(400).json({ error: 'No se encontraron datos de avance en el Sheet. Verifica el formato (columnas: Fecha, Nivel o niveles como NPB/N1/N2...).' });

        if (!confirmar) {
          return res.json({ ok: true, preview: true, ...parsed, csvUrl, mensaje: `Se encontraron ${parsed.entradas.length} filas con avance en ${Object.keys(parsed.m2PorNivel).length} niveles.` });
        }

        // Guardar confirmado
        const existente = leerProyecto(cc) || { cc, nombre: OBRA_NAMES[cc] || 'Obra ' + cc };
        const actualizado = {
          ...existente,
          _avanceDiario: parsed.entradas,
          _m2InstalPorNivel: parsed.m2PorNivel,
          _totalM2Instalado: parsed.totalM2Instalado,
          _gdocsUrl: csvUrl,
          _ultimaActAvance: new Date().toISOString().slice(0, 10)
        };
        guardarProyecto(cc, actualizado);
        actualizarMetaArchivo(cc, 'avanceDiario', 'Google Sheets · ' + new Date().toLocaleDateString('es-MX'));
        res.json({ ok: true, ...parsed, mensaje: `Avance diario actualizado: ${parsed.totalM2Instalado.toFixed(1)} m² en ${Object.keys(parsed.m2PorNivel).length} niveles.` });
      })
      .catch(e => res.status(500).json({ error: e.message }));
  });

  // Endpoint para refrescar el Sheet guardado (sin necesidad de pegar URL de nuevo)
  app.post('/api/upload/gdocs/:cc/refresh', (req, res) => {
    const { cc } = req.params;
    const proyecto = leerProyecto(cc);
    if (!proyecto?._gdocsUrl) return res.status(400).json({ error: 'No hay URL de Google Sheets guardada para esta obra.' });
    req.body = { url: proyecto._gdocsUrl, confirmar: true };
    req.params.cc = cc;
    // Re-use the handler above by re-routing — easier to just inline the logic
    const https = require('https');
    const fetchUrl = (u, redirects = 5) => new Promise((resolve, reject) => {
      if (redirects === 0) return reject(new Error('Demasiadas redirecciones'));
      const mod = u.startsWith('https') ? https : require('http');
      mod.get(u, { headers: { 'User-Agent': 'WallerDashboard/1.0' } }, r => {
        if (r.statusCode === 301 || r.statusCode === 302) { r.resume(); return resolve(fetchUrl(r.headers.location, redirects - 1)); }
        if (r.statusCode !== 200) { r.resume(); return reject(new Error('HTTP ' + r.statusCode)); }
        let data = ''; r.setEncoding('utf8'); r.on('data', c => data += c); r.on('end', () => resolve(data));
      }).on('error', reject);
    });
    fetchUrl(proyecto._gdocsUrl)
      .then(csv => {
        const parsed = parseAvanceDiarioCsv(csv);
        if (!parsed?.entradas.length) return res.status(400).json({ error: 'Sin datos de avance en el Sheet.' });
        const actualizado = { ...proyecto, _avanceDiario: parsed.entradas, _m2InstalPorNivel: parsed.m2PorNivel, _totalM2Instalado: parsed.totalM2Instalado, _ultimaActAvance: new Date().toISOString().slice(0, 10) };
        guardarProyecto(cc, actualizado);
        actualizarMetaArchivo(cc, 'avanceDiario', 'Google Sheets · ' + new Date().toLocaleDateString('es-MX'));
        res.json({ ok: true, ...parsed, mensaje: `Avance actualizado: ${parsed.totalM2Instalado.toFixed(1)} m² totales.` });
      })
      .catch(e => res.status(500).json({ error: e.message }));
  });

  // ── PDF — Interpretación con Gemini ────────────────────────────────────────
  app.post('/api/upload/pdf/:cc', upload.single('archivo'), async (req, res) => {
    if (!req.file) return res.status(400).json({ error: 'No se recibió ningún archivo' });
    const { cc } = req.params;
    const tipo = req.body.tipo || 'desconocido';
    const ANTHROPIC_KEY = process.env.ANTHROPIC_API_KEY || '';
    if (!ANTHROPIC_KEY) {
      try { fs.unlinkSync(req.file.path); } catch(e) {}
      return res.status(500).json({ error: 'ANTHROPIC_API_KEY no configurada en el servidor. Pide al administrador que la agregue con: pm2 set waller-obra-8085:ANTHROPIC_API_KEY sk-ant-...' });
    }

    try {
      const pdfBase64 = fs.readFileSync(req.file.path).toString('base64');
      const prompt = promptParaPdf(tipo, cc, OBRA_NAMES[cc] || 'Obra ' + cc);
      const bodyObj = {
        model: 'claude-haiku-4-5-20251001',
        max_tokens: 2048,
        messages: [{
          role: 'user',
          content: [
            { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: pdfBase64 } },
            { type: 'text', text: prompt }
          ]
        }]
      };
      const bodyStr = JSON.stringify(bodyObj);
      const https = require('https');
      const claudeRes = await new Promise((resolve, reject) => {
        const opts = {
          hostname: 'api.anthropic.com',
          path: '/v1/messages',
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'x-api-key': ANTHROPIC_KEY,
            'anthropic-version': '2023-06-01',
            'anthropic-beta': 'pdfs-2024-09-25',
            'Content-Length': Buffer.byteLength(bodyStr)
          }
        };
        const r = https.request(opts, resp => {
          let d = ''; resp.on('data', c => d += c); resp.on('end', () => resolve(d));
        });
        r.on('error', reject); r.write(bodyStr); r.end();
      });

      try { fs.unlinkSync(req.file.path); } catch(e) {}
      const cJson = JSON.parse(claudeRes);
      if (cJson.error) return res.status(400).json({ error: 'Claude: ' + (cJson.error.message || JSON.stringify(cJson.error)) });

      const rawText = cJson.content?.map(b => b.text || '').join('') || '{}';
      // Extraer JSON del texto (puede estar rodeado de markdown)
      const jsonMatch = rawText.match(/\{[\s\S]*\}/);
      let extraido;
      try { extraido = JSON.parse(jsonMatch ? jsonMatch[0] : rawText); } catch(e) { extraido = { _raw: rawText }; }

      if (!req.body.confirmar) {
        return res.json({ ok: true, preview: true, tipo, extraido, mensaje: 'PDF interpretado por Claude. Revisa los datos antes de guardar.' });
      }

      const existente = leerProyecto(cc) || { cc, nombre: OBRA_NAMES[cc] || 'Obra ' + cc };
      const actualizado = aplicarPdfAlProyecto(existente, tipo, extraido);
      guardarProyecto(cc, actualizado);
      actualizarMetaArchivo(cc, tipo, req.file.originalname || req.file.filename);
      res.json({ ok: true, tipo, extraido, mensaje: `PDF de ${tipo} interpretado y guardado correctamente.` });

    } catch(e) {
      try { fs.unlinkSync(req.file.path); } catch(ex) {}
      res.status(500).json({ error: 'Error interpretando PDF: ' + e.message });
    }
  });

  // ── Proxy Gemini (texto) ──────────────────────────────────────────────────
  app.post('/api/gemini', async (req, res) => {
    try {
      const { prompt } = req.body;
      if (!prompt) return res.status(400).json({ error: 'prompt requerido' });
      const GEMINI_KEY = process.env.GEMINI_KEY || '';
      const https = require('https');
      const body = JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] });
      const options = {
        hostname: 'generativelanguage.googleapis.com',
        path: '/v1beta/models/gemini-2.0-flash:generateContent',
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'X-goog-api-key': GEMINI_KEY, 'Content-Length': Buffer.byteLength(body) }
      };
      const gemReq = https.request(options, gemRes => {
        let data = '';
        gemRes.on('data', chunk => data += chunk);
        gemRes.on('end', () => {
          try {
            const json = JSON.parse(data);
            if (json.error) return res.status(400).json({ error: json.error.message });
            const parts = json.candidates?.[0]?.content?.parts || [];
            const texto = parts.map(p => p.text || '').join('').trim() || 'Sin respuesta.';
            res.json({ texto });
          } catch(e) { res.status(500).json({ error: e.message }); }
        });
      });
      gemReq.on('error', e => res.status(500).json({ error: e.message }));
      gemReq.write(body);
      gemReq.end();
    } catch(e) { res.status(500).json({ error: e.message }); }
  });

};
