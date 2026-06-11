const express = require('express');
const fs = require('fs');
const path = require('path');
const app = express();

app.use(express.json({ limit: '10mb' }));


const Database = require('better-sqlite3');
const db = new Database('./cotizaciones.db');

// Crear tabla si no existe
db.exec(`
  CREATE TABLE IF NOT EXISTS cotizaciones (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    folio TEXT NOT NULL,
    fecha TEXT,
    vendedor TEXT,
    cliente TEXT,
    telefono TEXT,
    email TEXT,
    razon_social TEXT,
    proyecto TEXT,
    ubicacion TEXT,
    muros_m2 REAL,
    muros_pm2 REAL,
    muros_imp REAL,
    losas_m2 REAL,
    losas_pm2 REAL,
    losas_imp REAL,
    panel_total REAL,
    comp_total REAL,
    flete_total REAL,
    subtotal REAL,
    iva REAL,
    total REAL,
    estado TEXT DEFAULT 'Original',
    folio_original TEXT,
    datos_json TEXT,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP
  )
`);

// Guardar cotización
app.post('/api/cotizaciones/guardar', (req, res) => {
  try {
    const D = req.body;
    const stmt = db.prepare(`
      INSERT INTO cotizaciones (
        folio, fecha, vendedor, cliente, telefono, email, razon_social,
        proyecto, ubicacion, muros_m2, muros_pm2, muros_imp,
        losas_m2, losas_pm2, losas_imp, panel_total, comp_total,
        flete_total, subtotal, iva, total, estado, folio_original, datos_json
      ) VALUES (
        @folio, @fecha, @vendedor, @cliente, @telefono, @email, @razon_social,
        @proyecto, @ubicacion, @muros_m2, @muros_pm2, @muros_imp,
        @losas_m2, @losas_pm2, @losas_imp, @panel_total, @comp_total,
        @flete_total, @subtotal, @iva, @total, @estado, @folio_original, @datos_json
      )
    `);
    stmt.run(D);
    res.json({ ok: true });
  } catch(e) {
    res.status(500).json({ error: e.message });
  }
});

// Obtener todas las cotizaciones
app.get('/api/cotizaciones', (req, res) => {
  try {
    const rows = db.prepare('SELECT * FROM cotizaciones ORDER BY created_at DESC').all();
    res.json(rows);
  } catch(e) {
    res.status(500).json({ error: e.message });
  }
});

// Obtener cotización por folio
app.get('/api/cotizaciones/folio/:folio', (req, res) => {
  try {
    const row = db.prepare('SELECT * FROM cotizaciones WHERE folio = ? ORDER BY created_at DESC LIMIT 1').get(req.params.folio);
    if(!row) return res.status(404).json({ error: 'No encontrada' });
    res.json(row);
  } catch(e) {
    res.status(500).json({ error: e.message });
  }
});

// Eliminar cotizaciones por IDs
app.post('/api/cotizaciones/eliminar', (req, res) => {
  try {
    const { ids, password } = req.body;
    if(password !== '01Proyectos') return res.status(401).json({ error: 'Contrasena incorrecta' });
    const stmt = db.prepare('DELETE FROM cotizaciones WHERE id = ?');
    ids.forEach(id => stmt.run(id));
    res.json({ ok: true, eliminados: ids.length });
  } catch(e) {
    res.status(500).json({ error: e.message });
  }
});


const jwt = require('jsonwebtoken');
const cookieParser = require('cookie-parser');
app.use(cookieParser());
app.use(express.static('public'));
require('./obras-routes')(app);
const JWT_SECRET = 'WallerSecret2026XK9mP';
const USERS = {'Vendedor01':{pass:'Ventas01',nombre:'Gisela Lemus',iniciales:'GL',rol:'vendedor'},'Administrador01':{pass:'AdminWaller',nombre:'Gustavo Rodriguez',iniciales:'GR',rol:'admin'},'WACOM01':{pass:'Asdf01',nombre:'Gisela Lemus',iniciales:'GL',rol:'vendedor'},'WACOM02':{pass:'Qwer02',nombre:'Juan Francisco Saavedra',iniciales:'JFS',rol:'vendedor'},'WAAI01':{pass:'Zxcv01',nombre:'Control AI',iniciales:'CAI',rol:'admin'},'Proyectos01':{pass:'01',nombre:'Proyectos',iniciales:'PR',rol:'vendedor'}};
app.post('/api/auth/login',(req,res)=>{const{usuario,password}=req.body;const user=USERS[usuario];if(!user||user.pass!==password)return res.status(401).json({error:'Usuario o contrasena incorrectos'});const token=jwt.sign({usuario,nombre:user.nombre,iniciales:user.iniciales,rol:user.rol},JWT_SECRET,{expiresIn:'8h'});res.cookie('waller_token',token,{httpOnly:true,secure:true,sameSite:'strict',maxAge:8*60*60*1000});res.json({ok:true,nombre:user.nombre,rol:user.rol,iniciales:user.iniciales});});
app.post('/api/auth/logout',(req,res)=>{res.clearCookie('waller_token');res.json({ok:true});});
app.get('/api/auth/me',(req,res)=>{const token=req.cookies&&req.cookies.waller_token;if(!token)return res.status(401).json({error:'No autenticado'});try{const decoded=jwt.verify(token,JWT_SECRET);res.json({ok:true,usuario:decoded.usuario,nombre:decoded.nombre,rol:decoded.rol,iniciales:decoded.iniciales});}catch(e){res.status(401).json({error:'Token invalido'});}});

const COUNTER_FILE = './counter.json';
function getCounter() {
  try { return JSON.parse(fs.readFileSync(COUNTER_FILE)); }
  catch { return { value: 50 }; }
}
function saveCounter(val) {
  fs.writeFileSync(COUNTER_FILE, JSON.stringify({ value: val }));
}
app.get('/api/counter', (req, res) => res.json(getCounter()));
app.post('/api/counter/increment', (req, res) => {
  const c = getCounter(); c.value += 1; saveCounter(c.value); res.json(c);
});

app.post('/api/pdf', async (req, res) => {
  let browser;
  try {
    const puppeteer = require('puppeteer-core');
    const sharp = require('sharp');
    const { state, filename } = req.body;
    const D = state;
    const esEscuela = D.proyecto === 'Escuela Tipo';

    

    const logoB64 = (await sharp(path.join(__dirname, 'public/img/logo.jpg')).resize(400).jpeg({quality:90}).toBuffer()).toString('base64');
    const plantaB64 = (await sharp(path.join(__dirname, esEscuela ? 'public/img/escuela.jpg' : 'public/img/planta.jpg')).resize(900).jpeg({quality:80}).toBuffer()).toString('base64');
    const axonoB64 = esEscuela ? null : (await sharp(path.join(__dirname, 'public/img/axono.jpg')).resize(900).jpeg({quality:80}).toBuffer()).toString('base64');

    const logoSrc = `data:image/jpeg;base64,${logoB64}`;
    const plantaSrc = `data:image/jpeg;base64,${plantaB64}`;
    const axonoSrc = axonoB64 ? `data:image/jpeg;base64,${axonoB64}` : '';

    const wm = `<svg viewBox="0 0 794 1123" preserveAspectRatio="xMidYMid slice" xmlns="http://www.w3.org/2000/svg"><defs><pattern id="WP" x="0" y="0" width="240" height="120" patternUnits="userSpaceOnUse" patternTransform="rotate(-22)"><image href="${logoSrc}" x="8" y="8" width="120" height="37" opacity="1"/><image href="${logoSrc}" x="128" y="68" width="120" height="37" opacity="1"/></pattern></defs><rect width="794" height="1123" fill="url(#WP)" opacity="0.07"/></svg>`;
    const compRows = D.comp_rows.map((c,i) => `<tr><td style="text-align:center;width:24px">${i+1}</td><td style="text-align:left">${c.nombre}</td><td style="text-align:center">${c.unidad}</td><td style="text-align:right">${c.cant}</td><td style="text-align:right">${c.pu}</td><td style="text-align:right">${c.imp}</td></tr>`).join('');
    const hdr = `<div class="hdr"><img src="${logoSrc}" alt="Waller"><div class="hdr-meta"><div class="hdr-folio">${D.folio}</div><div class="hdr-fecha">${D.fecha}</div></div></div>`;
    const ftr = `<div class="ftr"><span>Calle Jose Maria Vigil 2808 Int 6, Col. Providencia</span><span>waller.mx | @waller.mx</span><span>Oficina (33) 23 0303 5363</span></div>`;
    const wmDiv = `<div class="wm">${wm}</div>`;

    const imgsHtml = esEscuela
      ? `<div class="imgs-wrap"><div class="imgs-grid" style="grid-template-columns:1fr;">
          <div><div class="img-label">Vista del Proyecto</div><img src="${plantaSrc}"></div>
         </div></div>`
      : `<div class="imgs-wrap"><div class="imgs-grid">
          <div><div class="img-label">Planta de Cubierta</div><img src="${plantaSrc}"></div>
          <div><div class="img-label">Vista Axonometrica</div><img src="${axonoSrc}"></div>
         </div></div>`;

    const html = `<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><title>WALLER</title><style>
*{box-sizing:border-box;margin:0;padding:0;}
body{font-family:Arial,sans-serif;background:#fff;color:#222;}
.page{width:210mm;padding:14mm 16mm 12mm;background:#fff;position:relative;overflow:hidden;display:flex;flex-direction:column;}
.wm{position:absolute;top:0;left:0;right:0;bottom:0;pointer-events:none;z-index:0;overflow:hidden;}
.wm svg{width:100%;height:100%;}
.page>*:not(.wm){position:relative;z-index:1;}
.hdr{display:flex;justify-content:space-between;align-items:center;padding-bottom:10px;border-bottom:1.5px solid #2B2B2B;margin-bottom:20px;}
.hdr img{height:36px;width:auto;display:block;}
.hdr-meta{text-align:right;}
.hdr-folio{font-size:13px;font-weight:700;color:#0072BC;letter-spacing:1px;}
.hdr-fecha{font-size:10px;color:#555;margin-top:2px;}
.cbar{display:grid;grid-template-columns:1fr 1.6fr 1.3fr;background:#2B2B2B;color:#fff;margin-bottom:12px;}
.cbar div{padding:6px 10px;border-right:1px solid #444;}
.cbar div:last-child{border-right:none;}
.cbar .lbl{color:#999;font-size:8px;letter-spacing:1.5px;text-transform:uppercase;display:block;margin-bottom:1px;}
.cbar .val{font-weight:600;font-size:11px;}
.intro{font-size:10px;color:#444;margin-bottom:12px;line-height:1.6;}
.slbl{font-size:10px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:#0072BC;margin:10px 0 5px;padding-bottom:3px;border-bottom:1px solid #ddd;}
table{width:100%;border-collapse:collapse;margin-bottom:8px;font-size:10px;}
thead th{background:#2B2B2B;color:#fff;padding:5px 8px;font-size:10px;font-weight:600;text-align:right;}
thead th.left{text-align:left;}
thead th.center{text-align:center;}
tbody tr:nth-child(even){background:#f7f7f7;}
tbody td{padding:4px 8px;border-bottom:1px solid #eee;vertical-align:middle;text-align:right;}
tfoot td{padding:4px 8px;font-weight:600;font-size:10px;border-top:1.5px solid #ccc;background:#f2f2f2;text-align:right;}
tfoot tr.tr td{background:#0072BC;color:#fff;font-size:12px;font-weight:700;border-top:none;}
.notas{background:#f9f9f9;border-left:3px solid #0072BC;padding:8px 12px;font-size:7.5px;color:#555;line-height:1.5;margin-top:8px;}
.notas strong{display:block;margin-bottom:4px;color:#2B2B2B;font-size:9px;}
.aviso{background:#fafafa;border-left:3px solid #bbb;padding:8px 12px;font-size:6.8px;color:#888;line-height:1.5;margin-top:6px;}
.aviso strong{display:block;margin-bottom:3px;color:#aaa;font-size:8px;}
.ftr{border-top:1px solid #ddd;padding-top:6px;display:flex;justify-content:space-between;font-size:8px;color:#aaa;margin-top:16px;}
.imgs-wrap{flex:1;display:flex;flex-direction:column;justify-content:center;padding:16px 0;}
.imgs-grid{display:grid;grid-template-columns:1fr 1fr;gap:12px;}
.img-label{font-size:8px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:#555;margin-bottom:5px;text-align:center;}
.imgs-grid img{width:100%;border:1px solid #ddd;display:block;}
@page{size:A4 portrait;margin:0;}
</style></head><body>
<div class="page" style="min-height:297mm;">${wmDiv}${hdr}
<div class="cbar"><div><span class="lbl">Cliente</span><span class="val">${D.cliente}</span></div><div><span class="lbl">Proyecto</span><span class="val">${D.proyecto}</span></div><div><span class="lbl">Ubicacion</span><span class="val">${D.ubicacion}</span></div></div>
<p class="intro">Agradecemos la oportunidad para participar en el proyecto <strong>${D.proyecto}</strong>.<br>A continuacion le mostramos una cotizacion con los volumenes totales y precios vigentes para el desarrollo los muros de su proyecto.<br>El equipo de WALLER agradece su confianza en nosotros.</p>
<div class="slbl">Resumen</div>
<table>
<thead><tr><th class="left">Concepto</th><th>Importe</th></tr></thead>
<tbody>
<tr><td style="text-align:left">Costo total de Panel Waller (Muros + Losas)</td><td>${D.panel_total}</td></tr>
<tr><td style="text-align:left">Costo total de Complementarios</td><td>${D.comp_total}</td></tr>
<tr><td style="text-align:left">Costo total de Fletes</td><td>${D.flete_total}</td></tr>
<tr><td style="text-align:left">Costo por m2 de materiales (Panel, Complementarios y Fletes)</td><td>${D.m2_costo}</td></tr>
</tbody>
<tfoot>
<tr><td style="text-align:left">Subtotal sin IVA</td><td>${D.subtotal}</td></tr>
<tr><td style="text-align:left">IVA (16%)</td><td>${D.iva}</td></tr>
<tr class="tr"><td style="text-align:left">IMPORTE TOTAL</td><td>${D.total}</td></tr>
</tfoot>
</table>
${imgsHtml}
${ftr}</div>
<div style="page-break-before:always;"></div>
<div class="page">${wmDiv}${hdr}
<div class="slbl">Desglose - Muros</div>
<table>
<thead><tr><th class="left">Concepto</th><th>m2</th><th>$ x m2</th><th>Importe</th></tr></thead>
<tbody><tr><td style="text-align:left">Panel Waller 7.5 cm (0.61 x 2.44 m)</td><td>${D.muros_m2}</td><td>${D.muros_pm2}</td><td>${D.muros_imp}</td></tr></tbody>
<tfoot><tr><td colspan="4" style="text-align:left"><strong>Total m2 con Desperdicio: ${D.muros_m2}</strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<strong>Total de piezas Panel Waller: ${Math.ceil(parseFloat(D.muros_m2)/1.4884)}</strong></td></tr></tfoot>
</table>
<div class="slbl">Desglose - Losas</div>
<table>
<thead><tr><th class="left">Concepto</th><th>m2</th><th>$ x m2</th><th>Importe</th></tr></thead>
<tbody><tr><td style="text-align:left">Panel Waller 7.5 cm (0.61 x 2.44 m)</td><td>${D.losas_m2}</td><td>${D.losas_pm2}</td><td>${D.losas_imp}</td></tr></tbody>
<tfoot><tr><td colspan="4" style="text-align:left"><strong>Total m2 con Desperdicio: ${D.losas_m2}</strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<strong>Total de piezas Panel Waller: ${Math.ceil(parseFloat(D.losas_m2)/1.4884)}</strong></td></tr></tfoot>
</table>
<div class="slbl">Desglose de Complementarios</div>
<table>
<thead><tr><th class="center" style="width:24px">#</th><th class="left">Concepto</th><th>Unidad</th><th>Cantidad</th><th>P.Unit.</th><th>Importe</th></tr></thead>
<tbody>${compRows}</tbody>
<tfoot>
<tr><td colspan="5" style="text-align:left">Subtotal Complementarios</td><td>${D.comp_sub}</td></tr>
<tr><td colspan="5" style="text-align:left">IVA</td><td>${D.comp_iva}</td></tr>
<tr class="tr"><td colspan="5" style="text-align:left">Total</td><td>${D.comp_tot}</td></tr>
</tfoot>
</table>
<div class="slbl">Desglose de Fletes</div>
<table>
<thead><tr><th class="left">Concepto</th><th>Viajes</th><th>$ x Viaje</th><th>Importe</th></tr></thead>
<tbody><tr>
<td style="text-align:left">Envio de Panel Waller (6cm y/o 7.5cm y/o 9cm) de Tlajomulco de Zuniga, Jalisco, Mexico, Planta Waller a destino solicitado. Entrega a pie de calle. NO INCLUYE DESCARGA por parte de Waller.</td>
<td>${D.viajes}</td><td>${D.flete_xv}</td><td>${D.flete_imp}</td>
</tr></tbody>
<tfoot>
<tr><td colspan="3" style="text-align:left">Subtotal</td><td>${D.flete_sub}</td></tr>
<tr><td colspan="3" style="text-align:left">IVA</td><td>${D.flete_iva}</td></tr>
<tr class="tr"><td colspan="3" style="text-align:left">Total</td><td>${D.flete_tot}</td></tr>
</tfoot>
</table>
<div class="notas"><strong>Notas</strong>
- La entrega se programara de 3 a 5 dias habiles despues del pago total.<br>
- Esta cotizacion tiene vigencia de 30 dias naturales.<br>
- El rendimiento del calculo de los materiales es con base en la ficha tecnica de los mismos.<br>
- El costo es en moneda nacional (Pesos Mexicanos).<br>
- En caso de devaluaciones monetarias o inflaciones mayores al 5%, se realizara un ajuste a los precios.<br>
- Esta cotizacion es de caracter confidencial y unicamente valida para el destinatario.<br>
- El precio no incluye los refuerzos de acero necesarios para la ejecucion.<br>
- Esta cotizacion es realizada con base en los planos enviados por el cliente; cualquier cambio en obra sera notificado para ajuste de volumen.</div>
<div class="aviso"><strong>Aviso Legal</strong>
La informacion contenida en el presente documento tiene caracter exclusivamente tecnico, informativo y referencial, y se proporciona como guia general con base en la experiencia y criterios tecnicos de Waller prefabricados de concreto S.A. de C.V. e Instaladora de muros prefabricados S.A. de C.V. Su aplicacion debera ser evaluada y validada por los responsables tecnicos, estructurales y de ejecucion de cada proyecto, considerando las condiciones particulares de diseno, construccion, normatividad aplicable y uso especifico.<br><br>
En consecuencia, las recomendaciones aqui contenidas podran ser adaptadas, modificadas o complementadas conforme al criterio tecnico correspondiente y no constituyen una instruccion obligatoria ni garantizan resultados especificos. Waller no asume responsabilidad alguna por el uso, interpretacion, adaptacion o implementacion de la informacion contenida en este documento.<br><br>
Para cualquier aplicacion definitiva, debera realizarse una revision y validacion especifica por parte de profesionales competentes y autorizados.</div>
${ftr}</div>
</body></html>`;

    browser = await puppeteer.launch({
      executablePath: '/usr/bin/chromium-browser',
      args: ['--no-sandbox','--disable-setuid-sandbox','--disable-dev-shm-usage','--disable-gpu']
    });
    const page = await browser.newPage();
    await page.setContent(html, { waitUntil: 'networkidle0', timeout: 30000 });
    const pdf = await page.pdf({ format: 'A4', printBackground: true, margin: { top: 0, right: 0, bottom: 0, left: 0 } });
    await browser.close();

    res.set({
      'Content-Type': 'application/pdf',
      'Content-Disposition': `attachment; filename="${filename || 'cotizacion'}.pdf"`,
      'Content-Length': pdf.length
    });
    res.end(pdf);

  } catch (err) {
    if (browser) await browser.close().catch(() => {});
    console.error('PDF error:', err.message);
    res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 8080;
app.listen(PORT, () => console.log('Waller v2 running on port ' + PORT));
