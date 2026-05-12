const express = require('express');
const fs = require('fs');
const path = require('path');
const app = express();

app.use(express.json({ limit: '10mb' }));
app.use(express.static('public'));

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

    const logoB64 = (await sharp(path.join(__dirname, 'public/img/logo.jpg')).resize(400).jpeg({quality:90}).toBuffer()).toString('base64');
    const plantaB64 = (await sharp(path.join(__dirname, 'public/img/planta.jpg')).resize(900).jpeg({quality:80}).toBuffer()).toString('base64');
    const axonoB64 = (await sharp(path.join(__dirname, 'public/img/axono.jpg')).resize(900).jpeg({quality:80}).toBuffer()).toString('base64');

    const logoSrc = `data:image/jpeg;base64,${logoB64}`;
    const plantaSrc = `data:image/jpeg;base64,${plantaB64}`;
    const axonoSrc = `data:image/jpeg;base64,${axonoB64}`;
    const D = state;

    const wm = `<svg viewBox="0 0 794 1123" preserveAspectRatio="xMidYMid slice" xmlns="http://www.w3.org/2000/svg"><defs><pattern id="WP" x="0" y="0" width="240" height="120" patternUnits="userSpaceOnUse" patternTransform="rotate(-22)"><image href="${logoSrc}" x="8" y="8" width="120" height="37" opacity="1"/><image href="${logoSrc}" x="128" y="68" width="120" height="37" opacity="1"/></pattern></defs><rect width="794" height="1123" fill="url(#WP)" opacity="0.07"/></svg>`;
    const compRows = D.comp_rows.map((c,i) => `<tr><td>${i+1}</td><td>${c.nombre}</td><td>${c.unidad}</td><td style="text-align:right">${c.cant}</td><td style="text-align:right">${c.pu}</td><td style="text-align:right">${c.imp}</td></tr>`).join('');
    const hdr = `<div class="hdr"><img src="${logoSrc}" alt="Waller"><div class="hdr-meta"><div class="hdr-folio">${D.folio}</div><div class="hdr-fecha">${D.fecha}</div></div></div>`;
    const ftr = `<div class="ftr"><span>Calle Jose Maria Vigil 2808 Int 6, Col. Providencia</span><span>waller.mx | @waller.mx</span><span>Oficina (33) 23 0303 5363</span></div>`;
    const wmDiv = `<div class="wm">${wm}</div>`;

    const html = `<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><title>WALLER</title><style>
*{box-sizing:border-box;margin:0;padding:0;}
body{font-family:Arial,sans-serif;background:#fff;color:#222;}
.page{width:210mm;min-height:297mm;padding:14mm 16mm 12mm;background:#fff;position:relative;overflow:hidden;display:flex;flex-direction:column;page-break-after:always;}
.page:last-child{page-break-after:avoid;}
.wm{position:absolute;top:0;left:0;right:0;bottom:0;pointer-events:none;z-index:0;overflow:hidden;}
.wm svg{width:100%;height:100%;}
.page>*:not(.wm){position:relative;z-index:1;}
.hdr{display:flex;justify-content:space-between;align-items:center;padding-bottom:8px;border-bottom:1.5px solid #2B2B2B;margin-bottom:10px;}
.hdr img{height:36px;width:auto;display:block;}
.hdr-meta{text-align:right;}
.hdr-folio{font-size:13px;font-weight:700;color:#0072BC;letter-spacing:1px;}
.hdr-fecha{font-size:10px;color:#555;margin-top:2px;}
.cbar{display:grid;grid-template-columns:1fr 1.6fr 1.3fr;background:#2B2B2B;color:#fff;margin-bottom:10px;}
.cbar div{padding:6px 10px;border-right:1px solid #444;}
.cbar div:last-child{border-right:none;}
.cbar .lbl{color:#999;font-size:8px;letter-spacing:1.5px;text-transform:uppercase;display:block;margin-bottom:1px;}
.cbar .val{font-weight:600;font-size:11px;}
.intro{font-size:10px;color:#444;margin-bottom:10px;line-height:1.5;}
.slbl{font-size:10px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:#0072BC;margin:10px 0 5px;padding-bottom:3px;border-bottom:1px solid #ddd;}
table{width:100%;border-collapse:collapse;margin-bottom:8px;font-size:10px;}
thead th{background:#2B2B2B;color:#fff;padding:5px 8px;text-align:left;font-size:10px;font-weight:600;}
thead th:last-child{text-align:right;}
thead th:not(:first-child):not(:last-child){text-align:right;}
tbody tr:nth-child(even){background:#f7f7f7;}
tbody td{padding:4px 8px;border-bottom:1px solid #eee;vertical-align:middle;}
tbody td:last-child{text-align:right;}
tbody td:not(:first-child):not(:last-child){text-align:right;}
tfoot td{padding:4px 8px;font-weight:600;font-size:10px;border-top:1.5px solid #ccc;background:#f2f2f2;}
tfoot td:last-child{text-align:right;}
tfoot tr.tr td{background:#0072BC;color:#fff;font-size:12px;font-weight:700;border-top:none;}
.notas{background:#f9f9f9;border-left:3px solid #0072BC;padding:8px 12px;font-size:9px;color:#555;line-height:1.7;margin-top:8px;}
.notas strong{display:block;margin-bottom:3px;color:#2B2B2B;font-size:10px;}
.aviso{background:#fafafa;border-left:3px solid #bbb;padding:8px 12px;font-size:8.5px;color:#888;line-height:1.7;margin-top:6px;flex:1;}
.aviso strong{display:block;margin-bottom:3px;color:#aaa;font-size:9px;}
.ftr{border-top:1px solid #ddd;padding-top:6px;margin-top:auto;display:flex;justify-content:space-between;font-size:8px;color:#aaa;margin-top:8px;}
@page{size:A4 portrait;margin:0;}
</style></head><body>

<div class="page">${wmDiv}${hdr}
<div class="cbar"><div><span class="lbl">Cliente</span><span class="val">${D.cliente}</span></div><div><span class="lbl">Proyecto</span><span class="val">${D.proyecto}</span></div><div><span class="lbl">Ubicacion</span><span class="val">${D.ubicacion}</span></div></div>
<p class="intro">Agradecemos la oportunidad para participar en el proyecto <strong>${D.proyecto}</strong>.<br>A continuacion le mostramos una cotizacion con los volumenes totales y precios vigentes para el desarrollo los muros de su proyecto.<br>El equipo de WALLER agradece su confianza en nosotros.</p>
<div class="slbl">Resumen</div>
<table><thead><tr><th>Concepto</th><th>Importe</th></tr></thead><tbody>
<tr><td>Costo total de Panel Waller (Muros + Losas)</td><td>${D.panel_total}</td></tr>
<tr><td>Costo total de Complementarios</td><td>${D.comp_total}</td></tr>
<tr><td>Costo total de Fletes</td><td>${D.flete_total}</td></tr>
<tr><td>Costo por m2 de materiales (Panel, Complementarios y Fletes)</td><td>${D.m2_costo}</td></tr>
</tbody><tfoot>
<tr><td>Subtotal sin IVA</td><td>${D.subtotal}</td></tr>
<tr><td>IVA (16%)</td><td>${D.iva}</td></tr>
<tr class="tr"><td>IMPORTE TOTAL</td><td>${D.total}</td></tr>
</tfoot></table>
<div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-top:10px;">
  <div>
    <div style="font-size:8px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:#555;margin-bottom:4px;">Planta de Cubierta</div>
    <img src="${plantaSrc}" style="width:100%;border:1px solid #ddd;display:block;">
  </div>
  <div>
    <div style="font-size:8px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:#555;margin-bottom:4px;">Vista Axonometrica</div>
    <img src="${axonoSrc}" style="width:100%;border:1px solid #ddd;display:block;">
  </div>
</div>
${ftr}</div>

<div class="page">${wmDiv}${hdr}
<div class="slbl">Desglose - Muros</div>
<table><thead><tr><th>Concepto</th><th>m2</th><th>$ x m2</th><th>Importe</th></tr></thead><tbody>
<tr><td>Panel Waller 7.5 cm (0.61 x 2.44 m)</td><td>${D.muros_m2}</td><td>${D.muros_pm2}</td><td>${D.muros_imp}</td></tr>
</tbody><tfoot><tr><td colspan="3"><strong>Total m2 con Desperdicio: ${D.muros_m2}</strong></td><td></td></tr></tfoot></table>
<div class="slbl">Desglose - Losas</div>
<table><thead><tr><th>Concepto</th><th>m2</th><th>$ x m2</th><th>Importe</th></tr></thead><tbody>
<tr><td>Panel Waller 7.5 cm (0.61 x 2.44 m)</td><td>${D.losas_m2}</td><td>${D.losas_pm2}</td><td>${D.losas_imp}</td></tr>
</tbody><tfoot><tr><td colspan="3"><strong>Total m2 con Desperdicio: ${D.losas_m2}</strong></td><td></td></tr></tfoot></table>
<div class="slbl">Complementarios</div>
<table><thead><tr><th>#</th><th>Concepto</th><th>Unidad</th><th>Cantidad</th><th>P.Unit.</th><th>Importe</th></tr></thead>
<tbody>${compRows}</tbody>
<tfoot><tr><td colspan="5">Subtotal</td><td>${D.comp_sub}</td></tr><tr><td colspan="5">IVA</td><td>${D.comp_iva}</td></tr><tr class="tr"><td colspan="5">Total</td><td>${D.comp_tot}</td></tr></tfoot></table>
<div class="slbl">Fletes</div>
<table><thead><tr><th>Concepto</th><th>Viajes</th><th>$ x Viaje</th><th>Importe</th></tr></thead><tbody>
<tr><td>Traslado Panel Waller 7.5 cm, Guadalajara planta-obra. NO INCLUYE DESCARGA.</td><td>${D.viajes}</td><td>${D.flete_xv}</td><td>${D.flete_imp}</td></tr>
</tbody><tfoot><tr><td colspan="3">Subtotal</td><td>${D.flete_sub}</td></tr><tr><td colspan="3">IVA</td><td>${D.flete_iva}</td></tr><tr class="tr"><td colspan="3">Total</td><td>${D.flete_tot}</td></tr></tfoot></table>
${ftr}</div>

<div class="page">${wmDiv}${hdr}
<div class="notas"><strong>Notas</strong>
- La entrega se programara de 3 a 5 dias habiles despues del pago total.<br>
- Esta cotizacion tiene vigencia de 30 dias naturales.<br>
- El costo es en moneda nacional (Pesos Mexicanos).<br>
- Esta cotizacion es de caracter confidencial.<br>
- El precio no incluye refuerzos de acero.<br>
- Cualquier cambio en obra sera notificado para ajuste de volumen.</div>
<div class="aviso"><strong>Aviso Legal</strong>
La informacion contenida tiene caracter exclusivamente tecnico e informativo. Su aplicacion debera ser evaluada por los responsables tecnicos de cada proyecto. Waller no asume responsabilidad por el uso de esta informacion. Para cualquier aplicacion definitiva se requiere revision por profesionales autorizados.</div>
${ftr}</div>

</body></html>`;

    browser = await puppeteer.launch({
      executablePath: '/usr/bin/chromium-browser',
      args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage', '--disable-gpu']
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
