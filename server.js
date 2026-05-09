const express = require('express');
const fs = require('fs');
const path = require('path');
const app = express();

app.use(express.json({limit: '50mb'}));
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
  const c = getCounter();
  c.value += 1;
  saveCounter(c.value);
  res.json(c);
});

// Generate PDF using puppeteer
app.post('/api/pdf', async (req, res) => {
  try {
    const puppeteer = require('puppeteer');
    const { html, filename } = req.body;
    
    const browser = await puppeteer.launch({
      args: ['--no-sandbox', '--disable-setuid-sandbox']
    });
    const page = await browser.newPage();
    await page.setContent(html, { waitUntil: 'networkidle0' });
    
    const pdf = await page.pdf({
      format: 'A4',
      printBackground: true,
      margin: { top: 0, right: 0, bottom: 0, left: 0 }
    });
    
    await browser.close();
    
    res.set({
      'Content-Type': 'application/pdf',
      'Content-Disposition': `attachment; filename="${filename || 'cotizacion'}.pdf"`,
      'Content-Length': pdf.length
    });
    res.send(pdf);
  } catch (err) {
    console.error('PDF error:', err);
    res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 8080;
app.listen(PORT, () => console.log('Waller running on port ' + PORT));
