const express = require('express');
const fs = require('fs');
const app = express();

app.use(express.json());
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

const PORT = process.env.PORT || 8080;
app.listen(PORT, () => console.log('Waller running on port ' + PORT));
