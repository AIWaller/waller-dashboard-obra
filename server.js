const express = require('express');
const path = require('path');
const app = express();

app.use(express.json({ limit: '10mb' }));

// Raíz → dashboard de obras (antes del static para que no lo intercepte index.html)
app.get('/', (req, res) => res.redirect('/obras.html'));

app.use(express.static(path.join(__dirname, 'public')));

require('./obras-routes')(app);

const PORT = process.env.PORT || 8085;
app.listen(PORT, () => console.log('Waller obras running on port ' + PORT));
