# Proyecto: Cotizador INFEJAL — Comercial Waller

## Contexto general
Cotizador v2 para INFEJAL.
Estado: **🔧 En ajustes activos**
Producción en: **www.comercialwaller.mx**
Infraestructura: **GitHub + DigitalOcean**
Servidor: `waller-cotizador` (DigitalOcean)

---

## Stack
- **Backend:** Node.js + Express (`server.js`) — entry point principal
- **Frontend:** HTML plano (`/public/index.html`)
- **PDF:** Puppeteer + Puppeteer Core (requiere Chrome instalado)
- **Imágenes:** Sharp
- **Node:** >= 18.0.0 requerido

---

## Estructura del proyecto
```
waller-cotizador/
├── .claude/
├── public/
│   ├── index.html        — UI principal del cotizador
│   ├── index.html.bak    — Backup anterior (no borrar)
│   └── img/              — Imágenes y assets
├── server.js             — Entry point Express
├── package.json
├── counter.json          — Contador de cotizaciones generadas
├── nuevo.pdf             — PDF de prueba
├── sin_wm.pdf            — PDF sin marca de agua (prueba)
└── test.pdf              — PDF de prueba
```

---

## Reglas importantes
- Este es el cotizador **v2 en ajustes** — es el activo de desarrollo
- `index.html.bak` es un backup funcional anterior — **no borrar**
- Los PDFs de prueba (`nuevo.pdf`, `test.pdf`, `sin_wm.pdf`) son para validación
- Chrome debe estar instalado: `npm run build` lo instala vía Puppeteer
- Antes de cambios grandes, revisar si `index.html.bak` está actualizado

---

## Comandos útiles
```bash
# Instalar Chrome para Puppeteer
npm run build

# Iniciar servidor
npm start
```

---

## Flujo de trabajo acordado
1. **Una sesión = una tarea concreta**
2. **Iniciar sesión nueva** al terminar cada tarea
3. **Consultar este archivo** al inicio de cada sesión
4. Actualizar `index.html.bak` antes de cambios grandes

---

## Estado actual
- [x] Estructura base funcionando
- [x] Generación de PDF con Puppeteer
- [ ] Ajustes pendientes ← documentar aquí qué falta

---

## Notas de sesión
_(Agrega aquí decisiones importantes tomadas en cada sesión)_

-
