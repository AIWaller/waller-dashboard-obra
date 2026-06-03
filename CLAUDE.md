# Proyecto: Cotizador INFEJAL — Comercial Waller

## Contexto general
Cotizador para INFEJAL, integrado dentro del Dashboard principal.
Estado: **🔧 Verificación pendiente — revisar que todo esté bien**
Servidor: `waller-cotizador` (DigitalOcean)
Puerto: **8082** (mismo que dashboard — mismo proyecto)

---

## Stack
- **Backend:** Node.js + Express (`server.js`)
- **Frontend:** HTML plano (`/public/index.html`)
- **PDF:** Puppeteer + Puppeteer Core
- **Imágenes:** Sharp
- **Node:** >= 18.0.0

---

## Ubicación en servidor
Este cotizador vive dentro de `waller-cotizador-v2` junto con el dashboard.
```
waller-cotizador-v2/public/
├── index.html        — Cotizador INFEJAL (este archivo)
├── index.html.bak    — Backup anterior (no borrar)
└── ...resto del dashboard
```

## Respaldo
- `waller-cotizador` en puerto 8080 es la versión anterior de respaldo
- Si algo falla, ahí está la versión vieja funcional

---

## Reglas importantes
- Cambios en `index.html` afectan producción directamente
- Siempre actualizar `index.html.bak` antes de cambios grandes
- El respaldo en puerto 8080 es solo referencia — no modificar

---

## Estado actual
- [x] En producción en puerto 8082
- [ ] Verificación general pendiente — revisar que todo esté bien
- [ ] Pendiente separar en repo propio — futuro

---

## Flujo de trabajo
1. **Una sesión = una tarea concreta**
2. **Iniciar sesión nueva** al terminar cada tarea
3. **Consultar este archivo** al inicio de cada sesión
4. Backup de `index.html` antes de cada cambio
5. Anotar en "Notas de sesión" qué cambió

---

## Notas de sesión
_(Agrega aquí decisiones importantes tomadas en cada sesión)_

-
