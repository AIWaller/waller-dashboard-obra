# Waller Dashboard — Control de Obra

## Servidor
- **DigitalOcean:** `159.65.223.41:8085` (staging)
- **Carpeta:** `/root/waller-obra-8085/`
- **PM2:** `pm2 restart waller-obra-8085`
- **Repo:** `AIWaller/waller-dashboard-obra`

## Stack
- **Backend:** Node.js + Express (`obras-routes.js` + `server.js`)
- **Frontend:** HTML plano en `/public/`
- **Sin base de datos** — JSON planos por CC en `/root/data/proyectos/<CC>.json`
- **Archivos subidos:** `/root/data/uploads/`
- **Node:** >= 18.0.0
- **Librerías clave:** `xlsx`, `mammoth` (Word), `pdf-parse` (PDF)

## Proyectos activos (CC)
| CC  | Obra              |
|-----|-------------------|
| 510 | —                 |
| 511 | —                 |
| 512 | —                 |
| 513 | —                 |
| 514 | —                 |
| 515 | —                 |
| 516 | —                 |
| 517 | —                 |
| 518 | —                 |
| 519 | Matilde (activo)  |
| 520 | —                 |

## Rutas críticas

### Datos
```
/root/data/proyectos/<CC>.json   → JSON de cada obra (fuente de verdad)
/root/data/uploads/              → Archivos subidos (xlsx, pdf, docx)
```

### Backend
```
/root/waller-obra-8085/obras-routes.js   → Toda la lógica: parsers + endpoints
/root/waller-obra-8085/server.js         → Express setup, middlewares, static
/root/waller-obra-8085/public/           → Frontend HTML
```

### Endpoints principales
```
GET  /api/proyectos/:cc/metricas       → KPIs + EVM + estado resultados
GET  /api/proyectos/:cc/variaciones    → Variaciones de contrato
GET  /api/proyectos/:cc/evm            → Curva EV vs AC
GET  /api/proyectos/:cc/flujo-caja     → Flujo de caja
GET  /api/proyectos/:cc/cronograma     → Programa semanal + Curva S
GET  /api/corporativo/resumen          → Portafolio consolidado
POST /api/proyectos/:cc/upload         → Subir archivos (contrato/cronograma/cotización/estimación)
```

## Parsers (en obras-routes.js)

| Función                  | Tipo de archivo        | Qué extrae                          |
|--------------------------|------------------------|-------------------------------------|
| `parseCronograma()`      | XLSX cronograma        | `programaSemanal`, `fechaInicio`    |
| `parseCotizacionWaller()`| XLSX cotización        | `secciones`, `complementarios`      |
| `parseEstimacion()`      | XLSX estimación        | Partidas, importes, avance          |
| `parseContratoIA()`      | PDF / DOCX contrato    | Monto, fechas, plazo, alcance       |

## Estructura JSON de obra (campos clave)
```json
{
  "cc": "519",
  "nombre": "Matilde",
  "contrato": { "montoContrato": 0, "fechaInicio": "YYYY-MM-DD", "plazoSemanas": 0 },
  "cronograma": {
    "programaSemanal": { "Semana 1": 189.48, "Semana 2": 310.69 },
    "fechaInicio": "2026-03-30",
    "fechaFin": "YYYY-MM-DD",
    "totalSemanas": 21,
    "totalM2": 0
  },
  "cotizacion": {
    "secciones": [],
    "complementarios": [],
    "resumenCostos": {},
    "totalPaneles": 0,
    "totalPresupuesto": 0
  },
  "estimaciones": [],
  "bitacora": [],
  "avanceDiario": []
}
```

## Reglas de desarrollo (NO violar)
1. **Sin localStorage** — todo el estado viene del servidor
2. **Sin hardcodear CCs** — siempre dinámico por parámetro `:cc`
3. **Cada CC aislado** — nunca mezclar datos entre proyectos
4. **Cálculos en servidor** — `calcularMetricas()` en `obras-routes.js`, nunca en frontend
5. **Backup antes de tocar JSON de producción** — `cp /root/data/proyectos/<CC>.json /root/data/proyectos/<CC>.json.bak`
6. **`pm2 restart waller-obra-8085`** al terminar cada cambio en el servidor
7. **Commit a GitHub** después de cada tarea completada

## Estado actual (Junio 2026)

### ✅ Funcionando
- Motor EVM: CPI, EAC, curva EV vs AC
- Portafolio corporativo con KPIs consolidados
- Pestaña Contrato dinámica (lee JSON)
- Bitácora con formulario
- Parser contrato: PDF y Word (.docx)
- Gráfica egresado vs presupuesto (barras dobles)
- Estado de resultados en Resumen
- Archivos: subida y eliminación individual
- Formulario revisión antes de guardar contrato

### 🔧 Pendiente
1. **Parser cronograma** — `programaSemanal` no extrae correctamente del XLSX Matilde
   - Fila 15 (idx 14): etiquetas `Semana 1...Semana 21`
   - Fila 17 (idx 16): fechas reales por columna
   - Filas 19,30,41... (cada 11): niveles NPB,N1...S2 con m² por día
   - Resultado esperado: totales consolidados `{ "Semana 1": 189.48, ... }`
2. **Cotización — secciones vacías** — parser no lee secciones de paneles del Excel
3. **Curva S** — se activa al resolver punto 1
4. **Avance por nivel** — campo `avancePorNivel` no definido aún

## Flujo de trabajo
1. **Una sesión = una tarea concreta**
2. Consultar este archivo al inicio de cada sesión
3. Hacer backup de JSON antes de cambios en datos
4. `pm2 restart` + verificar endpoint después de cada cambio
5. Commit con mensaje descriptivo al terminar

## Notas de sesión
_(Actualizar aquí decisiones importantes)_

- **2026-06**: Fix parser cronograma (programaSemanal por semana consolidado), fix cotización secciones
