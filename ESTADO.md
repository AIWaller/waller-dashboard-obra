# Waller Dashboard — Estado del Sistema y Forma de Trabajo

## Cómo trabajamos
1. Los cambios de código se hacen en Claude Code en DO (/root/waller-obra-8085)
2. Verificamos resultados con la extensión Claude in Chrome apuntando a http://159.65.223.41:8085
3. Después de cada cambio: pm2 restart waller-obra-8085
4. Después de cada sesión: git push a AIWaller/waller-dashboard-obra

## Reglas absolutas — NUNCA romper estas
1. NUNCA modificar JSONs directamente en /root/data/proyectos/
2. NUNCA tocar el puerto 8083 (cotizador independiente)
3. NUNCA hardcodear datos en HTML
4. NUNCA usar localStorage
5. Todo cálculo en servidor (obras-routes.js), frontend solo pinta
6. Cada CC 100% aislado — imposible mezclar datos entre obras
7. Si el parser está bien funciona para CUALQUIER obra, no solo CC 519
8. Los datos se actualizan volviendo a subir archivos desde la UI (cargar.html)

## Errores que ya cometimos — no repetir
- Claude Code modificó JSONs directamente en lugar de arreglar el parser
- git pull de un chat nuevo rompió fixes del chat anterior
- Claude Code buscó datos en ./data/ en lugar de /root/data/
- El nuevo chat agregó login del cotizador al puerto 8085
- Se perdieron fixes al hacer git pull sin verificar qué traía

## Rutas críticas
- Proyecto: /root/waller-obra-8085/
- Datos: /root/data/proyectos/ (un JSON por CC)
- Uploads: /root/data/uploads/
- Catálogo: /root/data/catalogo.json
- Rutas API: /root/waller-obra-8085/obras-routes.js
- Frontend: /root/waller-obra-8085/public/obra.html, obras.html, cargar.html
- PM2: pm2 restart waller-obra-8085
- Repo: AIWaller/waller-dashboard-obra (token en CLAUDE.md — no poner aquí)

## Stack
Node.js + Express + JSON planos por CC. Sin base de datos. Sin autenticación todavía.

## Lo que funciona hoy
- Cada obra lee su propio JSON aislado por CC
- Motor de cálculo en servidor: calcularMetricas() con EVM, variaciones, flujo caja, edo resultados, por nivel
- Endpoints: /api/proyectos/:cc/metricas, /variaciones, /evm, /flujo-caja, /edo-resultados, /por-nivel, /cronograma, /contrato, /estimaciones, /cambios, /bitacora
- Portafolio consolidado: /api/corporativo/resumen
- Parser contrato acepta PDF y Word (.docx)
- Formulario de revisión antes de guardar contrato
- Cronograma: fechaInicio, fechaFin, 21 semanas, programaSemanal extraído del xlsx
- Gráfica egresado vs presupuesto con barras dobles
- Estado de resultados en Resumen
- EVM con CPI, EAC, curva EV vs AC
- Bitácora con formulario funcional
- Eliminar archivos desde sección Archivos

## Pendiente
1. Curva S solo muestra S1 — el render no itera las 21 semanas del programaSemanal
2. Muros/Panel sin presupuesto en Variaciones — cotizacion.secciones[] vacío, parser xlsx cotización no lee secciones de paneles
3. Por nivel sin avance instalado — campo avancePorNivel no existe, pendiente diseñar carga
4. Login/autenticación — pendiente
5. Separar cotizador (8083) y dashboard (8085) en droplets independientes

## Proyectos activos
CC 510 NHAOS, 511 Hidalma, 514 Deimare, 515 Praia, 516 Amatitán, 517 Artwalk, 518 IU Life, 519 Matilde (tiene contrato+cotización+cronograma+estimaciones), 520 Zendera Torre A (tiene contrato+cronograma)

## Para iniciar un chat nuevo
1. Lee este archivo ESTADO.md completo
2. Lee CLAUDE.md
3. NO hagas git pull sin verificar qué cambios trae
4. NO modifiques JSONs directamente
5. Pregunta antes de hacer cambios grandes
