# Skills - atm-formulario08-extractor

## Indice

- `roadmap-negocio.md`: alcance operativo del extractor F08 digital.
- `arquitectura-tecnica.md`: pipeline PDF -> plantilla Excel -> logs.
- `motor-matematico-logica.md`: parsing, normalizacion, validacion y mapeo.
- `carga-datos-inicial.md`: PDFs, muestras y templates.
- `ux-flujo.md`: mensajes CLI, errores y trazabilidad.
- `testing-bugs.md`: normalizadores, casos y salida esperada.

## Cuando activar cada skill

- Parsing, mapeo o validacion de campos: `motor-matematico-logica.md` + `testing-bugs.md`.
- Pipeline o estructura de salida: `arquitectura-tecnica.md`.
- Muestras o templates: `carga-datos-inicial.md`.
- Mensajes para operador: `ux-flujo.md`.
- Alcance de campos soportados: `roadmap-negocio.md`.

## Orden recomendado si hay varias capas

1. `roadmap-negocio.md` si cambia alcance.
2. `arquitectura-tecnica.md` si cambia pipeline.
3. `motor-matematico-logica.md` si cambia extraccion/mapeo.
4. `carga-datos-inicial.md` si hay muestras/templates.
5. `ux-flujo.md` si cambian mensajes/logs.
6. `testing-bugs.md` para cierre.

## Matriz de riesgo

- Bajo: README, mensajes, logs no estructurales.
- Medio: regex/normalizadores acotados, formato de logs.
- Alto: templates, placeholders, escritura Excel, PDFs reales, reglas de CUIT/dominio/fecha/monto.

## Comandos seguros

- `npm test`

## Comandos condicionales

- `npm run extractor`: solo con input/template controlados.
- `npm install`: solo si faltan dependencias.

## Comandos prohibidos o que requieren confirmacion

- Tocar templates reales: requiere confirmacion.
- Tocar PDFs reales o sensibles: requiere confirmacion.
- Borrar `input/`, `templates/`, `output/` o `logs/`: requiere confirmacion.
- `npm audit fix`: requiere confirmacion.
- Cambios en `package.json` o lockfile: requieren confirmacion.

## Ownership por archivos criticos

- `src/extract-formulario08.js`: Arquitectura para pipeline; Motor logico para parsing/mapeo; Testing para normalizadores.
- `templates/`: Carga de datos inicial.
- `input/`: Carga de datos inicial.
- `logs/`, mensajes CLI: UX + Testing.
- `README.md`: Roadmap + UX.

## Regla de seguridad

No inventar campos. Si un dato no se encuentra, debe quedar vacio y registrarse en log.

