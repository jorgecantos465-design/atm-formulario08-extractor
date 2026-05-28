# AGENTS.md - atm-formulario08-extractor

## Objetivo del proyecto

Extractor deterministico para leer PDFs de Formulario 08 y completar la fila 2 de una copia de una plantilla Excel existente, respetando placeholders y estructura del archivo.

## Stack tecnico detectado

- Node.js
- JavaScript CommonJS
- `pdf-parse` para lectura de PDF
- `exceljs` para manipular `.xlsx`/`.xlsm`
- Scripts CLI por npm

## Comandos

- Instalacion: `npm install`
- Ejecutar extractor: `npm run extractor`
- Alias: `npm run extract-f08`
- Test normalizadores: `npm test`
- Build: no hay script de build detectado
- Dev server: no aplica

## Estructura y archivos criticos

- `src/extract-formulario08.js`: extractor, normalizadores, mapeo y escritura de salida.
- `input/`: PDFs de entrada.
- `templates/`: plantilla Excel fuente; el MVP espera `.xlsx` o `.xlsm`.
- `output/`: Excel generado.
- `logs/`: logs de campos encontrados/no encontrados.
- `README.md`: reglas funcionales del MVP.

## Reglas de trabajo

- No tocar `node_modules`, `output` ni `logs` salvo pedido explicito.
- No modificar `package.json` sin avisar antes.
- No borrar archivos de `input/` o `templates/`.
- No cambiar encabezados, placeholders ni fila objetivo de la plantilla sin autorizacion.
- Mantener comportamiento deterministico: no inventar datos y registrar faltantes.
- Preservar rechazo de `.ods` salvo decision explicita.

## No tocar sin autorizacion

- Plantillas en `templates/`.
- PDFs reales o de prueba en `input/`.
- Mapeo de placeholders criticos.
- Campos que hoy deben quedar vacios por definicion del MVP.
- Reglas de validacion de dominio, CUIT, fechas y monto.

## Estrategia de commits

- Separar cambios por tipo: parsing PDF, normalizacion, mapeo Excel o logging.
- Todo cambio de normalizador debe pasar `npm test`.
- Incluir ejemplos de entrada/salida solo si no contienen datos sensibles.

## Criterios de finalizacion

- `npm test` pasa.
- `npm run extractor` genera salida esperada con un input controlado.
- La plantilla conserva estructura, estilos y placeholders fuera de la fila objetivo.
- El log informa campos no encontrados.

## Subagentes recomendados

Antes de activar un rol, revisar el indice operativo en `skills/README.md`.

- Roadmap y negocio: usar `skills/roadmap-negocio.md`.
- Arquitectura tecnica: usar `skills/arquitectura-tecnica.md`.
- Motor matematico / logica de negocio: usar `skills/motor-matematico-logica.md`.
- Datos iniciales / seed: usar `skills/carga-datos-inicial.md`.
- Testing y bugs: usar `skills/testing-bugs.md`.
- UX y flujo: usar `skills/ux-flujo.md`.
