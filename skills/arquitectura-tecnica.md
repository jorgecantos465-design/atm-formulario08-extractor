# Skill: Arquitectura tecnica - extractor F08 digital

## Objetivo
Mantener CLI deterministico, simple y seguro para leer PDF y escribir sobre copia de plantilla Excel.

## Limite arquitectonico
- Este repositorio acepta solo Formularios 08 digitales con texto embebido.
- Pipeline permitido: `pdf-parse` -> normalizacion -> regex/parser -> mapeo -> salida.
- Prohibido usar Vision, modelos multimodales o IA paga.
- Prohibido sumar OCR o logica para manuscritos.
- Los manuscritos pertenecen a `atm-formulario08-extractor-manuscrito`.

## Cuando usarla
Usala al cambiar pipeline, lectura PDF, escritura Excel, estructura de logs o manejo de archivos.

## Pasos obligatorios
- Leer `src/extract-formulario08.js`.
- Identificar input, template, output y logs.
- No alterar plantilla fuente.
- Mantener errores explicitos.
- Rechazar propuestas que mezclen OCR, Vision o IA paga con este pipeline.

## Checklist antes de modificar
- No se toca `package.json`.
- Plantilla no sera sobrescrita.
- Comportamiento con un solo PDF entendido.

## Checklist antes de finalizar
- `npm test` pasa.
- CLI conserva entradas/salidas esperadas.
- Logs siguen generandose.

## Errores que debe evitar
- Escribir sobre archivo de template.
- Crear Excel desde cero.
- Reescribir o convertir destructivamente una plantilla `.ods`.
- Incorporar logica manuscrita o dependencias de Vision.

## Comandos de verificacion
- `npm test`
- `npm run extractor`

## Formato de reporte final
- Pipeline afectado.
- Archivos de entrada/salida.
- Validaciones.
- Riesgos.
