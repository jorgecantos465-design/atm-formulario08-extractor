# Skill: Carga de datos inicial - extractor F08 digital

## Objetivo
Administrar PDFs y plantillas de prueba sin exponer datos sensibles ni romper el MVP.

## Cuando usarla
Usala al agregar/remover muestras, templates o datos de validacion.

## Pasos obligatorios
- Confirmar si el archivo contiene datos reales.
- Preferir muestras anonimizadas.
- No borrar `input/` ni `templates/`.
- Mantener una sola plantilla objetivo cuando el flujo lo requiera.

## Checklist antes de modificar
- Permiso para usar muestra.
- Datos sensibles anonimizados o aceptados.
- Template correcto identificado.

## Checklist antes de finalizar
- Extractor encuentra input/template.
- No hay datos sensibles nuevos sin aviso.
- README sigue cierto.

## Errores que debe evitar
- Subir documentos personales.
- Mezclar plantillas incompatibles.
- Eliminar muestras utiles.

## Comandos de verificacion
- `npm run extractor`
- `npm test`

## Formato de reporte final
- Archivos de muestra.
- Sensibilidad de datos.
- Compatibilidad.
- Validacion.

