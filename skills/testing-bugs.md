# Skill: Testing y bugs - extractor F08 digital

## Objetivo
Detectar fallas de parsing, normalizacion, validacion y mapeo a planilla.

## Cuando usarla
Usala ante PDFs nuevos, errores de campo, cambios en normalizadores o logs.

## Pasos obligatorios
- Reproducir con archivo de `input/`.
- Identificar campo esperado y valor obtenido.
- Ejecutar normalizadores.
- Verificar salida Excel y log.

## Checklist antes de modificar
- Dato sensible protegido.
- Resultado esperado escrito.
- Campo/placeholder identificado.

## Checklist antes de finalizar
- `npm test` pasa.
- `npm run extractor` procesa caso controlado.
- Log registra faltantes.

## Errores que debe evitar
- Ajustar regex para un solo PDF rompiendo otros.
- Inventar campos no encontrados.
- Ignorar formatos validos de dominio/CUIT.

## Comandos de verificacion
- `npm test`
- `npm run extractor`

## Formato de reporte final
- PDF/caso.
- Campo afectado.
- Causa.
- Verificacion.

