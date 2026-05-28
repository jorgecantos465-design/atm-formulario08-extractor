# Skill: Motor matematico / logica - extractor F08 digital

## Objetivo
Proteger reglas de extraccion, normalizacion, validacion de campos y mapeo a placeholders.

## Cuando usarla
Usala al tocar regex, CUIT, dominio, fechas, monto, nombres o placeholders.

## Pasos obligatorios
- Definir ejemplos validos e invalidos.
- Validar normalizacion sin perder evidencia.
- Mantener campos no encontrados vacios.
- Preservar placeholders existentes.

## Checklist antes de modificar
- Placeholder objetivo identificado.
- Regla actual entendida.
- Caso esperado escrito.

## Checklist antes de finalizar
- `npm test` pasa.
- Campos criticos mantienen formato esperado.
- No se completan datos dudosos.

## Errores que debe evitar
- Aceptar dominios invalidos.
- Reescribir fechas ambiguas sin certeza.
- Llenar campos definidos como vacios por MVP.

## Comandos de verificacion
- `npm test`
- `npm run extractor`

## Formato de reporte final
- Regla tocada.
- Ejemplos.
- Resultado antes/despues.
- Riesgo.

