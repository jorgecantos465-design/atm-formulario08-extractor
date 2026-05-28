# Extractor Formulario 08

MVP deterministico para leer un PDF de Formulario 08 y completar la fila 2 de una copia de una plantilla Excel existente.

## Estructura

- `input/`: colocar un solo PDF de prueba.
- `templates/`: colocar exactamente una plantilla `.xlsx` o `.xlsm`.
- `src/`: codigo del extractor.
- `output/`: se guarda la copia Excel completada.
- `logs/`: se guarda el log de campos encontrados y no encontrados.

## Uso

```powershell
npm install
npm run extractor
```

`npm run extract-f08` queda como alias compatible.

## Reglas del MVP

- Lee exactamente un PDF desde `input/`.
- No crea un Excel desde cero.
- Abre la plantilla existente desde `templates/`.
- No cambia encabezados, nombres de columnas ni placeholders.
- Detecta las columnas por los placeholders existentes.
- Completa solo la fila 2.
- Si la fila 2 contiene placeholders, se detiene para no romper la plantilla.
- Los campos `@atributo23@`, `@atributo24@`, `@atributo25@`, `@atributo26@`, `@atributo27@`, `@atributo32@` y `@atributo34@` quedan vacios.
- Si un dato no se encuentra en el PDF, queda vacio y se registra en el log.

## Campos extraidos

- `@usuario@`, `@atributo8@`: CUIT adquirente.
- `@apellido@`: apellido adquirente.
- `@nombre@`: nombres adquirente.
- `@email@`: email en minusculas.
- `@atributo9@`: nombre completo adquirente.
- `@atributo14@`: numero de Formulario 08.
- `@atributo15@`: dominio validado como `ABC123` o `AB123CD`.
- `@atributo16@`: domicilio legal o, si falta, domicilio real.
- `@atributo17@`: fecha de impresion `DD/MM/AAAA`.
- `@atributo18@`: modelo.
- `@atributo19@`: modelo/año si existe; si no, modelo.
- `@atributo20@`: año.
- `@atributo21@`: CUIT vendedor.
- `@atributo22@`: nombre vendedor.
- `@atributo33@`: monto operacion como numero.

## Nota sobre ODS

El MVP rechaza `.ods` para no reescribirlo con una libreria que pueda alterar estructura o estilos. Para esta primera version, usar `.xlsx` o `.xlsm`.
