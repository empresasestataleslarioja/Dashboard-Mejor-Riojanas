# Regresión de parsers de import

`parsers.test.js` corre en Node (no en el navegador) contra el código real de
`../index.html`: extrae las funciones de parseo (`tryParseVentasGastos`,
`tryParseERColumnasMeses`, `tryParseERMatricial`, etc.) y las ejecuta contra
workbooks armados en memoria que reproducen la **estructura** de bugs reales
ya corregidos — nunca datos reales de ninguna empresa.

Sirve para que esa clase de bug (fila de meses fuera de la ventana de
búsqueda, abreviatura de rótulo no reconocida, año no detectado y fila
descartada en silencio, nombre de empresa mal derivado) no vuelva a
reaparecer sin que algo lo marque.

## Correr

```
npm install
npm test
```

## Si un test falla después de tocar un parser

El fallo señala exactamente qué contrato se rompió — no lo hagas pasar
ajustando el assert; ajustá el código o, si el cambio de comportamiento es
intencional, actualizá el test y dejá en el mensaje del commit por qué.

## Si agregás un parser nuevo o encontrás un bug real de "no reconoce este
formato"

Sumá un `group()` con la estructura mínima que lo reproduce (no el archivo
real del usuario) y el resultado esperado. `extractFn()` lee la función
directamente de `index.html`, así que no hay que mantener una copia
separada del parser en el test.
