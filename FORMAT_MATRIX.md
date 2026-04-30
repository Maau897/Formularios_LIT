# Matriz De Variables F-LIT

Referencia operativa para seguir digitalizando formato por formato sin asumir que todas las hojas de una misma familia son iguales.

## F-LIT-21-03

Formato: `Congeladores`

Equipos detectados:
- `CONG-1`
- `CONG-2`
- `CONG-3`
- `CONG-5`
- `CONG-6`
- `CONG-7`

Campos variables por hoja:
- `Laboratorio`
- `Marca`
- `Modelo`
- `No. Serie`
- `Inventario / Código`
- `Condición normal`
- `Mínima`
- `Máxima`
- `Rangos de corrección`
- `Factores de corrección`

Hallazgos:
- Todas las hojas usan tres bandas de corrección.
- Las bandas cambian entre equipos.
- No conviene reutilizar un set fijo de rangos para toda la familia.

## F-LIT-22-03

Formato: `Refrigeradores`

Equipos detectados:
- `REFR-1`
- `REFR-2`
- `REFR-3`
- `REFR-4`

Campos variables por hoja:
- `Laboratorio`
- `Marca`
- `Modelo`
- `No. Serie`
- `Inventario / Código`
- `Condición normal`
- `Mínima`
- `Máxima`
- `Rangos de corrección`
- `Factores de corrección`

Hallazgos:
- Todas las hojas usan tres bandas.
- Las bandas y factores cambian por refrigerador.
- Ejemplo: `REFR-3` usa `1 - 4`, `4 - 7`, `7 - 10`.

## F-LIT-20-03

Formato: `Ultracongeladores`

Equipos detectados:
- `ULCO-1`
- `ULCO-2`
- `ULCO-3`

Campos variables por hoja:
- `Laboratorio`
- `Marca`
- `Modelo`
- `No. Serie`
- `Inventario / Código`
- `Condición normal`
- `Mínima`
- `Máxima`
- `Rangos de corrección`
- `Factores de corrección`

Hallazgos:
- `ULCO-1` y `ULCO-2` muestran `N/A` en el bloque de corrección.
- `ULCO-3` sí usa corrección real.
- `ULCO-3` tiene estas bandas:
  `-85 a -80`, `-80 a -70`, `-70 a -60`
- No se debe asumir que toda la familia `ULCO` comparte la misma lógica de corrección.

## F-LIT-23-03

Formato: `Incubadoras`

Equipos detectados:
- `ICO2-1`
- `ICO2-2`
- `ICO2-3`

Campos variables por hoja:
- `Laboratorio`
- `Marca`
- `Modelo`
- `No. Serie`
- `Inventario/Código`
- `Temperatura normal`
- `Temperatura mínima`
- `Temperatura máxima`
- `%CO2 normal`
- `%CO2 mínima`
- `%CO2 máxima`
- `Rango de corrección de temperatura`
- `Factor de corrección`

Hallazgos:
- Las incubadoras sí usan corrección, pero solo para temperatura.
- Todas las hojas revisadas usan una banda `36-38`.
- Además del valor de temperatura, este formato tiene una segunda variable:
  `%CO2`
- El layout del cuerpo no es igual a congeladores/refrigeradores porque inserta una fila extra para `%CO2`.

## F-LIT-09-04

Formato: `Condiciones ambientales`

Hojas detectadas:
- `TEMPERATURA 504`, `HUMEDAD 504`
- `TEMPERATURA 503`, `HUMEDAD 503`
- `TEMPERATURA 506`, `HUMEDAD 506`
- `TEMPERATURA 507`, `HUMEDAD 507`
- `TEMPERATURA 508`, `HUMEDAD 508`
- `TEMPERATURA 502`, `HUMEDAD 502`
- `TEMPERATURA 522`, `HUMEDAD 522`
- `TEMPERATURA 324`, `HUMEDAD 324`
- `TEMPERATURA 514`, `HUMEDAD 514`
- `TEMPERATURA 513`, `HUMEDAD 513`
- `TEMPERATURA 510`, `HUMEDAD 510`
- `TEMPERATURA UBM`, `HUMEDAD UBM`
- `TEMPERATURA UBM (CC)`, `HUMEDAD UBM (CC)`
- `TEMPERATURA UCF`, `HUMEDAD UCF`
- `TEMPERATURA UCF (UCE)`, `HUMEDAD UCF (UCE)`

Campos variables por hoja:
- `Laboratorio`
- `Instrumento`
- `Código`
- `Marca`
- `Modelo`
- `Serie`
- `Condición normal`
- `Mínima`
- `Máxima`
- `Rangos de corrección`
- `Factores de corrección`

Hallazgos:
- `TEMPERATURA` y `HUMEDAD` no son equivalentes uno a uno.
- Cambian posiciones de cabecera y los rangos.
- `TEMPERATURA` revisada:
  `10-15`, `15-20`, `20-25`, `25-30`, `30-35`
- `HUMEDAD` revisada:
  `10-20`, `20-30`, `30-40`, `40-50`, `50-60`, `60-70`, `70-80`, `80-100`
- Este formato debe tratarse como una familia aparte y no como “otro refrigerador”.

## Cierre

Conclusiones prácticas:
- No debemos modelar la corrección “por familia”.
- Debemos leer `rangos`, `factores`, `instrumento/equipo`, `inventario`, `serie` y `layout` por hoja.
- `ULCO-3` es el ejemplo más claro de por qué esto importa.

Archivos de apoyo generados:
- `format_inventory.json`
