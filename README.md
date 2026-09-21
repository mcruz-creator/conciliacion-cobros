# Conciliación de cobros con tarjeta

Página: https://mcruz-creator.github.io/conciliacion-cobros/

Cruza los recibos de caja (columnas TSNS, TSNI, TSNZ y TVIR) contra los movimientos
presentados a Payway, las transferencias QR de Payway y Mercado Pago.

- Los archivos se procesan en el navegador; no se suben a ningún servidor.
- Reglas (sin desempates ni tolerancias):
  1. **Par directo**: misma fecha + mismo importe exacto, único de ambos lados.
  2. **Grupo que cierra**: misma fecha + mismo importe, y la misma cantidad de recibos que de
     movimientos. El grupo cierra y se concilia en bloque; el apareo interno va por orden de carga.
  3. **Diferencia de centavos**: sobre lo que quedó sin conciliar, misma fecha y una diferencia de
     importe de hasta $0,99 en cualquier dirección, siempre que sea única de los dos lados.
  En las tres, TVIR solo se cruza contra MP. Lo que no cierra queda en "Para revisar".
  La columna "Apareo" del Excel indica de cuál de las tres reglas salió cada fila.
- Salida: Excel con Resumen, Conciliados, Para revisar, Recibos sin movimiento y Movimientos sin recibo.

`conciliar.py` es la versión de escritorio (Python) con las mismas reglas.
