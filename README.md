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
  4. **Neteo de anulaciones**: después de las tres anteriores, un recibo positivo y uno negativo
     sin conciliar, de la misma fecha, el mismo cliente y el mismo importe, se cancelan entre sí y
     salen del informe; después se vuelven a correr las reglas 1 a 3. Los negativos que ya
     conciliaron contra una devolución real del procesador no se netean.
  5. **Pagador Grupo Oroño**: antes de conciliar se apartan los movimientos cuyo campo
     "Pagador:" es Grupo Oroño, porque los hace la empresa y no salen de un recibo. Solo se
     mira el pagador: el texto "Producto de Grupooroño" no excluye nada.
  6. **Videoconsulta**: un movimiento cuyo detalle dice "Videoconsulta médica" solo concilia contra un
     recibo TVIR, nunca contra uno de mostrador, aunque la fecha y el importe coincidan.
  7. **Rendimientos de Mercado Pago**: los movimientos de MP que vienen sin medio de pago son el
     rendimiento diario de la cuenta (tampoco tienen local, caja, pagador ni comisión). Se apartan
     antes de conciliar. Se mira el medio de pago y no el número de identificación: ese también
     viene vacío en cobros reales de pacientes.
  En las tres primeras, TVIR solo se cruza contra MP. Lo que no cierra queda en "Para revisar".
  La columna "Apareo" del Excel indica de cuál de las tres primeras reglas salió cada fila.
- Salida: Excel con Resumen, Conciliados, Anulados, Pagador Grupo Oroño, Rendimientos MP, Hoja de trabajo, Para revisar,
  Recibos sin movimiento y Movimientos sin recibo. La **Hoja de trabajo** junta todo lo pendiente
  (las tres últimas hojas) agrupado por fecha e importe, con las columnas de "Para revisar" y una
  columna "Estado" que dice de cuál de las tres viene cada fila. En ambas hojas el importe del
  recibo y el del movimiento van en columnas separadas, para poder sumar cada lado por su cuenta.

`conciliar.py` es la versión de escritorio (Python) con las mismas reglas.
