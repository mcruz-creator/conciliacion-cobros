# Conciliación de cobros con tarjeta

Página: https://mcruz-creator.github.io/conciliacion-cobros/

Cruza los recibos de caja (columnas TSNS, TSNI, TSNZ y TVIR) contra los movimientos
presentados a Payway, las transferencias QR de Payway y Mercado Pago.

- Los archivos se procesan en el navegador; no se suben a ningún servidor.
- Regla: misma fecha + mismo importe exacto, única de ambos lados. TVIR solo contra MP.
  Todo lo ambiguo queda en "Para revisar". Sin desempates ni tolerancias.
- Salida: Excel con Resumen, Conciliados, Para revisar, Recibos sin movimiento y Movimientos sin recibo.

`conciliar.py` es la versión de escritorio (Python) con la misma regla.
