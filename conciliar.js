/*
 * Conciliación de cobros con tarjeta / QR / Mercado Pago - Grupo Oroño
 *
 * Reglas de conciliación (sin desempates ni tolerancias):
 *   1. Par directo: misma FECHA y mismo IMPORTE exacto, y esa combinación
 *      identifica a un único recibo y a un único movimiento.
 *   2. Grupo que cierra: dentro de una misma fecha e importe, si la cantidad de
 *      recibos es exactamente igual a la cantidad de movimientos, el grupo cierra
 *      y se concilia en bloque. El apareo dentro del grupo es por orden de carga
 *      (todos son del mismo día y del mismo importe): el total del grupo está
 *      verificado, el par individual no.
 *   3. Diferencia de centavos: sobre lo que quedó sin conciliar, misma FECHA y
 *      una diferencia de importe de hasta $0,99 en cualquier dirección, siempre
 *      que sea única de los dos lados (el QR redondea al peso entero).
 *   4. Neteo de anulaciones: después de las tres anteriores, si entre los recibos
 *      que quedaron SIN conciliar hay, en la misma FECHA y para el mismo CLIENTE,
 *      uno positivo y uno negativo del mismo importe, los dos se cancelan entre sí
 *      y salen del informe (la anulación nunca llegó al procesador). Después se
 *      vuelven a correr las reglas 1 a 3 sobre lo que quedó. Los recibos negativos
 *      que YA conciliaron contra una devolución real no se netean nunca.
 *   En las tres primeras, las columnas "solo MP" (videoconsultas) solo se cruzan contra
 *   Mercado Pago; un grupo que empata en cantidad pero no tiene suficientes
 *   movimientos de MP para sus videoconsultas NO cierra. Todo lo demás queda
 *   para revisión manual.
 *
 * Funciona en el navegador (usa window.XLSX) y en Node (para pruebas).
 */
(function (root) {
  "use strict";

  const XLSX = root.XLSX || (typeof require !== "undefined" ? require("xlsx-js-style") : null);

  const CONFIG_DEFAULT = {
    colsTarjeta: ["TSNS", "TSNI", "TSNZ"], // amarillas: tarjeta / QR
    colsSoloMP: ["TVIR"], // naranja: videoconsultas, solo Mercado Pago
  };

  const MODO_PAR = "Par directo";
  const MODO_TOL = "Diferencia hasta $0,99";
  const TOLERANCIA = 99; // centavos

  // ------------------------------------------------------------ utilidades
  const cents = (v) => Math.round(Number(v) * 100);
  const pesos = (c) => c / 100;

  function toNum(v) {
    if (v === null || v === undefined || v === "") return NaN;
    if (typeof v === "number") return v;
    return Number(String(v).trim());
  }

  function serialAFecha(serial) {
    // fecha de Excel (número de serie) -> "AAAA-MM-DD"
    const d = new Date(Date.UTC(1899, 11, 30) + Math.round(serial) * 86400000);
    return d.toISOString().slice(0, 10);
  }

  function ddmmyyyyAIso(s) {
    const [d, m, y] = String(s).trim().split("/");
    return `${y}-${m.padStart(2, "0")}-${d.padStart(2, "0")}`;
  }

  function decodificar(bytes) {
    try {
      return new TextDecoder("utf-8", { fatal: true }).decode(bytes);
    } catch (e) {
      return new TextDecoder("windows-1252").decode(bytes);
    }
  }

  function parseCsv(texto, saltear) {
    const lineas = texto.replace(/^﻿/, "").split(/\r?\n/).filter((l) => l.trim() !== "");
    const filas = lineas.slice(saltear).map((l) => l.split(";").map((c) => c.trim().replace(/^'|'$/g, "")));
    const cab = filas[0];
    return filas.slice(1).map((f) => Object.fromEntries(cab.map((c, i) => [c.trim(), f[i] ?? ""])));
  }

  function leerLibro(bytes) {
    return XLSX.read(bytes, { type: "array" });
  }

  function filasHoja(wb) {
    return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true, defval: "" });
  }

  // ------------------------------------------------------------ detección
  function detectar(nombre, bytes) {
    const ext = nombre.toLowerCase().split(".").pop();
    if (ext === "csv") {
      const cab = decodificar(bytes.slice(0, 3000));
      if (cab.includes("Detalle de Transferencias") || cab.includes("QR ID")) return "qr";
      if (cab.includes("NUM.CUPON")) return "payway";
      return null;
    }
    if (ext === "xls" || ext === "xlsx") {
      const filas = filasHoja(leerLibro(bytes)).slice(0, 8);
      const txt = filas.map((f) => f.join(" ")).join(" ");
      if (txt.includes("Listado de caja")) return "recibos";
      if (txt.toUpperCase().includes("MERCADO PAGO")) return "mp";
    }
    return null;
  }

  // ------------------------------------------------------------ lectura
  function leerRecibos(bytes, cfg) {
    const filas = filasHoja(leerLibro(bytes));
    const iCab = filas.findIndex((f) => String(f[0]).trim() === "Cajero");
    if (iCab < 0) throw new Error("No se encontró el encabezado 'Cajero' en el listado de caja");
    const cab = filas[iCab].map((c) => String(c).trim());
    const col = (n) => cab.indexOf(n);
    const cols = [...cfg.colsTarjeta, ...cfg.colsSoloMP];
    const faltan = cols.filter((c) => col(c) < 0);
    const salida = [];
    for (const f of filas.slice(iCab + 1)) {
      if (f[col("Fecha Caja")] === "" || f[col("Nro Recibo")] === "") continue;
      for (const c of cols) {
        if (col(c) < 0) continue;
        const v = toNum(f[col(c)]);
        if (!isNaN(v) && v !== 0) {
          const fc = f[col("Fecha Caja")];
          salida.push({
            fecha: typeof fc === "number" ? serialAFecha(fc) : ddmmyyyyAIso(fc),
            nroRecibo: String(f[col("Nro Recibo")]).trim(),
            cliente: String(f[col("Nombre Cliente")]).trim(),
            cajero: String(f[col("Cajero")]).trim(),
            columna: c,
            tipoAsiento: String(f[col("Tipo de Asiento")]).trim(),
            importe: cents(v),
          });
        }
      }
    }
    return { filas: salida, faltan };
  }

  function leerPayway(bytes) {
    return parseCsv(decodificar(bytes), 0).map((r) => ({
      fuente: "PAYWAY",
      fecha: ddmmyyyyAIso(r["COMPRA"]),
      importe: cents(r["MONTO_BRUTO"]),
      referencia: `Lote ${r["LOTE"]} / Cupón ${r["NUM.CUPON"]} / Aut ${r["NRO_AUT"]}`,
      detalle: `${r["TIPO"]} - ${r["MARCA"]} - ${r["DETALLE"]} - Tarj ${String(r["NUM.TARJETA"]).slice(-4)}`,
      terminal: `Estab ${r["ESTABLECIMIENTO"]}`,
    }));
  }

  function leerQr(bytes) {
    return parseCsv(decodificar(bytes), 1).map((r) => ({
      fuente: "QR",
      fecha: String(r["FECHA DE VENTA"]).slice(0, 10),
      importe: cents(r["MONTO BRUTO"]),
      referencia: `QR ${r["QR ID"]} / Cupón ${r["NRO CUPON"]}`,
      detalle: r["ESTADO"],
      terminal: `Terminal ${r["TERMINAL LOGICA"]}`,
    }));
  }

  function leerMp(bytes) {
    const filas = filasHoja(leerLibro(bytes));
    const cab = filas[0].map((c) => String(c).trim());
    const col = (n) => cab.indexOf(n);
    const g = (f, n) => (col(n) < 0 ? "" : String(f[col(n)] ?? "").trim());
    return filas.slice(1).filter((f) => g(f, "FECHA DE ORIGEN") !== "").map((f) => ({
      fuente: "MP",
      fecha: g(f, "FECHA DE ORIGEN").slice(0, 10),
      importe: cents(toNum(f[col("VALOR DE LA COMPRA")])),
      referencia: `MP ${g(f, "ID DE OPERACIÓN EN MERCADO PAGO")}`,
      detalle: [g(f, "TIPO DE OPERACIÓN"), g(f, "MEDIO DE PAGO"),
                g(f, "DETALLE DE LA VENTA").replace(/"/g, ""),
                "Pagador: " + g(f, "PAGADOR")].join(" - "),
      terminal: [g(f, "NOMBRE DE LOCAL"), g(f, "NOMBRE DE CAJA")].join(" ").trim(),
    }));
  }

  // ------------------------------------------------------------ conciliación
  function conciliar(recibos, movs, cfg) {
    const clave = (x) => `${x.fecha}|${x.importe}`;
    const porClave = new Map();
    movs.forEach((m, i) => {
      const k = clave(m);
      if (!porClave.has(k)) porClave.set(k, []);
      porClave.get(k).push(i);
    });

    // todos los pares posibles recibo-movimiento
    const candRec = recibos.map((r) =>
      (porClave.get(clave(r)) || []).filter((j) => !cfg.colsSoloMP.includes(r.columna) || movs[j].fuente === "MP"));
    const candMov = movs.map(() => []);
    candRec.forEach((js, i) => js.forEach((j) => candMov[j].push(i)));

    const pares = [];
    recibos.forEach((r, i) => {
      r.estado = candRec[i].length ? "Para revisar" : "Sin movimiento";
      r.modo = "";
    });
    movs.forEach((m, j) => {
      m.estado = candMov[j].length ? "Para revisar" : "Sin recibo";
      m.modo = "";
    });
    candRec.forEach((js, i) => {
      if (js.length === 1 && candMov[js[0]].length === 1) {
        recibos[i].estado = "Conciliado";
        movs[js[0]].estado = "Conciliado";
        recibos[i].modo = movs[js[0]].modo = MODO_PAR;
        pares.push([recibos[i], movs[js[0]], MODO_PAR]);
      }
    });

    // regla 2: grupos de la misma fecha e importe que empatan en cantidad
    const grupos = new Map();
    const bolsa = (k) => {
      if (!grupos.has(k)) grupos.set(k, { rec: [], mov: [] });
      return grupos.get(k);
    };
    recibos.forEach((r, i) => bolsa(clave(r)).rec.push(i));
    movs.forEach((m, j) => bolsa(clave(m)).mov.push(j));

    for (const g of grupos.values()) {
      if (!g.rec.length || g.rec.length !== g.mov.length) continue;
      if (g.rec.some((i) => recibos[i].estado === "Conciliado")) continue; // ya resuelto por la regla 1
      // las videoconsultas solo pueden ir contra Mercado Pago: tiene que haber
      // al menos tantos movimientos de MP como recibos "solo MP"
      const soloMp = g.rec.filter((i) => cfg.colsSoloMP.includes(recibos[i].columna));
      const resto = g.rec.filter((i) => !cfg.colsSoloMP.includes(recibos[i].columna));
      const mp = g.mov.filter((j) => movs[j].fuente === "MP");
      const otros = g.mov.filter((j) => movs[j].fuente !== "MP");
      if (soloMp.length > mp.length) continue;
      // primero las videoconsultas contra MP; el resto, en orden de carga
      const izq = [...soloMp, ...resto];
      const der = [...mp.slice(0, soloMp.length), ...otros, ...mp.slice(soloMp.length)];
      const modo = `Grupo de ${izq.length}`;
      izq.forEach((i, p) => {
        const j = der[p];
        recibos[i].estado = movs[j].estado = "Conciliado";
        recibos[i].modo = movs[j].modo = modo;
        pares.push([recibos[i], movs[j], modo]);
      });
    }

    // regla 3: misma fecha y diferencia de hasta 99 centavos, única de los dos
    // lados, sobre lo que quedó sin conciliar por las reglas 1 y 2
    const porFecha = new Map();
    movs.forEach((m, j) => {
      if (m.estado === "Conciliado") return;
      if (!porFecha.has(m.fecha)) porFecha.set(m.fecha, []);
      porFecha.get(m.fecha).push(j);
    });
    const candTol = new Map();
    const invTol = new Map();
    recibos.forEach((r, i) => {
      if (r.estado === "Conciliado") return;
      const js = (porFecha.get(r.fecha) || []).filter((j) =>
        Math.abs(r.importe - movs[j].importe) <= TOLERANCIA &&
        (!cfg.colsSoloMP.includes(r.columna) || movs[j].fuente === "MP"));
      candTol.set(i, js);
      js.forEach((j) => {
        if (!invTol.has(j)) invTol.set(j, []);
        invTol.get(j).push(i);
      });
    });
    for (const [i, js] of candTol) {
      if (js.length !== 1 || invTol.get(js[0]).length !== 1) continue;
      const j = js[0];
      recibos[i].estado = movs[j].estado = "Conciliado";
      recibos[i].modo = movs[j].modo = MODO_TOL;
      pares.push([recibos[i], movs[j], MODO_TOL]);
    }
    return pares;
  }

  /**
   * Regla 4: entre los recibos que quedaron sin conciliar, cancela los pares
   * positivo/negativo de la misma fecha, el mismo cliente y el mismo importe.
   * Devuelve los recibos que salen del informe, de a pares [positivo, negativo].
   */
  function netear(recibos) {
    const grupos = new Map();
    recibos.forEach((r) => {
      if (r.estado === "Conciliado" || !r.importe) return;
      const k = `${r.fecha}|${r.cliente}|${Math.abs(r.importe)}`;
      if (!grupos.has(k)) grupos.set(k, { pos: [], neg: [] });
      grupos.get(k)[r.importe > 0 ? "pos" : "neg"].push(r);
    });
    const orden = (a, b) => Number(a.nroRecibo) - Number(b.nroRecibo);
    const fuera = [];
    for (const g of grupos.values()) {
      const n = Math.min(g.pos.length, g.neg.length);
      if (!n) continue;
      g.pos.sort(orden);
      g.neg.sort(orden);
      for (let i = 0; i < n; i++) fuera.push([g.pos[i], g.neg[i]]);
    }
    return fuera;
  }

  // ------------------------------------------------------------ proceso completo
  /**
   * archivos: [{nombre, bytes: Uint8Array}]
   * devuelve {resultado, avisos, libro (workbook para descargar), nombreSalida}
   */
  function procesar(archivos, cfgUsuario) {
    const cfg = Object.assign({}, CONFIG_DEFAULT, cfgUsuario || {});
    const avisos = [];
    const leidos = [];
    let recibos = [];
    let movs = [];

    for (const a of archivos) {
      const tipo = detectar(a.nombre, a.bytes);
      if (!tipo) {
        avisos.push(`"${a.nombre}": no se reconoce el formato, se ignora.`);
        continue;
      }
      leidos.push({ tipo, nombre: a.nombre });
      if (tipo === "recibos") {
        const r = leerRecibos(a.bytes, cfg);
        if (r.faltan.length) avisos.push(`"${a.nombre}": no tiene las columnas ${r.faltan.join(", ")}.`);
        recibos = recibos.concat(r.filas);
      } else if (tipo === "payway") movs = movs.concat(leerPayway(a.bytes));
      else if (tipo === "qr") movs = movs.concat(leerQr(a.bytes));
      else if (tipo === "mp") movs = movs.concat(leerMp(a.bytes));
    }

    for (const t of ["recibos", "payway", "qr", "mp"]) {
      if (!leidos.some((l) => l.tipo === t)) avisos.push(`Falta el archivo de ${nombreTipo(t)}.`);
    }
    if (!recibos.length) throw new Error("No hay listado de caja (recibos): no se puede conciliar.");

    // si se cargó dos veces el mismo archivo, no duplicar
    recibos = dedup(recibos, (r) => [r.nroRecibo, r.columna, r.importe, r.fecha].join("|"));
    movs = dedup(movs, (m) => [m.fuente, m.referencia, m.importe, m.fecha].join("|"));

    let pares = conciliar(recibos, movs, cfg);
    const anulados = netear(recibos);
    if (anulados.length) {
      const fuera = new Set(anulados.flat());
      recibos = recibos.filter((r) => !fuera.has(r));
      pares = conciliar(recibos, movs, cfg);
    }
    const fechas = recibos.map((r) => r.fecha).sort();
    const desde = fechas[0];
    const hasta = fechas[fechas.length - 1];

    const resultado = resumir(recibos, movs, cfg, desde, hasta, anulados);
    const libro = armarLibro(recibos, movs, pares, resultado, leidos, cfg, anulados);
    return { resultado, avisos, leidos, libro, nombreSalida: `Conciliacion_${desde}_${hasta}.xlsx` };
  }

  function nombreTipo(t) {
    return { recibos: "recibos (listado de caja)", payway: "movimientos Payway", qr: "QR Payway", mp: "Mercado Pago" }[t];
  }

  function dedup(arr, fk) {
    const vistos = new Set();
    return arr.filter((x) => {
      const k = fk(x);
      if (vistos.has(k)) return false;
      vistos.add(k);
      return true;
    });
  }

  function suma(arr) {
    return arr.reduce((s, x) => s + x.importe, 0);
  }

  function resumir(recibos, movs, cfg, desde, hasta, anulados) {
    const est = (arr, e) => arr.filter((x) => x.estado === e);
    const porModo = (f) => {
      const g = est(recibos, "Conciliado").filter(f);
      return { rec: g.length, importe: suma(g) };
    };
    return {
      desde, hasta,
      recibosPorColumna: [...cfg.colsTarjeta, ...cfg.colsSoloMP]
        .map((c) => ({ columna: c, cant: recibos.filter((r) => r.columna === c).length,
                       importe: suma(recibos.filter((r) => r.columna === c)) }))
        .filter((x) => x.cant),
      movsPorFuente: ["PAYWAY", "QR", "MP"]
        .map((f) => ({ fuente: f, cant: movs.filter((m) => m.fuente === f).length,
                       importe: suma(movs.filter((m) => m.fuente === f)) }))
        .filter((x) => x.cant),
      totRecibos: { cant: recibos.length, importe: suma(recibos) },
      totMovs: { cant: movs.length, importe: suma(movs) },
      conciliado: { rec: est(recibos, "Conciliado").length, mov: est(movs, "Conciliado").length,
                    importe: suma(est(recibos, "Conciliado")), importeMov: suma(est(movs, "Conciliado")) },
      porPar: porModo((x) => x.modo === MODO_PAR),
      porGrupo: Object.assign(porModo((x) => String(x.modo).startsWith("Grupo")), {
        grupos: new Set(est(recibos, "Conciliado").filter((r) => String(r.modo).startsWith("Grupo"))
          .map((r) => `${r.fecha}|${r.importe}`)).size }),
      porTolerancia: porModo((x) => x.modo === MODO_TOL),
      revisar: { rec: est(recibos, "Para revisar").length, mov: est(movs, "Para revisar").length,
                 importe: suma(est(recibos, "Para revisar")), importeMov: suma(est(movs, "Para revisar")) },
      recSinMov: { cant: est(recibos, "Sin movimiento").length, importe: suma(est(recibos, "Sin movimiento")) },
      movSinRec: { cant: est(movs, "Sin recibo").length, importe: suma(est(movs, "Sin recibo")) },
      anulados: { rec: (anulados || []).length * 2, pares: (anulados || []).length,
                  importe: (anulados || []).reduce((s, p) => s + p[0].importe, 0) },
    };
  }

  // ------------------------------------------------------------ Excel de salida
  const ESTILO_CAB = { font: { bold: true, color: { rgb: "FFFFFF" } }, fill: { fgColor: { rgb: "1F4E78" } },
                       alignment: { vertical: "center", wrapText: true } };
  const FMT_IMPORTE = "#,##0.00";

  function fechaExcel(iso) {
    // número de serie de Excel (sin hora ni zona horaria)
    const [y, m, d] = iso.split("-").map(Number);
    return (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000;
  }

  function hoja(cabecera, filas, opciones) {
    const op = opciones || {};
    const datos = [cabecera, ...filas];
    const ws = XLSX.utils.aoa_to_sheet(datos);
    const rango = XLSX.utils.decode_range(ws["!ref"]);
    for (let c = rango.s.c; c <= rango.e.c; c++) {
      const cab = ws[XLSX.utils.encode_cell({ r: 0, c })];
      if (cab && !op.sinCabecera) cab.s = ESTILO_CAB;
      for (let r = 1; r <= rango.e.r; r++) {
        const cel = ws[XLSX.utils.encode_cell({ r, c })];
        if (!cel) continue;
        if (op.fechas && op.fechas.includes(c) && cel.t === "n") cel.z = "dd/mm/yyyy";
        else if (cel.t === "n" && op.importes && op.importes.includes(c)) cel.z = FMT_IMPORTE;
        if (op.relleno && op.relleno(r)) cel.s = Object.assign({}, cel.s, { fill: { fgColor: { rgb: "EDEDED" } } });
        if (op.negrita && op.negrita(r)) cel.s = Object.assign({}, cel.s, { font: { bold: true } });
      }
    }
    ws["!cols"] = cabecera.map((_, c) => {
      const largo = Math.max(...datos.slice(0, 300).map((f) => String(op.fechas && op.fechas.includes(c) ? "00/00/0000" : f[c] ?? "").length));
      return { wch: Math.min(Math.max(largo, 8) + 2, 60) };
    });
    if (!op.sinCabecera) {
      ws["!autofilter"] = { ref: ws["!ref"] };
      ws["!freeze"] = { xSplit: 0, ySplit: 1 };
      ws["!views"] = [{ state: "frozen", ySplit: 1 }];
    }
    return ws;
  }

  function armarLibro(recibos, movs, pares, R, leidos, cfg, anulados) {
    const wb = XLSX.utils.book_new();
    const f = (x) => fechaExcel(x);

    // Resumen
    const res = [
      ["Período", `${R.desde.split("-").reverse().join("/")} al ${R.hasta.split("-").reverse().join("/")}`],
      ["Procesado", new Date().toLocaleString("es-AR")],
      ["Regla 1", "Par directo: misma fecha + mismo importe exacto, único de ambos lados"],
      ["Regla 2", "Grupo que cierra: misma fecha + mismo importe, igual cantidad de recibos que de movimientos"],
      ["Regla 3", "Diferencia de centavos: misma fecha + diferencia de hasta $0,99, única de ambos lados"],
      ["Regla 4", "Neteo de anulaciones: recibo positivo y negativo sin conciliar, misma fecha, mismo cliente e importe"],
      ["", `En las tres primeras, ${cfg.colsSoloMP.join(", ")} solo se cruza contra Mercado Pago`],
      [],
      ["RECIBOS", "Líneas", "Importe"],
      ...R.recibosPorColumna.map((x) => [x.columna, x.cant, pesos(x.importe)]),
      ["Total", R.totRecibos.cant, pesos(R.totRecibos.importe)],
      [],
      ["MOVIMIENTOS", "Cantidad", "Importe"],
      ...R.movsPorFuente.map((x) => [x.fuente, x.cant, pesos(x.importe)]),
      ["Total", R.totMovs.cant, pesos(R.totMovs.importe)],
      [],
      ["RESULTADO", "Recibos", "Importe recibos", "Movimientos"],
      ["Conciliado", R.conciliado.rec, pesos(R.conciliado.importe), R.conciliado.mov],
      ["   por par directo (regla 1)", R.porPar.rec, pesos(R.porPar.importe), R.porPar.rec],
      [`   por grupo que cierra (regla 2, ${R.porGrupo.grupos} grupos)`, R.porGrupo.rec, pesos(R.porGrupo.importe), R.porGrupo.rec],
      ["   por diferencia de hasta $0,99 (regla 3)", R.porTolerancia.rec, pesos(R.porTolerancia.importe), R.porTolerancia.rec],
      ["Para revisar", R.revisar.rec, pesos(R.revisar.importe), R.revisar.mov],
      ["Recibos sin movimiento", R.recSinMov.cant, pesos(R.recSinMov.importe), ""],
      ["Movimientos sin recibo", "", pesos(R.movSinRec.importe), R.movSinRec.cant],
      ["Anulados y neteados (regla 4)", R.anulados.rec, 0, ""],
      [],
      ["ARCHIVOS LEÍDOS"],
      ...leidos.map((l) => [nombreTipo(l.tipo), l.nombre]),
    ];
    const anchoRes = 4;
    const resPad = res.map((r) => [...r, ...Array(anchoRes - r.length).fill("")]);
    XLSX.utils.book_append_sheet(wb, hoja(resPad[0], resPad.slice(1), {
      sinCabecera: true, importes: [2],
      negrita: (r) => /^[A-ZÁÉÍÓÚ ]+$/.test(String(resPad[r][0])) && String(resPad[r][0]).trim() !== "",
    }), "Resumen");

    // Conciliados
    const conc = pares
      .slice()
      .sort((a, b) => (a[0].fecha + a[0].nroRecibo).localeCompare(b[0].fecha + b[0].nroRecibo))
      .map(([r, m, modo]) => [f(r.fecha), pesos(r.importe), modo, Number(r.nroRecibo), r.cliente, r.cajero, r.columna,
                              r.tipoAsiento, m.fuente, pesos(m.importe), pesos(m.importe - r.importe),
                              m.referencia, m.detalle, m.terminal]);
    XLSX.utils.book_append_sheet(wb, hoja(
      ["Fecha", "Importe", "Apareo", "Nro Recibo", "Cliente", "Cajero", "Columna", "Tipo Asiento", "Fuente",
       "Importe mov", "Diferencia", "Referencia mov", "Detalle mov", "Terminal / Caja"], conc,
      { importes: [1, 9, 10], fechas: [0] }), "Conciliados");

    // Anulados por la regla 4
    const anu = (anulados || [])
      .slice()
      .sort((a, b) => (a[0].fecha + a[0].nroRecibo).localeCompare(b[0].fecha + b[0].nroRecibo))
      .map(([p, n], i) => [i + 1, f(p.fecha), p.cliente, pesos(p.importe),
                           Number(p.nroRecibo), p.cajero, p.columna, p.tipoAsiento,
                           Number(n.nroRecibo), n.cajero, n.columna, n.tipoAsiento, pesos(n.importe)]);
    XLSX.utils.book_append_sheet(wb, hoja(
      ["Par", "Fecha", "Cliente", "Importe", "Recibo", "Cajero", "Columna", "Tipo Asiento",
       "Recibo anulación", "Cajero", "Columna", "Tipo Asiento", "Importe anulación"], anu,
      { importes: [3, 12], fechas: [1], relleno: (r) => r % 2 === 0 }), "Anulados");

    // Para revisar: agrupado por fecha + importe
    const rr = recibos.filter((x) => x.estado === "Para revisar");
    const mr = movs.filter((x) => x.estado === "Para revisar");
    const claves = [...new Set([...rr, ...mr].map((x) => `${x.fecha}|${x.importe}`))].sort((a, b) => {
      const [fa, ia] = a.split("|"), [fb, ib] = b.split("|");
      return fa === fb ? Number(ia) - Number(ib) : fa.localeCompare(fb);
    });
    const rev = [];
    const grupoDeFila = [null];
    claves.forEach((k, n) => {
      const [fe, imp] = k.split("|");
      const gr = rr.filter((x) => x.fecha === fe && String(x.importe) === imp);
      const gm = mr.filter((x) => x.fecha === fe && String(x.importe) === imp);
      const base = [n + 1, f(fe), pesos(Number(imp)), gr.length, gm.length];
      gr.forEach((x) => { rev.push([...base, "Recibo", Number(x.nroRecibo), x.cliente, x.cajero, x.columna, x.tipoAsiento, "", "", "", ""]); grupoDeFila.push(n + 1); });
      gm.forEach((x) => { rev.push([...base, "Movimiento", "", "", "", "", "", x.fuente, x.referencia, x.detalle, x.terminal]); grupoDeFila.push(n + 1); });
    });
    XLSX.utils.book_append_sheet(wb, hoja(
      ["Grupo", "Fecha", "Importe", "Cant. recibos", "Cant. movimientos", "Lado", "Nro Recibo", "Cliente",
       "Cajero", "Columna", "Tipo Asiento", "Fuente", "Referencia", "Detalle", "Terminal / Caja"],
      rev, { importes: [2], fechas: [1], relleno: (r) => grupoDeFila[r] % 2 === 0 }), "Para revisar");

    // Pendientes
    const rs = recibos.filter((x) => x.estado === "Sin movimiento")
      .sort((a, b) => (a.fecha + a.nroRecibo).localeCompare(b.fecha + b.nroRecibo))
      .map((x) => [f(x.fecha), Number(x.nroRecibo), x.cliente, x.cajero, x.columna, x.tipoAsiento, pesos(x.importe)]);
    XLSX.utils.book_append_sheet(wb, hoja(
      ["Fecha", "Nro Recibo", "Cliente", "Cajero", "Columna", "Tipo Asiento", "Importe"], rs, { importes: [6], fechas: [0] }),
      "Recibos sin movimiento");

    const ms = movs.filter((x) => x.estado === "Sin recibo")
      .sort((a, b) => (a.fuente + a.fecha).localeCompare(b.fuente + b.fecha))
      .map((x) => [x.fuente, f(x.fecha), pesos(x.importe), x.referencia, x.detalle, x.terminal]);
    XLSX.utils.book_append_sheet(wb, hoja(
      ["Fuente", "Fecha", "Importe", "Referencia", "Detalle", "Terminal / Caja"], ms, { importes: [2], fechas: [1] }),
      "Movimientos sin recibo");

    return wb;
  }

  const api = { procesar, detectar, CONFIG_DEFAULT, pesos };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.Conciliador = api;
})(typeof window !== "undefined" ? window : globalThis);
