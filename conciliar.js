/*
 * Conciliación de cobros con tarjeta / QR / Mercado Pago - Grupo Oroño
 *
 * Regla de conciliación (única, sin desempates):
 *   Un recibo se concilia con un movimiento solo si tienen la MISMA FECHA y el
 *   MISMO IMPORTE exacto, y esa combinación identifica a un único recibo y a un
 *   único movimiento. Las columnas "solo MP" (videoconsultas) solo se cruzan
 *   contra Mercado Pago. Todo lo demás queda para revisión manual.
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
    });
    movs.forEach((m, j) => {
      m.estado = candMov[j].length ? "Para revisar" : "Sin recibo";
    });
    candRec.forEach((js, i) => {
      if (js.length === 1 && candMov[js[0]].length === 1) {
        recibos[i].estado = "Conciliado";
        movs[js[0]].estado = "Conciliado";
        pares.push([recibos[i], movs[js[0]]]);
      }
    });
    return pares;
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

    const pares = conciliar(recibos, movs, cfg);
    const fechas = recibos.map((r) => r.fecha).sort();
    const desde = fechas[0];
    const hasta = fechas[fechas.length - 1];

    const resultado = resumir(recibos, movs, cfg, desde, hasta);
    const libro = armarLibro(recibos, movs, pares, resultado, leidos, cfg);
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

  function resumir(recibos, movs, cfg, desde, hasta) {
    const est = (arr, e) => arr.filter((x) => x.estado === e);
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
                    importe: suma(est(recibos, "Conciliado")) },
      revisar: { rec: est(recibos, "Para revisar").length, mov: est(movs, "Para revisar").length,
                 importe: suma(est(recibos, "Para revisar")), importeMov: suma(est(movs, "Para revisar")) },
      recSinMov: { cant: est(recibos, "Sin movimiento").length, importe: suma(est(recibos, "Sin movimiento")) },
      movSinRec: { cant: est(movs, "Sin recibo").length, importe: suma(est(movs, "Sin recibo")) },
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

  function armarLibro(recibos, movs, pares, R, leidos, cfg) {
    const wb = XLSX.utils.book_new();
    const f = (x) => fechaExcel(x);

    // Resumen
    const res = [
      ["Período", `${R.desde.split("-").reverse().join("/")} al ${R.hasta.split("-").reverse().join("/")}`],
      ["Procesado", new Date().toLocaleString("es-AR")],
      ["Regla", `Misma fecha + mismo importe exacto, único de ambos lados (${cfg.colsSoloMP.join(", ")} solo contra MP)`],
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
      ["Para revisar", R.revisar.rec, pesos(R.revisar.importe), R.revisar.mov],
      ["Recibos sin movimiento", R.recSinMov.cant, pesos(R.recSinMov.importe), ""],
      ["Movimientos sin recibo", "", pesos(R.movSinRec.importe), R.movSinRec.cant],
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
      .map(([r, m]) => [f(r.fecha), pesos(r.importe), Number(r.nroRecibo), r.cliente, r.cajero, r.columna,
                        r.tipoAsiento, m.fuente, m.referencia, m.detalle, m.terminal]);
    XLSX.utils.book_append_sheet(wb, hoja(
      ["Fecha", "Importe", "Nro Recibo", "Cliente", "Cajero", "Columna", "Tipo Asiento", "Fuente",
       "Referencia mov", "Detalle mov", "Terminal / Caja"], conc, { importes: [1], fechas: [0] }), "Conciliados");

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
