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
 *   5. Pagador propio: los movimientos cuyo campo "Pagador:" es Grupo Oroño los
 *      hace la empresa, no salen de un recibo de caja. Se apartan ANTES de
 *      conciliar y se informan aparte. Solo se mira el pagador: "Producto de
 *      Grupooroño" aparece en cobros de pacientes reales y NO los excluye.
 *   6. Videoconsulta: un movimiento cuyo detalle dice "Videoconsulta médica" solo
 *      puede conciliar contra un recibo de una columna "solo MP" (TVIR). Nunca
 *      contra un recibo de mostrador, aunque la fecha y el importe coincidan.
 *   7. Rendimientos de Mercado Pago: los movimientos de MP sin MEDIO DE PAGO son
 *      el rendimiento diario de la cuenta, no un cobro: no tienen local, ni caja,
 *      ni pagador, ni comisión. Se apartan ANTES de conciliar y se informan
 *      aparte. Se mira el medio de pago, no el número de identificación: ese
 *      también viene vacío en cobros reales de pacientes.
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

  const PAGADOR_PROPIO = "orono"; // regla 5: sin acentos ni mayúsculas

  const MODO_PAR = "Par directo";
  const MODO_TOL = "Diferencia hasta $0,99";
  const TOLERANCIA = 99; // centavos

  // ------------------------------------------------------------ utilidades
  const sinAcentos = (s) => String(s == null ? "" : s).normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase();

  /** regla 6: ¿el movimiento es una videoconsulta? Solo puede ir contra un recibo virtual. */
  function esVideoconsulta(m) {
    return sinAcentos(m.detalle).includes("videoconsulta");
  }

  /** ¿este recibo y este movimiento pueden llegar a ser el mismo cobro? */
  function compatible(r, m, cfg) {
    const virtual = cfg.colsSoloMP.includes(r.columna);
    if (virtual && m.fuente !== "MP") return false;       // las videoconsultas se cobran solo por MP
    if (!virtual && esVideoconsulta(m)) return false;     // y un cobro de videoconsulta no es de mostrador
    return true;
  }

  /** regla 7: ¿es el rendimiento diario de la cuenta de MP? No tiene medio de pago. */
  function esRendimiento(m) {
    return m.fuente === "MP" && m.medio === "";
  }

  /** regla 5: ¿el movimiento lo pagó la propia empresa? Solo mira el campo "Pagador:". */
  function pagadorPropio(m) {
    const d = sinAcentos(m.detalle);
    const i = d.indexOf("pagador:");
    return i >= 0 && d.slice(i + "pagador:".length).includes(PAGADOR_PROPIO);
  }

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
      medio: g(f, "MEDIO DE PAGO"),
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
      (porClave.get(clave(r)) || []).filter((j) => compatible(r, movs[j], cfg)));
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
      const video = g.mov.filter((j) => esVideoconsulta(movs[j]) && movs[j].fuente === "MP");
      const mp = g.mov.filter((j) => movs[j].fuente === "MP" && !esVideoconsulta(movs[j]));
      const otros = g.mov.filter((j) => movs[j].fuente !== "MP" && !esVideoconsulta(movs[j]));
      // una videoconsulta que no sea de MP no puede aparear con nada: el grupo no cierra
      if (g.mov.length !== video.length + mp.length + otros.length) continue;
      if (video.length > soloMp.length) continue;         // regla 6: solo entran en recibos virtuales
      if (soloMp.length > video.length + mp.length) continue;
      // primero las videoconsultas, después el resto de MP; lo demás, en orden de carga
      const huecos = soloMp.length - video.length;
      const izq = [...soloMp, ...resto];
      const der = [...video, ...mp.slice(0, huecos), ...otros, ...mp.slice(huecos)];
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
        Math.abs(r.importe - movs[j].importe) <= TOLERANCIA && compatible(r, movs[j], cfg));
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

    // regla 5: los movimientos pagados por la propia empresa no se concilian
    const propios = movs.filter(pagadorPropio);
    movs = movs.filter((m) => !pagadorPropio(m));

    // regla 7: el rendimiento de la cuenta de MP no sale de un cobro, no tiene recibo
    const rendimientos = movs.filter(esRendimiento);
    movs = movs.filter((m) => !esRendimiento(m));

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

    const resultado = resumir(recibos, movs, cfg, desde, hasta, anulados, propios, rendimientos);
    const libro = armarLibro(recibos, movs, pares, resultado, leidos, cfg, anulados, propios, rendimientos);
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

  function resumir(recibos, movs, cfg, desde, hasta, anulados, propios, rendimientos) {
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
      pagadorPropio: { cant: (propios || []).length, importe: suma(propios || []) },
      rendimientos: { cant: (rendimientos || []).length, importe: suma(rendimientos || []) },
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
        if (op.rojo && op.rojo(r)) cel.s = Object.assign({}, cel.s, { font: { bold: true, color: { rgb: "9C0006" } } });
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

  function armarLibro(recibos, movs, pares, R, leidos, cfg, anulados, propios, rendimientos) {
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
      ["Regla 5", "Pagador Grupo Oroño: los movimientos que paga la empresa se apartan antes de conciliar"],
      ["Regla 6", "Videoconsulta: un movimiento que dice \"Videoconsulta médica\" solo concilia contra un recibo " + cfg.colsSoloMP.join(", ")],
      ["Regla 7", "Rendimientos: los movimientos de Mercado Pago sin medio de pago se apartan antes de conciliar"],
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
      ["Pagador Grupo Oroño (estos no tienen recibo)", "", pesos(R.pagadorPropio.importe), R.pagadorPropio.cant],
      ["Rendimientos de Mercado Pago (estos no tienen recibo)", "", pesos(R.rendimientos.importe), R.rendimientos.cant],
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

    // Regla 5: movimientos pagados por la propia empresa
    if ((propios || []).length) {
      const pr = propios
        .slice()
        .sort((a, b) => (a.fuente + a.fecha).localeCompare(b.fuente + b.fecha))
        .map((x) => [x.fuente, f(x.fecha), pesos(x.importe), x.referencia, x.detalle, x.terminal]);
      XLSX.utils.book_append_sheet(wb, hoja(
        ["Fuente", "Fecha", "Importe", "Referencia", "Detalle", "Terminal / Caja"], pr,
        { importes: [2], fechas: [1] }), "Pagador Grupo Oroño");
    }

    // Regla 7: rendimientos de la cuenta de Mercado Pago
    if ((rendimientos || []).length) {
      const rd = rendimientos
        .slice()
        .sort((a, b) => a.fecha.localeCompare(b.fecha))
        .map((x) => [f(x.fecha), pesos(x.importe), x.referencia]);
      XLSX.utils.book_append_sheet(wb, hoja(
        ["Fecha", "Importe", "Referencia"], rd, { importes: [1], fechas: [0] }), "Rendimientos MP");
    }

    // agrupa recibos y movimientos por fecha + importe, con las columnas de "Para revisar"
    function agrupado(rr, mr, conEstado) {
      const claves = [...new Set([...rr, ...mr].map((x) => `${x.fecha}|${x.importe}`))].sort((a, b) => {
        const [fa, ia] = a.split("|"), [fb, ib] = b.split("|");
        return fa === fb ? Number(ia) - Number(ib) : fa.localeCompare(fb);
      });
      const filas = [];
      const grupoDeFila = [null];
      // en la hoja de trabajo la primera columna dice de qué lista viene cada fila
      const est = (x, sola) => (conEstado ? [x.estado === "Para revisar" ? "Para revisar" : sola] : []);
      claves.forEach((k, n) => {
        const [fe, imp] = k.split("|");
        const gr = rr.filter((x) => x.fecha === fe && String(x.importe) === imp);
        const gm = mr.filter((x) => x.fecha === fe && String(x.importe) === imp);
        // el importe va en la columna del lado que corresponde, para poder sumar cada uno por separado
        const base = (lado) => [n + 1, f(fe), lado === "Recibo" ? pesos(Number(imp)) : "",
                                lado === "Movimiento" ? pesos(Number(imp)) : "", gr.length, gm.length];
        gr.forEach((x) => { filas.push([...est(x, "Recibo sin movimiento"), ...base("Recibo"), "Recibo", Number(x.nroRecibo), x.cliente, x.cajero, x.columna, x.tipoAsiento, "", "", "", ""]); grupoDeFila.push(n + 1); });
        gm.forEach((x) => { filas.push([...est(x, "Movimiento sin recibo"), ...base("Movimiento"), "Movimiento", "", "", "", "", "", x.fuente, x.referencia, x.detalle, x.terminal]); grupoDeFila.push(n + 1); });
      });
      return { filas, grupoDeFila };
    }

    const COLS_REV = ["Grupo", "Fecha", "Importe recibo", "Importe movimiento", "Cant. recibos",
                      "Cant. movimientos", "Lado", "Nro Recibo", "Cliente", "Cajero", "Columna",
                      "Tipo Asiento", "Fuente", "Referencia", "Detalle", "Terminal / Caja"];

    // Hoja de trabajo: todo lo pendiente junto (para revisar + lo que no tiene contrapartida)
    const tra = agrupado(recibos.filter((x) => x.estado !== "Conciliado"),
                         movs.filter((x) => x.estado !== "Conciliado"), true);
    XLSX.utils.book_append_sheet(wb, hoja(["Estado", ...COLS_REV], tra.filas,
      { importes: [3, 4], fechas: [2], relleno: (r) => tra.grupoDeFila[r] % 2 === 0 }), "Hoja de trabajo");

    // Para revisar: agrupado por fecha + importe
    const rev = agrupado(recibos.filter((x) => x.estado === "Para revisar"),
                         movs.filter((x) => x.estado === "Para revisar"), false);
    XLSX.utils.book_append_sheet(wb, hoja(COLS_REV, rev.filas,
      { importes: [2, 3], fechas: [1], relleno: (r) => rev.grupoDeFila[r] % 2 === 0 }), "Para revisar");

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

  const api = { procesar, detectar, CONFIG_DEFAULT, pesos, hoja, fechaExcel };
  root.Conciliador = api;
  if (typeof module !== "undefined" && module.exports) module.exports = api;
})(typeof window !== "undefined" ? window : globalThis);

/*
 * Contabilización en bancos - Grupo Oroño
 *
 * Paso 1: tres cruces, cada uno por una clave única verificada, sin desempates
 *         ni tolerancias.
 *   A. Liquidación Payway de tarjetas  <->  extracto Supervielle
 *      clave: número de establecimiento + fecha de pago.
 *      Es N:1 (el banco parte el día por marca, el PDF da un solo total del día).
 *   B. Transferencias QR Payway        <->  extracto Supervielle
 *      clave: fecha de venta + monto neto, apareo uno a uno por multiconjunto.
 *   C. Liberaciones Mercado Pago       <->  resumen de cuenta Mercado Pago
 *      clave: SOURCE_ID = REFERENCE_ID.
 *
 * Paso 2: asiento mensual consolidado, armado SOLO con lo que cruzó.
 *   Lo que no cruza no se fuerza: queda en la hoja "Pendientes" y se arrastra al
 *   mes siguiente (el Excel de un mes es archivo de entrada del mes que viene).
 *   Un pendiente se resuelve cuando reaparece POR SU MISMA CLAVE ÚNICA, nunca
 *   por importe parecido.
 *
 * Decisiones del usuario que este módulo aplica tal cual (no son criterio propio):
 *   - Cargo por Servicio Payway bonificado al 100%: no va al asiento.
 *   - Impuestos propios del banco (Imp. Déb. y Créd. e IIBB Acreditaciones): no
 *     van, se contabilizan en un asiento mensual aparte. Sí se muestran en el control.
 *   - Transferencias salientes de Mercado Pago (payout): fuera de esta sección.
 *   - Base gravada y exenta van juntas en "Gastos bancarios".
 *   - Rendimientos de Mercado Pago: cuenta "Rendimiento Inversiones".
 *
 * Funciona en el navegador (usa window.XLSX y un lector de PDF inyectado) y en
 * Node para las pruebas.
 */
(function (root) {
  "use strict";

  const XLSX = root.XLSX || (typeof require !== "undefined" ? require("xlsx-js-style") : null);

  // nombres de cuenta tal como los dio el usuario
  const CUENTAS = {
    banco: "Banco Supervielle acreditación tarjeta cuenta corriente",
    mp: "Cobros Mercado Pago",
    gastos: "Gastos bancarios",
    ivaCf: "IVA CF Bancos",
    retIibb: "ING BRUTOS retenciones",
    retAfip: "IVA retenciones",
    impDebCred: "Impuesto al Debito y Credito",
    tarjetas: "Tarjetas",
    naranja: "Tarjeta Naranja a Conciliar",
    rendimiento: "Rendimiento Inversiones",
    sinCuenta: "A DEFINIR — retención o percepción sin cuenta asignada",
  };

  // conceptos del extracto Supervielle, tal cual vienen escritos
  const CPT_PRISMA = "Comercios Prisma";
  const CPT_QR = "COBRO CON QR";
  // el debito con que el banco deshace un cobro con QR: la transferencia se dio vuelta
  const CPT_DEV = "DEVOLUCION PEI";
  // Tarjeta Naranja liquida por transferencia interbancaria, fuera de Payway. No hay
  // liquidación que respalde el crédito, así que no se puede abrir en bruto, arancel e
  // impuestos: entra entero a una cuenta a conciliar hasta que aparezca el documento.
  // La clave única es el CUIT en el detalle, no el texto del concepto: por "CRED BCA
  // ELECTRONICA INTERBANC" puede entrar plata de cualquiera.
  const CUIT_NARANJA = "30685376349";
  const CPT_IMP_CR = "Impuesto Débitos y Créditos/CR";
  const CPT_IIBB_AC = "IIBB- Acreditaciones Bancarias";

  const MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio",
                 "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"];

  // ------------------------------------------------------------ utilidades
  const cents = (v) => Math.round(num(v) * 100);
  const pesos = (c) => c / 100;

  /** cualquier celda -> número en pesos. Acepta 1234.56, "1.234,56" y "1234,56". */
  function num(v) {
    if (v === null || v === undefined || v === "") return 0;
    if (typeof v === "number") return Number.isFinite(v) && Math.abs(v) > 1e-9 ? v : 0;
    const s = String(v).trim().replace(/^'|'$/g, "");
    if (s === "") return 0;
    // formato argentino: el punto separa miles y la coma los decimales
    const ar = /^-?\d{1,3}(\.\d{3})*(,\d+)?$/.test(s) || /^-?\d+,\d+$/.test(s);
    const n = Number(ar ? s.replace(/\./g, "").replace(",", ".") : s);
    return Number.isFinite(n) && Math.abs(n) > 1e-9 ? n : 0;
  }

  function serialAIso(serial) {
    const d = new Date(Date.UTC(1899, 11, 30) + Math.round(serial) * 86400000);
    return d.toISOString().slice(0, 10);
  }

  /** celda de fecha -> "AAAA-MM-DD". Soporta serial de Excel y los textos de cada archivo. */
  function iso(v) {
    if (v === null || v === undefined || v === "") return "";
    if (typeof v === "number") return serialAIso(v);
    const s = String(v).trim();
    let m = s.match(/^(\d{4})-(\d\d)-(\d\d)/);            // 2026-08-31 / ISO largo
    if (m) return m[1] + "-" + m[2] + "-" + m[3];
    m = s.match(/^(\d\d)[/-](\d\d)[/-](\d{4})/);          // 31/08/2026 y 01-08-2026
    if (m) return m[3] + "-" + m[2] + "-" + m[1];
    return "";
  }

  const ddmm = (i) => (i ? i.slice(8, 10) + "/" + i.slice(5, 7) : "");
  const periodoDe = (i) => (i ? i.slice(0, 7) : "");
  function nombrePeriodo(p) {
    if (!p) return "";
    return MESES[Number(p.slice(5, 7)) - 1] + " " + p.slice(0, 4);
  }

  function leerLibro(bytes) { return XLSX.read(bytes, { type: "array" }); }

  function filas(wb, nombreHoja) {
    const h = wb.Sheets[nombreHoja || wb.SheetNames[0]];
    if (!h) return [];
    return XLSX.utils.sheet_to_json(h, { header: 1, raw: true, defval: "" });
  }

  /** convierte una matriz en objetos usando la fila `iCab` como encabezado */
  function objetos(mat, iCab) {
    const cab = (mat[iCab] || []).map((c) => String(c).trim());
    return mat.slice(iCab + 1)
      .filter((f) => f.some((c) => c !== "" && c !== null && c !== undefined))
      .map((f) => Object.fromEntries(cab.map((c, i) => [c, f[i]])));
  }

  // ------------------------------------------------------- lector de PDF
  // El navegador inyecta pdf.js; Node inyecta pdfjs-dist. El módulo no sabe cuál es.
  let leerPdf = null;
  function usarLectorPdf(fn) { leerPdf = fn; }

  // ------------------------------------------------------------ detección
  function detectar(nombre, bytes) {
    const ext = String(nombre).toLowerCase().split(".").pop();
    if (ext === "zip" || ext === "pdf") return "liqPayway";
    if (ext !== "xlsx" && ext !== "xls") return null;
    let wb;
    try { wb = leerLibro(bytes); } catch (e) { return null; }
    if (wb.SheetNames.indexOf("Pendientes") >= 0 && wb.SheetNames.indexOf("Asiento") >= 0) return "previo";
    const mat = filas(wb).slice(0, 6);
    const txt = mat.map((f) => f.map((c) => String(c).trim()).join("|")).join("\n");
    if (txt.indexOf("RECORD_TYPE") >= 0 && txt.indexOf("SOURCE_ID") >= 0) return "libMp";
    if (txt.indexOf("INITIAL_BALANCE") >= 0 && txt.indexOf("REFERENCE_ID") >= 0) return "resumenMp";
    if (txt.indexOf("QR ID") >= 0 && txt.indexOf("MONTO NETO") >= 0) return "qr";
    if (txt.indexOf("Concepto") >= 0 && txt.indexOf("Débito") >= 0 && txt.indexOf("Crédito") >= 0) return "extracto";
    return null;
  }

  // ------------------------------------------------- lectura: PDFs Payway
  /** clasifica una línea del DESGLOSE DE DESCUENTOS en su concepto contable */
  function clasificarDescuento(etiqueta) {
    const e = etiqueta.replace(/\s*\d+[,.]\d+\s*%/, "").trim();
    if (/^Arancel/i.test(e)) return "arancel";
    if (/^\d+\s+Ventas?\s/i.test(etiqueta)) return "financiero";  // "34 Ventas en 9 días"
    if (/^IVA$/i.test(e)) return "iva";
    if (/^Ret\.IB/i.test(e)) return "retIibb";
    if (/^Percep/i.test(e)) return "retAfip";
    return null;   // Tasa / Base Exenta (bases del IVA) y Cargo por Servicio con su bonificación
  }

  /**
   * Totales de la cabecera del resumen. Vienen en dos columnas (pesos y dólares)
   * y a veces comparten renglón con el domicilio, así que se busca el primer
   * renglón posterior al rótulo que TERMINE en dos importes.
   */
  function totalCabecera(lineas, rotulo) {
    const i = lineas.findIndex((l) => l.indexOf(rotulo + " $") >= 0);
    if (i < 0) return null;
    for (let j = i; j < Math.min(i + 8, lineas.length); j++) {
      const m = lineas[j].match(/(-?[\d.]+,\d\d)\s+(-?[\d.]+,\d\d)\s*$/);
      if (m) return { pesos: cents(m[1]), dolares: cents(m[2]) };
    }
    return null;
  }

  function leerLiquidacionPayway(texto, archivo) {
    const mEst = texto.match(/N[º°.]?\s*DE ESTABLECIMIENTO:\s*(\d+)/i);
    if (!mEst) throw new Error('"' + archivo + '": no se encontró el número de establecimiento');
    const est = mEst[1].replace(/^0+/, "");
    const lineas = texto.split("\n").map((l) => l.trim());
    const mPag = texto.match(/PAGADOR:\s*\d+\s+(.+)/);
    const cab = {
      pagador: mPag ? mPag[1].trim() : "",
      presentado: totalCabecera(lineas, "TOTAL PRESENTADO"),
      descuento: totalCabecera(lineas, "TOTAL DESCUENTO"),
      saldo: totalCabecera(lineas, "SALDO"),
    };
    const dias = [];
    let fecha = null;
    for (const l of lineas) {
      const mf = l.match(/^FECHA DE PAGO\s+(\d\d\/\d\d)/i);
      if (mf) { fecha = mf[1]; continue; }
      const mt = l.match(/^Total del d[ií]a\s+\$\s*([\d.]+,\d\d)\s+\$\s*(-?[\d.]+,\d\d)\s+\$\s*(-?[\d.]+,\d\d)/i);
      if (mt && fecha) {
        dias.push({ est: est, archivo: archivo, fecha: fecha,
                    presentado: cents(mt[1]), descuento: cents(mt[2]), neto: cents(mt[3]) });
      }
    }
    const desc = { arancel: 0, financiero: 0, iva: 0, retIibb: 0, retAfip: 0 };
    const detalle = [];
    if (texto.indexOf("DESGLOSE DE DESCUENTOS") >= 0) {
      const seg = texto.split("DESGLOSE DE DESCUENTOS")[1].split("SR. COMERCIANTE")[0];
      for (const linea of seg.split("\n")) {
        const m = linea.trim().match(/^(.+?)\s+\$\s*(-?[\d.]+,\d\d)$/);
        if (!m) continue;
        const cls = clasificarDescuento(m[1].trim());
        detalle.push({ est: est, archivo: archivo, etiqueta: m[1].trim(),
                       importe: cents(m[2]), concepto: cls || "(fuera del asiento)" });
        if (cls) desc[cls] += cents(m[2]);
      }
    }
    return { est: est, archivo: archivo, cab: cab, dias: dias, desc: desc, detalle: detalle };
  }

  // --------------------------------------------------- lectura: extracto
  function leerExtracto(bytes) {
    const mat = filas(leerLibro(bytes));
    const iCab = mat.findIndex((f) => String(f[0]).trim() === "Fecha" && String(f[2]).trim() === "Concepto");
    if (iCab < 0) throw new Error("El extracto del banco no tiene el encabezado Fecha / Hora / Concepto");
    return objetos(mat, iCab)
      .filter((r) => iso(r["Fecha"]))
      .map((r, i) => ({
        n: i,
        fecha: iso(r["Fecha"]),
        hora: String(r["Hora"] == null ? "" : r["Hora"]).trim(),
        concepto: String(r["Concepto"] == null ? "" : r["Concepto"]).trim(),
        detalle: String(r["Detalle"] == null ? "" : r["Detalle"]).trim(),
        importe: cents(r["Crédito"]) - cents(r["Débito"]),
      }));
  }

  // --------------------------------------------------------- lectura: QR
  const COLS_R_IIBB = /^R_INGRESOS BRUTOS /i;
  const COLS_R_OTRA = /^R_(IVA|IMPUESTO)/i;
  const COLS_P = /^P_/i;

  function leerQr(bytes) {
    const mat = filas(leerLibro(bytes));
    const iCab = mat.findIndex((f) => f.some((c) => String(c).trim() === "QR ID"));
    if (iCab < 0) throw new Error("El archivo de QR no tiene la fila de encabezados con 'QR ID'");
    return objetos(mat, iCab).filter((r) => String(r["QR ID"] == null ? "" : r["QR ID"]).trim() !== "").map((r, i) => {
      let otras = 0, perc = 0;
      const jur = [];
      for (const k of Object.keys(r)) {
        if (COLS_R_OTRA.test(k)) otras += cents(r[k]);
        else if (COLS_R_IIBB.test(k)) { if (cents(r[k])) jur.push(k.replace(/^R_INGRESOS BRUTOS /i, "")); }
        else if (COLS_P.test(k)) perc += cents(r[k]);
      }
      return {
        n: i,
        fecha: iso(r["FECHA DE VENTA"]),
        id: String(r["QR ID"]).trim(),
        terminal: String(r["TERMINAL LOGICA"] == null ? "" : r["TERMINAL LOGICA"]).trim(),
        cupon: String(r["NRO CUPON"] == null ? "" : r["NRO CUPON"]).trim(),
        estado: String(r["ESTADO"] == null ? "" : r["ESTADO"]).trim(),
        bruto: cents(r["MONTO BRUTO"]),
        neto: cents(r["MONTO NETO"]),
        arancel: cents(r["ARANCEL"]),
        ivaArancel: cents(r["IVA ARANCEL"]),
        retIibb: cents(r["TOTAL RETENCIONES (R)"]) - otras,
        retOtras: otras,
        percepciones: perc,
        jurisdicciones: jur,
      };
    });
  }

  // -------------------------------------------- lectura: Mercado Pago
  function leerLiberaciones(bytes) {
    const mat = filas(leerLibro(bytes));
    const iCab = mat.findIndex((f) => f.some((c) => String(c).trim() === "RECORD_TYPE"));
    if (iCab < 0) throw new Error("El archivo de liberaciones de Mercado Pago no tiene la columna RECORD_TYPE");
    return objetos(mat, iCab)
      .filter((r) => String(r["RECORD_TYPE"] == null ? "" : r["RECORD_TYPE"]).trim() === "release")
      .map((r, i) => ({
        n: i,
        fecha: iso(r["RELEASE_DATE"]),
        id: String(r["SOURCE_ID"] == null ? "" : r["SOURCE_ID"]).trim(),
        tipo: String(r["DESCRIPTION"] == null ? "" : r["DESCRIPTION"]).trim(),
        bruto: cents(r["GROSS_AMOUNT"]),
        comision: cents(r["MP_FEE_AMOUNT"]),
        impuestos: cents(r["TAXES_AMOUNT"]),
        detalleImp: String(r["TAX_DETAIL"] == null ? "" : r["TAX_DETAIL"]).trim(),
        neto: cents(r["NET_CREDIT_AMOUNT"]) - cents(r["NET_DEBIT_AMOUNT"]),
      }));
  }

  function leerResumenMp(bytes) {
    const mat = filas(leerLibro(bytes));
    const iCab = mat.findIndex((f) => f.some((c) => String(c).trim() === "REFERENCE_ID"));
    if (iCab < 0) throw new Error("El resumen de cuenta de Mercado Pago no tiene la columna REFERENCE_ID");
    return objetos(mat, iCab)
      .filter((r) => String(r["REFERENCE_ID"] == null ? "" : r["REFERENCE_ID"]).trim() !== "")
      .map((r, i) => ({
        n: i,
        fecha: iso(r["RELEASE_DATE"]),
        id: String(r["REFERENCE_ID"]).trim(),
        tipo: String(r["TRANSACTION_TYPE"] == null ? "" : r["TRANSACTION_TYPE"]).trim(),
        importe: cents(r["TRANSACTION_NET_AMOUNT"]),
      }));
  }

  // ------------------------------------------- lectura: Excel del mes anterior
  // La última columna la escribe una PERSONA, no el programa: si el mes que viene
  // ese Excel vuelve como entrada y la celda tiene texto, el pendiente se da por
  // cerrado a mano y deja de arrastrarse. Nunca entra al asiento por esa vía: el
  // motivo lo escribió alguien, no una clave única, y eso no alcanza para contabilizar.
  const CAB_PENDIENTES = ["Período de origen", "Origen", "Fecha", "Clave única", "Concepto", "Importe", "Períodos pendiente", "Cerrado a mano (motivo)"];

  function leerPendientes(bytes) {
    const wb = leerLibro(bytes);
    const mat = filas(wb, "Pendientes");
    const iCab = mat.findIndex((f) => String(f[0]).trim() === CAB_PENDIENTES[0]);
    if (iCab < 0) return [];
    return objetos(mat, iCab)
      .filter((r) => String(r["Clave única"] == null ? "" : r["Clave única"]).trim() !== "")
      .map((r) => ({
        periodo: String(r["Período de origen"] == null ? "" : r["Período de origen"]).trim(),
        origen: String(r["Origen"] == null ? "" : r["Origen"]).trim(),
        fecha: iso(r["Fecha"]) || String(r["Fecha"] == null ? "" : r["Fecha"]).trim(),
        clave: String(r["Clave única"]).trim(),
        concepto: String(r["Concepto"] == null ? "" : r["Concepto"]).trim(),
        importe: cents(r["Importe"]),
        meses: Number(r["Períodos pendiente"]) || 1,
        cerrado: String(r["Cerrado a mano (motivo)"] == null ? "" : r["Cerrado a mano (motivo)"]).trim(),
      }));
  }

  root.BancosLect = {
    num: num, iso: iso, ddmm: ddmm, periodoDe: periodoDe, nombrePeriodo: nombrePeriodo,
    cents: cents, pesos: pesos, detectar: detectar, usarLectorPdf: usarLectorPdf,
    leerLiquidacionPayway: leerLiquidacionPayway, leerExtracto: leerExtracto, leerQr: leerQr,
    leerLiberaciones: leerLiberaciones, leerResumenMp: leerResumenMp, leerPendientes: leerPendientes,
    CUENTAS: CUENTAS, CAB_PENDIENTES: CAB_PENDIENTES, CUIT_NARANJA: CUIT_NARANJA,
    CPT_PRISMA: CPT_PRISMA, CPT_QR: CPT_QR, CPT_DEV: CPT_DEV,
    CPT_IMP_CR: CPT_IMP_CR, CPT_IIBB_AC: CPT_IIBB_AC,
    filas: filas, objetos: objetos, leerLibro: leerLibro,
    get lectorPdf() { return leerPdf; },
  };
})(typeof window !== "undefined" ? window : globalThis);

/* Motor de la sección "Contabilización en bancos": los tres cruces, el asiento
 * mensual y el arrastre de pendientes. Ver bancos.js para los lectores. */
(function (root) {
  "use strict";

  const L = root.BancosLect;
  const XLSX = root.XLSX || (typeof require !== "undefined" ? require("xlsx-js-style") : null);
  const { cents, pesos, iso, ddmm, periodoDe, nombrePeriodo, CUENTAS } = L;

  // ------------------------------------------------------------------ cruces

  /** número de comercio que viene dentro del detalle del extracto */
  function comercioDe(detalle) {
    const m = String(detalle).match(/NRO COMERCIO:\s*(\d+)/i);
    return m ? m[1].replace(/^0+/, "") : "";
  }

  /** "03/08" del PDF -> "2026-08-03", tomando el año del período que se procesa */
  function fechaPagoAIso(ddmmTxt, periodo) {
    const [d, m] = ddmmTxt.split("/");
    let anio = Number(periodo.slice(0, 4));
    // un resumen de enero puede pagar algo fechado en diciembre del año anterior
    if (periodo.slice(5, 7) === "01" && m === "12") anio -= 1;
    return anio + "-" + m + "-" + d;
  }

  /**
   * Cruce A: liquidaciones Payway <-> extracto, por establecimiento + fecha de pago.
   * Es N:1 porque el banco acredita una línea por marca y el PDF da un solo total
   * del día. Cruza cuando la suma de las líneas del banco de ese establecimiento y
   * esa fecha es exactamente igual al total del día.
   * La unidad de pendiente es el ESTABLECIMIENTO: el desglose de descuentos es
   * mensual y no se puede partir por día, así que si un día no cruza queda afuera
   * del asiento todo el establecimiento.
   */
  function cruzarPayway(liqs, extracto, periodo) {
    const banco = new Map();          // est|fecha -> {importe, filas[]}
    for (const r of extracto) {
      if (r.concepto !== L.CPT_PRISMA) continue;
      const est = comercioDe(r.detalle);
      const k = est + "|" + r.fecha;
      if (!banco.has(k)) banco.set(k, { est: est, fecha: r.fecha, importe: 0, filas: [] });
      const g = banco.get(k);
      g.importe += r.importe;
      g.filas.push(r);
    }
    const usadas = new Set();
    const dias = [];
    for (const liq of liqs) {
      for (const d of liq.dias) {
        const f = fechaPagoAIso(d.fecha, periodo);
        const k = d.est + "|" + f;
        const g = banco.get(k);
        const cruza = !!g && g.importe === d.neto;
        if (cruza) usadas.add(k);
        dias.push(Object.assign({}, d, { fechaIso: f, clave: k, enBanco: g ? g.importe : null,
                                         lineasBanco: g ? g.filas.length : 0, cruza: cruza }));
      }
    }
    // establecimientos que entran al asiento: los que tienen TODOS sus días cruzados
    const ests = liqs.map((liq) => {
      const ds = dias.filter((d) => d.est === liq.est);
      return { liq: liq, dias: ds, cruza: ds.length > 0 && ds.every((d) => d.cruza),
               sinMovimientos: ds.length === 0 && liq.cab.presentado && liq.cab.presentado.pesos === 0 };
    });
    const huerfanas = [...banco.values()].filter((g) => !usadas.has(g.est + "|" + g.fecha));
    return { dias: dias, ests: ests, huerfanas: huerfanas };
  }

  /**
   * Cruce B: transferencias QR <-> extracto, por fecha + importe neto.
   * Se comparan las dos listas agrupadas: para cada (fecha, importe) cruza la
   * cantidad que coincide en ambos lados. No hay desempate posible ni necesario,
   * porque dos movimientos del mismo día por el mismo importe son intercambiables.
   * El sobrante de cualquiera de los dos lados queda pendiente.
   *
   * Hay dos bolsas separadas, porque en el extracto un cobro y una devolución
   * vienen con conceptos distintos: las aprobadas se buscan contra "COBRO CON QR"
   * y las devueltas contra "DEVOLUCION PEI". Al estar separadas no se pueden
   * mezclar: una devolución no puede consumir la acreditación de un cobro ni al
   * revés. Un estado que no sea ninguno de esos dos no se busca contra nada y
   * queda pendiente, para que un estado nuevo de Payway no se cuele sin que nadie
   * lo mire.
   */
  function cruzarQr(qr, extracto) {
    const bolsa = (concepto) => {       // fecha|importe -> filas del banco
      const m = new Map();
      for (const r of extracto) {
        if (r.concepto !== concepto) continue;
        const k = r.fecha + "|" + r.importe;
        if (!m.has(k)) m.set(k, []);
        m.get(k).push(r);
      }
      return m;
    };
    const cobros = bolsa(L.CPT_QR), devoluciones = bolsa(L.CPT_DEV);
    const filas = qr.map((r) => {
      const esCobro = r.estado === "Aprobada", esDev = r.estado === "Devuelto";
      if (!esCobro && !esDev) return Object.assign({}, r, { cruza: false, motivo: r.estado });
      const l = (esCobro ? cobros : devoluciones).get(r.fecha + "|" + r.neto);
      if (l && l.length) { l.shift(); return Object.assign({}, r, { cruza: true, motivo: "" }); }
      return Object.assign({}, r, { cruza: false, motivo: esCobro
        ? "sin acreditación en el extracto"
        : "devuelta por Payway, pero el extracto no tiene el débito de la devolución" });
    });
    const huerfanas = [];
    for (const m of [cobros, devoluciones]) for (const l of m.values()) for (const r of l) huerfanas.push(r);
    return { filas: filas, huerfanas: huerfanas };
  }

  /**
   * Cruce C: liberaciones de Mercado Pago <-> resumen de cuenta, por
   * SOURCE_ID = REFERENCE_ID. Lo que no cruza no entra al asiento.
   */
  const TIPOS_COBRO = ["payment", "refund"];
  const TIPOS_FUERA = ["payout", "reserve_for_payout", "reserve_for_refund"];

  function cruzarMp(liberaciones, resumen) {
    const ids = new Set(resumen.map((r) => r.id));
    return liberaciones.map((r) => {
      const alcance = TIPOS_COBRO.indexOf(r.tipo) >= 0 ? "cobro"
                    : r.tipo === "asset_management" ? "rendimiento"
                    : TIPOS_FUERA.indexOf(r.tipo) >= 0 ? "fuera" : "desconocido";
      const cruza = ids.has(r.id);
      return Object.assign({}, r, { alcance: alcance, cruza: cruza,
        motivo: cruza ? "" : "el SOURCE_ID no aparece en el resumen de cuenta" });
    });
  }

  /**
   * En las liberaciones, TAX_DETAIL trae el NOMBRE del impuesto ("debitos_creditos",
   * "santa_fe") y TAXES_AMOUNT el importe, uno por fila. Devuelve {clave, importe}
   * en positivo, o clave "" si la fila trae más de un impuesto y el importe no se
   * puede repartir (ahí se avisa y va a la línea "A DEFINIR", nunca se adivina).
   */
  function impuestoMp(r) {
    const claves = String(r.detalleImp).split(/[;,|]/).map((s) => s.trim()).filter(Boolean);
    return { claves: claves, importe: -r.impuestos };
  }

  // ------------------------------------------------------------- el asiento
  const CLAVES_IMP_MP = { debitos_creditos: "impDebCred", santa_fe: "retIibb" };

  /**
   * Los créditos de Tarjeta Naranja en el extracto, por CUIT en el detalle.
   * No se cruzan contra nada: no hay liquidación de Naranja todavía. Se los
   * identifica y se los manda enteros a una cuenta a conciliar, para que el
   * banco cierre y la plata quede a la vista hasta que aparezca el documento.
   */
  function creditosNaranja(extracto, periodo) {
    return extracto.filter((r) => r.importe > 0 && r.fecha.slice(0, 7) === periodo &&
                                  r.detalle.indexOf(L.CUIT_NARANJA) >= 0);
  }

  function armarAsiento(payway, qrCruce, mpCruce, resueltos, avisos, naranja) {
    const t = { banco: 0, mp: 0, gastos: 0, ivaCf: 0, retIibb: 0, retAfip: 0,
                impDebCred: 0, tarjetas: 0, naranja: 0, rendimiento: 0, sinCuenta: 0 };

    // --- Tarjeta Naranja: el crédito entra entero, sin abrir en descuentos
    for (const r of naranja || []) { t.banco += r.importe; t.naranja += r.importe; }

    // --- Payway: solo establecimientos enteros que cruzaron
    for (const e of payway.ests) {
      if (!e.cruza) continue;
      for (const d of e.dias) { t.banco += d.neto; t.tarjetas += d.presentado; }
      t.gastos += e.liq.desc.arancel + e.liq.desc.financiero;
      t.ivaCf += e.liq.desc.iva;
      t.retIibb += e.liq.desc.retIibb;
      t.retAfip += e.liq.desc.retAfip;
    }

    // --- QR: solo las transferencias que cruzaron
    for (const r of qrCruce.filas) {
      if (!r.cruza) continue;
      t.banco += r.neto;
      t.tarjetas += r.bruto;
      t.gastos += r.arancel;
      t.ivaCf += r.ivaArancel;
      t.retIibb += r.retIibb;
      if (r.retOtras || r.percepciones) {
        t.sinCuenta += r.retOtras + r.percepciones;
        avisos.push("QR " + r.id + ": tiene retenciones o percepciones sin cuenta asignada por " +
                    fmt(r.retOtras + r.percepciones) + ". Quedan en la línea \"A DEFINIR\".");
      }
    }

    // --- Mercado Pago: cobros y rendimientos que cruzaron
    for (const r of mpCruce) {
      if (!r.cruza) continue;
      if (r.alcance === "cobro") {
        t.mp += r.neto;
        t.tarjetas += r.bruto;
        t.gastos += -r.comision;
        const imp = impuestoMp(r);
        if (imp.importe) {
          const cuenta = imp.claves.length === 1 ? CLAVES_IMP_MP[imp.claves[0]] : null;
          if (cuenta) t[cuenta] += imp.importe;
          else {
            t.sinCuenta += imp.importe;
            avisos.push("Mercado Pago " + r.id + ": el impuesto \"" + r.detalleImp + "\" (" +
                        fmt(imp.importe) + ") no tiene cuenta asignada. Queda en la línea \"A DEFINIR\".");
          }
        }
      } else if (r.alcance === "rendimiento") {
        t.mp += r.neto;
        t.rendimiento += r.neto;
      }
    }

    // --- pendientes de períodos anteriores que se resolvieron este mes
    // Los de origen "Extracto" no contabilizan: son movimientos del banco sin
    // documento que los respalde, y el documento es el que dice a qué cuentas van.
    for (const v of resueltos || []) {
      if (!v.contabiliza) continue;
      t.banco += v.origen === "Mercado Pago" ? 0 : v.neto;
      t.mp += v.origen === "Mercado Pago" ? v.neto : 0;
      t.tarjetas += v.bruto;
      t.gastos += v.arancel + v.financiero;
      t.ivaCf += v.iva;
      t.retIibb += v.retIibb;
      t.retAfip += v.retAfip;
      t.impDebCred += v.impDebCred;
      t.rendimiento += v.rendimiento;
    }

    const lineas = [
      { cuenta: CUENTAS.banco, debe: t.banco, haber: 0 },
      { cuenta: CUENTAS.mp, debe: t.mp, haber: 0 },
      { cuenta: CUENTAS.gastos, debe: t.gastos, haber: 0 },
      { cuenta: CUENTAS.ivaCf, debe: t.ivaCf, haber: 0 },
      { cuenta: CUENTAS.retIibb, debe: t.retIibb, haber: 0 },
      { cuenta: CUENTAS.retAfip, debe: t.retAfip, haber: 0 },
      { cuenta: CUENTAS.impDebCred, debe: t.impDebCred, haber: 0 },
      { cuenta: CUENTAS.sinCuenta, debe: t.sinCuenta, haber: 0 },
      { cuenta: CUENTAS.tarjetas, debe: 0, haber: t.tarjetas },
      { cuenta: CUENTAS.naranja, debe: 0, haber: t.naranja },
      { cuenta: CUENTAS.rendimiento, debe: 0, haber: t.rendimiento },
    ].filter((l) => l.debe !== 0 || l.haber !== 0);

    const debe = lineas.reduce((a, l) => a + l.debe, 0);
    const haber = lineas.reduce((a, l) => a + l.haber, 0);
    if (debe !== haber) {
      avisos.push("El asiento no cierra: debe " + fmt(debe) + " contra haber " + fmt(haber) +
                  ", diferencia " + fmt(debe - haber) + ". No se corrigió nada, revisá los controles.");
    }
    return { lineas: lineas, debe: debe, haber: haber, dif: debe - haber, t: t };
  }

  const fmt = (c) => "$ " + pesos(c).toLocaleString("es-AR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

  // ----------------------------------------------------------- pendientes
  /**
   * Arma los pendientes de este mes y resuelve los que venían del anterior.
   * Un pendiente se resuelve solo si reaparece POR SU MISMA CLAVE ÚNICA.
   *
   * Hay una segunda salida, para los que nunca van a reaparecer porque el archivo
   * los trae con otro identificador: alguien escribe el motivo en la columna
   * "Cerrado a mano (motivo)" del Excel y el mes siguiente dejan de arrastrarse.
   * La clave le gana al cierre a mano: si el movimiento aparece de verdad se
   * resuelve por las buenas y se contabiliza, aunque la celda esté escrita. Un
   * cerrado a mano NO entra al asiento en ningún caso.
   */
  function armarPendientes(periodo, payway, qrCruce, mpCruce, previos) {
    const nuevos = [];
    const p = (o) => nuevos.push(Object.assign({ periodo: periodo, meses: 1, contabiliza: true,
      bruto: 0, neto: 0, arancel: 0, financiero: 0, iva: 0, retIibb: 0, retAfip: 0,
      impDebCred: 0, rendimiento: 0 }, o));

    for (const e of payway.ests) {
      if (e.cruza || e.sinMovimientos) continue;
      const neto = e.dias.reduce((a, d) => a + d.neto, 0);
      const noCruzan = e.dias.filter((d) => !d.cruza);
      p({ origen: "Payway", fecha: "", clave: "establecimiento " + e.liq.est,
          concepto: "Liquidación Payway del establecimiento " + e.liq.est + ": " + noCruzan.length +
                    " de " + e.dias.length + " días no cruzan con el extracto (" +
                    noCruzan.map((d) => d.fecha).join(", ") + ")",
          importe: neto,
          bruto: e.dias.reduce((a, d) => a + d.presentado, 0), neto: neto,
          arancel: e.liq.desc.arancel, financiero: e.liq.desc.financiero, iva: e.liq.desc.iva,
          retIibb: e.liq.desc.retIibb, retAfip: e.liq.desc.retAfip });
    }
    for (const g of payway.huerfanas) {
      p({ origen: "Extracto", fecha: g.fecha, clave: "prisma " + g.est + "|" + g.fecha,
          concepto: "Acreditación de Comercios Prisma del establecimiento " + g.est +
                    " sin liquidación que la respalde (" + g.filas.length + " líneas). Revisión manual.",
          importe: g.importe, contabiliza: false });
    }
    for (const r of qrCruce.filas) {
      if (r.cruza) continue;
      p({ origen: "QR", fecha: r.fecha, clave: "qr " + r.id,
          concepto: "Transferencia QR " + r.id + " en estado \"" + r.estado + "\": " + r.motivo,
          importe: r.neto, bruto: r.bruto, neto: r.neto, arancel: r.arancel, iva: r.ivaArancel,
          retIibb: r.retIibb });
    }
    for (const g of qrCruce.huerfanas) {
      const dev = g.concepto === L.CPT_DEV;
      p({ origen: "Extracto", fecha: g.fecha,
          clave: (dev ? "qrdev " : "qrext ") + g.fecha + "|" + g.importe,
          concepto: (dev
            ? "Devolución debitada en el extracto sin transferencia QR devuelta que la respalde: "
            : "Cobro con QR acreditado en el extracto sin transferencia que lo respalde: ") +
                    g.detalle + ". Revisión manual.",
          importe: g.importe, contabiliza: false });
    }
    for (const r of mpCruce) {
      if (r.cruza || r.alcance === "fuera") continue;
      if (r.alcance === "rendimiento") {
        p({ origen: "Mercado Pago", fecha: r.fecha, clave: "mp " + r.id,
            concepto: "Rendimiento de Mercado Pago: " + r.motivo,
            importe: r.neto, neto: r.neto, rendimiento: r.neto });
        continue;
      }
      const imp = impuestoMp(r);
      const clave = imp.claves.length === 1 ? CLAVES_IMP_MP[imp.claves[0]] : null;
      p({ origen: "Mercado Pago", fecha: r.fecha, clave: "mp " + r.id,
          concepto: "Liberación de Mercado Pago (" + r.tipo + "): " + r.motivo,
          importe: r.neto, bruto: r.bruto, neto: r.neto, arancel: -r.comision,
          retIibb: clave === "retIibb" ? imp.importe : 0,
          impDebCred: clave === "impDebCred" ? imp.importe : 0,
          contabiliza: r.alcance === "cobro" && (!imp.importe || !!clave) });
    }

    // --- resolución de los que venían del mes anterior
    const claves = {
      payway: new Set(payway.ests.filter((e) => e.cruza).map((e) => "establecimiento " + e.liq.est)),
      prisma: new Set(payway.dias.filter((d) => d.cruza).map((d) => "prisma " + d.clave)),
      qr: new Set(qrCruce.filas.filter((r) => r.cruza).map((r) => "qr " + r.id)),
      mp: new Set(mpCruce.filter((r) => r.cruza).map((r) => "mp " + r.id)),
    };
    const resueltos = [], siguen = [], cerrados = [];
    for (const v of previos) {
      const resuelto = claves.payway.has(v.clave) || claves.prisma.has(v.clave) ||
                       claves.qr.has(v.clave) || claves.mp.has(v.clave);
      if (resuelto) resueltos.push(v);
      else if (v.cerrado) cerrados.push(v);
      else siguen.push(Object.assign({}, v, { meses: v.meses + 1 }));
    }
    return { nuevos: nuevos, resueltos: resueltos, siguen: siguen, cerrados: cerrados };
  }

  root.BancosMotor = {
    comercioDe, fechaPagoAIso, cruzarPayway, cruzarQr, cruzarMp,
    impuestoMp, creditosNaranja, armarAsiento, armarPendientes, fmt, CLAVES_IMP_MP,
  };
})(typeof window !== "undefined" ? window : globalThis);

/* Orquestador y Excel de salida de la sección "Contabilización en bancos". */
(function (root) {
  "use strict";

  const L = root.BancosLect, M = root.BancosMotor;
  const XLSX = root.XLSX || (typeof require !== "undefined" ? require("xlsx-js-style") : null);
  const { cents, pesos, ddmm, nombrePeriodo, CUENTAS } = L;

  const { hoja, fechaExcel } = root.Conciliador;
  const fx = (i) => (i ? fechaExcel(i) : "");

  // --------------------------------------------------- PDF y ZIP de Payway
  /**
   * pdf.js no devuelve renglones sino fragmentos con su posición. Se agrupan por
   * la coordenada vertical y se ordenan por la horizontal; solo se mete un espacio
   * cuando hay hueco real entre un fragmento y el siguiente, porque si no
   * "1.234" + ",56" terminaría siendo "1.234 ,56".
   */
  async function textoDePdf(pdfjsLib, bytes) {
    const doc = await pdfjsLib.getDocument({ data: bytes }).promise;
    const lineas = [];
    for (let p = 1; p <= doc.numPages; p++) {
      const contenido = await (await doc.getPage(p)).getTextContent();
      const porY = new Map();
      for (const it of contenido.items) {
        if (!it.str) continue;
        const x = it.transform[4], y = Math.round(it.transform[5] / 2) * 2;
        let g = null;
        for (const k of porY.keys()) if (Math.abs(k - y) <= 2) { g = k; break; }
        if (g === null) { g = y; porY.set(g, []); }
        porY.get(g).push({ x: x, w: it.width || 0, s: it.str });
      }
      for (const y of [...porY.keys()].sort((a, b) => b - a)) {
        const items = porY.get(y).sort((a, b) => a.x - b.x);
        let linea = "", fin = null;
        for (const it of items) {
          if (fin !== null && it.x - fin > 0.6) linea += " ";
          linea += it.s;
          fin = it.x + it.w;
        }
        lineas.push(linea.trim());
      }
    }
    return lineas.join("\n");
  }

  /** los PDF que vienen dentro del ZIP de liquidaciones */
  function pdfsDelZip(fflate, bytes) {
    const zip = fflate.unzipSync(bytes);
    return Object.keys(zip)
      .filter((n) => /\.pdf$/i.test(n) && n.indexOf("__MACOSX") !== 0)
      .map((n) => ({ archivo: n.split("/").pop(), bytes: zip[n] }));
  }

  // ------------------------------------------------------------ orquestador
  /**
   * entradas = {
   *   liquidaciones: [{archivo, texto}],   // los PDF del ZIP, ya pasados a texto
   *   extracto, qr, libMp, resumenMp,      // Uint8Array
   *   previo                               // Uint8Array del Excel del mes anterior, o null
   * }
   */
  function procesar(entradas) {
    const avisos = [];
    const faltan = ["liquidaciones", "extracto", "qr", "libMp", "resumenMp"]
      .filter((k) => !entradas[k] || (Array.isArray(entradas[k]) && !entradas[k].length));
    if (faltan.length) throw new Error("Faltan archivos: " + faltan.join(", "));

    const extracto = L.leerExtracto(entradas.extracto);
    const periodos = {};
    for (const r of extracto) periodos[r.fecha.slice(0, 7)] = (periodos[r.fecha.slice(0, 7)] || 0) + 1;
    const orden = Object.entries(periodos).sort((a, b) => b[1] - a[1]);
    const periodo = orden[0][0];
    if (orden.length > 1) {
      avisos.push("El extracto abarca más de un mes (" + orden.map((x) => x[0] + ": " + x[1] + " mov.").join(", ") +
                  "). Se toma " + nombrePeriodo(periodo) + " y el resto queda afuera del cruce.");
    }

    const liqs = entradas.liquidaciones.map((x) => L.leerLiquidacionPayway(x.texto, x.archivo));
    for (const l of liqs) {
      const suma = l.dias.reduce((a, d) => a + d.neto, 0);
      if (l.cab.saldo && suma !== l.cab.saldo.pesos) {
        avisos.push("Liquidación " + l.archivo + ": la suma de los totales diarios (" + M.fmt(suma) +
                    ") no coincide con el SALDO de la cabecera (" + M.fmt(l.cab.saldo.pesos) + ").");
      }
      if (l.cab.saldo && l.cab.saldo.dolares) {
        avisos.push("Liquidación " + l.archivo + ": tiene importes en dólares. Esta sección solo contabiliza pesos.");
      }
    }

    const qrTodo = L.leerQr(entradas.qr);
    const qr = qrTodo.filter((r) => r.fecha.slice(0, 7) === periodo);
    if (qr.length !== qrTodo.length) {
      avisos.push("El archivo de QR trae " + (qrTodo.length - qr.length) + " transferencias de otros meses. " +
                  "Solo se cruzan las de " + nombrePeriodo(periodo) + ".");
    }
    const libMp = L.leerLiberaciones(entradas.libMp);
    const resumenMp = L.leerResumenMp(entradas.resumenMp);

    let previos = [];
    if (entradas.previo) {
      previos = L.leerPendientes(entradas.previo);
      if (!previos.length) avisos.push("El Excel del mes anterior no tiene pendientes para arrastrar.");
    } else {
      avisos.push("No adjuntaste el Excel del mes anterior: no se arrastra ningún pendiente. " +
                  "Si el mes pasado quedó algo sin cruzar, va a quedar afuera.");
    }

    const payway = M.cruzarPayway(liqs, extracto, periodo);
    const qrCruce = M.cruzarQr(qr, extracto);
    const mpCruce = M.cruzarMp(libMp, resumenMp);

    const desconocidos = mpCruce.filter((r) => r.alcance === "desconocido");
    for (const r of desconocidos) {
      avisos.push("Mercado Pago " + r.id + ': tipo de liberación "' + r.tipo +
                  '" que esta sección no conoce. No entra al asiento, queda en Pendientes.');
    }

    const naranja = M.creditosNaranja(extracto, periodo);
    if (naranja.length) {
      avisos.push(naranja.length + " crédito(s) de Tarjeta Naranja en el extracto por " +
                  M.fmt(naranja.reduce((a, r) => a + r.importe, 0)) + ". Naranja liquida por " +
                  "transferencia, fuera de Payway: no hay liquidación que abra el bruto y los " +
                  'descuentos, así que el crédito va entero a "' + L.CUENTAS.naranja + '".');
    }

    const pend = M.armarPendientes(periodo, payway, qrCruce, mpCruce, previos);
    const asiento = M.armarAsiento(payway, qrCruce, mpCruce, pend.resueltos, avisos, naranja);

    if (pend.cerrados.length) {
      avisos.push(pend.cerrados.length + " pendiente(s) que venían de antes figuran cerrados a mano en el " +
                  "Excel del mes anterior. No se arrastran más y no entran a este asiento. " +
                  'Están con su motivo en la hoja "Cerrados a mano".');
    }

    const vencidos = pend.siguen.filter((v) => v.meses >= 2);
    if (vencidos.length) {
      avisos.push(vencidos.length + " pendiente(s) llevan 2 períodos o más sin resolverse. Revisalos a mano.");
    }

    // impuestos que cobra el banco: se muestran, no van a este asiento
    const impBanco = [L.CPT_IMP_CR, L.CPT_IIBB_AC].map((c) => {
      const fs = extracto.filter((r) => r.concepto === c);
      return { concepto: c, cant: fs.length, importe: fs.reduce((a, r) => a + r.importe, 0) };
    });

    return { periodo, avisos, extracto, liqs, qr, libMp, resumenMp, payway, qrCruce, mpCruce,
             pendientes: pend, asiento, impBanco, naranja, previos };
  }

  // --------------------------------------------------------------- el Excel
  const CAB_PEND = L.CAB_PENDIENTES.concat(
    ["Contabiliza", "Bruto", "Neto", "Arancel", "Costo financiero", "IVA", "Ret. IIBB", "Ret. AFIP",
     "Imp. déb. y créd.", "Rendimiento"]);

  function filaPend(v) {
    return [nombrePeriodo(v.periodo), v.origen, fx(v.fecha), v.clave, v.concepto, pesos(v.importe), v.meses,
            v.cerrado || "",
            v.contabiliza ? "sí" : "no", pesos(v.bruto), pesos(v.neto), pesos(v.arancel), pesos(v.financiero),
            pesos(v.iva), pesos(v.retIibb), pesos(v.retAfip), pesos(v.impDebCred), pesos(v.rendimiento || 0)];
  }

  function armarLibro(R, leidos) {
    const wb = XLSX.utils.book_new();
    const A = R.asiento;

    // ---------------- Resumen
    const res = [
      ["Período", nombrePeriodo(R.periodo)],
      ["Procesado", new Date().toLocaleString("es-AR")],
      [],
      ["CÓMO SE CRUZA", "Clave única", "Cruzan", "No cruzan"],
      ["Liquidaciones Payway contra el extracto", "nº de establecimiento + fecha de pago",
       R.payway.dias.filter((d) => d.cruza).length, R.payway.dias.filter((d) => !d.cruza).length],
      ["Transferencias QR contra el extracto", "fecha de venta + importe neto",
       R.qrCruce.filas.filter((r) => r.cruza).length, R.qrCruce.filas.filter((r) => !r.cruza).length],
      ["Liberaciones de Mercado Pago contra el resumen de cuenta", "SOURCE_ID = REFERENCE_ID",
       R.mpCruce.filter((r) => r.cruza).length, R.mpCruce.filter((r) => !r.cruza).length],
      [],
      ["LO QUE ENTRA AL ASIENTO", "Bruto", "Neto acreditado"],
      ["Tarjetas Payway (establecimientos completos)",
       pesos(R.payway.ests.filter((e) => e.cruza).reduce((a, e) => a + e.dias.reduce((b, d) => b + d.presentado, 0), 0)),
       pesos(R.payway.ests.filter((e) => e.cruza).reduce((a, e) => a + e.dias.reduce((b, d) => b + d.neto, 0), 0))],
      ["Transferencias QR",
       pesos(R.qrCruce.filas.filter((r) => r.cruza).reduce((a, r) => a + r.bruto, 0)),
       pesos(R.qrCruce.filas.filter((r) => r.cruza).reduce((a, r) => a + r.neto, 0))],
      ["Cobros de Mercado Pago",
       pesos(R.mpCruce.filter((r) => r.cruza && r.alcance === "cobro").reduce((a, r) => a + r.bruto, 0)),
       pesos(R.mpCruce.filter((r) => r.cruza && r.alcance === "cobro").reduce((a, r) => a + r.neto, 0))],
      ["Créditos de Tarjeta Naranja (" + R.naranja.length + ", sin liquidación: a conciliar)", "",
       pesos(R.naranja.reduce((a, r) => a + r.importe, 0))],
      ["Rendimientos de Mercado Pago", "",
       pesos(R.mpCruce.filter((r) => r.cruza && r.alcance === "rendimiento").reduce((a, r) => a + r.neto, 0))],
      [],
      ["PENDIENTES", "Cantidad", "Importe"],
      ["Nuevos de este período", R.pendientes.nuevos.length,
       pesos(R.pendientes.nuevos.reduce((a, v) => a + v.importe, 0))],
      ["Que venían de antes y se resolvieron", R.pendientes.resueltos.length,
       pesos(R.pendientes.resueltos.reduce((a, v) => a + v.importe, 0))],
      ["Que venían de antes y siguen", R.pendientes.siguen.length,
       pesos(R.pendientes.siguen.reduce((a, v) => a + v.importe, 0))],
      ["Cerrados a mano en el Excel anterior (no entran al asiento)", R.pendientes.cerrados.length,
       pesos(R.pendientes.cerrados.reduce((a, v) => a + v.importe, 0))],
      [],
      ["FUERA DE ESTE ASIENTO", "Cantidad", "Importe"],
      ...R.impBanco.map((x) => [x.concepto + " (va en el asiento mensual de impuestos)", x.cant, pesos(x.importe)]),
      ["Transferencias salientes de Mercado Pago (fuera del alcance)",
       R.mpCruce.filter((r) => r.alcance === "fuera").length,
       pesos(R.mpCruce.filter((r) => r.alcance === "fuera").reduce((a, r) => a + r.neto, 0))],
      [],
      ["ARCHIVOS LEÍDOS"],
      ...leidos.map((x) => [x.tipo, x.nombre]),
    ];
    if (R.avisos.length) res.push([], ["AVISOS"], ...R.avisos.map((a) => [a]));
    const ancho = 4;
    const resPad = res.map((f) => f.concat(Array(Math.max(0, ancho - f.length)).fill("")));
    XLSX.utils.book_append_sheet(wb, hoja(resPad[0], resPad.slice(1), {
      sinCabecera: true, importes: [1, 2],
      negrita: (r) => /^[A-ZÁÉÍÓÚÑ ,.()]+$/.test(String(resPad[r][0])) && String(resPad[r][0]).trim() !== "",
    }), "Resumen");

    // ---------------- Asiento
    const asi = A.lineas.map((l) => [l.cuenta, l.debe ? pesos(l.debe) : "", l.haber ? pesos(l.haber) : ""]);
    asi.push(["TOTALES", pesos(A.debe), pesos(A.haber)]);
    asi.push(["Diferencia", pesos(A.dif), ""]);
    XLSX.utils.book_append_sheet(wb, hoja(["Cuenta", "Debe", "Haber"], asi, {
      importes: [1, 2],
      negrita: (r) => r >= asi.length - 1,
      rojo: (r) => r === asi.length && A.dif !== 0,
    }), "Asiento");

    // ---------------- Control tarjetas
    const ct = R.payway.dias.slice().sort((a, b) => (a.est + a.fechaIso).localeCompare(b.est + b.fechaIso))
      .map((d) => [d.est, fx(d.fechaIso), pesos(d.presentado), pesos(d.descuento), pesos(d.neto),
                   d.enBanco === null ? "" : pesos(d.enBanco), d.lineasBanco,
                   d.enBanco === null ? "" : pesos(d.enBanco - d.neto),
                   d.cruza ? "cruza" : "NO CRUZA", d.archivo]);
    for (const g of R.payway.huerfanas) {
      ct.push([g.est, fx(g.fecha), "", "", "", pesos(g.importe), g.filas.length, "",
               "SOLO EN EL EXTRACTO", ""]);
    }
    XLSX.utils.book_append_sheet(wb, hoja(
      ["Establecimiento", "Fecha de pago", "Presentado", "Descuento", "Neto de la liquidación",
       "Acreditado en el extracto", "Líneas del extracto", "Diferencia", "Estado", "Archivo"], ct,
      { importes: [2, 3, 4, 5, 7], fechas: [1], rojo: (r) => /NO CRUZA|SOLO EN/.test(String(ct[r - 1][8])) }),
      "Control tarjetas");

    // ---------------- Descuentos Payway
    const cd = [];
    for (const l of R.liqs) {
      const e = R.payway.ests.filter((x) => x.liq === l)[0];
      for (const d of l.detalle) {
        cd.push([l.est, d.etiqueta, pesos(d.importe), d.concepto,
                 e && e.cruza ? "entra al asiento" : "no entra", l.archivo]);
      }
    }
    XLSX.utils.book_append_sheet(wb, hoja(
      ["Establecimiento", "Concepto del resumen", "Importe", "Va a", "Estado", "Archivo"], cd,
      { importes: [2], rojo: (r) => String(cd[r - 1][4]) === "no entra" }), "Descuentos Payway");

    // ---------------- Control QR
    const cq = R.qrCruce.filas.map((r) => [fx(r.fecha), r.id, r.terminal, r.cupon, r.estado,
      pesos(r.bruto), pesos(r.arancel), pesos(r.ivaArancel), pesos(r.retIibb),
      pesos(r.retOtras), pesos(r.percepciones), pesos(r.neto),
      r.cruza ? "cruza" : "NO CRUZA", r.motivo]);
    for (const g of R.qrCruce.huerfanas) {
      cq.push([fx(g.fecha), "", "", "", "", "", "", "", "", "", "", pesos(g.importe),
               "SOLO EN EL EXTRACTO", g.concepto + ": " + g.detalle]);
    }
    XLSX.utils.book_append_sheet(wb, hoja(
      ["Fecha de venta", "QR ID", "Terminal lógica", "Nro cupón", "Estado", "Bruto", "Arancel",
       "IVA arancel", "Ret. IIBB", "Otras retenciones", "Percepciones", "Neto", "Cruce", "Motivo"], cq,
      { importes: [5, 6, 7, 8, 9, 10, 11], fechas: [0],
        rojo: (r) => /NO CRUZA|SOLO EN/.test(String(cq[r - 1][12])) }), "Control QR");

    // ---------------- Control Mercado Pago
    const cm = R.mpCruce.map((r) => [fx(r.fecha), r.id, r.tipo, r.alcance,
      pesos(r.bruto), pesos(-r.comision), r.detalleImp, pesos(-r.impuestos), pesos(r.neto),
      r.cruza ? "cruza" : "NO CRUZA", r.motivo]);
    XLSX.utils.book_append_sheet(wb, hoja(
      ["Fecha de liberación", "SOURCE_ID", "Tipo", "Alcance", "Bruto", "Comisión", "Impuesto",
       "Importe del impuesto", "Neto", "Cruce", "Motivo"], cm,
      { importes: [4, 5, 7, 8], fechas: [0], rojo: (r) => String(cm[r - 1][9]) === "NO CRUZA" }),
      "Control Mercado Pago");

    // ---------------- Impuestos del banco (fuera del asiento)
    const ci = [];
    for (const c of [L.CPT_IMP_CR, L.CPT_IIBB_AC]) {
      for (const r of R.extracto.filter((x) => x.concepto === c)) {
        ci.push([fx(r.fecha), r.concepto, r.detalle, pesos(r.importe)]);
      }
    }
    XLSX.utils.book_append_sheet(wb, hoja(
      ["Fecha", "Concepto", "Detalle", "Importe"], ci, { importes: [3], fechas: [0] }),
      "Impuestos del banco");

    // ---------------- Tarjeta Naranja
    if (R.naranja.length) {
      XLSX.utils.book_append_sheet(wb, hoja(
        ["Fecha", "Hora", "Concepto", "Detalle", "Importe acreditado"],
        R.naranja.map((r) => [fx(r.fecha), r.hora, r.concepto, r.detalle, pesos(r.importe)]),
        { importes: [4], fechas: [0] }), "Tarjeta Naranja");
    }

    // ---------------- Pendientes
    const pfilas = R.pendientes.siguen.concat(R.pendientes.nuevos).map(filaPend);
    XLSX.utils.book_append_sheet(wb, hoja(CAB_PEND, pfilas, {
      importes: [5, 9, 10, 11, 12, 13, 14, 15, 16, 17], fechas: [2],
      rojo: (r) => Number(pfilas[r - 1][6]) >= 2 }), "Pendientes");

    // ---------------- Pendientes resueltos
    if (R.pendientes.resueltos.length) {
      XLSX.utils.book_append_sheet(wb, hoja(CAB_PEND, R.pendientes.resueltos.map(filaPend), {
        importes: [5, 9, 10, 11, 12, 13, 14, 15, 16, 17], fechas: [2] }), "Pendientes resueltos");
    }

    // ---------------- Pendientes cerrados a mano
    if (R.pendientes.cerrados.length) {
      XLSX.utils.book_append_sheet(wb, hoja(CAB_PEND, R.pendientes.cerrados.map(filaPend), {
        importes: [5, 9, 10, 11, 12, 13, 14, 15, 16, 17], fechas: [2] }), "Cerrados a mano");
    }

    return wb;
  }

  root.Bancos = { procesar, armarLibro, textoDePdf, pdfsDelZip, CAB_PEND,
                  detectar: L.detectar, nombrePeriodo: nombrePeriodo };
  root.Conciliador.Bancos = root.Bancos;
})(typeof window !== "undefined" ? window : globalThis);
