"""
Conciliación de cobros con tarjeta / QR / Mercado Pago - Grupo Oroño

Cruza los recibos de caja (columnas de tarjeta y videoconsultas) contra:
  - Movimientos presentados a Payway (CSV)
  - QR transferencia de Payway (CSV)
  - Movimientos de Mercado Pago (XLSX)

Reglas de conciliación (sin desempates ni tolerancias):
  1. Par directo: misma FECHA y mismo IMPORTE exacto, y esa combinación
     identifica a un único recibo y a un único movimiento.
  2. Grupo que cierra: dentro de una misma fecha e importe, si la cantidad de
     recibos es exactamente igual a la de movimientos, el grupo cierra y se
     concilia en bloque. El apareo dentro del grupo va por orden de carga: el
     total del grupo está verificado, el par individual no.
  En las dos, los recibos TVIR (videoconsultas) solo se cruzan contra MP; un
  grupo que empata en cantidad pero no tiene movimientos de MP suficientes para
  sus videoconsultas NO cierra. Todo lo demás queda para revisión manual.

Uso: dejar los archivos en la carpeta "Entrada" y ejecutar. El resultado queda
en "Salida/Conciliacion_<desde>_<hasta>.xlsx".
"""
import sys
import warnings
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore")

COLS_TARJETA = ["TSNS", "TSNI", "TSNZ"]  # amarillas: cobros con tarjeta / QR
COLS_MP = ["TVIR"]  # naranja: videoconsultas, solo Mercado Pago

if getattr(sys, "frozen", False):
    BASE = Path(sys.executable).parent
else:
    BASE = Path(__file__).parent
ENTRADA = BASE / "Entrada"
SALIDA = BASE / "Salida"


# ---------------------------------------------------------------- lectura
def leer_csv(path, skiprows=0):
    for enc in ("utf-8-sig", "latin1"):
        try:
            return pd.read_csv(path, sep=";", skiprows=skiprows, encoding=enc, dtype=str)
        except UnicodeDecodeError:
            continue
    raise ValueError(f"No se pudo leer {path.name}")


def num(s):
    return pd.to_numeric(s, errors="coerce")


def leer_recibos(path):
    raw = pd.read_excel(path, header=None)
    fila = raw.index[raw.iloc[:, 0].astype(str).str.strip() == "Cajero"][0]
    df = pd.read_excel(path, header=fila)
    df.columns = [str(c).strip() for c in df.columns]
    faltan = [c for c in COLS_TARJETA + COLS_MP if c not in df.columns]
    if faltan:
        print(f"  ATENCIÓN: {path.name} no tiene las columnas {faltan}")
    filas = []
    for _, x in df.iterrows():
        for c in COLS_TARJETA + COLS_MP:
            if c not in df.columns:
                continue
            v = num(x[c])
            if pd.notna(v) and v != 0:
                filas.append({
                    "Fecha": pd.Timestamp(x["Fecha Caja"]).normalize(),
                    "Nro Recibo": int(x["Nro Recibo"]),
                    "Cliente": str(x["Nombre Cliente"]).strip(),
                    "Cajero": str(x["Cajero"]).strip(),
                    "Columna": c,
                    "Tipo Asiento": str(x["Tipo de Asiento"]).strip(),
                    "Importe": round(float(v), 2),
                    "Importe Total Recibo": round(float(num(x["Importe Total"])), 2),
                })
    return pd.DataFrame(filas)


def leer_payway(path):
    df = leer_csv(path)
    return pd.DataFrame({
        "Fuente": "PAYWAY",
        "Fecha": pd.to_datetime(df["COMPRA"], dayfirst=True),
        "Importe": num(df["MONTO_BRUTO"]).round(2),
        "Referencia": "Lote " + df["LOTE"].str.strip() + " / Cupón " + df["NUM.CUPON"].str.strip()
                      + " / Aut " + df["NRO_AUT"].str.strip(),
        "Detalle": df["TIPO"].str.strip() + " - " + df["MARCA"].str.strip() + " - "
                   + df["DETALLE"].str.strip() + " - Tarj " + df["NUM.TARJETA"].str[-4:],
        "Terminal / Caja": "Estab " + df["ESTABLECIMIENTO"].str.strip(),
    })


def leer_qr(path):
    df = leer_csv(path, skiprows=1)
    return pd.DataFrame({
        "Fuente": "QR",
        "Fecha": pd.to_datetime(df["FECHA DE VENTA"]),
        "Importe": num(df["MONTO BRUTO"]).round(2),
        "Referencia": "QR " + df["QR ID"].str.strip() + " / Cupón " + df["NRO CUPON"].str.strip(),
        "Detalle": df["ESTADO"].str.strip(),
        "Terminal / Caja": "Terminal " + df["TERMINAL LOGICA"].str.strip(),
    })


def leer_mp(path):
    df = pd.read_excel(path, dtype={"ID DE OPERACIÓN EN MERCADO PAGO": str})
    caja = (df["NOMBRE DE LOCAL"].fillna("") + " " + df["NOMBRE DE CAJA"].fillna("")).str.strip()
    detalle = (df["TIPO DE OPERACIÓN"].fillna("") + " - " + df["MEDIO DE PAGO"].fillna("")
               + " - " + df["DETALLE DE LA VENTA"].fillna("").str.replace('"', "")
               + " - Pagador: " + df["PAGADOR"].fillna(""))
    return pd.DataFrame({
        "Fuente": "MP",
        "Fecha": pd.to_datetime(df["FECHA DE ORIGEN"].str[:10]),
        "Importe": num(df["VALOR DE LA COMPRA"]).round(2),
        "Referencia": "MP " + df["ID DE OPERACIÓN EN MERCADO PAGO"].astype(str),
        "Detalle": detalle,
        "Terminal / Caja": caja,
    })


def detectar(path):
    """Identifica cada archivo por su contenido, no por el nombre."""
    ext = path.suffix.lower()
    if ext == ".csv":
        with open(path, encoding="latin1") as f:
            cab = f.read(3000)
        if "Detalle de Transferencias" in cab or "QR ID" in cab:
            return "qr"
        if "NUM.CUPON" in cab:
            return "payway"
    elif ext in (".xls", ".xlsx"):
        cab = pd.read_excel(path, header=None, nrows=8).astype(str).to_string()
        if "Listado de caja" in cab:
            return "recibos"
        if "MERCADO PAGO" in cab.upper():
            return "mp"
    return None


# ---------------------------------------------------------------- conciliación
def conciliar(rec, ext):
    """
    Arma todos los pares posibles recibo-movimiento con misma fecha e importe
    (TVIR solo contra MP). Regla 1: un par se concilia si ninguno de los dos
    tiene otro candidato. Regla 2: si dentro de una fecha + importe la cantidad
    de recibos es igual a la de movimientos, el grupo cierra y se concilia en
    bloque. El resto se agrupa para revisión manual.
    """
    rec = rec.reset_index(drop=True)
    ext = ext.reset_index(drop=True)
    rec["_r"] = rec.index
    ext["_e"] = ext.index

    pares = rec.merge(ext, on=["Fecha", "Importe"], suffixes=("", "_mov"))
    pares = pares[~(pares.Columna.isin(COLS_MP) & (pares.Fuente != "MP"))]

    cand_r = pares.groupby("_r").size()
    cand_e = pares.groupby("_e").size()
    pares["unico"] = pares._r.map(cand_r).eq(1) & pares._e.map(cand_e).eq(1)

    unicos = pares[pares.unico].copy()
    ambiguos = pares[~pares.unico]

    rec["Estado"] = "Sin movimiento con misma fecha e importe"
    rec.loc[rec._r.isin(ambiguos._r), "Estado"] = "Para revisar"
    rec.loc[rec._r.isin(unicos._r), "Estado"] = "Conciliado"
    ext["Estado"] = "Sin recibo con misma fecha e importe"
    ext.loc[ext._e.isin(ambiguos._e), "Estado"] = "Para revisar"
    ext.loc[ext._e.isin(unicos._e), "Estado"] = "Conciliado"
    unicos["Apareo"] = "Par directo"

    # regla 2: grupos de la misma fecha e importe que empatan en cantidad
    por_ext = ext.groupby(["Fecha", "Importe"])._e.apply(list)
    nuevos = []
    for clave, ridx in rec.groupby(["Fecha", "Importe"])._r.apply(list).items():
        eidx = por_ext.get(clave)
        if eidx is None or len(eidx) != len(ridx):
            continue
        if (rec.loc[ridx, "Estado"] == "Conciliado").any():
            continue  # ya resuelto por la regla 1
        # las videoconsultas solo pueden ir contra MP: tiene que haber al menos
        # tantos movimientos de MP como recibos TVIR
        solo_mp = [i for i in ridx if rec.at[i, "Columna"] in COLS_MP]
        resto = [i for i in ridx if rec.at[i, "Columna"] not in COLS_MP]
        mp = [j for j in eidx if ext.at[j, "Fuente"] == "MP"]
        otros = [j for j in eidx if ext.at[j, "Fuente"] != "MP"]
        if len(solo_mp) > len(mp):
            continue
        izq = solo_mp + resto
        der = mp[:len(solo_mp)] + otros + mp[len(solo_mp):]
        for i, j in zip(izq, der):
            rec.at[i, "Estado"] = "Conciliado"
            ext.at[j, "Estado"] = "Conciliado"
            nuevos.append((i, j, f"Grupo de {len(izq)}"))

    if nuevos:
        ri = [n[0] for n in nuevos]
        ei = [n[1] for n in nuevos]
        izq_df = rec.loc[ri].drop(columns=["Estado"]).reset_index(drop=True)
        der_df = ext.loc[ei].drop(columns=["Estado", "Fecha", "Importe"]).reset_index(drop=True)
        grupo = pd.concat([izq_df, der_df], axis=1)
        grupo["Apareo"] = [n[2] for n in nuevos]
        unicos = pd.concat([unicos, grupo], ignore_index=True)
    return rec, ext, unicos


# ---------------------------------------------------------------- salida
COLS_REC = ["Fecha", "Nro Recibo", "Cliente", "Cajero", "Columna", "Tipo Asiento", "Importe"]
COLS_EXT = ["Fuente", "Fecha", "Importe", "Referencia", "Detalle", "Terminal / Caja"]


def hoja_revisar(rec, ext):
    """Por cada fecha + importe con candidatos ambiguos: todos los recibos y movimientos juntos."""
    r = rec[rec.Estado == "Para revisar"]
    e = ext[ext.Estado == "Para revisar"]
    filas = []
    grupos = sorted(set(zip(r.Fecha, r.Importe)) | set(zip(e.Fecha, e.Importe)))
    for n, (f, imp) in enumerate(grupos, 1):
        gr = r[(r.Fecha == f) & (r.Importe == imp)]
        ge = e[(e.Fecha == f) & (e.Importe == imp)]
        base = {"Grupo": n, "Fecha": f, "Importe": imp,
                "Cant. recibos": len(gr), "Cant. movimientos": len(ge)}
        for _, x in gr.iterrows():
            filas.append({**base, "Lado": "Recibo", "Nro Recibo": x["Nro Recibo"],
                          "Cliente": x.Cliente, "Cajero": x.Cajero, "Columna": x.Columna,
                          "Tipo Asiento": x["Tipo Asiento"]})
        for _, x in ge.iterrows():
            filas.append({**base, "Lado": "Movimiento", "Fuente": x.Fuente,
                          "Referencia": x.Referencia, "Detalle": x.Detalle,
                          "Terminal / Caja": x["Terminal / Caja"]})
    cols = ["Grupo", "Fecha", "Importe", "Cant. recibos", "Cant. movimientos", "Lado",
            "Nro Recibo", "Cliente", "Cajero", "Columna", "Tipo Asiento", "Fuente",
            "Referencia", "Detalle", "Terminal / Caja"]
    df = pd.DataFrame(filas).reindex(columns=cols)
    df["Nro Recibo"] = df["Nro Recibo"].astype("Int64")
    return df


def armar_salida(rec, ext, unicos, desde, hasta, archivos):
    conc = unicos.rename(columns={"Referencia": "Referencia mov", "Detalle": "Detalle mov"})
    conc = conc[["Fecha", "Importe", "Apareo", "Nro Recibo", "Cliente", "Cajero", "Columna",
                 "Tipo Asiento", "Fuente", "Referencia mov", "Detalle mov", "Terminal / Caja"]]
    conc = conc.sort_values(["Fecha", "Nro Recibo"])

    revisar = hoja_revisar(rec, ext)
    rec_sin = rec[rec.Estado.str.startswith("Sin")][COLS_REC].sort_values(["Fecha", "Nro Recibo"])
    ext_sin = ext[ext.Estado.str.startswith("Sin")][COLS_EXT].sort_values(["Fuente", "Fecha"])

    res = [["Período", f"{desde:%d/%m/%Y} al {hasta:%d/%m/%Y}", "", ""],
           ["Procesado", datetime.now().strftime("%d/%m/%Y %H:%M"), "", ""],
           ["Regla 1", "Par directo: misma fecha + mismo importe exacto, único de ambos lados", "", ""],
           ["Regla 2", "Grupo que cierra: misma fecha + mismo importe, igual cantidad de recibos que de movimientos", "", ""],
           ["", "En las dos, TVIR solo se cruza contra Mercado Pago", "", ""],
           ["", "", "", ""],
           ["RECIBOS", "Líneas", "Importe", ""]]
    for col in COLS_TARJETA + COLS_MP:
        g = rec[rec.Columna == col]
        if len(g):
            res.append([col, len(g), round(g.Importe.sum(), 2), ""])
    res.append(["Total", len(rec), round(rec.Importe.sum(), 2), ""])
    res.append(["", "", "", ""])
    res.append(["MOVIMIENTOS", "Cantidad", "Importe", ""])
    for f in ["PAYWAY", "QR", "MP"]:
        g = ext[ext.Fuente == f]
        if len(g):
            res.append([f, len(g), round(g.Importe.sum(), 2), ""])
    res.append(["Total", len(ext), round(ext.Importe.sum(), 2), ""])
    res.append(["", "", "", ""])
    res.append(["RESULTADO", "Recibos", "Importe recibos", "Movimientos"])
    for estado in ["Conciliado", "Para revisar"]:
        g, h = rec[rec.Estado == estado], ext[ext.Estado == estado]
        res.append([estado, len(g), round(g.Importe.sum(), 2), len(h)])
        if estado == "Conciliado":
            for etiqueta, sel in [("   por par directo (regla 1)", unicos.Apareo == "Par directo"),
                                  ("   por grupo que cierra (regla 2)", unicos.Apareo != "Par directo")]:
                g = unicos[sel]
                res.append([etiqueta, len(g), round(g.Importe.sum(), 2), len(g)])
    g = rec[rec.Estado.str.startswith("Sin")]
    res.append(["Recibos sin movimiento", len(g), round(g.Importe.sum(), 2), ""])
    h = ext[ext.Estado.str.startswith("Sin")]
    res.append(["Movimientos sin recibo", "", round(h.Importe.sum(), 2), len(h)])
    res.append(["", "", "", ""])
    res.append(["ARCHIVOS LEÍDOS", "", "", ""])
    for tipo, nombre in archivos:
        res.append([tipo, nombre, "", ""])
    resumen = pd.DataFrame(res)

    return {"Resumen": resumen, "Conciliados": conc, "Para revisar": revisar,
            "Recibos sin movimiento": rec_sin, "Movimientos sin recibo": ext_sin}


def escribir_excel(hojas, destino):
    with pd.ExcelWriter(destino, engine="openpyxl") as w:
        for nombre, df in hojas.items():
            df.to_excel(w, sheet_name=nombre, index=False, header=(nombre != "Resumen"))
    wb = load_workbook(destino)
    azul = PatternFill("solid", fgColor="1F4E78")
    gris = PatternFill("solid", fgColor="EDEDED")
    for ws in wb.worksheets:
        es_resumen = ws.title == "Resumen"
        encabezados = [c.value for c in ws[1]]
        if not es_resumen:
            for cel in ws[1]:
                cel.font = Font(bold=True, color="FFFFFF")
                cel.fill = azul
                cel.alignment = Alignment(wrap_text=True, vertical="center")
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
        for col in ws.columns:
            letra = get_column_letter(col[0].column)
            titulo = str(col[0].value or "")
            largo = max(len(str(c.value)) if c.value is not None else 0 for c in col[:300])
            ws.column_dimensions[letra].width = min(max(largo, len(titulo), 8) + 2, 60)
            for c in (col if es_resumen else col[1:]):
                if isinstance(c.value, datetime):
                    c.number_format = "DD/MM/YYYY"
                elif isinstance(c.value, float):
                    c.number_format = "#,##0.00"
        if es_resumen:
            for row in ws.iter_rows():
                if row[0].value and str(row[0].value).isupper():
                    for c in row:
                        c.font = Font(bold=True)
        if ws.title == "Para revisar" and "Grupo" in encabezados:
            # sombreado alternado por grupo para leerlo más fácil
            for row in ws.iter_rows(min_row=2):
                if row[0].value and row[0].value % 2 == 0:
                    for c in row:
                        c.fill = gris
    wb.save(destino)


# ---------------------------------------------------------------- main
def main():
    ENTRADA.mkdir(exist_ok=True)
    SALIDA.mkdir(exist_ok=True)
    lectores = {"recibos": leer_recibos, "payway": leer_payway, "qr": leer_qr, "mp": leer_mp}
    datos = {k: [] for k in lectores}
    archivos = []

    print(f"Leyendo archivos de {ENTRADA}")
    for p in sorted(ENTRADA.iterdir()):
        if p.name.startswith("~$") or not p.is_file():
            continue
        tipo = detectar(p)
        if tipo is None:
            print(f"  - {p.name}: no reconocido, se ignora")
            continue
        print(f"  - {p.name}: {tipo}")
        archivos.append((tipo, p.name))
        datos[tipo].append(lectores[tipo](p))

    for req in lectores:
        if not datos[req]:
            print(f"  ATENCIÓN: falta el archivo de {req}")
    if not datos["recibos"]:
        sys.exit("No hay archivo de recibos: no se puede conciliar.")

    rec = pd.concat(datos["recibos"]).drop_duplicates()
    ext = pd.concat([d for k in ("payway", "qr", "mp") for d in datos[k]]).drop_duplicates()
    desde, hasta = rec.Fecha.min(), rec.Fecha.max()

    print(f"\nConciliando {len(rec)} líneas de recibos contra {len(ext)} movimientos...")
    rec, ext, unicos = conciliar(rec, ext)
    hojas = armar_salida(rec, ext, unicos, desde, hasta, archivos)

    destino = SALIDA / f"Conciliacion_{desde:%Y-%m-%d}_{hasta:%Y-%m-%d}.xlsx"
    escribir_excel(hojas, destino)

    for estado in ["Conciliado", "Para revisar"]:
        print(f"{estado:<24} recibos: {(rec.Estado == estado).sum():>4}   "
              f"movimientos: {(ext.Estado == estado).sum():>4}")
    print(f"{'Recibos sin movimiento':<24} {len(hojas['Recibos sin movimiento']):>13}")
    print(f"{'Movimientos sin recibo':<24} {len(hojas['Movimientos sin recibo']):>32}")
    print(f"\nResultado: {destino}")


if __name__ == "__main__":
    try:
        main()
    except SystemExit as e:
        print(e)
    finally:
        if getattr(sys, "frozen", False):
            input("\nPresioná Enter para cerrar...")
