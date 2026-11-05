import hmac
import io
import re
import unicodedata
import zipfile

import pandas as pd
import streamlit as st
import xlrd
from openpyxl import load_workbook
from xlutils.copy import copy as xl_copy

st.set_page_config(page_title="Completar facturas", layout="wide")


def comprobar_password():
    """Pide contraseña (guardada en st.secrets['APP_PASSWORD'], nunca en el código)."""
    if st.session_state.get("autenticado"):
        return True
    try:
        esperada = str(st.secrets["APP_PASSWORD"])
    except Exception:
        st.error("Falta configurar APP_PASSWORD en los secrets de la app.")
        return False
    st.title("Acceso")
    pwd = st.text_input("Contraseña", type="password")
    if pwd:
        if hmac.compare_digest(pwd.encode("utf-8"), esperada.encode("utf-8")):
            st.session_state["autenticado"] = True
            st.rerun()
        else:
            st.error("Contraseña incorrecta")
    return False


if not comprobar_password():
    st.stop()

# Posiciones (base 0) de las columnas en el fichero de pedidos
COL_PEDIDO_A = 0    # A
COL_PEDIDO_I = 8    # I
COL_CENTRO = 20     # U
COL_CLIENTE = 23    # X
COL_SERVICIO = 50   # AY

# Columnas (base 1) de destino en las facturas
FACT_COL_PEDIDO = 3  # C
FACT_COL_CENTRO = 6  # F
FACT_COL_CIF = 7     # G
FACT_COL_SERV = 8    # H


# ---------- utilidades ----------
def vacio(v):
    return v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() == ""


def limpiar(v):
    """Valor listo para escribir en Excel (sin NaN, enteros sin .0)."""
    if vacio(v):
        return ""
    if isinstance(v, float) and v.is_integer():
        return int(v)
    return v.strip() if isinstance(v, str) else v


def norm_pedido(v):
    if vacio(v):
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v).strip().upper()
    if s.isdigit():
        s = s.lstrip("0") or "0"
    return s


def norm_texto(v):
    if vacio(v):
        return ""
    s = unicodedata.normalize("NFKD", str(v))
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", s).strip().upper()


# ---------- carga de datos ----------
@st.cache_data(show_spinner=False)
def leer_excel(data: bytes, header):
    return pd.read_excel(io.BytesIO(data), header=header, dtype=object)


def construir_pedidos(df):
    if df.shape[1] <= COL_SERVICIO:
        raise ValueError(
            f"El fichero de pedidos tiene {df.shape[1]} columnas; "
            f"se necesitan al menos {COL_SERVICIO + 1} (hasta la AY)."
        )
    pedidos, duplicados = {}, set()
    for fila in df.itertuples(index=False, name=None):
        info = (limpiar(fila[COL_CENTRO]), limpiar(fila[COL_CLIENTE]), limpiar(fila[COL_SERVICIO]))
        for col in (COL_PEDIDO_A, COL_PEDIDO_I):
            clave = norm_pedido(fila[col])
            if not clave:
                continue
            if clave in pedidos and pedidos[clave] != info:
                duplicados.add(clave)
            pedidos.setdefault(clave, info)  # se queda con la primera aparición
    return pedidos, duplicados


def construir_cifs(df, col_cliente, col_cif):
    return {
        norm_texto(c): limpiar(cif)
        for c, cif in zip(df[col_cliente], df[col_cif])
        if not vacio(c) and not vacio(cif)
    }


# ---------- proceso de facturas ----------
def procesar_factura(nombre, data, pedidos, cifs, fila_cab):
    """fila_cab: fila (base 1) de la cabecera; los datos empiezan en la siguiente."""
    es_xls = nombre.lower().endswith(".xls")

    if es_xls:
        # .xls: se copia el libro con xlutils para conservar el formato original
        rb = xlrd.open_workbook(file_contents=data, formatting_info=True)
        rs = rb.sheet_by_index(0)
        wb = xl_copy(rb)
        ws = wb.get_sheet(0)
        n_filas = rs.nrows

        def leer(r):
            return rs.cell_value(r - 1, FACT_COL_PEDIDO - 1) if rs.ncols >= FACT_COL_PEDIDO else None

        def escribir(r, c, v):
            ws.write(r - 1, c - 1, v)
    else:
        wb = load_workbook(io.BytesIO(data), keep_vba=nombre.lower().endswith(".xlsm"))
        ws = wb.active
        n_filas = ws.max_row

        def leer(r):
            return ws.cell(r, FACT_COL_PEDIDO).value

        def escribir(r, c, v):
            ws.cell(r, c, v)

    if fila_cab > 0:
        escribir(fila_cab, FACT_COL_CENTRO, "Centro de coste")
        escribir(fila_cab, FACT_COL_CIF, "CIF cliente")
        escribir(fila_cab, FACT_COL_SERV, "Servicio")

    total = encontrados = 0
    sin_pedido, sin_cif = [], set()

    for r in range(fila_cab + 1, n_filas + 1):
        valor = leer(r)
        clave = norm_pedido(valor)
        if not clave:
            continue
        total += 1
        info = pedidos.get(clave)
        if info is None:
            sin_pedido.append(str(valor))
            continue
        centro, cliente, servicio = info
        cif = cifs.get(norm_texto(cliente), "")
        if not cif and not vacio(cliente):
            sin_cif.add(str(cliente))
        escribir(r, FACT_COL_CENTRO, centro)
        escribir(r, FACT_COL_CIF, cif)
        escribir(r, FACT_COL_SERV, servicio)
        encontrados += 1

    salida = io.BytesIO()
    wb.save(salida)
    return salida.getvalue(), {
        "Fichero": nombre,
        "Filas con pedido": total,
        "Completadas": encontrados,
        "Pedidos no encontrados": len(sin_pedido),
        "Clientes sin CIF": len(sin_cif),
        "_sin_pedido": sin_pedido,
        "_sin_cif": sorted(sin_cif),
    }


# ---------- interfaz ----------
st.title("Completar facturas con datos de pedidos")

c1, c2 = st.columns(2)
with c1:
    f_pedidos = st.file_uploader("1. Fichero de pedidos", type=["xlsx", "xlsm", "xls"])
with c2:
    f_cifs = st.file_uploader("2. Relación cliente - CIF", type=["xlsx", "xlsm", "xls"])

f_facturas = st.file_uploader(
    "3. Ficheros de facturas (.xls / .xlsx / .xlsm)", type=["xls", "xlsx", "xlsm"], accept_multiple_files=True
)

col_cli = col_cif = None
if f_cifs:
    df_cif = leer_excel(f_cifs.getvalue(), 0)
    cols = list(df_cif.columns)
    if len(cols) < 2:
        st.error("El fichero cliente-CIF necesita al menos dos columnas.")
        st.stop()
    a, b = st.columns(2)
    col_cli = a.selectbox("Columna con el nombre del cliente", cols, index=0)
    col_cif = b.selectbox("Columna con el CIF", cols, index=1)

fila_cab = st.number_input(
    "Fila de cabecera en las facturas (los datos empiezan en la siguiente)", min_value=0, value=4, step=1
)

if st.button("Procesar", type="primary", disabled=not (f_pedidos and f_cifs and f_facturas)):
    try:
        df_ped = leer_excel(f_pedidos.getvalue(), None)  # sin cabecera: se usan posiciones
        pedidos, duplicados = construir_pedidos(df_ped)
        cifs = construir_cifs(df_cif, col_cli, col_cif)

        resultados, resumen = {}, []
        barra = st.progress(0.0)
        for i, f in enumerate(f_facturas, 1):
            data, stats = procesar_factura(f.name, f.getvalue(), pedidos, cifs, int(fila_cab))
            resultados[f.name] = data
            resumen.append(stats)
            barra.progress(i / len(f_facturas))

        st.session_state["res"] = {
            "archivos": resultados,
            "resumen": resumen,
            "duplicados": sorted(duplicados),
        }
    except Exception as e:
        st.session_state.pop("res", None)
        st.error(f"Error al procesar: {e}")

res = st.session_state.get("res")
if res:
    st.subheader("Resultado")
    st.dataframe(
        pd.DataFrame([{k: v for k, v in r.items() if not k.startswith("_")} for r in res["resumen"]]),
        use_container_width=True,
        hide_index=True,
    )

    if res["duplicados"]:
        st.warning(
            "Pedidos repetidos en el fichero de pedidos con datos distintos "
            f"(se usó la primera aparición): {', '.join(res['duplicados'][:30])}"
        )

    for r in res["resumen"]:
        if r["_sin_pedido"] or r["_sin_cif"]:
            with st.expander(f"Incidencias: {r['Fichero']}"):
                if r["_sin_pedido"]:
                    st.write("Pedidos no encontrados:", ", ".join(r["_sin_pedido"]))
                if r["_sin_cif"]:
                    st.write("Clientes sin CIF:", ", ".join(r["_sin_cif"]))

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for nombre, data in res["archivos"].items():
            z.writestr(nombre, data)
    st.download_button(
        "Descargar todo (ZIP)", buf.getvalue(), "facturas_completadas.zip", "application/zip"
    )

    st.write("Descarga individual:")
    for nombre, data in res["archivos"].items():
        st.download_button(
            f"⬇ {nombre}",
            data,
            nombre,
            "application/vnd.ms-excel" if nombre.lower().endswith(".xls")
            else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"dl_{nombre}",
        )
