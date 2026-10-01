"""
App de Inventario de Costos — versión WEB (Streamlit)
=======================================================
Misma lógica que la app de escritorio, pero para verla en el navegador
y compartirla por link.

CÓMO PROBARLA EN TU LAPTOP:
    pip install streamlit pandas openpyxl
    streamlit run app.py

CÓMO PUBLICARLA GRATIS (Streamlit Community Cloud):
    Ver instrucciones completas al final de este archivo / en el LEEME.
"""

import os
import sqlite3

import pandas as pd
import streamlit as st


# ------------------------------------------------------------------
# Configuración de la página
# ------------------------------------------------------------------
st.set_page_config(page_title="Inventario de Costos — GIIC", layout="wide")

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "inventario.db")


# ------------------------------------------------------------------
# Capa de datos (idéntica a la app de escritorio)
# ------------------------------------------------------------------
SCHEMA = """
CREATE TABLE IF NOT EXISTS lista_costos (
    sku INTEGER PRIMARY KEY,
    nombre TEXT,
    costo_lista REAL,
    categoria TEXT,
    tipo TEXT
);
CREATE TABLE IF NOT EXISTS detalle_lotes (
    sku INTEGER, nombre TEXT, fecha_proceso TEXT,
    cantidad_mayor REAL, cantidad_menor REAL, cantidad REAL,
    costo_unimen REAL, valor_unimen REAL
);
CREATE TABLE IF NOT EXISTS detalle_lotes_existencia (
    sku INTEGER, nombre TEXT, fecha_proceso TEXT,
    cantidad_mayor REAL, cantidad_menor REAL, cantidad REAL,
    costo_unimen REAL, valor_unimen REAL
);
CREATE INDEX IF NOT EXISTS idx_dl_sku ON detalle_lotes(sku);
CREATE INDEX IF NOT EXISTS idx_dle_sku ON detalle_lotes_existencia(sku);
"""


def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.executescript(SCHEMA)
    return conn


GIIC_COLS = {
    "Material": "sku",
    "Material Descripción": "nombre",
    "Fecha Proceso": "fecha_proceso",
    "Cantidad Mayor": "cantidad_mayor",
    "Cantidad Menor": "cantidad_menor",
    "Cantidad": "cantidad",
    "Costo Unimen": "costo_unimen",
    "Valor Unimen": "valor_unimen",
}


def _read_giic_excel(file):
    df = pd.read_excel(file, sheet_name=0)
    faltantes = [c for c in GIIC_COLS if c not in df.columns]
    if faltantes:
        raise ValueError("Faltan columnas: " + ", ".join(faltantes))
    df = df.rename(columns=GIIC_COLS)
    df = df.dropna(subset=["sku"]).copy()
    df["sku"] = df["sku"].astype(float).astype(int)
    df["fecha_proceso"] = pd.to_datetime(df["fecha_proceso"], errors="coerce").dt.strftime("%Y-%m-%d")
    return df[list(GIIC_COLS.values())]


def importar_detalle_lotes(file, existencia=False):
    df = _read_giic_excel(file)
    tabla = "detalle_lotes_existencia" if existencia else "detalle_lotes"
    conn = get_conn()
    try:
        conn.execute(f"DELETE FROM {tabla}")
        df.to_sql(tabla, conn, if_exists="append", index=False)
        conn.commit()
    finally:
        conn.close()
    return len(df), df["sku"].nunique()


def importar_lista_costos(file):
    xls = pd.ExcelFile(file)
    sheet = xls.sheet_names[0]
    for s in xls.sheet_names:
        if "costo" in s.lower():
            sheet = s
            break
    raw = pd.read_excel(file, sheet_name=sheet, header=None)
    cols_lower = [str(c).strip().lower() for c in raw.iloc[0].tolist()] if len(raw) else []

    rows = []
    if {"sku", "nombre"}.issubset(set(cols_lower)):
        df = pd.read_excel(file, sheet_name=sheet)
        df.columns = [str(c).strip() for c in df.columns]
        col_sku = next(c for c in df.columns if c.lower() == "sku")
        col_nombre = next(c for c in df.columns if c.lower() == "nombre")
        col_costo = next((c for c in df.columns if "costo" in c.lower()), None)
        col_cat = next((c for c in df.columns if "categor" in c.lower()), None)
        col_tipo = next((c for c in df.columns if c.lower() == "tipo"), None)
        for _, r in df.iterrows():
            sku = r.get(col_sku)
            if pd.isna(sku):
                continue
            rows.append({
                "sku": int(float(sku)), "nombre": r.get(col_nombre),
                "costo_lista": r.get(col_costo) if col_costo else None,
                "categoria": r.get(col_cat) if col_cat else None,
                "tipo": r.get(col_tipo) if col_tipo else None,
            })
    else:
        tipo = "Primal"
        for i in range(len(raw)):
            a = raw.iat[i, 0] if raw.shape[1] > 0 else None
            if isinstance(a, str) and "procesada" in a.lower():
                tipo = "Procesada"
                continue
            left_sku = raw.iat[i, 0] if raw.shape[1] > 0 else None
            left_name = raw.iat[i, 1] if raw.shape[1] > 1 else None
            e_sku = raw.iat[i, 4] if raw.shape[1] > 4 else None
            e_name = raw.iat[i, 5] if raw.shape[1] > 5 else None
            e_costo = raw.iat[i, 6] if raw.shape[1] > 6 else None
            l_costo = raw.iat[i, 2] if raw.shape[1] > 2 else None
            for sku_v, name_v, costo_v, cat_v in [
                (left_sku, left_name, l_costo, "Engorda"),
                (e_sku, e_name, e_costo, "Media Engorda"),
            ]:
                try:
                    sku_int = int(float(sku_v))
                except (TypeError, ValueError):
                    continue
                if not isinstance(name_v, str):
                    continue
                rows.append({"sku": sku_int, "nombre": name_v, "costo_lista": costo_v,
                              "categoria": cat_v, "tipo": tipo})

    conn = get_conn()
    try:
        conn.execute("DELETE FROM lista_costos")
        conn.executemany(
            "INSERT OR REPLACE INTO lista_costos (sku, nombre, costo_lista, categoria, tipo) "
            "VALUES (:sku, :nombre, :costo_lista, :categoria, :tipo)", rows,
        )
        conn.commit()
    finally:
        conn.close()
    return len(rows)


def _info_sku(conn, sku):
    nombre = None
    for tabla in ("detalle_lotes_existencia", "detalle_lotes"):
        row = conn.execute(
            f"SELECT nombre FROM {tabla} WHERE sku=? ORDER BY fecha_proceso DESC LIMIT 1", (sku,)
        ).fetchone()
        if row and row[0]:
            nombre = row[0]
            break
    if nombre is None:
        row = conn.execute("SELECT nombre FROM lista_costos WHERE sku=?", (sku,)).fetchone()
        if row:
            nombre = row[0]

    en_costos = conn.execute("SELECT COUNT(*) FROM detalle_lotes WHERE sku=?", (sku,)).fetchone()[0] > 0
    en_exist = conn.execute("SELECT COUNT(*) FROM detalle_lotes_existencia WHERE sku=?", (sku,)).fetchone()[0] > 0
    en_lista = conn.execute("SELECT COUNT(*) FROM lista_costos WHERE sku=?", (sku,)).fetchone()[0] > 0
    fuente = "GIIC" if (en_costos or en_exist) else ("Lista de Costos (sin datos GIIC)" if en_lista else "Sin datos")

    r = conn.execute("SELECT SUM(costo_unimen*cantidad), SUM(cantidad) FROM detalle_lotes WHERE sku=?", (sku,)).fetchone()
    costo_hist = (r[0] / r[1]) if r[0] is not None and r[1] else None

    r = conn.execute("SELECT SUM(costo_unimen*cantidad), SUM(cantidad) FROM detalle_lotes_existencia WHERE sku=?", (sku,)).fetchone()
    costo_vivo = (r[0] / r[1]) if r[0] is not None and r[1] else 0.0

    r = conn.execute("SELECT costo_lista FROM lista_costos WHERE sku=?", (sku,)).fetchone()
    costo_lista = r[0] if r else None

    dif = (costo_vivo - costo_lista) if (costo_vivo and costo_lista) else None
    dif_pct = (dif / costo_lista) if (dif is not None and costo_lista) else None

    r = conn.execute("SELECT SUM(cantidad), SUM(valor_unimen), COUNT(*) FROM detalle_lotes WHERE sku=?", (sku,)).fetchone()
    kilos, valor_total, n_lotes = r

    r_min = conn.execute(
        "SELECT costo_unimen, fecha_proceso FROM detalle_lotes WHERE sku=? ORDER BY costo_unimen ASC LIMIT 1", (sku,)
    ).fetchone()
    r_max = conn.execute(
        "SELECT costo_unimen, fecha_proceso FROM detalle_lotes WHERE sku=? ORDER BY costo_unimen DESC LIMIT 1", (sku,)
    ).fetchone()

    r2 = conn.execute("SELECT SUM(cantidad), SUM(valor_unimen), COUNT(*) FROM detalle_lotes_existencia WHERE sku=?", (sku,)).fetchone()
    kilos_exist, valor_exist, n_lotes_exist = r2

    return {
        "SKU": sku, "Nombre": nombre or "", "Fuente": fuente,
        "Costo Ponderado Histórico": costo_hist,
        "Costo Vivo Ponderado (Existencia)": costo_vivo,
        "Costo Lista": costo_lista,
        "Dif. $ (Vivo-Lista)": dif, "Dif. %": dif_pct,
        "Kilos en Existencia": kilos_exist or 0, "Valor en Existencia": valor_exist or 0,
        "# Lotes en Existencia": n_lotes_exist or 0,
        "Kilos Histórico": kilos, "Valor Total Histórico": valor_total, "# Lotes Histórico": n_lotes,
        "Costo Mín.": r_min[0] if r_min else None, "Fecha Costo Mín.": r_min[1] if r_min else None,
        "Costo Máx.": r_max[0] if r_max else None, "Fecha Costo Máx.": r_max[1] if r_max else None,
    }


def buscar_inventario(termino):
    conn = get_conn()
    try:
        termino = (termino or "").strip()
        skus = set()
        if termino == "":
            cur = conn.execute("SELECT DISTINCT sku FROM detalle_lotes UNION SELECT DISTINCT sku FROM lista_costos")
            skus = {r[0] for r in cur.fetchall()}
        else:
            if termino.isdigit():
                for tabla in ("detalle_lotes", "lista_costos"):
                    cur = conn.execute(f"SELECT DISTINCT sku FROM {tabla} WHERE CAST(sku AS TEXT) LIKE ?", (f"%{termino}%",))
                    skus |= {r[0] for r in cur.fetchall()}
            like = f"%{termino}%"
            for tabla in ("detalle_lotes", "detalle_lotes_existencia", "lista_costos"):
                cur = conn.execute(f"SELECT DISTINCT sku FROM {tabla} WHERE nombre LIKE ? COLLATE NOCASE", (like,))
                skus |= {r[0] for r in cur.fetchall()}
        return [_info_sku(conn, sku) for sku in sorted(skus)]
    finally:
        conn.close()


def get_costo_referencia(sku):
    """
    Devuelve (costo, fuente_costo, nombre) para un SKU, usando esta prioridad:
    1) Costo Vivo Ponderado (lo que hay en existencia ahora) si es > 0
    2) Costo Ponderado Histórico (si no hay existencia pero sí historial)
    3) Costo Lista (si no hay nada de GIIC)
    Devuelve (None, None, None) si el SKU no existe en ningún lado.
    """
    conn = get_conn()
    try:
        info = _info_sku(conn, sku)
    finally:
        conn.close()

    if info["Nombre"] == "" and info["Fuente"] == "Sin datos":
        return None, None, None

    if info["Costo Vivo Ponderado (Existencia)"]:
        return info["Costo Vivo Ponderado (Existencia)"], "Vivo (Existencia)", info["Nombre"]
    if info["Costo Ponderado Histórico"]:
        return info["Costo Ponderado Histórico"], "Histórico", info["Nombre"]
    if info["Costo Lista"]:
        return info["Costo Lista"], "Lista", info["Nombre"]
    return None, None, info["Nombre"]


def contar_registros():
    conn = get_conn()
    try:
        n_lista = conn.execute("SELECT COUNT(*) FROM lista_costos").fetchone()[0]
        n_costos = conn.execute("SELECT COUNT(DISTINCT sku) FROM detalle_lotes").fetchone()[0]
        n_exist = conn.execute("SELECT COUNT(DISTINCT sku) FROM detalle_lotes_existencia").fetchone()[0]
        return n_lista, n_costos, n_exist
    finally:
        conn.close()


# ------------------------------------------------------------------
# Interfaz Streamlit
# ------------------------------------------------------------------

st.title("📦 Inventario de Costos — GIIC")

# --- Barra lateral: cargar / actualizar datos (protegido con clave simple) ---
with st.sidebar:
    st.header("Actualizar datos")
    clave = st.text_input("Clave de administrador", type="password")
    try:
        ADMIN_PASS = st.secrets["ADMIN_PASS"]
    except Exception:
        ADMIN_PASS = "cambia-esta-clave"

    if clave == ADMIN_PASS:
        st.success("Modo administrador activo")

        f_lista = st.file_uploader("Lista de Costos (Excel)", type=["xlsx", "xls"], key="f_lista")
        if f_lista is not None and st.button("Cargar Lista de Costos"):
            try:
                n = importar_lista_costos(f_lista)
                st.success(f"Lista de Costos cargada: {n} SKU.")
            except Exception as e:
                st.error(f"Error: {e}")

        f_costos = st.file_uploader("Detalle Lotes - Costos (Excel)", type=["xlsx", "xls"], key="f_costos")
        if f_costos is not None and st.button("Cargar Detalle Lotes (Costos)"):
            try:
                n_filas, n_skus = importar_detalle_lotes(f_costos, existencia=False)
                st.success(f"Cargado: {n_filas} filas, {n_skus} SKU.")
            except Exception as e:
                st.error(f"Error: {e}")

        f_exist = st.file_uploader("Detalle Lotes - Existencia (Excel)", type=["xlsx", "xls"], key="f_exist")
        if f_exist is not None and st.button("Cargar Detalle Lotes Existencia"):
            try:
                n_filas, n_skus = importar_detalle_lotes(f_exist, existencia=True)
                st.success(f"Cargado: {n_filas} filas, {n_skus} SKU.")
            except Exception as e:
                st.error(f"Error: {e}")
    else:
        if clave:
            st.error("Clave incorrecta.")
        st.caption("Solo un administrador puede cargar archivos nuevos. Cualquiera puede buscar sin clave.")

# --- Pestañas: Buscador / Cotizador de Pedido ---
n_lista, n_costos, n_exist = contar_registros()
st.caption(f"Datos cargados: {n_lista} SKU en Lista de Costos · {n_costos} SKU con historial de costos · {n_exist} SKU con existencia actual")

tab_buscar, tab_cotizador = st.tabs(["🔎 Buscador", "🧮 Cotizador de Pedido"])

with tab_buscar:
    termino = st.text_input("Buscar SKU o descripción", placeholder="Ej. 2001 o Chamberete")
    resultados = buscar_inventario(termino)

    if resultados:
        df = pd.DataFrame(resultados)

        df_fmt = df.copy()
        money_cols = ["Costo Ponderado Histórico", "Costo Vivo Ponderado (Existencia)", "Costo Lista",
                      "Dif. $ (Vivo-Lista)", "Valor en Existencia", "Valor Total Histórico", "Costo Mín.", "Costo Máx."]
        for c in money_cols:
            df_fmt[c] = df_fmt[c].map(lambda v: f"${v:,.2f}" if pd.notna(v) else "")
        df_fmt["Dif. %"] = df_fmt["Dif. %"].map(lambda v: f"{v*100:,.1f}%" if pd.notna(v) else "")
        for c in ["Kilos en Existencia", "Kilos Histórico"]:
            df_fmt[c] = df_fmt[c].map(lambda v: f"{v:,.2f}" if pd.notna(v) else "")

        st.dataframe(df_fmt, width='stretch', hide_index=True)
        st.caption(f"{len(resultados)} SKU encontrados.")
    else:
        st.info("No se encontraron resultados. Prueba con otro SKU o descripción.")

with tab_cotizador:
    st.markdown(
        "Agrega los SKU del pedido, su cantidad y el precio al que los quieres vender. "
        "El costo de referencia de cada SKU se toma automáticamente (primero el costo **vivo** de lo que "
        "hay en existencia; si no hay existencia, el **histórico**; si no hay nada de GIIC, el de **Lista**)."
    )

    if "pedido_df" not in st.session_state:
        st.session_state.pedido_df = pd.DataFrame(
            {"SKU": pd.Series(dtype="int"), "Cantidad (Kg)": pd.Series(dtype="float"),
             "Precio de Venta ($/Kg)": pd.Series(dtype="float")}
        )

    pedido_editado = st.data_editor(
        st.session_state.pedido_df,
        num_rows="dynamic",
        width='stretch',
        key="editor_pedido",
        column_config={
            "SKU": st.column_config.NumberColumn(format="%d"),
            "Cantidad (Kg)": st.column_config.NumberColumn(format="%.2f"),
            "Precio de Venta ($/Kg)": st.column_config.NumberColumn(format="%.2f"),
        },
    )
    st.session_state.pedido_df = pedido_editado

    filas_validas = pedido_editado.dropna(subset=["SKU"])
    filas_validas = filas_validas[filas_validas["SKU"] > 0]

    if len(filas_validas) > 0:
        calc_rows = []
        for _, row in filas_validas.iterrows():
            sku = int(row["SKU"])
            cantidad = float(row["Cantidad (Kg)"]) if pd.notna(row["Cantidad (Kg)"]) else 0.0
            precio = float(row["Precio de Venta ($/Kg)"]) if pd.notna(row["Precio de Venta ($/Kg)"]) else 0.0

            costo, fuente_costo, nombre = get_costo_referencia(sku)

            if costo is None:
                calc_rows.append({
                    "SKU": sku, "Nombre": "⚠️ No encontrado", "Costo Ref. ($/Kg)": None,
                    "Fuente Costo": "—", "Cantidad (Kg)": cantidad, "Precio Venta ($/Kg)": precio,
                    "Ganancia Unitaria ($/Kg)": None, "Margen %": None, "Ganancia Total ($)": None,
                    "Venta Total ($)": precio * cantidad,
                })
                continue

            ganancia_unit = precio - costo
            margen_pct = (ganancia_unit / precio) if precio else None
            ganancia_total = ganancia_unit * cantidad
            venta_total = precio * cantidad

            calc_rows.append({
                "SKU": sku, "Nombre": nombre, "Costo Ref. ($/Kg)": costo, "Fuente Costo": fuente_costo,
                "Cantidad (Kg)": cantidad, "Precio Venta ($/Kg)": precio,
                "Ganancia Unitaria ($/Kg)": ganancia_unit, "Margen %": margen_pct,
                "Ganancia Total ($)": ganancia_total, "Venta Total ($)": venta_total,
            })

        df_calc = pd.DataFrame(calc_rows)

        st.subheader("Resultado por SKU")
        df_calc_fmt = df_calc.copy()
        for c in ["Costo Ref. ($/Kg)", "Precio Venta ($/Kg)", "Ganancia Unitaria ($/Kg)",
                  "Ganancia Total ($)", "Venta Total ($)"]:
            df_calc_fmt[c] = df_calc_fmt[c].map(lambda v: f"${v:,.2f}" if pd.notna(v) else "")
        df_calc_fmt["Margen %"] = df_calc_fmt["Margen %"].map(lambda v: f"{v*100:,.1f}%" if pd.notna(v) else "")
        st.dataframe(df_calc_fmt, width='stretch', hide_index=True)

        venta_total_pedido = df_calc["Venta Total ($)"].sum()
        ganancia_total_pedido = df_calc["Ganancia Total ($)"].dropna().sum()
        costo_total_pedido = venta_total_pedido - ganancia_total_pedido
        margen_total_pedido = (ganancia_total_pedido / venta_total_pedido) if venta_total_pedido else 0

        if df_calc["Costo Ref. ($/Kg)"].isna().any():
            st.warning("Algunos SKU no se encontraron en la base — revisa que estén bien escritos. No se incluyen en los totales de ganancia.")

        c1, c2, c3 = st.columns(3)
        c1.metric("Venta Total del Pedido", f"${venta_total_pedido:,.2f}")
        c2.metric("Ganancia Total ($)", f"${ganancia_total_pedido:,.2f}")
        c3.metric("Margen % del Pedido", f"{margen_total_pedido*100:,.1f}%")

        st.divider()

        # --- Prorrateo por margen objetivo ---
        st.subheader("¿Cuánto cobrar para lograr un margen total específico?")
        st.markdown(
            "Pon el margen que quieres ganar en **todo el pedido** y te calculo el precio que le "
            "correspondería a cada SKU para que, en conjunto, el pedido cierre exactamente en ese margen "
            "(se aplica el mismo % de margen a cada SKU sobre su costo de referencia)."
        )
        margen_objetivo = st.number_input("Margen total objetivo (%)", min_value=0.0, max_value=95.0, value=20.0, step=1.0)

        if st.button("Calcular precios sugeridos"):
            filas_con_costo = df_calc[df_calc["Costo Ref. ($/Kg)"].notna()].copy()
            if len(filas_con_costo) == 0:
                st.error("No hay SKU con costo válido para calcular precios sugeridos.")
            else:
                m = margen_objetivo / 100.0
                filas_con_costo["Precio Sugerido ($/Kg)"] = filas_con_costo["Costo Ref. ($/Kg)"] / (1 - m)
                filas_con_costo["Ganancia Sugerida ($/Kg)"] = filas_con_costo["Precio Sugerido ($/Kg)"] - filas_con_costo["Costo Ref. ($/Kg)"]
                filas_con_costo["Ganancia Total Sugerida ($)"] = filas_con_costo["Ganancia Sugerida ($/Kg)"] * filas_con_costo["Cantidad (Kg)"]
                filas_con_costo["Venta Total Sugerida ($)"] = filas_con_costo["Precio Sugerido ($/Kg)"] * filas_con_costo["Cantidad (Kg)"]

                out = filas_con_costo[["SKU", "Nombre", "Costo Ref. ($/Kg)", "Cantidad (Kg)",
                                        "Precio Sugerido ($/Kg)", "Ganancia Total Sugerida ($)", "Venta Total Sugerida ($)"]].copy()
                out_fmt = out.copy()
                for c in ["Costo Ref. ($/Kg)", "Precio Sugerido ($/Kg)", "Ganancia Total Sugerida ($)", "Venta Total Sugerida ($)"]:
                    out_fmt[c] = out_fmt[c].map(lambda v: f"${v:,.2f}")
                st.dataframe(out_fmt, width='stretch', hide_index=True)

                venta_sug = out["Venta Total Sugerida ($)"].sum()
                ganancia_sug = out["Ganancia Total Sugerida ($)"].sum()
                st.success(
                    f"Con estos precios, el pedido vendería **${venta_sug:,.2f}**, ganarías "
                    f"**${ganancia_sug:,.2f}**, exactamente **{margen_objetivo:,.1f}%** de margen total."
                )
    else:
        st.info("Agrega al menos un SKU en la tabla de arriba (clic en el + al final de la tabla) para ver el cálculo.")


# ====================================================================
# CÓMO PUBLICAR ESTA APP GRATIS (Streamlit Community Cloud)
# ====================================================================
#
# 1. Crea una cuenta gratis en https://github.com (si no tienes) y en
#    https://streamlit.io/cloud (puedes entrar con tu cuenta de GitHub).
#
# 2. Sube estos 2 archivos a un repositorio nuevo en GitHub:
#       - app.py               (este archivo)
#       - requirements.txt     (con: streamlit / pandas / openpyxl)
#
# 3. En Streamlit Cloud: "New app" > selecciona tu repositorio >
#    archivo principal "app.py" > Deploy.
#
# 4. En unos minutos te da un link tipo:
#       https://tu-app.streamlit.app
#    Ese es el link que compartes con tu equipo — solo lo abren y
#    buscan, no necesitan instalar nada.
#
# 5. Clave de administrador: en Streamlit Cloud, ve a la configuración
#    de tu app > "Secrets" y agrega:
#       ADMIN_PASS = "la-clave-que-tú-quieras"
#    Esa es la clave que tú vas a escribir en la barra lateral cada vez
#    que quieras cargar archivos nuevos. Sin esa clave, la gente solo
#    puede buscar, no puede cargar/cambiar los datos.
#
# IMPORTANTE sobre los datos: en el plan gratis de Streamlit Cloud, el
# archivo inventario.db puede borrarse si la app se "duerme" por
# inactividad prolongada o si hay un redeploy. Es decir: sirve muy bien
# para que la gente busque, pero tú como administrador debes volver a
# cargar tus 3 archivos cada vez que notes que los datos desaparecieron
# (o revisa las opciones de almacenamiento persistente / base de datos
# externa de Streamlit si quieres que nunca se borre).
# ====================================================================
