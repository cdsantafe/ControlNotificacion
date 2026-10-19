"""
Dashboard de Control - TAREA CRÍTICA: NOTIFICACIÓN AL CLIENTE (BEES Deliver)
Ejecutar local:  streamlit run app.py
"""
from io import BytesIO
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

# ----------------------------------------------------------------------
# Configuración general
# ----------------------------------------------------------------------
OBJETIVO = 88.0                      # % objetivo de notificación > 2 min
ARCHIVO_DEFECTO = "routes-data_2026-10-19_14h1m.csv"
COLUMNAS_REQ = ["driver_name", "driver_notification_time", "arrived_at"]

st.set_page_config(page_title="Notificación al Cliente", page_icon="📲", layout="wide")

# ----------------------------------------------------------------------
# CSS personalizado (imita la planilla de control)
# ----------------------------------------------------------------------
st.markdown(
    """
    <style>
    .titulo {color:#D32F2F; font-size:2.4rem; font-weight:800; margin-bottom:0; letter-spacing:.5px;}
    .subtitulo {color:#444; font-size:1.2rem; font-weight:600; margin-top:0;}
    .resumen {width:100%; border-collapse:collapse; margin:10px 0 20px 0;}
    .resumen th {background:#FFC933; color:#222; padding:10px; text-align:center;
                 border:1px solid #d9a800; font-size:.95rem;}
    .resumen td {padding:12px; text-align:center; font-size:1.5rem; font-weight:700;
                 border:1px solid #ddd;}
    .ok  {background:#C8E6C9; color:#1B5E20;}
    .bad {background:#FFCDD2; color:#B71C1C;}
    .neu {background:#FFF8E1; color:#333;}
    </style>
    """,
    unsafe_allow_html=True,
)

# ----------------------------------------------------------------------
# Carga y preparación de datos
# ----------------------------------------------------------------------
def leer_archivo(origen) -> pd.DataFrame:
    """Lee CSV o Excel desde un uploader o una ruta."""
    nombre = getattr(origen, "name", str(origen)).lower()
    if nombre.endswith((".xlsx", ".xls")):
        return pd.read_excel(origen)
    return pd.read_csv(origen, sep=None, engine="python")  # detecta separador


def datos_demo() -> pd.DataFrame:
    """Datos de demostración cuando no hay archivo."""
    rng = np.random.default_rng(7)
    choferes = ["Chofer A", "Chofer B", "Chofer C", "Chofer D", "Chofer E"]
    filas = []
    for dia in pd.date_range("2026-10-12", periods=7):
        for ch in choferes:
            for _ in range(rng.integers(15, 30)):
                llegada = dia + pd.Timedelta(hours=int(rng.integers(8, 17)), minutes=int(rng.integers(0, 60)))
                anticipo = max(rng.normal(240 if ch < "D" else 100, 90), 0)
                filas.append({
                    "driver_name": ch,
                    "tour_display_id": f"P{choferes.index(ch)+1:03d}",
                    "tour_date": dia.date().isoformat(),
                    "driver_notification_time": (llegada - pd.Timedelta(seconds=anticipo)).isoformat() + "Z",
                    "arrived_at": llegada.isoformat() + "Z" if rng.random() > 0.05 else None,
                })
    return pd.DataFrame(filas)


def preparar(df: pd.DataFrame) -> pd.DataFrame:
    """Calcula anticipación y banderas de cumplimiento."""
    faltan = [c for c in COLUMNAS_REQ if c not in df.columns]
    if faltan:
        st.error(f"Faltan columnas en el archivo: {', '.join(faltan)}")
        st.stop()

    df = df.copy()
    df["arrived_at"] = pd.to_datetime(df["arrived_at"], utc=True, errors="coerce")
    df["driver_notification_time"] = pd.to_datetime(df["driver_notification_time"], utc=True, errors="coerce")

    # Fecha de la ruta: tour_date > trip_start_timestamp > fecha de llegada
    if "tour_date" in df.columns:
        fecha = pd.to_datetime(df["tour_date"], errors="coerce")
    elif "trip_start_timestamp" in df.columns:
        fecha = pd.to_datetime(df["trip_start_timestamp"], utc=True, errors="coerce").dt.tz_localize(None)
    else:
        fecha = df["arrived_at"].dt.tz_localize(None)
    df["fecha"] = fecha.dt.normalize()

    if "tour_display_id" not in df.columns:
        df["tour_display_id"] = "N/D"

    # Regla 1: excluir visitas sin llegada
    df = df[df["arrived_at"].notna() & df["fecha"].notna()].copy()

    # Reglas 2 y 3: anticipación en segundos (sin notificación => no cumple)
    df["anticipacion_seg"] = (df["arrived_at"] - df["driver_notification_time"]).dt.total_seconds()
    df["cumple_2"] = df["anticipacion_seg"] >= 120
    df["cumple_5"] = df["anticipacion_seg"] >= 300
    return df


def asignar_periodo(df: pd.DataFrame, frecuencia: str) -> pd.DataFrame:
    """Agrega la columna 'periodo' según la frecuencia elegida."""
    df = df.copy()
    if frecuencia == "Diario":
        df["periodo"] = df["fecha"].dt.strftime("%Y-%m-%d")
    elif frecuencia == "Semanal":
        iso = df["fecha"].dt.isocalendar()
        df["periodo"] = iso["year"].astype(str) + "-S" + iso["week"].astype(str).str.zfill(2)
    else:
        df["periodo"] = df["fecha"].dt.strftime("%Y-%m")
    return df


def pct(serie: pd.Series) -> float:
    return float(serie.mean() * 100) if len(serie) else np.nan


# ----------------------------------------------------------------------
# Estilos de tabla y tarjetas
# ----------------------------------------------------------------------
def color_celda(v):
    if pd.isna(v):
        return ""
    return "background-color:#C8E6C9;color:#1B5E20" if v >= OBJETIVO else "background-color:#FFCDD2;color:#B71C1C"


def estilizar(tabla: pd.DataFrame):
    cols = ["Notificaciones mayores a 2 minutos", "Notificaciones mayores a 5 minutos"]
    estilo = tabla.style.format({c: "{:.0f}%" for c in cols})
    aplicar = estilo.map if hasattr(estilo, "map") else estilo.applymap
    return aplicar(color_celda, subset=cols)


def clase(v):
    return "neu" if pd.isna(v) else ("ok" if v >= OBJETIVO else "bad")


def a_excel(tabla: pd.DataFrame) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        tabla.to_excel(w, index=False, sheet_name="Notificación")
    return buf.getvalue()


# ----------------------------------------------------------------------
# Encabezado
# ----------------------------------------------------------------------
st.markdown('<p class="titulo">TAREA CRÍTICA: NOTIFICACIÓN AL CLIENTE</p>', unsafe_allow_html=True)
st.markdown(f'<p class="subtitulo">Objetivo: {OBJETIVO:.0f}% notificado con más de 2 minutos</p>',
            unsafe_allow_html=True)

# ----------------------------------------------------------------------
# Carga de archivo (principal + sidebar; ambos comparten la misma fuente)
# ----------------------------------------------------------------------
sub_sidebar = st.sidebar.file_uploader("📂 Cargar CSV/Excel", type=["csv", "xlsx", "xls"], key="up_side")
sub_main = st.file_uploader("📂 Cargar o actualizar archivo de rutas (CSV/Excel)",
                            type=["csv", "xlsx", "xls"], key="up_main")
archivo = sub_main if sub_main is not None else sub_sidebar

if archivo is not None:
    bruto = leer_archivo(archivo)
    origen = f"Archivo cargado: {archivo.name}"
elif Path(ARCHIVO_DEFECTO).exists():
    bruto = leer_archivo(ARCHIVO_DEFECTO)
    origen = f"Archivo local: {ARCHIVO_DEFECTO}"
else:
    bruto = datos_demo()
    origen = "Datos de demostración (cargá un archivo para ver datos reales)"
st.caption(origen)

datos = preparar(bruto)
if datos.empty:
    st.warning("No hay visitas con registro de llegada para evaluar.")
    st.stop()

# ----------------------------------------------------------------------
# Filtros (sidebar)
# ----------------------------------------------------------------------
st.sidebar.header("Filtros")
frecuencia = st.sidebar.radio("Frecuencia", ["Diario", "Semanal", "Mensual"], horizontal=True)

f_min, f_max = datos["fecha"].min().date(), datos["fecha"].max().date()
rango = st.sidebar.date_input("Fecha / rango", value=(f_min, f_max), min_value=f_min, max_value=f_max)
if isinstance(rango, tuple) and len(rango) == 2:
    desde, hasta = rango
else:
    desde = hasta = rango[0] if isinstance(rango, tuple) else rango

choferes = sorted(datos["driver_name"].dropna().unique())
sel_ch = st.sidebar.multiselect("Chofer", choferes, default=choferes)
planillas = sorted(datos["tour_display_id"].dropna().astype(str).unique())
sel_pl = st.sidebar.multiselect("Planilla / Ruta", planillas, default=planillas)

mask = (
    (datos["fecha"].dt.date >= desde) & (datos["fecha"].dt.date <= hasta)
    & datos["driver_name"].isin(sel_ch)
    & datos["tour_display_id"].astype(str).isin(sel_pl)
)
d = asignar_periodo(datos[mask], frecuencia)
if d.empty:
    st.info("Sin datos para los filtros seleccionados.")
    st.stop()

# ----------------------------------------------------------------------
# Tarjeta consolidada
# ----------------------------------------------------------------------
g2, g5 = pct(d["cumple_2"]), pct(d["cumple_5"])
etiqueta = f"{desde:%d/%m/%Y}" if desde == hasta else f"{desde:%d/%m/%Y} - {hasta:%d/%m/%Y}"
st.markdown(
    f"""
    <table class="resumen">
      <tr><th>FECHA</th><th>% general mayor a 2 minutos</th><th>% general mayor a 5 minutos</th></tr>
      <tr><td class="neu">{etiqueta}</td>
          <td class="{clase(g2)}">{g2:.0f}%</td>
          <td class="{clase(g5)}">{g5:.0f}%</td></tr>
    </table>
    """,
    unsafe_allow_html=True,
)

# ----------------------------------------------------------------------
# Tabla por chofer
# ----------------------------------------------------------------------
tabla = (
    d.groupby("driver_name")
    .agg(Clientes=("arrived_at", "size"), c2=("cumple_2", "mean"), c5=("cumple_5", "mean"))
    .reset_index()
    .rename(columns={"driver_name": "CHOFER",
                     "c2": "Notificaciones mayores a 2 minutos",
                     "c5": "Notificaciones mayores a 5 minutos"})
)
tabla["Notificaciones mayores a 2 minutos"] *= 100
tabla["Notificaciones mayores a 5 minutos"] *= 100
tabla = tabla.sort_values("Notificaciones mayores a 2 minutos", ascending=False).reset_index(drop=True)

tab1, tab2, tab3 = st.tabs(["📋 Por chofer", "📈 Evolución", "⚠️ Visitas < 2 min"])

with tab1:
    st.dataframe(estilizar(tabla), use_container_width=True, hide_index=True)
    st.download_button("⬇️ Descargar Excel", a_excel(tabla.round(1)), "notificacion_por_chofer.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

with tab2:
    evo = (d.groupby("periodo")
             .agg(**{"% > 2 min": ("cumple_2", "mean"), "% > 5 min": ("cumple_5", "mean")})
             .mul(100).reset_index())
    fig = px.line(evo.melt("periodo", var_name="Métrica", value_name="%"),
                  x="periodo", y="%", color="Métrica", markers=True,
                  color_discrete_sequence=["#D32F2F", "#F9A825"])
    fig.add_hline(y=OBJETIVO, line_dash="dash", line_color="green", annotation_text=f"Objetivo {OBJETIVO:.0f}%")
    fig.update_layout(yaxis_range=[0, 105], xaxis_title=frecuencia, yaxis_title="% cumplimiento")
    st.plotly_chart(fig, use_container_width=True)

with tab3:
    inc = d[~d["cumple_2"]].copy()
    inc["Anticipación (seg)"] = inc["anticipacion_seg"].round(0)
    inc["arrived_at"] = inc["arrived_at"].dt.strftime("%Y-%m-%d %H:%M:%S")
    vista = (inc[["fecha", "driver_name", "tour_display_id", "arrived_at", "Anticipación (seg)"]]
             .rename(columns={"fecha": "Fecha", "driver_name": "Chofer",
                              "tour_display_id": "Planilla", "arrived_at": "Llegada"})
             .sort_values("Anticipación (seg)", na_position="first"))
    vista["Fecha"] = vista["Fecha"].dt.strftime("%Y-%m-%d")
    st.caption(f"{len(vista)} visitas sin cumplir (vacío = sin notificación registrada).")
    st.dataframe(vista, use_container_width=True, hide_index=True)
