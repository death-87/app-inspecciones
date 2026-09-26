"""Página adicional: ejecutar la aplicación principal app.py, no este archivo."""
import datetime as dt
import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="Indicadores | QA/QC", page_icon="📊", layout="wide")
if not st.session_state.get("autenticado", False) or st.session_state.get("usuario_actual") != "jnavarrete":
    st.info("Acceso privado. Abre la página principal e inicia sesión con jnavarrete.")
    st.stop()

INSPECTORES = ["Juan Navarrete", "Jorge Hernandez", "Arlem Sarmiento", "Harold Castillo", "Miguel Chirinos"]
ESTADOS = ["Conforme", "Con Hallazgos", "Rechazado"]
COLORES = {"Conforme": "#087F73", "Con Hallazgos": "#B77713", "Rechazado": "#BD4050"}
COLUMNAS = ["Fecha", "Semana", "Inspector", "Equipo_TAG", "Horas", "Estado", "Hallazgos"]


def preparar_datos(datos):
    df = datos.copy()
    df["Fecha"] = pd.to_datetime(df["Fecha"], errors="raise").dt.normalize()
    iso = df["Fecha"].dt.isocalendar()
    df["Semana"] = iso.week.astype(int)
    df["Año ISO"] = iso.year.astype(int)
    df["Periodo"] = df["Año ISO"].astype(str) + " · S" + df["Semana"].astype(str).str.zfill(2)
    df["Equipo_TAG"] = df["Equipo_TAG"].astype(str).str.strip().str.upper()
    for col in ["Horas", "Hallazgos"]:
        df[col] = pd.to_numeric(df[col], errors="raise")
    if df[COLUMNAS].isna().any().any() or df["Equipo_TAG"].eq("").any():
        raise ValueError("Hay campos vacíos en los registros.")
    if not df["Estado"].isin(ESTADOS).all():
        raise ValueError("Hay estados no reconocidos.")
    if not df["Horas"].between(0.5, 24).all():
        raise ValueError("Las horas deben estar entre 0,5 y 24 por registro.")
    if not ((df["Hallazgos"] >= 0) & (df["Hallazgos"] % 1 == 0)).all():
        raise ValueError("Los hallazgos deben ser enteros no negativos.")
    return df


def indicadores(df):
    n = len(df)
    return {
        "registros": n,
        "equipos": df["Equipo_TAG"].nunique(),
        "horas": float(df["Horas"].sum()),
        "promedio": float(df["Horas"].sum()) / n if n else None,
        "rechazo": 100 * df["Estado"].eq("Rechazado").sum() / n if n else None,
        "hallazgos": int(df["Hallazgos"].sum()),
    }


def estilo(fig, height=310):
    fig.update_layout(
        template="plotly_white", height=height,
        margin=dict(l=0, r=25, t=15, b=5),
        font=dict(family="Arial", size=12, color="#17324D"),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        legend=dict(title_text="", orientation="h", y=-0.22),
    )
    fig.update_xaxes(showgrid=True, gridcolor="#E2E8F0", zeroline=False)
    fig.update_yaxes(showgrid=False, zeroline=False)
    return fig


st.markdown("""<style>
.block-container { max-width: 1280px; padding-top: 2.5rem; padding-bottom: 1rem; }
[data-testid="stVerticalBlock"] { gap: .7rem; }
[data-testid="stMetric"] { background: white; border: 1px solid #dbe3eb; border-radius: 10px; padding: .7rem; }
[data-testid="stMetricValue"] { font-size: 1.65rem; }
h1 { font-size: 1.85rem !important; color: #17324d; }
h3 { font-size: 1.05rem !important; }
[data-testid="stForm"] { background: white; border-radius: 10px; }
@media (max-width: 640px) {
 [data-testid="stHorizontalBlock"] { flex-wrap: wrap; }
 [data-testid="stColumn"] { min-width: 100% !important; }
}
</style>""", unsafe_allow_html=True)
st.title("Indicadores de inspección")
st.caption("QA/QC · Actividad, dedicación y resultados de inspección")
st.warning("Modo temporal: estos registros no están conectados a Google Sheets. Descarga el CSV antes de cerrar o recargar la sesión.")

if "inspecciones_data" not in st.session_state:
    st.session_state.inspecciones_data = pd.DataFrame(columns=COLUMNAS)
elif not st.session_state.get("kpi_origen_revisado", False):
    st.info("Se encontraron datos de la sesión anterior. Si proceden del ejemplo original, son datos de demostración y no actividad real.")
st.session_state.kpi_origen_revisado = True
if "kpi_aviso" in st.session_state:
    st.success(st.session_state.pop("kpi_aviso"))

with st.expander("Registrar actividad", expanded=False):
    with st.form("kpi_registro", clear_on_submit=False):
        a, b, c = st.columns(3)
        fecha = a.date_input("Fecha", dt.date.today())
        inspector = b.selectbox("Inspector", INSPECTORES)
        equipo = c.text_input("TAG del equipo", placeholder="Ej: C701")
        d, e, f = st.columns(3)
        horas = d.number_input("Horas dedicadas", 0.5, 24.0, 4.0, 0.5)
        estado = e.selectbox("Resultado", ESTADOS)
        hallazgos = f.number_input("Hallazgos", min_value=0, value=0, step=1)
        st.caption("La semana y el año ISO se calculan a partir de la fecha. Un registro representa una actividad de un inspector.")
        enviado = st.form_submit_button("Registrar en esta sesión", type="primary")
    if enviado:
        tag = equipo.strip().upper()
        if not tag:
            st.error("Introduce el TAG del equipo.")
        elif estado == "Conforme" and hallazgos > 0:
            st.error("Un resultado Conforme debe tener cero hallazgos.")
        elif estado == "Con Hallazgos" and hallazgos == 0:
            st.error("Indica al menos un hallazgo para ese resultado.")
        else:
            fila = dict(Fecha=str(fecha), Semana=fecha.isocalendar().week,
                        Inspector=inspector, Equipo_TAG=tag, Horas=float(horas),
                        Estado=estado, Hallazgos=int(hallazgos))
            anterior = st.session_state.inspecciones_data
            nuevo = pd.DataFrame([fila])
            # Comparación normalizada: evita duplicados exactos por doble envío.
            comparacion = preparar_datos(pd.concat([anterior, nuevo], ignore_index=True))
            if comparacion.duplicated(subset=COLUMNAS).iloc[-1]:
                st.error("Este registro ya existe. Revisa la tabla antes de volver a guardarlo.")
            else:
                st.session_state.inspecciones_data = pd.concat([anterior, nuevo], ignore_index=True)
                st.session_state.kpi_aviso = "Actividad registrada en esta sesión. Ajusta los filtros si no aparece."
                st.rerun()

try:
    df = preparar_datos(st.session_state.inspecciones_data)
except (ValueError, KeyError, TypeError) as exc:
    st.error(f"No se pueden calcular los indicadores: {exc}")
    st.stop()

with st.sidebar:
    st.subheader("Filtros de indicadores")
    periodos = sorted(df["Periodo"].unique(), reverse=True)
    seleccion_periodos = st.multiselect("Año y semana ISO", periodos, default=periodos)
    nombres = sorted(set(INSPECTORES) | set(df["Inspector"].unique()))
    seleccion_inspectores = st.multiselect("Inspectores", nombres, default=nombres)
    seleccion_estados = st.multiselect("Resultados", ESTADOS, default=ESTADOS)
    buscar_tag = st.text_input("Buscar TAG", placeholder="Todo o parte del TAG").strip().upper()
    st.caption("Una selección vacía no muestra registros.")

filtrado = df[
    df["Periodo"].isin(seleccion_periodos)
    & df["Inspector"].isin(seleccion_inspectores)
    & df["Estado"].isin(seleccion_estados)
    & df["Equipo_TAG"].str.contains(buscar_tag, regex=False, na=False)
].copy()

if filtrado.empty:
    st.info("No hay registros para esta selección. Registra una actividad o ajusta los filtros.")
    st.stop()

k = indicadores(filtrado)
cols = st.columns(4)
cols[0].metric("Actividades registradas", k["registros"])
cols[1].metric("Equipos únicos", k["equipos"])
cols[2].metric("Horas dedicadas", f'{k["horas"]:.1f} h')
cols[3].metric("Rechazadas / registros", f'{k["rechazo"]:.1f}%')
st.caption(f'Promedio: {k["promedio"]:.2f} h por registro · Hallazgos reportados: {k["hallazgos"]} · Todos los indicadores respetan los filtros, incluido Resultado.')

resumen = filtrado.groupby("Inspector").agg(
    Registros=("Equipo_TAG", "size"), Horas=("Horas", "sum")
).reindex(seleccion_inspectores, fill_value=0).reset_index()
izq, der = st.columns(2)
with izq:
    st.subheader("Actividad por inspector")
    fig = px.bar(resumen.sort_values("Registros"), x="Registros", y="Inspector", orientation="h",
                 text="Registros", color_discrete_sequence=["#087F73"], hover_data={"Horas": ":.1f"})
    fig.update_traces(textposition="outside", cliponaxis=False)
    fig.update_xaxes(dtick=1, rangemode="tozero")
    st.plotly_chart(estilo(fig), use_container_width=True)
with der:
    st.subheader("Horas por inspector")
    fig = px.bar(resumen.sort_values("Horas"), x="Horas", y="Inspector", orientation="h",
                 text="Horas", color_discrete_sequence=["#355C83"], hover_data=["Registros"])
    fig.update_traces(texttemplate="%{x:.1f} h", textposition="outside", cliponaxis=False)
    fig.update_xaxes(rangemode="tozero")
    st.plotly_chart(estilo(fig), use_container_width=True)

izq, der = st.columns(2)
with izq:
    st.subheader("Resultados de inspección")
    resultados = filtrado.groupby("Estado").size().reindex(ESTADOS, fill_value=0).rename("Registros").reset_index()
    fig = px.bar(resultados, x="Estado", y="Registros", color="Estado", text="Registros",
                 color_discrete_map=COLORES, category_orders={"Estado": ESTADOS})
    fig.update_layout(showlegend=False)
    fig.update_yaxes(dtick=1, rangemode="tozero")
    st.plotly_chart(estilo(fig), use_container_width=True)
with der:
    st.subheader("Actividad diaria")
    diario = filtrado.groupby("Fecha").size().rename("Registros").reset_index()
    fig = px.bar(diario, x="Fecha", y="Registros", color_discrete_sequence=["#355C83"])
    fig.update_yaxes(dtick=1, rangemode="tozero")
    st.plotly_chart(estilo(fig), use_container_width=True)
    st.caption("Se muestran las fechas con registros seleccionados; los días sin datos no acreditan ausencia de trabajo.")

st.caption("El volumen y las horas describen carga de trabajo; no miden productividad sin considerar complejidad y alcance. Los hallazgos sumados pueden incluir reincidencias.")
st.subheader("Detalle de actividades")
detalle = filtrado.sort_values("Fecha", ascending=False).copy()
detalle["Fecha"] = detalle["Fecha"].dt.strftime("%Y-%m-%d")
st.dataframe(detalle, hide_index=True, use_container_width=True)
# Neutralizar fórmulas al abrir el CSV en una hoja de cálculo.
exportar = detalle.copy()
for col in exportar.select_dtypes(include="object"):
    exportar[col] = exportar[col].map(lambda v: "'" + v if isinstance(v, str) and v.lstrip().startswith(("=", "+", "-", "@")) else v)
st.download_button("Descargar selección CSV", exportar.to_csv(index=False).encode("utf-8-sig"),
                   file_name="indicadores_inspeccion.csv", mime="text/csv")
