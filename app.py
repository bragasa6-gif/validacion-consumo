import io
from datetime import date

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import load_workbook


st.set_page_config(
    page_title="Validación Consumo",
    page_icon="🦐",
    layout="centered"
)

BASE_FILE = "tabla_base.xlsx"


def to_float(value):
    if value is None:
        return np.nan

    if isinstance(value, str):
        text = value.strip()
        text = text.replace("%", "")
        text = text.replace(".", "")
        text = text.replace(",", ".")

        if text in ("", "-", "#DIV/0!", "#VALUE!", "#N/A", "#REF!", "#NAME?", "#NUM!", "#NULL!"):
            return np.nan

        try:
            return float(text)
        except ValueError:
            return np.nan

    try:
        return float(value)
    except Exception:
        return np.nan


def normalize_bw(value):
    value = to_float(value)

    if pd.isna(value):
        return np.nan

    return value / 100 if value > 1 else value


@st.cache_data(show_spinner=False)
def load_tables_from_bytes(file_bytes):
    wb = load_workbook(
        io.BytesIO(file_bytes),
        data_only=True,
        read_only=True
    )

    sheet_name = None

    for name in wb.sheetnames:
        if name.strip().lower() == "tabla":
            sheet_name = name
            break

    if sheet_name is None:
        raise ValueError("No encontré la hoja Tabla en el Excel.")

    ws = wb[sheet_name]

    normal_rows = []

    for row in ws.iter_rows(
        min_row=4,
        max_row=200,
        min_col=1,
        max_col=3,
        values_only=True
    ):
        peso, frio, calor = row

        peso = to_float(peso)

        if pd.isna(peso):
            continue

        normal_rows.append(
            {
                "peso": peso,
                "NormalFrio": normalize_bw(frio),
                "NormalCalor": normalize_bw(calor),
            }
        )

    min_rows = []

    for row in ws.iter_rows(
        min_row=3,
        max_row=600,
        min_col=6,
        max_col=8,
        values_only=True
    ):
        peso, calor, frio = row

        peso = to_float(peso)

        if pd.isna(peso):
            continue

        min_rows.append(
            {
                "peso": peso,
                "MinCalor": normalize_bw(calor),
                "MinFrio": normalize_bw(frio),
            }
        )

    normal = pd.DataFrame(normal_rows).sort_values("peso").reset_index(drop=True)
    minimo = pd.DataFrame(min_rows).sort_values("peso").reset_index(drop=True)

    if normal.empty or minimo.empty:
        raise ValueError("No pude leer las tablas A:C y F:H de la hoja Tabla.")

    return normal, minimo


def vlookup_approx(table, peso, column):
    valid = table[table["peso"] <= peso]

    if valid.empty:
        row = table.iloc[0]
    else:
        row = valid.iloc[-1]

    return float(row[column]), float(row["peso"])


def safe_div(a, b):
    if b is None or b == 0 or pd.isna(b):
        return np.nan

    return a / b


st.title("Validación de Población por Consumo")
st.caption("App web móvil para ingresar datos y calcular según la hoja Tabla del Excel.")

with st.sidebar:
    st.header("Tabla base")

    uploaded_base = st.file_uploader(
        "Subir Excel base",
        type=["xlsx"],
        help="Debe contener la hoja Tabla"
    )

    usar_incluida = st.toggle(
        "Usar tabla_base.xlsx incluida",
        value=True
    )

    if uploaded_base is not None:
        base_bytes = uploaded_base.getvalue()
        base_label = uploaded_base.name

    elif usar_incluida:
        with open(BASE_FILE, "rb") as f:
            base_bytes = f.read()

        base_label = BASE_FILE

    else:
        st.info("Sube el Excel base para iniciar.")
        st.stop()


try:
    normal_table, min_table = load_tables_from_bytes(base_bytes)

    st.sidebar.success("Tabla cargada")
    st.sidebar.caption(base_label)

except Exception as exc:
    st.error(f"No pude leer la hoja Tabla: {exc}")
    st.stop()


with st.expander("Ver tablas de referencia", expanded=False):
    st.write("Tabla Normal: Peso, NormalFrio, NormalCalor")
    st.dataframe(normal_table, width="stretch", hide_index=True)

    st.write("Tabla Mínima: Peso, MinCalor, MinFrio")
    st.dataframe(min_table, width="stretch", hide_index=True)


st.subheader("Ingreso de datos")

with st.form("validacion_form"):
    fecha = st.date_input(
        "Fecha",
        value=date.today()
    )

    psc = st.text_input(
        "PSC / Piscina",
        value=""
    )

    marca = st.text_input(
        "Marca alimento",
        value=""
    )

    peso_actual = st.number_input(
        "Peso actual (g)",
        min_value=0.01,
        value=3.50,
        step=0.10,
        format="%.2f"
    )

    kg_alimento_actual = st.number_input(
        "Kg alimento semanal actual",
        min_value=0.0,
        value=0.0,
        step=1.0
    )

    dias_alimentados = st.number_input(
        "Días alimentados",
        min_value=1.0,
        value=7.0,
        step=1.0
    )

    animales_vivos = st.number_input(
        "Animales vivos reportados (AL)",
        min_value=0.0,
        value=0.0,
        step=1000.0
    )

    supervivencia = st.number_input(
        "Supervivencia (%)",
        min_value=0.0,
        max_value=100.0,
        value=80.0,
        step=1.0
    )

    epoca = st.radio(
        "Época",
        ["Calor", "Frío"],
        horizontal=True
    )

    umbral_verde = st.number_input(
        "Semáforo verde hasta +/- %",
        min_value=0.0,
        value=5.0,
        step=1.0
    )

    umbral_amarillo = st.number_input(
        "Semáforo amarillo hasta +/- %",
        min_value=0.0,
        value=10.0,
        step=1.0
    )

    calcular = st.form_submit_button("Calcular")


if "historial" not in st.session_state:
    st.session_state.historial = []


if calcular:
    normal_col = "NormalFrio" if epoca == "Calor" else "NormalCalor"
    min_col = "MinCalor" if epoca == "Calor" else "MinFrio"

    bw_sug, peso_ref_normal = vlookup_approx(
        normal_table,
        peso_actual,
        normal_col
    )

    bw_min, peso_ref_min = vlookup_approx(
        min_table,
        peso_actual,
        min_col
    )

    biomasa_reportada = animales_vivos * peso_actual / 1000

    kg_sugerido = biomasa_reportada * bw_sug
    kg_minimo = biomasa_reportada * bw_min

    kg_diario_actual = safe_div(
        kg_alimento_actual,
        dias_alimentados
    )

    biomasa_por_alimento_sugerido = safe_div(
        kg_sugerido,
        bw_sug
    )

    biomasa_por_alimento_minimo = safe_div(
        kg_minimo,
        bw_min
    )

    biomasa_por_alimento_actual = safe_div(
        kg_diario_actual,
        bw_sug
    )

    animales_an = safe_div(
        biomasa_por_alimento_sugerido * 1000,
        peso_actual
    )

    animales_ap = safe_div(
        biomasa_por_alimento_minimo * 1000,
        peso_actual
    )

    animales_ao = safe_div(
        biomasa_por_alimento_actual * 1000,
        peso_actual
    )

    dif_ao_al = animales_ao - animales_vivos if animales_vivos else np.nan
    dif_ao_al_pct = safe_div(dif_ao_al, animales_vivos) * 100 if animales_vivos else np.nan

    def semaforo(pct):
        if pd.isna(pct):
            return "⚪ Sin referencia"

        if abs(pct) <= umbral_verde:
            return "🟢 Verde"

        if abs(pct) <= umbral_amarillo:
            return "🟡 Amarillo"

        return "🔴 Rojo"

    estado = semaforo(dif_ao_al_pct)

    st.subheader("Resultados")
    st.markdown(f"## {estado}")

    c1, c2 = st.columns(2)

    c1.metric(
        "BW sugerida",
        f"{bw_sug * 100:.2f}%"
    )

    c2.metric(
        "BW mínima",
        f"{bw_min * 100:.2f}%"
    )

    c3, c4 = st.columns(2)

    c3.metric(
        "Kg alimento sugerido (U)",
        f"{kg_sugerido:,.0f}"
    )

    c4.metric(
        "Kg mínimo (V)",
        f"{kg_minimo:,.0f}"
    )

    c5, c6 = st.columns(2)

    c5.metric(
        "Animales según alimentación sugerida (AN)",
        f"{animales_an:,.0f}"
    )

    c6.metric(
        "Animales según tabla 2 Min (AP)",
        f"{animales_ap:,.0f}"
    )

    c7, c8 = st.columns(2)

    c7.metric(
        "Animales según alimentación actual (AO)",
        f"{animales_ao:,.0f}"
    )

    c8.metric(
        "Animales vivos reportados (AL)",
        f"{animales_vivos:,.0f}"
    )

    if not pd.isna(dif_ao_al_pct):
        st.metric(
            "Diferencia AO vs AL",
            f"{dif_ao_al:,.0f}",
            f"{dif_ao_al_pct:.2f}%"
        )

    st.caption(
        f"BUSCARV aproximado: BW sugerida usa peso {peso_ref_normal} g | BW mínima usa peso {peso_ref_min} g"
    )

    st.caption(
        f"Kg diario actual calculado: {kg_diario_actual:,.2f} kg/día"
    )

    registro = {
        "Fecha": fecha.isoformat(),
        "PSC": psc,
        "Marca alimento": marca,
        "Peso actual": peso_actual,
        "Kg alimento semanal actual": kg_alimento_actual,
        "Días alimentados": dias_alimentados,
        "Kg diario actual": kg_diario_actual,
        "Animales vivos reportados (AL)": animales_vivos,
        "Supervivencia %": supervivencia,
        "Época": epoca,
        "BW sugerida %": bw_sug * 100,
        "BW mínima %": bw_min * 100,
        "Kg alimento sugerido (U)": kg_sugerido,
        "Kg mínimo (V)": kg_minimo,
        "Animales según alimentación sugerida (AN)": animales_an,
        "Animales según alimentación actual (AO)": animales_ao,
        "Animales según tabla 2 Min (AP)": animales_ap,
        "Diferencia AO vs AL": dif_ao_al,
        "Diferencia AO vs AL %": dif_ao_al_pct,
        "Semáforo": estado,
    }

    st.session_state.historial.append(registro)


st.subheader("Historial de esta sesión")

if st.session_state.historial:
    hist = pd.DataFrame(st.session_state.historial)

    st.dataframe(
        hist,
        width="stretch",
        hide_index=True
    )

    st.download_button(
        "Descargar historial CSV",
        data=hist.to_csv(index=False).encode("utf-8-sig"),
        file_name="historial_validacion_consumo.csv",
        mime="text/csv",
    )

else:
    st.info("Aún no hay registros calculados.")