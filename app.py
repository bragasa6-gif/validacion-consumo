import io
from datetime import date

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import load_workbook


st.set_page_config(
    page_title="Validación de Consumo",
    page_icon="🦐",
    layout="centered"
)

BASE_FILE = "tabla_base.xlsx"


st.markdown(
    """
    <style>
    .main-title {
        text-align: center;
        font-size: 38px;
        font-weight: 800;
        margin-bottom: 0px;
    }
    .subtitle {
        text-align: center;
        color: #6b7280;
        font-size: 16px;
        margin-bottom: 25px;
    }
    .card {
        background: #ffffff;
        padding: 18px;
        border-radius: 18px;
        box-shadow: 0 4px 18px rgba(0,0,0,0.08);
        margin-bottom: 14px;
        border: 1px solid #eef2f7;
    }
    .metric-label {
        color: #6b7280;
        font-size: 14px;
        margin-bottom: 4px;
    }
    .metric-value {
        font-size: 30px;
        font-weight: 800;
        color: #111827;
    }
    .semaforo {
        text-align: center;
        font-size: 42px;
        font-weight: 900;
        padding: 24px;
        border-radius: 24px;
        margin-bottom: 20px;
    }
    .verde {
        background: #dcfce7;
        color: #166534;
    }
    .amarillo {
        background: #fef9c3;
        color: #854d0e;
    }
    .rojo {
        background: #fee2e2;
        color: #991b1b;
    }
    .gris {
        background: #f3f4f6;
        color: #374151;
    }
    </style>
    """,
    unsafe_allow_html=True
)


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


def metric_card(label, value):
    st.markdown(
        f"""
        <div class="card">
            <div class="metric-label">{label}</div>
            <div class="metric-value">{value}</div>
        </div>
        """,
        unsafe_allow_html=True
    )


def semaforo_html(estado):
    if "Verde" in estado:
        css = "verde"
    elif "Amarillo" in estado:
        css = "amarillo"
    elif "Rojo" in estado:
        css = "rojo"
    else:
        css = "gris"

    st.markdown(
        f"""
        <div class="semaforo {css}">
            {estado}
        </div>
        """,
        unsafe_allow_html=True
    )


st.markdown("<div class='main-title'>🦐 Validación de Consumo</div>", unsafe_allow_html=True)
st.markdown("<div class='subtitle'>Cálculo rápido de población por consumo desde el celular</div>", unsafe_allow_html=True)


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
    st.write("Tabla Normal")
    st.dataframe(normal_table, hide_index=True)

    st.write("Tabla Mínima")
    st.dataframe(min_table, hide_index=True)


st.subheader("Ingreso de datos")

with st.form("validacion_form"):
    fecha = st.date_input("Fecha", value=date.today())

    psc = st.text_input("Piscina / PSC", value="")

    marca = st.text_input("Marca alimento", value="")

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
        "Animales vivos reportados",
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

    calcular = st.form_submit_button("Calcular validación")


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

    biomasa_sugerida = safe_div(
        kg_sugerido,
        bw_sug
    )

    biomasa_minima = safe_div(
        kg_minimo,
        bw_min
    )

    biomasa_actual = safe_div(
        kg_diario_actual,
        bw_sug
    )

    animales_an = safe_div(
        biomasa_sugerida * 1000,
        peso_actual
    )

    animales_ap = safe_div(
        biomasa_minima * 1000,
        peso_actual
    )

    animales_ao = safe_div(
        biomasa_actual * 1000,
        peso_actual
    )

    dif_ao_al = animales_ao - animales_vivos if animales_vivos else np.nan
    dif_ao_al_pct = safe_div(dif_ao_al, animales_vivos) * 100 if animales_vivos else np.nan

    if pd.isna(dif_ao_al_pct):
        estado = "⚪ Sin referencia"
    elif abs(dif_ao_al_pct) <= 5:
        estado = "🟢 Verde"
    elif abs(dif_ao_al_pct) <= 10:
        estado = "🟡 Amarillo"
    else:
        estado = "🔴 Rojo"

    st.subheader("Resultado de validación")

    semaforo_html(estado)

    metric_card("Diferencia entre alimentación actual y animales reportados", f"{dif_ao_al:,.0f}")
    metric_card("Diferencia porcentual", f"{dif_ao_al_pct:.2f}%")

    st.subheader("Resumen principal")

    c1, c2 = st.columns(2)
    with c1:
        metric_card("Animales según alimentación actual", f"{animales_ao:,.0f}")
    with c2:
        metric_card("Animales vivos reportados", f"{animales_vivos:,.0f}")

    c3, c4 = st.columns(2)
    with c3:
        metric_card("Kg alimento sugerido", f"{kg_sugerido:,.0f}")
    with c4:
        metric_card("Kg mínimo", f"{kg_minimo:,.0f}")

    st.subheader("Detalle técnico")

    c5, c6 = st.columns(2)
    with c5:
        metric_card("BW sugerida", f"{bw_sug * 100:.2f}%")
    with c6:
        metric_card("BW mínima", f"{bw_min * 100:.2f}%")

    c7, c8 = st.columns(2)
    with c7:
        metric_card("Animales según alimentación sugerida", f"{animales_an:,.0f}")
    with c8:
        metric_card("Animales según tabla 2 Min", f"{animales_ap:,.0f}")

    metric_card("Kg diario actual", f"{kg_diario_actual:,.2f} kg/día")

    st.caption(
        f"Referencia usada: BW sugerida con peso {peso_ref_normal} g | BW mínima con peso {peso_ref_min} g"
    )

    registro = {
        "Fecha": fecha.isoformat(),
        "Piscina / PSC": psc,
        "Marca alimento": marca,
        "Peso actual": peso_actual,
        "Kg alimento semanal actual": kg_alimento_actual,
        "Días alimentados": dias_alimentados,
        "Kg diario actual": kg_diario_actual,
        "Animales vivos reportados": animales_vivos,
        "Supervivencia %": supervivencia,
        "Época": epoca,
        "BW sugerida %": bw_sug * 100,
        "BW mínima %": bw_min * 100,
        "Kg alimento sugerido": kg_sugerido,
        "Kg mínimo": kg_minimo,
        "Animales según alimentación sugerida": animales_an,
        "Animales según alimentación actual": animales_ao,
        "Animales según tabla 2 Min": animales_ap,
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
        hide_index=True
    )

    st.download_button(
        "Descargar historial",
        data=hist.to_csv(index=False).encode("utf-8-sig"),
        file_name="historial_validacion_consumo.csv",
        mime="text/csv",
    )

else:
    st.info("Aún no hay registros calculados.")
