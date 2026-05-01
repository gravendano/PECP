import io
import math

import numpy as np
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt


# ============================================================
# Simulador de Priorización de Proyectos
# Constructora Altavista, S. A.
# ============================================================

st.set_page_config(
    page_title="Comité de Priorización de Proyectos",
    page_icon="🏗️",
    layout="wide",
)


# ------------------------------------------------------------
# Estilos visuales
# ------------------------------------------------------------
st.markdown(
    """
    <style>
    .main-title {
        background: #3f7f92;
        color: white;
        padding: 1.2rem 1.5rem;
        border-radius: 14px;
        margin-bottom: 1rem;
        text-align: center;
        font-size: 2rem;
        font-weight: 800;
        letter-spacing: .02em;
    }
    .subtitle-box {
        border: 1px solid #d8dee9;
        border-left: 8px solid #3f7f92;
        border-radius: 12px;
        padding: 1rem 1.25rem;
        background: #f8fafc;
        margin-bottom: 1rem;
    }
    .section-label {
        display: inline-block;
        background: #5c6f91;
        color: white;
        padding: .35rem .75rem;
        border-radius: 6px;
        font-weight: 800;
        letter-spacing: .08em;
        margin-top: 1rem;
        margin-bottom: .4rem;
    }
    .metric-card {
        border: 1px solid #d8dee9;
        border-radius: 14px;
        padding: 1rem;
        background: #ffffff;
        box-shadow: 0 1px 4px rgba(0,0,0,0.06);
    }
    .ok-box {
        border-left: 8px solid #198754;
        background: #f1f8f4;
        padding: .8rem 1rem;
        border-radius: 10px;
        margin-bottom: .5rem;
    }
    .bad-box {
        border-left: 8px solid #dc3545;
        background: #fff5f5;
        padding: .8rem 1rem;
        border-radius: 10px;
        margin-bottom: .5rem;
    }
    .warn-box {
        border-left: 8px solid #f0ad4e;
        background: #fff9ec;
        padding: .8rem 1rem;
        border-radius: 10px;
        margin-bottom: .5rem;
    }
    .small-note {
        color: #475569;
        font-size: .92rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ------------------------------------------------------------
# Datos base
# ------------------------------------------------------------
BASE_RATE = 0.10
YEARS = [0, 1, 2, 3, 4, 5]

PROJECTS = [
    {
        "codigo": "P1",
        "proyecto": "Edificio residencial en Zona 10",
        "tipo": "Construcción",
        "descripcion": "Torre de apartamentos de nivel medio-alto con locales comerciales en el primer nivel.",
        "construccion": True,
        "horas_tecnicas": 900,
        "horas_gerencia": 520,
        "flujos": [-4_500_000, 700_000, 950_000, 1_250_000, 1_500_000, 2_600_000],
        "scores": [5, 5, 3, 3, 2],
    },
    {
        "codigo": "P2",
        "proyecto": "Condominio horizontal en Carretera a El Salvador",
        "tipo": "Construcción",
        "descripcion": "Desarrollo de casas familiares con calles internas, garita, áreas verdes y amenidades básicas.",
        "construccion": True,
        "horas_tecnicas": 780,
        "horas_gerencia": 460,
        "flujos": [-3_800_000, 850_000, 1_000_000, 1_150_000, 1_300_000, 1_600_000],
        "scores": [4, 4, 4, 3, 3],
    },
    {
        "codigo": "P3",
        "proyecto": "Remodelación de oficinas corporativas",
        "tipo": "Construcción",
        "descripcion": "Remodelación integral para una empresa privada con contrato casi asegurado y ejecución de corto plazo.",
        "construccion": True,
        "horas_tecnicas": 300,
        "horas_gerencia": 180,
        "flujos": [-900_000, 450_000, 400_000, 300_000, 200_000, 150_000],
        "scores": [4, 3, 5, 5, 5],
    },
    {
        "codigo": "P4",
        "proyecto": "Bodega industrial en Villa Nueva",
        "tipo": "Construcción",
        "descripcion": "Construcción de bodega para renta logística cerca de corredores industriales.",
        "construccion": True,
        "horas_tecnicas": 520,
        "horas_gerencia": 330,
        "flujos": [-2_800_000, 500_000, 650_000, 750_000, 850_000, 1_300_000],
        "scores": [3, 4, 4, 4, 4],
    },
    {
        "codigo": "P5",
        "proyecto": "Digitalización del control de obra",
        "tipo": "Proyecto interno",
        "descripcion": "Implementación de software para presupuestos, avances, compras, reportes, control documental y seguimiento de obra.",
        "construccion": False,
        "horas_tecnicas": 250,
        "horas_gerencia": 250,
        "flujos": [-650_000, 180_000, 220_000, 260_000, 300_000, 340_000],
        "scores": [4, 5, 4, 5, 5],
    },
]

CRITERIA = [
    {"criterio": "Rentabilidad financiera", "peso": 10},
    {"criterio": "Alineación estratégica", "peso": 8},
    {"criterio": "Factibilidad técnica y legal", "peso": 7},
    {"criterio": "Riesgo del proyecto", "peso": 6},
    {"criterio": "Uso eficiente de recursos", "peso": 5},
]

LIMITS = {
    "capital": 7_500_000,
    "horas_tecnicas": 2_000,
    "horas_gerencia": 1_250,
    "max_proyectos": 3,
    "max_construccion": 2,
}


# ------------------------------------------------------------
# Funciones financieras
# ------------------------------------------------------------
def npv_excel(rate: float, cashflows: list[float]) -> float:
    """Calcula VPN como Excel: VNA(tasa, años 1:n) + flujo año 0."""
    initial = cashflows[0]
    future = cashflows[1:]
    return initial + sum(cf / ((1 + rate) ** i) for i, cf in enumerate(future, start=1))


def irr(cashflows: list[float]) -> float | None:
    """Calcula TIR usando numpy_financial si está disponible, con fallback por búsqueda."""
    try:
        import numpy_financial as npf
        value = npf.irr(cashflows)
        if value is None or np.isnan(value):
            return None
        return float(value)
    except Exception:
        pass

    # Fallback simple: busca una tasa entre -90% y 100% que aproxime VPN = 0.
    rates = np.linspace(-0.90, 1.00, 5000)
    values = [npv_excel(r, cashflows) for r in rates]
    for i in range(len(values) - 1):
        if values[i] == 0:
            return float(rates[i])
        if values[i] * values[i + 1] < 0:
            low, high = rates[i], rates[i + 1]
            for _ in range(60):
                mid = (low + high) / 2
                if npv_excel(low, cashflows) * npv_excel(mid, cashflows) <= 0:
                    high = mid
                else:
                    low = mid
            return float((low + high) / 2)
    return None


def payback(cashflows: list[float]) -> float | None:
    """Calcula payback simple con interpolación del año de recuperación."""
    cumulative = cashflows[0]
    if cumulative >= 0:
        return 0.0

    for year in range(1, len(cashflows)):
        previous = cumulative
        cumulative += cashflows[year]
        if cumulative >= 0:
            pending = abs(previous)
            flow = cashflows[year]
            if flow == 0:
                return None
            return (year - 1) + pending / flow
    return None


def financial_table(cashflows: list[float], rate: float) -> pd.DataFrame:
    rows = []
    cumulative = 0
    for year, cf in enumerate(cashflows):
        factor = 1 / ((1 + rate) ** year)
        present_value = cf * factor
        cumulative += cf
        rows.append(
            {
                "Año": year,
                "Flujo neto": cf,
                "Factor de descuento": factor,
                "Valor presente": present_value,
                "Flujo acumulado": cumulative,
            }
        )
    return pd.DataFrame(rows)


def format_q(value: float) -> str:
    return f"Q {value:,.0f}"


def format_pct(value: float | None) -> str:
    if value is None or np.isnan(value):
        return "No disponible"
    return f"{value * 100:.1f}%"


def format_years(value: float | None) -> str:
    if value is None:
        return "No recupera"
    return f"{value:.2f} años"


def get_project_by_code(code: str) -> dict:
    return next(p for p in PROJECTS if p["codigo"] == code)


def default_cashflow_df() -> pd.DataFrame:
    data = {"Año": YEARS}
    for project in PROJECTS:
        data[f'{project["codigo"]} - {project["proyecto"]}'] = project["flujos"]
    return pd.DataFrame(data)


def build_results(cashflow_df: pd.DataFrame, rate: float) -> pd.DataFrame:
    rows = []
    for project in PROJECTS:
        col = f'{project["codigo"]} - {project["proyecto"]}'
        flows = cashflow_df[col].astype(float).tolist()
        rows.append(
            {
                "Código": project["codigo"],
                "Proyecto": project["proyecto"],
                "Tipo": project["tipo"],
                "Inversión inicial": abs(flows[0]),
                "VPN": npv_excel(rate, flows),
                "TIR": irr(flows),
                "Payback": payback(flows),
                "Horas técnicas": project["horas_tecnicas"],
                "Horas gerencia": project["horas_gerencia"],
                "Proyecto de construcción": "Sí" if project["construccion"] else "No",
            }
        )
    return pd.DataFrame(rows)


def default_scores_df() -> pd.DataFrame:
    rows = []
    for project in PROJECTS:
        row = {"Código": project["codigo"], "Proyecto": project["proyecto"]}
        for criterion, score in zip(CRITERIA, project["scores"]):
            row[criterion["criterio"]] = score
        rows.append(row)
    return pd.DataFrame(rows)


def score_results(scores_df: pd.DataFrame, weights: dict[str, int]) -> pd.DataFrame:
    out = scores_df.copy()
    total = np.zeros(len(out))
    for criterion in weights:
        score_col = criterion
        weighted_col = f"{criterion} ponderado"
        out[weighted_col] = out[score_col].astype(float) * weights[criterion]
        total += out[weighted_col]
    out["Puntaje total"] = total
    out["Ranking"] = out["Puntaje total"].rank(ascending=False, method="min").astype(int)
    return out.sort_values(["Puntaje total", "Código"], ascending=[False, True])


def timeline_html(project: dict, flows: list[float]) -> str:
    labels = []
    for year, cf in enumerate(flows):
        if year == 0:
            continue
        labels.append(f"<div><strong>{format_q(cf)}</strong><br>↑<br>{year}</div>")

    timeline_items = "".join(labels)
    initial = format_q(abs(flows[0]))
    return f"""
    <div style="border: 2px solid #a7b0c4; border-radius: 24px; padding: 1.2rem; margin: .5rem 0 1rem 0; background: white;">
        <div style="display:inline-block; background:#5c6f91; color:white; padding:.3rem .7rem; font-weight:800; border-radius:4px; letter-spacing:.08em;">LÍNEA DE TIEMPO</div>
        <h3 style="margin-top:.8rem;">{project['codigo']} — {project['proyecto']}</h3>
        <div style="display:flex; align-items:end; gap:2rem; margin-top:1.3rem; overflow-x:auto;">
            <div style="text-align:center; min-width:110px;">
                <strong>Año 0</strong><br>
                ↓<br>
                <strong>{initial}</strong>
            </div>
            <div style="height:2px; background:#222; flex:1; min-width:80px; margin-bottom:2.4rem;"></div>
            <div style="display:flex; gap:2.5rem; text-align:center; min-width:650px;">
                {timeline_items}
            </div>
        </div>
        <p style="font-size:.92rem; color:#475569; margin-top:1rem;">Los flujos positivos se reciben al final de cada año. El flujo del año 0 representa la inversión inicial.</p>
    </div>
    """


def profile_npv_dataframe(cashflow_df: pd.DataFrame, selected_codes: list[str]) -> pd.DataFrame:
    rates = np.linspace(0, 0.30, 31)
    rows = []
    for rate in rates:
        row = {"Tasa": rate}
        for code in selected_codes:
            project = get_project_by_code(code)
            col = f'{project["codigo"]} - {project["proyecto"]}'
            flows = cashflow_df[col].astype(float).tolist()
            row[code] = npv_excel(rate, flows)
        rows.append(row)
    return pd.DataFrame(rows)


def create_excel_export(
    cashflow_df: pd.DataFrame,
    financial_results: pd.DataFrame,
    scores_ranked: pd.DataFrame,
    selected_portfolio: pd.DataFrame,
    validation_df: pd.DataFrame,
    recommendation: dict,
    rate: float,
) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary = pd.DataFrame(
            {
                "Dato": [
                    "Simulación",
                    "Tasa de descuento",
                    "Horizonte",
                    "Capital disponible",
                    "Horas técnicas disponibles",
                    "Horas de gerencia disponibles",
                    "Máximo de proyectos activos",
                    "Máximo de proyectos de construcción",
                ],
                "Valor": [
                    "Comité de Priorización de Proyectos — Constructora Altavista, S. A.",
                    f"{rate:.2%}",
                    "5 años",
                    LIMITS["capital"],
                    LIMITS["horas_tecnicas"],
                    LIMITS["horas_gerencia"],
                    LIMITS["max_proyectos"],
                    LIMITS["max_construccion"],
                ],
            }
        )
        summary.to_excel(writer, sheet_name="Resumen", index=False)
        cashflow_df.to_excel(writer, sheet_name="Flujos", index=False)
        financial_results.to_excel(writer, sheet_name="Indicadores", index=False)
        scores_ranked.to_excel(writer, sheet_name="Matriz ponderada", index=False)
        selected_portfolio.to_excel(writer, sheet_name="Portafolio", index=False)
        validation_df.to_excel(writer, sheet_name="Restricciones", index=False)
        pd.DataFrame([recommendation]).to_excel(writer, sheet_name="Recomendación", index=False)

        for sheet_name in writer.sheets:
            ws = writer.sheets[sheet_name]
            for column_cells in ws.columns:
                max_length = 0
                column_letter = column_cells[0].column_letter
                for cell in column_cells:
                    try:
                        max_length = max(max_length, len(str(cell.value)))
                    except Exception:
                        pass
                ws.column_dimensions[column_letter].width = min(max_length + 2, 55)

    return output.getvalue()


# ------------------------------------------------------------
# Estado inicial
# ------------------------------------------------------------
if "cashflow_df" not in st.session_state:
    st.session_state.cashflow_df = default_cashflow_df()

if "scores_df" not in st.session_state:
    st.session_state.scores_df = default_scores_df()


# ------------------------------------------------------------
# Encabezado
# ------------------------------------------------------------
st.markdown(
    '<div class="main-title">Comité de Priorización de Proyectos — Constructora Altavista, S. A.</div>',
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="subtitle-box">
    <strong>Simulación de 70 minutos.</strong> El comité debe calcular indicadores financieros, evaluar proyectos con una matriz ponderada y seleccionar un portafolio que respete restricciones de capital y capacidad interna.
    </div>
    """,
    unsafe_allow_html=True,
)


# ------------------------------------------------------------
# Barra lateral
# ------------------------------------------------------------
st.sidebar.header("Parámetros de la simulación")
rate = st.sidebar.slider(
    "Tasa de descuento anual",
    min_value=0.00,
    max_value=0.30,
    value=BASE_RATE,
    step=0.005,
    format="%.3f",
)

st.sidebar.caption("Todos los proyectos se evalúan con vida útil de 5 años. No se usa PMT/PAGO.")

st.sidebar.subheader("Pesos de criterios")
weights = {}
for criterion in CRITERIA:
    weights[criterion["criterio"]] = st.sidebar.number_input(
        criterion["criterio"],
        min_value=1,
        max_value=20,
        value=criterion["peso"],
        step=1,
    )

if st.sidebar.button("Restablecer datos base"):
    st.session_state.cashflow_df = default_cashflow_df()
    st.session_state.scores_df = default_scores_df()
    st.rerun()


# ------------------------------------------------------------
# Pestañas
# ------------------------------------------------------------
tabs = st.tabs(
    [
        "1. Caso",
        "2. Flujos y VPN",
        "3. Comparación financiera",
        "4. Matriz ponderada",
        "5. Portafolio",
        "6. Recomendación y Excel",
    ]
)


# ------------------------------------------------------------
# Tab 1: Caso
# ------------------------------------------------------------
with tabs[0]:
    st.markdown('<div class="section-label">CASO</div>', unsafe_allow_html=True)
    st.subheader("Constructora Altavista, S. A.")
    st.write(
        """
        Constructora Altavista, S. A. es una empresa guatemalteca dedicada al desarrollo y construcción de proyectos residenciales, comerciales e industriales de mediana escala. La empresa ha recibido varias oportunidades de inversión, pero no cuenta con suficiente capital ni capacidad interna para ejecutarlas todas al mismo tiempo.

        En años anteriores, algunas decisiones se tomaban principalmente por intuición, presión comercial, interés de inversionistas o urgencia del momento. Ahora, la Gerencia General desea formalizar el proceso de selección mediante un Comité de Priorización de Proyectos.

        El comité deberá calcular indicadores financieros, aplicar una matriz ponderada y seleccionar el mejor portafolio posible.
        """
    )

    st.markdown('<div class="section-label">OBJETIVOS ESTRATÉGICOS</div>', unsafe_allow_html=True)
    objectives = pd.DataFrame(
        [
            ["Crecer con rentabilidad", "Seleccionar proyectos con margen atractivo, retorno razonable y control de costos."],
            ["Fortalecer la reputación", "Priorizar proyectos que mejoren imagen, confianza y posicionamiento."],
            ["Reducir riesgos de ejecución", "Evitar permisos inciertos, diseños incompletos o alta exposición financiera."],
            ["Mejorar la eficiencia operativa", "Ordenar diseño, compras, supervisión, control de obra y cierre."],
            ["Diversificar el portafolio", "No depender de un solo tipo de proyecto, zona, cliente o fuente de ingresos."],
        ],
        columns=["Objetivo estratégico", "Descripción"],
    )
    st.dataframe(objectives, use_container_width=True, hide_index=True)

    st.markdown('<div class="section-label">PROYECTOS CANDIDATOS</div>', unsafe_allow_html=True)
    project_table = pd.DataFrame(
        [
            {
                "Código": p["codigo"],
                "Proyecto": p["proyecto"],
                "Tipo": p["tipo"],
                "Descripción": p["descripcion"],
            }
            for p in PROJECTS
        ]
    )
    st.dataframe(project_table, use_container_width=True, hide_index=True)


# ------------------------------------------------------------
# Tab 2: Flujos y VPN
# ------------------------------------------------------------
with tabs[1]:
    st.markdown('<div class="section-label">DATOS DE FLUJOS DE EFECTIVO</div>', unsafe_allow_html=True)
    st.write(
        "Los flujos están precargados, pero pueden editarse para simular cambios en inversión, ingresos, ahorros o recuperación del proyecto."
    )

    edited_cashflows = st.data_editor(
        st.session_state.cashflow_df,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        key="cashflow_editor",
    )
    st.session_state.cashflow_df = edited_cashflows

    project_options = [f'{p["codigo"]} — {p["proyecto"]}' for p in PROJECTS]
    selected_project_label = st.selectbox("Proyecto para visualizar", project_options)
    selected_code = selected_project_label.split(" — ")[0]
    selected_project = get_project_by_code(selected_code)
    selected_col = f'{selected_project["codigo"]} - {selected_project["proyecto"]}'
    selected_flows = edited_cashflows[selected_col].astype(float).tolist()

    st.markdown(timeline_html(selected_project, selected_flows), unsafe_allow_html=True)

    calc_df = financial_table(selected_flows, rate)
    display_calc = calc_df.copy()
    display_calc["Flujo neto"] = display_calc["Flujo neto"].map(format_q)
    display_calc["Factor de descuento"] = display_calc["Factor de descuento"].map(lambda x: f"{x:.4f}")
    display_calc["Valor presente"] = display_calc["Valor presente"].map(format_q)
    display_calc["Flujo acumulado"] = display_calc["Flujo acumulado"].map(format_q)

    st.markdown('<div class="section-label">DETERMINACIÓN DEL VALOR PRESENTE NETO</div>', unsafe_allow_html=True)
    st.dataframe(display_calc, use_container_width=True, hide_index=True)

    project_npv = npv_excel(rate, selected_flows)
    project_irr = irr(selected_flows)
    project_payback = payback(selected_flows)

    col1, col2, col3 = st.columns(3)
    col1.metric("VPN", format_q(project_npv))
    col2.metric("TIR", format_pct(project_irr))
    col3.metric("Payback", format_years(project_payback))

    st.info(
        "En Excel en español: VPN = VNA(tasa, flujos del año 1 al año 5) + flujo del año 0. La TIR se calcula con TIR(flujos del año 0 al año 5)."
    )


# ------------------------------------------------------------
# Tab 3: Comparación financiera
# ------------------------------------------------------------
with tabs[2]:
    financial_results = build_results(st.session_state.cashflow_df, rate)

    st.markdown('<div class="section-label">COMPARACIÓN FINANCIERA</div>', unsafe_allow_html=True)
    display_financial = financial_results.copy()
    display_financial["Inversión inicial"] = display_financial["Inversión inicial"].map(format_q)
    display_financial["VPN"] = display_financial["VPN"].map(format_q)
    display_financial["TIR"] = display_financial["TIR"].map(format_pct)
    display_financial["Payback"] = display_financial["Payback"].map(format_years)
    st.dataframe(display_financial, use_container_width=True, hide_index=True)

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("VPN por proyecto")
        fig, ax = plt.subplots()
        ax.bar(financial_results["Código"], financial_results["VPN"])
        ax.set_ylabel("VPN en Q")
        ax.set_xlabel("Proyecto")
        st.pyplot(fig)

    with col2:
        st.subheader("TIR por proyecto")
        fig, ax = plt.subplots()
        ax.bar(financial_results["Código"], financial_results["TIR"] * 100)
        ax.set_ylabel("TIR (%)")
        ax.set_xlabel("Proyecto")
        st.pyplot(fig)

    st.markdown('<div class="section-label">PERFIL DEL VPN</div>', unsafe_allow_html=True)
    st.write("Esta gráfica muestra cómo cambia el VPN cuando cambia la tasa de descuento. La TIR es el punto donde el VPN se acerca a cero.")

    selected_profile_codes = st.multiselect(
        "Seleccione proyectos para comparar",
        options=[p["codigo"] for p in PROJECTS],
        default=["P1", "P2", "P3"],
    )

    if selected_profile_codes:
        profile_df = profile_npv_dataframe(st.session_state.cashflow_df, selected_profile_codes)
        fig, ax = plt.subplots()
        for code in selected_profile_codes:
            ax.plot(profile_df["Tasa"] * 100, profile_df[code], label=code)
        ax.axhline(0, linewidth=1)
        ax.set_xlabel("Tasa de descuento (%)")
        ax.set_ylabel("VPN")
        ax.legend()
        st.pyplot(fig)


# ------------------------------------------------------------
# Tab 4: Matriz ponderada
# ------------------------------------------------------------
with tabs[3]:
    st.markdown('<div class="section-label">MATRIZ DE SELECCIÓN Y PRIORIZACIÓN</div>', unsafe_allow_html=True)
    st.write(
        "Edite las calificaciones de 1 a 5. En el criterio de riesgo, 5 significa bajo riesgo y 1 significa alto riesgo."
    )

    edited_scores = st.data_editor(
        st.session_state.scores_df,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        key="scores_editor",
        column_config={
            criterion["criterio"]: st.column_config.NumberColumn(
                criterion["criterio"],
                min_value=1,
                max_value=5,
                step=1,
            )
            for criterion in CRITERIA
        },
    )
    st.session_state.scores_df = edited_scores

    scores_ranked = score_results(edited_scores, weights)

    display_ranked = scores_ranked[
        ["Ranking", "Código", "Proyecto"]
        + [criterion["criterio"] for criterion in CRITERIA]
        + [f'{criterion["criterio"]} ponderado' for criterion in CRITERIA]
        + ["Puntaje total"]
    ].copy()
    st.dataframe(display_ranked, use_container_width=True, hide_index=True)

    st.subheader("Ranking ponderado")
    fig, ax = plt.subplots()
    ax.bar(scores_ranked["Código"], scores_ranked["Puntaje total"])
    ax.set_ylabel("Puntaje total")
    ax.set_xlabel("Proyecto")
    st.pyplot(fig)


# ------------------------------------------------------------
# Tab 5: Portafolio
# ------------------------------------------------------------
with tabs[4]:
    financial_results = build_results(st.session_state.cashflow_df, rate)
    scores_ranked = score_results(st.session_state.scores_df, weights)
    merged = financial_results.merge(
        scores_ranked[["Código", "Puntaje total", "Ranking"]],
        on="Código",
        how="left",
    )

    st.markdown('<div class="section-label">SELECCIÓN DEL PORTAFOLIO</div>', unsafe_allow_html=True)
    st.write("Seleccione hasta tres proyectos y revise si el portafolio cumple las restricciones.")

    selected_codes = []
    cols = st.columns(len(PROJECTS))
    for col, project in zip(cols, PROJECTS):
        with col:
            selected = st.checkbox(
                f'{project["codigo"]}',
                value=project["codigo"] in ["P1", "P3", "P5"],
                help=project["proyecto"],
            )
            if selected:
                selected_codes.append(project["codigo"])

    selected_portfolio = merged[merged["Código"].isin(selected_codes)].copy()

    if selected_portfolio.empty:
        st.warning("Seleccione al menos un proyecto para evaluar el portafolio.")
    else:
        capital_used = selected_portfolio["Inversión inicial"].sum()
        vpn_total = selected_portfolio["VPN"].sum()
        tech_used = selected_portfolio["Horas técnicas"].sum()
        management_used = selected_portfolio["Horas gerencia"].sum()
        active_projects = len(selected_portfolio)
        construction_projects = (selected_portfolio["Proyecto de construcción"] == "Sí").sum()
        total_score = selected_portfolio["Puntaje total"].sum()
        avg_score = selected_portfolio["Puntaje total"].mean()

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Capital usado", format_q(capital_used))
        m2.metric("VPN total", format_q(vpn_total))
        m3.metric("Horas técnicas", f"{tech_used:,.0f}")
        m4.metric("Horas gerencia", f"{management_used:,.0f}")

        m5, m6, m7 = st.columns(3)
        m5.metric("Proyectos activos", f"{active_projects}")
        m6.metric("Proyectos construcción", f"{construction_projects}")
        m7.metric("Puntaje promedio", f"{avg_score:.1f}")

        validation_rows = [
            ["Capital", capital_used, LIMITS["capital"], capital_used <= LIMITS["capital"]],
            ["Horas técnicas", tech_used, LIMITS["horas_tecnicas"], tech_used <= LIMITS["horas_tecnicas"]],
            ["Horas de gerencia", management_used, LIMITS["horas_gerencia"], management_used <= LIMITS["horas_gerencia"]],
            ["Proyectos activos", active_projects, LIMITS["max_proyectos"], active_projects <= LIMITS["max_proyectos"]],
            ["Proyectos de construcción", construction_projects, LIMITS["max_construccion"], construction_projects <= LIMITS["max_construccion"]],
        ]
        validation_df = pd.DataFrame(validation_rows, columns=["Restricción", "Uso", "Límite", "Cumple"])

        st.markdown('<div class="section-label">VALIDACIÓN DE RESTRICCIONES</div>', unsafe_allow_html=True)
        for _, row in validation_df.iterrows():
            if row["Cumple"]:
                st.markdown(
                    f'<div class="ok-box">✅ <strong>{row["Restricción"]}</strong>: cumple. Uso: {row["Uso"]:,.0f} / Límite: {row["Límite"]:,.0f}</div>',
                    unsafe_allow_html=True,
                )
            else:
                st.markdown(
                    f'<div class="bad-box">❌ <strong>{row["Restricción"]}</strong>: no cumple. Uso: {row["Uso"]:,.0f} / Límite: {row["Límite"]:,.0f}</div>',
                    unsafe_allow_html=True,
                )

        display_portfolio = selected_portfolio.copy()
        display_portfolio["Inversión inicial"] = display_portfolio["Inversión inicial"].map(format_q)
        display_portfolio["VPN"] = display_portfolio["VPN"].map(format_q)
        display_portfolio["TIR"] = display_portfolio["TIR"].map(format_pct)
        display_portfolio["Payback"] = display_portfolio["Payback"].map(format_years)
        st.dataframe(display_portfolio, use_container_width=True, hide_index=True)


# ------------------------------------------------------------
# Tab 6: Recomendación y exportación
# ------------------------------------------------------------
with tabs[5]:
    financial_results = build_results(st.session_state.cashflow_df, rate)
    scores_ranked = score_results(st.session_state.scores_df, weights)
    merged = financial_results.merge(
        scores_ranked[["Código", "Puntaje total", "Ranking"]],
        on="Código",
        how="left",
    )

    st.markdown('<div class="section-label">RECOMENDACIÓN EJECUTIVA</div>', unsafe_allow_html=True)

    selected_export_codes = st.multiselect(
        "Proyectos seleccionados para la recomendación",
        options=[p["codigo"] for p in PROJECTS],
        default=["P1", "P3", "P5"],
    )
    selected_export_portfolio = merged[merged["Código"].isin(selected_export_codes)].copy()

    project_priority = st.selectbox(
        "Proyecto prioritario número 1",
        options=[p["codigo"] for p in PROJECTS],
    )
    discarded_project = st.selectbox(
        "Proyecto atractivo que quedó fuera",
        options=[p["codigo"] for p in PROJECTS],
        index=1,
    )
    main_constraint = st.selectbox(
        "Restricción más importante",
        options=[
            "Capital disponible",
            "Horas técnicas",
            "Horas de gerencia",
            "Máximo de proyectos activos",
            "Máximo de proyectos de construcción",
        ],
    )

    justification_financial = st.text_area("Justificación financiera", height=100)
    justification_strategy = st.text_area("Justificación estratégica", height=100)
    justification_operational = st.text_area("Justificación operativa", height=100)
    main_risk = st.text_area("Riesgo principal", height=80)
    final_recommendation = st.text_area("Recomendación final", height=120)

    recommendation = {
        "Proyectos seleccionados": ", ".join(selected_export_codes),
        "Proyecto prioritario": project_priority,
        "Proyecto descartado relevante": discarded_project,
        "Restricción más importante": main_constraint,
        "Justificación financiera": justification_financial,
        "Justificación estratégica": justification_strategy,
        "Justificación operativa": justification_operational,
        "Riesgo principal": main_risk,
        "Recomendación final": final_recommendation,
    }

    st.markdown('<div class="section-label">EXPORTACIÓN A EXCEL</div>', unsafe_allow_html=True)

    capital_used = selected_export_portfolio["Inversión inicial"].sum() if not selected_export_portfolio.empty else 0
    tech_used = selected_export_portfolio["Horas técnicas"].sum() if not selected_export_portfolio.empty else 0
    management_used = selected_export_portfolio["Horas gerencia"].sum() if not selected_export_portfolio.empty else 0
    active_projects = len(selected_export_portfolio)
    construction_projects = (selected_export_portfolio["Proyecto de construcción"] == "Sí").sum() if not selected_export_portfolio.empty else 0
    validation_df = pd.DataFrame(
        [
            ["Capital", capital_used, LIMITS["capital"], capital_used <= LIMITS["capital"]],
            ["Horas técnicas", tech_used, LIMITS["horas_tecnicas"], tech_used <= LIMITS["horas_tecnicas"]],
            ["Horas de gerencia", management_used, LIMITS["horas_gerencia"], management_used <= LIMITS["horas_gerencia"]],
            ["Proyectos activos", active_projects, LIMITS["max_proyectos"], active_projects <= LIMITS["max_proyectos"]],
            ["Proyectos de construcción", construction_projects, LIMITS["max_construccion"], construction_projects <= LIMITS["max_construccion"]],
        ],
        columns=["Restricción", "Uso", "Límite", "Cumple"],
    )

    excel_bytes = create_excel_export(
        st.session_state.cashflow_df,
        financial_results,
        scores_ranked,
        selected_export_portfolio,
        validation_df,
        recommendation,
        rate,
    )

    st.download_button(
        label="Descargar resultados en Excel",
        data=excel_bytes,
        file_name="simulador_altavista_resultados.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.caption(
        "El archivo incluye hojas para resumen, flujos, indicadores financieros, matriz ponderada, portafolio, restricciones y recomendación ejecutiva."
    )
