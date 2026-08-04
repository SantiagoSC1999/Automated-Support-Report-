import os
import glob
from datetime import datetime

import streamlit as st

from report_engine import generate_report, ValidationError, load_target_items

st.set_page_config(
    page_title="Generador de Reportes Mensuales",
    page_icon="📊",
    layout="wide",
)

CUSTOM_CSS = """
<style>
#MainMenu, footer, header {visibility: hidden;}

.hero {
    padding: 2rem 2.2rem;
    border-radius: 18px;
    background: linear-gradient(135deg, #6D28D9 0%, #4338CA 55%, #1E3A8A 100%);
    color: white;
    margin-bottom: 1.6rem;
    box-shadow: 0 10px 30px rgba(67, 56, 202, 0.25);
}
.hero h1 {
    margin: 0 0 0.3rem 0;
    font-size: 2rem;
}
.hero p {
    margin: 0;
    opacity: 0.9;
    font-size: 1.05rem;
}

.metric-card {
    background: var(--background-color, #ffffff);
    border: 1px solid rgba(120, 120, 120, 0.15);
    border-left: 4px solid #9CA3AF;
    border-radius: 14px;
    padding: 1rem 1.2rem;
    text-align: center;
}
.metric-card.number { border-left-color: #1560BD; }
.metric-card.percent { border-left-color: #7B2D8E; }
.metric-card .value {
    font-size: 1.8rem;
    font-weight: 700;
}
.metric-card .label {
    font-size: 0.85rem;
    opacity: 0.7;
}

.section-label {
    font-size: 0.8rem;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    opacity: 0.55;
    margin: 0.6rem 0 0.4rem 0;
}

div[data-testid="stFileUploaderDropzone"] {
    border-radius: 14px;
}

/* --- AI status animation --- */
.ai-status-card {
    display: flex;
    align-items: center;
    gap: 0.9rem;
    padding: 1rem 1.2rem;
    border-radius: 14px;
    background: linear-gradient(135deg, rgba(109,40,217,0.08), rgba(30,58,138,0.08));
    border: 1px solid rgba(109,40,217,0.25);
}
.ai-orb {
    flex-shrink: 0;
    width: 30px;
    height: 30px;
    border-radius: 50%;
    background: linear-gradient(135deg, #8B5CF6, #4338CA, #1560BD);
    background-size: 200% 200%;
    animation: ai-pulse 1.6s ease-in-out infinite, ai-gradient 3s ease infinite;
}
.ai-status-text {
    font-weight: 700;
    color: #4338CA;
}
.ai-status-sub {
    font-size: 0.88rem;
    opacity: 0.75;
    margin-top: 0.1rem;
}
.dot-flashing span {
    animation: dot-flash 1.4s infinite;
    opacity: 0.2;
}
.dot-flashing span:nth-child(2) { animation-delay: 0.2s; }
.dot-flashing span:nth-child(3) { animation-delay: 0.4s; }

@keyframes ai-pulse {
    0% { box-shadow: 0 0 0 0 rgba(139,92,246,0.45); }
    70% { box-shadow: 0 0 0 12px rgba(139,92,246,0); }
    100% { box-shadow: 0 0 0 0 rgba(139,92,246,0); }
}
@keyframes ai-gradient {
    0% { background-position: 0% 50%; }
    50% { background-position: 100% 50%; }
    100% { background-position: 0% 50%; }
}
@keyframes dot-flash {
    0%, 80%, 100% { opacity: 0.2; }
    40% { opacity: 1; }
}
</style>
"""

st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


def render_ai_status(placeholder, message: str):
    placeholder.markdown(
        f"""
        <div class="ai-status-card">
            <div class="ai-orb"></div>
            <div>
                <div class="ai-status-text">
                    IA analizando tus datos
                    <span class="dot-flashing"><span>.</span><span>.</span><span>.</span></span>
                </div>
                <div class="ai-status-sub">{message}</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

st.markdown(
    """
    <div class="hero">
        <h1>📊 Generador de Reportes Mensuales</h1>
        <p>Sube el Excel de tickets (Freshservice) y deja que la IA arme el reporte por ti.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

if "result" not in st.session_state:
    st.session_state.result = None
if "errors" not in st.session_state:
    st.session_state.errors = None

with st.sidebar:
    st.subheader("⚙️ Configuración")
    api_key_set = bool(os.getenv("GEMINI_API_KEY"))
    if api_key_set:
        st.success("Clave de Gemini detectada")
    else:
        st.warning("No hay GEMINI_API_KEY en .env — se usarán textos por defecto")

    target_items = load_target_items()
    if target_items:
        with st.expander(f"Plataformas rastreadas ({len(target_items)})"):
            for item in target_items:
                st.write(f"- {item['name']}")

    st.divider()
    st.caption("Plantilla Word: `inputs/template_report.docx`")
    st.caption("Reportes generados: carpeta `outputs/`")

col_left, col_right = st.columns([1.3, 1])

with col_left:
    st.markdown("### 1. Sube tu archivo")
    uploaded_file = st.file_uploader(
        "Arrastra o selecciona el Excel exportado de Freshservice",
        type=["xlsx"],
    )

    generate_clicked = st.button(
        "🚀 Generar reporte",
        type="primary",
        disabled=uploaded_file is None,
        use_container_width=True,
    )

with col_right:
    st.markdown("### 2. Estado del análisis")
    status_placeholder = st.empty()
    progress_placeholder = st.empty()

    if not uploaded_file and not st.session_state.result and not st.session_state.errors:
        status_placeholder.info("Esperando que subas un archivo Excel para comenzar.")

if generate_clicked and uploaded_file is not None:
    st.session_state.result = None
    st.session_state.errors = None

    os.makedirs("inputs", exist_ok=True)
    saved_path = os.path.join("inputs", uploaded_file.name)
    with open(saved_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    progress_bar = progress_placeholder.progress(0)
    render_ai_status(status_placeholder, "Iniciando análisis...")

    def on_progress(fraction: float, message: str):
        progress_bar.progress(min(int(fraction * 100), 100))
        render_ai_status(status_placeholder, message)

    try:
        result = generate_report(saved_path, on_progress=on_progress)
        st.session_state.result = result
        status_placeholder.success(f"¡Listo! Reporte de {result['month_label']} generado.")
    except ValidationError as e:
        st.session_state.errors = e.errors
        status_placeholder.error("Se encontraron errores en el Excel. Corrígelos y vuelve a intentarlo.")
    except Exception as e:
        st.session_state.errors = [f"Error inesperado: {e}"]
        status_placeholder.error("Ocurrió un error inesperado durante el análisis.")

if st.session_state.errors:
    st.markdown("### ❌ Errores encontrados")
    for err in st.session_state.errors:
        st.error(err)

if st.session_state.result:
    result = st.session_state.result
    metrics = result["metrics"]

    st.markdown("### 3. Resultados")

    def metric_row(items):
        cols = st.columns(len(items))
        for col, value, label, kind in items:
            col.markdown(
                f"""
                <div class="metric-card {kind}">
                    <div class="value">{value}</div>
                    <div class="label">{label}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    st.markdown('<div class="section-label">Números</div>', unsafe_allow_html=True)
    cols = st.columns(4)
    metric_row([
        (cols[0], metrics["cantidad_tickets"], "Tickets totales", "number"),
        (cols[1], metrics["resolved_closed_tickets"], "Resueltos / Cerrados", "number"),
        (cols[2], metrics["tickets_resolved_in_first_contact"], "Resueltos en 1er contacto", "number"),
        (cols[3], metrics["surveys_answered"], "Encuestas respondidas", "number"),
    ])

    st.markdown('<div class="section-label">Porcentajes</div>', unsafe_allow_html=True)
    cols = st.columns(4)
    metric_row([
        (cols[0], metrics["sla_resolution"], "SLA cumplido (Resolución)", "percent"),
        (cols[1], f"{metrics['fr_sla_percent']:.2f}%", "SLA cumplido (1a Respuesta)", "percent"),
        (cols[2], metrics["percentage_of_resolved_tickets_in_fr"], "% Resueltos 1er contacto", "percent"),
        (cols[3], metrics["percentage_of_surveys_answered"], "% Encuestas respondidas", "percent"),
    ])

    st.markdown('<div class="section-label">Tiempos promedio</div>', unsafe_allow_html=True)
    cols = st.columns(4)
    metric_row([
        (cols[0], metrics["avg_fr_time"], "Primera respuesta", "number"),
        (cols[1], metrics["avg_resolution"], "Resolución", "number"),
        (cols[2], f"{metrics['percentage_of_resolved_tickets']}%", "% Resueltos/Cerrados", "percent"),
        (cols[3], f"{metrics['percentage_of_blank_tickets']}%", "% Sin plataforma asignada", "percent"),
    ])

    if metrics.get("target_items_display"):
        st.markdown('<div class="section-label">Tickets por plataforma</div>', unsafe_allow_html=True)
        table_rows = "".join(
            f"<tr><td>{item['name']}</td><td style='text-align:center'>{item['tickets']}</td>"
            f"<td style='text-align:center'>{item['percent']}%</td></tr>"
            for item in metrics["target_items_display"]
        )
        st.markdown(
            f"""
            <table style="width:100%; border-collapse:collapse;">
                <thead>
                    <tr style="border-bottom:2px solid rgba(120,120,120,0.25);">
                        <th style="text-align:left; padding:0.4rem 0.2rem;">Plataforma</th>
                        <th style="padding:0.4rem 0.2rem;">Tickets</th>
                        <th style="padding:0.4rem 0.2rem;">%</th>
                    </tr>
                </thead>
                <tbody>{table_rows}</tbody>
            </table>
            """,
            unsafe_allow_html=True,
        )

    st.write("")

    with open(result["output_path"], "rb") as f:
        st.download_button(
            "⬇️ Descargar reporte Word",
            data=f.read(),
            file_name=os.path.basename(result["output_path"]),
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True,
        )

    graph_files = sorted(glob.glob(os.path.join("inputs", "tickets_by_*.png")))
    if graph_files:
        with st.expander("Ver gráficas generadas"):
            grid = st.columns(3)
            for i, path in enumerate(graph_files):
                grid[i % 3].image(path, use_container_width=True)
