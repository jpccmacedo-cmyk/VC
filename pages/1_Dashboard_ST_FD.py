import streamlit as st
import streamlit.components.v1 as components
from pathlib import Path

st.set_page_config(
    page_title="Dashboard ST & FD",
    page_icon="📈",
    layout="wide"
)

st.title("📈 Dashboard ST & FD")

st.markdown(
    """
    Faça upload do Excel dentro do dashboard abaixo para visualizar os indicadores ST e FD.
    """
)

html_path = Path(__file__).resolve().parent.parent / "dashboard_st_fd.html"

if not html_path.exists():
    st.error(
        "Arquivo dashboard_st_fd.html não encontrado. "
        "Verifique se ele está na mesma pasta do app.py."
    )
else:
    html_dashboard = html_path.read_text(encoding="utf-8")

    components.html(
        html_dashboard,
        height=1200,
        scrolling=True
    )
