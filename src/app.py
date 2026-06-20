"""Main application entry point for the Shoes Production Dashboard."""

import pandas as pd
import streamlit as st

from src.config.settings import DEFAULT_APP_CONFIG
from src.utils.data_loader import load_data, calculate_kpis
from src.components.kpi import render_kpi_sidebar, render_all_kpis
from src.components.charts import (
    render_oee_components_chart,
    render_oee_trend_chart,
    render_production_metrics_charts,
    render_material_analysis_charts,
)
from src.components.step_detail import render_step_detail_tab
from src.components.data_table import render_data_table, render_summary_stats


def configure_page() -> None:
    """Configure Streamlit page settings."""
    st.set_page_config(
        page_title=DEFAULT_APP_CONFIG.page_title,
        page_icon=DEFAULT_APP_CONFIG.page_icon,
        layout=DEFAULT_APP_CONFIG.layout,
        initial_sidebar_state=DEFAULT_APP_CONFIG.initial_sidebar_state,
        menu_items={
            'Get Help': 'https://www.streamlit.io/community',
            'Report a bug': "https://www.streamlit.io/community",
            'About': "# Shoes Production Dashboard\nComprehensive OEE dashboard for production monitoring."
        }
    )

    st.markdown("""
    <style>
        .stApp {
            background-color: #0e1117;
            color: #fafafa;
        }
        .stMetric {
            background-color: #262730;
            padding: 1rem;
            border-radius: 0.5rem;
            border: 1px solid #464853;
        }
        .stTabs [data-baseweb="tab-list"] {
            background-color: #262730;
        }
        .stTabs [data-baseweb="tab"] {
            background-color: #262730;
            color: #fafafa;
        }
        .stSelectbox > div > div {
            background-color: #262730;
            color: #fafafa;
        }
        .stMultiSelect > div > div {
            background-color: #262730;
            color: #fafafa;
        }
        .stSlider > div > div {
            background-color: #262730;
        }
    </style>
    """, unsafe_allow_html=True)


def render_header() -> None:
    """Render application header."""
    st.title(f"{DEFAULT_APP_CONFIG.page_icon} {DEFAULT_APP_CONFIG.page_title}")
    st.markdown("""
    **Author:** CHEN JUNG-LU  
    **Email:** E1582484@u.nus.edu 
    """)
    st.markdown("---")


def render_charts_tabs(df: pd.DataFrame, oee_target: float) -> None:
    """Render all chart tabs."""
    st.header("📈 Performance Analytics")
    
    tab1, tab2, tab3, tab4 = st.tabs([
        "OEE Analysis", 
        "Production Metrics", 
        "Material Analysis", 
        "Step-by-Step View"
    ])
    
    with tab1:
        render_oee_components_chart(df)
        render_oee_trend_chart(df, oee_target)
    
    with tab2:
        render_production_metrics_charts(df)
    
    with tab3:
        render_material_analysis_charts(df)
    
    with tab4:
        render_step_detail_tab(df)


def main() -> None:
    """Main application function."""
    configure_page()
    render_header()
    
    # Load data
    df, _ = load_data()
    
    # Get KPI targets from sidebar
    targets = render_kpi_sidebar()
    
    # Calculate KPIs
    kpis = calculate_kpis(df)
    
    # Render KPIs
    render_all_kpis(kpis, targets)
    
    # Render charts
    render_charts_tabs(df, targets['oee_target'])
    
    st.markdown("---")
    
    # Render data table
    filtered_df = render_data_table(df)
    render_summary_stats(filtered_df)


if __name__ == "__main__":
    main()