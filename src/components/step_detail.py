"""Step detail view component for the Shoes Production Dashboard."""

import streamlit as st
import pandas as pd

from src.components.charts import render_step_detail_charts


def render_step_selector(df: pd.DataFrame) -> pd.Series:
    """Render step selector and return selected step data."""
    selected_step = st.selectbox(
        'Select a Step for Detailed View:',
        df['Steps'].unique()
    )
    
    step_data = df[df['Steps'] == selected_step].iloc[0]
    return step_data


def render_step_info(step_data: pd.Series) -> None:
    """Render step information panels."""
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.subheader("📋 Step Information")
        st.write(f"**Step:** {step_data['Steps']}")
        st.write(f"**Running Status:** {step_data['Running Status']}")
        st.write(f"**Current Lot Number:** {step_data['Current lot Number']}")
        st.write(f"**Units Produced:** {step_data['Units produced']:,.0f}")
        
    with col2:
        st.subheader("⏱️ Time Metrics")
        st.write(f"**Expected Run Time:** {step_data['Expected Lot Run Time (Hours)']:.1f} hours")
        st.write(f"**Actual Run Time:** {step_data['Current Lot Run Time (Hours)']:.1f} hours")
        st.write(f"**Process Down Time:** {step_data['Process Down time']:.1f} hours")
        st.write(f"**Failure Rate:** {step_data['Failure Rate']:.1%}")
        
    with col3:
        st.subheader("🏭 Material Metrics")
        st.write(f"**Material Used:** {step_data['Material Used (KG)']:.1f} KG")
        st.write(f"**Waste Materials:** {step_data['Waste Materials (KG)']:.1f} KG")
        waste_rate = (step_data['Waste Materials (KG)'] / step_data['Material Used (KG)']) * 100 if step_data['Material Used (KG)'] > 0 else 0
        st.write(f"**Waste Rate:** {waste_rate:.1f}%")


def render_step_detail_tab(df: pd.DataFrame) -> None:
    """Render the complete step-by-step detail tab."""
    step_data = render_step_selector(df)
    render_step_info(step_data)
    render_step_detail_charts(step_data)