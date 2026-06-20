"""Data table components for the Shoes Production Dashboard."""

import streamlit as st
import pandas as pd
from typing import List, Optional

from src.config.settings import DISPLAY_COLUMNS, OEE_COLUMNS
from src.utils.data_loader import filter_data


def render_data_filters(df: pd.DataFrame) -> tuple:
    """Render filter controls and return filter values."""
    col1, col2 = st.columns(2)
    
    with col1:
        status_filter = st.multiselect(
            'Filter by Running Status:',
            options=df['Running Status'].unique(),
            default=df['Running Status'].unique()
        )
        
    with col2:
        oee_threshold = st.slider(
            'Minimum OEE Threshold:',
            min_value=0,
            max_value=100,
            value=0,
            step=5,
            format='%d%%'
        ) / 100
    
    return status_filter, oee_threshold


def format_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Apply formatting to DataFrame for display."""
    display_df = df[DISPLAY_COLUMNS].copy()
    
    format_dict = {
        'OEE': '{:.1%}',
        'Availability': '{:.1%}',
        'Performance': '{:.1%}',
        'Quality': '{:.1%}',
        'Failure Rate': '{:.1%}',
        'Material Used (KG)': '{:.1f}',
        'Waste Materials (KG)': '{:.1f}',
        'Current Lot Run Time (Hours)': '{:.1f}',
        'Expected Lot Run Time (Hours)': '{:.1f}',
        'Process Down time': '{:.1f}'
    }
    
    styled = display_df.style.format(format_dict).background_gradient(subset=OEE_COLUMNS)
    
    return styled


def render_data_table(df: pd.DataFrame) -> pd.DataFrame:
    """Render filtered data table with formatting."""
    st.header("📋 Detailed Data View")
    
    status_filter, oee_threshold = render_data_filters(df)
    
    filtered_df = filter_data(df, status_filter, oee_threshold)
    
    styled_df = format_dataframe(filtered_df)
    
    st.dataframe(styled_df, use_container_width=True)
    
    return filtered_df


def render_summary_stats(filtered_df: pd.DataFrame) -> None:
    """Render summary statistics for filtered data."""
    st.subheader("📊 Summary Statistics")
    col1, col2 = st.columns(2)
    
    with col1:
        st.write("**Filtered Data Summary:**")
        if len(filtered_df) > 0:
            st.write(f"- Total Steps: {len(filtered_df)}")
            st.write(f"- Average OEE: {filtered_df['OEE'].mean():.1%}")
            st.write(f"- Total Units Produced: {filtered_df['Units produced'].sum():,.0f}")
            st.write(f"- Total Material Used: {filtered_df['Material Used (KG)'].sum():,.1f} KG")
        else:
            st.write("- No steps match the current filter criteria")
            st.write("- Please adjust the Running Status or OEE threshold filters")
    
    with col2:
        st.write("**Performance Ranges:**")
        if len(filtered_df) > 0:
            st.write(f"- OEE Range: {filtered_df['OEE'].min():.1%} - {filtered_df['OEE'].max():.1%}")
            best_step = filtered_df.loc[filtered_df['OEE'].idxmax(), 'Steps']
            worst_step = filtered_df.loc[filtered_df['OEE'].idxmin(), 'Steps']
            st.write(f"- Best Performing Step: {best_step}")
            st.write(f"- Lowest Performing Step: {worst_step}")
            waste_rate = (filtered_df['Waste Materials (KG)'].sum() / filtered_df['Material Used (KG)'].sum() * 100)
            st.write(f"- Average Waste Rate: {waste_rate:.1f}%")
        else:
            st.write("- No data matches the current filter criteria")
            st.write("- Please adjust the filters to see performance ranges")