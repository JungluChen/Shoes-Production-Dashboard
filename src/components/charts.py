"""Chart components for the Shoes Production Dashboard."""

import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import pandas as pd
import streamlit as st
from typing import Optional, Union

from src.config.settings import DEFAULT_CHART_CONFIG, STATUS_COLORS, GAUGE_THRESHOLDS


def render_oee_components_chart(df: pd.DataFrame) -> None:
    """Render OEE components bar charts in a 2x2 grid."""
    fig_oee = make_subplots(
        rows=2, cols=2,
        subplot_titles=('OEE by Step', 'Availability by Step', 'Performance by Step', 'Quality by Step'),
        specs=[[{"secondary_y": False}, {"secondary_y": False}],
               [{"secondary_y": False}, {"secondary_y": False}]]
    )
    
    # OEE
    fig_oee.add_trace(
        go.Bar(x=df['Steps'], y=df['OEE'], name='OEE', marker_color=DEFAULT_CHART_CONFIG.color_palette['primary']),
        row=1, col=1
    )
    
    # Availability
    fig_oee.add_trace(
        go.Bar(x=df['Steps'], y=df['Availability'], name='Availability', marker_color=DEFAULT_CHART_CONFIG.color_palette['secondary']),
        row=1, col=2
    )
    
    # Performance
    fig_oee.add_trace(
        go.Bar(x=df['Steps'], y=df['Performance'], name='Performance', marker_color=DEFAULT_CHART_CONFIG.color_palette['success']),
        row=2, col=1
    )
    
    # Quality
    fig_oee.add_trace(
        go.Bar(x=df['Steps'], y=df['Quality'], name='Quality', marker_color=DEFAULT_CHART_CONFIG.color_palette['danger']),
        row=2, col=2
    )
    
    fig_oee.update_layout(
        height=DEFAULT_CHART_CONFIG.height_large,
        showlegend=False,
        title_text="OEE Components Analysis"
    )
    fig_oee.update_yaxes(range=[0, 1], tickformat='.0%')
    st.plotly_chart(fig_oee, use_container_width=True)


def render_oee_trend_chart(df: pd.DataFrame, target: float = 0.85) -> None:
    """Render OEE trend line chart across production steps."""
    fig_trend = px.line(
        df, x='Steps', y='OEE',
        title='OEE Trend Across Production Steps',
        markers=True, line_shape='spline'
    )
    fig_trend.add_hline(
        y=target, line_dash="dash", line_color="red",
        annotation_text=f"Target: {target:.0%}"
    )
    fig_trend.update_yaxes(tickformat='.0%', range=[0, 1])
    st.plotly_chart(fig_trend, use_container_width=True)


def render_production_metrics_charts(df: pd.DataFrame) -> None:
    """Render production metrics charts (units, status, time)."""
    col1, col2 = st.columns(2)
    
    with col1:
        # Units Produced by Step
        fig_units = px.bar(
            df, x='Steps', y='Units produced',
            title='Units Produced by Step',
            color='Units produced',
            color_continuous_scale='Blues'
        )
        fig_units.update_layout(height=DEFAULT_CHART_CONFIG.height_default)
        st.plotly_chart(fig_units, use_container_width=True)
        
    with col2:
        # Running Status Distribution
        status_counts = df['Running Status'].value_counts()
        fig_status = px.pie(
            values=status_counts.values, names=status_counts.index,
            title='Running Status Distribution',
            color_discrete_map=STATUS_COLORS
        )
        fig_status.update_layout(height=DEFAULT_CHART_CONFIG.height_default)
        st.plotly_chart(fig_status, use_container_width=True)
    
    # Time Analysis
    fig_time = go.Figure()
    fig_time.add_trace(go.Bar(
        name='Expected Time', x=df['Steps'], y=df['Expected Lot Run Time (Hours)'],
        marker_color=DEFAULT_CHART_CONFIG.color_palette['info']
    ))
    fig_time.add_trace(go.Bar(
        name='Actual Time', x=df['Steps'], y=df['Current Lot Run Time (Hours)'],
        marker_color=DEFAULT_CHART_CONFIG.color_palette['primary']
    ))
    fig_time.add_trace(go.Bar(
        name='Down Time', x=df['Steps'], y=df['Process Down time'],
        marker_color=DEFAULT_CHART_CONFIG.color_palette['danger']
    ))
    fig_time.update_layout(
        barmode='group',
        title='Time Analysis by Step',
        height=DEFAULT_CHART_CONFIG.height_default
    )
    st.plotly_chart(fig_time, use_container_width=True)


def render_material_analysis_charts(df: pd.DataFrame) -> None:
    """Render material analysis charts (usage, waste, efficiency)."""
    col1, col2 = st.columns(2)
    
    with col1:
        # Material Usage
        fig_material = px.bar(
            df, x='Steps', y='Material Used (KG)',
            title='Material Usage by Step',
            color='Material Used (KG)',
            color_continuous_scale='Greens'
        )
        fig_material.update_layout(height=DEFAULT_CHART_CONFIG.height_default)
        st.plotly_chart(fig_material, use_container_width=True)
        
    with col2:
        # Waste Analysis
        fig_waste = px.bar(
            df, x='Steps', y='Waste Materials (KG)',
            title='Waste Materials by Step',
            color='Waste Materials (KG)',
            color_continuous_scale='Reds'
        )
        fig_waste.update_layout(height=DEFAULT_CHART_CONFIG.height_default)
        st.plotly_chart(fig_waste, use_container_width=True)
    
    # Material Efficiency
    fig_efficiency = px.line(
        df, x='Steps', y='Material_Efficiency',
        title='Material Efficiency by Step',
        markers=True
    )
    fig_efficiency.update_yaxes(tickformat='.0%')
    st.plotly_chart(fig_efficiency, use_container_width=True)


def render_gauge_chart(metric_name: str, value: float, column) -> None:
    """Render a gauge chart for a single metric."""
    thresholds = GAUGE_THRESHOLDS.get(metric_name, GAUGE_THRESHOLDS['OEE'])
    
    fig_gauge = go.Figure(go.Indicator(
        mode="gauge+number+delta",
        value=value * 100,
        domain={'x': [0, 1], 'y': [0, 1]},
        title={'text': metric_name},
        delta={'reference': thresholds['target']},
        gauge={
            'axis': {'range': [None, 100]},
            'bar': {'color': "darkblue"},
            'steps': [
                {'range': [0, thresholds['warning']], 'color': "lightgray"},
                {'range': [thresholds['warning'], thresholds['good']], 'color': "yellow"},
                {'range': [thresholds['good'], 100], 'color': "green"}
            ],
            'threshold': {
                'line': {'color': "red", 'width': 4},
                'thickness': 0.75,
                'value': thresholds['target']
            }
        }
    ))
    fig_gauge.update_layout(height=300)
    column.plotly_chart(fig_gauge, use_container_width=True)


def render_step_detail_charts(step_data: Union[pd.Series, dict]) -> None:
    """Render gauge charts for a specific step's OEE metrics."""
    col1, col2, col3, col4 = st.columns(4)
    
    metrics = [
        ('OEE', step_data['OEE'], col1),
        ('Availability', step_data['Availability'], col2),
        ('Performance', step_data['Performance'], col3),
        ('Quality', step_data['Quality'], col4)
    ]
    
    for metric_name, value, col in metrics:
        with col:
            render_gauge_chart(metric_name, value, col)