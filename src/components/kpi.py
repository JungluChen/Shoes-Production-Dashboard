"""KPI components for the Shoes Production Dashboard."""

import streamlit as st
import pandas as pd
from typing import Dict, Any

from src.config.settings import DEFAULT_KPI_TARGETS, OEE_COLUMNS


def render_kpi_sidebar() -> Dict[str, float]:
    """Render KPI target inputs in sidebar and return targets."""
    st.sidebar.header("KPI Targets")
    
    targets = {
        'oee_target': st.sidebar.slider("OEE Target (%)", 0, 100, int(DEFAULT_KPI_TARGETS.oee_target * 100)) / 100,
        'availability_target': st.sidebar.slider("Availability Target (%)", 0, 100, int(DEFAULT_KPI_TARGETS.availability_target * 100)) / 100,
        'performance_target': st.sidebar.slider("Performance Target (%)", 0, 100, int(DEFAULT_KPI_TARGETS.performance_target * 100)) / 100,
        'quality_target': st.sidebar.slider("Quality Target (%)", 0, 100, int(DEFAULT_KPI_TARGETS.quality_target * 100)) / 100,
        'time_efficiency_target': st.sidebar.slider("Time Efficiency Target (%)", 0, 100, int(DEFAULT_KPI_TARGETS.time_efficiency_target * 100)) / 100,
        'material_efficiency_target': st.sidebar.slider("Material Efficiency Target (%)", 0, 100, int(DEFAULT_KPI_TARGETS.material_efficiency_target * 100)) / 100,
        'waste_target': st.sidebar.slider("Waste Target (%)", 0, 100, int(DEFAULT_KPI_TARGETS.waste_target * 100)) / 100,
    }
    
    return targets


def calculate_delta(current: float, target: float, higher_is_better: bool = True) -> str:
    """Calculate delta string for metric display."""
    diff = abs((current - target) * 100)
    if higher_is_better:
        return f"{diff:.1f}%" if current >= target else f"-{diff:.1f}%"
    else:
        return f"{diff:.1f}%" if current <= target else f"-{diff:.1f}%"


def render_core_oee_kpis(kpis: Dict[str, float], targets: Dict[str, float]) -> None:
    """Render core OEE metrics (OEE, Availability, Performance, Quality)."""
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric(
            label="🎯 Overall Equipment Effectiveness (OEE)",
            value=f"{kpis['avg_oee']:.1%}",
            delta=calculate_delta(kpis['avg_oee'], targets['oee_target']),
        )
        
    with col2:
        st.metric(
            label="⚡ Availability",
            value=f"{kpis['avg_availability']:.1%}",
            delta=calculate_delta(kpis['avg_availability'], targets['availability_target']),
        )
        
    with col3:
        st.metric(
            label="🚀 Performance",
            value=f"{kpis['avg_performance']:.1%}",
            delta=calculate_delta(kpis['avg_performance'], targets['performance_target']),
        )
        
    with col4:
        st.metric(
            label="✅ Quality",
            value=f"{kpis['avg_quality']:.1%}",
            delta=calculate_delta(kpis['avg_quality'], targets['quality_target']),
        )


def render_production_kpis(kpis: Dict[str, float], targets: Dict[str, float]) -> None:
    """Render production metrics (units, time, downtime, efficiency)."""
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric(
            label="📦 Total Units Produced",
            value=f"{kpis['total_units']:,.0f}",
        )
        
    with col2:
        st.metric(
            label="⏱️ Total Production Time",
            value=f"{kpis.get('total_actual_time', kpis['total_downtime'] + kpis['total_expected_time']):.1f} hrs",
        )
        
    with col3:
        st.metric(
            label="⏸️ Total Downtime",
            value=f"{kpis['total_downtime']:.1f} hrs",
        )
        
    with col4:
        st.metric(
            label="📈 Time Efficiency",
            value=f"{kpis['time_efficiency']:.1f}%",
            delta=calculate_delta(kpis['time_efficiency'] / 100, targets['time_efficiency_target']),
        )


def render_material_quality_kpis(kpis: Dict[str, float], targets: Dict[str, float]) -> None:
    """Render material and quality metrics."""
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric(
            label="🏭 Total Material Used",
            value=f"{kpis['total_material_used']:,.1f} KG",
        )
        
    with col2:
        st.metric(
            label="🗑️ Total Waste",
            value=f"{kpis['total_waste']:,.1f} KG",
        )
        
    with col3:
        st.metric(
            label="❌ Average Failure Rate",
            value=f"{kpis['avg_failure_rate']:.2%}",
            delta=calculate_delta(kpis['avg_failure_rate'], targets['waste_target'], higher_is_better=False),
        )
        
    with col4:
        st.metric(
            label="♻️ Material Efficiency",
            value=f"{kpis['material_efficiency']:.1f}%",
            delta=calculate_delta(kpis['material_efficiency'] / 100, targets['material_efficiency_target']),
        )


def render_status_kpis(kpis: Dict[str, float]) -> None:
    """Render operational status metrics."""
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric(label="🟢 Running Steps", value=f"{kpis['running_steps']}")
        
    with col2:
        st.metric(label="🔴 Stopped Steps", value=f"{kpis['stopped_steps']}")
        
    with col3:
        st.metric(label="🟡 Idle Steps", value=f"{kpis['idle_steps']}")
        
    with col4:
        st.metric(label="📊 Total Production Steps", value=f"{kpis['total_steps']}")


def render_all_kpis(kpis: Dict[str, float], targets: Dict[str, float]) -> None:
    """Render all KPI sections."""
    st.header("📊 Key Performance Indicators")
    
    render_core_oee_kpis(kpis, targets)
    st.markdown("")
    render_production_kpis(kpis, targets)
    st.markdown("")
    render_material_quality_kpis(kpis, targets)
    st.markdown("")
    render_status_kpis(kpis)
    st.markdown("---")