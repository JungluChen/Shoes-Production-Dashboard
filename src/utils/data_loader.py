"""Data loading and preprocessing utilities."""

import pandas as pd
import streamlit as st
from typing import Tuple, Dict, Any

from src.config.settings import DEFAULT_APP_CONFIG, COLUMN_MAPPING, OEE_COLUMNS


@st.cache_data(ttl=DEFAULT_APP_CONFIG.cache_ttl)
def load_data(filepath: str = None) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    """
    Load and preprocess production data from Excel file.
    
    Args:
        filepath: Path to the Excel file. Uses default if not provided.
        
    Returns:
        Tuple of (processed DataFrame, data dictionary for quick access)
    """
    if filepath is None:
        filepath = DEFAULT_APP_CONFIG.data_file
    
    df = pd.read_excel(filepath)
    
    # Compute OEE components
    df["Availability"] = df["Current Lot Run Time (Hours)"] / df["Expected Lot Run Time (Hours)"]
    df["Performance"] = (df["Expected Lot Run Time (Hours)"] - df["Process Down time"]) / df["Expected Lot Run Time (Hours)"]
    df["Quality"] = 1 - df["Failure Rate"]
    
    # Clamp values to valid [0,1] range
    for col in ["Availability", "Performance", "Quality"]:
        df[col] = df[col].clip(0, 1)
    
    # Calculate OEE
    df["OEE"] = df["Availability"] * df["Performance"] * df["Quality"]
    df["OEE"] = df["OEE"].fillna(0)
    
    # Calculate material efficiency
    df["Material_Efficiency"] = (
        df["Material Used (KG)"] - df["Waste Materials (KG)"]
    ) / df["Material Used (KG)"]
    df["Material_Efficiency"] = df["Material_Efficiency"].clip(0, 1).fillna(1)
    
    # Create dictionary for quick access
    data_dict = {
        'Steps': df['Steps'].to_dict(),
        'Running_Status': df['Running Status'].to_dict(),
        'Current_Lot_Number': df['Current lot Number'].to_dict(),
        'Material_Used': df['Material Used (KG)'].to_dict(),
        'Waste_Materials': df['Waste Materials (KG)'].to_dict(),
        'Current_Lot_Run_Time': df['Current Lot Run Time (Hours)'].to_dict(),
        'Expected_Lot_Run_Time': df['Expected Lot Run Time (Hours)'].to_dict(),
        'Process_Down_Time': df['Process Down time'].to_dict(),
        'Failure_Rate': df['Failure Rate'].to_dict(),
        'Units_Produced': df['Units produced'].to_dict(),
        'Availability': df['Availability'].to_dict(),
        'Performance': df['Performance'].to_dict(),
        'Quality': df['Quality'].to_dict(),
        'OEE': df['OEE'].to_dict(),
        'Material_Efficiency': df['Material_Efficiency'].to_dict(),
    }
    
    return df, data_dict


def calculate_kpis(df: pd.DataFrame) -> Dict[str, float]:
    """
    Calculate overall KPIs from the production data.
    
    Args:
        df: Processed DataFrame with OEE columns.
        
    Returns:
        Dictionary of calculated KPIs.
    """
    total_material = df['Material Used (KG)'].sum()
    total_waste = df['Waste Materials (KG)'].sum()
    total_actual_time = df['Current Lot Run Time (Hours)'].sum()
    total_expected_time = df['Expected Lot Run Time (Hours)'].sum()
    
    kpis = {
        'avg_oee': df['OEE'].mean(),
        'avg_availability': df['Availability'].mean(),
        'avg_performance': df['Performance'].mean(),
        'avg_quality': df['Quality'].mean(),
        'total_units': df['Units produced'].sum(),
        'total_material_used': total_material,
        'total_waste': total_waste,
        'waste_percentage': (total_waste / total_material * 100) if total_material > 0 else 0,
        'total_downtime': df['Process Down time'].sum(),
        'avg_downtime': df['Process Down time'].mean(),
        'time_efficiency': (total_actual_time / total_expected_time * 100) if total_expected_time > 0 else 0,
        'material_efficiency': ((total_material - total_waste) / total_material * 100) if total_material > 0 else 0,
        'running_steps': len(df[df['Running Status'] == 'Running']),
        'stopped_steps': len(df[df['Running Status'] == 'Stopped']),
        'idle_steps': len(df[df['Running Status'] == 'Idle']),
        'total_steps': len(df),
        'productivity_rate': df['Units produced'].sum() / total_actual_time if total_actual_time > 0 else 0,
        'avg_failure_rate': df['Failure Rate'].mean(),
    }
    
    return kpis


def filter_data(df: pd.DataFrame, status_filter: list = None, oee_threshold: float = 0) -> pd.DataFrame:
    """
    Filter DataFrame by running status and OEE threshold.
    
    Args:
        df: DataFrame to filter.
        status_filter: List of running statuses to include.
        oee_threshold: Minimum OEE value (0-1).
        
    Returns:
        Filtered DataFrame.
    """
    if status_filter is None:
        status_filter = df['Running Status'].unique().tolist()
    
    filtered = df[
        (df['Running Status'].isin(status_filter)) &
        (df['OEE'] >= oee_threshold)
    ].copy()
    
    return filtered