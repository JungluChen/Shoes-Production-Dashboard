"""Configuration settings for the Shoes Production Dashboard."""

from dataclasses import dataclass
from typing import Dict, Any


@dataclass
class KPITargets:
    """KPI target thresholds for the dashboard."""
    oee_target: float = 0.85
    availability_target: float = 0.90
    performance_target: float = 0.95
    quality_target: float = 0.99
    time_efficiency_target: float = 0.95
    material_efficiency_target: float = 0.90
    waste_target: float = 0.05
    productivity_target: float = 100.0
    downtime_target: float = 1.0
    running_ratio_target: float = 0.80
    active_ratio_target: float = 0.90


@dataclass
class ChartConfig:
    """Chart configuration constants."""
    height_default: int = 400
    height_large: int = 600
    color_palette: Dict[str, str] = None
    
    def __post_init__(self):
        if self.color_palette is None:
            self.color_palette = {
                'primary': '#1f77b4',
                'secondary': '#ff7f0e',
                'success': '#2ca02c',
                'danger': '#d62728',
                'warning': '#ffbb78',
                'info': '#98df8a',
            }


@dataclass
class AppConfig:
    """Main application configuration."""
    page_title: str = "Shoes Production Dashboard"
    page_icon: str = "👟"
    layout: str = "wide"
    initial_sidebar_state: str = "expanded"
    data_file: str = "data/AS2 5001.xlsx"
    cache_ttl: int = 3600  # seconds


# Default configurations
DEFAULT_KPI_TARGETS = KPITargets()
DEFAULT_CHART_CONFIG = ChartConfig()
DEFAULT_APP_CONFIG = AppConfig()


# Column name mappings for consistency
COLUMN_MAPPING = {
    'Steps': 'step',
    'Running Status': 'running_status',
    'Current lot Number': 'lot_number',
    'Material Used (KG)': 'material_used_kg',
    'Waste Materials (KG)': 'waste_materials_kg',
    'Current Lot Run Time (Hours)': 'actual_run_time_hrs',
    'Expected Lot Run Time (Hours)': 'expected_run_time_hrs',
    'Process Down time': 'downtime_hrs',
    'Failure Rate': 'failure_rate',
    'Units produced': 'units_produced',
}

# OEE component columns
OEE_COLUMNS = ['Availability', 'Performance', 'Quality', 'OEE']

# Display column order for data table
DISPLAY_COLUMNS = [
    'Steps', 'OEE', 'Availability', 'Performance', 'Quality',
    'Running Status', 'Current lot Number', 'Units produced',
    'Material Used (KG)', 'Waste Materials (KG)',
    'Current Lot Run Time (Hours)', 'Expected Lot Run Time (Hours)',
    'Process Down time', 'Failure Rate'
]

# Status color mapping
STATUS_COLORS = {
    'Running': '#2ca02c',
    'Stopped': '#d62728',
    'Idle': '#ff7f0e',
}

# Gauge chart thresholds
GAUGE_THRESHOLDS = {
    'OEE': {'target': 85, 'warning': 50, 'good': 80},
    'Availability': {'target': 90, 'warning': 50, 'good': 80},
    'Performance': {'target': 90, 'warning': 50, 'good': 80},
    'Quality': {'target': 90, 'warning': 50, 'good': 80},
}