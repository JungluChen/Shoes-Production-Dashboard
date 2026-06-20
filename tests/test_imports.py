"""Test imports for the Shoes Production Dashboard."""

import sys
sys.path.insert(0, 'src')

def test_config_import():
    from src.config.settings import (
        DEFAULT_APP_CONFIG, DEFAULT_KPI_TARGETS, DEFAULT_CHART_CONFIG,
        COLUMN_MAPPING, OEE_COLUMNS, DISPLAY_COLUMNS, STATUS_COLORS, GAUGE_THRESHOLDS
    )
    assert DEFAULT_APP_CONFIG.page_title == "Shoes Production Dashboard"
    assert DEFAULT_KPI_TARGETS.oee_target == 0.85

def test_utils_import():
    from src.utils.data_loader import load_data, calculate_kpis, filter_data
    assert callable(load_data)
    assert callable(calculate_kpis)
    assert callable(filter_data)

def test_components_import():
    from src.components.charts import (
        render_oee_components_chart, render_oee_trend_chart,
        render_production_metrics_charts, render_material_analysis_charts,
        render_gauge_chart, render_step_detail_charts
    )
    from src.components.kpi import (
        render_kpi_sidebar, render_core_oee_kpis,
        render_production_kpis, render_material_quality_kpis,
        render_status_kpis, render_all_kpis
    )
    from src.components.step_detail import (
        render_step_selector, render_step_info, render_step_detail_tab
    )
    from src.components.data_table import (
        render_data_filters, format_dataframe,
        render_data_table, render_summary_stats
    )
    assert all(callable(f) for f in [
        render_oee_components_chart, render_oee_trend_chart,
        render_production_metrics_charts, render_material_analysis_charts,
        render_gauge_chart, render_step_detail_charts,
        render_kpi_sidebar, render_core_oee_kpis,
        render_production_kpis, render_material_quality_kpis,
        render_status_kpis, render_all_kpis,
        render_step_selector, render_step_info, render_step_detail_tab,
        render_data_filters, format_dataframe,
        render_data_table, render_summary_stats
    ])

def test_app_import():
    from src.app import main, configure_page, render_header, render_charts_tabs
    assert callable(main)
    assert callable(configure_page)
    assert callable(render_header)
    assert callable(render_charts_tabs)

if __name__ == "__main__":
    test_config_import()
    test_utils_import()
    test_components_import()
    test_app_import()
    print("All imports successful!")