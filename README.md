# Shoes Production Dashboard

A comprehensive **Overall Equipment Effectiveness (OEE)** monitoring dashboard for shoe production lines. Built with Streamlit, Plotly, and Pandas for real-time production analytics and visualization.

## Features

### Core OEE Analytics
- **OEE Calculation**: Automatically computes Availability, Performance, Quality, and Overall OEE
- **Real-time KPI Monitoring**: Track OEE, Availability, Performance, Quality with configurable targets
- **Production Metrics**: Units produced, production time, downtime, time efficiency
- **Material Analysis**: Material usage, waste tracking, material efficiency rates
- **Operational Status**: Running, stopped, and idle step monitoring

### Visualization
- **OEE Components Breakdown**: 2x2 grid of bar charts for each OEE component
- **OEE Trend Analysis**: Line chart with target threshold across production steps
- **Production Metrics**: Units by step, status distribution, time analysis
- **Material Analysis**: Usage, waste, and efficiency charts
- **Step-by-Step Detail View**: Interactive gauges for individual step performance

### Interactive Features
- **Configurable KPI Targets**: Adjustable thresholds via sidebar
- **Data Filtering**: Filter by running status and OEE threshold
- **Formatted Data Table**: Sortable, styled table with gradient highlighting
- **Summary Statistics**: Best/worst performing steps, performance ranges

## Project Structure

```
Shoes-Production-Dashboard/
├── app.py                      # Entry point
├── requirements.txt            # Python dependencies
├── README.md                   # This file
├── data/
│   └── AS2 5001.xlsx          # Production data (Excel)
├── src/
│   ├── __init__.py            # Package initialization
│   ├── app.py                 # Main Streamlit application
│   ├── config/
│   │   ├── __init__.py
│   │   └── settings.py        # Configuration constants & targets
│   ├── utils/
│   │   ├── __init__.py
│   │   └── data_loader.py     # Data loading, KPI calculation, filtering
│   └── components/
│       ├── __init__.py
│       ├── charts.py          # Plotly chart components
│       ├── kpi.py             # KPI metric components
│       ├── step_detail.py     # Step detail view
│       └── data_table.py      # Data table with filters
└── tests/
    └── test_imports.py        # Import verification tests
```

## Quick Start

### Prerequisites
- Python 3.10+
- pip package manager

### Installation

```bash
# Clone the repository
git clone https://github.com/JungluChen/Shoes-Production-Dashboard.git
cd Shoes-Production-Dashboard

# Create virtual environment (recommended)
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

### Running the Dashboard

```bash
# Run the Streamlit app
streamlit run app.py
```

The dashboard will open at `http://localhost:8501`

## Data Format

The dashboard expects an Excel file (`data/AS2 5001.xlsx`) with the following columns:

| Column | Description | Type |
|--------|-------------|------|
| Steps | Production step name | String |
| Running Status | Current status (Running/Stopped/Idle) | String |
| Current lot Number | Active lot identifier | Integer |
| Material Used (KG) | Material consumed | Float |
| Waste Materials (KG) | Waste generated | Float |
| Current Lot Run Time (Hours) | Actual run time | Float |
| Expected Lot Run Time (Hours) | Planned run time | Float |
| Process Down time | Downtime hours | Float |
| Failure Rate | Defect rate (0-1) | Float |
| Units produced | Output quantity | Integer |

### OEE Calculations

The dashboard automatically computes:

- **Availability** = Actual Run Time / Expected Run Time
- **Performance** = (Expected Run Time - Downtime) / Expected Run Time
- **Quality** = 1 - Failure Rate
- **OEE** = Availability × Performance × Quality
- **Material Efficiency** = (Material Used - Waste) / Material Used

All values are clamped to [0, 1] range.

## Configuration

### KPI Targets (Sidebar)
Adjust targets in the sidebar:
- OEE Target (default: 85%)
- Availability Target (default: 90%)
- Performance Target (default: 95%)
- Quality Target (default: 99%)
- Time Efficiency Target (default: 95%)
- Material Efficiency Target (default: 90%)
- Waste Target (default: 5%)

### Application Settings
Modify `src/config/settings.py` for:
- Page title, icon, layout
- Data file path
- Cache TTL
- Chart colors and dimensions
- Gauge thresholds

## Usage Guide

### 1. Overview Tab
View all KPIs at a glance. Metrics show delta from targets (green = meeting target, red = below target).

### 2. Performance Analytics
- **OEE Analysis**: Component breakdown and trend
- **Production Metrics**: Units, status, time analysis
- **Material Analysis**: Usage, waste, efficiency trends
- **Step-by-Step View**: Select a step for detailed gauge charts

### 3. Data Table
- Filter by running status
- Set minimum OEE threshold
- View formatted data with gradient highlighting
- See summary statistics for filtered data

## Testing

```bash
# Run import tests
python -m pytest tests/ -v
```

## Technology Stack

- **Streamlit** - Web application framework
- **Plotly** - Interactive visualizations
- **Pandas** - Data manipulation
- **NumPy** - Numerical computations
- **OpenPyXL** - Excel file reading

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit changes (`git commit -m 'Add amazing feature'`)
4. Push to branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Author

**CHEN JUNG-LU**  
Email: E1582484@u.nus.edu

## Acknowledgments

- Streamlit community for the excellent framework
- Plotly for interactive charting capabilities
- Pandas team for data analysis library