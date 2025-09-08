#%%
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# Only load Excel file once when program starts
@st.cache_data
def load_data():
    df = pd.read_excel('AS2 5001.xlsx')
    
    # --- Step 4: Compute OEE components ---
    df["Availability"] = df["Current Lot Run Time (Hours)"] / df["Expected Lot Run Time (Hours)"]
    df["Performance"] = (df["Expected Lot Run Time (Hours)"] - df["Process Down time"]) / df["Expected Lot Run Time (Hours)"]
    df["Quality"] = 1 - df["Failure Rate"]
    # Clamp values to valid [0,1] range
    df["Availability"] = df["Availability"].clip(0, 1)
    df["Performance"] = df["Performance"].clip(0, 1)
    df["Quality"] = df["Quality"].clip(0, 1)
    
    # --- Step 5: Calculate OEE ---
    df["OEE"] = df["Availability"] * df["Performance"] * df["Quality"]
    df["OEE"].fillna(0, inplace=True)

    # Create dictionary to store all initial and calculated data
    data_dict = {
        # Initial data columns
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
        # Claculate OEE
        'Availability': df['Availability'].to_dict(),
        'Performance': df['Performance'].to_dict(),
        'Quality': df['Quality'].to_dict(),
        'OEE': df['OEE'].to_dict()
    }
    return df, data_dict


df, data_dict = load_data()
#%%
st.set_page_config(page_title="Shoes Production Dashboard", page_icon="👟", layout="wide")
st.title("👟 Shoes Production Dashboard")
st.markdown("""
**Author:** CHEN JUNG-LU  
**Email:** E1582484@u.nus.edu 
""")
st.markdown("---")

# === KPI Section ===
st.header("📊 Key Performance Indicators")

# Calculate overall KPIs
avg_oee = df['OEE'].mean()
avg_availability = df['Availability'].mean()
avg_performance = df['Performance'].mean()
avg_quality = df['Quality'].mean()
total_units = df['Units produced'].sum()
total_material_used = df['Material Used (KG)'].sum()
total_waste = df['Waste Materials (KG)'].sum()
waste_percentage = (total_waste / total_material_used) * 100 if total_material_used > 0 else 0

# Display KPIs in columns
col1, col2, col3, col4 = st.columns(4)

with col1:
    st.metric(
        label="Overall Equipment Effectiveness (OEE)",
        value=f"{avg_oee:.1%}",
        delta=f"Target: 85%" if avg_oee >= 0.85 else f"Below Target: 85%"
    )
    
with col2:
    st.metric(
        label="Availability",
        value=f"{avg_availability:.1%}",
        delta=f"Target: 90%" if avg_availability >= 0.90 else f"Below Target: 90%"
    )
    
with col3:
    st.metric(
        label="Performance",
        value=f"{avg_performance:.1%}",
        delta=f"Target: 95%" if avg_performance >= 0.95 else f"Below Target: 95%"
    )
    
with col4:
    st.metric(
        label="Quality",
        value=f"{avg_quality:.1%}",
        delta=f"Target: 99%" if avg_quality >= 0.99 else f"Below Target: 99%"
    )

# Additional KPIs
col5, col6, col7, col8 = st.columns(4)

with col5:
    st.metric(
        label="Total Units Produced",
        value=f"{total_units:,.0f}"
    )
    
with col6:
    st.metric(
        label="Total Material Used (KG)",
        value=f"{total_material_used:,.1f}"
    )
    
with col7:
    st.metric(
        label="Total Waste (KG)",
        value=f"{total_waste:,.1f}"
    )
    
with col8:
    st.metric(
        label="Waste Percentage",
        value=f"{waste_percentage:.1f}%",
        delta="Good" if waste_percentage < 5 else "High"
    )

st.markdown("---")

# === Charts Section ===
st.header("📈 Performance Analytics")

# Create tabs for different chart categories
tab1, tab2, tab3, tab4 = st.tabs(["OEE Analysis", "Production Metrics", "Material Analysis", "Step-by-Step View"])

with tab1:
    # OEE Components Chart
    fig_oee = make_subplots(
        rows=2, cols=2,
        subplot_titles=('OEE by Step', 'Availability by Step', 'Performance by Step', 'Quality by Step'),
        specs=[[{"secondary_y": False}, {"secondary_y": False}],
               [{"secondary_y": False}, {"secondary_y": False}]]
    )
    
    # OEE
    fig_oee.add_trace(
        go.Bar(x=df['Steps'], y=df['OEE'], name='OEE', marker_color='#1f77b4'),
        row=1, col=1
    )
    
    # Availability
    fig_oee.add_trace(
        go.Bar(x=df['Steps'], y=df['Availability'], name='Availability', marker_color='#ff7f0e'),
        row=1, col=2
    )
    
    # Performance
    fig_oee.add_trace(
        go.Bar(x=df['Steps'], y=df['Performance'], name='Performance', marker_color='#2ca02c'),
        row=2, col=1
    )
    
    # Quality
    fig_oee.add_trace(
        go.Bar(x=df['Steps'], y=df['Quality'], name='Quality', marker_color='#d62728'),
        row=2, col=2
    )
    
    fig_oee.update_layout(height=600, showlegend=False, title_text="OEE Components Analysis")
    fig_oee.update_yaxes(range=[0, 1], tickformat='.0%')
    st.plotly_chart(fig_oee, use_container_width=True)
    
    # OEE Trend Line
    fig_trend = px.line(df, x='Steps', y='OEE', title='OEE Trend Across Production Steps',
                       markers=True, line_shape='spline')
    fig_trend.add_hline(y=0.85, line_dash="dash", line_color="red", 
                       annotation_text="Target: 85%")
    fig_trend.update_yaxes(tickformat='.0%', range=[0, 1])
    st.plotly_chart(fig_trend, use_container_width=True)

with tab2:
    col1, col2 = st.columns(2)
    
    with col1:
        # Units Produced by Step
        fig_units = px.bar(df, x='Steps', y='Units produced', 
                          title='Units Produced by Step',
                          color='Units produced',
                          color_continuous_scale='Blues')
        fig_units.update_layout(height=400)
        st.plotly_chart(fig_units, use_container_width=True)
        
    with col2:
        # Running Status Distribution
        status_counts = df['Running Status'].value_counts()
        fig_status = px.pie(values=status_counts.values, names=status_counts.index,
                           title='Running Status Distribution')
        fig_status.update_layout(height=400)
        st.plotly_chart(fig_status, use_container_width=True)
    
    # Time Analysis
    fig_time = go.Figure()
    fig_time.add_trace(go.Bar(name='Expected Time', x=df['Steps'], y=df['Expected Lot Run Time (Hours)']))
    fig_time.add_trace(go.Bar(name='Actual Time', x=df['Steps'], y=df['Current Lot Run Time (Hours)']))
    fig_time.add_trace(go.Bar(name='Down Time', x=df['Steps'], y=df['Process Down time']))
    fig_time.update_layout(barmode='group', title='Time Analysis by Step', height=400)
    st.plotly_chart(fig_time, use_container_width=True)

with tab3:
    col1, col2 = st.columns(2)
    
    with col1:
        # Material Usage
        fig_material = px.bar(df, x='Steps', y='Material Used (KG)',
                             title='Material Usage by Step',
                             color='Material Used (KG)',
                             color_continuous_scale='Greens')
        fig_material.update_layout(height=400)
        st.plotly_chart(fig_material, use_container_width=True)
        
    with col2:
        # Waste Analysis
        fig_waste = px.bar(df, x='Steps', y='Waste Materials (KG)',
                          title='Waste Materials by Step',
                          color='Waste Materials (KG)',
                          color_continuous_scale='Reds')
        fig_waste.update_layout(height=400)
        st.plotly_chart(fig_waste, use_container_width=True)
    
    # Material Efficiency
    df['Material_Efficiency'] = (df['Material Used (KG)'] - df['Waste Materials (KG)']) / df['Material Used (KG)']
    fig_efficiency = px.line(df, x='Steps', y='Material_Efficiency',
                            title='Material Efficiency by Step',
                            markers=True)
    fig_efficiency.update_yaxes(tickformat='.0%')
    st.plotly_chart(fig_efficiency, use_container_width=True)

with tab4:
    # Step selector
    selected_step = st.selectbox('Select a Step for Detailed View:', df['Steps'].unique())
    
    # Filter data for selected step
    step_data = df[df['Steps'] == selected_step].iloc[0]
    
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
        waste_rate = (step_data['Waste Materials (KG)'] / step_data['Material Used (KG)']) * 100
        st.write(f"**Waste Rate:** {waste_rate:.1f}%")
    
    # Performance gauge charts for selected step
    col1, col2, col3, col4 = st.columns(4)
    
    metrics = [
        ('OEE', step_data['OEE'], col1),
        ('Availability', step_data['Availability'], col2),
        ('Performance', step_data['Performance'], col3),
        ('Quality', step_data['Quality'], col4)
    ]
    
    for metric_name, value, col in metrics:
        with col:
            fig_gauge = go.Figure(go.Indicator(
                mode = "gauge+number+delta",
                value = value * 100,
                domain = {'x': [0, 1], 'y': [0, 1]},
                title = {'text': metric_name},
                delta = {'reference': 85 if metric_name == 'OEE' else 90},
                gauge = {
                    'axis': {'range': [None, 100]},
                    'bar': {'color': "darkblue"},
                    'steps': [
                        {'range': [0, 50], 'color': "lightgray"},
                        {'range': [50, 80], 'color': "yellow"},
                        {'range': [80, 100], 'color': "green"}
                    ],
                    'threshold': {
                        'line': {'color': "red", 'width': 4},
                        'thickness': 0.75,
                        'value': 85 if metric_name == 'OEE' else 90
                    }
                }
            ))
            fig_gauge.update_layout(height=300)
            st.plotly_chart(fig_gauge, use_container_width=True)

st.markdown("---")

# === Data Table Section ===
st.header("📋 Detailed Data View")

# Add filters
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

# Apply filters
filtered_df = df[
    (df['Running Status'].isin(status_filter)) &
    (df['OEE'] >= oee_threshold)
]

# Display filtered data
columns_order = ['Steps', 'OEE', 'Availability', 'Performance', 'Quality'] + [col for col in filtered_df.columns if col not in ['Steps', 'OEE', 'Availability', 'Performance', 'Quality']]
filtered_df = filtered_df[columns_order]

st.dataframe(
    filtered_df.style.format({
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
    }).background_gradient(subset=['OEE', 'Availability', 'Performance', 'Quality']),
    use_container_width=True,
)

# Summary statistics
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
        st.write(f"- Best Performing Step: {filtered_df.loc[filtered_df['OEE'].idxmax(), 'Steps']}")
        st.write(f"- Lowest Performing Step: {filtered_df.loc[filtered_df['OEE'].idxmin(), 'Steps']}")
        st.write(f"- Average Waste Rate: {(filtered_df['Waste Materials (KG)'].sum() / filtered_df['Material Used (KG)'].sum() * 100):.1f}%")
    else:
        st.write("- No data matches the current filter criteria")
        st.write("- Please adjust the filters to see performance ranges")

