"""
Enterprise Portfolio Analytics Dashboard - Enhanced Edition
Streamlit UI for multi-source project analysis with Persona-Based Insights

NEW FEATURES:
- Persona Selector (Executive / VP / Manager)
- 10 Insight Categories with Filtering
- Strategic / Operational / Risk View Separation
- Project-Level Persona-Based Insights (NEW)
- Data Validation & Accuracy Governance (NEW)
- Enhanced Visualizations
- Hover Tooltips for Project Details
- Variance Calculation Explanations
- Improved Color Schemes
- Project Dropdown Tables in Insights (NEW)
"""

import streamlit as st
import plotly.graph_objects as go
import plotly.express as px
import pandas as pd
from datetime import datetime
import json
from ai_engine import PortfolioAIEngine
import os
import numpy as np
from collections import defaultdict


def load_custom_css():
    """Apply custom CSS styling for enterprise dashboard"""
    st.markdown("""
    <style>
        /* ==================== FORCE LIGHT MODE ==================== */
        :root {
            color-scheme: light !important;
        }
        
        html, body {
            background-color: #f5f7fa !important;
            color: #1a202c !important;
        }
        
        /* ==================== GLOBAL ELEMENTS ==================== */
        * {
            border-color: #e2e8f0 !important;
        }
        
        /* ==================== STREAMLIT APP ==================== */
        .stApp, 
        [data-testid="stAppViewContainer"], 
        [data-testid="stHeader"],
        [data-testid="stToolbar"] {
            background-color: #f5f7fa !important;
            background: linear-gradient(135deg, #f5f7fa 0%, #f0f4f8 100%) !important;
        }
        
        [data-testid="stSidebar"],
        [data-testid="stSidebarContent"] {
            background-color: #ffffff !important;
            border-right: 1px solid #e2e8f0 !important;
        }
        
        /* ==================== TEXT ELEMENTS ==================== */
        p, span, div, label, h1, h2, h3, h5, h6 {
            color: #1a202c !important;
        }
        
        /* h4 excluded - will be controlled by white-header-text class */
        h4 {
            color: #1a202c !important;
        }
        
        /* White header class - override h4 color */
        .white-header-text h4 {
            color: #ffffff !important;
        }
        
        h1, h2, h3 {
            font-weight: 700 !important;
            letter-spacing: -0.5px !important;
        }
        
        h4, h5, h6 {
            font-weight: 600 !important;
        }
        
        /* ==================== MAIN HEADER (GRADIENT) ==================== */
        .main-header {
            font-size: 2.8rem !important;
            font-weight: 800 !important;
            margin-bottom: 1.5rem !important;
            text-align: center !important;
            background: linear-gradient(135deg, #6366f1 0%, #8b5cf6 50%, #ec4899 100%) !important;
            -webkit-background-clip: text !important;
            -webkit-text-fill-color: transparent !important;
            background-clip: text !important;
            letter-spacing: -1px !important;
        }
        
        /* ==================== SECTION HEADERS ==================== */
        .section-header {
            font-size: 1.75rem !important;
            font-weight: 700 !important;
            color: #1e293b !important;
            margin-top: 2.5rem !important;
            margin-bottom: 1.5rem !important;
            padding-bottom: 1rem !important;
            border-bottom: 3px solid !important;
            border-image: linear-gradient(90deg, #6366f1 0%, #8b5cf6 100%) 1 !important;
            position: relative !important;
        }
        
        /* ==================== INPUT FIELDS ==================== */
        .stTextInput > div > div > input,
        .stNumberInput > div > div > input,
        .stDateInput > div > div > input,
        .stTimeInput > div > div > input,
        .stSelectbox [data-baseweb="select"],
        .stMultiSelect [data-baseweb="select"] {
            background-color: #ffffff !important;
            color: #1a202c !important;
            border: 2px solid #e2e8f0 !important;
            border-radius: 8px !important;
            padding: 10px 12px !important;
            font-size: 14px !important;
            transition: all 0.3s ease !important;
        }
        
        .stTextInput > div > div > input:hover,
        .stNumberInput > div > div > input:hover,
        .stSelectbox [data-baseweb="select"]:hover,
        .stMultiSelect [data-baseweb="select"]:hover {
            border-color: #cbd5e1 !important;
            box-shadow: 0 4px 12px rgba(99, 102, 241, 0.1) !important;
        }
        
        .stTextInput > div > div > input:focus,
        .stNumberInput > div > div > input:focus,
        .stSelectbox [data-baseweb="select"]:focus,
        .stMultiSelect [data-baseweb="select"]:focus {
            border-color: #6366f1 !important;
            box-shadow: 0 0 0 3px rgba(99, 102, 241, 0.15) !important;
        }
        
        /* ==================== BUTTONS ==================== */
        .stButton > button,
        .stDownloadButton > button,
        button[kind="secondary"] {
            background: linear-gradient(135deg, #6366f1 0%, #8b5cf6 100%) !important;
            color: #ffffff !important;
            border: none !important;
            border-radius: 8px !important;
            padding: 10px 24px !important;
            font-weight: 600 !important;
            font-size: 14px !important;
            cursor: pointer !important;
            transition: all 0.3s ease !important;
            box-shadow: 0 4px 12px rgba(99, 102, 241, 0.3) !important;
        }
        
        .stButton > button:hover,
        .stDownloadButton > button:hover {
            transform: translateY(-2px) !important;
            box-shadow: 0 6px 20px rgba(99, 102, 241, 0.4) !important;
        }
        
        .stButton > button:active,
        .stDownloadButton > button:active {
            transform: translateY(0) !important;
        }
        
        /* ==================== FILE UPLOADER ==================== */
        [data-testid="stFileUploader"],
        [data-testid="stFileUploadDropzone"] {
            background-color: #ffffff !important;
        }
        
        [data-testid="stFileUploader"] section {
            background-color: #f8f9fa !important;
            border: 2px dashed #cbd5e1 !important;
            border-radius: 12px !important;
            padding: 30px !important;
            transition: all 0.3s ease !important;
        }
        
        [data-testid="stFileUploader"] section:hover {
            border-color: #6366f1 !important;
            background-color: #f3f4f6 !important;
        }
        
        [data-testid="stFileUploader"] button {
            background: linear-gradient(135deg, #6366f1 0%, #8b5cf6 100%) !important;
            color: #ffffff !important;
            border: none !important;
            border-radius: 8px !important;
            font-weight: 600 !important;
        }
        
        [data-testid="stFileUploader"] small {
            color: #64748b !important;
        }
        
        /* ==================== SELECTBOX & DROPDOWNS - FIXED ==================== */
        .stSelectbox,
        .stMultiSelect {
            color: #1a202c !important;
        }
        
        /* Selectbox main container */
        [data-baseweb="select"] > div {
            background-color: #ffffff !important;
            color: #1a202c !important;
            border-radius: 8px !important;
            border: 2px solid #e2e8f0 !important;
        }
        
        /* CRITICAL FIX: Dropdown menu (the list that opens) */
        [data-baseweb="popover"],
        [data-baseweb="menu"],
        [role="listbox"] {
            background-color: #ffffff !important;
            color: #000000 !important;
            border: 2px solid #e2e8f0 !important;
            border-radius: 8px !important;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15) !important;
        }
        
        /* Dropdown menu items */
        [data-baseweb="menu"] ul,
        [role="listbox"] ul {
            background-color: #ffffff !important;
        }
        
        /* Individual dropdown options */
        [data-baseweb="menu"] li,
        [role="option"] {
            background-color: #ffffff !important;
            color: #000000 !important;
            padding: 10px 16px !important;
        }
        
        /* Dropdown option hover state */
        [data-baseweb="menu"] li:hover,
        [role="option"]:hover,
        [aria-selected="true"] {
            background-color: #f3f4f6 !important;
            color: #000000 !important;
        }
        
        /* Selected value text in closed dropdown */
        [data-baseweb="select"] [role="button"] {
            color: #000000 !important;
        }
        
        /* Dropdown arrow icon */
        [data-baseweb="select"] svg {
            color: #000000 !important;
            fill: #000000 !important;
        }
        
        /* ==================== RADIO BUTTONS & CHECKBOXES ==================== */
        .stRadio > label,
        .stCheckbox > label {
            color: #1a202c !important;
            font-weight: 500 !important;
        }
        
        .stRadio [type="radio"],
        .stCheckbox [type="checkbox"] {
            accent-color: #6366f1 !important;
        }
        
        /* ==================== TAGS/CHIPS (MULTISELECT) ==================== */
        .stMultiSelect [data-baseweb="tag"] {
            background-color: #e0e7ff !important;
            color: #3730a3 !important;
            border-radius: 6px !important;
            font-weight: 600 !important;
        }
        
        /* ==================== EXPANDER ==================== */
        .streamlit-expanderHeader,
        [data-testid="stExpanderDetails"] {
            background-color: #f8f9fa !important;
            color: #1a202c !important;
            border-radius: 8px !important;
            border: 1px solid #e2e8f0 !important;
        }
        
        .streamlit-expanderHeader:hover {
            background-color: #f3f4f6 !important;
            border-color: #cbd5e1 !important;
        }
        
        /* ==================== DATAFRAMES & TABLES - SIMPLE WHITE & BLACK ==================== */
        
        [data-testid="stDataFrame"] {
            background: #ffffff !important;
        }
        
        [data-testid="stDataFrame"] thead {
            background: #ffffff !important;
        }
        
        [data-testid="stDataFrame"] thead th {
            background: #f5f5f5 !important;
            color: #000000 !important;
        }
        
        [data-testid="stDataFrame"] tbody {
            background: #ffffff !important;
        }
        
        [data-testid="stDataFrame"] td {
            background: #ffffff !important;
            color: #000000 !important;
        }
        
        [data-testid="stDataFrame"] th {
            color: #000000 !important;
        }
        
        /* ==================== METRIC IMPROVEMENTS ==================== */
        [data-testid="stMetric"] {
            background: linear-gradient(135deg, #ffffff 0%, #f8f9fa 100%) !important;
            border: 1px solid #e2e8f0 !important;
            border-radius: 12px !important;
            padding: 20px !important;
            box-shadow: 0 1px 3px rgba(0, 0, 0, 0.05) !important;
            transition: all 0.3s ease !important;
        }
        
        [data-testid="stMetric"]:hover {
            border-color: #cbd5e1 !important;
            box-shadow: 0 4px 12px rgba(99, 102, 241, 0.1) !important;
        }
        
        /* ==================== METRICS ==================== */
        [data-testid="stMetricValue"] {
            color: #1a202c !important;
            font-size: 2.5rem !important;
            font-weight: 800 !important;
        }
        
        [data-testid="stMetricLabel"] {
            color: #64748b !important;
            font-weight: 600 !important;
            font-size: 0.95rem !important;
        }
        
        [data-testid="stMetricDelta"] {
            color: #059669 !important;
            font-weight: 700 !important;
        }
        
        /* ==================== TABS ==================== */
        .stTabs [data-baseweb="tab-list"] {
            background-color: #ffffff !important;
            border-bottom: 2px solid #e2e8f0 !important;
        }
        
        .stTabs [data-baseweb="tab"] {
            color: #64748b !important;
            font-weight: 600 !important;
            border-bottom: 3px solid transparent !important;
            transition: all 0.3s ease !important;
        }
        
        .stTabs [data-baseweb="tab"][aria-selected="true"] {
            color: #6366f1 !important;
            border-bottom-color: #6366f1 !important;
        }
        
        .stTabs [data-baseweb="tab"]:hover {
            color: #1a202c !important;
        }
        
        /* ==================== ALERTS & NOTIFICATIONS ==================== */
        .stAlert,
        [data-testid="stAlert"] {
            background-color: #f8f9fa !important;
            border: 1px solid #e2e8f0 !important;
            border-radius: 8px !important;
            padding: 16px !important;
            color: #1a202c !important;
            border-left: 4px solid #6366f1 !important;
        }
        
        /* ==================== CARDS & CONTAINERS ==================== */
        .card-container {
            background-color: #ffffff !important;
            border: 1px solid #e2e8f0 !important;
            border-radius: 12px !important;
            padding: 20px !important;
            box-shadow: 0 1px 3px rgba(0, 0, 0, 0.05) !important;
            transition: all 0.3s ease !important;
        }
        
        .card-container:hover {
            border-color: #cbd5e1 !important;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08) !important;
        }
        
        /* ==================== SCROLLBARS ==================== */
        ::-webkit-scrollbar {
            width: 8px !important;
            height: 8px !important;
        }
        
        ::-webkit-scrollbar-track {
            background: #f5f7fa !important;
        }
        
        ::-webkit-scrollbar-thumb {
            background: #cbd5e1 !important;
            border-radius: 4px !important;
        }
        
        ::-webkit-scrollbar-thumb:hover {
            background: #94a3b8 !important;
        }
        
        /* ==================== INSIGHT BOXES ==================== */
        .insight-box {
            background: linear-gradient(135deg, #f3f4f6 0%, #f8f9fa 100%) !important;
            border: 2px solid #e2e8f0 !important;
            border-radius: 12px !important;
            padding: 20px !important;
            margin-bottom: 16px !important;
            transition: all 0.3s ease !important;
        }
        
        .insight-box:hover {
            border-color: #6366f1 !important;
            box-shadow: 0 4px 12px rgba(99, 102, 241, 0.15) !important;
        }
        
        .insight-critical {
            border-left: 5px solid #dc2626 !important;
            background: linear-gradient(135deg, #fef2f2 0%, #fee2e2 100%) !important;
        }
        
        .insight-high {
            border-left: 5px solid #ea580c !important;
            background: linear-gradient(135deg, #fff7ed 0%, #fef3c7 100%) !important;
        }
        
        .insight-warning {
            border-left: 5px solid #eab308 !important;
            background: linear-gradient(135deg, #fefce8 0%, #fef08a 100%) !important;
        }
        
        /* ==================== BADGES ==================== */
        .badge {
            display: inline-block !important;
            padding: 4px 12px !important;
            border-radius: 20px !important;
            font-size: 12px !important;
            font-weight: 600 !important;
        }
        
        .badge-critical {
            background-color: #fecaca !important;
            color: #7f1d1d !important;
        }
        
        .badge-high {
            background-color: #fed7aa !important;
            color: #92400e !important;
        }
        
        .badge-warning {
            background-color: #fef08a !important;
            color: #713f12 !important;
        }
        
        .badge-success {
            background-color: #bbf7d0 !important;
            color: #065f46 !important;
        }
        
        .badge-info {
            background-color: #bfdbfe !important;
            color: #1e3a8a !important;
        }
        
        /* ==================== ALERT BOXES (for insight cards) ==================== */
        .alert-critical {
            background: linear-gradient(135deg, #fef2f2 0%, #fee2e2 100%) !important;
            border: 1px solid #fecaca !important;
            border-left: 4px solid #dc2626 !important;
            border-radius: 8px !important;
            padding: 12px 16px !important;
            margin: 8px 0 !important;
            color: #7f1d1d !important;
        }
        
        .alert-warning {
            background: linear-gradient(135deg, #fff7ed 0%, #fef3c7 100%) !important;
            border: 1px solid #fed7aa !important;
            border-left: 4px solid #ea580c !important;
            border-radius: 8px !important;
            padding: 12px 16px !important;
            margin: 8px 0 !important;
            color: #92400e !important;
        }
        
        .alert-info {
            background: linear-gradient(135deg, #f0f9ff 0%, #e0f2fe 100%) !important;
            border: 1px solid #bae6fd !important;
            border-left: 4px solid #0284c7 !important;
            border-radius: 8px !important;
            padding: 12px 16px !important;
            margin: 8px 0 !important;
            color: #0c4a6e !important;
        }
        
        /* ==================== PLOTLY CHARTS - FORCE BLACK TEXT ==================== */
        .plotly-graph-div {
            background-color: #ffffff !important;
        }
        
        .plotly {
            background-color: #ffffff !important;
        }
        
        .plotly-graph-div text {
            fill: #000000 !important;
            color: #000000 !important;
        }
        
        svg text {
            fill: #000000 !important;
        }
        
        /* ==================== FORCE WHITE TEXT ON COLORED HEADERS ==================== */
        .white-header-text {
            color: #ffffff !important;
        }
        
        .white-header-text h4 {
            color: #ffffff !important;
            text-shadow: 0 2px 4px rgba(0,0,0,0.6) !important;
        }
        
        /* ==================== LINKS ==================== */
        a {
            color: #6366f1 !important;
            text-decoration: none !important;
            transition: color 0.3s ease !important;
        }
        
        a:hover {
            color: #8b5cf6 !important;
            text-decoration: underline !important;
        }
        
        /* ==================== ULTIMATE WHITE TEXT OVERRIDE ==================== */
        .white-header-text,
        .white-header-text * {
            color: #ffffff !important;
        }
        
        .white-header-text h4,
        div[class*="white-header"] h4 {
            color: #ffffff !important;
            background: transparent !important;
        }
        
        /* Override any inherited dark color on h4 inside white-header-text */
        .white-header-text h4 {
            color: #ffffff !important;
            -webkit-text-fill-color: #ffffff !important;
            caret-color: #ffffff !important;
        }
        
        /* ==================== RESPONSIVE ==================== */
        @media (max-width: 768px) {
            .main-header {
                font-size: 2rem !important;
            }
            
            .section-header {
                font-size: 1.5rem !important;
            }
            
            .stButton > button,
            .stDownloadButton > button {
                padding: 8px 16px !important;
                font-size: 13px !important;
            }
        }
    </style>
    """, unsafe_allow_html=True)

def create_status_distribution_chart(summary):
    """Create pie chart of project status distribution"""
    status_dist = summary.get('status_distribution', {})
    
    if not status_dist:
        return None
    
    colors = {
        'On Track': '#10b981',
        'At Risk': '#f59e0b',
        'Delayed': '#ef4444',
        'Unknown': '#9ca3af'
    }
    
    labels = list(status_dist.keys())
    values = list(status_dist.values())
    chart_colors = [colors.get(label, '#94a3b8') for label in labels]
    
    fig = go.Figure(data=[go.Pie(
        labels=labels,
        values=values,
        marker=dict(colors=chart_colors, line=dict(color='white', width=2)),
        textinfo='label+percent',
        textposition='auto',
        textfont=dict(size=14, color='#000000'),
        hovertemplate='<b>%{label}</b><br>Count: %{value}<br>Percentage: %{percent}<extra></extra>'
    )])
    
    fig.update_layout(
        title='Portfolio Status Distribution',
        height=500,
        plot_bgcolor='#ffffff',
        paper_bgcolor='#ffffff',
        font=dict(color='#000000', size=14)
    )
    
    return fig


def create_health_distribution_chart(summary):
    """Create bar chart of health indicators with enforced color mapping"""
    health_dist = summary.get('health_distribution', {})
    
    if not health_dist:
        return None
    
    colors_map = {
        'Green': '#10b981',
        'green': '#10b981',
        'Yellow': '#f59e0b',
        'yellow': '#f59e0b',
        'Red': '#ef4444',
        'red': '#ef4444',
        'Unknown': '#9ca3af',
        'unknown': '#9ca3af'
    }
    
    df = pd.DataFrame(list(health_dist.items()), columns=['Health', 'Count'])
    df['Color'] = df['Health'].apply(lambda x: colors_map.get(x, colors_map.get(x.lower(), '#9ca3af')))
    
    fig = go.Figure(data=[go.Bar(
        x=df['Health'],
        y=df['Count'],
        marker_color=df['Color'],
        marker_line=dict(color='white', width=1),
        text=df['Count'],
        textposition='outside',
        textfont=dict(size=12, color='#000000'),
        hovertemplate='<b>%{x}</b><br>Projects: %{y}<extra></extra>'
    )])
    
    fig.update_layout(
        title='Health Indicator Distribution',
        xaxis_title='Health Status',
        yaxis_title='Number of Projects',
        height=500,
        hovermode='x unified',
        plot_bgcolor='#ffffff',
        paper_bgcolor='#ffffff',
        font=dict(color='#000000', size=12)
    )
    
    return fig


def create_budget_variance_chart(projects):
    """Create chart showing budget variance across projects"""
    data = []
    
    for project_id, project_data in projects.items():
        derived = project_data.get('derived_metrics', {})
        metadata = project_data.get('project_metadata', {})
        
        cost_var = derived.get('cost_variance_pct')
        if cost_var is not None:
            data.append({
                'Project': metadata.get('project_name', project_id),
                'Project ID': project_id,
                'Variance %': cost_var
            })
    
    if not data:
        return None
    
    df = pd.DataFrame(data)
    df = df.sort_values('Variance %')
    
    colors = []
    for x in df['Variance %']:
        if x < -10:
            colors.append('#ef4444')
        elif x > 5:
            colors.append('#10b981')
        else:
            colors.append('#f59e0b')
    
    fig = go.Figure(data=[go.Bar(
        x=df['Project'],
        y=df['Variance %'],
        marker_color=colors,
        marker_line=dict(color='white', width=1),
        text=df['Variance %'].apply(lambda x: f"{x:.1f}%"),
        textposition='outside',
        textfont=dict(size=11, color='#000000'),
        customdata=df['Project ID'],
        hovertemplate='<b>%{customdata}</b><br>' +
                      'Project: %{x}<br>' +
                      'Variance: %{y:.1f}%<br>' +
                      '<i>Calculation: ((Actual Cost - Baseline) / Baseline) × 100</i><extra></extra>'
    )])
    
    fig.update_layout(
        title='Budget Variance by Project<br><sub>Negative = Over Budget | Positive = Under Budget</sub>',
        xaxis_title='Projects',
        yaxis_title='Variance % (negative = overrun)',
        height=600,
        xaxis={'tickangle': -45},
        hovermode='x unified',
        plot_bgcolor='#ffffff',
        paper_bgcolor='#ffffff',
        font=dict(color='#000000', size=11)
    )
    
    fig.add_hline(y=0, line_dash="dash", line_color="gray", annotation_text="Break Even")
    
    return fig


def create_schedule_variance_chart(projects):
    """Create chart showing schedule variance"""
    data = []
    
    for project_id, project_data in projects.items():
        derived = project_data.get('derived_metrics', {})
        metadata = project_data.get('project_metadata', {})
        
        schedule_var = derived.get('schedule_variance_days')
        if schedule_var is not None:
            data.append({
                'Project': metadata.get('project_name', project_id),
                'Project ID': project_id,
                'Delay (Days)': schedule_var
            })
    
    if not data:
        return None
    
    df = pd.DataFrame(data)
    df = df.sort_values('Delay (Days)', ascending=False)
    
    colors = []
    for x in df['Delay (Days)']:
        if x > 30:
            colors.append('#ef4444')
        elif x <= 0:
            colors.append('#10b981')
        else:
            colors.append('#f59e0b')
    
    fig = go.Figure(data=[go.Bar(
        x=df['Project'],
        y=df['Delay (Days)'],
        marker_color=colors,
        marker_line=dict(color='white', width=1),
        text=df['Delay (Days)'].apply(lambda x: f"{x:.0f}d"),
        textposition='outside',
        textfont=dict(size=11, color='#000000'),
        customdata=df['Project ID'],
        hovertemplate='<b>%{customdata}</b><br>' +
                      'Project: %{x}<br>' +
                      'Schedule Variance: %{y:.0f} days<br>' +
                      '<i>Calculation: Actual End Date - Baseline End Date</i><extra></extra>'
    )])
    
    fig.update_layout(
        title='Schedule Variance by Project<br><sub>Positive = Delayed | Negative = Ahead of Schedule</sub>',
        xaxis_title='Projects',
        yaxis_title='Days (positive = delayed)',
        height=600,
        xaxis={'tickangle': -45},
        hovermode='x unified',
        plot_bgcolor='#ffffff',
        paper_bgcolor='#ffffff',
        font=dict(color='#000000', size=11)
    )
    
    fig.add_hline(y=0, line_dash="dash", line_color="gray", annotation_text="On Time")
    
    return fig


def create_data_completeness_chart(summary):
    """Create chart showing data source coverage"""
    completeness = summary.get('data_completeness', {})
    
    if not completeness:
        return None
    
    labels = ['Full Data\n(3 sources)', 'Partial Data\n(2 sources)', 'Minimal Data\n(1 source)']
    values = [
        completeness.get('full_data', 0),
        completeness.get('partial_data', 0),
        completeness.get('minimal_data', 0)
    ]
    colors = ['#10b981', '#f59e0b', '#ef4444']
    
    fig = go.Figure(data=[go.Bar(
        x=labels,
        y=values,
        marker_color=colors,
        marker_line=dict(color='white', width=1),
        text=values,
        textposition='outside',
        textfont=dict(size=12, color='#000000'),
        hovertemplate='<b>%{x}</b><br>Projects: %{y}<extra></extra>'
    )])
    
    fig.update_layout(
        title='Data Source Completeness',
        xaxis_title='Coverage Level',
        yaxis_title='Number of Projects',
        height=500,
        hovermode='x unified',
        plot_bgcolor='#ffffff',
        paper_bgcolor='#ffffff',
        font=dict(color='#000000', size=12)
    )
    
    return fig


def create_portfolio_metrics_summary(summary):
    """Create portfolio-level metrics display"""
    metrics = summary.get('portfolio_metrics', {})
    
    if not metrics:
        return None
    
    total_budget = metrics.get('total_baseline_budget', 0)
    total_actual = metrics.get('total_actual_cost', 0)
    variance_pct = metrics.get('portfolio_variance_pct')
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric(
            "Portfolio Baseline Budget",
            f"${total_budget/1_000_000:.1f}M" if total_budget > 1_000_000 else f"${total_budget/1000:.0f}K"
        )
    
    with col2:
        st.metric(
            "Total Actual Cost",
            f"${total_actual/1_000_000:.1f}M" if total_actual > 1_000_000 else f"${total_actual/1000:.0f}K"
        )
    
    with col3:
        if variance_pct is not None:
            st.metric(
                "Portfolio Variance",
                f"{variance_pct:.1f}%",
                delta=f"{'Over' if variance_pct < 0 else 'Under'} Budget"
            )
    
    st.markdown("""
    <div style="background: linear-gradient(135deg, #f0f9ff 0%, #e0f2fe 100%); border: 2px solid #bae6fd; border-radius: 8px; padding: 16px; margin-top: 16px; border-left: 4px solid #0284c7;">
        <p style="margin: 0 0 8px 0; font-weight: 700; color: #0c4a6e;">📊 Variance Calculation</p>
        <p style="margin: 8px 0; color: #0c4a6e; font-family: monospace; font-size: 12px;">Portfolio Variance = ((Total Actual Cost - Total Baseline Budget) / Total Baseline Budget) × 100</p>
        <p style="margin: 8px 0; color: #0c4a6e; font-size: 13px;"><strong>Interpretation:</strong> Negative % = Over Budget | Positive % = Under Budget</p>
    </div>
    """, unsafe_allow_html=True)


def display_insight_card(insight: dict, projects_map: dict = None):
    """Display a single insight card with appropriate styling and detailed project breakdown in expandable tables"""
    severity = insight.get('severity', 'info')
    confidence = insight.get('confidence', 'Unknown')
    
    severity_colors = {
        'critical': 'alert-critical',
        'high': 'alert-warning',
        'warning': 'alert-warning',
        'info': 'alert-info'
    }
    
    severity_icons = {
        'critical': '🚨',
        'high': '⚠️',
        'warning': '⚠️',
        'info': 'ℹ️'
    }
    
    alert_class = severity_colors.get(severity, 'alert-info')
    icon = severity_icons.get(severity, 'ℹ️')
    
    title = insight['title']
    metrics = insight.get('metrics', {})
    
    severity_background = {
        'critical': 'linear-gradient(135deg, #dc2626 0%, #991b1b 100%)',
        'high': 'linear-gradient(135deg, #ea580c 0%, #b45309 100%)',
        'warning': 'linear-gradient(135deg, #eab308 0%, #b45309 100%)',
        'info': 'linear-gradient(135deg, #0284c7 0%, #075985 100%)'
    }
    
    bg_style = severity_background.get(severity, 'linear-gradient(135deg, #6366f1 0%, #7c3aed 100%)')
    
    header_html = f'<div class="white-header-text" style="background: {bg_style}; padding: 1.5rem; border-radius: 10px; color: #ffffff !important; margin: 0.5rem 0; box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);"><h4 style="margin: 0 !important; color: #ffffff !important; font-size: 1.3rem !important; font-weight: 700 !important; text-shadow: 0 2px 4px rgba(0,0,0,0.6) !important; background: transparent !important;">{icon} {title}</h4></div>'
    st.markdown(header_html, unsafe_allow_html=True)
    
    st.markdown(f"""
    <div class="insight-box insight-{severity}">
        <p style="margin: 8px 0; color: #1a202c;"><strong>📁 Category:</strong> {insight['category'].replace('_', ' ').title()}</p>
        <p style="margin: 8px 0; color: #1a202c;"><strong>📊 Confidence:</strong> {confidence}</p>
        <p style="margin: 12px 0; color: #1a202c;"><strong>📝 Description:</strong></p>
        <p style="margin: 8px 0; padding-left: 12px; border-left: 3px solid #6366f1; color: #1a202c;">{insight['description']}</p>
        <p style="margin: 12px 0; color: #1a202c;"><strong>💥 Impact:</strong></p>
        <p style="margin: 8px 0; padding-left: 12px; border-left: 3px solid #ea580c; color: #1a202c;">{insight['impact']}</p>
        <p style="margin: 12px 0; color: #1a202c;"><strong>✅ Recommendation:</strong></p>
        <p style="margin: 8px 0; padding-left: 12px; border-left: 3px solid #10b981; color: #1a202c;">{insight['recommendation']}</p>
    </div>
    """, unsafe_allow_html=True)
    
    if isinstance(metrics, dict) and projects_map:
        project_keys = {
            'flagged_projects': '🚩 Flagged Projects',
            'leakage_projects': '💧 Leakage Projects',
            'on_track_projects': '✅ On Track',
            'at_risk_projects': '⚠️ At Risk',
            'delayed_projects': '🔴 Delayed',
            'green_projects': '✅ Healthy',
            'yellow_projects': '⚠️ Warning',
            'red_projects': '🔴 Critical',
            'affected_projects': '📋 Affected',
            'uncovered_initiatives': '❌ Uncovered',
            'worst_offenders': '⚠️ Worst Offenders',
            'overloaded_managers': '👤 Overloaded Managers',
            'projects': '📋 Projects',
            'project_ids': '📋 Projects',
            'top_contributors': '📊 Top Contributors',
        }
        
        for key, label in project_keys.items():
            if key in metrics and metrics[key]:
                projects_list = metrics[key]
                
                if isinstance(projects_list, list):
                    if projects_list and isinstance(projects_list[0], dict):
                        project_ids = [item.get('project_id', item.get('Project ID', item)) for item in projects_list]
                    else:
                        project_ids = projects_list
                    
                    with st.expander(f"{label} ({len(project_ids)} projects)"):
                        df_projects = pd.DataFrame([
                            {
                                'Project ID': pid,
                                'Project Name': projects_map.get(pid, 'Unknown') if isinstance(pid, str) else str(pid)
                            }
                            for pid in project_ids if pid
                        ])
                        if not df_projects.empty:
                            st.dataframe(df_projects, use_container_width=True, hide_index=True)
                        else:
                            st.info("No projects found")
    
    if metrics:
        with st.expander("📊 Supporting Metrics"):
            st.json(metrics)


def display_project_assessment(project_data):
    """Display detailed project assessment"""
    
    assessment = project_data.get('assessment', {})
    metadata = project_data.get('project_metadata', {})
    
    st.markdown(f"### Project Assessment: {metadata.get('project_name', 'Unknown')}")
    st.caption(f"Project ID: {metadata.get('project_id')} | Analysis: {metadata.get('analysis_timestamp', '')}")
    
    st.markdown("#### 📋 Overall Assessment")
    
    overall = assessment.get('overall_assessment', {})
    
    status = overall.get('status', 'Unknown')
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("Status", status)
    
    with col2:
        st.metric("Health", overall.get('health', 'Unknown'))
    
    with col3:
        st.metric("Confidence", overall.get('confidence_level', 'Unknown'))
    
    with col4:
        st.metric("Data Sources", f"{overall.get('data_sources_available', 0)}/3")
    
    if overall.get('summary'):
        st.markdown(f"""
        <div class="insight-box">
            <p style="margin: 0; padding: 12px; background: linear-gradient(135deg, #f0f9ff 0%, #e0f2fe 100%); border-radius: 8px; border-left: 4px solid #0284c7; color: #0c4a6e;">
                {overall['summary']}
            </p>
        </div>
        """, unsafe_allow_html=True)
    
    baseline = project_data.get('baseline_metrics', {})
    wave = project_data.get('latest_wave_snapshot', {})
    actuals = project_data.get('actuals_summary', {})
    
    st.markdown("**Available Data:**")
    badges = []
    if baseline:
        badges.append('<span class="data-source-badge source-smartsheet">Smartsheet</span>')
    if wave:
        badges.append('<span class="data-source-badge source-wave">Wave</span>')
    if actuals:
        badges.append('<span class="data-source-badge source-tick">Tick</span>')
    
    st.markdown(" ".join(badges), unsafe_allow_html=True)
    
    st.markdown("#### 🎯 Key Drivers")
    drivers = assessment.get('key_drivers', [])
    if drivers:
        for driver in drivers:
            st.markdown(f"• {driver}")
    else:
        st.caption("No key drivers identified")
    
    st.markdown("#### 🔗 Cross-Source Observations")
    observations = assessment.get('cross_source_observations', [])
    if observations:
        for obs in observations:
            st.markdown(f"• {obs}")
    else:
        st.caption("Limited cross-source data")
    
    st.markdown("#### ⚠️ Risks & Early Warnings")
    risks = assessment.get('risks_warnings', [])
    if risks:
        for risk in risks:
            if '[CRITICAL]' in risk:
                st.markdown(f'<div class="alert-critical">{risk}</div>', unsafe_allow_html=True)
            elif '[WARNING]' in risk:
                st.markdown(f'<div class="alert-warning">{risk}</div>', unsafe_allow_html=True)
            else:
                st.markdown(f'<div class="alert-info">{risk}</div>', unsafe_allow_html=True)
    else:
        st.success("No significant risks detected")
    
    st.markdown("#### ✅ Positive Signals")
    positives = assessment.get('positive_signals', [])
    if positives:
        for pos in positives:
            st.markdown(f"• {pos}")
    else:
        st.caption("No strong positive signals")
    
    st.markdown("#### 📊 Data Gaps / Assumptions")
    gaps = assessment.get('data_gaps', [])
    if gaps:
        for gap in gaps:
            st.warning(gap)
    else:
        st.success("Complete data from all sources")
    
    if assessment.get('recommendations'):
        st.markdown("#### 💡 Recommendations")
        for rec in assessment['recommendations']:
            st.info(rec)
    
    with st.expander("📈 View Detailed Metrics"):
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("**Baseline Metrics (Smartsheet)**")
            if baseline:
                st.json(baseline)
            else:
                st.caption("No baseline data")
            
            st.markdown("**Wave Snapshot**")
            if wave:
                st.json(wave)
            else:
                st.caption("No Wave data")
        
        with col2:
            st.markdown("**Actuals Summary (Tick)**")
            if actuals:
                st.json(actuals)
            else:
                st.caption("No Tick data")
            
            st.markdown("**Derived Metrics**")
            derived = project_data.get('derived_metrics', {})
            if derived:
                st.json(derived)
            else:
                st.caption("No derived metrics")
    
    rules = project_data.get('rule_evaluations', [])
    if rules:
        with st.expander("🔍 View Consistency Rule Evaluations"):
            for rule in rules:
                st.markdown(f"**{rule.get('rule')}** - *{rule.get('severity')}*")
                st.markdown(f"Description: {rule.get('description')}")
                if rule.get('recommendation'):
                    st.markdown(f"Recommendation: {rule.get('recommendation')}")
                st.markdown("---")


def remove_duplicate_insights(insights):
    """Remove duplicate insights while preserving order"""
    seen = set()
    unique_insights = []
    
    for insight in insights:
        insight_key = (insight.get('title'), insight.get('category'), insight.get('severity'))
        if insight_key not in seen:
            seen.add(insight_key)
            unique_insights.append(insight)
    
    return unique_insights


def main():
    """Main application"""
    
    st.set_page_config(
        page_title="Enterprise Portfolio Analytics",
        page_icon="📊",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    load_custom_css()
    
    st.markdown('<p class="main-header">Enterprise Data Analytics Dashboard</p>', unsafe_allow_html=True)
    st.markdown("**AI-Powered Multi-Source Project Analysis: Smartsheet + Wave + Tick**")
    st.markdown("---")
    
    with st.sidebar:
        st.header("📁 Upload Data Sources")
        
        st.markdown("""
        Upload data from three systems:
        - **Smartsheet**: Baseline & governance
        - **Wave**: Weekly forecasts & status
        - **Tick**: Actual execution data
        """)
        
        smartsheet_file = st.file_uploader(
            "📘 Smartsheet Export (Baseline)",
            type=['xlsx', 'xlsm', 'csv'],
            help="Project baseline, budgets, and governance data"
        )
        
        wave_file = st.file_uploader(
            "📙 Wave Export (Forecasts)",
            type=['xlsx', 'xlsm', 'csv'],
            help="Weekly status snapshots and forecasts"
        )
        
        tick_file = st.file_uploader(
            "📗 Tick Export (Actuals)",
            type=['xlsx', 'xlsm', 'csv'],
            help="Actual hours and costs from execution"
        )
        
        st.markdown("---")
        
        sheet_config = {}
        
        if smartsheet_file:
            st.markdown("**Smartsheet Configuration**")
            sheet_config['smartsheet_sheet'] = st.text_input(
                "Sheet Name (optional)",
                key="smartsheet_sheet",
                help="Leave blank to use first sheet"
            )
        
        if wave_file:
            st.markdown("**Wave Configuration**")
            sheet_config['wave_sheet'] = st.text_input(
                "Sheet Name (optional)",
                key="wave_sheet",
                help="Leave blank to use first sheet"
            )
        
        if tick_file:
            st.markdown("**Tick Configuration**")
            sheet_config['tick_sheet'] = st.text_input(
                "Sheet Name (optional)",
                key="tick_sheet",
                help="Leave blank to use first sheet"
            )
        
        st.markdown("---")
        
        analyze_button = st.button(
            "🚀 Analyze Portfolio",
            type="primary",
            disabled=not (smartsheet_file or wave_file or tick_file)
        )
        
        if 'engine' in st.session_state:
            st.success("✅ Analysis Complete")
            if 'portfolio_summary' in st.session_state:
                summary = st.session_state['portfolio_summary']
                st.metric("Projects Analyzed", summary['portfolio_overview']['total_projects'])
    
    if analyze_button:
        try:
            with st.spinner("🔄 Loading data and analyzing portfolio..."):
                
                engine = PortfolioAIEngine()
                
                if smartsheet_file:
                    sheet_name = sheet_config.get('smartsheet_sheet') or None
                    if smartsheet_file.name.endswith('.csv'):
                        df = pd.read_csv(smartsheet_file)
                    else:
                        df = pd.read_excel(smartsheet_file, sheet_name=sheet_name)
                    engine.load_smartsheet(df)
                    st.success("✅ Smartsheet loaded")
                
                if wave_file:
                    sheet_name = sheet_config.get('wave_sheet') or None
                    if wave_file.name.endswith('.csv'):
                        df = pd.read_csv(wave_file)
                    else:
                        df = pd.read_excel(wave_file, sheet_name=sheet_name)
                    engine.load_wave(df)
                    st.success("✅ Wave loaded")
                
                if tick_file:
                    sheet_name = sheet_config.get('tick_sheet') or None
                    if tick_file.name.endswith('.csv'):
                        df = pd.read_csv(tick_file)
                    else:
                        df = pd.read_excel(tick_file, sheet_name=sheet_name)
                    engine.load_tick(df)
                    st.success("✅ Tick loaded")
                
                projects = engine.analyze_all_projects()
                
                engine.generate_all_insights()
                
                summary = engine.get_portfolio_summary()
                
                st.session_state['engine'] = engine
                st.session_state['projects'] = projects
                st.session_state['portfolio_summary'] = summary
                
                st.success("✅ Portfolio analysis complete!")
                st.balloons()
                
        except Exception as e:
            st.error(f"❌ Analysis failed: {str(e)}")
            with st.expander("📋 Error Details"):
                import traceback
                st.code(traceback.format_exc())
    
    if 'portfolio_summary' in st.session_state:
        
        summary = st.session_state['portfolio_summary']
        projects = st.session_state['projects']
        engine = st.session_state['engine']
        
        projects_map = {}
        for project_id, project_data in projects.items():
            metadata = project_data.get('project_metadata', {})
            projects_map[project_id] = metadata.get('project_name', 'Unknown')
        
        st.markdown('<p class="section-header">👤 Select Your Persona</p>', unsafe_allow_html=True)
        
        persona = st.radio(
            "View insights tailored to your role:",
            ["Executive (C-Level)", "VP / Portfolio Owner", "Manager / Delivery Lead", "All Insights"],
            horizontal=True
        )
        
        st.markdown('<p class="section-header">📊 Portfolio Overview</p>', unsafe_allow_html=True)
        
        overview = summary['portfolio_overview']
        st.info(f"**Total Projects:** {overview['total_projects']} | **Analysis Time:** {overview['analysis_timestamp']}")
        
        create_portfolio_metrics_summary(summary)
        
        st.markdown('<p class="section-header">💡 Decision-Grade Insights</p>', unsafe_allow_html=True)
        
        if persona == "Executive (C-Level)":
            insights = engine.get_executive_insights()
            st.markdown("**🎯 Strategic & Portfolio-Level Insights**")
        elif persona == "VP / Portfolio Owner":
            insights = engine.get_vp_insights()
            st.markdown("**📈 Portfolio Management & Risk Insights**")
        elif persona == "Manager / Delivery Lead":
            insights = engine.get_manager_insights()
            st.markdown("**🔧 Operational & Execution Insights**")
        else:
            exec_insights = engine.get_executive_insights()
            vp_insights = engine.get_vp_insights()
            mgr_insights = engine.get_manager_insights()
            
            st.markdown("**All Personas Combined:**")
            insights = exec_insights + vp_insights + mgr_insights
            insights = remove_duplicate_insights(insights)
        
        if insights:
            categories = list(set([i['category'] for i in insights]))
            selected_categories = st.multiselect(
                "Filter by Category:",
                options=categories,
                default=categories,
                format_func=lambda x: x.replace('_', ' ').title()
            )
            
            filtered_insights = [i for i in insights if i['category'] in selected_categories]
            
            severities = ['critical', 'high', 'warning', 'info']
            selected_severity = st.multiselect(
                "Filter by Severity:",
                options=severities,
                default=['critical', 'high', 'warning'],
                format_func=lambda x: x.title()
            )
            
            filtered_insights = [i for i in filtered_insights if i['severity'] in selected_severity]
            
            st.markdown(f"**Showing {len(filtered_insights)} insights**")
            
            for insight in filtered_insights:
                display_insight_card(insight, projects_map)
        else:
            st.info("No insights generated yet. Complete the analysis to see insights.")
        
        if summary.get('top_concerns'):
            st.markdown('<p class="section-header">🚨 Top Portfolio Concerns</p>', unsafe_allow_html=True)
            for concern in summary['top_concerns']:
                st.markdown(f"""
                <div class="insight-box insight-warning">
                    <p style="margin: 0; padding-left: 12px; border-left: 4px solid #ea580c; color: #1a202c;">{concern}</p>
                </div>
                """, unsafe_allow_html=True)
        
        if summary.get('critical_issues'):
            st.markdown('<p class="section-header">🚨 Critical Issues Requiring Attention</p>', unsafe_allow_html=True)
            for issue in summary['critical_issues']:
                st.markdown(f"""
                <div class="insight-box insight-critical" style="border-left: 5px solid #dc2626;">
                    <p style="margin: 8px 0; color: #1a202c;"><strong>🎯 {issue['project_name']}</strong></p>
                    <p style="margin: 8px 0; color: #1a202c;">Project ID: <code>{issue['project_id']}</code></p>
                    <p style="margin: 12px 0; padding: 8px; background: rgba(220, 38, 38, 0.1); border-radius: 4px; border-left: 3px solid #dc2626; color: #7f1d1d;">{issue['issue']}</p>
                    <p style="margin: 8px 0; padding-left: 12px; border-left: 3px solid #10b981; color: #1a202c;"><strong>✅ Recommendation:</strong> {issue['recommendation']}</p>
                </div>
                """, unsafe_allow_html=True)
        
        st.markdown('<p class="section-header">📈 Portfolio Visualizations</p>', unsafe_allow_html=True)
        
        viz_tabs = st.tabs([
            "Status Distribution",
            "Health Distribution",
            "Budget Variance",
            "Schedule Variance",
            "Data Completeness"
        ])
        
        with viz_tabs[0]:
            fig = create_status_distribution_chart(summary)
            if fig:
                st.plotly_chart(fig, use_container_width=True)
        
        with viz_tabs[1]:
            fig = create_health_distribution_chart(summary)
            if fig:
                st.plotly_chart(fig, use_container_width=True)
        
        with viz_tabs[2]:
            fig = create_budget_variance_chart(projects)
            if fig:
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Insufficient budget data for visualization")
        
        with viz_tabs[3]:
            fig = create_schedule_variance_chart(projects)
            if fig:
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Insufficient schedule data for visualization")
        
        with viz_tabs[4]:
            fig = create_data_completeness_chart(summary)
            if fig:
                st.plotly_chart(fig, use_container_width=True)
        
        st.markdown('<p class="section-header">📋 Project Details</p>', unsafe_allow_html=True)
        
        project_list = []
        for project_id, project_data in projects.items():
            metadata = project_data.get('project_metadata', {})
            assessment = project_data.get('assessment', {})
            overall = assessment.get('overall_assessment', {})
            
            project_list.append({
                'Project ID': project_id,
                'Project Name': metadata.get('project_name', 'Unknown'),
                'Status': overall.get('status', 'Unknown'),
                'Health': overall.get('health', 'Unknown'),
                'Confidence': overall.get('confidence_level', 'Unknown')
            })
        
        df_projects = pd.DataFrame(project_list)
        
        st.dataframe(df_projects, use_container_width=True, hide_index=True)
        
        st.markdown("### 🔍 Detailed Project Analysis")
        
        selected_project = st.selectbox(
            "Select a project to view detailed assessment:",
            options=[p['Project ID'] for p in project_list],
            format_func=lambda x: f"{x} - {next((p['Project Name'] for p in project_list if p['Project ID'] == x), 'Unknown')}"
        )
        
        if selected_project:
            st.markdown("---")
            display_project_assessment(projects[selected_project])
            
            st.markdown("---")
            st.markdown("### 🎯 Persona-Based Insights for This Project")
            
            project_persona = st.radio(
                "Select persona to view project-specific insights:",
                ["Executive (C-Level)", "VP / Portfolio Owner", "Manager / Delivery Lead"],
                horizontal=True,
                key=f"project_persona_{selected_project}"
            )
            
            if project_persona == "Executive (C-Level)":
                project_insights = engine.get_project_executive_insights(selected_project)
                st.markdown("**🎯 Strategic Insights for This Project**")
            elif project_persona == "VP / Portfolio Owner":
                project_insights = engine.get_project_vp_insights(selected_project)
                st.markdown("**📈 Portfolio Management Insights for This Project**")
            else:
                project_insights = engine.get_project_manager_insights(selected_project)
                st.markdown("**🔧 Operational Insights for This Project**")
            
            if project_insights:
                st.markdown(f"**Showing {len(project_insights)} project-level insights**")
                
                for insight in project_insights:
                    display_insight_card(insight, projects_map)
            else:
                st.info("No project-specific insights for this persona.")
        
        st.markdown('<p class="section-header">📊 Portfolio Summary Report</p>', unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("**Status Distribution**")
            status_dist = summary.get('status_distribution', {})
            for status, count in status_dist.items():
                st.metric(status, count)
        
        with col2:
            st.markdown("**Portfolio Risks**")
            risks = summary.get('portfolio_risks', [])
            if risks:
                for risk in risks:
                    severity_class = 'insight-critical' if risk['severity'] == 'high' else 'insight-warning'
                    st.markdown(f"""
                    <div class="insight-box {severity_class}">
                        <p style="margin: 8px 0; color: #1a202c;"><strong>⚠️ {risk['risk'].upper()}</strong></p>
                        <p style="margin: 8px 0; color: #1a202c;">{risk['description']}</p>
                        <p style="margin: 8px 0; padding: 8px; background: rgba(0, 0, 0, 0.05); border-radius: 4px; color: #1a202c;"><strong>Impact:</strong> {risk['impact']}</p>
                    </div>
                    """, unsafe_allow_html=True)
            else:
                st.success("No portfolio-level risks identified")
        
        st.markdown('<p class="section-header">💾 Export Results</p>', unsafe_allow_html=True)
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            portfolio_json = json.dumps({
                'portfolio_summary': summary,
                'projects': projects,
                'insights': {
                    'executive': engine.get_executive_insights(),
                    'vp': engine.get_vp_insights(),
                    'manager': engine.get_manager_insights()
                }
            }, indent=2, default=str)
            
            st.download_button(
                label="📥 Download Portfolio Analysis",
                data=portfolio_json,
                file_name=f"portfolio_analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
                mime="application/json"
            )
        
        with col2:
            csv = df_projects.to_csv(index=False)
            st.download_button(
                label="📥 Download Project List (CSV)",
                data=csv,
                file_name=f"project_list_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv"
            )
        
        with col3:
            if selected_project:
                project_json = json.dumps(projects[selected_project], indent=2, default=str)
                st.download_button(
                    label=f"📥 Download {selected_project} Details",
                    data=project_json,
                    file_name=f"project_{selected_project}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
                    mime="application/json"
                )
    
    else:
        st.info("👈 Upload data files from the sidebar to begin portfolio analysis")
        
        st.markdown("""
        ## 🔬 Enterprise Portfolio Analytics
        
        ### Three-System Analysis Approach
        
        This dashboard analyzes projects from three critical data sources:
        
        1. **📘 Smartsheet (Baseline Truth)**
           - Project definition and ownership
           - Approved budgets (CapEx, OpEx, EAC)
           - Baseline schedule (start, finish, duration)
           - Official health indicators (schedule, budget, risk, quality)
           - RAID logs, approvals, prioritization
        
        2. **📙 Wave (Management Perception)**
           - Weekly status snapshots
           - Delivery stage and lifecycle progression
           - IT delivery forecasts
           - Status trends over time
           - Interdependencies
        
        3. **📗 Tick (Execution Reality)**
           - Actual work performed
           - Hours logged by resources
           - Actual costs incurred
           - Vendor and task-level execution
        
        ### 🎯 Persona-Based Insights (Portfolio & Project Level)
        
        **10 Decision-Grade Insight Categories:**
        1. Strategic Alignment (strategy vs execution vs effort)
        2. Value Leakage & Value Realization
        3. Execution Health & Delivery Risk
        4. Resource Utilization & Productivity
        5. Portfolio Prioritization (accelerate / pause / stop)
        6. Governance & Ownership Gaps
        7. Time, Velocity & Execution Drag
        8. Predictive Risk Signals
        9. Data Hygiene & Integrity Issues
        10. Change-Management & Behavioral Patterns
        
        **Tailored for 3 Personas:**
        - 🎯 **C-Level Executive**: Strategic alignment, portfolio-wide trends, investment decisions
        - 📈 **VP / Portfolio Owner**: Risk management, resource allocation, prioritization
        - 🔧 **Manager / Delivery Lead**: Operational execution, team productivity, data quality
        
        **Data Accuracy & Governance:**
        - Conservative approach: No false positives from missing data
        - Evidence-based only: No hallucination or inference
        - Confidence degradation: Explicit when data is incomplete
        - Project and portfolio level insights: Separate views for different needs
        
        ### 🎨 Analysis Capabilities
        
        **Cross-Source Validation:**
        - Compares baseline vs forecast vs actuals
        - Identifies status vs reality mismatches
        - Detects inconsistencies in reporting
        
        **Evidence-Based Insights:**
        - Only fact-based analysis
        - Explicit about data gaps and limitations
        - Confidence levels for all assessments
        
        **Risk Detection:**
        - Budget overrun early warnings
        - Schedule delay indicators
        - Status-to-actuals consistency checks
        - Burn rate projections
        
        **Decision Support:**
        - Portfolio-level metrics and trends
        - Project-specific drill-down
        - Actionable recommendations
        - Critical issue flagging
        
        ### 📋 How to Use
        
        1. **Upload** data exports from Smartsheet, Wave, and/or Tick
        2. **Configure** sheet names if needed (optional)
        3. **Analyze** - AI automatically correlates and analyzes projects
        4. **Select Persona** - View insights tailored to your role
        5. **Explore** portfolio dashboards and project details
        6. **Drill Down** - Click on a project to see persona-specific insights for that project
        7. **Export** results for reporting or further analysis
        
        ### ✨ Key Features
        
        - **Universal Correlation**: Automatically matches projects across systems
        - **Multi-Source**: Works with 1, 2, or all 3 data sources
        - **Evidence-Based**: Every insight tied to specific metrics
        - **Consistency Rules**: Validates cross-source data quality
        - **Executive-Ready**: Clear, actionable summaries
        - **Persona-Driven**: Insights mapped to decision-maker roles at both portfolio and project level
        - **Data Accuracy**: Conservative governance prevents false positives
        
        ### 🔍 Analysis Principles
        
        - Accuracy over optimism
        - Evidence over assumptions
        - Contradictions are explicitly called out
        - Missing data is NOT treated as zero
        - Smartsheet = baseline truth
        - Tick = execution reality
        - Wave = management perception
        
        ---
        
        **Ready to start?** Upload your data files using the sidebar controls.
        """)


if __name__ == "__main__":
    main()