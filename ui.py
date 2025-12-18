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
        .main-header {
            font-size: 2.5rem;
            font-weight: bold;
            color: #1e3a8a;
            margin-bottom: 0.5rem;
            text-align: center;
        }
        .section-header {
            font-size: 1.8rem;
            font-weight: 600;
            color: #1e40af;
            margin-top: 2rem;
            margin-bottom: 1rem;
            border-bottom: 3px solid #3b82f6;
            padding-bottom: 0.5rem;
        }
        .project-card {
            background: white;
            padding: 1.5rem;
            border-radius: 10px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            border-left: 5px solid #3b82f6;
            margin: 1rem 0;
        }
        .status-on-track {
            color: #10b981;
            font-weight: bold;
        }
        .status-at-risk {
            color: #f59e0b;
            font-weight: bold;
        }
        .status-delayed {
            color: #ef4444;
            font-weight: bold;
        }
        .metric-container {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            padding: 1.5rem;
            border-radius: 10px;
            color: white;
            text-align: center;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }
        .alert-critical {
            background-color: #fee2e2;
            border-left: 5px solid #dc2626;
            padding: 1.5rem;
            border-radius: 8px;
            margin: 1rem 0;
        }
        .alert-warning {
            background-color: #fef3c7;
            border-left: 5px solid #f59e0b;
            padding: 1.5rem;
            border-radius: 8px;
            margin: 1rem 0;
        }
        .alert-info {
            background-color: #dbeafe;
            border-left: 5px solid #3b82f6;
            padding: 1.5rem;
            border-radius: 8px;
            margin: 1rem 0;
        }
        .insight-box {
            background: linear-gradient(135deg, #3b82f6 0%, #1e40af 100%);
            color: white;
            padding: 2rem;
            border-radius: 10px;
            font-size: 1.05rem;
            line-height: 1.8;
            margin: 1.5rem 0;
        }
        .data-source-badge {
            display: inline-block;
            padding: 0.25rem 0.75rem;
            border-radius: 12px;
            font-size: 0.85rem;
            font-weight: 600;
            margin: 0.25rem;
        }
        .source-smartsheet {
            background-color: #dbeafe;
            color: #1e40af;
        }
        .source-wave {
            background-color: #fef3c7;
            color: #92400e;
        }
        .source-tick {
            background-color: #d1fae5;
            color: #065f46;
        }
        .persona-tab {
            background-color: #f3f4f6;
            padding: 1rem;
            border-radius: 8px;
            margin: 0.5rem 0;
        }
        .insight-category-badge {
            display: inline-block;
            padding: 0.5rem 1rem;
            border-radius: 20px;
            font-size: 0.9rem;
            font-weight: 600;
            margin: 0.25rem;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
        }
        .confidence-badge {
            display: inline-block;
            padding: 0.25rem 0.75rem;
            border-radius: 12px;
            font-size: 0.8rem;
            font-weight: 600;
            margin: 0.25rem;
            background-color: #e0e7ff;
            color: #4338ca;
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
        'Unknown': '#6b7280'
    }
    
    labels = list(status_dist.keys())
    values = list(status_dist.values())
    chart_colors = [colors.get(label, '#6b7280') for label in labels]
    
    fig = go.Figure(data=[go.Pie(
        labels=labels,
        values=values,
        marker=dict(colors=chart_colors),
        textinfo='label+percent',
        textfont_size=14
    )])
    
    fig.update_layout(
        title='Portfolio Status Distribution',
        height=400
    )
    
    return fig


def create_health_distribution_chart(summary):
    """Create bar chart of health indicators"""
    health_dist = summary.get('health_distribution', {})
    
    if not health_dist:
        return None
    
    colors_map = {
        'Green': '#10b981',
        'Yellow': '#f59e0b',
        'Red': '#ef4444',
        'Unknown': '#6b7280'
    }
    
    df = pd.DataFrame(list(health_dist.items()), columns=['Health', 'Count'])
    df['Color'] = df['Health'].map(colors_map)
    
    fig = go.Figure(data=[go.Bar(
        x=df['Health'],
        y=df['Count'],
        marker_color=df['Color'],
        text=df['Count'],
        textposition='outside'
    )])
    
    fig.update_layout(
        title='Health Indicator Distribution',
        xaxis_title='Health Status',
        yaxis_title='Number of Projects',
        height=400
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
                'Variance %': cost_var
            })
    
    if not data:
        return None
    
    df = pd.DataFrame(data)
    df = df.sort_values('Variance %')
    
    colors = ['#ef4444' if x < -10 else '#10b981' if x > 5 else '#f59e0b' for x in df['Variance %']]
    
    fig = go.Figure(data=[go.Bar(
        x=df['Project'],
        y=df['Variance %'],
        marker_color=colors,
        text=df['Variance %'].apply(lambda x: f"{x:.1f}%"),
        textposition='outside'
    )])
    
    fig.update_layout(
        title='Budget Variance by Project',
        xaxis_title='Projects',
        yaxis_title='Variance % (negative = overrun)',
        height=500,
        xaxis={'tickangle': -45}
    )
    
    fig.add_hline(y=0, line_dash="dash", line_color="gray")
    
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
                'Delay (Days)': schedule_var
            })
    
    if not data:
        return None
    
    df = pd.DataFrame(data)
    df = df.sort_values('Delay (Days)', ascending=False)
    
    colors = ['#ef4444' if x > 30 else '#10b981' if x <= 0 else '#f59e0b' for x in df['Delay (Days)']]
    
    fig = go.Figure(data=[go.Bar(
        x=df['Project'],
        y=df['Delay (Days)'],
        marker_color=colors,
        text=df['Delay (Days)'].apply(lambda x: f"{x:.0f}d"),
        textposition='outside'
    )])
    
    fig.update_layout(
        title='Schedule Variance by Project',
        xaxis_title='Projects',
        yaxis_title='Days (positive = delayed)',
        height=500,
        xaxis={'tickangle': -45}
    )
    
    fig.add_hline(y=0, line_dash="dash", line_color="gray")
    
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
        text=values,
        textposition='outside'
    )])
    
    fig.update_layout(
        title='Data Source Completeness',
        xaxis_title='Coverage Level',
        yaxis_title='Number of Projects',
        height=400
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


def display_insight_card(insight: dict):
    """Display a single insight card with appropriate styling"""
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
    
    st.markdown(f"""
    <div class="{alert_class}">
        <h4>{icon} {insight['title']}</h4>
        <p><strong>Category:</strong> <span class="insight-category-badge">{insight['category'].replace('_', ' ').title()}</span></p>
        <p><strong>Confidence:</strong> <span class="confidence-badge">{confidence}</span></p>
        <p><strong>Description:</strong> {insight['description']}</p>
        <p><strong>Impact:</strong> {insight['impact']}</p>
        <p><strong>Recommendation:</strong> {insight['recommendation']}</p>
    </div>
    """, unsafe_allow_html=True)
    
    if insight.get('metrics'):
        with st.expander("📊 Supporting Metrics"):
            st.json(insight['metrics'])


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
        st.info(overall['summary'])
    
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


def main():
    """Main application"""
    
    st.set_page_config(
        page_title="Enterprise Portfolio Analytics",
        page_icon="📊",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    load_custom_css()
    
    st.markdown('<p class="main-header">📊 Enterprise Portfolio Analytics</p>', unsafe_allow_html=True)
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
                display_insight_card(insight)
        else:
            st.info("No insights generated yet. Complete the analysis to see insights.")
        
        if summary.get('top_concerns'):
            st.markdown("### 🚨 Top Portfolio Concerns")
            for concern in summary['top_concerns']:
                st.warning(concern)
        
        if summary.get('critical_issues'):
            st.markdown("### ⚠️ Critical Issues Requiring Attention")
            for issue in summary['critical_issues']:
                st.markdown(f"""
                <div class="alert-critical">
                    <strong>{issue['project_name']}</strong> ({issue['project_id']})<br/>
                    {issue['issue']}<br/>
                    <em>Recommendation: {issue['recommendation']}</em>
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
                    display_insight_card(insight)
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
                    severity_color = 'critical' if risk['severity'] == 'high' else 'warning'
                    st.markdown(f"""
                    <div class="alert-{severity_color}">
                        <strong>{risk['risk'].upper()}</strong><br/>
                        {risk['description']}<br/>
                        <em>Impact: {risk['impact']}</em>
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