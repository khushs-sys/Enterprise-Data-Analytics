"""
Enterprise Portfolio Analytics AI Engine - Enhanced Edition with Explicit Formulas
Analyzes multi-source project data: Smartsheet (baseline) + Wave (forecast) + Tick (actuals)

NEW FEATURES:
- 16 Explicit Formula-Based Insights (Tier 1-4)
- Conservative Data Governance (No hallucination)
- Formula Traceability & Audit Trail
- Confidence Degradation on Missing Data
"""

import json
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from typing import Dict, List, Any, Optional, Tuple
from collections import defaultdict
from difflib import SequenceMatcher
import warnings
warnings.filterwarnings('ignore')


class PortfolioAIEngine:
    """
    Enterprise Portfolio Analytics Engine with Formula-Based Insights
    
    Analyzes projects from three data sources:
    1. Smartsheet: Baseline truth (approved plan, budget, governance)
    2. Wave: Management perception (weekly snapshots, forecasts, status)
    3. Tick: Execution reality (actual hours, costs, work performed)
    
    NEW: 16 Explicit Formula-Based Insights (Tier 1-4)
    TIER-1: Board-Level (Value Leakage, Strategy Coverage, Top/Bottom Analysis)
    TIER-2: Portfolio & P&L (Cost per Outcome, Execution Drag, Investment Map)
    TIER-3: Operational Excellence (Effort Mismatch, Utilization, Span Control)
    TIER-4: Execution Hygiene (Phantom Work, Task Hygiene, Idle Capacity)
    """
    
    def __init__(self):
        """Initialize the Portfolio AI Engine"""
        self.smartsheet_data = None
        self.wave_data = None
        self.tick_data = None
        self.projects = {}
        self.portfolio_insights = {}
        
        # Column mappings (auto-detected)
        self.smartsheet_cols = {}
        self.wave_cols = {}
        self.tick_cols = {}
        
        # Insight storage by persona (with formula traceability)
        self.executive_insights = []
        self.vp_insights = []
        self.manager_insights = []
        
    def _is_valid_project_id(self, project_id) -> bool:
        """Check if project_id is valid (not None, empty, or Unknown)"""
        if project_id is None:
            return False
        if isinstance(project_id, str):
            normalized = str(project_id).strip().upper()
            if not normalized or normalized == 'UNKNOWN' or normalized == 'NOT SPECIFIED':
                return False
        return True
    
    def _detect_column(self, df: pd.DataFrame, patterns: List[str], col_type: str = "column") -> Optional[str]:
        """
        Intelligently detect a column by trying multiple patterns
        """
        if df is None or df.empty:
            return None
        
        for col in df.columns:
            col_lower = str(col).lower().strip()
            for pattern in patterns:
                if col_lower == pattern.lower().strip():
                    return col
        
        for col in df.columns:
            col_lower = str(col).lower().strip()
            for pattern in patterns:
                pattern_lower = pattern.lower().strip()
                if pattern_lower in col_lower or col_lower in pattern_lower:
                    return col
        
        for col in df.columns:
            col_normalized = str(col).lower().strip().replace('_', '').replace(' ', '').replace('-', '')
            for pattern in patterns:
                pattern_normalized = pattern.lower().strip().replace('_', '').replace(' ', '').replace('-', '')
                if col_normalized == pattern_normalized or pattern_normalized in col_normalized:
                    return col
        
        return None
    
    def _detect_columns_for_source(self, df: pd.DataFrame, source: str) -> Dict[str, str]:
        """Detect all relevant columns for a data source"""
        
        detected = {}
        
        id_patterns = [
            'project_id', 'projectid', 'project id', 'project code', 'project_code',
            'id', 'project', 'project_name', 'project name', 'name', 'projectname',
            'project no', 'project_no', 'projectno', 'project number', 'project_number',
            'wbs', 'wbs code', 'wbs_code'
        ]
        detected['id'] = self._detect_column(df, id_patterns, "project ID")
        
        wave_num_patterns = ['wave #', 'wave#', 'wave_#', 'wave', 'wave_number', 'wave number', '#']
        detected['wave_num'] = self._detect_column(df, wave_num_patterns, "wave number")
        
        name_patterns = [
            'project_name', 'project name', 'projectname', 'name',
            'project title', 'project_title', 'title', 'description'
        ]
        name_col = self._detect_column(df, name_patterns, "project name")
        if name_col and name_col != detected['id']:
            detected['name'] = name_col
        
        start_patterns = [
            'baseline_start', 'baseline start', 'start_date', 'start date', 'start',
            'planned_start', 'planned start', 'begin_date', 'begin date',
            'project_start', 'project start', 'estimated start date', 'estimated start date of it work'
        ]
        detected['start_date'] = self._detect_column(df, start_patterns, "start date")
        
        finish_patterns = [
            'baseline_finish', 'baseline finish', 'baseline_end', 'baseline end',
            'end_date', 'end date', 'finish_date', 'finish date', 'finish', 'end',
            'planned_finish', 'planned finish', 'planned_end', 'planned end',
            'project_end', 'project end', 'completion_date', 'completion date',
            'estimated end date', 'estimated end date of it work'
        ]
        detected['finish_date'] = self._detect_column(df, finish_patterns, "finish date")
        
        forecast_finish_patterns = [
            'forecast_finish', 'forecast finish', 'forecast_end', 'forecast end',
            'forecasted_finish', 'forecasted finish', 'estimated_finish', 'estimated finish',
            'projected_finish', 'projected finish', 'forecast completion', 'forecast_completion',
            'l4 forecast date', 'l4 forecast', 'l4_forecast_date'
        ]
        detected['forecast_finish'] = self._detect_column(df, forecast_finish_patterns, "forecast finish")
        
        budget_patterns = [
            'total_budget', 'total budget', 'budget', 'approved_budget', 'approved budget',
            'baseline_budget', 'baseline budget', 'planned_budget', 'planned budget',
            'total_cost', 'total cost', 'project_budget', 'project budget',
            'overall total budget', 'total overall it costs'
        ]
        detected['budget'] = self._detect_column(df, budget_patterns, "budget")
        
        capex_patterns = ['capex', 'cap_ex', 'capital_expense', 'capital expense', 'total budget capex']
        detected['capex'] = self._detect_column(df, capex_patterns, "capex")
        
        opex_patterns = ['opex', 'op_ex', 'operating_expense', 'operating expense', 'total budget opex']
        detected['opex'] = self._detect_column(df, opex_patterns, "opex")
        
        eac_patterns = ['eac', 'total eac', 'estimate at completion']
        detected['eac'] = self._detect_column(df, eac_patterns, "eac")
        
        actual_cost_patterns = [
            'actual_cost', 'actual cost', 'cost', 'amount', 'actual_amount', 'actual amount',
            'actuals', 'spent', 'expenditure', 'total_actual', 'total actual', 'total'
        ]
        detected['actual_cost'] = self._detect_column(df, actual_cost_patterns, "actual cost")
        
        hours_patterns = [
            'planned_hours', 'planned hours', 'hours', 'effort', 'estimated_hours', 'estimated hours',
            'baseline_hours', 'baseline hours', 'total_hours', 'total hours'
        ]
        detected['hours'] = self._detect_column(df, hours_patterns, "hours")
        
        actual_hours_patterns = [
            'actual_hours', 'actual hours', 'hours', 'worked_hours', 'worked hours',
            'time_spent', 'time spent', 'logged_hours', 'logged hours'
        ]
        detected['actual_hours'] = self._detect_column(df, actual_hours_patterns, "actual hours")
        
        status_patterns = [
            'status', 'project_status', 'project status', 'state', 'phase',
            'current_status', 'current status', 'rag', 'rag_status', 'rag status',
            'weekly status', 'high level status'
        ]
        detected['status'] = self._detect_column(df, status_patterns, "status")
        
        schedule_health_patterns = [
            'schedule_health', 'schedule health', 'schedule_status', 'schedule status',
            'schedule_rag', 'schedule rag', 'timeline_status', 'timeline status'
        ]
        detected['schedule_health'] = self._detect_column(df, schedule_health_patterns, "schedule health")
        
        budget_health_patterns = [
            'budget_health', 'budget health', 'budget_status', 'budget status',
            'budget_rag', 'budget rag', 'cost_status', 'cost status', 'financial_health'
        ]
        detected['budget_health'] = self._detect_column(df, budget_health_patterns, "budget health")
        
        risk_patterns = [
            'risk_level', 'risk level', 'risk', 'risk_status', 'risk status',
            'risk_rag', 'risk rag', 'overall_risk', 'overall risk', 'risk health'
        ]
        detected['risk'] = self._detect_column(df, risk_patterns, "risk level")
        
        snapshot_date_patterns = [
            'snapshot_date', 'snapshot date', 'report_date', 'report date',
            'as_of_date', 'as of date', 'date', 'week', 'reporting_date'
        ]
        detected['snapshot_date'] = self._detect_column(df, snapshot_date_patterns, "snapshot date")
        
        completion_patterns = [
            'completion_pct', 'completion pct', 'completion', 'percent_complete', 'percent complete',
            '% complete', 'pct_complete', 'progress', 'completion_%', 'completion %',
            'overall % complete', 'pct_complete_normalized'
        ]
        detected['completion'] = self._detect_column(df, completion_patterns, "completion %")
        
        stage_patterns = ['stage', 'lifecycle', 'lifecycle_stage', 'delivery_stage', 'phase', 'wave stage']
        detected['stage'] = self._detect_column(df, stage_patterns, "stage")
        
        owner_patterns = [
            'owner', 'project_owner', 'project owner', 'manager', 'project_manager',
            'project manager', 'pm', 'responsible', 'lead', 'it project manager',
            'initiative owner', 'accountable workstream'
        ]
        detected['owner'] = self._detect_column(df, owner_patterns, "owner")
        
        strategic_patterns = [
            'strategic_alignment', 'strategic alignment', 'strategic goal', 'strategy',
            'strategic_goal', 'priority', 'prioritization score'
        ]
        detected['strategic_alignment'] = self._detect_column(df, strategic_patterns, "strategic alignment")
        
        benefit_patterns = [
            'benefit', 'benefits', 'net recurring benefits', 'one-time benefits',
            '5 yr rev impact', '5 yr cost savings', 'benefit quantification'
        ]
        detected['benefits'] = self._detect_column(df, benefit_patterns, "benefits")
        
        value_lever_patterns = [
            'value_lever', 'value lever', 'value_driver', 'value driver', 'outcome', 'business outcome'
        ]
        detected['value_lever'] = self._detect_column(df, value_lever_patterns, "value lever")
        
        approval_date_patterns = [
            'approval_date', 'approval date', 'approved_date', 'approved date', 'start_approval'
        ]
        detected['approval_date'] = self._detect_column(df, approval_date_patterns, "approval date")
        
        interdep_patterns = ['interdependencies', 'dependencies', 'project interdependencies']
        detected['interdependencies'] = self._detect_column(df, interdep_patterns, "interdependencies")
        
        complexity_patterns = ['complexity', 'it implementation complexity', 'implementation complexity']
        detected['complexity'] = self._detect_column(df, complexity_patterns, "complexity")
        
        task_patterns = ['task', 'task_name', 'task name', 'activity', 'work_item']
        detected['task'] = self._detect_column(df, task_patterns, "task")
        
        resource_patterns = ['resource', 'user', 'assigned_to', 'assigned to', 'team_member']
        detected['resource'] = self._detect_column(df, resource_patterns, "resource")
        
        detected = {k: v for k, v in detected.items() if v is not None}
        
        return detected
        
    def load_smartsheet(self, df: pd.DataFrame):
        """Load Smartsheet data (baseline & governance)"""
        if isinstance(df, dict):
            df = list(df.values())[0]
        
        self.smartsheet_data = df.copy()
        self.smartsheet_cols = self._detect_columns_for_source(df, 'smartsheet')
        
        print(f"✅ Loaded Smartsheet: {len(df)} projects")
        if self.smartsheet_cols.get('id'):
            print(f"   📌 Project ID column: '{self.smartsheet_cols['id']}'")
        else:
            print(f"   ⚠️  No project ID column detected - using first column")
            self.smartsheet_cols['id'] = df.columns[0]
        
        other_cols = [f"{k}:'{v}'" for k, v in self.smartsheet_cols.items() if k != 'id' and k in ['budget', 'start_date', 'finish_date', 'wave_num']]
        if other_cols:
            print(f"   📊 Key columns: {', '.join(other_cols[:4])}")
        
    def load_wave(self, df: pd.DataFrame):
        """Load Wave data (weekly snapshots & forecasts)"""
        if isinstance(df, dict):
            df = list(df.values())[0]
        
        self.wave_data = df.copy()
        self.wave_cols = self._detect_columns_for_source(df, 'wave')
        
        print(f"✅ Loaded Wave: {len(df)} snapshots")
        if self.wave_cols.get('id'):
            print(f"   📌 Project ID column: '{self.wave_cols['id']}'")
        elif self.wave_cols.get('wave_num'):
            print(f"   📌 Wave # column: '{self.wave_cols['wave_num']}'")
        else:
            print(f"   ⚠️  No project ID column detected - using first column")
            self.wave_cols['id'] = df.columns[0]
        
        other_cols = [f"{k}:'{v}'" for k, v in self.wave_cols.items() if k != 'id' and k in ['status', 'snapshot_date', 'forecast_finish', 'stage']]
        if other_cols:
            print(f"   📊 Key columns: {', '.join(other_cols[:4])}")
        
    def load_tick(self, df: pd.DataFrame):
        """Load Tick data (actual execution)"""
        if isinstance(df, dict):
            df = list(df.values())[0]
        
        self.tick_data = df.copy()
        self.tick_cols = self._detect_columns_for_source(df, 'tick')
        
        print(f"✅ Loaded Tick: {len(df)} actuals")
        if self.tick_cols.get('id'):
            print(f"   📌 Project ID column: '{self.tick_cols['id']}'")
        elif self.tick_cols.get('wave_num'):
            print(f"   📌 Wave column: '{self.tick_cols['wave_num']}'")
        else:
            print(f"   ⚠️  No project ID column detected - using first column")
            self.tick_cols['id'] = df.columns[0]
        
        other_cols = [f"{k}:'{v}'" for k, v in self.tick_cols.items() if k != 'id' and k in ['actual_cost', 'actual_hours', 'wave_num']]
        if other_cols:
            print(f"   📊 Key columns: {', '.join(other_cols[:3])}")
    
    def _normalize_project_id(self, project_id):
        """Normalize project ID for matching across systems"""
        if pd.isna(project_id):
            return None
        return str(project_id).strip().upper()
    
    def _fuzzy_match_score(self, s1: str, s2: str) -> float:
        """Calculate similarity score between two strings"""
        if pd.isna(s1) or pd.isna(s2):
            return 0.0
        s1 = str(s1).lower().strip()
        s2 = str(s2).lower().strip()
        return SequenceMatcher(None, s1, s2).ratio()
    
    def _safe_get(self, df, column, default=None):
        """Safely get column value with fallback"""
        if df is None or column not in df.columns:
            return default
        val = df[column].iloc[0] if len(df) > 0 else default
        return val if pd.notna(val) else default
    
    def _safe_numeric(self, value, default=None):
        """Safely convert to numeric, return None if not available"""
        try:
            if pd.isna(value):
                return None
            if isinstance(value, str):
                value = value.replace(',', '').replace('$', '').replace('€', '').strip()
            return float(value)
        except:
            return None
    
    def _safe_date(self, value):
        """Safely convert to date"""
        try:
            if pd.isna(value):
                return None
            if isinstance(value, (datetime, pd.Timestamp)):
                return value
            return pd.to_datetime(value)
        except:
            return None
    
    def _calculate_variance_pct(self, actual, baseline):
        """Calculate variance percentage - returns None if inputs invalid"""
        if baseline is None or baseline == 0 or actual is None:
            return None
        return ((actual - baseline) / abs(baseline)) * 100
    
    def _classify_health(self, schedule_health, budget_health, risk_level):
        """Classify overall project health"""
        health_map = {
            'green': 3, 'yellow': 2, 'red': 1,
            'on track': 3, 'at risk': 2, 'delayed': 1,
            'low': 3, 'medium': 2, 'high': 1
        }
        
        schedule_score = health_map.get(str(schedule_health).lower().strip(), 2) if schedule_health else 2
        budget_score = health_map.get(str(budget_health).lower().strip(), 2) if budget_health else 2
        risk_score = health_map.get(str(risk_level).lower().strip(), 2) if risk_level else 2
        
        avg_score = (schedule_score + budget_score + risk_score) / 3
        
        if avg_score >= 2.5:
            return "On Track"
        elif avg_score >= 1.5:
            return "At Risk"
        else:
            return "Delayed"
    
    def _get_latest_wave_snapshot(self, project_id):
        """Get most recent Wave snapshot for project"""
        if self.wave_data is None:
            return None
        
        if self.wave_cols.get('wave_num'):
            wave_col = self.wave_cols['wave_num']
            project_waves = self.wave_data[
                self.wave_data[wave_col].apply(self._normalize_project_id) == 
                self._normalize_project_id(project_id)
            ]
            
            if len(project_waves) > 0:
                return project_waves.iloc[-1]
        
        if self.wave_cols.get('name'):
            name_col = self.wave_cols['name']
            best_match_idx = None
            best_score = 0.6
            
            for idx, wave_name in self.wave_data[name_col].items():
                score = self._fuzzy_match_score(project_id, wave_name)
                if score > best_score:
                    best_score = score
                    best_match_idx = idx
            
            if best_match_idx is not None:
                return self.wave_data.loc[best_match_idx]
        
        return None
    
    def _get_wave_trends(self, project_id, lookback_weeks=8):
        """Analyze Wave trends over time"""
        if self.wave_data is None:
            return None
        
        if self.wave_cols.get('wave_num'):
            wave_col = self.wave_cols['wave_num']
            project_waves = self.wave_data[
                self.wave_data[wave_col].apply(self._normalize_project_id) == 
                self._normalize_project_id(project_id)
            ]
        else:
            return None
        
        if len(project_waves) < 2:
            return None
        
        trends = {
            'snapshot_count': len(project_waves),
            'date_range': {}
        }
        
        if self.wave_cols.get('status'):
            status_col = self.wave_cols['status']
            status_changes = project_waves[status_col].value_counts().to_dict()
            trends['status_distribution'] = status_changes
            
            statuses = project_waves[status_col].astype(str).str.lower().tolist()
            if len(statuses) >= 2:
                if 'red' in statuses[-2:] or 'delayed' in statuses[-2:]:
                    trends['recent_deterioration'] = True
        
        return trends
    
    def _get_tick_actuals(self, project_id):
        """Get aggregated Tick actuals for project"""
        if self.tick_data is None:
            return None
        
        if self.tick_cols.get('wave_num'):
            wave_col = self.tick_cols['wave_num']
            project_ticks = self.tick_data[
                self.tick_data[wave_col].apply(self._normalize_project_id) == 
                self._normalize_project_id(project_id)
            ]
        elif self.tick_cols.get('id'):
            id_col = self.tick_cols['id']
            best_matches = []
            for proj_name in self.tick_data[id_col].unique():
                score = self._fuzzy_match_score(project_id, proj_name)
                if score > 0.6:
                    best_matches.append(proj_name)
            
            if best_matches:
                project_ticks = self.tick_data[self.tick_data[id_col].isin(best_matches)]
            else:
                return None
        else:
            return None
        
        if len(project_ticks) == 0:
            return None
        
        actuals = {
            'transaction_count': len(project_ticks)
        }
        
        if self.tick_cols.get('actual_hours'):
            hours_col = self.tick_cols['actual_hours']
            total_hours = project_ticks[hours_col].fillna(0).sum()
            actuals['total_hours'] = float(total_hours)
        elif self.tick_cols.get('hours'):
            hours_col = self.tick_cols['hours']
            total_hours = project_ticks[hours_col].fillna(0).sum()
            actuals['total_hours'] = float(total_hours)
        
        if self.tick_cols.get('actual_cost'):
            cost_col = self.tick_cols['actual_cost']
            total_cost = project_ticks[cost_col].fillna(0).sum()
            actuals['total_cost'] = float(total_cost)
        
        if 'user' in [c.lower() for c in project_ticks.columns]:
            user_col = [c for c in project_ticks.columns if c.lower() == 'user'][0]
            actuals['unique_resources'] = project_ticks[user_col].nunique()
        
        date_cols = [v for k, v in self.tick_cols.items() if 'date' in k.lower()]
        if date_cols:
            for col in date_cols:
                if col in project_ticks.columns:
                    dates = pd.to_datetime(project_ticks[col], errors='coerce').dropna()
                    if len(dates) > 0:
                        actuals['work_start_date'] = str(dates.min().date())
                        actuals['work_end_date'] = str(dates.max().date())
                        actuals['work_span_days'] = (dates.max() - dates.min()).days
                        break
        
        return actuals
    
    def _evaluate_consistency_rules(self, project_data: Dict) -> List[Dict]:
        """Evaluate cross-source consistency rules"""
        rules_violated = []
        
        baseline = project_data.get('baseline_metrics', {})
        wave = project_data.get('latest_wave_snapshot', {})
        actuals = project_data.get('actuals_summary', {})
        derived = project_data.get('derived_metrics', {})
        
        if wave:
            status = str(wave.get('status', '')).lower()
            cost_variance = derived.get('cost_variance_pct')
            if ('green' in status or 'on track' in status) and cost_variance and cost_variance < -10:
                rules_violated.append({
                    'rule': 'status_cost_mismatch',
                    'severity': 'warning',
                    'description': f"Status is '{wave.get('status')}' but cost variance is {cost_variance:.1f}%",
                    'recommendation': 'Review status accuracy or investigate cost drivers'
                })
        
        if baseline and actuals:
            baseline_budget = baseline.get('total_budget')
            actual_cost = actuals.get('total_cost')
            completion_pct = derived.get('completion_pct')
            
            if baseline_budget and baseline_budget > 0 and actual_cost and actual_cost > 0 and completion_pct and completion_pct > 0:
                projected_total = actual_cost / (completion_pct / 100)
                if projected_total > baseline_budget * 1.2:
                    rules_violated.append({
                        'rule': 'burn_rate_overrun',
                        'severity': 'critical',
                        'description': f"Current burn rate projects {projected_total/1000:.0f}K total cost vs {baseline_budget/1000:.0f}K budget",
                        'recommendation': 'Immediate budget review required'
                    })
        
        schedule_variance = derived.get('schedule_variance_days')
        if schedule_variance and schedule_variance > 30:
            schedule_health = baseline.get('schedule_health', '').lower() if baseline.get('schedule_health') else ''
            if 'green' in schedule_health or 'on track' in schedule_health:
                rules_violated.append({
                    'rule': 'schedule_health_mismatch',
                    'severity': 'warning',
                    'description': f"Schedule delayed {schedule_variance} days but health shows '{baseline.get('schedule_health')}'",
                    'recommendation': 'Update schedule health indicator'
                })
        
        if wave and not actuals:
            status = str(wave.get('status', '')).lower()
            if 'active' in status or 'in progress' in status or 'green' in status:
                rules_violated.append({
                    'rule': 'missing_actuals',
                    'severity': 'info',
                    'description': 'Project marked active but no execution data found',
                    'recommendation': 'Verify project has started or update status'
                })
        
        if baseline and actuals:
            completion = derived.get('completion_pct')
            actual_hours = actuals.get('total_hours')
            baseline_hours = baseline.get('planned_hours')
            
            if baseline_hours and baseline_hours > 0 and actual_hours and actual_hours > 0 and completion:
                implied_completion = (actual_hours / baseline_hours) * 100
                if abs(completion - implied_completion) > 20:
                    rules_violated.append({
                        'rule': 'completion_effort_mismatch',
                        'severity': 'warning',
                        'description': f"Reported {completion:.0f}% complete but hours suggest {implied_completion:.0f}%",
                        'recommendation': 'Reconcile completion % with actual effort'
                    })
        
        return rules_violated
    
    def _calculate_derived_metrics(self, baseline: Dict, wave: Dict, actuals: Dict) -> Dict:
        """Calculate derived metrics from source data"""
        derived = {}
        
        baseline_budget = baseline.get('total_budget')
        actual_cost = actuals.get('total_cost')
        eac = baseline.get('eac')
        
        if baseline_budget and baseline_budget > 0 and actual_cost and actual_cost > 0:
            derived['cost_variance_pct'] = self._calculate_variance_pct(actual_cost, baseline_budget)
            derived['cost_variance_amount'] = actual_cost - baseline_budget
        
        if baseline_budget and baseline_budget > 0 and eac and eac > 0:
            derived['eac_variance_pct'] = self._calculate_variance_pct(eac, baseline_budget)
            derived['eac_variance_amount'] = eac - baseline_budget
        
        baseline_end = self._safe_date(baseline.get('baseline_finish'))
        forecast_end = self._safe_date(wave.get('forecast_finish') if wave else None)
        if baseline_end and forecast_end:
            schedule_variance = (forecast_end - baseline_end).days
            derived['schedule_variance_days'] = schedule_variance
        
        if actual_cost and actual_cost > 0 and actuals.get('work_span_days'):
            if actuals['work_span_days'] > 0:
                derived['daily_burn_rate'] = actual_cost / actuals['work_span_days']
        
        if wave:
            derived['completion_pct'] = wave.get('completion_pct')
        elif baseline.get('completion_pct'):
            derived['completion_pct'] = baseline.get('completion_pct')
        
        if baseline_budget and actual_cost:
            derived['remaining_budget'] = baseline_budget - actual_cost
            if derived['remaining_budget'] < 0:
                derived['budget_overrun'] = True
        
        return derived
    
    def _generate_project_assessment(self, project_data: Dict) -> Dict:
        """Generate comprehensive project assessment"""
        
        metadata = project_data.get('project_metadata', {})
        baseline = project_data.get('baseline_metrics', {})
        wave = project_data.get('latest_wave_snapshot', {})
        actuals = project_data.get('actuals_summary', {})
        derived = project_data.get('derived_metrics', {})
        rules = project_data.get('rule_evaluations', [])
        trends = project_data.get('wave_trends', {})
        
        assessment = {
            'project_id': metadata.get('project_id'),
            'project_name': metadata.get('project_name'),
            'timestamp': datetime.now().isoformat()
        }
        
        status = self._classify_health(
            baseline.get('schedule_health'),
            baseline.get('budget_health'),
            baseline.get('risk_level')
        )
        
        data_completeness = 0
        if baseline: data_completeness += 1
        if wave: data_completeness += 1
        if actuals: data_completeness += 1
        
        confidence = 'High' if data_completeness == 3 else ('Medium' if data_completeness == 2 else 'Low')
        
        assessment['overall_assessment'] = {
            'status': status,
            'health': baseline.get('budget_health', 'Unknown'),
            'confidence_level': confidence,
            'data_sources_available': data_completeness,
            'summary': self._generate_summary(status, baseline, wave, actuals, derived, rules)
        }
        
        drivers = []
        
        cost_var = derived.get('cost_variance_pct')
        if cost_var:
            if cost_var < -10:
                drivers.append(f"Cost overrun: {abs(cost_var):.1f}% over baseline budget")
            elif cost_var > 10:
                drivers.append(f"Cost underrun: {cost_var:.1f}% under baseline budget")
        
        schedule_var = derived.get('schedule_variance_days')
        if schedule_var:
            if schedule_var > 0:
                drivers.append(f"Schedule delay: {schedule_var} days behind baseline")
            elif schedule_var < 0:
                drivers.append(f"Schedule ahead: {abs(schedule_var)} days ahead of baseline")
        
        if derived.get('daily_burn_rate'):
            burn = derived['daily_burn_rate']
            drivers.append(f"Daily burn rate: ${burn:,.0f}/day")
        
        assessment['key_drivers'] = drivers[:3] if drivers else ['Insufficient data for key drivers']
        
        observations = []
        
        if wave and actuals:
            observations.append(f"Wave reports {wave.get('status', 'unknown')} status")
            if actuals.get('total_cost'):
                observations.append(f"Tick shows ${actuals['total_cost']:,.0f} actual cost incurred")
        
        if trends and trends.get('recent_deterioration'):
            observations.append("Wave data shows recent status deterioration")
        
        if not actuals and baseline:
            observations.append("No execution data (Tick) found - project may not have started")
        
        assessment['cross_source_observations'] = observations if observations else ['Single source data - limited cross-validation']
        
        risks = []
        for rule in rules:
            if rule['severity'] in ['critical', 'warning']:
                risks.append(f"[{rule['severity'].upper()}] {rule['description']}")
        
        assessment['risks_warnings'] = risks if risks else ['No significant risks detected']
        
        positives = []
        
        if cost_var and cost_var >= 0:
            positives.append("Cost tracking within or under budget")
        
        if schedule_var and schedule_var <= 0:
            positives.append("Schedule on track or ahead")
        
        if baseline.get('budget_health', '').lower() == 'green':
            positives.append("Budget health marked as green")
        
        assessment['positive_signals'] = positives if positives else ['No strong positive signals identified']
        
        gaps = []
        
        if not baseline:
            gaps.append("Missing Smartsheet baseline data")
        if not wave:
            gaps.append("Missing Wave forecast/status data")
        if not actuals:
            gaps.append("Missing Tick actual execution data")
        
        if not derived.get('cost_variance_pct'):
            gaps.append("Cannot calculate cost variance - missing budget or actuals")
        
        if not derived.get('schedule_variance_days'):
            gaps.append("Cannot calculate schedule variance - missing dates")
        
        assessment['data_gaps'] = gaps if gaps else ['Complete data from all three sources']
        
        recommendations = []
        for rule in rules:
            if rule.get('recommendation'):
                recommendations.append(rule['recommendation'])
        
        assessment['recommendations'] = recommendations[:3] if recommendations else ['Continue monitoring with current data']
        
        return assessment
    
    def _generate_summary(self, status, baseline, wave, actuals, derived, rules):
        """Generate 2-3 sentence executive summary"""
        parts = []
        
        parts.append(f"Project classified as '{status}' based on cross-source analysis.")
        
        cost_var = derived.get('cost_variance_pct')
        schedule_var = derived.get('schedule_variance_days')
        
        if cost_var and abs(cost_var) > 10:
            parts.append(f"Cost variance of {cost_var:.1f}% indicates {'overrun' if cost_var < 0 else 'underrun'}.")
        elif schedule_var and abs(schedule_var) > 15:
            parts.append(f"Schedule variance of {schedule_var} days shows {'delay' if schedule_var > 0 else 'acceleration'}.")
        else:
            if actuals:
                parts.append(f"Execution tracking shows {actuals.get('total_hours', 0):.0f} hours logged to date.")
            else:
                parts.append("Limited execution data available for detailed analysis.")
        
        critical_rules = [r for r in rules if r['severity'] == 'critical']
        if critical_rules:
            parts.append(f"Critical issue: {critical_rules[0]['description']}")
        
        return " ".join(parts[:3])
    
    def analyze_project(self, project_id: str) -> Dict:
        """
        Analyze a single project across all three data sources
        Returns structured assessment with evidence-based insights
        """
        
        print(f"\n🔍 Analyzing Project: {project_id}")
        
        project_data = {
            'project_metadata': {
                'project_id': project_id,
                'analysis_timestamp': datetime.now().isoformat()
            },
            'baseline_metrics': {},
            'latest_wave_snapshot': {},
            'wave_trends': {},
            'actuals_summary': {},
            'derived_metrics': {},
            'rule_evaluations': []
        }
        
        if self.smartsheet_data is not None and self.smartsheet_cols.get('id'):
            id_col = self.smartsheet_cols['id']
            project_smartsheet = self.smartsheet_data[
                self.smartsheet_data[id_col].apply(self._normalize_project_id) == 
                self._normalize_project_id(project_id)
            ]
            
            if len(project_smartsheet) > 0:
                row = project_smartsheet.iloc[0]
                
                if self.smartsheet_cols.get('name'):
                    project_data['project_metadata']['project_name'] = self._safe_get(project_smartsheet, self.smartsheet_cols['name'])
                else:
                    project_data['project_metadata']['project_name'] = project_id
                
                baseline_metrics = {}
                
                if self.smartsheet_cols.get('start_date'):
                    baseline_metrics['baseline_start'] = str(self._safe_date(row.get(self.smartsheet_cols['start_date'])))
                
                if self.smartsheet_cols.get('finish_date'):
                    baseline_metrics['baseline_finish'] = str(self._safe_date(row.get(self.smartsheet_cols['finish_date'])))
                
                if self.smartsheet_cols.get('budget'):
                    budget_val = self._safe_numeric(row.get(self.smartsheet_cols['budget']))
                    if budget_val is not None:
                        baseline_metrics['total_budget'] = budget_val
                
                if self.smartsheet_cols.get('capex'):
                    capex_val = self._safe_numeric(row.get(self.smartsheet_cols['capex']))
                    if capex_val is not None:
                        baseline_metrics['capex'] = capex_val
                
                if self.smartsheet_cols.get('opex'):
                    opex_val = self._safe_numeric(row.get(self.smartsheet_cols['opex']))
                    if opex_val is not None:
                        baseline_metrics['opex'] = opex_val
                
                if self.smartsheet_cols.get('eac'):
                    eac_val = self._safe_numeric(row.get(self.smartsheet_cols['eac']))
                    if eac_val is not None:
                        baseline_metrics['eac'] = eac_val
                
                if self.smartsheet_cols.get('hours'):
                    hours_val = self._safe_numeric(row.get(self.smartsheet_cols['hours']))
                    if hours_val is not None:
                        baseline_metrics['planned_hours'] = hours_val
                
                if self.smartsheet_cols.get('schedule_health'):
                    baseline_metrics['schedule_health'] = self._safe_get(project_smartsheet, self.smartsheet_cols['schedule_health'])
                
                if self.smartsheet_cols.get('budget_health'):
                    baseline_metrics['budget_health'] = self._safe_get(project_smartsheet, self.smartsheet_cols['budget_health'])
                
                if self.smartsheet_cols.get('risk'):
                    baseline_metrics['risk_level'] = self._safe_get(project_smartsheet, self.smartsheet_cols['risk'])
                
                if self.smartsheet_cols.get('owner'):
                    baseline_metrics['owner'] = self._safe_get(project_smartsheet, self.smartsheet_cols['owner'])
                
                if self.smartsheet_cols.get('strategic_alignment'):
                    baseline_metrics['strategic_alignment'] = self._safe_get(project_smartsheet, self.smartsheet_cols['strategic_alignment'])
                
                if self.smartsheet_cols.get('benefits'):
                    baseline_metrics['benefits'] = self._safe_get(project_smartsheet, self.smartsheet_cols['benefits'])
                
                if self.smartsheet_cols.get('completion'):
                    completion_val = self._safe_numeric(row.get(self.smartsheet_cols['completion']))
                    if completion_val is not None:
                        baseline_metrics['completion_pct'] = completion_val
                
                if self.smartsheet_cols.get('stage'):
                    baseline_metrics['stage'] = self._safe_get(project_smartsheet, self.smartsheet_cols['stage'])
                
                if self.smartsheet_cols.get('interdependencies'):
                    baseline_metrics['interdependencies'] = self._safe_get(project_smartsheet, self.smartsheet_cols['interdependencies'])
                
                project_data['baseline_metrics'] = baseline_metrics
                print(f"  ✓ Smartsheet baseline loaded")
            else:
                print(f"  ⚠️  No Smartsheet data found")
        
        latest_wave = self._get_latest_wave_snapshot(project_id)
        if latest_wave is not None:
            wave_snapshot = {}
            
            if self.wave_cols.get('snapshot_date'):
                wave_snapshot['snapshot_date'] = str(self._safe_date(latest_wave.get(self.wave_cols['snapshot_date'])))
            
            if self.wave_cols.get('status'):
                wave_snapshot['status'] = str(latest_wave.get(self.wave_cols['status']))
            
            if self.wave_cols.get('stage'):
                wave_snapshot['stage'] = str(latest_wave.get(self.wave_cols['stage']))
            
            if self.wave_cols.get('forecast_finish'):
                wave_snapshot['forecast_finish'] = str(self._safe_date(latest_wave.get(self.wave_cols['forecast_finish'])))
            
            if self.wave_cols.get('completion'):
                completion_val = self._safe_numeric(latest_wave.get(self.wave_cols['completion']))
                if completion_val is not None:
                    wave_snapshot['completion_pct'] = completion_val
            
            if self.wave_cols.get('complexity'):
                wave_snapshot['complexity'] = str(latest_wave.get(self.wave_cols['complexity']))
            
            if self.wave_cols.get('owner'):
                wave_snapshot['owner'] = str(latest_wave.get(self.wave_cols['owner']))
            
            if self.wave_cols.get('budget'):
                budget_val = self._safe_numeric(latest_wave.get(self.wave_cols['budget']))
                if budget_val is not None:
                    wave_snapshot['budget'] = budget_val
            
            if self.wave_cols.get('value_lever'):
                wave_snapshot['value_lever'] = str(latest_wave.get(self.wave_cols['value_lever']))
            
            if self.wave_cols.get('approval_date'):
                wave_snapshot['approval_date'] = str(self._safe_date(latest_wave.get(self.wave_cols['approval_date'])))
            
            project_data['latest_wave_snapshot'] = wave_snapshot
            print(f"  ✓ Wave snapshot loaded")
        else:
            print(f"  ⚠️  No Wave data found")
        
        wave_trends = self._get_wave_trends(project_id)
        if wave_trends:
            project_data['wave_trends'] = wave_trends
            print(f"  ✓ Wave trends analyzed ({wave_trends.get('snapshot_count', 0)} snapshots)")
        
        tick_actuals = self._get_tick_actuals(project_id)
        if tick_actuals:
            project_data['actuals_summary'] = tick_actuals
            print(f"  ✓ Tick actuals loaded ({tick_actuals.get('transaction_count', 0)} transactions)")
        else:
            print(f"  ⚠️  No Tick data found")
        
        derived_metrics = self._calculate_derived_metrics(
            project_data['baseline_metrics'],
            project_data['latest_wave_snapshot'],
            project_data['actuals_summary']
        )
        project_data['derived_metrics'] = derived_metrics
        print(f"  ✓ Derived metrics calculated")
        
        rule_violations = self._evaluate_consistency_rules(project_data)
        project_data['rule_evaluations'] = rule_violations
        if rule_violations:
            print(f"  ⚠️  {len(rule_violations)} consistency rules triggered")
        
        assessment = self._generate_project_assessment(project_data)
        project_data['assessment'] = assessment
        
        print(f"✅ Analysis complete for {project_id}")
        
        self.projects[project_id] = project_data
        
        return project_data
    
    def analyze_all_projects(self) -> Dict[str, Dict]:
        """Analyze all projects in the portfolio"""
        
        print("\n" + "="*60)
        print("PORTFOLIO-WIDE ANALYSIS")
        print("="*60)
        
        project_ids = set()
        
        if self.smartsheet_data is not None and self.smartsheet_cols.get('id'):
            id_col = self.smartsheet_cols['id']
            project_ids.update(
                self.smartsheet_data[id_col].apply(self._normalize_project_id).dropna().unique()
            )
        
        if self.smartsheet_data is not None and self.smartsheet_cols.get('wave_num'):
            wave_col = self.smartsheet_cols['wave_num']
            wave_ids = self.smartsheet_data[wave_col].apply(self._normalize_project_id).dropna().unique()
            project_ids.update(wave_ids)
        
        if self.wave_data is not None and self.wave_cols.get('wave_num'):
            wave_col = self.wave_cols['wave_num']
            project_ids.update(
                self.wave_data[wave_col].apply(self._normalize_project_id).dropna().unique()
            )
        
        if self.tick_data is not None and self.tick_cols.get('wave_num'):
            wave_col = self.tick_cols['wave_num']
            tick_waves = self.tick_data[wave_col].apply(self._normalize_project_id).dropna().unique()
            tick_waves = [w for w in tick_waves if w != 'NOT SPECIFIED']
            project_ids.update(tick_waves)
        
        project_ids = sorted(list(project_ids))
        
        print(f"\n📊 Found {len(project_ids)} unique projects across all sources")
        
        for project_id in project_ids:
            try:
                self.analyze_project(project_id)
            except Exception as e:
                print(f"❌ Error analyzing {project_id}: {str(e)}")
                import traceback
                traceback.print_exc()
        
        print(f"\n✅ Portfolio analysis complete: {len(self.projects)} projects analyzed")
        
        return self.projects
    
    def get_portfolio_summary(self) -> Dict:
        """Generate portfolio-level summary and insights"""
        
        if not self.projects:
            return {'error': 'No projects analyzed yet'}
        
        summary = {
            'portfolio_overview': {
                'total_projects': len(self.projects),
                'analysis_timestamp': datetime.now().isoformat()
            },
            'status_distribution': defaultdict(int),
            'health_distribution': defaultdict(int),
            'confidence_distribution': defaultdict(int),
            'data_completeness': {
                'full_data': 0,
                'partial_data': 0,
                'minimal_data': 0
            },
            'critical_issues': [],
            'portfolio_risks': [],
            'top_concerns': []
        }
        
        total_budget = 0
        total_actuals = 0
        total_overruns = 0
        projects_with_delays = 0
        
        for project_id, project_data in self.projects.items():
            assessment = project_data.get('assessment', {})
            overall = assessment.get('overall_assessment', {})
            
            status = overall.get('status', 'Unknown')
            summary['status_distribution'][status] += 1
            
            health = overall.get('health', 'Unknown')
            summary['health_distribution'][health] += 1
            
            confidence = overall.get('confidence_level', 'Unknown')
            summary['confidence_distribution'][confidence] += 1
            
            data_sources = overall.get('data_sources_available', 0)
            if data_sources == 3:
                summary['data_completeness']['full_data'] += 1
            elif data_sources == 2:
                summary['data_completeness']['partial_data'] += 1
            else:
                summary['data_completeness']['minimal_data'] += 1
            
            baseline = project_data.get('baseline_metrics', {})
            actuals = project_data.get('actuals_summary', {})
            derived = project_data.get('derived_metrics', {})
            
            if baseline.get('total_budget'):
                total_budget += baseline['total_budget']
            if actuals.get('total_cost'):
                total_actuals += actuals['total_cost']
            
            if derived.get('budget_overrun'):
                total_overruns += 1
            
            if derived.get('schedule_variance_days', 0) > 0:
                projects_with_delays += 1
            
            rules = project_data.get('rule_evaluations', [])
            for rule in rules:
                if rule['severity'] == 'critical':
                    summary['critical_issues'].append({
                        'project_id': project_id,
                        'project_name': project_data['project_metadata'].get('project_name'),
                        'issue': rule['description'],
                        'recommendation': rule['recommendation']
                    })
        
        summary['portfolio_metrics'] = {
            'total_baseline_budget': total_budget,
            'total_actual_cost': total_actuals,
            'portfolio_variance_pct': self._calculate_variance_pct(total_actuals, total_budget) if total_budget > 0 else None,
            'projects_over_budget': total_overruns,
            'projects_delayed': projects_with_delays
        }
        
        if total_overruns > len(self.projects) * 0.3:
            summary['portfolio_risks'].append({
                'risk': 'budget_trend',
                'severity': 'high',
                'description': f"{total_overruns} of {len(self.projects)} projects showing budget overruns",
                'impact': 'Portfolio-wide cost pressure'
            })
        
        if projects_with_delays > len(self.projects) * 0.4:
            summary['portfolio_risks'].append({
                'risk': 'schedule_trend',
                'severity': 'high',
                'description': f"{projects_with_delays} of {len(self.projects)} projects experiencing delays",
                'impact': 'Portfolio delivery timeline at risk'
            })
        
        at_risk_count = summary['status_distribution'].get('At Risk', 0)
        delayed_count = summary['status_distribution'].get('Delayed', 0)
        
        if at_risk_count + delayed_count > len(self.projects) * 0.5:
            summary['top_concerns'].append(
                f"Over 50% of projects ({at_risk_count + delayed_count}/{len(self.projects)}) are At Risk or Delayed"
            )
        
        if summary['data_completeness']['minimal_data'] > len(self.projects) * 0.3:
            summary['top_concerns'].append(
                f"{summary['data_completeness']['minimal_data']} projects have incomplete data across sources"
            )
        
        if len(summary['critical_issues']) > 0:
            summary['top_concerns'].append(
                f"{len(summary['critical_issues'])} critical issues requiring immediate attention"
            )
        
        self.portfolio_insights = summary
        return summary
    
    def generate_all_insights(self):
        """Generate all formula-based insights with persona mapping"""
        print("\n" + "="*60)
        print("GENERATING FORMULA-BASED INSIGHTS")
        print("="*60)
        
        self.executive_insights = []
        self.vp_insights = []
        self.manager_insights = []
        
        # TIER-1: Board-Level Insights
        self._formula_value_leakage_index()
        self._formula_strategy_execution_coverage()
        self._formula_top_bottom_analysis()
        self._formula_delivery_confidence_forecast()
        
        # TIER-2: Portfolio & P&L Insights
        self._formula_cost_per_strategic_outcome()
        self._formula_execution_drag_index()
        self._formula_investment_map()
        self._formula_hidden_dependency_risk()
        
        # TIER-3: Operational Excellence Insights
        self._formula_effort_progress_mismatch()
        self._formula_resource_utilization_quality()
        self._formula_managerial_span_effectiveness()
        self._formula_burnout_risk_radar()
        
        # TIER-4: Execution Hygiene Insights
        self._formula_phantom_work_detection()
        self._formula_task_hygiene_score()
        self._formula_idle_capacity_hotspots()
        self._formula_execution_velocity_by_team()
        
        print(f"\n✅ Generated {len(self.executive_insights)} Executive insights")
        print(f"✅ Generated {len(self.vp_insights)} VP insights")
        print(f"✅ Generated {len(self.manager_insights)} Manager insights")
    
    # ========================================
    # TIER-1: BOARD-LEVEL INSIGHTS
    # ========================================
    
    def _formula_value_leakage_index(self):
        """
        Formula: Value Leakage % = (Effort/Cost on projects with no Wave value lever OR stalled status) / Total Effort/Cost
        Requirements: Tick effort OR Smartsheet cost + Wave value lever mapping
        """
        
        total_effort = 0
        leakage_effort = 0
        leakage_projects = []
        
        data_sources_available = set()
        
        for proj_id, proj_data in self.projects.items():
            if not self._is_valid_project_id(proj_id):
                continue
            
            actuals = proj_data.get('actuals_summary', {})
            wave = proj_data.get('latest_wave_snapshot', {})
            baseline = proj_data.get('baseline_metrics', {})
            
            effort = actuals.get('total_hours') or actuals.get('total_cost') or baseline.get('total_budget')
            
            if effort and effort > 0:
                total_effort += effort
                
                if actuals:
                    data_sources_available.add('tick')
                if wave:
                    data_sources_available.add('wave')
                if baseline:
                    data_sources_available.add('smartsheet')
                
                value_lever = wave.get('value_lever') if wave else None
                status = wave.get('status', '').lower() if wave else ''
                smartsheet_status = baseline.get('schedule_health', '').lower() if baseline else ''
                
                is_stalled = 'stalled' in status or 'red' in smartsheet_status or 'delayed' in status
                has_no_value_lever = not value_lever or str(value_lever).strip() == '' or str(value_lever).lower() == 'none'
                
                if has_no_value_lever or is_stalled:
                    leakage_effort += effort
                    leakage_projects.append({
                        'project_id': proj_id,
                        'project_name': proj_data['project_metadata'].get('project_name'),
                        'effort': effort,
                        'reason': 'No value lever' if has_no_value_lever else 'Stalled status'
                    })
        
        if total_effort > 0 and len(data_sources_available) >= 1:
            leakage_pct = (leakage_effort / total_effort) * 100
            
            confidence = 'High' if len(data_sources_available) >= 2 else 'Medium'
            
            if leakage_pct > 20:
                self.executive_insights.append({
                    'category': 'value_leakage',
                    'title': f'Value Leakage Index: {leakage_pct:.1f}% Portfolio Effort at Risk',
                    'severity': 'critical',
                    'description': f"{leakage_pct:.1f}% of portfolio effort/cost is on projects with no clear value lever or stalled status",
                    'impact': f"${leakage_effort/1000:.0f}K effort potentially not delivering value",
                    'recommendation': 'Review value mapping for all projects and address stalled initiatives',
                    'metrics': {
                        'leakage_pct': leakage_pct,
                        'leakage_effort': leakage_effort,
                        'total_effort': total_effort,
                        'leakage_project_count': len(leakage_projects),
                        'top_contributors': leakage_projects
                    },
                    'formula_used': 'Value Leakage % = (Effort on no-value/stalled projects) / Total Effort',
                    'data_sources_used': list(data_sources_available),
                    'project_id': None,
                    'confidence': confidence
                })
                self.vp_insights.append(self.executive_insights[-1].copy())
    
    def _formula_strategy_execution_coverage(self):
        """
        Formula: Coverage % = Wave initiatives with (Smartsheet tasks AND Tick hours) / Total Wave initiatives
        Requirements: Wave initiatives + Smartsheet tasks + Tick hours
        """
        
        total_wave_initiatives = 0
        covered_initiatives = 0
        uncovered_initiatives = []
        
        data_sources_available = set()
        
        for proj_id, proj_data in self.projects.items():
            if not self._is_valid_project_id(proj_id):
                continue
            
            wave = proj_data.get('latest_wave_snapshot', {})
            baseline = proj_data.get('baseline_metrics', {})
            actuals = proj_data.get('actuals_summary', {})
            
            if wave:
                total_wave_initiatives += 1
                data_sources_available.add('wave')
                
                has_smartsheet = bool(baseline)
                has_tick = bool(actuals and actuals.get('total_hours'))
                
                if has_smartsheet:
                    data_sources_available.add('smartsheet')
                if has_tick:
                    data_sources_available.add('tick')
                
                if has_smartsheet and has_tick:
                    covered_initiatives += 1
                else:
                    uncovered_initiatives.append({
                        'project_id': proj_id,
                        'project_name': proj_data['project_metadata'].get('project_name'),
                        'missing': 'Smartsheet' if not has_smartsheet else 'Tick hours'
                    })
        
        if total_wave_initiatives > 0 and 'wave' in data_sources_available:
            coverage_pct = (covered_initiatives / total_wave_initiatives) * 100
            
            confidence = 'High' if len(data_sources_available) == 3 else 'Low'
            
            if coverage_pct < 60:
                self.executive_insights.append({
                    'category': 'strategic_alignment',
                    'title': f'Strategy-Execution Coverage: Only {coverage_pct:.1f}% Fully Linked',
                    'severity': 'critical',
                    'description': f"Only {covered_initiatives} of {total_wave_initiatives} Wave initiatives have both Smartsheet baseline and Tick execution data",
                    'impact': 'Strategic initiatives not fully traceable from plan to execution',
                    'recommendation': 'Establish full traceability for all Wave initiatives through Smartsheet and Tick',
                    'metrics': {
                        'coverage_pct': coverage_pct,
                        'covered_count': covered_initiatives,
                        'total_count': total_wave_initiatives,
                        'uncovered_initiatives': uncovered_initiatives
                    },
                    'formula_used': 'Coverage % = Initiatives with (Smartsheet AND Tick) / Total Wave initiatives',
                    'data_sources_used': list(data_sources_available),
                    'project_id': None,
                    'confidence': confidence
                })
            elif coverage_pct < 80:
                self.vp_insights.append({
                    'category': 'strategic_alignment',
                    'title': f'Strategy-Execution Coverage: {coverage_pct:.1f}% Linked',
                    'severity': 'warning',
                    'description': f"{covered_initiatives} of {total_wave_initiatives} Wave initiatives have full traceability",
                    'impact': 'Some strategic initiatives lack complete tracking',
                    'recommendation': 'Improve data linkage for remaining initiatives',
                    'metrics': {
                        'coverage_pct': coverage_pct,
                        'covered_count': covered_initiatives,
                        'total_count': total_wave_initiatives,
                        'uncovered_initiatives': uncovered_initiatives
                    },
                    'formula_used': 'Coverage % = Initiatives with (Smartsheet AND Tick) / Total Wave initiatives',
                    'data_sources_used': list(data_sources_available),
                    'project_id': None,
                    'confidence': confidence
                })
    
    def _formula_top_bottom_analysis(self):
        """
        Formula: Rank projects by Tick effort (top 10%) vs Smartsheet progress/Wave value (bottom 10%)
        Intersection = flagged projects
        """
        
        projects_with_data = []
        
        for proj_id, proj_data in self.projects.items():
            if not self._is_valid_project_id(proj_id):
                continue
            
            actuals = proj_data.get('actuals_summary', {})
            baseline = proj_data.get('baseline_metrics', {})
            wave = proj_data.get('latest_wave_snapshot', {})
            derived = proj_data.get('derived_metrics', {})
            
            effort = actuals.get('total_hours') or actuals.get('total_cost')
            progress = derived.get('completion_pct') or baseline.get('completion_pct')
            value_lever = wave.get('value_lever') if wave else None
            
            if effort and effort > 0:
                projects_with_data.append({
                    'project_id': proj_id,
                    'project_name': proj_data['project_metadata'].get('project_name'),
                    'effort': effort,
                    'progress': progress if progress else 0,
                    'has_value_lever': bool(value_lever and str(value_lever).strip() and str(value_lever).lower() != 'none')
                })
        
        if len(projects_with_data) >= 10:
            sorted_by_effort = sorted(projects_with_data, key=lambda x: x['effort'], reverse=True)
            sorted_by_progress = sorted(projects_with_data, key=lambda x: x['progress'])
            
            top_10_pct_count = max(1, len(projects_with_data) // 10)
            
            top_effort = sorted_by_effort[:top_10_pct_count]
            bottom_progress = sorted_by_progress[:top_10_pct_count]
            
            flagged_projects = []
            for te in top_effort:
                for bp in bottom_progress:
                    if te['project_id'] == bp['project_id']:
                        flagged_projects.append(te)
                        break
            
            if flagged_projects:
                self.executive_insights.append({
                    'category': 'value_leakage',
                    'title': f'Top 10% Effort / Bottom 10% Outcome: {len(flagged_projects)} Projects Flagged',
                    'severity': 'critical',
                    'description': f"{len(flagged_projects)} projects consuming highest effort but delivering lowest progress/value",
                    'impact': 'Significant resource investment with minimal return',
                    'recommendation': 'Immediate review for scope reduction, reprioritization, or termination',
                    'metrics': {
                        'flagged_count': len(flagged_projects),
                        'flagged_projects': flagged_projects
                    },
                    'formula_used': 'Top 10% effort INTERSECT Bottom 10% progress',
                    'data_sources_used': ['tick', 'smartsheet', 'wave'],
                    'project_id': None,
                    'confidence': 'High'
                })
                self.vp_insights.append(self.executive_insights[-1].copy())
    
    def _formula_delivery_confidence_forecast(self):
        """
        Formula: Risk Score = Low velocity + High burn rate + Task slippage trend
        Binary: Likely to Miss / On Track
        Requirement: >=2 sources
        """
        
        at_risk_projects = []
        
        for proj_id, proj_data in self.projects.items():
            if not self._is_valid_project_id(proj_id):
                continue
            
            actuals = proj_data.get('actuals_summary', {})
            derived = proj_data.get('derived_metrics', {})
            trends = proj_data.get('wave_trends', {})
            
            data_sources = 0
            risk_signals = []
            
            completion = derived.get('completion_pct')
            work_days = actuals.get('work_span_days')
            
            if completion and work_days and work_days > 30:
                data_sources += 1
                velocity = completion / work_days
                if velocity < 0.5:
                    risk_signals.append('Low velocity')
            
            burn_rate = derived.get('daily_burn_rate')
            remaining_budget = derived.get('remaining_budget')
            
            if burn_rate and burn_rate > 0 and remaining_budget:
                data_sources += 1
                days_remaining = remaining_budget / burn_rate if burn_rate > 0 else 999
                if days_remaining < 30:
                    risk_signals.append('High burn rate')
            
            if trends and trends.get('recent_deterioration'):
                data_sources += 1
                risk_signals.append('Task slippage trend')
            
            if data_sources >= 2 and len(risk_signals) >= 2:
                at_risk_projects.append({
                    'project_id': proj_id,
                    'project_name': proj_data['project_metadata'].get('project_name'),
                    'risk_signals': risk_signals,
                    'forecast': 'Likely to Miss'
                })
        
        if at_risk_projects:
            self.executive_insights.append({
                'category': 'predictive_risk',
                'title': f'Delivery Confidence Forecast: {len(at_risk_projects)} Projects Likely to Miss',
                'severity': 'critical',
                'description': f"{len(at_risk_projects)} projects show multiple risk signals indicating delivery failure",
                'impact': 'Portfolio delivery commitments at risk',
                'recommendation': 'Implement recovery plans or adjust expectations immediately',
                'metrics': {
                    'at_risk_count': len(at_risk_projects),
                    'at_risk_projects': at_risk_projects
                },
                'formula_used': 'Risk Score = Low velocity + High burn + Slippage (binary: Likely to Miss)',
                'data_sources_used': ['tick', 'smartsheet', 'wave'],
                'project_id': None,
                'confidence': 'High'
            })
            self.vp_insights.append(self.executive_insights[-1].copy())
    
    # ========================================
    # TIER-2: PORTFOLIO & P&L INSIGHTS
    # ========================================
    
    def _formula_cost_per_strategic_outcome(self):
        """
        Formula: Cost per Value Lever = Tick effort cost / Wave value delivered
        """
        
        value_lever_costs = defaultdict(lambda: {'cost': 0, 'projects': []})
        
        for proj_id, proj_data in self.projects.items():
            if not self._is_valid_project_id(proj_id):
                continue
            
            actuals = proj_data.get('actuals_summary', {})
            wave = proj_data.get('latest_wave_snapshot', {})
            
            cost = actuals.get('total_cost')
            value_lever = wave.get('value_lever') if wave else None
            
            if cost and cost > 0 and value_lever and str(value_lever).strip():
                value_lever_costs[value_lever]['cost'] += cost
                value_lever_costs[value_lever]['projects'].append({
                    'project_id': proj_id,
                    'project_name': proj_data['project_metadata'].get('project_name'),
                    'cost': cost
                })
        
        if value_lever_costs:
            sorted_levers = sorted(value_lever_costs.items(), key=lambda x: x[1]['cost'], reverse=True)
            
            self.vp_insights.append({
                'category': 'value_leakage',
                'title': f'Cost per Strategic Outcome: {len(value_lever_costs)} Value Levers Analyzed',
                'severity': 'info',
                'description': f"Investment distribution across {len(value_lever_costs)} strategic value levers",
                'impact': 'Enables value-based portfolio optimization',
                'recommendation': 'Rebalance investment toward highest-value levers',
                'metrics': {
                    'value_lever_count': len(value_lever_costs),
                    'top_investments': [{'lever': k, 'cost': v['cost'], 'project_count': len(v['projects'])} for k, v in sorted_levers]
                },
                'formula_used': 'Cost per Value Lever = Total Tick cost / Value Lever',
                'data_sources_used': ['tick', 'wave'],
                'project_id': None,
                'confidence': 'High'
            })
    
    def _formula_execution_drag_index(self):
        """
        Formula: Drag Days = (First Smartsheet task start) - (Wave approval date)
        """
        
        drag_projects = []
        
        for proj_id, proj_data in self.projects.items():
            if not self._is_valid_project_id(proj_id):
                continue
            
            baseline = proj_data.get('baseline_metrics', {})
            wave = proj_data.get('latest_wave_snapshot', {})
            
            task_start = self._safe_date(baseline.get('baseline_start'))
            approval_date = self._safe_date(wave.get('approval_date'))
            
            if task_start and approval_date and task_start > approval_date:
                drag_days = (task_start - approval_date).days
                
                if drag_days > 30:
                    drag_projects.append({
                        'project_id': proj_id,
                        'project_name': proj_data['project_metadata'].get('project_name'),
                        'drag_days': drag_days,
                        'approval_date': str(approval_date.date()),
                        'start_date': str(task_start.date())
                    })
        
        if drag_projects:
            avg_drag = sum(p['drag_days'] for p in drag_projects) / len(drag_projects)
            
            self.vp_insights.append({
                'category': 'velocity',
                'title': f'Execution Drag Index: {avg_drag:.0f} Days Average Delay',
                'severity': 'warning',
                'description': f"{len(drag_projects)} projects show significant delay between approval and execution start",
                'impact': f"Average {avg_drag:.0f} days lost between approval and start",
                'recommendation': 'Streamline project kickoff and resource allocation processes',
                'metrics': {
                    'avg_drag_days': avg_drag,
                    'affected_projects': len(drag_projects),
                    'worst_offenders': sorted(drag_projects, key=lambda x: x['drag_days'], reverse=True)
                },
                'formula_used': 'Drag Days = Task Start Date - Approval Date',
                'data_sources_used': ['smartsheet', 'wave'],
                'project_id': None,
                'confidence': 'High'
            })
            self.manager_insights.append(self.vp_insights[-1].copy())
    
    def _formula_investment_map(self):
        """
        Logic: Over-invested (high effort + low value), Under-invested (low effort + high value)
        """
        
        over_invested = []
        under_invested = []
        
        projects_with_data = []
        
        for proj_id, proj_data in self.projects.items():
            if not self._is_valid_project_id(proj_id):
                continue
            
            actuals = proj_data.get('actuals_summary', {})
            wave = proj_data.get('latest_wave_snapshot', {})
            derived = proj_data.get('derived_metrics', {})
            
            effort = actuals.get('total_hours') or actuals.get('total_cost')
            value_lever = wave.get('value_lever') if wave else None
            progress = derived.get('completion_pct')
            
            if effort and effort > 0:
                has_value = bool(value_lever and str(value_lever).strip() and str(value_lever).lower() != 'none')
                
                projects_with_data.append({
                    'project_id': proj_id,
                    'project_name': proj_data['project_metadata'].get('project_name'),
                    'effort': effort,
                    'has_value': has_value,
                    'progress': progress if progress else 0
                })
        
        if len(projects_with_data) >= 4:
            median_effort = sorted([p['effort'] for p in projects_with_data])[len(projects_with_data)//2]
            
            for proj in projects_with_data:
                if proj['effort'] > median_effort and not proj['has_value']:
                    over_invested.append(proj)
                elif proj['effort'] < median_effort and proj['has_value']:
                    under_invested.append(proj)
            
            if over_invested:
                self.vp_insights.append({
                    'category': 'prioritization',
                    'title': f'Over-Investment Alert: {len(over_invested)} Projects',
                    'severity': 'warning',
                    'description': f"{len(over_invested)} projects consuming above-median effort with no clear value lever",
                    'impact': 'Inefficient resource allocation',
                    'recommendation': 'Review value proposition or reduce investment',
                    'metrics': {
                        'over_invested_count': len(over_invested),
                        'projects': over_invested
                    },
                    'formula_used': 'Over-invested = High effort + No value lever',
                    'data_sources_used': ['tick', 'wave'],
                    'project_id': None,
                    'confidence': 'High'
                })
            
            if under_invested:
                self.vp_insights.append({
                    'category': 'prioritization',
                    'title': f'Under-Investment Opportunity: {len(under_invested)} Projects',
                    'severity': 'info',
                    'description': f"{len(under_invested)} high-value projects receiving below-median investment",
                    'impact': 'Potential for accelerated value delivery',
                    'recommendation': '⚡ ACCELERATE: Consider increasing investment',
                    'metrics': {
                        'under_invested_count': len(under_invested),
                        'projects': under_invested
                    },
                    'formula_used': 'Under-invested = Low effort + Has value lever',
                    'data_sources_used': ['tick', 'wave'],
                    'project_id': None,
                    'confidence': 'High'
                })
    
    def _formula_hidden_dependency_risk(self):
        """
        Rule: Smartsheet status = Green AND (Tick effort = 0 OR upstream dependency stalled)
        """
        
        at_risk_projects = []
        
        for proj_id, proj_data in self.projects.items():
            if not self._is_valid_project_id(proj_id):
                continue
            
            baseline = proj_data.get('baseline_metrics', {})
            actuals = proj_data.get('actuals_summary', {})
            
            status = baseline.get('schedule_health', '').lower()
            has_effort = bool(actuals and actuals.get('total_hours'))
            has_dependencies = baseline.get('interdependencies')
            
            if 'green' in status and (not has_effort or has_dependencies):
                at_risk_projects.append({
                    'project_id': proj_id,
                    'project_name': proj_data['project_metadata'].get('project_name'),
                    'reason': 'No execution despite green status' if not has_effort else 'Has dependencies',
                    'dependencies': has_dependencies if has_dependencies else 'None listed'
                })
        
        if at_risk_projects:
            self.vp_insights.append({
                'category': 'execution_health',
                'title': f'Hidden Dependency Risk: {len(at_risk_projects)} Projects',
                'severity': 'warning',
                'description': f"{len(at_risk_projects)} projects report green status but show no execution or have dependencies",
                'impact': 'False confidence in project health',
                'recommendation': 'Validate dependency status and execution progress',
                'metrics': {
                    'at_risk_count': len(at_risk_projects),
                    'projects': at_risk_projects
                },
                'formula_used': 'Green status AND (No Tick effort OR Has dependencies)',
                'data_sources_used': ['smartsheet', 'tick'],
                'project_id': None,
                'confidence': 'High'
            })
            self.manager_insights.append(self.vp_insights[-1].copy())
    
    # ========================================
    # TIER-3: OPERATIONAL EXCELLENCE INSIGHTS
    # ========================================
    
    def _formula_effort_progress_mismatch(self):
        """
        Rule: Increasing Tick hours + Flat Smartsheet % complete
        """
        
        mismatch_projects = []
        
        for proj_id, proj_data in self.projects.items():
            if not self._is_valid_project_id(proj_id):
                continue
            
            actuals = proj_data.get('actuals_summary', {})
            derived = proj_data.get('derived_metrics', {})
            baseline = proj_data.get('baseline_metrics', {})
            
            actual_hours = actuals.get('total_hours')
            planned_hours = baseline.get('planned_hours')
            completion = derived.get('completion_pct')
            
            if actual_hours and planned_hours and completion:
                implied_completion = (actual_hours / planned_hours) * 100
                
                if actual_hours > planned_hours * 0.5 and completion < 40:
                    mismatch_projects.append({
                        'project_id': proj_id,
                        'project_name': proj_data['project_metadata'].get('project_name'),
                        'actual_hours': actual_hours,
                        'planned_hours': planned_hours,
                        'reported_completion': completion,
                        'implied_completion': implied_completion,
                        'gap': implied_completion - completion
                    })
        
        if mismatch_projects:
            self.manager_insights.append({
                'category': 'data_hygiene',
                'title': f'Effort-Progress Mismatch: {len(mismatch_projects)} Projects',
                'severity': 'warning',
                'description': f"{len(mismatch_projects)} projects show high effort but low reported progress",
                'impact': 'Potential productivity issues or data quality problems',
                'recommendation': 'Reconcile completion % with actual work performed',
                'metrics': {
                    'mismatch_count': len(mismatch_projects),
                    'projects': mismatch_projects
                },
                'formula_used': 'Actual hours > 50% planned AND Completion < 40%',
                'data_sources_used': ['tick', 'smartsheet'],
                'project_id': None,
                'confidence': 'High'
            })
    
    def _formula_resource_utilization_quality(self):
        """
        Formula: Strategic Utilization = Effort on Wave-linked work / Total effort
        """
        
        total_effort = 0
        strategic_effort = 0
        
        for proj_id, proj_data in self.projects.items():
            if not self._is_valid_project_id(proj_id):
                continue
            
            actuals = proj_data.get('actuals_summary', {})
            wave = proj_data.get('latest_wave_snapshot', {})
            
            effort = actuals.get('total_hours')
            
            if effort and effort > 0:
                total_effort += effort
                
                if wave:
                    strategic_effort += effort
        
        if total_effort > 0:
            strategic_util_pct = (strategic_effort / total_effort) * 100
            
            if strategic_util_pct < 70:
                self.vp_insights.append({
                    'category': 'resource_utilization',
                    'title': f'Resource Utilization Quality: {strategic_util_pct:.1f}% Strategic',
                    'severity': 'warning',
                    'description': f"Only {strategic_util_pct:.1f}% of total effort is linked to Wave strategic initiatives",
                    'impact': 'Significant effort on non-strategic work',
                    'recommendation': 'Review unlinked effort and validate strategic alignment',
                    'metrics': {
                        'strategic_util_pct': strategic_util_pct,
                        'strategic_effort': strategic_effort,
                        'total_effort': total_effort,
                        'non_strategic_effort': total_effort - strategic_effort
                    },
                    'formula_used': 'Strategic Utilization = Wave-linked effort / Total effort',
                    'data_sources_used': ['tick', 'wave'],
                    'project_id': None,
                    'confidence': 'High'
                })
    
    def _formula_managerial_span_effectiveness(self):
        """
        Rule: High # initiatives per manager + Correlated with slippage/inefficiency
        """
        
        manager_loads = defaultdict(lambda: {'projects': [], 'total_delay': 0, 'over_budget_count': 0})
        
        for proj_id, proj_data in self.projects.items():
            if not self._is_valid_project_id(proj_id):
                continue
            
            baseline = proj_data.get('baseline_metrics', {})
            derived = proj_data.get('derived_metrics', {})
            
            owner = baseline.get('owner')
            schedule_variance = derived.get('schedule_variance_days', 0)
            budget_overrun = derived.get('budget_overrun', False)
            
            if owner and str(owner).strip():
                manager_loads[owner]['projects'].append(proj_id)
                if schedule_variance > 0:
                    manager_loads[owner]['total_delay'] += schedule_variance
                if budget_overrun:
                    manager_loads[owner]['over_budget_count'] += 1
        
        overloaded_managers = []
        
        for manager, data in manager_loads.items():
            project_count = len(data['projects'])
            
            if project_count >= 5:
                avg_delay = data['total_delay'] / project_count if project_count > 0 else 0
                
                if avg_delay > 30 or data['over_budget_count'] > project_count * 0.5:
                    overloaded_managers.append({
                        'manager': manager,
                        'project_count': project_count,
                        'avg_delay_days': avg_delay,
                        'over_budget_count': data['over_budget_count']
                    })
        
        if overloaded_managers:
            self.vp_insights.append({
                'category': 'resource_utilization',
                'title': f'Managerial Span Effectiveness: {len(overloaded_managers)} Overloaded Managers',
                'severity': 'warning',
                'description': f"{len(overloaded_managers)} managers with high project counts show correlation with delays and overruns",
                'impact': 'Managerial overload driving poor project outcomes',
                'recommendation': 'Rebalance project assignments and consider additional management support',
                'metrics': {
                    'overloaded_count': len(overloaded_managers),
                    'managers': overloaded_managers
                },
                'formula_used': 'High project count + Correlated delays/overruns',
                'data_sources_used': ['smartsheet'],
                'project_id': None,
                'confidence': 'High'
            })
    
    def _formula_burnout_risk_radar(self):
        """
        Rule: Sustained high effort + Low delivery movement
        """
        
        burnout_risk_projects = []
        
        for proj_id, proj_data in self.projects.items():
            if not self._is_valid_project_id(proj_id):
                continue
            
            actuals = proj_data.get('actuals_summary', {})
            derived = proj_data.get('derived_metrics', {})
            
            total_hours = actuals.get('total_hours')
            work_span = actuals.get('work_span_days')
            completion = derived.get('completion_pct')
            unique_resources = actuals.get('unique_resources')
            
            if total_hours and work_span and work_span > 60 and unique_resources:
                avg_hours_per_resource = total_hours / unique_resources
                
                if avg_hours_per_resource > 200 and completion and completion < 50:
                    burnout_risk_projects.append({
                        'project_id': proj_id,
                        'project_name': proj_data['project_metadata'].get('project_name'),
                        'total_hours': total_hours,
                        'unique_resources': unique_resources,
                        'avg_hours_per_resource': avg_hours_per_resource,
                        'completion': completion,
                        'work_span_days': work_span
                    })
        
        if burnout_risk_projects:
            self.vp_insights.append({
                'category': 'resource_utilization',
                'title': f'Burnout Risk Radar: {len(burnout_risk_projects)} Projects at Risk',
                'severity': 'critical',
                'description': f"{len(burnout_risk_projects)} projects show sustained high effort with low progress",
                'impact': 'Team burnout and attrition risk',
                'recommendation': 'Review team health, scope, and consider resource rotation',
                'metrics': {
                    'at_risk_count': len(burnout_risk_projects),
                    'projects': burnout_risk_projects
                },
                'formula_used': 'Sustained effort (>200 hrs/person) + Low progress (<50%)',
                'data_sources_used': ['tick', 'smartsheet'],
                'project_id': None,
                'confidence': 'High'
            })
            self.manager_insights.append(self.vp_insights[-1].copy())
    
    # ========================================
    # TIER-4: EXECUTION HYGIENE INSIGHTS
    # ========================================
    
    def _formula_phantom_work_detection(self):
        """
        Rule: Tick hours with no Smartsheet task AND no Wave mapping
        """
        
        if self.tick_data is None:
            return
        
        phantom_hours = 0
        phantom_projects = []
        
        for proj_id, proj_data in self.projects.items():
            if not self._is_valid_project_id(proj_id):
                continue
            
            actuals = proj_data.get('actuals_summary', {})
            baseline = proj_data.get('baseline_metrics', {})
            wave = proj_data.get('latest_wave_snapshot', {})
            
            has_tick = bool(actuals and actuals.get('total_hours'))
            has_smartsheet = bool(baseline)
            has_wave = bool(wave)
            
            if has_tick and not has_smartsheet and not has_wave:
                phantom_hours += actuals.get('total_hours', 0)
                phantom_projects.append({
                    'project_id': proj_id,
                    'project_name': proj_data['project_metadata'].get('project_name'),
                    'hours': actuals.get('total_hours', 0)
                })
        
        if phantom_hours > 0:
            self.manager_insights.append({
                'category': 'data_hygiene',
                'title': f'Phantom Work Detection: {phantom_hours:.0f} Hours Unaccounted',
                'severity': 'warning',
                'description': f"{phantom_hours:.0f} hours logged in Tick with no Smartsheet task or Wave mapping",
                'impact': 'Work being performed outside of approved project scope',
                'recommendation': 'Investigate unlinked work and ensure proper project setup',
                'metrics': {
                    'phantom_hours': phantom_hours,
                    'phantom_project_count': len(phantom_projects),
                    'projects': phantom_projects
                },
                'formula_used': 'Tick hours AND (No Smartsheet task AND No Wave mapping)',
                'data_sources_used': ['tick'],
                'project_id': None,
                'confidence': 'High'
            })
    
    def _formula_task_hygiene_score(self):
        """
        Formula: Hygiene % = Tasks with (owner + dates + effort) / Total tasks
        """
        
        if self.smartsheet_data is None:
            return
        
        total_tasks = len(self.smartsheet_data)
        complete_tasks = 0
        
        for idx, row in self.smartsheet_data.iterrows():
            has_owner = False
            has_dates = False
            has_effort = False
            
            if self.smartsheet_cols.get('owner'):
                owner_val = row.get(self.smartsheet_cols['owner'])
                has_owner = bool(owner_val and str(owner_val).strip())
            
            if self.smartsheet_cols.get('start_date') and self.smartsheet_cols.get('finish_date'):
                start_val = self._safe_date(row.get(self.smartsheet_cols['start_date']))
                end_val = self._safe_date(row.get(self.smartsheet_cols['finish_date']))
                has_dates = bool(start_val and end_val)
            
            if self.smartsheet_cols.get('hours'):
                effort_val = self._safe_numeric(row.get(self.smartsheet_cols['hours']))
                has_effort = bool(effort_val and effort_val > 0)
            
            if has_owner and has_dates and has_effort:
                complete_tasks += 1
        
        if total_tasks > 0:
            hygiene_pct = (complete_tasks / total_tasks) * 100
            
            if hygiene_pct < 70:
                self.manager_insights.append({
                    'category': 'data_hygiene',
                    'title': f'Task Hygiene Score: {hygiene_pct:.1f}% Complete',
                    'severity': 'warning',
                    'description': f"Only {hygiene_pct:.1f}% of Smartsheet tasks have owner, dates, and effort defined",
                    'impact': 'Incomplete task definition impairs planning and tracking',
                    'recommendation': 'Enforce task completeness standards in Smartsheet',
                    'metrics': {
                        'hygiene_pct': hygiene_pct,
                        'complete_tasks': complete_tasks,
                        'total_tasks': total_tasks,
                        'incomplete_tasks': total_tasks - complete_tasks
                    },
                    'formula_used': 'Hygiene % = Tasks with (owner + dates + effort) / Total tasks',
                    'data_sources_used': ['smartsheet'],
                    'project_id': None,
                    'confidence': 'High'
                })
    
    def _formula_idle_capacity_hotspots(self):
        """
        Rule: Resources with availability + No strategic assignment
        """
        
        if self.tick_data is None:
            return
        
        resource_col = self.tick_cols.get('resource')
        if not resource_col:
            return
        
        resource_hours = defaultdict(lambda: {'total_hours': 0, 'projects': set()})
        
        for idx, row in self.tick_data.iterrows():
            resource = row.get(resource_col)
            hours = self._safe_numeric(row.get(self.tick_cols.get('actual_hours') or self.tick_cols.get('hours')))
            
            proj_id_col = self.tick_cols.get('wave_num') or self.tick_cols.get('id')
            proj_id = row.get(proj_id_col) if proj_id_col else None
            
            if resource and hours:
                resource_hours[resource]['total_hours'] += hours
                if proj_id:
                    resource_hours[resource]['projects'].add(self._normalize_project_id(proj_id))
        
        low_utilization_resources = []
        
        for resource, data in resource_hours.items():
            if data['total_hours'] < 100:
                low_utilization_resources.append({
                    'resource': resource,
                    'total_hours': data['total_hours'],
                    'project_count': len(data['projects'])
                })
        
        if len(low_utilization_resources) > 5:
            total_idle_hours = sum(r['total_hours'] for r in low_utilization_resources)
            
            self.manager_insights.append({
                'category': 'resource_utilization',
                'title': f'Idle Capacity Hotspots: {len(low_utilization_resources)} Under-Utilized Resources',
                'severity': 'info',
                'description': f"{len(low_utilization_resources)} resources show low utilization with ~{total_idle_hours:.0f} total hours",
                'impact': 'Potential capacity for strategic initiatives',
                'recommendation': 'Review availability and consider strategic assignments',
                'metrics': {
                    'low_util_count': len(low_utilization_resources),
                    'total_idle_hours': total_idle_hours,
                    'resources': low_utilization_resources
                },
                'formula_used': 'Resources with <100 hours logged',
                'data_sources_used': ['tick'],
                'project_id': None,
                'confidence': 'Medium'
            })
    
    def _formula_execution_velocity_by_team(self):
        """
        Formula: Velocity = Tasks completed / Effort hours
        """
        
        team_velocities = defaultdict(lambda: {'completed': 0, 'hours': 0, 'projects': []})
        
        for proj_id, proj_data in self.projects.items():
            if not self._is_valid_project_id(proj_id):
                continue
            
            baseline = proj_data.get('baseline_metrics', {})
            actuals = proj_data.get('actuals_summary', {})
            derived = proj_data.get('derived_metrics', {})
            
            owner = baseline.get('owner')
            completion = derived.get('completion_pct')
            hours = actuals.get('total_hours')
            
            if owner and completion and hours and hours > 0:
                team_velocities[owner]['completed'] += completion
                team_velocities[owner]['hours'] += hours
                team_velocities[owner]['projects'].append(proj_id)
        
        velocity_analysis = []
        
        for team, data in team_velocities.items():
            if data['hours'] > 0:
                velocity = data['completed'] / data['hours']
                velocity_analysis.append({
                    'team': team,
                    'velocity': velocity,
                    'total_completion': data['completed'],
                    'total_hours': data['hours'],
                    'project_count': len(data['projects'])
                })
        
        if velocity_analysis:
            sorted_velocity = sorted(velocity_analysis, key=lambda x: x['velocity'])
            
            if len(sorted_velocity) >= 3:
                lowest_performers = sorted_velocity[:3]
                
                self.manager_insights.append({
                    'category': 'velocity',
                    'title': f'Execution Velocity by Team: {len(velocity_analysis)} Teams Analyzed',
                    'severity': 'info',
                    'description': f"Velocity analysis across {len(velocity_analysis)} teams/owners",
                    'impact': 'Identifies high and low performing teams',
                    'recommendation': 'Share best practices from high-velocity teams with low performers',
                    'metrics': {
                        'team_count': len(velocity_analysis),
                        'lowest_velocity': lowest_performers,
                        'all_teams': velocity_analysis
                    },
                    'formula_used': 'Velocity = Completion % / Effort hours',
                    'data_sources_used': ['tick', 'smartsheet'],
                    'project_id': None,
                    'confidence': 'High'
                })
    
    def get_executive_insights(self) -> List[Dict]:
        """Get insights for C-Level executives"""
        return sorted(self.executive_insights, key=lambda x: {'critical': 0, 'high': 1, 'warning': 2, 'info': 3}.get(x['severity'], 3))
    
    def get_vp_insights(self) -> List[Dict]:
        """Get insights for VP / Portfolio Owners"""
        return sorted(self.vp_insights, key=lambda x: {'critical': 0, 'high': 1, 'warning': 2, 'info': 3}.get(x['severity'], 3))
    
    def get_manager_insights(self) -> List[Dict]:
        """Get insights for Managers / Delivery Leads"""
        return sorted(self.manager_insights, key=lambda x: {'critical': 0, 'high': 1, 'warning': 2, 'info': 3}.get(x['severity'], 3))
    
    def get_project_executive_insights(self, project_id: str) -> List[Dict]:
        """Get executive insights for a specific project"""
        if not self._is_valid_project_id(project_id):
            return []
        filtered = [i for i in self.executive_insights if i.get('project_id') == project_id]
        return sorted(filtered, key=lambda x: {'critical': 0, 'high': 1, 'warning': 2, 'info': 3}.get(x['severity'], 3))
    
    def get_project_vp_insights(self, project_id: str) -> List[Dict]:
        """Get VP insights for a specific project"""
        if not self._is_valid_project_id(project_id):
            return []
        filtered = [i for i in self.vp_insights if i.get('project_id') == project_id]
        return sorted(filtered, key=lambda x: {'critical': 0, 'high': 1, 'warning': 2, 'info': 3}.get(x['severity'], 3))
    
    def get_project_manager_insights(self, project_id: str) -> List[Dict]:
        """Get manager insights for a specific project"""
        if not self._is_valid_project_id(project_id):
            return []
        filtered = [i for i in self.manager_insights if i.get('project_id') == project_id]
        return sorted(filtered, key=lambda x: {'critical': 0, 'high': 1, 'warning': 2, 'info': 3}.get(x['severity'], 3))
    
    def export_project_analysis(self, project_id: str, filepath: str):
        """Export single project analysis to JSON"""
        if project_id not in self.projects:
            print(f"❌ Project {project_id} not found")
            return
        
        with open(filepath, 'w') as f:
            json.dump(self.projects[project_id], f, indent=2, default=str)
        print(f"✅ Exported {project_id} analysis to {filepath}")
    
    def export_portfolio_analysis(self, filepath: str):
        """Export full portfolio analysis to JSON"""
        portfolio_data = {
            'portfolio_summary': self.get_portfolio_summary(),
            'projects': self.projects,
            'insights': {
                'executive': self.get_executive_insights(),
                'vp': self.get_vp_insights(),
                'manager': self.get_manager_insights()
            }
        }
        
        with open(filepath, 'w') as f:
            json.dump(portfolio_data, f, indent=2, default=str)
        print(f"✅ Exported portfolio analysis to {filepath}")


if __name__ == "__main__":
    
    engine = PortfolioAIEngine()
    
    smartsheet_df = pd.read_excel("smartsheet_export.xlsx")
    wave_df = pd.read_excel("wave_export.xlsx")
    tick_df = pd.read_excel("tick_export.xlsx")
    
    engine.load_smartsheet(smartsheet_df)
    engine.load_wave(wave_df)
    engine.load_tick(tick_df)
    
    projects = engine.analyze_all_projects()
    
    engine.generate_all_insights()
    
    exec_insights = engine.get_executive_insights()
    vp_insights = engine.get_vp_insights()
    manager_insights = engine.get_manager_insights()
    
    summary = engine.get_portfolio_summary()
    
    engine.export_portfolio_analysis("portfolio_analysis.json")
    
    print("\n" + "="*60)
    print("INSIGHTS SUMMARY")
    print("="*60)
    print(f"Executive Insights: {len(exec_insights)}")
    print(f"VP Insights: {len(vp_insights)}")
    print(f"Manager Insights: {len(manager_insights)}")