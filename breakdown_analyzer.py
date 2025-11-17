# breakdown_analyzer.py

import pandas as pd
from datetime import datetime
from typing import Dict, List, Any
import logging

# Assumendo che il tuo logger sia configurato in un file logger_config.py
from logger_config import setup_logger
logger = setup_logger('BreakdownAnalyzer')
logger = logging.getLogger('BreakdownAnalyzer')


class BreakdownAnalyzer:
    """
    Analyzes production line stoppages (breakdowns) using statistical and AI methods.
    """

    def __init__(self, ai_analyzer: Any):
        """
        Initializes the Breakdown Analyzer.

        Args:
            ai_analyzer: An instance of AIAnalyzer for AI-driven insights.
        """
        self.ai_analyzer = ai_analyzer
        logger.info("BreakdownAnalyzer initialized.")

    def get_breakdown_data(self, db_connection: Any, start_date: str, end_date: str) -> List[Dict]:
        """
        Retrieves line stoppage data from the database using the specified query.

        Args:
            db_connection: An active database connection object.
            start_date: The start date for the analysis ('YYYY-MM-DD').
            end_date: The end date for the analysis ('YYYY-MM-DD').

        Returns:
            A list of dictionaries, where each dictionary represents a stoppage event.
        """
        query = """
               SELECT  r.BreakDownProblemLogId, r.DateReport,
                      r.HourReport,
                      r.UserName,
                      ia.IssueArea,
                      wa.AreaName,
                      r.WorkingEquipmentsID,
                      wl.WorkingLineName,
                      ws.AreaSubName,
                      i.DescriptionRO,
                      r.FromHour,
                      r.ToHour,
                      r.Lost_OR_Gain,
                      r.Hours,
                      r.PoNumber,
                      r.ProductCode,
                      r.IssueProblemsPerLineId,
                      r.Note,
                      r.ActionPlan,
                      r.PlannedTime
               FROM [ResetServices].[BreakDown].[ReportIssueLogs] R
                   INNER JOIN [ResetServices].[BreakDown].IssuesAreas ia 
               ON ia.IssueAreaId = R.IssueAreaId
                   INNER JOIN [ResetServices].[BreakDown].WorkingAreas wa ON wa.WorkingAreaID = R.WorkingAreaID
                   INNER JOIN [ResetServices].[BreakDown].WorkingLines wl ON wl.WorkingLineID = R.WorkingLineID
                   INNER JOIN [ResetServices].[BreakDown].WorkingSubAreas ws ON ws.WorkingSubAreaID = R.WorkingSubAreaID
                   INNER JOIN [ResetServices].[BreakDown].IssueProblems i ON i.IssueProblemId = R.IssueProblemId
               WHERE r.DateReport BETWEEN ? AND ? and DescriptionRO <>'TOT BINE'
               ORDER BY r.DateReport 
                """
        try:
            cursor = db_connection.cursor()
            logger.info(f"Executing breakdown query for period: {start_date} to {end_date}")
            cursor.execute(query, start_date, end_date)

            columns = [column[0] for column in cursor.description]
            breakdowns = [dict(zip(columns, row)) for row in cursor.fetchall()]

            logger.info(f"Successfully retrieved {len(breakdowns)} breakdown records.")
            logger.info(f'Ai analysis ....')
            return breakdowns
        except Exception as e:
            logger.error(f"Failed to retrieve breakdown data: {e}", exc_info=True)
            return []

    def analyze_breakdowns(self, breakdown_data: List[Dict], production_data: Dict, period_type: str) -> Dict[str, Any]:
        """
        Performs a full analysis of breakdown data, including statistics and AI insights.

        Args:
            breakdown_data: The list of raw breakdown events.
            production_data: Dictionary with production context (e.g., NrBoards).
            period_type: The analysis period ('weekly' or 'monthly').

        Returns:
            A dictionary containing the full analysis results.
        """
        try:
            # Calcola le statistiche prima di tutto
            statistics = self._calculate_breakdown_statistics(breakdown_data, production_data)

            if not breakdown_data:
                logger.warning("No breakdown data to analyze.")
                return {
                    'statistics': statistics,  # Usa le statistiche (vuote) già calcolate
                    'ai_insights': {'executive_summary': 'No breakdown data available for this period.'},
                    'period_type': period_type,
                    'raw_data': [],
                    'success': True  # È un successo, semplicemente non c'erano dati
                }

            # 1. --- CORREZIONE NOME METODO ---
            # Il metodo corretto in AIAnalyzer si chiama 'analyze_breakdowns'
            logger.info("Calling AI analyzer for breakdown insights...")
            ai_insights = self.ai_analyzer.analyze_breakdowns(statistics, production_data, period_type)

            return {
                'statistics': statistics,
                'ai_insights': ai_insights or {'executive_summary': 'AI analysis failed.'},
                # Assicura che non sia mai None
                'period_type': period_type,
                'raw_data': breakdown_data,
                'success': True
            }
        except Exception as e:
            logger.error(f"An error occurred during breakdown analysis: {e}", exc_info=True)
            # 2. --- CORREZIONE CHIAMATA DI FALLBACK ---
            # Calcola comunque le statistiche passando 'production_data'
            stats_fallback = self._calculate_breakdown_statistics(breakdown_data, production_data)
            return {
                'statistics': stats_fallback,
                'ai_insights': {'executive_summary': 'A critical error occurred during analysis.', 'error': str(e)},
                'period_type': period_type,
                'raw_data': breakdown_data,
                'success': False
            }

    def _calculate_breakdown_statistics(self, breakdown_data: List[Dict], production_data: Dict) -> Dict:
        """
        Calculates key statistics from the breakdown data.

        Args:
            breakdown_data: A list of breakdown events.

        Returns:
            A dictionary containing aggregated statistics.
        """
        if not breakdown_data:
            return {
                'total_stoppages': 0, 'total_downtime_hours': 0,
                'total_boards_produced': production_data.get('NrBoards',0),
                'avg_downtime_minutes': 0, 'problem_frequency': {},
                'top_problems_by_freq': [], 'top_problems_by_time': [],
                'top_lines_by_time': []
            }

        df = pd.DataFrame(breakdown_data)
        df['Hours'] = pd.to_numeric(df['Hours'], errors='coerce').fillna(0)

        total_stoppages = len(df)
        total_downtime_hours = df['Hours'].sum()
        total_boards = production_data.get('NrBoards', 1)

        problem_freq = df['DescriptionRO'].value_counts().to_dict()
        problem_time = df.groupby('DescriptionRO')['Hours'].sum().sort_values(ascending=False).to_dict()
        line_time = df.groupby('WorkingLineName')['Hours'].sum().sort_values(ascending=False).to_dict()

        return {
            'total_stoppages': total_stoppages,
            'total_downtime_hours': round(total_downtime_hours, 2),
            'total_boards_produced': production_data.get('NrBoards', 0),
            'total_orders': production_data.get('NrOrders', 0),
            'avg_downtime_minutes': round((total_downtime_hours * 60) / total_stoppages if total_stoppages > 0 else 0,
                                          2),
            'unique_problems_count': len(problem_freq),
            'problem_frequency': problem_freq,
            'top_problems_by_freq': list(problem_freq.items())[:5],
            'top_problems_by_time': list(problem_time.items())[:5],
            'top_lines_by_time': list(line_time.items())[:5]
        }
