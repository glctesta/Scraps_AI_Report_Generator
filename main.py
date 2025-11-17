"""
AI Quality Analysis - Main Orchestrator
VANDEWIELE ROMANIA SRL - PTHM Department

This application orchestrates three distinct analyses:
1. Scrap Analysis (Weekly)
2. Production Fail Analysis (Weekly & Monthly)
3. Line Breakdown/Stoppage Analysis (Weekly & Monthly)

Each analysis is self-contained, generating its own set of reports (Excel, PDF, and a potential Kaizen PDF)
and sending a dedicated, professional email in English with tabular summaries.
"""
import sys
from datetime import datetime, timedelta
from typing import Dict, Any, List

# Import custom modules
from config_manager import ConfigManager
from db_connection import DatabaseConnection
from utils import get_email_recipients, send_email
from logger_config import setup_logger
from ai_analyzer import AIAnalyzer
from excel_generator import ExcelReportGenerator
from pdf_generator import PDFReportGenerator
from breakdown_analyzer import BreakdownAnalyzer
from fail_analyzer import FailAnalyzer

# Setup main logger
logger = setup_logger('AIScrapAnalysis', 'logs/ai_scrap_analysis.log')


class AIScrapAnalysisApp:
    """Main application orchestrator."""

    def __init__(self):
        """Initializes all components of the application."""
        logger.info("=" * 80)
        logger.info("Initializing AI Quality Analysis Orchestrator")
        logger.info("=" * 80)
        try:
            self.config_manager = ConfigManager(key_file='encryption_key.key', config_file='db_config.enc')
            self.db = DatabaseConnection(self.config_manager)
            self.db.connect()
            logger.info("Database Connection successful.")

            self.email_recipients = get_email_recipients(self.db.connection, 'Sys_email_Quality')
            logger.info(f"Found {len(self.email_recipients)} email recipients.")

            config = self.config_manager.load_config()
            ollama_url = config.get('ollama_url', 'http://localhost:11434')
            self.ai_analyzer = AIAnalyzer(base_url=ollama_url)
            logger.info(f"AIAnalyzer initialized for Ollama at {ollama_url}")

            self.excel_gen = ExcelReportGenerator()
            self.pdf_gen = PDFReportGenerator()
            self.fail_analyzer = FailAnalyzer(self.ai_analyzer)
            self.breakdown_analyzer = BreakdownAnalyzer(self.ai_analyzer)
            logger.info("All components initialized successfully.")

        except Exception as e:
            logger.critical(f"CRITICAL: Application initialization failed: {e}", exc_info=True)
            raise

    def run_complete_analysis(self):
        """Executes a full analysis run for all modules based on the schedule."""
        try:
            current_date = datetime.now()
            is_first_week_of_month = current_date.day <= 7

            logger.info("### 1. STARTING WEEKLY SCRAP ANALYSIS ###")
            self.run_scrap_analysis()

            logger.info("### 2. STARTING WEEKLY FAIL ANALYSIS ###")
            self._run_fail_analysis(period_type='weekly')

            logger.info("### 3. STARTING WEEKLY BREAKDOWN ANALYSIS ###")
            self._run_breakdown_analysis(period_type='weekly')

            if is_first_week_of_month:
                logger.info("="*50)
                logger.info("First week of the month detected. Running monthly reports with YTD data.")
                logger.info("="*50)
                
                logger.info("### 4. STARTING MONTHLY FAIL ANALYSIS ###")
                self._run_fail_analysis(period_type='monthly')

                logger.info("### 5. STARTING MONTHLY BREAKDOWN ANALYSIS ###")
                self._run_breakdown_analysis(period_type='monthly')

            logger.info("All scheduled analyses completed.")
            return {'success': True}

        except Exception as e:
            logger.error(f"A critical error occurred during the complete analysis run: {e}", exc_info=True)
            return {'success': False, 'error': str(e)}

    # ===================================================================
    # --- ANALYSIS ORCHESTRATION METHODS ---
    # ===================================================================

    def run_scrap_analysis(self):
        """Executes scrap analysis."""
        try:
            start_date, end_date, period_str = self._get_dates_for_period('weekly')
            logger.info(f"Running SCRAP analysis for period: {period_str}")

            production_data = self._get_production_data(start_date, end_date)
            scraps_data = self._get_scraps_data(start_date, end_date)
            if not scraps_data:
                logger.warning("No scrap data found. Skipping.")
                return

            top_defects = self._calculate_top_defects(scraps_data)
            statistics = self._calculate_scrap_statistics(production_data, scraps_data)
            ai_insights = self.ai_analyzer.analyze_defects(top_defects, production_data)
            
            report_data = self._prepare_scrap_report_data(period_str, production_data, scraps_data, top_defects, ai_insights, statistics)

            attachments = self._generate_reports_and_get_paths(report_data, "Scrap_Analysis", "weekly")
            
            email_subject, email_body = self._generate_email_for_analysis(report_data)
            if self.email_recipients and attachments:
                send_email(recipients=self.email_recipients, subject=email_subject, body=email_body, is_html=True, attachments=attachments)
                logger.info("Scrap analysis email sent successfully.")
        except Exception as e:
            logger.error(f"Error during SCRAP analysis orchestration: {e}", exc_info=True)

    def _run_fail_analysis(self, period_type: str):
        """Runs fail analysis for a given period."""
        try:
            start_date, end_date, period_str = self._get_dates_for_period(period_type)
            logger.info(f"Running {period_type.upper()} FAIL analysis for: {period_str}")

            production_data = self._get_production_data(start_date, end_date)
            fail_data = self.fail_analyzer.get_fail_data(self.db.connection, start_date, end_date)
            if not fail_data:
                logger.warning(f"No fail data found. Skipping.")
                return

            analysis_result = self.fail_analyzer.analyze_fails(fail_data, production_data, period_type)
            report_data = self._prepare_fail_report_data(analysis_result, period_str, fail_data)

            if period_type == 'monthly':
                report_data['ytd_data'] = self._get_ytd_fail_data()

            attachments = self._generate_reports_and_get_paths(report_data, "Fail_Analysis", period_type)
            
            email_subject, email_body = self._generate_email_for_analysis(report_data)
            if self.email_recipients and attachments:
                send_email(recipients=self.email_recipients, subject=email_subject, body=email_body, is_html=True, attachments=attachments)
                logger.info(f"{period_type.title()} fail analysis email sent successfully.")
        except Exception as e:
            logger.error(f"Failed to run {period_type} FAIL analysis: {e}", exc_info=True)

    def _run_breakdown_analysis(self, period_type: str):
        """Runs breakdown analysis for a given period."""
        try:
            start_date, end_date, period_str = self._get_dates_for_period(period_type)
            logger.info(f"Running {period_type.upper()} BREAKDOWN analysis for: {period_str}")

            production_data = self._get_production_data(start_date, end_date)
            breakdown_data = self.breakdown_analyzer.get_breakdown_data(self.db.connection, start_date, end_date)
            if not breakdown_data:
                logger.warning(f"No breakdown data found. Skipping.")
                return

            analysis_result = self.breakdown_analyzer.analyze_breakdowns(breakdown_data, production_data, period_type)
            report_data = self._prepare_breakdown_report_data(analysis_result, production_data, period_str, breakdown_data)

            if period_type == 'monthly':
                 report_data['ytd_data'] = self._get_ytd_breakdown_data()
            
            attachments = self._generate_reports_and_get_paths(report_data, "Breakdown_Analysis", period_type)

            email_subject, email_body = self._generate_email_for_analysis(report_data)
            if self.email_recipients and attachments:
                send_email(recipients=self.email_recipients, subject=email_subject, body=email_body, is_html=True, attachments=attachments)
                logger.info(f"{period_type.title()} breakdown analysis email sent successfully.")
        except Exception as e:
            logger.error(f"Failed to run {period_type} BREAKDOWN analysis: {e}", exc_info=True)
            
    # ===================================================================
    # --- HELPER & UTILITY METHODS ---
    # ===================================================================
    
    def _generate_email_for_analysis(self, report_data: Dict) -> tuple[str, str]:
        """
        Generates a professional, data-rich HTML email for any analysis type.
        """
        analysis_type = report_data.get('analysis_type', 'Quality Report')
        period = report_data.get('period')
        subject = f"AI {analysis_type} - {period}"
        summary = report_data.get('executive_summary', 'AI analysis summary could not be generated.')
        stats = report_data.get('statistics', {})
        
        # --- Build Metrics Table ---
        metrics_html = ""
        if 'scrap_rate' in stats:
            metrics_html += f"<tr><td>Total Scraps</td><td style='text-align:center;'>{stats.get('total_scraps', 0)}</td></tr>"
            metrics_html += f"<tr><td>Scrap Rate</td><td style='text-align:center;'>{stats.get('scrap_rate', 0):.2f}%</td></tr>"
        if 'fail_rate' in stats:
            metrics_html += f"<tr><td>Total Fails</td><td style='text-align:center;'>{stats.get('total_fails', 0)}</td></tr>"
            metrics_html += f"<tr><td>Fail Rate</td><td style='text-align:center;'>{stats.get('fail_rate', 0):.2f}%</td></tr>"
        if 'total_downtime_hours' in stats:
            metrics_html += f"<tr><td>Total Stoppages</td><td style='text-align:center;'>{stats.get('total_stoppages', 0)}</td></tr>"
            metrics_html += f"<tr><td>Total Downtime</td><td style='text-align:center;'>{stats.get('total_downtime_hours', 0):.2f} hrs</td></tr>"

        # --- Build Top Issues Table ---
        chart_data = report_data.get('chart_data', [])
        top_issues_html = ""
        if chart_data:
            for item in chart_data[:5]:
                top_issues_html += f"<tr><td>{item.get('label', 'N/A')}</td><td style='text-align:center;'>{item.get('value', 0)}</td></tr>"
        
        # --- Build Recommendations Section ---
        recommendations_html = ""
        recommendations = report_data.get('recommendations', [])
        if recommendations:
            for rec in recommendations[:3]: # Show top 3
                p_color = {'High': '#D9534F', 'Medium': '#F0AD4E', 'Low': '#5CB85C'}.get(rec.get('priority'), '#777777')
                recommendations_html += f"""
                <div style="margin-bottom: 15px; border-left: 4px solid {p_color}; padding-left: 10px;">
                    <strong style="color: {p_color};">[{rec.get('priority', 'N/A')}] {rec.get('title', 'Recommendation')}</strong><br>
                    <span style="font-size: 0.9em; color: #555;">{rec.get('description', 'No description.')}</span>
                </div>
                """

        # --- Build Kaizen Proposal Box ---
        kaizen_html = ""
        kaizen = report_data.get('kaizen_proposal')
        if kaizen and kaizen.get('project_title'):
             kaizen_html = f"""
            <h3 style="color:#0056b3; border-bottom: 1px solid #ddd; padding-bottom: 5px;">Kaizen Project Proposal</h3>
            <div style="background-color: #e9f5ff; padding: 15px; border-radius: 5px;">
                <p style="margin:0; font-size:1.1em;"><strong>Title:</strong> {kaizen.get('project_title', 'N/A')}</p>
                <p style="margin-top:10px;"><strong>Goal:</strong> {kaizen.get('goal', 'N/A')}</p>
                <p style="font-size:0.9em; margin-top:15px;"><i>A dedicated PDF charter for this project is attached.</i></p>
            </div>
            """

        # --- Assemble Full Email Body ---
        body = f"""
        <html>
          <body style="font-family: Arial, sans-serif; color: #333; background-color: #f4f4f4; padding: 20px;">
            <div style="max-width: 800px; margin: auto; background-color: #ffffff; border: 1px solid #ddd; padding: 30px; border-radius: 8px;">
              <h1 style="color: #0056b3; text-align: center; border-bottom: 2px solid #0056b3; padding-bottom: 10px;">{analysis_type} Report</h1>
              <p style="text-align: center; color: #555;"><strong>Period:</strong> {period}</p>
              
              <h2 style="color:#0056b3; border-bottom: 1px solid #ddd; padding-bottom: 5px;">AI Executive Summary</h2>
              <p style="font-size: 1.05em; line-height: 1.6;">{summary}</p>
              
              <table style="width: 100%; margin-top: 20px;">
                <tr>
                  <td style="width: 48%; vertical-align: top;">
                    <h3 style="color:#0056b3;">Key Metrics</h3>
                    <table style="width: 100%; border-collapse: collapse;">{metrics_html}</table>
                  </td>
                  <td style="width: 4%;"></td>
                  <td style="width: 48%; vertical-align: top;">
                    <h3 style="color:#0056b3;">Top 5 Issues</h3>
                    <table style="width: 100%; border-collapse: collapse;">{top_issues_html}</table>
                  </td>
                </tr>
              </table>

              <h2 style="color:#0056b3; border-bottom: 1px solid #ddd; padding-bottom: 5px; margin-top: 30px;">AI Recommendations</h2>
              {recommendations_html or "<p><i>No specific recommendations provided by AI.</i></p>"}

              {kaizen_html}

              <hr style="margin-top: 30px;">
              <p style="text-align: center; font-size: 0.9em; color: #777;">
                The full analysis, including detailed data and charts, is available in the attached PDF and Excel reports.<br>
                This is an automated report generated by the AI Quality Analysis System.
              </p>
            </div>
          </body>
        </html>
        """
        return subject, body

    def _generate_reports_and_get_paths(self, report_data: Dict, prefix: str, period_type: str) -> List[str]:
        """Generates all standard and conditional reports and returns their paths."""
        attachments = []
        filename_stamp = f"{prefix}_{period_type}_{datetime.now().strftime('%Y%m%d')}"
        
        excel_path = self.excel_gen.generate_report(report_data, f"reports/{filename_stamp}.xlsx")
        if excel_path: attachments.append(excel_path)
        
        pdf_path = self._generate_generic_pdf_report(report_data, f"{prefix.replace('_', ' ')} - {period_type.title()}")
        if pdf_path: attachments.append(pdf_path)

        kaizen_data = report_data.get('kaizen_proposal')
        if kaizen_data and kaizen_data.get('project_title'):
            logger.info("Kaizen proposal found. Generating Kaizen Charter PDF.")
            kaizen_pdf_path = self.pdf_gen.generate_kaizen_pdf(kaizen_data)
            if kaizen_pdf_path: attachments.append(kaizen_pdf_path)
            
        return attachments

    # ... (Il resto dei metodi come _get_dates_for_period, _prepare_*, _get_*, _calculate_* sono stati omessi per brevità, ma DEVONO rimanere nel tuo file) ...
    # Assicurati di avere tutti gli altri metodi che erano qui prima!
    
# METODI OMESSI PER BREVITA (DEVONO ESSERE MANTENUTI NEL TUO FILE)
    def _get_dates_for_period(self, period_type: str) -> tuple[str, str, str]:
        # Implementation from previous version
        end_date_dt = datetime.now()
        if period_type == 'weekly':
            start_date_dt = end_date_dt - timedelta(days=7)
            period_str = f"{start_date_dt.strftime('%Y-%m-%d')} to {end_date_dt.strftime('%Y-%m-%d')}"
        else:
            end_of_last_month = end_date_dt.replace(day=1) - timedelta(days=1)
            start_date_dt = end_of_last_month.replace(day=1)
            period_str = start_date_dt.strftime('%B %Y')
        return start_date_dt.strftime('%Y-%m-%d'), end_date_dt.strftime('%Y-%m-%d'), period_str

    def _prepare_scrap_report_data(self, period_str, production_data, scraps_data, top_defects, ai_insights, statistics) -> Dict[str, Any]:
        chart_data = [{'label': d.get('DefectName', 'N/A'), 'value': d.get('Count', 0)} for d in top_defects]
        return {'analysis_type': 'Scrap Analysis','period': period_str,'generation_date': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),'production_data': production_data,'statistics': statistics,'executive_summary': ai_insights.get('executive_summary', "AI summary not available."),'root_causes': ai_insights.get('root_causes', []),'recommendations': ai_insights.get('recommendations', []),'raw_data': scraps_data, 'chart_data': chart_data}
    
    def _prepare_fail_report_data(self, analysis_result: Dict, period_str: str, raw_data: List) -> Dict[str, Any]:
        stats = analysis_result.get('statistics', {})
        ai = analysis_result.get('ai_insights', {})
        chart_data = [{'label': d.get('defect', 'N/A'), 'value': d.get('count', 0)} for d in stats.get('top_defects', [])]
        return {'analysis_type': 'Production Fail Analysis','period': period_str,'generation_date': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),'statistics': stats,'executive_summary': ai.get('executive_summary', "AI summary not available."),'root_causes': ai.get('root_causes', []),'recommendations': ai.get('recommendations', []),'kaizen_proposal': ai.get('kaizen_project_proposal'),'raw_data': raw_data,'chart_data': chart_data}

    def _prepare_breakdown_report_data(self, analysis_result: Dict, production_data: Dict, period_str: str, raw_data: List) -> Dict[str, Any]:
        stats = analysis_result['statistics']
        ai = analysis_result['ai_insights']
        chart_data = [{'label': item[0], 'value': item[1]} for item in stats.get('top_problems_by_time', [])]
        return {'analysis_type': 'Line Stoppage Analysis','period': period_str,'generation_date': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),'production_data': production_data,'statistics': stats,'executive_summary': ai.get('executive_summary', "AI summary not available."),'root_causes': ai.get('root_causes', []),'recommendations': ai.get('recommendations', []),'kaizen_proposal': ai.get('kaizen_project_proposal'),'raw_data': raw_data,'chart_data': chart_data}
    
    def _generate_generic_pdf_report(self, report_data: Dict, title: str) -> str:
        try:
            self.pdf_gen.title = title
            pdf_path = self.pdf_gen.generate_report(report_data) 
            logger.info(f"Successfully delegated PDF generation: {pdf_path}")
            return pdf_path
        except Exception as e:
            logger.error(f"The PDF generation process failed: {e}", exc_info=True)
            return ""

    def _get_ytd_fail_data(self) -> List[Dict]:
        current_year = datetime.now().year
        query = f"""SELECT FORMAT(CAST(DataVerify AS DATE), 'yyyy-MM') AS Month, COUNT(*) AS TotalFails FROM traceability_rs.dbo.QualityVerify WHERE IsPass = 0 AND YEAR(DataVerify) = {current_year} GROUP BY FORMAT(CAST(DataVerify AS DATE), 'yyyy-MM') ORDER BY Month;"""
        try:
            cursor = self.db.connection.cursor()
            cursor.execute(query)
            return [dict(zip([c[0] for c in cursor.description], row)) for row in cursor.fetchall()]
        except Exception as e:
            logger.error(f"Failed to get YTD fail data: {e}", exc_info=True)
            return []

    def _get_ytd_breakdown_data(self) -> List[Dict]:
        current_year = datetime.now().year
        query = f"""SELECT FORMAT(CAST(DateReport AS DATE), 'yyyy-MM') AS Month, COUNT(*) AS TotalStoppages, SUM(CAST(Hours AS float)) AS TotalDowntime FROM [ResetServices].[BreakDown].[ReportIssueLogs] WHERE YEAR(DateReport) = {current_year} AND DescriptionRO <> 'TOT BINE' GROUP BY FORMAT(CAST(DateReport AS DATE), 'yyyy-MM') ORDER BY Month;"""
        try:
            cursor = self.db.connection.cursor()
            cursor.execute(query)
            return [dict(zip([c[0] for c in cursor.description], row)) for row in cursor.fetchall()]
        except Exception as e:
            logger.error(f"Failed to get YTD breakdown data: {e}", exc_info=True)
            return []

    def _get_production_data(self, start_date: str, end_date: str) -> dict:
        query = """SELECT COUNT(DISTINCT o.OrderNumber) AS TotalOrders, COUNT(distinct b.IDBoard) AS TotalBoards FROM [Traceability_RS].[dbo].Orders o LEFT JOIN [Traceability_RS].[dbo].Boards b ON o.IDOrder = b.IDOrder WHERE cast(b.CreationDate as date) BETWEEN ? AND ?"""
        try:
            conn = self.db.connection
            cursor = conn.cursor()
            cursor.execute(query, (start_date, end_date))
            row = cursor.fetchone()
            cursor.close()
            return {'NrOrders': row.TotalOrders or 0, 'NrBoards': row.TotalBoards or 0}
        except Exception as e:
            logger.error(f"Failed to get production data: {e}", exc_info=True)
            return {'NrOrders': 0, 'NrBoards': 0}

    def _get_scraps_data(self, start_date: str, end_date: str) -> list:
        query = """SELECT s.ScrapDeclarationId, s.[User] as DeclaredBy, FORMAT(s.DateIn, 'dd/MM/yyyy') as [Date], o.OrderNumber, l.labelcod, p.productCode as Product, A.AreaName, d.DefectNameRO as Defect
                   FROM [Traceability_RS].[dbo].ScarpDeclarations S
                   INNER JOIN Traceability_RS.dbo.LabelCodes L ON l.IDLabelCode = s.IdLabelCode
                   INNER JOIN [Traceability_RS].[dbo].Areas A ON a.IDArea = s.IDParentPhase
                   INNER JOIN [Traceability_RS].[dbo].defects D ON d.IDDefect = s.ScrapReasonId
                   INNER JOIN [Traceability_RS].[dbo].boards B ON l.IDBoard = b.IDBoard
                   INNER JOIN [Traceability_RS].[dbo].orders o ON o.idorder = b.IDOrder
                   INNER JOIN traceability_rs.dbo.products P on p.idproduct=o.idproduct
                   WHERE (s.Refuzed IS NULL OR s.Refuzed = 0) AND CAST(s.DateIn as date) BETWEEN ? AND ?
                   ORDER BY S.DateIn DESC"""
        try:
            conn = self.db.connection
            cursor = conn.cursor()
            cursor.execute(query, (start_date, end_date))
            columns = [column[0] for column in cursor.description]
            return [dict(zip(columns, row)) for row in cursor.fetchall()]
        except Exception as e:
            logger.error(f"Failed to get scraps data: {e}", exc_info=True)
            return []

    def _calculate_top_defects(self, scraps_data: list) -> list:
        if not scraps_data: return []
        defect_counts = {}
        for scrap in scraps_data:
            defect_name = scrap['Defect']
            defect_counts[defect_name] = defect_counts.get(defect_name, 0) + 1
        top_defects = [{'DefectName': k, 'Count': v} for k, v in defect_counts.items()]
        return sorted(top_defects, key=lambda x: x['Count'], reverse=True)

    def _calculate_scrap_statistics(self, production_data: dict, scraps_data: list) -> dict:
        nr_boards = production_data.get('NrBoards', 1)
        scrap_count = len(scraps_data)
        scrap_rate = (scrap_count / nr_boards * 100) if nr_boards > 0 else 0
        return {'scrap_rate': scrap_rate, 'total_scraps': scrap_count, 'total_boards': nr_boards}


if __name__ == "__main__":
    try:
        app = AIScrapAnalysisApp()
        app.run_complete_analysis()
    except Exception as e:
        logger.critical(f"FATAL: Application failed to run. Error: {e}", exc_info=True)
        sys.exit(1)