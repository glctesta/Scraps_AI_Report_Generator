"""
AI Scrap Analysis - Main Orchestrator
VANDEWIELE ROMANIA SRL - PTHM Department
"""
import sys
from datetime import datetime, timedelta
from pathlib import Path

# Import tuoi moduli esistenti
from config_manager import ConfigManager
from db_connection import DatabaseConnection
from email_sender import EmailSender
from utils import get_email_recipients

# Import nuovi moduli
from logger_config import setup_logger
from ai_analyzer import AIAnalyzer
from report_generator import ReportGenerator
from pdf_generator import PDFReportGenerator, generate_pdf_report
from excel_generator import ExcelReportGenerator, generate_excel_report

# Setup logger
logger = setup_logger('AIScrapAnalysis', 'logs/ai_scrap_analysis.log')


class AIScrapAnalysisApp:
    """Applicazione principale per AI Scrap Analysis"""

    def __init__(self):
        """Inizializza applicazione"""
        logger.info("=" * 80)
        logger.info("Avvio AI Scrap Analysis")
        logger.info("=" * 80)

        try:
            # Config Manager (usa il TUO file .enc)
            self.config_manager = ConfigManager(
                key_file='encryption_key.key',
                config_file='db_config.enc'
            )
            logger.info("ConfigManager inizializzato")

            # Database Connection (usa la TUA classe)
            self.db = DatabaseConnection(self.config_manager)
            self.db.connect()
            logger.info("DatabaseConnection connesso")

            # DEBUG: Verifica connessione DB
            self._test_database_connection()

            # Recupera destinatari email dal DB (usa la TUA funzione utils)
            logger.info("Recupero destinatari email...")
            self.email_recipients = get_email_recipients(self.db.connection, 'Sys_email_Quality')
            logger.info(f"Destinatari email trovati: {len(self.email_recipients)}")
            if self.email_recipients:
                logger.info(f"Lista destinatari: {self.email_recipients}")

            # AI Analyzer (opzionale - se hai API key)
            ai_api_key = self._load_ai_config()
            self.ai_analyzer = AIAnalyzer(
                api_key=ai_api_key.get('api_key'),
                provider=ai_api_key.get('provider', 'anthropic')
            )
            logger.info("AIAnalyzer inizializzato")

            # Report Generator
            self.report_gen = ReportGenerator(output_dir='reports')
            logger.info("ReportGenerator inizializzato")

            # PDF Generator
            self.pdf_gen = PDFReportGenerator(title="AI Scrap Analysis Report")
            logger.info("PDFReportGenerator inizializzato")

            # Excel Generator
            self.excel_gen = ExcelReportGenerator(title="AI Scrap Analysis Report")
            logger.info("ExcelReportGenerator inizializzato")

            # Email Sender (usa la TUA classe email_sender.py)
            self.email_sender = self._init_email_sender()
            logger.info("EmailSender inizializzato")

        except Exception as e:
            logger.error(f"Errore inizializzazione: {e}")
            raise

    def _test_database_connection(self):
        """Testa la connessione al database e verifica le tabelle"""
        try:
            # Usa la connessione dal DB connection - VERIFICA che esista
            if hasattr(self.db, 'connection') and self.db.connection:
                cursor = self.db.connection.cursor()

                # Test query generica
                cursor.execute("SELECT TOP 1 name FROM sys.databases")
                result = cursor.fetchone()
                logger.info(f"Connessione DB testata: {result[0] if result else 'OK'}")

                # Verifica esistenza tabella settings
                cursor.execute("""
                               SELECT TABLE_NAME
                               FROM [traceability_rs].INFORMATION_SCHEMA.TABLES
                               WHERE TABLE_NAME = 'settings'
                               """)
                settings_table = cursor.fetchone()
                if settings_table:
                    logger.info("Tabella 'settings' trovata")
                else:
                    logger.warning("Tabella 'settings' non trovata")

                cursor.close()
            else:
                logger.warning("Connessione DB non disponibile per test")

        except Exception as e:
            logger.error(f"Errore test connessione DB: {e}")

    def _load_ai_config(self) -> dict:
        """
        Carica configurazione AI - Ora supporta Ollama
        """
        ai_config_file = Path('ai_config.json')

        if ai_config_file.exists():
            import json
            with open(ai_config_file, 'r') as f:
                config = json.load(f)
                # Supporta configurazione Ollama
                if config.get('provider') == 'ollama':
                    config['api_key'] = None  # Ollama non richiede API key
                return config
        else:
            logger.warning("File ai_config.json non trovato - Uso Ollama con llama3.2")
            return {
                'api_key': None,
                'provider': 'ollama',
                'model': 'llama3.2'
            }

    def _init_email_sender(self):
        """
        Inizializza EmailSender usando la configurazione esistente
        """
        try:
            logger.info("Inizializzazione EmailSender con relay server interno...")

            # NON inizializzare EmailSender qui - usa direttamente utils.send_email()
            # che già gestisce tutto correttamente
            logger.info("EmailSender configurato tramite utils.send_email()")
            return None  # Non serve restituire un'istanza

        except Exception as e:
            logger.error(f"Errore inizializzazione EmailSender: {e}")
            raise

    def run_analysis(self, start_date: str = None, end_date: str = None):
        """
        Esegue analisi completa
        """
        try:
            # Date di default: ultima settimana
            if not start_date:
                start_date = (datetime.now() - timedelta(days=7)).strftime('%Y-%m-%d')
            if not end_date:
                end_date = datetime.now().strftime('%Y-%m-%d')

            logger.info(f"Periodo analisi: {start_date} -> {end_date}")

            # 1. Recupera dati produzione
            logger.info("Recupero dati produzione...")
            production_data = self._get_production_data(start_date, end_date)
            logger.info(f"Ordini: {production_data['NrOrders']}, Schede: {production_data['NrBoards']}")

            # 2. Recupera difetti
            logger.info("Recupero difetti...")
            scraps_data = self._get_scraps_data(start_date, end_date)
            logger.info(f"Difetti totali: {len(scraps_data)}")

            # 3. Calcola top defects
            logger.info("Calcolo top defects...")
            top_defects = self._calculate_top_defects(scraps_data)

            # 4. Analisi AI
            logger.info("Analisi AI in corso...")
            ai_insights = self.ai_analyzer.analyze_defects(top_defects, production_data)
            logger.info(f"Root causes: {len(ai_insights.get('root_causes', []))}")
            logger.info(f"Recommendations: {len(ai_insights.get('recommendations', []))}")

            # 5. Calcola statistiche
            analysis = self._calculate_statistics(production_data, scraps_data)

            # 6. Prepara dati report completi
            report_data = self._prepare_comprehensive_report_data(
                start_date, end_date, production_data, scraps_data,
                top_defects, ai_insights, analysis
            )

            # 7. Genera Excel avanzato
            logger.info("Generazione report Excel avanzato...")
            excel_data = self.excel_gen.generate_report(report_data)

            # Salva Excel temporaneamente
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            excel_filename = f"AI_Scrap_Analysis_{timestamp}.xlsx"
            excel_path = Path('reports') / excel_filename
            with open(excel_path, 'wb') as f:
                f.write(excel_data)
            logger.info(f"Excel salvato: {excel_path}")

            # 8. Genera PDF avanzato
            logger.info("Generazione report PDF...")
            pdf_data = self.pdf_gen.generate_report(report_data)

            # Salva PDF temporaneamente
            pdf_filename = f"AI_Scrap_Analysis_{timestamp}.pdf"
            pdf_path = Path('reports') / pdf_filename
            with open(pdf_path, 'wb') as f:
                f.write(pdf_data)
            logger.info(f"PDF salvato: {pdf_path}")

            # 9. Genera corpo email professionale in inglese
            logger.info("Generazione corpo email professionale...")
            email_body, email_subject = self._generate_professional_email(report_data)

            # 10. Invia email con allegati
            if self.email_recipients:
                logger.info(f"Invio email a {len(self.email_recipients)} destinatari...")

                # Usa la funzione send_email da utils.py che già funziona
                from utils import send_email

                send_email(
                    recipients=self.email_recipients,
                    subject=email_subject,
                    body=email_body,  # Usa l'HTML già generato
                    is_html=True,  # Specifica che è HTML
                    attachments=[str(excel_path), str(pdf_path)]
                )

                logger.info("Email inviata con successo!")
            else:
                logger.warning("Nessun destinatario email configurato - Invio saltato")

            # 11. Pulisci file temporanei
            try:
                excel_path.unlink()
                pdf_path.unlink()
                logger.info("File temporanei puliti")
            except Exception as e:
                logger.warning(f"Errore nella pulizia file temporanei: {e}")

            logger.info("=" * 80)
            logger.info("Analisi completata con successo!")
            logger.info("=" * 80)

            return {
                'success': True,
                'excel_path': str(excel_path),
                'pdf_path': str(pdf_path),
                'report_data': report_data
            }

        except Exception as e:
            logger.error(f"Errore durante analisi: {e}", exc_info=True)
            return {
                'success': False,
                'error': str(e)
            }
        finally:
            # Chiudi connessione DB
            self.db.disconnect()

    def _prepare_comprehensive_report_data(self, start_date, end_date, production_data,
                                         scraps_data, top_defects, ai_insights, analysis):
        """Prepara dati completi per report PDF ed Excel"""

        # Calcola metriche avanzate
        scrap_rate = analysis.get('scrap_rate', 0)
        total_scraps = analysis.get('total_scraps', 0)
        total_boards = analysis.get('total_boards', 0)

        # Prepara executive summary
        executive_summary = self._generate_executive_summary_text(
            start_date, end_date, production_data, scrap_rate, total_scraps, ai_insights
        )

        # Prepara metriche dettagliate
        metrics = self._prepare_metrics_data(production_data, analysis, ai_insights)

        # Prepara raccomandazioni
        recommendations = self._prepare_recommendations_data(ai_insights)

        # Prepara dati raw per tabelle
        raw_data = self._prepare_raw_data(scraps_data, top_defects)

        return {
            'version': '1.0',
            'analysis_type': 'AI Scrap Analysis',
            'period': f"{start_date} to {end_date}",
            'executive_summary': executive_summary,
            'production_data': production_data,
            'scrap_data': {
                'total_scraps': total_scraps,
                'scrap_rate': scrap_rate,
                'total_boards': total_boards
            },
            'top_defects': top_defects,
            'ai_insights': ai_insights,
            'analysis': analysis,
            'metrics': metrics,
            'recommendations': recommendations,
            'raw_data': raw_data,
            'generation_date': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        }

    def _generate_executive_summary_text(self, start_date, end_date, production_data, scrap_rate, total_scraps, ai_insights):
        """Genera testo executive summary"""

        status = "EXCELLENT" if scrap_rate < 1.0 else "GOOD" if scrap_rate < 3.0 else "REQUIRES ATTENTION"

        summary = f"""
AI Scrap Analysis Report for period {start_date} to {end_date}

PRODUCTION OVERVIEW:
- Total Orders: {production_data.get('NrOrders', 0):,}
- Total Boards: {production_data.get('NrBoards', 0):,}
- Total Defects: {total_scraps:,}
- Scrap Rate: {scrap_rate:.2f}%

STATUS: {status}

KEY FINDINGS:
"""

        # Aggiungi insights principali dall'AI
        root_causes = ai_insights.get('root_causes', [])
        if root_causes:
            summary += "\nPRIMARY ROOT CAUSES IDENTIFIED:\n"
            for cause in root_causes[:3]:
                summary += f"- {cause.get('cause', '')}\n"

        recommendations = ai_insights.get('recommendations', [])
        if recommendations:
            summary += "\nTOP RECOMMENDATIONS:\n"
            for rec in recommendations[:3]:
                summary += f"- {rec.get('title', '')} (Priority: {rec.get('priority', 'Medium')})\n"

        return summary

    def _prepare_metrics_data(self, production_data, analysis, ai_insights):
        """Prepara dati metriche per report"""

        scrap_rate = analysis.get('scrap_rate', 0)
        total_scraps = analysis.get('total_scraps', 0)

        # Metriche di base
        metrics = [
            {
                'name': 'Scrap Rate',
                'value': f"{scrap_rate:.2f}%",
                'target': '< 2.0%',
                'achieved': scrap_rate < 2.0
            },
            {
                'name': 'Total Defects',
                'value': f"{total_scraps:,}",
                'target': 'Minimize',
                'achieved': total_scraps < 100
            },
            {
                'name': 'Production Volume',
                'value': f"{production_data.get('NrBoards', 0):,}",
                'target': 'Maximize',
                'achieved': True
            }
        ]

        # Aggiungi metriche specifiche dall'AI analysis
        if ai_insights.get('analysis_type') == 'ai':
            metrics.append({
                'name': 'AI Analysis Quality',
                'value': 'High',
                'target': 'High',
                'achieved': True
            })

        return metrics

    def _prepare_recommendations_data(self, ai_insights):
        """Prepara dati raccomandazioni"""
        recommendations = ai_insights.get('recommendations', [])

        enhanced_recommendations = []
        for rec in recommendations:
            enhanced_recommendations.append({
                'priority': rec.get('priority', 'Medium'),
                'description': rec.get('description', ''),
                'title': rec.get('title', ''),
                'impact': rec.get('impact', 'Process Improvement')
            })

        return enhanced_recommendations

    def _prepare_raw_data(self, scraps_data, top_defects):
        """Prepara dati raw per tabelle"""
        raw_data = []

        # Tabella top defects
        for defect in top_defects[:10]:
            raw_data.append({
                'Defect Name': defect.get('DefectName', ''),
                'Count': defect.get('Count', 0),
                'Percentage': f"{(defect.get('Count', 0) / sum(d['Count'] for d in top_defects) * 100):.1f}%",
                'Top Area': defect.get('TopArea', ''),
                'Products Affected': defect.get('ProductCount', 0)
            })

        return raw_data

    def _generate_professional_email(self, report_data):
        """Genera email professionale in inglese"""

        period = report_data.get('period', 'N/A')
        scrap_rate = report_data.get('scrap_data', {}).get('scrap_rate', 0)
        total_scraps = report_data.get('scrap_data', {}).get('total_scraps', 0)
        total_boards = report_data.get('scrap_data', {}).get('total_boards', 0)

        # Determina status basato su scrap rate
        if scrap_rate < 1.0:
            status_text = "EXCELLENT"
            status_color = "#28a745"
        elif scrap_rate < 3.0:
            status_text = "GOOD"
            status_color = "#ffc107"
        else:
            status_text = "REQUIRES ATTENTION"
            status_color = "#dc3545"

        subject = f"AI Scrap Analysis Report - {period} - {status_text}"

        # Top defects per email
        top_defects = report_data.get('top_defects', [])[:5]
        defects_html = ""
        for i, defect in enumerate(top_defects, 1):
            defects_html += f"""
            <tr>
                <td style="padding: 8px; border-bottom: 1px solid #ddd;">{i}. {defect.get('DefectName', '')}</td>
                <td style="padding: 8px; border-bottom: 1px solid #ddd; text-align: center;">{defect.get('Count', 0)}</td>
                <td style="padding: 8px; border-bottom: 1px solid #ddd; text-align: center;">{(defect.get('Count', 0) / total_scraps * 100):.1f}%</td>
            </tr>
            """

        # Root causes
        root_causes = report_data.get('ai_insights', {}).get('root_causes', [])[:3]
        causes_html = ""
        for cause in root_causes:
            causes_html += f"""
            <div style="background: #fff3cd; padding: 10px; margin: 5px 0; border-left: 4px solid #ffc107;">
                <strong>{cause.get('category', '')}:</strong> {cause.get('cause', '')}
            </div>
            """

        # Recommendations
        recommendations = report_data.get('ai_insights', {}).get('recommendations', [])[:3]
        rec_html = ""
        for rec in recommendations:
            priority_color = {
                'High': '#dc3545',
                'Medium': '#fd7e14',
                'Low': '#198754'
            }.get(rec.get('priority', 'Medium'), '#6c757d')

            rec_html += f"""
            <div style="background: #e7f3ff; padding: 10px; margin: 5px 0; border-left: 4px solid {priority_color};">
                <strong>{rec.get('title', '')}</strong> 
                <span style="color: {priority_color}; font-weight: bold;">[{rec.get('priority', 'Medium')}]</span><br>
                {rec.get('description', '')}
            </div>
            """

        html_body = f"""
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <style>
        body {{ font-family: Arial, sans-serif; line-height: 1.6; color: #333; }}
        .container {{ max-width: 800px; margin: 0 auto; padding: 20px; }}
        .header {{ background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 30px; text-align: center; }}
        .status {{ background: #f8f9fa; padding: 20px; border-radius: 5px; margin: 20px 0; border-left: 5px solid {status_color}; }}
        .section {{ margin: 25px 0; }}
        .metric-grid {{ display: grid; grid-template-columns: repeat(2, 1fr); gap: 15px; }}
        .metric-card {{ background: white; padding: 15px; border-radius: 5px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }}
        table {{ width: 100%; border-collapse: collapse; margin: 15px 0; }}
        th {{ background: #f8f9fa; padding: 12px; text-align: left; }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>AI Scrap Analysis Report</h1>
            <p>VANDEWIELE ROMANIA SRL - PTHM Department</p>
            <p>Period: {period}</p>
        </div>

        <div class="status">
            <h2>Overall Status: {status_text}</h2>
            <p>Scrap Rate: <strong>{scrap_rate:.2f}%</strong></p>
        </div>

        <div class="section">
            <h3>Key Production Metrics</h3>
            <div class="metric-grid">
                <div class="metric-card">
                    <h4>Total Boards</h4>
                    <p style="font-size: 24px; font-weight: bold; color: #667eea;">{total_boards:,}</p>
                </div>
                <div class="metric-card">
                    <h4>Total Defects</h4>
                    <p style="font-size: 24px; font-weight: bold; color: {'#dc3545' if total_scraps > 100 else '#28a745'};">{total_scraps:,}</p>
                </div>
                <div class="metric-card">
                    <h4>Scrap Rate</h4>
                    <p style="font-size: 24px; font-weight: bold; color: {'#dc3545' if scrap_rate > 3.0 else '#28a745'};">{scrap_rate:.2f}%</p>
                </div>
                <div class="metric-card">
                    <h4>AI Analysis</h4>
                    <p style="font-size: 24px; font-weight: bold; color: #667eea;">{report_data.get('ai_insights', {}).get('analysis_type', 'Statistical').title()}</p>
                </div>
            </div>
        </div>

        <div class="section">
            <h3>Top 5 Defects</h3>
            <table>
                <thead>
                    <tr>
                        <th>Defect Type</th>
                        <th style="text-align: center;">Count</th>
                        <th style="text-align: center;">Percentage</th>
                    </tr>
                </thead>
                <tbody>
                    {defects_html}
                </tbody>
            </table>
        </div>

        <div class="section">
            <h3>AI Root Cause Analysis</h3>
            {causes_html}
        </div>

        <div class="section">
            <h3>AI Recommendations</h3>
            {rec_html}
        </div>

        <div class="section">
            <h3>Attached Reports</h3>
            <ul>
                <li><strong>PDF Report:</strong> Comprehensive analysis with detailed insights and visualizations</li>
                <li><strong>Excel Report:</strong> Raw data and detailed metrics for further analysis</li>
            </ul>
        </div>

        <div style="margin-top: 30px; padding: 20px; background: #f8f9fa; border-radius: 5px;">
            <p><strong>Next Steps:</strong></p>
            <ul>
                <li>Review the attached detailed reports</li>
                <li>Implement high-priority recommendations</li>
                <li>Schedule follow-up analysis for next period</li>
            </ul>
        </div>

        <footer style="margin-top: 30px; padding-top: 20px; border-top: 1px solid #ddd; text-align: center; color: #666;">
            <p>Generated automatically by AI Scrap Analysis System</p>
            <p>VANDEWIELE ROMANIA SRL • {datetime.now().strftime('%B %d, %Y at %H:%M')}</p>
        </footer>
    </div>
</body>
</html>
        """

        return html_body, subject

    def _generate_executive_summary(self, report_data):
        """Genera sommario esecutivo per email"""
        scrap_rate = report_data.get('scrap_data', {}).get('scrap_rate', 0)
        total_scraps = report_data.get('scrap_data', {}).get('total_scraps', 0)

        summary = f"""
AI Scrap Analysis completed for period {report_data.get('period', 'N/A')}

Key Findings:
- Scrap Rate: {scrap_rate:.2f}%
- Total Defects: {total_scraps:,}
- Analysis Type: {report_data.get('ai_insights', {}).get('analysis_type', 'Statistical').title()}

Top recommendations have been generated to improve production quality and reduce defects.
Please review the attached PDF and Excel reports for detailed analysis and actionable insights.
        """

        return summary

    def _get_production_data(self, start_date: str, end_date: str) -> dict:
        """Recupera dati produzione dal DB usando la TUA connessione"""
        query = """
        SELECT
            COUNT(DISTINCT o.OrderNumber) AS TotalOrders,
            COUNT(distinct s.IDBoard) AS TotalBoards
            FROM [Traceability_RS].[dbo].Scannings s
            inner join [Traceability_RS].[dbo].Boards b on b.IDBoard=s.IDBoard
            inner join [Traceability_RS].[dbo].Orders o on o.IDOrder=b.IDOrder
                    WHERE CAST( s.ScanTimeFinish AS DATE) BETWEEN ? AND ?
        """

        # Usa la TUA connessione
        conn = self.db.connection
        cursor = conn.cursor()
        cursor.execute(query, (start_date, end_date))

        row = cursor.fetchone()
        cursor.close()

        return {
            'NrOrders': row.TotalOrders or 0,
            'NrBoards': row.TotalBoards or 0
        }

    def _get_scraps_data(self, start_date: str, end_date: str) -> list:
        """
        Recupera difetti dal DB usando la query corretta ScrapDeclarations

        Args:
            start_date: Data inizio (YYYY-MM-DD)
            end_date: Data fine (YYYY-MM-DD)

        Returns:
            Lista di dizionari con dati difetti
        """
        query = """
        SELECT s.ScrapDeclarationId, 
                                 s.[User] as DECLAREDBY, 
                                 FORMAT(s.DateIn, 'dd/MM/yyyy') as [Date], 
                o.OrderNumber,
                l.labelcod, 
                p.productCode as Product, 
                '' as ProductDescription,
                A.AreaName,
                '' as AreaDescription,
                d.DefectNameRO as Defect,
                d.DefectNameRO as DefectDescription,
                1 as Qty,
                '' as Comments
                          FROM [Traceability_RS].[dbo].ScarpDeclarations S
                              INNER JOIN Traceability_RS.dbo.LabelCodes L 
                          ON l.IDLabelCode = s.IdLabelCode
                              INNER JOIN [Traceability_RS].dbo.Areas A ON a.IDArea = s.IDParentPhase
                              INNER JOIN [Traceability_RS].dbo.defects D ON d.IDDefect = s.ScrapReasonId
                              INNER JOIN [Traceability_RS].dbo.boards B ON l.IDBoard = b.IDBoard
                              INNER JOIN [Traceability_RS].dbo.orders o ON o.idorder = b.IDOrder
                              INNER JOIN traceability_rs.dbo.products P on p.idproduct=o.idproduct
                          WHERE (s.Refuzed IS NULL 
                             OR s.Refuzed = 0)
                            AND CAST (s.DateIn as date) BETWEEN ? 
                            AND ?
                          ORDER BY S.DateIn DESC 
        """

        try:
            # Usa la TUA connessione
            conn = self.db.connection
            cursor = conn.cursor()

            logger.info(f"Esecuzione query ScrapDeclarations: {start_date} -> {end_date}")
            cursor.execute(query, (start_date, end_date))

            scraps = []
            for row in cursor.fetchall():
                scraps.append({
                    'ScrapDeclarationId': row.ScrapDeclarationId,
                    'DeclaredBy': row.DECLAREDBY,
                    'Date': row.Date,
                    'OrderNumber': row.OrderNumber,
                    'LabelCode': row.labelcod,
                    'Product': row.Product,
                    'ProductDescription': row.ProductDescription,
                    'AreaName': row.AreaName,
                    'AreaDescription': row.AreaDescription,
                    'Defect': row.Defect,
                    'DefectDescription': row.DefectDescription,
                    'Qty': row.Qty,
                    'Comments': row.Comments
                })

            cursor.close()
            logger.info(f"Recuperati {len(scraps)} difetti")
            return scraps

        except Exception as e:
            logger.error(f"Errore recupero scraps: {e}", exc_info=True)
            raise

    def _calculate_top_defects(self, scraps_data: list) -> list:
        """Calcola top defects per conteggio"""
        defect_counts = {}
        defect_details = {}

        for scrap in scraps_data:
            defect_name = scrap['Defect']  # ⚠️ Usa 'Defect' dalla query
            area_name = scrap['AreaName']

            # Conteggio per difetto
            if defect_name not in defect_counts:
                defect_counts[defect_name] = 0
                defect_details[defect_name] = {
                    'areas': {},
                    'products': set(),
                    'operators': set()
                }

            defect_counts[defect_name] += 1

            # Conteggio per area
            if area_name not in defect_details[defect_name]['areas']:
                defect_details[defect_name]['areas'][area_name] = 0
            defect_details[defect_name]['areas'][area_name] += 1

            # Raccogli prodotti e operatori
            defect_details[defect_name]['products'].add(scrap['Product'])
            defect_details[defect_name]['operators'].add(scrap['DeclaredBy'])

            # Ordina per conteggio decrescente
        top_defects = []
        for defect_name, count in sorted(defect_counts.items(), key=lambda x: x[1], reverse=True):
            details = defect_details[defect_name]

            # Area più frequente
            top_area = max(details['areas'].items(), key=lambda x: x[1]) if details['areas'] else ('N/A', 0)

            top_defects.append({
                'DefectName': defect_name,
                'Count': count,
                'TopArea': top_area[0],
                'AreaCount': top_area[1],
                'ProductCount': len(details['products']),
                'OperatorCount': len(details['operators'])
            })

        logger.info(f"Top defects calcolati: {len(top_defects)}")
        return top_defects

    def _calculate_statistics(self, production_data: dict, scraps_data: list) -> dict:
        """
        Calcola statistiche avanzate

        Args:
            production_data: Dati produzione
            scraps_data: Lista difetti

        Returns:
            Dizionario con statistiche
        """
        nr_boards = production_data.get('NrBoards', 0)
        scrap_count = len(scraps_data)

        # Scrap rate
        scrap_rate = (scrap_count / nr_boards * 100) if nr_boards > 0 else 0

        # Analisi per area
        area_stats = {}
        for scrap in scraps_data:
            area = scrap['AreaName']
            if area not in area_stats:
                area_stats[area] = {'count': 0, 'defects': set()}
            area_stats[area]['count'] += 1
            area_stats[area]['defects'].add(scrap['Defect'])

        # Analisi per prodotto
        product_stats = {}
        for scrap in scraps_data:
            product = scrap['Product']
            if product not in product_stats:
                product_stats[product] = 0
            product_stats[product] += 1

        # Top 5 aree
        top_areas = sorted(
            [{'area': k, 'count': v['count'], 'unique_defects': len(v['defects'])}
             for k, v in area_stats.items()],
            key=lambda x: x['count'],
            reverse=True
        )[:5]

        # Top 5 prodotti
        top_products = sorted(
            [{'product': k, 'count': v} for k, v in product_stats.items()],
            key=lambda x: x['count'],
            reverse=True
        )[:5]

        logger.info(
            f"Statistiche: Scrap Rate={scrap_rate:.2f}%, Aree={len(area_stats)}, Prodotti={len(product_stats)}")

        return {
            'scrap_rate': scrap_rate,
            'total_scraps': scrap_count,
            'total_boards': nr_boards,
            'area_stats': area_stats,
            'product_stats': product_stats,
            'top_areas': top_areas,
            'top_products': top_products
        }


def main():
    """Entry point"""
    try:
        app = AIScrapAnalysisApp()

        # Analisi ultima settimana
        result = app.run_analysis()

        if result['success']:
            print("\nAnalisi completata con successo!")
            print(f"Report Excel: {result['excel_path']}")
            print(f"Report PDF: {result['pdf_path']}")
        else:
            print(f"\nErrore: {result.get('error')}")
            sys.exit(1)

    except Exception as e:
        logger.error(f"Errore fatale: {e}", exc_info=True)
        sys.exit(1)


if __name__ == "__main__":
    main()