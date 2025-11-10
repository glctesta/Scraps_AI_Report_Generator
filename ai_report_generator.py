import logging
import pandas as pd
from datetime import datetime, timedelta
import os
from pathlib import Path

# Import necessari

from db_connection import DatabaseConnection
from utils import *
import utils
from cryptography.fernet import Fernet
from config_manager import ConfigManager

class AIReportGenerator:
    """Generatore di report AI per analisi scraps con dati da SQL Server"""

    def __init__(self, config_manager=None):
        """
        Inizializza il generatore di report

        Args:
            config_manager: Istanza di ConfigManager (opzionale)
        """
        logging.info("=" * 80)
        logging.info("Inizializzazione AIReportGenerator")
        logging.info("=" * 80)

        # Inizializza ConfigManager
        if config_manager is None:
            self.config_manager = ConfigManager()
        else:
            self.config_manager = config_manager

        # Inizializza DatabaseConnection
        self.db = DatabaseConnection(self.config_manager)

        # Verifica configurazione email
        smtp_config = utils.get_email_recipients('Sys_email_Quality')
        if smtp_config and all([
            smtp_config.get('server'),
            smtp_config.get('port'),
            smtp_config.get('username'),
            smtp_config.get('password')
        ]):
            self.email_enabled = True
            logging.info("✓ Configurazione email attiva")
        else:
            self.email_enabled = False
            logging.info("Nessuna configurazione SMTP fornita - funzionalità email disabilitate")

        logging.info("AIReportGenerator inizializzato con successo")

    def get_connection(self):
        """Crea connessione al database usando DatabaseConnection"""
        try:
            conn = self.db.connect()
            logging.info("Connessione al database stabilita")
            return conn
        except Exception as e:
            logging.error(f"Errore connessione database: {e}")
            raise

    def get_data(self, start_date, end_date):
        """
        Recupera tutti i dati necessari da SQL Server

        Args:
            start_date: Data inizio periodo (formato: 'YYYY-MM-DD')
            end_date: Data fine periodo (formato: 'YYYY-MM-DD')

        Returns:
            tuple: (production_data dict, scrap_data DataFrame)
        """
        conn = self.db.connect()
        cursor = conn.cursor()

        try:
            # Query dati produzione
            production_query = """
            DECLARE @DateStart AS DATE = ?;
            DECLARE @DateFinish AS DATE = ?;
            
            SELECT
                COUNT(DISTINCT o.OrderNumber) AS TotalOrders,
                COUNT(DISTINCT s.IDBoard) AS TotalBoards  
            FROM [Traceability_RS].[dbo].Scannings s
            INNER JOIN [Traceability_RS].[dbo].Boards b ON b.IDBoard = s.IDBoard
            INNER JOIN [Traceability_RS].[dbo].Orders o ON o.IDOrder = b.IDOrder
            WHERE s.ScanTimeFinish BETWEEN @DateStart AND @DateFinish;
            """

            cursor.execute(production_query, start_date, end_date)
            production_result = cursor.fetchone()
            production_data = {
                'NrOrders': production_result[0] if production_result and production_result[0] is not None else 0,
                'NrBoards': production_result[1] if production_result and production_result[1] is not None else 0
            }

            logging.info(f"✓ Dati produzione recuperati: {production_data}")

            # Query dati scraps
            scrap_query = """
            SELECT 
                s.ScrapDeclarationId, 
                s.[User] AS DECLAREDBY, 
                FORMAT(s.DateIn, 'dd/MM/yyyy') AS [Date], 
                o.OrderNumber,
                l.labelcod, 
                p.productCode AS Product, 
                '' AS ProductDescription,
                A.AreaName,
                '' AS AreaDescription,
                d.DefectNameRO AS Defect,
                d.DefectNameRO AS DefectDescription,
                1 AS Qty,
                '' AS Comments
            FROM [Traceability_RS].[dbo].ScarpDeclarations S
            INNER JOIN Traceability_RS.dbo.LabelCodes L ON l.IDLabelCode = s.IdLabelCode
            INNER JOIN [Traceability_RS].dbo.Areas A ON a.IDArea = s.IDParentPhase
            INNER JOIN [Traceability_RS].dbo.defects D ON d.IDDefect = s.ScrapReasonId
            INNER JOIN [Traceability_RS].dbo.boards B ON l.IDBoard = b.IDBoard
            INNER JOIN [Traceability_RS].dbo.orders o ON o.idorder = b.IDOrder
            INNER JOIN traceability_rs.dbo.products P ON p.idproduct = o.idproduct
            WHERE (s.Refuzed IS NULL OR s.Refuzed = 0)
              AND CAST(s.DateIn AS DATE) BETWEEN ? AND ?
            ORDER BY S.DateIn DESC 
            """

            cursor.execute(scrap_query, start_date, end_date)
            columns = [column[0] for column in cursor.description]
            scrap_rows = cursor.fetchall()
            scrap_data = pd.DataFrame.from_records(scrap_rows, columns=columns)

            logging.info(f"✓ Dati scraps recuperati: {len(scrap_data)} record")

            return production_data, scrap_data

        except Exception as e:
            logging.error(f"Errore nel recupero dati: {e}")
            raise
        finally:
            cursor.close()
            conn.close()

    def get_email_recipients(self):
        """Recupera i destinatari email dal database usando utils"""
        try:
            logging.info("📧 Recupero destinatari email...")

            # Usa il context manager per la connessione
            with DatabaseConnection(self.config_manager) as conn:
                if conn is None:
                    logging.error("❌ Connessione non valida")
                    return self._get_fallback_recipients()

                attributes_to_try = ['Sys_email_Quality']

                for attribute in attributes_to_try:
                    try:
                        logging.info(f"🔍 Cerco attributo: {attribute}")
                        recipients = utils.get_email_recipients(conn, attribute)

                        if recipients:
                            logging.info(f"✅ Destinatari trovati: {recipients}")
                            return recipients
                        else:
                            logging.warning(f"⚠️ Nessun destinatario per {attribute}")

                    except Exception as e:
                        logging.error(f"❌ Errore con {attribute}: {e}")
                        continue

                logging.warning("⚠️ Nessun destinatario trovato")
                return self._get_fallback_recipients()

        except Exception as e:
            logging.error(f"❌ Errore generale: {e}")
            import traceback
            logging.error(traceback.format_exc())
            return self._get_fallback_recipients()

    def _get_fallback_recipients(self):
        """Restituisce destinatari di fallback"""
        fallback = ['quality@example.com']  # Modifica con email reale
        logging.info(f"📧 Uso destinatari fallback: {fallback}")
        return fallback

    def generate_complete_report(self, start_date=None, end_date=None, output_dir="output"):
        """
        Genera report completo con dati da SQL Server

        Args:
            start_date: Data inizio (formato 'YYYY-MM-DD'). Default: 7 giorni fa
            end_date: Data fine (formato 'YYYY-MM-DD'). Default: oggi
            output_dir: Directory output per i report

        Returns:
            dict: Risultato con percorsi file generati e statistiche
        """
        try:
            # Date di default se non specificate
            if end_date is None:
                end_date = datetime.now().strftime('%Y-%m-%d')
            if start_date is None:
                start_date = (datetime.now() - timedelta(days=7)).strftime('%Y-%m-%d')

            logging.info(f"📅 Periodo analisi: {start_date} → {end_date}")

            # Crea directory output
            Path(output_dir).mkdir(parents=True, exist_ok=True)

            # 1. Recupera dati da SQL Server
            logging.info("📊 Recupero dati da SQL Server...")
            production_data, scrap_data = self.get_data(start_date, end_date)

            # 2. Verifica dati
            if scrap_data.empty:
                logging.warning("⚠️ Nessun dato scrap trovato nel periodo specificato")
                return {
                    'success': False,
                    'message': 'Nessun dato disponibile per il periodo specificato',
                    'production_data': production_data
                }

            # 3. Analisi dati
            logging.info("🔍 Analisi dati in corso...")
            analysis_results = self._analyze_scraps(scrap_data, production_data)

            # 4. Genera report Excel
            logging.info("📄 Generazione report Excel...")
            excel_path = self._generate_excel_report(
                scrap_data,
                production_data,
                analysis_results,
                output_dir,
                start_date,
                end_date
            )

            # 5. Genera report PDF (opzionale)
            # pdf_path = self._generate_pdf_report(...)

            # 6. Invia email se configurato
            if self.email_enabled:
                logging.info("📧 Invio email...")
                recipients = self.get_email_recipients()
                self._send_report_email(recipients, excel_path, start_date, end_date)

            result = {
                'success': True,
                'excel_path': excel_path,
                'production_data': production_data,
                'scrap_count': len(scrap_data),
                'analysis': analysis_results,
                'period': f"{start_date} to {end_date}"
            }

            logging.info("✅ Report generato con successo")
            return result

        except Exception as e:
            logging.error(f"❌ Errore durante generazione report: {e}")
            import traceback
            logging.error(traceback.format_exc())
            raise

    def _analyze_scraps(self, scrap_data, production_data):
        """Analizza i dati scraps e calcola statistiche"""
        # Implementa analisi personalizzata
        analysis = {
            'total_scraps': len(scrap_data),
            'scrap_rate': (len(scrap_data) / production_data['NrBoards'] * 100) if production_data['NrBoards'] > 0 else 0,
            'top_defects': scrap_data['Defect'].value_counts().head(5).to_dict(),
            'top_areas': scrap_data['AreaName'].value_counts().head(5).to_dict(),
            'top_products': scrap_data['Product'].value_counts().head(5).to_dict()
        }
        return analysis

    def _generate_excel_report(self, scrap_data, production_data, analysis, output_dir, start_date, end_date):
        """Genera report Excel"""
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"Scrap_Report_{start_date}_to_{end_date}_{timestamp}.xlsx"
        filepath = os.path.join(output_dir, filename)

        with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
            # Sheet 1: Raw Data
            scrap_data.to_excel(writer, sheet_name='Scrap Data', index=False)

            # Sheet 2: Summary
            summary_df = pd.DataFrame({
                'Metric': ['Total Orders', 'Total Boards', 'Total Scraps', 'Scrap Rate (%)'],
                'Value': [
                    production_data['NrOrders'],
                    production_data['NrBoards'],
                    analysis['total_scraps'],
                    round(analysis['scrap_rate'], 2)
                ]
            })
            summary_df.to_excel(writer, sheet_name='Summary', index=False)

            # Sheet 3: Top Defects
            top_defects_df = pd.DataFrame(list(analysis['top_defects'].items()),
                                         columns=['Defect', 'Count'])
            top_defects_df.to_excel(writer, sheet_name='Top Defects', index=False)

        logging.info(f"✓ Report Excel salvato: {filepath}")
        return filepath

    def _send_report_email(self, recipients, excel_path, start_date, end_date):
        """Invia report via email usando utils.send_email esistente"""
        try:
            subject = f"Scrap Analysis Report - {start_date} to {end_date}"
            body = f"""
            <html>
            <body>
                <h2>Scrap Analysis Report</h2>
                <p>Period: <strong>{start_date}</strong> to <strong>{end_date}</strong></p>
                <p>Please find the attached report.</p>
            </body>
            </html>
            """

            # USA utils.send_email che già funziona
            from utils import send_email

            send_email(
                recipients=recipients,
                subject=subject,
                body=body,
                is_html=True,
                attachments=[excel_path]
            )

            logging.info(f"✓ Email inviata a: {', '.join(recipients)}")

        except Exception as e:
            logging.error(f"❌ Errore invio email: {e}")
