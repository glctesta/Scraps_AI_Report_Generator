"""
Modulo per la generazione di report Excel con analisi AI
Utilizza openpyxl per creare file Excel con formattazione avanzata
"""

import os
from datetime import datetime
from io import BytesIO

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from openpyxl.chart import BarChart, PieChart, LineChart, Reference
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.table import Table, TableStyleInfo

    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

from logger_config import setup_logger

# Inizializza logger
logger = setup_logger('ExcelGenerator')


class ExcelReportGenerator:
    """
    Classe per generare report Excel con analisi AI
    """

    def __init__(self, title="Report Analisi AI"):
        """
        Inizializza il generatore Excel

        Args:
            title (str): Titolo del documento
        """
        logger.info(f"Inizializzazione ExcelReportGenerator - Titolo: {title}")

        if not OPENPYXL_AVAILABLE:
            logger.error("openpyxl non disponibile. Installare con: pip install openpyxl")
            raise ImportError("openpyxl non installato")

        self.title = title
        self.workbook = Workbook()

        # Rimuovi il foglio di default
        if 'Sheet' in self.workbook.sheetnames:
            del self.workbook['Sheet']

        # Definisci stili
        self._setup_styles()

        logger.info("ExcelReportGenerator inizializzato con successo")

    def _setup_styles(self):
        """Definisce gli stili per il documento Excel"""
        logger.debug("Configurazione stili Excel")

        # Stile header
        self.header_font = Font(name='Calibri', size=14, bold=True, color='FFFFFF')
        self.header_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
        self.header_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

        # Stile titolo
        self.title_font = Font(name='Calibri', size=18, bold=True, color='1F4E78')
        self.title_alignment = Alignment(horizontal='center', vertical='center')

        # Stile sottotitolo
        self.subtitle_font = Font(name='Calibri', size=12, bold=True, color='44546A')

        # Stile dati
        self.data_font = Font(name='Calibri', size=11)
        self.data_alignment = Alignment(horizontal='left', vertical='center')

        # Bordi
        thin_border = Side(border_style="thin", color="000000")
        self.border = Border(left=thin_border, right=thin_border, top=thin_border, bottom=thin_border)

        # Riempimenti
        self.fill_light = PatternFill(start_color='E7E6E6', end_color='E7E6E6', fill_type='solid')
        self.fill_success = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
        self.fill_warning = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')
        self.fill_error = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')

        logger.debug("Stili Excel configurati")

    def generate_report(self, analysis_data, output_path=None):
        """
        Genera il report Excel completo

        Args:
            analysis_data (dict): Dizionario con i dati dell'analisi
            output_path (str): Percorso del file di output (opzionale)

        Returns:
            bytes: Contenuto Excel come bytes se output_path è None
            str: Percorso del file se output_path è specificato
        """
        logger.info("Inizio generazione report Excel")
        logger.debug(f"Dati analisi ricevuti: {list(analysis_data.keys())}")

        try:
            # Crea fogli
            self._create_summary_sheet(analysis_data)

            if 'metrics' in analysis_data:
                self._create_metrics_sheet(analysis_data['metrics'])

            if 'detailed_analysis' in analysis_data:
                self._create_analysis_sheet(analysis_data['detailed_analysis'])

            if 'recommendations' in analysis_data:
                self._create_recommendations_sheet(analysis_data['recommendations'])

            if 'raw_data' in analysis_data:
                self._create_raw_data_sheet(analysis_data['raw_data'])

            # Salva o restituisci bytes
            if output_path:
                logger.info(f"Salvataggio Excel su file: {output_path}")
                self.workbook.save(output_path)
                logger.info(f"Excel salvato in: {output_path}")
                return output_path
            else:
                logger.info("Generazione Excel in memoria (BytesIO)")
                buffer = BytesIO()
                self.workbook.save(buffer)
                excel_bytes = buffer.getvalue()
                buffer.close()
                logger.info(f"Excel generato in memoria: {len(excel_bytes)} bytes")
                return excel_bytes

        except Exception as e:
            logger.error(f"Errore durante la generazione Excel: {str(e)}", exc_info=True)
            raise

    def _create_summary_sheet(self, data):
        """Crea il foglio di riepilogo"""
        logger.debug("Creazione foglio Summary")

        ws = self.workbook.create_sheet("Summary", 0)

        # Titolo
        ws['A1'] = self.title
        ws['A1'].font = self.title_font
        ws['A1'].alignment = self.title_alignment
        ws.merge_cells('A1:D1')
        ws.row_dimensions[1].height = 30

        # Informazioni documento
        row = 3
        ws[f'A{row}'] = "Data Generazione:"
        ws[f'B{row}'] = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
        ws[f'A{row}'].font = self.subtitle_font

        row += 1
        ws[f'A{row}'] = "Versione:"
        ws[f'B{row}'] = data.get('version', '1.0')
        ws[f'A{row}'].font = self.subtitle_font

        row += 1
        ws[f'A{row}'] = "Tipo Analisi:"
        ws[f'B{row}'] = data.get('analysis_type', 'Standard')
        ws[f'A{row}'].font = self.subtitle_font

        # Sommario esecutivo
        row += 2
        ws[f'A{row}'] = "Sommario Esecutivo"
        ws[f'A{row}'].font = Font(name='Calibri', size=14, bold=True, color='1F4E78')

        row += 1
        if 'executive_summary' in data:
            summary = data['executive_summary']
            if isinstance(summary, str):
                ws[f'A{row}'] = summary
                ws.merge_cells(f'A{row}:D{row}')
                ws[f'A{row}'].alignment = Alignment(wrap_text=True, vertical='top')
            elif isinstance(summary, dict):
                for key, value in summary.items():
                    ws[f'A{row}'] = key
                    ws[f'B{row}'] = str(value)
                    ws[f'A{row}'].font = self.subtitle_font
                    row += 1

        # Statistiche chiave
        if 'key_stats' in data:
            row += 2
            ws[f'A{row}'] = "Statistiche Chiave"
            ws[f'A{row}'].font = Font(name='Calibri', size=14, bold=True, color='1F4E78')

            row += 1
            stats = data['key_stats']
            for stat_name, stat_value in stats.items():
                ws[f'A{row}'] = stat_name
                ws[f'B{row}'] = stat_value
                ws[f'A{row}'].font = self.subtitle_font
                row += 1

        # Formattazione colonne
        ws.column_dimensions['A'].width = 25
        ws.column_dimensions['B'].width = 40
        ws.column_dimensions['C'].width = 20
        ws.column_dimensions['D'].width = 20

        logger.debug("Foglio Summary creato")

    def _create_metrics_sheet(self, metrics_data):
        """Crea il foglio metriche e KPI"""
        logger.debug("Creazione foglio Metrics")

        ws = self.workbook.create_sheet("Metrics & KPI")

        # Titolo
        ws['A1'] = "Metriche e KPI"
        ws['A1'].font = self.title_font
        ws['A1'].alignment = self.title_alignment
        ws.merge_cells('A1:E1')
        ws.row_dimensions[1].height = 30

        # Header tabella
        headers = ['Metrica', 'Valore Attuale', 'Target', 'Differenza', 'Status']
        row = 3
        for col, header in enumerate(headers, start=1):
            cell = ws.cell(row=row, column=col)
            cell.value = header
            cell.font = self.header_font
            cell.fill = self.header_fill
            cell.alignment = self.header_alignment
            cell.border = self.border

        # Dati
        row += 1
        for metric in metrics_data:
            name = metric.get('name', 'N/A')
            value = metric.get('value', 0)
            target = metric.get('target', 0)

            # Calcola differenza
            try:
                if isinstance(value, str):
                    value_num = float(value.replace('%', ''))
                else:
                    value_num = float(value)

                if isinstance(target, str):
                    target_num = float(target.replace('%', ''))
                else:
                    target_num = float(target)

                diff = value_num - target_num
                achieved = diff >= 0
            except:
                diff = 'N/A'
                achieved = False

            # Scrivi dati
            ws.cell(row=row, column=1).value = name
            ws.cell(row=row, column=2).value = value
            ws.cell(row=row, column=3).value = target
            ws.cell(row=row, column=4).value = diff if diff != 'N/A' else 'N/A'
            ws.cell(row=row, column=5).value = '✓ Raggiunto' if achieved else '✗ Non Raggiunto'

            # Formattazione
            for col in range(1, 6):
                cell = ws.cell(row=row, column=col)
                cell.border = self.border
                cell.alignment = Alignment(horizontal='center', vertical='center')

                # Colora status
                if col == 5:
                    cell.fill = self.fill_success if achieved else self.fill_error

            row += 1

        # Aggiungi grafico
        if len(metrics_data) > 0:
            self._add_metrics_chart(ws, len(metrics_data) + 3)

        # Formattazione colonne
        ws.column_dimensions['A'].width = 30
        ws.column_dimensions['B'].width = 15
        ws.column_dimensions['C'].width = 15
        ws.column_dimensions['D'].width = 15
        ws.column_dimensions['E'].width = 20

        logger.debug(f"Foglio Metrics creato con {len(metrics_data)} metriche")

    def _add_metrics_chart(self, ws, data_rows):
        """Aggiunge un grafico a barre per le metriche"""
        logger.debug("Aggiunta grafico metriche")

        try:
            chart = BarChart()
            chart.title = "Confronto Metriche vs Target"
            chart.style = 10
            chart.y_axis.title = 'Valore'
            chart.x_axis.title = 'Metriche'

            # Dati
            data = Reference(ws, min_col=2, min_row=3, max_row=data_rows, max_col=3)
            cats = Reference(ws, min_col=1, min_row=4, max_row=data_rows)

            chart.add_data(data, titles_from_data=True)
            chart.set_categories(cats)

            # Posiziona grafico
            ws.add_chart(chart, f"G3")

            logger.debug("Grafico metriche aggiunto")
        except Exception as e:
            logger.warning(f"Impossibile aggiungere grafico metriche: {str(e)}")

    def _create_analysis_sheet(self, analysis_data):
        """Crea il foglio analisi dettagliata"""
        logger.debug("Creazione foglio Detailed Analysis")

        ws = self.workbook.create_sheet("Detailed Analysis")

        # Titolo
        ws['A1'] = "Analisi Dettagliata"
        ws['A1'].font = self.title_font
        ws['A1'].alignment = self.title_alignment
        ws.merge_cells('A1:C1')
        ws.row_dimensions[1].height = 30

        row = 3

        if isinstance(analysis_data, dict):
            for section, content in analysis_data.items():
                # Titolo sezione
                ws[f'A{row}'] = section
                ws[f'A{row}'].font = Font(name='Calibri', size=12, bold=True, color='1F4E78')
                ws[f'A{row}'].fill = self.fill_light
                ws.merge_cells(f'A{row}:C{row}')
                row += 1

                # Contenuto
                if isinstance(content, list):
                    for item in content:
                        ws[f'A{row}'] = f"• {item}"
                        ws.merge_cells(f'A{row}:C{row}')
                        ws[f'A{row}'].alignment = Alignment(wrap_text=True, vertical='top')
                        row += 1
                else:
                    ws[f'A{row}'] = str(content)
                    ws.merge_cells(f'A{row}:C{row}')
                    ws[f'A{row}'].alignment = Alignment(wrap_text=True, vertical='top')
                    row += 1

                row += 1  # Spazio tra sezioni

        # Formattazione colonne
        ws.column_dimensions['A'].width = 80

        logger.debug("Foglio Detailed Analysis creato")

    def _create_recommendations_sheet(self, recommendations_data):
        """Crea il foglio raccomandazioni"""
        logger.debug("Creazione foglio Recommendations")

        ws = self.workbook.create_sheet("Recommendations")

        # Titolo
        ws['A1'] = "Raccomandazioni AI"
        ws['A1'].font = self.title_font
        ws['A1'].alignment = self.title_alignment
        ws.merge_cells('A1:D1')
        ws.row_dimensions[1].height = 30

        # Header
        headers = ['#', 'Priorità', 'Descrizione', 'Impatto Atteso']
        row = 3
        for col, header in enumerate(headers, start=1):
            cell = ws.cell(row=row, column=col)
            cell.value = header
            cell.font = self.header_font
            cell.fill = self.header_fill
            cell.alignment = self.header_alignment
            cell.border = self.border

        # Dati
        row += 1
        for i, rec in enumerate(recommendations_data, start=1):
            priority = rec.get('priority', 'Media')
            description = rec.get('description', 'N/A')
            impact = rec.get('impact', 'N/A')

            ws.cell(row=row, column=1).value = i
            ws.cell(row=row, column=2).value = priority
            ws.cell(row=row, column=3).value = description
            ws.cell(row=row, column=4).value = impact

            # Formattazione
            for col in range(1, 5):
                cell = ws.cell(row=row, column=col)
                cell.border = self.border
                cell.alignment = Alignment(wrap_text=True, vertical='top')

            # Colora priorità
            priority_cell = ws.cell(row=row, column=2)
            if priority == 'Alta':
                priority_cell.fill = self.fill_error
                priority_cell.font = Font(bold=True, color='C00000')
            elif priority == 'Media':
                priority_cell.fill = self.fill_warning
                priority_cell.font = Font(bold=True, color='C65911')
            else:
                priority_cell.fill = self.fill_success
                priority_cell.font = Font(bold=True, color='006100')

            row += 1

        # Formattazione colonne
        ws.column_dimensions['A'].width = 5
        ws.column_dimensions['B'].width = 15
        ws.column_dimensions['C'].width = 50
        ws.column_dimensions['D'].width = 20

        logger.debug(f"Foglio Recommendations creato con {len(recommendations_data)} raccomandazioni")

    def _create_raw_data_sheet(self, raw_data):
        """Crea il foglio con dati grezzi"""
        logger.debug("Creazione foglio Raw Data")

        ws = self.workbook.create_sheet("Raw Data")

        # Titolo
        ws['A1'] = "Dati Grezzi"
        ws['A1'].font = self.title_font
        ws['A1'].alignment = self.title_alignment

        row = 3

        if isinstance(raw_data, list) and len(raw_data) > 0:
            # Assumi che il primo elemento contenga le chiavi
            if isinstance(raw_data[0], dict):
                # Header
                headers = list(raw_data[0].keys())
                for col, header in enumerate(headers, start=1):
                    cell = ws.cell(row=row, column=col)
                    cell.value = header
                    cell.font = self.header_font
                    cell.fill = self.header_fill
                    cell.alignment = self.header_alignment
                    cell.border = self.border

                # Dati
                row += 1
                for data_row in raw_data:
                    for col, header in enumerate(headers, start=1):
                        cell = ws.cell(row=row, column=col)
                        cell.value = data_row.get(header, '')
                        cell.border = self.border
                    row += 1

                # Applica tabella Excel
                tab = Table(displayName="RawDataTable", ref=f"A3:{get_column_letter(len(headers))}{row - 1}")
                style = TableStyleInfo(
                    name="TableStyleMedium9",
                    showFirstColumn=False,
                    showLastColumn=False,
                    showRowStripes=True,
                    showColumnStripes=False
                )
                tab.tableStyleInfo = style
                ws.add_table(tab)

                # Auto-fit colonne
                for col in range(1, len(headers) + 1):
                    ws.column_dimensions[get_column_letter(col)].width = 15

        logger.debug("Foglio Raw Data creato")


# Funzione di utilità per uso rapido
def generate_excel_report(analysis_data, output_path=None, title="Report Analisi AI"):
    """
    Funzione di utilità per generare rapidamente un report Excel

    Args:
        analysis_data (dict): Dati dell'analisi
        output_path (str): Percorso output (opzionale)
        title (str): Titolo del report

    Returns:
        bytes o str: Contenuto Excel o percorso file
    """
    logger.info(f"Generazione rapida Excel: {title}")
    try:
        generator = ExcelReportGenerator(title=title)
        result = generator.generate_report(analysis_data, output_path)
        logger.info("Excel generato con successo tramite funzione di utilità")
        return result
    except Exception as e:
        logger.error(f"Errore nella generazione rapida Excel: {str(e)}", exc_info=True)
        raise


if __name__ == "__main__":
    # Test del modulo
    logger.info("Test ExcelGenerator in esecuzione...")

    test_data = {
        'version': '1.0',
        'analysis_type': 'Test Report',
        'executive_summary': 'Questo è un report di test per verificare la generazione Excel.',
        'key_stats': {
            'Totale Analisi': 150,
            'Successi': 142,
            'Fallimenti': 8
        },
        'metrics': [
            {'name': 'Efficienza', 'value': '95%', 'target': '90%'},
            {'name': 'Qualità', 'value': '88%', 'target': '95%'},
            {'name': 'Tempo Ciclo', 'value': '45 min', 'target': '50 min'},
        ],
        'detailed_analysis': {
            'Sezione 1': ['Punto 1', 'Punto 2', 'Punto 3'],
            'Sezione 2': 'Analisi dettagliata della sezione 2'
        },
        'recommendations': [
            {'priority': 'Alta', 'description': 'Migliorare il processo X', 'impact': 'Alto'},
            {'priority': 'Media', 'description': 'Ottimizzare il sistema Y', 'impact': 'Medio'},
            {'priority': 'Bassa', 'description': 'Aggiornare documentazione', 'impact': 'Basso'},
        ],
        'raw_data': [
            {'Prodotto': 'A', 'Quantità': 1000, 'Difetti': 5, 'FPY': '99.5%'},
            {'Prodotto': 'B', 'Quantità': 1500, 'Difetti': 8, 'FPY': '99.47%'},
            {'Prodotto': 'C', 'Quantità': 800, 'Difetti': 3, 'FPY': '99.63%'},
        ]
    }

    try:
        excel_bytes = generate_excel_report(test_data, title="Test Report Excel")
        logger.info(f"Test completato: Excel generato ({len(excel_bytes)} bytes)")

        # Salva su file per test
        with open('test_report.xlsx', 'wb') as f:
            f.write(excel_bytes)
        logger.info("File test_report.xlsx salvato con successo")

    except Exception as e:
        logger.error(f"Test fallito: {str(e)}", exc_info=True)
