"""
Modulo per la generazione di report PDF con analisi AI
Utilizza ReportLab per creare PDF professionali con grafici e tabelle
"""

import os
from datetime import datetime
from io import BytesIO
import base64

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, letter
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch, cm
    from reportlab.platypus import (
        SimpleDocTemplate, Table, TableStyle, Paragraph,
        Spacer, PageBreak, Image, KeepTogether
    )
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT, TA_JUSTIFY
    from reportlab.pdfgen import canvas

    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

from logger_config import setup_logger

# Inizializza logger
logger = setup_logger('PDFGenerator')


class PDFReportGenerator:
    """
    Classe per generare report PDF con analisi AI
    """

    def __init__(self, title="Report Analisi AI", author="AI Report System"):
        """
        Inizializza il generatore PDF

        Args:
            title (str): Titolo del documento
            author (str): Autore del documento
        """
        logger.info(f"Inizializzazione PDFReportGenerator - Titolo: {title}")

        if not REPORTLAB_AVAILABLE:
            logger.error("ReportLab non disponibile. Installare con: pip install reportlab")
            raise ImportError("ReportLab non installato")

        self.title = title
        self.author = author
        self.styles = getSampleStyleSheet()
        self._setup_custom_styles()

        logger.info("PDFReportGenerator inizializzato con successo")

    def _setup_custom_styles(self):
        """Configura stili personalizzati per il PDF"""
        logger.debug("Configurazione stili personalizzati PDF")

        # Stile per titolo principale
        self.styles.add(ParagraphStyle(
            name='CustomTitle',
            parent=self.styles['Heading1'],
            fontSize=24,
            textColor=colors.HexColor('#1a5490'),
            spaceAfter=30,
            alignment=TA_CENTER,
            fontName='Helvetica-Bold'
        ))

        # Stile per sottotitoli
        self.styles.add(ParagraphStyle(
            name='CustomHeading2',
            parent=self.styles['Heading2'],
            fontSize=16,
            textColor=colors.HexColor('#2c5aa0'),
            spaceAfter=12,
            spaceBefore=12,
            fontName='Helvetica-Bold'
        ))

        # Stile per paragrafi
        self.styles.add(ParagraphStyle(
            name='CustomBody',
            parent=self.styles['BodyText'],
            fontSize=11,
            alignment=TA_JUSTIFY,
            spaceAfter=10
        ))

        # Stile per evidenziazioni
        self.styles.add(ParagraphStyle(
            name='Highlight',
            parent=self.styles['BodyText'],
            fontSize=12,
            textColor=colors.HexColor('#d9534f'),
            fontName='Helvetica-Bold',
            spaceAfter=10
        ))

        logger.debug("Stili personalizzati configurati")

    def generate_report(self, analysis_data, output_path=None):
        """
        Genera il report PDF completo

        Args:
            analysis_data (dict): Dizionario con i dati dell'analisi
            output_path (str): Percorso del file di output (opzionale)

        Returns:
            bytes: Contenuto del PDF come bytes se output_path è None
            str: Percorso del file se output_path è specificato
        """
        logger.info("Inizio generazione report PDF")
        logger.debug(f"Dati analisi ricevuti: {list(analysis_data.keys())}")

        try:
            # Crea buffer o file
            if output_path:
                logger.info(f"Generazione PDF su file: {output_path}")
                buffer = output_path
            else:
                logger.info("Generazione PDF in memoria (BytesIO)")
                buffer = BytesIO()

            # Crea documento
            doc = SimpleDocTemplate(
                buffer,
                pagesize=A4,
                rightMargin=2 * cm,
                leftMargin=2 * cm,
                topMargin=2 * cm,
                bottomMargin=2 * cm,
                title=self.title,
                author=self.author
            )

            # Costruisci contenuto
            story = []

            # Header
            story.extend(self._create_header(analysis_data))

            # Sommario esecutivo
            if 'executive_summary' in analysis_data:
                story.extend(self._create_executive_summary(analysis_data['executive_summary']))

            # Analisi dettagliata
            if 'detailed_analysis' in analysis_data:
                story.extend(self._create_detailed_analysis(analysis_data['detailed_analysis']))

            # Metriche e KPI
            if 'metrics' in analysis_data:
                story.extend(self._create_metrics_section(analysis_data['metrics']))

            # Raccomandazioni AI
            if 'recommendations' in analysis_data:
                story.extend(self._create_recommendations(analysis_data['recommendations']))

            # Grafici
            if 'charts' in analysis_data:
                story.extend(self._create_charts_section(analysis_data['charts']))

            # Tabelle dati
            if 'tables' in analysis_data:
                story.extend(self._create_tables_section(analysis_data['tables']))

            # Footer
            story.extend(self._create_footer())

            # Genera PDF
            logger.info("Building PDF document...")
            doc.build(story, onFirstPage=self._add_page_number, onLaterPages=self._add_page_number)

            logger.info("Report PDF generato con successo")

            if output_path:
                logger.info(f"PDF salvato in: {output_path}")
                return output_path
            else:
                pdf_bytes = buffer.getvalue()
                buffer.close()
                logger.info(f"PDF generato in memoria: {len(pdf_bytes)} bytes")
                return pdf_bytes

        except Exception as e:
            logger.error(f"Errore durante la generazione del PDF: {str(e)}", exc_info=True)
            raise

    def _create_header(self, data):
        """Crea l'intestazione del report"""
        logger.debug("Creazione header PDF")
        elements = []

        # Titolo
        title = Paragraph(self.title, self.styles['CustomTitle'])
        elements.append(title)
        elements.append(Spacer(1, 0.3 * inch))

        # Informazioni documento
        info_data = [
            ['Data Generazione:', datetime.now().strftime('%d/%m/%Y %H:%M:%S')],
            ['Autore:', self.author],
            ['Versione:', data.get('version', '1.0')],
            ['Tipo Analisi:', data.get('analysis_type', 'Standard')]
        ]

        info_table = Table(info_data, colWidths=[4 * cm, 12 * cm])
        info_table.setStyle(TableStyle([
            ('FONT', (0, 0), (0, -1), 'Helvetica-Bold', 10),
            ('FONT', (1, 0), (1, -1), 'Helvetica', 10),
            ('TEXTCOLOR', (0, 0), (0, -1), colors.HexColor('#555555')),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ]))

        elements.append(info_table)
        elements.append(Spacer(1, 0.5 * inch))

        # Linea separatore
        elements.append(self._create_line())
        elements.append(Spacer(1, 0.3 * inch))

        logger.debug("Header PDF creato")
        return elements

    def _create_executive_summary(self, summary_data):
        """Crea la sezione sommario esecutivo"""
        logger.debug("Creazione executive summary")
        elements = []

        # Titolo sezione
        elements.append(Paragraph("Sommario Esecutivo", self.styles['CustomHeading2']))
        elements.append(Spacer(1, 0.2 * inch))

        # Contenuto
        if isinstance(summary_data, str):
            elements.append(Paragraph(summary_data, self.styles['CustomBody']))
        elif isinstance(summary_data, dict):
            for key, value in summary_data.items():
                elements.append(Paragraph(f"<b>{key}:</b> {value}", self.styles['CustomBody']))

        elements.append(Spacer(1, 0.3 * inch))
        logger.debug("Executive summary creato")
        return elements

    def _create_detailed_analysis(self, analysis_data):
        """Crea la sezione analisi dettagliata"""
        logger.debug("Creazione analisi dettagliata")
        elements = []

        elements.append(Paragraph("Analisi Dettagliata", self.styles['CustomHeading2']))
        elements.append(Spacer(1, 0.2 * inch))

        if isinstance(analysis_data, str):
            elements.append(Paragraph(analysis_data, self.styles['CustomBody']))
        elif isinstance(analysis_data, list):
            for item in analysis_data:
                elements.append(Paragraph(f"• {item}", self.styles['CustomBody']))
        elif isinstance(analysis_data, dict):
            for section, content in analysis_data.items():
                elements.append(Paragraph(f"<b>{section}</b>", self.styles['Heading3']))
                elements.append(Paragraph(str(content), self.styles['CustomBody']))
                elements.append(Spacer(1, 0.1 * inch))

        elements.append(Spacer(1, 0.3 * inch))
        logger.debug("Analisi dettagliata creata")
        return elements

    def _create_metrics_section(self, metrics_data):
        """Crea la sezione metriche e KPI"""
        logger.debug("Creazione sezione metriche")
        elements = []

        elements.append(Paragraph("Metriche e KPI", self.styles['CustomHeading2']))
        elements.append(Spacer(1, 0.2 * inch))

        # Crea tabella metriche
        table_data = [['Metrica', 'Valore', 'Target', 'Status']]

        for metric in metrics_data:
            status = '✓' if metric.get('achieved', False) else '✗'
            table_data.append([
                metric.get('name', 'N/A'),
                str(metric.get('value', 'N/A')),
                str(metric.get('target', 'N/A')),
                status
            ])

        metrics_table = Table(table_data, colWidths=[5 * cm, 3 * cm, 3 * cm, 2 * cm])
        metrics_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1a5490')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 10),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey]),
        ]))

        elements.append(metrics_table)
        elements.append(Spacer(1, 0.3 * inch))

        logger.debug(f"Sezione metriche creata con {len(metrics_data)} metriche")
        return elements

    def _create_recommendations(self, recommendations_data):
        """Crea la sezione raccomandazioni AI"""
        logger.debug("Creazione sezione raccomandazioni")
        elements = []

        elements.append(Paragraph("Raccomandazioni AI", self.styles['CustomHeading2']))
        elements.append(Spacer(1, 0.2 * inch))

        if isinstance(recommendations_data, list):
            for i, rec in enumerate(recommendations_data, 1):
                priority = rec.get('priority', 'Media')
                color = self._get_priority_color(priority)

                rec_text = f"""
                <b>Raccomandazione {i}</b> [Priorità: <font color="{color}">{priority}</font>]<br/>
                {rec.get('description', 'N/A')}<br/>
                <i>Impatto atteso: {rec.get('impact', 'N/A')}</i>
                """
                elements.append(Paragraph(rec_text, self.styles['CustomBody']))
                elements.append(Spacer(1, 0.15 * inch))

        logger.debug(f"Sezione raccomandazioni creata con {len(recommendations_data)} items")
        return elements

    def _create_charts_section(self, charts_data):
        """Crea la sezione grafici"""
        logger.debug("Creazione sezione grafici")
        elements = []

        elements.append(PageBreak())
        elements.append(Paragraph("Analisi Grafiche", self.styles['CustomHeading2']))
        elements.append(Spacer(1, 0.2 * inch))

        for chart in charts_data:
            try:
                # Titolo grafico
                elements.append(Paragraph(chart.get('title', 'Grafico'), self.styles['Heading3']))

                # Immagine grafico (se presente come path o base64)
                if 'image_path' in chart:
                    img = Image(chart['image_path'], width=6 * inch, height=4 * inch)
                    elements.append(img)
                elif 'image_base64' in chart:
                    # Decodifica base64 e crea immagine
                    img_data = base64.b64decode(chart['image_base64'])
                    img_buffer = BytesIO(img_data)
                    img = Image(img_buffer, width=6 * inch, height=4 * inch)
                    elements.append(img)

                # Descrizione
                if 'description' in chart:
                    elements.append(Spacer(1, 0.1 * inch))
                    elements.append(Paragraph(chart['description'], self.styles['CustomBody']))

                elements.append(Spacer(1, 0.3 * inch))

            except Exception as e:
                logger.warning(f"Errore nel caricamento grafico: {str(e)}")
                elements.append(Paragraph(f"[Grafico non disponibile: {chart.get('title', 'N/A')}]",
                                          self.styles['CustomBody']))

        logger.debug(f"Sezione grafici creata con {len(charts_data)} grafici")
        return elements

    def _create_tables_section(self, tables_data):
        """Crea la sezione tabelle dati"""
        logger.debug("Creazione sezione tabelle")
        elements = []

        elements.append(PageBreak())
        elements.append(Paragraph("Dati Dettagliati", self.styles['CustomHeading2']))
        elements.append(Spacer(1, 0.2 * inch))

        for table_info in tables_data:
            # Titolo tabella
            elements.append(Paragraph(table_info.get('title', 'Tabella'), self.styles['Heading3']))
            elements.append(Spacer(1, 0.1 * inch))

            # Dati tabella
            data = table_info.get('data', [])
            if data:
                # Calcola larghezza colonne
                num_cols = len(data[0]) if data else 0
                col_width = (17 * cm) / num_cols if num_cols > 0 else 3 * cm

                table = Table(data, colWidths=[col_width] * num_cols)
                table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 10),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                    ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                    ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                    ('FONTSIZE', (0, 1), (-1, -1), 8),
                    ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey]),
                ]))

                elements.append(table)

            elements.append(Spacer(1, 0.3 * inch))

        logger.debug(f"Sezione tabelle creata con {len(tables_data)} tabelle")
        return elements

    def _create_footer(self):
        """Crea il footer del documento"""
        logger.debug("Creazione footer")
        elements = []

        elements.append(Spacer(1, 0.5 * inch))
        elements.append(self._create_line())
        elements.append(Spacer(1, 0.1 * inch))

        footer_text = f"""
        <i>Report generato automaticamente dal sistema AI Report Generator<br/>
        © {datetime.now().year} - Tutti i diritti riservati</i>
        """
        elements.append(Paragraph(footer_text, self.styles['Normal']))

        logger.debug("Footer creato")
        return elements

    def _create_line(self):
        """Crea una linea separatrice"""
        from reportlab.platypus import HRFlowable
        return HRFlowable(width="100%", thickness=1, color=colors.grey, spaceAfter=0.1 * inch)

    def _get_priority_color(self, priority):
        """Restituisce il colore in base alla priorità"""
        colors_map = {
            'Alta': '#d9534f',
            'Media': '#f0ad4e',
            'Bassa': '#5cb85c'
        }
        return colors_map.get(priority, '#777777')

    def _add_page_number(self, canvas, doc):
        """Aggiunge il numero di pagina"""
        page_num = canvas.getPageNumber()
        text = f"Pagina {page_num}"
        canvas.saveState()
        canvas.setFont('Helvetica', 9)
        canvas.drawRightString(19 * cm, 1 * cm, text)
        canvas.restoreState()


# Funzione di utilità per uso rapido
def generate_pdf_report(analysis_data, output_path=None, title="Report Analisi AI"):
    """
    Funzione di utilità per generare rapidamente un report PDF

    Args:
        analysis_data (dict): Dati dell'analisi
        output_path (str): Percorso output (opzionale)
        title (str): Titolo del report

    Returns:
        bytes o str: Contenuto PDF o percorso file
    """
    logger.info(f"Generazione rapida PDF: {title}")
    try:
        generator = PDFReportGenerator(title=title)
        result = generator.generate_report(analysis_data, output_path)
        logger.info("PDF generato con successo tramite funzione di utilità")
        return result
    except Exception as e:
        logger.error(f"Errore nella generazione rapida PDF: {str(e)}", exc_info=True)
        raise


if __name__ == "__main__":
    # Test del modulo
    logger.info("Test PDFGenerator in esecuzione...")

    test_data = {
        'version': '1.0',
        'analysis_type': 'Test Report',
        'executive_summary': 'Questo è un report di test per verificare la generazione PDF.',
        'detailed_analysis': {
            'Sezione 1': 'Contenuto della prima sezione di analisi',
            'Sezione 2': 'Contenuto della seconda sezione di analisi'
        },
        'metrics': [
            {'name': 'Efficienza', 'value': '95%', 'target': '90%', 'achieved': True},
            {'name': 'Qualità', 'value': '88%', 'target': '95%', 'achieved': False},
        ],
        'recommendations': [
            {'priority': 'Alta', 'description': 'Migliorare il processo X', 'impact': 'Alto'},
            {'priority': 'Media', 'description': 'Ottimizzare il sistema Y', 'impact': 'Medio'},
        ],
        'tables': [
            {
                'title': 'Dati Produzione',
                'data': [
                    ['Prodotto', 'Quantità', 'Difetti'],
                    ['Prodotto A', '1000', '5'],
                    ['Prodotto B', '1500', '8'],
                ]
            }
        ]
    }

    try:
        pdf_bytes = generate_pdf_report(test_data, title="Test Report PDF")
        logger.info(f"Test completato: PDF generato ({len(pdf_bytes)} bytes)")

        # Salva su file per test
        with open('test_report.pdf', 'wb') as f:
            f.write(pdf_bytes)
        logger.info("File test_report.pdf salvato con successo")

    except Exception as e:
        logger.error(f"Test fallito: {str(e)}", exc_info=True)
