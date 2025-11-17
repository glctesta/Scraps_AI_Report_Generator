"""
Module for generating PDF reports with AI analysis.
Uses ReportLab to create professional PDFs with charts and tables.
"""

from datetime import datetime
from io import BytesIO
import base64
from pathlib import Path

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.platypus import (
        SimpleDocTemplate, Table, TableStyle, Paragraph,
        Spacer, PageBreak, Image, HRFlowable
    )
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY

    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

from logger_config import setup_logger

# Initialize logger
logger = setup_logger('PDFGenerator')


class PDFReportGenerator:
    """
    Class to generate PDF reports with AI analysis.
    """

    def __init__(self, title="AI Analysis Report", author="AI Report System"):
        """
        Initializes the PDF generator.

        Args:
            title (str): The title of the document.
            author (str): The author of the document.
        """
        logger.info(f"Initializing PDFReportGenerator - Title: {title}")

        if not REPORTLAB_AVAILABLE:
            logger.error("ReportLab not available. Install with: pip install reportlab")
            raise ImportError("ReportLab not installed")

        self.title = title
        self.author = author
        self.styles = getSampleStyleSheet()
        self._setup_custom_styles()

        logger.info("PDFReportGenerator initialized successfully")

    def _setup_custom_styles(self):
        """Configures custom styles for the PDF."""
        logger.debug("Configuring custom PDF styles")

        self.styles.add(ParagraphStyle(
            name='CustomTitle', parent=self.styles['Heading1'], fontSize=24,
            textColor=colors.HexColor('#1a5490'), spaceAfter=20, alignment=TA_CENTER,
            fontName='Helvetica-Bold'
        ))
        self.styles.add(ParagraphStyle(
            name='CustomHeading2', parent=self.styles['Heading2'], fontSize=16,
            textColor=colors.HexColor('#2c5aa0'), spaceAfter=12, spaceBefore=12,
            fontName='Helvetica-Bold'
        ))
        self.styles.add(ParagraphStyle(
            name='CustomBody', parent=self.styles['BodyText'], fontSize=11,
            alignment=TA_JUSTIFY, spaceAfter=10, leading=14
        ))
        self.styles.add(ParagraphStyle(
            name='Highlight', parent=self.styles['BodyText'], fontSize=12,
            textColor=colors.HexColor('#d9534f'), fontName='Helvetica-Bold', spaceAfter=10
        ))
        self.styles.add(ParagraphStyle(
            name='FooterStyle', parent=self.styles['Normal'], fontSize=9,
            alignment=TA_CENTER, textColor=colors.grey
        ))

        logger.debug("Custom styles configured")

    def generate_report(self, analysis_data, output_path=None):
        """
        Generates the generic, complete PDF report.

        Args:
            analysis_data (dict): Dictionary with the analysis data.
            output_path (str): The output file path (optional).

        Returns:
            str: The path to the file if output_path is specified.
        """
        logger.info("Starting PDF report generation")
        if not output_path:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            output_path = Path('reports') / f"Analysis_Report_{timestamp}.pdf"

        doc = SimpleDocTemplate(str(output_path), pagesize=A4, rightMargin=2 * cm,
                                leftMargin=2 * cm, topMargin=2 * cm, bottomMargin=2 * cm)

        story = self._build_story(analysis_data)

        doc.build(story, onFirstPage=self._add_page_number, onLaterPages=self._add_page_number)
        logger.info(f"PDF report generated successfully: {output_path}")
        return str(output_path)

    def _build_story(self, data):
        """Builds the sequence of elements (story) for the PDF."""
        story = []
        story.extend(self._create_header(data))
        if data.get('executive_summary'):
            story.extend(self._create_executive_summary(data['executive_summary']))
        if data.get('root_causes'):
            story.extend(self._create_root_causes_section(data['root_causes']))
        if data.get('recommendations'):
            story.extend(self._create_recommendations(data['recommendations']))
        if data.get('charts'):
            story.extend(self._create_charts_section(data['charts']))
        if data.get('tables'):
            story.extend(self._create_tables_section(data['tables']))
        story.extend(self._create_footer())
        return story

    def _create_header(self, data):
        """Creates the report header."""
        return [
            Paragraph(self.title, self.styles['CustomTitle']),
            Spacer(1, 0.5 * cm),
        ]

    def _create_executive_summary(self, summary_text):
        """Creates the executive summary section."""
        return [
            Paragraph("Executive Summary", self.styles['CustomHeading2']),
            Paragraph(summary_text, self.styles['CustomBody']),
            Spacer(1, 0.5 * cm),
        ]

    def _create_root_causes_section(self, root_causes):
        """Creates the root causes section."""
        elements = [Paragraph("Root Cause Analysis", self.styles['CustomHeading2'])]
        for cause in root_causes:
            cause_text = f"<b>{cause.get('problem_area', cause.get('category', 'General'))}:</b> {cause.get('cause_description', cause.get('cause', 'N/A'))}"
            elements.append(Paragraph(cause_text, self.styles['CustomBody']))
        elements.append(Spacer(1, 0.5 * cm))
        return elements

    def _create_recommendations(self, recommendations_data):
        """Creates the AI recommendations section."""
        elements = [Paragraph("AI Recommendations", self.styles['CustomHeading2'])]
        for i, rec in enumerate(recommendations_data, 1):
            priority = rec.get('priority', 'Medium').title()
            color = self._get_priority_color(priority)
            rec_text = f"<b>{i}. {rec.get('title', 'Recommendation')}</b> [Priority: <font color='{color}'>{priority}</font>]<br/>{rec.get('description', 'N/A')}"
            elements.append(Paragraph(rec_text, self.styles['CustomBody']))
            elements.append(Spacer(1, 0.2 * cm))
        elements.append(Spacer(1, 0.5 * cm))
        return elements

    def _create_charts_section(self, charts_data):
        """Creates the charts section."""
        elements = [PageBreak(), Paragraph("Graphical Analysis", self.styles['CustomHeading2'])]
        for chart in charts_data:
            try:
                if 'image_base64' in chart:
                    img_data = base64.b64decode(chart['image_base64'])
                    img = Image(BytesIO(img_data), width=16 * cm, height=10 * cm)
                    elements.append(img)
                    elements.append(Spacer(1, 0.5 * cm))
            except Exception as e:
                logger.warning(f"Error loading chart: {e}")
        return elements

    def _create_tables_section(self, tables_data):
        """Creates the detailed data tables section."""
        elements = [PageBreak(), Paragraph("Detailed Data", self.styles['CustomHeading2'])]
        for table_info in tables_data:
            elements.append(Paragraph(table_info.get('title', 'Table'), self.styles['Heading3']))
            data = table_info.get('data', [[]])
            if data:
                table = Table(data, hAlign='LEFT')
                table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#4F81BD')),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black),
                    ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.lightgrey, colors.white]),
                ]))
                elements.append(table)
                elements.append(Spacer(1, 0.5 * cm))
        return elements

    def _create_footer(self):
        """Creates the document footer."""
        return [
            Spacer(1, 1 * cm),
            HRFlowable(width="100%", thickness=0.5, color=colors.grey),
            Paragraph(
                f"Report automatically generated on {datetime.now().strftime('%d/%m/%Y at %H:%M')} © {datetime.now().year}",
                self.styles['FooterStyle'])
        ]

    def _get_priority_color(self, priority):
        """Returns the color based on priority string (handles English and Italian)."""
        return {
            'High': '#d9534f', 'Alta': '#d9534f',
            'Medium': '#f0ad4e', 'Media': '#f0ad4e',
            'Low': '#5cb85c', 'Bassa': '#5cb85c',
        }.get(priority, '#777777')

    def _add_page_number(self, canvas, doc):
        """Adds the page number at the bottom of the page."""
        page_num = canvas.getPageNumber()
        text = f"Page {page_num}"
        canvas.saveState()
        canvas.setFont('Helvetica', 9)
        canvas.drawRightString(A4[0] - 2 * cm, 1.5 * cm, text)
        canvas.restoreState()

    def generate_kaizen_pdf(self, kaizen_data: dict) -> str | None:
        """
        Generates a structured "Kaizen Project Charter" PDF using ReportLab.

        Args:
            kaizen_data (dict): Dictionary with the Kaizen proposal.

        Returns:
            str | None: Path to the generated PDF or None on error.
        """
        if not kaizen_data:
            logger.warning("No Kaizen data provided. Skipping PDF generation.")
            return None

        try:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f"Kaizen_Charter_{timestamp}.pdf"
            filepath = Path('reports') / filename

            doc = SimpleDocTemplate(str(filepath), pagesize=A4,
                                    rightMargin=2 * cm, leftMargin=2 * cm,
                                    topMargin=2 * cm, bottomMargin=2 * cm)

            story = []

            # Title
            story.append(Paragraph("Kaizen Project Charter", self.styles['CustomTitle']))
            story.append(Paragraph(kaizen_data.get('project_title', 'N/A'), self.styles['h2']))
            story.append(HRFlowable(width="100%", thickness=1, color=colors.grey))
            story.append(Spacer(1, 1 * cm))

            # Sections
            story.append(Paragraph("1. Problem Statement", self.styles['CustomHeading2']))
            story.append(
                Paragraph(kaizen_data.get('problem_statement', 'No description provided.'), self.styles['CustomBody']))
            story.append(Spacer(1, 0.5 * cm))

            story.append(Paragraph("2. Project Goal (SMART Goal)", self.styles['CustomHeading2']))
            story.append(Paragraph(kaizen_data.get('goal', 'No goal defined.'), self.styles['CustomBody']))
            story.append(Spacer(1, 0.5 * cm))

            story.append(Paragraph("3. Suggested Team", self.styles['CustomHeading2']))
            team_members = kaizen_data.get('suggested_team', [])
            for member in team_members:
                story.append(Paragraph(f"• {member}", self.styles['CustomBody']))
            story.append(Spacer(1, 0.5 * cm))

            story.append(Paragraph("4. Initial Steps (PDCA: Plan Phase)", self.styles['CustomHeading2']))
            initial_steps = kaizen_data.get('initial_steps', [])
            for i, step in enumerate(initial_steps, 1):
                story.append(Paragraph(f"{i}. {step}", self.styles['CustomBody']))
            story.append(Spacer(1, 1 * cm))

            # Signature block (on a new page for cleanliness)
            story.append(PageBreak())
            story.append(Spacer(1, 15 * cm))  # Spacer to position at the bottom

            signature_data = [
                ['Project Sponsor:', 'Team Leader:'],
                ['\n\n____________________', '\n\n____________________'],
                ['(Signature)', '(Signature)']
            ]

            table = Table(signature_data, colWidths=[8 * cm, 8 * cm])
            table.setStyle(TableStyle([
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTNAME', (0, 2), (-1, 2), 'Helvetica-Oblique'),
                ('FONTSIZE', (0, 2), (-1, 2), 8),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ]))
            story.append(table)

            # Build the document
            doc.build(story)

            logger.info(f"Kaizen Charter PDF generated successfully: {filepath}")
            return str(filepath)

        except Exception as e:
            logger.error(f"Error during Kaizen PDF generation: {e}", exc_info=True)
            return None


# Utility function for quick use (outside the class)
def generate_pdf_utility(analysis_data, output_path=None, title="AI Analysis Report"):
    """
    Utility function to quickly generate a PDF report.
    """
    logger.info(f"Quick PDF generation: {title}")
    try:
        generator = PDFReportGenerator(title=title)
        result = generator.generate_report(analysis_data, output_path)
        return result
    except Exception as e:
        logger.error(f"Error in quick PDF generation: {e}", exc_info=True)
        return None


if __name__ == "__main__":
    # Test the module
    logger.info("Running PDFGenerator test...")

    if not Path('reports').exists():
        Path('reports').mkdir()

    # Test data for the generic report
    test_data = {
        'version': '1.0',
        'analysis_type': 'Test Report',
        'executive_summary': 'This is a test report to verify the PDF generation with ReportLab. The summary describes the performance of the period, highlighting the main critical issues and successes achieved.',
        'root_causes': [
            {'problem_area': 'Machinery', 'cause_description': 'Wear of component XY on line 3.'},
            {'problem_area': 'Material', 'cause_description': 'Non-conforming material batch from supplier Z.'}
        ],
        'recommendations': [
            {'priority': 'High', 'title': 'Maintenance for Component XY',
             'description': 'Schedule extraordinary maintenance to replace the worn component.'},
            {'priority': 'Medium', 'title': 'Audit Supplier Z',
             'description': 'Verify the quality processes of supplier Z.'},
        ],
        'tables': [{
            'title': 'Production Data by Product',
            'data': [['Product', 'Quantity', 'Defects (%)'], ['Product A', '1000', '1.5'], ['Product B', '1500', '2.1']]
        }]
    }

    # Test data for the Kaizen Charter
    test_kaizen_data = {
        'project_title': 'Machine Downtime Reduction for Component XY Wear',
        'problem_statement': 'Line 3 has recorded 15 hours of machine downtime in the last month due to recurring failures of component XY, causing an estimated production loss of 5000 units.',
        'goal': 'Reduce machine downtime due to component XY by 90% by the end of the next quarter, from 15 hours/month to less than 1.5 hours/month.',
        'suggested_team': ['Maintenance Manager', 'Line 3 Operator', 'Process Engineer', 'Quality Control'],
        'initial_steps': [
            'Perform 5 Whys analysis to determine the root cause of accelerated wear.',
            'Evaluate alternative materials or suppliers for component XY.',
            'Define a specific preventive maintenance plan for the component.',
            'Install sensors to monitor the vibration and temperature of the component.'
        ]
    }

    try:
        # Generate the main analysis report
        generator = PDFReportGenerator(title="Test Report - General Analysis")
        report_path = generator.generate_report(test_data)
        if report_path:
            logger.info(f"Test report file saved successfully at: {report_path}")

        # Generate the Kaizen charter
        kaizen_path = generator.generate_kaizen_pdf(test_kaizen_data)
        if kaizen_path:
            logger.info(f"Test Kaizen charter file saved successfully at: {kaizen_path}")

    except Exception as e:
        logger.error(f"Test failed: {e}", exc_info=True)