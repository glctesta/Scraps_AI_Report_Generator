"""
Modulo per l'invio di email con allegati (PDF, Excel)
Supporta SMTP con TLS/SSL e autenticazione
"""

import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
from typing import List, Union, Dict
import mimetypes

from logger_config import setup_logger

# Inizializza logger
logger = setup_logger('EmailSender')


class EmailSender:
    """
    Classe per l'invio di email con allegati
    """

    def __init__(self, smtp_config: Dict[str, any]):
        """
        Inizializza il sender email

        Args:
            smtp_config (dict): Configurazione SMTP
                {
                    'server': 'smtp.gmail.com',
                    'port': 587,
                    'username': 'user@example.com',
                    'password': 'password',
                    'use_tls': True,
                    'use_ssl': False,
                    'from_address': 'sender@example.com',
                    'from_name': 'AI Report System'
                }
        """
        logger.info("Inizializzazione EmailSender")
        logger.debug(f"Configurazione SMTP: server={smtp_config.get('server')}, port={smtp_config.get('port')}")

        self.smtp_server = smtp_config.get('server')
        self.smtp_port = smtp_config.get('port', 587)
        self.username = smtp_config.get('username')
        self.password = smtp_config.get('password')
        self.use_tls = smtp_config.get('use_tls', True)
        self.use_ssl = smtp_config.get('use_ssl', False)
        self.from_address = smtp_config.get('from_address', self.username)
        self.from_name = smtp_config.get('from_name', 'AI Report System')

        # Validazione configurazione
        if not all([self.smtp_server, self.smtp_port, self.username, self.password]):
            logger.error("Configurazione SMTP incompleta")
            raise ValueError("Configurazione SMTP incompleta. Verificare server, port, username, password")

        logger.info("EmailSender inizializzato con successo")

    def send_email(self,
                   to_addresses: Union[str, List[str]],
                   subject: str,
                   body: str,
                   attachments: List[Dict[str, any]] = None,
                   cc_addresses: Union[str, List[str]] = None,
                   bcc_addresses: Union[str, List[str]] = None,
                   html_body: str = None) -> bool:
        """
        Invia email con allegati

        Args:
            to_addresses: Destinatario/i (stringa o lista)
            subject: Oggetto email
            body: Corpo email (testo)
            attachments: Lista di allegati [{'filename': 'report.pdf', 'data': bytes, 'path': 'file.pdf'}]
            cc_addresses: Destinatari in copia
            bcc_addresses: Destinatari in copia nascosta
            html_body: Corpo email in HTML (opzionale)

        Returns:
            bool: True se invio riuscito, False altrimenti
        """
        logger.info(f"Preparazione invio email - Oggetto: '{subject}'")

        try:
            # Normalizza destinatari
            to_list = self._normalize_addresses(to_addresses)
            cc_list = self._normalize_addresses(cc_addresses) if cc_addresses else []
            bcc_list = self._normalize_addresses(bcc_addresses) if bcc_addresses else []

            logger.info(f"Destinatari: {len(to_list)} TO, {len(cc_list)} CC, {len(bcc_list)} BCC")
            logger.debug(f"TO: {to_list}")

            # Crea messaggio
            msg = MIMEMultipart('alternative')
            msg['From'] = f"{self.from_name} <{self.from_address}>"
            msg['To'] = ', '.join(to_list)
            msg['Subject'] = subject
            msg['Date'] = datetime.now().strftime('%a, %d %b %Y %H:%M:%S %z')

            if cc_list:
                msg['Cc'] = ', '.join(cc_list)

            # Corpo email
            if html_body:
                logger.debug("Aggiunta corpo HTML")
                msg.attach(MIMEText(body, 'plain', 'utf-8'))
                msg.attach(MIMEText(html_body, 'html', 'utf-8'))
            else:
                logger.debug("Aggiunta corpo testo semplice")
                msg.attach(MIMEText(body, 'plain', 'utf-8'))

            # Allegati
            if attachments:
                logger.info(f"Aggiunta {len(attachments)} allegati")
                for attachment in attachments:
                    self._add_attachment(msg, attachment)

            # Lista completa destinatari
            all_recipients = to_list + cc_list + bcc_list
            logger.debug(f"Totale destinatari (inclusi BCC): {len(all_recipients)}")

            # Invio email
            logger.info(f"Connessione a server SMTP: {self.smtp_server}:{self.smtp_port}")

            if self.use_ssl:
                logger.debug("Utilizzo SSL")
                server = smtplib.SMTP_SSL(self.smtp_server, self.smtp_port, timeout=30)
            else:
                logger.debug("Utilizzo connessione standard")
                server = smtplib.SMTP(self.smtp_server, self.smtp_port, timeout=30)

                if self.use_tls:
                    logger.debug("Avvio STARTTLS")
                    server.starttls()

            # Login
            logger.info(f"Login con username: {self.username}")
            server.login(self.username, self.password)
            logger.info("Login SMTP riuscito")

            # Invio
            logger.info("Invio email in corso...")
            server.send_message(msg, from_addr=self.from_address, to_addrs=all_recipients)
            server.quit()

            logger.info(f"✓ Email inviata con successo a {len(all_recipients)} destinatari")
            return True

        except smtplib.SMTPAuthenticationError as e:
            logger.error(f"Errore autenticazione SMTP: {str(e)}", exc_info=True)
            return False
        except smtplib.SMTPException as e:
            logger.error(f"Errore SMTP: {str(e)}", exc_info=True)
            return False
        except Exception as e:
            logger.error(f"Errore generico nell'invio email: {str(e)}", exc_info=True)
            return False

    def _normalize_addresses(self, addresses: Union[str, List[str]]) -> List[str]:
        """Normalizza gli indirizzi email in una lista"""
        if isinstance(addresses, str):
            # Separa per virgola o punto e virgola
            return [addr.strip() for addr in addresses.replace(';', ',').split(',') if addr.strip()]
        elif isinstance(addresses, list):
            return [addr.strip() for addr in addresses if addr.strip()]
        return []

    def _add_attachment(self, msg: MIMEMultipart, attachment: Dict[str, any]):
        """
        Aggiunge un allegato al messaggio

        Args:
            msg: Messaggio MIME
            attachment: Dict con 'filename' e ('data' o 'path')
        """
        filename = attachment.get('filename')

        if not filename:
            logger.warning("Allegato senza nome file, ignorato")
            return

        logger.debug(f"Aggiunta allegato: {filename}")

        try:
            # Ottieni dati allegato
            if 'data' in attachment:
                # Dati già in memoria (bytes)
                data = attachment['data']
                logger.debug(f"Allegato da dati in memoria: {len(data)} bytes")
            elif 'path' in attachment:
                # Leggi da file
                file_path = attachment['path']
                if not os.path.exists(file_path):
                    logger.error(f"File allegato non trovato: {file_path}")
                    return

                with open(file_path, 'rb') as f:
                    data = f.read()
                logger.debug(f"Allegato da file: {file_path} ({len(data)} bytes)")
            else:
                logger.warning("Allegato senza 'data' o 'path', ignorato")
                return

            # Determina MIME type
            mime_type, _ = mimetypes.guess_type(filename)
            if mime_type is None:
                mime_type = 'application/octet-stream'

            logger.debug(f"MIME type: {mime_type}")

            # Crea parte allegato
            maintype, subtype = mime_type.split('/', 1)
            part = MIMEBase(maintype, subtype)
            part.set_payload(data)
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename="{filename}"')

            msg.attach(part)
            logger.debug(f"✓ Allegato '{filename}' aggiunto con successo")

        except Exception as e:
            logger.error(f"Errore nell'aggiunta allegato '{filename}': {str(e)}", exc_info=True)

    def send_report_email(self,
                         to_addresses: Union[str, List[str]],
                         report_title: str,
                         report_summary: str,
                         pdf_data: bytes = None,
                         excel_data: bytes = None,
                         cc_addresses: Union[str, List[str]] = None) -> bool:
        """
        Metodo di convenienza per inviare email con report PDF/Excel

        Args:
            to_addresses: Destinatari
            report_title: Titolo del report
            report_summary: Sommario del report
            pdf_data: Dati PDF (bytes)
            excel_data: Dati Excel (bytes)
            cc_addresses: Destinatari in copia

        Returns:
            bool: True se invio riuscito
        """
        logger.info(f"Invio report email: {report_title}")

        # Crea oggetto
        subject = f"Report AI: {report_title}"

        # Crea corpo email
        body = f"""
Gentile utente,

in allegato trovi il report di analisi AI generato automaticamente.

Titolo: {report_title}
Data generazione: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}

SOMMARIO:
{report_summary}

---
Questo è un messaggio automatico del sistema AI Report Generator.
Per qualsiasi domanda, contatta il supporto tecnico.

Cordiali saluti,
{self.from_name}
        """

        # Crea corpo HTML
        html_body = f"""
<html>
<head>
    <style>
        body {{ font-family: Arial, sans-serif; line-height: 1.6; color: #333; }}
        .header {{ background-color: #1a5490; color: white; padding: 20px; text-align: center; }}
        .content {{ padding: 20px; }}
        .summary {{ background-color: #f4f4f4; padding: 15px; border-left: 4px solid #1a5490; margin: 20px 0; }}
        .footer {{ background-color: #f4f4f4; padding: 10px; text-align: center; font-size: 12px; color: #666; }}
    </style>
</head>
<body>
    <div class="header">
        <h1>Report Analisi AI</h1>
    </div>
    <div class="content">
        <p>Gentile utente,</p>
        <p>in allegato trovi il report di analisi AI generato automaticamente.</p>
        
        <p><strong>Titolo:</strong> {report_title}<br>
        <strong>Data generazione:</strong> {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}</p>
        
        <div class="summary">
            <h3>SOMMARIO</h3>
            <p>{report_summary.replace(chr(10), '<br>')}</p>
        </div>
        
        <p>I report sono disponibili nei seguenti formati:</p>
        <ul>
            {'<li>📄 PDF - Report completo con grafici</li>' if pdf_data else ''}
            {'<li>📊 Excel - Dati analitici ed elaborabili</li>' if excel_data else ''}
        </ul>
    </div>
    <div class="footer">
        <p>Questo è un messaggio automatico del sistema AI Report Generator.<br>
        Per qualsiasi domanda, contatta il supporto tecnico.</p>
        <p>© {datetime.now().year} {self.from_name}</p>
    </div>
</body>
</html>
        """

        # Prepara allegati
        attachments = []

        if pdf_data:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            attachments.append({
                'filename': f'Report_AI_{timestamp}.pdf',
                'data': pdf_data
            })
            logger.info(f"Allegato PDF preparato: {len(pdf_data)} bytes")

        if excel_data:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            attachments.append({
                'filename': f'Report_AI_{timestamp}.xlsx',
                'data': excel_data
            })
            logger.info(f"Allegato Excel preparato: {len(excel_data)} bytes")

        if not attachments:
            logger.warning("Nessun allegato specificato per l'email del report")

        # Invia email
        return self.send_email(
            to_addresses=to_addresses,
            subject=subject,
            body=body,
            html_body=html_body,
            attachments=attachments,
            cc_addresses=cc_addresses
        )

    def test_connection(self) -> bool:
        """
        Testa la connessione SMTP

        Returns:
            bool: True se connessione riuscita
        """
        logger.info("Test connessione SMTP in corso...")

        try:
            if self.use_ssl:
                server = smtplib.SMTP_SSL(self.smtp_server, self.smtp_port, timeout=10)
            else:
                server = smtplib.SMTP(self.smtp_server, self.smtp_port, timeout=10)
                if self.use_tls:
                    server.starttls()

            server.login(self.username, self.password)
            server.quit()

            logger.info("✓ Test connessione SMTP riuscito")
            return True

        except smtplib.SMTPAuthenticationError:
            logger.error("✗ Test connessione fallito: Errore autenticazione")
            return False
        except smtplib.SMTPException as e:
            logger.error(f"✗ Test connessione fallito: {str(e)}")
            return False
        except Exception as e:
            logger.error(f"✗ Test connessione fallito: {str(e)}")
            return False


# Funzione di utilità per invio rapido
def send_quick_email(smtp_config: Dict, to_addresses: str, subject: str, body: str,
                     attachments: List[Dict] = None) -> bool:
    """
    Funzione di utilità per invio rapido email

    Args:
        smtp_config: Configurazione SMTP
        to_addresses: Destinatari
        subject: Oggetto
        body: Corpo email
        attachments: Lista allegati

    Returns:
        bool: True se invio riuscito
    """
    logger.info(f"Invio rapido email: {subject}")
    try:
        sender = EmailSender(smtp_config)
        result = sender.send_email(to_addresses, subject, body, attachments)
        if result:
            logger.info("Email inviata con successo tramite funzione rapida")
        else:
            logger.warning("Invio email fallito tramite funzione rapida")
        return result
    except Exception as e:
        logger.error(f"Errore nell'invio rapido email: {str(e)}", exc_info=True)
        return False


if __name__ == "__main__":
    # Test del modulo
    logger.info("Test EmailSender in esecuzione...")

    # Configurazione di test (MODIFICARE CON DATI REALI)
    test_config = {
        'server': 'smtp.gmail.com',
        'port': 587,
        'username': 'your_email@gmail.com',
        'password': 'your_app_password',
        'use_tls': True,
        'from_address': 'your_email@gmail.com',
        'from_name': 'AI Report System Test'
    }

    try:
        sender = EmailSender(test_config)

        # Test connessione
        if sender.test_connection():
            logger.info("Test connessione superato")

            # Test invio email semplice
            # ATTENZIONE: Decommentare solo per test reale
            # result = sender.send_email(
            #     to_addresses='recipient@example.com',
            #     subject='Test Email AI Report System',
            #     body='Questa è una email di test dal sistema AI Report.',
            #     html_body='<h1>Test</h1><p>Questa è una email di test dal sistema AI Report.</p>'
            # )
            #
            # if result:
            #     logger.info("Test invio email superato")
            # else:
            #     logger.error("Test invio email fallito")
        else:
            logger.error("Test connessione fallito")

    except Exception as e:
        logger.error(f"Test fallito: {str(e)}", exc_info=True)