# utils.py
from email_connector import EmailSender
import logging
import re
from typing import List, Optional


logger = logging.getLogger("TraceabilityRS")  # usa la config fatta in main.py


def get_email_recipients(conn, attribute: str = 'Sys_Email_Purchase') -> List[str]:
    """
    Recupera gli indirizzi email dei destinatari dal database per lo specifico attributo.
    Esempi di attributo: 'Sys_Email_Purchase', 'Sys_email_submission'
    """
    cursor = None
    try:
        # ✅ VERIFICA CONNESSIONE
        if conn is None:
            logger.error("Connessione database non valida (None)")
            return []

        # ✅ TEST CONNESSIONE
        try:
            cursor = conn.cursor()
        except Exception as e:
            logger.error(f"Connessione chiusa o non valida: {str(e)}")
            return []

        query = """
                SELECT [VALUE]
                FROM traceability_rs.dbo.settings
                WHERE atribute = ? \
                """

        logger.info(f"Eseguo query per attributo: {attribute}")
        cursor.execute(query, attribute)
        results = cursor.fetchall()

        logger.info(f"Query eseguita, trovate {len(results)} righe")

        email_list = [row[0] for row in results if row[0]]

        valid_emails = []
        for email in email_list:
            chunks = []
            if ';' in email:
                chunks = [e.strip() for e in email.split(';')]
            elif ',' in email:
                chunks = [e.strip() for e in email.split(',')]
            else:
                chunks = [email.strip()]
            valid_emails.extend([e for e in chunks if e and '@' in e])

        logger.info(f"Indirizzi email validi trovati per {attribute}: {valid_emails}")
        return valid_emails

    except Exception as e:
        logger.error(f"Errore nel recupero degli indirizzi email ({attribute}): {str(e)}")
        import traceback
        logger.error(traceback.format_exc())
        return []  # ← Restituisci lista vuota invece di raise

    finally:
        # ✅ CHIUDI SOLO IL CURSOR, NON LA CONNESSIONE
        if cursor is not None:
            try:
                cursor.close()
                logger.debug("Cursor chiuso correttamente")
            except Exception as e:
                logger.warning(f"Errore nella chiusura del cursor: {e}")
# def get_email_recipients(conn, attribute: str = 'Sys_Email_Purchase') -> List[str]:
#     """
#     Recupera gli indirizzi email dei destinatari dal database per lo specifico attributo.
#     Esempi di attributo: 'Sys_Email_Purchase', 'Sys_email_submission'
#     """
#     cursor = None
#     try:
#         # ✅ VERIFICA CONNESSIONE
#         if conn is None:
#             logger.error("Connessione database non valida (None)")
#             return []
#
#         # ✅ TEST CONNESSIONE
#         try:
#             cursor = conn.cursor()
#         except Exception as e:
#             logger.error(f"Connessione chiusa o non valida: {str(e)}")
#             return []
#
#         query = """
#                 SELECT [VALUE]
#                 FROM traceability_rs.dbo.settings
#                 WHERE atribute = ? \
#                 """
#
#         logger.info(f"Eseguo query per attributo: {attribute}")
#         cursor.execute(query, attribute)
#         results = cursor.fetchall()
#
#         logger.info(f"Query eseguita, trovate {len(results)} righe")
#
#         # DEBUG: Mostra i risultati grezzi
#         for i, row in enumerate(results):
#             logger.info(f"Risultato {i + 1}: {row[0] if row else 'None'}")
#
#         email_list = [row[0] for row in results if row[0]]
#
#         valid_emails = []
#         for email in email_list:
#             chunks = []
#             if ';' in email:
#                 chunks = [e.strip() for e in email.split(';')]
#             elif ',' in email:
#                 chunks = [e.strip() for e in email.split(',')]
#             else:
#                 chunks = [email.strip()]
#
#             # Filtra email valide
#             for e in chunks:
#                 if e and '@' in e:
#                     # Pulizia aggiuntiva
#                     clean_email = e.strip()
#                     if clean_email and re.match(r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$', clean_email):
#                         valid_emails.append(clean_email)
#                     else:
#                         logger.warning(f"Email non valida scartata: {clean_email}")
#
#         logger.info(f"Indirizzi email validi trovati per {attribute}: {valid_emails}")
#         return valid_emails
#
#     except Exception as e:
#         logger.error(f"Errore nel recupero degli indirizzi email ({attribute}): {str(e)}")
#         import traceback
#         logger.error(traceback.format_exc())
#         return []  # ← Restituisci lista vuota invece di raise
#
#     finally:
#         # ✅ CHIUDI SOLO IL CURSOR, NON LA CONNESSIONE
#         if cursor is not None:
#             try:
#                 cursor.close()
#                 logger.debug("Cursor chiuso correttamente")
#             except Exception as e:
#                 logger.warning(f"Errore nella chiusura del cursor: {e}")
def send_email(
    recipients: List[str],
    subject: str,
    body: str,
    smtp_host: str = "vandewiele-com.mail.protection.outlook.com",
    smtp_port: int = 25,
    is_html: bool = False,
    attachments: List[str] = None,  # <-- NUOVO parametro per allegati
    timeout: int = 15
) -> None:
    """
    Invia l'email ai destinatari specificati.

    Args:
        recipients: Lista di indirizzi email destinatari
        subject: Oggetto dell'email
        body: Corpo dell'email (testo o HTML se is_html=True)
        smtp_host: Host SMTP
        smtp_port: Porta SMTP
        is_html: Se True invia il corpo come HTML (default: False)
        attachments: Lista di percorsi file da allegare (default: None)

    Note: Usa EmailSender già presente nel progetto.
    """
    if not recipients:
        logger.error("Nessun destinatario specificato per l'email")
        return

    try:
        sender = EmailSender(smtp_host, smtp_port)

        # ATTENZIONE: credenziali hardcoded – idealmente spostarle in config sicura
        sender.save_credentials(
            "Accounting@Eutron.it",
            "9jHgFhSs7Vf+"
        )

        sender.send_email(
            to_email=', '.join(recipients),
            subject=subject,
            body=body,
            is_html=is_html,
            attachments=attachments  # <-- Passa gli allegati
        )
        logger.info("Email inviata con successo a %d destinatari", len(recipients))
        print("email inviata")
    except Exception as e:
        logger.error("Errore nell'invio dell'email: %s", str(e))
        raise