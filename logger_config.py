"""
Logger configuration for AI Scrap Analysis
"""
import logging
import sys
from pathlib import Path

class UnicodeStreamHandler(logging.StreamHandler):
    """Stream handler che gestisce correttamente Unicode su Windows"""

    def emit(self, record):
        try:
            msg = self.format(record)
            stream = self.stream
            # Encoding sicuro per Windows
            stream.write(msg.encode('utf-8', errors='replace').decode('utf-8') + self.terminator)
            self.flush()
        except Exception:
            self.handleError(record)

def setup_logger(name, log_file=None, level=logging.INFO):
    """Setup logger with console and file handlers"""

    # Create logger
    logger = logging.getLogger(name)
    logger.setLevel(level)

    # Avoid adding handlers multiple times
    if logger.handlers:
        return logger

    # Create formatter
    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )

    # Console handler con gestione Unicode
    console_handler = UnicodeStreamHandler(sys.stdout)
    console_handler.setLevel(level)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    # File handler (if log file specified)
    if log_file:
        log_path = Path(log_file)
        log_path.parent.mkdir(parents=True, exist_ok=True)

        file_handler = logging.FileHandler(log_path, encoding='utf-8')
        file_handler.setLevel(level)
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)

    return logger