import os
import json
from pathlib import Path
from cryptography.fernet import Fernet
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
import sys
import logging

logger = logging.getLogger('ConfigManager')


class ConfigManager:
    def __init__(self, key_file='encryption_key.key', config_file='db_config.enc'):
        """
        Inizializza il gestore di configurazione

        Args:
            key_file: Nome file chiave di cifratura
            config_file: Nome file configurazione cifrato
        """
        self.logger = logger
        self.key_file = key_file
        self.config_file = config_file

    def _get_base_path(self):
        """Restituisce il path base corretto per l'eseguibile o script"""
        try:
            # Se siamo in un eseguibile PyInstaller
            if getattr(sys, 'frozen', False):
                base_path = Path(sys._MEIPASS)
            else:
                base_path = Path(__file__).parent

            self.logger.debug(f"Base path: {base_path}")
            return base_path
        except Exception as e:
            self.logger.warning(f"Errore nel determinare base path: {e}")
            return Path.cwd()

    def _get_file_path(self, filename):
        """Restituisce il path completo del file, cercando in diverse posizioni"""
        base_path = self._get_base_path()

        # Cerca nella directory corrente (per eseguibile)
        current_dir = Path.cwd()
        file_in_current = current_dir / filename
        if file_in_current.exists():
            self.logger.debug(f"Trovato {filename} in: {current_dir}")
            return file_in_current

        # Cerca nella directory base
        file_in_base = base_path / filename
        if file_in_base.exists():
            self.logger.debug(f"Trovato {filename} in: {base_path}")
            return file_in_base

        # Cerca nella directory dello script
        script_dir = Path(__file__).parent
        file_in_script = script_dir / filename
        if file_in_script.exists():
            self.logger.debug(f"Trovato {filename} in: {script_dir}")
            return file_in_script

        self.logger.error(f"File {filename} non trovato in:")
        self.logger.error(f"  - {current_dir}")
        self.logger.error(f"  - {base_path}")
        self.logger.error(f"  - {script_dir}")

        return None

    def load_config(self):
        """
        Carica e decifra la configurazione del database

        Returns:
            dict: Configurazione del database

        Raises:
            FileNotFoundError: Se i file di configurazione non sono trovati
        """
        try:
            # Ottieni i path dei file
            key_path = self._get_file_path(self.key_file)
            config_path = self._get_file_path(self.config_file)

            if not key_path or not config_path:
                raise FileNotFoundError(
                    f"File di configurazione non trovato. Cercati: {self.key_file}, {self.config_file}")

            self.logger.info(f"Caricamento configurazione da: {config_path}")

            # Carica la chiave
            with open(key_path, 'rb') as f:
                key = f.read()

            # Carica e decifra la configurazione
            with open(config_path, 'rb') as f:
                encrypted_data = f.read()

            fernet = Fernet(key)
            decrypted_data = fernet.decrypt(encrypted_data)
            config = json.loads(decrypted_data.decode())

            self.logger.info("Configurazione caricata con successo")
            return config

        except Exception as e:
            self.logger.error(f"Errore nel caricamento configurazione: {e}")
            raise