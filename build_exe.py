# build_exe.py - Script Python per generare l'eseguibile
import os
import subprocess
import sys
import shutil


def build_executable():
    print("=" * 50)
    print("BUILD SCARPS AI ANALYSIS EXECUTABLE")
    print("=" * 50)

    # Verifica che i file necessari esistano
    required_files = [
        'main.py',
        'Logo.png',
        'db_config.enc',
        'encryption_key.key'
    ]

    missing_files = []
    for file in required_files:
        if not os.path.exists(file):
            missing_files.append(file)

    if missing_files:
        print(f"ERRORE: File mancanti: {missing_files}")
        return False

    print("File richiesti trovati ✓")

    # Pulisci build precedenti
    if os.path.exists('build'):
        shutil.rmtree('build')
        print("Cartella build pulita ✓")

    if os.path.exists('dist'):
        shutil.rmtree('dist')
        print("Cartella dist pulita ✓")

    # Comando PyInstaller
    cmd = [
        'pyinstaller',
        '--onedir',
        '--name', 'Scarps_AI_Analisys',
        '--icon=Logo.png',
        '--add-data', 'Logo.png;.',
        '--add-data', 'db_config.enc;.',
        '--add-data', 'encryption_key.key;.',
        '--add-data', 'email_credentials.enc;.',
        '--add-data', 'email_key.key;.',
        '--add-data', 'ai_config.json;.',
        '--hidden-import=email.mime.multipart',
        '--hidden-import=email.mime.text',
        '--hidden-import=email.mime.base',
        '--hidden-import=email.encoders',
        '--hidden-import=cryptography.fernet',
        '--hidden-import=cryptography.hazmat.backends.openssl.backend',
        '--hidden-import=cryptography.hazmat.primitives',
        '--hidden-import=cryptography.hazmat.primitives.kdf.pbkdf2',
        '--hidden-import=pyodbc',
        '--hidden-import=openpyxl',
        '--hidden-import=reportlab',
        '--hidden-import=reportlab.lib',
        '--hidden-import=reportlab.pdfgen',
        '--hidden-import=reportlab.platypus',
        '--hidden-import=PIL',
        '--hidden-import=PIL.Image',
        '--hidden-import=PIL.ImageOps',
        '--hidden-import=requests',
        '--hidden-import=ollama',
        '--collect-all=reportlab',
        '--collect-all=PIL',
        '--noconsole',
        'main.py'
    ]

    # File opzionali - aggiungi solo se esistono
    optional_files = [
        'email_credentials.enc',
        'email_key.key',
        'ai_config.json'
    ]

    for file in optional_files:
        if os.path.exists(file):
            cmd.extend(['--add-data', f'{file};.'])
            print(f"File opzionale incluso: {file} ✓")

    print("\nAvvio build con PyInstaller...")
    print(" ".join(cmd))

    try:
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
        print("Build completata con successo! ✓")

        # Mostra il percorso dell'eseguibile
        exe_path = os.path.join('dist', 'Scarps_AI_Analisys', 'Scarps_AI_Analisys.exe')
        if os.path.exists(exe_path):
            print(f"\nEseguibile generato: {exe_path}")

            # Elenca i file inclusi
            dist_dir = os.path.join('dist', 'Scarps_AI_Analisys')
            print(f"\nFile inclusi nella distribuzione:")
            for root, dirs, files in os.walk(dist_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    rel_path = os.path.relpath(file_path, dist_dir)
                    print(f"  - {rel_path}")

            return True
        else:
            print("ERRORE: Eseguibile non generato")
            return False

    except subprocess.CalledProcessError as e:
        print(f"ERRORE durante la build: {e}")
        print(f"Output: {e.stdout}")
        print(f"Error: {e.stderr}")
        return False
    except FileNotFoundError:
        print("ERRORE: PyInstaller non trovato. Installa con: pip install pyinstaller")
        return False


if __name__ == "__main__":
    if build_executable():
        print("\n" + "=" * 50)
        print("BUILD COMPLETATO CON SUCCESSO!")
        print("=" * 50)
    else:
        print("\n" + "=" * 50)
        print("BUILD FALLITA!")
        print("=" * 50)
        sys.exit(1)