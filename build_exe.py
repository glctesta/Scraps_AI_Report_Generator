# build_simple.py - Versione ultra-semplice
import os
import subprocess
import sys


def build_simple():
    print("Build Scarps AI Analysis - Versione Semplice")
    print("Assicurati di essere nella cartella con main.py e Logo.png")
    print()

    # Lista file nella directory corrente
    print("File nella directory corrente:")
    for file in os.listdir('.'):
        print(f"  - {file}")

    input("\nPremi INVIO per avviare la build...")

    cmd = [
        'pyinstaller',
        '--onedir',
        '--name', 'Scarps_AI_Analisys',
        '--icon=Logo.png',
        '--add-data', 'Logo.png;.',
        '--add-data', 'db_config.enc;.',
        '--add-data', 'encryption_key.key;.',
        '--noconsole',
        'main.py'
    ]

    try:
        subprocess.run(cmd, check=True)
        print("\n✅ Build completata!")
        print("Eseguibile in: dist/Scarps_AI_Analisys/")
    except Exception as e:
        print(f"\n❌ Errore: {e}")


if __name__ == "__main__":
    build_simple()