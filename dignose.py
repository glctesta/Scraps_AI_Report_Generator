"""
Script di diagnosi per Analisi_scraps
"""
import sys
import os
from pathlib import Path

print("=" * 60)
print("DIAGNOSI AMBIENTE - Analisi Scraps")
print("=" * 60)

# 1. Verifica Python
print(f"\n1. Python Version: {sys.version}")
print(f"   Executable: {sys.executable}")

# 2. Verifica Directory
print(f"\n2. Directory corrente: {os.getcwd()}")
print(f"   Script path: {Path(__file__).parent}")

# 3. Verifica Moduli
print("\n3. Moduli installati:")
try:
    import reportlab

    print(f"   ✓ reportlab: {reportlab.Version}")
except ImportError:
    print("   ✗ reportlab: NON INSTALLATO")

try:
    import openpyxl

    print(f"   ✓ openpyxl: {openpyxl.__version__}")
except ImportError:
    print("   ✗ openpyxl: NON INSTALLATO")

# 4. Verifica File Progetto
print("\n4. File progetto:")
project_files = [
    "main.py",
    "ai_report_generator.py",
    "pdf_generator.py",
    "excel_generator.py",
    ".env"
]

for file in project_files:
    exists = Path(file).exists()
    status = "✓" if exists else "✗"
    print(f"   {status} {file}")

# 5. Verifica Directory
print("\n5. Directory:")
directories = ["data", "data/input", "output", "logs"]
for dir_path in directories:
    exists = Path(dir_path).exists()
    status = "✓" if exists else "✗"
    print(f"   {status} {dir_path}")

# 6. Verifica main.py
print("\n6. Analisi main.py:")
try:
    with open("main.py", "r", encoding="utf-8") as f:
        content = f.read()

    has_main_block = 'if __name__ == "__main__"' in content
    has_main_function = "def main(" in content
    has_class = "class MainApplication" in content or "class Main" in content

    print(f"   - Blocco if __name__: {'✓' if has_main_block else '✗'}")
    print(f"   - Funzione main(): {'✓' if has_main_function else '✗'}")
    print(f"   - Classe Main: {'✓' if has_class else '✗'}")

    # Mostra ultime 20 righe
    lines = content.split('\n')
    print(f"\n   Ultime 20 righe di main.py:")
    print("   " + "-" * 50)
    for line in lines[-20:]:
        print(f"   {line}")

except FileNotFoundError:
    print("   ✗ main.py NON TROVATO!")
except Exception as e:
    print(f"   ✗ Errore lettura: {e}")

# 7. Prova import
print("\n7. Test import moduli:")
try:
    import main

    print("   ✓ import main: OK")
    print(f"   - Attributi: {[x for x in dir(main) if not x.startswith('_')]}")
except Exception as e:
    print(f"   ✗ import main: ERRORE - {e}")

print("\n" + "=" * 60)
print("DIAGNOSI COMPLETATA")
print("=" * 60)
