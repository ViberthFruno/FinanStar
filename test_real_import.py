#!/usr/bin/env python3
"""
Test importando los módulos REALES de la aplicación
"""
import sys
import os
from pathlib import Path

# Añadir el directorio actual al path
sys.path.insert(0, str(Path(__file__).parent))

print("=" * 70)
print("TEST CON IMPORTACIÓN REAL DE MÓDULOS")
print("=" * 70)

# Importar el caso_handler real
print("\n1. Importando case_handler...")
try:
    from case_handler import CaseHandler
    print("   ✓ case_handler importado correctamente")
except Exception as e:
    print(f"   ✗ Error importando: {e}")
    sys.exit(1)

# Importar logger mock
print("\n2. Creando logger mock...")
class MockLogger:
    def log(self, message, level="INFO"):
        print(f"   [{level}] {message}")

logger = MockLogger()

# Crear instancia de CaseHandler
print("\n3. Creando instancia de CaseHandler...")
try:
    handler = CaseHandler()
    print("   ✓ CaseHandler instanciado correctamente")
except Exception as e:
    print(f"   ✗ Error al instanciar: {e}")
    sys.exit(1)

# Mostrar casos cargados
print("\n4. Casos cargados (en orden):")
for i, case_name in enumerate(handler.cases.keys(), 1):
    print(f"   {i}. {case_name}")

# Test de matching con caso 12
print("\n5. TEST DE MATCHING CON CASO 12")
print("-" * 70)

subjects = [
    "FinanStar caso 1 01/11/2025 30/11/2025",
    "FinanStar caso 10 01/11/2025 30/11/2025",
    "FinanStar caso 11 01/11/2025 30/11/2025",
    "FinanStar caso 12 01/11/2025 30/11/2025",
]

for subject in subjects:
    print(f"\nSubject: '{subject}'")
    result = handler.find_matching_case(subject, logger)

    # Extraer el número esperado
    parts = subject.split()
    expected_num = parts[2]
    expected_case = f"case{expected_num}"

    if result == expected_case:
        print(f"   ✓ CORRECTO: Detectado como {result}")
    else:
        print(f"   ✗ ERROR: Detectado como {result}, esperaba {expected_case}")

# Test específico detallado del caso 12
print("\n\n6. TEST DETALLADO DEL CASO 12")
print("=" * 70)

subject_12 = "FinanStar caso 12 01/11/2025 30/11/2025"
print(f"Subject: '{subject_12}'")
print("\nBuscando match...")

result = handler.find_matching_case(subject_12, logger)

print(f"\nResultado final: {result}")

if result == 'case12':
    print("\n" + "=" * 70)
    print("✓✓✓ ÉXITO: EL CASO 12 SE DETECTA CORRECTAMENTE ✓✓✓")
    print("=" * 70)
    print("\nSi la aplicación GUI sigue fallando, necesitas:")
    print("  1. Cerrar COMPLETAMENTE la aplicación (kill process)")
    print("  2. Ejecutar: python3 main.py")
    print("  3. Si usas un .exe compilado, DEBES RECOMPILARLO")
else:
    print("\n" + "=" * 70)
    print(f"✗✗✗ ERROR: Se detectó como {result} en lugar de case12 ✗✗✗")
    print("=" * 70)
