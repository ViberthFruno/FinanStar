#!/usr/bin/env python3
"""
Script de diagnóstico para verificar el reconocimiento del Caso 12
"""
import sys
import json
import re
from pathlib import Path

print("=" * 70)
print("DIAGNÓSTICO DEL CASO 12")
print("=" * 70)

# 1. Verificar config.json
print("\n1. VERIFICANDO CONFIG.JSON")
print("-" * 70)
config_path = Path(__file__).parent / "config.json"
print(f"Ruta: {config_path}")
print(f"Existe: {config_path.exists()}")

if config_path.exists():
    with open(config_path, 'r', encoding='utf-8') as f:
        config = json.load(f)

    search_params = config.get('search_params', {})
    print(f"\nKeywords configuradas:")
    for caso_key in ['caso1', 'caso10', 'caso11', 'caso12']:
        keyword = search_params.get(caso_key, '')
        print(f"  {caso_key}: '{keyword}'")
else:
    print("⚠ WARNING: config.json NO EXISTE")
    sys.exit(1)

# 2. Verificar case_handler.py
print("\n2. VERIFICANDO CASE_HANDLER.PY")
print("-" * 70)
handler_path = Path(__file__).parent / "case_handler.py"
print(f"Ruta: {handler_path}")

with open(handler_path, 'r', encoding='utf-8') as f:
    handler_content = f.read()

# Verificar que tiene word boundaries
has_word_boundary = r'\b' in handler_content and 're.escape' in handler_content
print(f"Tiene word boundaries (\\b + re.escape): {has_word_boundary}")

# Verificar orden de carga
has_correct_order = "'case12', 'case11', 'case10'" in handler_content
print(f"Orden correcto de carga (case12 primero): {has_correct_order}")

if not has_word_boundary or not has_correct_order:
    print("\n⚠ WARNING: El código NO tiene los cambios aplicados")
    sys.exit(1)

# 3. Test de matching
print("\n3. TEST DE MATCHING")
print("-" * 70)

subjects_to_test = [
    "FinanStar caso 1 01/11/2025 30/11/2025",
    "FinanStar caso 10 01/11/2025 30/11/2025",
    "FinanStar caso 11 01/11/2025 30/11/2025",
    "FinanStar caso 12 01/11/2025 30/11/2025",
]

# Simular find_matching_case
case_order = [
    'case12', 'case11', 'case10', 'case9', 'case8', 'case7',
    'case6', 'case5', 'case4', 'case3', 'case2', 'case1'
]

def find_matching_case_test(subject):
    """Simulación exacta del método find_matching_case"""
    for case_name in case_order:
        caso_key = case_name.replace('case', 'caso')
        keyword = search_params.get(caso_key, '').strip()

        if not keyword:
            continue

        # Exactamente como está en el código
        pattern = r'\b' + re.escape(keyword.lower()) + r'\b'
        if re.search(pattern, subject.lower()):
            return case_name, keyword

    return None, None

for subject in subjects_to_test:
    case_name, keyword = find_matching_case_test(subject)
    expected = subject.split()[2]  # Extrae "1", "10", "11", "12"
    expected_case = f"case{expected}"

    status = "✓" if case_name == expected_case else "✗"
    print(f"{status} Subject: '{subject}'")
    print(f"   → Detectado: {case_name} (keyword: '{keyword}')")
    print(f"   → Esperado: {expected_case}")
    print()

# 4. Test específico caso 12
print("\n4. TEST ESPECÍFICO CASO 12")
print("-" * 70)
subject_caso12 = "FinanStar caso 12 01/11/2025 30/11/2025"
keyword_caso1 = search_params.get('caso1', '').strip()
keyword_caso12 = search_params.get('caso12', '').strip()

print(f"Subject: '{subject_caso12}'")
print(f"\nKeyword Caso 1: '{keyword_caso1}'")
pattern1 = r'\b' + re.escape(keyword_caso1.lower()) + r'\b'
match1 = re.search(pattern1, subject_caso12.lower())
print(f"Pattern: {pattern1}")
print(f"¿Hace match? {bool(match1)}")

print(f"\nKeyword Caso 12: '{keyword_caso12}'")
pattern12 = r'\b' + re.escape(keyword_caso12.lower()) + r'\b'
match12 = re.search(pattern12, subject_caso12.lower())
print(f"Pattern: {pattern12}")
print(f"¿Hace match? {bool(match12)}")

# 5. Verificar módulos case12.py
print("\n5. VERIFICANDO MÓDULO CASE12.PY")
print("-" * 70)
case12_path = Path(__file__).parent / "case12.py"
print(f"Ruta: {case12_path}")
print(f"Existe: {case12_path.exists()}")

if case12_path.exists():
    with open(case12_path, 'r', encoding='utf-8') as f:
        case12_content = f.read()

    has_get_keywords = 'def get_search_keywords' in case12_content
    print(f"Tiene get_search_keywords(): {has_get_keywords}")

    uses_caso12_key = "'caso12'" in case12_content or '"caso12"' in case12_content
    print(f"Usa la key 'caso12' en config: {uses_caso12_key}")

# 6. CONCLUSIÓN
print("\n" + "=" * 70)
print("CONCLUSIÓN")
print("=" * 70)

case_detected, _ = find_matching_case_test("FinanStar caso 12 01/11/2025 30/11/2025")

if case_detected == 'case12':
    print("✓ El código FUNCIONA CORRECTAMENTE")
    print("✓ El caso 12 se detecta correctamente")
    print("\nSi la aplicación sigue fallando:")
    print("  1. Asegúrate de REINICIAR completamente la aplicación")
    print("  2. Si usas PyInstaller, necesitas RECOMPILAR el ejecutable")
    print("  3. Verifica que la app esté usando este config.json")
else:
    print(f"✗ ERROR: Se detecta como {case_detected} en lugar de case12")
    print("\nRevisa los mensajes de warning arriba")
