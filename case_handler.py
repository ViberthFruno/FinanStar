# Archivo: case_handler.py
# Ubicación: raíz del proyecto
# Descripción: Manejador principal para cargar y ejecutar casos de respuesta automática

import os
import sys
import importlib.util


class CaseHandler:
    def __init__(self):
        """Inicializa el manejador de casos"""
        self.cases = {}
        self.load_cases()

    def _get_base_path(self):
        """Obtiene la ruta base correcta para desarrollo y PyInstaller"""
        if getattr(sys, 'frozen', False):
            # Corriendo como .exe empaquetado por PyInstaller
            return sys._MEIPASS
        else:
            # Corriendo en desarrollo
            return os.path.dirname(os.path.abspath(__file__))

    def load_cases(self):
        """Carga todos los archivos de casos disponibles"""
        try:
            # Intentar primero la importación explícita (funciona en .exe)
            if self._load_cases_explicit():
                return

            # Si falla, intentar carga dinámica (funciona en desarrollo)
            self._load_cases_dynamic()

        except Exception as e:
            print(f"Error crítico al cargar casos: {str(e)}")

    def _load_cases_explicit(self):
        """Carga casos mediante importación explícita (compatible con PyInstaller)"""
        try:
            # Cargar casos en orden inverso para priorizar números más altos
            # Esto evita que "caso 1" haga match antes que "caso 12"
            case_modules = [
                'case12', 'case11', 'case10', 'case9', 'case8', 'case7',
                'case6', 'case5', 'case4', 'case3', 'case2', 'case1'
            ]

            loaded_count = 0
            for case_name in case_modules:
                try:
                    # Importar el módulo explícitamente
                    case_module = __import__(case_name)

                    # Verificar que el módulo tenga la clase Case
                    if hasattr(case_module, 'Case'):
                        self.cases[case_name] = case_module.Case()
                        print(f"Caso cargado: {case_name}")
                        loaded_count += 1
                    else:
                        print(f"Advertencia: {case_name} no tiene la clase Case")

                except ImportError:
                    # El módulo no existe, continuar con el siguiente
                    continue
                except Exception as e:
                    print(f"Error al cargar {case_name}: {str(e)}")

            # Retornar True si se cargó al menos un caso
            return loaded_count > 0

        except Exception as e:
            print(f"Error en carga explícita de casos: {str(e)}")
            return False

    def _load_cases_dynamic(self):
        """Carga casos dinámicamente desde archivos (solo desarrollo)"""
        try:
            current_dir = self._get_base_path()

            # Verificar si el directorio existe y es accesible
            if not os.path.exists(current_dir):
                print(f"Advertencia: Directorio no encontrado: {current_dir}")
                return

            case_files = [f for f in os.listdir(current_dir) if
                          f.startswith('case') and f.endswith('.py') and f != 'case_handler.py']

            if not case_files:
                print("No se encontraron archivos de casos para cargar")
                return

            for case_file in case_files:
                try:
                    case_name = case_file[:-3]  # Remover .py
                    case_path = os.path.join(current_dir, case_file)

                    # Cargar el módulo dinámicamente
                    spec = importlib.util.spec_from_file_location(case_name, case_path)
                    if spec is None or spec.loader is None:
                        print(f"No se pudo crear spec para {case_file}")
                        continue

                    case_module = importlib.util.module_from_spec(spec)
                    spec.loader.exec_module(case_module)

                    # Verificar que el módulo tenga la clase Case
                    if hasattr(case_module, 'Case'):
                        self.cases[case_name] = case_module.Case()
                        print(f"Caso cargado: {case_name}")
                    else:
                        print(f"Error: {case_file} no tiene la clase Case")

                except Exception as e:
                    print(f"Error al cargar caso {case_file}: {str(e)}")

        except Exception as e:
            print(f"Error en carga dinámica de casos: {str(e)}")

    def get_available_cases(self):
        """Obtiene la lista de casos disponibles"""
        return list(self.cases.keys())

    def get_case_info(self, case_name):
        """Obtiene información de un caso específico"""
        if case_name in self.cases:
            case_obj = self.cases[case_name]
            return {
                'name': case_obj.get_name(),
                'description': case_obj.get_description(),
                'search_keywords': case_obj.get_search_keywords()
            }
        return None

    def execute_case(self, case_name, email_data, logger):
        """Ejecuta un caso específico"""
        if case_name in self.cases:
            try:
                case_obj = self.cases[case_name]
                return case_obj.process_email(email_data, logger)
            except Exception as e:
                logger.log(f"Error al ejecutar caso {case_name}: {str(e)}", level="ERROR")
                return False
        else:
            logger.log(f"Caso no encontrado: {case_name}", level="ERROR")
            return False

    def find_matching_case(self, subject, logger, allowed_cases=None):
        """Busca el primer caso que coincida con el asunto del email"""
        import re

        allowed_set = set(allowed_cases) if allowed_cases else None

        # Log inicial mejorado
        logger.log(f"Buscando caso para: '{subject}'", level="INFO")
        if allowed_set:
            logger.log(f"Casos permitidos: {sorted(allowed_set)}", level="INFO")

        for case_name, case_obj in self.cases.items():
            if allowed_set is not None and case_name not in allowed_set:
                continue
            try:
                keywords = case_obj.get_search_keywords()
                if not keywords:
                    logger.log(f"⚠ {case_name}: SIN KEYWORDS configuradas", level="WARNING")
                    continue

                for keyword in keywords:
                    # Log de cada intento
                    pattern = r'\b' + re.escape(keyword.lower()) + r'\b'
                    match = re.search(pattern, subject.lower())

                    if match:
                        logger.log(f"✓ MATCH: {case_name} | keyword: '{keyword}'", level="INFO")
                        return case_name
                    else:
                        logger.log(f"  ✗ no match: {case_name} | keyword: '{keyword}'", level="DEBUG")

            except Exception as e:
                logger.log(f"Error al verificar caso {case_name}: {str(e)}", level="ERROR")
                continue

        logger.log("⚠ NO SE ENCONTRÓ NINGÚN CASO MATCHING", level="WARNING")
        return None

    def reload_cases(self):
        """Recarga todos los casos disponibles"""
        self.cases.clear()
        self.load_cases()

    def get_case_keywords(self):
        """Obtiene las palabras clave configuradas para cada caso"""
        keywords_map = {}
        for case_name, case_obj in self.cases.items():
            cleaned_keywords = []
            try:
                raw_keywords = case_obj.get_search_keywords()
                if isinstance(raw_keywords, str):
                    raw_keywords = [raw_keywords]
                for keyword in raw_keywords:
                    if not isinstance(keyword, str):
                        continue
                    cleaned = keyword.strip()
                    if cleaned:
                        cleaned_keywords.append(cleaned)

                # Log para debugging
                if cleaned_keywords:
                    print(f"Keywords para {case_name}: {cleaned_keywords}")
                else:
                    print(f"⚠ WARNING: {case_name} NO tiene keywords configuradas")

            except Exception as e:
                print(f"Error al obtener palabras clave para {case_name}: {str(e)}")
                cleaned_keywords = []

            keywords_map[case_name] = cleaned_keywords

        return keywords_map